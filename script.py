from __future__ import annotations

import csv
import re
import shutil
import sys
import time
from dataclasses import dataclass
from pathlib import Path

import pythoncom
import pywintypes
import win32com.client as win32
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.formula import ArrayFormula

from auto import STOCK_ID_BY_TICKER, load_finchat_cache, repair_missing_overview_metrics, save_finchat_cache, scrape_finchat

try:
    import matplotlib

    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
except Exception:
    plt = None


ROOT = Path(__file__).resolve().parent
TEMPLATE_PATH = ROOT / "Reverse DCF Template.xlsm"
COMPANIES_PATH = ROOT / "companies.txt"
GENERATED_DIR = ROOT / "generated"
CSV_DIR = GENERATED_DIR / "csv"
CHART_DIR = GENERATED_DIR / "charts"
INPUTS_PATH = GENERATED_DIR / "reverse_dcf_inputs.xlsx"
SUMMARY_PATH = GENERATED_DIR / "reverse_dcf_summary.xlsx"

INPUT_HEADERS = [
    "Company",
    "Ticker",
    "Price Lookup Ticker",
    "Current Price",
    "Raw Shares Outstanding",
    "Share Unit",
    "Shares Outstanding (million)",
    "FCF",
    "StockBased",
    "Cap",
    "Dep",
    "Diff",
    "FCF Year 1",
    "Perpetuity Growth Rate",
    "Desired Return",
    "Expected EPS LT Growth",
    "Currency",
    "Exchange",
    "ISIN / Code",
    "Price Source",
    "Price Timestamp (UTC)",
]

SUMMARY_HEADERS = [
    "Company",
    "P/E Gap vs 5Y",
    "Growth Gap",
    "Growth Gap (Historical FCF)",
    "Growth Gap (FCF Only)",
    "Growth Gap (Historical FCF, FCF Only)",
    "P/E vs 5Y",
    "Valuation",
    "Expected EPS LT Growth",
    "Historical FCF Growth",
    "Implied Growth Rate",
    "Implied Growth Rate (FCF Only)",
    "Current Forward P/E",
    "P/E Avg 5Y",
    "Comment",
]

DEFAULT_PERPETUITY_GROWTH = 0.03
DEFAULT_DESIRED_RETURN = 0.10
DEFAULT_GROWTH_RATE = 0.27
DEFAULT_TEST_VALUES = {
    "shares": 104.8,
    "fcf": 3916.0,
    "stockbased": 139.0,
    "cap": 1002.0,
    "dep": 621.0,
}

HEADER_FILL_BLUE = PatternFill("solid", fgColor="D9EAF7")
HEADER_FILL_GREEN = PatternFill("solid", fgColor="E2F0D9")
VALUE_FILL_GREEN = PatternFill("solid", fgColor="E2F0D9")
VALUE_FILL_RED = PatternFill("solid", fgColor="FCE4D6")
HEADER_FONT = Font(bold=True)


@dataclass
class CompanyRow:
    company: str
    ticker: str
    finchat_stock_id: str
    lookup_ticker: str
    currency: str
    isin: str
    exchange: str
    generated_filename: str


@dataclass
class InputValues:
    current_price_formula: str | None
    lookup_ticker: str
    current_price: float | None
    raw_shares_formula: str | None
    raw_shares: float | None
    share_unit: str
    shares: float | None
    fcf: float | None
    stockbased: float | None
    cap: float | None
    dep: float | None
    perpetuity_growth: float
    desired_return: float
    growth_rate: float
    historical_fcf_growth: float | None
    implied_growth_rate: float | None
    implied_growth_rate_fcf_only: float | None
    current_forward_pe: float | None
    avg_pe_5y: float | None
    avg_pe_10y: float | None
    pe_valuation: str
    pe_chart_path: str
    pe_history_points: list[dict]
    fcf_history_points: list[dict]
    price_source: str
    price_timestamp: str


def normalize_text(value: str) -> str:
    return (
        value.replace("\u2011", "-")
        .replace("\u2013", "-")
        .replace("\u2014", "-")
        .replace("\u2212", "-")
        .strip()
    )


def sanitize_filename(value: str) -> str:
    cleaned = re.sub(r'[<>:"/\\|?*]+', "", value)
    cleaned = re.sub(r"\s+", " ", cleaned).strip().rstrip(".")
    return cleaned or "company"


def parse_primary_ticker(raw_ticker: str) -> str:
    ticker = normalize_text(raw_ticker)
    if not ticker:
        return ""
    ticker = ticker.split("/")[0].strip()
    if not ticker:
        return ""
    parts = ticker.split()
    if not parts:
        return ""
    ticker = parts[0].strip()
    ticker = ticker.replace("?", "")
    return ticker


def finchat_ticker_key(raw_ticker: str) -> str:
    return parse_primary_ticker(raw_ticker).replace("?", "")


def derive_lookup_ticker(ticker: str, exchange: str) -> str:
    base = parse_primary_ticker(ticker)
    exchange_text = normalize_text(exchange).lower()

    suffix_map = [
        ("london stock exchange", ".L"),
        ("toronto stock exchange", ".TO"),
        ("tsx", ".TO"),
        ("euronext amsterdam", ".AS"),
        ("copenhagen", ".CO"),
        ("warsaw stock exchange", ".WA"),
        ("swedish", ".ST"),
        ("stockholm", ".ST"),
    ]

    for needle, suffix in suffix_map:
        if needle in exchange_text:
            if "." in base:
                return base
            return f"{base}{suffix}"

    return base


def load_companies(path: Path) -> list[CompanyRow]:
    companies: list[CompanyRow] = []
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle, delimiter="\t")
        for row in reader:
            company = normalize_text(row["Stock Name / Company"])
            ticker = normalize_text(row["Ticker"])
            currency = normalize_text(row["Currency"])
            isin = normalize_text(row["ISIN / Code"])
            exchange = normalize_text(row["Exchange"])
            lookup = derive_lookup_ticker(ticker, exchange)
            finchat_key = finchat_ticker_key(ticker)
            filename = f"{sanitize_filename(company)} - Reverse DCF.xlsm"
            companies.append(
                CompanyRow(
                    company=company,
                    ticker=ticker,
                    finchat_stock_id=STOCK_ID_BY_TICKER.get(finchat_key, ""),
                    lookup_ticker=lookup,
                    currency=currency,
                    isin=isin,
                    exchange=exchange,
                    generated_filename=filename,
                )
            )
    return companies


def load_companies_from_inputs(path: Path) -> list[CompanyRow]:
    formula_wb = load_workbook(path, data_only=False)
    value_wb = load_workbook(path, data_only=True)
    if "Inputs" not in formula_wb.sheetnames or "Inputs" not in value_wb.sheetnames:
        return []

    formula_ws = formula_wb["Inputs"]
    value_ws = value_wb["Inputs"]
    headers = {str(formula_ws.cell(row=1, column=col).value or ""): col for col in range(1, formula_ws.max_column + 1)}
    company_col = headers.get("Company", 1)
    ticker_col = headers.get("Ticker", 2)
    lookup_col = headers.get("Price Lookup Ticker", 3)
    currency_col = headers.get("Currency", 17)
    exchange_col = headers.get("Exchange", 18)
    isin_col = headers.get("ISIN / Code", 19)

    companies: list[CompanyRow] = []
    for row_idx in range(2, formula_ws.max_row + 1):
        company = normalize_text(str(value_ws.cell(row=row_idx, column=company_col).value or formula_ws.cell(row=row_idx, column=company_col).value or ""))
        ticker = normalize_text(str(value_ws.cell(row=row_idx, column=ticker_col).value or formula_ws.cell(row=row_idx, column=ticker_col).value or ""))
        if not company or not ticker:
            continue

        currency = normalize_text(str(value_ws.cell(row=row_idx, column=currency_col).value or formula_ws.cell(row=row_idx, column=currency_col).value or ""))
        exchange = normalize_text(str(value_ws.cell(row=row_idx, column=exchange_col).value or formula_ws.cell(row=row_idx, column=exchange_col).value or ""))
        isin = normalize_text(str(value_ws.cell(row=row_idx, column=isin_col).value or formula_ws.cell(row=row_idx, column=isin_col).value or ""))
        lookup = clean_text_value(value_ws.cell(row=row_idx, column=lookup_col).value)
        if not lookup or lookup.startswith("#"):
            lookup = derive_lookup_ticker(ticker, exchange)
        finchat_key = finchat_ticker_key(ticker)
        filename = f"{sanitize_filename(company)} - Reverse DCF.xlsm"

        companies.append(
            CompanyRow(
                company=company,
                ticker=ticker,
                finchat_stock_id=STOCK_ID_BY_TICKER.get(finchat_key, ""),
                lookup_ticker=lookup,
                currency=currency,
                isin=isin,
                exchange=exchange,
                generated_filename=filename,
            )
        )

    return companies


def load_company_roster() -> list[CompanyRow]:
    if INPUTS_PATH.exists():
        companies = load_companies_from_inputs(INPUTS_PATH)
        if companies:
            return companies
        print(f"Inputs workbook has no usable company rows, falling back to: {COMPANIES_PATH.name}")

    if not COMPANIES_PATH.exists():
        print(f"Company list not found: {COMPANIES_PATH}")
        return []
    return load_companies(COMPANIES_PATH)


def ensure_generated_dir() -> None:
    GENERATED_DIR.mkdir(exist_ok=True)
    CSV_DIR.mkdir(exist_ok=True)
    CHART_DIR.mkdir(exist_ok=True)


def autosize_columns(ws) -> None:
    widths: dict[int, int] = {}
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is None:
                continue
            widths[cell.column] = max(widths.get(cell.column, 0), min(len(str(cell.value)), 60))

    for col_idx, width in widths.items():
        ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = width + 2


def used_rows_for_csv(ws) -> list[list[object]]:
    max_row = 0
    max_col = 0
    for row in ws.iter_rows():
        for cell in row:
            if cell.value not in (None, ""):
                max_row = max(max_row, cell.row)
                max_col = max(max_col, cell.column)

    rows: list[list[object]] = []
    if max_row == 0 or max_col == 0:
        return rows

    for row in ws.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
        rows.append([cell.value for cell in row])
    return rows


def export_rows_to_csv(rows: list[list[object]], csv_path: Path) -> None:
    with csv_path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        for row in rows:
            writer.writerow("" if value is None else value for value in row)


def export_workbook_sheet_to_csv(workbook_path: Path, sheet_name: str, csv_path: Path, data_only: bool) -> None:
    workbook = load_workbook(workbook_path, data_only=data_only)
    worksheet = workbook[sheet_name]
    export_rows_to_csv(used_rows_for_csv(worksheet), csv_path)


def excel_call_with_retry(func, attempts: int = 8, delay_seconds: float = 1.0):
    last_error = None
    for _ in range(attempts):
        try:
            return func()
        except pywintypes.com_error as exc:
            last_error = exc
            time.sleep(delay_seconds)
    if last_error is not None:
        raise last_error


def style_header_row(ws, headers: list[str], fill: PatternFill) -> None:
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = HEADER_FONT
        cell.fill = fill


def parse_numeric(value) -> float | None:
    if value in (None, ""):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def clean_text_value(value) -> str:
    text = str(value or "").strip()
    if text in {"", "#VALUE!", "#N/A"}:
        return ""
    return text


def normalize_growth_rate(value: float | None) -> float | None:
    if value is None:
        return None
    if abs(value) > 1:
        return value / 100.0
    return value


def with_zero_default(value: float | None) -> float:
    return 0.0 if value is None else value


def average(values: list[float]) -> float | None:
    if not values:
        return None
    return sum(values) / len(values)


def compute_pe_averages(history_points: list[dict]) -> tuple[float | None, float | None]:
    historical_values: list[float] = []
    for point in history_points:
        label = str(point.get("label", "")).strip().upper()
        value = parse_numeric(point.get("value"))
        if value is None:
            continue
        if label == "LTM":
            continue
        historical_values.append(value)
    avg_5y = average(historical_values[:5])
    avg_10y = average(historical_values[:10])
    return avg_5y, avg_10y


def compute_history_cagr(history_points: list[dict], max_points: int = 10) -> float | None:
    series: list[float] = []
    for point in history_points:
        label = str(point.get("label", "")).strip().upper()
        value = parse_numeric(point.get("value"))
        if value is None or value <= 0 or label == "LTM":
            continue
        series.append(value)
    if len(series) < 2:
        return None
    series = series[:max_points]
    newest = series[0]
    oldest = series[-1]
    periods = len(series) - 1
    if oldest <= 0 or newest <= 0 or periods <= 0:
        return None
    try:
        return (newest / oldest) ** (1 / periods) - 1
    except Exception:
        return None


def classify_pe_valuation(current_forward_pe: float | None, avg_pe_5y: float | None, avg_pe_10y: float | None) -> str:
    comparisons = [value for value in [avg_pe_5y, avg_pe_10y] if value is not None]
    if current_forward_pe is None or not comparisons:
        return ""
    if all(current_forward_pe < value for value in comparisons):
        return "Undervalued vs P/E history"
    if all(current_forward_pe > value for value in comparisons):
        return "Overvalued vs P/E history"
    return "Neutral vs P/E history"


def classify_pe_vs_5y(current_forward_pe: float | None, avg_pe_5y: float | None) -> str:
    if current_forward_pe is None or avg_pe_5y is None:
        return ""
    if current_forward_pe < avg_pe_5y:
        return "Underpriced vs 5Y P/E"
    if current_forward_pe > avg_pe_5y:
        return "Overpriced vs 5Y P/E"
    return "Fairly priced vs 5Y P/E"


def compute_pe_gap_vs_5y(current_forward_pe: float | None, avg_pe_5y: float | None) -> float | None:
    if current_forward_pe is None or avg_pe_5y in (None, 0):
        return None
    return (current_forward_pe - avg_pe_5y) / avg_pe_5y


def first_numeric(*values) -> float | None:
    for value in values:
        parsed = parse_numeric(value)
        if parsed is not None:
            return parsed
    return None


def suggest_share_unit(raw_shares: float | None) -> str:
    if raw_shares is None:
        return "shares"
    if raw_shares >= 1_000_000:
        return "shares"
    if raw_shares >= 1:
        return "millions"
    return "billions"


def extract_formula(value) -> str | None:
    if isinstance(value, ArrayFormula):
        return value.text
    if isinstance(value, str) and value.startswith("="):
        return value
    return None


def cell_value_or_formula(formula_ws, value_ws, row_idx: int, col_idx: int):
    formula_value = formula_ws.cell(row=row_idx, column=col_idx).value
    formula = extract_formula(formula_value)
    if formula is not None:
        return formula
    value = value_ws.cell(row=row_idx, column=col_idx).value
    if value not in (None, ""):
        return value
    return formula_value


def load_existing_inputs() -> dict[str, InputValues]:
    if not INPUTS_PATH.exists():
        return {}

    formula_wb = load_workbook(INPUTS_PATH, data_only=False)
    value_wb = load_workbook(INPUTS_PATH, data_only=True)
    formula_ws = formula_wb["Inputs"]
    value_ws = value_wb["Inputs"]
    headers = {str(formula_ws.cell(row=1, column=col).value or ""): col for col in range(1, formula_ws.max_column + 1)}
    has_share_unit = "Share Unit" in headers
    raw_shares_col = headers.get("Raw Shares Outstanding")
    share_unit_col = headers.get("Share Unit")
    shares_col = headers.get("Shares Outstanding (million)", 5)
    fcf_col = headers.get("FCF", 6 if not has_share_unit else 8)
    stockbased_col = headers.get("StockBased", 7 if not has_share_unit else 9)
    cap_col = headers.get("Cap", 8 if not has_share_unit else 10)
    dep_col = headers.get("Dep", 9 if not has_share_unit else 11)
    pg_col = headers.get("Perpetuity Growth Rate", 12 if not has_share_unit else 14)
    dr_col = headers.get("Desired Return", 13 if not has_share_unit else 15)
    gr_col = headers.get("Growth Rate Year 1-10", 14 if not has_share_unit else 16)
    price_source_col = headers.get("Price Source", 18 if not has_share_unit else 20)
    price_timestamp_col = headers.get("Price Timestamp (UTC)", 19 if not has_share_unit else 21)

    existing: dict[str, InputValues] = {}
    for row_idx in range(2, formula_ws.max_row + 1):
        company = formula_ws.cell(row=row_idx, column=1).value
        if not company:
            continue
        ticker_key = finchat_ticker_key(str(formula_ws.cell(row=row_idx, column=2).value or ""))
        current_price_formula = extract_formula(formula_ws.cell(row=row_idx, column=4).value)
        raw_shares_formula = (
            extract_formula(formula_ws.cell(row=row_idx, column=raw_shares_col).value)
            if raw_shares_col
            else extract_formula(formula_ws.cell(row=row_idx, column=shares_col).value)
        )
        share_unit = str(formula_ws.cell(row=row_idx, column=share_unit_col).value or "millions").strip().lower() if share_unit_col else "millions"
        lookup_ticker = clean_text_value(formula_ws.cell(row=row_idx, column=3).value)
        values = InputValues(
            current_price_formula=current_price_formula,
            lookup_ticker=lookup_ticker,
            current_price=first_numeric(
                value_ws.cell(row=row_idx, column=4).value,
                formula_ws.cell(row=row_idx, column=4).value,
            ),
            raw_shares_formula=raw_shares_formula,
            raw_shares=first_numeric(
                value_ws.cell(row=row_idx, column=raw_shares_col).value if raw_shares_col else None,
                formula_ws.cell(row=row_idx, column=raw_shares_col).value if raw_shares_col else None,
                value_ws.cell(row=row_idx, column=shares_col).value,
                formula_ws.cell(row=row_idx, column=shares_col).value,
            ),
            share_unit=share_unit,
            shares=first_numeric(
                value_ws.cell(row=row_idx, column=shares_col).value,
                formula_ws.cell(row=row_idx, column=shares_col).value,
            ),
            fcf=first_numeric(
                value_ws.cell(row=row_idx, column=fcf_col).value,
                formula_ws.cell(row=row_idx, column=fcf_col).value,
            ),
            stockbased=first_numeric(
                value_ws.cell(row=row_idx, column=stockbased_col).value,
                formula_ws.cell(row=row_idx, column=stockbased_col).value,
            ),
            cap=first_numeric(
                value_ws.cell(row=row_idx, column=cap_col).value,
                formula_ws.cell(row=row_idx, column=cap_col).value,
            ),
            dep=first_numeric(
                value_ws.cell(row=row_idx, column=dep_col).value,
                formula_ws.cell(row=row_idx, column=dep_col).value,
            ),
            perpetuity_growth=first_numeric(
                value_ws.cell(row=row_idx, column=pg_col).value,
                formula_ws.cell(row=row_idx, column=pg_col).value,
            ) or DEFAULT_PERPETUITY_GROWTH,
            desired_return=first_numeric(
                value_ws.cell(row=row_idx, column=dr_col).value,
                formula_ws.cell(row=row_idx, column=dr_col).value,
            ) or DEFAULT_DESIRED_RETURN,
            growth_rate=normalize_growth_rate(first_numeric(
                value_ws.cell(row=row_idx, column=gr_col).value,
                formula_ws.cell(row=row_idx, column=gr_col).value,
            )),
            historical_fcf_growth=None,
            implied_growth_rate=None,
            implied_growth_rate_fcf_only=None,
            current_forward_pe=None,
            avg_pe_5y=None,
            avg_pe_10y=None,
            pe_valuation="",
            pe_chart_path="",
            pe_history_points=[],
            fcf_history_points=[],
            price_source=str(formula_ws.cell(row=row_idx, column=price_source_col).value or "Excel STOCKHISTORY"),
            price_timestamp=str(formula_ws.cell(row=row_idx, column=price_timestamp_col).value or ""),
        )
        existing[str(company)] = values
        if ticker_key:
            existing[f"ticker::{ticker_key}"] = values
    return existing


def build_input_values(company: CompanyRow, existing: dict[str, InputValues]) -> InputValues:
    saved = existing.get(company.company) or existing.get(f"ticker::{finchat_ticker_key(company.ticker)}")
    if saved:
        inferred_share_unit = saved.share_unit or suggest_share_unit(saved.raw_shares)
        if (
            saved.raw_shares_formula
            and saved.share_unit.strip().lower() == "millions"
            and (saved.raw_shares or 0) >= 1_000_000
        ):
            inferred_share_unit = "shares"
        return InputValues(
            current_price_formula=saved.current_price_formula,
            lookup_ticker=clean_text_value(saved.lookup_ticker) or company.lookup_ticker,
            current_price=saved.current_price,
            raw_shares_formula=saved.raw_shares_formula,
            raw_shares=saved.raw_shares,
            share_unit=inferred_share_unit,
            shares=saved.shares,
            fcf=with_zero_default(saved.fcf),
            stockbased=with_zero_default(saved.stockbased),
            cap=with_zero_default(saved.cap),
            dep=with_zero_default(saved.dep),
            perpetuity_growth=saved.perpetuity_growth,
            desired_return=saved.desired_return,
            growth_rate=saved.growth_rate,
            historical_fcf_growth=saved.historical_fcf_growth,
            implied_growth_rate=saved.implied_growth_rate,
            implied_growth_rate_fcf_only=saved.implied_growth_rate_fcf_only,
            current_forward_pe=saved.current_forward_pe,
            avg_pe_5y=saved.avg_pe_5y,
            avg_pe_10y=saved.avg_pe_10y,
            pe_valuation=saved.pe_valuation,
            pe_chart_path=saved.pe_chart_path,
            pe_history_points=saved.pe_history_points,
            fcf_history_points=saved.fcf_history_points,
            price_source="Excel STOCKHISTORY",
            price_timestamp=saved.price_timestamp,
        )

    return InputValues(
        current_price_formula=None,
        lookup_ticker=company.lookup_ticker,
        current_price=None,
        raw_shares_formula=None,
        raw_shares=None,
        share_unit="shares",
        shares=DEFAULT_TEST_VALUES["shares"],
        fcf=DEFAULT_TEST_VALUES["fcf"],
        stockbased=DEFAULT_TEST_VALUES["stockbased"],
        cap=DEFAULT_TEST_VALUES["cap"],
        dep=DEFAULT_TEST_VALUES["dep"],
        perpetuity_growth=DEFAULT_PERPETUITY_GROWTH,
        desired_return=DEFAULT_DESIRED_RETURN,
        growth_rate=None,
        historical_fcf_growth=None,
        implied_growth_rate=None,
        implied_growth_rate_fcf_only=None,
        current_forward_pe=None,
        avg_pe_5y=None,
        avg_pe_10y=None,
        pe_valuation="",
        pe_chart_path="",
        pe_history_points=[],
        fcf_history_points=[],
        price_source="Excel STOCKHISTORY",
        price_timestamp="",
    )


def create_default_inputs_workbook(
    companies: list[CompanyRow],
    existing: dict[str, InputValues] | None = None,
) -> dict[str, InputValues]:
    existing = existing or {}
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Inputs"
    style_header_row(sheet, INPUT_HEADERS, HEADER_FILL_BLUE)

    values_by_company: dict[str, InputValues] = {}
    for row_idx, company in enumerate(companies, start=2):
        values = build_input_values(company, existing)
        values_by_company[company.company] = values

        sheet.cell(row=row_idx, column=1, value=company.company)
        sheet.cell(row=row_idx, column=2, value=company.ticker)
        sheet.cell(row=row_idx, column=3, value=values.lookup_ticker)
        current_price_formula = values.current_price_formula or f'=_FV(B{row_idx},"Price")'
        raw_shares_formula = values.raw_shares_formula or f'=_FV(B{row_idx},"Shares outstanding",TRUE)'
        raw_shares_value = raw_shares_formula if values.raw_shares is None else values.raw_shares
        sheet.cell(row=row_idx, column=4, value=current_price_formula)
        sheet.cell(row=row_idx, column=5, value=raw_shares_value)
        sheet.cell(row=row_idx, column=6, value=values.share_unit or suggest_share_unit(values.raw_shares))
        sheet.cell(
            row=row_idx,
            column=7,
            value=f'=IF(E{row_idx}="","",IF(LOWER(F{row_idx})="shares",E{row_idx}/1000000,IF(LOWER(F{row_idx})="billions",E{row_idx}*1000,E{row_idx})))',
        )
        sheet.cell(row=row_idx, column=8, value=values.fcf)
        sheet.cell(row=row_idx, column=9, value=values.stockbased)
        sheet.cell(row=row_idx, column=10, value=values.cap)
        sheet.cell(row=row_idx, column=11, value=values.dep)
        sheet.cell(row=row_idx, column=12, value=f'=IF(OR(J{row_idx}="",K{row_idx}=""),"",J{row_idx}-K{row_idx})')
        sheet.cell(row=row_idx, column=13, value=f'=IF(COUNTA(H{row_idx}:L{row_idx})<5,"",H{row_idx}-I{row_idx}+L{row_idx})')
        sheet.cell(row=row_idx, column=14, value=values.perpetuity_growth)
        sheet.cell(row=row_idx, column=15, value=values.desired_return)
        sheet.cell(row=row_idx, column=16, value=values.growth_rate)
        sheet.cell(row=row_idx, column=17, value=company.currency)
        sheet.cell(row=row_idx, column=18, value=company.exchange)
        sheet.cell(row=row_idx, column=19, value=company.isin)
        sheet.cell(row=row_idx, column=20, value=values.price_source)
        sheet.cell(row=row_idx, column=21, value=values.price_timestamp)

    note_row = len(companies) + 4
    sheet.cell(
        row=note_row,
        column=1,
        value="Column D and E keep the live Excel stock formulas. Column F is the share unit, and column G normalizes to millions for the template sync.",
    )
    sheet.cell(
        row=note_row + 1,
        column=1,
        value="Use Share Unit values like shares, millions, or billions. Save the master workbook in Excel, then rerun this script to sync all company files and the summary.",
    )

    for cell in sheet["D"]:
        cell.number_format = "0.00"
    for col_letter in "EGHIJKLM":
        for cell in sheet[col_letter]:
            cell.number_format = "0.00"
    for col_letter in "NOP":
        for cell in sheet[col_letter]:
            cell.number_format = "0.00%"

    autosize_columns(sheet)
    workbook.save(INPUTS_PATH)
    export_workbook_sheet_to_csv(INPUTS_PATH, "Inputs", CSV_DIR / "reverse_dcf_inputs.csv", data_only=True)
    return values_by_company


def load_inputs_as_source_of_truth(companies: list[CompanyRow]) -> dict[str, InputValues]:
    if not INPUTS_PATH.exists():
        return create_default_inputs_workbook(companies)

    existing = load_existing_inputs()
    headers_wb = load_workbook(INPUTS_PATH, data_only=False)
    headers_ws = headers_wb["Inputs"]
    headers = {str(headers_ws.cell(row=1, column=col).value or ""): col for col in range(1, headers_ws.max_column + 1)}

    if "Share Unit" not in headers or "Raw Shares Outstanding" not in headers:
        values_by_company = create_default_inputs_workbook(companies, existing)
    else:
        export_workbook_sheet_to_csv(INPUTS_PATH, "Inputs", CSV_DIR / "reverse_dcf_inputs.csv", data_only=True)
        values_by_company = {}
        for company in companies:
            values_by_company[company.company] = build_input_values(company, existing)

    return values_by_company


def apply_finchat_cache_to_inputs(companies: list[CompanyRow]) -> None:
    if not INPUTS_PATH.exists():
        return

    finchat_cache = load_finchat_cache()
    if not finchat_cache:
        return

    pythoncom.CoInitialize()
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    workbook = None
    try:
        try:
            workbook = excel_call_with_retry(lambda: excel.Workbooks.Open(str(INPUTS_PATH.resolve()), UpdateLinks=0))
        except pywintypes.com_error:
            print(f"Skipped FinChat writeback because the input workbook is locked: {INPUTS_PATH.name}")
            return
        sheet = excel_call_with_retry(lambda: workbook.Worksheets("Inputs"))
        excel.ScreenUpdating = False
        excel.EnableEvents = False

        def set_excel_cell(row_idx: int, col_idx: int, value) -> None:
            excel_call_with_retry(lambda: setattr(sheet.Cells(row_idx, col_idx), "Value", value))

        def clear_excel_cell(row_idx: int, col_idx: int) -> None:
            excel_call_with_retry(lambda: sheet.Cells(row_idx, col_idx).ClearContents())

        for row_idx, company in enumerate(companies, start=2):
            stock_id = company.finchat_stock_id
            if not stock_id or stock_id not in finchat_cache:
                continue
            cached = finchat_cache[stock_id]

            set_excel_cell(row_idx, 8, cached["fcf_value"] if cached.get("fcf_value") is not None else 0)
            set_excel_cell(row_idx, 9, cached["stockbased_value"] if cached.get("stockbased_value") is not None else 0)
            set_excel_cell(row_idx, 10, abs(cached["cap_value"]) if cached.get("cap_value") is not None else 0)
            set_excel_cell(row_idx, 11, cached["dep_value"] if cached.get("dep_value") is not None else 0)
            if cached.get("eps_lt_growth_value") is not None:
                set_excel_cell(row_idx, 16, normalize_growth_rate(cached["eps_lt_growth_value"]))
            else:
                clear_excel_cell(row_idx, 16)

        excel_call_with_retry(lambda: excel.CalculateFullRebuild())
        excel_call_with_retry(lambda: workbook.Save())
    except pywintypes.com_error:
        print(f"Skipped FinChat writeback because Excel COM became unavailable: {INPUTS_PATH.name}")
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        try:
            excel.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()

    export_workbook_sheet_to_csv(INPUTS_PATH, "Inputs", CSV_DIR / "reverse_dcf_inputs.csv", data_only=True)


def recalculate_workbooks_with_excel(paths: list[Path]) -> None:
    if not paths:
        return

    pythoncom.CoInitialize()
    excel = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        excel.Calculation = -4105  # xlCalculationAutomatic

        for path in paths:
            workbook = None
            try:
                workbook = excel_call_with_retry(lambda: excel.Workbooks.Open(str(path.resolve()), UpdateLinks=0))
                workbook.ForceFullCalculation = True
            except Exception:
                try:
                    if workbook is not None:
                        workbook.Close(SaveChanges=False)
                except Exception:
                    pass
                continue

            try:
                excel_call_with_retry(lambda: workbook.Application.CalculateFullRebuild())
                excel_call_with_retry(lambda: workbook.Save())
            except Exception:
                pass
            finally:
                try:
                    workbook.Close(SaveChanges=False)
                except Exception:
                    pass
    except Exception:
        print("Skipped Excel recalc pass for generated workbooks.")
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def refresh_finchat_cache(companies: list[CompanyRow], headless: bool = True) -> bool:
    stock_ids = [company.finchat_stock_id for company in companies if company.finchat_stock_id]
    if not stock_ids:
        return False
    rows = scrape_finchat(stock_ids, headless=headless)
    save_finchat_cache(rows)
    try:
        repair_missing_overview_metrics(headless=headless)
    except Exception as exc:
        print(f"Overview metric repair skipped: {exc}")
    return True


def compute_reverse_dcf(
    values: InputValues,
    growth_rate: float | None = None,
    fcf_year_1_override: float | None = None,
) -> tuple[float | None, float | None]:
    chosen_growth_rate = growth_rate if growth_rate is not None else values.implied_growth_rate if values.implied_growth_rate is not None else values.growth_rate
    required = [
        values.shares,
        values.fcf,
        values.stockbased,
        values.cap,
        values.dep,
        values.perpetuity_growth,
        values.desired_return,
        chosen_growth_rate,
    ]
    if any(v is None for v in required):
        return None, None
    if not values.shares or values.desired_return <= values.perpetuity_growth or chosen_growth_rate is None:
        return None, None

    diff = values.cap - values.dep
    fcf_year_1 = fcf_year_1_override if fcf_year_1_override is not None else values.fcf - values.stockbased + diff
    yearly_fcfs = []
    current_fcf = fcf_year_1
    for _ in range(10):
        yearly_fcfs.append(current_fcf)
        current_fcf *= 1 + chosen_growth_rate

    discounted_fcfs = [
        fcf / ((1 + values.desired_return) ** year)
        for year, fcf in enumerate(yearly_fcfs, start=1)
    ]
    terminal_value = (yearly_fcfs[-1] * (1 + values.perpetuity_growth)) / (
        values.desired_return - values.perpetuity_growth
    )
    present_value_tv = terminal_value / ((1 + values.desired_return) ** 10)
    total_equity_value = sum(discounted_fcfs) + present_value_tv
    intrinsic_value = total_equity_value / values.shares

    if values.current_price in (None, 0):
        upside = None
    else:
        upside = (intrinsic_value - values.current_price) / values.current_price

    return intrinsic_value, upside


def solve_implied_growth_rate(values: InputValues, fcf_year_1_override: float | None = None) -> float | None:
    if values.current_price in (None, 0):
        return None

    low = -0.99
    high = 1.5
    low_value, _ = compute_reverse_dcf(values, growth_rate=low, fcf_year_1_override=fcf_year_1_override)
    high_value, _ = compute_reverse_dcf(values, growth_rate=high, fcf_year_1_override=fcf_year_1_override)
    if low_value is None or high_value is None:
        return None

    target = values.current_price
    low_diff = low_value - target
    high_diff = high_value - target

    for _ in range(12):
        if low_diff == 0:
            return low
        if high_diff == 0:
            return high
        if low_diff * high_diff < 0:
            break
        high += 1.0
        high_value, _ = compute_reverse_dcf(values, growth_rate=high, fcf_year_1_override=fcf_year_1_override)
        if high_value is None:
            return None
        high_diff = high_value - target
    else:
        return None

    for _ in range(80):
        mid = (low + high) / 2
        mid_value, _ = compute_reverse_dcf(values, growth_rate=mid, fcf_year_1_override=fcf_year_1_override)
        if mid_value is None:
            return None
        mid_diff = mid_value - target
        if abs(mid_diff) < 0.01:
            return mid
        if low_diff * mid_diff <= 0:
            high = mid
            high_diff = mid_diff
        else:
            low = mid
            low_diff = mid_diff

    return (low + high) / 2


def build_summary_comment(values: InputValues) -> str:
    notes: list[str] = []
    if values.current_price in (None, 0):
        notes.append("Current price missing from Excel stock data")
    if values.shares in (None, 0):
        notes.append("Shares outstanding missing or zero")
    if values.fcf in (None, 0):
        notes.append("FCF missing in scrape, defaulted to 0")
    if values.stockbased in (None, 0):
        notes.append("Stock-based compensation missing in scrape, defaulted to 0")
    if values.cap in (None, 0):
        notes.append("Capex missing in scrape, defaulted to 0")
    if values.dep in (None, 0):
        notes.append("Depreciation missing in scrape, defaulted to 0")
    if values.growth_rate is None:
        notes.append("Expected EPS LT growth not found on Fiscal.ai")
    if values.historical_fcf_growth is None:
        notes.append("Historical FCF growth not available from Fiscal.ai")
    if values.implied_growth_rate is None:
        notes.append("Implied growth could not be solved from current inputs")
    if values.implied_growth_rate_fcf_only is None:
        notes.append("FCF-only implied growth could not be solved from current inputs")
    if values.current_forward_pe is None:
        notes.append("Current forward P/E not found on Fiscal.ai")
    if values.avg_pe_5y is None and values.avg_pe_10y is None:
        notes.append("Historical P/E series not available for 5Y/10Y average")
    return "; ".join(notes)


def apply_finchat_market_metrics(
    companies: list[CompanyRow],
    values_by_company: dict[str, InputValues],
) -> None:
    finchat_cache = load_finchat_cache()
    for company in companies:
        values = values_by_company[company.company]
        cached = finchat_cache.get(company.finchat_stock_id, {})
        history_points = cached.get("pe_history_values") or []
        fcf_history_points = cached.get("fcf_history_values") or []
        avg_pe_5y, avg_pe_10y = compute_pe_averages(history_points)
        cached_growth = normalize_growth_rate(parse_numeric(cached.get("eps_lt_growth_value")))
        if cached_growth is not None:
            values.growth_rate = cached_growth
        values.historical_fcf_growth = compute_history_cagr(fcf_history_points)
        current_forward_pe = parse_numeric(cached.get("forward_pe_value"))
        if current_forward_pe is None and history_points:
            current_forward_pe = parse_numeric(history_points[0].get("value"))
        values.current_forward_pe = current_forward_pe
        values.avg_pe_5y = avg_pe_5y
        values.avg_pe_10y = avg_pe_10y
        values.pe_history_points = history_points
        values.fcf_history_points = fcf_history_points
        values.pe_valuation = classify_pe_valuation(values.current_forward_pe, avg_pe_5y, avg_pe_10y)


def create_company_workbooks(companies: list[CompanyRow], values_by_company: dict[str, InputValues]) -> None:
    for company in companies:
        values = values_by_company[company.company]
        target_path = GENERATED_DIR / company.generated_filename
        try:
            shutil.copyfile(TEMPLATE_PATH, target_path)
        except PermissionError:
            print(f"Skipped locked workbook: {target_path.name}")
            continue

        workbook = load_workbook(target_path, keep_vba=True)
        workbook.calculation.calcMode = "auto"
        workbook.calculation.fullCalcOnLoad = True
        workbook.calculation.forceFullCalc = True

        sheet = workbook.worksheets[0]
        for extra_sheet in workbook.worksheets[1:]:
            workbook.remove(extra_sheet)

        def populate_sheet(target_sheet, *, fcf_only: bool) -> None:
            target_sheet["A1"] = company.company if not fcf_only else f"{company.company} (FCF Only)"
            target_sheet["A1"].font = Font(bold=True)
            target_sheet["C5"] = values.current_price
            target_sheet["C6"] = values.shares
            target_sheet["E7"] = values.fcf
            target_sheet["F7"] = 0 if fcf_only else values.stockbased
            target_sheet["G7"] = 0 if fcf_only else values.cap
            target_sheet["H7"] = 0 if fcf_only else values.dep
            target_sheet["C7"] = values.perpetuity_growth
            target_sheet["C8"] = values.desired_return
            target_sheet["C11"] = (
                values.implied_growth_rate_fcf_only if fcf_only and values.implied_growth_rate_fcf_only is not None
                else values.implied_growth_rate if not fcf_only and values.implied_growth_rate is not None
                else values.growth_rate
            )

        populate_sheet(sheet, fcf_only=False)

        fcf_only_sheet = workbook.copy_worksheet(sheet)
        fcf_only_sheet.title = "Reverse DCF - FCF Only"
        populate_sheet(fcf_only_sheet, fcf_only=True)

        workbook.save(target_path)
        export_rows_to_csv(used_rows_for_csv(sheet), CSV_DIR / f"{target_path.stem}.csv")
        export_rows_to_csv(
            used_rows_for_csv(fcf_only_sheet),
            CSV_DIR / f"{target_path.stem} - FCF Only.csv",
        )


def create_pe_chart(company: CompanyRow, values: InputValues) -> str:
    if plt is None or values.current_forward_pe is None or not values.pe_history_points:
        return ""

    points = [
        point
        for point in values.pe_history_points
        if str(point.get("label", "")).strip().upper() != "LTM" and parse_numeric(point.get("value")) is not None
    ]
    if not points:
        return ""

    labels = [str(point["label"]) for point in points]
    series = [float(point["value"]) for point in points]
    fig, ax = plt.subplots(figsize=(10, 5))
    if len(points) <= 2:
        comparison_labels = [f"Hist {labels[-1]}", "Current Fwd P/E"]
        comparison_values = [series[-1], values.current_forward_pe]
        comparison_colors = ["#2F6690", "#C84C09"]
        if values.avg_pe_5y is not None:
            comparison_labels.append("5Y Avg")
            comparison_values.append(values.avg_pe_5y)
            comparison_colors.append("#2A9D8F")
        if values.avg_pe_10y is not None:
            comparison_labels.append("10Y Avg")
            comparison_values.append(values.avg_pe_10y)
            comparison_colors.append("#6D597A")

        bars = ax.bar(comparison_labels, comparison_values, color=comparison_colors, width=0.62)
        for bar, value in zip(bars, comparison_values):
            ax.text(
                bar.get_x() + bar.get_width() / 2,
                value,
                f"{value:.1f}",
                ha="center",
                va="bottom",
                fontsize=10,
            )
        subtitle = "Limited Fiscal.ai history available"
        if values.avg_pe_5y is not None:
            gap_5y = (values.current_forward_pe / values.avg_pe_5y - 1) * 100
            subtitle = f"Current forward P/E is {gap_5y:.1f}% vs 5Y average"
        ax.set_title(f"{company.company} Forward P/E Snapshot\n{subtitle}")
        ax.set_xlabel("")
    else:
        ax.plot(labels, series, color="#2F6690", marker="o", linewidth=2, label="Historical P/E")
        ax.axhline(values.current_forward_pe, color="#C84C09", linestyle="--", linewidth=2, label="Current Forward P/E")
        if values.avg_pe_5y is not None:
            ax.axhline(values.avg_pe_5y, color="#2A9D8F", linestyle=":", linewidth=2, label="5Y Avg P/E")
        if values.avg_pe_10y is not None:
            ax.axhline(values.avg_pe_10y, color="#6D597A", linestyle="-.", linewidth=2, label="10Y Avg P/E")
        ax.set_title(f"{company.company} P/E vs History")
        ax.set_xlabel("Period")
        ax.tick_params(axis="x", rotation=45)
        ax.legend()

    ax.set_ylabel("P/E")
    ax.grid(axis="y", alpha=0.25)
    fig.tight_layout()

    chart_name = f"{sanitize_filename(company.company)} - PE History.png"
    chart_path = CHART_DIR / chart_name
    fig.savefig(chart_path, dpi=160)
    plt.close(fig)
    return str(chart_path.name)


def create_summary_workbook(companies: list[CompanyRow], values_by_company: dict[str, InputValues]) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Summary"
    style_header_row(sheet, SUMMARY_HEADERS, HEADER_FILL_GREEN)

    for row_idx, company in enumerate(companies, start=2):
        values = values_by_company[company.company]
        expected_growth = values.growth_rate
        historical_fcf_growth = values.historical_fcf_growth
        implied_growth = values.implied_growth_rate
        implied_growth_fcf_only = values.implied_growth_rate_fcf_only
        growth_gap = None if expected_growth is None or implied_growth is None else expected_growth - implied_growth
        growth_gap_historical = None if historical_fcf_growth is None or implied_growth is None else historical_fcf_growth - implied_growth
        growth_gap_fcf_only = None if expected_growth is None or implied_growth_fcf_only is None else expected_growth - implied_growth_fcf_only
        growth_gap_historical_fcf_only = None if historical_fcf_growth is None or implied_growth_fcf_only is None else historical_fcf_growth - implied_growth_fcf_only
        comment = build_summary_comment(values)
        pe_gap_vs_5y = compute_pe_gap_vs_5y(values.current_forward_pe, values.avg_pe_5y)
        pe_vs_5y = classify_pe_vs_5y(values.current_forward_pe, values.avg_pe_5y)
        if growth_gap is None:
            valuation = ""
        elif growth_gap > 0:
            valuation = "Undervalued"
        else:
            valuation = "Overvalued"
        expected_growth_display = "NA" if expected_growth is None else expected_growth

        sheet.cell(row=row_idx, column=1, value=company.company)
        sheet.cell(row=row_idx, column=2, value=pe_gap_vs_5y)
        sheet.cell(row=row_idx, column=3, value=growth_gap)
        sheet.cell(row=row_idx, column=4, value=growth_gap_historical)
        sheet.cell(row=row_idx, column=5, value=growth_gap_fcf_only)
        sheet.cell(row=row_idx, column=6, value=growth_gap_historical_fcf_only)
        sheet.cell(row=row_idx, column=7, value=pe_vs_5y)
        sheet.cell(row=row_idx, column=8, value=valuation)
        sheet.cell(row=row_idx, column=9, value=expected_growth_display)
        sheet.cell(row=row_idx, column=10, value="NA" if historical_fcf_growth is None else historical_fcf_growth)
        sheet.cell(row=row_idx, column=11, value=implied_growth)
        sheet.cell(row=row_idx, column=12, value=implied_growth_fcf_only)
        sheet.cell(row=row_idx, column=13, value=values.current_forward_pe)
        sheet.cell(row=row_idx, column=14, value=values.avg_pe_5y)
        sheet.cell(row=row_idx, column=15, value=comment)

    for col_letter in "MN":
        for cell in sheet[col_letter]:
            cell.number_format = "0.00"
    for col_letter in "BCDEFJKL":
        for cell in sheet[col_letter]:
            cell.number_format = "0.00%"

    for row_idx in range(2, sheet.max_row + 1):
        pe_gap = sheet.cell(row=row_idx, column=2).value
        dcf_gap = sheet.cell(row=row_idx, column=3).value
        dcf_gap_historical = sheet.cell(row=row_idx, column=4).value
        dcf_gap_fcf_only = sheet.cell(row=row_idx, column=5).value
        dcf_gap_historical_fcf_only = sheet.cell(row=row_idx, column=6).value
        pe_valuation = sheet.cell(row=row_idx, column=7)
        dcf_valuation = sheet.cell(row=row_idx, column=8)

        if isinstance(pe_gap, (int, float)):
            sheet.cell(row=row_idx, column=2).fill = VALUE_FILL_GREEN if pe_gap < 0 else VALUE_FILL_RED
        if isinstance(dcf_gap, (int, float)):
            sheet.cell(row=row_idx, column=3).fill = VALUE_FILL_GREEN if dcf_gap > 0 else VALUE_FILL_RED
        if isinstance(dcf_gap_historical, (int, float)):
            sheet.cell(row=row_idx, column=4).fill = VALUE_FILL_GREEN if dcf_gap_historical > 0 else VALUE_FILL_RED
        if isinstance(dcf_gap_fcf_only, (int, float)):
            sheet.cell(row=row_idx, column=5).fill = VALUE_FILL_GREEN if dcf_gap_fcf_only > 0 else VALUE_FILL_RED
        if isinstance(dcf_gap_historical_fcf_only, (int, float)):
            sheet.cell(row=row_idx, column=6).fill = VALUE_FILL_GREEN if dcf_gap_historical_fcf_only > 0 else VALUE_FILL_RED
        if dcf_valuation.value == "Undervalued":
            dcf_valuation.fill = VALUE_FILL_GREEN
        elif dcf_valuation.value == "Overvalued":
            dcf_valuation.fill = VALUE_FILL_RED

        if isinstance(pe_valuation.value, str):
            if pe_valuation.value.startswith("Underpriced"):
                pe_valuation.fill = VALUE_FILL_GREEN
            elif pe_valuation.value.startswith("Overpriced"):
                pe_valuation.fill = VALUE_FILL_RED

    autosize_columns(sheet)
    export_rows_to_csv(used_rows_for_csv(sheet), CSV_DIR / "reverse_dcf_summary.csv")
    try:
        workbook.save(SUMMARY_PATH)
    except PermissionError:
        print(f"Skipped locked summary workbook: {SUMMARY_PATH.name}")


def sync_from_inputs(refresh_finchat: bool = False) -> int:
    if not TEMPLATE_PATH.exists():
        print(f"Template not found: {TEMPLATE_PATH}")
        return 1

    companies = load_company_roster()
    if not companies:
        print("No companies found in the inputs workbook or companies.txt.")
        return 1
    ensure_generated_dir()
    if refresh_finchat:
        try:
            refreshed = refresh_finchat_cache(companies)
            if refreshed:
                print("Refreshed FinChat cache.")
        except Exception as exc:
            print(f"FinChat refresh skipped: {exc}")
    values_by_company = load_inputs_as_source_of_truth(companies)
    apply_finchat_cache_to_inputs(companies)
    values_by_company = load_inputs_as_source_of_truth(companies)
    apply_finchat_market_metrics(companies, values_by_company)
    for values in values_by_company.values():
        values.implied_growth_rate = solve_implied_growth_rate(values)
        values.implied_growth_rate_fcf_only = solve_implied_growth_rate(values, fcf_year_1_override=values.fcf)
    for company in companies:
        values_by_company[company.company].pe_chart_path = create_pe_chart(company, values_by_company[company.company])
    create_company_workbooks(companies, values_by_company)
    recalculate_workbooks_with_excel([GENERATED_DIR / company.generated_filename for company in companies if (GENERATED_DIR / company.generated_filename).exists()])
    create_summary_workbook(companies, values_by_company)

    print(f"Using input source: {INPUTS_PATH}")
    print(f"Created: {SUMMARY_PATH}")
    print(f"Created {len(companies)} company workbooks in: {GENERATED_DIR}")
    print("Open the master workbook in Excel, let formulas calculate, save it, and this sync will copy the saved values into the company files and summary.")
    return 0


def watch_inputs(poll_seconds: float = 2.0) -> int:
    print(f"Watching for changes: {INPUTS_PATH}")
    print("Save the input workbook in Excel and the sync will rerun automatically.")

    last_mtime = None
    if INPUTS_PATH.exists():
        last_mtime = INPUTS_PATH.stat().st_mtime

    initial_result = sync_from_inputs()
    if initial_result != 0:
        return initial_result

    while True:
        try:
            if INPUTS_PATH.exists():
                current_mtime = INPUTS_PATH.stat().st_mtime
                if last_mtime is None or current_mtime > last_mtime:
                    last_mtime = current_mtime
                    print("")
                    print("Detected input workbook save, syncing...")
                    result = sync_from_inputs()
                    if result != 0:
                        print("Sync failed. Fix the issue and save the workbook again to retry.")
            time.sleep(poll_seconds)
        except KeyboardInterrupt:
            print("")
            print("Stopped watch mode.")
            return 0


def main() -> int:
    if len(sys.argv) > 1 and sys.argv[1] == "--watch":
        return watch_inputs()
    if len(sys.argv) > 1 and sys.argv[1] == "--refresh-finchat":
        return sync_from_inputs(refresh_finchat=True)
    return sync_from_inputs()


if __name__ == "__main__":
    sys.exit(main())
