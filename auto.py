from __future__ import annotations

import csv
import json
import re
import sys
from dataclasses import dataclass
from pathlib import Path


ROOT = Path(__file__).resolve().parent
GENERATED_DIR = ROOT / "generated"
FINCHAT_CSV_PATH = GENERATED_DIR / "finchat_scraped_data.csv"
FINCHAT_JSON_PATH = GENERATED_DIR / "finchat_scraped_data.json"
FISCAL_PROFILE_DIR = GENERATED_DIR / "fiscal_profile"
    # "BRO": "NYSE-BRO",
    # "NOVO-B": "CPSE-NOVO%20B",
STOCK_ID_BY_TICKER = {
    "WISE": "LSE-WISE",
    "MELI": "NasdaqGS-MELI",
    "AMZN": "NasdaqGS-AMZN",
    "TOCP": "TSXV-TOI",
    "GOOGL": "NasdaqGS-GOOGL",
    "CSU": "TSX-CSU",
    "RKO": "OM-ROKO%20B",
    "CPRT": "NasdaqGS-CPRT",
    "MEDP": "NasdaqGS-MEDP",
    "CHAPS": "XTRA-CHG",
    "DIN": "WSE-DNP",
    "ASML": "ENXTAM-ASML",
    "DUOL": "NasdaqGS-DUOL",
    "III": "LSE-III",
    "RMS": "ENXTPA-RMS",
    "SPGI": "NYSE-SPGI",
    
}

METRIC_LABELS = {
    "fcf": "Free Cash Flow (Year 1 Mean)",
    "fcf_history": "FCF History",
    "stockbased": "Share-Based Compensation Expense",
    "cap": "Capital Expenditure",
    "dep": "Depreciation & Amortization",
    "shares": "Total Shares Outstanding",
    "eps_lt_growth": "EPS LT Growth Est.",
    "forward_pe": "Forward P/E (NTM)",
    "pe_history": "P/E History",
}

METRIC_ALIASES = {
    "shares": ["Total Shares Outstanding"],
    "stockbased": ["Stock-Based Compensation", "Share-Based Compensation Expense"],
    "cap": ["Capital Expenditure", "Capital Expenditures"],
    "dep": ["Depreciation & Amortization", "Depreciation & Amortization, Total"],
}

FINANCIAL_TABS = {
    "income-statement": [METRIC_LABELS["shares"]],
    "cash-flow-statement": [
        METRIC_LABELS["stockbased"],
        METRIC_LABELS["cap"],
        METRIC_LABELS["dep"],
    ],
}


@dataclass
class FinchatMetrics:
    stock_id: str
    current_price_text: str = ""
    shares_text: str = ""
    fcf_text: str = ""
    fcf_history_text: str = ""
    stockbased_text: str = ""
    cap_text: str = ""
    dep_text: str = ""
    eps_lt_growth_text: str = ""
    forward_pe_text: str = ""
    pe_history_text: str = ""

    def to_row(self) -> dict[str, str]:
        return {
            "Stock": self.stock_id,
            "Current Price": self.current_price_text,
            METRIC_LABELS["shares"]: self.shares_text,
            METRIC_LABELS["fcf"]: self.fcf_text,
            METRIC_LABELS["fcf_history"]: self.fcf_history_text,
            METRIC_LABELS["stockbased"]: self.stockbased_text,
            METRIC_LABELS["cap"]: self.cap_text,
            METRIC_LABELS["dep"]: self.dep_text,
            METRIC_LABELS["eps_lt_growth"]: self.eps_lt_growth_text,
            METRIC_LABELS["forward_pe"]: self.forward_pe_text,
            METRIC_LABELS["pe_history"]: self.pe_history_text,
        }


def parse_metric_text(value: str) -> float | None:
    if value is None:
        return None
    text = str(value).strip()
    if not text or text.upper() == "N/A":
        return None
    text = text.replace("\xa0", " ")
    text = text.replace("US$", "").replace(",", "").replace("$", "").replace("€", "").replace("£", "")
    for token in ["US$", "CAD", "USD", "EUR", "PLN", "DKK", "SEK", "GBP"]:
        text = text.replace(token, "")
    negative = False
    if text.startswith("(") and text.endswith(")"):
        negative = True
        text = text[1:-1]
    text = text.strip()
    multiplier = 1.0
    suffix_map = {
        "K": 1e3,
        "M": 1e6,
        "B": 1e9,
        "T": 1e12,
    }
    if text and text[-1].upper() in suffix_map:
        multiplier = suffix_map[text[-1].upper()]
        text = text[:-1]
    text = text.strip()
    if text.endswith("%"):
        text = text[:-1]
    if not re.fullmatch(r"[-+]?\d*\.?\d+", text):
        return None
    number = float(text) * multiplier
    return -number if negative else number


from fiscal_tables import (
    extract_metric, extract_metric_history, extract_ratio_history,
    extract_estimate_mean_for_first_future_column, extract_overview_growth_metric,
    extract_ntm_valuation_metric, overview_value, current_price, set_millions,
)


def extract_first_matching_metric(page, labels: list[str], pick: str = "latest") -> str:
    for label in labels:
        value = extract_metric(page, label, pick=pick)
        if value != "N/A":
            return value
    return "N/A"


def goto_ready(page, url: str, wait_ms: int = 1500) -> None:
    print(f"[fiscal.ai] Open: {url}")
    try:
        page.goto(url, wait_until="load", timeout=120000)
    except Exception:
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=120000)
        except Exception:
            raise
    page.wait_for_timeout(wait_ms)


def scrape_overview_metrics(stock_id: str, headless: bool = True) -> dict[str, str]:
    try:
        from playwright.sync_api import sync_playwright
    except ImportError as exc:
        raise RuntimeError("playwright is not installed") from exc

    with sync_playwright() as p:
        if FISCAL_PROFILE_DIR.exists():
            context = p.chromium.launch_persistent_context(
                str(FISCAL_PROFILE_DIR),
                headless=headless,
                viewport={"width": 1800, "height": 2200},
            )
            page = prepare_persistent_page(context)
            browser = None
        else:
            browser = p.chromium.launch(headless=headless)
            context = browser.new_context(viewport={"width": 1800, "height": 2200})
            page = context.new_page()
        print(f"[fiscal.ai] Repair overview -> {stock_id}")
        goto_ready(page, f"https://fiscal.ai/company/{stock_id}/")
        result = {
            "eps_lt_growth_text": extract_overview_growth_metric(page, "EPS LT Growth Est"),
            "forward_pe_text": extract_ntm_valuation_metric(page, "P/E"),
        }
        context.close()
        if browser is not None:
            browser.close()
        return result


def repair_missing_overview_metrics(headless: bool = True) -> bool:
    if not FINCHAT_CSV_PATH.exists():
        return False

    rows: list[dict[str, str]] = []
    with FINCHAT_CSV_PATH.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        rows = list(reader)

    unsupported = set()
    changed = False
    for row in rows:
        stock_id = row.get("Stock", "").strip()
        if not stock_id or stock_id in unsupported:
            continue
        missing_growth = (row.get(METRIC_LABELS["eps_lt_growth"], "") or "").strip() in {"", "N/A", "—"}
        missing_pe = (row.get(METRIC_LABELS["forward_pe"], "") or "").strip() in {"", "N/A", "—"}
        if not (missing_growth or missing_pe):
            continue
        print(f"[fiscal.ai] Repair missing overview metrics for {stock_id}")
        overview = scrape_overview_metrics(stock_id, headless=headless)
        if missing_growth and overview["eps_lt_growth_text"] not in {"", "N/A", "—"}:
            row[METRIC_LABELS["eps_lt_growth"]] = overview["eps_lt_growth_text"]
            changed = True
        if missing_pe and overview["forward_pe_text"] not in {"", "N/A", "—"}:
            row[METRIC_LABELS["forward_pe"]] = overview["forward_pe_text"]
            changed = True

    if changed:
        save_finchat_cache(rows)
    return changed


def scrape_finchat(stock_ids: list[str], headless: bool = True) -> list[dict[str, str]]:
    try:
        from playwright.sync_api import sync_playwright
    except ImportError as exc:
        raise RuntimeError("playwright is not installed") from exc

    results_by_stock: dict[str, dict[str, str]] = {
        row.get("Stock", "").strip(): row
        for row in load_finchat_cache_rows()
        if row.get("Stock", "").strip()
    }
    with sync_playwright() as p:
        if FISCAL_PROFILE_DIR.exists():
            context = p.chromium.launch_persistent_context(
                str(FISCAL_PROFILE_DIR),
                headless=headless,
                viewport={"width": 1800, "height": 2200},
            )
            page = prepare_persistent_page(context)
            browser = None
            print(f"[fiscal.ai] Using saved browser profile: {FISCAL_PROFILE_DIR}")
        else:
            browser = p.chromium.launch(headless=headless)
            context = browser.new_context(viewport={"width": 1800, "height": 2200})
            page = context.new_page()
            print("[fiscal.ai] No saved browser profile found. Running without a persisted login session.")

        total = len(stock_ids)
        for index, stock_id in enumerate(stock_ids, start=1):
            print(f"[fiscal.ai] ({index}/{total}) Start {stock_id}")
            stock_data = FinchatMetrics(stock_id=stock_id)
            try:
                for tab, metrics in FINANCIAL_TABS.items():
                    print(f"[fiscal.ai] ({index}/{total}) {stock_id} -> financials/{tab}")
                    goto_ready(page, f"https://fiscal.ai/company/{stock_id}/financials/{tab}/annual/")
                    set_millions(page)
                    if tab == "income-statement":
                        stock_data.shares_text = extract_first_matching_metric(page, METRIC_ALIASES["shares"], pick="latest")
                    elif tab == "cash-flow-statement":
                        stock_data.fcf_history_text = extract_metric_history(page, "Free Cash Flow")
                        stock_data.stockbased_text = extract_first_matching_metric(page, METRIC_ALIASES["stockbased"], pick="latest")
                        stock_data.cap_text = extract_first_matching_metric(page, METRIC_ALIASES["cap"], pick="latest")
                        stock_data.dep_text = extract_first_matching_metric(page, METRIC_ALIASES["dep"], pick="latest")

                print(f"[fiscal.ai] ({index}/{total}) {stock_id} -> estimates/free-cash-flow")
                goto_ready(page, f"https://fiscal.ai/company/{stock_id}/estimates/free-cash-flow/", wait_ms=2500)
                stock_data.fcf_text = extract_estimate_mean_for_first_future_column(page, "Free Cash Flow")

                print(f"[fiscal.ai] ({index}/{total}) {stock_id} -> overview")
                goto_ready(page, f"https://fiscal.ai/company/{stock_id}/")
                stock_data.current_price_text = current_price(page, stock_id)
                overview_shares = overview_value(page, "Shares Out")
                if overview_shares != "N/A":
                    stock_data.shares_text = overview_shares
                stock_data.eps_lt_growth_text = extract_overview_growth_metric(page, "EPS LT Growth Est")
                stock_data.forward_pe_text = extract_ntm_valuation_metric(page, "P/E")

                print(f"[fiscal.ai] ({index}/{total}) {stock_id} -> ratios/annual (forward P/E history)")
                goto_ready(page, f"https://fiscal.ai/company/{stock_id}/financials/ratios/annual/", wait_ms=2500)
                stock_data.pe_history_text = extract_ratio_history(page, "Forward Valuation", "Forward P/E")

                results_by_stock[stock_id] = stock_data.to_row()
                save_finchat_cache([results_by_stock[sid] for sid in stock_ids if sid in results_by_stock])
                print(
                    f"[fiscal.ai] ({index}/{total}) Done {stock_id} | "
                    f"FCF={stock_data.fcf_text or 'N/A'} | "
                    f"EPS LT={stock_data.eps_lt_growth_text or 'N/A'} | "
                    f"Fwd PE={stock_data.forward_pe_text or 'N/A'}"
                )
            except Exception as exc:
                print(f"[fiscal.ai] ({index}/{total}) Skipped {stock_id}: {exc}")
                continue

        context.close()
        if browser is not None:
            browser.close()

    return [results_by_stock[sid] for sid in stock_ids if sid in results_by_stock]


def login_fiscal_ai() -> int:
    try:
        from playwright.sync_api import sync_playwright
    except ImportError as exc:
        raise RuntimeError("playwright is not installed") from exc

    GENERATED_DIR.mkdir(exist_ok=True)
    FISCAL_PROFILE_DIR.mkdir(exist_ok=True)
    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            str(FISCAL_PROFILE_DIR),
            headless=False,
            viewport={"width": 1600, "height": 1000},
        )
        page = prepare_persistent_page(context)
        page.goto("https://fiscal.ai/", wait_until="load", timeout=120000)
        print("")
        print("[fiscal.ai] Browser opened with a persistent profile.")
        print("[fiscal.ai] Log in to Fiscal.ai in this browser window.")
        print("[fiscal.ai] After you can open company pages normally, press Enter here to save the session and close the browser.")
        input()
        try:
            context.close()
        except Exception:
            # If the user already closed the browser window, the persistent
            # profile may still have been saved successfully.
            pass
    print(f"[fiscal.ai] Saved login profile to {FISCAL_PROFILE_DIR}")
    return 0


def prepare_persistent_page(context):
    page = None
    for candidate in context.pages:
        if candidate.url and candidate.url != "about:blank":
            page = candidate
            break
    if page is None:
        page = context.pages[0] if context.pages else context.new_page()
    for candidate in list(context.pages):
        if candidate is page:
            continue
        try:
            if candidate.url == "about:blank":
                candidate.close()
        except Exception:
            pass
    return page


def save_finchat_cache(rows: list[dict[str, str]]) -> None:
    GENERATED_DIR.mkdir(exist_ok=True)
    print(f"[fiscal.ai] Saving CSV cache to {FINCHAT_CSV_PATH}")
    with FINCHAT_CSV_PATH.open("w", encoding="utf-8-sig", newline="") as handle:
        fieldnames = [
            "Stock",
            "Current Price",
            METRIC_LABELS["shares"],
            METRIC_LABELS["fcf"],
            METRIC_LABELS["fcf_history"],
            METRIC_LABELS["stockbased"],
            METRIC_LABELS["cap"],
            METRIC_LABELS["dep"],
            METRIC_LABELS["eps_lt_growth"],
            METRIC_LABELS["forward_pe"],
            METRIC_LABELS["pe_history"],
        ]
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)

    print(f"[fiscal.ai] Saving JSON cache to {FINCHAT_JSON_PATH}")
    with FINCHAT_JSON_PATH.open("w", encoding="utf-8") as handle:
        json.dump(rows, handle, indent=2)


def load_finchat_cache_rows() -> list[dict[str, str]]:
    if not FINCHAT_CSV_PATH.exists():
        return []
    with FINCHAT_CSV_PATH.open("r", encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def load_finchat_cache() -> dict[str, dict[str, float | str | None]]:
    if not FINCHAT_CSV_PATH.exists():
        return {}

    cache: dict[str, dict[str, float | str | None]] = {}
    with FINCHAT_CSV_PATH.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        for row in reader:
            stock_id = row.get("Stock", "").strip()
            if not stock_id:
                continue
            cache[stock_id] = {
                "shares_text": row.get(METRIC_LABELS["shares"], ""),
                "shares_value": parse_metric_text(row.get(METRIC_LABELS["shares"], "")),
                "fcf_text": row.get(METRIC_LABELS["fcf"], ""),
                "fcf_value": parse_metric_text(row.get(METRIC_LABELS["fcf"], "")),
                "fcf_history_text": row.get(METRIC_LABELS["fcf_history"], ""),
                "fcf_history_values": json.loads(row.get(METRIC_LABELS["fcf_history"], "[]") or "[]"),
                "stockbased_text": row.get(METRIC_LABELS["stockbased"], ""),
                "stockbased_value": parse_metric_text(row.get(METRIC_LABELS["stockbased"], "")),
                "cap_text": row.get(METRIC_LABELS["cap"], ""),
                "cap_value": parse_metric_text(row.get(METRIC_LABELS["cap"], "")),
                "dep_text": row.get(METRIC_LABELS["dep"], ""),
                "dep_value": parse_metric_text(row.get(METRIC_LABELS["dep"], "")),
                "eps_lt_growth_text": row.get(METRIC_LABELS["eps_lt_growth"], ""),
                "eps_lt_growth_value": parse_metric_text(row.get(METRIC_LABELS["eps_lt_growth"], "")),
                "forward_pe_text": row.get(METRIC_LABELS["forward_pe"], ""),
                "forward_pe_value": parse_metric_text(row.get(METRIC_LABELS["forward_pe"], "")),
                "pe_history_text": row.get(METRIC_LABELS["pe_history"], ""),
                "pe_history_values": json.loads(row.get(METRIC_LABELS["pe_history"], "[]") or "[]"),
            }
    return cache


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--login":
        raise SystemExit(login_fiscal_ai())
    rows = scrape_finchat(list(STOCK_ID_BY_TICKER.values()), headless=False)
    save_finchat_cache(rows)
    print(f"Saved FinChat cache to {FINCHAT_CSV_PATH}")
