# Reverse DCF

Small Windows-based workflow for generating reverse DCF Excel workbooks from the master inputs workbook, a macro-enabled template, Excel formulas, and optional Fiscal.ai scraped data.

## Requirements

- Windows
- Microsoft Excel installed
- Python 3.11+ recommended

## Setup

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install openpyxl pywin32 matplotlib playwright
python -m playwright install chromium
```

## Basic Commands

Generate the input workbook, per-company DCF files, and summary files:

```powershell
python script.py
```

Regenerate and refresh the Fiscal.ai cache first:

```powershell
python script.py --refresh-finchat
```

Watch the master input workbook and resync automatically after each save:

```powershell
python script.py --watch
```

Open a browser and save a persistent Fiscal.ai login session:

```powershell
python auto.py --login
```

Scrape Fiscal.ai data manually and write the local cache:

```powershell
python auto.py
```

## Typical Flow

1. Update `companies.txt`.
2. If needed, log in once with `python auto.py --login`.
3. Run `python script.py` or `python script.py --refresh-finchat`.
4. Open `generated\reverse_dcf_inputs.xlsx` in Excel, let formulas calculate, and save.
5. Rerun `python script.py`, or keep `python script.py --watch` running.

`companies.txt` is mainly useful for bootstrapping the first `generated\reverse_dcf_inputs.xlsx`. After that, the script can read the company roster directly from the inputs workbook.

## Generated Output

- `generated\reverse_dcf_inputs.xlsx`
- `generated\reverse_dcf_summary.xlsx`
- `generated\*.xlsm`
- `generated\csv\*.csv`
- `generated\charts\*.png`
- `generated\finchat_scraped_data.csv`
- `generated\finchat_scraped_data.json`
