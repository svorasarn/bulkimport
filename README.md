# Konsole Bulk Import Validator

Web-based validator for Konsole bulk import Excel files. Checks your file against Konsole's exact parsing rules **before** you upload it, giving clear error messages instead of cryptic import failures.

## What it checks

- **Column names** — exact case match (Konsole's `BulkImportService.mapHeader()` is case-sensitive)
- **Required fields** — flags empty cells on `@NotEmpty` columns
- **Value patterns** — validates against `@Pattern` regex from Java annotations
- **Instrument Type** — case-sensitive `mapFormOfUndertaking()` switch (Demand guarantee / Dependent undertaking / Standby letter of credit)
- **Date format** — must be dd/MM/yyyy
- **Amount format** — must be numeric
- **Header row position** — Konsole reads from row 2
- **Minimum column count** — needs >= 20 or rows are silently skipped

## Supported message types

- **GTI** — BG Issued / SB Import (24 columns)
- **LCE** — LC Export (19 columns)
- **LCI** — LC Import (21 columns)
- **GTR** — BG Received / SB Export (26 columns)

## Supported file formats

- `.xlsx` (Excel 2007+)
- `.xls` (Excel 97-2003, including encrypted files)
- `.csv`

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy to Streamlit Cloud

1. Push this repo to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Connect your GitHub repo
4. Set main file path to `app.py`
5. Deploy
