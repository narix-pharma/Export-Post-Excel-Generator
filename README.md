# Narix Pharmaceuticals — Export Bill Generator

Internal tool for generating customs-ready Excel files from MARG invoice PDFs.

## What it does

1. Upload a MARG PDF invoice
2. Review and fill in generic names / manufacturers (auto-filled from product DB)
3. Build the packing list with per-customer allocation
4. Download two files:
   - **Invoice + Packing List Excel** (customs format)
   - **LP Excel** (label print format for India Post)

## Setup (Windows)

1. Install [Python 3.11+](https://www.python.org/downloads/)
2. Double-click **`START - Double Click Me.bat`**
3. The app opens in your browser automatically at `http://localhost:5000`

## Updating the product database

Place your updated `PRODUCT_DATA.xlsx` in the `narix-app` folder and double-click  
**`UPDATE PRODUCTS - Double Click Me.bat`**, then restart the app.

### PRODUCT_DATA.xlsx format

| Column | Header | Description |
|--------|--------|-------------|
| A | Brand Name | Product name as it appears in MARG |
| B | Generic Name | INN / generic name |
| C | Manufacturer | Manufacturer name |
| D | HSN Code | HSN code (default 3004) |

## Files

| File | Purpose |
|------|---------|
| `app.py` | Flask backend — PDF parsing, Excel generation |
| `static/index.html` | Frontend UI |
| `template.xlsx` | Invoice + Packing List Excel template |
| `lp_template.xlsx` | LP (Label Print) Excel template |
| `requirements.txt` | Python dependencies |

## Dependencies

```
flask
pdfplumber
openpyxl
python-dateutil
```
