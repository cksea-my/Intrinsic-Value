# Intrinsic Value Dashboard

A Flask web dashboard that automatically discovers Excel files in `data/` named `<TICKER>_Intrinsic_Value.xlsx`. Each ticker becomes a top navigation tab, and every worksheet becomes a sub-tab. The original workbook remains downloadable.

## Folder structure

```
intrinsic_value_dashboard/
├── app.py
├── requirements.txt
├── render.yaml
├── Procfile
├── data/
│   └── AAPL_Intrinsic_Value.xlsx
├── templates/
└── static/
```

## Add or update stocks

1. Place files in `data/`.
2. Use the exact format `TICKER_Intrinsic_Value.xlsx` (for example, `MSFT_Intrinsic_Value.xlsx`).
3. Commit and push to GitHub. Render redeploys automatically when Auto-Deploy is enabled.

> `.xlx` is not a valid modern Excel extension. Use `.xlsx` or `.xlsm`.

## Run locally

```bash
python -m venv .venv
# Windows: .venv\Scripts\activate
# macOS/Linux: source .venv/bin/activate
pip install -r requirements.txt
python app.py
```
Open http://127.0.0.1:5000

## Deploy to Render

1. Unzip this package.
2. Create a new GitHub repository and upload all contents (not the outer folder only).
3. In Render, choose **New > Blueprint** and connect the repository. Render reads `render.yaml`.
4. Alternatively choose **New > Web Service** with:
   - Build command: `pip install -r requirements.txt`
   - Start command: `gunicorn app:app --workers 2 --threads 4 --timeout 120`
5. Deploy.

## Important behavior

- Formula results are read from Excel's last saved cached values. Open/recalculate/save the workbook in Excel before uploading if formulas changed.
- This application is read-only; it never modifies workbooks.
- Large workbooks are rendered on demand. Free Render services may cold-start after inactivity.
- Do not commit confidential financial data to a public GitHub repository. Use a private repository if needed.
