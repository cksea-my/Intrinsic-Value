# Intrinsic Value Calculator

Fills in a 20-year / 10-year discounted cash flow workbook automatically,
using free data from Yahoo Finance. You give it a ticker (and optionally
the fiscal year); it fetches free cash flow, debt, cash, growth rates,
shares outstanding, beta-based discount rate, and last close price, then
writes them into `Intrinsic_Value_Calculator.xlsx`. Every projection
formula in the workbook (PV of cash flows, intrinsic value, final
intrinsic value per share, discount/premium) is left untouched and
recalculates automatically in Excel.

Two ways to use it:

| | |
|---|---|
| **Desktop GUI** | `iv_calculator_gui.py` — a small Tkinter app you run locally on Windows/Mac/Linux |
| **Web portal**  | `web/` — a Flask app + one-page form you can deploy online |

Both share the same fetch logic in `iv_core.py`, so results are identical.
The web portal computes the full DCF in Python and **displays the results
directly on the page** (name, FCF, debt, cash, growth rates, discount rate,
10-year and 20-year intrinsic value, and the year-by-year FCF projection) —
no need to open Excel to see them. A "Download as Excel" button is also
available for anyone who wants the editable workbook.

## Repo layout

```
├── iv_core.py                     # shared fetch + Excel-write logic
├── iv_calculator_gui.py           # desktop GUI (tkinter)
├── Intrinsic_Value_Calculator.xlsx# template workbook (VMI 20y / 10y tabs)
├── requirements.txt
├── Procfile                       # for Render/Railway/Heroku-style hosts
└── web/
    ├── app.py                     # Flask backend (serves page + /api/generate)
    └── templates/
        └── index.html             # single-page form
```

## Rules the script applies (matches the template's own methodology)

- **Free Cash Flow (Current):** if Operating Cash Flow (last 4 quarters) >
  1.5 × Net Income (latest fiscal year) → use Free Cash Flow (last 4
  quarters); otherwise use Net Income (latest fiscal year).
- **Growth Yr 1–5:** analyst long-term growth estimate, falling back to
  historical EPS CAGR, then trailing earnings growth.
- **Growth Yr 6–10:** the lesser of 15% or half of the Yr 1–5 rate.
- **Growth Yr 11–20** *(20-year tab only)*: long-term GDP growth + 1% —
  4.18% for US stocks, 7% for China/Hong Kong stocks, defaulting to 4.18%
  (flagged) for anything else.
- **Discount rate:** Risk Free Rate + Beta × Market Risk Premium, using
  the same US vs. China/HK constants built into the template's reference
  tables, with beta clamped to the template's 0.8–1.6 band.
- **Last close price:** latest Yahoo Finance quote (overridable).

Every auto-filled cell gets an Excel comment noting where the number came
from, so you can sanity-check any run.

## Running the desktop GUI

```bash
pip install -r requirements.txt
python iv_calculator_gui.py
```

Enter a ticker (year defaults to the current year), click **Generate**,
pick where to save, done.

## Running the web portal locally

```bash
pip install -r requirements.txt
cd web
python app.py
```

Open `http://localhost:5000`. Enter a ticker (Year and Discount Rate are
optional — leave blank to auto-fill), click **Analyze**, and the page
shows the full valuation inline. Switch between the 10-year and 20-year
tabs with the toggle, and use **Download as Excel** if you want the
editable workbook.

## Deploying the web portal online

The frontend is one HTML page, but fetching live market data needs a
small Python backend (Yahoo Finance can't be called securely straight
from a browser, and GitHub Pages only hosts static files) — so the
simplest free path is: push this repo to GitHub, then point a small
Python host at it.

**1. Push to GitHub**
```bash
git init
git add .
git commit -m "Intrinsic Value Calculator"
git branch -M main
git remote add origin https://github.com/<your-username>/<repo-name>.git
git push -u origin main
```

**2. Deploy on [Render](https://render.com) (free tier works)**
1. New → Web Service → connect this GitHub repo.
2. Build command: `pip install -r requirements.txt`
3. Start command: `gunicorn --chdir web app:app`
4. Deploy. Render gives you a URL like `https://iv-calculator.onrender.com`
   — that's your one-page portal.

Render's free tier spins down when idle (first request after a while
takes ~30s to wake up) — fine for personal/light use. For an always-on
host, Railway.app or Fly.io work the same way (same build/start commands).

**3. (Optional) Custom domain**
Most of these hosts let you attach a custom domain for free — add it
under the service's Settings → Custom Domain once deployed.

## Notes / limitations

- Data is free, delayed Yahoo Finance data — not for time-sensitive
  trading decisions.
- Tickers outside the US (e.g. `0700.HK`, `VOD.L`) generally work; Yahoo's
  free data coverage for growth estimates and betas is thinner outside
  large-cap US/HK/China names, so check the Excel comments on those cells.
- This tool doesn't give investment advice — the intrinsic value model's
  assumptions (growth rates, discount rate) are estimates you should
  review, not guarantees.
