"""
web/app.py — Flask backend for the Intrinsic Value Calculator online portal.

Serves a single-page form (templates/index.html) and a /api/generate
endpoint that fetches live data from Yahoo Finance, fills in the workbook,
and streams the result back as a downloadable .xlsx.

Run locally:
    pip install -r requirements.txt
    python app.py
    -> open http://localhost:5000

Deploy: see the top-level README.md ("Deploying the web portal").
"""

import os
import sys
import tempfile
import traceback

from flask import Flask, jsonify, render_template, request, send_file

# so `import iv_core` works whether this is run from /web or the repo root
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import iv_core  # noqa: E402

app = Flask(__name__)

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE_PATH = os.path.join(REPO_ROOT, "Intrinsic_Value_Calculator.xlsx")

MAX_TICKER_LEN = 12


@app.route("/")
def index():
    return render_template("index.html", default_year=iv_core.default_year())


@app.route("/api/generate", methods=["POST"])
def generate():
    payload = request.get_json(silent=True) or request.form

    ticker = (payload.get("ticker") or "").strip().upper()
    if not ticker or len(ticker) > MAX_TICKER_LEN:
        return jsonify({"error": "Please provide a valid ticker symbol."}), 400

    year_raw = payload.get("year")
    try:
        year = int(year_raw) if year_raw else iv_core.default_year()
    except (TypeError, ValueError):
        return jsonify({"error": "Year must be a whole number."}), 400

    def parse_optional_float(key):
        raw = payload.get(key)
        if raw in (None, ""):
            return None
        try:
            return float(raw)
        except ValueError:
            return None

    manual_last_close = parse_optional_float("last_close")
    manual_discount_rate = parse_optional_float("discount_rate")

    logs = []

    tmp_fd, tmp_path = tempfile.mkstemp(suffix=".xlsx")
    os.close(tmp_fd)

    try:
        iv_core.build_workbook(
            TEMPLATE_PATH,
            tmp_path,
            ticker,
            year,
            manual_last_close=manual_last_close,
            manual_discount_rate=manual_discount_rate,
            log=logs.append,
        )
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": f"Could not generate the workbook for '{ticker}': {e}", "log": logs}), 502

    download_name = f"{ticker}_Intrinsic_Value.xlsx"
    response = send_file(
        tmp_path,
        as_attachment=True,
        download_name=download_name,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    @response.call_on_close
    def _cleanup():
        try:
            os.remove(tmp_path)
        except OSError:
            pass

    return response


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
