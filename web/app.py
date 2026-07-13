"""
web/app.py — Flask backend for the Intrinsic Value Calculator online portal.

Two endpoints:
  POST /api/analyze   -> fetches live data + computes the DCF (10y & 20y),
                          returns JSON so the page can render the results
                          inline (no Excel needed to view them).
  POST /api/generate  -> same fetch, but returns a filled-in copy of the
                          .xlsx template as a download (for anyone who
                          wants the editable workbook).

Run locally:
    pip install -r requirements.txt
    python app.py
    -> open http://localhost:5000
"""

import os
import sys
import tempfile
import traceback

from flask import Flask, jsonify, render_template, request, send_file

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import iv_core  # noqa: E402

app = Flask(__name__)

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE_PATH = os.path.join(REPO_ROOT, "Intrinsic_Value_Calculator.xlsx")

MAX_TICKER_LEN = 12


def _parse_request_args(payload):
    ticker = (payload.get("ticker") or "").strip().upper()
    if not ticker or len(ticker) > MAX_TICKER_LEN:
        return None, None, None, ("Please provide a valid ticker symbol.", 400)

    year_raw = payload.get("year")
    year = None
    if year_raw not in (None, ""):
        try:
            year = int(year_raw)
        except (TypeError, ValueError):
            return None, None, None, ("Year must be a whole number.", 400)

    dr_raw = payload.get("discount_rate")
    discount_rate = None
    if dr_raw not in (None, ""):
        try:
            discount_rate = float(dr_raw)
        except (TypeError, ValueError):
            return None, None, None, ("Discount rate must be a number, e.g. 0.086.", 400)

    return ticker, year, discount_rate, None


@app.route("/")
def index():
    return render_template("index.html", default_year=iv_core.default_year())


@app.route("/api/analyze", methods=["POST"])
def analyze():
    payload = request.get_json(silent=True) or request.form
    ticker, year, discount_rate, error = _parse_request_args(payload)
    if error:
        msg, code = error
        return jsonify({"error": msg}), code

    try:
        result = iv_core.analyze_ticker(ticker, year=year, manual_discount_rate=discount_rate)
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": f"Could not fetch/compute data for '{ticker}': {e}"}), 502

    return jsonify(result)


@app.route("/api/generate", methods=["POST"])
def generate():
    payload = request.get_json(silent=True) or request.form
    ticker, year, discount_rate, error = _parse_request_args(payload)
    if error:
        msg, code = error
        return jsonify({"error": msg}), code

    logs = []
    tmp_fd, tmp_path = tempfile.mkstemp(suffix=".xlsx")
    os.close(tmp_fd)

    try:
        iv_core.build_workbook(
            TEMPLATE_PATH, tmp_path, ticker, year,
            manual_last_close=None,
            manual_discount_rate=discount_rate,
            log=logs.append,
        )
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": f"Could not generate the workbook for '{ticker}': {e}", "log": logs}), 502

    download_name = f"{ticker}_Intrinsic_Value.xlsx"
    response = send_file(
        tmp_path, as_attachment=True, download_name=download_name,
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
