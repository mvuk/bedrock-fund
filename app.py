#!/usr/bin/env python3
"""
app.py — Flask web interface for the Bedrock Fund Portfolio Analyzer.

Run: python app.py
Then open http://localhost:5000
"""

import os
import time
import zipfile
import traceback

import numpy as np
import pandas as pd
import yfinance as yf
from flask import Flask, render_template, request, redirect, url_for, send_from_directory

from compute import (
    calculate_returns,
    calculate_mean_and_covariance,
    calculate_frontier_parameters,
    generate_frontier_points,
    calculate_mvp,
    calculate_market_portfolio,
    calculate_zc_portfolio,
    generate_cml_points,
    verify_matrix_inversion,
)
from plot_graph import plot_frontier
from generate_outputs import generate_excel, ASSET_NAMES
from generate_word import generate_report

# ── App setup ────────────────────────────────────────────────────────────────

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
DATA_DIR = os.path.join(BASE_DIR, "data")

app = Flask(__name__)

SHEET_DEFS = [
    ("sheet1_daily_prices.csv", "Daily Prices", "prices", False),
    ("sheet2_daily_returns.csv", "Daily Returns", "returns", False),
    ("sheet3_mean_and_covariance.csv", "Mean & Covariance", "full", True),
    ("sheet4_epf.csv", "EPF", "returns", False),
    ("sheet5_riskless_asset.csv", "Riskless Asset", "returns", False),
    ("sheet6_portfolios.csv", "Portfolios", "full", True),
]

DEFAULT_TICKERS = "NVDA,TSM,EQIX,VRT,LIN,ALB,NEE,CAT,UNP,PLD,DE,WM,UNH,TMO,ISRG,LMT,FCX,NEM,COST,BRK-B"
DEFAULT_RISKLESS = "BIL"
DEFAULT_START = "2025-01-02"
DEFAULT_END = "2025-12-31"

# Store latest results in memory so /results can render them without re-computing
_latest_results = {}


# ── Helpers ──────────────────────────────────────────────────────────────────

def _download_prices(tickers, start, end):
    """Download adjusted close prices. Returns DataFrame or raises."""
    # yfinance end date is exclusive, add one day
    end_exclusive = pd.Timestamp(end) + pd.Timedelta(days=1)
    end_str = end_exclusive.strftime("%Y-%m-%d")

    df = yf.download(tickers, start=start, end=end_str, auto_adjust=True)

    if isinstance(df.columns, pd.MultiIndex):
        df = df["Close"]
    else:
        df = df[["Close"]]
        df.columns = tickers if isinstance(tickers, list) else [tickers]

    return df


def _parse_tickers(raw):
    """Parse comma/space/newline-separated tickers into a clean list."""
    raw = raw.replace("\n", ",").replace("\r", ",").replace(" ", ",")
    tickers = [t.strip().upper() for t in raw.split(",") if t.strip()]
    # Deduplicate while preserving order
    seen = set()
    unique = []
    for t in tickers:
        if t not in seen:
            seen.add(t)
            unique.append(t)
    return unique


def _load_sheet_data():
    """Load CSV sheet previews for the tabbed display."""
    csv_dir = os.path.join(OUTPUT_DIR, "csv")
    sheets = []
    for fname, label, fmt, show_full in SHEET_DEFS:
        path = os.path.join(csv_dir, fname)
        if not os.path.exists(path):
            sheets.append({"label": label, "headers": [], "rows": [], "total": 0, "truncated": False})
            continue

        import csv
        with open(path, newline="") as f:
            reader = csv.reader(f)
            all_rows = list(reader)

        if not all_rows:
            sheets.append({"label": label, "headers": [], "rows": [], "total": 0, "truncated": False})
            continue

        headers = all_rows[0]
        data_rows = all_rows[1:]
        total = len(data_rows)

        # Format numbers based on type
        def _fmt_cell(val, col_idx):
            if not val or val == "None":
                return ""
            try:
                fv = float(val)
                if fmt == "prices":
                    return f"{fv:.2f}"
                else:
                    # returns / covariance: 6 decimals, but keep larger numbers shorter
                    if abs(fv) >= 100:
                        return f"{fv:.2f}"
                    elif abs(fv) >= 1:
                        return f"{fv:.4f}"
                    else:
                        return f"{fv:.6f}"
            except (ValueError, TypeError):
                # Date or string — truncate dates to 10 chars
                if len(val) > 10 and "T" not in val and val[:4].isdigit():
                    return val[:10]
                return val

        formatted = []
        for row in data_rows:
            formatted.append([_fmt_cell(c, i) for i, c in enumerate(row)])

        if show_full or total <= 20:
            display_rows = formatted
            truncated = False
        else:
            display_rows = formatted[:10] + [None] + formatted[-5:]  # None = separator
            truncated = True

        sheets.append({
            "label": label,
            "headers": headers,
            "rows": display_rows,
            "total": total,
            "truncated": truncated,
        })

    return sheets


# ── Routes ───────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    error = request.args.get("error")
    return render_template(
        "index.html",
        default_tickers=DEFAULT_TICKERS,
        default_riskless=DEFAULT_RISKLESS,
        default_start=DEFAULT_START,
        default_end=DEFAULT_END,
        error=error,
    )


@app.route("/run", methods=["POST"])
def run_analysis():
    global _latest_results
    _latest_results = {}

    tickers_raw = request.form.get("tickers", DEFAULT_TICKERS)
    riskless = request.form.get("riskless", DEFAULT_RISKLESS).strip().upper()
    start_date = request.form.get("start_date", DEFAULT_START)
    end_date = request.form.get("end_date", DEFAULT_END)

    tickers = _parse_tickers(tickers_raw)

    # ── Validation ───────────────────────────────────────────────────────
    if len(tickers) < 2:
        return redirect(url_for("index", error="Need at least 2 risky asset tickers."))
    if len(tickers) > 30:
        return redirect(url_for("index", error="Maximum 30 tickers supported."))
    if not riskless:
        return redirect(url_for("index", error="Riskless asset ticker is required."))

    try:
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        os.makedirs(DATA_DIR, exist_ok=True)

        # ── 1. Download data ─────────────────────────────────────────────
        risky_prices = _download_prices(tickers, start_date, end_date)

        # Check columns match requested tickers
        missing = [t for t in tickers if t not in risky_prices.columns]
        if missing:
            return redirect(url_for("index",
                error=f"Failed to download data for: {', '.join(missing)}. "
                      "Check that these are valid US-listed ticker symbols."))

        risky_prices = risky_prices[tickers]

        # Check we got enough data
        if len(risky_prices) < 20:
            return redirect(url_for("index",
                error=f"Only {len(risky_prices)} trading days found. "
                      "Need at least 20. Check your date range."))

        # Check for missing data
        null_counts = risky_prices.isnull().sum()
        bad = null_counts[null_counts > 0]
        if len(bad) > 0:
            details = ", ".join(f"{t}: {c} missing" for t, c in bad.items())
            return redirect(url_for("index",
                error=f"Missing price data: {details}. Try a different date range."))

        bil_prices = _download_prices([riskless], start_date, end_date)

        if riskless not in bil_prices.columns:
            return redirect(url_for("index",
                error=f"Failed to download data for riskless asset '{riskless}'."))

        # Save to CSV for generate_outputs compatibility
        risky_prices.to_csv(os.path.join(DATA_DIR, "risky_prices.csv"))
        bil_prices.to_csv(os.path.join(DATA_DIR, "bil_prices.csv"))

        # ── 2. Compute ───────────────────────────────────────────────────
        risky_returns = calculate_returns(risky_prices)
        bil_returns = calculate_returns(bil_prices)

        mean_vector, cov_matrix = calculate_mean_and_covariance(risky_returns)
        rf = bil_returns.mean().values[0]

        params = calculate_frontier_parameters(mean_vector, cov_matrix)
        A, B, C, D = params["A"], params["B"], params["C"], params["D"]

        max_dev, cond_num = verify_matrix_inversion(cov_matrix, params["Sigma_inv"])

        frontier = generate_frontier_points(params)
        mvp = calculate_mvp(params, cov_matrix)

        # Check Case 1
        mu_mvp = A / C
        if rf >= mu_mvp:
            _latest_results = {"error":
                f"The risk-free rate ({rf:.8f}) is greater than or equal to A/C ({mu_mvp:.8f}). "
                "This means the tangency portfolio lies on the inefficient branch (Case 2/3 "
                "in Huang & Litzenberger). The standard Market Portfolio construction does not "
                "apply. Try a different set of assets or date range where risky assets "
                "outperform the risk-free rate on average."}
            return redirect(url_for("results"))

        market = calculate_market_portfolio(params, mean_vector, cov_matrix, rf)
        zc = calculate_zc_portfolio(params, market["expected_return"], mean_vector, cov_matrix)
        cml = generate_cml_points(rf, market)

        # ── 3. Generate outputs ──────────────────────────────────────────
        graph_path = plot_frontier(frontier, mvp, market, zc, cml, rf, output_dir=OUTPUT_DIR)

        generate_excel(
            risky_prices=risky_prices, risky_returns=risky_returns,
            mean_vector=mean_vector, cov_matrix=cov_matrix, frontier=frontier,
            bil_prices=bil_prices, bil_returns=bil_returns,
            mvp=mvp, market=market, zc=zc,
            tickers=tickers, rf=rf, params=params,
            graph_path=graph_path, output_dir=OUTPUT_DIR,
        )

        generate_report(
            risky_prices=risky_prices, risky_returns=risky_returns,
            mean_vector=mean_vector, cov_matrix=cov_matrix, params=params,
            frontier=frontier, bil_prices=bil_prices, bil_returns=bil_returns,
            mvp=mvp, market=market, zc=zc,
            tickers=tickers, rf=rf,
            graph_path=graph_path, output_dir=OUTPUT_DIR,
            inv_max_dev=max_dev, inv_cond_num=cond_num,
        )

        # Create CSV zip
        csv_dir = os.path.join(OUTPUT_DIR, "csv")
        zip_path = os.path.join(OUTPUT_DIR, "csv_all.zip")
        if os.path.isdir(csv_dir):
            with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
                for fname in sorted(os.listdir(csv_dir)):
                    if fname.endswith(".csv"):
                        zf.write(os.path.join(csv_dir, fname), fname)

        # ── 4. Build results context ─────────────────────────────────────
        n_assets = len(tickers)
        daily_std = np.sqrt(np.diag(cov_matrix))
        ann_mean = mean_vector * 252
        ann_std = daily_std * np.sqrt(252)
        sort_idx = np.argsort(ann_mean)[::-1]

        top5 = [(tickers[i], ann_mean[i], ann_std[i]) for i in sort_idx[:5]]
        bottom5 = [(tickers[i], ann_mean[i], ann_std[i]) for i in sort_idx[-5:]]

        mkt_sort = np.argsort(np.abs(market["weights"]))[::-1]
        market_weights = [
            (tickers[i], ASSET_NAMES.get(tickers[i], tickers[i]), market["weights"][i])
            for i in mkt_sort
        ]

        mvp_sort = np.argsort(np.abs(mvp["weights"]))[::-1]
        mvp_weights = [
            (tickers[i], ASSET_NAMES.get(tickers[i], tickers[i]), mvp["weights"][i])
            for i in mvp_sort
        ]

        gross_exposure = np.sum(np.abs(market["weights"]))
        long_sum = np.sum(market["weights"][market["weights"] > 0])
        short_sum = np.sum(market["weights"][market["weights"] < 0])
        mvp_sharpe = (mvp["expected_return"] - rf) / mvp["std_dev"]
        zc_sharpe = (zc["expected_return"] - rf) / zc["std_dev"]

        # Wrap dicts as SimpleNamespace so Jinja can use dot notation
        class _NS:
            def __init__(self, d):
                self.__dict__.update(d)

        sheets = _load_sheet_data()

        # JSON data for the client-side investment calculator
        import json
        calc_assets = []
        for i in mkt_sort:
            calc_assets.append({
                "ticker": tickers[i],
                "name": ASSET_NAMES.get(tickers[i], tickers[i]),
                "weight": float(market["weights"][i]),
            })
        calc_json = json.dumps({
            "assets": calc_assets,
            "mu_m_daily": float(market["expected_return"]),
            "sigma_m_daily": float(market["std_dev"]),
            "rf_daily": float(rf),
        })

        _latest_results = {
            "sheets": sheets,
            "A": A, "B": B, "C": C, "D": D, "H": market["H"], "rf": rf,
            "inv_max_dev": max_dev, "inv_cond_num": cond_num,
            "n_days": len(risky_prices), "n_returns": len(risky_returns),
            "mvp": _NS(mvp), "market": _NS(market), "zc": _NS(zc),
            "mvp_sharpe": mvp_sharpe, "zc_sharpe": zc_sharpe,
            "top5": top5, "bottom5": bottom5,
            "market_weights": market_weights, "mvp_weights": mvp_weights,
            "gross_exposure": gross_exposure,
            "long_sum": long_sum, "short_sum": short_sum,
            "cache_bust": int(time.time()),
            "calc_json": calc_json,
        }

    except AssertionError as e:
        _latest_results = {"error": f"Computation assertion failed: {e}"}
    except np.linalg.LinAlgError as e:
        _latest_results = {"error":
            f"Covariance matrix is singular or ill-conditioned: {e}. "
            "This typically means two assets have perfectly correlated returns. "
            "Try removing one of the correlated assets."}
    except Exception as e:
        tb = traceback.format_exc()
        _latest_results = {"error": f"Unexpected error: {e}\n\nTraceback:\n{tb}"}

    return redirect(url_for("results"))


@app.route("/results")
def results():
    if not _latest_results:
        return redirect(url_for("index", error="No analysis results available. Run an analysis first."))
    return render_template("results.html", **_latest_results)


@app.route("/download/<path:filename>")
def download(filename):
    # Security: only serve from output directory, no path traversal
    safe_name = os.path.basename(filename)
    if not os.path.exists(os.path.join(OUTPUT_DIR, safe_name)):
        return "File not found", 404
    return send_from_directory(OUTPUT_DIR, safe_name, as_attachment=(safe_name != "graph1.png"))


# ── Main ─────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    print("Starting Bedrock Fund Portfolio Analyzer...")
    print("Open http://localhost:8080 in your browser")
    app.run(host='0.0.0.0', port=8080, debug=False)
