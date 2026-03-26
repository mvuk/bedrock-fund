#!/usr/bin/env python3
"""
generate_outputs.py — Produce the final Excel deliverable for the Bedrock Fund.

Creates bedrock_fund.xlsx with 6 sheets matching the EC310R assignment spec.

EC310R Financial Economic Theory — Wilfrid Laurier University
Professor Doron Nisani | Huang & Litzenberger
"""

import os
import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XlImage

# Display names for the 20 risky assets
ASSET_NAMES = {
    "NVDA": "NVIDIA", "TSM": "TSMC", "EQIX": "Equinix", "VRT": "Vertiv",
    "LIN": "Linde", "ALB": "Albemarle", "NEE": "NextEra Energy",
    "CAT": "Caterpillar", "UNP": "Union Pacific", "PLD": "Prologis",
    "DE": "Deere", "WM": "Waste Management", "UNH": "UnitedHealth",
    "TMO": "Thermo Fisher", "ISRG": "Intuitive Surgical",
    "LMT": "Lockheed Martin", "FCX": "Freeport-McMoRan", "NEM": "Newmont",
    "COST": "Costco", "BRK-B": "Berkshire Hathaway",
}


def generate_excel(
    risky_prices,
    risky_returns,
    mean_vector,
    cov_matrix,
    frontier,
    bil_prices,
    bil_returns,
    mvp,
    market,
    zc,
    tickers,
    rf,
    params,
    graph_path=None,
    output_dir="output",
):
    """
    Generate the bedrock_fund.xlsx file with all 6 required sheets.

    Sheets:
        1. Daily Prices       — date + 20 asset price columns
        2. Daily Returns       — simple returns for all 20 assets
        3. Mean and Covariance — mean return row, then 20×20 cov matrix
        4. EPF                 — (σ, μ) frontier points
        5. Riskless Asset      — BIL prices and returns
        6. Portfolios          — MVP, Market, ZC weights and statistics
    """
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, "bedrock_fund.xlsx")

    labels = [f"{t} ({ASSET_NAMES.get(t, t)})" for t in tickers]

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:

        # ── Sheet 1: Daily Prices ────────────────────────────────────────
        prices_out = risky_prices.copy()
        prices_out.columns = labels
        prices_out.to_excel(writer, sheet_name="Daily Prices")

        # ── Sheet 2: Daily Returns ───────────────────────────────────────
        returns_out = risky_returns.copy()
        returns_out.columns = labels
        returns_out.to_excel(writer, sheet_name="Daily Returns")

        # ── Sheet 3: Mean and Covariance ─────────────────────────────────
        mean_df = pd.DataFrame([mean_vector], columns=labels, index=["Mean Return"])
        cov_df = pd.DataFrame(cov_matrix, index=labels, columns=labels)
        combined = pd.concat([mean_df, pd.DataFrame(index=[""]), cov_df])
        combined.to_excel(writer, sheet_name="Mean and Covariance")

        # ── Sheet 4: EPF ─────────────────────────────────────────────────
        epf_df = pd.DataFrame({
            "σ (Std Dev)": frontier["sigma"],
            "μ (Expected Return)": frontier["mu"],
        })
        epf_df.to_excel(writer, sheet_name="EPF", index=False)

        # ── Sheet 5: Riskless Asset ──────────────────────────────────────
        bil_combined = pd.DataFrame(index=bil_prices.index)
        bil_combined["BIL Price"] = bil_prices.values.flatten()
        if bil_returns is not None and len(bil_returns) > 0:
            bil_ret_series = pd.Series(bil_returns.values.flatten(), index=bil_returns.index)
            bil_combined["BIL Return"] = bil_ret_series
        bil_combined.to_excel(writer, sheet_name="Riskless Asset")

        # ── Sheet 6: Portfolios ──────────────────────────────────────────
        port_data = []
        for i, t in enumerate(tickers):
            port_data.append({
                "Ticker": t,
                "Name": ASSET_NAMES.get(t, t),
                "MVP Weight": mvp["weights"][i],
                "Market Weight": market["weights"][i],
                "ZC Weight": zc["weights"][i],
            })
        port_df = pd.DataFrame(port_data)

        # Summary statistics rows
        summary_rows = pd.DataFrame([
            {"Ticker": "", "Name": "", "MVP Weight": "", "Market Weight": "", "ZC Weight": ""},
            {"Ticker": "Sum of Weights", "Name": "",
             "MVP Weight": np.sum(mvp["weights"]),
             "Market Weight": np.sum(market["weights"]),
             "ZC Weight": np.sum(zc["weights"])},
            {"Ticker": "Expected Return (μ)", "Name": "",
             "MVP Weight": mvp["expected_return"],
             "Market Weight": market["expected_return"],
             "ZC Weight": zc["expected_return"]},
            {"Ticker": "Std Dev (σ)", "Name": "",
             "MVP Weight": mvp["std_dev"],
             "Market Weight": market["std_dev"],
             "ZC Weight": zc["std_dev"]},
            {"Ticker": "Sharpe Ratio", "Name": "",
             "MVP Weight": (mvp["expected_return"] - rf) / mvp["std_dev"],
             "Market Weight": market["sharpe_ratio"],
             "ZC Weight": (zc["expected_return"] - rf) / zc["std_dev"]},
        ])
        port_full = pd.concat([port_df, summary_rows], ignore_index=True)
        port_full.to_excel(writer, sheet_name="Portfolios", index=False)

    # Embed graph image if available
    if graph_path and os.path.exists(graph_path):
        wb = load_workbook(output_path)
        ws = wb["EPF"]
        img = XlImage(graph_path)
        img.width = 800
        img.height = 500
        ws.add_image(img, "D2")
        wb.save(output_path)

    print(f"Excel file saved to {output_path}")

    # ── Export each sheet as CSV ─────────────────────────────────────────
    csv_dir = os.path.join(output_dir, "csv")
    os.makedirs(csv_dir, exist_ok=True)

    prices_out = risky_prices.copy()
    prices_out.columns = labels
    prices_out.to_csv(os.path.join(csv_dir, "sheet1_daily_prices.csv"))

    returns_out = risky_returns.copy()
    returns_out.columns = labels
    returns_out.to_csv(os.path.join(csv_dir, "sheet2_daily_returns.csv"))

    mean_df = pd.DataFrame([mean_vector], columns=labels, index=["Mean Return"])
    cov_df = pd.DataFrame(cov_matrix, index=labels, columns=labels)
    combined = pd.concat([mean_df, pd.DataFrame(index=[""]), cov_df])
    combined.to_csv(os.path.join(csv_dir, "sheet3_mean_and_covariance.csv"))

    epf_df = pd.DataFrame({
        "σ (Std Dev)": frontier["sigma"],
        "μ (Expected Return)": frontier["mu"],
    })
    epf_df.to_csv(os.path.join(csv_dir, "sheet4_epf.csv"), index=False)

    bil_combined = pd.DataFrame(index=bil_prices.index)
    bil_combined["BIL Price"] = bil_prices.values.flatten()
    if bil_returns is not None and len(bil_returns) > 0:
        bil_ret_series = pd.Series(bil_returns.values.flatten(), index=bil_returns.index)
        bil_combined["BIL Return"] = bil_ret_series
    bil_combined.to_csv(os.path.join(csv_dir, "sheet5_riskless_asset.csv"))

    port_full.to_csv(os.path.join(csv_dir, "sheet6_portfolios.csv"), index=False)

    print(f"CSV exports saved to {csv_dir}/")

    return output_path
