#!/usr/bin/env python3
"""
main.py — Orchestrator for the Bedrock Fund project.

Loads price data, computes all portfolio theory quantities, generates the
graph and the final Excel deliverable.

EC310R Financial Economic Theory — Wilfrid Laurier University
Professor Doron Nisani | Huang & Litzenberger

Usage:
    python main.py
"""

import os
import numpy as np
import pandas as pd

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
from generate_outputs import generate_excel
from generate_word import generate_report

# ── Paths ────────────────────────────────────────────────────────────────────

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

RISKY_TICKERS = [
    "NVDA", "TSM", "EQIX", "VRT", "LIN",
    "ALB", "NEE", "CAT", "UNP", "PLD",
    "DE", "WM", "UNH", "TMO", "ISRG",
    "LMT", "FCX", "NEM", "COST", "BRK-B",
]

ASSET_NAMES = {
    "NVDA": "NVIDIA", "TSM": "TSMC", "EQIX": "Equinix", "VRT": "Vertiv",
    "LIN": "Linde", "ALB": "Albemarle", "NEE": "NextEra Energy",
    "CAT": "Caterpillar", "UNP": "Union Pacific", "PLD": "Prologis",
    "DE": "Deere", "WM": "Waste Management", "UNH": "UnitedHealth",
    "TMO": "Thermo Fisher", "ISRG": "Intuitive Surgical",
    "LMT": "Lockheed Martin", "FCX": "Freeport-McMoRan", "NEM": "Newmont",
    "COST": "Costco", "BRK-B": "Berkshire Hathaway",
}


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # ── 1. Load data ─────────────────────────────────────────────────────
    print("Loading price data...")
    risky_prices = pd.read_csv(
        os.path.join(DATA_DIR, "risky_prices.csv"), index_col=0, parse_dates=True
    )
    bil_prices = pd.read_csv(
        os.path.join(DATA_DIR, "bil_prices.csv"), index_col=0, parse_dates=True
    )
    print(f"  Risky assets: {risky_prices.shape[0]} days x {risky_prices.shape[1]} assets")
    print(f"  BIL:          {bil_prices.shape[0]} days")

    # ── 2. Compute returns ───────────────────────────────────────────────
    print("\nComputing returns...")
    risky_returns = calculate_returns(risky_prices)
    bil_returns = calculate_returns(bil_prices)
    print(f"  Return observations: {len(risky_returns)}")

    # ── 3. Mean and covariance ───────────────────────────────────────────
    print("\nComputing mean vector and covariance matrix...")
    mean_vector, cov_matrix = calculate_mean_and_covariance(risky_returns)

    # Risk-free rate: average daily return of BIL
    rf = bil_returns.mean().values[0]
    print(f"  Risk-free rate (daily avg of BIL): {rf:.8f}")

    # ── 4. Frontier parameters ───────────────────────────────────────────
    print("\nComputing frontier parameters (A, B, C, D)...")
    params = calculate_frontier_parameters(mean_vector, cov_matrix)
    A, B, C, D = params["A"], params["B"], params["C"], params["D"]

    # ── 4b. Verify matrix inversion ────────────────────────────────────
    max_dev, cond_num = verify_matrix_inversion(cov_matrix, params["Sigma_inv"])
    print(f"  Matrix inversion check:")
    print(f"    Max |Σ·Σ⁻¹ - I| = {max_dev:.2e}")
    print(f"    Condition number  = {cond_num:.2f}")
    if cond_num < 100:
        cond_desc = "well-conditioned"
    elif cond_num < 10000:
        cond_desc = "moderately conditioned"
    else:
        cond_desc = "ill-conditioned"
    print(f"    Assessment: {cond_desc}")

    # ── 5. Frontier points ───────────────────────────────────────────────
    print("Generating frontier points...")
    frontier = generate_frontier_points(params)

    # ── 6. MVP ───────────────────────────────────────────────────────────
    print("Computing Minimum Variance Portfolio...")
    mvp = calculate_mvp(params, cov_matrix)

    # ── 7. Market Portfolio ──────────────────────────────────────────────
    print("Computing Market (Tangency) Portfolio...")
    market = calculate_market_portfolio(params, mean_vector, cov_matrix, rf)

    # ── 8. Zero-Covariance Portfolio ─────────────────────────────────────
    print("Computing Zero-Covariance Portfolio for Market Portfolio...")
    zc = calculate_zc_portfolio(params, market["expected_return"], mean_vector, cov_matrix)

    # Verify zero covariance
    cov_mkt_zc = market["weights"] @ cov_matrix @ zc["weights"]
    print(f"  Cov(Market, ZC) = {cov_mkt_zc:.2e} (should be ≈ 0)")

    # ── 9. CML ───────────────────────────────────────────────────────────
    print("Generating Capital Market Line points...")
    cml = generate_cml_points(rf, market)

    # ── 10. Plot ─────────────────────────────────────────────────────────
    print("\nGenerating graph...")
    graph_path = plot_frontier(frontier, mvp, market, zc, cml, rf, output_dir=OUTPUT_DIR)

    # ── 11. Excel ────────────────────────────────────────────────────────
    print("Generating Excel file...")
    generate_excel(
        risky_prices=risky_prices,
        risky_returns=risky_returns,
        mean_vector=mean_vector,
        cov_matrix=cov_matrix,
        frontier=frontier,
        bil_prices=bil_prices,
        bil_returns=bil_returns,
        mvp=mvp,
        market=market,
        zc=zc,
        tickers=RISKY_TICKERS,
        rf=rf,
        params=params,
        graph_path=graph_path,
        output_dir=OUTPUT_DIR,
    )

    # ── 12. Word report ─────────────────────────────────────────────────
    print("Generating Word report...")
    report_path = generate_report(
        risky_prices=risky_prices,
        risky_returns=risky_returns,
        mean_vector=mean_vector,
        cov_matrix=cov_matrix,
        params=params,
        frontier=frontier,
        bil_prices=bil_prices,
        bil_returns=bil_returns,
        mvp=mvp,
        market=market,
        zc=zc,
        tickers=RISKY_TICKERS,
        rf=rf,
        graph_path=graph_path,
        output_dir=OUTPUT_DIR,
        inv_max_dev=max_dev,
        inv_cond_num=cond_num,
    )

    # ── 13. Summary ──────────────────────────────────────────────────────
    print("\n" + "=" * 70)
    print("BEDROCK FUND — RESULTS SUMMARY")
    print("=" * 70)

    print(f"\nFrontier Parameters:")
    print(f"  A = {A:.10f}")
    print(f"  B = {B:.10f}")
    print(f"  C = {C:.10f}")
    print(f"  D = {D:.10f}")

    print(f"\nRisk-Free Rate (μ_F, daily): {rf:.8f}")
    print(f"  H = {market['H']:.10f}")

    print(f"\nMinimum Variance Portfolio (MVP):")
    print(f"  μ_MVP = {mvp['expected_return']:.8f}")
    print(f"  σ_MVP = {mvp['std_dev']:.8f}")
    print(f"  Weights:")
    for i, t in enumerate(RISKY_TICKERS):
        print(f"    {t:6s} ({ASSET_NAMES.get(t, t):22s}): {mvp['weights'][i]:+.6f}")

    print(f"\nMarket (Tangency) Portfolio:")
    print(f"  μ_M   = {market['expected_return']:.8f}")
    print(f"  σ_M   = {market['std_dev']:.8f}")
    print(f"  Sharpe = {market['sharpe_ratio']:.6f}")
    print(f"  Weights:")
    for i, t in enumerate(RISKY_TICKERS):
        print(f"    {t:6s} ({ASSET_NAMES.get(t, t):22s}): {market['weights'][i]:+.6f}")

    print(f"\nZero-Covariance Portfolio (for Market Portfolio):")
    print(f"  μ_ZC  = {zc['expected_return']:.8f}")
    print(f"  σ_ZC  = {zc['std_dev']:.8f}")
    print(f"  Cov(M, ZC) = {cov_mkt_zc:.2e}")
    print(f"  Weights:")
    for i, t in enumerate(RISKY_TICKERS):
        print(f"    {t:6s} ({ASSET_NAMES.get(t, t):22s}): {zc['weights'][i]:+.6f}")

    # ── 13. Diagnostics ──────────────────────────────────────────────────
    print(f"\n{'─' * 70}")
    print("DIAGNOSTIC: Market Portfolio — Top 5 weights by |w|")
    print(f"{'─' * 70}")
    abs_weights = np.abs(market["weights"])
    top5_idx = np.argsort(abs_weights)[::-1][:5]
    for rank, idx in enumerate(top5_idx, 1):
        t = RISKY_TICKERS[idx]
        w = market["weights"][idx]
        print(f"  {rank}. {t:6s} ({ASSET_NAMES.get(t, t):22s}): {w:+.4f}  (|w| = {abs_weights[idx]:.4f})")
    print(f"  Sum of all |w|: {np.sum(abs_weights):.4f}  (leverage ratio)")

    print(f"\n{'─' * 70}")
    print("DIAGNOSTIC: Annualized asset statistics (×252 for mean, ×√252 for σ)")
    print(f"{'─' * 70}")
    daily_std = np.sqrt(np.diag(cov_matrix))
    ann_mean = mean_vector * 252
    ann_std = daily_std * np.sqrt(252)
    # Sort by annualized mean descending
    sort_idx = np.argsort(ann_mean)[::-1]
    print(f"  {'Ticker':<8s} {'Name':<24s} {'Ann. μ':>8s} {'Ann. σ':>8s} {'Daily μ':>10s}")
    for idx in sort_idx:
        t = RISKY_TICKERS[idx]
        print(f"  {t:<8s} {ASSET_NAMES.get(t, t):<24s} {ann_mean[idx]:>+7.2%} {ann_std[idx]:>7.2%} {mean_vector[idx]:>+.6f}")

    # ── 14. Cluster analysis ────────────────────────────────────────────
    CLUSTERS = {
        "AI/Compute Infrastructure": ["NVDA", "TSM", "EQIX", "VRT"],
        "Energy Transition Enablers": ["LIN", "ALB", "NEE", "CAT"],
        "Global Logistics & Physical Infra": ["UNP", "PLD", "DE", "WM"],
        "Healthcare & Life Sciences Infra": ["UNH", "TMO", "ISRG"],
        "Defense & Hard Asset Enablers": ["LMT", "FCX", "NEM", "COST", "BRK-B"],
    }
    ticker_idx = {t: i for i, t in enumerate(RISKY_TICKERS)}

    for port_name, weights in [("Market Portfolio", market["weights"]),
                                ("Minimum Variance Portfolio", mvp["weights"])]:
        print(f"\n{'━' * 70}")
        print(f"  {port_name} — Weights by Cluster")
        print(f"{'━' * 70}")
        print(f"  {'Ticker':<8s} {'Name':<24s} {'Weight':>10s}")
        print(f"  {'─' * 44}")
        grand_total = 0.0
        for cluster_name, cluster_tickers in CLUSTERS.items():
            cluster_total = 0.0
            print(f"\n  {cluster_name}")
            for t in cluster_tickers:
                i = ticker_idx[t]
                w = weights[i]
                cluster_total += w
                print(f"    {t:<8s} {ASSET_NAMES.get(t, t):<24s} {w:>+10.4f}")
            print(f"    {'':─<32s} ──────────")
            print(f"    {'Cluster Total':<32s} {cluster_total:>+10.4f}")
            grand_total += cluster_total
        print(f"\n  {'─' * 44}")
        print(f"  {'GRAND TOTAL':<34s} {grand_total:>+10.4f}")

    # ── 15. Cross-cluster correlation matrix ─────────────────────────────
    print(f"\n{'━' * 70}")
    print("  Cross-Cluster Daily Return Correlations")
    print(f"{'━' * 70}")
    # Compute equal-weighted daily cluster returns
    cluster_returns = {}
    for cluster_name, cluster_tickers in CLUSTERS.items():
        cols = [RISKY_TICKERS.index(t) for t in cluster_tickers]
        cluster_returns[cluster_name] = risky_returns.iloc[:, cols].mean(axis=1)
    cluster_ret_df = pd.DataFrame(cluster_returns)
    corr = cluster_ret_df.corr()
    # Short labels for printing
    short = ["AI/Comp", "Energy", "Logist", "Health", "Defens"]
    print(f"  {'':>10s}", end="")
    for s in short:
        print(f" {s:>8s}", end="")
    print()
    for i, (cname, _) in enumerate(CLUSTERS.items()):
        print(f"  {short[i]:>10s}", end="")
        for j in range(len(CLUSTERS)):
            print(f" {corr.iloc[i, j]:>+8.3f}", end="")
        print()

    # ── 16. Interpretation ───────────────────────────────────────────────
    print(f"\n{'━' * 70}")
    print("  Interpretation")
    print(f"{'━' * 70}")
    # Compute annualized cluster stats
    print(f"\n  {'Cluster':<38s} {'Ann. μ':>8s}  {'Mkt w':>8s}  {'MVP w':>8s}")
    print(f"  {'─' * 66}")
    for cluster_name, cluster_tickers in CLUSTERS.items():
        cols = [ticker_idx[t] for t in cluster_tickers]
        cl_ann_mean = np.mean(ann_mean[cols])
        cl_mkt_w = np.sum(market["weights"][cols])
        cl_mvp_w = np.sum(mvp["weights"][cols])
        print(f"  {cluster_name:<38s} {cl_ann_mean:>+7.2%}  {cl_mkt_w:>+8.4f}  {cl_mvp_w:>+8.4f}")

    print(f"\n{'=' * 70}")
    print("Deliverables:")
    print(f"  Graph:  {os.path.abspath(graph_path)}")
    print(f"  Excel:  {os.path.abspath(os.path.join(OUTPUT_DIR, 'bedrock_fund.xlsx'))}")
    print(f"  Report: {os.path.abspath(report_path)}")
    print(f"{'=' * 70}")


if __name__ == "__main__":
    main()
