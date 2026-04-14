#!/usr/bin/env python3
"""
presentation_charts.py — Generate presentation charts for the Bedrock Fund.

Creates three charts:
1. Annualized returns by asset (horizontal bar chart, colored by cluster)
2. Correlation matrix heatmap
3. MVP weights (horizontal bar chart)

EC310R Financial Economic Theory — Wilfrid Laurier University
"""

import matplotlib
matplotlib.use('Agg')

import os
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.patches import Patch

from compute import (
    calculate_returns,
    calculate_mean_and_covariance,
    calculate_frontier_parameters,
    calculate_mvp,
)

# ── Cluster definitions ──────────────────────────────────────────────────────

CLUSTERS = {
    "AI/Compute": {
        "tickers": ["NVDA", "TSM", "EQIX", "VRT"],
        "color": "#4472C4",  # blue
    },
    "Energy Transition": {
        "tickers": ["LIN", "ALB", "NEE", "CAT"],
        "color": "#70AD47",  # green
    },
    "Logistics & Infra": {
        "tickers": ["UNP", "PLD", "DE", "WM", "COST"],
        "color": "#ED7D31",  # orange
    },
    "Healthcare": {
        "tickers": ["UNH", "TMO", "ISRG"],
        "color": "#BF4B4B",  # red
    },
    "Defense & Hard Assets": {
        "tickers": ["LMT", "FCX", "NEM", "BRK-B"],
        "color": "#7B5EA7",  # purple
    },
}


def get_ticker_color(ticker):
    """Return the cluster color for a given ticker."""
    for cluster_name, cluster_data in CLUSTERS.items():
        if ticker in cluster_data["tickers"]:
            return cluster_data["color"]
    return "#888888"  # fallback gray


def get_ticker_cluster(ticker):
    """Return the cluster name for a given ticker."""
    for cluster_name, cluster_data in CLUSTERS.items():
        if ticker in cluster_data["tickers"]:
            return cluster_name
    return "Other"


def chart_returns_by_asset(returns_df, tickers, output_dir="output"):
    """
    Chart 1: Horizontal bar chart of annualized returns for all 20 assets.
    Sorted from highest to lowest return, colored by cluster.
    """
    # Calculate annualized returns (252 trading days)
    mean_daily = returns_df.mean()
    ann_returns = mean_daily * 252

    # Create DataFrame for sorting
    data = pd.DataFrame({
        "ticker": tickers,
        "ann_return": [ann_returns[t] for t in tickers],
    })
    data = data.sort_values("ann_return", ascending=True)  # ascending for horizontal bar

    fig, ax = plt.subplots(figsize=(10, 8))

    # Create bars with cluster colors
    colors = [get_ticker_color(t) for t in data["ticker"]]
    bars = ax.barh(data["ticker"], data["ann_return"] * 100, color=colors, edgecolor="white")

    # Add percentage labels at end of each bar
    for bar, val in zip(bars, data["ann_return"] * 100):
        if val >= 0:
            ax.text(val + 1, bar.get_y() + bar.get_height() / 2,
                    f"{val:.1f}%", va="center", ha="left", fontsize=9)
        else:
            ax.text(val - 1, bar.get_y() + bar.get_height() / 2,
                    f"{val:.1f}%", va="center", ha="right", fontsize=9)

    # Vertical line at 0%
    ax.axvline(x=0, color="black", linewidth=0.8, linestyle="-")

    # Legend for clusters
    legend_elements = [
        Patch(facecolor=cluster["color"], label=name)
        for name, cluster in CLUSTERS.items()
    ]
    ax.legend(handles=legend_elements, loc="lower right", fontsize=9)

    ax.set_xlabel("Annualized Return (%)", fontsize=12)
    ax.set_title("2025 Annualized Returns by Asset", fontsize=14, fontweight="bold")
    ax.grid(axis="x", alpha=0.3)

    # Adjust x-axis limits to accommodate labels
    xmin, xmax = ax.get_xlim()
    ax.set_xlim(xmin - 5, xmax + 15)

    fig.tight_layout()
    output_path = os.path.join(output_dir, "chart_returns_by_asset.png")
    fig.savefig(output_path, dpi=300, bbox_inches="tight")
    plt.close(fig)
    print(f"Chart saved: {output_path}")
    return output_path


def chart_correlation_heatmap(returns_df, tickers, output_dir="output"):
    """
    Chart 2: Correlation matrix heatmap (20x20).
    Uses diverging colormap: blue (negative) -> white (zero) -> red (positive).
    """
    # Calculate correlation matrix (normalized to [-1, 1])
    corr_matrix = returns_df.corr()

    fig, ax = plt.subplots(figsize=(10, 9))

    # Use diverging colormap: blue-white-red
    cmap = plt.cm.RdBu_r  # reversed so red=positive, blue=negative
    im = ax.imshow(corr_matrix.values, cmap=cmap, vmin=-1, vmax=1, aspect="auto")

    # Add colorbar
    cbar = fig.colorbar(im, ax=ax, shrink=0.8)
    cbar.set_label("Correlation", fontsize=11)

    # Set tick labels
    ax.set_xticks(range(len(tickers)))
    ax.set_yticks(range(len(tickers)))
    ax.set_xticklabels(tickers, rotation=45, ha="right", fontsize=9)
    ax.set_yticklabels(tickers, fontsize=9)

    ax.set_title("Correlation Matrix \u2014 20 Bedrock Fund Assets", fontsize=14, fontweight="bold")

    fig.tight_layout()
    output_path = os.path.join(output_dir, "chart_correlation_heatmap.png")
    fig.savefig(output_path, dpi=300, bbox_inches="tight")
    plt.close(fig)
    print(f"Chart saved: {output_path}")
    return output_path


def chart_mvp_weights(mvp_weights, tickers, output_dir="output"):
    """
    Chart 3: Horizontal bar chart of MVP weights.
    All 20 assets, sorted by weight descending.
    """
    # Create DataFrame for sorting
    data = pd.DataFrame({
        "ticker": tickers,
        "weight": mvp_weights,
    })
    data = data.sort_values("weight", ascending=True)  # ascending for horizontal bar

    fig, ax = plt.subplots(figsize=(10, 6))

    bars = ax.barh(data["ticker"], data["weight"] * 100, color="#4472C4", edgecolor="white")

    # Add percentage labels
    for bar, val in zip(bars, data["weight"] * 100):
        if val >= 0:
            ax.text(val + 0.3, bar.get_y() + bar.get_height() / 2,
                    f"{val:.1f}%", va="center", ha="left", fontsize=9)
        else:
            ax.text(val - 0.3, bar.get_y() + bar.get_height() / 2,
                    f"{val:.1f}%", va="center", ha="right", fontsize=9)

    # Vertical line at 0%
    ax.axvline(x=0, color="black", linewidth=0.8, linestyle="-")

    ax.set_xlabel("Weight (%)", fontsize=12)
    ax.set_title("Minimum Variance Portfolio \u2014 Asset Weights", fontsize=14, fontweight="bold")
    ax.grid(axis="x", alpha=0.3)

    # Adjust x-axis limits to accommodate labels
    xmin, xmax = ax.get_xlim()
    ax.set_xlim(xmin - 3, xmax + 5)

    fig.tight_layout()
    output_path = os.path.join(output_dir, "chart_mvp_weights.png")
    fig.savefig(output_path, dpi=300, bbox_inches="tight")
    plt.close(fig)
    print(f"Chart saved: {output_path}")
    return output_path


def main():
    """Generate all presentation charts."""
    # Paths
    base_dir = os.path.dirname(os.path.abspath(__file__))
    data_dir = os.path.join(base_dir, "data")
    output_dir = os.path.join(base_dir, "output")
    os.makedirs(output_dir, exist_ok=True)

    # Load price data
    prices_path = os.path.join(data_dir, "risky_prices.csv")
    if not os.path.exists(prices_path):
        print(f"Error: {prices_path} not found. Run the analysis first.")
        return

    prices_df = pd.read_csv(prices_path, index_col=0, parse_dates=True)
    tickers = list(prices_df.columns)

    print(f"Loaded {len(prices_df)} days of price data for {len(tickers)} assets")

    # Calculate returns and statistics
    returns_df = calculate_returns(prices_df)
    mean_vector, cov_matrix = calculate_mean_and_covariance(returns_df)

    # Calculate MVP
    params = calculate_frontier_parameters(mean_vector, cov_matrix)
    mvp = calculate_mvp(params, cov_matrix)

    print(f"Computed statistics: mean, covariance, MVP")

    # Generate charts
    print("\nGenerating charts...")
    chart_returns_by_asset(returns_df, tickers, output_dir)
    chart_correlation_heatmap(returns_df, tickers, output_dir)
    chart_mvp_weights(mvp["weights"], tickers, output_dir)

    print("\nAll charts generated successfully!")


if __name__ == "__main__":
    main()
