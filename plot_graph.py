#!/usr/bin/env python3
"""
plot_graph.py — Generate Graph 1 for the Bedrock Fund assignment.

Plots the Efficient Portfolio Frontier, MVP, Market Portfolio,
Zero-Covariance Portfolio, and Capital Market Line in (σ, μ) space.

EC310R Financial Economic Theory — Wilfrid Laurier University
Professor Doron Nisani | Huang & Litzenberger
"""

import matplotlib
matplotlib.use('Agg')

import os
import matplotlib.pyplot as plt
import numpy as np


def plot_frontier(frontier, mvp, market, zc, cml, rf, output_dir="output"):
    """
    Create the main (σ, μ) graph with all required elements.

    Parameters
    ----------
    frontier : dict
        'sigma' and 'mu' arrays from generate_frontier_points().
    mvp : dict
        MVP with 'expected_return' and 'std_dev'.
    market : dict
        Market Portfolio with 'expected_return' and 'std_dev'.
    zc : dict
        Zero-Covariance Portfolio with 'expected_return' and 'std_dev'.
    cml : dict
        CML 'sigma' and 'mu' arrays from generate_cml_points().
    rf : float
        Risk-free rate (daily).
    output_dir : str
        Directory to save the graph.
    """
    os.makedirs(output_dir, exist_ok=True)

    fig, ax = plt.subplots(figsize=(12, 8))

    # ── EPF: both branches ───────────────────────────────────────────────
    # Split into efficient (upper) and inefficient (lower) at the MVP
    mu_mvp = mvp["expected_return"]
    upper = frontier["mu"] >= mu_mvp - 1e-12
    lower = frontier["mu"] <= mu_mvp + 1e-12

    ax.plot(frontier["sigma"][upper], frontier["mu"][upper],
            color="navy", linewidth=2, label="Efficient Frontier (upper)")
    ax.plot(frontier["sigma"][lower], frontier["mu"][lower],
            color="navy", linewidth=1.2, linestyle="--", alpha=0.6,
            label="Inefficient Frontier (lower)")

    # ── CML ──────────────────────────────────────────────────────────────
    ax.plot(cml["sigma"], cml["mu"],
            color="firebrick", linewidth=1.8, linestyle="-.",
            label="Capital Market Line (CML)")

    # ── Risk-free rate point ─────────────────────────────────────────────
    ax.plot(0, rf, "D", color="firebrick", markersize=8, zorder=5,
            label=f"Risk-Free Rate (μ_F = {rf:.6f})")

    # ── MVP ──────────────────────────────────────────────────────────────
    ax.plot(mvp["std_dev"], mvp["expected_return"],
            "s", color="green", markersize=10, zorder=5,
            label=f"MVP (σ={mvp['std_dev']:.4f}, μ={mvp['expected_return']:.6f})")
    ax.annotate("MVP",
                xy=(mvp["std_dev"], mvp["expected_return"]),
                xytext=(mvp["std_dev"] + 0.001, mvp["expected_return"] + 0.0002),
                fontsize=10, fontweight="bold", color="green")

    # ── Market Portfolio ─────────────────────────────────────────────────
    ax.plot(market["std_dev"], market["expected_return"],
            "^", color="red", markersize=12, zorder=5,
            label=f"Market Portfolio (σ={market['std_dev']:.4f}, μ={market['expected_return']:.6f})")
    ax.annotate("Market\nPortfolio",
                xy=(market["std_dev"], market["expected_return"]),
                xytext=(market["std_dev"] + 0.001, market["expected_return"] + 0.0002),
                fontsize=10, fontweight="bold", color="red")

    # ── ZC Portfolio ─────────────────────────────────────────────────────
    ax.plot(zc["std_dev"], zc["expected_return"],
            "o", color="purple", markersize=10, zorder=5,
            label=f"ZC Portfolio (σ={zc['std_dev']:.4f}, μ={zc['expected_return']:.6f})")
    ax.annotate("ZC Portfolio",
                xy=(zc["std_dev"], zc["expected_return"]),
                xytext=(zc["std_dev"] + 0.001, zc["expected_return"] - 0.0004),
                fontsize=10, fontweight="bold", color="purple")

    # ── Formatting ───────────────────────────────────────────────────────
    ax.set_xlabel("σ (Standard Deviation)", fontsize=13)
    ax.set_ylabel("μ (Expected Return)", fontsize=13)
    ax.set_title("Bedrock Fund — Efficient Portfolio Frontier", fontsize=15, fontweight="bold")
    ax.legend(loc="upper left", fontsize=9, framealpha=0.9)
    ax.grid(True, alpha=0.3)
    ax.tick_params(labelsize=11)

    fig.tight_layout()

    output_path = os.path.join(output_dir, "graph1.png")
    fig.savefig(output_path, dpi=300, bbox_inches="tight")
    plt.close(fig)
    print(f"Graph saved to {output_path}")

    return output_path
