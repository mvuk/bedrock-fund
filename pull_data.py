#!/usr/bin/env python3
"""
pull_data.py — Download historical price data for the Bedrock Fund.

Downloads daily adjusted close prices from Yahoo Finance for 20 risky assets
and 1 riskless asset (BIL) over the period 2025-03-24 to 2026-03-23.
Saves raw CSV files to the data/ directory.

EC310R Financial Economic Theory — Wilfrid Laurier University
Professor Doron Nisani | Huang & Litzenberger
"""

import os
import yfinance as yf
import pandas as pd

# ── Ticker definitions ──────────────────────────────────────────────────────

RISKY_TICKERS = [
    "NVDA", "TSM", "EQIX", "VRT", "LIN",
    "ALB", "NEE", "CAT", "UNP", "PLD",
    "DE", "WM", "UNH", "TMO", "ISRG",
    "LMT", "FCX", "NEM", "COST", "BRK-B",
]

RISKLESS_TICKER = "BIL"

START_DATE = "2025-01-02"
END_DATE = "2026-01-01"  # yfinance end is exclusive, so this gives up to 2025-12-31

DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")


def download_prices(tickers, start, end):
    """Download adjusted close prices for a list of tickers."""
    df = yf.download(tickers, start=start, end=end, auto_adjust=True)

    # yf.download returns MultiIndex columns when multiple tickers are passed.
    # For a single ticker it returns flat columns.
    if isinstance(df.columns, pd.MultiIndex):
        df = df["Close"]
    else:
        df = df[["Close"]]
        df.columns = tickers if isinstance(tickers, list) else [tickers]

    return df


def main():
    os.makedirs(DATA_DIR, exist_ok=True)

    # ── Download risky asset prices ──────────────────────────────────────
    print(f"Downloading risky asset prices for {len(RISKY_TICKERS)} tickers...")
    risky_prices = download_prices(RISKY_TICKERS, START_DATE, END_DATE)

    # Ensure column order matches our ticker list
    risky_prices = risky_prices[RISKY_TICKERS]

    risky_path = os.path.join(DATA_DIR, "risky_prices.csv")
    risky_prices.to_csv(risky_path)
    print(f"  Saved risky prices to {risky_path}")
    print(f"  Shape: {risky_prices.shape[0]} trading days x {risky_prices.shape[1]} assets")

    # ── Download riskless asset (BIL) prices ─────────────────────────────
    print(f"\nDownloading riskless asset ({RISKLESS_TICKER}) prices...")
    bil_prices = download_prices([RISKLESS_TICKER], START_DATE, END_DATE)

    bil_path = os.path.join(DATA_DIR, "bil_prices.csv")
    bil_prices.to_csv(bil_path)
    print(f"  Saved BIL prices to {bil_path}")
    print(f"  Shape: {bil_prices.shape[0]} trading days x {bil_prices.shape[1]} column")

    # ── Diagnostics ──────────────────────────────────────────────────────
    print(f"\n{'='*60}")
    print(f"Trading days collected (risky): {risky_prices.shape[0]}")
    print(f"Trading days collected (BIL):   {bil_prices.shape[0]}")

    if risky_prices.shape[0] != bil_prices.shape[0]:
        print("WARNING: Row count mismatch between risky assets and BIL!")

    # Check for missing data per ticker
    missing = risky_prices.isnull().sum()
    tickers_with_missing = missing[missing > 0]
    if len(tickers_with_missing) > 0:
        print("\nWARNING: The following tickers have missing data:")
        for ticker, count in tickers_with_missing.items():
            print(f"  {ticker}: {count} missing values")
    else:
        print("All risky tickers have complete data.")

    bil_missing = bil_prices.isnull().sum().sum()
    if bil_missing > 0:
        print(f"WARNING: BIL has {bil_missing} missing values!")
    else:
        print("BIL has complete data.")

    print(f"{'='*60}")


if __name__ == "__main__":
    main()
