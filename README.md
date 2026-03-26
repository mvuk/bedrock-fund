# Bedrock Fund — Portfolio Optimizer

Mean-variance portfolio optimization tool built for EC310R Financial Economic Theory (Wilfrid Laurier University).

## What it does
- Downloads daily price data from Yahoo Finance for any set of US-listed stocks
- Computes returns, covariance matrix, and the Efficient Portfolio Frontier
- Finds the Minimum Variance Portfolio, Market Portfolio (tangency), and Zero-Correlation Portfolio
- Generates the Capital Market Line and a combined graph
- Exports Excel (.xlsx), Word report (.docx), Markdown report, and CSVs
- Interactive Flask web UI with investment calculator demonstrating Two-Fund Separation

## Setup
```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Usage
### Web UI
```bash
python app.py
# Open http://localhost:5000
```

### Command line
```bash
python pull_data.py   # Download price data
python main.py        # Run full pipeline, outputs to output/
```

## Theory
Based on Huang & Litzenberger, "Foundations for Financial Economics" — implementing the N-asset mean-variance optimization (Chapters 3-5).
