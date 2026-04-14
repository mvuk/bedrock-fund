#!/usr/bin/env python3
"""
compute.py — Core computation module for the Bedrock Fund.

Implements mean-variance portfolio theory following Huang & Litzenberger
and the EC310R course material (Professor Doron Nisani, Wilfrid Laurier University).

All matrix operations use NumPy. The @ operator is used for matrix multiplication.
Notation mirrors Huang & Litzenberger exactly.
"""

import numpy as np
import pandas as pd


def calculate_returns(prices_df):
    """
    Compute simple (arithmetic) daily returns from a price DataFrame.

    Course reference: Chapter 6 — Statistical Foundations (Weeks 1-2)

    Formula:
        R_t = (P_t - P_{t-1}) / P_{t-1}

    We use simple returns (not log returns) because the portfolio return is
    the weighted sum of individual simple returns:
        R_P = w_1 R_1 + w_2 R_2 + ... + w_n R_n

    This additivity property does NOT hold for log returns, making simple
    returns the correct choice for mean-variance portfolio optimization.

    Parameters
    ----------
    prices_df : pd.DataFrame
        DataFrame of asset prices with dates as index, tickers as columns.

    Returns
    -------
    pd.DataFrame
        DataFrame of simple returns. First row is dropped (NaN from differencing).
    """
    returns = prices_df.pct_change().dropna()
    return returns


def calculate_mean_and_covariance(returns_df):
    """
    Compute the sample mean vector and sample covariance matrix of returns.

    Course reference: Chapters 6-8 (Weeks 1-3)

    The mean vector μ = (μ_1, μ_2, ..., μ_n)^T contains the average daily
    return for each of the n assets.

    The covariance matrix Σ is the n×n sample covariance matrix, computed
    with (n-1) denominator (Bessel's correction) for an unbiased estimate.

    Key properties of Σ:
        - Square (n×n)
        - Symmetric: Σ = Σ^T
        - Positive semi-definite: x^T Σ x ≥ 0 for all x
        - Invertible (required for frontier calculations; verified downstream)

    Parameters
    ----------
    returns_df : pd.DataFrame
        DataFrame of daily returns (T observations × n assets).

    Returns
    -------
    mean_vector : np.ndarray
        (n,) array of mean daily returns.
    cov_matrix : np.ndarray
        (n, n) sample covariance matrix (ddof=1).
    """
    mean_vector = returns_df.mean().values
    cov_matrix = returns_df.cov().values  # pandas .cov() uses ddof=1 by default
    return mean_vector, cov_matrix


def calculate_frontier_parameters(mean_vector, cov_matrix):
    """
    Compute the four scalar parameters (A, B, C, D) that fully characterize
    the efficient portfolio frontier for n risky assets.

    Course reference: Chapter 8 — N-Asset Portfolio Selection (Weeks 2-3)
    Huang & Litzenberger, Chapters 3-4

    Derivation:
        The investor minimizes portfolio variance σ²_P = w^T Σ w subject to
        w^T μ = μ_P (target return) and w^T 1 = 1 (full investment).

        The Lagrangian yields the optimal weight vector:
            w* = Σ⁻¹ [λ₁ μ + λ₂ 1]

        where the multipliers depend on four scalar constants derived from Σ⁻¹:

        Let 1 be the n×1 vector of ones, μ the n×1 mean vector.

        A = 1^T Σ⁻¹ μ  = μ^T Σ⁻¹ 1    (scalar; symmetry of Σ⁻¹)
        B = μ^T Σ⁻¹ μ                    (scalar; must be > 0)
        C = 1^T Σ⁻¹ 1                    (scalar; must be > 0)
        D = B·C - A²                      (scalar; must be > 0)

    The parameter D > 0 ensures the frontier is a proper hyperbola (not
    degenerate). If D = 0, all assets have the same expected return.

    Parameters
    ----------
    mean_vector : np.ndarray
        (n,) mean return vector μ.
    cov_matrix : np.ndarray
        (n, n) covariance matrix Σ.

    Returns
    -------
    dict
        Keys: 'A', 'B', 'C', 'D', 'Sigma_inv'
    """
    n = len(mean_vector)
    ones = np.ones(n)
    mu = mean_vector

    # Invert the covariance matrix
    Sigma_inv = np.linalg.inv(cov_matrix)

    # Compute the four frontier constants
    A = ones @ Sigma_inv @ mu          # A = 1'Σ⁻¹μ
    B = mu @ Sigma_inv @ mu            # B = μ'Σ⁻¹μ
    C = ones @ Sigma_inv @ ones        # C = 1'Σ⁻¹1
    D = B * C - A ** 2                 # D = BC - A²

    # Assertion checks (Huang & Litzenberger requirements)
    assert B > 0, f"B must be > 0, got B = {B}"
    assert C > 0, f"C must be > 0, got C = {C}"
    assert D > 0, f"D must be > 0, got D = {D}"

    return {"A": A, "B": B, "C": C, "D": D, "Sigma_inv": Sigma_inv}


def generate_frontier_points(params, n_points=500):
    """
    Generate (σ, μ) points along the entire efficient portfolio frontier,
    including both the upper (efficient) and lower (inefficient) branches.

    Course reference: Chapter 8 — The Efficient Portfolio Frontier

    The frontier in (σ², μ) space is given by:
        σ²_P(μ_P) = (1/D)(C·μ_P² - 2A·μ_P + B)

    Equivalently, in hyperbola form:
        σ²_P = (C/D)(μ_P - A/C)² + 1/C

    The vertex of the hyperbola is the Minimum Variance Portfolio (MVP)
    at (σ_MVP, μ_MVP) = (√(1/C), A/C).

    In (σ, μ) space the frontier forms two branches radiating from the MVP:
        - Upper branch (μ > μ_MVP): the efficient frontier
        - Lower branch (μ < μ_MVP): the inefficient frontier

    We sweep μ_P from well below the MVP to well above, computing
    σ_P = √(σ²_P) for each value.

    Parameters
    ----------
    params : dict
        Frontier parameters from calculate_frontier_parameters().
    n_points : int
        Number of points per branch (total = 2 * n_points - 1 after dedup).

    Returns
    -------
    dict
        'sigma': np.ndarray of standard deviations,
        'mu': np.ndarray of expected returns.
        Points are ordered from lowest μ to highest μ.
    """
    A, B, C, D = params["A"], params["B"], params["C"], params["D"]

    mu_mvp = A / C

    # Determine a reasonable range for μ_P
    # Go ±3 standard deviations of individual asset returns around the MVP
    mu_range = 3 * np.sqrt(B / C)  # heuristic spread
    mu_min = mu_mvp - mu_range
    mu_max = mu_mvp + mu_range

    mu_values = np.linspace(mu_min, mu_max, 2 * n_points - 1)

    # σ²_P = (1/D)(C·μ² - 2A·μ + B)
    sigma_sq = (1.0 / D) * (C * mu_values ** 2 - 2.0 * A * mu_values + B)

    # Numerical guard: clamp tiny negatives to zero
    sigma_sq = np.maximum(sigma_sq, 0.0)
    sigma_values = np.sqrt(sigma_sq)

    return {"sigma": sigma_values, "mu": mu_values}


def calculate_mvp(params, cov_matrix):
    """
    Compute the Minimum Variance Portfolio (MVP).

    Course reference: Chapter 8 — Minimum Variance Portfolio

    The MVP is the portfolio on the frontier with the lowest possible
    variance. It is the vertex of the mean-variance hyperbola.

    Formulas:
        μ_MVP = A / C
        σ²_MVP = 1 / C   →   σ_MVP = √(1/C)
        w_MVP = Σ⁻¹ 1 / C

    The MVP weights sum to 1 by construction:
        1^T w_MVP = 1^T (Σ⁻¹ 1 / C) = C / C = 1  ✓

    The MVP is the boundary between the efficient (upper) and inefficient
    (lower) portions of the frontier. All rational investors hold portfolios
    on or above the MVP.

    Parameters
    ----------
    params : dict
        Frontier parameters (must contain 'A', 'C', 'Sigma_inv').
    cov_matrix : np.ndarray
        (n, n) covariance matrix (used for verification).

    Returns
    -------
    dict
        'weights': np.ndarray (n,), portfolio weights.
        'expected_return': float, μ_MVP.
        'std_dev': float, σ_MVP.
    """
    A, C, Sigma_inv = params["A"], params["C"], params["Sigma_inv"]
    n = cov_matrix.shape[0]
    ones = np.ones(n)

    weights = Sigma_inv @ ones / C
    expected_return = A / C
    std_dev = np.sqrt(1.0 / C)

    return {
        "weights": weights,
        "expected_return": expected_return,
        "std_dev": std_dev,
    }


def calculate_market_portfolio(params, mean_vector, cov_matrix, rf):
    """
    Compute the Market (Tangency) Portfolio — the risky portfolio that
    maximizes the Sharpe ratio when a riskless asset is available.

    Course reference: Chapter 9 — The Risk-Free Asset and the Capital Market Line
    Huang & Litzenberger, Chapter 5

    The tangency portfolio is found by maximizing:
        Sharpe = (μ_P - μ_F) / σ_P

    The solution yields raw (unnormalized) weights:
        z = Σ⁻¹ (μ - μ_F · 1)

    These are then normalized to sum to 1:
        w_M = z / (1^T z)

    The quantity H measures the squared maximum Sharpe ratio:
        H = (μ - μ_F · 1)^T Σ⁻¹ (μ - μ_F · 1)

    For the tangency portfolio to lie on the EFFICIENT portion of the
    frontier, we require μ_F < A/C (Case 1 in Huang & Litzenberger).
    If μ_F > A/C, the tangent touches the inefficient branch.

    Parameters
    ----------
    params : dict
        Frontier parameters (must contain 'A', 'C', 'Sigma_inv').
    mean_vector : np.ndarray
        (n,) mean return vector μ.
    cov_matrix : np.ndarray
        (n, n) covariance matrix Σ.
    rf : float
        Risk-free rate μ_F (daily, matching the return frequency).

    Returns
    -------
    dict
        'weights': np.ndarray (n,), normalized portfolio weights.
        'expected_return': float, μ_M.
        'std_dev': float, σ_M.
        'sharpe_ratio': float, (μ_M - μ_F) / σ_M.
        'H': float, squared maximum Sharpe ratio.
    """
    A, C, Sigma_inv = params["A"], params["C"], params["Sigma_inv"]
    n = len(mean_vector)
    ones = np.ones(n)
    mu = mean_vector

    # Verify Case 1: μ_F < A/C
    mu_mvp = A / C
    if rf >= mu_mvp:
        print(f"WARNING: rf ({rf:.6f}) >= A/C ({mu_mvp:.6f}). "
              f"Tangency portfolio may be on the inefficient branch.")

    # Excess return vector
    excess = mu - rf * ones

    # H = (μ - μ_F·1)' Σ⁻¹ (μ - μ_F·1)
    H = excess @ Sigma_inv @ excess
    assert H > 0, f"H must be > 0, got H = {H}"

    # Raw weights and normalization
    z = Sigma_inv @ excess
    w_M = z / np.sum(z)

    # Portfolio statistics
    mu_M = w_M @ mu
    sigma_M = np.sqrt(w_M @ cov_matrix @ w_M)
    sharpe = (mu_M - rf) / sigma_M

    return {
        "weights": w_M,
        "expected_return": mu_M,
        "std_dev": sigma_M,
        "sharpe_ratio": sharpe,
        "H": H,
    }


def calculate_zc_portfolio(params, market_return, mean_vector, cov_matrix):
    """
    Compute the Zero-Covariance (ZC) Portfolio corresponding to the
    Market Portfolio.

    Course reference: Chapter 8, Proposition 2 — Zero Covariance Portfolio

    For any frontier portfolio P with return μ_P ≠ μ_MVP, there exists a
    unique frontier portfolio ZC(P) such that Cov(R_P, R_ZC) = 0.

    The expected return of the ZC portfolio is:
        μ_ZC = A/C - (D / C²) / (μ_P - A/C)

    The ZC portfolio weights are obtained from the general frontier weight
    formula. Any frontier portfolio with target return μ_ZC has weights:
        w* = g + h · μ_ZC

    where:
        g = (1/D)(B · Σ⁻¹ 1 - A · Σ⁻¹ μ)
        h = (1/D)(C · Σ⁻¹ μ - A · Σ⁻¹ 1)

    Verification: w_M^T Σ w_ZC should be ≈ 0 (up to floating-point error).

    The ZC portfolio generalizes the role of the risk-free asset: in the
    absence of a riskless asset, the ZC portfolio of the tangency portfolio
    yields the same two-fund separation result.

    Parameters
    ----------
    params : dict
        Frontier parameters (must contain 'A', 'B', 'C', 'D', 'Sigma_inv').
    market_return : float
        Expected return μ_M of the Market Portfolio.
    mean_vector : np.ndarray
        (n,) mean return vector μ.
    cov_matrix : np.ndarray
        (n, n) covariance matrix Σ.

    Returns
    -------
    dict
        'weights': np.ndarray (n,), portfolio weights.
        'expected_return': float, μ_ZC.
        'std_dev': float, σ_ZC.
    """
    A, B, C, D = params["A"], params["B"], params["C"], params["D"]
    Sigma_inv = params["Sigma_inv"]
    n = len(mean_vector)
    ones = np.ones(n)
    mu = mean_vector

    mu_mvp = A / C

    # μ_ZC = A/C - (D/C²) / (μ_M - A/C)
    mu_zc = mu_mvp - (D / (C ** 2)) / (market_return - mu_mvp)

    # General frontier weight formula: w* = g + h * μ_target
    # g = (1/D)(B Σ⁻¹ 1 - A Σ⁻¹ μ)
    # h = (1/D)(C Σ⁻¹ μ - A Σ⁻¹ 1)
    Sigma_inv_ones = Sigma_inv @ ones
    Sigma_inv_mu = Sigma_inv @ mu

    g = (1.0 / D) * (B * Sigma_inv_ones - A * Sigma_inv_mu)
    h = (1.0 / D) * (C * Sigma_inv_mu - A * Sigma_inv_ones)

    w_zc = g + h * mu_zc

    # Portfolio standard deviation
    sigma_zc = np.sqrt(w_zc @ cov_matrix @ w_zc)

    return {
        "weights": w_zc,
        "expected_return": mu_zc,
        "std_dev": sigma_zc,
    }


def verify_matrix_inversion(Sigma, Sigma_inv):
    """
    Verify the quality of the covariance matrix inversion.

    Computes Σ · Σ⁻¹ and checks how close the product is to the identity
    matrix. Also reports the condition number of Σ, which measures how
    sensitive the inverse is to perturbations in the input.

    Parameters
    ----------
    Sigma : np.ndarray
        (n, n) covariance matrix.
    Sigma_inv : np.ndarray
        (n, n) inverse covariance matrix.

    Returns
    -------
    max_deviation : float
        Maximum absolute deviation of Σ · Σ⁻¹ from the identity matrix.
        Returns None if verification fails.
    condition_number : float
        Condition number of Σ (ratio of largest to smallest singular value).
        Returns None if verification fails.
    """
    try:
        n = Sigma.shape[0]
        product = Sigma @ Sigma_inv
        identity = np.eye(n)
        max_deviation = np.max(np.abs(product - identity))
        condition_number = np.linalg.cond(Sigma)
        return max_deviation, condition_number
    except (FloatingPointError, ZeroDivisionError, OverflowError) as e:
        print(f"WARNING: Matrix verification failed: {e}")
        return None, None


def generate_cml_points(rf, market_portfolio, n_points=100):
    """
    Generate (σ, μ) points along the Capital Market Line (CML).

    Course reference: Chapter 9 — Capital Market Line

    When a riskless asset with return μ_F is available, investors can combine
    it with the Market Portfolio to achieve any point on the CML:

        μ_P = μ_F + [(μ_M - μ_F) / σ_M] · σ_P
            = μ_F + √H · σ_P

    The CML:
        - Starts at (0, μ_F) — 100% in the riskless asset
        - Passes through (σ_M, μ_M) — 100% in the Market Portfolio
        - Extends beyond — leveraged positions (borrowing at μ_F)

    The slope of the CML is √H = (μ_M - μ_F) / σ_M, the maximum attainable
    Sharpe ratio. This is the "price of risk" in the economy.

    Parameters
    ----------
    rf : float
        Risk-free rate μ_F.
    market_portfolio : dict
        Must contain 'expected_return', 'std_dev', 'H'.
    n_points : int
        Number of points to generate.

    Returns
    -------
    dict
        'sigma': np.ndarray of standard deviations (from 0 to ~2× σ_M),
        'mu': np.ndarray of expected returns.
    """
    sigma_M = market_portfolio["std_dev"]
    H = market_portfolio["H"]
    slope = np.sqrt(H)

    # Extend to 2× the Market Portfolio's σ
    sigma_values = np.linspace(0, 2.0 * sigma_M, n_points)
    mu_values = rf + slope * sigma_values

    return {"sigma": sigma_values, "mu": mu_values}
