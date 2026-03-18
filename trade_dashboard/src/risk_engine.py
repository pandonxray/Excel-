from __future__ import annotations

import numpy as np
import pandas as pd


def rolling_zscore(series: pd.Series, window: int) -> pd.Series:
    mean = series.rolling(window).mean()
    std = series.rolling(window).std(ddof=0)
    return (series - mean) / std


def historical_percentile(series: pd.Series, window: int | None = None) -> float:
    sample = series.dropna()
    if window is not None:
        sample = sample.iloc[-window:]
    if sample.empty:
        return np.nan
    return float((sample <= sample.iloc[-1]).mean() * 100)


def percentile_of_value(series: pd.Series, value: float, window: int | None = None) -> float:
    sample = series.dropna()
    if window is not None:
        sample = sample.iloc[-window:]
    if sample.empty:
        return np.nan
    return float((sample <= value).mean() * 100)


def zscore_of_value(series: pd.Series, value: float, window: int) -> float:
    sample = series.dropna().iloc[-window:]
    if sample.empty:
        return np.nan
    std = float(sample.std(ddof=0))
    if std == 0:
        return np.nan
    mean = float(sample.mean())
    return (value - mean) / std


def _return_series(series: pd.Series) -> tuple[pd.Series, str]:
    clean = series.dropna()
    if clean.empty:
        return pd.Series(dtype=float), "pct_change"
    if (clean > 0).all():
        return series.pct_change(), "pct_change"
    return series.diff(), "diff"


def var_es(returns: pd.Series, confidence: float = 0.95) -> tuple[float, float]:
    r = returns.dropna()
    if r.empty:
        return np.nan, np.nan
    var = float(np.percentile(r, (1 - confidence) * 100))
    es = float(r[r <= var].mean()) if (r <= var).any() else var
    return var, es


def var_es_over_window(
    series: pd.Series,
    lookback_window: int,
    horizon: int = 1,
    confidence: float = 0.95,
) -> tuple[float, float, str]:
    returns, return_basis = _return_series(series)
    sample = returns.dropna().iloc[-lookback_window:]
    if horizon > 1:
        sample = sample.rolling(horizon).sum().dropna()
    var_value, es_value = var_es(sample, confidence)
    return var_value, es_value, return_basis


def rolling_volatility(returns: pd.Series, window: int) -> pd.Series:
    return returns.rolling(window).std(ddof=0) * np.sqrt(252)


def max_drawdown(series: pd.Series) -> float:
    s = series.dropna()
    if s.empty:
        return np.nan
    cummax = s.cummax()
    drawdown = (s - cummax) / cummax.abs().replace(0, np.nan)
    return float(drawdown.min())


def rolling_correlation(series_a: pd.Series, series_b: pd.Series, window: int = 60) -> pd.Series:
    return series_a.rolling(window).corr(series_b)


def build_risk_report(series: pd.Series, risk_config: dict | None = None) -> dict[str, object]:
    risk_config = risk_config or {}
    percentile_windows = risk_config.get("percentile_windows", [20, 60, 120, 250])
    zscore_windows = risk_config.get("zscore_windows", [20, 60, 120])
    volatility_windows = risk_config.get("volatility_windows", [20, 60, 120])
    mdd_windows = risk_config.get("mdd_windows", [60, 120, 250])
    var_horizons = risk_config.get("var_horizons", [1, 5])
    var_conf_levels = risk_config.get("var_confidence_levels", [0.95, 0.99])

    returns, return_basis = _return_series(series)
    clean = series.dropna()

    report: dict[str, object] = {
        "current_value": float(clean.iloc[-1]) if not clean.empty else np.nan,
        "full_history_percentile": historical_percentile(series),
        "return_basis": return_basis,
        "window_percentiles": {},
        "zscores": {},
        "volatility": {},
        "max_drawdown": {"full": max_drawdown(series)},
        "var": {},
        "es": {},
    }

    report["window_percentiles"] = {
        int(window): historical_percentile(series, int(window)) for window in percentile_windows
    }
    report["zscores"] = {
        int(window): float(rolling_zscore(series, int(window)).iloc[-1]) for window in zscore_windows
    }
    report["volatility"] = {
        int(window): float(rolling_volatility(returns, int(window)).iloc[-1]) for window in volatility_windows
    }

    for window in mdd_windows:
        sample = series.dropna().iloc[-int(window) :]
        report["max_drawdown"][int(window)] = max_drawdown(sample)

    for horizon in var_horizons:
        horizon = int(horizon)
        horizon_returns = returns if horizon == 1 else returns.rolling(horizon).sum()
        for confidence in var_conf_levels:
            label = f"{horizon}D_{int(confidence * 100)}"
            var_value, es_value = var_es(horizon_returns, float(confidence))
            report["var"][label] = var_value
            report["es"][label] = es_value

    return report


def summarize_risk_metrics(series: pd.Series, risk_config: dict | None = None) -> dict[str, float | str]:
    report = build_risk_report(series, risk_config=risk_config)
    summary: dict[str, float | str] = {
        "Current": report["current_value"],
        "HistPct_All": report["full_history_percentile"],
        "ReturnBasis": str(report["return_basis"]),
        "MaxDrawdown_Full": report["max_drawdown"]["full"],
    }

    for window, value in report["window_percentiles"].items():
        summary[f"Pct_{window}D"] = value
    for window, value in report["zscores"].items():
        summary[f"Z_{window}D"] = value
    for window, value in report["volatility"].items():
        summary[f"Vol_{window}D_Ann"] = value
    for window, value in report["max_drawdown"].items():
        if window == "full":
            continue
        summary[f"MDD_{window}D"] = value
    for label, value in report["var"].items():
        summary[f"VaR_{label}"] = value
    for label, value in report["es"].items():
        summary[f"ES_{label}"] = value

    return summary
