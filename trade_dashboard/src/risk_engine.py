from __future__ import annotations

import numpy as np
import pandas as pd


def rolling_zscore(series: pd.Series, window: int) -> pd.Series:
    mean = series.rolling(window).mean()
    std = series.rolling(window).std(ddof=0)
    return (series - mean) / std


def historical_percentile(series: pd.Series, window: int) -> float:
    sample = series.dropna().iloc[-window:]
    if sample.empty:
        return np.nan
    return float((sample <= sample.iloc[-1]).mean() * 100)


def var_es(returns: pd.Series, confidence: float = 0.95) -> tuple[float, float]:
    r = returns.dropna()
    if r.empty:
        return np.nan, np.nan
    var = float(np.percentile(r, (1 - confidence) * 100))
    es = float(r[r <= var].mean()) if (r <= var).any() else var
    return var, es


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


def summarize_risk_metrics(series: pd.Series) -> dict[str, float]:
    returns = series.pct_change()
    var_1d, es_1d = var_es(returns, 0.95)
    var_5d, es_5d = var_es(returns.rolling(5).sum(), 0.95)
    return {
        "当前值": float(series.dropna().iloc[-1]) if not series.dropna().empty else np.nan,
        "历史分位(120日)": historical_percentile(series, 120),
        "ZScore(60日)": float(rolling_zscore(series, 60).iloc[-1]),
        "VaR_1D(95%)": var_1d,
        "VaR_5D(95%)": var_5d,
        "ES_1D(95%)": es_1d,
        "ES_5D(95%)": es_5d,
        "波动率(60日年化)": float(rolling_volatility(returns, 60).iloc[-1]),
        "最大回撤": max_drawdown(series),
    }
