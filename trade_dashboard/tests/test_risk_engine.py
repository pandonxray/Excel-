import pandas as pd

from src.risk_engine import (
    build_risk_report,
    max_drawdown,
    percentile_of_value,
    summarize_risk_metrics,
    var_es_over_window,
    zscore_of_value,
)


def test_max_drawdown_negative():
    series = pd.Series([100, 110, 90, 95])
    assert max_drawdown(series) < 0


def test_summary_has_keys():
    series = pd.Series(range(1, 300))
    metrics = summarize_risk_metrics(series)
    assert "Current" in metrics
    assert "VaR_1D_95" in metrics


def test_build_risk_report_contains_percentiles_and_var():
    series = pd.Series(range(50, 400))
    report = build_risk_report(series, {"percentile_windows": [20], "var_horizons": [1], "var_confidence_levels": [0.95]})
    assert report["current_value"] == 399.0
    assert 20 in report["window_percentiles"]
    assert "1D_95" in report["var"]


def test_percentile_of_value_and_zscore_of_value():
    series = pd.Series([1, 2, 3, 4, 5])
    assert percentile_of_value(series, 4, 5) == 80.0
    assert round(zscore_of_value(series, 5, 5), 5) > 1.0


def test_var_es_over_window():
    series = pd.Series([100, 101, 102, 100, 99, 98, 97, 99, 100, 101], dtype=float)
    var_value, es_value, basis = var_es_over_window(series, lookback_window=5, horizon=1, confidence=0.95)
    assert basis == "pct_change"
    assert pd.notna(var_value)
    assert pd.notna(es_value)
