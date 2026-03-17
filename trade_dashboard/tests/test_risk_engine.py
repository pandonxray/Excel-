import pandas as pd

from src.risk_engine import max_drawdown, summarize_risk_metrics


def test_max_drawdown_negative():
    series = pd.Series([100, 110, 90, 95])
    assert max_drawdown(series) < 0


def test_summary_has_keys():
    series = pd.Series(range(1, 300))
    metrics = summarize_risk_metrics(series)
    assert "当前值" in metrics
    assert "VaR_1D(95%)" in metrics
