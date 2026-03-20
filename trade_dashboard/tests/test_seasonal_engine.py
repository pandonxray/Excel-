import pandas as pd

from src.seasonal_engine import remove_feb29, seasonal_matrix


def test_remove_feb29():
    idx = pd.to_datetime(["2024-02-28", "2024-02-29", "2024-03-01"])
    df = pd.DataFrame({"x": [1, 2, 3]}, index=idx)
    out = remove_feb29(df)
    assert len(out) == 2


def test_seasonal_matrix_is_continuous_after_reindex():
    idx = pd.to_datetime(["2024-01-01", "2024-01-03", "2025-01-01", "2025-01-03"])
    series = pd.Series([1.0, 3.0, 2.0, 4.0], index=idx)
    matrix = seasonal_matrix(series, years=2, interpolate=True)
    assert "01-02" in matrix.index
    assert matrix.loc["01-02"].notna().all()


def test_seasonal_matrix_does_not_fill_future_tail():
    idx = pd.to_datetime(["2026-01-01", "2026-01-02"])
    series = pd.Series([10.0, 11.0], index=idx)
    matrix = seasonal_matrix(series, years=1, interpolate=True)
    assert pd.isna(matrix.loc["01-05", 2026])
