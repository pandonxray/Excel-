import pandas as pd

from src.seasonal_engine import remove_feb29


def test_remove_feb29():
    idx = pd.to_datetime(["2024-02-28", "2024-02-29", "2024-03-01"])
    df = pd.DataFrame({"x": [1, 2, 3]}, index=idx)
    out = remove_feb29(df)
    assert len(out) == 2
