from __future__ import annotations

import numpy as np
import pandas as pd


def remove_feb29(df: pd.DataFrame) -> pd.DataFrame:
    mask = ~((df.index.month == 2) & (df.index.day == 29))
    return df.loc[mask]


def seasonal_matrix(series: pd.Series, years: int = 5) -> pd.DataFrame:
    s = series.dropna()
    if s.empty:
        return pd.DataFrame()

    cutoff_year = s.index.max().year - years + 1
    s = s[s.index.year >= cutoff_year]

    frame = pd.DataFrame({"value": s.values}, index=s.index)
    frame["year"] = frame.index.year
    frame["doy"] = frame.index.strftime("%m-%d")
    return frame.pivot(index="doy", columns="year", values="value").sort_index()


def seasonal_stats(series: pd.Series, years: int = 5) -> dict[str, float]:
    matrix = seasonal_matrix(series, years=years)
    if matrix.empty:
        return {"同期分位": np.nan, "同期偏离": np.nan}

    today = series.dropna().index.max().strftime("%m-%d")
    row = matrix.loc[today].dropna() if today in matrix.index else pd.Series(dtype=float)
    current = series.dropna().iloc[-1] if not series.dropna().empty else np.nan
    if row.empty:
        return {"同期分位": np.nan, "同期偏离": np.nan}

    percentile = float((row <= current).mean() * 100)
    dev = float(current - row.mean())
    return {"同期分位": percentile, "同期偏离": dev}
