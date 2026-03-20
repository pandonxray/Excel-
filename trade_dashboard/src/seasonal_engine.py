from __future__ import annotations

import numpy as np
import pandas as pd


SEASONAL_YEAR = 2001


def remove_feb29(df: pd.DataFrame) -> pd.DataFrame:
    mask = ~((df.index.month == 2) & (df.index.day == 29))
    return df.loc[mask]


def _continuous_seasonal_index() -> pd.Index:
    return pd.date_range(f"{SEASONAL_YEAR}-01-01", f"{SEASONAL_YEAR}-12-31", freq="D").strftime("%m-%d")


def seasonal_matrix(series: pd.Series, years: int = 5, interpolate: bool = True) -> pd.DataFrame:
    s = series.dropna()
    if s.empty:
        return pd.DataFrame()

    cutoff_year = s.index.max().year - years + 1
    s = s[s.index.year >= cutoff_year]

    frame = pd.DataFrame({"value": s.values}, index=s.index)
    frame["year"] = frame.index.year
    frame["doy"] = frame.index.strftime("%m-%d")
    matrix = frame.pivot(index="doy", columns="year", values="value")
    matrix = matrix.reindex(_continuous_seasonal_index())

    if interpolate:
        # Only fill gaps between known observations; keep future tail gaps empty.
        matrix = matrix.interpolate(method="linear", limit_area="inside")

    return matrix


def seasonal_stats(series: pd.Series, years: int = 5) -> dict[str, float]:
    matrix = seasonal_matrix(series, years=years)
    if matrix.empty:
        return {"seasonal_percentile": np.nan, "seasonal_deviation": np.nan}

    latest = series.dropna()
    if latest.empty:
        return {"seasonal_percentile": np.nan, "seasonal_deviation": np.nan}

    today = latest.index.max().strftime("%m-%d")
    row = matrix.loc[today].dropna() if today in matrix.index else pd.Series(dtype=float)
    current = latest.iloc[-1]
    if row.empty:
        return {"seasonal_percentile": np.nan, "seasonal_deviation": np.nan}

    percentile = float((row <= current).mean() * 100)
    deviation = float(current - row.mean())
    return {"seasonal_percentile": percentile, "seasonal_deviation": deviation}
