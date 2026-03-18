from __future__ import annotations

import logging
from pathlib import Path
from typing import Iterable

import pandas as pd

logger = logging.getLogger(__name__)
MISSING_MARKERS = {"", "NA", "N/A", "NULL", "NONE", "NAN", "#N/A"}


def _flatten_columns(columns: pd.Index, column_name_row: int) -> list[str]:
    if not isinstance(columns, pd.MultiIndex):
        return [str(col).strip() for col in columns]

    flattened: list[str] = []
    for col in columns:
        parts = [str(part).strip() for part in col]
        selected = parts[column_name_row].strip()
        if selected.startswith("Unnamed:"):
            selected = ""
        flattened.append(selected)
    return flattened


def _coerce_excel_dates(values: pd.Series) -> pd.Series:
    cleaned = values.astype("object").where(~values.isna(), None)
    cleaned = cleaned.map(lambda x: None if isinstance(x, str) and x.strip().upper() in MISSING_MARKERS else x)
    values = pd.Series(cleaned, index=values.index)
    numeric = pd.to_numeric(values, errors="coerce")
    parsed = pd.to_datetime(values, errors="coerce")
    excel_serials = numeric.notna() & numeric.between(20000, 60000)
    if excel_serials.any():
        parsed.loc[excel_serials] = pd.to_datetime(
            numeric.loc[excel_serials],
            unit="D",
            origin="1899-12-30",
            errors="coerce",
        )
    return parsed


def load_timeseries_from_excel(
    workbook_path: str | Path,
    data_sheet: str,
    date_column: str = "Date",
    header_rows: int | Iterable[int] = 0,
    column_name_row: int = 0,
) -> pd.DataFrame:
    try:
        df = pd.read_excel(workbook_path, sheet_name=data_sheet, header=header_rows)
    except Exception as exc:
        logger.exception("Failed to read Excel data: %s", exc)
        raise

    df = df.replace(r"^\s*$", pd.NA, regex=True)
    df = df.replace(list(MISSING_MARKERS), pd.NA)

    df.columns = _flatten_columns(df.columns, column_name_row)
    df = df.loc[:, [bool(str(col).strip()) for col in df.columns]]

    if date_column not in df.columns:
        raise ValueError(f"Missing date column: {date_column}")

    df[date_column] = _coerce_excel_dates(df[date_column])
    df = df.dropna(subset=[date_column]).drop_duplicates(subset=[date_column])
    df = df.set_index(date_column).sort_index()

    numeric_cols = [c for c in df.columns if c]
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    df = df.dropna(axis=1, how="all")

    logger.info("Loaded timeseries data with %s rows and %s columns", len(df), len(df.columns))
    return df


def load_strategy_table(workbook_path: str | Path, strategy_sheet: str) -> pd.DataFrame:
    df = pd.read_excel(workbook_path, sheet_name=strategy_sheet)
    required = {"StrategyName", "Formula", "Enabled"}
    missing = required.difference(df.columns)
    if missing:
        raise ValueError(f"Missing strategy fields: {missing}")

    df["Enabled"] = df["Enabled"].astype(str).str.upper().isin(["Y", "YES", "TRUE", "1"])
    return df
