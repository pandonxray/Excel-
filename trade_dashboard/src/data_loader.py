from __future__ import annotations

import logging
from pathlib import Path

import pandas as pd

logger = logging.getLogger(__name__)


def load_timeseries_from_excel(
    workbook_path: str | Path,
    data_sheet: str,
    date_column: str = "Date",
) -> pd.DataFrame:
    try:
        df = pd.read_excel(workbook_path, sheet_name=data_sheet)
    except Exception as exc:
        logger.exception("读取 Excel 数据失败: %s", exc)
        raise

    if date_column not in df.columns:
        raise ValueError(f"缺少日期列: {date_column}")

    df[date_column] = pd.to_datetime(df[date_column], errors="coerce")
    df = df.dropna(subset=[date_column]).drop_duplicates(subset=[date_column])
    df = df.set_index(date_column).sort_index()

    numeric_cols = [c for c in df.columns if c]
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    logger.info("加载时间序列完成，行数=%s，列数=%s", len(df), len(df.columns))
    return df


def load_strategy_table(workbook_path: str | Path, strategy_sheet: str) -> pd.DataFrame:
    df = pd.read_excel(workbook_path, sheet_name=strategy_sheet)
    required = {"StrategyName", "Formula", "Enabled"}
    missing = required.difference(df.columns)
    if missing:
        raise ValueError(f"策略表缺失字段: {missing}")

    df["Enabled"] = df["Enabled"].astype(str).str.upper().isin(["Y", "YES", "TRUE", "1"])
    return df
