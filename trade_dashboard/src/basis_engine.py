from __future__ import annotations

import logging

import numpy as np
import pandas as pd

logger = logging.getLogger(__name__)

PP_NORTH_CHINA_SPOT = "PP：拉丝：市场价：华北地区（日）"
PP_EAST_CHINA_SPOT_COLUMNS = [
    "PP：拉丝：PPH-T30S：自提价：宁波：宁波富德（日）",
    "PP：拉丝：PPH-T30S：自提价：上海（日）",
    "PP：拉丝：PPH-T03：自提价：余姚：镇海炼化（日）",
    "PP：拉丝：PPH-T03：自提价：杭州：镇海炼化（日）",
    "PP：拉丝：PPH-T30S：自提价：杭州：宁波富德（日）",
    "PP：拉丝：PPH-T03：自提价：台州：镇海炼化（日）",
]
MA_TAICANG_IMPORT_SPOT = "甲醇：进口：市场价：太仓（日）"
NEA_TARIFF_SWITCH_DATE = pd.Timestamp("2024-01-01")
TARGET_COLUMNS = [
    "PP_basis_north_china",
    "PP_basis_east_china",
    "PL_basis_shandong",
    "PL_basis_nea",
    "PL_basis_sea",
    "MA_basis_taicang",
]


def _empty_series(index: pd.Index, dtype: str = "float64") -> pd.Series:
    return pd.Series(index=index, dtype=dtype)


def _reindex_series(frame: pd.DataFrame, column: str, index: pd.Index, dtype: str = "float64") -> pd.Series:
    if column not in frame.columns:
        logger.warning("Basis source column not found; filling with NaN: %s", column)
        return _empty_series(index, dtype=dtype)
    return frame[column].reindex(index)


def _build_pp_east_min_frame(spot_df: pd.DataFrame, index: pd.Index) -> tuple[pd.Series, pd.Series]:
    east_frame = pd.DataFrame({column: _reindex_series(spot_df, column, index) for column in PP_EAST_CHINA_SPOT_COLUMNS}, index=index)
    min_value = east_frame.min(axis=1, skipna=True)
    min_source = pd.Series(index=index, dtype="object")
    available = east_frame.notna().any(axis=1)
    if available.any():
        min_source.loc[available] = east_frame.loc[available].idxmin(axis=1)
    return min_value, min_source


def _latest_non_null(series: pd.Series) -> str:
    clean = series.dropna()
    if clean.empty:
        return "N/A"
    return str(clean.iloc[-1])


def build_basis_tables(
    wind_df: pd.DataFrame,
    wind_continue_df: pd.DataFrame,
    manual_df: pd.DataFrame,
    spot_df: pd.DataFrame,
    fx_column: str = "USDCHY",
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    union_index = wind_df.index.union(wind_continue_df.index).union(manual_df.index).union(spot_df.index).sort_values()
    basis_df = pd.DataFrame(index=union_index)

    basis_df["PP_spot_north_china"] = _reindex_series(spot_df, PP_NORTH_CHINA_SPOT, union_index)
    basis_df["PP_c"] = _reindex_series(wind_continue_df, "PP_c", union_index)
    basis_df["PP_basis_north_china"] = basis_df["PP_spot_north_china"] - basis_df["PP_c"]

    east_min_value, east_min_source = _build_pp_east_min_frame(spot_df, union_index)
    basis_df["PP_spot_east_china_min"] = east_min_value
    basis_df["PP_spot_east_china_min_source"] = east_min_source
    basis_df["PP_basis_east_china"] = basis_df["PP_spot_east_china_min"] - basis_df["PP_c"]

    basis_df["PL_shandongspot"] = _reindex_series(wind_continue_df, "PL_shandongspot", union_index)
    basis_df["PL_c"] = _reindex_series(wind_continue_df, "PL_c", union_index)
    basis_df["PL_basis_shandong"] = basis_df["PL_shandongspot"] - basis_df["PL_c"]

    fx_series = manual_df[fx_column].reindex(union_index) if fx_column in manual_df.columns else _empty_series(union_index)
    if fx_series.isna().all():
        fx_series = wind_df[fx_column].reindex(union_index) if fx_column in wind_df.columns else _empty_series(union_index)
    else:
        if fx_column in wind_df.columns:
            fx_series = fx_series.combine_first(wind_df[fx_column].reindex(union_index))
    basis_df["USDCHY"] = fx_series
    basis_df["ICIS_NEA"] = _reindex_series(manual_df, "ICIS_NEA", union_index)
    basis_df["ICIS_SEA"] = _reindex_series(manual_df, "ICIS_SEA", union_index)
    basis_df["PL_nea_vat"] = 1.13
    basis_df["PL_nea_tariff"] = np.where(union_index < NEA_TARIFF_SWITCH_DATE, 1.02, 1.0)
    basis_df["PL_spot_nea_import_taxed"] = basis_df["ICIS_NEA"] * basis_df["USDCHY"] * basis_df["PL_nea_vat"] * basis_df["PL_nea_tariff"]
    basis_df["PL_basis_nea"] = basis_df["PL_spot_nea_import_taxed"] - basis_df["PL_c"]

    basis_df["PL_sea_vat"] = 1.13
    basis_df["PL_spot_sea_import_taxed"] = basis_df["ICIS_SEA"] * basis_df["USDCHY"] * basis_df["PL_sea_vat"]
    basis_df["PL_basis_sea"] = basis_df["PL_spot_sea_import_taxed"] - basis_df["PL_c"]

    basis_df["MA_spot_taicang_import"] = _reindex_series(spot_df, MA_TAICANG_IMPORT_SPOT, union_index)
    basis_df["MA_c"] = _reindex_series(wind_continue_df, "MA_c", union_index)
    basis_df["MA_basis_taicang"] = basis_df["MA_spot_taicang_import"] - basis_df["MA_c"]

    display_df = basis_df[TARGET_COLUMNS].copy()
    east_latest_source = _latest_non_null(basis_df["PP_spot_east_china_min_source"])
    basis_meta = pd.DataFrame(
        [
            {
                "metric": "PP_basis_north_china",
                "label": "PP基差_华北",
                "formula": "PP：拉丝：市场价：华北地区（日） - PP_c",
                "note": "基差定义为现货 - 期货，期货腿使用 PP 活跃合约收盘价 PP_c。",
            },
            {
                "metric": "PP_basis_east_china",
                "label": "PP基差_华东",
                "formula": "min(宁波富德宁波, 上海, 镇海炼化余姚, 镇海炼化杭州, 宁波富德杭州, 镇海炼化台州) - PP_c",
                "note": (
                    "华东现货腿取六个华东报价的每日最小值；完整的日度最小值来源保存在 "
                    "PP_spot_east_china_min_source 列。"
                    f" 最新有效来源: {east_latest_source}。"
                ),
            },
            {
                "metric": "PL_basis_shandong",
                "label": "丙烯基差_山东",
                "formula": "PL_shandongspot - PL_c",
                "note": "山东丙烯现货腿使用 wind_continue 表中的 PL_shandongspot。",
            },
            {
                "metric": "PL_basis_nea",
                "label": "丙烯基差_NEA",
                "formula": "(ICIS_NEA * USDCHY * 1.13 * tariff) - PL_c",
                "note": "这里按 2024-01-01 为切换点：2024-01-01 之前 tariff=1.02，2024-01-01 起 tariff=1.00。",
            },
            {
                "metric": "PL_basis_sea",
                "label": "丙烯基差_SEA",
                "formula": "(ICIS_SEA * USDCHY * 1.13) - PL_c",
                "note": "SEA 口径按外盘美元价 * 汇率 * 1.13 计算完税现货，再减去 PL 连续合约。",
            },
            {
                "metric": "MA_basis_taicang",
                "label": "MA基差_太仓进口",
                "formula": "甲醇：进口：市场价：太仓（日） - MA_c",
                "note": "基差定义为太仓进口货现货价减去 MA 活跃合约收盘价。",
            },
        ]
    )
    return basis_df, display_df, basis_meta
