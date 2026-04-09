from __future__ import annotations

from dataclasses import dataclass

import pandas as pd

from .formula_engine import evaluate_formula


@dataclass(frozen=True)
class IndustryMetric:
    name: str
    category: str
    formula: str
    note: str = ""


RAW_ALIAS_MAP: dict[str, list[str]] = {
    "po_shandong": ["环氧丙烷：市场价：山东（日）"],
    "propylene_shandong": ["丙烯：市场价：山东（日）", "山东丙烯"],
    "chlorine_shandong": ["液氯：市场价：山东（日）"],
    "h2o2_shandong": ["双氧水：50%：市场价：山东（日）"],
    "polyether_soft_shandong": ["软泡聚醚：市场价：山东（日）"],
    "acrylic_acid_shandong": ["丙烯酸：普通级：市场价：华北地区（日）"],
    "acrylonitrile_shandong": ["丙烯腈：市场价：山东（日）"],
    "ammonia_shandong": ["合成氨：自提价：山东（日）"],
    "butanol_shandong": ["正丁醇：市场价：华北地区（日）"],
    "octanol_shandong": ["辛醇：市场价：华北地区（日）"],
    "phenol_shandong": ["苯酚：市场价：山东（日）"],
    "acetone_shandong": ["丙酮：市场价：山东（日）"],
    "benzene_shandong": ["纯苯：自提价：山东（日）"],
    "ech_shandong": ["环氧氯丙烷", "环氧氯丙烷：市场价：山东（日）"],
}

PP_POWDER_KEYWORDS = ["PP粉：", "停-PP粉："]

PROFIT_METRICS: list[IndustryMetric] = [
    IndustryMetric(
        "PO利润-氯醇法",
        "利润",
        "po_shandong - 0.85 * (propylene_shandong + 100) - 1.4 * (chlorine_shandong + 100) - chlorohydrin_fee",
        "山东PO - 0.85×(山东丙烯+100) - 1.4×(山东液氯+100) - 氯醇法加工费",
    ),
    IndustryMetric(
        "HPPO利润",
        "利润",
        "po_shandong - 0.8 * (propylene_shandong + 100) - 1.4 * h2o2_shandong - hppo_fee",
        "山东PO - 0.8×(山东丙烯+100) - 1.4×山东双氧水 - 1800",
    ),
    IndustryMetric(
        "软泡聚醚-PO价差",
        "利润",
        "polyether_soft_shandong - po_shandong",
        "山东软泡聚醚 - 山东PO",
    ),
    IndustryMetric(
        "丙烯酸利润",
        "利润",
        "acrylic_acid_shandong - 0.71 * (propylene_shandong + 100) - acrylic_fee",
        "山东丙烯酸 - 0.71×(山东丙烯+100) - 1800",
    ),
    IndustryMetric(
        "丙烯腈利润",
        "利润",
        "acrylonitrile_shandong - 1.05 * (propylene_shandong + 100) - 0.5 * ammonia_shandong - acrylonitrile_fee",
        "山东丙烯腈 - 1.05×(山东丙烯+100) - 0.5×山东合成氨 - 1500",
    ),
    IndustryMetric(
        "丁醇利润",
        "利润",
        "butanol_shandong - 0.62 * (propylene_shandong + 100) - butanol_fee",
        "山东丁醇 - 0.62×(山东丙烯+100) - 1800",
    ),
    IndustryMetric(
        "辛醇利润",
        "利润",
        "octanol_shandong - 0.72 * (propylene_shandong + 100) - octanol_fee",
        "山东辛醇 - 0.72×(山东丙烯+100) - 2300",
    ),
    IndustryMetric(
        "酚酮综合利润",
        "利润",
        "(phenol_shandong + acetone_shandong * 0.6) - (benzene_shandong * 0.92 + 0.505 * (propylene_shandong + 100) + phenol_acetone_fee)",
        "(山东苯酚×1 + 山东丙酮×0.6) - (山东纯苯×0.92 + 0.505×(山东丙烯+100) + 1500)",
    ),
    IndustryMetric(
        "粉料利润",
        "利润",
        "(powder_proxy_shandong - 50) - (propylene_shandong + 100) - powder_fee",
        "(山东粉料-50) - (山东丙烯+100) - 300",
    ),
]

NETBACK_METRICS: list[IndustryMetric] = [
    IndustryMetric(
        "氯醇法-PO净回值",
        "净回值",
        "(po_shandong - 1.4 * (chlorine_shandong + 100) - chlorohydrin_fee) / 0.85 - 100",
        "(山东PO - 1.4×(山东液氯+100) - 氯醇法加工费) / 0.85 - 100",
    ),
    IndustryMetric(
        "丙烯酸净回值",
        "净回值",
        "(acrylic_acid_shandong - acrylic_fee) / 0.71 - 100",
        "(山东丙烯酸 - 1800) / 0.71 - 100",
    ),
    IndustryMetric(
        "丙烯腈净回值",
        "净回值",
        "(acrylonitrile_shandong - 0.5 * ammonia_shandong - acrylonitrile_fee) / 1.05 - 100",
        "(山东丙烯腈 - 0.5×山东合成氨 - 1500) / 1.05 - 100",
    ),
    IndustryMetric(
        "丁醇净回值",
        "净回值",
        "(butanol_shandong - butanol_fee) / 0.62 - 100",
        "(山东丁醇 - 1800) / 0.62 - 100",
    ),
    IndustryMetric(
        "辛醇净回值",
        "净回值",
        "(octanol_shandong - octanol_fee) / 0.72 - 100",
        "(山东辛醇 - 2300) / 0.72 - 100",
    ),
    IndustryMetric(
        "酚酮净回值",
        "净回值",
        "((phenol_shandong + 0.6 * acetone_shandong) - 0.92 * benzene_shandong - phenol_acetone_fee) / 0.505 - 100",
        "((山东苯酚 + 0.6×山东丙酮) - 0.92×山东纯苯 - 1500) / 0.505 - 100",
    ),
    IndustryMetric(
        "粉料净回值",
        "净回值",
        "(powder_proxy_shandong - 50) - powder_fee - 100",
        "(山东粉料 - 50) - 300 - 100",
    ),
]

PROFIT_WEIGHTS = {
    "PO利润-氯醇法": 0.2165754,
    "丙烯酸利润": 0.1008041,
    "丙烯腈利润": 0.2109155,
    "丁醇利润": 0.0593774,
    "辛醇利润": 0.0823884,
    "酚酮综合利润": 0.0474092,
    "粉料利润": 0.2693376,
    "环氧氯丙烷": 0.0089241,
}

NETBACK_WEIGHTS = {
    "氯醇法-PO净回值": 0.2165754,
    "丙烯酸净回值": 0.1008041,
    "丙烯腈净回值": 0.2109155,
    "丁醇净回值": 0.0593774,
    "辛醇净回值": 0.0823884,
    "酚酮净回值": 0.0474092,
    "粉料净回值": 0.2693376,
}


def _resolve_first_available(frame: pd.DataFrame, candidates: list[str]) -> tuple[pd.Series, str | None]:
    for name in candidates:
        if name in frame.columns:
            return pd.to_numeric(frame[name], errors="coerce"), name
    return pd.Series(float("nan"), index=frame.index), None


def _build_powder_proxy(frame: pd.DataFrame) -> tuple[pd.Series, str]:
    exact = [column for column in frame.columns if "华北粉料" in str(column) or "临沂" in str(column)]
    if exact:
        series = pd.to_numeric(frame[exact[0]], errors="coerce")
        return series, f"使用列：{exact[0]}"

    powder_cols = [
        column
        for column in frame.columns
        if any(keyword in str(column) for keyword in PP_POWDER_KEYWORDS) and "停-" not in str(column)
    ]
    if powder_cols:
        series = frame[powder_cols].apply(pd.to_numeric, errors="coerce").mean(axis=1)
        return series, "华北粉料（临沂）列缺失，改用 PP粉现货列均值代理"

    fallback = [column for column in frame.columns if "PP：拉丝：市场价：华北地区（日）" in str(column)]
    if fallback:
        series = pd.to_numeric(frame[fallback[0]], errors="coerce")
        return series, "华北粉料（临沂）列缺失，改用华北PP拉丝价格代理"

    return pd.Series(float("nan"), index=frame.index), "未找到粉料相关价格列"


def _weighted_sum_strict(frame: pd.DataFrame, weights: dict[str, float]) -> pd.Series:
    weighted = pd.DataFrame(index=frame.index)
    for column, weight in weights.items():
        weighted[column] = frame[column] * weight
    return weighted.sum(axis=1, min_count=len(weights))


def build_propylene_profit_dashboard(spot_df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    alias_df = pd.DataFrame(index=spot_df.index)
    metadata_rows: list[dict[str, object]] = []

    for alias, candidates in RAW_ALIAS_MAP.items():
        series, source_name = _resolve_first_available(spot_df, candidates)
        alias_df[alias] = series
        metadata_rows.append(
            {
                "metric": alias,
                "category": "原始映射",
                "formula": source_name or "未找到对应列",
                "note": "" if source_name else "原始列缺失，当前以 NaN 处理",
            }
        )

    powder_proxy, powder_note = _build_powder_proxy(spot_df)
    alias_df["powder_proxy_shandong"] = powder_proxy
    metadata_rows.append(
        {"metric": "powder_proxy_shandong", "category": "原始映射", "formula": "粉料代理价格", "note": powder_note}
    )

    constants = {
        "chlorohydrin_fee": 1800.0,
        "hppo_fee": 1800.0,
        "acrylic_fee": 1800.0,
        "acrylonitrile_fee": 1500.0,
        "butanol_fee": 1800.0,
        "octanol_fee": 2300.0,
        "phenol_acetone_fee": 1500.0,
        "powder_fee": 300.0,
    }
    for name, value in constants.items():
        alias_df[name] = value

    result = pd.DataFrame(index=spot_df.index)
    for metric in PROFIT_METRICS + NETBACK_METRICS:
        result[metric.name] = evaluate_formula(alias_df, metric.formula)
        metadata_rows.append(
            {"metric": metric.name, "category": metric.category, "formula": metric.formula, "note": metric.note}
        )

    composite_profit_frame = pd.DataFrame(index=result.index)
    for metric_name in PROFIT_WEIGHTS:
        if metric_name == "环氧氯丙烷":
            composite_profit_frame[metric_name] = alias_df["ech_shandong"]
        else:
            composite_profit_frame[metric_name] = result[metric_name]
    result["下游综合利润"] = _weighted_sum_strict(composite_profit_frame, PROFIT_WEIGHTS)

    composite_netback_frame = pd.DataFrame(index=result.index)
    for metric_name in NETBACK_WEIGHTS:
        composite_netback_frame[metric_name] = result[metric_name]
    result["综合净回值"] = _weighted_sum_strict(composite_netback_frame, NETBACK_WEIGHTS)

    ech_note = ""
    if alias_df["ech_shandong"].dropna().empty:
        ech_note = "工作簿中未找到环氧氯丙烷列，因此对应日期的下游综合利润暂不计算"

    metadata_rows.append(
        {
            "metric": "下游综合利润",
            "category": "综合",
            "formula": "按给定权重汇总利润项，并额外叠加 0.0089241×环氧氯丙烷；任一项缺失时该日不计算",
            "note": ech_note,
        }
    )
    metadata_rows.append(
        {
            "metric": "综合净回值",
            "category": "综合",
            "formula": "按同样权重汇总各净回值（不含环氧氯丙烷净回值）；任一项缺失时该日不计算",
            "note": "",
        }
    )

    return result.sort_index(), pd.DataFrame(metadata_rows)
