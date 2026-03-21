from __future__ import annotations

import logging
import re
import sys
import tempfile
from pathlib import Path

import numpy as np
import pandas as pd

if not hasattr(np, "unicode_"):
    np.unicode_ = np.str_

import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

try:
    from .data_loader import load_timeseries_from_excel
    from .driver_engine import (
        build_driver_diagnostics,
        build_driver_package,
        compute_factor_sensitivity,
        decompose_change,
        decompose_change_between_dates,
        run_driver_scenarios,
    )
    from .excel_refresh import refresh_excel_workbook
    from .portfolio_engine import build_portfolios
    from .risk_engine import (
        build_risk_report,
        percentile_of_value,
        var_es_over_window,
        zscore_of_value,
    )
    from .seasonal_engine import remove_feb29, seasonal_matrix, seasonal_stats
    from .utils import load_yaml, setup_logging
except ImportError:
    from src.data_loader import load_timeseries_from_excel
    from src.driver_engine import (
        build_driver_diagnostics,
        build_driver_package,
        compute_factor_sensitivity,
        decompose_change,
        decompose_change_between_dates,
        run_driver_scenarios,
    )
    from src.excel_refresh import refresh_excel_workbook
    from src.portfolio_engine import build_portfolios
    from src.risk_engine import (
        build_risk_report,
        percentile_of_value,
        var_es_over_window,
        zscore_of_value,
    )
    from src.seasonal_engine import remove_feb29, seasonal_matrix, seasonal_stats
    from src.utils import load_yaml, setup_logging


if getattr(sys, "frozen", False):
    BASE_DIR = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parents[1]))
else:
    BASE_DIR = Path(__file__).resolve().parents[1]

APP_CONFIG = load_yaml(BASE_DIR / "config" / "app.yaml")
METRIC_CONFIG = load_yaml(BASE_DIR / "config" / "metric.yaml")
setup_logging(APP_CONFIG.get("logging", {}).get("level", "INFO"), APP_CONFIG.get("logging", {}).get("file"))
logger = logging.getLogger(__name__)

SOURCE_LABELS = {"wind": "Wind", "manual": "Manual"}


def _save_uploaded_excel(uploaded_file) -> str:
    suffix = Path(uploaded_file.name).suffix or ".xlsx"
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(uploaded_file.getbuffer())
        return tmp.name


@st.cache_data(show_spinner=False)
def load_all_data(workbook_path: str) -> tuple[dict[str, pd.DataFrame], pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    excel_cfg = APP_CONFIG["excel"]
    workbook = Path(workbook_path)

    wind_df = load_timeseries_from_excel(
        workbook,
        excel_cfg["data_sheet"],
        excel_cfg["date_column"],
        excel_cfg.get("header_rows", 0),
        excel_cfg.get("column_name_row", 0),
    )
    manual_df = load_timeseries_from_excel(
        workbook,
        excel_cfg.get("manual_sheet", "manual_data"),
        excel_cfg.get("manual_date_column", "price_date"),
        excel_cfg.get("manual_header_rows", 0),
        excel_cfg.get("manual_column_name_row", 0),
    )

    merged_for_formula = wind_df.join(manual_df, how="outer").sort_index()

    strategy_cfg = load_yaml(BASE_DIR / "config" / "strategy.yaml")
    strategy_df = pd.DataFrame(strategy_cfg.get("strategies", []))
    strategy_df = strategy_df.rename(
        columns={"name": "StrategyName", "formula": "Formula", "enabled": "Enabled", "category": "Category", "notes": "Notes"}
    )
    portfolios = build_portfolios(merged_for_formula, strategy_df) if not strategy_df.empty else pd.DataFrame(index=merged_for_formula.index)

    return {"wind": wind_df, "manual": manual_df}, portfolios, strategy_df, merged_for_formula


def _format_metric(value: float | str, style: str = "number") -> str:
    if isinstance(value, str):
        return value
    if pd.isna(value):
        return "N/A"
    if style == "percentile":
        return f"{value:.2f}%"
    if style == "ratio_pct":
        return f"{value * 100:.2f}%"
    return f"{value:.4f}"


def _series_groups(columns: list[str]) -> dict[str, list[str]]:
    groups: dict[str, list[str]] = {}
    for col in columns:
        match = re.match(r"([A-Za-z]+)", col)
        prefix = match.group(1) if match else "Other"
        groups.setdefault(prefix, []).append(col)
    return dict(sorted(groups.items()))


def _grouped_selectbox(label_prefix: str, columns: list[str], key_prefix: str) -> str:
    groups = _series_groups(columns)
    group_names = list(groups.keys())
    group = st.sidebar.selectbox(f"{label_prefix}类别", group_names, key=f"{key_prefix}_group")
    return st.sidebar.selectbox(f"{label_prefix}合约", groups[group], key=f"{key_prefix}_item")


def _grouped_strategy_selectbox(strategy_df: pd.DataFrame, key_prefix: str) -> str:
    grouped: dict[str, list[str]] = {}
    for _, row in strategy_df.iterrows():
        category = str(row.get("Category", "Other") or "Other")
        grouped.setdefault(category, []).append(str(row["StrategyName"]))

    category_names = sorted(grouped.keys())
    category = st.sidebar.selectbox("预设组合分类", category_names, key=f"{key_prefix}_category")
    return st.sidebar.selectbox("预设组合", grouped[category], key=f"{key_prefix}_name")


def _lookup_strategy_row(strategy_df: pd.DataFrame, name: str) -> pd.Series | None:
    matched = strategy_df[strategy_df["StrategyName"] == name]
    if matched.empty:
        return None
    return matched.iloc[0]


def _month_spread_candidates(columns: list[str]) -> dict[str, list[str]]:
    groups: dict[str, list[str]] = {}
    for column in columns:
        match = re.match(r"^([A-Za-z]+)(\d{2})$", str(column))
        if not match:
            continue
        product = match.group(1)
        groups.setdefault(product, []).append(str(column))

    result: dict[str, list[str]] = {}
    for product, product_columns in groups.items():
        result[product] = sorted(product_columns, key=lambda value: int(value[-2:]))
    return dict(sorted(result.items()))


def _build_month_spread_target(source_df: pd.DataFrame) -> tuple[str, pd.DataFrame, str, pd.DataFrame]:
    candidates = _month_spread_candidates(source_df.columns.tolist())
    if not candidates:
        empty = pd.DataFrame(index=source_df.index, columns=["target"])
        coverage = pd.DataFrame(columns=["指标", "开始", "结束", "有效点数"])
        return "月差", empty, "", coverage

    product = st.sidebar.selectbox("选择品种", list(candidates.keys()), key="month_spread_product")
    product_columns = candidates[product]
    front_leg = st.sidebar.selectbox("近月合约", product_columns, key="month_spread_front")
    default_back_index = 1 if len(product_columns) > 1 else 0
    back_leg = st.sidebar.selectbox("远月合约", product_columns, index=default_back_index, key="month_spread_back")

    frame = source_df[[front_leg, back_leg]].copy()
    frame["target"] = frame[front_leg] - frame[back_leg]
    coverage = _build_coverage_table(frame, [front_leg, back_leg, "target"])
    formula = f"{front_leg} - {back_leg}"
    target_name = f"{product} {front_leg[-2:]}-{back_leg[-2:]} 月差"
    return target_name, frame, formula, coverage


def _series_valid_range(series: pd.Series) -> tuple[pd.Timestamp | None, pd.Timestamp | None]:
    clean = series.dropna()
    if clean.empty:
        return None, None
    return clean.index.min(), clean.index.max()


def _build_coverage_table(frame: pd.DataFrame, required_cols: list[str]) -> pd.DataFrame:
    rows: list[dict[str, object]] = []
    for col in required_cols:
        start, end = _series_valid_range(frame[col])
        rows.append(
            {
                "区间": f"{col} 有效区间",
                "开始": start.date().isoformat() if start is not None else "N/A",
                "结束": end.date().isoformat() if end is not None else "N/A",
                "样本数": int(frame[col].notna().sum()),
            }
        )

    overlap = frame.dropna(subset=required_cols)
    overlap_start, overlap_end = _series_valid_range(overlap["target"]) if "target" in overlap.columns else (None, None)
    rows.append(
        {
            "区间": "全部项共同有值区间",
            "开始": overlap_start.date().isoformat() if overlap_start is not None else "N/A",
            "结束": overlap_end.date().isoformat() if overlap_end is not None else "N/A",
            "样本数": int(len(overlap)),
        }
    )
    return pd.DataFrame(rows)


def _apply_date_filter(frame: pd.DataFrame, required_cols: list[str], key_prefix: str) -> tuple[pd.DataFrame, str]:
    clean_index = frame.dropna(how="all").index
    if clean_index.empty:
        return frame.iloc[0:0], "空区间"

    mode = st.sidebar.selectbox("分析区间", ["全区间", "共同有值区间", "自定义区间"], key=f"{key_prefix}_range_mode")
    if mode == "共同有值区间":
        return frame.dropna(subset=required_cols), mode

    if mode == "自定义区间":
        start_default, end_default = clean_index.min().date(), clean_index.max().date()
        start_date, end_date = st.sidebar.date_input(
            "选择开始/结束日期",
            value=(start_default, end_default),
            min_value=start_default,
            max_value=end_default,
            key=f"{key_prefix}_date_input",
        )
        filtered = frame.loc[(frame.index >= pd.Timestamp(start_date)) & (frame.index <= pd.Timestamp(end_date))]
        return filtered, mode

    return frame, mode


def _render_metric_table(title: str, data: dict[str, float], style: str = "number") -> None:
    frame = pd.DataFrame({"指标": list(data.keys()), "数值": [_format_metric(value, style=style) for value in data.values()]})
    st.subheader(title)
    st.dataframe(frame, use_container_width=True, hide_index=True)


def _risk_controls(series: pd.Series) -> dict[str, float]:
    st.sidebar.markdown("### 风控参数")
    max_window = max(20, min(len(series.dropna()), 2000))
    percentile_window = int(st.sidebar.number_input("百分位窗口", min_value=20, max_value=max_window, value=min(250, max_window), step=10))
    zscore_window = int(st.sidebar.number_input("ZScore窗口", min_value=20, max_value=max_window, value=min(60, max_window), step=10))
    var_lookback = int(st.sidebar.number_input("VaR回看窗口", min_value=20, max_value=max_window, value=min(250, max_window), step=10))
    var_horizon = int(st.sidebar.number_input("VaR期限(日)", min_value=1, max_value=20, value=5, step=1))
    var_confidence = float(st.sidebar.slider("VaR置信度", min_value=0.80, max_value=0.995, value=0.95, step=0.005))
    default_value = float(series.dropna().iloc[-1]) if not series.dropna().empty else 0.0
    custom_value = float(st.sidebar.number_input("自定义数值查看分位", value=default_value))
    return {
        "percentile_window": percentile_window,
        "zscore_window": zscore_window,
        "var_lookback": var_lookback,
        "var_horizon": var_horizon,
        "var_confidence": var_confidence,
        "custom_value": custom_value,
    }


def _render_custom_risk_panel(series: pd.Series, controls: dict[str, float]) -> None:
    custom_percentile = percentile_of_value(series, controls["custom_value"], int(controls["percentile_window"]))
    custom_zscore = zscore_of_value(series, controls["custom_value"], int(controls["zscore_window"]))
    custom_var, custom_es, return_basis = var_es_over_window(
        series,
        int(controls["var_lookback"]),
        int(controls["var_horizon"]),
        float(controls["var_confidence"]),
    )

    cols = st.columns(4)
    cols[0].metric("输入值", _format_metric(controls["custom_value"]))
    cols[1].metric("输入值百分位", _format_metric(custom_percentile, style="percentile"))
    cols[2].metric("输入值ZScore", _format_metric(custom_zscore))
    cols[3].metric("VaR收益口径", return_basis)

    lower = st.columns(2)
    lower[0].metric(f"VaR ({int(controls['var_horizon'])}D, {controls['var_confidence']:.1%})", _format_metric(custom_var))
    lower[1].metric(f"ES ({int(controls['var_horizon'])}D, {controls['var_confidence']:.1%})", _format_metric(custom_es))


def _render_risk_section(series: pd.Series, controls: dict[str, float]) -> None:
    risk_cfg = {
        "percentile_windows": sorted(set(METRIC_CONFIG.get("risk", {}).get("percentile_windows", []) + [int(controls["percentile_window"])])),
        "zscore_windows": sorted(set(METRIC_CONFIG.get("risk", {}).get("zscore_windows", []) + [int(controls["zscore_window"])])),
        "volatility_windows": METRIC_CONFIG.get("risk", {}).get("volatility_windows", [20, 60, 120]),
        "mdd_windows": METRIC_CONFIG.get("risk", {}).get("mdd_windows", [60, 120, 250]),
        "var_horizons": sorted(set(METRIC_CONFIG.get("risk", {}).get("var_horizons", []) + [int(controls["var_horizon"])])),
        "var_confidence_levels": sorted(set(METRIC_CONFIG.get("risk", {}).get("var_confidence_levels", []) + [float(controls["var_confidence"])])),
    }
    report = build_risk_report(series, risk_cfg)

    top = st.columns(4)
    top[0].metric("当前值", _format_metric(report["current_value"]))
    top[1].metric("全历史百分位", _format_metric(report["full_history_percentile"], style="percentile"))
    top[2].metric("收益口径", _format_metric(report["return_basis"]))
    top[3].metric("全历史最大回撤", _format_metric(report["max_drawdown"]["full"], style="ratio_pct"))

    _render_custom_risk_panel(series, controls)

    col1, col2 = st.columns(2)
    with col1:
        _render_metric_table("滚动百分位", {f"{k}日": v for k, v in report["window_percentiles"].items()}, style="percentile")
        _render_metric_table("滚动 Z-Score", {f"{k}日": v for k, v in report["zscores"].items()})
    with col2:
        _render_metric_table("历史 VaR", report["var"])
        _render_metric_table("历史 ES", report["es"])

    lower_left, lower_right = st.columns(2)
    with lower_left:
        _render_metric_table("年化波动率", {f"{k}日": v for k, v in report["volatility"].items()}, style="ratio_pct")
    with lower_right:
        drawdown_metrics = {("全历史" if k == "full" else f"近{k}日"): v for k, v in report["max_drawdown"].items()}
        _render_metric_table("最大回撤", drawdown_metrics, style="ratio_pct")


def _render_metric_explainer() -> None:
    with st.expander("术语说明与计算逻辑", expanded=False):
        st.markdown(
            "`当前值`：所选序列最新一个有效日度值。\n\n"
            "`历史分位`：当前值在回看窗口中的相对位置，数值越高说明越靠近历史高位。\n\n"
            "`Z-Score`：`(当前值 - 窗口均值) / 窗口标准差`，衡量偏离均值的程度。\n\n"
            "`VaR`：在给定置信度和持有期下，可能承受的最大损失阈值。\n\n"
            "`ES`：超过 VaR 后的平均损失，通常比 VaR 更保守。\n\n"
            "`波动率`：根据历史收益波动估算出的风险强弱。\n\n"
            "`最大回撤`：从历史高点回落到低点的最大跌幅。\n\n"
            "`季节性分位`：当前值在同季节历史样本中的位置。\n\n"
            "`季节性偏离`：当前值相对于季节性均值的偏离程度。"
        )
        st.markdown(
            "计算逻辑概览：\n\n"
            "1. 分位数：统计窗口内小于等于当前值的样本占比。\n\n"
            "2. Z-Score：用窗口均值和标准差标准化当前值。\n\n"
            "3. VaR / ES：基于历史收益分布，在给定窗口、置信度和持有期下估算。\n\n"
            "4. 最大回撤：根据历史净值路径，计算从阶段高点到低点的最大跌幅。\n\n"
            "5. 季节图：把不同年份映射到同一自然年坐标后进行对比。"
        )


def _render_seasonality_section_v2(series: pd.Series) -> None:
    seasonal_years = APP_CONFIG.get("analysis", {}).get("seasonal_years", 5)
    years = st.slider("季节性回看年数", 3, 15, seasonal_years, key="seasonality_years_v2")
    seasonal_source = series.dropna()
    if APP_CONFIG.get("analysis", {}).get("remove_feb29", True):
        seasonal_source = remove_feb29(seasonal_source.to_frame("value"))["value"]

    matrix = seasonal_matrix(seasonal_source, years, interpolate=True)
    if matrix.empty:
        st.info("当前序列样本不足，暂时无法生成季节图。")
        return

    axis_left, axis_right = st.columns(2)
    y_min_text = axis_left.text_input("季节图纵轴下限", value="", key="seasonality_y_min_v2")
    y_max_text = axis_right.text_input("季节图纵轴上限", value="", key="seasonality_y_max_v2")
    y_range = None
    try:
        if y_min_text.strip() or y_max_text.strip():
            y_range = [
                float(y_min_text) if y_min_text.strip() else None,
                float(y_max_text) if y_max_text.strip() else None,
            ]
    except ValueError:
        st.warning("季节图纵轴范围输入无效，请输入数字。")

    continuous_index = pd.to_datetime("2001-" + matrix.index)
    matrix_plot = matrix.copy()
    matrix_plot.index = continuous_index
    fig_matrix = px.line(matrix_plot, title="季节路径对比")
    if y_range is not None:
        fig_matrix.update_yaxes(range=y_range)
    st.plotly_chart(fig_matrix, use_container_width=True)

    mean = matrix.mean(axis=1)
    std = matrix.std(axis=1)
    band = pd.DataFrame({"均值": mean, "+1σ": mean + std, "-1σ": mean - std}, index=continuous_index)
    if band.dropna(how="all").empty:
        st.info("当前样本不足，暂时无法生成季节均值带。")
    else:
        fig_band = px.line(band.dropna(how="all"), title="季节均值与波动带")
        if y_range is not None:
            fig_band.update_yaxes(range=y_range)
        st.plotly_chart(fig_band, use_container_width=True)

    monthly = seasonal_source.to_frame("value")
    monthly["month"] = monthly.index.month
    fig_box = px.box(monthly, x="month", y="value", title="月度分布箱线图")
    if y_range is not None:
        fig_box.update_yaxes(range=y_range)
    st.plotly_chart(fig_box, use_container_width=True)

    s_metrics = seasonal_stats(seasonal_source, years)
    cols = st.columns(2)
    cols[0].metric("季节性分位", _format_metric(s_metrics["seasonal_percentile"], style="percentile"))
    cols[1].metric("季节性偏离", _format_metric(s_metrics["seasonal_deviation"]))


def _render_seasonality_section(series: pd.Series) -> None:
    seasonal_years = APP_CONFIG.get("analysis", {}).get("seasonal_years", 5)
    years = st.slider("季节性回看年数", 3, 15, seasonal_years)
    seasonal_source = series.dropna()
    if APP_CONFIG.get("analysis", {}).get("remove_feb29", True):
        seasonal_source = remove_feb29(seasonal_source.to_frame("value"))["value"]

    matrix = seasonal_matrix(seasonal_source, years, interpolate=True)
    if matrix.empty:
        st.info("当前序列可用于季节性分析的数据不足。")
        return

    continuous_index = pd.to_datetime("2001-" + matrix.index)
    matrix_plot = matrix.copy()
    matrix_plot.index = continuous_index
    axis_left, axis_right = st.columns(2)
    st.caption("如需放大或压缩季节图波动，可手动输入纵轴上下限；留空则使用默认范围。")
    y_min_text = axis_left.text_input("季节图纵轴下限", value="", key="seasonality_y_min")
    y_max_text = axis_right.text_input("季节图纵轴上限", value="", key="seasonality_y_max")
    st.plotly_chart(px.line(matrix_plot, title="历年季节性曲线（连续化）"), use_container_width=True)

    mean = matrix.mean(axis=1)
    std = matrix.std(axis=1)
    band = pd.DataFrame({"均值": mean, "+1σ": mean + std, "-1σ": mean - std}, index=continuous_index)
    if band.dropna(how="all").empty:
        st.info("当前区间有效季节性样本不足，暂时无法生成季节性均值带。")
    else:
        st.plotly_chart(px.line(band.dropna(how="all"), title="季节性均值带"), use_container_width=True)

    monthly = seasonal_source.to_frame("value")
    monthly["month"] = monthly.index.month
    st.plotly_chart(px.box(monthly, x="month", y="value", title="月度分布"), use_container_width=True)

    s_metrics = seasonal_stats(seasonal_source, years)
    cols = st.columns(2)
    cols[0].metric("当前值同期百分位", _format_metric(s_metrics["seasonal_percentile"], style="percentile"))
    cols[1].metric("当前值同期偏离", _format_metric(s_metrics["seasonal_deviation"]))


def _build_term_expression(coefficient: float, column: str, multiplier_col: str | None) -> str:
    coeff_prefix = "" if coefficient == 1 else f"{coefficient:g} * "
    expr = f"{coeff_prefix}{column}"
    if multiplier_col:
        expr = f"{expr} * {multiplier_col}"
    return expr


def _collect_term_configs(sources: dict[str, pd.DataFrame], mode: str) -> list[dict[str, object]]:
    term_count = st.sidebar.selectbox("项数", [2, 3], key=f"{mode}_term_count")
    fx_default = APP_CONFIG["excel"].get("fx_column", "USDCHY")
    wind_columns = sources["wind"].columns.tolist()
    terms: list[dict[str, object]] = []

    for idx in range(term_count):
        st.sidebar.markdown(f"#### 第 {idx + 1} 项")
        source_key = st.sidebar.selectbox(
            f"第 {idx + 1} 项数据源",
            ["wind", "manual"],
            format_func=lambda x: SOURCE_LABELS[x],
            key=f"{mode}_term_{idx}_source",
        )
        column = _grouped_selectbox(f"第 {idx + 1} 项", sources[source_key].columns.tolist(), f"{mode}_term_{idx}")
        coefficient = st.sidebar.number_input(
            f"第 {idx + 1} 项系数",
            value=(1.0 if idx == 0 else (-1.0 if mode == "价差" else 1.0)),
            step=0.1,
            format="%.2f",
            key=f"{mode}_term_{idx}_coef",
        )
        multiplier_col = None
        if st.sidebar.checkbox(f"第 {idx + 1} 项乘 Wind 列", value=False, key=f"{mode}_term_{idx}_mult_flag"):
            default_index = wind_columns.index(fx_default) if fx_default in wind_columns else 0
            multiplier_col = st.sidebar.selectbox(
                f"第 {idx + 1} 项乘数列",
                wind_columns,
                index=default_index,
                key=f"{mode}_term_{idx}_mult_col",
            )

        side = "numerator"
        if mode == "价格比":
            default_side = "numerator" if idx == 0 else "denominator"
            side = st.sidebar.selectbox(
                f"第 {idx + 1} 项归属",
                ["numerator", "denominator"],
                format_func=lambda x: "分子" if x == "numerator" else "分母",
                index=0 if default_side == "numerator" else 1,
                key=f"{mode}_term_{idx}_side",
            )

        terms.append(
            {
                "source": source_key,
                "column": column,
                "coefficient": float(coefficient),
                "multiplier_col": multiplier_col,
                "side": side,
            }
        )

    return terms


def _build_combo_frame(
    sources: dict[str, pd.DataFrame],
    mode: str,
    terms: list[dict[str, object]],
) -> tuple[pd.DataFrame, str, str, list[str]]:
    frame = pd.DataFrame(index=sources["wind"].index.union(sources["manual"].index))
    required_cols: list[str] = []
    numerator_exprs: list[str] = []
    denominator_exprs: list[str] = []
    spread_exprs: list[str] = []
    target_parts: list[pd.Series] = []
    denominator_parts: list[pd.Series] = []
    name_tokens: list[str] = []

    for idx, term in enumerate(terms, start=1):
        source_key = str(term["source"])
        column = str(term["column"])
        coefficient = float(term["coefficient"])
        multiplier_col = term["multiplier_col"]

        base_col_name = f"term_{idx}_{column}"
        base_series = sources[source_key][column].rename(base_col_name)
        frame = frame.join(base_series, how="outer")
        required_cols.append(base_col_name)
        value_series = frame[base_col_name]

        if multiplier_col:
            mult_col_name = f"term_{idx}_{multiplier_col}"
            if mult_col_name not in frame.columns:
                frame = frame.join(sources["wind"][str(multiplier_col)].rename(mult_col_name), how="outer")
            required_cols.append(mult_col_name)
            value_series = value_series * frame[mult_col_name]

        expr = _build_term_expression(coefficient, column, str(multiplier_col) if multiplier_col else None)
        calc_col = f"calc_term_{idx}"
        frame[calc_col] = coefficient * value_series
        required_cols.append(calc_col)
        name_tokens.append(column)

        if mode == "价格比":
            if term["side"] == "numerator":
                target_parts.append(frame[calc_col])
                numerator_exprs.append(expr)
            else:
                denominator_parts.append(frame[calc_col])
                denominator_exprs.append(expr)
        else:
            target_parts.append(frame[calc_col])
            spread_exprs.append(expr)

    if mode == "价格比":
        numerator = sum(target_parts[1:], target_parts[0]) if len(target_parts) > 1 else target_parts[0]
        denominator = sum(denominator_parts[1:], denominator_parts[0]) if len(denominator_parts) > 1 else denominator_parts[0]
        frame["target"] = numerator / denominator.replace(0, pd.NA)
        formula = f"({' + '.join(numerator_exprs)}) / ({' + '.join(denominator_exprs)})"
        target_name = "_".join(name_tokens) + "_ratio"
    else:
        target = sum(target_parts[1:], target_parts[0]) if len(target_parts) > 1 else target_parts[0]
        frame["target"] = target
        formula = " + ".join(spread_exprs).replace("+ -", "- ")
        target_name = "_".join(name_tokens) + "_spread"

    return frame, target_name, formula, required_cols


def _sidebar_excel_path() -> str:
    default_path = APP_CONFIG["excel"]["workbook_path"]
    st.sidebar.markdown("### Excel 数据源")
    source_mode = st.sidebar.radio("选择方式", ["本地路径", "拖拽/上传 Excel"], key="excel_source_mode")

    if source_mode == "本地路径":
        if "excel_path" not in st.session_state:
            st.session_state["excel_path"] = default_path
        excel_path = st.sidebar.text_input("Excel 路径", value=st.session_state["excel_path"])
        st.session_state["excel_path"] = excel_path.strip() or default_path
        return st.session_state["excel_path"]

    uploaded_file = st.sidebar.file_uploader(
        "拖拽或选择 Excel 文件",
        type=["xlsx", "xlsm", "xls"],
        key="excel_uploader",
        help="支持直接拖拽 Excel 到这里，或点击后从本地文件夹选择。",
    )
    if uploaded_file is None:
        return default_path

    upload_key = f"{uploaded_file.name}_{uploaded_file.size}"
    if st.session_state.get("uploaded_excel_key") != upload_key:
        st.session_state["uploaded_excel_key"] = upload_key
        st.session_state["uploaded_excel_path"] = _save_uploaded_excel(uploaded_file)
        load_all_data.clear()
    return st.session_state["uploaded_excel_path"]


def _select_analysis_target(
    sources: dict[str, pd.DataFrame],
    portfolios: pd.DataFrame,
    strategy_df: pd.DataFrame,
) -> tuple[str, pd.Series, str, pd.DataFrame | None, pd.DataFrame | None, pd.Series | None]:
    mode = st.sidebar.radio("分析对象", ["单品种", "跨表/价差组合", "预设组合"])

    if mode == "单品种":
        source_key = st.sidebar.selectbox("数据来源", ["wind", "manual"], format_func=lambda x: SOURCE_LABELS[x])
        source_df = sources[source_key]
        column = _grouped_selectbox("选择", source_df.columns.tolist(), "single")
        frame = source_df[[column]].rename(columns={column: "target"})
        frame, _ = _apply_date_filter(frame, ["target"], "single")
        return f"{SOURCE_LABELS[source_key]}:{column}", frame["target"], column, None, None, None

    if mode == "跨表/价差组合":
        combo_mode = st.sidebar.selectbox("组合方式", ["价差", "价格比"])
        terms = _collect_term_configs(sources, combo_mode)
        combo_frame, default_name, formula, required_cols = _build_combo_frame(sources, combo_mode, terms)
        filtered_frame, _ = _apply_date_filter(combo_frame, required_cols, "combo")
        coverage = _build_coverage_table(combo_frame, required_cols)
        custom_name = st.sidebar.text_input("组合名称", value="")
        return custom_name.strip() or default_name, filtered_frame["target"], formula, combo_frame, coverage, None

    if portfolios.empty:
        fallback = sources["wind"].columns[0]
        return fallback, sources["wind"][fallback], fallback, None, None, None

    name = _grouped_strategy_selectbox(strategy_df, "preset")
    formula_map = dict(zip(strategy_df["StrategyName"], strategy_df["Formula"]))
    frame = portfolios[[name]].rename(columns={name: "target"})
    frame, _ = _apply_date_filter(frame, ["target"], "preset")
    return name, frame["target"], formula_map.get(name, name), None, None, _lookup_strategy_row(strategy_df, name)


def _normalize_series_frame(frame: pd.DataFrame, start_date: pd.Timestamp | None = None) -> pd.DataFrame:
    filtered = frame.copy()
    if start_date is not None:
        filtered = filtered.loc[filtered.index >= pd.Timestamp(start_date)]

    normalized = pd.DataFrame(index=filtered.index)
    for column in filtered.columns:
        clean = filtered[column].dropna()
        if clean.empty:
            continue
        base = float(clean.iloc[0])
        if pd.isna(base) or abs(base) < 1e-12:
            normalized[column] = filtered[column]
        else:
            normalized[column] = filtered[column] / base * 100
    return normalized.dropna(how="all")


def _render_driver_decomposition_tab(formula_data: pd.DataFrame, strategy_row: pd.Series | None) -> None:
    if strategy_row is None:
        st.info("驱动拆解页仅对预设组合开放。")
        return

    package = build_driver_package(formula_data, strategy_row)
    if package is None:
        st.info("当前预设组合暂未配置驱动拆解模板。")
        return

    st.subheader(f"{package.target_label} 驱动拆解")
    st.caption(package.formula)
    st.markdown(
        "这一页回答三个问题："
        "当前价差主要由哪些底层因子组成；"
        "最近一段时间是谁在推动变化；"
        "如果某个驱动继续波动，目标序列大概会怎么变。"
    )

    target_clean = package.target_series.dropna()
    if target_clean.empty:
        st.warning("当前策略暂无可用于驱动拆解的数据。")
        return

    contribution_20d = decompose_change(package, window=20)
    non_residual = contribution_20d[contribution_20d["component"] != "residual"] if not contribution_20d.empty else contribution_20d
    main_driver = "N/A"
    if not non_residual.empty:
        main_driver = str(non_residual.iloc[non_residual["contribution"].abs().idxmax()]["label"])

    summary = st.columns(5)
    summary[0].metric("当前值", _format_metric(float(target_clean.iloc[-1])))
    summary[1].metric("5日变化", _format_metric(float(target_clean.iloc[-1] - target_clean.iloc[-6])) if len(target_clean) >= 6 else "N/A")
    summary[2].metric("20日变化", _format_metric(float(target_clean.iloc[-1] - target_clean.iloc[-21])) if len(target_clean) >= 21 else "N/A")
    summary[3].metric("主导因子", main_driver)
    summary[4].metric("样本区间", f"{target_clean.index.min().date()} 至 {target_clean.index.max().date()}")

    with st.expander("怎么看这一页", expanded=False):
        st.markdown(
            "`右侧合成项`：公式里被减掉或被比较的那一侧，通常代表进口成本、原料成本或对价腿。\n\n"
            "`驱动路径标准化`：把各条序列都换算成起点=100，只看相对变化速度和方向，不看绝对价格高低。\n\n"
            "`贡献拆解`：回答近 5/20/60 日的总变化里，分别有多少是由国内端、海外端、汇率、税费等驱动造成的。\n\n"
            "`组件编码`：内部计算用的字段键值，用来区分同名项；真正看研究时优先看 `驱动项` 列。\n\n"
            "`驱动诊断`：看每个驱动当前处在历史什么位置，包括分位和 Z-Score。\n\n"
            "`敏感度`：假设某个驱动单独变动 1%，目标序列会理论变化多少。\n\n"
            "`情景分析`：把某个驱动上调或下调 5%，估算目标值会落到哪里。"
        )

    current_rows: list[dict[str, object]] = []
    for component in package.components + package.derived_components:
        clean = component.series.dropna()
        if clean.empty:
            continue
        role = "基础驱动"
        if component in package.derived_components:
            role = "派生结果"
        if component.key == "rhs_total" or "import" in component.key:
            role = "右侧合成项"
        current_rows.append({"序列": component.label, "类型": role, "当前值": float(clean.iloc[-1])})
    if current_rows:
        st.markdown("#### 当前组件快照")
        st.caption("先看基础驱动，再看派生结果。`右侧合成项`通常可以理解为成本侧或对价侧。")
        st.dataframe(pd.DataFrame(current_rows), use_container_width=True, hide_index=True)

    plot_frame = pd.DataFrame(index=package.target_series.index)
    plot_frame[package.target_label] = package.target_series
    for component in package.derived_components[:2]:
        plot_frame[component.label] = component.series
    if not package.derived_components:
        for component in package.components[:3]:
            plot_frame[component.label] = component.series
    plot_frame = plot_frame.dropna(how="all")
    normalized_source = plot_frame.dropna(how="all")
    if not normalized_source.empty:
        available_dates = normalized_source.index.date.tolist()
        default_start = available_dates[max(0, len(available_dates) - min(len(available_dates), 252))]
        selected_start = st.date_input(
            "标准化起点",
            value=default_start,
            min_value=available_dates[0],
            max_value=available_dates[-1],
            key=f"normalize_start_{package.strategy_name}",
        )
        normalized = _normalize_series_frame(normalized_source, pd.Timestamp(selected_start))
    else:
        normalized = pd.DataFrame()

    if not normalized.empty:
        st.markdown("#### 驱动路径标准化")
        st.caption("把所有序列从你选定的起点统一成 100 后，就能看出谁涨得更快、谁回落更明显，以及目标序列和成本侧是否出现背离。")
        st.plotly_chart(px.line(normalized, title="驱动路径标准化对比（基准=100）"), use_container_width=True)

    st.markdown("#### 贡献拆解")
    valid_frame = pd.DataFrame({component.key: component.series for component in package.components}, index=package.target_series.index)
    valid_frame["target"] = package.target_series
    valid_frame = valid_frame.dropna()
    if valid_frame.empty:
        contribution = pd.DataFrame()
        selected_start_date = None
        selected_end_date = None
    else:
        valid_dates = valid_frame.index.date.tolist()
        default_end = valid_dates[-1]
        default_start = valid_dates[max(0, len(valid_dates) - min(len(valid_dates), 60))]
        date_cols = st.columns(2)
        selected_start_date = date_cols[0].date_input(
            "分析开始日期",
            value=default_start,
            min_value=valid_dates[0],
            max_value=valid_dates[-1],
            key=f"driver_range_start_{package.strategy_name}",
        )
        selected_end_date = date_cols[1].date_input(
            "分析结束日期",
            value=default_end,
            min_value=valid_dates[0],
            max_value=valid_dates[-1],
            key=f"driver_range_end_{package.strategy_name}",
        )
        if selected_start_date > selected_end_date:
            st.warning("分析开始日期不能晚于结束日期。")
            contribution = pd.DataFrame()
        else:
            contribution = decompose_change_between_dates(package, pd.Timestamp(selected_start_date), pd.Timestamp(selected_end_date))
    if contribution.empty:
        st.info("当前选择区间样本不足，暂时无法计算贡献拆解。")
    else:
        total_change = float(contribution.attrs.get("total_change", float("nan")))
        st.caption(f"区间 {selected_start_date} 至 {selected_end_date} 总变化：{_format_metric(total_change)}")
        st.markdown(
            "这部分是在问：如果把你选定区间内的变化拆开看，国内端、海外端、汇率、税费这些因素，各自大概贡献了多少。"
            "正值表示推高目标序列，负值表示压低目标序列。"
        )
        contribution_chart = contribution.copy()
        figure = go.Figure(
            go.Bar(
                x=contribution_chart["label"],
                y=contribution_chart["contribution"],
                marker_color=["#1f77b4" if value >= 0 else "#d62728" for value in contribution_chart["contribution"]],
            )
        )
        figure.update_layout(title="区间驱动贡献", xaxis_title="", yaxis_title="贡献值")
        st.plotly_chart(figure, use_container_width=True)

        contribution_abs = contribution_chart.copy()
        contribution_abs["abs_contribution"] = contribution_abs["contribution"].abs()
        contribution_abs = contribution_abs[contribution_abs["abs_contribution"] > 0]
        if not contribution_abs.empty:
            donut = px.pie(
                contribution_abs,
                names="label",
                values="abs_contribution",
                hole=0.55,
                title="驱动影响占比（按绝对值）",
            )
            st.plotly_chart(donut, use_container_width=True)

        contribution_display = contribution.copy()
        contribution_display["pct_of_total"] = contribution_display["pct_of_total"] * 100
        contribution_display = contribution_display.rename(
            columns={"component": "组件编码", "label": "驱动项", "contribution": "贡献值", "pct_of_total": "贡献占比(%)"}
        )
        st.caption(
            "`组件编码` 是内部识别字段；`驱动项` 是实际看的名称；`贡献值` 是该因子对近阶段总变化的估算贡献；"
            "`贡献占比` 是它占总变化的比例。"
        )
        st.dataframe(contribution_display, use_container_width=True, hide_index=True)

    diagnostics = build_driver_diagnostics(package)
    if not diagnostics.empty:
        diagnostics = diagnostics.rename(columns={"series": "序列", "current": "当前值"})
        st.markdown("#### 驱动诊断")
        st.caption(
            "`pct_252 / pct_756 / pct_1260` 分别代表大约 1 年、3 年、5 年历史分位；"
            "`z_60` 代表相对近 60 个交易日均值偏离了多少个标准差。"
        )
        st.dataframe(diagnostics, use_container_width=True, hide_index=True)

    lower_left, lower_right = st.columns(2)
    with lower_left:
        sensitivity = compute_factor_sensitivity(package)
        st.markdown("#### 敏感度（+1%）")
        st.caption("假设某个驱动单独上升 1%，其他都不变，目标序列理论上会变化多少。适合判断当前最该盯哪个变量。")
        if sensitivity.empty:
            st.info("当前策略暂时无法计算敏感度。")
        else:
            sensitivity = sensitivity.rename(
                columns={
                    "component": "组件编码",
                    "label": "驱动项",
                    "base_value": "当前值",
                    "bump_pct": "冲击幅度(%)",
                    "target_change": "目标变化",
                    "target_change_pct": "目标变化率(%)",
                }
            )
            st.dataframe(sensitivity, use_container_width=True, hide_index=True)
    with lower_right:
        scenarios = run_driver_scenarios(package)
        st.markdown("#### 情景分析（+/-5%）")
        st.caption("把某个驱动上调或下调 5%，观察目标值会落到哪里，用于做路径推演和交易预案。")
        if scenarios.empty:
            st.info("当前策略暂时无法计算情景分析。")
        else:
            scenarios = scenarios.rename(columns={"scenario": "情景", "target_value": "目标值", "target_change": "目标变化"})
            st.dataframe(scenarios, use_container_width=True, hide_index=True)


def _select_analysis_target_v2(
    sources: dict[str, pd.DataFrame],
    portfolios: pd.DataFrame,
    strategy_df: pd.DataFrame,
) -> tuple[str, pd.Series, str, pd.DataFrame | None, pd.DataFrame | None, pd.Series | None]:
    mode = st.sidebar.radio("分析对象", ["单品种", "月差", "跨表/价差组合", "预设组合"], key="analysis_mode_v2")

    if mode == "单品种":
        source_key = st.sidebar.selectbox("数据源", ["wind", "manual"], format_func=lambda x: SOURCE_LABELS[x], key="single_source_v2")
        source_df = sources[source_key]
        column = _grouped_selectbox("选择", source_df.columns.tolist(), "single_v2")
        frame = source_df[[column]].rename(columns={column: "target"})
        frame, _ = _apply_date_filter(frame, ["target"], "single_v2")
        return f"{SOURCE_LABELS[source_key]}:{column}", frame["target"], column, None, None, None

    if mode == "月差":
        source_key = st.sidebar.selectbox("数据源", ["wind", "manual"], format_func=lambda x: SOURCE_LABELS[x], key="month_spread_source_v2")
        target_name, frame, formula, coverage = _build_month_spread_target(sources[source_key])
        if frame.empty or "target" not in frame.columns:
            return target_name, pd.Series(dtype=float), formula, frame, coverage, None
        required_cols = [column for column in frame.columns if column != "target"] + ["target"]
        filtered_frame, _ = _apply_date_filter(frame, required_cols, "month_spread_v2")
        return target_name, filtered_frame["target"], formula, frame, coverage, None

    if mode == "跨表/价差组合":
        combo_mode = st.sidebar.selectbox("组合方式", ["价差组合", "比值组合"], key="combo_mode_v2")
        terms = _collect_term_configs(sources, combo_mode)
        combo_frame, default_name, formula, required_cols = _build_combo_frame(sources, combo_mode, terms)
        filtered_frame, _ = _apply_date_filter(combo_frame, required_cols, "combo_v2")
        coverage = _build_coverage_table(combo_frame, required_cols)
        custom_name = st.sidebar.text_input("组合名称", value="", key="combo_name_v2")
        return custom_name.strip() or default_name, filtered_frame["target"], formula, combo_frame, coverage, None

    if portfolios.empty:
        fallback = sources["wind"].columns[0]
        return fallback, sources["wind"][fallback], fallback, None, None, None

    name = _grouped_strategy_selectbox(strategy_df, "preset_v2")
    formula_map = dict(zip(strategy_df["StrategyName"], strategy_df["Formula"]))
    frame = portfolios[[name]].rename(columns={name: "target"})
    frame, _ = _apply_date_filter(frame, ["target"], "preset_v2")
    return name, frame["target"], formula_map.get(name, name), None, None, _lookup_strategy_row(strategy_df, name)


def main() -> None:
    st.set_page_config(page_title="Wind 看板", layout="wide")
    st.title("Wind 与手工价格联合分析看板")
    st.caption("支持 Wind sheet、manual sheet、2-3 项价差/价格比组合、重叠区间分析，以及按 USDCHY 参与计算。")

    workbook_path = _sidebar_excel_path()
    workbook = Path(workbook_path)

    if st.button("刷新 Excel 数据"):
        ok = refresh_excel_workbook(workbook, APP_CONFIG["excel"].get("refresh_timeout_sec", 180))
        load_all_data.clear()
        if ok:
            st.success("Excel 已刷新并重新加载。")
        else:
            st.warning("Excel 自动刷新失败或被跳过，请检查本机 Excel/pywin32 环境。")

    try:
        sources, portfolios, strategy_df, formula_data = load_all_data(workbook_path)
    except Exception as exc:
        st.error(f"数据加载失败: {exc}")
        logger.exception("数据加载失败")
        return

    target_name, series, formula, combo_frame, coverage, selected_strategy_row = _select_analysis_target_v2(sources, portfolios, strategy_df)
    series = series.dropna()
    if series.empty:
        st.warning("当前选择下没有可分析的数据。")
        return

    controls = _risk_controls(series)
    st.sidebar.markdown("### 当前定义")
    st.sidebar.code(formula)

    overview_tab, risk_tab, seasonal_tab, data_tab = st.tabs(["总览", "风控分析", "季节性", "数据浏览"])

    with overview_tab:
        st.subheader(target_name)
        st.caption(f"数据文件: {workbook_path}")
        st.caption(f"样本区间: {series.index.min().date()} 至 {series.index.max().date()}")
        if coverage is not None:
            st.markdown("#### 可用价格区间")
            st.dataframe(coverage, use_container_width=True, hide_index=True)
        _render_risk_section(series, controls)
        st.plotly_chart(px.line(series, title=f"{target_name} 历史走势"), use_container_width=True)
        st.plotly_chart(px.histogram(series, nbins=50, title=f"{target_name} 数值分布"), use_container_width=True)

    with risk_tab:
        st.subheader(f"{target_name} 风控细节")
        if coverage is not None:
            st.markdown("#### 各项有效区间")
            st.dataframe(coverage, use_container_width=True, hide_index=True)
        _render_risk_section(series, controls)

    with seasonal_tab:
        st.subheader(f"{target_name} 季节性分析")
        _render_seasonality_section(series)

    with data_tab:
        st.subheader("Wind 数据预览")
        st.dataframe(sources["wind"].tail(200), use_container_width=True)
        st.subheader("Manual 数据预览")
        st.dataframe(sources["manual"].tail(200), use_container_width=True)
        if not portfolios.empty:
            st.subheader("预设组合预览")
            st.dataframe(portfolios.tail(200), use_container_width=True)
        if combo_frame is not None:
            st.subheader("当前组合对齐数据")
            st.dataframe(combo_frame.tail(200), use_container_width=True)
        st.subheader("当前分析序列")
        st.dataframe(series.tail(200).rename(target_name).to_frame(), use_container_width=True)


def run_dashboard_app() -> None:
    st.set_page_config(page_title="交易研究看板", layout="wide")
    st.title("交易研究看板")
    st.caption("面向日度时间序列研究，支持单品种、手工组合、预设组合和驱动拆解。")

    workbook_path = _sidebar_excel_path()
    workbook = Path(workbook_path)

    if st.button("刷新 Excel 工作簿"):
        ok = refresh_excel_workbook(workbook, APP_CONFIG["excel"].get("refresh_timeout_sec", 180))
        load_all_data.clear()
        if ok:
            st.success("Excel 工作簿刷新成功。")
        else:
            st.warning("Excel 刷新失败，请检查工作簿路径和本机 Excel 环境。")

    try:
        sources, portfolios, strategy_df, formula_data = load_all_data(workbook_path)
    except Exception as exc:
        st.error(f"读取工作簿数据失败：{exc}")
        logger.exception("Unable to load workbook data")
        return

    target_name, series, formula, combo_frame, coverage, selected_strategy_row = _select_analysis_target_v2(sources, portfolios, strategy_df)
    series = series.dropna()
    if series.empty:
        st.warning("当前选择没有可用数据。")
        return

    controls = _risk_controls(series)
    st.sidebar.markdown("### 当前公式")
    st.sidebar.code(formula)
    _render_metric_explainer()

    overview_tab, risk_tab, seasonal_tab, driver_tab, data_tab = st.tabs(
        ["总览", "风控分析", "季节性", "驱动拆解", "数据浏览"]
    )

    with overview_tab:
        st.subheader(target_name)
        _render_metric_explainer()
        st.caption(f"工作簿：{workbook_path}")
        st.caption(f"数据区间：{series.index.min().date()} 至 {series.index.max().date()}")
        if coverage is not None:
            st.markdown("#### 覆盖区间")
            st.dataframe(coverage, use_container_width=True, hide_index=True)
        _render_risk_section(series, controls)
        st.plotly_chart(px.line(series, title=f"{target_name} 时间序列"), use_container_width=True)
        st.plotly_chart(px.histogram(series, nbins=50, title=f"{target_name} 分布"), use_container_width=True)

    with risk_tab:
        st.subheader(f"{target_name} 风控分析")
        if coverage is not None:
            st.markdown("#### 覆盖区间")
            st.dataframe(coverage, use_container_width=True, hide_index=True)
        _render_risk_section(series, controls)

    with seasonal_tab:
        st.subheader(f"{target_name} 季节性")
        _render_seasonality_section_v2(series)

    with driver_tab:
        _render_driver_decomposition_tab(formula_data, selected_strategy_row)

    with data_tab:
        st.subheader("Wind 数据")
        st.dataframe(sources["wind"].tail(200), use_container_width=True)
        st.subheader("Manual 数据")
        st.dataframe(sources["manual"].tail(200), use_container_width=True)
        if not portfolios.empty:
            st.subheader("预设组合数据")
            st.dataframe(portfolios.tail(200), use_container_width=True)
        if combo_frame is not None:
            st.subheader("自定义组合中间表")
            st.dataframe(combo_frame.tail(200), use_container_width=True)
        st.subheader("当前选择序列")
        st.dataframe(series.tail(200).rename(target_name).to_frame(), use_container_width=True)


if __name__ == "__main__":
    run_dashboard_app()
