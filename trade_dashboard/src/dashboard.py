from __future__ import annotations

import logging
import re
import sys
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

try:
    from .data_loader import load_strategy_table, load_timeseries_from_excel
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
    from src.data_loader import load_strategy_table, load_timeseries_from_excel
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


@st.cache_data(show_spinner=False)
def load_all_data(workbook_path: str) -> tuple[pd.DataFrame, pd.DataFrame]:
    excel_cfg = APP_CONFIG["excel"]
    workbook = Path(workbook_path)
    data = load_timeseries_from_excel(
        workbook,
        excel_cfg["data_sheet"],
        excel_cfg["date_column"],
        excel_cfg.get("header_rows", 0),
        excel_cfg.get("column_name_row", 0),
    )

    strategy_sheet = excel_cfg.get("strategy_sheet")
    if strategy_sheet:
        strategy_df = load_strategy_table(workbook, strategy_sheet)
    else:
        strategy_cfg = load_yaml(BASE_DIR / "config" / "strategy.yaml")
        strategy_df = pd.DataFrame(strategy_cfg.get("strategies", []))
        strategy_df = strategy_df.rename(columns={"name": "StrategyName", "formula": "Formula", "enabled": "Enabled"})

    portfolios = build_portfolios(data, strategy_df) if not strategy_df.empty else pd.DataFrame(index=data.index)
    return data, portfolios


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


def _grouped_selectbox(label_prefix: str, columns: list[str], key_prefix: str, exclude: str | None = None) -> str:
    groups = _series_groups(columns)
    group_names = list(groups.keys())

    default_group = group_names[0]
    if exclude:
        for group_name, members in groups.items():
            if any(item != exclude for item in members):
                default_group = group_name
                break

    group = st.sidebar.selectbox(f"{label_prefix}类别", group_names, key=f"{key_prefix}_group", index=group_names.index(default_group))
    choices = [item for item in groups[group] if item != exclude]
    if not choices:
        choices = columns
    return st.sidebar.selectbox(f"{label_prefix}合约", choices, key=f"{key_prefix}_item")


def _build_ratio_series(
    data: pd.DataFrame,
    left_col: str,
    right_col: str,
    left_weight: float,
    right_weight: float,
    mode: str,
) -> tuple[str, pd.Series, str]:
    left_expr = f"{left_weight:g} * {left_col}" if left_weight != 1 else left_col
    right_expr = f"{right_weight:g} * {right_col}" if right_weight != 1 else right_col

    if mode == "价格比":
        denominator = (data[right_col] * right_weight).replace(0, pd.NA)
        series = (data[left_col] * left_weight) / denominator
        formula = f"({left_expr}) / ({right_expr})"
        suffix = "ratio"
    else:
        series = (data[left_col] * left_weight) - (data[right_col] * right_weight)
        formula = f"({left_expr}) - ({right_expr})"
        suffix = "spread"

    name = f"{left_col}_{right_col}_{suffix}"
    return name, series.rename(name), formula


def _render_metric_table(title: str, data: dict[str, float], style: str = "number") -> None:
    frame = pd.DataFrame(
        {
            "指标": list(data.keys()),
            "数值": [_format_metric(value, style=style) for value in data.values()],
        }
    )
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
    st.plotly_chart(px.line(matrix_plot, title="历年季节性曲线（连续化）"), use_container_width=True)

    mean = matrix.mean(axis=1)
    std = matrix.std(axis=1)
    band = pd.DataFrame({"均值": mean, "+1σ": mean + std, "-1σ": mean - std}, index=continuous_index)
    st.plotly_chart(px.line(band, title="季节性均值带"), use_container_width=True)

    monthly = seasonal_source.to_frame("value")
    monthly["month"] = monthly.index.month
    st.plotly_chart(px.box(monthly, x="month", y="value", title="月度分布"), use_container_width=True)

    s_metrics = seasonal_stats(seasonal_source, years)
    cols = st.columns(2)
    cols[0].metric("当前值同期百分位", _format_metric(s_metrics["seasonal_percentile"], style="percentile"))
    cols[1].metric("当前值同期偏离", _format_metric(s_metrics["seasonal_deviation"]))


def _select_analysis_series(raw_data: pd.DataFrame, portfolios: pd.DataFrame) -> tuple[str, pd.Series, str]:
    mode = st.sidebar.radio("分析对象", ["单品种", "价格比组合", "预设组合"])
    raw_columns = raw_data.columns.tolist()

    if mode == "单品种":
        column = _grouped_selectbox("选择", raw_columns, "single")
        return column, raw_data[column], column

    if mode == "价格比组合":
        combo_mode = st.sidebar.selectbox("组合方式", ["价格比", "价差"])
        left_col = _grouped_selectbox("左腿", raw_columns, "left")
        right_col = _grouped_selectbox("右腿", raw_columns, "right", exclude=left_col)
        left_weight = st.sidebar.number_input("左腿系数", value=1.0, step=0.1, format="%.2f")
        right_weight = st.sidebar.number_input("右腿系数", value=1.0, step=0.1, format="%.2f")
        custom_name = st.sidebar.text_input("组合名称", value="")
        default_name, series, formula = _build_ratio_series(raw_data, left_col, right_col, left_weight, right_weight, combo_mode)
        return custom_name.strip() or default_name, series, formula

    if portfolios.empty:
        st.sidebar.warning("当前没有可用的预设组合，已切回单品种模式。")
        column = _grouped_selectbox("选择", raw_columns, "fallback")
        return column, raw_data[column], column

    portfolio_name = st.sidebar.selectbox("选择预设组合", portfolios.columns.tolist())
    formula_lookup = load_yaml(BASE_DIR / "config" / "strategy.yaml").get("strategies", [])
    formula_map = {item.get("name"): item.get("formula", "") for item in formula_lookup}
    return portfolio_name, portfolios[portfolio_name], formula_map.get(portfolio_name, portfolio_name)


def _sidebar_excel_path() -> str:
    default_path = APP_CONFIG["excel"]["workbook_path"]
    if "excel_path" not in st.session_state:
        st.session_state["excel_path"] = default_path

    st.sidebar.markdown("### Excel 数据源")
    excel_path = st.sidebar.text_input("Excel 路径", value=st.session_state["excel_path"])
    st.session_state["excel_path"] = excel_path.strip() or default_path
    return st.session_state["excel_path"]


def main() -> None:
    st.set_page_config(page_title="Wind 看板", layout="wide")
    st.title("Wind 品种与组合分析看板")
    st.caption("支持单品种、临时价格比/价差组合、预设组合，以及可调风控与季节性分析。")

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
        raw_data, portfolios = load_all_data(workbook_path)
    except Exception as exc:
        st.error(f"数据加载失败: {exc}")
        logger.exception("数据加载失败")
        return

    target_name, series, formula = _select_analysis_series(raw_data, portfolios)
    series = series.dropna()
    if series.empty:
        st.warning("当前选择没有可用数据。")
        return

    controls = _risk_controls(series)

    st.sidebar.markdown("### 当前定义")
    st.sidebar.code(formula)

    overview_tab, risk_tab, seasonal_tab, data_tab = st.tabs(["总览", "风控分析", "季节性", "数据浏览"])

    with overview_tab:
        st.subheader(target_name)
        st.caption(f"数据文件: {workbook_path}")
        st.caption(f"样本区间: {series.index.min().date()} 至 {series.index.max().date()}")
        _render_risk_section(series, controls)
        st.plotly_chart(px.line(series, title=f"{target_name} 历史走势"), use_container_width=True)
        st.plotly_chart(px.histogram(series, nbins=50, title=f"{target_name} 数值分布"), use_container_width=True)

    with risk_tab:
        st.subheader(f"{target_name} 风控细节")
        _render_risk_section(series, controls)
        returns = series.pct_change().dropna() if (series > 0).all() else series.diff().dropna()
        if not returns.empty:
            st.plotly_chart(px.histogram(returns, nbins=60, title="收益/变动分布"), use_container_width=True)
            rolling_rank = series.rank(pct=True) * 100
            st.plotly_chart(px.line(rolling_rank, title="全样本历史百分位路径"), use_container_width=True)

    with seasonal_tab:
        st.subheader(f"{target_name} 季节性分析")
        _render_seasonality_section(series)

    with data_tab:
        st.subheader("原始数据预览")
        st.dataframe(raw_data.tail(200), use_container_width=True)
        if not portfolios.empty:
            st.subheader("预设组合预览")
            st.dataframe(portfolios.tail(200), use_container_width=True)
        st.subheader("当前分析序列")
        st.dataframe(series.tail(200).rename(target_name).to_frame(), use_container_width=True)


if __name__ == "__main__":
    main()
