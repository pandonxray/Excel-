from __future__ import annotations

import logging
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

from .data_loader import load_strategy_table, load_timeseries_from_excel
from .excel_refresh import refresh_excel_workbook
from .portfolio_engine import build_portfolios
from .risk_engine import summarize_risk_metrics
from .seasonal_engine import remove_feb29, seasonal_matrix, seasonal_stats
from .utils import load_yaml, setup_logging


BASE_DIR = Path(__file__).resolve().parents[1]
APP_CONFIG = load_yaml(BASE_DIR / "config" / "app.yaml")
setup_logging(APP_CONFIG.get("logging", {}).get("level", "INFO"), APP_CONFIG.get("logging", {}).get("file"))
logger = logging.getLogger(__name__)


def load_all_data() -> tuple[pd.DataFrame, pd.DataFrame]:
    excel_cfg = APP_CONFIG["excel"]
    workbook = BASE_DIR / excel_cfg["workbook_path"]
    data = load_timeseries_from_excel(workbook, excel_cfg["data_sheet"], excel_cfg["date_column"])

    strategy_sheet = excel_cfg.get("strategy_sheet")
    if strategy_sheet:
        strategy_df = load_strategy_table(workbook, strategy_sheet)
    else:
        strategy_cfg = load_yaml(BASE_DIR / "config" / "strategy.yaml")
        strategy_df = pd.DataFrame(strategy_cfg.get("strategies", []))
        strategy_df = strategy_df.rename(columns={"name": "StrategyName", "formula": "Formula", "enabled": "Enabled"})

    portfolios = build_portfolios(data, strategy_df)
    return data, portfolios


def main() -> None:
    st.set_page_config(page_title="组合分析与风控看板", layout="wide")
    st.title("中文多品种组合分析与风控看板")

    excel_cfg = APP_CONFIG["excel"]
    workbook = BASE_DIR / excel_cfg["workbook_path"]

    col_a, col_b = st.columns([1, 3])
    with col_a:
        if st.button("刷新 Excel 数据"):
            ok = refresh_excel_workbook(workbook, excel_cfg.get("refresh_timeout_sec", 180))
            st.success("刷新完成") if ok else st.warning("刷新失败或跳过，请检查日志")

    try:
        raw_data, portfolios = load_all_data()
    except Exception as exc:
        st.error(f"数据加载失败: {exc}")
        logger.exception("加载失败")
        return

    if APP_CONFIG.get("analysis", {}).get("remove_feb29", True):
        portfolios = remove_feb29(portfolios)

    page = st.sidebar.radio("页面", ["总览页", "组合分析页", "季节分析页", "数据浏览页"])
    if portfolios.empty:
        st.warning("没有可用组合，请检查策略配置。")
        return

    selected = st.sidebar.selectbox("选择组合", portfolios.columns.tolist())
    series = portfolios[selected]

    if page == "总览页":
        metrics = summarize_risk_metrics(series)
        cards = st.columns(4)
        for idx, (k, v) in enumerate(metrics.items()):
            cards[idx % 4].metric(k, f"{v:.4f}" if pd.notna(v) else "N/A")
        st.plotly_chart(px.line(series, title=f"{selected} 历史走势"), use_container_width=True)

    elif page == "组合分析页":
        st.plotly_chart(px.line(series, title=f"{selected} 走势"), use_container_width=True)
        st.plotly_chart(px.histogram(series.dropna(), nbins=40, title="分布图"), use_container_width=True)

    elif page == "季节分析页":
        years = st.slider("季节回看年数", 3, 10, APP_CONFIG.get("analysis", {}).get("seasonal_years", 5))
        matrix = seasonal_matrix(series, years)
        if matrix.empty:
            st.info("季节数据不足")
            return

        st.plotly_chart(px.line(matrix, title=f"{selected} 历年季节图"), use_container_width=True)
        mean = matrix.mean(axis=1)
        std = matrix.std(axis=1)
        band = pd.DataFrame({"均值": mean, "+1σ": mean + std, "-1σ": mean - std})
        st.plotly_chart(px.line(band, title="季节均值带"), use_container_width=True)
        monthly = series.dropna().to_frame("value")
        monthly["month"] = monthly.index.month
        st.plotly_chart(px.box(monthly, x="month", y="value", title="月度箱线图"), use_container_width=True)
        s_metrics = seasonal_stats(series, years)
        col1, col2 = st.columns(2)
        col1.metric("当前值同期分位", f"{s_metrics['同期分位']:.2f}%" if pd.notna(s_metrics["同期分位"]) else "N/A")
        col2.metric("当前值同期偏离", f"{s_metrics['同期偏离']:.4f}" if pd.notna(s_metrics["同期偏离"]) else "N/A")

    else:
        st.subheader("原始数据浏览")
        st.dataframe(raw_data.tail(200), use_container_width=True)
        st.subheader("组合数据浏览")
        st.dataframe(portfolios.tail(200), use_container_width=True)


if __name__ == "__main__":
    main()
