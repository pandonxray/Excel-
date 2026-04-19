from __future__ import annotations

import logging
import re
import sys
import tempfile
from itertools import count
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

if not hasattr(np, "unicode_"):
    np.unicode_ = np.str_

_PLOTLY_KEY_COUNTER = count()

try:
    from .basis_engine import build_basis_tables
    from .data_loader import load_timeseries_from_excel
    from .driver_engine import (
        build_driver_diagnostics,
        build_driver_package,
        compute_factor_sensitivity,
        decompose_change_between_dates,
        run_driver_scenarios,
    )
    from .excel_refresh import refresh_excel_workbook
    from .industry_engine import build_propylene_profit_dashboard
    from .portfolio_engine import build_portfolios
    from .risk_engine import build_risk_report, percentile_of_value, var_es_over_window, zscore_of_value
    from .seasonal_engine import remove_feb29, seasonal_matrix, seasonal_stats
    from .utils import load_yaml, setup_logging
except ImportError:
    from src.basis_engine import build_basis_tables
    from src.data_loader import load_timeseries_from_excel
    from src.driver_engine import (
        build_driver_diagnostics,
        build_driver_package,
        compute_factor_sensitivity,
        decompose_change_between_dates,
        run_driver_scenarios,
    )
    from src.excel_refresh import refresh_excel_workbook
    from src.industry_engine import build_propylene_profit_dashboard
    from src.portfolio_engine import build_portfolios
    from src.risk_engine import build_risk_report, percentile_of_value, var_es_over_window, zscore_of_value
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

SOURCE_LABELS = {
    "wind": "Wind期货",
    "wind_continue": "连续/活跃合约",
    "manual": "Manual外盘",
    "spot": "现货产业链",
    "basis": "基差",
    "downstream": "下游利润",
}
SOURCE_COLORS = {
    "wind": "#6f8fa8",
    "wind_continue": "#97a7c3",
    "manual": "#87a9b6",
    "spot": "#b88968",
    "basis": "#8aa08d",
    "downstream": "#7aa38f",
}
PLOT_TEMPLATE = "plotly_white"


def _inject_theme() -> None:
    st.markdown(
        """
        <style>
        :root {
            --bg: #f5f1ea;
            --bg-elevated: #fbf8f2;
            --bg-muted: #ebe4da;
            --panel: rgba(252, 249, 244, 0.92);
            --panel-soft: rgba(247, 242, 235, 0.88);
            --line: rgba(88, 103, 122, 0.14);
            --text: #243446;
            --muted: #6d7c8e;
            --accent: #809ab2;
            --accent-soft: rgba(128, 154, 178, 0.16);
            --good: #7aa38f;
        }
        .stApp {
            background:
                radial-gradient(circle at top left, rgba(128, 154, 178, 0.18), transparent 26%),
                radial-gradient(circle at top right, rgba(122, 163, 143, 0.12), transparent 24%),
                linear-gradient(180deg, #f8f5f0 0%, #f1ece4 100%);
            color: var(--text);
            font-family: "Aptos", "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif;
        }
        .block-container {
            padding-top: 1.25rem;
            padding-bottom: 2rem;
            max-width: 1500px;
        }
        [data-testid="stSidebar"] {
            background:
                linear-gradient(180deg, rgba(247, 242, 235, 0.98) 0%, rgba(241, 235, 227, 0.98) 100%);
            border-right: 1px solid var(--line);
        }
        [data-testid="stSidebar"] > div:first-child {
            background: transparent;
        }
        [data-testid="stSidebar"] * {
            color: var(--text);
        }
        [data-testid="stSidebar"] [data-baseweb="select"] > div,
        [data-testid="stSidebar"] [data-baseweb="input"] > div,
        [data-testid="stSidebar"] [data-baseweb="textarea"] > div,
        [data-testid="stSidebar"] .stDateInput > div > div,
        [data-testid="stSidebar"] .stFileUploader > div {
            background: rgba(255, 255, 255, 0.58);
            border: 1px solid rgba(88, 103, 122, 0.12);
            border-radius: 14px;
        }
        [data-testid="stSidebar"] .stRadio label,
        [data-testid="stSidebar"] .stCheckbox label {
            color: var(--muted);
        }
        [data-testid="stSidebar"] .stButton button {
            background: linear-gradient(135deg, var(--accent) 0%, #9fb4c6 100%);
            color: #fffdf9;
            border: none;
            border-radius: 999px;
            font-weight: 700;
            min-height: 2.75rem;
        }
        .workspace-shell {
            padding: 1.25rem 1.4rem 1.4rem 1.4rem;
            border: 1px solid var(--line);
            background:
                linear-gradient(135deg, rgba(253, 250, 245, 0.96) 0%, rgba(245, 239, 232, 0.96) 100%);
            border-radius: 28px;
            box-shadow: 0 18px 40px rgba(108, 106, 99, 0.08);
            margin-bottom: 1.1rem;
        }
        .workspace-kicker {
            text-transform: uppercase;
            letter-spacing: 0.14em;
            font-size: 0.72rem;
            color: var(--accent);
            margin-bottom: 0.65rem;
        }
        .workspace-title {
            font-size: 2.15rem;
            line-height: 1.02;
            font-weight: 700;
            color: var(--text);
            margin: 0;
        }
        .workspace-note {
            margin-top: 0.9rem;
            max-width: 60rem;
            font-size: 0.98rem;
            line-height: 1.7;
            color: var(--muted);
        }
        .workspace-strip {
            display: grid;
            grid-template-columns: repeat(4, minmax(0, 1fr));
            gap: 0.8rem;
            margin-top: 1.2rem;
        }
        .workspace-stat {
            padding: 0.95rem 1rem;
            border-radius: 18px;
            background: rgba(255, 255, 255, 0.55);
            border: 1px solid rgba(88, 103, 122, 0.10);
        }
        .workspace-stat-label {
            font-size: 0.78rem;
            color: var(--muted);
            margin-bottom: 0.35rem;
        }
        .workspace-stat-value {
            font-size: 1rem;
            color: var(--text);
            font-weight: 600;
        }
        .hero-card {
            padding: 1.15rem 1.2rem 1.25rem 1.2rem;
            border-radius: 24px;
            background:
                linear-gradient(135deg, rgba(253, 250, 246, 0.92) 0%, rgba(247, 243, 238, 0.92) 100%);
            border: 1px solid var(--line);
            color: var(--text);
            margin-bottom: 0.9rem;
            box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.45);
        }
        .hero-grid {
            display: grid;
            grid-template-columns: minmax(0, 1.7fr) minmax(18rem, 0.9fr);
            gap: 1rem;
            align-items: start;
        }
        .hero-kicker {
            text-transform: uppercase;
            letter-spacing: 0.12em;
            font-size: 0.72rem;
            color: var(--accent);
            margin-bottom: 0.55rem;
        }
        .hero-title {
            font-size: 2.2rem;
            font-weight: 700;
            margin-bottom: 0.5rem;
            line-height: 1.02;
        }
        .hero-note {
            font-size: 0.96rem;
            line-height: 1.6;
            color: var(--muted);
        }
        .hero-meta {
            display: grid;
            gap: 0.75rem;
        }
        .hero-meta-block {
            padding-top: 0.05rem;
        }
        .hero-meta-label {
            font-size: 0.74rem;
            text-transform: uppercase;
            letter-spacing: 0.1em;
            color: var(--muted);
            margin-bottom: 0.28rem;
        }
        .hero-meta-value {
            color: var(--text);
            font-size: 0.95rem;
            line-height: 1.5;
            word-break: break-word;
        }
        div[data-testid="stMetric"] {
            background: var(--panel-soft);
            border: 1px solid var(--line);
            border-radius: 18px;
            padding: 0.95rem 1rem;
            box-shadow: 0 8px 18px rgba(108, 106, 99, 0.05);
        }
        div[data-testid="stMetric"] label,
        div[data-testid="stMetric"] [data-testid="stMetricLabel"] {
            color: var(--muted);
        }
        div[data-testid="stMetricValue"] {
            color: var(--text);
        }
        .section-chip {
            display: inline-block;
            padding: 0.22rem 0.7rem;
            border-radius: 999px;
            background: var(--accent-soft);
            color: var(--accent);
            font-weight: 600;
            font-size: 0.78rem;
            margin-bottom: 0.45rem;
        }
        .section-title {
            font-size: 1.2rem;
            color: var(--text);
            font-weight: 650;
            margin-bottom: 0.2rem;
        }
        .section-note {
            color: var(--muted);
            font-size: 0.92rem;
            margin-bottom: 0.9rem;
        }
        .section-divider {
            height: 1px;
            width: 100%;
            background: linear-gradient(90deg, rgba(211, 166, 90, 0.24) 0%, rgba(211, 166, 90, 0) 100%);
            margin: 0.55rem 0 1rem 0;
        }
        .formula-box {
            padding: 0.95rem 1rem;
            border: 1px solid var(--line);
            background: rgba(255, 255, 255, 0.55);
            border-radius: 18px;
            margin: 0.75rem 0 1rem 0;
            color: var(--text);
        }
        .formula-box strong {
            color: var(--accent);
        }
        .stTabs [data-baseweb="tab-list"] {
            gap: 0.45rem;
            border-bottom: 1px solid var(--line);
            margin-bottom: 1rem;
        }
        .stTabs [data-baseweb="tab"] {
            height: 2.55rem;
            border-radius: 999px;
            background: transparent;
            color: var(--muted);
            padding: 0 1rem;
        }
        .stTabs [aria-selected="true"] {
            background: rgba(255, 255, 255, 0.75) !important;
            color: var(--text) !important;
            border: 1px solid var(--line);
        }
        .stDataFrame, div[data-testid="stTable"] {
            border: 1px solid var(--line);
            border-radius: 18px;
            overflow: hidden;
        }
        div[data-testid="stPlotlyChart"] {
            border: 1px solid var(--line);
            border-radius: 22px;
            background: rgba(255, 255, 255, 0.55);
            padding: 0.4rem 0.55rem;
        }
        h1, h2, h3, label, p {
            color: var(--text);
        }
        @media (max-width: 980px) {
            .workspace-strip,
            .hero-grid {
                grid-template-columns: 1fr;
            }
            .workspace-title,
            .hero-title {
                font-size: 1.7rem;
            }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def _save_uploaded_excel(uploaded_file) -> str:
    suffix = Path(uploaded_file.name).suffix or ".xlsx"
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(uploaded_file.getbuffer())
        return tmp.name


@st.cache_data(show_spinner=False)
def load_all_data(
    workbook_path: str,
) -> tuple[dict[str, pd.DataFrame], pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
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
        excel_cfg["manual_sheet"],
        excel_cfg.get("manual_date_column", "price_date"),
        excel_cfg.get("manual_header_rows", 0),
        excel_cfg.get("manual_column_name_row", 0),
    )
    spot_df = load_timeseries_from_excel(
        workbook,
        excel_cfg["spot_sheet"],
        excel_cfg.get("spot_date_column", "指标名称"),
        excel_cfg.get("spot_header_rows", [0, 1, 2, 3]),
        excel_cfg.get("spot_column_name_row", 1),
    )
    wind_continue_df = load_timeseries_from_excel(
        workbook,
        excel_cfg.get("wind_continue_sheet", "wind_continue"),
        excel_cfg["date_column"],
        excel_cfg.get("header_rows", 0),
        excel_cfg.get("column_name_row", 0),
    )

    downstream_df, downstream_meta = build_propylene_profit_dashboard(spot_df)
    basis_formula_df, basis_display_df, basis_meta = build_basis_tables(
        wind_df,
        wind_continue_df,
        manual_df,
        spot_df,
        fx_column=excel_cfg.get("fx_column", "USDCHY"),
    )
    base_formula_cols = set(wind_df.columns) | set(wind_continue_df.columns) | set(manual_df.columns) | set(spot_df.columns)
    basis_helper_df = basis_formula_df.drop(columns=[col for col in basis_formula_df.columns if col in base_formula_cols], errors="ignore")
    merged_for_formula = (
        wind_df.join(wind_continue_df, how="outer")
        .join(manual_df, how="outer")
        .join(spot_df, how="outer")
        .join(basis_helper_df, how="outer")
        .sort_index()
    )

    strategy_cfg = load_yaml(BASE_DIR / "config" / "strategy.yaml")
    strategy_df = pd.DataFrame(strategy_cfg.get("strategies", []))
    if not strategy_df.empty:
        strategy_df = strategy_df.rename(
            columns={
                "name": "StrategyName",
                "formula": "Formula",
                "enabled": "Enabled",
                "category": "Category",
                "notes": "Notes",
            }
        )
        portfolios = build_portfolios(merged_for_formula, strategy_df)
    else:
        portfolios = pd.DataFrame(index=merged_for_formula.index)

    sources = {
        "wind": wind_df,
        "wind_continue": wind_continue_df,
        "manual": manual_df,
        "spot": spot_df,
        "basis": basis_display_df,
        "downstream": downstream_df,
    }
    return sources, portfolios, strategy_df, downstream_meta, basis_meta, basis_formula_df


def _format_metric(value: float | str, style: str = "number") -> str:
    if isinstance(value, str):
        return value
    if pd.isna(value):
        return "N/A"
    if style == "percentile":
        return f"{value:.2f}%"
    if style == "ratio_pct":
        return f"{value * 100:.2f}%"
    return f"{value:.2f}"


def _render_section_intro(chip: str, title: str, note: str = "") -> None:
    note_html = f"<div class='section-note'>{note}</div>" if note else ""
    st.markdown(
        f"""
        <div class="section-chip">{chip}</div>
        <div class="section-title">{title}</div>
        {note_html}
        <div class="section-divider"></div>
        """,
        unsafe_allow_html=True,
    )


def _series_summary(series: pd.Series) -> dict[str, str]:
    clean = series.dropna()
    if clean.empty:
        return {
            "current": "N/A",
            "daily_change": "N/A",
            "range": "N/A",
            "samples": "0",
        }

    current = clean.iloc[-1]
    previous = clean.iloc[-2] if len(clean) > 1 else np.nan
    daily_change = current - previous if pd.notna(previous) else np.nan
    start, end = clean.index.min(), clean.index.max()
    return {
        "current": _format_metric(current),
        "daily_change": _format_metric(daily_change),
        "range": f"{start.date().isoformat()} - {end.date().isoformat()}",
        "samples": f"{len(clean):,}",
    }


def _render_workspace_shell(workbook_path: str) -> None:
    workbook = Path(workbook_path)
    workbook_name = workbook.name or workbook_path
    workbook_folder = str(workbook.parent) if workbook.parent else workbook_path
    st.markdown(
        f"""
        <div class="workspace-shell">
            <div class="workspace-kicker">Trade Research Workspace</div>
            <div class="workspace-title">丙烯产业链研究工作台</div>
            <div class="workspace-note">
                统一查看期货、外盘、现货产业链、下游利润和自定义组合。这个界面现在更偏交易台，
                首先服务筛选、监控和判断，而不是展示一堆卡片。
            </div>
            <div class="workspace-strip">
                <div class="workspace-stat">
                    <div class="workspace-stat-label">当前工作簿</div>
                    <div class="workspace-stat-value">{workbook_name}</div>
                </div>
                <div class="workspace-stat">
                    <div class="workspace-stat-label">数据目录</div>
                    <div class="workspace-stat-value">{workbook_folder}</div>
                </div>
                <div class="workspace-stat">
                    <div class="workspace-stat-label">分析引擎</div>
                    <div class="workspace-stat-value">Wind / Continue / Manual / Spot / Basis / Downstream</div>
                </div>
                <div class="workspace-stat">
                    <div class="workspace-stat-label">工作方式</div>
                    <div class="workspace-stat-value">单序列、预设组合、自定义组合</div>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def _series_groups(source_key: str, columns: list[str], downstream_meta: pd.DataFrame | None = None) -> dict[str, list[str]]:
    if source_key == "downstream" and downstream_meta is not None and not downstream_meta.empty:
        grouped: dict[str, list[str]] = {}
        meta = downstream_meta[downstream_meta["metric"].isin(columns)]
        for _, row in meta.iterrows():
            category = str(row["category"])
            if category == "原始映射":
                continue
            grouped.setdefault(category, []).append(str(row["metric"]))
        return {group: sorted(items) for group, items in grouped.items()}

    if source_key == "spot":
        rules = {
            "PO链": ["环氧丙烷", "液氯", "双氧水", "聚醚"],
            "丙烯下游": ["丙烯酸", "丙烯腈", "正丁醇", "辛醇", "合成氨", "苯酚", "丙酮", "纯苯", "丙烯"],
            "甲醇链": ["甲醇"],
            "乙烯与汇率": ["乙烯", "汇率"],
            "PP粉料": ["PP粉", "PP：拉丝", "停-PP粉"],
        }
        grouped: dict[str, list[str]] = {name: [] for name in rules}
        grouped["其他"] = []
        for column in columns:
            placed = False
            for group, keywords in rules.items():
                if any(keyword in column for keyword in keywords):
                    grouped[group].append(column)
                    placed = True
                    break
            if not placed:
                grouped["其他"].append(column)
        return {group: sorted(items) for group, items in grouped.items() if items}

    if source_key == "basis":
        rules = {
            "PP基差": ["PP_basis_"],
            "丙烯基差": ["PL_basis_"],
            "MA基差": ["MA_basis_"],
        }
        grouped: dict[str, list[str]] = {name: [] for name in rules}
        grouped["其他"] = []
        for column in columns:
            placed = False
            for group, keywords in rules.items():
                if any(keyword in column for keyword in keywords):
                    grouped[group].append(column)
                    placed = True
                    break
            if not placed:
                grouped["其他"].append(column)
        return {group: sorted(items) for group, items in grouped.items() if items}

    grouped: dict[str, list[str]] = {}
    for col in columns:
        match = re.match(r"([A-Za-z]+)", col)
        prefix = match.group(1) if match else "Other"
        grouped.setdefault(prefix, []).append(col)
    return {group: sorted(items) for group, items in sorted(grouped.items())}


def _source_selectbox(label: str, key: str, allowed_sources: list[str]) -> str:
    return st.sidebar.selectbox(label, allowed_sources, format_func=lambda x: SOURCE_LABELS[x], key=key)


def _grouped_series_select(source_key: str, frame: pd.DataFrame, key_prefix: str, downstream_meta: pd.DataFrame | None = None) -> str:
    groups = _series_groups(source_key, frame.columns.tolist(), downstream_meta)
    group = st.sidebar.selectbox("选择类别", list(groups.keys()), key=f"{key_prefix}_group")
    return st.sidebar.selectbox("选择指标", groups[group], key=f"{key_prefix}_item")


def _grouped_strategy_selectbox(strategy_df: pd.DataFrame, key_prefix: str) -> str:
    grouped: dict[str, list[str]] = {}
    for _, row in strategy_df.iterrows():
        category = str(row.get("Category", "Other") or "Other")
        grouped.setdefault(category, []).append(str(row["StrategyName"]))
    category = st.sidebar.selectbox("预设组合分类", sorted(grouped), key=f"{key_prefix}_category")
    return st.sidebar.selectbox("预设组合", sorted(grouped[category]), key=f"{key_prefix}_name")


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
                "序列": col,
                "开始": start.date().isoformat() if start is not None else "N/A",
                "结束": end.date().isoformat() if end is not None else "N/A",
                "有效样本": int(frame[col].notna().sum()),
            }
        )
    overlap = frame.dropna(subset=required_cols)
    start, end = _series_valid_range(overlap[required_cols[0]]) if not overlap.empty else (None, None)
    rows.append(
        {
            "序列": "共同有效区间",
            "开始": start.date().isoformat() if start is not None else "N/A",
            "结束": end.date().isoformat() if end is not None else "N/A",
            "有效样本": int(len(overlap)),
        }
    )
    return pd.DataFrame(rows)


def _apply_date_filter(frame: pd.DataFrame, required_cols: list[str], key_prefix: str) -> tuple[pd.DataFrame, str]:
    clean_index = frame.dropna(how="all").index
    if clean_index.empty:
        return frame.iloc[0:0], "空区间"

    mode = st.sidebar.selectbox("分析区间", ["全区间", "共同有效区间", "自定义区间"], key=f"{key_prefix}_range_mode")
    if mode == "共同有效区间":
        return frame.dropna(subset=required_cols), mode
    if mode == "自定义区间":
        start_default = clean_index.min().date()
        end_default = clean_index.max().date()
        start_date, end_date = st.sidebar.date_input(
            "选择日期区间",
            value=(start_default, end_default),
            min_value=start_default,
            max_value=end_default,
            key=f"{key_prefix}_date_input",
        )
        filtered = frame.loc[(frame.index >= pd.Timestamp(start_date)) & (frame.index <= pd.Timestamp(end_date))]
        return filtered, mode
    return frame, mode


def _sidebar_excel_path() -> str:
    default_path = APP_CONFIG["excel"]["workbook_path"]
    st.sidebar.markdown("### Excel 数据源")
    source_mode = st.sidebar.radio("选择方式", ["本地路径", "拖拽/上传Excel"], key="excel_source_mode")
    if source_mode == "本地路径":
        current = st.session_state.get("excel_path", default_path)
        excel_path = st.sidebar.text_input("Excel 路径", value=current)
        st.session_state["excel_path"] = excel_path.strip() or default_path
        return st.session_state["excel_path"]

    uploaded_file = st.sidebar.file_uploader("拖拽或选择Excel文件", type=["xlsx", "xlsm", "xls"], key="excel_uploader")
    if uploaded_file is None:
        return default_path

    upload_key = f"{uploaded_file.name}_{uploaded_file.size}"
    if st.session_state.get("uploaded_excel_key") != upload_key:
        st.session_state["uploaded_excel_key"] = upload_key
        st.session_state["uploaded_excel_path"] = _save_uploaded_excel(uploaded_file)
        load_all_data.clear()
    return st.session_state["uploaded_excel_path"]


def _build_term_expression(coefficient: float, column: str, multiplier_label: str | None = None) -> str:
    coefficient_text = "" if abs(coefficient - 1.0) < 1e-12 else f"{coefficient:g} × "
    base = f"{coefficient_text}{column}"
    if multiplier_label:
        base = f"{base} × {multiplier_label}"
    return base


def _collect_term_configs(sources: dict[str, pd.DataFrame], downstream_meta: pd.DataFrame) -> list[dict[str, object]]:
    mode = st.sidebar.selectbox("组合方式", ["价差组合", "价格比组合"], key="combo_mode")
    term_count = st.sidebar.selectbox("项数", [2, 3], key="combo_term_count")
    terms: list[dict[str, object]] = []

    for idx in range(term_count):
        st.sidebar.markdown(f"#### 第 {idx + 1} 项")
        source_key = _source_selectbox(f"第 {idx + 1} 项来源", f"combo_source_{idx}", list(SOURCE_LABELS))
        column = _grouped_series_select(source_key, sources[source_key], f"combo_{idx}", downstream_meta)
        coefficient = st.sidebar.number_input(
            f"第 {idx + 1} 项系数",
            value=(1.0 if idx == 0 else (-1.0 if mode == "价差组合" else 1.0)),
            step=0.1,
            format="%.2f",
            key=f"combo_coef_{idx}",
        )

        multiplier_source = None
        multiplier_col = None
        if st.sidebar.checkbox(f"第 {idx + 1} 项乘以其他列", key=f"combo_mult_flag_{idx}"):
            multiplier_source = _source_selectbox(f"第 {idx + 1} 项乘数字段来源", f"combo_mult_source_{idx}", list(SOURCE_LABELS))
            multiplier_col = _grouped_series_select(multiplier_source, sources[multiplier_source], f"combo_mult_{idx}", downstream_meta)

        side = "numerator"
        if mode == "价格比组合":
            side = st.sidebar.selectbox(
                f"第 {idx + 1} 项归属",
                ["numerator", "denominator"],
                format_func=lambda x: "分子" if x == "numerator" else "分母",
                index=0 if idx == 0 else 1,
                key=f"combo_side_{idx}",
            )

        terms.append(
            {
                "mode": mode,
                "source": source_key,
                "column": column,
                "coefficient": float(coefficient),
                "multiplier_source": multiplier_source,
                "multiplier_col": multiplier_col,
                "side": side,
            }
        )
    return terms


def _build_combo_frame(sources: dict[str, pd.DataFrame], terms: list[dict[str, object]]) -> tuple[pd.DataFrame, str, str, list[str]]:
    union_index = pd.DatetimeIndex([])
    for frame in sources.values():
        union_index = union_index.union(frame.index)
    frame = pd.DataFrame(index=union_index.sort_values())
    required_cols: list[str] = []
    numerator_parts: list[pd.Series] = []
    denominator_parts: list[pd.Series] = []
    spread_parts: list[pd.Series] = []
    expr_parts: list[str] = []
    name_tokens: list[str] = []
    mode = str(terms[0]["mode"])

    for idx, term in enumerate(terms, start=1):
        source_key = str(term["source"])
        column = str(term["column"])
        coefficient = float(term["coefficient"])
        base_name = f"term_{idx}_{column}"
        frame = frame.join(sources[source_key][column].rename(base_name), how="left")
        value = frame[base_name]
        required_cols.append(base_name)

        multiplier_label = None
        if term["multiplier_col"]:
            multiplier_source = str(term["multiplier_source"])
            multiplier_col = str(term["multiplier_col"])
            multiplier_name = f"term_{idx}_{multiplier_col}"
            if multiplier_name not in frame.columns:
                frame = frame.join(sources[multiplier_source][multiplier_col].rename(multiplier_name), how="left")
            value = value * frame[multiplier_name]
            required_cols.append(multiplier_name)
            multiplier_label = multiplier_col

        calc_name = f"calc_{idx}"
        frame[calc_name] = coefficient * value
        required_cols.append(calc_name)
        expr_parts.append(_build_term_expression(coefficient, column, multiplier_label))
        name_tokens.append(column)

        if mode == "价格比组合":
            if term["side"] == "numerator":
                numerator_parts.append(frame[calc_name])
            else:
                denominator_parts.append(frame[calc_name])
        else:
            spread_parts.append(frame[calc_name])

    if mode == "价格比组合":
        numerator = sum(numerator_parts[1:], numerator_parts[0]) if len(numerator_parts) > 1 else numerator_parts[0]
        denominator = sum(denominator_parts[1:], denominator_parts[0]) if len(denominator_parts) > 1 else denominator_parts[0]
        frame["target"] = numerator / denominator.replace(0, pd.NA)
        num_expr = [
            _build_term_expression(float(t["coefficient"]), str(t["column"]), str(t["multiplier_col"]) if t["multiplier_col"] else None)
            for t in terms
            if t["side"] == "numerator"
        ]
        den_expr = [
            _build_term_expression(float(t["coefficient"]), str(t["column"]), str(t["multiplier_col"]) if t["multiplier_col"] else None)
            for t in terms
            if t["side"] == "denominator"
        ]
        formula = f"({' + '.join(num_expr)}) / ({' + '.join(den_expr)})"
        target_name = "_".join(name_tokens) + "_ratio"
    else:
        frame["target"] = sum(spread_parts[1:], spread_parts[0]) if len(spread_parts) > 1 else spread_parts[0]
        formula = " + ".join(expr_parts).replace("+ -", "- ")
        target_name = "_".join(name_tokens) + "_spread"

    return frame, target_name, formula, required_cols


def _risk_controls(series: pd.Series) -> dict[str, float]:
    st.sidebar.markdown("### 风控参数")
    max_window = max(20, min(len(series.dropna()), 2000))
    percentile_window = int(st.sidebar.number_input("百分位窗口", min_value=20, max_value=max_window, value=min(250, max_window), step=10))
    zscore_window = int(st.sidebar.number_input("ZScore窗口", min_value=20, max_value=max_window, value=min(60, max_window), step=10))
    var_lookback = int(st.sidebar.number_input("VaR回看窗口", min_value=20, max_value=max_window, value=min(250, max_window), step=10))
    var_horizon = int(st.sidebar.number_input("VaR期限(日)", min_value=1, max_value=20, value=5, step=1))
    var_confidence = float(st.sidebar.slider("VaR置信度", min_value=0.80, max_value=0.995, value=0.95, step=0.005))
    default_value = float(series.dropna().iloc[-1]) if not series.dropna().empty else 0.0
    custom_value = float(st.sidebar.number_input("输入值", value=default_value))
    return {
        "percentile_window": percentile_window,
        "zscore_window": zscore_window,
        "var_lookback": var_lookback,
        "var_horizon": var_horizon,
        "var_confidence": var_confidence,
        "custom_value": custom_value,
    }


def _render_metric_table(title: str, data: dict[str, float], style: str = "number") -> None:
    frame = pd.DataFrame({"指标": list(data.keys()), "数值": [_format_metric(v, style) for v in data.values()]})
    st.subheader(title)
    st.dataframe(frame, use_container_width=True, hide_index=True)


def _apply_plot_style(fig, title: str) -> None:
    fig.update_layout(
        margin=dict(l=18, r=18, t=56, b=18),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(255,255,255,0.18)",
        title=title,
        title_font=dict(size=19, color="#243446"),
        font=dict(color="#4f6275"),
        xaxis=dict(showgrid=False, zeroline=False),
        yaxis=dict(gridcolor="rgba(88, 103, 122, 0.10)", zeroline=False),
    )


def _render_formula_box(formula: str, note: str = "") -> None:
    note_html = f"<div style='margin-top:0.55rem;color:var(--muted)'>{note}</div>" if note else ""
    st.markdown(
        f"<div class='formula-box'><strong>计算逻辑</strong><br>{formula}{note_html}</div>",
        unsafe_allow_html=True,
    )


def _render_market_header(title: str, source_key: str, workbook_path: str, formula: str, note: str, series: pd.Series) -> None:
    start, end = _series_valid_range(series)
    color = SOURCE_COLORS[source_key]
    summary = _series_summary(series)
    workbook = Path(workbook_path)
    st.markdown(
        f"""
        <div class="hero-card" style="border-top: 3px solid {color};">
            <div class="hero-grid">
                <div>
                    <div class="hero-kicker">{SOURCE_LABELS[source_key]}</div>
                    <div class="hero-title">{title}</div>
                    <div class="hero-note">
                        当前工作面围绕一条目标序列展开，主图负责看路径，右侧快照负责看最新状态。
                        下面的风控和季节性是同一条序列的延伸阅读，不再拆成零碎组件。
                    </div>
                </div>
                <div class="hero-meta">
                    <div class="hero-meta-block">
                        <div class="hero-meta-label">当前值</div>
                        <div class="hero-meta-value">{summary["current"]}</div>
                    </div>
                    <div class="hero-meta-block">
                        <div class="hero-meta-label">日变动</div>
                        <div class="hero-meta-value">{summary["daily_change"]}</div>
                    </div>
                    <div class="hero-meta-block">
                        <div class="hero-meta-label">样本区间</div>
                        <div class="hero-meta-value">{start.date().isoformat() if start else 'N/A'} 至 {end.date().isoformat() if end else 'N/A'}</div>
                    </div>
                    <div class="hero-meta-block">
                        <div class="hero-meta-label">工作簿</div>
                        <div class="hero-meta-value">{workbook.name or workbook_path}</div>
                    </div>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    _render_formula_box(formula, note)


def _render_risk_section(series: pd.Series, controls: dict[str, float]) -> None:
    _render_section_intro("Risk", "风险与定位", "把绝对值、分位、波动和尾部风险收在同一组判断里。")
    risk_cfg = {
        "percentile_windows": sorted(set(METRIC_CONFIG.get("risk", {}).get("percentile_windows", []) + [int(controls["percentile_window"])])),
        "zscore_windows": sorted(set(METRIC_CONFIG.get("risk", {}).get("zscore_windows", []) + [int(controls["zscore_window"])])),
        "volatility_windows": METRIC_CONFIG.get("risk", {}).get("volatility_windows", [20, 60, 120]),
        "mdd_windows": METRIC_CONFIG.get("risk", {}).get("mdd_windows", [60, 120, 250]),
        "var_horizons": sorted(set(METRIC_CONFIG.get("risk", {}).get("var_horizons", []) + [int(controls["var_horizon"])])),
        "var_confidence_levels": sorted(set(METRIC_CONFIG.get("risk", {}).get("var_confidence_levels", []) + [float(controls["var_confidence"])])),
    }
    report = build_risk_report(series, risk_cfg)
    custom_percentile_all = percentile_of_value(series, controls["custom_value"])
    custom_percentile_window = percentile_of_value(series, controls["custom_value"], int(controls["percentile_window"]))
    custom_zscore = zscore_of_value(series, controls["custom_value"], int(controls["zscore_window"]))
    custom_var, custom_es, return_basis = var_es_over_window(
        series,
        int(controls["var_lookback"]),
        int(controls["var_horizon"]),
        float(controls["var_confidence"]),
    )

    top = st.columns(4)
    top[0].metric("当前值", _format_metric(report["current_value"]))
    top[1].metric("全历史百分位", _format_metric(report["full_history_percentile"], style="percentile"))
    top[2].metric(f"最近{int(controls['percentile_window'])}日百分位", _format_metric(report["window_percentiles"][int(controls["percentile_window"])], style="percentile"))
    top[3].metric("全历史最大回撤", _format_metric(report["max_drawdown"]["full"], style="ratio_pct"))

    middle = st.columns(5)
    middle[0].metric("输入值", _format_metric(controls["custom_value"]))
    middle[1].metric("输入值全历史百分位", _format_metric(custom_percentile_all, style="percentile"))
    middle[2].metric("输入值窗口百分位", _format_metric(custom_percentile_window, style="percentile"))
    middle[3].metric("输入值ZScore", _format_metric(custom_zscore))
    middle[4].metric("收益口径", return_basis)

    lower = st.columns(2)
    lower[0].metric(f"VaR ({int(controls['var_horizon'])}D, {controls['var_confidence']:.1%})", _format_metric(custom_var))
    lower[1].metric(f"ES ({int(controls['var_horizon'])}D, {controls['var_confidence']:.1%})", _format_metric(custom_es))

    col1, col2 = st.columns(2)
    with col1:
        _render_metric_table("滚动百分位", {f"{k}日": v for k, v in report["window_percentiles"].items()}, style="percentile")
        _render_metric_table("滚动ZScore", {f"{k}日": v for k, v in report["zscores"].items()})
    with col2:
        _render_metric_table("历史VaR", report["var"])
        _render_metric_table("历史ES", report["es"])

    bottom_left, bottom_right = st.columns(2)
    with bottom_left:
        _render_metric_table("年化波动率", {f"{k}日": v for k, v in report["volatility"].items()}, style="ratio_pct")
    with bottom_right:
        mdd = {("全历史" if k == "full" else f"{k}日"): v for k, v in report["max_drawdown"].items()}
        _render_metric_table("最大回撤", mdd, style="ratio_pct")


def _render_time_series_chart(series: pd.Series, title: str, color: str, key: str | None = None) -> None:
    key = key or f"plotly_chart_{next(_PLOTLY_KEY_COUNTER)}"
    frame = series.dropna().rename("value").to_frame()
    fig = px.line(frame, x=frame.index, y="value", title=title, template=PLOT_TEMPLATE)
    fig.update_traces(line=dict(color=color, width=3))
    _apply_plot_style(fig, title)
    st.plotly_chart(fig, use_container_width=True, key=key)


def _render_distribution_chart(series: pd.Series, title: str, color: str, key: str | None = None) -> None:
    key = key or f"plotly_chart_{next(_PLOTLY_KEY_COUNTER)}"
    fig = px.histogram(series.dropna().to_frame("value"), x="value", nbins=45, title=title, template=PLOT_TEMPLATE)
    fig.update_traces(marker_color=color, marker_line_width=0)
    _apply_plot_style(fig, title)
    st.plotly_chart(fig, use_container_width=True, key=key)


def _render_seasonality_section(series: pd.Series) -> None:
    _render_section_intro("Seasonality", "季节路径", "同一指标按年对齐后看路径、均值带和当前季节位置。")
    years = st.slider("季节性回看年数", 3, 15, APP_CONFIG.get("analysis", {}).get("seasonal_years", 5))
    seasonal_source = series.dropna()
    if APP_CONFIG.get("analysis", {}).get("remove_feb29", True):
        seasonal_source = remove_feb29(seasonal_source.to_frame("value"))["value"]
    matrix = seasonal_matrix(seasonal_source, years, interpolate=True)
    if matrix.empty:
        st.info("当前样本不足，暂时无法生成季节性图。")
        return

    plot_frame = matrix.copy()
    plot_frame.index = pd.to_datetime("2001-" + plot_frame.index)
    fig = px.line(plot_frame, title="历年季节性曲线", template=PLOT_TEMPLATE)
    fig.update_layout(
        margin=dict(l=18, r=18, t=56, b=18),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        title_font=dict(size=20, color="#edf3ff"),
        font=dict(color="#cfd9e8"),
        xaxis=dict(showgrid=False, zeroline=False),
        yaxis=dict(gridcolor="rgba(255,255,255,0.08)", zeroline=False),
    )
    st.plotly_chart(fig, use_container_width=True, key=f"plotly_chart_{next(_PLOTLY_KEY_COUNTER)}")

    mean = matrix.mean(axis=1)
    std = matrix.std(axis=1)
    band = pd.DataFrame({"均值": mean, "+1σ": mean + std, "-1σ": mean - std})
    band.index = pd.to_datetime("2001-" + band.index)
    if not band.dropna(how="all").empty:
        fig_band = px.line(band.dropna(how="all"), title="季节性均值带", template=PLOT_TEMPLATE)
        fig_band.update_layout(
            margin=dict(l=18, r=18, t=56, b=18),
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            title_font=dict(size=20, color="#edf3ff"),
            font=dict(color="#cfd9e8"),
            xaxis=dict(showgrid=False, zeroline=False),
            yaxis=dict(gridcolor="rgba(255,255,255,0.08)", zeroline=False),
        )
        st.plotly_chart(fig_band, use_container_width=True, key=f"plotly_chart_{next(_PLOTLY_KEY_COUNTER)}")

    metrics = seasonal_stats(seasonal_source, years)
    cols = st.columns(2)
    cols[0].metric("季节性分位", _format_metric(metrics["seasonal_percentile"], style="percentile"))
    cols[1].metric("季节性偏离", _format_metric(metrics["seasonal_deviation"]))


def _render_seasonality_section_soft(series: pd.Series) -> None:
    _render_section_intro("Seasonality", "季节路径", "把历年路径放在同一时间轴上，看当前点位在季节上偏高还是偏低。")
    years = st.slider("季节性回看年数", 3, 15, APP_CONFIG.get("analysis", {}).get("seasonal_years", 5))
    seasonal_source = series.dropna()
    if APP_CONFIG.get("analysis", {}).get("remove_feb29", True):
        seasonal_source = remove_feb29(seasonal_source.to_frame("value"))["value"]
    matrix = seasonal_matrix(seasonal_source, years, interpolate=True)
    if matrix.empty:
        st.info("当前样本不足，暂时无法生成季节性图。")
        return

    plot_frame = matrix.copy()
    plot_frame.index = pd.to_datetime("2001-" + plot_frame.index)
    fig = px.line(plot_frame, title="历年季节性曲线", template=PLOT_TEMPLATE)
    _apply_plot_style(fig, "历年季节性曲线")
    st.plotly_chart(fig, use_container_width=True, key=f"plotly_chart_{next(_PLOTLY_KEY_COUNTER)}")

    mean = matrix.mean(axis=1)
    std = matrix.std(axis=1)
    band = pd.DataFrame({"均值": mean, "+1σ": mean + std, "-1σ": mean - std})
    band.index = pd.to_datetime("2001-" + band.index)
    if not band.dropna(how="all").empty:
        fig_band = px.line(band.dropna(how="all"), title="季节性均值带", template=PLOT_TEMPLATE)
        _apply_plot_style(fig_band, "季节性均值带")
        st.plotly_chart(fig_band, use_container_width=True, key=f"plotly_chart_{next(_PLOTLY_KEY_COUNTER)}")

    metrics = seasonal_stats(seasonal_source, years)
    cols = st.columns(2)
    cols[0].metric("季节性分位", _format_metric(metrics["seasonal_percentile"], style="percentile"))
    cols[1].metric("季节性偏离", _format_metric(metrics["seasonal_deviation"]))


def _build_driver_frame(package) -> pd.DataFrame:
    frame = pd.DataFrame({component.key: component.series for component in package.components}, index=package.target_series.index)
    frame["target"] = package.target_series
    return frame.dropna()


def _render_driver_analysis(package) -> None:
    _render_section_intro("Decomposition", "变动拆解", "把目标序列拆成底层驱动，查看区间贡献、敏感性和情景冲击。")
    frame = _build_driver_frame(package)
    if frame.empty or len(frame) < 2:
        st.info("当前样本不足，暂时无法做变动拆解。")
        return

    max_back = min(len(frame) - 1, 250)
    default_back = min(60, max_back)
    days_back = st.slider("拆解回看区间", 1, max_back, default_back)
    start_date = frame.index[-days_back - 1]
    end_date = frame.index[-1]
    contribution = decompose_change_between_dates(package, start_date, end_date)

    if contribution.empty:
        st.info("当前区间没有足够的共同有效样本。")
        return

    total_change = contribution.attrs.get("total_change", float("nan"))
    summary_cols = st.columns(3)
    summary_cols[0].metric("拆解区间", f"{start_date.date().isoformat()} -> {end_date.date().isoformat()}")
    summary_cols[1].metric("目标变动", _format_metric(total_change))
    summary_cols[2].metric("驱动项数量", str(len(contribution)))

    chart_frame = contribution.copy()
    chart_frame["direction"] = np.where(chart_frame["contribution"] >= 0, "Positive", "Negative")
    fig = px.bar(
        chart_frame,
        x="label",
        y="contribution",
        color="direction",
        color_discrete_map={"Positive": "#7aa38f", "Negative": "#b88968"},
        title="区间贡献拆解",
        template=PLOT_TEMPLATE,
    )
    _apply_plot_style(fig, "区间贡献拆解")
    st.plotly_chart(fig, use_container_width=True, key=f"plotly_chart_{next(_PLOTLY_KEY_COUNTER)}")
    st.dataframe(contribution, use_container_width=True, hide_index=True)

    diagnostics = build_driver_diagnostics(package, windows=(60, 120, 250), z_window=60)
    sensitivity = compute_factor_sensitivity(package, bump_pct=0.01)
    scenarios = run_driver_scenarios(package, shock_pct=0.05)

    diag_col, sens_col = st.columns(2)
    with diag_col:
        st.subheader("定位诊断")
        st.dataframe(diagnostics, use_container_width=True, hide_index=True)
    with sens_col:
        st.subheader("1% 因子敏感性")
        st.dataframe(sensitivity, use_container_width=True, hide_index=True)

    if not sensitivity.empty:
        fig_sens = px.bar(
            sensitivity.sort_values("target_change"),
            x="target_change",
            y="label",
            orientation="h",
            title="敏感性排序",
            template=PLOT_TEMPLATE,
            color="target_change",
            color_continuous_scale=["#b88968", "#e6ddd1", "#7aa38f"],
        )
        _apply_plot_style(fig_sens, "敏感性排序")
        st.plotly_chart(fig_sens, use_container_width=True, key=f"plotly_chart_{next(_PLOTLY_KEY_COUNTER)}")

    st.subheader("±5% 情景冲击")
    st.dataframe(scenarios, use_container_width=True, hide_index=True)


def _lookup_strategy_row(strategy_df: pd.DataFrame, name: str) -> pd.Series | None:
    if strategy_df.empty:
        return None
    matched = strategy_df[strategy_df["StrategyName"] == name]
    if matched.empty:
        return None
    return matched.iloc[0]


def _lookup_downstream_meta(meta: pd.DataFrame, metric_name: str) -> tuple[str, str]:
    if meta.empty:
        return metric_name, ""
    matched = meta[meta["metric"] == metric_name]
    if matched.empty:
        return metric_name, ""
    row = matched.iloc[0]
    return str(row.get("formula", metric_name) or metric_name), str(row.get("note", "") or "")


def _lookup_basis_meta(meta: pd.DataFrame, metric_name: str) -> tuple[str, str]:
    if meta.empty:
        return metric_name, ""
    matched = meta[meta["metric"] == metric_name]
    if matched.empty:
        return metric_name, ""
    row = matched.iloc[0]
    return str(row.get("formula", metric_name) or metric_name), str(row.get("note", "") or "")


def _merge_notes(*notes: str) -> str:
    cleaned = [note.strip() for note in notes if isinstance(note, str) and note.strip()]
    return "<br><br>".join(cleaned)


def _infer_source_from_formula(sources: dict[str, pd.DataFrame], formula: str) -> str:
    for source_key, frame in sources.items():
        if formula in frame.columns:
            return source_key
    return "wind"


def _build_analysis_target(
    sources: dict[str, pd.DataFrame],
    portfolios: pd.DataFrame,
    strategy_df: pd.DataFrame,
    downstream_meta: pd.DataFrame,
    basis_meta: pd.DataFrame,
) -> tuple[str, str, pd.Series, str, pd.DataFrame | None, pd.DataFrame | None, str, pd.Series | None]:
    mode = st.sidebar.radio("市场序列模式", ["单序列", "自定义组合", "预设组合"], key="market_mode")

    if mode == "单序列":
        source_key = _source_selectbox("数据板块", "single_source", list(SOURCE_LABELS))
        frame = sources[source_key]
        column = _grouped_series_select(source_key, frame, "single_select", downstream_meta)
        aligned = frame[[column]].rename(columns={column: "target"})
        filtered, _ = _apply_date_filter(aligned, ["target"], "single")
        if source_key == "downstream":
            formula, note = _lookup_downstream_meta(downstream_meta, column)
        elif source_key == "basis":
            formula, note = _lookup_basis_meta(basis_meta, column)
        else:
            formula, note = column, ""
        return source_key, column, filtered["target"], formula, None, None, note, None

    if mode == "自定义组合":
        terms = _collect_term_configs(sources, downstream_meta)
        combo_frame, default_name, formula, required_cols = _build_combo_frame(sources, terms)
        filtered_frame, _ = _apply_date_filter(combo_frame, required_cols, "combo")
        coverage = _build_coverage_table(combo_frame, required_cols)
        custom_name = st.sidebar.text_input("组合名称", value="", key="combo_name")
        return "spot", custom_name.strip() or default_name, filtered_frame["target"], formula, combo_frame, coverage, "", None

    if portfolios.empty:
        fallback = sources["wind"].columns[0]
        return "wind", fallback, sources["wind"][fallback], fallback, None, None, "", None

    name = _grouped_strategy_selectbox(strategy_df, "preset")
    formula_map = dict(zip(strategy_df["StrategyName"], strategy_df["Formula"]))
    frame = portfolios[[name]].rename(columns={name: "target"})
    filtered, _ = _apply_date_filter(frame, ["target"], "preset")
    row = _lookup_strategy_row(strategy_df, name)
    note = str(row.get("Notes", "") or "") if row is not None else ""
    formula = formula_map.get(name, name)
    basis_formula, basis_note = _lookup_basis_meta(basis_meta, formula)
    source_key = _infer_source_from_formula(sources, formula)
    if basis_formula != formula or basis_note:
        formula = basis_formula
        note = _merge_notes(note, basis_note)
    return source_key, name, filtered["target"], formula, None, None, note, row


def _render_market_view(
    workbook_path: str,
    source_key: str,
    title: str,
    series: pd.Series,
    formula: str,
    note: str,
    combo_frame: pd.DataFrame | None,
    coverage: pd.DataFrame | None,
    driver_strategy_row: pd.Series | None,
    merged_for_formula: pd.DataFrame,
) -> None:
    if series.dropna().empty:
        st.warning("当前选择没有可分析的数据。")
        return

    controls = _risk_controls(series)
    driver_package = build_driver_package(merged_for_formula, driver_strategy_row) if driver_strategy_row is not None else None
    _render_market_header(title, source_key, workbook_path, formula, note, series)

    overview_tab, driver_tab, risk_tab, seasonal_tab = st.tabs(["概览", "拆解分析", "风控分析", "季节性"])
    with overview_tab:
        _render_section_intro("Overview", "主图与结构快照", "先看路径和分布，再看自定义组合的可用区间与对齐结果。")
        top_left, top_right = st.columns([1.9, 1.1])
        with top_left:
            _render_time_series_chart(series, f"{title} 历史走势", SOURCE_COLORS.get(source_key, "#6f8fa8"))
        with top_right:
            _render_distribution_chart(series, f"{title} 数值分布", SOURCE_COLORS.get(source_key, "#6f8fa8"))

        if coverage is not None:
            st.markdown('<div class="section-chip">可用区间</div>', unsafe_allow_html=True)
            st.dataframe(coverage, use_container_width=True, hide_index=True)
        if combo_frame is not None:
            st.markdown('<div class="section-chip">组合对齐数据</div>', unsafe_allow_html=True)
            st.dataframe(combo_frame.tail(120), use_container_width=True)

        _render_risk_section(series, controls)

    with driver_tab:
        if driver_package is None:
            st.info("当前视图没有可自动拆解的驱动配置。预设组合会优先支持这部分分析。")
        else:
            _render_driver_analysis(driver_package)

    with risk_tab:
        _render_risk_section(series, controls)
    with seasonal_tab:
        _render_seasonality_section_soft(series)


def _render_downstream_board(downstream_df: pd.DataFrame, downstream_meta: pd.DataFrame) -> None:
    _render_section_intro("Downstream", "下游利润板块", "把利润、净回值和综合指标压成一个连续工作面。")
    latest = downstream_df.dropna(how="all").iloc[-1]
    top = st.columns(4)
    top[0].metric("下游综合利润", _format_metric(latest.get("下游综合利润")))
    top[1].metric("综合净回值", _format_metric(latest.get("综合净回值")))
    top[2].metric("PO利润-氯醇法", _format_metric(latest.get("PO利润-氯醇法")))
    top[3].metric("丙烯腈利润", _format_metric(latest.get("丙烯腈利润")))

    snapshot_rows: list[dict[str, object]] = []
    for category in ["利润", "净回值", "综合"]:
        meta_rows = downstream_meta[downstream_meta["category"] == category]
        for _, row in meta_rows.iterrows():
            metric_name = str(row["metric"])
            if metric_name not in downstream_df.columns:
                continue
            snapshot_rows.append({"分类": category, "指标": metric_name, "最新值": _format_metric(latest.get(metric_name))})

    left, right = st.columns([1.1, 1.9])
    with left:
        st.subheader("最新快照")
        st.dataframe(pd.DataFrame(snapshot_rows), use_container_width=True, hide_index=True)
    with right:
        selection = st.columns(2)
        category = selection[0].selectbox("利润分类", ["利润", "净回值", "综合"], key="downstream_category")
        metric_options = downstream_meta.loc[downstream_meta["category"] == category, "metric"].tolist()
        metric_name = selection[1].selectbox("指标", metric_options, key="downstream_metric")
        formula, note = _lookup_downstream_meta(downstream_meta, metric_name)
        _render_formula_box(formula, note)
        metric_series = downstream_df[metric_name].dropna()
        _render_time_series_chart(metric_series, f"{metric_name} 历史走势", SOURCE_COLORS["downstream"])

        compare_metrics = downstream_meta.loc[downstream_meta["category"] == category, "metric"].tolist()
        compare_frame = downstream_df[compare_metrics].dropna(how="all").tail(180)
        if not compare_frame.empty and len(compare_frame) > 1:
            normalized = compare_frame.divide(compare_frame.iloc[0]).multiply(100)
            fig = px.line(normalized, title=f"{category} 近180日相对路径（起点=100）", template=PLOT_TEMPLATE)
            _apply_plot_style(fig, f"{category} 近180日相对路径（起点=100）")
            st.plotly_chart(fig, use_container_width=True, key=f"plotly_chart_{next(_PLOTLY_KEY_COUNTER)}")


def _render_data_preview(
    sources: dict[str, pd.DataFrame],
    portfolios: pd.DataFrame,
    downstream_meta: pd.DataFrame,
    basis_meta: pd.DataFrame,
    basis_formula_df: pd.DataFrame,
) -> None:
    _render_section_intro("Data", "数据预览", "这里保留原始表和组合结果，方便抽查口径，不抢主工作面的视觉重心。")
    tab_names = [
        "Wind",
        "Continuous Futures",
        "Manual",
        "Spot Industry",
        "Basis",
        "Downstream Profit",
        "Preset Portfolio",
        "Mapping",
    ]
    wind_tab, wind_continue_tab, manual_tab, spot_tab, basis_tab, downstream_tab, preset_tab, meta_tab = st.tabs(tab_names)
    with wind_tab:
        st.dataframe(sources["wind"].tail(200), use_container_width=True)
    with wind_continue_tab:
        st.dataframe(sources["wind_continue"].tail(200), use_container_width=True)
    with manual_tab:
        st.dataframe(sources["manual"].tail(200), use_container_width=True)
    with spot_tab:
        st.dataframe(sources["spot"].tail(200), use_container_width=True)
    with basis_tab:
        st.dataframe(basis_formula_df.tail(200), use_container_width=True)
    with downstream_tab:
        st.dataframe(sources["downstream"].tail(200), use_container_width=True)
    with preset_tab:
        if portfolios.empty:
            st.info("当前没有可用的预设组合。")
        else:
            st.dataframe(portfolios.tail(200), use_container_width=True)
    with meta_tab:
        meta_view = pd.concat(
            [
                downstream_meta.assign(source="downstream"),
                basis_meta.assign(source="basis"),
            ],
            ignore_index=True,
            sort=False,
        )
        st.dataframe(meta_view, use_container_width=True, hide_index=True)


def run_dashboard_app() -> None:
    st.set_page_config(page_title="丙烯研究看板", layout="wide")
    _inject_theme()

    workbook_path = _sidebar_excel_path()
    workbook = Path(workbook_path)

    _render_workspace_shell(workbook_path)

    if st.sidebar.button("刷新Excel数据"):
        ok = refresh_excel_workbook(workbook, APP_CONFIG["excel"].get("refresh_timeout_sec", 180))
        load_all_data.clear()
        if ok:
            st.sidebar.success("Excel 已刷新并重新加载。")
        else:
            st.sidebar.warning("Excel 自动刷新失败或被跳过，请检查本机 Excel / pywin32 环境。")

    try:
        sources, portfolios, strategy_df, downstream_meta, basis_meta, basis_formula_df = load_all_data(workbook_path)
    except Exception as exc:
        st.error(f"数据加载失败：{exc}")
        logger.exception("Failed to load workbook data")
        return

    merged_for_formula = (
        sources["wind"]
        .join(sources["wind_continue"], how="outer")
        .join(sources["manual"], how="outer")
        .join(sources["spot"], how="outer")
        .join(
            basis_formula_df.drop(
                columns=[
                    col
                    for col in basis_formula_df.columns
                    if col in set(sources["wind"].columns)
                    | set(sources["wind_continue"].columns)
                    | set(sources["manual"].columns)
                    | set(sources["spot"].columns)
                ],
                errors="ignore",
            ),
            how="outer",
        )
        .sort_index()
    )

    market_tab, downstream_tab, data_tab = st.tabs(["市场序列", "下游利润", "数据预览"])

    with market_tab:
        source_key, title, series, formula, combo_frame, coverage, note, driver_strategy_row = _build_analysis_target(
            sources,
            portfolios,
            strategy_df,
            downstream_meta,
            basis_meta,
        )
        series = series.dropna()
        if series.empty:
            st.warning("当前选择没有可用样本。")
        else:
            _render_market_view(workbook_path, source_key, title, series, formula, note, combo_frame, coverage, driver_strategy_row, merged_for_formula)

    with downstream_tab:
        _render_downstream_board(sources["downstream"], downstream_meta)

    with data_tab:
        _render_data_preview(sources, portfolios, downstream_meta, basis_meta, basis_formula_df)


def main() -> None:
    run_dashboard_app()


if __name__ == "__main__":
    main()
