from __future__ import annotations

import logging
import re
import sys
import tempfile
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

if not hasattr(np, "unicode_"):
    np.unicode_ = np.str_

try:
    from .data_loader import load_timeseries_from_excel
    from .excel_refresh import refresh_excel_workbook
    from .industry_engine import build_propylene_profit_dashboard
    from .portfolio_engine import build_portfolios
    from .risk_engine import build_risk_report, percentile_of_value, var_es_over_window, zscore_of_value
    from .seasonal_engine import remove_feb29, seasonal_matrix, seasonal_stats
    from .utils import load_yaml, setup_logging
except ImportError:
    from src.data_loader import load_timeseries_from_excel
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
    "manual": "Manual外盘",
    "spot": "现货产业链",
    "downstream": "下游利润",
}
SOURCE_COLORS = {
    "wind": "#0f766e",
    "manual": "#2563eb",
    "spot": "#ea580c",
    "downstream": "#7c3aed",
}
PLOT_TEMPLATE = "plotly_white"


def _inject_theme() -> None:
    st.markdown(
        """
        <style>
        .stApp {
            background:
                radial-gradient(circle at top left, rgba(15, 118, 110, 0.10), transparent 30%),
                radial-gradient(circle at top right, rgba(234, 88, 12, 0.10), transparent 28%),
                linear-gradient(180deg, #f7f4ed 0%, #fcfbf8 100%);
            color: #18212b;
            font-family: "Avenir Next", "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif;
        }
        [data-testid="stSidebar"] {
            background: linear-gradient(180deg, #fbfaf6 0%, #f2ede3 100%);
            border-right: 1px solid rgba(24, 33, 43, 0.08);
        }
        .hero-card {
            padding: 1.4rem 1.5rem;
            border-radius: 24px;
            background: linear-gradient(135deg, rgba(24, 33, 43, 0.96) 0%, rgba(31, 53, 82, 0.92) 100%);
            color: #f7f3ea;
            box-shadow: 0 16px 40px rgba(24, 33, 43, 0.14);
            margin-bottom: 1rem;
        }
        .hero-kicker {
            text-transform: uppercase;
            letter-spacing: 0.08em;
            font-size: 0.78rem;
            opacity: 0.75;
            margin-bottom: 0.5rem;
        }
        .hero-title {
            font-size: 2rem;
            font-weight: 700;
            margin-bottom: 0.4rem;
        }
        .hero-note {
            font-size: 0.95rem;
            line-height: 1.6;
            opacity: 0.9;
        }
        div[data-testid="stMetric"] {
            background: rgba(255, 255, 255, 0.92);
            border: 1px solid rgba(24, 33, 43, 0.08);
            border-radius: 20px;
            padding: 0.9rem 1rem;
            box-shadow: 0 8px 22px rgba(24, 33, 43, 0.05);
        }
        .section-chip {
            display: inline-block;
            padding: 0.28rem 0.7rem;
            border-radius: 999px;
            background: rgba(15, 118, 110, 0.10);
            color: #0f766e;
            font-weight: 600;
            font-size: 0.82rem;
            margin-bottom: 0.5rem;
        }
        .formula-box {
            padding: 0.9rem 1rem;
            border-left: 4px solid #0f766e;
            background: rgba(15, 118, 110, 0.08);
            border-radius: 14px;
            margin: 0.75rem 0 1rem 0;
            color: #1f2937;
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

    downstream_df, downstream_meta = build_propylene_profit_dashboard(spot_df)
    merged_for_formula = wind_df.join(manual_df, how="outer").join(spot_df, how="outer").sort_index()

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

    sources = {"wind": wind_df, "manual": manual_df, "spot": spot_df, "downstream": downstream_df}
    return sources, portfolios, strategy_df, downstream_meta


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


def _render_formula_box(formula: str, note: str = "") -> None:
    note_html = f"<div style='margin-top:0.45rem;color:#4b5563'>{note}</div>" if note else ""
    st.markdown(
        f"<div class='formula-box'><strong>计算逻辑</strong><br>{formula}{note_html}</div>",
        unsafe_allow_html=True,
    )


def _render_market_header(title: str, source_key: str, workbook_path: str, formula: str, note: str, series: pd.Series) -> None:
    start, end = _series_valid_range(series)
    color = SOURCE_COLORS[source_key]
    st.markdown(
        f"""
        <div class="hero-card" style="border-left: 6px solid {color};">
            <div class="hero-kicker">{SOURCE_LABELS[source_key]}</div>
            <div class="hero-title">{title}</div>
            <div class="hero-note">数据文件：{workbook_path}<br>样本区间：{start.date().isoformat() if start else 'N/A'} 至 {end.date().isoformat() if end else 'N/A'}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    _render_formula_box(formula, note)


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


def _render_time_series_chart(series: pd.Series, title: str, color: str) -> None:
    frame = series.dropna().rename("value").to_frame()
    fig = px.line(frame, x=frame.index, y="value", title=title, template=PLOT_TEMPLATE)
    fig.update_traces(line=dict(color=color, width=2.5))
    fig.update_layout(margin=dict(l=20, r=20, t=50, b=20), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(255,255,255,0.85)")
    st.plotly_chart(fig, use_container_width=True)


def _render_distribution_chart(series: pd.Series, title: str, color: str) -> None:
    fig = px.histogram(series.dropna().to_frame("value"), x="value", nbins=45, title=title, template=PLOT_TEMPLATE)
    fig.update_traces(marker_color=color, marker_line_width=0)
    fig.update_layout(margin=dict(l=20, r=20, t=50, b=20), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(255,255,255,0.85)")
    st.plotly_chart(fig, use_container_width=True)


def _render_seasonality_section(series: pd.Series) -> None:
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
    fig.update_layout(margin=dict(l=20, r=20, t=50, b=20), paper_bgcolor="rgba(0,0,0,0)")
    st.plotly_chart(fig, use_container_width=True)

    mean = matrix.mean(axis=1)
    std = matrix.std(axis=1)
    band = pd.DataFrame({"均值": mean, "+1σ": mean + std, "-1σ": mean - std})
    band.index = pd.to_datetime("2001-" + band.index)
    if not band.dropna(how="all").empty:
        fig_band = px.line(band.dropna(how="all"), title="季节性均值带", template=PLOT_TEMPLATE)
        fig_band.update_layout(margin=dict(l=20, r=20, t=50, b=20), paper_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig_band, use_container_width=True)

    metrics = seasonal_stats(seasonal_source, years)
    cols = st.columns(2)
    cols[0].metric("季节性分位", _format_metric(metrics["seasonal_percentile"], style="percentile"))
    cols[1].metric("季节性偏离", _format_metric(metrics["seasonal_deviation"]))


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


def _build_analysis_target(
    sources: dict[str, pd.DataFrame],
    portfolios: pd.DataFrame,
    strategy_df: pd.DataFrame,
    downstream_meta: pd.DataFrame,
) -> tuple[str, str, pd.Series, str, pd.DataFrame | None, pd.DataFrame | None, str]:
    mode = st.sidebar.radio("市场序列模式", ["单序列", "自定义组合", "预设组合"], key="market_mode")

    if mode == "单序列":
        source_key = _source_selectbox("数据板块", "single_source", list(SOURCE_LABELS))
        frame = sources[source_key]
        column = _grouped_series_select(source_key, frame, "single_select", downstream_meta)
        aligned = frame[[column]].rename(columns={column: "target"})
        filtered, _ = _apply_date_filter(aligned, ["target"], "single")
        formula, note = _lookup_downstream_meta(downstream_meta, column) if source_key == "downstream" else (column, "")
        return source_key, column, filtered["target"], formula, None, None, note

    if mode == "自定义组合":
        terms = _collect_term_configs(sources, downstream_meta)
        combo_frame, default_name, formula, required_cols = _build_combo_frame(sources, terms)
        filtered_frame, _ = _apply_date_filter(combo_frame, required_cols, "combo")
        coverage = _build_coverage_table(combo_frame, required_cols)
        custom_name = st.sidebar.text_input("组合名称", value="", key="combo_name")
        return "spot", custom_name.strip() or default_name, filtered_frame["target"], formula, combo_frame, coverage, ""

    if portfolios.empty:
        fallback = sources["wind"].columns[0]
        return "wind", fallback, sources["wind"][fallback], fallback, None, None, ""

    name = _grouped_strategy_selectbox(strategy_df, "preset")
    formula_map = dict(zip(strategy_df["StrategyName"], strategy_df["Formula"]))
    frame = portfolios[[name]].rename(columns={name: "target"})
    filtered, _ = _apply_date_filter(frame, ["target"], "preset")
    row = _lookup_strategy_row(strategy_df, name)
    note = str(row.get("Notes", "") or "") if row is not None else ""
    return "wind", name, filtered["target"], formula_map.get(name, name), None, None, note


def _render_market_view(
    workbook_path: str,
    source_key: str,
    title: str,
    series: pd.Series,
    formula: str,
    note: str,
    combo_frame: pd.DataFrame | None,
    coverage: pd.DataFrame | None,
) -> None:
    if series.dropna().empty:
        st.warning("当前选择没有可分析的数据。")
        return

    controls = _risk_controls(series)
    _render_market_header(title, source_key, workbook_path, formula, note, series)

    overview_tab, risk_tab, seasonal_tab = st.tabs(["概览", "风控分析", "季节性"])
    with overview_tab:
        top_left, top_right = st.columns([1.9, 1.1])
        with top_left:
            _render_time_series_chart(series, f"{title} 历史走势", SOURCE_COLORS.get(source_key, "#0f766e"))
        with top_right:
            _render_distribution_chart(series, f"{title} 数值分布", SOURCE_COLORS.get(source_key, "#0f766e"))

        if coverage is not None:
            st.markdown('<div class="section-chip">可用区间</div>', unsafe_allow_html=True)
            st.dataframe(coverage, use_container_width=True, hide_index=True)
        if combo_frame is not None:
            st.markdown('<div class="section-chip">组合对齐数据</div>', unsafe_allow_html=True)
            st.dataframe(combo_frame.tail(120), use_container_width=True)

        _render_risk_section(series, controls)

    with risk_tab:
        _render_risk_section(series, controls)
    with seasonal_tab:
        _render_seasonality_section(series)


def _render_downstream_board(downstream_df: pd.DataFrame, downstream_meta: pd.DataFrame) -> None:
    st.markdown('<div class="section-chip">丙烯下游利润板块</div>', unsafe_allow_html=True)
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
            fig.update_layout(margin=dict(l=20, r=20, t=50, b=20), paper_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig, use_container_width=True)


def _render_data_preview(sources: dict[str, pd.DataFrame], portfolios: pd.DataFrame, downstream_meta: pd.DataFrame) -> None:
    tab_names = ["Wind", "Manual", "Spot Industry", "Downstream Profit", "Preset Portfolio", "Profit Mapping"]
    wind_tab, manual_tab, spot_tab, downstream_tab, preset_tab, meta_tab = st.tabs(tab_names)
    with wind_tab:
        st.dataframe(sources["wind"].tail(200), use_container_width=True)
    with manual_tab:
        st.dataframe(sources["manual"].tail(200), use_container_width=True)
    with spot_tab:
        st.dataframe(sources["spot"].tail(200), use_container_width=True)
    with downstream_tab:
        st.dataframe(sources["downstream"].tail(200), use_container_width=True)
    with preset_tab:
        if portfolios.empty:
            st.info("当前没有可用的预设组合。")
        else:
            st.dataframe(portfolios.tail(200), use_container_width=True)
    with meta_tab:
        st.dataframe(downstream_meta, use_container_width=True, hide_index=True)


def run_dashboard_app() -> None:
    st.set_page_config(page_title="丙烯研究看板", layout="wide")
    _inject_theme()

    workbook_path = _sidebar_excel_path()
    workbook = Path(workbook_path)

    st.markdown(
        """
        <div class="hero-card">
            <div class="hero-kicker">Propylene Research Dashboard</div>
            <div class="hero-title">丙烯产业链研究看板</div>
            <div class="hero-note">把期货、外盘、现货产业链和下游利润放进同一个研究界面里。你可以直接看单序列、做组合，也可以单独盯住丙烯下游利润与综合净回值。</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if st.sidebar.button("刷新Excel数据"):
        ok = refresh_excel_workbook(workbook, APP_CONFIG["excel"].get("refresh_timeout_sec", 180))
        load_all_data.clear()
        if ok:
            st.sidebar.success("Excel 已刷新并重新加载。")
        else:
            st.sidebar.warning("Excel 自动刷新失败或被跳过，请检查本机 Excel / pywin32 环境。")

    try:
        sources, portfolios, strategy_df, downstream_meta = load_all_data(workbook_path)
    except Exception as exc:
        st.error(f"数据加载失败：{exc}")
        logger.exception("Failed to load workbook data")
        return

    market_tab, downstream_tab, data_tab = st.tabs(["市场序列", "下游利润", "数据预览"])

    with market_tab:
        source_key, title, series, formula, combo_frame, coverage, note = _build_analysis_target(
            sources,
            portfolios,
            strategy_df,
            downstream_meta,
        )
        series = series.dropna()
        if series.empty:
            st.warning("当前选择没有可用样本。")
        else:
            _render_market_view(workbook_path, source_key, title, series, formula, note, combo_frame, coverage)

    with downstream_tab:
        _render_downstream_board(sources["downstream"], downstream_meta)

    with data_tab:
        _render_data_preview(sources, portfolios, downstream_meta)


def main() -> None:
    run_dashboard_app()


if __name__ == "__main__":
    main()
