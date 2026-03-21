from __future__ import annotations

from dataclasses import dataclass
import math
import re
from typing import Any

import numexpr as ne
import pandas as pd

from .formula_engine import evaluate_formula

_SIMPLE_FORMULA = re.compile(r"^\s*(?P<lhs>[A-Za-z_][A-Za-z0-9_]*)\s*-\s*(?P<rhs>.+?)\s*$")
_TOKEN = re.compile(r"[A-Za-z_][A-Za-z0-9_]*")


@dataclass
class DriverComponent:
    key: str
    label: str
    expr: str
    series: pd.Series


@dataclass
class DriverPackage:
    strategy_name: str
    target_label: str
    formula: str
    target_expr: str
    target_series: pd.Series
    components: list[DriverComponent]
    derived_components: list[DriverComponent]


def _normalize_expr(expr: str) -> str:
    return expr.replace("`", "").strip()


def _evaluate_scalar_expr(expr: str, context: dict[str, float]) -> float:
    values = ne.evaluate(_normalize_expr(expr), local_dict=context)
    if hasattr(values, "item"):
        return float(values.item())
    return float(values)


def _coerce_strategy_row(strategy_row: pd.Series | dict[str, Any]) -> dict[str, Any]:
    if isinstance(strategy_row, pd.Series):
        return strategy_row.to_dict()
    return dict(strategy_row)


def infer_decomposition(strategy_row: pd.Series | dict[str, Any]) -> dict[str, Any] | None:
    row = _coerce_strategy_row(strategy_row)
    formula = str(row.get("Formula") or row.get("formula") or "").strip()
    if not formula or "DATE_INT" in formula:
        return None

    match = _SIMPLE_FORMULA.match(formula)
    if not match:
        return None

    lhs_expr = match.group("lhs").strip()
    rhs_expr = match.group("rhs").strip()
    rhs_tokens = [token.strip() for token in rhs_expr.split("*")]
    components = [{"key": "lhs", "label": lhs_expr, "expr": lhs_expr}]
    rhs_keys: list[str] = []
    for idx, token in enumerate(rhs_tokens, start=1):
        key = "rhs" if idx == 1 else f"rhs_{idx}"
        rhs_keys.append(key)
        components.append({"key": key, "label": token, "expr": token})

    rhs_expr_with_keys = " * ".join(rhs_keys) if rhs_keys else "0"
    target_expr = f"lhs - ({rhs_expr_with_keys})"
    return {
        "type": "spread",
        "target_label": str(row.get("StrategyName") or row.get("name") or "Target"),
        "target_expr": target_expr,
        "components": components,
        "derived": [{"key": "rhs_total", "label": "Right-hand side", "expr": rhs_expr_with_keys}],
    }


def get_decomposition_config(strategy_row: pd.Series | dict[str, Any]) -> dict[str, Any] | None:
    row = _coerce_strategy_row(strategy_row)
    config = row.get("decomposition")
    if isinstance(config, dict):
        return config
    return infer_decomposition(row)


def _build_component(data: pd.DataFrame, spec: dict[str, Any]) -> DriverComponent:
    expr = str(spec["expr"])
    return DriverComponent(
        key=str(spec["key"]),
        label=str(spec.get("label") or spec["key"]),
        expr=expr,
        series=evaluate_formula(data, expr),
    )


def build_driver_package(data: pd.DataFrame, strategy_row: pd.Series | dict[str, Any]) -> DriverPackage | None:
    row = _coerce_strategy_row(strategy_row)
    config = get_decomposition_config(row)
    if config is None:
        return None

    formula = str(row.get("Formula") or row.get("formula") or "")
    target_series = evaluate_formula(data, formula)
    components = [_build_component(data, spec) for spec in config.get("components", [])]
    component_df = pd.DataFrame({component.key: component.series for component in components}, index=data.index)
    derived_components = [_build_component(component_df, spec) for spec in config.get("derived", [])]

    return DriverPackage(
        strategy_name=str(row.get("StrategyName") or row.get("name") or "Target"),
        target_label=str(config.get("target_label") or row.get("StrategyName") or row.get("name") or "Target"),
        formula=formula,
        target_expr=str(config.get("target_expr") or " + ".join(component.key for component in components)),
        target_series=target_series.rename(str(row.get("StrategyName") or row.get("name") or "Target")),
        components=components,
        derived_components=derived_components,
    )


def _latest_valid_row(package: DriverPackage) -> pd.DataFrame:
    frame = pd.DataFrame({component.key: component.series for component in package.components}, index=package.target_series.index)
    frame["target"] = package.target_series
    return frame.dropna()


def decompose_change(package: DriverPackage, window: int) -> pd.DataFrame:
    frame = _latest_valid_row(package)
    if len(frame) <= window:
        return pd.DataFrame(columns=["component", "label", "contribution", "pct_of_total"])

    start_row = frame.iloc[-window - 1]
    end_row = frame.iloc[-1]
    keys = [component.key for component in package.components]
    start_context = {key: float(start_row[key]) for key in keys}
    end_context = {key: float(end_row[key]) for key in keys}

    start_value = _evaluate_scalar_expr(package.target_expr, start_context)
    end_value = _evaluate_scalar_expr(package.target_expr, end_context)
    total_change = end_value - start_value

    rows: list[dict[str, Any]] = []
    contribution_sum = 0.0
    for component in package.components:
        forward_context = dict(start_context)
        forward_context[component.key] = end_context[component.key]
        backward_context = dict(end_context)
        backward_context[component.key] = start_context[component.key]
        forward_value = _evaluate_scalar_expr(package.target_expr, forward_context)
        backward_value = _evaluate_scalar_expr(package.target_expr, backward_context)
        contribution = 0.5 * ((forward_value - start_value) + (end_value - backward_value))
        contribution_sum += contribution
        pct_of_total = contribution / total_change if not math.isclose(total_change, 0.0, abs_tol=1e-12) else float("nan")
        rows.append(
            {
                "component": component.key,
                "label": component.label,
                "contribution": contribution,
                "pct_of_total": pct_of_total,
            }
        )

    residual = total_change - contribution_sum
    if not math.isclose(residual, 0.0, abs_tol=1e-9):
        rows.append(
            {
                "component": "residual",
                "label": "Residual",
                "contribution": residual,
                "pct_of_total": residual / total_change if not math.isclose(total_change, 0.0, abs_tol=1e-12) else float("nan"),
            }
        )

    result = pd.DataFrame(rows)
    result.attrs["window"] = window
    result.attrs["total_change"] = total_change
    result.attrs["start_value"] = start_value
    result.attrs["end_value"] = end_value
    return result


def decompose_change_between_dates(package: DriverPackage, start_date: pd.Timestamp, end_date: pd.Timestamp) -> pd.DataFrame:
    frame = _latest_valid_row(package)
    if frame.empty:
        return pd.DataFrame(columns=["component", "label", "contribution", "pct_of_total"])

    filtered = frame.loc[(frame.index >= pd.Timestamp(start_date)) & (frame.index <= pd.Timestamp(end_date))]
    if len(filtered) < 2:
        return pd.DataFrame(columns=["component", "label", "contribution", "pct_of_total"])

    start_row = filtered.iloc[0]
    end_row = filtered.iloc[-1]
    keys = [component.key for component in package.components]
    start_context = {key: float(start_row[key]) for key in keys}
    end_context = {key: float(end_row[key]) for key in keys}

    start_value = _evaluate_scalar_expr(package.target_expr, start_context)
    end_value = _evaluate_scalar_expr(package.target_expr, end_context)
    total_change = end_value - start_value

    rows: list[dict[str, Any]] = []
    contribution_sum = 0.0
    for component in package.components:
        forward_context = dict(start_context)
        forward_context[component.key] = end_context[component.key]
        backward_context = dict(end_context)
        backward_context[component.key] = start_context[component.key]
        forward_value = _evaluate_scalar_expr(package.target_expr, forward_context)
        backward_value = _evaluate_scalar_expr(package.target_expr, backward_context)
        contribution = 0.5 * ((forward_value - start_value) + (end_value - backward_value))
        contribution_sum += contribution
        pct_of_total = contribution / total_change if not math.isclose(total_change, 0.0, abs_tol=1e-12) else float("nan")
        rows.append(
            {
                "component": component.key,
                "label": component.label,
                "contribution": contribution,
                "pct_of_total": pct_of_total,
            }
        )

    residual = total_change - contribution_sum
    if not math.isclose(residual, 0.0, abs_tol=1e-9):
        rows.append(
            {
                "component": "residual",
                "label": "Residual",
                "contribution": residual,
                "pct_of_total": residual / total_change if not math.isclose(total_change, 0.0, abs_tol=1e-12) else float("nan"),
            }
        )

    result = pd.DataFrame(rows)
    result.attrs["total_change"] = total_change
    result.attrs["start_value"] = start_value
    result.attrs["end_value"] = end_value
    result.attrs["start_date"] = filtered.index[0]
    result.attrs["end_date"] = filtered.index[-1]
    return result


def build_driver_diagnostics(package: DriverPackage, windows: tuple[int, ...] = (252, 756, 1260), z_window: int = 60) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []
    for label, series in [(package.target_label, package.target_series)] + [
        (component.label, component.series) for component in package.components + package.derived_components
    ]:
        clean = series.dropna()
        if clean.empty:
            continue
        current = float(clean.iloc[-1])
        row: dict[str, Any] = {"series": label, "current": current}
        for window in windows:
            sample = clean.tail(window)
            if sample.empty:
                row[f"pct_{window}"] = float("nan")
            else:
                row[f"pct_{window}"] = float((sample <= current).mean() * 100)
        z_sample = clean.tail(z_window)
        if len(z_sample) >= 2 and float(z_sample.std(ddof=0)) != 0.0:
            row[f"z_{z_window}"] = float((current - z_sample.mean()) / z_sample.std(ddof=0))
        else:
            row[f"z_{z_window}"] = float("nan")
        rows.append(row)
    return pd.DataFrame(rows)


def compute_factor_sensitivity(package: DriverPackage, bump_pct: float = 0.01) -> pd.DataFrame:
    frame = _latest_valid_row(package)
    if frame.empty:
        return pd.DataFrame(columns=["component", "label", "base_value", "bump_pct", "target_change", "target_change_pct"])

    latest = frame.iloc[-1]
    context = {component.key: float(latest[component.key]) for component in package.components}
    base_value = _evaluate_scalar_expr(package.target_expr, context)
    rows: list[dict[str, Any]] = []
    for component in package.components:
        bumped = dict(context)
        base_component_value = bumped[component.key]
        bump_amount = 1.0 if math.isclose(base_component_value, 0.0, abs_tol=1e-12) else abs(base_component_value) * bump_pct
        bumped[component.key] = base_component_value + bump_amount
        bumped_value = _evaluate_scalar_expr(package.target_expr, bumped)
        target_change = bumped_value - base_value
        rows.append(
            {
                "component": component.key,
                "label": component.label,
                "base_value": base_component_value,
                "bump_pct": bump_pct * 100,
                "target_change": target_change,
                "target_change_pct": target_change / base_value * 100 if not math.isclose(base_value, 0.0, abs_tol=1e-12) else float("nan"),
            }
        )
    return pd.DataFrame(rows)


def run_driver_scenarios(package: DriverPackage, shock_pct: float = 0.05) -> pd.DataFrame:
    frame = _latest_valid_row(package)
    if frame.empty:
        return pd.DataFrame(columns=["scenario", "target_value", "target_change"])

    latest = frame.iloc[-1]
    context = {component.key: float(latest[component.key]) for component in package.components}
    base_value = _evaluate_scalar_expr(package.target_expr, context)
    rows: list[dict[str, Any]] = []
    for component in package.components:
        for direction in (1.0, -1.0):
            shocked = dict(context)
            shocked[component.key] = shocked[component.key] * (1 + direction * shock_pct)
            scenario_name = f"{component.label} {'+' if direction > 0 else '-'}{shock_pct * 100:.0f}%"
            target_value = _evaluate_scalar_expr(package.target_expr, shocked)
            rows.append(
                {
                    "scenario": scenario_name,
                    "target_value": target_value,
                    "target_change": target_value - base_value,
                }
            )
    return pd.DataFrame(rows).sort_values("target_change", ascending=False).reset_index(drop=True)
