import pandas as pd

from src.driver_engine import (
    build_driver_diagnostics,
    build_driver_package,
    compute_factor_sensitivity,
    decompose_change,
    run_driver_scenarios,
)


def test_build_driver_package_supports_inferred_simple_spread():
    df = pd.DataFrame(
        {
            "PP01": [100, 101, 103],
            "L01": [90, 92, 94],
        },
        index=pd.to_datetime(["2025-03-28", "2025-03-31", "2025-04-01"]),
    )
    strategy_row = {
        "StrategyName": "PP01_L01_spread",
        "Formula": "PP01 - L01",
    }

    package = build_driver_package(df, strategy_row)

    assert package is not None
    assert package.target_expr == "lhs - (rhs)"
    assert [component.key for component in package.components] == ["lhs", "rhs"]
    assert package.target_series.iloc[-1] == 9


def test_driver_engine_supports_date_switched_decomposition():
    df = pd.DataFrame(
        {
            "LPG01": [4500, 4600],
            "FEI01": [500, 520],
            "USDCHY": [7.10, 7.20],
        },
        index=pd.to_datetime(["2025-03-31", "2025-04-01"]),
    )
    strategy_row = {
        "StrategyName": "LPG_FEI_01_spread",
        "Formula": "(DATE_INT < 20250401) * (LPG01 - FEI01 * USDCHY * 1.09 * 1.01) + (DATE_INT >= 20250401) * (LPG01 - FEI01 * USDCHY * 1.09 * 1.11)",
        "decomposition": {
            "target_label": "LPG内外价差",
            "target_expr": "domestic - fei * fx * vat * tariff",
            "components": [
                {"key": "domestic", "label": "国内LPG", "expr": "LPG01"},
                {"key": "fei", "label": "FEI", "expr": "FEI01"},
                {"key": "fx", "label": "汇率", "expr": "USDCHY"},
                {"key": "vat", "label": "增值税", "expr": "1.09"},
                {
                    "key": "tariff",
                    "label": "关税",
                    "expr": "(DATE_INT < 20250401) * 1.01 + (DATE_INT >= 20250401) * 1.11",
                },
            ],
            "derived": [
                {
                    "key": "import_cost",
                    "label": "进口成本",
                    "expr": "fei * fx * vat * tariff",
                }
            ],
        },
    }

    package = build_driver_package(df, strategy_row)

    assert package is not None
    assert round(package.target_series.iloc[0], 4) == round(4500 - 500 * 7.10 * 1.09 * 1.01, 4)
    assert round(package.target_series.iloc[1], 4) == round(4600 - 520 * 7.20 * 1.09 * 1.11, 4)
    assert package.derived_components[0].label == "进口成本"


def test_driver_analytics_outputs_non_empty_frames():
    df = pd.DataFrame(
        {
            "PP01": [100, 102, 104, 106, 108],
            "L01": [92, 93, 95, 97, 99],
        },
        index=pd.date_range("2025-01-01", periods=5, freq="D"),
    )
    strategy_row = {
        "StrategyName": "PP01_L01_spread",
        "Formula": "PP01 - L01",
    }
    package = build_driver_package(df, strategy_row)

    contribution = decompose_change(package, window=2)
    diagnostics = build_driver_diagnostics(package, windows=(3,), z_window=3)
    sensitivity = compute_factor_sensitivity(package, bump_pct=0.01)
    scenarios = run_driver_scenarios(package, shock_pct=0.05)

    assert not contribution.empty
    assert "contribution" in contribution.columns
    assert not diagnostics.empty
    assert "pct_3" in diagnostics.columns
    assert not sensitivity.empty
    assert "target_change" in sensitivity.columns
    assert not scenarios.empty
    assert "scenario" in scenarios.columns

