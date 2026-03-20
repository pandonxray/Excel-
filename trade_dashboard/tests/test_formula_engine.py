import pandas as pd

from src.formula_engine import evaluate_formula


def test_evaluate_formula_basic():
    df = pd.DataFrame({"PP": [10, 20], "PG": [2, 5]})
    df.index = pd.to_datetime(["2024-01-01", "2024-01-02"])
    series = evaluate_formula(df, "PP / PG")
    assert series.iloc[0] == 5
    assert series.iloc[1] == 4


def test_evaluate_formula_supports_date_int_switch():
    df = pd.DataFrame({"PP01": [100, 100], "FEI01": [10, 10], "USDCHY": [7, 7]})
    df.index = pd.to_datetime(["2025-03-31", "2025-04-01"])
    formula = "(DATE_INT < 20250401) * (PP01 - FEI01 * USDCHY * 1.01) + (DATE_INT >= 20250401) * (PP01 - FEI01 * USDCHY * 1.11)"
    series = evaluate_formula(df, formula)
    assert round(series.iloc[0], 4) == round(100 - 10 * 7 * 1.01, 4)
    assert round(series.iloc[1], 4) == round(100 - 10 * 7 * 1.11, 4)
