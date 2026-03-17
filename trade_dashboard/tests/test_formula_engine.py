import pandas as pd

from src.formula_engine import evaluate_formula


def test_evaluate_formula_basic():
    df = pd.DataFrame({"PP": [10, 20], "PG": [2, 5]})
    series = evaluate_formula(df, "PP / PG")
    assert series.iloc[0] == 5
    assert series.iloc[1] == 4
