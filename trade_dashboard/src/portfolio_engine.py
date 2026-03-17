from __future__ import annotations

import pandas as pd

from .formula_engine import evaluate_formula


def build_portfolios(data: pd.DataFrame, strategy_df: pd.DataFrame) -> pd.DataFrame:
    result = pd.DataFrame(index=data.index)
    enabled = strategy_df[strategy_df["Enabled"]]

    for _, row in enabled.iterrows():
        name = row["StrategyName"]
        formula = row["Formula"]
        result[name] = evaluate_formula(data, formula)

    return result
