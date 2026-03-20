from __future__ import annotations

import logging
import re

import numexpr as ne
import pandas as pd

logger = logging.getLogger(__name__)
_TOKEN = re.compile(r"[\u4e00-\u9fffA-Za-z_][\u4e00-\u9fffA-Za-z0-9_\/]*")


def _normalize_formula(formula: str) -> str:
    return formula.replace("`", "")


def evaluate_formula(df: pd.DataFrame, formula: str) -> pd.Series:
    formula = _normalize_formula(formula)
    names = set(_TOKEN.findall(formula))
    context: dict[str, pd.Series] = {
        "DATE_INT": pd.Series(df.index.strftime("%Y%m%d").astype(int), index=df.index),
    }

    for name in names:
        if name in {"and", "or", "DATE_INT"}:
            continue
        if name not in df.columns:
            logger.warning("Formula field not found; filling with NaN: %s", name)
            context[name] = pd.Series(float("nan"), index=df.index)
        else:
            context[name] = df[name]

    try:
        values = ne.evaluate(formula, local_dict=context)
        return pd.Series(values, index=df.index, name=formula)
    except Exception as exc:
        logger.exception("Formula evaluation failed: %s", exc)
        raise ValueError(f"Unable to evaluate formula: {formula}") from exc
