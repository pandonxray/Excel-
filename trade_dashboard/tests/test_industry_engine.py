import pandas as pd

from src.industry_engine import build_propylene_profit_dashboard


def _base_frame(index: pd.DatetimeIndex) -> pd.DataFrame:
    return pd.DataFrame(
        {
            "环氧丙烷：市场价：山东（日）": [7000 + i * 100 for i in range(len(index))],
            "液氯：市场价：山东（日）": [100 + i * 20 for i in range(len(index))],
            "双氧水：50%：市场价：山东（日）": [800 + i * 20 for i in range(len(index))],
            "软泡聚醚：市场价：山东（日）": [7600 + i * 50 for i in range(len(index))],
            "丙烯腈：市场价：山东（日）": [9000 + i * 100 for i in range(len(index))],
            "丙酮：市场价：山东（日）": [5000 + i * 50 for i in range(len(index))],
            "纯苯：自提价：山东（日）": [6200 + i * 50 for i in range(len(index))],
            "苯酚：市场价：山东（日）": [7600 + i * 20 for i in range(len(index))],
            "丙烯酸：普通级：市场价：华北地区（日）": [6900 + i * 100 for i in range(len(index))],
            "正丁醇：市场价：华北地区（日）": [7300 + i * 50 for i in range(len(index))],
            "辛醇：市场价：华北地区（日）": [8600 + i * 100 for i in range(len(index))],
            "合成氨：自提价：山东（日）": [2600 + i * 50 for i in range(len(index))],
            "丙烯：市场价：山东（日）": [6500 + i * 100 for i in range(len(index))],
            "PP粉：225：出厂价：山东：东方宏业（日）": [6900 + i * 20 for i in range(len(index))],
            "PP粉：300：出厂价：山东：山东凯日（日）": [6880 + i * 30 for i in range(len(index))],
        },
        index=index,
    )


def test_build_propylene_profit_dashboard_computes_key_metrics():
    index = pd.to_datetime(["2026-04-08", "2026-04-09"])
    frame = _base_frame(index)

    result, meta = build_propylene_profit_dashboard(frame)

    expected_po_profit = 7100 - 0.85 * (6600 + 100) - 1.4 * (120 + 100) - 1800
    expected_netback = (7000 - 1800) / 0.71 - 100

    assert round(result.loc[pd.Timestamp("2026-04-09"), "PO利润-氯醇法"], 6) == round(expected_po_profit, 6)
    assert round(result.loc[pd.Timestamp("2026-04-09"), "丙烯酸净回值"], 6) == round(expected_netback, 6)
    assert "下游综合利润" in result.columns
    assert "综合净回值" in result.columns
    assert not meta[meta["metric"] == "下游综合利润"].empty


def test_build_propylene_profit_dashboard_records_missing_ech_note():
    index = pd.to_datetime(["2026-04-09"])
    frame = _base_frame(index)

    _, meta = build_propylene_profit_dashboard(frame)

    note = meta.loc[meta["metric"] == "下游综合利润", "note"].iloc[0]
    assert "环氧氯丙烷" in note


def test_build_propylene_profit_dashboard_renormalizes_composite_when_ech_missing():
    index = pd.to_datetime(["2026-04-09"])
    frame = _base_frame(index)

    result, _ = build_propylene_profit_dashboard(frame)

    expected = (
        result.loc[pd.Timestamp("2026-04-09"), "PO利润-氯醇法"] * 0.2165754
        + result.loc[pd.Timestamp("2026-04-09"), "丙烯酸利润"] * 0.1008041
        + result.loc[pd.Timestamp("2026-04-09"), "丙烯腈利润"] * 0.2109155
        + result.loc[pd.Timestamp("2026-04-09"), "丁醇利润"] * 0.0593774
        + result.loc[pd.Timestamp("2026-04-09"), "辛醇利润"] * 0.0823884
        + result.loc[pd.Timestamp("2026-04-09"), "酚酮综合利润"] * 0.0474092
        + result.loc[pd.Timestamp("2026-04-09"), "粉料利润"] * 0.2693376
    ) / (0.2165754 + 0.1008041 + 0.2109155 + 0.0593774 + 0.0823884 + 0.0474092 + 0.2693376)

    assert round(float(result.loc[pd.Timestamp("2026-04-09"), "下游综合利润"]), 6) == round(float(expected), 6)
