import pandas as pd

from src.industry_engine import build_propylene_profit_dashboard


def test_build_propylene_profit_dashboard_computes_key_metrics():
    index = pd.to_datetime(["2026-04-08", "2026-04-09"])
    frame = pd.DataFrame(
        {
            "环氧丙烷：市场价：山东（日）": [7000, 7100],
            "液氯：市场价：山东（日）": [100, 120],
            "双氧水：50%：市场价：山东（日）": [800, 820],
            "软泡聚醚：市场价：山东（日）": [7600, 7650],
            "丙烯腈：市场价：山东（日）": [9000, 9100],
            "丙酮：市场价：山东（日）": [5000, 5050],
            "纯苯：自提价：山东（日）": [6200, 6250],
            "苯酚：市场价：山东（日）": [7600, 7620],
            "丙烯酸：普通级：市场价：华北地区（日）": [6900, 7000],
            "正丁醇：市场价：华北地区（日）": [7300, 7350],
            "辛醇：市场价：华北地区（日）": [8600, 8700],
            "合成氨：自提价：山东（日）": [2600, 2650],
            "丙烯：市场价：山东（日）": [6500, 6600],
            "PP粉：225：出厂价：山东：东方宏业（日）": [6900, 6920],
            "PP粉：300：出厂价：山东：山东凯日（日）": [6880, 6910],
        },
        index=index,
    )

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
    frame = pd.DataFrame(
        {
            "环氧丙烷：市场价：山东（日）": [7100],
            "液氯：市场价：山东（日）": [120],
            "双氧水：50%：市场价：山东（日）": [820],
            "软泡聚醚：市场价：山东（日）": [7650],
            "丙烯腈：市场价：山东（日）": [9100],
            "丙酮：市场价：山东（日）": [5050],
            "纯苯：自提价：山东（日）": [6250],
            "苯酚：市场价：山东（日）": [7620],
            "丙烯酸：普通级：市场价：华北地区（日）": [7000],
            "正丁醇：市场价：华北地区（日）": [7350],
            "辛醇：市场价：华北地区（日）": [8700],
            "合成氨：自提价：山东（日）": [2650],
            "丙烯：市场价：山东（日）": [6600],
            "PP粉：225：出厂价：山东：东方宏业（日）": [6920],
        },
        index=index,
    )

    _, meta = build_propylene_profit_dashboard(frame)

    note = meta.loc[meta["metric"] == "下游综合利润", "note"].iloc[0]
    assert "环氧氯丙烷" in note


def test_build_propylene_profit_dashboard_keeps_composite_nan_when_required_input_missing():
    index = pd.to_datetime(["2026-04-09"])
    frame = pd.DataFrame(
        {
            "环氧丙烷：市场价：山东（日）": [7100],
            "液氯：市场价：山东（日）": [120],
            "双氧水：50%：市场价：山东（日）": [820],
            "软泡聚醚：市场价：山东（日）": [7650],
            "丙烯腈：市场价：山东（日）": [9100],
            "丙酮：市场价：山东（日）": [5050],
            "纯苯：自提价：山东（日）": [6250],
            "苯酚：市场价：山东（日）": [7620],
            "丙烯酸：普通级：市场价：华北地区（日）": [7000],
            "正丁醇：市场价：华北地区（日）": [7350],
            "辛醇：市场价：华北地区（日）": [8700],
            "合成氨：自提价：山东（日）": [2650],
            "丙烯：市场价：山东（日）": [6600],
            "PP粉：225：出厂价：山东：东方宏业（日）": [6920],
        },
        index=index,
    )

    result, _ = build_propylene_profit_dashboard(frame)

    assert pd.isna(result.loc[pd.Timestamp("2026-04-09"), "下游综合利润"])
