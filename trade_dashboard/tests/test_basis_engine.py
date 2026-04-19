import pandas as pd

from src.basis_engine import build_basis_tables


def test_build_basis_tables_computes_basis_and_east_source():
    index = pd.to_datetime(["2023-12-31", "2024-01-02"])
    wind_df = pd.DataFrame({"USDCHY": [7.0, 7.1]}, index=index)
    wind_continue_df = pd.DataFrame(
        {
            "PP_c": [7300, 7350],
            "PL_c": [6800, 6900],
            "MA_c": [2500, 2520],
            "PL_shandongspot": [7000, 7080],
        },
        index=index,
    )
    manual_df = pd.DataFrame(
        {
            "USDCHY": [7.0, 7.1],
            "ICIS_NEA": [900, 910],
            "ICIS_SEA": [880, 885],
        },
        index=index,
    )
    spot_df = pd.DataFrame(
        {
            "PP：拉丝：市场价：华北地区（日）": [7600, 7620],
            "PP：拉丝：PPH-T30S：自提价：宁波：宁波富德（日）": [7500, 7510],
            "PP：拉丝：PPH-T30S：自提价：上海（日）": [7490, 7520],
            "PP：拉丝：PPH-T03：自提价：余姚：镇海炼化（日）": [7510, 7495],
            "PP：拉丝：PPH-T03：自提价：杭州：镇海炼化（日）": [7520, 7505],
            "PP：拉丝：PPH-T30S：自提价：杭州：宁波富德（日）": [7480, 7488],
            "PP：拉丝：PPH-T03：自提价：台州：镇海炼化（日）": [7530, 7499],
            "甲醇：进口：市场价：太仓（日）": [2600, 2615],
        },
        index=index,
    )

    basis_df, display_df, basis_meta = build_basis_tables(wind_df, wind_continue_df, manual_df, spot_df)

    assert round(basis_df.loc[index[0], "PP_basis_north_china"], 4) == 300
    assert round(basis_df.loc[index[0], "PP_basis_east_china"], 4) == 180
    assert basis_df.loc[index[0], "PP_spot_east_china_min_source"] == "PP：拉丝：PPH-T30S：自提价：杭州：宁波富德（日）"
    assert basis_df.loc[index[1], "PP_spot_east_china_min_source"] == "PP：拉丝：PPH-T30S：自提价：杭州：宁波富德（日）"
    assert round(basis_df.loc[index[0], "PL_basis_shandong"], 4) == 200
    assert round(basis_df.loc[index[1], "MA_basis_taicang"], 4) == 95
    assert display_df.columns.tolist() == [
        "PP_basis_north_china",
        "PP_basis_east_china",
        "PL_basis_shandong",
        "PL_basis_nea",
        "PL_basis_sea",
        "MA_basis_taicang",
    ]
    assert "PP_spot_east_china_min_source" not in display_df.columns
    assert "最新有效来源" in basis_meta.loc[basis_meta["metric"] == "PP_basis_east_china", "note"].iloc[0]


def test_build_basis_tables_switches_nea_tariff_on_2024_boundary():
    index = pd.to_datetime(["2023-12-31", "2024-01-01"])
    wind_df = pd.DataFrame({"USDCHY": [7.0, 7.0]}, index=index)
    wind_continue_df = pd.DataFrame(
        {
            "PP_c": [0, 0],
            "PL_c": [1000, 1000],
            "MA_c": [0, 0],
            "PL_shandongspot": [0, 0],
        },
        index=index,
    )
    manual_df = pd.DataFrame(
        {
            "ICIS_NEA": [100, 100],
            "ICIS_SEA": [100, 100],
        },
        index=index,
    )
    spot_df = pd.DataFrame(
        {
            "PP：拉丝：市场价：华北地区（日）": [0, 0],
            "PP：拉丝：PPH-T30S：自提价：宁波：宁波富德（日）": [0, 0],
            "PP：拉丝：PPH-T30S：自提价：上海（日）": [0, 0],
            "PP：拉丝：PPH-T03：自提价：余姚：镇海炼化（日）": [0, 0],
            "PP：拉丝：PPH-T03：自提价：杭州：镇海炼化（日）": [0, 0],
            "PP：拉丝：PPH-T30S：自提价：杭州：宁波富德（日）": [0, 0],
            "PP：拉丝：PPH-T03：自提价：台州：镇海炼化（日）": [0, 0],
            "甲醇：进口：市场价：太仓（日）": [0, 0],
        },
        index=index,
    )

    basis_df, _, _ = build_basis_tables(wind_df, wind_continue_df, manual_df, spot_df)

    assert round(basis_df.loc[index[0], "PL_nea_tariff"], 4) == 1.02
    assert round(basis_df.loc[index[1], "PL_nea_tariff"], 4) == 1.0
    assert round(basis_df.loc[index[0], "PL_spot_nea_import_taxed"], 4) == round(100 * 7.0 * 1.13 * 1.02, 4)
    assert round(basis_df.loc[index[1], "PL_spot_nea_import_taxed"], 4) == round(100 * 7.0 * 1.13, 4)
