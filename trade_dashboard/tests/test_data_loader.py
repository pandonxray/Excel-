import pandas as pd
from openpyxl import Workbook

from src.data_loader import load_timeseries_from_excel


def test_load_timeseries_from_excel_supports_multi_header_and_excel_serial_dates(tmp_path):
    workbook = tmp_path / "wind_like.xlsx"
    rows = [
        ["Wind", "PP01", "LPG01"],
        ["指标名称", "期货收盘价(1月交割连续):聚丙烯", "期货收盘价(1月交割连续):LPG"],
        [46098, 7771, 4387],
        [46097, 7870, 4412],
    ]
    pd.DataFrame(rows).to_excel(workbook, sheet_name="wind_raw_data", header=False, index=False)

    df = load_timeseries_from_excel(
        workbook,
        "wind_raw_data",
        date_column="Wind",
        header_rows=[0, 1],
        column_name_row=0,
    )

    assert list(df.columns) == ["PP01", "LPG01"]
    assert df.index.tolist() == [pd.Timestamp("2026-03-16"), pd.Timestamp("2026-03-17")]
    assert df.loc[pd.Timestamp("2026-03-17"), "PP01"] == 7771


def test_load_timeseries_from_excel_ignores_na_and_blank_values(tmp_path):
    workbook = tmp_path / "wind_like_missing.xlsx"
    rows = [
        ["Wind", "PP01", "LPG01", "EMPTY"],
        ["指标名称", "聚丙烯", "LPG", "空列"],
        [46098, "7771", "NA", ""],
        [46097, " ", "4380", "N/A"],
    ]
    pd.DataFrame(rows).to_excel(workbook, sheet_name="wind_raw_data", header=False, index=False)

    df = load_timeseries_from_excel(
        workbook,
        "wind_raw_data",
        date_column="Wind",
        header_rows=[0, 1],
        column_name_row=0,
    )

    assert list(df.columns) == ["PP01", "LPG01"]
    assert pd.isna(df.loc[pd.Timestamp("2026-03-17"), "LPG01"])
    assert pd.isna(df.loc[pd.Timestamp("2026-03-16"), "PP01"])


def test_load_timeseries_from_excel_drops_duplicate_date_columns(tmp_path):
    workbook = tmp_path / "spot_like.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "spot_data_propylene_industry"

    rows = [
        ["钢联数据", "钢联数据", "钢联数据", "钢联数据"],
        ["指标名称", "环氧丙烷：市场价：山东（日）", "指标名称", "丙烯：市场价：山东（日）"],
        ["频度", "日", "频度", "日"],
        ["指标描述", "", "指标描述", ""],
        [pd.Timestamp("2026-04-09"), 7000, pd.Timestamp("2026-04-09"), 6600],
        [pd.Timestamp("2026-04-08"), 6950, pd.Timestamp("2026-04-08"), 6550],
    ]
    for row in rows:
        ws.append(row)
    wb.save(workbook)

    df = load_timeseries_from_excel(
        workbook,
        "spot_data_propylene_industry",
        date_column="指标名称",
        header_rows=[0, 1, 2, 3],
        column_name_row=1,
    )

    assert list(df.columns) == ["环氧丙烷：市场价：山东（日）", "丙烯：市场价：山东（日）"]
    assert df.loc[pd.Timestamp("2026-04-09"), "环氧丙烷：市场价：山东（日）"] == 7000
