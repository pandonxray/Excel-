import pandas as pd

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
