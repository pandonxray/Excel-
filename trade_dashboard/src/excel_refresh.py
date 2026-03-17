from __future__ import annotations

import logging
import time
from pathlib import Path

logger = logging.getLogger(__name__)


def refresh_excel_workbook(workbook_path: str | Path, timeout_sec: int = 180) -> bool:
    """在 Windows 环境使用 Excel COM 执行 RefreshAll 并保存。"""
    workbook_path = str(Path(workbook_path).resolve())

    try:
        import win32com.client  # type: ignore
    except ImportError:
        logger.warning("未安装 pywin32，跳过 Excel 自动刷新。")
        return False

    excel = None
    wb = None
    try:
        logger.info("开始刷新 Excel: %s", workbook_path)
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(workbook_path)
        wb.RefreshAll()

        start = time.time()
        while time.time() - start < timeout_sec:
            try:
                excel.CalculateUntilAsyncQueriesDone()
                break
            except Exception:
                time.sleep(1)

        wb.Save()
        logger.info("Excel 刷新并保存完成。")
        return True
    except Exception as exc:
        logger.exception("Excel 刷新失败: %s", exc)
        return False
    finally:
        if wb is not None:
            wb.Close(SaveChanges=True)
        if excel is not None:
            excel.Quit()
