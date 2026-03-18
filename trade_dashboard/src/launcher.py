from __future__ import annotations

import os
import sys
import threading
import time
import webbrowser
from pathlib import Path

from streamlit.web import cli as stcli


def _base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(getattr(sys, "_MEIPASS", Path(sys.executable).resolve().parent))
    return Path(__file__).resolve().parents[1]


def _open_browser(url: str) -> None:
    time.sleep(2)
    try:
        webbrowser.open(url)
    except Exception:
        pass


def main() -> None:
    base_dir = _base_dir()
    os.chdir(base_dir)

    dashboard_path = base_dir / "src" / "dashboard.py"
    port = "8501"
    url = f"http://localhost:{port}"

    threading.Thread(target=_open_browser, args=(url,), daemon=True).start()
    sys.argv = [
        "streamlit",
        "run",
        str(dashboard_path),
        "--server.headless",
        "true",
        "--server.port",
        port,
        "--global.developmentMode",
        "false",
    ]
    raise SystemExit(stcli.main())


if __name__ == "__main__":
    main()
