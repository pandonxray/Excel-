from __future__ import annotations

import logging
from pathlib import Path
from typing import Any

import yaml


def ensure_parent(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)


def setup_logging(level: str = "INFO", file_path: str | None = None) -> None:
    handlers: list[logging.Handler] = [logging.StreamHandler()]
    if file_path:
        log_path = Path(file_path)
        ensure_parent(log_path)
        handlers.append(logging.FileHandler(log_path, encoding="utf-8"))

    logging.basicConfig(
        level=getattr(logging, level.upper(), logging.INFO),
        format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
        handlers=handlers,
        force=True,
    )


def load_yaml(path: str | Path) -> dict[str, Any]:
    with Path(path).open("r", encoding="utf-8") as f:
        return yaml.safe_load(f) or {}
