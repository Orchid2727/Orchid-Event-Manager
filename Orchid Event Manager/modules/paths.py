from __future__ import annotations

import os
from pathlib import Path
import shutil
import sys


APP_DATA_FOLDER = "Orchid Purchase Manager 2 Data"
V12_DATA_FOLDER = "Orchid Purchase Manager Data"
LEGACY_DATA_FOLDER = "Orchid Event Manager Data"


def resource_root() -> Path:
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent.parent


def source_root() -> Path:
    return Path(__file__).resolve().parent.parent


def data_dir() -> Path:
    """Return the shared Orchid Purchase Manager 2.x data folder."""
    override = os.environ.get("ORCHID_DATA_DIR", "").strip()
    path = Path(override).expanduser() if override else Path.home() / "Documents" / APP_DATA_FOLDER
    path.mkdir(parents=True, exist_ok=True)
    return path


def v12_data_dir() -> Path:
    return Path.home() / "Documents" / V12_DATA_FOLDER


def legacy_data_dir() -> Path:
    override = os.environ.get("ORCHID_LEGACY_DATA_DIR", "").strip()
    return Path(override).expanduser() if override else Path.home() / "Documents" / LEGACY_DATA_FOLDER


def seed_product_master_path() -> Path:
    return resource_root() / "data" / "product_master.csv"


def backups_dir() -> Path:
    path = data_dir() / "backups"
    path.mkdir(parents=True, exist_ok=True)
    return path


def ensure_seed_data() -> None:
    dst = data_dir() / "product_master.csv"
    if dst.exists():
        return
    src = seed_product_master_path()
    if src.exists():
        shutil.copy2(src, dst)
    else:
        dst.write_text("", encoding="utf-8")


def product_master_path() -> Path:
    ensure_seed_data()
    return data_dir() / "product_master.csv"


def product_master_update_marker_path() -> Path:
    """A lightweight cross-process signal written whenever Product Master changes."""
    return data_dir() / "product_master_updated.json"


def legacy_product_master_path() -> Path:
    return legacy_data_dir() / "product_master.csv"


def reports_dir() -> Path:
    path = data_dir() / "reports"
    path.mkdir(parents=True, exist_ok=True)
    return path
