from __future__ import annotations

import json
import os
from pathlib import Path
import shutil
import sys
import time


APP_DATA_FOLDER = "Orchid Purchase Manager 2 Data"
V12_DATA_FOLDER = "Orchid Purchase Manager Data"
LEGACY_DATA_FOLDER = "Orchid Event Manager Data"
DATA_LOCATION_MIGRATION_FILE = ".orchid_data_location_migration.json"


def resource_root() -> Path:
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent.parent


def source_root() -> Path:
    return Path(__file__).resolve().parent.parent


def documents_data_dir(home: Path | None = None) -> Path:
    """Return the pre-4.9.38 Documents location without creating it."""
    home = Path(home) if home is not None else Path.home()
    return home / "Documents" / APP_DATA_FOLDER


def application_support_data_dir(home: Path | None = None) -> Path:
    """Return Orchid's non-cloud working-data location without creating it."""
    home = Path(home) if home is not None else Path.home()
    return home / "Library" / "Application Support" / APP_DATA_FOLDER


def _has_visible_or_hidden_content(path: Path) -> bool:
    try:
        next(path.iterdir())
        return True
    except (StopIteration, FileNotFoundError, NotADirectoryError):
        return False


def _rebase_copied_state_paths(root: Path, source_root: Path, destination_root: Path) -> None:
    """Point copied active-event state at its new local files.

    The current-event state deliberately stores absolute paths to the review
    workbook, completed packet, and Event Archive.  Once the data folder has
    been safely copied, those values must follow it; otherwise a reopened app
    could still select the old iCloud Documents packet.
    """
    state_path = Path(root) / "current_review_state.json"
    if not state_path.is_file():
        return
    try:
        payload = json.loads(state_path.read_text(encoding="utf-8"))
    except (OSError, ValueError, TypeError):
        return

    source_root = Path(source_root)
    destination_root = Path(destination_root)
    changed = False

    def rebase(value):
        nonlocal changed
        if isinstance(value, str):
            try:
                relative = Path(value).expanduser().relative_to(source_root)
            except ValueError:
                return value
            updated = str(destination_root / relative)
            if updated != value:
                changed = True
            return updated
        if isinstance(value, list):
            return [rebase(item) for item in value]
        if isinstance(value, dict):
            return {key: rebase(item) for key, item in value.items()}
        return value

    updated_payload = rebase(payload)
    if changed:
        state_path.write_text(json.dumps(updated_payload, indent=2) + "\n", encoding="utf-8")


def _copy_tree_verified(source: Path, destination: Path) -> None:
    """Copy ``source`` to ``destination`` and prove every source file arrived.

    Purchase packets must not disappear because a migration was interrupted.  A
    new data folder is therefore staged beside its destination, verified by
    relative file name and byte size, and promoted only after the full copy is
    present.  The original Documents folder is intentionally never changed.
    """
    source = Path(source)
    destination = Path(destination)
    staging = destination.parent / f".{destination.name}.migration-{time.time_ns()}"
    shutil.copytree(source, staging, copy_function=shutil.copy2)
    try:
        source_files = {
            item.relative_to(source): item.stat().st_size
            for item in source.rglob("*")
            if item.is_file()
        }
        copied_files = {
            item.relative_to(staging): item.stat().st_size
            for item in staging.rglob("*")
            if item.is_file()
        }
        if copied_files != source_files:
            raise OSError("The Orchid data copy could not be verified.")
        _rebase_copied_state_paths(staging, source, destination)
        staging.rename(destination)
    except Exception:
        # Preserve the staged files for recovery if an OS or cloud-provider
        # interruption occurs.  The old Documents data is never removed.
        raise


def _migrate_documents_data(source: Path, destination: Path) -> None:
    """Make one safe, non-destructive copy of existing Documents data.

    macOS can sync Desktop and Documents through iCloud.  Orchid's packet
    manifests and PDFs need a local, application-managed home so Finder never
    opens an incomplete cloud-provider copy.  Only a brand-new or empty target
    is populated automatically; a non-empty Application Support folder is
    already authoritative and is not overwritten.
    """
    source = Path(source)
    destination = Path(destination)
    marker = destination / DATA_LOCATION_MIGRATION_FILE
    if marker.is_file():
        return

    if source.is_dir() and _has_visible_or_hidden_content(source):
        destination.parent.mkdir(parents=True, exist_ok=True)
        if not destination.exists():
            _copy_tree_verified(source, destination)
        elif not _has_visible_or_hidden_content(destination):
            # The target directory was created but has no data yet.  Copy into
            # it without removing or modifying the Documents source.
            shutil.copytree(source, destination, dirs_exist_ok=True, copy_function=shutil.copy2)
            source_files = {
                item.relative_to(source): item.stat().st_size
                for item in source.rglob("*")
                if item.is_file()
            }
            copied_files = {
                item.relative_to(destination): item.stat().st_size
                for item in destination.rglob("*")
                if item.is_file()
            }
            if copied_files != source_files:
                raise OSError("The Orchid data copy could not be verified.")
            _rebase_copied_state_paths(destination, source, destination)
    destination.mkdir(parents=True, exist_ok=True)
    marker.write_text(
        json.dumps(
            {
                "source": str(source),
                "reason": "Moved Orchid working files out of iCloud-synced Documents",
                "completed_at": time.time(),
            },
            indent=2,
        )
        + "\n",
        encoding="utf-8",
    )


def data_dir() -> Path:
    """Return Orchid's durable local data folder.

    An explicit environment override stays available for support and tests.
    Standard macOS installs now use Application Support rather than Documents,
    which may be iCloud-synced.  Existing Documents data is copied once before
    it is used; the original is retained as an extra recovery copy.
    """
    override = os.environ.get("ORCHID_DATA_DIR", "").strip()
    if override:
        path = Path(override).expanduser()
        path.mkdir(parents=True, exist_ok=True)
        return path

    path = application_support_data_dir()
    _migrate_documents_data(documents_data_dir(), path)
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
