from __future__ import annotations

from datetime import datetime
from hashlib import sha256
from pathlib import Path

import pandas as pd

from modules.paths import product_master_path
from modules.product_resolver import load_extended_master, normalize_style
from modules.outsource_rules import ROUTING_POLICY_VERSION


_SIGNATURE_CACHE_KEY: tuple[str, int, int, int, int] | None = None
_SIGNATURE_CACHE_VALUE: dict[str, object] | None = None


def live_product_master_path() -> Path:
    """Return the active Product Master path every time it is needed."""
    return product_master_path()


def load_live_product_master() -> tuple[Path, pd.DataFrame]:
    """Reload Product Master from disk. No module-level cache is used."""
    path = live_product_master_path()
    return path, load_extended_master(path)


def product_master_signature(path: Path | None = None) -> dict[str, object]:
    global _SIGNATURE_CACHE_KEY, _SIGNATURE_CACHE_VALUE
    path = Path(path or live_product_master_path())
    if not path.exists():
        return {
            "path": str(path),
            "exists": False,
            "modified": "",
            "size": 0,
            "sha256": "",
            "records": 0,
            "styles": 0,
        }
    stat = path.stat()
    override_path = path.parent / "never_outsource_overrides.json"
    if override_path.exists():
        override_stat = override_path.stat()
        override_mtime = int(override_stat.st_mtime_ns)
        override_size = int(override_stat.st_size)
    else:
        override_mtime = 0
        override_size = 0
    cache_key = (str(path.resolve()), int(stat.st_mtime_ns), int(stat.st_size), override_mtime, override_size)
    if cache_key == _SIGNATURE_CACHE_KEY and _SIGNATURE_CACHE_VALUE is not None:
        return dict(_SIGNATURE_CACHE_VALUE)

    payload = path.read_bytes()
    override_payload = override_path.read_bytes() if override_path.exists() else b"{}\n"
    try:
        frame = pd.read_csv(path, dtype=str).fillna("")
    except Exception:
        frame = pd.DataFrame()
    styles: set[str] = set()
    if not frame.empty:
        for _, row in frame.iterrows():
            style = normalize_style(row.get("Style Number", ""))
            name = " ".join(str(row.get("Product Name", "") or "").split()).casefold()
            styles.add(f"style:{style}" if style else f"product:{name}")
    modified = datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %I:%M:%S %p")
    result = {
        "path": str(path),
        "exists": True,
        "modified": modified,
        "size": len(payload),
        "sha256": sha256(payload).hexdigest(),
        "override_sha256": sha256(override_payload).hexdigest(),
        "routing_sha256": sha256(payload + b"\x00" + override_payload).hexdigest(),
        "records": int(len(frame)),
        "styles": int(len(styles)),
    }
    _SIGNATURE_CACHE_KEY = cache_key
    _SIGNATURE_CACHE_VALUE = dict(result)
    return result


def review_uses_current_product_master(
    review_workbook_path: Path | None,
    product_master: Path | None = None,
) -> bool:
    """Return whether a Purchase Review uses the current Master and routing policy.

    Final reports must never be generated from a stale review after a Product
    Master routing change. Older workbooks without System Info intentionally
    return False so Orchid refreshes them once before creating reports.
    """
    if not review_workbook_path or not Path(review_workbook_path).exists():
        return False
    try:
        from modules.xlsx_reader import read_sheet_rows

        rows = read_sheet_rows(Path(review_workbook_path), "System Info")
        saved = {
            " ".join(str(row[0] or "").split()).casefold(): str(row[1] or "").strip()
            for row in rows
            if len(row) >= 2 and str(row[0] or "").strip()
        }
    except Exception:
        return False
    signature = product_master_signature(product_master)
    saved_routing_hash = saved.get("product master routing sha256", "")
    current_routing_hash = str(signature.get("routing_sha256", ""))
    saved_hash = saved.get("product master sha256", "")
    current_hash = str(signature.get("sha256", ""))
    saved_policy = saved.get("routing policy version", "")
    hashes_match = (
        bool(saved_routing_hash and current_routing_hash and saved_routing_hash == current_routing_hash)
        if saved_routing_hash else bool(saved_hash and current_hash and saved_hash == current_hash)
    )
    return bool(hashes_match and saved_policy == ROUTING_POLICY_VERSION)


def style_exists(style_number: object, frame: pd.DataFrame | None = None) -> bool:
    style = normalize_style(style_number)
    if not style:
        return False
    if frame is None:
        _, frame = load_live_product_master()
    if frame.empty or "Style Number" not in frame.columns:
        return False
    return frame["Style Number"].map(normalize_style).eq(style).any()
