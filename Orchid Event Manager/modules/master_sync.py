from __future__ import annotations

from datetime import datetime
from hashlib import sha256
from pathlib import Path

import pandas as pd

from modules.paths import product_master_path
from modules.product_resolver import load_extended_master, normalize_style


_SIGNATURE_CACHE_KEY: tuple[str, int, int] | None = None
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
    cache_key = (str(path.resolve()), int(stat.st_mtime_ns), int(stat.st_size))
    if cache_key == _SIGNATURE_CACHE_KEY and _SIGNATURE_CACHE_VALUE is not None:
        return dict(_SIGNATURE_CACHE_VALUE)

    payload = path.read_bytes()
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
        "records": int(len(frame)),
        "styles": int(len(styles)),
    }
    _SIGNATURE_CACHE_KEY = cache_key
    _SIGNATURE_CACHE_VALUE = dict(result)
    return result


def style_exists(style_number: object, frame: pd.DataFrame | None = None) -> bool:
    style = normalize_style(style_number)
    if not style:
        return False
    if frame is None:
        _, frame = load_live_product_master()
    if frame.empty or "Style Number" not in frame.columns:
        return False
    return frame["Style Number"].map(normalize_style).eq(style).any()
