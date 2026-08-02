from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Iterable

import pandas as pd

from modules.paths import product_master_path

DEFAULT_THREAD_INK_COLORS = [
    "Black",
    "Gold",
    "Gray",
    "Green",
    "Khaki",
    "Light Blue",
    "Maroon",
    "Navy",
    "Orange",
    "Purple",
    "Red",
    "Royal",
    "Silver",
    "Tan",
    "White",
]


def clean_color(value: object) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return re.sub(r"\s+", " ", str(value).strip())


def thread_ink_color_file(master_path: Path | None = None) -> Path:
    master = Path(master_path or product_master_path())
    return master.parent / "thread_ink_colors.json"


def _dedupe_sorted(values: Iterable[object], *, include_blank: bool = True) -> list[str]:
    registry: dict[str, str] = {}
    for raw in values:
        value = clean_color(raw)
        if not value:
            continue
        registry.setdefault(value.casefold(), value)
    ordered = sorted(registry.values(), key=lambda value: value.casefold())
    return ([""] + ordered) if include_blank else ordered


def _saved_colors(path: Path) -> list[str]:
    if not path.exists():
        return []
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return []
    if isinstance(payload, dict):
        values = payload.get("colors", [])
    else:
        values = payload
    return [clean_color(value) for value in values if clean_color(value)] if isinstance(values, list) else []


def _master_colors(master_path: Path) -> list[str]:
    if not master_path.exists():
        return []
    try:
        frame = pd.read_csv(master_path, dtype=str).fillna("")
    except Exception:
        return []
    if "Decoration Color" not in frame.columns:
        return []
    return [clean_color(value) for value in frame["Decoration Color"] if clean_color(value)]


def load_thread_ink_colors(
    master_path: Path | None = None,
    *,
    extra_values: Iterable[object] = (),
    include_blank: bool = True,
) -> list[str]:
    master = Path(master_path or product_master_path())
    values = list(DEFAULT_THREAD_INK_COLORS)
    values.extend(_saved_colors(thread_ink_color_file(master)))
    values.extend(_master_colors(master))
    values.extend(extra_values)
    return _dedupe_sorted(values, include_blank=include_blank)


def save_thread_ink_colors(
    values: Iterable[object],
    master_path: Path | None = None,
) -> list[str]:
    master = Path(master_path or product_master_path())
    path = thread_ink_color_file(master)
    merged = load_thread_ink_colors(master, extra_values=values, include_blank=False)
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(json.dumps({"colors": merged}, indent=2) + "\n", encoding="utf-8")
    temporary.replace(path)
    return [""] + merged


def register_thread_ink_color(value: object, master_path: Path | None = None) -> tuple[str, bool, list[str]]:
    cleaned = clean_color(value)
    if not cleaned:
        return "", False, load_thread_ink_colors(master_path)
    current = load_thread_ink_colors(master_path, include_blank=False)
    existing = next((item for item in current if item.casefold() == cleaned.casefold()), "")
    if existing:
        return existing, False, [""] + current
    updated = save_thread_ink_colors([cleaned], master_path)
    selected = next((item for item in updated if item.casefold() == cleaned.casefold()), cleaned)
    return selected, True, updated
