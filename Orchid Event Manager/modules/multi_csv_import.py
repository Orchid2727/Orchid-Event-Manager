from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import hashlib
import os
import re
import shutil
import time

import pandas as pd


TRACE_SOURCE_COLUMN = "Orchid Source CSV"
TRACE_BATCH_COLUMN = "Orchid Import Batch"

# These fields describe the business contents of an order. Import trace columns,
# parser-only helper columns, and arbitrary export columns are intentionally
# excluded so that the same Shopify order exported on two different days is
# still recognized as a duplicate.
ORDER_COMPARE_FIELDS = (
    "Lineitem name",
    "Lineitem sku",
    "Vendor",
    "Variant Color",
    "Variant Size",
    "Variant Title",
    "Variant Option 1",
    "Variant Option 2",
    "Variant Option 3",
    "Lineitem quantity",
    "Lineitem price",
    "Lineitem discount",
    "Lineitem net sales",
    "Line Notes",
    "Notes",
    "Billing Company",
    "Billing Name",
    "Created at",
    "Total",
    "Subtotal",
)


@dataclass(frozen=True)
class ImportPreview:
    new_orders: tuple[str, ...]
    duplicate_orders: tuple[str, ...]
    changed_orders: tuple[str, ...]
    current_order_count: int
    incoming_order_count: int
    current_line_count: int
    incoming_line_count: int

    @property
    def new_order_count(self) -> int:
        return len(self.new_orders)

    @property
    def duplicate_order_count(self) -> int:
        return len(self.duplicate_orders)

    @property
    def changed_order_count(self) -> int:
        return len(self.changed_orders)


def _text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    return str(value).strip()


def _normalized_order_name(value) -> str:
    return _text(value).casefold()


def _order_names(frame: pd.DataFrame) -> list[str]:
    if "Name" not in frame.columns:
        return []
    names: list[str] = []
    seen: set[str] = set()
    for value in frame["Name"].tolist():
        display = _text(value)
        key = display.casefold()
        if not key or key in seen:
            continue
        seen.add(key)
        names.append(display)
    return names


def _row_signature(row: pd.Series) -> str:
    parts = [_text(row.get(field, "")).casefold() for field in ORDER_COMPARE_FIELDS]
    return hashlib.sha256("\x1f".join(parts).encode("utf-8")).hexdigest()


def order_signatures(frame: pd.DataFrame) -> dict[str, str]:
    """Return a stable content signature per order number.

    Row order is ignored, which prevents harmless export sorting differences
    from turning an unchanged order into a replacement candidate.
    """
    if "Name" not in frame.columns:
        raise ValueError("The normalized order data is missing the Name column.")
    signatures: dict[str, str] = {}
    names = frame["Name"].map(_text)
    for order_name, group in frame.assign(_orchid_order_name=names).groupby(
        "_orchid_order_name", sort=False, dropna=False
    ):
        display = _text(order_name)
        if not display:
            continue
        row_hashes = sorted(_row_signature(row) for _, row in group.iterrows())
        signatures[display.casefold()] = hashlib.sha256("|".join(row_hashes).encode("utf-8")).hexdigest()
    return signatures


def preview_additional_orders(current: pd.DataFrame, incoming: pd.DataFrame) -> ImportPreview:
    current_signatures = order_signatures(current)
    incoming_signatures = order_signatures(incoming)
    current_display = {_normalized_order_name(name): name for name in _order_names(current)}
    incoming_display = {_normalized_order_name(name): name for name in _order_names(incoming)}

    new_keys = [key for key in incoming_display if key not in current_signatures]
    duplicate_keys = [
        key for key in incoming_display
        if key in current_signatures and current_signatures[key] == incoming_signatures.get(key)
    ]
    changed_keys = [
        key for key in incoming_display
        if key in current_signatures and current_signatures[key] != incoming_signatures.get(key)
    ]

    return ImportPreview(
        new_orders=tuple(incoming_display[key] for key in new_keys),
        duplicate_orders=tuple(incoming_display[key] for key in duplicate_keys),
        changed_orders=tuple(incoming_display[key] for key in changed_keys),
        current_order_count=len(current_signatures),
        incoming_order_count=len(incoming_signatures),
        current_line_count=len(current),
        incoming_line_count=len(incoming),
    )


def annotate_import_source(frame: pd.DataFrame, source_name: str, batch_label: str) -> pd.DataFrame:
    result = frame.copy(deep=True)
    if TRACE_SOURCE_COLUMN not in result.columns:
        result[TRACE_SOURCE_COLUMN] = source_name
    else:
        result[TRACE_SOURCE_COLUMN] = result[TRACE_SOURCE_COLUMN].map(_text)
        blank = result[TRACE_SOURCE_COLUMN] == ""
        result.loc[blank, TRACE_SOURCE_COLUMN] = source_name
    if TRACE_BATCH_COLUMN not in result.columns:
        result[TRACE_BATCH_COLUMN] = batch_label
    else:
        result[TRACE_BATCH_COLUMN] = result[TRACE_BATCH_COLUMN].map(_text)
        blank = result[TRACE_BATCH_COLUMN] == ""
        result.loc[blank, TRACE_BATCH_COLUMN] = batch_label
    return result


def merge_additional_orders(
    current: pd.DataFrame,
    incoming: pd.DataFrame,
    *,
    incoming_source_name: str,
    current_source_name: str,
    replace_changed_orders: bool,
    batch_label: str | None = None,
) -> tuple[pd.DataFrame, ImportPreview, dict]:
    preview = preview_additional_orders(current, incoming)
    batch_label = batch_label or datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    current_annotated = annotate_import_source(current, current_source_name, "Original import")
    incoming_annotated = annotate_import_source(incoming, incoming_source_name, batch_label)

    incoming_order_keys = incoming_annotated["Name"].map(_normalized_order_name)
    accepted_keys = {_normalized_order_name(name) for name in preview.new_orders}
    replaced_keys: set[str] = set()
    if replace_changed_orders:
        replaced_keys = {_normalized_order_name(name) for name in preview.changed_orders}
        accepted_keys.update(replaced_keys)

    current_order_keys = current_annotated["Name"].map(_normalized_order_name)
    if replaced_keys:
        current_annotated = current_annotated.loc[~current_order_keys.isin(replaced_keys)].copy()
    accepted_incoming = incoming_annotated.loc[incoming_order_keys.isin(accepted_keys)].copy()

    columns = list(current_annotated.columns)
    for column in accepted_incoming.columns:
        if column not in columns:
            columns.append(column)
    merged = pd.concat(
        [current_annotated.reindex(columns=columns), accepted_incoming.reindex(columns=columns)],
        ignore_index=True,
        sort=False,
    ).fillna("")

    details = {
        "new_orders_added": preview.new_order_count,
        "duplicate_orders_skipped": preview.duplicate_order_count,
        "changed_orders_found": preview.changed_order_count,
        "changed_orders_replaced": len(replaced_keys),
        "changed_orders_kept": preview.changed_order_count - len(replaced_keys),
        "lines_added": len(accepted_incoming),
        "final_order_count": len(order_signatures(merged)),
        "final_line_count": len(merged),
        "batch_label": batch_label,
    }
    return merged, preview, details


def safe_event_stem(value: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", _text(value)).strip("._-")
    return cleaned[:80] or "Current_Event"


def write_combined_csv_atomic(frame: pd.DataFrame, destination: Path) -> Path:
    destination = Path(destination)
    destination.parent.mkdir(parents=True, exist_ok=True)
    temp_path = destination.with_name(f".{destination.name}.{os.getpid()}.{int(time.time() * 1000)}.tmp")
    backup_path = destination.with_suffix(destination.suffix + ".backup")
    try:
        if destination.exists():
            shutil.copy2(destination, backup_path)
        frame.to_csv(temp_path, index=False)
        with temp_path.open("rb") as handle:
            os.fsync(handle.fileno())
        os.replace(temp_path, destination)
        try:
            directory_fd = os.open(str(destination.parent), os.O_RDONLY)
            try:
                os.fsync(directory_fd)
            finally:
                os.close(directory_fd)
        except OSError:
            pass
        return destination
    finally:
        try:
            temp_path.unlink(missing_ok=True)
        except Exception:
            pass
