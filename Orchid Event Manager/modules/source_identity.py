from __future__ import annotations

"""Immutable Shopify source-line identities used across review and purchasing.

Generated Line IDs are presentation identifiers and can be reused after a review
workbook is regenerated.  Source IDs are derived only from the original import
fields plus an occurrence number for exact duplicate rows, so saved edits can be
replayed only onto the same Shopify source line.
"""

from hashlib import sha256
import math
import re
from typing import Mapping, Any


def clean_identity(value: object) -> str:
    if value is None:
        return ""
    try:
        if isinstance(value, float) and math.isnan(value):
            return ""
    except Exception:
        pass
    return re.sub(r"\s+", " ", str(value).strip()).casefold()


def canonical_quantity(value: object) -> str:
    text = clean_identity(value)
    if not text:
        return "0"
    try:
        number = float(text.replace(",", ""))
        return str(int(number)) if number.is_integer() else format(number, ".12g")
    except (TypeError, ValueError):
        return text


def _first(row: Mapping[str, Any], *names: str) -> object:
    for name in names:
        if name in row:
            value = row.get(name, "")
            if clean_identity(value):
                return value
    return ""


def source_base_parts(row: Mapping[str, Any]) -> tuple[str, ...]:
    """Return the immutable source fields before duplicate occurrence is added."""
    return (
        clean_identity(_first(row, "Original Order Number", "Order Number", "Name")),
        clean_identity(_first(row, "Original Order Date", "Order Date", "Created at")),
        clean_identity(_first(row, "Original Company", "Company", "Billing Company")),
        clean_identity(_first(row, "Original Employee Name", "Employee Name", "Billing Name")),
        clean_identity(_first(row, "Original Shopify Line", "Original Line Item", "Lineitem name")),
        canonical_quantity(_first(row, "Original Quantity", "Lineitem quantity", "Quantity")),
        clean_identity(_first(row, "Original Shopify SKU", "Shopify SKU", "Lineitem sku")),
        clean_identity(_first(row, "Shopify Order Notes", "Notes")),
        clean_identity(_first(row, "Shopify Line Notes", "Line Notes")),
        clean_identity(_first(row, "Original Variant Title", "Variant Title")),
        clean_identity(_first(row, "Original Variant Option 1", "Variant Option 1")),
        clean_identity(_first(row, "Original Variant Option 2", "Variant Option 2")),
        clean_identity(_first(row, "Original Variant Option 3", "Variant Option 3")),
        clean_identity(_first(row, "Original Variant Color", "Variant Color")),
        clean_identity(_first(row, "Original Variant Size", "Variant Size")),
    )


def source_base_signature(row: Mapping[str, Any]) -> str:
    parts = source_base_parts(row)
    if not any(parts):
        return ""
    return "\x1f".join(parts)


def build_source_id(row: Mapping[str, Any], occurrence: int | str = 1) -> str:
    base = source_base_signature(row)
    if not base:
        return ""
    try:
        occurrence_value = max(int(occurrence), 1)
    except (TypeError, ValueError):
        occurrence_value = 1
    digest = sha256(f"{base}\x1f{occurrence_value}".encode("utf-8")).hexdigest()[:24].upper()
    return f"SRC-{digest}"


def strict_legacy_signature(row: Mapping[str, Any]) -> str:
    """Identity for pre-4.8.96 workbooks that do not contain Source ID.

    Quantity is deliberately excluded because an old corrupt saved edit may have
    overwritten it.  The signature is accepted only when it is unique in both the
    previous and regenerated workbooks.
    """
    parts = (
        clean_identity(_first(row, "Original Order Number", "Order Number", "Name")),
        clean_identity(_first(row, "Original Order Date", "Order Date", "Created at")),
        clean_identity(_first(row, "Original Company", "Company", "Billing Company")),
        clean_identity(_first(row, "Original Employee Name", "Employee Name", "Billing Name")),
        clean_identity(_first(row, "Original Shopify Line", "Original Line Item", "Lineitem name")),
        clean_identity(_first(row, "Original Shopify SKU", "Shopify SKU", "Lineitem sku")),
        clean_identity(_first(row, "Shopify Order Notes", "Notes")),
        clean_identity(_first(row, "Shopify Line Notes", "Line Notes")),
    )
    if not any(parts):
        return ""
    return sha256("\x1f".join(parts).encode("utf-8")).hexdigest()


def manual_entry_correction_signature(row: Mapping[str, Any]) -> str:
    """Safely correlate one line after a corrected Shopify style/title typo.

    This deliberately omits the human-entered product title while retaining the
    order, customer, quantity, SKU, notes, and every variant field.  Callers
    must still require a unique match in both workbooks and must never replay
    parser/Product-Master-owned product identity fields through this fallback.
    """
    parts = (
        clean_identity(_first(row, "Original Order Number", "Order Number", "Name")),
        clean_identity(_first(row, "Original Order Date", "Order Date", "Created at")),
        clean_identity(_first(row, "Original Company", "Company", "Billing Company")),
        clean_identity(_first(row, "Original Employee Name", "Employee Name", "Billing Name")),
        canonical_quantity(_first(row, "Original Quantity", "Lineitem quantity", "Quantity")),
        clean_identity(_first(row, "Original Shopify SKU", "Shopify SKU", "Lineitem sku")),
        clean_identity(_first(row, "Shopify Order Notes", "Notes")),
        clean_identity(_first(row, "Shopify Line Notes", "Line Notes")),
        clean_identity(_first(row, "Original Variant Title", "Variant Title")),
        clean_identity(_first(row, "Original Variant Option 1", "Variant Option 1")),
        clean_identity(_first(row, "Original Variant Option 2", "Variant Option 2")),
        clean_identity(_first(row, "Original Variant Option 3", "Variant Option 3")),
        clean_identity(_first(row, "Original Variant Color", "Variant Color")),
        clean_identity(_first(row, "Original Variant Size", "Variant Size")),
    )
    if not any(parts):
        return ""
    return sha256("\x1f".join(parts).encode("utf-8")).hexdigest()


def source_identity_from_row(row: Mapping[str, Any]) -> str:
    """Return a workbook-safe identity, preferring the persisted Source ID."""
    source_id = str(row.get("Source ID", "") or "").strip()
    if source_id:
        return source_id
    line_id = str(row.get("Line ID", "") or "").strip()
    if line_id:
        return f"LEGACY-LINE:{line_id}"
    signature = strict_legacy_signature(row)
    return f"LEGACY-SIG:{signature}" if signature else ""
