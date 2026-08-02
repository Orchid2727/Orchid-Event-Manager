from __future__ import annotations

import re
from typing import Any

import pandas as pd

from modules.purchase_rules import is_known_insole_style

NEVER_OUTSOURCE_COLUMN = "Never Outsource"
YES = "Yes"
NO = "No"
# Increment when a routing policy change requires existing Purchase Reviews to
# be rebuilt before reports are generated.
# Version 4 adds permanent insole style recognition.  Existing reviews must be
# rebuilt once so their no-color / never-outsource routing is refreshed.
ROUTING_POLICY_VERSION = "4"

# These vendors ship to Orchid for in-house handling rather than directly to an
# outside decorator.  Include both the catalog name and its common short form
# so imported data and Product Master selections receive the same default.
NEVER_OUTSOURCE_VENDORS = {
    "vf",
    "big top tees",
    "bigtop tees",
    "edwards",
    "berne",
    "berne apparel",
}


def clean(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value).strip())


def normalize_bool(value: Any, default: bool = False) -> bool:
    text = clean(value).casefold()
    if text in {"yes", "y", "true", "1", "never outsource"}:
        return True
    if text in {"no", "n", "false", "0", "allow outsource"}:
        return False
    return bool(default)


def vendor_never_outsource(vendor: Any = "") -> bool:
    """Return True for vendors that Orchid always ships to itself for in-house handling."""
    return clean(vendor).casefold() in NEVER_OUTSOURCE_VENDORS


def is_headwear_product(product_name: Any = "", category: Any = "", style_number: Any = "") -> bool:
    category_text = clean(category).casefold()
    if category_text == "hats / headwear":
        return True
    text = f"{clean(product_name)} {clean(style_number)}".casefold()
    return bool(re.search(r"\b(hat|hats|cap|caps|beanie|beanies|boonie|boonies|booney|booneys|visor|headwear)\b", text))


def is_pants_jeans_shorts_category(category: Any = "") -> bool:
    """Return True for Orchid's saved bottoms category."""
    category_text = re.sub(r"\s*/\s*", " / ", clean(category).casefold())
    return category_text == "pants / jeans / shorts"


def is_flame_resistant_category(category: Any = "") -> bool:
    """Return True for Orchid's saved flame-resistant category."""
    return clean(category).casefold() == "flame resistant"


def is_bags_category(category: Any = "") -> bool:
    """Return True for Orchid's saved bags category."""
    return clean(category).casefold() == "bags"


def is_boots_category(category: Any = "") -> bool:
    """Return True for Orchid's saved boots and footwear category."""
    return clean(category).casefold() == "boots"


def is_boots_product(product_name: Any = "", category: Any = "", style_number: Any = "") -> bool:
    """Return True for a boot, boot accessory, or insole.

    Category is checked first so saved Product Master records are decisive;
    the title/style fallback protects a newly imported boot before setup.
    """
    if is_boots_category(category):
        return True
    if is_known_insole_style(style_number):
        return True
    text = f"{clean(product_name)} {clean(style_number)}".casefold()
    return bool(re.search(r"\b(boot|boots|insole|insoles)\b", text))


def default_never_outsource(product_name: Any = "", category: Any = "", style_number: Any = "") -> bool:
    """Return the Orchid shipping default for products that must stay in house."""
    return (
        is_headwear_product(product_name, category, style_number)
        or is_pants_jeans_shorts_category(category)
        or is_flame_resistant_category(category)
        or is_bags_category(category)
        or is_boots_product(product_name, category, style_number)
    )


def resolve_never_outsource(
    value: Any,
    product_name: Any = "",
    category: Any = "",
    style_number: Any = "",
) -> bool:
    # Boots are an Orchid-only purchasing category.  This intentionally
    # outranks a stale explicit No saved by a build before the Boots policy.
    if is_boots_product(product_name, category, style_number):
        return True
    text = clean(value)
    if text:
        return normalize_bool(text, False)
    return default_never_outsource(product_name, category, style_number)


def apply_never_outsource_defaults(frame: pd.DataFrame) -> pd.DataFrame:
    result = frame.copy()
    if NEVER_OUTSOURCE_COLUMN not in result.columns:
        result[NEVER_OUTSOURCE_COLUMN] = ""
    if result.empty:
        return result
    for index, row in result.iterrows():
        current = clean(row.get(NEVER_OUTSOURCE_COLUMN, ""))
        if is_boots_product(row.get("Product Name", ""), row.get("Product Category", ""), row.get("Style Number", "")):
            result.at[index, NEVER_OUTSOURCE_COLUMN] = YES
        elif current:
            result.at[index, NEVER_OUTSOURCE_COLUMN] = YES if normalize_bool(current) else NO
        else:
            result.at[index, NEVER_OUTSOURCE_COLUMN] = YES if (
                vendor_never_outsource(row.get("Vendor", ""))
                or default_never_outsource(
                    row.get("Product Name", ""), row.get("Product Category", ""), row.get("Style Number", "")
                )
            ) else NO
    return result


def never_outsource_override_path(product_master_path):
    """Return the Product Master sidecar used by the editor for explicit overrides."""
    from pathlib import Path
    return Path(product_master_path).parent / "never_outsource_overrides.json"


def load_never_outsource_overrides(product_master_path) -> dict[str, bool]:
    import json
    path = never_outsource_override_path(product_master_path)
    if not path.exists():
        return {}
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {}
    if not isinstance(payload, dict):
        return {}
    return {str(key).strip(): bool(value) for key, value in payload.items() if str(key).strip()}


def apply_never_outsource_overrides(frame: pd.DataFrame, product_master_path) -> pd.DataFrame:
    """Apply explicit editor sidecar values to the routing frame.

    The sidecar exists so a deliberate Never Outsource selection survives color
    row edits. It is part of the live routing source and must be honored by the
    parser, not only displayed by the Product Master editor.
    """
    result = frame.copy()
    overrides = load_never_outsource_overrides(product_master_path)
    if not overrides or result.empty:
        return result
    if NEVER_OUTSOURCE_COLUMN not in result.columns:
        result[NEVER_OUTSOURCE_COLUMN] = ""
    for index, row in result.iterrows():
        style = re.sub(r"\s+", "", clean(row.get("Style Number", ""))).upper()
        product = re.sub(r"\s+", " ", clean(row.get("Product Name", "")).casefold()).strip()
        key = f"style:{style}" if style else f"product:{product}"
        if is_boots_product(row.get("Product Name", ""), row.get("Product Category", ""), row.get("Style Number", "")):
            result.at[index, NEVER_OUTSOURCE_COLUMN] = YES
        elif key in overrides:
            result.at[index, NEVER_OUTSOURCE_COLUMN] = YES if overrides[key] else NO
    return result
