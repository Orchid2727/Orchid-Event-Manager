from __future__ import annotations

import re
from typing import Any

import pandas as pd

NEVER_OUTSOURCE_COLUMN = "Never Outsource"
YES = "Yes"
NO = "No"

NEVER_OUTSOURCE_VENDORS = {"vf"}


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


def default_never_outsource(product_name: Any = "", category: Any = "", style_number: Any = "") -> bool:
    """Return the Orchid shipping default for products that must stay in house."""
    return (
        is_headwear_product(product_name, category, style_number)
        or is_pants_jeans_shorts_category(category)
        or is_flame_resistant_category(category)
        or is_bags_category(category)
    )


def resolve_never_outsource(
    value: Any,
    product_name: Any = "",
    category: Any = "",
    style_number: Any = "",
) -> bool:
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
        if current:
            result.at[index, NEVER_OUTSOURCE_COLUMN] = YES if normalize_bool(current) else NO
        else:
            result.at[index, NEVER_OUTSOURCE_COLUMN] = YES if default_never_outsource(
                row.get("Product Name", ""), row.get("Product Category", ""), row.get("Style Number", "")
            ) else NO
    return result
