from __future__ import annotations

import re
from typing import Any

import pandas as pd

from modules.blank_garment_rules import BLANK_DECORATION_LABEL, is_blank_decoration
from modules.decoration_inference import EMBROIDERY, SCREEN_PRINT, infer_decoration_color
from modules.decoration_locations import (
    BEANIE, HAT, LEFT_CHEST, default_decoration_location, normalize_decoration_location,
)
from modules.purchase_rules import bool_text, category_defaults, infer_category
from modules.routing_rules import VF_STYLE_OVERRIDES, preferred_vendor_for_brand_text
from modules.internal_services import (
    infer_in_house_decoration, is_in_house_decoration, is_in_house_service_product,
    is_internal_service_style,
)


def clean(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value).strip())


def _vendor_key(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", "", clean(value).casefold())


def infer_default_decoration(category: Any, product_name: Any = "", style_number: Any = "") -> str:
    """Return Orchid's normal decoration method for a confidently identified product.

    These are defaults, not locks. The Product Master dropdown remains editable.
    Unknown or mixed-use categories stay blank so Orchid can make the decision.
    """
    category_text = clean(category) or infer_category(product_name, style_number)
    text = f"{clean(product_name)} {clean(style_number)}".casefold()

    if category_text in {"Pants / Jeans / Shorts", "Flame Resistant"}:
        return BLANK_DECORATION_LABEL
    if category_text == "Hats / Headwear":
        return EMBROIDERY
    if category_text == "Polos / Shirts":
        return EMBROIDERY
    if category_text == "Bags":
        return EMBROIDERY
    if category_text == "T-Shirts":
        return SCREEN_PRINT
    if category_text == "Safety Workwear":
        return SCREEN_PRINT
    if category_text == "Towels / Blankets":
        return EMBROIDERY
    if category_text == "Outerwear":
        # Orchid normally screen prints sweatshirts/hoodies, while jackets,
        # fleece, vests, coats and pullovers default to embroidery.
        if re.search(r"\b(hoodie|hooded sweatshirt|sweatshirt|crewneck|crew neck)\b", text):
            return SCREEN_PRINT
        return EMBROIDERY
    if category_text == "Other Apparel":
        if re.search(r"\b(hoodie|sweatshirt|crewneck|crew neck|tee|t[- ]?shirt)\b", text):
            return SCREEN_PRINT
        if re.search(r"\b(woven|button[- ]?down|work shirt|dress shirt|blouse|polo|jacket|coat|vest|fleece)\b", text):
            return EMBROIDERY
    return ""


def apply_product_intelligence(frame: pd.DataFrame) -> pd.DataFrame:
    """Fill only confident Orchid defaults while preserving manual overrides.

    Preferred-brand routing may replace S&S with SanMar because Orchid explicitly
    prefers SanMar when both distributors carry the brand. Red Kap always routes
    to VF. Other manually selected vendors remain untouched.
    """
    if frame.empty:
        return frame.copy()

    result = frame.copy()
    for column in [
        "Product Name", "Style Number", "Vendor", "Product Category",
        "Decoration Type", "Decoration Location", "Decoration Placement Instructions", "Decoration Color", "Garment Color",
        "Requires Size", "Requires Color", "Requires Decoration",
    ]:
        if column not in result.columns:
            result[column] = ""

    # Product Master is the permanent user-controlled catalog. Vendor inference
    # may fill a blank vendor, but it must never overwrite a vendor selected by
    # the user. Earlier builds treated VF as replaceable, so a saved VF choice
    # for styles such as 3339DN was silently changed back to S&S during cleanup.
    replaceable_vendor_keys = {""}

    for index, row in result.iterrows():
        product_name = clean(row.get("Product Name", ""))
        style_number = clean(row.get("Style Number", ""))
        style_key = re.sub(r"\s+", "", style_number).upper()
        description_text = " | ".join(value for value in [product_name, style_number] if value)

        # Repair two known records even when older Product Master versions saved
        # conflicting category/requirement values.
        if style_key == "6572":
            result.at[index, "Product Category"] = "Hats / Headwear"
            result.at[index, "Requires Size"] = "No"
            result.at[index, "Requires Color"] = "Yes"
            result.at[index, "Requires Decoration"] = "Yes"
            if not clean(row.get("Decoration Type", "")):
                result.at[index, "Decoration Type"] = EMBROIDERY
        elif style_key in {"1104", "CT102804"}:
            if style_key == "1104":
                result.at[index, "Vendor"] = "Tru-Spec"
            result.at[index, "Product Category"] = "Pants / Jeans / Shorts"
            result.at[index, "Requires Size"] = "Yes"
            result.at[index, "Requires Color"] = "Yes"
            result.at[index, "Requires Decoration"] = "No"
            result.at[index, "Decoration Type"] = BLANK_DECORATION_LABEL
            result.at[index, "Decoration Location"] = ""
            result.at[index, "Decoration Placement Instructions"] = ""
            result.at[index, "Decoration Color"] = ""

        brand, preferred_vendor = preferred_vendor_for_brand_text(description_text)
        current_vendor = clean(row.get("Vendor", ""))
        if preferred_vendor:
            current_key = _vendor_key(current_vendor)
            if current_key in replaceable_vendor_keys:
                result.at[index, "Vendor"] = preferred_vendor

        # Exact Orchid style rules outrank catalog-brand inference. SP3A is a
        # Red Kap Pro Airflow work shirt and must always be purchased from VF,
        # even when the Shopify title does not include the Red Kap brand name.
        if style_key in VF_STYLE_OVERRIDES:
            result.at[index, "Vendor"] = "VF"

        category = clean(row.get("Product Category", ""))
        inferred_category = infer_category(product_name, style_number)
        if not category or (category == "Other" and inferred_category != "Other"):
            category = inferred_category
            result.at[index, "Product Category"] = category

        defaults = category_defaults(category, product_name, style_number)
        internal_service_style = is_internal_service_style(style_number)
        inferred_service_type = infer_in_house_decoration(product_name, row.get("Original Line Item", ""))
        if inferred_service_type or internal_service_style:
            defaults = {"requires_size": False, "requires_color": False, "requires_decoration": False}
        if not clean(row.get("Requires Size", "")) or inferred_service_type or internal_service_style:
            result.at[index, "Requires Size"] = bool_text(defaults["requires_size"])
        if not clean(row.get("Requires Color", "")) or inferred_service_type or internal_service_style:
            result.at[index, "Requires Color"] = bool_text(defaults["requires_color"])

        decoration = clean(row.get("Decoration Type", ""))
        if not decoration:
            decoration = infer_in_house_decoration(product_name, row.get("Original Line Item", ""))
            if not decoration:
                decoration = infer_default_decoration(category, product_name, style_number)
            if decoration:
                result.at[index, "Decoration Type"] = decoration

        if decoration:
            service_only = internal_service_style or is_in_house_service_product(
                product_name, row.get("Original Line Item", ""), decoration, style_number=style_number
            )
            requires_decoration = (
                False if service_only
                else not is_blank_decoration(decoration) and not is_in_house_decoration(decoration)
            )
            if not clean(row.get("Requires Decoration", "")) or service_only or is_in_house_decoration(decoration):
                result.at[index, "Requires Decoration"] = bool_text(requires_decoration)

            current_location = clean(row.get("Decoration Location", ""))
            if requires_decoration:
                suggested_location = default_decoration_location(product_name, category, decoration)
                # Repair old Left Chest defaults for hats/beanies, but preserve any
                # more specific manually saved route.
                if suggested_location == BEANIE and current_location in {"", LEFT_CHEST, HAT}:
                    current_location = BEANIE
                elif suggested_location == HAT and current_location in {"", LEFT_CHEST}:
                    current_location = HAT
                result.at[index, "Decoration Location"] = normalize_decoration_location(
                    current_location, decoration, requires_decoration=True,
                    product_name=product_name, category=category,
                )
            else:
                result.at[index, "Decoration Location"] = ""
                result.at[index, "Decoration Placement Instructions"] = ""

            if not requires_decoration:
                result.at[index, "Decoration Color"] = ""
            elif not clean(row.get("Decoration Color", "")):
                color = clean(row.get("Garment Color", ""))
                inferred_color = infer_decoration_color(decoration, color)
                if inferred_color:
                    result.at[index, "Decoration Color"] = inferred_color

    return result
