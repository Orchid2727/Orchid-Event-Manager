from __future__ import annotations

import re
from typing import Mapping

import pandas as pd

BLANK_DECORATION_LABEL = "Blank Garment (No Decoration)"
LEGACY_BLANK_DECORATION_VALUES = {
    "none",
    "blank garment",
    "blank garments",
    "blank garment (no decoration)",
    "no decoration",
}

# Orchid normally orders these products undecorated. The text rule applies only
# when no explicit decoration type exists, so intentional exceptions remain possible.
BLANK_GARMENT_TERMS = (
    "pant",
    "pants",
    "jean",
    "jeans",
    "trouser",
    "trousers",
    "shorts",
    "bib overall",
    "bib overalls",
    "overall",
    "overalls",
)

# These Shopify/Product Master records currently contain only a product number as
# the description, so name-based detection cannot identify them reliably.
KNOWN_BLANK_GARMENT_STYLES = {
    "01MWXDB",
    "02MCWST",
    "874",
    "PX62",
    "WRT20JH",
    "CT102804",
}


def clean_text(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value).strip())


def normalize_style(value: object) -> str:
    return re.sub(r"\s+", "", clean_text(value)).upper()


def is_blank_decoration(value: object) -> bool:
    return clean_text(value).casefold() in LEGACY_BLANK_DECORATION_VALUES


def normalize_decoration_type(value: object) -> str:
    text = clean_text(value)
    return BLANK_DECORATION_LABEL if is_blank_decoration(text) else text


def is_blank_garment_product(
    product_name: object = "",
    original_line: object = "",
    style_number: object = "",
) -> bool:
    if normalize_style(style_number) in KNOWN_BLANK_GARMENT_STYLES:
        return True
    text = " ".join([clean_text(product_name), clean_text(original_line)]).casefold()
    # "Short sleeve shirt" is apparel, not a pair of shorts. Older builds used
    # the singular word "short" and incorrectly converted many tees and shirts
    # into blank garments. Keep only true lower-body garment matches here.
    text_without_short_sleeve = re.sub(r"\bshort[- ]sleeve\b", "sleeve", text)
    return any(
        re.search(rf"\b{re.escape(term)}\b", text_without_short_sleeve)
        for term in BLANK_GARMENT_TERMS
    )


def blank_garment_defaults(record: Mapping[str, object]) -> dict:
    updated = dict(record)
    product_name = clean_text(updated.get("Product Name", ""))
    original_line = clean_text(updated.get("Original Line Item", ""))
    style_number = clean_text(updated.get("Style Number", ""))
    decoration_type = clean_text(updated.get("Decoration Type", ""))

    if is_blank_garment_product(product_name, original_line, style_number) and not decoration_type:
        updated["Decoration Type"] = BLANK_DECORATION_LABEL
        updated["Decoration Color"] = ""
    elif is_blank_decoration(decoration_type):
        updated["Decoration Type"] = BLANK_DECORATION_LABEL
        updated["Decoration Color"] = ""
    return updated


def apply_blank_garment_defaults(frame: pd.DataFrame) -> pd.DataFrame:
    if frame.empty:
        return frame

    result = frame.copy()
    for column in ["Product Name", "Style Number", "Decoration Type", "Decoration Color", "Decoration Location", "Decoration Placement Instructions"]:
        if column not in result.columns:
            result[column] = ""

    blank_mask = result.apply(
        lambda row: is_blank_garment_product(
            row.get("Product Name", ""),
            row.get("Original Line Item", ""),
            row.get("Style Number", ""),
        ),
        axis=1,
    )
    missing_deco = result["Decoration Type"].map(clean_text).eq("")
    auto_mask = blank_mask & missing_deco
    if auto_mask.any():
        result.loc[auto_mask, "Decoration Type"] = BLANK_DECORATION_LABEL
        result.loc[auto_mask, "Decoration Color"] = ""
        result.loc[auto_mask, "Decoration Location"] = ""
        result.loc[auto_mask, "Decoration Placement Instructions"] = ""

    blank_deco_mask = result["Decoration Type"].map(is_blank_decoration)
    if blank_deco_mask.any():
        result.loc[blank_deco_mask, "Decoration Type"] = BLANK_DECORATION_LABEL
        result.loc[blank_deco_mask, "Decoration Color"] = ""
        result.loc[blank_deco_mask, "Decoration Location"] = ""
        result.loc[blank_deco_mask, "Decoration Placement Instructions"] = ""
    return result
