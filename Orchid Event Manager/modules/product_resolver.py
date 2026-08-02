from __future__ import annotations

from dataclasses import dataclass
from difflib import SequenceMatcher
from pathlib import Path
import re
import weakref

import pandas as pd

from modules.blank_garment_rules import (
    BLANK_DECORATION_LABEL,
    is_blank_decoration,
    apply_blank_garment_defaults,
)
from modules.decoration_locations import normalize_decoration_location
from modules.outsource_rules import (
    NEVER_OUTSOURCE_COLUMN, apply_never_outsource_defaults, apply_never_outsource_overrides,
    never_outsource_override_path, resolve_never_outsource, vendor_never_outsource,
)
from modules.purchase_rules import (
    RULE_COLUMNS,
    apply_purchase_rule_defaults,
    normalize_bool,
    row_rules,
)


_MASTER_CACHE_KEY: tuple[str, int, int, int, int] | None = None
_MASTER_CACHE_FRAME: pd.DataFrame | None = None
_LOOKUP_VERSION = 1
_MASTER_LOOKUPS: dict[int, tuple[weakref.ReferenceType, dict]] = {}

BASE_COLUMNS = [
    "Product Name",
    "Style Number",
    "Garment Color",
    "Vendor",
    "Decoration Type",
    "Decoration Location",
    "Decoration Placement Instructions",
    "Decoration Color",
]
ALIAS_COLUMNS = [
    "Product ID",
    "Product Aliases",
    "Vendor Color Code",
    "Color Aliases",
    "Purchasing Style Number",
]
MASTER_COLUMNS = BASE_COLUMNS + ALIAS_COLUMNS + RULE_COLUMNS + [NEVER_OUTSOURCE_COLUMN, "Setup Required"]


def clean(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value).strip())


def normalize_style(value: object) -> str:
    return re.sub(r"\s+", "", clean(value)).upper()


def normalize_phrase(value: object) -> str:
    text = clean(value).casefold()
    text = text.replace("®", " ").replace("™", " ").replace("&", " and ")
    text = text.replace("grey", "gray")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def split_aliases(value: object) -> list[str]:
    raw = clean(value)
    if not raw:
        return []
    return [clean(part) for part in re.split(r"[|;,\n]+", raw) if clean(part)]


def default_product_id(style: object, product: object) -> str:
    style_key = normalize_style(style)
    if style_key:
        return style_key
    phrase = normalize_phrase(product)
    return re.sub(r"\s+", "-", phrase).upper()[:48] or "UNASSIGNED"


def ensure_master_columns(frame: pd.DataFrame) -> pd.DataFrame:
    result = frame.copy()
    for column in MASTER_COLUMNS:
        if column not in result.columns:
            result[column] = ""
    result = result[MASTER_COLUMNS].fillna("")
    for column in MASTER_COLUMNS:
        result[column] = result[column].map(clean)
    missing_id = result["Product ID"].eq("")
    if missing_id.any():
        result.loc[missing_id, "Product ID"] = result.loc[missing_id].apply(
            lambda row: default_product_id(row.get("Style Number", ""), row.get("Product Name", "")),
            axis=1,
        )
    result = apply_never_outsource_defaults(apply_purchase_rule_defaults(result))
    return result


def _prepare_master_lookup(frame: pd.DataFrame) -> pd.DataFrame:
    """Build a read-only lookup index without adding columns or large attrs.

    A weak-reference cache keeps the index tied to the exact in-memory DataFrame.
    Temporary candidate DataFrames therefore stay small and fast to copy.
    """
    frame_id = id(frame)
    cached = _MASTER_LOOKUPS.get(frame_id)
    if cached is not None and cached[0]() is frame:
        return frame

    style_index: dict[str, list[object]] = {}
    exact_product_index: dict[str, list[object]] = {}
    product_terms_by_index: dict[object, tuple[str, ...]] = {}
    for index, row in frame.iterrows():
        style = normalize_style(row.get("Style Number", ""))
        if style:
            style_index.setdefault(style, []).append(index)
        terms = tuple(dict.fromkeys(
            normalized for normalized in (normalize_phrase(term) for term in _product_terms(row))
            if normalized
        ))
        product_terms_by_index[index] = terms
        for term in terms:
            exact_product_index.setdefault(term, []).append(index)

    lookup = {
        "version": _LOOKUP_VERSION,
        "style": style_index,
        "exact_product": exact_product_index,
        "product_terms": product_terms_by_index,
    }

    def _discard(_reference, key=frame_id):
        _MASTER_LOOKUPS.pop(key, None)

    _MASTER_LOOKUPS[frame_id] = (weakref.ref(frame, _discard), lookup)
    return frame


def _master_lookup(frame: pd.DataFrame) -> dict:
    _prepare_master_lookup(frame)
    cached = _MASTER_LOOKUPS.get(id(frame))
    return cached[1] if cached is not None and cached[0]() is frame else {}


def load_extended_master(path: Path) -> pd.DataFrame:
    global _MASTER_CACHE_KEY, _MASTER_CACHE_FRAME
    path = Path(path)
    override_path = never_outsource_override_path(path)
    if path.exists():
        stat = path.stat()
        if override_path.exists():
            override_stat = override_path.stat()
            override_mtime = int(override_stat.st_mtime_ns)
            override_size = int(override_stat.st_size)
        else:
            override_mtime = 0
            override_size = 0
        cache_key = (str(path.resolve()), int(stat.st_mtime_ns), int(stat.st_size), override_mtime, override_size)
    else:
        cache_key = (str(path.resolve()), 0, 0, 0, 0)
    if cache_key == _MASTER_CACHE_KEY and _MASTER_CACHE_FRAME is not None:
        returned = _MASTER_CACHE_FRAME.copy(deep=True)
        _prepare_master_lookup(returned)
        return returned
    try:
        frame = pd.read_csv(path, dtype=str).fillna("")
    except Exception:
        frame = pd.DataFrame(columns=MASTER_COLUMNS)
    result = ensure_master_columns(apply_purchase_rule_defaults(apply_blank_garment_defaults(frame)))
    result = apply_never_outsource_overrides(result, path)
    _prepare_master_lookup(result)
    _MASTER_CACHE_KEY = cache_key
    _MASTER_CACHE_FRAME = result.copy(deep=True)
    returned = result.copy(deep=True)
    _prepare_master_lookup(returned)
    return returned


def _color_terms(row: pd.Series) -> list[str]:
    terms = [clean(row.get("Garment Color", "")), clean(row.get("Vendor Color Code", ""))]
    terms.extend(split_aliases(row.get("Color Aliases", "")))
    return [term for term in terms if term]


def _product_terms(row: pd.Series) -> list[str]:
    terms = [clean(row.get("Product Name", "")), clean(row.get("Product ID", ""))]
    terms.extend(split_aliases(row.get("Product Aliases", "")))
    return [term for term in terms if term]


def _exact_term_match(value: object, terms: list[str]) -> bool:
    key = normalize_phrase(value)
    return bool(key) and any(key == normalize_phrase(term) for term in terms)


def _contains_term_match(value: object, terms: list[str]) -> bool:
    key = normalize_phrase(value)
    if not key:
        return False
    padded = f" {key} "
    return any(f" {normalize_phrase(term)} " in padded for term in terms if normalize_phrase(term))


def _similarity(value: object, term: object) -> float:
    left = normalize_phrase(value)
    right = normalize_phrase(term)
    if not left or not right:
        return 0.0
    return SequenceMatcher(None, left, right).ratio()


def _consistent_value(frame: pd.DataFrame, column: str) -> str:
    values = [clean(value) for value in frame[column] if clean(value)] if column in frame.columns else []
    unique = list(dict.fromkeys(values))
    return unique[0] if len(unique) == 1 else ""


def _best_color_row(candidates: pd.DataFrame, parsed_color: object) -> tuple[pd.Series | None, str]:
    if candidates.empty:
        return None, ""
    color = clean(parsed_color)
    if color:
        for _, row in candidates.iterrows():
            if _exact_term_match(color, _color_terms(row)):
                return row, "Exact style/color alias"
        for _, row in candidates.iterrows():
            if _contains_term_match(color, _color_terms(row)):
                return row, "Contained color alias"
        scored: list[tuple[float, int]] = []
        for index, row in candidates.iterrows():
            score = max((_similarity(color, term) for term in _color_terms(row)), default=0.0)
            scored.append((score, index))
        if scored:
            best_score, best_index = max(scored)
            if best_score >= 0.90:
                return candidates.loc[best_index], "Fuzzy color alias"
        return None, ""

    # Missing Shopify color: safely fill only when the style has one purchasing color.
    distinct_rows = candidates.copy()
    distinct_rows["_purchase_color"] = distinct_rows.apply(
        lambda row: clean(row.get("Vendor Color Code", "")) or clean(row.get("Garment Color", "")),
        axis=1,
    )
    unique_colors = [value for value in distinct_rows["_purchase_color"].unique() if clean(value)]
    if len(unique_colors) == 1:
        row = distinct_rows[distinct_rows["_purchase_color"].eq(unique_colors[0])].iloc[0]
        return row, "Single-color style inference"
    if len(candidates) == 1:
        return candidates.iloc[0], "Single-row style inference"
    return None, ""


@dataclass
class ResolveResult:
    matched: bool
    match_type: str
    product_id: str
    style: str
    product_name: str
    product_aliases: str
    garment_color: str
    vendor: str
    decoration_type: str
    decoration_location: str
    decoration_placement_instructions: str
    decoration_color: str
    master_color: str
    vendor_color_code: str
    purchasing_style_number: str
    product_category: str
    requires_size: bool
    requires_color: bool
    requires_decoration: bool
    never_outsource: bool
    issue: str

    def as_dict(self) -> dict:
        return {
            "Master Match": self.match_type if self.matched else "No",
            "Master Product ID": self.product_id,
            "Master Style Number": self.style,
            "Master Product Name": self.product_name,
            "Product Aliases": self.product_aliases,
            "Master Garment Color": self.master_color,
            "Master Vendor Color Code": self.vendor_color_code,
            # This is the supplier's orderable style for the matched color.
            # Shopify may use one friendly style for every color, while the
            # vendor assigns each color its own ordering number.
            "Vendor Product #": self.purchasing_style_number,
            "Product Category": self.product_category,
            "Requires Size": "Yes" if self.requires_size else "No",
            "Requires Color": "Yes" if self.requires_color else "No",
            "Requires Decoration": "Yes" if self.requires_decoration else "No",
            NEVER_OUTSOURCE_COLUMN: "Yes" if self.never_outsource else "No",
            "Resolved Garment Color": self.garment_color,
            "Vendor": self.vendor,
            "Decoration Type": self.decoration_type,
            "Decoration Location": self.decoration_location,
            "Decoration Placement Instructions": self.decoration_placement_instructions,
            "Decoration Color": self.decoration_color,
            "Resolver Issue": self.issue,
        }


def resolve_product(
    parsed_style: object,
    parsed_product: object,
    parsed_color: object,
    master: pd.DataFrame,
    master_prepared: bool = False,
) -> ResolveResult:
    frame = master if master_prepared else ensure_master_columns(master)
    style_key = normalize_style(parsed_style)
    product_key = normalize_phrase(parsed_product)
    parsed_color_text = clean(parsed_color)

    candidates = pd.DataFrame(columns=frame.columns)
    base_match_type = ""
    lookup = _master_lookup(frame)
    if style_key:
        indexes = lookup.get("style", {}).get(style_key, [])
        if indexes:
            candidates = frame.loc[indexes].copy()
            base_match_type = "Style"

    if candidates.empty and product_key:
        indexes = lookup.get("exact_product", {}).get(product_key, [])
        if indexes:
            candidates = frame.loc[indexes].copy()
            base_match_type = "Product alias"

    if candidates.empty and product_key:
        padded = f" {product_key} "
        indexes = [
            index for index, terms in lookup.get("product_terms", {}).items()
            if any(f" {term} " in padded for term in terms)
        ]
        if len(indexes) == 1:
            candidates = frame.loc[indexes].copy()
            base_match_type = "Contained product alias"
        elif len(indexes) > 1:
            candidates = pd.DataFrame(columns=frame.columns)

    if candidates.empty:
        inferred = row_rules({
            "Product Name": clean(parsed_product),
            "Style Number": style_key,
            "Decoration Type": "",
        })
        return ResolveResult(
            False, "No", "", style_key, clean(parsed_product), "", parsed_color_text,
            "", "", "", "", "", "", "", "",
            inferred["Product Category"],
            normalize_bool(inferred["Requires Size"], True),
            normalize_bool(inferred["Requires Color"], True),
            normalize_bool(inferred["Requires Decoration"], True),
            resolve_never_outsource("", clean(parsed_product), inferred["Product Category"], style_key),
            "No Product Master match",
        )

    color_row, color_match_type = _best_color_row(candidates, parsed_color_text)
    routing_source = color_row if color_row is not None else candidates.iloc[0]
    # Setup Required is color-row specific. A blank or unfinished unrelated color
    # must not block a complete color that is actually present on the order.
    setup_required = clean(routing_source.get("Setup Required", "")).casefold() in {"yes", "y", "true", "1"}
    rules = row_rules(routing_source)
    product_category = rules["Product Category"]
    requires_size = normalize_bool(rules["Requires Size"], True)
    requires_color = normalize_bool(rules["Requires Color"], True)
    requires_decoration = normalize_bool(rules["Requires Decoration"], True)
    vendor = clean(routing_source.get("Vendor", "")) or _consistent_value(candidates, "Vendor")
    deco_type = clean(routing_source.get("Decoration Type", "")) or _consistent_value(candidates, "Decoration Type")
    deco_location = clean(routing_source.get("Decoration Location", "")) or _consistent_value(candidates, "Decoration Location")
    deco_placement = clean(routing_source.get("Decoration Placement Instructions", "")) or _consistent_value(candidates, "Decoration Placement Instructions")
    deco_color = clean(routing_source.get("Decoration Color", ""))
    if not deco_color:
        deco_color = _consistent_value(candidates, "Decoration Color")
    if not requires_decoration and not deco_type:
        deco_type = BLANK_DECORATION_LABEL
    deco_location = normalize_decoration_location(
        deco_location, deco_type, requires_decoration=requires_decoration
    )
    if is_blank_decoration(deco_type) or not requires_decoration:
        deco_color = ""

    product_name = clean(routing_source.get("Product Name", "")) or clean(parsed_product)
    product_aliases = clean(routing_source.get("Product Aliases", "")) or _consistent_value(candidates, "Product Aliases")
    master_style = clean(routing_source.get("Style Number", "")) or style_key
    product_id = clean(routing_source.get("Product ID", "")) or default_product_id(master_style, product_name)
    master_color = clean(routing_source.get("Garment Color", ""))
    vendor_code = clean(routing_source.get("Vendor Color Code", ""))
    purchasing_style_number = clean(routing_source.get("Purchasing Style Number", ""))
    resolved_color = parsed_color_text
    issue = "Product Master setup required" if setup_required else ""

    if color_row is not None:
        # Preserve the human-friendly master color; use vendor code only when the master color is absent.
        resolved_color = master_color or vendor_code or parsed_color_text
    elif parsed_color_text:
        # Style is known and Shopify supplied a usable purchasing color. Route by style-level defaults.
        resolved_color = parsed_color_text
        color_match_type = "Style routing with Shopify color"
        if deco_type and not is_blank_decoration(deco_type) and not deco_color:
            issue = "; ".join(filter(None, [issue, "Color is not mapped to a decoration color"]))
    else:
        if requires_color:
            issue = "; ".join(filter(None, [issue, "Missing garment color"]))

    match_type = base_match_type
    if color_match_type:
        match_type = f"{base_match_type} - {color_match_type}"
    elif base_match_type:
        match_type = f"{base_match_type} routing"

    return ResolveResult(
        True,
        match_type,
        product_id,
        master_style,
        product_name,
        product_aliases,
        resolved_color,
        vendor,
        deco_type,
        deco_location,
        deco_placement,
        deco_color,
        master_color,
        vendor_code,
        purchasing_style_number,
        product_category,
        requires_size,
        requires_color,
        requires_decoration,
        (
            resolve_never_outsource(
                routing_source.get(NEVER_OUTSOURCE_COLUMN, ""), product_name, product_category, master_style
            )
            or vendor_never_outsource(vendor)
        ),
        issue,
    )
