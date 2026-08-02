from __future__ import annotations

import re
import shutil
from datetime import datetime
from pathlib import Path
import json
import os
import subprocess
import sys
import traceback
from tkinter import messagebox, simpledialog
from PIL import Image as PILImage

import customtkinter as ctk
import pandas as pd

from modules.paths import product_master_path, product_master_update_marker_path
from modules.outsource_rules import (
    NEVER_OUTSOURCE_COLUMN, apply_never_outsource_defaults, default_never_outsource,
    is_boots_product, vendor_never_outsource,
)
from modules.product_resolver import default_product_id
from modules.blank_garment_rules import (
    BLANK_DECORATION_LABEL,
    apply_blank_garment_defaults,
    is_blank_decoration,
    is_blank_garment_product,
    normalize_decoration_type,
)
from modules.scroll_support import install_native_scroll_support
from modules.decoration_locations import (
    DECORATION_LOCATIONS,
    STANDARD_DECORATION_LOCATIONS,
    LEFT_CHEST,
    HAT,
    OTHER_CUSTOM,
    NOT_APPLICABLE_IN_HOUSE,
    NOT_APPLICABLE_NO_DECORATION,
    default_decoration_location,
    location_choice_and_custom,
    normalize_decoration_location,
    resolve_location_choice,
)
from modules.product_intelligence import apply_product_intelligence, infer_default_decoration
from modules.routing_rules import preferred_vendor_for_brand_text
from modules.shopify_parser import is_purchasing_color
from modules.thread_ink_colors import (
    load_thread_ink_colors, register_thread_ink_color, save_thread_ink_colors,
)
from modules.color_setup_guidance import (
    is_new_color_setup_row,
    is_setup_required,
    new_color_setup_guidance,
    new_color_setup_message,
    prioritize_new_color_setup_rows,
)
from modules.internal_services import (
    HEMMING_ALTERATION_LABEL, SEW_ON_PATCH_LABEL,
    infer_in_house_decoration, is_in_house_decoration, is_in_house_service_product,
    is_internal_service_style, is_phantom_product_label,
)
from modules.purchase_rules import (
    CATEGORY_OPTIONS,
    RULE_COLUMNS,
    apply_purchase_rule_defaults,
    bool_text,
    category_defaults,
    infer_category,
    normalize_bool,
    row_rules,
)

PRODUCT_MASTER_PATH = product_master_path()
PRODUCT_MASTER_PID_FILE = PRODUCT_MASTER_PATH.parent / "product_master.pid"
PRODUCT_MASTER_REQUEST_FILE = PRODUCT_MASTER_PATH.parent / "product_master_open_request.json"
PRODUCT_MASTER_UPDATE_MARKER = product_master_update_marker_path()
NEVER_OUTSOURCE_OVERRIDE_FILE = PRODUCT_MASTER_PATH.parent / "never_outsource_overrides.json"
COLUMNS = [
    "Product Name",
    "Style Number",
    "Garment Color",
    "Vendor",
    "Decoration Type",
    "Decoration Location",
    "Decoration Placement Instructions",
    "Decoration Color",
    "Product ID",
    "Product Aliases",
    "Vendor Color Code",
    "Color Aliases",
    "Purchasing Style Number",
    "Product Category",
    "Requires Size",
    "Requires Color",
    "Requires Decoration",
    NEVER_OUTSOURCE_COLUMN,
    "Setup Required",
]

WORKBOOK_VENDORS = [
    "Burnside",
    "Cutter & Buck",
    "Outdoor Cap",
    "Richardson",
    "S&S Activewear",
    "SanMar",
    "Tru-Spec",
    "VF",
    "Wrangler",
]
# A vendor added in Product Master is an Orchid-wide purchasing choice, not a
# temporary value for whichever style happens to be open.  Keep Rocky Boots
# canonical as well, so prior entries such as "rocky boots" and "Rocky-Boots"
# do not split its styles into separate vendor groups.
CUSTOM_VENDOR_CANONICAL_NAMES = {
    "rockyboot": "Rocky Boots",
    "rockyboots": "Rocky Boots",
}
PREFERRED_VENDOR_CASE = {vendor.casefold(): vendor for vendor in WORKBOOK_VENDORS}
PREFERRED_VENDOR_CASE.update({value.casefold(): value for value in CUSTOM_VENDOR_CANONICAL_NAMES.values()})
PURCHASE_VENDOR_REGISTRY_FILE = PRODUCT_MASTER_PATH.parent / "purchase_vendor_options.json"
DECORATION_TYPES = ["", "Embroidery", "Screen Print", SEW_ON_PATCH_LABEL, HEMMING_ALTERATION_LABEL, BLANK_DECORATION_LABEL]

PURPLE = "#5B2AA8"
PURPLE_DARK = "#3D176F"
PURPLE_LIGHT = "#F2ECFB"
PURPLE_BORDER = "#D7C8EE"
KEY_VENDOR_MAPPING_BG = "#F7F0FF"
KEY_VENDOR_MAPPING_BADGE_BG = "#E9DDFB"
TEXT_DARK = "#201A2D"
TEXT_MUTED = "#6F667A"
CARD_BG = "#FFFFFF"
WINDOW_BG = "#F7F5FA"
ROW_ALT = "#FAF7FE"
SUCCESS = "#2D7A46"
WARNING = "#B26A00"
DANGER = "#B42318"
MAX_DASHBOARD_ROWS = 60

# The editor no longer displays a second Product Catalog beside the selected
# product.  That reclaimed width lets every purchasing-color control stay on
# one easy-to-read left-to-right line.  The label and field share the exact
# same grid column, so they cannot drift out of alignment.
COLOR_TABLE_COLUMNS = (
    ("Purchasing Color", 150),
    ("Vendor Color Code", 135),
    ("Shopify Color Alias", 180),
    ("Purchase As Style #", 150),
    ("Thread / Ink", 140),
)
COLOR_TABLE_ACTION_WIDTH = 76
PURCHASE_AS_STYLE_COLUMN = 3


def configure_color_table_columns(widget):
    """Set the shared full-width columns for one purchasing-color mapping."""
    for column, (_label, minimum) in enumerate(COLOR_TABLE_COLUMNS):
        widget.grid_columnconfigure(column, weight=1, minsize=minimum)
    widget.grid_columnconfigure(len(COLOR_TABLE_COLUMNS), weight=0, minsize=COLOR_TABLE_ACTION_WIDTH)




def preserve_previous_color_alias(previous_color: object, new_color: object, aliases: object) -> str:
    """Keep the imported Shopify color searchable when a purchasing color is corrected.

    Product Master's Garment Color is the purchasing color. When the user replaces an
    imported Shopify color with the vendor purchasing color, the original value must
    remain in Color Aliases. Otherwise replacing an event with the same CSV can appear
    to introduce a new color and incorrectly return a completed product to Needs Setup.
    """
    previous = normalize_space(previous_color)
    current = normalize_space(new_color)
    values = [normalize_space(value) for value in re.split(r"[|;,\n]+", normalize_space(aliases)) if normalize_space(value)]
    seen = {value.casefold() for value in values}
    if previous and current and previous.casefold() != current.casefold() and previous.casefold() not in seen:
        values.append(previous)
    return " | ".join(values)

def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_space(value):
    return re.sub(r"\s+", " ", clean_text(value)).strip()


def vendor_identity_key(value):
    """Return a punctuation-insensitive key for matching vendor name variants."""
    value = normalize_space(value).casefold()
    return re.sub(r"[^a-z0-9]+", "", value)


def canonical_vendor_name(value):
    """Return one consistent spelling for a vendor, ignoring case, spaces, and punctuation."""
    value = normalize_space(value)
    if not value:
        return ""
    preferred_by_identity = {vendor_identity_key(vendor): vendor for vendor in WORKBOOK_VENDORS}
    preferred_by_identity.update(CUSTOM_VENDOR_CANONICAL_NAMES)
    return preferred_by_identity.get(vendor_identity_key(value), PREFERRED_VENDOR_CASE.get(value.casefold(), value))


def alphabetical_options(values, *, include_blank=False):
    """Return case-insensitive, duplicate-free dropdown values in alphabetical order."""
    registry = {}
    for raw in values:
        value = normalize_space(raw)
        if value:
            registry.setdefault(value.casefold(), value)
    ordered = sorted(registry.values(), key=str.casefold)
    return ([""] + ordered) if include_blank else ordered


def append_vendor_option(options, value):
    value = canonical_vendor_name(value)
    if value and not any(existing.casefold() == value.casefold() for existing in options):
        options.append(value)
    options[:] = alphabetical_options(options)
    return value


def canonicalize_vendor_values(values):
    """Normalize all vendor spellings case-insensitively, including custom vendors."""
    registry = dict(PREFERRED_VENDOR_CASE)
    result = []
    for value in values:
        value = canonical_vendor_name(value)
        if not value:
            result.append("")
            continue
        key = value.casefold()
        canonical = registry.get(key)
        if canonical is None:
            registry[key] = value
            canonical = value
        result.append(canonical)
    return result


def _load_purchase_vendor_registry() -> list[str]:
    """Read custom purchase vendors saved independently of any one style.

    Before this registry existed, ``+ Add Vendor`` only changed the dropdown in
    the current Product Master window.  Closing that window before a style save
    made a custom vendor appear to vanish.  The registry is deliberately small
    and stores only the permanent dropdown choices; vendor assignments continue
    to live on the Product Master styles themselves.
    """
    try:
        payload = json.loads(PURCHASE_VENDOR_REGISTRY_FILE.read_text(encoding="utf-8"))
    except (FileNotFoundError, OSError, TypeError, json.JSONDecodeError):
        return []
    values = payload.get("vendors", []) if isinstance(payload, dict) else payload
    if not isinstance(values, list):
        return []
    return canonicalize_vendor_values(values)


def _save_purchase_vendor_registry(values) -> list[str]:
    """Persist canonical custom vendor options without touching Product Master."""
    saved = alphabetical_options(canonicalize_vendor_values(values))
    try:
        PURCHASE_VENDOR_REGISTRY_FILE.parent.mkdir(parents=True, exist_ok=True)
        temporary = PURCHASE_VENDOR_REGISTRY_FILE.with_suffix(".tmp")
        temporary.write_text(
            json.dumps({"vendors": saved}, indent=2, sort_keys=True) + "\n",
            encoding="utf-8",
        )
        temporary.replace(PURCHASE_VENDOR_REGISTRY_FILE)
    except OSError:
        # A vendor assignment saved on a product remains authoritative even if
        # this convenience registry cannot be written on a particular launch.
        pass
    return saved


def purchase_vendor_options(master=None) -> list[str]:
    """Return all permanent vendor choices, including saved custom vendors."""
    saved_custom = _load_purchase_vendor_registry()
    master_vendors = []
    if master is not None and "Vendor" in master.columns:
        master_vendors = master["Vendor"].tolist()
        # Seed the new registry from pre-v4.9.75 Product Master assignments.
        # This is what carries the user's already-assigned Rocky Boots styles
        # forward as a permanent dropdown choice after the first upgraded open.
        if master_vendors:
            saved_custom = _save_purchase_vendor_registry([*saved_custom, *master_vendors])
    values = list(WORKBOOK_VENDORS) + saved_custom + master_vendors
    return alphabetical_options(canonicalize_vendor_values(values))


def register_purchase_vendor(value) -> str:
    """Add one permanent dropdown vendor and return its canonical name."""
    value = canonical_vendor_name(value)
    if not value:
        return ""
    _save_purchase_vendor_registry([*_load_purchase_vendor_registry(), value])
    return value


def normalize_style(value):
    return re.sub(r"\s+", "", clean_text(value)).upper()


def normalize_color(value):
    return normalize_space(value).casefold()


def normalize_search_text(value):
    """Normalize punctuation and spacing so style/brand searches are forgiving."""
    value = normalize_space(value).casefold()
    value = value.replace("™", " ").replace("®", " ").replace("+", " and ").replace("&", " and ")
    return re.sub(r"[^a-z0-9]+", " ", value).strip()


def search_tokens(value):
    return [token for token in normalize_search_text(value).split() if token]


def search_match_score_values(style, product_id, name, vendor, combined, query):
    """Return a lower-is-better search score for pre-normalized catalog values."""
    tokens = search_tokens(query)
    if not tokens:
        return 50
    if not all(token in combined for token in tokens):
        return None
    normalized_query = " ".join(tokens)
    if normalized_query == style:
        return 0
    if normalized_query == product_id:
        return 1
    if style.startswith(normalized_query):
        return 2
    if product_id.startswith(normalized_query):
        return 3
    if name.startswith(normalized_query):
        return 4
    if normalized_query in name:
        return 5
    if normalized_query in vendor:
        return 6
    return 10


def search_match_score(rows, query):
    """Return a lower-is-better search score, or None when the row does not match."""
    return search_match_score_values(
        normalize_search_text(first_nonblank(rows.get("Style Number", []))),
        normalize_search_text(first_nonblank(rows.get("Product ID", []))),
        normalize_search_text(best_product_name(rows.get("Product Name", []))),
        normalize_search_text(first_nonblank(rows.get("Vendor", []))),
        normalize_search_text(" ".join(clean_text(value) for value in rows[COLUMNS].to_numpy().flatten())),
        query,
    )


def canonical_product_name(value):
    name = normalize_space(value)
    return re.sub(r"[\s\.\-–—:|/]+$", "", name).strip()


def first_nonblank(values):
    for value in values:
        value = clean_text(value)
        if value:
            return value
    return ""


def best_product_name(values):
    cleaned = [canonical_product_name(value) for value in values]
    cleaned = [value for value in cleaned if value]
    if not cleaned:
        return ""
    counts = {}
    for value in cleaned:
        counts[value] = counts.get(value, 0) + 1
    return sorted(counts, key=lambda value: (-counts[value], len(value), value.casefold()))[0]


def catalog_key(row):
    style = normalize_style(row["Style Number"])
    color = normalize_color(row["Garment Color"])
    if style:
        return f"style:{style}|color:{color}"
    return f"product:{canonical_product_name(row['Product Name']).casefold()}|color:{color}"


def style_group_key(row):
    style = normalize_style(row["Style Number"])
    if style:
        return f"style:{style}"
    return f"product:{canonical_product_name(row['Product Name']).casefold()}"


def clear_legacy_size_like_purchasing_colors(master):
    """Remove values that Orchid already knows cannot be garment colors.

    Versions before 4.9.67 could save a Shopify size or footwear-width marker
    in Product Master's Garment Color column.  The current importer rejects
    those values, but existing catalog rows are also read whenever the editor
    opens.  Leaving a stale value such as ``15`` there makes it look like a
    new purchasing color and can create an unnecessary thread/ink decision.

    This migration is intentionally narrow: it clears only values rejected by
    ``is_purchasing_color``.  Numeric vendor color *codes* are not touched,
    because a vendor may legitimately use a numeric code.
    """
    if master.empty:
        return master

    result = master.copy()
    if "Garment Color" not in result.columns:
        result["Garment Color"] = ""
    if "Color Aliases" not in result.columns:
        result["Color Aliases"] = ""

    for index, raw_color in result["Garment Color"].items():
        color = normalize_space(raw_color)
        if not color or is_purchasing_color(color):
            continue

        # The Product Master is not the place to preserve an order's size.
        # The original Shopify line retains the size for Purchase Review and
        # reporting; Product Master keeps a blank base row for no-color items.
        result.at[index, "Garment Color"] = ""

        aliases = []
        seen = set()
        for alias in re.split(r"[|;,\n]+", normalize_space(result.at[index, "Color Aliases"])):
            alias = normalize_space(alias)
            if not alias or not is_purchasing_color(alias):
                continue
            key = alias.casefold()
            if key not in seen:
                aliases.append(alias)
                seen.add(key)
        result.at[index, "Color Aliases"] = " | ".join(aliases)

    return result


def settle_complete_boot_setup_rows(master):
    """Clear the false setup flag left by a discarded boot size color.

    Boots require a size but no garment color or decoration.  Once their base
    details are present, a stale `Setup Required` flag created only by a bogus
    size-color row should not keep the style in the Needs Setup queue.
    """
    if master.empty:
        return master

    result = master.copy()
    for index, row in result.iterrows():
        if not is_setup_required(row.get("Setup Required", "")):
            continue
        if not is_boots_product(
            row.get("Product Name", ""), row.get("Product Category", ""), row.get("Style Number", "")
        ):
            continue
        rules = row_rules(row)
        base_details_complete = all(
            clean_text(row.get(column, ""))
            for column in ("Product Name", "Style Number", "Vendor", "Product Category")
        )
        if (
            base_details_complete
            and not normalize_bool(rules.get("Requires Color", ""), True)
            and not normalize_bool(rules.get("Requires Decoration", ""), True)
        ):
            result.at[index, "Setup Required"] = "No"
    return result


def remove_blank_color_placeholders(master):
    """Remove a stale blank placeholder once a color-required style has real colors."""
    if master.empty:
        return master
    result = master.copy()
    result["_style_group"] = result.apply(style_group_key, axis=1)
    drop_indices = []
    shared_columns = [
        "Product Name", "Style Number", "Vendor", "Decoration Type",
        "Decoration Location", "Decoration Placement Instructions", "Product ID",
        "Product Aliases", "Product Category", "Requires Size", "Requires Color",
        "Requires Decoration", NEVER_OUTSOURCE_COLUMN,
    ]
    for _, group in result.groupby("_style_group", sort=False, dropna=False):
        blank_mask = group["Garment Color"].astype(str).str.strip().eq("")
        if not blank_mask.any() or blank_mask.all():
            continue
        requires_color = normalize_bool(first_nonblank(group["Requires Color"]), True)
        if not requires_color:
            continue
        blank_rows = group[blank_mask]
        colored_indices = group[~blank_mask].index
        for column in shared_columns:
            fallback = first_nonblank(blank_rows[column])
            if not fallback:
                continue
            for index in colored_indices:
                if not clean_text(result.at[index, column]):
                    result.at[index, column] = fallback
        drop_indices.extend(blank_rows.index.tolist())
    if drop_indices:
        result = result.drop(index=drop_indices)
    return result.drop(columns=["_style_group"]).reset_index(drop=True)


def remove_internal_service_product_records(master):
    """Keep decoration charges out of the permanent purchasing catalog.

    A Shopify decoration service is a valid sales/order-total line, but Orchid
    never buys it from a garment vendor.  It therefore must not survive as a
    Product Master setup item.  ``load_product_master`` already makes a dated
    backup before writing a cleanup, so older service rows remain recoverable
    without cluttering the live purchasing catalog.
    """
    if master.empty:
        return master
    service_mask = master.apply(
        lambda row: is_in_house_service_product(
            row.get("Product Name", ""),
            row.get("Original Line Item", ""),
            row.get("Decoration Type", ""),
            style_number=row.get("Style Number", ""),
            garment_color=row.get("Garment Color", ""),
        ) or is_phantom_product_label(row.get("Product Name", ""))
          or is_phantom_product_label(row.get("Style Number", "")),
        axis=1,
    )
    return master.loc[~service_mask].copy().reset_index(drop=True)


def clean_and_deduplicate_master(master):
    master = master.copy()
    for column in COLUMNS:
        if column not in master.columns:
            master[column] = ""
    master = master[COLUMNS].fillna("")
    for column in COLUMNS:
        master[column] = master[column].map(clean_text)
    master["Product Name"] = master["Product Name"].map(canonical_product_name)
    master["Style Number"] = master["Style Number"].map(normalize_style)
    master["Garment Color"] = master["Garment Color"].map(normalize_space)
    master = clear_legacy_size_like_purchasing_colors(master)
    # A decoration fee may have been materialized by an older import.  Remove
    # it on upgrade so it cannot remain in All Products or Needs Setup.
    master = remove_internal_service_product_records(master)
    # Repair duplicate spellings such as SanMar / Sanmar before saving or routing.
    master["Vendor"] = canonicalize_vendor_values(master["Vendor"])
    if master.empty:
        return pd.DataFrame(columns=COLUMNS)

    master["_key"] = master.apply(catalog_key, axis=1)
    rows = []
    for _, group in master.groupby("_key", sort=False, dropna=False):
        rows.append(
            {
                "Product Name": best_product_name(group["Product Name"]),
                "Style Number": first_nonblank(group["Style Number"]),
                "Garment Color": first_nonblank(group["Garment Color"]),
                "Vendor": first_nonblank(group["Vendor"]),
                "Decoration Type": first_nonblank(group["Decoration Type"]),
                "Decoration Location": first_nonblank(group["Decoration Location"]),
                "Decoration Placement Instructions": first_nonblank(group["Decoration Placement Instructions"]),
                "Decoration Color": first_nonblank(group["Decoration Color"]),
                "Product ID": first_nonblank(group["Product ID"]),
                "Product Aliases": first_nonblank(group["Product Aliases"]),
                "Vendor Color Code": first_nonblank(group["Vendor Color Code"]),
                "Color Aliases": first_nonblank(group["Color Aliases"]),
                "Purchasing Style Number": first_nonblank(group["Purchasing Style Number"]),
                "Product Category": first_nonblank(group["Product Category"]),
                "Requires Size": first_nonblank(group["Requires Size"]),
                "Requires Color": first_nonblank(group["Requires Color"]),
                "Requires Decoration": first_nonblank(group["Requires Decoration"]),
                NEVER_OUTSOURCE_COLUMN: first_nonblank(group[NEVER_OUTSOURCE_COLUMN]),
                "Setup Required": first_nonblank(group["Setup Required"]),
            }
        )
    result = remove_blank_color_placeholders(pd.DataFrame(rows, columns=COLUMNS))
    if not result.empty:
        missing_id = result["Product ID"].astype(str).str.strip().eq("")
        result.loc[missing_id, "Product ID"] = result.loc[missing_id].apply(
            lambda row: default_product_id(row.get("Style Number", ""), row.get("Product Name", "")),
            axis=1,
        )
    result = apply_never_outsource_defaults(apply_product_intelligence(apply_purchase_rule_defaults(apply_blank_garment_defaults(result))))
    result = settle_complete_boot_setup_rows(result)
    return result.sort_values(
        by=["Style Number", "Product Name", "Garment Color"],
        key=lambda series: series.astype(str).str.casefold(),
    ).reset_index(drop=True)


def clean_changed_style_groups(master, style_keys):
    """Clean only edited product groups after the catalog was cleaned on load.

    Normal Product Master saves change one style at a time. Re-cleaning every
    historical style on each Save made a large, already-valid catalog feel slow.
    The unchanged portion has already passed ``clean_and_deduplicate_master``
    when it was opened; this function runs the same rules on the edited style
    (and a newly merged style, if the product number changed) before writing the
    complete, sorted file back to disk.
    """
    keys = {clean_text(key) for key in style_keys if clean_text(key)}
    if not keys:
        return clean_and_deduplicate_master(master)
    working = master.copy()
    for column in COLUMNS:
        if column not in working.columns:
            working[column] = ""
    working = working[COLUMNS].fillna("")
    group_keys = working.apply(style_group_key, axis=1)
    changed_mask = group_keys.isin(keys)
    if not changed_mask.any():
        return clean_and_deduplicate_master(working)
    unchanged = working.loc[~changed_mask, COLUMNS].copy()
    changed = clean_and_deduplicate_master(working.loc[changed_mask, COLUMNS])
    combined = pd.concat([unchanged, changed], ignore_index=True)
    return combined.sort_values(
        by=["Style Number", "Product Name", "Garment Color"],
        key=lambda series: series.astype(str).str.casefold(),
    ).reset_index(drop=True)


def product_master_file_stamp():
    try:
        stat = PRODUCT_MASTER_PATH.stat()
        return (int(stat.st_mtime_ns), int(stat.st_size))
    except OSError:
        return (0, 0)


def _atomic_write_product_master(frame: pd.DataFrame) -> None:
    """Write Product Master transactionally so interruption cannot corrupt it."""
    PRODUCT_MASTER_PATH.parent.mkdir(parents=True, exist_ok=True)
    temporary = PRODUCT_MASTER_PATH.with_name(
        f".{PRODUCT_MASTER_PATH.name}.{os.getpid()}.tmp"
    )
    try:
        with temporary.open("w", encoding="utf-8", newline="") as handle:
            frame.to_csv(handle, index=False)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temporary, PRODUCT_MASTER_PATH)
    finally:
        try:
            temporary.unlink(missing_ok=True)
        except OSError:
            pass


def load_product_master():
    PRODUCT_MASTER_PATH.parent.mkdir(parents=True, exist_ok=True)
    if not PRODUCT_MASTER_PATH.exists():
        _atomic_write_product_master(pd.DataFrame(columns=COLUMNS))
    try:
        master = pd.read_csv(PRODUCT_MASTER_PATH, dtype=str).fillna("")
    except pd.errors.EmptyDataError:
        master = pd.DataFrame(columns=COLUMNS)
    except Exception as error:
        messagebox.showerror("Unable to Open Product Master", str(error))
        return pd.DataFrame(columns=COLUMNS)
    cleaned = clean_and_deduplicate_master(master)
    comparable_original = master.copy()
    for column in COLUMNS:
        if column not in comparable_original.columns:
            comparable_original[column] = ""
    comparable_original = comparable_original[COLUMNS].fillna("")
    for column in COLUMNS:
        comparable_original[column] = comparable_original[column].map(clean_text)
    changed = not comparable_original.reset_index(drop=True).equals(cleaned.reset_index(drop=True))
    if changed and PRODUCT_MASTER_PATH.exists():
        backup_dir = PRODUCT_MASTER_PATH.parent / "backups"
        backup_dir.mkdir(parents=True, exist_ok=True)
        migration_backup = backup_dir / f"product_master_before_v50_purchase_rules_{datetime.now():%Y%m%d_%H%M%S}.csv"
        shutil.copy2(PRODUCT_MASTER_PATH, migration_backup)
        _atomic_write_product_master(cleaned)
    # Do not rewrite an unchanged catalog merely because it was opened. Besides
    # saving time, this keeps file stamps stable so the main app can reuse caches.
    return cleaned


def save_product_master(master, *, expected_style="", expected_vendor="", affected_style_keys=None):
    """Save the live Product Master and verify critical values from disk.

    Product Master saves must not report success until the exact live CSV can be
    re-opened and the requested vendor is present for the saved style. This
    catches cleanup rules, path drift, and failed writes before the user closes
    the application.
    """
    try:
        cleaned = (
            clean_changed_style_groups(master, affected_style_keys)
            if affected_style_keys is not None
            else clean_and_deduplicate_master(master)
        )
        style_key = normalize_style(expected_style)
        vendor_key = canonical_vendor_name(expected_vendor)
        if style_key and vendor_key:
            candidate = cleaned[cleaned["Style Number"].map(normalize_style).eq(style_key)]
            saved_vendors = {canonical_vendor_name(value) for value in candidate["Vendor"].tolist()}
            if not candidate.empty and vendor_key not in saved_vendors:
                raise ValueError(
                    f"Save verification stopped before writing: {style_key} would not retain vendor {vendor_key}."
                )
        _atomic_write_product_master(cleaned)

        # Re-open the physical live file instead of trusting the in-memory frame.
        disk = pd.read_csv(PRODUCT_MASTER_PATH, dtype=str).fillna("")
        if style_key and vendor_key:
            candidate = disk[disk["Style Number"].map(normalize_style).eq(style_key)]
            disk_vendors = {canonical_vendor_name(value) for value in candidate["Vendor"].tolist()}
            if candidate.empty or vendor_key not in disk_vendors:
                raise IOError(
                    f"Save verification failed after writing. {style_key} was not reloaded from "
                    f"{PRODUCT_MASTER_PATH} with vendor {vendor_key}."
                )
        # ``cleaned`` was the exact frame written above. Returning the physical
        # file avoids doing a second whole-catalog cleanup after every save while
        # still proving the saved vendor can be read back from disk.
        for column in COLUMNS:
            if column not in disk.columns:
                disk[column] = ""
        return disk[COLUMNS].fillna("")
    except Exception as error:
        messagebox.showerror("Unable to Save Product Master", str(error))
        return None


def load_never_outsource_overrides():
    """Load only deliberate user choices, separate from legacy generated defaults."""
    try:
        payload = json.loads(NEVER_OUTSOURCE_OVERRIDE_FILE.read_text(encoding="utf-8"))
        if isinstance(payload, dict):
            return {str(key): bool(value) for key, value in payload.items()}
    except (FileNotFoundError, json.JSONDecodeError, OSError, TypeError):
        pass
    return {}


def save_never_outsource_overrides(overrides):
    try:
        NEVER_OUTSOURCE_OVERRIDE_FILE.parent.mkdir(parents=True, exist_ok=True)
        temporary = NEVER_OUTSOURCE_OVERRIDE_FILE.with_suffix(".tmp")
        temporary.write_text(json.dumps(overrides, indent=2, sort_keys=True), encoding="utf-8")
        temporary.replace(NEVER_OUTSOURCE_OVERRIDE_FILE)
        return True
    except OSError:
        return False


def editor_never_outsource_value(key, overrides, saved_value, product_name, category, style_number, vendor=""):
    """Resolve the editor switch while distinguishing legacy No from a manual override."""
    # Boots are never allowed to route to an outside decorator.  Check this
    # before the legacy override sidecar so an old saved No cannot reappear.
    if is_boots_product(product_name, category, style_number):
        return True
    if key in overrides:
        return bool(overrides[key])
    if vendor_never_outsource(vendor) or default_never_outsource(product_name, category, style_number):
        return True
    return normalize_bool(saved_value, False)


class ProductMasterV2(ctk.CTk):
    """Product Master editor. The class name remains compatible with older app versions."""

    def __init__(self, prefill_product: str = "", search_product: str = "", setup_product: str = ""):
        super().__init__()
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")
        self.title("Orchid Purchase Manager - Product Master 4.9.104")
        self.geometry("1320x930")
        self.minsize(1280, 780)
        self.configure(fg_color=WINDOW_BG)
        self.after(80, self._maximize_window)
        self._last_open_request = 0.0
        try:
            PRODUCT_MASTER_PID_FILE.parent.mkdir(parents=True, exist_ok=True)
            PRODUCT_MASTER_PID_FILE.write_text(str(os.getpid()), encoding="utf-8")
        except Exception:
            pass

        self.master_data = load_product_master()
        self._loaded_master_stamp = product_master_file_stamp()
        self._loading_style = False
        self._style_render_generation = 0
        self._style_render_after_id = None
        self._catalog_refresh_after_id = None
        self._dashboard_refresh_after_id = None
        self._editor_filter_after_id = None
        self.session_start_complete = 0
        self.session_start_total = 0
        self.session_saved_styles = set()
        self.session_backup_path = None
        self._save_toast_after_id = None
        self.never_outsource_overrides = load_never_outsource_overrides()
        self._never_outsource_touched = False

        self.vendor_options = purchase_vendor_options(self.master_data)
        self.decoration_color_options = load_thread_ink_colors(PRODUCT_MASTER_PATH)
        self.active_color_control_index = None

        self.all_style_keys = []
        self.filtered_style_keys = []
        self.style_position = 0
        self.color_controls = []
        self.current_style_key = None

        self.search_var = ctk.StringVar()
        self.catalog_filter_var = ctk.StringVar(value="All Products")
        self.incomplete_only_var = ctk.BooleanVar(value=False)  # backward-compatible internal flag
        self.dashboard_search_var = ctk.StringVar()
        self.dashboard_vendor_filter_var = ctk.StringVar(value="All Vendors")
        self.vendor_var = ctk.StringVar()
        self.decoration_type_var = ctk.StringVar()
        self.decoration_location_var = ctk.StringVar(value=LEFT_CHEST)
        self.custom_decoration_location_var = ctk.StringVar()
        self.decoration_placement_instructions_var = ctk.StringVar()
        self.product_name_var = ctk.StringVar()
        self.style_number_var = ctk.StringVar()
        self.product_id_var = ctk.StringVar()
        self.product_aliases_var = ctk.StringVar()
        self.category_var = ctk.StringVar()
        self.requires_size_var = ctk.BooleanVar(value=True)
        self.requires_color_var = ctk.BooleanVar(value=True)
        self.requires_decoration_var = ctk.BooleanVar(value=True)
        self.never_outsource_var = ctk.BooleanVar(value=False)
        self.advanced_visible = False
        self.apply_color_var = ctk.StringVar(value="")
        self._suspend_editor_filter_trace = False

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.build_style_list()
        self.session_start_total = len(self.all_style_keys)
        self.session_start_complete = self.completed_style_count()

        self.dashboard_frame = ctk.CTkFrame(self, fg_color=WINDOW_BG, corner_radius=0)
        self.editor_frame = ctk.CTkFrame(self, fg_color=WINDOW_BG, corner_radius=0)
        self.dashboard_frame.grid(row=0, column=0, sticky="nsew")
        self.editor_frame.grid(row=0, column=0, sticky="nsew")
        self.build_dashboard()
        self.build_editor()
        install_native_scroll_support(
            self,
            [self.queue_scroll, self.vendor_stats_frame, self.left_scroll, self.colors_scroll],
        )
        self.product_name_var.trace_add("write", self.on_product_name_changed)
        self.decoration_type_var.trace_add("write", self.on_decoration_type_changed)
        self.vendor_var.trace_add("write", self.on_vendor_changed)
        self.show_dashboard()
        if clean_text(prefill_product):
            self.after(150, lambda value=clean_text(prefill_product): self.add_new_product(value, prompt=False))
        elif clean_text(setup_product):
            self.after(150, lambda value=clean_text(setup_product): self._open_setup_product(value))
        elif clean_text(search_product):
            self.after(150, lambda value=clean_text(search_product): self._open_search_result(value))
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.after(40, self._bring_to_front)
        self.after(300, self._poll_open_requests)

    def _maximize_window(self):
        """Open at the largest normal window size without entering macOS full-screen mode."""
        try:
            self.state("zoomed")
            return
        except Exception:
            pass
        try:
            self.update_idletasks()
            width = max(self.winfo_screenwidth(), self.winfo_width())
            height = max(self.winfo_screenheight() - 70, self.winfo_height())
            self.geometry(f"{width}x{height}+0+0")
        except Exception:
            pass

    def _bring_to_front(self):
        try:
            self.deiconify()
            self.lift()
            self.attributes("-topmost", True)
            self.focus_force()
            self.after(700, lambda: self.attributes("-topmost", False))
        except Exception:
            pass
        if sys.platform == "darwin":
            script = (
                'tell application "System Events" to set frontmost of '
                f'(first process whose unix id is {os.getpid()}) to true'
            )
            subprocess.run(
                ["/usr/bin/osascript", "-e", script], check=False,
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            )

    def _poll_open_requests(self):
        try:
            if PRODUCT_MASTER_REQUEST_FILE.exists():
                modified = PRODUCT_MASTER_REQUEST_FILE.stat().st_mtime
                if modified > self._last_open_request:
                    self._last_open_request = modified
                    payload = json.loads(PRODUCT_MASTER_REQUEST_FILE.read_text(encoding="utf-8"))
                    PRODUCT_MASTER_REQUEST_FILE.unlink(missing_ok=True)
                    self._bring_to_front()
                    action = str(payload.get("action", "show"))
                    value = clean_text(payload.get("value", ""))
                    if action == "new" and value:
                        self.add_new_product(value, prompt=False)
                    elif action == "setup" and value:
                        self._open_setup_product(value)
                    elif action == "search" and value:
                        self._open_search_result(value)
        except Exception:
            pass
        try:
            self.after(300, self._poll_open_requests)
        except Exception:
            pass

    # ---------- data and health ----------
    def build_style_list(self):
        """Build an index once so large Product Masters do not require repeated full-file scans."""
        self.all_style_keys = []
        self.style_rows_cache = {}
        self.style_search_cache = {}
        self.style_display_cache = {}
        self.style_vendor_cache = {}
        self.style_category_cache = {}
        self.style_editor_metadata_cache = {}
        # Product completeness is derived from the current master rows. Clear
        # it only when those rows are reindexed, then reuse it while browsing.
        self.style_complete_cache = {}
        self.style_issues_cache = {}
        self._complete_count_cache = None
        for index, row in self.master_data.iterrows():
            # Defense in depth for a catalog changed by another process while
            # this editor is open.  Decoration charges never belong in the
            # Product Master navigation or its setup counts.
            if is_in_house_service_product(
                row.get("Product Name", ""),
                row.get("Original Line Item", ""),
                row.get("Decoration Type", ""),
                style_number=row.get("Style Number", ""),
                garment_color=row.get("Garment Color", ""),
            ):
                continue
            key = style_group_key(row)
            if key not in self.style_rows_cache:
                self.style_rows_cache[key] = []
                self.all_style_keys.append(key)
            self.style_rows_cache[key].append(index)
        # Search used to normalize every field of every row on each keystroke.
        # Build that small search index only when the data actually changes.
        for key in self.all_style_keys:
            indices = self.style_rows_cache[key]
            rows = self.master_data.loc[indices]
            style = first_nonblank(rows["Style Number"])
            product_id = first_nonblank(rows["Product ID"])
            name = best_product_name(rows["Product Name"])
            vendor = first_nonblank(rows["Vendor"])
            category = first_nonblank(rows["Product Category"])
            self.style_search_cache[key] = {
                "style": normalize_search_text(style),
                "product_id": normalize_search_text(product_id),
                "name": normalize_search_text(name),
                "vendor": normalize_search_text(vendor),
                "combined": normalize_search_text(
                    " ".join(clean_text(value) for value in rows[COLUMNS].to_numpy().flatten())
                ),
            }
            self.style_display_cache[key] = (
                f"{style} — {name}" if style and name else (style or name or "Unnamed product")
            )
            self.style_vendor_cache[key] = vendor or "Unassigned"
            self.style_category_cache[key] = category or "Category not assigned"
            self.style_editor_metadata_cache[key] = {
                "style": style or "No product number",
                "name": name or "Unnamed product",
            }

    def get_style_rows(self, style_key):
        if self.master_data.empty:
            return self.master_data.copy()
        indices = self.style_rows_cache.get(style_key, [])
        if not indices:
            return self.master_data.iloc[0:0].copy()
        return self.master_data.loc[indices].copy()

    def style_issues(self, rows):
        issues = []
        if rows.empty:
            return ["No product record"]
        decoration_type = first_nonblank(rows["Decoration Type"])
        style_number = first_nonblank(rows["Style Number"])
        if is_in_house_service_product(
            first_nonblank(rows["Product Name"]),
            decoration_type=decoration_type,
            style_number=style_number,
        ):
            return []
        new_color_guidance = new_color_setup_guidance(rows)
        if new_color_guidance["new_count"]:
            names = ", ".join((new_color_guidance["thread_ink_color_names"] or new_color_guidance["new_color_names"])[:2])
            remaining = new_color_guidance["new_count"] - min(
                2, len(new_color_guidance["thread_ink_color_names"] or new_color_guidance["new_color_names"])
            )
            if remaining:
                names += f" +{remaining} more"
            if new_color_guidance["thread_ink_color_names"]:
                issues.append(f"New color needs thread/ink: {names}")
            else:
                issues.append(f"New color setup must be reviewed: {names}")
        elif "Setup Required" in rows.columns and rows["Setup Required"].map(is_setup_required).any():
            issues.append("New product setup must be reviewed and saved")
        if not first_nonblank(rows["Product Name"]):
            issues.append("Missing product name")
        if not first_nonblank(rows["Style Number"]):
            issues.append("Missing product number")
        if not first_nonblank(rows["Product Category"]):
            issues.append("Missing product category")
        rules = row_rules(rows.iloc[0])
        requires_color = normalize_bool(rules["Requires Color"], True)
        requires_decoration = normalize_bool(rules["Requires Decoration"], True)
        if requires_color and rows["Garment Color"].astype(str).str.strip().eq("").any():
            issues.append("Missing purchasing color")
        service_only = is_internal_service_style(style_number) or is_in_house_service_product(
            first_nonblank(rows["Product Name"]),
            decoration_type=decoration_type,
            style_number=style_number,
        )
        # Vendor, decoration type, and placement are style-level settings. A blank
        # duplicate color row inherits the shared value instead of making every
        # complete color look unfinished.
        shared_vendor = first_nonblank(rows["Vendor"])
        shared_decoration = first_nonblank(rows["Decoration Type"])
        shared_location = first_nonblank(rows["Decoration Location"])
        if not service_only and not shared_vendor:
            issues.append("Missing purchase vendor")
        if not service_only and not shared_decoration:
            issues.append("Missing decoration type")
        decorated = rows[rows["Decoration Type"].map(lambda value: not is_blank_decoration(value) and not is_in_house_decoration(value))]
        if requires_decoration and not decorated.empty:
            if not shared_location or shared_location == OTHER_CUSTOM:
                issues.append("Missing decoration location")
            # Thread/ink color remains color-specific, but only actual purchasing
            # colors are validated. Empty placeholder rows do not block the style.
            actual_colors = decorated[decorated["Garment Color"].astype(str).str.strip().ne("")]
            if not actual_colors.empty and actual_colors["Decoration Color"].astype(str).str.strip().eq("").any():
                shared_thread = first_nonblank(actual_colors["Decoration Color"])
                if not shared_thread:
                    issues.append("Missing thread/ink color")
        return issues

    def style_is_complete(self, rows):
        return not self.style_issues(rows)

    def completed_style_count(self):
        if self._complete_count_cache is None:
            self._complete_count_cache = sum(
                1 for key in self.all_style_keys if self.style_key_is_complete(key)
            )
        return self._complete_count_cache

    def style_key_is_complete(self, key):
        """Cache catalog health until Product Master data actually changes."""
        if key not in self.style_complete_cache:
            self.style_complete_cache[key] = not self.style_key_issues(key)
        return self.style_complete_cache[key]

    def style_key_issues(self, key):
        """Reuse a style's health result across search, dashboard, and editor cards."""
        if key not in self.style_issues_cache:
            self.style_issues_cache[key] = self.style_issues(self.get_style_rows(key))
        return self.style_issues_cache[key]

    def incomplete_style_keys(self):
        return [key for key in self.all_style_keys if not self.style_key_is_complete(key)]

    def vendor_style_counts(self):
        counts = {}
        for key in self.all_style_keys:
            vendor = self.style_vendor_cache.get(key, "Unassigned")
            counts[vendor] = counts.get(vendor, 0) + 1
        return sorted(counts.items(), key=lambda item: (-item[1], item[0].casefold()))

    def style_display(self, key):
        return self.style_display_cache.get(key, "Unnamed product")

    # ---------- dashboard ----------
    def _open_search_result(self, value: str):
        value = clean_text(value)
        if not value:
            return
        self.dashboard_search_var.set(value)
        self.refresh_dashboard()
        matches = self.dashboard_matching_style_keys()
        if matches:
            self.open_style(matches[0])

    def _open_setup_product(self, value: str):
        """Open the requested candidate inside the Needs Setup queue."""
        value = clean_text(value)
        if not value:
            return
        self.reload_live_file(silent=True)
        self.catalog_filter_var.set("Needs Setup")
        self.incomplete_only_var.set(True)
        self.search_var.set("")
        self.dashboard_search_var.set(value)
        matches = self.dashboard_matching_style_keys()
        if not matches:
            self.add_new_product(value, prompt=False)
            return
        target = matches[0]
        self.show_editor(target_key=target)
        self.set_catalog_filter("Needs Setup")
        if target in self.filtered_style_keys:
            self.style_position = self.filtered_style_keys.index(target)
            self.show_current_style()

    def build_dashboard(self):
        frame = self.dashboard_frame
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(3, weight=1)

        header = ctk.CTkFrame(frame, fg_color=CARD_BG, corner_radius=0, height=104)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_columnconfigure(1, weight=1)
        flower_path = Path(__file__).resolve().parent.parent / "assets" / "orchid_flower_purple.png"
        if flower_path.exists():
            flower_source = PILImage.open(flower_path).convert("RGBA")
            self.product_master_flower_image = ctk.CTkImage(light_image=flower_source, dark_image=flower_source, size=(62, 62))
            ctk.CTkLabel(header, text="", image=self.product_master_flower_image, width=70).grid(row=0, column=0, rowspan=2, padx=(28, 8), pady=18)
        else:
            # The branded flower asset ships with every build. Avoid substituting
            # a generic glyph if an installation is incomplete.
            ctk.CTkLabel(header, text="", width=70).grid(
                row=0, column=0, rowspan=2, padx=(28, 8), pady=18
            )
        ctk.CTkLabel(header, text="Product Master", text_color=TEXT_DARK, font=ctk.CTkFont(size=31, weight="bold"), anchor="w").grid(row=0, column=1, sticky="sw", pady=(18, 0))
        ctk.CTkLabel(header, text="Define where products are purchased and what information is required to place an order.", text_color=TEXT_MUTED, font=ctk.CTkFont(size=14), anchor="w").grid(row=1, column=1, sticky="nw", pady=(2, 16))
        ctk.CTkButton(header, text="+ Add New Product", command=self.add_new_product, width=155, height=42, fg_color="transparent", hover_color=PURPLE_LIGHT, border_width=1, border_color=PURPLE, text_color=PURPLE, font=ctk.CTkFont(size=13, weight="bold")).grid(row=0, column=2, rowspan=2, padx=(8, 6), pady=28)
        ctk.CTkButton(header, text="Refresh", command=self.reload_live_file, width=145, height=42, fg_color="transparent", hover_color=PURPLE_LIGHT, border_width=1, border_color=PURPLE, text_color=PURPLE, font=ctk.CTkFont(size=13, weight="bold")).grid(row=0, column=3, rowspan=2, padx=6, pady=28)
        ctk.CTkButton(header, text="Browse Catalog  ▶", command=self.browse_all, width=205, height=44, fg_color=PURPLE, hover_color=PURPLE_DARK, font=ctk.CTkFont(size=14, weight="bold")).grid(row=0, column=4, rowspan=2, padx=(6, 22), pady=28)

        self.dashboard_stats = ctk.CTkFrame(frame, fg_color="transparent")
        self.dashboard_stats.grid(row=1, column=0, sticky="ew", padx=20, pady=(16, 8))
        for column in range(4):
            self.dashboard_stats.grid_columnconfigure(column, weight=1)
        self.total_card = self.create_stat_card(self.dashboard_stats, 0, "Total Styles", "0")
        self.complete_card = self.create_stat_card(self.dashboard_stats, 1, "Fully Configured", "0")
        # Needs Setup is intentionally an action, not just a status.  When an
        # import identifies a few unfinished styles among a large catalog, the
        # operator should be able to go directly to that short queue.
        self.incomplete_card = self.create_stat_card(
            self.dashboard_stats,
            2,
            "Needs Setup",
            "0",
            command=self.open_needs_setup_queue,
        )
        self.percent_card = self.create_stat_card(self.dashboard_stats, 3, "Catalog Status", "Ready")

        health = ctk.CTkFrame(frame, fg_color=CARD_BG, border_width=1, border_color=PURPLE_BORDER, corner_radius=12)
        health.grid(row=2, column=0, sticky="ew", padx=20, pady=8)
        health.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(health, text="Product Master Progress", text_color=TEXT_DARK, font=ctk.CTkFont(size=15, weight="bold")).grid(row=0, column=0, padx=(20, 14), pady=18)
        self.dashboard_progress = ctk.CTkProgressBar(health, progress_color=PURPLE, fg_color="#EDE8F3")
        self.dashboard_progress.grid(row=0, column=1, sticky="ew", padx=10, pady=18)
        self.dashboard_health_text = ctk.CTkLabel(health, text="", text_color=TEXT_MUTED, font=ctk.CTkFont(size=13))
        self.dashboard_health_text.grid(row=0, column=2, padx=(14, 20), pady=18)

        body = ctk.CTkFrame(frame, fg_color="transparent")
        body.grid(row=3, column=0, sticky="nsew", padx=20, pady=(8, 16))
        body.grid_columnconfigure(0, weight=2)
        body.grid_columnconfigure(1, weight=1)
        body.grid_rowconfigure(0, weight=1)

        queue_card = ctk.CTkFrame(body, fg_color=CARD_BG, border_width=1, border_color=PURPLE_BORDER, corner_radius=12)
        queue_card.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        queue_card.grid_columnconfigure(0, weight=1)
        queue_card.grid_rowconfigure(3, weight=1)
        ctk.CTkLabel(queue_card, text="Product Catalog", text_color=PURPLE, font=ctk.CTkFont(size=18, weight="bold"), anchor="w").grid(row=0, column=0, sticky="ew", padx=20, pady=(18, 4))
        ctk.CTkLabel(queue_card, text="Search or browse saved product master details. Only gaps affecting the current import appear in Purchase Review.", text_color=TEXT_MUTED, font=ctk.CTkFont(size=12), anchor="w").grid(row=1, column=0, sticky="ew", padx=20, pady=(0, 8))
        search_row = ctk.CTkFrame(queue_card, fg_color="transparent")
        search_row.grid(row=2, column=0, sticky="ew", padx=20, pady=(0, 10))
        search_row.grid_columnconfigure(0, weight=1)
        self.dashboard_search_entry = ctk.CTkEntry(search_row, textvariable=self.dashboard_search_var, placeholder_text="Search product #, ID, alias, description, vendor, or color...", height=42, corner_radius=9, border_color=PURPLE_BORDER, fg_color="#FFFFFF", text_color=TEXT_DARK)
        self.dashboard_search_entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        self.dashboard_search_entry.bind("<Return>", lambda _event: self.open_first_dashboard_match())
        dashboard_vendor_values = ["All Vendors"] + self.vendor_options + ["Unassigned"]
        self.dashboard_vendor_filter = ctk.CTkComboBox(search_row, variable=self.dashboard_vendor_filter_var, values=list(dict.fromkeys(dashboard_vendor_values)), width=170, height=38, command=lambda _value: self.refresh_dashboard(), border_color=PURPLE_BORDER, button_color="#EAE3F5", button_hover_color="#DED2EF", fg_color="#FFFFFF", text_color=TEXT_DARK)
        self.dashboard_vendor_filter.grid(row=0, column=1, padx=(8, 0))
        self.dashboard_setup_button = ctk.CTkButton(
            search_row,
            text="Set Up Next  ➜",
            command=self.open_needs_setup_queue,
            width=142,
            height=38,
            fg_color=PURPLE,
            hover_color=PURPLE_DARK,
            font=ctk.CTkFont(size=11, weight="bold"),
        )
        self.dashboard_setup_button.grid(row=0, column=2, padx=(8, 0))
        self.dashboard_search_results = ctk.CTkLabel(
            search_row, text="", text_color=TEXT_MUTED, font=ctk.CTkFont(size=11), anchor="w"
        )
        self.dashboard_search_results.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(5, 0))
        self.dashboard_search_var.trace_add("write", lambda *_: self._schedule_dashboard_refresh())
        self.queue_scroll = ctk.CTkScrollableFrame(queue_card, fg_color="#FFFFFF", corner_radius=8, border_width=1, border_color=PURPLE_BORDER)
        self.queue_scroll.grid(row=3, column=0, sticky="nsew", padx=20, pady=(0, 18))
        self.queue_scroll.grid_columnconfigure(0, weight=1)

        vendor_card = ctk.CTkFrame(body, fg_color=CARD_BG, border_width=1, border_color=PURPLE_BORDER, corner_radius=12)
        vendor_card.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        vendor_card.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(vendor_card, text="Styles by Vendor", text_color=PURPLE, font=ctk.CTkFont(size=18, weight="bold"), anchor="w").grid(row=0, column=0, sticky="ew", padx=20, pady=(18, 4))
        self.vendor_stats_frame = ctk.CTkScrollableFrame(vendor_card, fg_color="#FFFFFF", corner_radius=8, border_width=1, border_color=PURPLE_BORDER)
        self.vendor_stats_frame.grid(row=1, column=0, sticky="nsew", padx=20, pady=(10, 12))
        self.vendor_stats_frame.grid_columnconfigure(0, weight=1)
        vendor_card.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(vendor_card, text="Click a vendor to filter the Product Catalog.", text_color=TEXT_MUTED, font=ctk.CTkFont(size=11), justify="left", wraplength=330).grid(row=2, column=0, sticky="ew", padx=20, pady=(0, 18))

    def create_stat_card(self, parent, column, title, value, command=None):
        card = ctk.CTkFrame(parent, fg_color=CARD_BG, border_width=1, border_color=PURPLE_BORDER, corner_radius=12)
        card.grid(row=0, column=column, sticky="ew", padx=(0 if column == 0 else 7, 0 if column == 3 else 7))
        title_label = ctk.CTkLabel(card, text=title, text_color=TEXT_MUTED, font=ctk.CTkFont(size=12, weight="bold"))
        title_label.pack(anchor="w", padx=18, pady=(15, 3))
        label = ctk.CTkLabel(card, text=value, text_color=TEXT_DARK, font=ctk.CTkFont(size=27, weight="bold"))
        label.pack(anchor="w", padx=18, pady=(0, 15))
        if command is not None:
            # Bind each visible part of the card so clicking the count or its
            # heading produces the same one-click Needs Setup view.
            for widget in (card, title_label, label):
                widget.bind("<Button-1>", lambda _event: command())
        return label

    def refresh_dashboard(self):
        # The style index is rebuilt whenever master_data changes. Rebuilding
        # it here made every search and catalog repaint discard the cached
        # completeness result and re-evaluate the entire Product Master.
        total = len(self.all_style_keys)
        complete = self.completed_style_count()
        incomplete = total - complete
        percent = complete / total if total else 0
        self.total_card.configure(text=str(total))
        self.complete_card.configure(text=str(complete), text_color=SUCCESS)
        self.incomplete_card.configure(text=str(incomplete), text_color=WARNING if incomplete else SUCCESS)
        self.percent_card.configure(
            text="Ready" if not incomplete else "Needs Setup",
            text_color=SUCCESS if not incomplete else WARNING,
        )
        self.dashboard_progress.set(percent)
        if incomplete:
            self.dashboard_health_text.configure(text=f"{complete} complete • {incomplete} need attention • {total} total styles • {round(percent * 100)}%")
        else:
            self.dashboard_health_text.configure(text=f"{total} complete • Product Master ready • 100%")
        if hasattr(self, "dashboard_vendor_filter"):
            values = ["All Vendors"] + self.vendor_options + ["Unassigned"]
            self.dashboard_vendor_filter.configure(values=list(dict.fromkeys(values)))
        if hasattr(self, "dashboard_setup_button"):
            if incomplete:
                self.dashboard_setup_button.configure(
                    text=f"Set Up Next ({incomplete})  ➜",
                    state="normal",
                )
            else:
                self.dashboard_setup_button.configure(text="Product Master Ready", state="disabled")
        if hasattr(self, "filter_buttons"):
            self.filter_buttons["All Products"].configure(text=f"All Products ({total})")
            self.filter_buttons["Needs Setup"].configure(text=f"Needs Setup ({incomplete})")
            self.filter_buttons["Complete"].configure(text=f"Complete ({complete})")
            self._refresh_filter_button_styles()

        for widget in self.queue_scroll.winfo_children():
            widget.destroy()
        display_keys = self.dashboard_matching_style_keys()
        searching = bool(self.dashboard_search_var.get().strip()) or self.dashboard_vendor_filter_var.get() != "All Vendors"
        if hasattr(self, "dashboard_search_results"):
            label = f"{len(display_keys)} result{'s' if len(display_keys) != 1 else ''}"
            if searching:
                label += " — press Return to open the best match"
            self.dashboard_search_results.configure(text=label)
        if not display_keys:
            search_value = self.dashboard_search_var.get().strip()
            if searching and search_value:
                ctk.CTkLabel(self.queue_scroll, text=f"{search_value} is not in Product Master", text_color=WARNING, font=ctk.CTkFont(size=18, weight="bold")).grid(row=0, column=0, padx=20, pady=(34, 8))
                ctk.CTkLabel(self.queue_scroll, text="Add it as a new permanent product master detail, then complete the vendor, color, size requirement, and decoration information.", text_color=TEXT_MUTED, font=ctk.CTkFont(size=12), wraplength=590).grid(row=1, column=0, padx=20, pady=(0, 14))
                ctk.CTkButton(self.queue_scroll, text=f"+ Add {search_value} as New Product", command=lambda value=search_value: self.add_new_product(value, prompt=False), width=235, height=38, fg_color=PURPLE, hover_color=PURPLE_DARK, font=ctk.CTkFont(size=13, weight="bold")).grid(row=2, column=0, padx=20, pady=(0, 34))
            else:
                message = "No products match this search" if searching else "No products are saved in the catalog yet"
                color = TEXT_MUTED
                ctk.CTkLabel(self.queue_scroll, text=message, text_color=color, font=ctk.CTkFont(size=18, weight="bold")).grid(row=0, column=0, padx=20, pady=40)
        else:
            visible_keys = display_keys[:MAX_DASHBOARD_ROWS]
            row_offset = 0
            if len(display_keys) > MAX_DASHBOARD_ROWS:
                ctk.CTkLabel(
                    self.queue_scroll,
                    text=f"Showing the first {MAX_DASHBOARD_ROWS} of {len(display_keys)} products. Use search or the vendor filter to find a specific style without loading every card at once.",
                    text_color=WARNING,
                    font=ctk.CTkFont(size=11, weight="bold"),
                    justify="left",
                    wraplength=590,
                ).grid(row=0, column=0, sticky="ew", padx=14, pady=(12, 8))
                row_offset = 1
            for row_index, key in enumerate(visible_keys, start=row_offset):
                rows = self.get_style_rows(key)
                item = ctk.CTkFrame(self.queue_scroll, fg_color=ROW_ALT if row_index % 2 else "#FFFFFF", corner_radius=8)
                item.grid(row=row_index, column=0, sticky="ew", padx=4, pady=3)
                item.grid_columnconfigure(0, weight=1)
                ctk.CTkLabel(item, text=self.style_display(key), text_color=TEXT_DARK, font=ctk.CTkFont(size=13, weight="bold"), anchor="w", wraplength=500).grid(row=0, column=0, sticky="ew", padx=14, pady=(10, 2))
                issues = self.style_key_issues(key)
                has_vendor = bool(first_nonblank(rows["Vendor"]))
                has_decoration = bool(first_nonblank(rows["Decoration Type"]))
                if not issues:
                    badge_text, badge_text_color, badge_fill = "COMPLETE", SUCCESS, "#EAF8EF"
                    detail_text = "All required purchasing details are configured"
                    button_text = "Edit"
                elif has_vendor or has_decoration:
                    badge_text, badge_text_color, badge_fill = "PARTIAL", WARNING, "#FFF1DC"
                    detail_text = " • ".join(issues[:2]) + (f" • +{len(issues) - 2} more" if len(issues) > 2 else "")
                    button_text = "Continue"
                else:
                    badge_text, badge_text_color, badge_fill = "NEEDS SETUP", DANGER, "#FDECEC"
                    detail_text = " • ".join(issues[:2]) + (f" • +{len(issues) - 2} more" if len(issues) > 2 else "")
                    button_text = "Set Up"
                ctk.CTkLabel(item, text=detail_text, text_color=TEXT_MUTED, font=ctk.CTkFont(size=11), anchor="w", wraplength=500).grid(row=1, column=0, sticky="ew", padx=14, pady=(0, 10))
                ctk.CTkLabel(item, text=badge_text, height=25, corner_radius=12, fg_color=badge_fill, text_color=badge_text_color, font=ctk.CTkFont(size=10, weight="bold")).grid(row=0, column=1, rowspan=2, padx=(8, 4), pady=10)
                ctk.CTkButton(item, text=button_text, command=lambda target=key: self.open_style(target), width=82, height=32, fg_color="transparent", hover_color=PURPLE_LIGHT, border_width=1, border_color=PURPLE, text_color=PURPLE).grid(row=0, column=2, rowspan=2, padx=(4, 12), pady=10)

        for widget in self.vendor_stats_frame.winfo_children():
            widget.destroy()
        vendor_rows = [("All Vendors", len(self.all_style_keys))] + self.vendor_style_counts()
        for index, (vendor, count) in enumerate(vendor_rows):
            selected = self.dashboard_vendor_filter_var.get().casefold() == vendor.casefold()
            row_fill = PURPLE_LIGHT if selected else (ROW_ALT if index % 2 else "#FFFFFF")
            row = ctk.CTkFrame(self.vendor_stats_frame, fg_color=row_fill, corner_radius=6)
            row.grid(row=index, column=0, sticky="ew", pady=1)
            row.grid_columnconfigure(0, weight=1)
            ctk.CTkButton(
                row,
                text=vendor,
                command=lambda value=vendor: self.filter_dashboard_by_vendor(value),
                fg_color="transparent",
                hover_color=PURPLE_LIGHT,
                text_color=PURPLE_DARK if selected else TEXT_DARK,
                anchor="w",
                height=32,
                font=ctk.CTkFont(size=12, weight="bold" if selected else "normal"),
            ).grid(row=0, column=0, sticky="ew", padx=(4, 0), pady=2)
            ctk.CTkButton(
                row,
                text=str(count),
                command=lambda value=vendor: self.filter_dashboard_by_vendor(value),
                width=48,
                height=28,
                fg_color="transparent",
                hover_color=PURPLE_LIGHT,
                text_color=PURPLE,
                font=ctk.CTkFont(size=12, weight="bold"),
            ).grid(row=0, column=1, padx=(0, 4), pady=2)

    def filter_dashboard_by_vendor(self, vendor: str):
        """Filter the Product Catalog when a vendor summary row is clicked."""
        self.dashboard_vendor_filter_var.set(vendor or "All Vendors")
        self.refresh_dashboard()

    def open_needs_setup_queue(self):
        """Open only unfinished styles and select the next one to complete."""
        incomplete = self.incomplete_style_keys()
        if not incomplete:
            messagebox.showinfo("Product Master Complete", "Every saved style is fully configured.", parent=self)
            return
        self._suspend_editor_filter_trace = True
        try:
            self.dashboard_search_var.set("")
            self.dashboard_vendor_filter_var.set("All Vendors")
            self.search_var.set("")
            self.catalog_filter_var.set("Needs Setup")
            self.incomplete_only_var.set(True)
        finally:
            self._suspend_editor_filter_trace = False
        self.show_editor(target_key=incomplete[0])

    def dashboard_matching_style_keys(self):
        query = self.dashboard_search_var.get().strip()
        vendor_filter = self.dashboard_vendor_filter_var.get().strip() or "All Vendors"
        ranked = []
        for position, key in enumerate(self.all_style_keys):
            vendor = self.style_vendor_cache.get(key, "Unassigned")
            if vendor_filter != "All Vendors" and vendor.casefold() != vendor_filter.casefold():
                continue
            search = self.style_search_cache.get(key, {})
            score = search_match_score_values(
                search.get("style", ""), search.get("product_id", ""),
                search.get("name", ""), search.get("vendor", ""),
                search.get("combined", ""), query,
            )
            if score is None:
                continue
            ranked.append((score, position, key))
        ranked.sort(key=lambda item: (item[0], item[1]))
        return [key for _, _, key in ranked]

    def _schedule_dashboard_refresh(self):
        """Wait for a short typing pause before rebuilding visual catalog cards."""
        if self._dashboard_refresh_after_id:
            try:
                self.after_cancel(self._dashboard_refresh_after_id)
            except Exception:
                pass

        def refresh():
            self._dashboard_refresh_after_id = None
            self.refresh_dashboard()

        self._dashboard_refresh_after_id = self.after(140, refresh)

    def open_first_dashboard_match(self):
        matches = self.dashboard_matching_style_keys()
        if matches:
            self.open_style(matches[0])

    def show_dashboard(self):
        """Return reliably from Style Editor to the Product Catalog dashboard."""
        self.editor_frame.grid_remove()
        self.dashboard_frame.grid(sticky="nsew")
        self.dashboard_frame.tkraise()
        # Reload only when another window actually changed the catalog. The
        # previous unconditional reload reread and rebuilt the full catalog every
        # time Product Catalog was opened.
        if product_master_file_stamp() != getattr(self, "_loaded_master_stamp", (0, 0)):
            self.reload_live_file(silent=True)
        else:
            self.refresh_dashboard()
        self.update_idletasks()

    def reload_live_file(self, silent: bool = False):
        """Reload the active Product Master from disk so this window cannot show stale data."""
        self.master_data = load_product_master()
        self._loaded_master_stamp = product_master_file_stamp()
        self.vendor_options = purchase_vendor_options(self.master_data)
        self.decoration_color_options = load_thread_ink_colors(PRODUCT_MASTER_PATH)
        if hasattr(self, "dashboard_vendor_filter"):
            self.dashboard_vendor_filter.configure(values=["All Vendors"] + self.vendor_options + ["Unassigned"])
        if hasattr(self, "vendor_combo"):
            self.vendor_combo.configure(values=self.vendor_options)
        self._refresh_thread_ink_dropdowns()
        self.build_style_list()
        self.refresh_dashboard()
        if not silent:
            messagebox.showinfo("Product Master Reloaded", f"Reloaded the live Product Master catalog from:\n{PRODUCT_MASTER_PATH}")

    def add_vendor_option(self):
        """Add and select a vendor without leaving the Style Editor."""
        value = simpledialog.askstring(
            "Add Purchase Vendor",
            "Enter the vendor name exactly as it should appear on purchase orders.",
            parent=self,
        )
        value = canonical_vendor_name(value or "")
        if not value:
            return
        # Adding a vendor is itself a permanent Product Master action.  Do not
        # make the operator save an unrelated style merely to keep the option.
        value = register_purchase_vendor(value)
        existing = next((vendor for vendor in self.vendor_options if vendor.casefold() == value.casefold()), "")
        selected = existing or value
        if not existing:
            append_vendor_option(self.vendor_options, value)
            if hasattr(self, "vendor_combo"):
                self.vendor_combo.configure(values=self.vendor_options)
            if hasattr(self, "dashboard_vendor_filter"):
                self.dashboard_vendor_filter.configure(
                    values=list(dict.fromkeys(["All Vendors"] + self.vendor_options + ["Unassigned"]))
                )
        self.vendor_options = alphabetical_options(self.vendor_options)
        if hasattr(self, "vendor_combo"):
            self.vendor_combo.configure(values=self.vendor_options)
        if hasattr(self, "dashboard_vendor_filter"):
            self.dashboard_vendor_filter.configure(values=["All Vendors"] + self.vendor_options + ["Unassigned"])
        self.vendor_var.set(selected)
        self.status_label.configure(text=f"Vendor {selected} added and selected. Save this product to keep it in future dropdowns.")

    def _set_active_color_row(self, index):
        self.active_color_control_index = index

    def _refresh_thread_ink_dropdowns(self):
        options = getattr(self, "decoration_color_options", load_thread_ink_colors(PRODUCT_MASTER_PATH))
        self.decoration_color_options = alphabetical_options(options, include_blank=True)
        if hasattr(self, "apply_color_combo"):
            self.apply_color_combo.configure(values=self.decoration_color_options)
        for control in getattr(self, "color_controls", []):
            combo = control.get("decoration_combo")
            if combo is not None:
                combo.configure(values=self.decoration_color_options)

    def add_thread_ink_color_option(self):
        """Create one reusable thread/ink color and select it on the active color row."""
        value = simpledialog.askstring(
            "Add Thread / Ink Color",
            "Enter the thread or ink color exactly as it should appear on purchasing and decoration reports.\n\nExamples: Metallic Gold, Dark Silver, Safety Green",
            parent=self,
        )
        selected, added, options = register_thread_ink_color(value or "", PRODUCT_MASTER_PATH)
        if not selected:
            return
        self.decoration_color_options = options
        self._refresh_thread_ink_dropdowns()

        target = next((
            control for control in self.color_controls
            if control.get("index") == self.active_color_control_index
        ), None)
        if target is None:
            target = next((control for control in self.color_controls if not control["decoration_var"].get().strip()), None)
        if target is None and self.color_controls:
            target = self.color_controls[0]
        if target is not None:
            target["decoration_var"].set(selected)
            self.active_color_control_index = target.get("index")
        self.apply_color_var.set(selected)
        message = (
            f"Thread/ink color {selected} added and selected."
            if added else f"Thread/ink color {selected} already exists and is now selected."
        )
        self.status_label.configure(text=message)

    def add_new_product(self, seed: str | None = None, prompt: bool = True):
        """Create a new Product Master record and open it for completion."""
        if prompt:
            seed = simpledialog.askstring(
                "Add New Product",
                "Enter the product/style number or a product description.\n\nExample: 1104",
                parent=self,
            )
        seed = normalize_space(seed or "")
        if not seed:
            return

        style_candidate = normalize_style(seed) if re.fullmatch(r"[A-Za-z0-9._/-]+", seed) else ""
        if style_candidate:
            existing_key = f"style:{style_candidate}"
            if existing_key in self.all_style_keys:
                self.open_style(existing_key)
                messagebox.showinfo("Product Already Exists", f"{style_candidate} is already in Product Master.")
                return
            product_name = ""
        else:
            product_name = canonical_product_name(seed)
            existing_key = f"product:{product_name.casefold()}"
            if existing_key in self.all_style_keys:
                self.open_style(existing_key)
                messagebox.showinfo("Product Already Exists", f"{product_name} is already in Product Master.")
                return

        category = infer_category(product_name, style_candidate)
        defaults = category_defaults(category, product_name, style_candidate)
        row = {
            "Product Name": product_name,
            "Style Number": style_candidate,
            "Garment Color": "",
            "Vendor": "",
            "Decoration Type": "",
            "Decoration Location": "",
            "Decoration Placement Instructions": "",
            "Decoration Color": "",
            "Product ID": default_product_id(style_candidate, product_name),
            "Product Aliases": "",
            "Vendor Color Code": "",
            "Color Aliases": "",
            "Product Category": category,
            "Requires Size": bool_text(defaults["requires_size"]),
            "Requires Color": bool_text(defaults["requires_color"]),
            "Requires Decoration": bool_text(defaults["requires_decoration"]),
            NEVER_OUTSOURCE_COLUMN: bool_text(default_never_outsource(product_name, category, style_candidate)),
            "Setup Required": "Yes",
        }
        self.master_data = pd.concat([self.master_data, pd.DataFrame([row], columns=COLUMNS)], ignore_index=True)
        self.build_style_list()
        new_key = style_group_key(pd.Series(row))
        self.dashboard_search_var.set("")
        self.search_var.set("")
        self.incomplete_only_var.set(False)
        self.show_editor(target_key=new_key)
        if style_candidate and not product_name:
            self.category_var.set("")
        self.status_label.configure(text="New product created. Complete the product name, category, vendor, color, and decoration rules, then click Save.")


    def delete_current_product(self):
        """Permanently remove the selected Product Master style/color group."""
        if not self.current_style_key:
            messagebox.showwarning("No Product Selected", "Select a Product Master record before deleting it.", parent=self)
            return
        rows = self.get_style_rows(self.current_style_key)
        if rows.empty:
            messagebox.showwarning("Product Record Unavailable", "This Product Master record is no longer available.", parent=self)
            return
        style = first_nonblank(rows["Style Number"]) or "No product number"
        description = first_nonblank(rows["Product Name"]) or "No description"
        color_count = max(1, len(rows.index))
        confirmed = messagebox.askyesno(
            "Delete Product Permanently?",
            f"Delete this Product Master record?\n\nProduct: {style}\nDescription: {description}\nPurchasing color rows: {color_count}\n\nThis removes the product from future matching and setup lists. Existing exported reports are not changed. If it is used by the current event, regenerate Purchase Review after deletion.\n\nThis action cannot be undone.",
            icon="warning",
            parent=self,
        )
        if not confirmed:
            return
        self.ensure_session_backup()
        old_key = self.current_style_key
        self.master_data = self.master_data.drop(index=rows.index).reset_index(drop=True)
        saved = save_product_master(self.master_data)
        if saved is None:
            messagebox.showerror("Unable to Delete Product", "Orchid could not save Product Master after removing this record.", parent=self)
            return
        self.master_data = saved
        self.never_outsource_overrides.pop(old_key, None)
        save_never_outsource_overrides(self.never_outsource_overrides)
        self.build_style_list()
        self.current_style_key = None
        self.refresh_progress()
        self._mark_catalog_updated(style)
        self.status_label.configure(text=f"Deleted {style} from Product Master.")
        if self.all_style_keys:
            self.apply_filters(reset_position=True)
            if self.filtered_style_keys:
                self.style_position = 0
                self.show_current_style()
            else:
                self.show_dashboard()
        else:
            self.show_dashboard()
        messagebox.showinfo("Product Deleted", f"{style} — {description} was removed from Product Master.", parent=self)

    def start_incomplete_queue(self):
        incomplete = self.incomplete_style_keys()
        if not incomplete:
            messagebox.showinfo("Product Master", "Every saved style currently has all optional purchasing fields completed.")
            return
        self.search_var.set("")
        self.incomplete_only_var.set(True)
        self.show_editor(target_key=incomplete[0])

    def open_style(self, key):
        """Open the exact catalog record clicked, regardless of the prior queue filter."""
        if key not in self.all_style_keys:
            self.reload_live_file(silent=True)
        if key not in self.all_style_keys:
            messagebox.showwarning("Product Not Found", "That Product Master record is no longer available.", parent=self)
            return

        # A completed item cannot appear inside Needs Setup.  Reset both filter
        # variables before building the editor list so target_key cannot fall
        # back to the first incomplete record (historically American Flag).
        self._suspend_editor_filter_trace = True
        try:
            self.search_var.set("")
            self.catalog_filter_var.set("All Products")
            self.incomplete_only_var.set(False)
        finally:
            self._suspend_editor_filter_trace = False

        # All Products with no search is already represented by all_style_keys.
        # Use that index directly instead of rescanning every style twice.
        self.filtered_style_keys = list(self.all_style_keys)
        self.style_position = self.filtered_style_keys.index(key)
        self.dashboard_frame.grid_remove()
        self.editor_frame.grid(sticky="nsew")
        self.editor_frame.tkraise()
        self._refresh_filter_button_styles()
        # Give macOS a paint cycle before the form and catalog controls are
        # rebuilt.  Without this pause, selecting a product could show a white
        # window until the user moved the mouse or clicked again.
        self.show_current_style(defer=True)
        self.after_idle(self.refresh_progress)

    # ---------- editor ----------
    def build_editor(self):
        frame = self.editor_frame
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(4, weight=1)
        self.build_editor_header(frame)
        self.build_filter_bar(frame)
        self.build_style_banner(frame)
        self.build_content_area(frame)
        self.build_action_bar(frame)
        self.build_footer(frame)
        self.build_style_loading_overlay(frame)
        self.search_var.trace_add("write", self._on_editor_search_changed)

    def build_style_loading_overlay(self, parent):
        """Create the lightweight status card shown before a product redraw."""
        self.style_loading_card = ctk.CTkFrame(
            parent,
            fg_color="#FFFFFF",
            border_width=1,
            border_color=PURPLE_BORDER,
            corner_radius=14,
        )
        self.style_loading_label = ctk.CTkLabel(
            self.style_loading_card,
            text="Loading Product Master…",
            text_color=PURPLE_DARK,
            font=ctk.CTkFont(size=15, weight="bold"),
        )
        self.style_loading_label.pack(padx=28, pady=(16, 3))
        self.style_loading_detail = ctk.CTkLabel(
            self.style_loading_card,
            text="Preparing the selected product.",
            text_color=TEXT_MUTED,
            font=ctk.CTkFont(size=11),
        )
        self.style_loading_detail.pack(padx=28, pady=(0, 16))

    def _selected_style_label(self):
        """Return a safe display label for the style about to be rendered."""
        if not self.filtered_style_keys:
            return ""
        position = max(0, min(self.style_position, len(self.filtered_style_keys) - 1))
        rows = self.get_style_rows(self.filtered_style_keys[position])
        return first_nonblank(rows["Style Number"]) or best_product_name(rows["Product Name"])

    def _show_style_loading(self):
        if not hasattr(self, "style_loading_card"):
            return
        style = self._selected_style_label()
        self.style_loading_label.configure(
            text=f"Loading {style}…" if style else "Loading Product Master…"
        )
        self.style_loading_card.place(relx=0.5, rely=0.55, anchor="center")
        self.style_loading_card.lift()
        try:
            self.status_label.configure(text="Loading selected product…")
            # Process geometry requests only.  The actual style render is
            # deliberately scheduled for the next event-loop turn below.
            self.update_idletasks()
        except Exception:
            pass

    def _hide_style_loading(self, generation=None):
        if generation is not None and generation != self._style_render_generation:
            return
        if hasattr(self, "style_loading_card"):
            self.style_loading_card.place_forget()

    def _schedule_current_style_render(self):
        """Render only the newest requested product after macOS can repaint."""
        self._style_render_generation += 1
        generation = self._style_render_generation
        if self._style_render_after_id:
            try:
                self.after_cancel(self._style_render_after_id)
            except Exception:
                pass
        self._show_style_loading()

        def render_if_current():
            if generation != self._style_render_generation:
                return
            self._style_render_after_id = None
            self._render_current_style(generation)

        # A small real delay is intentional: after_idle can run before macOS
        # receives the expose event that draws the editor window.
        self._style_render_after_id = self.after(18, render_if_current)

    def _schedule_editor_catalog_refresh(self, generation):
        """Compatibility no-op after removing the redundant in-editor catalog."""
        return

    def _on_editor_search_changed(self, *_):
        if not self._suspend_editor_filter_trace:
            if self._editor_filter_after_id:
                try:
                    self.after_cancel(self._editor_filter_after_id)
                except Exception:
                    pass

            def apply_after_typing_pause():
                self._editor_filter_after_id = None
                self.apply_filters(reset_position=True)

            # Avoid redrawing the selected form and up to 60 catalog cards for
            # every character while a user is typing a style or employee search.
            self._editor_filter_after_id = self.after(140, apply_after_typing_pause)

    def build_editor_header(self, parent):
        header = ctk.CTkFrame(parent, fg_color=CARD_BG, corner_radius=0, height=92)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_columnconfigure(1, weight=1)

        brand = ctk.CTkFrame(header, fg_color="transparent")
        brand.grid(row=0, column=0, rowspan=2, padx=(24, 12), pady=16, sticky="w")
        flower_path = Path(__file__).resolve().parent.parent / "assets" / "orchid_flower_purple.png"
        if flower_path.exists():
            source = PILImage.open(flower_path).convert("RGBA")
            self.editor_header_flower = ctk.CTkImage(light_image=source, dark_image=source, size=(46, 46))
            ctk.CTkLabel(brand, text="", image=self.editor_header_flower, width=50).pack(side="left", padx=(0, 10))
        ctk.CTkButton(
            brand, text="← Catalog", command=self.show_dashboard, width=96, height=36,
            fg_color="transparent", hover_color=PURPLE_LIGHT, border_width=1,
            border_color=PURPLE_BORDER, text_color=PURPLE,
            font=ctk.CTkFont(size=12, weight="bold"),
        ).pack(side="left")

        ctk.CTkLabel(
            header, text="Product Master", text_color=TEXT_DARK,
            font=ctk.CTkFont(size=27, weight="bold"), anchor="w",
        ).grid(row=0, column=1, sticky="sw", pady=(15, 0))
        ctk.CTkLabel(
            header, text="Manage permanent purchasing and decoration defaults for every future order.",
            text_color=TEXT_MUTED, font=ctk.CTkFont(size=12), anchor="w",
        ).grid(row=1, column=1, sticky="nw", pady=(2, 14))

        action_buttons = ctk.CTkFrame(header, fg_color="transparent")
        action_buttons.grid(row=0, column=2, rowspan=2, padx=(12, 8), pady=20)
        ctk.CTkButton(
            action_buttons, text="+ Add New Product", command=self.add_new_product, width=148, height=36,
            fg_color="transparent", hover_color=PURPLE_LIGHT, border_width=1,
            border_color=PURPLE, text_color=PURPLE,
            font=ctk.CTkFont(size=12, weight="bold"),
        ).pack(pady=(0, 5))
        ctk.CTkButton(
            action_buttons, text="Delete Product", command=self.delete_current_product, width=148, height=32,
            fg_color="transparent", hover_color="#FDECEC", border_width=1,
            border_color="#C94B4B", text_color="#A93232",
            font=ctk.CTkFont(size=11, weight="bold"),
        ).pack()

        progress_card = ctk.CTkFrame(
            header, fg_color=PURPLE_LIGHT, border_width=1, border_color=PURPLE_BORDER,
            corner_radius=12, width=290, height=58,
        )
        progress_card.grid(row=0, column=3, rowspan=2, padx=(8, 24), pady=17)
        self.progress_text = ctk.CTkLabel(
            progress_card, text="", text_color=TEXT_DARK,
            font=ctk.CTkFont(size=12, weight="bold"),
        )
        self.progress_text.pack(padx=16, pady=(10, 3))
        self.progress_bar = ctk.CTkProgressBar(
            progress_card, width=240, height=8, progress_color=PURPLE, fg_color="#E1D8EC",
        )
        self.progress_bar.pack(padx=16, pady=(0, 10))

    def set_catalog_filter(self, value: str, open_first: bool = False):
        self.catalog_filter_var.set(value)
        self.incomplete_only_var.set(value == "Needs Setup")
        self._refresh_filter_button_styles()
        self.apply_filters(reset_position=True)
        if open_first and self.filtered_style_keys:
            self.style_position = 0
            self.show_current_style()

    def _refresh_filter_button_styles(self):
        selected = self.catalog_filter_var.get()
        for value, button in getattr(self, "filter_buttons", {}).items():
            active = value == selected
            button.configure(
                fg_color=PURPLE if active else PURPLE_LIGHT,
                hover_color=PURPLE_DARK if active else "#E4D6F5",
                text_color="#FFFFFF" if active else TEXT_DARK,
            )

    def continue_setup(self):
        self.open_needs_setup_queue()

    def build_filter_bar(self, parent):
        bar = ctk.CTkFrame(parent, fg_color="transparent", corner_radius=0)
        bar.grid(row=1, column=0, sticky="ew", padx=20, pady=(12, 6))
        bar.grid_columnconfigure(0, weight=1)

        self.search_entry = ctk.CTkEntry(
            bar, textvariable=self.search_var,
            placeholder_text="Search product number, name, vendor, category, alias, or color...",
            height=44, corner_radius=10, border_color=PURPLE_BORDER,
            fg_color="#FFFFFF", text_color=TEXT_DARK, placeholder_text_color="#9A92A4",
        )
        self.search_entry.grid(row=0, column=0, sticky="ew", padx=(0, 12))
        self.search_entry.bind("<Return>", lambda _event: self.open_first_editor_match())

        self.filter_buttons = {}
        total = len(self.all_style_keys)
        incomplete = len(self.incomplete_style_keys())
        complete = total - incomplete
        labels = {
            "All Products": f"{total} Products",
            "Needs Setup": f"Needs Setup ({incomplete})",
            "Complete": f"{complete} Ready",
        }
        for col, value in enumerate(("All Products", "Needs Setup", "Complete"), start=1):
            button = ctk.CTkButton(
                bar, text=labels[value], width=126, height=40,
                command=lambda v=value: self.set_catalog_filter(v, open_first=(v == "Needs Setup")),
                corner_radius=20, border_width=1, border_color=PURPLE_BORDER,
                font=ctk.CTkFont(size=11, weight="bold"),
            )
            button.grid(row=0, column=col, padx=3)
            self.filter_buttons[value] = button
        self.set_up_next_button = ctk.CTkButton(
            bar,
            text="Set Up Next  ➜",
            command=self.open_needs_setup_queue,
            width=132,
            height=40,
            fg_color=PURPLE,
            hover_color=PURPLE_DARK,
            font=ctk.CTkFont(size=11, weight="bold"),
        )
        self.set_up_next_button.grid(row=0, column=4, padx=(8, 0))
        self._refresh_filter_button_styles()

        ctk.CTkButton(
            bar, text="Clear", command=self.browse_all, width=72, height=40,
            fg_color="transparent", hover_color=PURPLE_LIGHT,
            border_width=1, border_color=PURPLE_BORDER, text_color=PURPLE,
        ).grid(row=0, column=5, padx=(8, 0))

    def build_style_banner(self, parent):
        banner = ctk.CTkFrame(
            parent, fg_color=CARD_BG, border_width=1,
            border_color=PURPLE_BORDER, corner_radius=14,
        )
        banner.grid(row=2, column=0, sticky="ew", padx=20, pady=6)
        banner.grid_columnconfigure(1, weight=1)

        ctk.CTkButton(
            banner, text="◀", command=self.previous_style, width=42, height=42,
            fg_color="transparent", hover_color=PURPLE_LIGHT,
            border_width=1, border_color=PURPLE_BORDER, text_color=PURPLE,
            font=ctk.CTkFont(size=17, weight="bold"),
        ).grid(row=0, column=0, rowspan=3, padx=(18, 12), pady=16)

        self.style_count_label = ctk.CTkLabel(
            banner, text="", text_color=TEXT_MUTED,
            font=ctk.CTkFont(size=11, weight="bold"), anchor="w",
        )
        self.style_count_label.grid(row=0, column=1, sticky="sw", pady=(12, 0))
        self.style_number_label = ctk.CTkLabel(
            banner, text="", text_color=PURPLE_DARK,
            font=ctk.CTkFont(size=25, weight="bold"), anchor="w",
        )
        self.style_number_label.grid(row=1, column=1, sticky="w", pady=(0, 0))
        self.product_name_label = ctk.CTkLabel(
            banner, text="", text_color=TEXT_DARK,
            font=ctk.CTkFont(size=15, weight="bold"), anchor="w", wraplength=650,
        )
        self.product_name_label.grid(row=2, column=1, sticky="nw", pady=(0, 12))

        status_area = ctk.CTkFrame(banner, fg_color="transparent")
        status_area.grid(row=0, column=2, rowspan=3, padx=12, pady=14, sticky="e")
        self.product_status_badge = ctk.CTkLabel(
            status_area, text="", height=28, corner_radius=14,
            fg_color="#EAF8EF", text_color=SUCCESS,
            font=ctk.CTkFont(size=10, weight="bold"),
        )
        self.product_status_badge.pack(anchor="e", pady=(2, 5), padx=6)
        self.issue_label = ctk.CTkLabel(
            status_area, text="", text_color=TEXT_MUTED,
            font=ctk.CTkFont(size=10), justify="right", wraplength=330,
        )
        self.issue_label.pack(anchor="e", padx=6)

        flower_path = Path(__file__).resolve().parent.parent / "assets" / "orchid_flower_purple.png"
        if flower_path.exists():
            flower = PILImage.open(flower_path).convert("RGBA")
            alpha = flower.getchannel("A").point(lambda value: int(value * 0.48))
            flower.putalpha(alpha)
            self.editor_watermark_image = ctk.CTkImage(
                light_image=flower, dark_image=flower, size=(84, 84),
            )
            ctk.CTkLabel(
                banner, text="", image=self.editor_watermark_image, width=88,
            ).grid(row=0, column=3, rowspan=3, padx=(2, 8), pady=4)

        ctk.CTkButton(
            banner, text="▶", command=self.next_style, width=42, height=42,
            fg_color="transparent", hover_color=PURPLE_LIGHT,
            border_width=1, border_color=PURPLE_BORDER, text_color=PURPLE,
            font=ctk.CTkFont(size=17, weight="bold"),
        ).grid(row=0, column=4, rowspan=3, padx=(4, 18), pady=16)

    def build_content_area(self, parent):
        content = ctk.CTkFrame(parent, fg_color="transparent")
        content.grid(row=4, column=0, sticky="nsew", padx=20, pady=8)
        content.grid_columnconfigure(0, weight=1)
        content.grid_rowconfigure(0, weight=1)

        workspace = ctk.CTkFrame(content, fg_color="transparent", corner_radius=0)
        workspace.grid(row=0, column=0, sticky="nsew")
        workspace.grid_columnconfigure(0, weight=38, minsize=410)
        workspace.grid_columnconfigure(1, weight=62, minsize=860)
        workspace.grid_rowconfigure(0, weight=1)
        self.build_left_panel(workspace)
        self.build_color_panel(workspace)

    def build_editor_catalog(self, parent):
        catalog = ctk.CTkFrame(
            parent, fg_color=CARD_BG, border_width=1,
            border_color=PURPLE_BORDER, corner_radius=14,
        )
        catalog.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        catalog.grid_columnconfigure(0, weight=1)
        catalog.grid_rowconfigure(2, weight=1)

        title_row = ctk.CTkFrame(catalog, fg_color="transparent")
        title_row.grid(row=0, column=0, sticky="ew", padx=16, pady=(15, 4))
        title_row.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            title_row, text="Product Catalog", text_color=TEXT_DARK,
            font=ctk.CTkFont(size=17, weight="bold"), anchor="w",
        ).grid(row=0, column=0, sticky="ew")
        ctk.CTkButton(
            title_row, text="+", command=self.add_new_product, width=34, height=32,
            corner_radius=16, fg_color=PURPLE, hover_color=PURPLE_DARK,
            font=ctk.CTkFont(size=17, weight="bold"),
        ).grid(row=0, column=1)
        self.catalog_result_label = ctk.CTkLabel(
            catalog, text="", text_color=TEXT_MUTED,
            font=ctk.CTkFont(size=10), anchor="w",
        )
        self.catalog_result_label.grid(row=1, column=0, sticky="ew", padx=16, pady=(0, 8))

        self.editor_catalog_scroll = ctk.CTkScrollableFrame(
            catalog, fg_color="#FBFAFD", corner_radius=10,
            border_width=1, border_color="#E8E0F1",
        )
        self.editor_catalog_scroll.grid(row=2, column=0, sticky="nsew", padx=12, pady=(0, 12))
        self.editor_catalog_scroll.grid_columnconfigure(0, weight=1)

    def select_editor_catalog_style(self, key):
        if key not in self.filtered_style_keys:
            return
        self.style_position = self.filtered_style_keys.index(key)
        self.show_current_style(defer=True)

    def refresh_editor_catalog(self):
        scroll = getattr(self, "editor_catalog_scroll", None)
        if scroll is None:
            return
        for widget in scroll.winfo_children():
            widget.destroy()
        keys = list(self.filtered_style_keys)
        self.catalog_result_label.configure(
            text=f"{len(keys)} product{'s' if len(keys) != 1 else ''} in this view"
        )
        if not keys:
            ctk.CTkLabel(
                scroll, text="No matching products", text_color=TEXT_MUTED,
                font=ctk.CTkFont(size=13, weight="bold"),
            ).grid(row=0, column=0, padx=16, pady=32)
            return
        visible = keys[:MAX_EDITOR_CATALOG_ROWS]
        for row_index, key in enumerate(visible):
            metadata = self.style_editor_metadata_cache.get(key, {})
            style = metadata.get("style", "No product number")
            name = metadata.get("name", "Unnamed product")
            vendor = self.style_vendor_cache.get(key, "Unassigned")
            category = self.style_category_cache.get(key, "Category not assigned")
            issues = self.style_key_issues(key)
            selected = key == self.current_style_key
            fill = "#F1E8FC" if selected else ("#FFFFFF" if row_index % 2 == 0 else "#FCFAFE")
            border = PURPLE if selected else "#E9E2F0"
            item = ctk.CTkFrame(
                scroll, fg_color=fill, border_width=2 if selected else 1,
                border_color=border, corner_radius=10,
            )
            item.grid(row=row_index, column=0, sticky="ew", padx=4, pady=4)
            item.grid_columnconfigure(0, weight=1)
            text = f"{style}\n{name}"
            ctk.CTkButton(
                item, text=text, command=lambda target=key: self.select_editor_catalog_style(target),
                fg_color="transparent", hover_color=PURPLE_LIGHT,
                text_color=TEXT_DARK, anchor="w", height=52,
                font=ctk.CTkFont(size=12, weight="bold"),
            ).grid(row=0, column=0, sticky="ew", padx=(6, 2), pady=(5, 0))
            detail = f"{vendor}  •  {category}"
            ctk.CTkLabel(
                item, text=detail, text_color=TEXT_MUTED,
                font=ctk.CTkFont(size=9), anchor="w", wraplength=225,
            ).grid(row=1, column=0, sticky="ew", padx=14, pady=(0, 7))
            badge_text = "READY" if not issues else "NEEDS SETUP"
            badge_fill = "#EAF8EF" if not issues else "#FFF1DC"
            badge_color = SUCCESS if not issues else WARNING
            ctk.CTkLabel(
                item, text=badge_text, height=24, corner_radius=12,
                fg_color=badge_fill, text_color=badge_color,
                font=ctk.CTkFont(size=9, weight="bold"),
            ).grid(row=0, column=1, rowspan=2, padx=(2, 8), pady=12)
        if len(keys) > len(visible):
            ctk.CTkLabel(
                scroll, text=f"Showing first {len(visible)} products. Refine the search to see more.",
                text_color=WARNING, font=ctk.CTkFont(size=10, weight="bold"),
                wraplength=230,
            ).grid(row=len(visible), column=0, padx=12, pady=12)

    def build_left_panel(self, parent):
        left = ctk.CTkScrollableFrame(parent, fg_color="transparent", corner_radius=0)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        left.grid_columnconfigure(0, weight=1)
        self.left_scroll = left

        info = ctk.CTkFrame(left, fg_color=CARD_BG, border_width=1, border_color=PURPLE_BORDER, corner_radius=12)
        info.grid(row=0, column=0, sticky="ew")
        info.grid_columnconfigure(0, weight=1)
        info.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(info, text="Product Information", text_color=PURPLE, font=ctk.CTkFont(size=17, weight="bold"), anchor="w").grid(row=0, column=0, columnspan=2, sticky="ew", padx=20, pady=(16, 10))

        ctk.CTkLabel(info, text="Product Number *", text_color=TEXT_DARK, font=ctk.CTkFont(size=12, weight="bold"), anchor="w").grid(row=1, column=0, sticky="ew", padx=(20, 8), pady=(3, 4))
        ctk.CTkLabel(info, text="Product Description *", text_color=TEXT_DARK, font=ctk.CTkFont(size=12, weight="bold"), anchor="w").grid(row=1, column=1, sticky="ew", padx=(8, 20), pady=(3, 4))
        ctk.CTkEntry(info, textvariable=self.style_number_var, height=42, corner_radius=9, border_color=PURPLE_BORDER, fg_color="#FFFFFF", text_color=TEXT_DARK).grid(row=2, column=0, sticky="ew", padx=(20, 8))
        ctk.CTkEntry(info, textvariable=self.product_name_var, height=42, corner_radius=9, border_color=PURPLE_BORDER, fg_color="#FFFFFF", text_color=TEXT_DARK).grid(row=2, column=1, sticky="ew", padx=(8, 20))
        ctk.CTkLabel(info, text="Example: K500, 1062, ST660", text_color=TEXT_MUTED, font=ctk.CTkFont(size=10), anchor="w").grid(row=3, column=0, sticky="ew", padx=(20, 8), pady=(3, 8))
        ctk.CTkLabel(info, text="Permanent catalog description", text_color=TEXT_MUTED, font=ctk.CTkFont(size=10), anchor="w").grid(row=3, column=1, sticky="ew", padx=(8, 20), pady=(3, 8))

        ctk.CTkLabel(info, text="Purchase Vendor *", text_color=TEXT_DARK, font=ctk.CTkFont(size=12, weight="bold"), anchor="w").grid(row=4, column=0, columnspan=2, sticky="ew", padx=20, pady=(4, 4))
        vendor_row = ctk.CTkFrame(info, fg_color="transparent")
        vendor_row.grid(row=5, column=0, columnspan=2, sticky="ew", padx=20)
        vendor_row.grid_columnconfigure(0, weight=1)
        self.vendor_combo = ctk.CTkComboBox(vendor_row, variable=self.vendor_var, values=self.vendor_options, height=42, corner_radius=9, border_color=PURPLE_BORDER, button_color="#EAE3F5", button_hover_color="#DED2EF", fg_color="#FFFFFF", text_color=TEXT_DARK, dropdown_fg_color="#FFFFFF", dropdown_text_color=TEXT_DARK, command=self.on_vendor_selected)
        self.vendor_combo.grid(row=0, column=0, sticky="ew", padx=(0, 7))
        ctk.CTkButton(
            vendor_row, text="+ Add Vendor", command=self.add_vendor_option,
            width=108, height=38, fg_color="transparent", hover_color=PURPLE_LIGHT,
            border_width=1, border_color=PURPLE, text_color=PURPLE,
            font=ctk.CTkFont(size=11, weight="bold"),
        ).grid(row=0, column=1)
        ctk.CTkLabel(info, text="Choose a saved vendor, type a name, or use Add Vendor. Saving this product keeps the vendor in future dropdowns.", text_color=TEXT_MUTED, font=ctk.CTkFont(size=10), anchor="w", wraplength=520).grid(row=6, column=0, columnspan=2, sticky="ew", padx=20, pady=(3, 8))

        ctk.CTkLabel(info, text="Product Category *", text_color=TEXT_DARK, font=ctk.CTkFont(size=12, weight="bold"), anchor="w").grid(row=7, column=0, columnspan=2, sticky="ew", padx=20, pady=(4, 4))
        self.category_combo = ctk.CTkComboBox(info, variable=self.category_var, values=alphabetical_options(CATEGORY_OPTIONS), height=38, command=self.on_category_selected, border_color=PURPLE_BORDER, button_color="#EAE3F5", button_hover_color="#DED2EF", fg_color="#FFFFFF", text_color=TEXT_DARK, dropdown_fg_color="#FFFFFF", dropdown_text_color=TEXT_DARK)
        self.category_combo.grid(row=8, column=0, columnspan=2, sticky="ew", padx=20)
        ctk.CTkLabel(info, text="Category supplies sensible defaults for size, color, and decoration requirements.", text_color=TEXT_MUTED, font=ctk.CTkFont(size=10), anchor="w", wraplength=520).grid(row=9, column=0, columnspan=2, sticky="ew", padx=20, pady=(3, 14))

        rules = ctk.CTkFrame(left, fg_color=CARD_BG, border_width=1, border_color=PURPLE_BORDER, corner_radius=12)
        rules.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        rules.grid_columnconfigure(0, weight=1)
        rules.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(rules, text="Purchasing Rules", text_color=PURPLE, font=ctk.CTkFont(size=17, weight="bold"), anchor="w").grid(row=0, column=0, columnspan=2, sticky="ew", padx=18, pady=(14, 6))
        ctk.CTkLabel(rules, text="Orchid only flags information that is required for this product.", text_color=TEXT_MUTED, font=ctk.CTkFont(size=10), anchor="w", wraplength=520).grid(row=1, column=0, columnspan=2, sticky="ew", padx=18, pady=(0, 8))
        self.requires_size_switch = ctk.CTkSwitch(rules, text="Size is required", variable=self.requires_size_var, onvalue=True, offvalue=False, progress_color=PURPLE)
        self.requires_size_switch.grid(row=2, column=0, sticky="w", padx=18, pady=5)
        self.requires_color_switch = ctk.CTkSwitch(rules, text="Color is required", variable=self.requires_color_var, onvalue=True, offvalue=False, progress_color=PURPLE)
        self.requires_color_switch.grid(row=2, column=1, sticky="w", padx=18, pady=5)
        self.requires_decoration_switch = ctk.CTkSwitch(rules, text="Decoration is required", variable=self.requires_decoration_var, onvalue=True, offvalue=False, progress_color=PURPLE, command=self.on_requires_decoration_changed)
        self.requires_decoration_switch.grid(row=3, column=0, sticky="w", padx=18, pady=5)
        self.never_outsource_switch = ctk.CTkSwitch(
            rules, text="Never outsource — ship to Orchid", variable=self.never_outsource_var,
            onvalue=True, offvalue=False, progress_color=PURPLE,
            command=self.on_never_outsource_changed,
        )
        self.never_outsource_switch.grid(row=3, column=1, sticky="w", padx=18, pady=(5, 3))
        ctk.CTkLabel(
            rules, text="Edwards items, headwear, bottoms, bags, and Flame Resistant products default to Yes. Turn off only for a deliberate exception.",
            text_color=TEXT_MUTED, font=ctk.CTkFont(size=10), anchor="w", wraplength=520,
        ).grid(row=4, column=0, columnspan=2, sticky="ew", padx=18, pady=(0, 14))

        self.advanced_button = ctk.CTkButton(left, text="Show Advanced Matching Fields", command=self.toggle_advanced_fields, height=36, fg_color="transparent", hover_color=PURPLE_LIGHT, border_width=1, border_color=PURPLE, text_color=PURPLE)
        self.advanced_button.grid(row=2, column=0, sticky="ew", pady=(12, 0))
        self.advanced_frame = ctk.CTkFrame(left, fg_color=CARD_BG, border_width=1, border_color=PURPLE_BORDER, corner_radius=12)
        self.advanced_frame.grid(row=3, column=0, sticky="ew", pady=(8, 0))
        self.advanced_frame.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(self.advanced_frame, text="Advanced Matching", text_color=PURPLE, font=ctk.CTkFont(size=16, weight="bold"), anchor="w").grid(row=0, column=0, sticky="ew", padx=18, pady=(14, 6))
        self.add_labeled_entry(self.advanced_frame, 1, "Product ID", self.product_id_var, "Generated internal identity; usually leave unchanged")
        self.add_labeled_entry(self.advanced_frame, 4, "Product Aliases", self.product_aliases_var, "Optional alternate names separated by commas")
        self.advanced_frame.grid_remove()

        tools = ctk.CTkFrame(left, fg_color=PURPLE_LIGHT, border_width=1, border_color=PURPLE_BORDER, corner_radius=12)
        tools.grid(row=4, column=0, sticky="ew", pady=(12, 8))
        tools.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(tools, text="Quick Tools", text_color=PURPLE, font=ctk.CTkFont(size=16, weight="bold"), anchor="w").grid(row=0, column=0, sticky="ew", padx=18, pady=(14, 8))
        self.apply_color_combo = ctk.CTkComboBox(tools, variable=self.apply_color_var, values=self.decoration_color_options, height=40, corner_radius=9, border_color=PURPLE_BORDER, button_color="#EAE3F5", button_hover_color="#DED2EF", fg_color="#FFFFFF", text_color=TEXT_DARK)
        self.apply_color_combo.grid(row=1, column=0, sticky="ew", padx=18)
        ctk.CTkButton(tools, text="Apply Thread/Ink Color to All", command=self.apply_color_to_all, height=36, fg_color=PURPLE, hover_color=PURPLE_DARK).grid(row=2, column=0, sticky="ew", padx=18, pady=(8, 6))
        ctk.CTkButton(tools, text="Copy First Completed Color to All", command=self.copy_first_color_to_all, height=34, fg_color="transparent", hover_color="#E7DCF6", border_width=1, border_color=PURPLE, text_color=PURPLE).grid(row=3, column=0, sticky="ew", padx=18, pady=(0, 14))

    def add_labeled_entry(self, parent, row, label, variable, helper):
        ctk.CTkLabel(parent, text=label, text_color=TEXT_DARK, font=ctk.CTkFont(size=12, weight="bold"), anchor="w").grid(row=row, column=0, sticky="ew", padx=20, pady=(3, 4))
        ctk.CTkEntry(parent, textvariable=variable, height=42, corner_radius=9, border_color=PURPLE_BORDER, fg_color="#FFFFFF", text_color=TEXT_DARK).grid(row=row + 1, column=0, sticky="ew", padx=20)
        ctk.CTkLabel(parent, text=helper, text_color=TEXT_MUTED, font=ctk.CTkFont(size=10), anchor="w").grid(row=row + 2, column=0, sticky="ew", padx=20, pady=(3, 8))

    def build_color_panel(self, parent):
        panel = ctk.CTkFrame(parent, fg_color=CARD_BG, border_width=1, border_color=PURPLE_BORDER, corner_radius=12)
        panel.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        panel.grid_columnconfigure(0, weight=1)
        # The color list is the final grid row.  Giving row 4 the flexible
        # height left a dead strip below it whenever the Product Master window
        # was tall.  Keep the purchase-color work area usable all the way down
        # to the action bar, including for styles with many color mappings.
        panel.grid_rowconfigure(3, weight=1)
        title_row = ctk.CTkFrame(panel, fg_color="transparent")
        title_row.grid(row=0, column=0, sticky="ew", padx=22, pady=(16, 4))
        title_row.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(title_row, text="Decoration Defaults & Purchasing Colors", text_color=PURPLE, font=ctk.CTkFont(size=17, weight="bold"), anchor="w").grid(row=0, column=0, sticky="ew")
        ctk.CTkLabel(
            title_row,
            text="Key Vendor Mapping: Use Purchase As Style # when the vendor order number differs from your Shopify style — often by color.",
            text_color=PURPLE_DARK,
            font=ctk.CTkFont(size=10, weight="bold"),
            anchor="w",
            wraplength=690,
        ).grid(row=1, column=0, sticky="ew", pady=(2, 0))
        ctk.CTkButton(title_row, text="+ Add Garment Color", command=self.add_color_row, width=150, height=34, fg_color="transparent", hover_color=PURPLE_LIGHT, border_width=1, border_color=PURPLE, text_color=PURPLE).grid(row=0, column=1, rowspan=2, sticky="n")
        helper_row = ctk.CTkFrame(panel, fg_color="transparent")
        helper_row.grid(row=1, column=0, sticky="ew", padx=22, pady=(0, 8))
        helper_row.grid_columnconfigure(0, weight=1)
        self.color_setup_summary_label = ctk.CTkLabel(
            helper_row,
            text="Set the permanent decoration route first, then maintain purchasing colors below.",
            text_color=TEXT_MUTED,
            font=ctk.CTkFont(size=12),
            anchor="w",
        )
        self.color_setup_summary_label.grid(row=0, column=0, sticky="ew")
        ctk.CTkButton(
            helper_row, text="+ Add Thread/Ink Color", command=self.add_thread_ink_color_option,
            width=158, height=32, fg_color="transparent", hover_color=PURPLE_LIGHT,
            border_width=1, border_color=PURPLE, text_color=PURPLE,
        ).grid(row=0, column=1, padx=(8, 0))

        decoration_settings = ctk.CTkFrame(panel, fg_color=PURPLE_LIGHT, border_width=1, border_color=PURPLE_BORDER, corner_radius=10)
        decoration_settings.grid(row=2, column=0, sticky="ew", padx=22, pady=(0, 10))
        for column in range(3):
            decoration_settings.grid_columnconfigure(column, weight=1)

        ctk.CTkLabel(decoration_settings, text="Decoration Type *", text_color=TEXT_DARK, font=ctk.CTkFont(size=11, weight="bold"), anchor="w").grid(row=0, column=0, sticky="ew", padx=(14, 7), pady=(10, 4))
        ctk.CTkLabel(decoration_settings, text="Default Location *", text_color=TEXT_DARK, font=ctk.CTkFont(size=11, weight="bold"), anchor="w").grid(row=0, column=1, sticky="ew", padx=7, pady=(10, 4))
        self.custom_location_label = ctk.CTkLabel(decoration_settings, text="Custom Location *", text_color=TEXT_DARK, font=ctk.CTkFont(size=11, weight="bold"), anchor="w")
        self.custom_location_label.grid(row=0, column=2, sticky="ew", padx=(7, 14), pady=(10, 4))

        self.decoration_combo = ctk.CTkComboBox(decoration_settings, variable=self.decoration_type_var, values=DECORATION_TYPES, height=40, corner_radius=9, border_color=PURPLE_BORDER, button_color="#EAE3F5", button_hover_color="#DED2EF", fg_color="#FFFFFF", text_color=TEXT_DARK, dropdown_fg_color="#FFFFFF", dropdown_text_color=TEXT_DARK)
        self.decoration_combo.grid(row=1, column=0, sticky="ew", padx=(14, 7))
        self.decoration_location_combo = ctk.CTkComboBox(
            decoration_settings, variable=self.decoration_location_var, values=DECORATION_LOCATIONS, height=36,
            command=self.on_decoration_location_selected,
            border_color=PURPLE_BORDER, button_color="#EAE3F5", button_hover_color="#DED2EF",
            fg_color="#FFFFFF", text_color=TEXT_DARK, dropdown_fg_color="#FFFFFF", dropdown_text_color=TEXT_DARK,
        )
        self.decoration_location_combo.grid(row=1, column=1, sticky="ew", padx=7)
        self.custom_location_entry = ctk.CTkEntry(
            decoration_settings, textvariable=self.custom_decoration_location_var, height=36,
            placeholder_text="Right Chest or Center Back",
            border_color=PURPLE_BORDER, fg_color="#FFFFFF", text_color=TEXT_DARK,
        )
        self.custom_location_entry.grid(row=1, column=2, sticky="ew", padx=(7, 14))

        ctk.CTkLabel(decoration_settings, text="Placement Instructions (optional)", text_color=TEXT_DARK, font=ctk.CTkFont(size=11, weight="bold"), anchor="w").grid(row=2, column=0, columnspan=3, sticky="ew", padx=14, pady=(9, 4))
        self.placement_instructions_entry = ctk.CTkEntry(
            decoration_settings, textvariable=self.decoration_placement_instructions_var, height=36,
            placeholder_text="Example: 2 inches below shoulder seam",
            border_color=PURPLE_BORDER, fg_color="#FFFFFF", text_color=TEXT_DARK,
        )
        self.placement_instructions_entry.grid(row=3, column=0, columnspan=3, sticky="ew", padx=14, pady=(0, 10))
        self.on_decoration_location_selected(self.decoration_location_var.get())

        self.colors_scroll = ctk.CTkScrollableFrame(panel, fg_color="#FFFFFF", corner_radius=8, border_width=1, border_color=PURPLE_BORDER)
        self.colors_scroll.grid(row=3, column=0, sticky="nsew", padx=22, pady=(6, 18))
        self.colors_scroll.grid_columnconfigure(0, weight=1)

    def build_action_bar(self, parent):
        actions = ctk.CTkFrame(
            parent, fg_color=CARD_BG, border_width=1,
            border_color=PURPLE_BORDER, corner_radius=14,
        )
        actions.grid(row=5, column=0, sticky="ew", padx=20, pady=(4, 10))
        actions.grid_columnconfigure(1, weight=1)
        actions.grid_columnconfigure(2, weight=1)
        actions.grid_columnconfigure(3, weight=1)

        ctk.CTkButton(
            actions, text="← Product Catalog", command=self.show_dashboard,
            width=150, height=44, fg_color="transparent", hover_color=PURPLE_LIGHT,
            border_width=1, border_color=PURPLE_BORDER, text_color=PURPLE,
            font=ctk.CTkFont(size=12, weight="bold"),
        ).grid(row=0, column=0, padx=(14, 8), pady=12)
        ctk.CTkButton(
            actions, text="◀ Previous", command=self.previous_style,
            height=44, fg_color="transparent", hover_color=PURPLE_LIGHT,
            border_width=1, border_color=PURPLE_BORDER, text_color=PURPLE,
            font=ctk.CTkFont(size=12, weight="bold"),
        ).grid(row=0, column=1, sticky="ew", padx=6, pady=12)
        self.save_button = ctk.CTkButton(
            actions, text="Save Changes", command=self.save_current_style,
            height=44, fg_color="#EEE7F7", hover_color="#E3D7F2",
            border_width=1, border_color=PURPLE_BORDER, text_color=PURPLE_DARK,
            font=ctk.CTkFont(size=13, weight="bold"),
        )
        self.save_button.grid(row=0, column=2, sticky="ew", padx=6, pady=12)
        self.save_next_button = ctk.CTkButton(
            actions, text="Save & Next Setup Item", command=self.save_and_next_style,
            height=44, fg_color=PURPLE, hover_color=PURPLE_DARK,
            font=ctk.CTkFont(size=13, weight="bold"),
        )
        self.save_next_button.grid(row=0, column=3, sticky="ew", padx=6, pady=12)
        ctk.CTkButton(
            actions, text="Next ▶", command=self.next_style,
            width=110, height=44, fg_color="transparent", hover_color=PURPLE_LIGHT,
            border_width=1, border_color=PURPLE_BORDER, text_color=PURPLE,
            font=ctk.CTkFont(size=12, weight="bold"),
        ).grid(row=0, column=4, padx=(8, 14), pady=12)

    def build_footer(self, parent):
        footer = ctk.CTkFrame(parent, fg_color=PURPLE_DARK, corner_radius=0, height=64)
        footer.grid(row=6, column=0, sticky="ew")
        footer.grid_columnconfigure(1, weight=1)
        footer_brand = ctk.CTkFrame(footer, fg_color="transparent")
        footer_brand.grid(row=0, column=0, padx=26, pady=12, sticky="w")
        footer_flower = Path(__file__).resolve().parent.parent / "assets" / "orchid_flower_white.png"
        if footer_flower.exists():
            source = PILImage.open(footer_flower).convert("RGBA")
            self.footer_flower_image = ctk.CTkImage(light_image=source, dark_image=source, size=(28, 28))
            ctk.CTkLabel(footer_brand, text="", image=self.footer_flower_image, width=32).pack(side="left", padx=(0, 7))
        ctk.CTkLabel(footer_brand, text="ORCHID PURCHASING SYSTEM", text_color="#FFFFFF", font=ctk.CTkFont(size=14, weight="bold")).pack(side="left")
        self.status_label = ctk.CTkLabel(footer, text="", text_color="#FFFFFF", font=ctk.CTkFont(size=12, weight="bold"))
        self.status_label.grid(row=0, column=1, pady=18)
        self.footer_record_label = ctk.CTkLabel(footer, text="", text_color="#FFFFFF", font=ctk.CTkFont(size=11))
        self.footer_record_label.grid(row=0, column=2, padx=26, pady=18, sticky="e")

    # ---------- editor behavior ----------
    def show_editor(self, target_key=None):
        self.dashboard_frame.grid_remove()
        self.editor_frame.grid()
        self.apply_filters(reset_position=True, target_key=target_key)

    def browse_all(self):
        self.search_var.set("")
        self.incomplete_only_var.set(False)
        self.catalog_filter_var.set("All Products")
        self._refresh_filter_button_styles()
        self.apply_filters(reset_position=True)

    def get_matching_style_keys(self):
        query = self.search_var.get().strip()
        catalog_filter = self.catalog_filter_var.get() if hasattr(self, "catalog_filter_var") else ("Needs Setup" if self.incomplete_only_var.get() else "All Products")
        ranked = []
        for position, key in enumerate(self.all_style_keys):
            is_complete = self.style_key_is_complete(key)
            if catalog_filter == "Needs Setup" and is_complete:
                continue
            if catalog_filter == "Complete" and not is_complete:
                continue
            search = self.style_search_cache.get(key, {})
            score = search_match_score_values(
                search.get("style", ""), search.get("product_id", ""),
                search.get("name", ""), search.get("vendor", ""),
                search.get("combined", ""), query,
            )
            if score is None:
                continue
            ranked.append((score, position, key))
        ranked.sort(key=lambda item: (item[0], item[1]))
        return [key for _, _, key in ranked]

    def open_first_editor_match(self):
        matches = self.get_matching_style_keys()
        if matches:
            self.show_editor(target_key=matches[0])
        else:
            self.status_label.configure(text="No Product Master records match that search.")

    def apply_filters(self, reset_position=True, target_key=None):
        if self._editor_filter_after_id:
            try:
                self.after_cancel(self._editor_filter_after_id)
            except Exception:
                pass
            self._editor_filter_after_id = None
        self.filtered_style_keys = self.get_matching_style_keys()
        if target_key and target_key in self.filtered_style_keys:
            self.style_position = self.filtered_style_keys.index(target_key)
        elif reset_position:
            self.style_position = 0
        self.refresh_progress()
        self.show_current_style()

    def refresh_progress(self):
        total = len(self.all_style_keys)
        complete = self.completed_style_count()
        remaining = total - complete
        percent = complete / total if total else 0
        self.progress_text.configure(text=f"{total} saved styles • {complete} fully configured")
        self.progress_bar.set(percent)

    def _scroll_to_top(self):
        """Reset both editor panels for the newly displayed product."""
        for frame in (getattr(self, "left_scroll", None), getattr(self, "colors_scroll", None)):
            canvas = getattr(frame, "_parent_canvas", None) if frame is not None else None
            if canvas is not None:
                try:
                    canvas.yview_moveto(0.0)
                except Exception:
                    pass

    def _scroll_colors_to_bottom(self):
        canvas = getattr(getattr(self, "colors_scroll", None), "_parent_canvas", None)
        if canvas is not None:
            try:
                canvas.yview_moveto(1.0)
            except Exception:
                pass

    def clear_color_rows(self):
        for widget in self.colors_scroll.winfo_children():
            widget.destroy()
        self.color_controls = []

    def show_current_style(self, defer: bool = False):
        """Display the selected style, optionally yielding once for a clean redraw."""
        if defer:
            self._schedule_current_style_render()
            return
        self._style_render_generation += 1
        generation = self._style_render_generation
        if self._style_render_after_id:
            try:
                self.after_cancel(self._style_render_after_id)
            except Exception:
                pass
            self._style_render_after_id = None
        self._render_current_style(generation)

    def _render_current_style(self, generation=None):
        self._loading_style = True
        self.clear_color_rows()
        self.status_label.configure(text="")
        if not self.filtered_style_keys:
            self.current_style_key = None
            self.style_count_label.configure(text="No matching styles")
            if hasattr(self, "style_number_label"):
                self.style_number_label.configure(text="")
            self.product_name_label.configure(text="")
            if hasattr(self, "product_status_badge"):
                self.product_status_badge.configure(text="NO MATCHES", fg_color="#F2EFF5", text_color=TEXT_MUTED)
            self.issue_label.configure(text="")
            self.footer_record_label.configure(text="No records")
            self.product_name_var.set("")
            self.style_number_var.set("")
            self.vendor_var.set("")
            self.decoration_type_var.set("")
            self.decoration_location_var.set(LEFT_CHEST)
            self.custom_decoration_location_var.set("")
            self.decoration_placement_instructions_var.set("")
            self.product_id_var.set("")
            self.product_aliases_var.set("")
            self.category_var.set("")
            self.requires_size_var.set(True)
            self.requires_color_var.set(True)
            self.requires_decoration_var.set(True)
            self.never_outsource_var.set(False)
            self._loading_style = False
            self._schedule_editor_catalog_refresh(generation)
            self._hide_style_loading(generation)
            return

        self.style_position = max(0, min(self.style_position, len(self.filtered_style_keys) - 1))
        key = self.filtered_style_keys[self.style_position]
        self.current_style_key = key
        rows = self.get_style_rows(key)
        issues = self.style_issues(rows)
        self.style_count_label.configure(text=f"Product {self.style_position + 1} of {len(self.filtered_style_keys)} in this view")
        self.footer_record_label.configure(text=f"{len(self.all_style_keys)} saved styles")
        style_number = first_nonblank(rows["Style Number"]) or "No product number"
        product_name = best_product_name(rows["Product Name"]) or "Unnamed product"
        if hasattr(self, "style_number_label"):
            self.style_number_label.configure(text=style_number)
        self.product_name_label.configure(text=product_name)
        if hasattr(self, "product_status_badge"):
            self.product_status_badge.configure(
                text="READY" if not issues else "NEEDS SETUP",
                fg_color="#EAF8EF" if not issues else "#FFF1DC",
                text_color=SUCCESS if not issues else WARNING,
            )
        self.issue_label.configure(
            text="All permanent purchasing details are configured." if not issues else " • ".join(issues),
            text_color=SUCCESS if not issues else TEXT_MUTED,
        )

        self.product_name_var.set(best_product_name(rows["Product Name"]))
        self.style_number_var.set(first_nonblank(rows["Style Number"]))
        self.product_id_var.set(first_nonblank(rows["Product ID"]) or default_product_id(first_nonblank(rows["Style Number"]), best_product_name(rows["Product Name"])))
        self.product_aliases_var.set(first_nonblank(rows["Product Aliases"]))
        rules = row_rules(rows.iloc[0])
        self.category_var.set(rules["Product Category"])
        self.requires_size_var.set(normalize_bool(rules["Requires Size"], True))
        self.requires_color_var.set(normalize_bool(rules["Requires Color"], True))
        self.requires_decoration_var.set(normalize_bool(rules["Requires Decoration"], True))
        # Older Orchid versions wrote No into every blank row, which is not
        # evidence that the user deliberately overrode a newer category default.
        never_outsource = editor_never_outsource_value(
            key,
            self.never_outsource_overrides,
            first_nonblank(rows[NEVER_OUTSOURCE_COLUMN]),
            self.product_name_var.get(),
            self.category_var.get(),
            self.style_number_var.get(),
            first_nonblank(rows["Vendor"]),
        )
        self.never_outsource_var.set(never_outsource)
        self._never_outsource_touched = False
        vendors = list(dict.fromkeys(clean_text(v) for v in rows["Vendor"] if clean_text(v)))
        decorations = list(dict.fromkeys(clean_text(v) for v in rows["Decoration Type"] if clean_text(v)))
        locations = list(dict.fromkeys(clean_text(v) for v in rows["Decoration Location"] if clean_text(v)))
        self.vendor_var.set(vendors[0] if len(vendors) == 1 else "")
        self.decoration_type_var.set(decorations[0] if len(decorations) == 1 else "")
        stored_location = normalize_decoration_location(
            locations[0] if len(locations) == 1 else "",
            self.decoration_type_var.get(),
            requires_decoration=self.requires_decoration_var.get(),
            product_name=self.product_name_var.get(), category=self.category_var.get(),
        )
        location_choice, custom_location = location_choice_and_custom(stored_location, self.decoration_type_var.get())
        self.decoration_location_var.set(location_choice)
        self.custom_decoration_location_var.set(custom_location)
        instructions = list(dict.fromkeys(
            clean_text(v) for v in rows.get("Decoration Placement Instructions", pd.Series(dtype=str)) if clean_text(v)
        ))
        self.decoration_placement_instructions_var.set(instructions[0] if len(instructions) == 1 else "")
        self.on_decoration_location_selected(location_choice)

        # Apply visible defaults immediately for any legacy/incomplete record that
        # predates the Product Intelligence Engine. These remain editable.
        detected_brand, preferred_vendor = preferred_vendor_for_brand_text(
            f"{self.product_name_var.get()} {self.style_number_var.get()}"
        )
        if not self.vendor_var.get().strip() and preferred_vendor:
            self.vendor_var.set(preferred_vendor)
        if not self.decoration_type_var.get().strip():
            inferred_decoration = infer_default_decoration(
                self.category_var.get(), self.product_name_var.get(), self.style_number_var.get()
            )
            if inferred_decoration:
                self.decoration_type_var.set(inferred_decoration)

        self._update_color_setup_guidance(rows)
        sorted_rows = prioritize_new_color_setup_rows(rows)
        for display_index, (data_index, row) in enumerate(sorted_rows.iterrows(), start=1):
            self.render_color_row(display_index, data_index, row)
        self.update_decoration_color_state()
        self._loading_style = False
        self.on_vendor_changed()
        self.after_idle(self._scroll_to_top)
        self._schedule_editor_catalog_refresh(generation)
        self._hide_style_loading(generation)

    def render_color_row(self, display_index, data_index, row):
        is_new_color = is_new_color_setup_row(row)
        row_color = "#FFF6E8" if is_new_color else (ROW_ALT if display_index % 2 == 0 else "#FFFFFF")
        frame = ctk.CTkFrame(
            self.colors_scroll,
            fg_color=row_color,
            border_width=1 if is_new_color else 0,
            border_color="#E3A64C" if is_new_color else row_color,
            corner_radius=8,
            height=108,
        )
        frame.grid(row=display_index - 1, column=0, sticky="ew", padx=4, pady=4)
        configure_color_table_columns(frame)

        # Keep every mapping control on one left-to-right row.  This is possible
        # now that the duplicate Product Catalog is no longer using the first
        # third of the editor window.
        garment_var = ctk.StringVar(value=clean_text(row.get("Garment Color", "")))
        vendor_code_var = ctk.StringVar(value=clean_text(row.get("Vendor Color Code", "")))
        aliases_var = ctk.StringVar(value=clean_text(row.get("Color Aliases", "")))
        purchasing_style_var = ctk.StringVar(value=clean_text(row.get("Purchasing Style Number", "")))
        decoration_var = ctk.StringVar(value=clean_text(row.get("Decoration Color", "")))
        decision_border = "#D98924" if is_new_color else PURPLE_BORDER
        label_font = ctk.CTkFont(size=10, weight="bold")
        for column, (label, _minimum) in enumerate(COLOR_TABLE_COLUMNS):
            if column == PURCHASE_AS_STYLE_COLUMN:
                mapping_heading = ctk.CTkFrame(
                    frame, fg_color=KEY_VENDOR_MAPPING_BADGE_BG,
                    corner_radius=7,
                )
                mapping_heading.grid(
                    row=0, column=column, sticky="ew", padx=5, pady=(8, 3),
                )
                mapping_heading.grid_columnconfigure(0, weight=1)
                ctk.CTkLabel(
                    mapping_heading, text=label, text_color=PURPLE_DARK,
                    font=label_font, anchor="w",
                ).grid(row=0, column=0, sticky="ew", padx=8, pady=(4, 0))
                ctk.CTkLabel(
                    mapping_heading, text="Vendor order number", text_color=PURPLE,
                    font=ctk.CTkFont(size=9), anchor="w",
                ).grid(row=1, column=0, sticky="ew", padx=8, pady=(0, 4))
                continue
            ctk.CTkLabel(
                frame,
                text=label,
                text_color="#8A4F00" if is_new_color else PURPLE,
                font=label_font,
                anchor="w",
            ).grid(row=0, column=column, sticky="ew", padx=(10 if column == 0 else 5, 5), pady=(10, 3))
        ctk.CTkLabel(
            frame,
            text="Actions",
            text_color="#8A4F00" if is_new_color else PURPLE,
            font=label_font,
            anchor="center",
        ).grid(row=0, column=len(COLOR_TABLE_COLUMNS), sticky="ew", padx=(5, 10), pady=(10, 3))
        garment_entry = ctk.CTkEntry(frame, textvariable=garment_var, height=38, corner_radius=9, border_color=decision_border, fg_color="#FFFFFF", text_color=TEXT_DARK)
        garment_entry.grid(row=1, column=0, sticky="ew", padx=(10, 5), pady=(0, 10))
        ctk.CTkEntry(frame, textvariable=vendor_code_var, height=38, corner_radius=9, border_color=PURPLE_BORDER, fg_color="#FFFFFF", text_color=TEXT_DARK).grid(row=1, column=1, sticky="ew", padx=5, pady=(0, 10))
        ctk.CTkEntry(frame, textvariable=aliases_var, placeholder_text="e.g. Dark Indigo, Navy", height=38, corner_radius=9, border_color=PURPLE_BORDER, fg_color="#FFFFFF", text_color=TEXT_DARK).grid(row=1, column=2, sticky="ew", padx=5, pady=(0, 10))
        purchase_as_style_entry = ctk.CTkEntry(
            frame, textvariable=purchasing_style_var, placeholder_text="e.g. 1062",
            height=38, corner_radius=9, border_width=2, border_color=PURPLE,
            fg_color=KEY_VENDOR_MAPPING_BG, text_color=PURPLE_DARK,
        )
        purchase_as_style_entry.grid(
            row=1, column=PURCHASE_AS_STYLE_COLUMN, sticky="ew", padx=5, pady=(0, 10),
        )
        purchase_as_style_badge = ctk.CTkLabel(
            frame, text="Custom Vendor Style", height=20, corner_radius=10,
            fg_color=KEY_VENDOR_MAPPING_BADGE_BG, text_color=PURPLE_DARK,
            font=ctk.CTkFont(size=9, weight="bold"),
        )

        def refresh_purchase_as_style_badge(*_args):
            if normalize_space(purchasing_style_var.get()):
                purchase_as_style_badge.grid(
                    row=2, column=PURCHASE_AS_STYLE_COLUMN, sticky="w", padx=7, pady=(0, 7),
                )
            else:
                purchase_as_style_badge.grid_remove()

        purchasing_style_var.trace_add("write", refresh_purchase_as_style_badge)
        refresh_purchase_as_style_badge()
        decoration_combo = ctk.CTkComboBox(
            frame, variable=decoration_var, values=self.decoration_color_options, height=38, corner_radius=9,
            border_color=decision_border, button_color="#EAE3F5", button_hover_color="#DED2EF",
            fg_color="#FFFFFF", text_color=TEXT_DARK, dropdown_fg_color="#FFFFFF", dropdown_text_color=TEXT_DARK,
            command=lambda _value, index=data_index: self._set_active_color_row(index),
        )
        decoration_combo.grid(row=1, column=4, sticky="ew", padx=5, pady=(0, 10))
        decoration_combo.bind("<Button-1>", lambda _event, index=data_index: self._set_active_color_row(index), add="+")
        current_rows = self.get_style_rows(self.current_style_key) if self.current_style_key else pd.DataFrame()
        is_boot_base_row = (
            len(current_rows) <= 1
            and is_boots_product(
                self.product_name_var.get(), self.category_var.get(), self.style_number_var.get()
            )
        )
        delete_button = ctk.CTkButton(
            frame, text="Clear" if is_boot_base_row else "Delete", width=64, height=38,
            fg_color="#FFF1F0", hover_color="#FDE2E0",
            border_width=1, border_color="#E5A09A", text_color=DANGER,
            command=lambda index=data_index: self.delete_color_row(index),
        )
        delete_button.grid(row=1, column=len(COLOR_TABLE_COLUMNS), sticky="ew", padx=(5, 10), pady=(0, 10))
        self.color_controls.append({
            "index": data_index,
            "garment_var": garment_var,
            "garment_entry": garment_entry,
            "vendor_code_var": vendor_code_var,
            "aliases_var": aliases_var,
            "purchasing_style_var": purchasing_style_var,
            "purchase_as_style_entry": purchase_as_style_entry,
            "purchase_as_style_badge": purchase_as_style_badge,
            "decoration_var": decoration_var,
            "decoration_combo": decoration_combo,
            "delete_button": delete_button,
            "is_new_color": is_new_color,
        })

    def toggle_advanced_fields(self):
        self.advanced_visible = not self.advanced_visible
        if self.advanced_visible:
            self.advanced_frame.grid()
            self.advanced_button.configure(text="Hide Advanced Matching Fields")
        else:
            self.advanced_frame.grid_remove()
            self.advanced_button.configure(text="Show Advanced Matching Fields")

    def on_vendor_selected(self, selected_vendor):
        self.on_vendor_changed()

    def on_vendor_changed(self, *_args):
        """Apply vendor defaults while preserving a deliberate saved override."""
        if getattr(self, "_loading_style", False) or self._never_outsource_touched:
            return
        if is_boots_product(
            self.product_name_var.get(), self.category_var.get(), self.style_number_var.get()
        ):
            self.never_outsource_var.set(True)
            return
        if self.current_style_key in self.never_outsource_overrides:
            return
        if vendor_never_outsource(self.vendor_var.get()):
            self.never_outsource_var.set(True)
            if hasattr(self, "status_label"):
                self.status_label.configure(
                    text=f"{self.vendor_var.get().strip()} defaults to Never Outsource — ship to Orchid."
                )

    def on_category_selected(self, selected_category):
        defaults = category_defaults(selected_category, self.product_name_var.get(), self.style_number_var.get())
        self.requires_size_var.set(defaults["requires_size"])
        self.requires_color_var.set(defaults["requires_color"])
        self.requires_decoration_var.set(defaults["requires_decoration"])
        self.never_outsource_var.set(
            vendor_never_outsource(self.vendor_var.get())
            or default_never_outsource(
                self.product_name_var.get(), selected_category, self.style_number_var.get()
            )
        )
        if not defaults["requires_decoration"]:
            self.decoration_type_var.set(BLANK_DECORATION_LABEL)
        elif is_blank_decoration(self.decoration_type_var.get()):
            self.decoration_type_var.set("")
        if not self.decoration_type_var.get().strip():
            inferred_decoration = infer_default_decoration(
                selected_category, self.product_name_var.get(), self.style_number_var.get()
            )
            if inferred_decoration:
                self.decoration_type_var.set(inferred_decoration)
        if selected_category == "Hats / Headwear" and defaults["requires_decoration"]:
            self.decoration_location_var.set(HAT)
            self.custom_decoration_location_var.set("")
        self.update_decoration_color_state()
        self.status_label.configure(text=f"Applied {selected_category} purchase defaults.")

    def on_never_outsource_changed(self):
        """Remember that this switch value was an intentional user choice."""
        self._never_outsource_touched = True

    def on_decoration_location_selected(self, choice: str):
        is_custom = clean_text(choice) == OTHER_CUSTOM
        if hasattr(self, "custom_location_entry"):
            self.custom_location_entry.configure(state="normal" if is_custom else "disabled")
        if hasattr(self, "custom_location_label"):
            self.custom_location_label.configure(text_color=PURPLE if is_custom else TEXT_MUTED)
        if not is_custom:
            self.custom_decoration_location_var.set("")

    def resolved_decoration_location(self, decoration_type: str, requires_decoration: bool) -> str:
        if is_blank_decoration(decoration_type) or is_in_house_decoration(decoration_type):
            return ""
        choice = self.decoration_location_var.get()
        location = resolve_location_choice(choice, self.custom_decoration_location_var.get())
        return normalize_decoration_location(
            location, decoration_type, requires_decoration=requires_decoration,
            product_name=self.product_name_var.get(), category=self.category_var.get(),
        )

    def on_requires_decoration_changed(self):
        if not self.requires_decoration_var.get() and not is_in_house_decoration(self.decoration_type_var.get()):
            self.decoration_type_var.set(BLANK_DECORATION_LABEL)
        elif self.requires_decoration_var.get() and is_blank_decoration(self.decoration_type_var.get()):
            self.decoration_type_var.set("")
        self.update_decoration_color_state()

    def on_product_name_changed(self, *_):
        if not self.category_var.get().strip():
            self.category_var.set(infer_category(self.product_name_var.get(), self.style_number_var.get()))
        inferred_service = infer_in_house_decoration(self.product_name_var.get())
        if inferred_service and not self.decoration_type_var.get().strip():
            self.decoration_type_var.set(inferred_service)
            self.requires_size_var.set(False)
            self.requires_color_var.set(False)
            self.requires_decoration_var.set(False)
        elif is_blank_garment_product(self.product_name_var.get(), style_number=self.style_number_var.get()) and not self.decoration_type_var.get().strip():
            self.decoration_type_var.set(BLANK_DECORATION_LABEL)

    def on_decoration_type_changed(self, *_):
        self.update_decoration_color_state()

    def update_decoration_color_state(self):
        decoration_type = self.decoration_type_var.get()
        in_house = is_in_house_decoration(decoration_type)
        none_selected = is_blank_decoration(decoration_type) or in_house or not self.requires_decoration_var.get()
        if in_house:
            self.requires_decoration_var.set(False)
        if hasattr(self, "decoration_location_combo"):
            if none_selected:
                display = NOT_APPLICABLE_IN_HOUSE if in_house else NOT_APPLICABLE_NO_DECORATION
                self.decoration_location_var.set(display)
                self.custom_decoration_location_var.set("")
                self.decoration_location_combo.configure(state="disabled")
                if hasattr(self, "custom_location_entry"):
                    self.custom_location_entry.configure(state="disabled")
                if hasattr(self, "placement_instructions_entry"):
                    self.placement_instructions_entry.configure(state="disabled")
                    self.decoration_placement_instructions_var.set("")
            else:
                current = self.decoration_location_var.get().strip()
                if not current or current in {NOT_APPLICABLE_IN_HOUSE, NOT_APPLICABLE_NO_DECORATION}:
                    self.decoration_location_var.set(default_decoration_location(
                        self.product_name_var.get(), self.category_var.get(), decoration_type
                    ))
                self.decoration_location_combo.configure(state="normal")
                if hasattr(self, "placement_instructions_entry"):
                    self.placement_instructions_entry.configure(state="normal")
                self.on_decoration_location_selected(self.decoration_location_var.get())
        for control in self.color_controls:
            combo = control.get("decoration_combo")
            if none_selected:
                control["decoration_var"].set("")
                if combo is not None:
                    combo.configure(state="disabled")
            elif combo is not None:
                combo.configure(state="normal")
        if hasattr(self, "apply_color_combo"):
            self.apply_color_combo.configure(state="disabled" if none_selected else "normal")

    def _update_color_setup_guidance(self, rows):
        """Show exactly which imported color needs attention, if any."""
        if not hasattr(self, "color_setup_summary_label"):
            return
        text, needs_attention = new_color_setup_message(rows)
        self.color_setup_summary_label.configure(
            text=text,
            text_color=WARNING if needs_attention else TEXT_MUTED,
        )

    def _sync_visible_color_controls_to_master_data(self) -> None:
        """Preserve unsaved color edits before the color table is rebuilt.

        Adding or deleting a garment-color row calls ``show_current_style``, which
        recreates every entry widget. Earlier builds rebuilt from the last saved
        DataFrame and could therefore erase colors the user had just typed.
        """
        for control in getattr(self, "color_controls", []):
            index = control.get("index")
            if index not in self.master_data.index:
                continue
            self.master_data.at[index, "Garment Color"] = normalize_space(control["garment_var"].get())
            self.master_data.at[index, "Vendor Color Code"] = normalize_space(control["vendor_code_var"].get())
            self.master_data.at[index, "Color Aliases"] = normalize_space(control["aliases_var"].get())
            self.master_data.at[index, "Purchasing Style Number"] = normalize_space(control["purchasing_style_var"].get())
            self.master_data.at[index, "Decoration Color"] = normalize_space(control["decoration_var"].get())

    def delete_color_row(self, data_index):
        self._sync_visible_color_controls_to_master_data()
        rows = self.get_style_rows(self.current_style_key) if self.current_style_key else pd.DataFrame()
        if len(rows) <= 1:
            if is_boots_product(
                self.product_name_var.get(), self.category_var.get(), self.style_number_var.get()
            ):
                # Boots must keep the underlying product row but never a shoe
                # size in Purchasing Color. Clear the mistaken value safely.
                self.master_data.at[data_index, "Garment Color"] = ""
                self.master_data.at[data_index, "Vendor Color Code"] = ""
                self.master_data.at[data_index, "Color Aliases"] = ""
                self.master_data.at[data_index, "Purchasing Style Number"] = ""
                self.master_data.at[data_index, "Decoration Color"] = ""
                saved = save_product_master(
                    self.master_data,
                    affected_style_keys=[self.current_style_key],
                )
                if saved is None:
                    return
                self.master_data = saved
                self._loaded_master_stamp = product_master_file_stamp()
                self.build_style_list()
                self.filtered_style_keys = self.get_matching_style_keys()
                if self.current_style_key in self.filtered_style_keys:
                    self.style_position = self.filtered_style_keys.index(self.current_style_key)
                self.show_current_style()
                self.refresh_progress()
                self.status_label.configure(text="Cleared the mistaken boot purchasing color.")
                return
            messagebox.showinfo("Color Required", "A product must keep at least one garment-color row.")
            return
        if data_index not in self.master_data.index:
            return
        color = clean_text(self.master_data.at[data_index, "Garment Color"]) or "blank color row"
        if not messagebox.askyesno("Delete Garment Color", f"Delete {color!r} from this product?"):
            return
        self.ensure_session_backup()
        current_key = self.current_style_key
        self.master_data = self.master_data.drop(index=data_index).reset_index(drop=True)
        saved = save_product_master(self.master_data, affected_style_keys=[current_key])
        if saved is None:
            return
        self.master_data = saved
        self.build_style_list()
        self.filtered_style_keys = self.get_matching_style_keys()
        if current_key in self.filtered_style_keys:
            self.style_position = self.filtered_style_keys.index(current_key)
        elif self.filtered_style_keys:
            self.style_position = min(self.style_position, len(self.filtered_style_keys) - 1)
        self.show_current_style()
        self.refresh_progress()
        self.status_label.configure(text=f"Deleted {color}.")

    def add_color_row(self):
        if not self.current_style_key:
            return
        # Keep every color currently typed into the form before rebuilding the
        # color list with the new blank row. This prevents earlier draft colors
        # from disappearing after repeated Add Garment Color clicks.
        self._sync_visible_color_controls_to_master_data()
        new_row = {
            "Product Name": self.product_name_var.get().strip(),
            "Style Number": self.style_number_var.get().strip(),
            "Garment Color": "",
            "Vendor": self.vendor_var.get().strip(),
            "Decoration Type": self.decoration_type_var.get().strip(),
            "Decoration Location": self.resolved_decoration_location(
                self.decoration_type_var.get(), self.requires_decoration_var.get()
            ),
            "Decoration Placement Instructions": normalize_space(self.decoration_placement_instructions_var.get()),
            "Decoration Color": "",
            "Product ID": self.product_id_var.get().strip() or default_product_id(self.style_number_var.get(), self.product_name_var.get()),
            "Product Aliases": self.product_aliases_var.get().strip(),
            "Vendor Color Code": "",
            "Color Aliases": "",
            "Purchasing Style Number": "",
            "Product Category": self.category_var.get().strip() or infer_category(self.product_name_var.get(), self.style_number_var.get()),
            "Requires Size": bool_text(self.requires_size_var.get()),
            "Requires Color": bool_text(self.requires_color_var.get()),
            "Requires Decoration": bool_text(self.requires_decoration_var.get()),
            NEVER_OUTSOURCE_COLUMN: bool_text(self.never_outsource_var.get()),
            "Setup Required": "Yes",
        }
        self.master_data = pd.concat([self.master_data, pd.DataFrame([new_row])], ignore_index=True)
        self.build_style_list()
        new_key = style_group_key(pd.Series(new_row))
        if new_key not in self.filtered_style_keys:
            self.filtered_style_keys.append(new_key)
        self.current_style_key = new_key
        self.style_position = self.filtered_style_keys.index(new_key)
        self.show_current_style()
        def focus_new_row():
            self._scroll_colors_to_bottom()
            blank_controls = [control for control in self.color_controls if control["index"] == self.master_data.index[-1]]
            if blank_controls:
                blank_controls[0]["garment_entry"].focus_set()
        self.after(120, focus_new_row)
        self.status_label.configure(text="Blank garment-color row added. Complete it before saving.")

    def apply_color_to_all(self):
        value = self.apply_color_var.get().strip()
        for control in self.color_controls:
            control["decoration_var"].set(value)
        self.status_label.configure(text=f"Applied {value or 'blank'} to all garment colors.")

    def copy_first_color_to_all(self):
        value = next((control["decoration_var"].get().strip() for control in self.color_controls if control["decoration_var"].get().strip()), "")
        if not value:
            messagebox.showinfo("No Completed Color", "Enter one thread or ink color first, then use this button.")
            return
        self.apply_color_var.set(value)
        self.apply_color_to_all()

    def ensure_session_backup(self):
        if self.session_backup_path or not PRODUCT_MASTER_PATH.exists():
            return True
        try:
            backup_dir = PRODUCT_MASTER_PATH.parent / "backups"
            backup_dir.mkdir(parents=True, exist_ok=True)
            backup_path = backup_dir / f"product_master_before_pro_session_{datetime.now():%Y%m%d_%H%M%S}.csv"
            shutil.copy2(PRODUCT_MASTER_PATH, backup_path)
            self.session_backup_path = backup_path
            return True
        except Exception as error:
            # A cloud-synced Documents folder can temporarily refuse the optional
            # session backup.  Do not silently abort the user's actual save.
            self._write_save_diagnostic("session backup", error)
            self.status_label.configure(text="Backup was unavailable; continuing with the Product Master save.")
            return False

    def _write_save_diagnostic(self, stage: str, error: Exception) -> Path | None:
        """Write a useful diagnostic without allowing logging itself to block saving."""
        try:
            log_path = PRODUCT_MASTER_PATH.parent / "product_master_save_error.log"
            with log_path.open("a", encoding="utf-8") as handle:
                handle.write(f"\n[{datetime.now().isoformat(timespec='seconds')}] {stage}\n")
                handle.write("".join(traceback.format_exception(type(error), error, error.__traceback__)))
            return log_path
        except Exception:
            return None

    def _mark_catalog_updated(self, style_name: str = ""):
        try:
            payload = {
                "saved_at": datetime.now().isoformat(timespec="seconds"),
                "style": style_name,
                "complete": self.completed_style_count(),
                "total": len(self.all_style_keys),
            }
            PRODUCT_MASTER_UPDATE_MARKER.parent.mkdir(parents=True, exist_ok=True)
            temporary = PRODUCT_MASTER_UPDATE_MARKER.with_suffix(".tmp")
            temporary.write_text(json.dumps(payload, indent=2), encoding="utf-8")
            temporary.replace(PRODUCT_MASTER_UPDATE_MARKER)
        except Exception:
            pass

    def _show_save_confirmation(self, message: str, complete: bool):
        if not hasattr(self, "save_toast"):
            self.save_toast = ctk.CTkLabel(
                self.editor_frame, text="", height=38, corner_radius=19,
                fg_color="#EAF8EF", text_color=SUCCESS,
                font=ctk.CTkFont(size=13, weight="bold"),
            )
        if self._save_toast_after_id:
            try:
                self.after_cancel(self._save_toast_after_id)
            except Exception:
                pass
        self.save_toast.configure(
            text=message,
            fg_color="#EAF8EF" if complete else "#FFF1DC",
            text_color=SUCCESS if complete else WARNING,
        )
        self.save_toast.place(relx=0.5, y=112, anchor="n")
        self.save_toast.lift()
        self._save_toast_after_id = self.after(1900, self.save_toast.place_forget)

    def save_current_style(self):
        """Save from a button click and always give visible success or failure feedback."""
        try:
            return self._save_current_style_impl()
        except Exception as error:
            log_path = self._write_save_diagnostic("Product Master save", error)
            details = f"\n\nDiagnostic saved to:\n{log_path}" if log_path else ""
            try:
                self.status_label.configure(text="Save failed — Orchid displayed the error so it can be corrected.")
            except Exception:
                pass
            messagebox.showerror(
                "Unable to Save Product",
                f"Orchid could not save this Product Master record.\n\n{error}{details}",
                parent=self,
            )
            return False

    def _save_current_style_impl(self):
        if not self.current_style_key:
            messagebox.showwarning(
                "No Product Selected",
                "Select a Product Master record before saving.",
                parent=self,
            )
            return False
        original_key = self.current_style_key
        rows = self.get_style_rows(original_key)
        if rows.empty:
            messagebox.showwarning(
                "Product Record Unavailable",
                "This Product Master record is no longer available. Return to Product Catalog and open it again.",
                parent=self,
            )
            return False
        self.ensure_session_backup()

        product_name = canonical_product_name(self.product_name_var.get())
        style_number = normalize_style(self.style_number_var.get())
        vendor = canonical_vendor_name(self.vendor_var.get())
        vendor = next(
            (existing for existing in self.vendor_options if existing.casefold() == vendor.casefold()),
            vendor,
        )
        product_id = self.product_id_var.get().strip() or default_product_id(style_number, product_name)
        product_aliases = normalize_space(self.product_aliases_var.get())
        category = self.category_var.get().strip() or infer_category(product_name, style_number)
        requires_size = self.requires_size_var.get()
        requires_color = self.requires_color_var.get()
        requires_decoration = self.requires_decoration_var.get()
        if is_boots_product(product_name, category, style_number):
            # Product Master always saves Boots in their permanent purchasing
            # configuration, even if an older record or a click changed a
            # switch before Save.
            requires_size = True
            requires_color = False
            requires_decoration = False
            self.requires_size_var.set(True)
            self.requires_color_var.set(False)
            self.requires_decoration_var.set(False)
            self.never_outsource_var.set(True)
        decoration_type = self.decoration_type_var.get().strip()
        if is_in_house_decoration(decoration_type):
            requires_decoration = False
            self.requires_decoration_var.set(False)
        elif not requires_decoration:
            decoration_type = BLANK_DECORATION_LABEL
            self.decoration_type_var.set(BLANK_DECORATION_LABEL)
        if is_blank_garment_product(product_name, style_number=style_number) and not decoration_type:
            decoration_type = BLANK_DECORATION_LABEL
            self.decoration_type_var.set(BLANK_DECORATION_LABEL)
        if requires_decoration and not is_blank_decoration(decoration_type) and not is_in_house_decoration(decoration_type) and self.decoration_location_var.get() == OTHER_CUSTOM and not normalize_space(self.custom_decoration_location_var.get()):
            messagebox.showwarning(
                "Custom Location Required",
                "Enter the custom decoration location before saving this product.",
            )
            self.custom_location_entry.focus_set()
            return False
        decoration_location = self.resolved_decoration_location(decoration_type, requires_decoration)
        placement_instructions = normalize_space(self.decoration_placement_instructions_var.get())
        if not requires_decoration or is_blank_decoration(decoration_type) or is_in_house_decoration(decoration_type):
            placement_instructions = ""
            self.decoration_placement_instructions_var.set("")
        choice, custom = location_choice_and_custom(decoration_location, decoration_type)
        self.decoration_location_var.set(choice)
        self.custom_decoration_location_var.set(custom)

        # A repeat import may add one new color to a fully configured style.
        # Preserve the 33 saved color decisions and require a choice only for
        # the newly added color before clearing its Setup Required marker.
        pending_thread_controls = []
        if requires_decoration and not is_blank_decoration(decoration_type) and not is_in_house_decoration(decoration_type):
            for control in self.color_controls:
                index = control["index"]
                if index not in self.master_data.index:
                    continue
                source_row = self.master_data.loc[index]
                purchasing_color = normalize_space(control["garment_var"].get())
                thread_ink = normalize_space(control["decoration_var"].get())
                if is_new_color_setup_row(source_row) and purchasing_color and not thread_ink:
                    pending_thread_controls.append(control)
        if pending_thread_controls:
            colors = ", ".join(
                normalize_space(control["garment_var"].get()) for control in pending_thread_controls[:3]
            )
            if len(pending_thread_controls) > 3:
                colors += f" +{len(pending_thread_controls) - 3} more"
            messagebox.showwarning(
                "Thread / Ink Needed for New Color",
                "Choose the Thread / Ink color only for the new purchasing color before saving:\n\n"
                + colors,
                parent=self,
            )
            pending_thread_controls[0]["decoration_combo"].focus_set()
            return False

        for index in rows.index:
            self.master_data.at[index, "Product Name"] = product_name
            self.master_data.at[index, "Style Number"] = style_number
            self.master_data.at[index, "Vendor"] = vendor
            self.master_data.at[index, "Product ID"] = product_id
            self.master_data.at[index, "Product Aliases"] = product_aliases
            self.master_data.at[index, "Product Category"] = category
            self.master_data.at[index, "Requires Size"] = bool_text(requires_size)
            self.master_data.at[index, "Requires Color"] = bool_text(requires_color)
            self.master_data.at[index, "Requires Decoration"] = bool_text(requires_decoration)
            self.master_data.at[index, NEVER_OUTSOURCE_COLUMN] = bool_text(self.never_outsource_var.get())
            if "Setup Required" in self.master_data.columns:
                self.master_data.at[index, "Setup Required"] = "No"
            self.master_data.at[index, "Decoration Type"] = normalize_decoration_type(decoration_type)
            self.master_data.at[index, "Decoration Location"] = normalize_decoration_location(
                decoration_location, decoration_type, requires_decoration=requires_decoration
            )
            self.master_data.at[index, "Decoration Placement Instructions"] = placement_instructions
        for control in self.color_controls:
            index = control["index"]
            if index in self.master_data.index:
                previous_color = normalize_space(self.master_data.at[index, "Garment Color"])
                purchasing_color = normalize_space(control["garment_var"].get())
                color_aliases = preserve_previous_color_alias(
                    previous_color, purchasing_color, control["aliases_var"].get()
                )
                self.master_data.at[index, "Garment Color"] = purchasing_color
                self.master_data.at[index, "Vendor Color Code"] = normalize_space(control["vendor_code_var"].get())
                self.master_data.at[index, "Color Aliases"] = color_aliases
                control["aliases_var"].set(color_aliases)
                self.master_data.at[index, "Purchasing Style Number"] = normalize_space(control["purchasing_style_var"].get())
                self.master_data.at[index, "Decoration Color"] = "" if (is_blank_decoration(decoration_type) or is_in_house_decoration(decoration_type)) else control["decoration_var"].get().strip()

        used_thread_colors = [
            control["decoration_var"].get().strip() for control in self.color_controls
            if control["decoration_var"].get().strip()
        ]
        if used_thread_colors:
            self.decoration_color_options = save_thread_ink_colors(used_thread_colors, PRODUCT_MASTER_PATH)
            self._refresh_thread_ink_dropdowns()

        new_key = f"style:{style_number}" if style_number else f"product:{product_name.casefold()}"
        saved = save_product_master(
            self.master_data,
            expected_style=style_number,
            expected_vendor=vendor,
            affected_style_keys=[original_key, new_key],
        )
        if saved is None:
            return False
        self.master_data = saved
        self._loaded_master_stamp = product_master_file_stamp()
        # A typed vendor is just as permanent as one added with the button.
        # Register it after the Product Master write has succeeded so a failed
        # product save cannot leave a misleading vendor option behind.
        if vendor:
            register_purchase_vendor(vendor)
            self.vendor_options = purchase_vendor_options(self.master_data)
        if hasattr(self, "vendor_combo"):
            self.vendor_combo.configure(values=self.vendor_options)
        if hasattr(self, "dashboard_vendor_filter"):
            self.dashboard_vendor_filter.configure(values=["All Vendors"] + self.vendor_options + ["Unassigned"])
        self.build_style_list()
        self.current_style_key = new_key if new_key in self.all_style_keys else None
        if self._never_outsource_touched:
            self.never_outsource_overrides[new_key] = bool(self.never_outsource_var.get())
            if original_key != new_key:
                self.never_outsource_overrides.pop(original_key, None)
            save_never_outsource_overrides(self.never_outsource_overrides)
            self._never_outsource_touched = False
        self.session_saved_styles.add(style_number or product_name or self.current_style_key or "Unknown")
        self.refresh_progress()

        current_rows = self.get_style_rows(new_key) if new_key in self.all_style_keys else pd.DataFrame(columns=COLUMNS)
        issues = self.style_issues(current_rows)
        saved_name = style_number or product_name or "Product"
        self._mark_catalog_updated(saved_name)
        if issues:
            message = "Saved — additional setup available: " + ", ".join(issues)
            self.status_label.configure(text=message)
            self._show_save_confirmation(f"Saved {saved_name} — setup still needed", complete=False)
        else:
            self.status_label.configure(text=f"Saved ✓ Verified in live Product Master: {PRODUCT_MASTER_PATH}")
            self._show_save_confirmation(f"✓ {saved_name} saved", complete=True)
        # Remove the temporary NEW marker immediately after a successful save.
        # Save & Next may select another item synchronously; resolving the
        # current key on the idle cycle keeps both actions safe.
        self.after_idle(self.show_current_style)
        return True

    def save_and_next_style(self):
        current_key = self.current_style_key
        try:
            global_position = self.all_style_keys.index(current_key)
        except (ValueError, TypeError):
            global_position = -1
        if not self.save_current_style():
            return
        saved_key = self.current_style_key
        saved_rows = self.get_style_rows(saved_key) if saved_key else pd.DataFrame(columns=COLUMNS)
        remaining_issues = self.style_issues(saved_rows)
        if remaining_issues:
            messagebox.showwarning(
                "Setup Still Required",
                "This product was saved, but it still needs the following before Orchid can move to the next setup item:\n\n• "
                + "\n• ".join(remaining_issues),
                parent=self,
            )
            self.status_label.configure(text="Saved — complete the listed setup before moving to the next item.")
            return
        self.filtered_style_keys = self.get_matching_style_keys()
        if not self.filtered_style_keys:
            self.show_dashboard()
            messagebox.showinfo("Setup View Complete", "There are no additional setup items remaining in the current view.")
            return
        next_key = next((key for key in self.all_style_keys[global_position + 1 :] if key in self.filtered_style_keys), None)
        if next_key is None:
            next_key = self.filtered_style_keys[0]
        self.style_position = self.filtered_style_keys.index(next_key)
        self.show_current_style(defer=True)
        self.after(120, self._scroll_to_top)
        self.refresh_progress()

    def previous_style(self):
        if self.filtered_style_keys and self.style_position > 0:
            self.style_position -= 1
            self.show_current_style(defer=True)

    def next_style(self):
        if self.filtered_style_keys and self.style_position < len(self.filtered_style_keys) - 1:
            self.style_position += 1
            self.show_current_style(defer=True)

    def on_close(self):
        try:
            if PRODUCT_MASTER_PID_FILE.exists() and PRODUCT_MASTER_PID_FILE.read_text(encoding="utf-8").strip() == str(os.getpid()):
                PRODUCT_MASTER_PID_FILE.unlink(missing_ok=True)
            PRODUCT_MASTER_REQUEST_FILE.unlink(missing_ok=True)
        except Exception:
            pass
        if self.session_saved_styles:
            self.build_style_list()
            complete_now = self.completed_style_count()
            gained = max(0, complete_now - self.session_start_complete)
            messagebox.showinfo(
                "Product Master Session",
                f"Styles saved this session: {len(self.session_saved_styles)}\n"
                f"Newly completed styles: {gained}\n"
                f"Product Master health: {complete_now} of {len(self.all_style_keys)} complete\n\n"
                f"Active file:\n{PRODUCT_MASTER_PATH}",
            )
        self.destroy()


if __name__ == "__main__":
    import sys
    prefill = ""
    search = ""
    setup = ""
    if "--new-product" in sys.argv:
        try:
            prefill = sys.argv[sys.argv.index("--new-product") + 1]
        except (ValueError, IndexError):
            prefill = ""
    if "--search-product" in sys.argv:
        try:
            search = sys.argv[sys.argv.index("--search-product") + 1]
        except (ValueError, IndexError):
            search = ""
    if "--setup-product" in sys.argv:
        try:
            setup = sys.argv[sys.argv.index("--setup-product") + 1]
        except (ValueError, IndexError):
            setup = ""
    ProductMasterV2(prefill_product=prefill, search_product=search, setup_product=setup).mainloop()
