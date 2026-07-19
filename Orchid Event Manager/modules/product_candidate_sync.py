from __future__ import annotations

import re
import shutil
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd

from modules.blank_garment_rules import BLANK_DECORATION_LABEL, is_blank_decoration
from modules.product_intelligence import infer_default_decoration
from modules.decoration_locations import default_decoration_location, normalize_decoration_location
from modules.purchase_rules import bool_text, category_defaults, infer_category
from modules.routing_rules import preferred_vendor_for_brand_text
from modules.xlsx_reader import read_table

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
    "Product Category",
    "Requires Size",
    "Requires Color",
    "Requires Decoration",
    "Setup Required",
]


def clean(value: Any) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return re.sub(r"\s+", " ", str(value).strip())


def normalize_style(value: Any) -> str:
    return re.sub(r"\s+", "", clean(value)).upper()


def normalize_color(value: Any) -> str:
    return clean(value).casefold()


def _split_colors(value: Any) -> list[str]:
    text = clean(value)
    if not text:
        return [""]
    output: list[str] = []
    for part in re.split(r"\s*,\s*|\s*;\s*", text):
        color = clean(part)
        if color and color.casefold() not in {item.casefold() for item in output}:
            output.append(color)
    return output or [""]


def _product_id(style: str, name: str) -> str:
    return style or re.sub(r"[^A-Z0-9]+", "-", clean(name).upper()).strip("-")[:48]


def _read_master(path: Path) -> pd.DataFrame:
    if path.exists():
        frame = pd.read_csv(path, dtype=str).fillna("")
    else:
        frame = pd.DataFrame(columns=COLUMNS)
    for column in COLUMNS:
        if column not in frame.columns:
            frame[column] = ""
    return frame[COLUMNS].copy()


def _looks_like_style_code(value: Any) -> bool:
    text = clean(value)
    if not text or len(text) > 30 or " " in text:
        return False
    return bool(re.search(r"[A-Za-z]", text) and re.search(r"[0-9]", text)) or text.isdigit()


def _candidate_rows(review_path: Path) -> list[dict[str, str]]:
    """Read new-style candidates from both current and legacy review layouts.

    Professional 4.3.9 could place an unknown style only on Review & Edit when the
    same order also had an event-specific problem such as a missing size. Reading
    that legacy surface lets 4.4 repair an already active packet immediately.
    """
    rows: list[dict[str, str]] = []
    for sheet in ("Product Master", "Product Master Needed", "Purchasing Rules Needed"):
        try:
            rows.extend(read_table(review_path, sheet, {"Product #", "Missing Information"}))
            break
        except Exception:
            continue

    try:
        legacy_rows = read_table(review_path, "Review & Edit", {"Product #", "Review Reason"})
    except Exception:
        legacy_rows = []
    try:
        source_rows = read_table(review_path, "All PO Lines", {"Line ID", "Master Match"})
    except Exception:
        source_rows = []
    source_by_line = {clean(row.get("Line ID", "")): row for row in source_rows if clean(row.get("Line ID", ""))}

    for record in legacy_rows:
        style = normalize_style(record.get("Product #", ""))
        reason = clean(record.get("Review Reason", "")).casefold()
        resolution = clean(record.get("Resolution", "")).casefold()
        include = clean(record.get("Include", "Yes")).casefold()
        source = source_by_line.get(clean(record.get("Line ID", "")), {})
        source_master_match = clean(source.get("Master Match", "")).casefold()
        corrected_unknown_style = bool(
            _looks_like_style_code(style)
            and source_master_match == "no"
            and clean(source.get("Source", "")).casefold() != "smart custom item"
        )
        if include in {"no", "n", "false", "0", "exclude"}:
            continue
        if "one-time manual item" in resolution:
            continue
        if not _looks_like_style_code(style):
            continue
        if "not found in product master" not in reason and not corrected_unknown_style:
            continue
        rows.append({
            "Product #": style,
            "Description": clean(record.get("Description", "")),
            "Garment Color(s)": clean(record.get("Garment Color", "")),
            "Current Vendor": clean(record.get("Purchase Vendor", "")),
            "Missing Information": clean(record.get("Review Reason", "")) or "Product not found in Product Master",
            "Source": "Legacy Purchase Review",
        })

    deduped: list[dict[str, str]] = []
    seen: set[tuple[str, str]] = set()
    for record in rows:
        key = (normalize_style(record.get("Product #", "")), normalize_color(record.get("Garment Color(s)", record.get("Garment Color", ""))))
        if key in seen:
            continue
        seen.add(key)
        deduped.append(record)
    return deduped


def sync_review_product_candidates(review_path: Path, master_path: Path) -> dict[str, Any]:
    """Materialize review candidates into Product Master as safe incomplete records.

    Purchase Review previously listed genuinely new styles without adding them to the
    editable Product Master. This function adds only those deduplicated candidates,
    preserving every existing manual value and adding known garment colors as separate
    rows. The Product Master editor can then open the exact style and use its normal
    Save & Next setup workflow.
    """
    review_path = Path(review_path)
    master_path = Path(master_path)
    if not review_path.exists():
        return {"changed": False, "added_styles": 0, "added_colors": 0, "updated_fields": 0}

    candidates = _candidate_rows(review_path)
    if not candidates:
        return {"changed": False, "added_styles": 0, "added_colors": 0, "updated_fields": 0}

    frame = _read_master(master_path)
    changed = False
    added_styles = 0
    added_colors = 0
    updated_fields = 0

    for candidate in candidates:
        style = normalize_style(candidate.get("Product #", ""))
        description = clean(candidate.get("Description", ""))
        if not style and not description:
            continue
        product_name = description or style
        colors = _split_colors(candidate.get("Garment Color(s)", candidate.get("Garment Color", "")))
        current_vendor = clean(candidate.get("Current Vendor", candidate.get("Suggested Vendor", "")))
        _, inferred_vendor = preferred_vendor_for_brand_text(f"{product_name} {style}")
        vendor = current_vendor or inferred_vendor
        category = infer_category(product_name, style)
        defaults = category_defaults(category, product_name, style)
        decoration = infer_default_decoration(category, product_name, style)
        requires_decoration = defaults["requires_decoration"]
        if not requires_decoration:
            decoration = BLANK_DECORATION_LABEL

        if style:
            style_mask = frame["Style Number"].map(normalize_style).eq(style)
        else:
            style_mask = frame["Product Name"].map(lambda value: clean(value).casefold()).eq(product_name.casefold())
        style_exists = bool(style_mask.any())
        if not style_exists:
            added_styles += 1

        template = None
        if style_exists:
            template = frame.loc[style_mask].iloc[0].to_dict()

        for color in colors:
            if style_exists:
                color_mask = style_mask & frame["Garment Color"].map(normalize_color).eq(normalize_color(color))
            else:
                color_mask = pd.Series(False, index=frame.index)
            if color_mask.any():
                indexes = frame.index[color_mask]
                for index in indexes:
                    fill_values = {
                        "Product Name": product_name,
                        "Style Number": style,
                        "Vendor": vendor,
                        "Product Category": category,
                        "Decoration Type": decoration,
                        "Decoration Location": normalize_decoration_location(
                            default_decoration_location(product_name, category, decoration),
                            decoration, requires_decoration=requires_decoration,
                            product_name=product_name, category=category,
                        ),
                        "Requires Size": bool_text(defaults["requires_size"]),
                        "Requires Color": bool_text(defaults["requires_color"]),
                        "Requires Decoration": bool_text(requires_decoration),
                        "Setup Required": "Yes",
                    }
                    for field, value in fill_values.items():
                        if value and not clean(frame.at[index, field]):
                            frame.at[index, field] = value
                            updated_fields += 1
                            changed = True
                continue

            row = {column: "" for column in COLUMNS}
            if template:
                row.update({column: clean(template.get(column, "")) for column in COLUMNS})
            row.update({
                "Product Name": clean(row.get("Product Name")) or product_name,
                "Style Number": clean(row.get("Style Number")) or style,
                "Garment Color": color,
                "Vendor": clean(row.get("Vendor")) or vendor,
                "Decoration Type": clean(row.get("Decoration Type")) or decoration,
                "Decoration Location": normalize_decoration_location(
                    clean(row.get("Decoration Location")) or default_decoration_location(product_name, category, decoration),
                    clean(row.get("Decoration Type")) or decoration,
                    requires_decoration=requires_decoration,
                    product_name=product_name, category=category,
                ),
                "Decoration Color": "" if is_blank_decoration(decoration) else clean(row.get("Decoration Color")),
                "Product ID": clean(row.get("Product ID")) or _product_id(style, product_name),
                "Product Category": clean(row.get("Product Category")) or category,
                "Requires Size": clean(row.get("Requires Size")) or bool_text(defaults["requires_size"]),
                "Requires Color": clean(row.get("Requires Color")) or bool_text(defaults["requires_color"]),
                "Requires Decoration": clean(row.get("Requires Decoration")) or bool_text(requires_decoration),
                "Setup Required": "Yes",
            })
            frame = pd.concat([frame, pd.DataFrame([row], columns=COLUMNS)], ignore_index=True)
            style_exists = True
            if style:
                style_mask = frame["Style Number"].map(normalize_style).eq(style)
            else:
                style_mask = frame["Product Name"].map(lambda value: clean(value).casefold()).eq(product_name.casefold())
            added_colors += 1
            changed = True

    if changed:
        master_path.parent.mkdir(parents=True, exist_ok=True)
        if master_path.exists():
            backup_dir = master_path.parent / "backups"
            backup_dir.mkdir(parents=True, exist_ok=True)
            backup_path = backup_dir / f"product_master_before_candidate_sync_{datetime.now():%Y%m%d_%H%M%S}.csv"
            shutil.copy2(master_path, backup_path)
        frame.to_csv(master_path, index=False)

    return {
        "changed": changed,
        "added_styles": added_styles,
        "added_colors": added_colors,
        "updated_fields": updated_fields,
    }
