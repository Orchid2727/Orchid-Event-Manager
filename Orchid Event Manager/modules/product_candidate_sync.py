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
from modules.outsource_rules import NEVER_OUTSOURCE_COLUMN
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
    NEVER_OUTSOURCE_COLUMN,
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
    extra_columns = [column for column in frame.columns if column not in COLUMNS]
    return frame[COLUMNS + extra_columns].copy()


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



def _normalized_catalog_text(value: Any) -> str:
    text = clean(value).casefold().replace("®", " ").replace("™", " ")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _catalog_description_candidate(value: Any, style: str) -> str:
    description = clean(value)
    normalized = _normalized_catalog_text(description)
    if not normalized or normalized == _normalized_catalog_text(style):
        return ""
    if len(normalized) < 4 or not re.search(r"[a-z]", normalized):
        return ""
    return description


def _best_catalog_description(group: pd.DataFrame, style: str) -> str:
    candidates = [
        _catalog_description_candidate(value, style)
        for value in group.get("Product Name", pd.Series(dtype=str))
    ]
    candidates = [value for value in candidates if value]
    if not candidates:
        return ""
    counts: dict[str, int] = {}
    display: dict[str, str] = {}
    for value in candidates:
        key = _normalized_catalog_text(value)
        counts[key] = counts.get(key, 0) + 1
        if key not in display or len(value) > len(display[key]):
            display[key] = value
    best_key = max(counts, key=lambda key: (counts[key], len(display[key])))
    return display[best_key]


def _color_terms(row: pd.Series) -> set[str]:
    values = [row.get("Garment Color", ""), row.get("Vendor Color Code", "")]
    values.extend(re.split(r"[|;,\n]+", clean(row.get("Color Aliases", ""))))
    return {_normalized_catalog_text(value).replace("grey", "gray") for value in values if clean(value)}


_SETUP_TRUE = {"yes", "y", "true", "1"}
_COLOR_SPECIFIC_FIELDS = {
    "Garment Color",
    "Vendor Color Code",
    "Color Aliases",
    "Decoration Color",
    "Setup Required",
}


def _is_setup_required(row: pd.Series) -> bool:
    return clean(row.get("Setup Required", "")).casefold() in _SETUP_TRUE


def _matching_color_indexes(rows: pd.DataFrame, color: Any) -> list[Any]:
    key = _normalized_catalog_text(color).replace("grey", "gray")
    if not key or rows.empty:
        return []
    return [index for index, row in rows.iterrows() if key in _color_terms(row)]


def _comparison_value(column: str, value: Any) -> str:
    if column == "Style Number":
        return normalize_style(value)
    if column in {"Garment Color", "Vendor Color Code", "Color Aliases"}:
        return _normalized_catalog_text(value).replace("grey", "gray")
    return clean(value).casefold()


def _merge_color_aliases(existing: Any, *values: Any) -> str:
    output: list[str] = []
    seen: set[str] = set()
    for raw in [existing, *values]:
        for piece in re.split(r"[|;,\n]+", clean(raw)):
            value = clean(piece)
            key = _normalized_catalog_text(value).replace("grey", "gray")
            if value and key and key not in seen:
                output.append(value)
                seen.add(key)
    return " | ".join(output)


def _safe_redundant_setup_row(candidate: pd.Series, keeper: pd.Series, columns: list[str]) -> bool:
    """Return True only when a setup row is clearly a duplicate of a completed row.

    The candidate garment color must already be covered by the completed row's
    purchasing color, vendor color code, or Shopify color aliases. Conflicting
    business settings prevent automatic removal so uncertain records remain visible.
    """
    candidate_color = clean(candidate.get("Garment Color", ""))
    candidate_key = _normalized_catalog_text(candidate_color).replace("grey", "gray")
    if not candidate_key or candidate_key not in _color_terms(keeper):
        return False

    candidate_vendor_code = clean(candidate.get("Vendor Color Code", ""))
    if candidate_vendor_code:
        vendor_key = _normalized_catalog_text(candidate_vendor_code).replace("grey", "gray")
        if vendor_key not in _color_terms(keeper):
            return False

    candidate_decoration = clean(candidate.get("Decoration Color", ""))
    keeper_decoration = clean(keeper.get("Decoration Color", ""))
    if candidate_decoration and keeper_decoration and candidate_decoration.casefold() != keeper_decoration.casefold():
        return False

    for column in columns:
        if column in _COLOR_SPECIFIC_FIELDS:
            continue
        candidate_value = clean(candidate.get(column, ""))
        keeper_value = clean(keeper.get(column, ""))
        if candidate_value and keeper_value and _comparison_value(column, candidate_value) != _comparison_value(column, keeper_value):
            return False
    return True


def _collapse_redundant_setup_color_rows(frame: pd.DataFrame) -> tuple[pd.DataFrame, int, list[dict[str, str]]]:
    """Remove only setup-required color rows already covered by a completed row.

    Professional 4.9.8 RC9 correctly retained imported Shopify colors as aliases,
    but a later candidate-sync pass compared only Garment Color and could recreate
    the imported color as a second setup row. This repair is intentionally narrow:
    it never merges two completed rows and never removes a row with conflicting
    business settings.
    """
    if frame.empty or "Style Number" not in frame.columns:
        return frame, 0, []

    working = frame.copy()
    remove_indexes: list[Any] = []
    details: list[dict[str, str]] = []
    style_keys = working["Style Number"].map(normalize_style)
    for style in [value for value in dict.fromkeys(style_keys.tolist()) if value]:
        group = working.loc[style_keys.eq(style)]
        completed = group.loc[~group.apply(_is_setup_required, axis=1)]
        pending = group.loc[group.apply(_is_setup_required, axis=1)]
        if completed.empty or pending.empty:
            continue
        for candidate_index, candidate in pending.iterrows():
            color = clean(candidate.get("Garment Color", ""))
            keeper_indexes = _matching_color_indexes(completed, color)
            for keeper_index in keeper_indexes:
                keeper = working.loc[keeper_index]
                if not _safe_redundant_setup_row(candidate, keeper, list(working.columns)):
                    continue

                # Preserve any harmless information before removing the redundant row.
                aliases = _merge_color_aliases(
                    keeper.get("Color Aliases", ""),
                    candidate.get("Garment Color", ""),
                    candidate.get("Color Aliases", ""),
                )
                working.at[keeper_index, "Color Aliases"] = aliases
                if not clean(working.at[keeper_index, "Decoration Color"]) and clean(candidate.get("Decoration Color", "")):
                    working.at[keeper_index, "Decoration Color"] = clean(candidate.get("Decoration Color", ""))
                for column in working.columns:
                    if column in _COLOR_SPECIFIC_FIELDS:
                        continue
                    if not clean(working.at[keeper_index, column]) and clean(candidate.get(column, "")):
                        working.at[keeper_index, column] = clean(candidate.get(column, ""))

                remove_indexes.append(candidate_index)
                details.append({
                    "style": style,
                    "removed_color": color,
                    "kept_color": clean(keeper.get("Garment Color", "")),
                })
                break

    if remove_indexes:
        working = working.drop(index=list(dict.fromkeys(remove_indexes))).reset_index(drop=True)
    return working, len(set(remove_indexes)), details


def _write_master_atomic(frame: pd.DataFrame, master_path: Path, backup_label: str) -> None:
    master_path.parent.mkdir(parents=True, exist_ok=True)
    if master_path.exists():
        backup_dir = master_path.parent / "backups"
        backup_dir.mkdir(parents=True, exist_ok=True)
        backup_path = backup_dir / f"product_master_before_{backup_label}_{datetime.now():%Y%m%d_%H%M%S}.csv"
        shutil.copy2(master_path, backup_path)
    temporary = master_path.with_suffix(master_path.suffix + ".tmp")
    frame.fillna("").to_csv(temporary, index=False)
    temporary.replace(master_path)


def sync_shopify_catalog_enrichment(parsed_orders: pd.DataFrame, master_path: Path) -> dict[str, Any]:
    """Enrich known styles from Shopify while preserving Orchid-owned setup.

    The style number is the merge key. Shopify may refresh the catalog description
    and introduce new garment colors. Vendor, category, decoration settings,
    purchasing colors, aliases, placement instructions, requirement switches, and
    Never Outsource values are copied from the existing style and never overwritten.
    """
    master_path = Path(master_path)
    result: dict[str, Any] = {
        "changed": False,
        "enriched_styles": 0,
        "updated_descriptions": 0,
        "added_colors": 0,
        "repaired_duplicate_colors": 0,
        "details": [],
        "repair_details": [],
    }
    if parsed_orders is None or parsed_orders.empty or not master_path.exists():
        return result

    frame = _read_master(master_path)
    if frame.empty or "Style Number" not in frame.columns:
        return result

    frame, repaired_count, repair_details = _collapse_redundant_setup_color_rows(frame)
    result["repaired_duplicate_colors"] = repaired_count
    result["repair_details"] = repair_details

    source = parsed_orders.copy()
    for column in ("Style Number", "Product Name", "Garment Color", "Parser Source"):
        if column not in source.columns:
            source[column] = ""
        source[column] = source[column].fillna("").map(clean)
    # Manually typed custom order lines are not authoritative catalog data. They
    # remain usable for the current order but cannot replace a permanent Product
    # Master description or create permanent color rows.
    source = source[~source["Parser Source"].str.casefold().eq("smart custom item")].copy()
    source["_style_key"] = source["Style Number"].map(normalize_style)
    source = source[source["_style_key"].ne("")].copy()

    changed = repaired_count > 0
    details: list[dict[str, Any]] = []
    if source.empty:
        if changed:
            _write_master_atomic(frame, master_path, "repeat_import_repair")
        result["changed"] = changed
        return result

    for style, group in source.groupby("_style_key", sort=False):
        style_mask = frame["Style Number"].map(normalize_style).eq(style)
        if not style_mask.any():
            continue
        style_rows = frame.loc[style_mask]
        description = _best_catalog_description(group, style)
        description_updated = False
        if description:
            existing_names = {clean(value) for value in style_rows["Product Name"] if clean(value)}
            if existing_names != {description}:
                frame.loc[style_mask, "Product Name"] = description
                description_updated = True
                result["updated_descriptions"] += 1
                changed = True

        existing_terms: set[str] = set()
        for _, row in frame.loc[style_mask].iterrows():
            existing_terms.update(_color_terms(row))
        incoming_colors: list[str] = []
        incoming_keys: set[str] = set()
        for raw in group["Garment Color"]:
            color = clean(raw)
            key = _normalized_catalog_text(color).replace("grey", "gray")
            if color and key and key not in incoming_keys:
                incoming_colors.append(color)
                incoming_keys.add(key)

        colors_added: list[str] = []
        for color in incoming_colors:
            color_key = _normalized_catalog_text(color).replace("grey", "gray")
            if color_key in existing_terms:
                continue
            template_rows = frame.loc[style_mask]
            completed_templates = template_rows.loc[~template_rows.apply(_is_setup_required, axis=1)]
            template = (completed_templates.iloc[0] if not completed_templates.empty else template_rows.iloc[0]).to_dict()
            row = {column: clean(template.get(column, "")) for column in frame.columns}
            row["Product Name"] = description or clean(row.get("Product Name", ""))
            row["Style Number"] = style
            row["Garment Color"] = color
            row["Vendor Color Code"] = ""
            row["Color Aliases"] = ""
            # Thread/ink is color-specific and must be confirmed for a newly seen garment color.
            row["Decoration Color"] = ""
            row["Setup Required"] = "Yes"
            frame = pd.concat([frame, pd.DataFrame([row], columns=frame.columns)], ignore_index=True)
            style_mask = frame["Style Number"].map(normalize_style).eq(style)
            existing_terms.add(color_key)
            colors_added.append(color)
            result["added_colors"] += 1
            changed = True

        if description_updated or colors_added:
            result["enriched_styles"] += 1
            details.append({
                "style": style,
                "description_updated": description_updated,
                "description": description if description_updated else "",
                "colors_added": colors_added,
            })

    if changed:
        backup_label = "repeat_import_repair" if repaired_count else "shopify_enrichment"
        _write_master_atomic(frame, master_path, backup_label)
    result["changed"] = changed
    result["details"] = details
    return result


def sync_review_product_candidates(review_path: Path, master_path: Path) -> dict[str, Any]:
    """Materialize only genuinely new review candidates into Product Master.

    Every candidate color is matched against the completed purchasing color,
    vendor color code, and Shopify color aliases before a setup row may be added.
    This keeps identical CSV replacements deterministic while still surfacing a
    truly new garment color for explicit setup.
    """
    review_path = Path(review_path)
    master_path = Path(master_path)
    empty_result = {
        "changed": False,
        "added_styles": 0,
        "added_colors": 0,
        "updated_fields": 0,
        "matched_existing_colors": 0,
        "repaired_duplicate_colors": 0,
        "repair_details": [],
    }
    if not review_path.exists():
        return empty_result

    frame = _read_master(master_path)
    frame, repaired_count, repair_details = _collapse_redundant_setup_color_rows(frame)
    candidates = _candidate_rows(review_path)
    if not candidates:
        if repaired_count:
            _write_master_atomic(frame, master_path, "repeat_import_repair")
        result = dict(empty_result)
        result.update({
            "changed": repaired_count > 0,
            "repaired_duplicate_colors": repaired_count,
            "repair_details": repair_details,
        })
        return result

    changed = repaired_count > 0
    added_styles = 0
    added_colors = 0
    updated_fields = 0
    matched_existing_colors = 0

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
            style_rows = frame.loc[style_mask]
            completed_templates = style_rows.loc[~style_rows.apply(_is_setup_required, axis=1)]
            template = (completed_templates.iloc[0] if not completed_templates.empty else style_rows.iloc[0]).to_dict()

        for color in colors:
            if style_exists:
                style_rows = frame.loc[style_mask]
                matching_indexes = _matching_color_indexes(style_rows, color)
            else:
                matching_indexes = []

            if matching_indexes:
                complete_matches = [index for index in matching_indexes if not _is_setup_required(frame.loc[index])]
                if complete_matches:
                    # The review workbook may still contain a stale candidate, but the
                    # live Product Master already has an authoritative completed match.
                    matched_existing_colors += 1
                    continue

                # The same incomplete row already exists. Fill only blank defaults and
                # keep it setup-required; do not create another row.
                for index in matching_indexes:
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
                # Thread/ink is color-specific and must be confirmed for a genuinely new color.
                "Decoration Color": "",
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
        backup_label = "repeat_import_repair" if repaired_count else "candidate_sync"
        _write_master_atomic(frame, master_path, backup_label)

    return {
        "changed": changed,
        "added_styles": added_styles,
        "added_colors": added_colors,
        "updated_fields": updated_fields,
        "matched_existing_colors": matched_existing_colors,
        "repaired_duplicate_colors": repaired_count,
        "repair_details": repair_details,
    }
