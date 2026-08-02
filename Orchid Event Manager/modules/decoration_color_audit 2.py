from __future__ import annotations

from datetime import datetime
from pathlib import Path
import os
import re
import shutil
from typing import Iterable

import pandas as pd

from modules.blank_garment_rules import is_blank_decoration
from modules.internal_services import is_in_house_decoration

COMMON_DECORATION_COLORS = [
    "", "White", "Black", "Navy", "Red", "Royal", "Gold", "Silver", "Gray",
    "Orange", "Purple", "Green", "Tan", "Khaki", "Maroon", "Light Blue",
]

# Cache exact Product Master decoration rules and fallback color defaults. The file
# stamp is part of the key, so a successful audit save automatically forces a reload.
_MASTER_LOOKUP_CACHE: dict[tuple[str, int, int], tuple[dict, dict, dict]] = {}


def clean(value: object) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return re.sub(r"\s+", " ", str(value).strip())


def normalize_style(value: object) -> str:
    return re.sub(r"\s+", "", clean(value)).upper()


def normalize_style_relaxed(value: object) -> str:
    return re.sub(r"[^A-Z0-9]+", "", normalize_style(value))


def normalize_color(value: object) -> str:
    return clean(value).casefold()


def normalize_product(value: object) -> str:
    text = clean(value).casefold().replace("®", " ").replace("™", " ")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def split_aliases(value: object) -> list[str]:
    return [clean(part) for part in re.split(r"[|;,\n]+", clean(value)) if clean(part)]


def decoration_family(value: object) -> str:
    text = clean(value).casefold()
    if "embroider" in text:
        return "Embroidery"
    if "screen" in text:
        return "Screen Print"
    return ""


def audit_item_needs_attention(
    item: dict,
    pending: dict | None = None,
    failure: dict | None = None,
) -> bool:
    """Return the single authoritative unresolved state for an audit row.

    A checked box is only a request to save; it cannot hide a row until the
    current-event verification is persisted. A saved event decision is
    authoritative for this event, while an actual recorded Product Master or
    current-event save failure remains unresolved and visible for correction.
    """
    if not isinstance(item, dict):
        return False
    pending = pending if isinstance(pending, dict) else {}
    failure = failure if isinstance(failure, dict) else {}
    if clean(failure.get("error", "")):
        return True
    if not bool(item.get("suspicious")):
        return False
    persisted_verified = bool(item.get("verified"))
    if pending.get("verified") is False:
        return True
    return not persisted_verified


def build_audit_view_snapshot(
    records: list[dict] | None,
    pending_edits: dict | None = None,
    failures: dict | None = None,
    selected_type: object = "All Decoration",
    issues_only: bool = True,
) -> dict:
    """Build metrics and visible rows from the exact same canonical list."""
    normalized_records = [item for item in (records or []) if isinstance(item, dict)]
    pending_edits = pending_edits if isinstance(pending_edits, dict) else {}
    failures = failures if isinstance(failures, dict) else {}
    selected = clean(selected_type)
    if selected not in {"Embroidery", "Screen Print"}:
        selected = "All Decoration"

    unresolved: list[dict] = []
    for item in normalized_records:
        key = clean(item.get("key", ""))
        pending = pending_edits.get(key, {})
        failure = failures.get(key, {})
        if audit_item_needs_attention(item, pending, failure):
            unresolved.append(item)

    unresolved_keys = {clean(item.get("key", "")) for item in unresolved}
    visible: list[dict] = []
    for item in normalized_records:
        if selected != "All Decoration" and item.get("decoration_type") != selected:
            continue
        if issues_only and clean(item.get("key", "")) not in unresolved_keys:
            continue
        visible.append(item)

    # Attention-first recovery: when a narrow decoration-type filter contains no
    # unresolved rows, show the complete unresolved list instead of presenting a
    # blank table beside a nonzero Need Attention count.
    if issues_only and unresolved and not visible:
        selected = "All Decoration"
        visible = list(unresolved)

    return {
        "records": normalized_records,
        "unresolved": unresolved,
        "visible": visible,
        "selected_type": selected,
        "total": len(normalized_records),
        "needs_attention": len(unresolved),
        "verified": sum(1 for item in normalized_records if bool(item.get("verified"))),
        "affected_lines": sum(len(item.get("line_records", [])) for item in normalized_records),
    }


def audit_group_key(
    decoration_type: object,
    style: object,
    product_name: object,
    garment_color: object,
    location: object,
) -> str:
    identity = normalize_style(style) or f"product:{normalize_product(product_name)}"
    return "|".join(
        (
            decoration_family(decoration_type).casefold(),
            identity.casefold(),
            normalize_color(garment_color),
            clean(location).casefold(),
        )
    )


def suggest_decoration_color(garment_color: object) -> str:
    """Return a visible suggestion only; never apply it automatically."""
    text = normalize_color(garment_color)
    if not text:
        return ""
    dark_terms = (
        "black", "navy", "dark", "charcoal", "royal", "purple", "maroon",
        "forest", "brown", "red", "burgundy", "graphite", "midnight",
    )
    light_terms = (
        "white", "natural", "cream", "ivory", "khaki", "tan", "sand",
        "silver", "light gray", "light grey", "yellow", "gold",
    )
    if any(term in text for term in dark_terms):
        return "White"
    if any(term in text for term in light_terms):
        return "Black"
    return ""


def _prefer_nonblank(existing: dict, field: str, value: object) -> None:
    text = clean(value)
    if text and not clean(existing.get(field, "")):
        existing[field] = text


def _master_lookups(master: pd.DataFrame) -> tuple[dict, dict, dict]:
    """Return exact rules, product defaults, and category color defaults."""
    exact_rules: dict[tuple[str, str], dict] = {}
    defaults: dict[str, dict] = {}
    category_candidates: dict[tuple[str, str], set[str]] = {}
    if master.empty:
        return exact_rules, defaults, {}

    for row in master.to_dict("records"):
        style = normalize_style(row.get("Style Number", ""))
        product = normalize_product(row.get("Product Name", ""))
        identity = style or f"product:{product}"
        if not style and not product:
            continue
        garment = normalize_color(row.get("Garment Color", ""))
        rule = {
            "color": clean(row.get("Decoration Color", "")),
            "location": clean(row.get("Decoration Location", "")),
            "decoration_type": clean(row.get("Decoration Type", "")),
        }
        category = clean(row.get("Product Category", "")).casefold()
        if garment:
            existing = exact_rules.setdefault((identity, garment), {})
            existing["_exists"] = True
            for field, value in rule.items():
                _prefer_nonblank(existing, field, value)
        else:
            existing = defaults.setdefault(identity, {})
            for field, value in rule.items():
                _prefer_nonblank(existing, field, value)
        if category and garment and rule["color"]:
            category_candidates.setdefault((category, garment), set()).add(rule["color"])
    category_defaults = {
        key: next(iter(values)) for key, values in category_candidates.items() if len(values) == 1
    }
    return exact_rules, defaults, category_defaults


def _cached_master_lookups(master_path: Path, *, force: bool = False) -> tuple[dict, dict, dict]:
    path = Path(master_path)
    if not path.exists():
        return {}, {}, {}
    stat = path.stat()
    cache_key = (str(path.resolve()), int(stat.st_mtime_ns), int(stat.st_size))
    cached = None if force else _MASTER_LOOKUP_CACHE.get(cache_key)
    if cached is not None:
        return cached
    try:
        master = pd.read_csv(path, dtype=str).fillna("")
    except Exception:
        master = pd.DataFrame()
    result = _master_lookups(master)
    _MASTER_LOOKUP_CACHE.clear()
    _MASTER_LOOKUP_CACHE[cache_key] = result
    return result


def build_decoration_color_audit(
    records: list[dict],
    master_path: Path,
    verified_map: dict | None = None,
) -> list[dict]:
    """Build style/color audit groups and validate them against exact master rules."""
    verified_map = verified_map if isinstance(verified_map, dict) else {}
    exact_master, default_master, category_master = _cached_master_lookups(master_path)

    groups: dict[str, dict] = {}
    for raw in records:
        include = clean(raw.get("Include", "Yes")).casefold()
        if include not in {"yes", "y", "true", "1"}:
            continue
        deco_type = decoration_family(raw.get("Decoration Type", ""))
        if (
            not deco_type
            or is_blank_decoration(raw.get("Decoration Type", ""))
            or is_in_house_decoration(raw.get("Decoration Type", ""))
        ):
            continue
        style = clean(raw.get("Product #", raw.get("Style Number", "")))
        product = clean(raw.get("Description", raw.get("Product Name", "")))
        garment = clean(raw.get("Garment Color", raw.get("Color", "")))
        location = clean(raw.get("Decoration Location", ""))
        key = audit_group_key(deco_type, style, product, garment, location)
        group = groups.setdefault(
            key,
            {
                "key": key,
                "decoration_type": deco_type,
                "style": style,
                "product_name": product,
                "garment_color": garment,
                "location": location,
                "product_category": clean(raw.get("Product Category", "")),
                "line_ids": [],
                "line_records": [],
                "orders": set(),
                "employees": set(),
                "colors": set(),
                "quantity": 0.0,
            },
        )
        line_id = clean(raw.get("Line ID", ""))
        if line_id:
            group["line_ids"].append(line_id)
        group["line_records"].append(raw)
        order = clean(raw.get("Order Number", ""))
        employee = clean(raw.get("Employee Name", ""))
        if order:
            group["orders"].add(order)
        if employee:
            group["employees"].add(employee)
        color = clean(raw.get("Decoration Color", ""))
        if color:
            group["colors"].add(color)
        try:
            group["quantity"] += float(raw.get("Quantity", 0) or 0)
        except Exception:
            pass

    output: list[dict] = []
    for key, group in groups.items():
        colors = sorted(group.pop("colors"), key=str.casefold)
        current_color = colors[0] if len(colors) == 1 else ""
        identity = normalize_style(group["style"]) or f"product:{normalize_product(group['product_name'])}"
        normalized_garment = normalize_color(group["garment_color"])
        exact_rule = dict(exact_master.get((identity, normalized_garment), {}) or {})
        default_rule = dict(default_master.get(identity, {}) or {})
        category_default = category_master.get(
            (clean(group.get("product_category", "")).casefold(), normalized_garment), ""
        )
        exact_color = clean(exact_rule.get("color", ""))
        exact_location = clean(exact_rule.get("location", ""))
        exact_type = decoration_family(exact_rule.get("decoration_type", ""))
        fallback_color = clean(default_rule.get("color", "")) or category_default
        master_color = exact_color or fallback_color
        exact_exists = bool(exact_rule.get("_exists"))
        master_source = (
            "Exact style + garment color" if exact_exists else
            ("Product-level default" if clean(default_rule.get("color", "")) else
             ("Category + garment color default" if category_default else "Not saved"))
        )

        event_location = clean(group["location"])
        master_synchronized = bool(
            exact_exists
            and current_color
            and normalize_color(exact_color) == normalize_color(current_color)
            and exact_location.casefold() == event_location.casefold()
            and (not exact_type or exact_type == group["decoration_type"])
        )

        issues: list[str] = []
        if not colors:
            issues.append("Decoration color missing")
        elif len(colors) > 1:
            issues.append("Multiple event colors: " + ", ".join(colors))
        if current_color and normalize_color(current_color) == normalize_color(group["garment_color"]):
            issues.append("Decoration matches garment color")
        if exact_exists:
            if not exact_color:
                issues.append("Exact Product Master decoration color missing")
            elif current_color and normalize_color(exact_color) != normalize_color(current_color):
                issues.append("Event differs from exact Product Master color")
            if not exact_location:
                issues.append("Exact Product Master decoration location missing")
            elif exact_location.casefold() != event_location.casefold():
                issues.append("Event differs from exact Product Master location")
        else:
            if fallback_color and current_color and normalize_color(fallback_color) != normalize_color(current_color):
                issues.append("Product or category default differs from this garment color")
            issues.append("No exact style + garment color rule")

        verified_record = verified_map.get(key, {})
        verified_color = clean(verified_record.get("color", "")) if isinstance(verified_record, dict) else ""
        verified_location = clean(verified_record.get("location", "")) if isinstance(verified_record, dict) else ""
        # A saved event decision must complete the current-event audit even when
        # Product Master has not yet produced an exact style + garment-color
        # match. That exact-rule status is still displayed for catalog follow-up,
        # but it cannot reopen an already-approved event decision forever.
        # Actual write failures are persisted separately and remain blocking.
        verified = bool(
            verified_record
            and normalize_color(verified_color) == normalize_color(current_color)
            and (not verified_location or verified_location.casefold() == event_location.casefold())
        )
        quantity = group["quantity"]
        group.update(
            {
                "colors_found": colors,
                "current_color": current_color,
                "master_color": master_color,
                "master_exact_color": exact_color,
                "master_location": exact_location,
                "master_source": master_source,
                "exact_master_exists": exact_exists,
                "master_synchronized": master_synchronized,
                "suggested_color": suggest_decoration_color(group["garment_color"]),
                "issues": issues,
                "suspicious": bool(issues),
                "verified": verified,
                "order_count": len(group.pop("orders")),
                "employee_count": len(group.pop("employees")),
                "quantity": int(quantity) if float(quantity).is_integer() else quantity,
            }
        )
        output.append(group)
    return sorted(
        output,
        key=lambda item: (
            0 if item["decoration_type"] == "Embroidery" else 1,
            normalize_style(item["style"]),
            normalize_color(item["garment_color"]),
            clean(item["location"]).casefold(),
        ),
    )


_SIZE_SUFFIXES = (
    "XXXXL", "XXXL", "XXL", "XL", "XS", "5XL", "4XL", "3XL", "2XL",
    "SM", "MD", "LG", "S", "M", "L",
)


def _style_variants(value: object) -> list[str]:
    raw = normalize_style(value)
    variants: list[str] = []
    if raw:
        variants.append(raw)
    relaxed = normalize_style_relaxed(raw)
    if relaxed and relaxed not in variants:
        variants.append(relaxed)
    # Shopify/vendor exports sometimes append a color or size after a delimiter.
    # Only strip explicit delimiters here so legitimate style numbers stay intact.
    for separator in ("-", "_", "/", "."):
        if separator in raw:
            pieces = [piece for piece in raw.split(separator) if piece]
            for stop in range(len(pieces) - 1, 0, -1):
                candidate = separator.join(pieces[:stop])
                if candidate and candidate not in variants:
                    variants.append(candidate)
                relaxed_candidate = normalize_style_relaxed(candidate)
                if relaxed_candidate and relaxed_candidate not in variants:
                    variants.append(relaxed_candidate)
    for suffix in _SIZE_SUFFIXES:
        if relaxed.endswith(suffix) and len(relaxed) > len(suffix) + 2:
            candidate = relaxed[:-len(suffix)]
            if candidate and candidate not in variants:
                variants.append(candidate)
    return variants


def _change_id(change: dict) -> str:
    return clean(change.get("key", "")) or audit_group_key(
        change.get("decoration_type", ""),
        change.get("style", ""),
        change.get("product_name", ""),
        change.get("garment_color", ""),
        change.get("location", ""),
    )


def _first_line_record(change: dict) -> dict:
    records = change.get("line_records", [])
    if isinstance(records, list):
        for record in records:
            if isinstance(record, dict):
                return record
    return {}


def _first_value(record: dict, *keys: str) -> str:
    for key in keys:
        value = clean(record.get(key, ""))
        if value:
            return value
    return ""


def _build_master_indexes(master: pd.DataFrame) -> dict[str, dict]:
    style_map: dict[str, list[int]] = {}
    relaxed_style_map: dict[str, list[int]] = {}
    relaxed_style_sources: dict[str, set[str]] = {}
    alias_map: dict[str, list[int]] = {}
    product_map: dict[str, list[int]] = {}
    exact_style_map: dict[tuple[str, str], list[int]] = {}
    exact_product_map: dict[tuple[str, str], list[int]] = {}
    for index, row in master.iterrows():
        style = normalize_style(row.get("Style Number", ""))
        relaxed = normalize_style_relaxed(style)
        product = normalize_product(row.get("Product Name", ""))
        garment = normalize_color(row.get("Garment Color", ""))
        if style:
            style_map.setdefault(style, []).append(index)
            exact_style_map.setdefault((style, garment), []).append(index)
        if relaxed:
            relaxed_style_map.setdefault(relaxed, []).append(index)
            relaxed_style_sources.setdefault(relaxed, set()).add(style)
        aliases: list[str] = []
        aliases.extend(_style_variants(row.get("Product ID", "")))
        for alias in split_aliases(row.get("Product Aliases", "")):
            aliases.extend(_style_variants(alias))
        for alias in dict.fromkeys(aliases):
            alias_map.setdefault(alias, []).append(index)
        if product:
            product_map.setdefault(product, []).append(index)
            exact_product_map.setdefault((product, garment), []).append(index)
    return {
        "style": style_map,
        "relaxed_style": relaxed_style_map,
        "relaxed_style_sources": relaxed_style_sources,
        "alias": alias_map,
        "product": product_map,
        "exact_style": exact_style_map,
        "exact_product": exact_product_map,
    }


def _unique_indexes(indexes: Iterable[int]) -> list[int]:
    return list(dict.fromkeys(int(index) for index in indexes))


def _source_indices(change: dict, indexes: dict[str, dict]) -> list[int]:
    style = clean(change.get("style", ""))
    product_key = normalize_product(change.get("product_name", ""))
    for variant in _style_variants(style):
        candidates = indexes["style"].get(variant, []) or indexes["alias"].get(variant, [])
        if candidates:
            return _unique_indexes(candidates)
        relaxed = normalize_style_relaxed(variant)
        candidates = indexes["relaxed_style"].get(relaxed, [])
        sources = indexes.get("relaxed_style_sources", {}).get(relaxed, set())
        if candidates and len(sources) == 1:
            return _unique_indexes(candidates)
    product_candidates = indexes["product"].get(product_key, []) if product_key else []
    return _unique_indexes(product_candidates) if len(set(product_candidates)) == 1 else []


def _exact_indices(change: dict, indexes: dict[str, dict]) -> list[int]:
    garment_key = normalize_color(change.get("garment_color", ""))
    style = clean(change.get("style", ""))
    for variant in _style_variants(style):
        exact = indexes["exact_style"].get((variant, garment_key), [])
        if exact:
            return _unique_indexes(exact)
    product_key = normalize_product(change.get("product_name", ""))
    exact = indexes["exact_product"].get((product_key, garment_key), []) if product_key else []
    return _unique_indexes(exact)


def _new_master_row(master: pd.DataFrame, change: dict) -> pd.Series:
    record = _first_line_record(change)
    row = pd.Series({column: "" for column in master.columns}, dtype="object")
    style = clean(change.get("style", "")) or _first_value(record, "Product #", "Style Number")
    product = clean(change.get("product_name", "")) or _first_value(record, "Description", "Product Name")
    garment = clean(change.get("garment_color", "")) or _first_value(record, "Garment Color", "Color")
    vendor = _first_value(record, "Purchase Vendor", "Vendor")
    category = clean(change.get("product_category", "")) or _first_value(record, "Product Category")

    values = {
        "Product Name": product,
        "Style Number": style,
        "Product ID": normalize_style(style),
        "Garment Color": garment,
        "Vendor": vendor,
        "Product Category": category,
        "Decoration Type": clean(change.get("decoration_type", "")),
        "Decoration Location": clean(change.get("location", "")),
        "Decoration Color": clean(change.get("decoration_color", "")),
        "Requires Size": _first_value(record, "Requires Size"),
        "Requires Color": _first_value(record, "Requires Color") or "Yes",
        "Requires Decoration": _first_value(record, "Requires Decoration") or "Yes",
        "Never Outsource": _first_value(record, "Never Outsource", "Do Not Outsource"),
    }
    for column, value in values.items():
        if column in row.index and value:
            row[column] = value
    complete = bool(style and product and garment and vendor and category and values["Decoration Type"] and values["Decoration Location"] and values["Decoration Color"])
    if "Setup Required" in row.index:
        row["Setup Required"] = "No" if complete else "Yes"
    return row


def apply_product_master_decoration_colors(master_path: Path, changes: list[dict]) -> dict:
    """Save exact style+garment-color rules, continuing past individual failures.

    Every successful change is written in one atomic Product Master replacement.
    Missing base styles are safely created from the current-event routing data so
    a completed audit can become a permanent exact rule instead of failing at the
    first new style.
    """
    master_path = Path(master_path)
    if not master_path.exists():
        raise FileNotFoundError(f"Product Master was not found: {master_path}")

    # Always reload from disk immediately before saving. This avoids stale lookup
    # indexes when Product Master was edited earlier in the same app session.
    _MASTER_LOOKUP_CACHE.clear()
    master = pd.read_csv(master_path, dtype=str).fillna("")
    required_columns = [
        "Product Name", "Style Number", "Product ID", "Product Aliases",
        "Garment Color", "Vendor", "Product Category", "Decoration Type",
        "Decoration Location", "Decoration Color", "Setup Required",
    ]
    for column in required_columns:
        if column not in master.columns:
            master[column] = ""

    indexes = _build_master_indexes(master)
    updated = 0
    added = 0
    created_styles = 0
    saved_keys: list[str] = []
    failures: list[dict] = []

    for change in changes:
        change_id = _change_id(change)
        style_display = clean(change.get("style", "")) or clean(change.get("product_name", "")) or "Unknown product"
        garment = clean(change.get("garment_color", ""))
        decoration_color = clean(change.get("decoration_color", ""))
        decoration_type = clean(change.get("decoration_type", ""))
        location = clean(change.get("location", ""))
        try:
            if not decoration_color:
                raise ValueError("Decoration color is blank.")
            if not garment:
                raise ValueError("Garment color is blank.")
            if not location:
                raise ValueError("Decoration location is blank.")

            exact_indices = _exact_indices(change, indexes)
            if exact_indices:
                for index in exact_indices:
                    master.at[index, "Decoration Color"] = decoration_color
                    if decoration_type:
                        master.at[index, "Decoration Type"] = decoration_type
                    master.at[index, "Decoration Location"] = location
                    master.at[index, "Setup Required"] = "No"
                    updated += 1
                saved_keys.append(change_id)
                continue

            source_indices = _source_indices(change, indexes)
            if source_indices:
                source_index = next(
                    (index for index in source_indices if not clean(master.at[index, "Garment Color"])),
                    source_indices[0],
                )
                new_row = master.loc[source_index].copy()
                new_row["Garment Color"] = garment
                new_row["Decoration Color"] = decoration_color
                if "Vendor Color Code" in new_row.index:
                    new_row["Vendor Color Code"] = ""
                if "Color Aliases" in new_row.index:
                    new_row["Color Aliases"] = ""
                if decoration_type:
                    new_row["Decoration Type"] = decoration_type
                new_row["Decoration Location"] = location
                new_row["Setup Required"] = "No"
            else:
                new_row = _new_master_row(master, change)
                if not clean(new_row.get("Style Number", "")) and not clean(new_row.get("Product Name", "")):
                    raise ValueError("No style number or product name was available to create a Product Master rule.")
                created_styles += 1

            master = pd.concat([master, pd.DataFrame([new_row])], ignore_index=True)
            added += 1
            indexes = _build_master_indexes(master)
            saved_keys.append(change_id)
        except Exception as error:
            failures.append({
                "key": change_id,
                "style": style_display,
                "garment_color": garment,
                "error": str(error) or error.__class__.__name__,
            })

    backup = None
    if saved_keys:
        backup_dir = master_path.parent / "backups"
        backup_dir.mkdir(parents=True, exist_ok=True)
        backup = backup_dir / f"product_master_before_decoration_color_audit_{datetime.now():%Y%m%d_%H%M%S}.csv"
        shutil.copy2(master_path, backup)

        temp = master_path.with_name(f".{master_path.name}.decoration-audit.{os.getpid()}.tmp")
        try:
            master.to_csv(temp, index=False)
            with temp.open("rb") as handle:
                os.fsync(handle.fileno())
            os.replace(temp, master_path)
        finally:
            try:
                temp.unlink(missing_ok=True)
            except Exception:
                pass
        _MASTER_LOOKUP_CACHE.clear()

    return {
        "updated": updated,
        "added": added,
        "created_styles": created_styles,
        "saved_keys": saved_keys,
        "failed": failures,
        "backup": backup,
    }
