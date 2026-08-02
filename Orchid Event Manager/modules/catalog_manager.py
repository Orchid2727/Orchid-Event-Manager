from __future__ import annotations

from datetime import datetime
from pathlib import Path
import re
import shutil

import pandas as pd

from modules.blank_garment_rules import (
    KNOWN_BLANK_GARMENT_STYLES,
    apply_blank_garment_defaults,
    is_blank_decoration,
    is_blank_garment_product,
)
from modules.paths import backups_dir, seed_product_master_path
from modules.outsource_rules import apply_never_outsource_defaults
from modules.purchase_rules import apply_purchase_rule_defaults, category_defaults, infer_category
from modules.product_intelligence import apply_product_intelligence
from modules.product_resolver import MASTER_COLUMNS as COLUMNS, default_product_id, ensure_master_columns
from modules.internal_services import is_in_house_service_product, is_phantom_product_label


def clean_text(value) -> str:
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()


def normalize_space(value) -> str:
    return re.sub(r"\s+", " ", clean_text(value)).strip()


def normalize_style(value) -> str:
    return re.sub(r"\s+", "", clean_text(value)).upper()


def canonical_product_name(value) -> str:
    name = normalize_space(value)
    return re.sub(r"[\s\.\-–—:|/]+$", "", name).strip()


def first_nonblank(values) -> str:
    for value in values:
        value = clean_text(value)
        if value:
            return value
    return ""


def clean_master(master: pd.DataFrame) -> pd.DataFrame:
    master = ensure_master_columns(master)
    for col in COLUMNS:
        master[col] = master[col].map(clean_text)
    master["Product Name"] = master["Product Name"].map(canonical_product_name)
    master["Style Number"] = master["Style Number"].map(normalize_style)
    master["Garment Color"] = master["Garment Color"].map(normalize_space)
    # Keep the Settings-page cleanup consistent with the Product Master editor.
    # Otherwise an old service charge could disappear when Product Master opens
    # but return if the catalog was later cleaned from Settings.
    internal_mask = master.apply(
        lambda row: is_in_house_service_product(
            row.get("Product Name", ""),
            decoration_type=row.get("Decoration Type", ""),
            style_number=row.get("Style Number", ""),
            garment_color=row.get("Garment Color", ""),
        ) or is_phantom_product_label(row.get("Product Name", ""))
          or is_phantom_product_label(row.get("Style Number", "")),
        axis=1,
    )
    master = master.loc[~internal_mask].copy()

    def key(row):
        style = normalize_style(row["Style Number"])
        color = normalize_space(row["Garment Color"]).casefold()
        product = canonical_product_name(row["Product Name"]).casefold()
        return ("style", style, color) if style else ("product", product, color)

    if master.empty:
        return pd.DataFrame(columns=COLUMNS)

    master["_key"] = master.apply(key, axis=1)
    rows = []
    for _, group in master.groupby("_key", sort=False, dropna=False):
        record = {column: first_nonblank(group[column]) for column in COLUMNS}
        if not record["Product ID"]:
            record["Product ID"] = default_product_id(record["Style Number"], record["Product Name"])
        rows.append(record)

    result = pd.DataFrame(rows, columns=COLUMNS)

    # Repair the legacy "short sleeve" classification bug. Older versions treated
    # the word "short" in "short sleeve shirt" as a pair of shorts, which saved
    # tees and shirts as blank garments. Repair only that clear conflict and leave
    # intentional Orchid assignments untouched.
    for index, row in result.iterrows():
        product_name = clean_text(row.get("Product Name", ""))
        style = normalize_style(row.get("Style Number", ""))
        current_category = clean_text(row.get("Product Category", ""))
        inferred_category = infer_category(product_name, style)
        truly_blank = is_blank_garment_product(product_name, style_number=style)
        legacy_conflict = (
            current_category == "Pants / Jeans / Shorts"
            and inferred_category != "Pants / Jeans / Shorts"
            and not truly_blank
        )
        if legacy_conflict:
            result.at[index, "Product Category"] = inferred_category
            defaults = category_defaults(inferred_category, product_name, style)
            result.at[index, "Requires Size"] = "Yes" if defaults["requires_size"] else "No"
            result.at[index, "Requires Color"] = "Yes" if defaults["requires_color"] else "No"
            result.at[index, "Requires Decoration"] = "Yes" if defaults["requires_decoration"] else "No"
            if is_blank_decoration(row.get("Decoration Type", "")):
                result.at[index, "Decoration Type"] = ""
                result.at[index, "Decoration Color"] = ""
        elif current_category in {"", "Other"} and inferred_category not in {"", "Other"}:
            result.at[index, "Product Category"] = inferred_category

        # Some Wrangler/VF style numbers already contain the color code. Requiring
        # a separate color name makes otherwise order-ready pants appear blocked.
        if style in KNOWN_BLANK_GARMENT_STYLES and not clean_text(row.get("Garment Color", "")):
            result.at[index, "Requires Color"] = "No"

    # Apply permanent purchasing rules before saving the cleaned Product Master.
    # This ensures a known insole is repaired to Boots / no color / no
    # decoration even when the record was created in an older app version.
    result = apply_never_outsource_defaults(
        apply_product_intelligence(
            apply_purchase_rule_defaults(
                apply_blank_garment_defaults(result)
            )
        )
    )
    if result.empty:
        return result
    return result.sort_values(
        ["Style Number", "Product Name", "Garment Color"],
        key=lambda s: s.astype(str).str.casefold(),
    ).reset_index(drop=True)


def backup_file(path: Path, label: str = "backup") -> Path | None:
    path = Path(path)
    if not path.exists():
        return None
    backup_root = backups_dir()
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_label = re.sub(r"[^A-Za-z0-9_-]+", "_", label).strip("_") or "backup"
    backup = backup_root / f"product_master_{safe_label}_{stamp}.csv"
    shutil.copy2(path, backup)
    return backup


def validate_product_master_file(path: Path) -> pd.DataFrame:
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"Product Master file was not found: {path}")
    try:
        frame = pd.read_csv(path, dtype=str).fillna("")
    except pd.errors.EmptyDataError as exc:
        raise ValueError("The selected Product Master file is empty.") from exc
    required = {"Product Name", "Style Number", "Vendor", "Garment Color"}
    missing = sorted(required.difference(frame.columns))
    if missing:
        raise ValueError("The selected CSV is not a Product Master file. Missing: " + ", ".join(missing))
    return clean_master(frame)


def restore_product_master(active_path: Path, source_path: Path, label: str = "before_restore") -> dict:
    active_path = Path(active_path)
    source_path = Path(source_path)
    restored = validate_product_master_file(source_path)
    active_path.parent.mkdir(parents=True, exist_ok=True)
    backup = backup_file(active_path, label=label) if active_path.exists() else None
    temp_path = active_path.with_suffix(".restore.tmp")
    restored.to_csv(temp_path, index=False)
    temp_path.replace(active_path)
    return {
        "backup": backup,
        "source": source_path,
        "records": len(restored),
        "styles": count_styles(restored),
    }


def restore_seed_product_master(active_path: Path) -> dict:
    return restore_product_master(active_path, seed_product_master_path(), label="before_clean_restore")


def count_styles(frame: pd.DataFrame) -> int:
    if frame.empty:
        return 0
    keys: set[str] = set()
    for _, row in frame.iterrows():
        style = normalize_style(row.get("Style Number", ""))
        product = canonical_product_name(row.get("Product Name", "")).casefold()
        keys.add(f"style:{style}" if style else f"product:{product}")
    return len(keys)


def clean_product_master(path: Path, make_backup: bool = True):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    if not path.exists():
        empty = pd.DataFrame(columns=COLUMNS)
        empty.to_csv(path, index=False)
        return {"before": 0, "after": 0, "removed": 0, "backup": None}
    original = pd.read_csv(path, dtype=str).fillna("")
    backup = backup_file(path, label="before_cleanup") if make_backup else None
    cleaned = clean_master(original)
    cleaned.to_csv(path, index=False)
    return {"before": len(original), "after": len(cleaned), "removed": len(original) - len(cleaned), "backup": backup}


def _master_style_keys(master: pd.DataFrame) -> set[str]:
    keys: set[str] = set()
    if master.empty:
        return keys
    for _, row in master.iterrows():
        style = normalize_style(row.get("Style Number", ""))
        name = canonical_product_name(row.get("Product Name", "")).casefold()
        if style:
            keys.add(f"style:{style}")
        elif name:
            keys.add(f"product:{name}")
    return keys


def build_product_candidate_list(master_path: Path, parsed: pd.DataFrame) -> pd.DataFrame:
    """Create a safe preview of genuinely new products without changing Product Master.

    One candidate is produced per style/product. Shopify colors and quantities stay
    in the current purchase review; they are not mass-added as incomplete permanent
    Product Master records.
    """
    if Path(master_path).exists():
        master = pd.read_csv(master_path, dtype=str).fillna("")
    else:
        master = pd.DataFrame(columns=COLUMNS)
    known = _master_style_keys(master)

    source = parsed.copy()
    if source.empty:
        return pd.DataFrame(columns=["Product Name", "Style Number", "Example Color", "Order Lines", "Action"])
    if "Parser Source" in source.columns:
        source = source[~source["Parser Source"].astype(str).str.casefold().eq("smart custom item")].copy()

    candidates: dict[str, dict[str, object]] = {}
    for _, row in source.iterrows():
        style = normalize_style(row.get("Style Number", ""))
        name = canonical_product_name(row.get("Product Name", ""))
        color = normalize_space(row.get("Garment Color", ""))
        key = f"style:{style}" if style else f"product:{name.casefold()}"
        if not (style or name) or key in known:
            continue
        entry = candidates.setdefault(key, {
            "Product Name": name,
            "Style Number": style,
            "Example Color": color,
            "Order Lines": 0,
            "Action": "Add manually in Product Master only if this is a permanent catalog product",
        })
        entry["Order Lines"] = int(entry["Order Lines"]) + 1
        if not entry["Example Color"] and color:
            entry["Example Color"] = color

    frame = pd.DataFrame(candidates.values(), columns=["Product Name", "Style Number", "Example Color", "Order Lines", "Action"])
    if frame.empty:
        return frame
    return frame.sort_values(
        ["Style Number", "Product Name"], key=lambda series: series.astype(str).str.casefold()
    ).reset_index(drop=True)


def save_product_candidate_list(master_path: Path, parsed: pd.DataFrame, reports_root: Path) -> dict:
    candidates = build_product_candidate_list(master_path, parsed)
    folder = Path(reports_root) / "Product Master Candidates"
    folder.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output = folder / f"purchasing_rules_candidates_{stamp}.csv"
    candidates.to_csv(output, index=False)
    return {"output_path": output, "candidate_count": len(candidates)}


def merge_parsed_products(master_path: Path, parsed: pd.DataFrame):
    """Deprecated compatibility wrapper.

    2.0 never mass-adds Shopify order lines to Product Master. Return a preview
    result so older callers cannot accidentally recreate hundreds of incomplete
    permanent records.
    """
    candidates = build_product_candidate_list(master_path, parsed)
    if Path(master_path).exists():
        master = pd.read_csv(master_path, dtype=str).fillna("")
    else:
        master = pd.DataFrame(columns=COLUMNS)
    return {
        "existing_rows": len(master),
        "incoming_rows": len(parsed),
        "final_rows": len(master),
        "candidate_rows": len(candidates),
        "backup": None,
    }
