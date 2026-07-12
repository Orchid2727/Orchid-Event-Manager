from __future__ import annotations

from datetime import datetime
from pathlib import Path
import re
import shutil
import pandas as pd

COLUMNS = [
    "Product Name", "Style Number", "Garment Color",
    "Vendor", "Decoration Type", "Decoration Color",
]


def clean_text(value) -> str:
    if pd.isna(value):
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
    master = master.copy()
    for col in COLUMNS:
        if col not in master.columns:
            master[col] = ""
    master = master[COLUMNS].fillna("")
    for col in COLUMNS:
        master[col] = master[col].map(clean_text)

    # Canonicalize fields before grouping.
    master["Product Name"] = master["Product Name"].map(canonical_product_name)
    master["Style Number"] = master["Style Number"].map(normalize_style)
    master["Garment Color"] = master["Garment Color"].map(normalize_space)

    # Use style+color when style exists. Otherwise product+color.
    def key(row):
        style = normalize_style(row["Style Number"])
        color = normalize_space(row["Garment Color"]).casefold()
        product = canonical_product_name(row["Product Name"]).casefold()
        return ("style", style, color) if style else ("product", product, color)

    master["_key"] = master.apply(key, axis=1)
    rows = []
    for _, group in master.groupby("_key", sort=False, dropna=False):
        rows.append({
            "Product Name": first_nonblank(group["Product Name"]),
            "Style Number": first_nonblank(group["Style Number"]),
            "Garment Color": first_nonblank(group["Garment Color"]),
            "Vendor": first_nonblank(group["Vendor"]),
            "Decoration Type": first_nonblank(group["Decoration Type"]),
            "Decoration Color": first_nonblank(group["Decoration Color"]),
        })

    result = pd.DataFrame(rows, columns=COLUMNS)
    if result.empty:
        return result
    return result.sort_values(
        ["Style Number", "Product Name", "Garment Color"],
        key=lambda s: s.astype(str).str.casefold(),
    ).reset_index(drop=True)


def backup_file(path: Path) -> Path | None:
    if not path.exists():
        return None
    backup_dir = path.parent / "backups"
    backup_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = backup_dir / f"product_master_backup_{stamp}.csv"
    shutil.copy2(path, backup)
    return backup


def clean_product_master(path: Path, make_backup: bool = True):
    path.parent.mkdir(parents=True, exist_ok=True)
    if not path.exists():
        empty = pd.DataFrame(columns=COLUMNS)
        empty.to_csv(path, index=False)
        return {"before": 0, "after": 0, "removed": 0, "backup": None}

    original = pd.read_csv(path, dtype=str).fillna("")
    backup = backup_file(path) if make_backup else None
    cleaned = clean_master(original)
    cleaned.to_csv(path, index=False)
    return {
        "before": len(original),
        "after": len(cleaned),
        "removed": len(original) - len(cleaned),
        "backup": backup,
    }


def merge_parsed_products(master_path: Path, parsed: pd.DataFrame):
    if master_path.exists():
        master = pd.read_csv(master_path, dtype=str).fillna("")
    else:
        master = pd.DataFrame(columns=COLUMNS)

    incoming = pd.DataFrame({
        "Product Name": parsed.get("Product Name", ""),
        "Style Number": parsed.get("Style Number", ""),
        "Garment Color": parsed.get("Garment Color", ""),
        "Vendor": "",
        "Decoration Type": "",
        "Decoration Color": "",
    })

    combined = pd.concat([master, incoming], ignore_index=True)
    backup = backup_file(master_path) if master_path.exists() else None
    cleaned = clean_master(combined)
    cleaned.to_csv(master_path, index=False)
    return {
        "existing_rows": len(master),
        "incoming_rows": len(incoming),
        "final_rows": len(cleaned),
        "backup": backup,
    }
