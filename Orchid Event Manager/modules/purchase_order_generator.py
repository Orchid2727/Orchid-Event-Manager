from __future__ import annotations

from datetime import datetime
from pathlib import Path
import re
import pandas as pd

from modules.shopify_parser import parse_shopify_orders

MASTER_COLUMNS = [
    "Product Name", "Style Number", "Garment Color",
    "Vendor", "Decoration Type", "Decoration Color",
]

OUTPUT_COLUMNS = [
    "Vendor", "Decoration Type", "Decoration Color",
    "Style Number", "Product Name", "Garment Color", "Size", "Quantity",
]


def clean_text(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_style(value) -> str:
    return re.sub(r"\s+", "", clean_text(value)).upper()


def normalize_color(value) -> str:
    return re.sub(r"\s+", " ", clean_text(value)).strip().casefold()


def safe_filename(value) -> str:
    value = clean_text(value) or "Unassigned"
    value = re.sub(r"[^A-Za-z0-9._-]+", "_", value)
    return value.strip("_") or "Unassigned"


def load_master(product_master_path: Path) -> pd.DataFrame:
    master = pd.read_csv(product_master_path, dtype=str).fillna("")
    for column in MASTER_COLUMNS:
        if column not in master.columns:
            master[column] = ""
    master = master[MASTER_COLUMNS].copy()
    master["_style_key"] = master["Style Number"].map(normalize_style)
    master["_color_key"] = master["Garment Color"].map(normalize_color)

    # Keep the first populated assignment for each exact style/color combination.
    master["_score"] = (
        master["Vendor"].map(lambda x: bool(clean_text(x))).astype(int) * 4
        + master["Decoration Type"].map(lambda x: bool(clean_text(x))).astype(int) * 2
        + master["Decoration Color"].map(lambda x: bool(clean_text(x))).astype(int)
    )
    master = master.sort_values("_score", ascending=False)
    return master.drop_duplicates(["_style_key", "_color_key"], keep="first")


def build_purchase_order_data(
    shopify_csv_path: Path,
    product_master_path: Path,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    parsed = parse_shopify_orders(shopify_csv_path, product_master_path)
    master = load_master(product_master_path)

    parsed = parsed.copy()
    parsed["_style_key"] = parsed["Style Number"].map(normalize_style)
    parsed["_color_key"] = parsed["Garment Color"].map(normalize_color)

    merged = parsed.merge(
        master,
        on=["_style_key", "_color_key"],
        how="left",
        suffixes=("", "_master"),
    )

    for column in ["Vendor", "Decoration Type", "Decoration Color"]:
        merged[column] = merged[column].fillna("").map(clean_text)

    merged["Style Number"] = merged["Style Number"].map(clean_text)
    merged["Product Name"] = merged["Product Name"].map(clean_text)
    merged["Garment Color"] = merged["Garment Color"].map(clean_text)
    merged["Size"] = merged["Size"].map(clean_text)
    merged["Quantity"] = pd.to_numeric(merged["Quantity"], errors="coerce").fillna(0).astype(int)

    missing_reason = []
    for _, row in merged.iterrows():
        reasons = []
        if clean_text(row.get("Needs Review", "")):
            reasons.append(clean_text(row["Needs Review"]))
        if not clean_text(row["Style Number"]):
            reasons.append("Missing style number")
        if not clean_text(row["Garment Color"]):
            reasons.append("Missing garment color")
        if not clean_text(row["Vendor"]):
            reasons.append("Vendor not assigned in Product Master")
        if not clean_text(row["Decoration Type"]):
            reasons.append("Decoration type not assigned in Product Master")
        decoration_type = clean_text(row["Decoration Type"]).casefold()
        if decoration_type and decoration_type != "none" and not clean_text(row["Decoration Color"]):
            reasons.append("Decoration color not assigned in Product Master")
        missing_reason.append("; ".join(dict.fromkeys(reasons)))

    merged["Review Reason"] = missing_reason
    ready = merged[merged["Review Reason"].eq("")].copy()
    review = merged[merged["Review Reason"].ne("")].copy()

    if ready.empty:
        grouped = pd.DataFrame(columns=OUTPUT_COLUMNS)
    else:
        grouped = (
            ready.groupby(
                [
                    "Vendor", "Decoration Type", "Decoration Color",
                    "Style Number", "Product Name", "Garment Color", "Size",
                ],
                dropna=False,
                as_index=False,
            )["Quantity"]
            .sum()
            .sort_values(
                ["Vendor", "Decoration Type", "Decoration Color", "Style Number", "Garment Color", "Size"],
                key=lambda series: series.astype(str).str.casefold(),
            )
            .reset_index(drop=True)
        )

    review_columns = [
        "Order Number", "Company", "Original Line Item", "Product Name",
        "Style Number", "Garment Color", "Size", "Quantity", "Review Reason",
    ]
    review = review[review_columns].copy() if not review.empty else pd.DataFrame(columns=review_columns)

    summary_rows = []
    if not grouped.empty:
        for (vendor, decoration_type, decoration_color), group in grouped.groupby(
            ["Vendor", "Decoration Type", "Decoration Color"], dropna=False
        ):
            summary_rows.append({
                "Vendor": vendor,
                "Decoration Type": decoration_type,
                "Decoration Color": decoration_color,
                "Unique Lines": len(group),
                "Total Quantity": int(group["Quantity"].sum()),
            })
    summary = pd.DataFrame(
        summary_rows,
        columns=["Vendor", "Decoration Type", "Decoration Color", "Unique Lines", "Total Quantity"],
    )
    return grouped, review, summary


def generate_purchase_orders(
    shopify_csv_path: Path,
    product_master_path: Path,
    reports_root: Path,
) -> dict:
    grouped, review, summary = build_purchase_order_data(
        Path(shopify_csv_path), Path(product_master_path)
    )

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_dir = Path(reports_root) / "purchase_orders" / stamp
    output_dir.mkdir(parents=True, exist_ok=True)

    created_files = []

    summary_path = output_dir / "Purchase_Order_Summary.csv"
    summary.to_csv(summary_path, index=False)
    created_files.append(summary_path)

    if not review.empty:
        review_path = output_dir / "Review_Needed.csv"
        review.to_csv(review_path, index=False)
        created_files.append(review_path)

    if not grouped.empty:
        for (vendor, decoration_type, decoration_color), group in grouped.groupby(
            ["Vendor", "Decoration Type", "Decoration Color"], dropna=False
        ):
            parts = [safe_filename(vendor), safe_filename(decoration_type)]
            if clean_text(decoration_type).casefold() != "none":
                parts.append(safe_filename(decoration_color))
            filename = "__".join(parts) + ".csv"
            path = output_dir / filename
            group[OUTPUT_COLUMNS].to_csv(path, index=False)
            created_files.append(path)

    return {
        "output_dir": output_dir,
        "files": created_files,
        "ready_lines": len(grouped),
        "ready_quantity": int(grouped["Quantity"].sum()) if not grouped.empty else 0,
        "review_lines": len(review),
        "routes": len(summary),
    }
