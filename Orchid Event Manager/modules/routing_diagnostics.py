from __future__ import annotations

from datetime import datetime
from pathlib import Path
import re

import pandas as pd
import xlsxwriter

from modules.purchase_order_generator import (
    clean_text,
    identity_key,
    load_master,
    normalize_style,
    normalize_color,
    canonical_product,
    safe_filename,
)
from modules.shopify_parser import parse_shopify_orders
from modules.routing_rules import apply_routing_overrides, is_temporary_truspec_item
from modules.blank_garment_rules import BLANK_DECORATION_LABEL, is_blank_decoration
from modules.product_resolver import load_extended_master, resolve_product, normalize_phrase
from modules.purchase_rules import normalize_bool, row_rules

PURPLE = "#5B2AA8"
PURPLE_DARK = "#3D176F"
PURPLE_LIGHT = "#E8DDF6"
PURPLE_PALE = "#F4EFFB"
BORDER = "#CFC7D8"
TEXT = "#201A2D"
MUTED = "#6F667A"
YELLOW = "#FFF2CC"
RED_PALE = "#FCE8E6"
GREEN_PALE = "#E7F4EA"
BLUE_PALE = "#EAF2F8"


def _missing_fields(row: pd.Series) -> list[str]:
    missing = []
    if not clean_text(row.get("Vendor", "")):
        missing.append("Vendor")
    if not clean_text(row.get("Decoration Type", "")):
        missing.append("Decoration Type")
    deco = clean_text(row.get("Decoration Type", "")).casefold()
    if deco and not is_blank_decoration(deco) and not clean_text(row.get("Decoration Color", "")):
        missing.append("Decoration Color")
    return missing


def _join_unique(values) -> str:
    return ", ".join(dict.fromkeys(clean_text(v) for v in values if clean_text(v)))


def build_routing_diagnostics(shopify_csv_path: Path, product_master_path: Path) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    parsed = parse_shopify_orders(Path(shopify_csv_path), Path(product_master_path)).copy()
    master = load_extended_master(Path(product_master_path))

    for col in [
        "Style Number", "Product Name", "Garment Color", "Size", "Needs Review",
        "Original Line Item", "Order Number", "Company", "Employee Name",
        "Parser Source", "Parse Confidence", "Detected Vendor",
    ]:
        if col not in parsed.columns:
            parsed[col] = ""
        parsed[col] = parsed[col].fillna("").map(clean_text)
    if "Quantity" not in parsed.columns:
        parsed["Quantity"] = 0
    parsed["Quantity"] = pd.to_numeric(parsed["Quantity"], errors="coerce").fillna(0).astype(int)

    rows = []
    for idx, source in parsed.iterrows():
        result = resolve_product(
            source.get("Style Number", ""),
            source.get("Product Name", ""),
            source.get("Garment Color", ""),
            master,
        )
        record = source.to_dict()
        record.update(result.as_dict())
        if not result.matched and clean_text(source.get("Parser Source", "")).casefold() == "smart custom item":
            record["Master Match"] = "Smart Custom"
            record["Vendor"] = clean_text(source.get("Detected Vendor", ""))
        record["Product Name"] = clean_text(record.get("Master Product Name", "")) or clean_text(source.get("Product Name", ""))
        record["Style Number"] = clean_text(source.get("Style Number", "")) or clean_text(record.get("Master Style Number", ""))
        record["Garment Color"] = clean_text(record.get("Resolved Garment Color", "")) or clean_text(source.get("Garment Color", ""))
        routed = apply_routing_overrides(pd.DataFrame([record])).iloc[0]

        style = clean_text(routed.get("Style Number", ""))
        product = clean_text(routed.get("Product Name", ""))
        color = clean_text(routed.get("Garment Color", ""))
        size = clean_text(routed.get("Size", ""))
        vendor = clean_text(routed.get("Vendor", ""))
        deco_type = clean_text(routed.get("Decoration Type", ""))
        deco_color = clean_text(routed.get("Decoration Color", ""))
        parser_source = clean_text(routed.get("Parser Source", ""))

        rules = row_rules({
            "Product Name": product,
            "Style Number": style,
            "Product Category": routed.get("Product Category", ""),
            "Requires Size": routed.get("Requires Size", ""),
            "Requires Color": routed.get("Requires Color", ""),
            "Requires Decoration": routed.get("Requires Decoration", ""),
            "Decoration Type": deco_type,
        })
        requires_size = normalize_bool(rules["Requires Size"], True)
        requires_color = normalize_bool(rules["Requires Color"], True)
        requires_decoration = normalize_bool(rules["Requires Decoration"], True)

        blockers = []
        if not style:
            blockers.append("Missing product number")
        if not product:
            blockers.append("Missing description")
        if requires_color and not color:
            blockers.append("Missing garment color")
        if requires_size and not size:
            blockers.append("Missing size")
        if not vendor:
            blockers.append("Missing vendor")
        if requires_decoration and not deco_type:
            blockers.append("Missing decoration type")
        if requires_decoration and deco_type and not is_blank_decoration(deco_type) and not deco_color:
            blockers.append("Missing decoration color")

        master_match = clean_text(routed.get("Master Match", ""))
        if not blockers:
            severity = "OK"
            diagnostic_result = f"Purchase-order ready - {master_match or 'routed'}"
            action = "No action needed"
        elif master_match == "No" and parser_source.casefold() != "smart custom item":
            severity = "NEW PRODUCT"
            diagnostic_result = "No Product Master match"
            action = "Add or correct this product in Product Master"
        elif any(item in blockers for item in ["Missing vendor", "Missing decoration type", "Missing decoration color"]):
            severity = "FIX MASTER"
            diagnostic_result = "Product matched, but permanent routing is incomplete"
            action = "Complete the listed purchasing fields in Product Master"
        else:
            severity = "ORDER BLOCKER"
            diagnostic_result = "Purchase line is missing required order information"
            action = "Correct the Shopify/custom item or edit this line in the review workbook"

        rows.append({
            "Diagnostic ID": f"D{idx+1:05d}",
            "Severity": severity,
            "Order Number": clean_text(routed.get("Order Number", "")),
            "Company": clean_text(routed.get("Company", "")),
            "Employee": clean_text(routed.get("Employee Name", "")),
            "Parsed Product #": clean_text(source.get("Style Number", "")),
            "Resolved Product #": style,
            "Resolved Description": product,
            "Resolved Garment Color": color,
            "Parsed Size": size,
            "Quantity": int(routed.get("Quantity", 0)),
            "Master Match": master_match,
            "Product ID": clean_text(routed.get("Master Product ID", "")),
            "Vendor Color Code": clean_text(routed.get("Master Vendor Color Code", "")),
            "Product Category": rules["Product Category"],
            "Requires Size": rules["Requires Size"],
            "Requires Color": rules["Requires Color"],
            "Requires Decoration": rules["Requires Decoration"],
            "Matched Vendor": vendor,
            "Matched Decoration Type": deco_type,
            "Matched Decoration Color": deco_color,
            "Parser Source": parser_source,
            "Parse Confidence": clean_text(routed.get("Parse Confidence", "")),
            "Blocking Fields": "; ".join(blockers),
            "Diagnostic Result": diagnostic_result,
            "Recommended Action": action,
            "Original Shopify Line": clean_text(routed.get("Original Line Item", "")),
        })

    diagnostics = pd.DataFrame(rows)
    issues = diagnostics[diagnostics["Severity"].ne("OK")].copy() if not diagnostics.empty else diagnostics.copy()

    duplicate_frame = master.copy()
    duplicate_frame["_alias_key"] = duplicate_frame.apply(
        lambda row: f"{normalize_style(row.get('Style Number', ''))}|{normalize_phrase(row.get('Garment Color', ''))}",
        axis=1,
    )
    duplicate_counts = duplicate_frame.groupby("_alias_key", dropna=False).size().reset_index(name="Record Count")
    duplicate_counts = duplicate_counts[(duplicate_counts["Record Count"] > 1) & duplicate_counts["_alias_key"].ne("|")]
    if duplicate_counts.empty:
        duplicates = pd.DataFrame(columns=["Match Key", "Record Count", "Product #", "Description", "Garment Color"])
    else:
        duplicates = duplicate_counts.merge(duplicate_frame, on="_alias_key", how="left")
        duplicates = duplicates.rename(columns={"_alias_key": "Match Key", "Style Number": "Product #", "Product Name": "Description"})
        duplicates = duplicates[["Match Key", "Record Count", "Product #", "Description", "Garment Color"]]
    return diagnostics, issues, duplicates

def _write_table(workbook, ws, frame: pd.DataFrame, start_row: int, formats: dict, widths: dict | None = None):
    widths = widths or {}
    for col, name in enumerate(frame.columns):
        ws.write(start_row, col, name, formats["header"])
    for r, values in enumerate(frame.itertuples(index=False, name=None), start=start_row + 1):
        for c, value in enumerate(values):
            name = frame.columns[c]
            fmt = formats["qty"] if name in {"Quantity", "Record Count"} else formats["wrap"]
            ws.write(r, c, value, fmt)
    end_row = max(start_row + 1, start_row + len(frame))
    if len(frame.columns):
        ws.autofilter(start_row, 0, end_row, len(frame.columns)-1)
    for c, name in enumerate(frame.columns):
        ws.set_column(c, c, widths.get(name, 18))


def generate_routing_diagnostic_workbook(shopify_csv_path: Path, product_master_path: Path, reports_root: Path) -> dict:
    diagnostics, issues, duplicates = build_routing_diagnostics(shopify_csv_path, product_master_path)
    out_dir = Path(reports_root) / "Diagnostics"
    out_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = out_dir / f"Orchid_Routing_Diagnostics__{stamp}.xlsx"

    workbook = xlsxwriter.Workbook(output_path)
    formats = {
        "title": workbook.add_format({"bold": True, "font_size": 18, "font_color": PURPLE_DARK, "bg_color": PURPLE_LIGHT, "align": "left", "valign": "vcenter"}),
        "header": workbook.add_format({"bold": True, "font_color": TEXT, "bg_color": "#D8C8EA", "border": 1, "border_color": BORDER, "align": "center", "valign": "vcenter", "text_wrap": True}),
        "label": workbook.add_format({"bold": True, "bg_color": PURPLE_PALE, "border": 1, "border_color": BORDER}),
        "value": workbook.add_format({"border": 1, "border_color": BORDER}),
        "wrap": workbook.add_format({"border": 1, "border_color": BORDER, "valign": "top", "text_wrap": True}),
        "qty": workbook.add_format({"border": 1, "border_color": BORDER, "align": "center", "num_format": "0"}),
        "note": workbook.add_format({"font_color": MUTED, "bg_color": BLUE_PALE, "border": 1, "border_color": BORDER, "text_wrap": True}),
    }

    summary = workbook.add_worksheet("Diagnostic Summary")
    summary.hide_gridlines(2)
    summary.merge_range("A1:F1", "ORCHID ROUTING DIAGNOSTICS", formats["title"])
    summary.set_row(0, 34)
    metrics = [
        ("Shopify product lines analyzed", len(diagnostics)),
        ("Exact complete matches", int((diagnostics.get("Severity", pd.Series(dtype=str)) == "OK").sum())),
        ("Lines requiring attention", len(issues)),
        ("No Product Master match", int((diagnostics.get("Severity", pd.Series(dtype=str)) == "NEW PRODUCT").sum())),
        ("Order blockers", int((diagnostics.get("Severity", pd.Series(dtype=str)) == "ORDER BLOCKER").sum())),
        ("Matched records missing routing", int((diagnostics.get("Severity", pd.Series(dtype=str)) == "FIX MASTER").sum())),
        ("Duplicate Product Master keys", len(duplicates)),
    ]
    for r, (label, value) in enumerate(metrics, start=2):
        summary.write(r, 0, label, formats["label"])
        summary.write(r, 1, value, formats["value"])
    summary.merge_range("A11:F13", "Start with the Issues Only tab. Each row explains the exact key Orchid attempted, whether a Product Master match was found, and the recommended correction. Fix permanent product information in Product Master, then regenerate a new diagnostic workbook.", formats["note"])
    summary.set_column("A:A", 34)
    summary.set_column("B:B", 16)
    summary.set_column("C:F", 18)

    issue_ws = workbook.add_worksheet("Issues Only")
    issue_ws.hide_gridlines(2)
    issue_ws.merge_range(0, 0, 0, max(len(issues.columns)-1, 5), "ISSUES REQUIRING ATTENTION", formats["title"])
    issue_ws.freeze_panes(2, 0)
    _write_table(workbook, issue_ws, issues, 1, formats, {
        "Diagnostic ID": 12, "Severity": 18, "Order Number": 14, "Company": 24,
        "Employee": 22, "Parsed Product #": 16, "Parsed Description": 38,
        "Parsed Garment Color": 22, "Parsed Size": 12, "Quantity": 10,
        "Attempted Match Key": 45, "Exact Match Found": 15,
        "Matched Master Product #": 18, "Matched Master Description": 38,
        "Matched Master Color": 22, "Matched Vendor": 18,
        "Matched Decoration Type": 18, "Matched Decoration Color": 20,
        "Parser Review Reason": 40, "Diagnostic Result": 42,
        "Recommended Action": 55, "Original Shopify Line": 55,
    })
    if not issues.empty:
        last = len(issues) + 1
        issue_ws.conditional_format(2, 1, last, 1, {"type": "text", "criteria": "containing", "value": "NEW PRODUCT", "format": workbook.add_format({"bg_color": RED_PALE})})
        issue_ws.conditional_format(2, 1, last, 1, {"type": "text", "criteria": "containing", "value": "ORDER BLOCKER", "format": workbook.add_format({"bg_color": YELLOW})})
        issue_ws.conditional_format(2, 1, last, 1, {"type": "text", "criteria": "containing", "value": "FIX MASTER", "format": workbook.add_format({"bg_color": BLUE_PALE})})

    all_ws = workbook.add_worksheet("All Match Details")
    all_ws.hide_gridlines(2)
    all_ws.merge_range(0, 0, 0, max(len(diagnostics.columns)-1, 5), "ALL MATCH DETAILS", formats["title"])
    all_ws.freeze_panes(2, 0)
    _write_table(workbook, all_ws, diagnostics, 1, formats)

    dup_ws = workbook.add_worksheet("Duplicate Master Keys")
    dup_ws.hide_gridlines(2)
    dup_ws.merge_range(0, 0, 0, max(len(duplicates.columns)-1, 5), "DUPLICATE PRODUCT MASTER MATCH KEYS", formats["title"])
    _write_table(workbook, dup_ws, duplicates, 1, formats, {"Match Key": 50, "Description": 38, "Garment Color": 22})

    workbook.close()
    return {
        "output_path": output_path,
        "total_lines": len(diagnostics),
        "issue_lines": len(issues),
        "complete_matches": int((diagnostics.get("Severity", pd.Series(dtype=str)) == "OK").sum()),
        "new_products": int((diagnostics.get("Severity", pd.Series(dtype=str)) == "NEW PRODUCT").sum()),
        "mismatches": int((diagnostics.get("Severity", pd.Series(dtype=str)) == "ORDER BLOCKER").sum()),
        "incomplete_matches": int((diagnostics.get("Severity", pd.Series(dtype=str)) == "FIX MASTER").sum()),
    }
