from __future__ import annotations

from datetime import datetime
from pathlib import Path
import json
import re

import pandas as pd
import xlsxwriter

from modules.blank_garment_rules import BLANK_DECORATION_LABEL, is_blank_decoration
from modules.decoration_locations import (
    DECORATION_LOCATIONS,
    normalize_decoration_location,
)
from modules.note_rules import (
    combine_purchase_instructions,
    decision_note_requires_review,
    decoration_note_requires_review,
    decoration_note_targets_line,
    decoration_note_is_global,
    decoration_target_families,
    decoration_target_styles,
    note_action_label,
    note_requires_review,
    note_requires_review_for_line,
)

from modules.purchase_order_generator import (
    clean_text,
    identity_key,
    load_master,
    normalize_size,
    normalize_style,
    split_embedded_size,
    safe_filename,
)
from modules.shopify_parser import is_decoration_service, normalize_order_export, parse_shopify_orders
from modules.routing_rules import apply_routing_overrides
from modules.product_resolver import load_extended_master, resolve_product
from modules.purchase_rules import normalize_bool, row_rules
from modules.internal_services import (
    HEMMING_ALTERATION_LABEL, SEW_ON_PATCH_LABEL,
    is_in_house_decoration, is_in_house_service_product,
)
from modules.report_modes import GENERAL_SALES_PERIOD, UNIFORM_SIZING_EVENT, normalize_report_mode
from modules.decoration_fulfillment import (
    STANDARD_ORCHID_WORKFLOW, is_entire_order_outsourced, normalize_decoration_fulfillment,
)
from modules.master_sync import product_master_signature
from modules.outsource_rules import vendor_never_outsource
from modules.decoration_inference import (
    count_unique_decisions,
    infer_order_decorations,
    service_totals_from_raw,
    unresolved_decision_key,
)

PURPLE = "#5B2AA8"
PURPLE_DARK = "#3D176F"
PURPLE_LIGHT = "#E8DDF6"
PURPLE_PALE = "#F4EFFB"
BORDER = "#CFC7D8"
TEXT = "#201A2D"
MUTED = "#6F667A"
WHITE = "#FFFFFF"
YELLOW = "#FFF2CC"
RED_PALE = "#FCE8E6"
GREEN_PALE = "#E7F4EA"
BLUE_PALE = "#EAF2F8"

PREFERRED_VENDOR_NAMES = [
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
PREFERRED_VENDOR_CASE = {vendor.casefold(): vendor for vendor in PREFERRED_VENDOR_NAMES}

PERMANENT_GAP_LABELS = {
    "missing purchase vendor",
    "missing decoration type",
    "missing decoration color",
    "product master setup required",
}
EVENT_GAP_LABELS = {
    "missing product number",
    "missing description",
    "missing garment color",
    "missing size",
    "quantity must be greater than zero",
    "missing customer or order name",
    "customer decision required",
}


def _reason_parts(value: object) -> list[str]:
    return [part.strip() for part in clean_text(value).split(";") if part.strip()]


def _looks_like_style_code(value: object) -> bool:
    # Style matching is intentionally case-insensitive and whitespace-insensitive.
    text = normalize_style(value)
    if not text or len(text) > 30:
        return False
    return bool(re.search(r"[A-Za-z]", text) and re.search(r"[0-9]", text)) or text.isdigit()


def _issue_buckets(row: pd.Series) -> tuple[list[str], list[str], bool]:
    """Split one line's blockers into permanent setup and order-specific work.

    A line can legitimately need both. For example, a newly imported style can need
    a permanent vendor assignment while the individual order also needs a size.
    Product Master must receive the reusable setup first; Purchase Review keeps only
    the event-specific correction.
    """
    parts = _reason_parts(row.get("Review Reason", ""))
    permanent: list[str] = []
    event: list[str] = []
    for part in parts:
        lowered = part.casefold()
        if lowered in PERMANENT_GAP_LABELS:
            permanent.append(part)
        elif lowered in EVENT_GAP_LABELS:
            event.append(part)
        else:
            # Unknown warnings are safer as current-order review unless a more
            # specific rule below identifies a genuine style candidate.
            event.append(part)

    style = clean_text(row.get("Product #", ""))
    source = clean_text(row.get("Source", "")).casefold()
    master_match = clean_text(row.get("Master Match", "")).casefold()
    unknown_candidate = bool(
        style
        and _looks_like_style_code(style)
        and master_match == "no"
        and source != "smart custom item"
    )
    if unknown_candidate:
        permanent.insert(0, "Product not found in Product Master")
    return list(dict.fromkeys(permanent)), list(dict.fromkeys(event)), unknown_candidate


def _fix_location(row: pd.Series) -> str:
    permanent, event, _unknown_candidate = _issue_buckets(row)
    if permanent and event:
        return "Product Master + Purchase Review"
    if permanent:
        return "Product Master"
    return "Purchase Review"


def _product_master_decision_key(row: pd.Series) -> str:
    style = normalize_style(row.get("Product #", ""))
    description = clean_text(row.get("Description", "")).casefold()
    return f"product:{style or description or clean_text(row.get('Line ID', ''))}"


def _action_required(row: pd.Series) -> str:
    permanent, event, unknown_candidate = _issue_buckets(row)
    style = clean_text(row.get("Product #", ""))
    actions: list[str] = []
    if permanent:
        permanent_reason = "; ".join(permanent)
        if unknown_candidate:
            actions.append(
                f"Add {style or 'this product'} to Product Master and complete: {permanent_reason}."
            )
        else:
            actions.append(
                f"Update {style or 'this product'} in Product Master: {permanent_reason}."
            )
    if event:
        event_reason = "; ".join(event)
        if "missing product number" in event_reason.casefold():
            actions.append("Correct the product number for this order, or mark it as a one-time manual item.")
        else:
            actions.append(f"Then correct this order in Purchase Review: {event_reason}.")
    return " ".join(actions) or "Complete the missing information."


ALL_COLUMNS = [
    "Line ID",
    "Include",
    "Do Not Outsource",
    "Purchase Vendor",
    "Decoration Type",
    "Decoration Location",
    "Decoration Placement Instructions",
    "Decoration Color",
    "Product #",
    "Description",
    "Garment Color",
    "Size",
    "Quantity",
    "Company",
    "Employee Name",
    "Order Number",
    "Shopify Order Notes",
    "Purchase Instructions",
    "Note Action",
    "Review Status",
    "Review Reason",
    "Product Category",
    "Requires Size",
    "Requires Color",
    "Requires Decoration",
    "Decoration Source",
    "Decision Key",
    "Source",
    "Parse Confidence",
    "Master Match",
    "Original Shopify Line",
    "Order Date",
    "Shopify Line Notes",
    "Decoration Decision",
    "Mandatory Note Review",
]


def _line_note_context(row: pd.Series) -> str:
    return " ".join(filter(None, [
        clean_text(row.get("Product Name", row.get("Description", ""))),
        clean_text(row.get("Original Line Item", row.get("Original Shopify Line", ""))),
        clean_text(row.get("Product Category", "")),
        clean_text(row.get("Style Number", row.get("Product #", ""))),
    ]))


def _mandatory_note_review_flags(frame: pd.DataFrame) -> pd.Series:
    """Mark the exact purchase lines that must stop for a note decision.

    Order-level decoration notes are routed to named garment families when the note
    identifies one. If the wording cannot be matched to any line, all lines remain
    blocked so a safety-critical instruction can never disappear silently.
    """
    flags = pd.Series(False, index=frame.index, dtype=bool)
    if frame.empty:
        return flags
    grouped = frame.groupby("Order Number", sort=False, dropna=False) if "Order Number" in frame.columns else [("", frame)]
    for _order, group in grouped:
        indices = list(group.index)
        order_note = next((clean_text(value) for value in group.get("Shopify Order Notes", pd.Series(dtype=str)) if clean_text(value)), "")

        for index, row in group.iterrows():
            line_note = clean_text(row.get("Shopify Line Notes", ""))
            if note_requires_review(line_note):
                flags.at[index] = True

        if decision_note_requires_review(order_note):
            flags.loc[indices] = True
            continue
        if not decoration_note_requires_review(order_note):
            continue

        style_targets = decoration_target_styles(order_note)
        family_targets = decoration_target_families(order_note)
        is_global = decoration_note_is_global(order_note)
        if style_targets or family_targets or is_global:
            matches = [
                index for index, row in group.iterrows()
                if decoration_note_targets_line(order_note, _line_note_context(row))
            ]
        else:
            # An unclear order-level note becomes one review decision instead of
            # being duplicated on every product. The user can still see the full
            # order note and correct the affected line without reviewing unrelated
            # items repeatedly.
            matches = indices[:1]
        if not matches:
            matches = indices[:1]
        flags.loc[matches] = True
    return flags


def _sheet_name(value: str) -> str:
    value = re.sub(r"[\\/*?:\[\]]+", "-", clean_text(value))
    return (value or "Unassigned")[:31]


def _unique_sheet_name(workbook, value: str) -> str:
    """Return a worksheet name that is unique under Excel's case-insensitive rules."""
    base = _sheet_name(value)
    sheetnames = getattr(workbook, "sheetnames", {})
    names = sheetnames.keys() if isinstance(sheetnames, dict) else sheetnames
    used = {clean_text(name).casefold() for name in names}
    if base.casefold() not in used:
        return base
    counter = 2
    while True:
        suffix = f" ({counter})"
        candidate = f"{base[:31 - len(suffix)]}{suffix}"
        if candidate.casefold() not in used:
            return candidate
        counter += 1


def _canonical_vendor(value: object, registry: dict[str, str] | None = None) -> str:
    value = re.sub(r"\s+", " ", clean_text(value)).strip()
    if not value:
        return ""
    key = value.casefold()
    preferred = PREFERRED_VENDOR_CASE.get(key)
    if preferred:
        if registry is not None:
            registry[key] = preferred
        return preferred
    if registry is not None:
        existing = registry.get(key)
        if existing:
            return existing
        registry[key] = value
    return value


def _build_review_data(
    shopify_csv_path: Path, product_master_path: Path, decoration_fulfillment: str = STANDARD_ORCHID_WORKFLOW
):
    raw = normalize_order_export(shopify_csv_path)
    if "Name" in raw.columns:
        raw["Name"] = raw["Name"].replace("", pd.NA).ffill().fillna("")
    for column in ["Billing Company", "Billing Name", "Notes", "Created at"]:
        if column in raw.columns and "Name" in raw.columns:
            raw[column] = (
                raw[column].replace("", pd.NA)
                .groupby(raw["Name"], dropna=False)
                .transform(lambda values: values.ffill().bfill())
                .fillna("")
            )

    parsed = parse_shopify_orders(shopify_csv_path, product_master_path).copy()
    master = load_extended_master(product_master_path)

    for column in [
        "Order Number", "Order Date", "Company", "Employee Name",
        "Shopify Order Notes", "Shopify Line Notes", "Original Line Item", "Product Name",
        "Style Number", "Garment Color", "Size", "Needs Review",
        "Parser Source", "Parse Confidence", "Detected Vendor",
    ]:
        if column not in parsed.columns:
            parsed[column] = ""
        parsed[column] = parsed[column].fillna("").map(clean_text)

    if "Quantity" not in parsed.columns:
        parsed["Quantity"] = 0
    parsed["Quantity"] = pd.to_numeric(parsed["Quantity"], errors="coerce").fillna(0).astype(int)

    resolved_records = []
    for _, source in parsed.iterrows():
        result = resolve_product(
            source.get("Style Number", ""),
            source.get("Product Name", ""),
            source.get("Garment Color", ""),
            master,
        )
        record = source.to_dict()
        record.update(result.as_dict())

        parsed_style = clean_text(source.get("Style Number", ""))
        parsed_product = clean_text(source.get("Product Name", ""))
        parsed_color = clean_text(source.get("Garment Color", ""))

        # Product Master is authoritative when the parser only returned a style code
        # as the description. Otherwise preserve the more specific Shopify description.
        generic_description = not parsed_product or parsed_product.casefold() == parsed_style.casefold()
        if generic_description and clean_text(record.get("Master Product Name", "")):
            record["Product Name"] = clean_text(record.get("Master Product Name", ""))
        else:
            record["Product Name"] = parsed_product or clean_text(record.get("Master Product Name", ""))

        record["Style Number"] = parsed_style or clean_text(record.get("Master Style Number", ""))
        record["Garment Color"] = clean_text(record.get("Resolved Garment Color", "")) or parsed_color

        # Smart custom items can be purchase-ready without becoming permanent master records.
        if not result.matched and clean_text(source.get("Parser Source", "")).casefold() == "smart custom item":
            record["Master Match"] = "Smart Custom"
            record["Vendor"] = clean_text(source.get("Detected Vendor", ""))

        resolved_records.append(record)

    merged = pd.DataFrame(resolved_records)
    merged = apply_routing_overrides(merged)
    for column in [
        "Vendor", "Decoration Type", "Decoration Location", "Decoration Placement Instructions", "Decoration Color", "Product Name",
        "Style Number", "Garment Color", "Size", "Master Match",
        "Parser Source", "Parse Confidence", "Detected Vendor", "Resolver Issue",
        "Product Category", "Requires Size", "Requires Color", "Requires Decoration", "Never Outsource",
        "Decoration Source",
    ]:
        if column not in merged.columns:
            merged[column] = ""
        merged[column] = merged[column].fillna("").map(clean_text)

    merged["Size"] = merged["Size"].map(normalize_size)
    if not merged.empty:
        corrected = merged.apply(
            lambda row: split_embedded_size(row["Garment Color"], row["Size"]),
            axis=1,
            result_type="expand",
        )
        corrected.columns = ["Garment Color", "Size"]
        merged[["Garment Color", "Size"]] = corrected

    # Decoration service lines are excluded from purchasing, but their quantities
    # are valuable routing data. Use them to classify product lines before asking
    # the user to make repetitive style-by-style decisions.
    merged = infer_order_decorations(merged, service_totals_from_raw(raw))
    merged["_Mandatory Note Review"] = _mandatory_note_review_flags(merged)

    rows = []
    for _, row in merged.iterrows():
        reasons = []
        style = clean_text(row.get("Style Number", ""))
        description = clean_text(row.get("Product Name", ""))
        color = clean_text(row.get("Garment Color", ""))
        size = clean_text(row.get("Size", ""))
        vendor = clean_text(row.get("Vendor", ""))
        deco_type = clean_text(row.get("Decoration Type", ""))
        deco_location_raw = clean_text(row.get("Decoration Location", ""))
        placement_instructions = clean_text(row.get("Decoration Placement Instructions", ""))
        deco_color = clean_text(row.get("Decoration Color", ""))
        qty = int(row.get("Quantity", 0))
        customer_identity = (
            clean_text(row.get("Employee Name", ""))
            or clean_text(row.get("Company", ""))
            or clean_text(row.get("Order Number", ""))
        )
        purchase_rules = row_rules({
            "Product Name": description,
            "Style Number": style,
            "Product Category": row.get("Product Category", ""),
            "Requires Size": row.get("Requires Size", ""),
            "Requires Color": row.get("Requires Color", ""),
            "Requires Decoration": row.get("Requires Decoration", ""),
            "Decoration Type": deco_type,
        })
        requires_size = normalize_bool(purchase_rules["Requires Size"], True)
        requires_color = normalize_bool(purchase_rules["Requires Color"], True)
        requires_decoration = normalize_bool(purchase_rules["Requires Decoration"], True)
        deco_location = normalize_decoration_location(
            deco_location_raw, deco_type, requires_decoration=requires_decoration,
            product_name=description, category=row.get("Product Category", ""),
        )
        # When Product Master says color is not required, do not carry a parsed
        # waist/size fragment (for example 58 from 58x34) as a garment color.
        # This also keeps pants/jeans/bibs from generating false review decisions.
        if not requires_color:
            color = ""
            row["Garment Color"] = ""
        if not requires_decoration and not deco_type:
            deco_type = BLANK_DECORATION_LABEL
            row["Decoration Type"] = deco_type
            deco_color = ""
        if is_blank_decoration(deco_type) or is_in_house_decoration(deco_type):
            deco_location = ""
            placement_instructions = ""
            deco_color = ""

        # Only information required to place a vendor order is a blocker.
        if not style:
            reasons.append("Missing product number")
        if not description:
            reasons.append("Missing description")
        if requires_color and not color:
            reasons.append("Missing garment color")
        if requires_size and not size:
            reasons.append("Missing size")
        if not vendor:
            reasons.append("Missing purchase vendor")
        if requires_decoration and not deco_type:
            reasons.append("Missing decoration type")
        if requires_decoration and deco_type and not is_blank_decoration(deco_type) and not is_in_house_decoration(deco_type) and not deco_color:
            reasons.append("Missing decoration color")
        if qty <= 0:
            reasons.append("Quantity must be greater than zero")
        if not customer_identity:
            reasons.append("Missing customer or order name")

        resolver_issue = clean_text(row.get("Resolver Issue", ""))
        if resolver_issue and resolver_issue.casefold() not in {reason.casefold() for reason in reasons}:
            # Resolver warnings only block if they correspond to a required field
            # or a newly synchronized style still needs the user's permanent review.
            if "product master setup required" in resolver_issue.casefold():
                reasons.append("Product Master setup required")
            if "missing garment color" in resolver_issue.casefold() and requires_color and not color:
                reasons.append("Missing garment color")
            elif "decoration color" in resolver_issue.casefold() and deco_type and not is_blank_decoration(deco_type) and not is_in_house_decoration(deco_type) and not deco_color:
                reasons.append("Missing decoration color")

        raw_note = clean_text(row.get("Shopify Order Notes", ""))
        line_note = clean_text(row.get("Shopify Line Notes", ""))
        purchase_instructions = combine_purchase_instructions(raw_note, line_note)
        note_action = note_action_label(purchase_instructions)
        mandatory_note_review = bool(row.get("_Mandatory Note Review", False))
        if mandatory_note_review:
            reasons.append("Customer decision required")

        reason = "; ".join(dict.fromkeys(reasons))
        status = "Needs Review" if reason else "Ready"
        rows.append({
            "Line ID": f"L{len(rows)+1:05d}",
            "Include": "Yes",
            "Do Not Outsource": (
                "Yes" if is_entire_order_outsourced(decoration_fulfillment)
                and (
                    normalize_bool(row.get("Never Outsource", ""), False)
                    or vendor_never_outsource(vendor)
                ) else "No"
            ),
            "Purchase Vendor": vendor,
            "Decoration Type": deco_type,
            "Decoration Location": deco_location,
            "Decoration Placement Instructions": placement_instructions,
            "Decoration Color": deco_color,
            "Product #": style,
            "Description": description,
            "Garment Color": color,
            "Size": size,
            "Quantity": qty,
            "Company": clean_text(row.get("Company", "")),
            "Employee Name": clean_text(row.get("Employee Name", "")),
            "Order Number": clean_text(row.get("Order Number", "")),
            "Shopify Order Notes": raw_note,
            "Purchase Instructions": purchase_instructions,
            "Note Action": note_action,
            "Review Status": status,
            "Review Reason": reason,
            "Product Category": purchase_rules["Product Category"],
            "Requires Size": purchase_rules["Requires Size"],
            "Requires Color": purchase_rules["Requires Color"],
            "Requires Decoration": purchase_rules["Requires Decoration"],
            "Decoration Source": clean_text(row.get("Decoration Source", "")),
            "Decision Key": "",
            "Source": clean_text(row.get("Parser Source", "")) or "Shopify Product",
            "Parse Confidence": clean_text(row.get("Parse Confidence", "")) or "High",
            "Master Match": clean_text(row.get("Master Match", "")) or "No",
            "Original Shopify Line": clean_text(row.get("Original Line Item", "")),
            "Order Date": clean_text(row.get("Order Date", "")),
            "Shopify Line Notes": line_note,
            "Decoration Decision": "",
            "Mandatory Note Review": "Yes" if mandatory_note_review else "No",
        })

    detail = pd.DataFrame(rows, columns=ALL_COLUMNS)
    if not detail.empty:
        detail["Decision Key"] = detail.apply(unresolved_decision_key, axis=1)
    review = detail[detail["Review Status"].eq("Needs Review")].copy() if not detail.empty else pd.DataFrame(columns=ALL_COLUMNS)

    excluded_rows = []
    for _, row in raw.iterrows():
        item = clean_text(row.get("Lineitem name", ""))
        if item and is_decoration_service(item):
            qty = pd.to_numeric(row.get("Lineitem quantity", 0), errors="coerce")
            excluded_rows.append({
                "Order Number": clean_text(row.get("Name", "")),
                "Company": clean_text(row.get("Billing Company", "")),
                "Employee Name": clean_text(row.get("Billing Name", "")),
                "Service / Fee": item,
                "Quantity": 0 if pd.isna(qty) else int(qty),
                "Shopify Order Notes": clean_text(row.get("Notes", "")),
            })
    excluded = pd.DataFrame(
        excluded_rows,
        columns=["Order Number", "Company", "Employee Name", "Service / Fee", "Quantity", "Shopify Order Notes"],
    )
    return detail, review, excluded, raw


def _revalidate_detail(detail: pd.DataFrame) -> pd.DataFrame:
    """Recalculate blockers after carried-forward Review & Edit corrections.

    Product Master changes remain authoritative, while user-entered order-specific
    corrections such as size, color, product number, and include/exclude choices
    survive a Purchase Review regeneration.
    """
    if detail.empty:
        return detail.copy()
    result = detail.copy()
    for index, row in result.iterrows():
        style = clean_text(row.get("Product #", ""))
        description = clean_text(row.get("Description", ""))
        color = clean_text(row.get("Garment Color", ""))
        size = normalize_size(row.get("Size", ""))
        vendor = clean_text(row.get("Purchase Vendor", ""))
        deco_type = clean_text(row.get("Decoration Type", ""))
        deco_location_raw = clean_text(row.get("Decoration Location", ""))
        placement_instructions = clean_text(row.get("Decoration Placement Instructions", ""))
        deco_color = clean_text(row.get("Decoration Color", ""))
        try:
            qty = int(float(row.get("Quantity", 0) or 0))
        except (TypeError, ValueError):
            qty = 0

        purchase_rules = row_rules({
            "Product Name": description,
            "Style Number": style,
            "Product Category": row.get("Product Category", ""),
            "Requires Size": row.get("Requires Size", ""),
            "Requires Color": row.get("Requires Color", ""),
            "Requires Decoration": row.get("Requires Decoration", ""),
            "Decoration Type": deco_type,
        })
        requires_size = normalize_bool(purchase_rules["Requires Size"], True)
        requires_color = normalize_bool(purchase_rules["Requires Color"], True)
        requires_decoration = normalize_bool(purchase_rules["Requires Decoration"], True)
        deco_location = normalize_decoration_location(
            deco_location_raw, deco_type, requires_decoration=requires_decoration,
            product_name=description, category=row.get("Product Category", ""),
        )
        if not requires_color:
            color = ""
        if not requires_decoration and not deco_type:
            deco_type = BLANK_DECORATION_LABEL
            deco_color = ""
        if is_blank_decoration(deco_type) or is_in_house_decoration(deco_type):
            deco_location = ""
            placement_instructions = ""
            deco_color = ""

        reasons: list[str] = []
        if not style:
            reasons.append("Missing product number")
        if not description:
            reasons.append("Missing description")
        if requires_color and not color:
            reasons.append("Missing garment color")
        if requires_size and not size:
            reasons.append("Missing size")
        if not vendor:
            reasons.append("Missing purchase vendor")
        if requires_decoration and not deco_type:
            reasons.append("Missing decoration type")
        if requires_decoration and deco_type and not is_blank_decoration(deco_type) and not is_in_house_decoration(deco_type) and not deco_color:
            reasons.append("Missing decoration color")
        if qty <= 0:
            reasons.append("Quantity must be greater than zero")
        customer_identity = (
            clean_text(row.get("Employee Name", ""))
            or clean_text(row.get("Company", ""))
            or clean_text(row.get("Order Number", ""))
        )
        if not customer_identity:
            reasons.append("Missing customer or order name")

        raw_note = clean_text(row.get("Shopify Order Notes", ""))
        line_note = clean_text(row.get("Shopify Line Notes", ""))
        purchase_instructions = clean_text(row.get("Purchase Instructions", "")) or combine_purchase_instructions(raw_note, line_note)
        mandatory_note_review = normalize_bool(row.get("Mandatory Note Review", ""), False)
        if not mandatory_note_review:
            mandatory_note_review = note_requires_review_for_line(
                raw_note,
                line_note,
                " ".join(filter(None, [description, clean_text(row.get("Original Shopify Line", "")), clean_text(row.get("Product Category", "")), style])),
            )
        if mandatory_note_review:
            reasons.append("Customer decision required")

        reason = "; ".join(dict.fromkeys(reasons))
        result.at[index, "Product #"] = style
        result.at[index, "Description"] = description
        result.at[index, "Garment Color"] = color
        result.at[index, "Size"] = size
        result.at[index, "Purchase Vendor"] = vendor
        result.at[index, "Decoration Type"] = deco_type
        result.at[index, "Decoration Location"] = deco_location
        result.at[index, "Decoration Placement Instructions"] = placement_instructions
        result.at[index, "Decoration Color"] = deco_color
        result.at[index, "Quantity"] = qty
        result.at[index, "Purchase Instructions"] = purchase_instructions
        result.at[index, "Note Action"] = note_action_label(purchase_instructions)
        result.at[index, "Shopify Line Notes"] = line_note
        result.at[index, "Mandatory Note Review"] = "Yes" if mandatory_note_review else "No"
        result.at[index, "Product Category"] = purchase_rules["Product Category"]
        result.at[index, "Requires Size"] = purchase_rules["Requires Size"]
        result.at[index, "Requires Color"] = purchase_rules["Requires Color"]
        result.at[index, "Requires Decoration"] = purchase_rules["Requires Decoration"]
        result.at[index, "Review Reason"] = reason
        result.at[index, "Review Status"] = "Needs Review" if reason else "Ready"

    result["Decision Key"] = result.apply(unresolved_decision_key, axis=1)
    return result


def _apply_previous_review_edits(detail: pd.DataFrame, previous_review_path: Path | None) -> pd.DataFrame:
    """Carry visible Review & Edit corrections into a regenerated workbook.

    Completed order-specific decisions must stay completed after Product Master
    changes trigger a regeneration.  A prior Ready status is preserved only when
    the carried values satisfy every hard purchasing requirement.  The one manual
    exception is ``Customer decision required``: saving that row is the user's
    explicit confirmation, so it remains Ready across regeneration.
    """
    if detail.empty or not previous_review_path:
        return detail
    previous = Path(previous_review_path).expanduser()
    if not previous.exists():
        return detail
    try:
        from modules.xlsx_reader import read_table
        edits = read_table(previous, "Review & Edit", {"Line ID"})
    except Exception:
        return detail
    if not edits:
        return detail

    result = detail.copy()
    index_by_line = {
        clean_text(value): index
        for index, value in result["Line ID"].items()
        if clean_text(value)
    }
    carry_fields = {
        "Include", "Do Not Outsource", "Product #", "Description", "Garment Color", "Size", "Quantity",
        "Company", "Employee Name", "Order Number", "Decoration Decision",
        "Decoration Location", "Decoration Placement Instructions",
    }
    permanent_override_fields = {"Purchase Vendor", "Decoration Type", "Decoration Location", "Decoration Placement Instructions", "Decoration Color"}
    previously_ready: set[str] = set()
    for edit in edits:
        line_id = clean_text(edit.get("Line ID", ""))
        index = index_by_line.get(line_id)
        if index is None:
            continue
        was_ready = clean_text(edit.get("Review Status", "")).casefold() == "ready"
        if was_ready:
            old_source_signature = "|".join([
                clean_text(edit.get("Original Shopify Line", "")),
                clean_text(edit.get("Shopify Order Notes", "")),
                clean_text(edit.get("Shopify Line Notes", "")),
            ])
            new_source_signature = "|".join([
                clean_text(result.at[index, "Original Shopify Line"] if "Original Shopify Line" in result.columns else ""),
                clean_text(result.at[index, "Shopify Order Notes"] if "Shopify Order Notes" in result.columns else ""),
                clean_text(result.at[index, "Shopify Line Notes"] if "Shopify Line Notes" in result.columns else ""),
            ])
            if old_source_signature == new_source_signature:
                previously_ready.add(line_id)
            else:
                was_ready = False
        # Product Master remains authoritative while a decision is unresolved.
        # Vendor/decoration values are carried only after the user deliberately
        # completed the order-specific row and marked it Ready.
        fields = set(carry_fields)
        if was_ready:
            fields.update(permanent_override_fields)
        for field in fields:
            if field not in edit:
                continue
            value = edit.get(field, "")
            if field == "Quantity":
                try:
                    value = int(float(value or 0))
                except (TypeError, ValueError):
                    value = 0
            if field in {"Include", "Do Not Outsource", "Quantity"} or clean_text(value):
                result.at[index, field] = value

    result = _revalidate_detail(result)

    # Revalidation intentionally recreates hard blockers. Preserve a user's
    # completed manual decision only when no hard blocker remains.
    manual_only_reason = "customer decision required"
    for line_id in previously_ready:
        index = index_by_line.get(line_id)
        if index is None:
            continue
        include = clean_text(result.at[index, "Include"]).casefold()
        if include not in {"yes", "y", "true", "1", "include"}:
            result.at[index, "Review Status"] = "Ready"
            continue
        reasons = [
            clean_text(reason).casefold()
            for reason in clean_text(result.at[index, "Review Reason"]).split(";")
            if clean_text(reason)
        ]
        hard_reasons = [reason for reason in reasons if reason != manual_only_reason]
        mandatory_note = normalize_bool(result.at[index, "Mandatory Note Review"], False)
        decision = clean_text(result.at[index, "Decoration Decision"]).casefold()
        valid_decisions = {"follow note as written", "keep product master default", "no decoration", "embroidery", "screen print", "sew on patch", "hemming / alteration", "do not outsource — ship to orchid"}
        note_confirmed = not mandatory_note or decision in valid_decisions
        if not hard_reasons and note_confirmed:
            result.at[index, "Review Status"] = "Ready"

    return result

def _formats(workbook):
    return {
        "title": workbook.add_format({
            "bold": True, "font_size": 16, "font_color": PURPLE_DARK,
            "bg_color": PURPLE_LIGHT, "align": "left", "valign": "vcenter",
        }),
        "header": workbook.add_format({
            "bold": True, "font_color": TEXT, "bg_color": "#D8C8EA",
            "border": 1, "border_color": BORDER, "align": "center",
            "valign": "vcenter", "text_wrap": True,
        }),
        "section": workbook.add_format({
            "bold": True, "font_size": 13, "font_color": PURPLE_DARK, "bg_color": PURPLE_PALE,
            "border": 1, "border_color": BORDER,
        }),
        "label": workbook.add_format({
            "bold": True, "font_color": TEXT, "bg_color": PURPLE_PALE,
            "border": 1, "border_color": BORDER,
        }),
        "value": workbook.add_format({"border": 1, "border_color": BORDER}),
        "body": workbook.add_format({
            "font_color": TEXT, "border": 1, "border_color": BORDER,
            "valign": "top",
        }),
        "text": workbook.add_format({
            "font_color": TEXT, "border": 1, "border_color": BORDER,
            "valign": "top", "align": "left", "num_format": "@",
        }),
        "body_wrap": workbook.add_format({
            "font_color": TEXT, "border": 1, "border_color": BORDER,
            "valign": "top", "text_wrap": True,
        }),
        "qty": workbook.add_format({
            "font_color": TEXT, "border": 1, "border_color": BORDER,
            "align": "center", "num_format": "0",
        }),
        "currency": workbook.add_format({
            "font_color": TEXT, "border": 1, "border_color": BORDER,
            "align": "right", "num_format": "$#,##0.00",
        }),
        "instruction": workbook.add_format({
            "font_color": MUTED, "bg_color": BLUE_PALE, "text_wrap": True,
            "valign": "top", "border": 1, "border_color": BORDER,
        }),
        "warning": workbook.add_format({
            "font_color": TEXT, "bg_color": YELLOW, "text_wrap": True,
            "border": 1, "border_color": BORDER,
        }),
    }


def _write_dataframe_sheet(workbook, sheet_name, title, frame, formats, widths=None, note=None):
    ws = workbook.add_worksheet(_unique_sheet_name(workbook, sheet_name))
    ws.hide_gridlines(2)
    ws.set_zoom(88)
    ws.freeze_panes(3, 0)
    columns = list(frame.columns)
    last_col = max(len(columns) - 1, 0)
    ws.merge_range(0, 0, 0, max(last_col, 5), title, formats["title"])
    ws.set_row(0, 32)
    if note:
        ws.merge_range(1, 0, 1, max(last_col, 5), note, formats["instruction"])
        ws.set_row(1, 34)
    for col, name in enumerate(columns):
        ws.write(2, col, name, formats["header"])
    ws.set_row(2, 32)
    for r, values in enumerate(frame.itertuples(index=False, name=None), start=3):
        for c, value in enumerate(values):
            if columns[c] in {"Quantity", "Total Pieces", "Pieces", "Garments"}:
                fmt = formats["qty"]
            elif columns[c] in {"Order Total", "Discount", "Total"}:
                fmt = formats["currency"]
            elif columns[c] in {"Product #", "Line ID", "Order Number", "Order Number(s)", "PO Number"}:
                fmt = formats["text"]
            else:
                fmt = (
                    formats["body_wrap"] if columns[c] in {
                        "Shopify Order Notes", "Purchase Instructions", "Review Reason",
                        "Original Shopify Line", "Description",
                    } else formats["body"]
                )
            ws.write(r, c, value, fmt)
        ws.set_row(r, 28 if any(clean_text(v) for i, v in enumerate(values) if columns[i] in {"Shopify Order Notes", "Purchase Instructions", "Review Reason"}) else 20)

    end_row = max(3, len(frame) + 2)
    if columns:
        ws.autofilter(2, 0, end_row, last_col)
    default_widths = {
        "Line ID": 11, "Include": 10, "Purchase Vendor": 18, "Decoration Type": 17,
        "Decoration Location": 20, "Decoration Placement Instructions": 30, "Decoration Color": 17, "Product #": 14, "Description": 36,
        "Garment Color": 20, "Size": 12, "Quantity": 10,
        "Company": 25, "Employee Name": 24, "Order Number": 13,
        "Shopify Order Notes": 44, "Purchase Instructions": 30,
        "Review Status": 15, "Review Reason": 38, "Note Action": 16,
        "Source": 18, "Parse Confidence": 16, "Master Match": 16,
        "Original Shopify Line": 45, "Order Date": 22,
        "Shopify Line Notes": 40, "Decoration Decision": 28, "Mandatory Note Review": 18,
        "Product Category": 22, "Requires Size": 14, "Requires Color": 14, "Requires Decoration": 18,
        "Vendor": 22, "Report": 24, "Total Pieces": 14, "PO Number": 22, "Action": 20,
        "Company / Department": 28, "Order Number(s)": 18, "Order Total": 16, "Discount": 14, "Total": 16,
    }
    for c, name in enumerate(columns):
        ws.set_column(c, c, (widths or {}).get(name, default_widths.get(name, 18)))
    return ws


def _add_edit_controls(workbook, ws, row_count, settings_rows):
    if row_count < 1:
        row_count = 50
    start = 3
    end = start + row_count - 1
    col = {name: index for index, name in enumerate(ALL_COLUMNS)}
    ws.data_validation(start, col["Include"], end, col["Include"], {
        "validate": "list", "source": ["Yes", "No"],
    })
    if settings_rows["vendors"]:
        ws.data_validation(start, col["Purchase Vendor"], end, col["Purchase Vendor"], {
            "validate": "list", "source": "=Settings!$A$2:$A$%d" % (len(settings_rows["vendors"]) + 1),
        })
    ws.data_validation(start, col["Decoration Type"], end, col["Decoration Type"], {
        "validate": "list", "source": ["Embroidery", "Screen Print", SEW_ON_PATCH_LABEL, HEMMING_ALTERATION_LABEL, BLANK_DECORATION_LABEL],
    })
    ws.data_validation(start, col["Decoration Location"], end, col["Decoration Location"], {
        "validate": "list", "source": DECORATION_LOCATIONS,
        "error_type": "warning", "error_title": "Custom Location",
        "error_message": "Choose a standard location or type the exact custom location.",
    })
    ws.data_validation(start, col["Review Status"], end, col["Review Status"], {
        "validate": "list", "source": ["Ready", "Needs Review"],
    })
    ws.data_validation(start, col["Decoration Decision"], end, col["Decoration Decision"], {
        "validate": "list",
        "source": ["Follow Note as Written", "Keep Product Master Default", "No Decoration", "Embroidery", "Screen Print", "Sew On Patch", "Hemming / Alteration", "Do Not Outsource — Ship to Orchid"],
    })
    note_format = workbook.add_format({"bg_color": YELLOW, "text_wrap": True})
    review_format = workbook.add_format({"bg_color": RED_PALE, "font_color": "#9C0006"})
    ready_format = workbook.add_format({"bg_color": GREEN_PALE, "font_color": "#006100"})
    note_column = "Purchase Instructions" if "Purchase Instructions" in col else "Shopify Order Notes"
    note_letter = xlsxwriter.utility.xl_col_to_name(col[note_column])
    ws.conditional_format(start, col[note_column], end, col[note_column], {
        "type": "formula", "criteria": f"=LEN({note_letter}{start + 1})>0", "format": note_format,
    })
    ws.conditional_format(start, col["Review Status"], end, col["Review Status"], {
        "type": "text", "criteria": "containing", "value": "Needs Review",
        "format": review_format,
    })
    ws.conditional_format(start, col["Review Status"], end, col["Review Status"], {
        "type": "text", "criteria": "containing", "value": "Ready",
        "format": ready_format,
    })

def _build_employee_totals(raw: pd.DataFrame) -> pd.DataFrame:
    columns = ["Company / Department", "Employee Name", "Order Number(s)", "Order Total", "Discount", "Total"]
    if raw.empty or "Name" not in raw.columns:
        return pd.DataFrame(columns=columns)

    source = raw.copy()
    source["Name"] = source["Name"].replace("", pd.NA).ffill().fillna("")
    for column in ["Billing Company", "Billing Name"]:
        if column not in source.columns:
            source[column] = ""
        source[column] = (
            source[column].replace("", pd.NA)
            .groupby(source["Name"], dropna=False)
            .transform(lambda values: values.ffill().bfill())
            .fillna("")
        )

    rows = []
    for order_number, group in source.groupby("Name", dropna=False, sort=False):
        order_number = clean_text(order_number)
        if not order_number:
            continue
        company = clean_text(group["Billing Company"].iloc[0])
        employee = clean_text(group["Billing Name"].iloc[0])
        total = 0.0
        for total_column in ["Total", "Subtotal"]:
            if total_column not in group.columns:
                continue
            values = pd.to_numeric(group[total_column], errors="coerce").dropna()
            if not values.empty:
                total = float(values.iloc[0])
                break
        rows.append({
            "Company / Department": company,
            "Employee Name": employee,
            "Order Number(s)": order_number,
            "Order Total": total,
            "Discount": 0.0,
            "Total": total,
        })
    employee_orders = pd.DataFrame(rows, columns=columns)
    if employee_orders.empty:
        return employee_orders
    grouped_rows = []
    for (company, employee), group in employee_orders.groupby(
        ["Company / Department", "Employee Name"], dropna=False, sort=False
    ):
        order_numbers = ", ".join(dict.fromkeys(
            clean_text(value) for value in group["Order Number(s)"] if clean_text(value)
        ))
        grouped_rows.append({
            "Company / Department": clean_text(company),
            "Employee Name": clean_text(employee),
            "Order Number(s)": order_numbers,
            "Order Total": float(group["Order Total"].sum()),
            "Discount": 0.0,
            "Total": float(group["Order Total"].sum()),
        })
    return pd.DataFrame(grouped_rows, columns=columns)


def generate_review_workbook(
    shopify_csv_path: Path,
    product_master_path: Path,
    reports_root: Path,
    report_mode: str = GENERAL_SALES_PERIOD,
    event_name: str = "",
    decoration_fulfillment: str = STANDARD_ORCHID_WORKFLOW,
    regenerate_helper: Path | None = None,
    previous_review_path: Path | None = None,
) -> dict:
    report_mode = normalize_report_mode(report_mode)
    decoration_fulfillment = normalize_decoration_fulfillment(decoration_fulfillment)
    event_name = clean_text(event_name) if report_mode == UNIFORM_SIZING_EVENT else ""
    product_master_path = Path(product_master_path).expanduser().resolve()
    master_signature = product_master_signature(product_master_path)
    # Always reload the active Product Master from disk immediately before building.
    detail, review, excluded, raw = _build_review_data(
        Path(shopify_csv_path), product_master_path, decoration_fulfillment
    )
    now = datetime.now()
    stamp = now.strftime("%Y%m%d_%H%M%S_%f")
    output_dir = Path(reports_root) / "Review Workbooks"
    output_dir.mkdir(parents=True, exist_ok=True)
    event_part = f"__{safe_filename(event_name)}" if event_name else ""
    output_path = output_dir / f"Orchid_Purchase_Review{event_part}__{stamp}.xlsx"
    regeneration_request_path = output_path.with_suffix(".orchidregen")

    detail = _apply_previous_review_edits(detail, previous_review_path)

    master = load_extended_master(Path(product_master_path))
    vendor_registry = dict(PREFERRED_VENDOR_CASE)
    if "Vendor" in master.columns:
        master["Vendor"] = master["Vendor"].map(lambda value: _canonical_vendor(value, vendor_registry))
    if not detail.empty and "Purchase Vendor" in detail.columns:
        detail["Purchase Vendor"] = detail["Purchase Vendor"].map(
            lambda value: _canonical_vendor(value, vendor_registry)
        )
    review = detail[detail["Review Status"].eq("Needs Review")].copy() if not detail.empty else pd.DataFrame(columns=ALL_COLUMNS)

    vendor_values = {_canonical_vendor(v, vendor_registry) for v in master.get("Vendor", []) if clean_text(v)}
    if not detail.empty:
        vendor_values.update(_canonical_vendor(v, vendor_registry) for v in detail.get("Purchase Vendor", []) if clean_text(v))
    vendors = sorted({value for value in vendor_values if value}, key=str.casefold)
    route_frame = pd.DataFrame(columns=["Vendor", "Report Type", "Pieces", "PO Number", "Status"])
    if not detail.empty:
        route_source = detail[detail["Include"].str.casefold().eq("yes")].copy()
        if "Do Not Outsource" in route_source.columns:
            route_source = route_source[
                ~route_source["Do Not Outsource"].map(clean_text).str.casefold().isin({"yes", "y", "true", "1"})
            ].copy()
        if not route_source.empty:
            route_source["Vendor"] = route_source["Purchase Vendor"].map(lambda value: _canonical_vendor(value, vendor_registry)).replace("", "Unassigned")
            deco = route_source["Decoration Type"].map(clean_text)
            # One purchase order route per vendor. Embroidery, screen printing,
            # and blank garments are separated into sections inside the PDF.
            route_source["Report Type"] = "Combined Vendor Order"
            route_source["_needs_review"] = (
                route_source["Review Status"].map(clean_text).str.casefold().eq("needs review")
                | deco.eq("")
            )
            # Missing-vendor lines belong in Review & Edit, not in a fake vendor
            # purchase order. They will join the correct route after the user
            # assigns a vendor and Orchid rereads the saved workbook.
            route_source = route_source[~route_source["Vendor"].eq("Unassigned")].copy()
            route_frame = (
                route_source.groupby(["Vendor", "Report Type"], dropna=False, as_index=False)
                .agg(Pieces=("Quantity", "sum"), _needs_review=("_needs_review", "max"))
            )
            route_frame["PO Number"] = ""
            route_frame["Status"] = route_frame["_needs_review"].map(lambda value: "Needs Review" if value else "Ready")
            route_frame = route_frame[["Vendor", "Report Type", "Pieces", "PO Number", "Status"]]
            route_frame = route_frame.sort_values(["Vendor", "Report Type"], key=lambda col: col.str.casefold()).reset_index(drop=True)
            def _abbr(value: str, max_len: int = 4) -> str:
                words = re.findall(r"[A-Za-z0-9]+", clean_text(value))
                if not words:
                    return "PO"
                if len(words) > 1:
                    return "".join(word[0] for word in words).upper()[:max_len]
                return words[0].upper()[:max_len]
            event_code = _abbr(event_name or "General Sales", 4)
            for idx in route_frame.index:
                vendor_code = _abbr(route_frame.at[idx, "Vendor"], 4)
                route_frame.at[idx, "PO Number"] = (
                    "" if route_frame.at[idx, "Vendor"] == "Needs Vendor Assignment"
                    else f"{event_code}-{vendor_code}"
                )


    if detail.empty:
        permanent_mask = pd.Series(dtype=bool)
        exception_mask = pd.Series(dtype=bool)
        detail["Fix In"] = pd.Series(dtype=str)
        detail["Action Required"] = pd.Series(dtype=str)
        detail["Permanent Review Reason"] = pd.Series(dtype=str)
        detail["Event Review Reason"] = pd.Series(dtype=str)
    else:
        buckets = detail.apply(_issue_buckets, axis=1)
        detail["Permanent Review Reason"] = buckets.map(lambda value: "; ".join(value[0]))
        detail["Event Review Reason"] = buckets.map(lambda value: "; ".join(value[1]))
        detail["Fix In"] = detail.apply(_fix_location, axis=1)
        detail["Action Required"] = detail.apply(_action_required, axis=1)
        included_mask = ~detail["Include"].map(clean_text).str.casefold().eq("no")
        # Permanent setup is independent of current-order corrections. Unknown
        # style-like products are setup items even when all order fields happen to
        # be present, so they can never bypass Product Master readiness.
        permanent_mask = detail["Permanent Review Reason"].map(clean_text).ne("") & included_mask
        exception_mask = detail["Event Review Reason"].map(clean_text).ne("") & included_mask

    workbook = xlsxwriter.Workbook(str(output_path))
    workbook.set_properties({
        "title": "Orchid Purchase Manager Purchase Review",
        "subject": "Review and approve vendor purchase order lines",
        "author": "Orchid Uniforms & Apparel",
        "company": "Orchid Uniforms & Apparel",
    })
    fmt = _formats(workbook)

    start = workbook.add_worksheet("Start Here")
    start.hide_gridlines(2)
    start.set_column("A:A", 4)
    start.set_column("B:B", 34)
    start.set_column("C:C", 92)
    start.merge_range("B2:C2", "ORCHID PURCHASE MANAGER", fmt["title"])
    start.set_row(1, 36)
    start.write(2, 1, "Purchase Order Mode", fmt["label"])
    start.write(2, 2, report_mode, fmt["value"])
    start.set_row(2, 24)
    if event_name:
        start.write(3, 1, "Event Name", fmt["label"])
        start.write(3, 2, event_name, fmt["value"])
        start.set_row(3, 24)
    instructions = [
        ("1", "Check Dashboard. It shows what is ready and what still blocks a purchase order."),
        ("2", "If Product Master Needs Attention, use the Product Master list to update those products in Orchid Purchase Manager, then recreate Purchase Review."),
        ("3", "Use Review & Edit only for current-order corrections, actionable customer notes, and one-time/manual items."),
        ("4", "Financial allowance and discount text is removed from the visible Purchase Instructions field while the original Shopify note remains preserved in hidden source data."),
        ("5", "Enter or edit each PO number on Dashboard."),
        ("6", "For Uniform Sizing Events, review Employee Totals and enter any discount."),
        ("7", "Save this workbook, then choose Create Final Purchase Orders in Orchid Purchase Manager."),
    ]
    instruction_start = 5 if event_name else 4
    for i, (number, instruction_text) in enumerate(instructions, start=instruction_start):
        start.write(i, 1, number, fmt["section"])
        start.write(i, 2, instruction_text, fmt["instruction"])
        start.set_row(i, 34)
    important_row = instruction_start + len(instructions) + 2
    start.write(important_row, 1, "Purchase-line rule", fmt["section"])
    start.write(important_row, 2, "A line is Ready when Orchid has the vendor, product number, description, color, size, quantity, customer identity, and valid decoration routing. Old parser warnings do not keep a complete line in review.", fmt["warning"])
    start.set_row(important_row, 48)

    summary = workbook.add_worksheet("Dashboard")
    summary.activate()
    start.hide()
    summary.hide_gridlines(2)
    summary.set_landscape()
    summary.set_zoom(95)
    summary.set_column("A:B", 17)
    summary.set_column("C:D", 17)
    summary.set_column("E:F", 17)
    summary.set_column("G:H", 17)
    summary.set_column("I:J", 17)

    dashboard_title = workbook.add_format({
        "bold": True, "font_size": 25, "font_color": PURPLE_DARK,
        "bg_color": PURPLE_LIGHT, "align": "left", "valign": "vcenter",
        "left": 1, "right": 1, "top": 1, "bottom": 1, "border_color": BORDER,
    })
    context_format = workbook.add_format({
        "bold": True, "font_size": 13, "font_color": TEXT,
        "bg_color": PURPLE_PALE, "align": "left", "valign": "vcenter",
        "border": 1, "border_color": BORDER,
    })
    status_ready = workbook.add_format({
        "bold": True, "font_size": 22, "font_color": "#166534",
        "bg_color": GREEN_PALE, "align": "center", "valign": "vcenter",
        "border": 1, "border_color": "#86B98A",
    })
    status_blocked = workbook.add_format({
        "bold": True, "font_size": 22, "font_color": "#9C0006",
        "bg_color": RED_PALE, "align": "center", "valign": "vcenter",
        "border": 1, "border_color": "#D99A96",
    })
    status_sub_ready = workbook.add_format({
        "font_size": 12, "font_color": "#166534", "bg_color": GREEN_PALE,
        "align": "center", "valign": "vcenter", "text_wrap": True,
        "border": 1, "border_color": "#86B98A",
    })
    status_sub_blocked = workbook.add_format({
        "font_size": 12, "font_color": "#7F1D1D", "bg_color": RED_PALE,
        "align": "center", "valign": "vcenter", "text_wrap": True,
        "border": 1, "border_color": "#D99A96",
    })
    card_title = workbook.add_format({
        "bold": True, "font_size": 11, "font_color": MUTED,
        "bg_color": PURPLE_PALE, "align": "center", "valign": "vcenter",
        "border": 1, "border_color": BORDER,
    })
    card_value_good = workbook.add_format({
        "bold": True, "font_size": 18, "font_color": "#166534",
        "bg_color": WHITE, "align": "center", "valign": "vcenter",
        "border": 1, "border_color": BORDER,
    })
    card_value_warn = workbook.add_format({
        "bold": True, "font_size": 18, "font_color": "#9C0006",
        "bg_color": WHITE, "align": "center", "valign": "vcenter",
        "border": 1, "border_color": BORDER,
    })
    card_caption = workbook.add_format({
        "font_size": 10, "font_color": MUTED, "bg_color": WHITE,
        "align": "center", "valign": "vcenter", "text_wrap": True,
        "border": 1, "border_color": BORDER,
    })
    action_button = workbook.add_format({
        "bold": True, "font_size": 13, "font_color": WHITE, "bg_color": PURPLE,
        "align": "center", "valign": "vcenter", "border": 1,
        "border_color": PURPLE_DARK, "underline": False, "text_wrap": True,
    })
    action_button_good = workbook.add_format({
        "bold": True, "font_size": 12, "font_color": "#166534", "bg_color": GREEN_PALE,
        "align": "center", "valign": "vcenter", "border": 1,
        "border_color": "#86B98A", "underline": False, "text_wrap": True,
    })
    action_button_disabled = workbook.add_format({
        "bold": True, "font_size": 12, "font_color": "#8A8293", "bg_color": "#EEEAF2",
        "align": "center", "valign": "vcenter", "border": 1,
        "border_color": BORDER, "underline": False, "text_wrap": True,
    })
    action_button_secondary = workbook.add_format({
        "bold": True, "font_size": 12, "font_color": PURPLE_DARK, "bg_color": PURPLE_LIGHT,
        "align": "center", "valign": "vcenter", "border": 1,
        "border_color": "#BFA9DD", "underline": False, "text_wrap": True,
    })
    action_button_warn = workbook.add_format({
        "bold": True, "font_size": 12, "font_color": "#7A3E00", "bg_color": "#FFF4E8",
        "align": "center", "valign": "vcenter", "border": 1,
        "border_color": "#E6B77A", "underline": False, "text_wrap": True,
    })
    next_step_text = workbook.add_format({
        "font_size": 13, "font_color": TEXT, "bg_color": WHITE,
        "align": "left", "valign": "vcenter", "text_wrap": True,
        "border": 1, "border_color": BORDER,
    })
    po_summary = workbook.add_format({
        "bold": True, "font_color": PURPLE_DARK, "bg_color": PURPLE_PALE,
        "align": "right", "valign": "vcenter", "border": 1,
        "border_color": BORDER,
    })

    event_company = ""
    if not detail.empty:
        companies = [c for c in detail["Company"].dropna().map(clean_text).unique() if c]
        event_company = companies[0] if len(companies) == 1 else "Multiple Companies / Orders"
    ready_reports = int((route_frame["Status"] == "Ready").sum()) if not route_frame.empty else 0
    review_reports = int((route_frame["Status"] == "Needs Review").sum()) if not route_frame.empty else 0
    ready_lines = int(detail["Review Status"].eq("Ready").sum()) if not detail.empty else 0
    review_lines_count = int(detail["Review Status"].eq("Needs Review").sum()) if not detail.empty else 0
    total_lines = len(detail)
    permanent_product_count = (
        detail.loc[permanent_mask].apply(_product_master_decision_key, axis=1).nunique()
        if not detail.empty and permanent_mask.any() else 0
    )
    exception_count = count_unique_decisions(detail.loc[exception_mask]) if not detail.empty else 0
    # The visible work queues are the source of truth. Informational metadata
    # in hidden source rows must not keep the Dashboard blocked.
    is_ready = permanent_product_count == 0 and exception_count == 0 and review_reports == 0
    review_complete_initial = exception_count == 0 and review_reports == 0

    # Hidden live counters keep the visible Dashboard uncluttered while allowing
    # the status cards to update as Review & Edit rows are changed to Ready.
    summary.write_formula("L1", "=MAX(0,COUNTA('Product Master'!$A:$A)-3)", None, permanent_product_count)
    summary.write_formula(
        "M1",
        """=SUMPRODUCT(IFERROR((('Review & Edit'!$B$4:$B$503<>"Ready")*('Review & Edit'!$V$4:$V$503<>"")*('Review & Edit'!$U$4:$U$503<>"No"))/COUNTIFS('Review & Edit'!$V$4:$V$503,'Review & Edit'!$V$4:$V$503,'Review & Edit'!$B$4:$B$503,"<>Ready",'Review & Edit'!$U$4:$U$503,"<>No"),0))""",
        None,
        exception_count,
    )
    summary.write_number("N1", total_lines)
    summary.set_column("L:O", None, None, {"hidden": True})

    summary.merge_range("A1:J2", "ORCHID PURCHASE MANAGER", dashboard_title)
    summary.set_row(0, 28)
    summary.set_row(1, 28)
    summary.merge_range("A3:E3", f"{event_name or event_company or 'General Sales Period'}", context_format)
    summary.merge_range("F3:J3", f"{report_mode}  |  Generated {now.strftime('%Y-%m-%d %I:%M %p')}", fmt["instruction"])

    summary.merge_range("A5:J6", "", status_ready if is_ready else status_blocked)
    summary.write_formula(
        "A5",
        '=IF(AND($L$1=0,$M$1=0,$O$1=0),"READY TO GENERATE PURCHASE ORDERS","PURCHASE REVIEW IN PROGRESS")',
        status_ready if is_ready else status_blocked,
        "READY TO GENERATE PURCHASE ORDERS" if is_ready else "PURCHASE REVIEW IN PROGRESS",
    )
    summary.set_row(4, 34)
    summary.set_row(5, 34)
    summary.merge_range("A7:J7", "", status_sub_ready if is_ready else status_sub_blocked)
    summary.write_formula(
        "A7",
        '=IF(AND($L$1=0,$M$1=0,$O$1=0),"All required purchasing information is complete. Confirm PO numbers below, then generate purchase orders.",IF($L$1+$M$1>0,$L$1+$M$1&" item(s) still need attention. ","")&IF($O$1>0,$O$1&" purchase order(s) remain blocked. ","")&"Use the Next Step section below to finish the review.")',
        status_sub_ready if is_ready else status_sub_blocked,
        (
            "All required purchasing information is complete. Confirm PO numbers below, then generate purchase orders."
            if is_ready else
            (
                (f"{permanent_product_count + exception_count} item(s) still need attention. "
                 if permanent_product_count + exception_count else "")
                + (f"{review_reports} purchase order(s) remain blocked. " if review_reports else "")
                + "Use the Next Step section below to finish the review."
            )
        ),
    )
    summary.set_row(6, 30)

    # Four-stage workflow path. It updates as the Product Master and Review & Edit
    # counters change, while Excel remains the trusted detailed workspace.
    stage_complete = workbook.add_format({
        "bold": True, "font_size": 11, "font_color": "#166534",
        "bg_color": GREEN_PALE, "align": "center", "valign": "vcenter",
        "border": 1, "border_color": "#86B98A",
    })
    stage_active = workbook.add_format({
        "bold": True, "font_size": 12, "font_color": PURPLE_DARK,
        "bg_color": PURPLE_LIGHT, "align": "center", "valign": "vcenter",
        "border": 1, "border_color": PURPLE,
    })
    stage_attention = workbook.add_format({
        "bold": True, "font_size": 11, "font_color": "#9A4D00",
        "bg_color": "#FFF4E8", "align": "center", "valign": "vcenter",
        "border": 1, "border_color": "#E6B77A",
    })
    stage_pending = workbook.add_format({
        "bold": True, "font_size": 11, "font_color": MUTED,
        "bg_color": "#F0EDF4", "align": "center", "valign": "vcenter",
        "border": 1, "border_color": BORDER,
    })
    summary.merge_range("A8:B8", "✓ 1  IMPORT ORDERS", stage_complete)
    summary.merge_range("C8:E8", "", stage_attention if permanent_product_count else stage_complete)
    summary.write_formula(
        "C8", '=IF($L$1=0,"✓ 2  PRODUCT MASTER","▶ 2  PRODUCT MASTER")',
        stage_attention if permanent_product_count else stage_complete,
        "▶ 2  PRODUCT MASTER" if permanent_product_count else "✓ 2  PRODUCT MASTER",
    )
    summary.merge_range("F8:G8", "", stage_pending if permanent_product_count else (stage_attention if exception_count or review_reports else stage_complete))
    summary.write_formula(
        "F8", '=IF($L$1>0,"3  PURCHASE REVIEW",IF(OR($M$1>0,$O$1>0),"▶ 3  PURCHASE REVIEW","✓ 3  PURCHASE REVIEW"))',
        stage_pending if permanent_product_count else (stage_attention if exception_count or review_reports else stage_complete),
        ("3  PURCHASE REVIEW" if permanent_product_count else
         ("▶ 3  PURCHASE REVIEW" if exception_count or review_reports else "✓ 3  PURCHASE REVIEW")),
    )
    summary.merge_range("H8:J8", "", stage_active if is_ready else stage_pending)
    summary.write_formula(
        "H8", '=IF(AND($L$1=0,$M$1=0,$O$1=0),"▶ 4  PURCHASE ORDERS","4  PURCHASE ORDERS")',
        stage_active if is_ready else stage_pending,
        "▶ 4  PURCHASE ORDERS" if is_ready else "4  PURCHASE ORDERS",
    )
    summary.set_row(7, 30)
    summary.conditional_format("C8:E8", {"type": "formula", "criteria": "=$L$1=0", "format": stage_complete})
    summary.conditional_format("C8:E8", {"type": "formula", "criteria": "=$L$1>0", "format": stage_attention})
    summary.conditional_format("F8:G8", {"type": "formula", "criteria": "=$L$1>0", "format": stage_pending})
    summary.conditional_format("F8:G8", {"type": "formula", "criteria": "=AND($L$1=0,OR($M$1>0,$O$1>0))", "format": stage_attention})
    summary.conditional_format("F8:G8", {"type": "formula", "criteria": "=AND($L$1=0,$M$1=0,$O$1=0)", "format": stage_complete})
    summary.conditional_format("H8:J8", {"type": "formula", "criteria": "=AND($L$1=0,$M$1=0,$O$1=0)", "format": stage_active})
    summary.conditional_format("H8:J8", {"type": "formula", "criteria": "=OR($L$1>0,$M$1>0,$O$1>0)", "format": stage_pending})

    # Status cards summarize the three working areas. Navigation is intentionally
    # kept in one centered Next Step control below, avoiding duplicate buttons.
    summary.merge_range("A9:C9", "PRODUCT MASTER", card_title)
    summary.merge_range("A10:C11", "", card_value_good if permanent_product_count == 0 else card_value_warn)
    summary.write_formula(
        "A10",
        '=IF($L$1=0,"CATALOG COMPLETE",$L$1&" NEW STYLE(S) NEED SETUP")',
        card_value_good if permanent_product_count == 0 else card_value_warn,
        "CATALOG COMPLETE" if permanent_product_count == 0 else f"{permanent_product_count} NEW STYLE(S) NEED SETUP",
    )
    summary.merge_range(11, 0, 13, 2, "Permanent product details reused for future orders", card_caption)

    summary.merge_range("D9:F9", "PURCHASE REVIEW", card_title)
    summary.merge_range("D10:F11", "", card_value_good if review_complete_initial else card_value_warn)
    summary.write_formula(
        "D10",
        '=IF($M$1>0,$M$1&" ORDER DECISION(S) REQUIRED",IF($O$1>0,"WAITING ON "&$O$1&" DECISION(S)","REVIEW COMPLETE"))',
        card_value_good if review_complete_initial else card_value_warn,
        (
            f"{exception_count} ORDER DECISION(S) REQUIRED" if exception_count
            else (f"WAITING ON {review_reports} DECISION(S)" if review_reports else "REVIEW COMPLETE")
        ),
    )
    summary.merge_range(11, 3, 13, 5, "Order-specific blockers and customer decisions", card_caption)

    if report_mode == UNIFORM_SIZING_EVENT:
        summary.merge_range("G9:J9", "EMPLOYEE TOTALS", card_title)
        summary.merge_range("G10:J11", "MANAGED IN APP", card_value_good)
        summary.merge_range(11, 6, 13, 9, "Adjust discounts and verify sizing-event totals on the Purchase Manager dashboard", card_caption)
    else:
        summary.merge_range("G9:J9", "PO NUMBERS", card_title)
        summary.merge_range("G10:J11", "", card_value_good if is_ready else card_value_warn)
        summary.write_formula(
            "G10",
            '=IF(AND($L$1=0,$M$1=0,$O$1=0),"CONFIRM PO NUMBERS","AVAILABLE AFTER REVIEW")',
            card_value_good if is_ready else card_value_warn,
            "CONFIRM PO NUMBERS" if is_ready else "AVAILABLE AFTER REVIEW",
        )
        summary.merge_range(11, 6, 13, 9, "PO numbers are the final step after Product Master and Purchase Review are complete", card_caption)

    workflow_title = workbook.add_format({
        "bold": True, "font_size": 11, "font_color": PURPLE_DARK,
        "bg_color": PURPLE_PALE, "align": "center", "valign": "vcenter",
        "border": 1, "border_color": BORDER,
    })
    workflow_text = workbook.add_format({
        "font_size": 13, "font_color": TEXT, "bg_color": WHITE,
        "align": "center", "valign": "vcenter", "text_wrap": True,
        "border": 1, "border_color": BORDER,
    })
    summary.merge_range("A16:J16", "WORKFLOW ACTIONS", workflow_title)

    # Professional 4.0 preserves direct navigation while preserving the new
    # one-click regeneration workflow. Product Master and Purchase Review can
    # both be opened when both types of work remain.
    summary.merge_range("A17:C18", "", action_button if permanent_product_count else action_button_disabled)
    if permanent_product_count:
        summary.write_formula(
            "A17",
            '=HYPERLINK("#\'Product Master\'!A1","OPEN PRODUCT MASTER"&CHAR(10)&$L$1&" STYLE(S) NEED SETUP")',
            action_button,
            f"OPEN PRODUCT MASTER\n{permanent_product_count} STYLE(S) NEED SETUP",
        )
    else:
        summary.write("A17", "PRODUCT MASTER COMPLETE", action_button_good)

    review_action_format = (
        action_button_secondary if permanent_product_count and (exception_count or review_reports)
        else action_button if (exception_count or review_reports)
        else action_button_good
    )
    summary.merge_range("D17:F18", "", review_action_format)
    summary.write_formula(
        "D17",
        '=IF(OR($M$1>0,$O$1>0),HYPERLINK("#\'Review & Edit\'!A1","OPEN PURCHASE REVIEW"&CHAR(10)&($M$1+$O$1)&" DECISION(S) REMAIN"),"PURCHASE REVIEW COMPLETE")',
        review_action_format,
        (
            f"OPEN PURCHASE REVIEW\n{exception_count + review_reports} DECISION(S) REMAIN"
            if exception_count or review_reports else "PURCHASE REVIEW COMPLETE"
        ),
    )
    summary.conditional_format("D17:F18", {
        "type": "formula", "criteria": "=AND($M$1=0,$O$1=0)", "format": action_button_good,
    })

    summary.merge_range("G17:J18", "", action_button_warn if (permanent_product_count and regenerate_helper) else action_button_disabled)
    if permanent_product_count and regenerate_helper:
        # Professional 4.3 uses a registered Orchid regeneration request file.
        # macOS opens this file with Orchid Purchase Manager, which saves/closes
        # the current workbook, carries forward Review & Edit corrections, and
        # opens the refreshed workbook. No .command file is involved.
        summary.write_url(
            "G17", f"external:{regeneration_request_path.name}", action_button_warn,
            "REGENERATE PURCHASE REVIEW\nAFTER PRODUCT MASTER CHANGES",
        )
    elif is_ready:
        summary.write_formula(
            "G17", '=HYPERLINK("#\'Dashboard\'!A22","CONFIRM PO NUMBERS")',
            action_button_good, "CONFIRM PO NUMBERS",
        )
    else:
        summary.write("G17", "REGENERATE AFTER PRODUCT MASTER CHANGES", action_button_disabled)

    summary.merge_range("A19:J19", "", workflow_text)
    summary.write_formula(
        "A19",
        '=IF($L$1>0,"Open Product Master to complete reusable product setup. Open Purchase Review for any separate order decisions. Then regenerate to apply Product Master changes.",IF(OR($M$1>0,$O$1>0),"Open Purchase Review and resolve the remaining order-specific decisions.","All review work is complete. Confirm PO numbers below, then generate final purchase orders in Orchid Purchase Manager."))',
        workflow_text,
        (
            "Open Product Master to complete reusable product setup. Open Purchase Review for any separate order decisions. Then regenerate to apply Product Master changes."
            if permanent_product_count else
            ("Open Purchase Review and resolve the remaining order-specific decisions."
             if exception_count or review_reports else
             "All review work is complete. Confirm PO numbers below, then generate final purchase orders in Orchid Purchase Manager.")
        ),
    )
    summary.set_row(15, 24)
    summary.set_row(16, 34)
    summary.set_row(17, 34)
    summary.set_row(18, 38)

    # The purchase-order table now spans the full Dashboard width, uses larger
    # type and taller rows, and begins immediately below the action bar.
    po_header = workbook.add_format({
        "bold": True, "font_size": 12, "font_color": WHITE,
        "bg_color": PURPLE, "align": "center", "valign": "vcenter",
        "border": 1, "border_color": PURPLE_DARK,
    })
    po_vendor = workbook.add_format({
        "font_size": 12, "font_color": TEXT, "bg_color": WHITE,
        "align": "left", "valign": "vcenter", "border": 1,
        "border_color": BORDER,
    })
    po_text = workbook.add_format({
        "font_size": 12, "font_color": TEXT, "bg_color": WHITE,
        "align": "center", "valign": "vcenter", "border": 1,
        "border_color": BORDER,
    })
    po_number_fmt = workbook.add_format({
        "bold": True, "font_size": 12, "font_color": PURPLE_DARK,
        "bg_color": WHITE, "align": "center", "valign": "vcenter",
        "border": 1, "border_color": BORDER, "num_format": "@",
    })
    po_status_fmt = workbook.add_format({
        "bold": True, "font_size": 13, "font_color": TEXT,
        "bg_color": WHITE, "align": "center", "valign": "vcenter",
        "border": 1, "border_color": BORDER,
    })
    po_status_ready_badge = workbook.add_format({
        "bold": True, "font_size": 13, "font_color": "#166534",
        "bg_color": "#DDF4E4", "align": "center", "valign": "vcenter",
        "border": 2, "border_color": "#6FA77A",
    })
    po_status_review_badge = workbook.add_format({
        "bold": True, "font_size": 13, "font_color": "#8A3F00",
        "bg_color": "#FFE9C9", "align": "center", "valign": "vcenter",
        "border": 2, "border_color": "#DDA35C",
    })

    dashboard_routes = route_frame[["Vendor", "Report Type", "PO Number", "Status"]].copy()
    route_section_row = 20
    route_header_row = 21
    route_data_start = route_header_row + 1
    route_data_end = route_data_start + max(len(dashboard_routes) - 1, 0)
    status_col_letter = "I" if report_mode == UNIFORM_SIZING_EVENT else "H"

    summary.merge_range(route_section_row, 0, route_section_row, 5, "PURCHASE ORDERS", fmt["section"])
    summary.merge_range(route_section_row, 6, route_section_row, 9, "", po_summary)
    if len(dashboard_routes):
        first_excel = route_data_start + 1
        last_excel = route_data_end + 1
        summary.write_formula(
            route_section_row, 6,
            f'=COUNTIF(${status_col_letter}${first_excel}:${status_col_letter}${last_excel},"Ready")&" ready  |  "&COUNTIF(${status_col_letter}${first_excel}:${status_col_letter}${last_excel},"Needs Review")&" blocked"',
            po_summary,
            f"{ready_reports} ready  |  {review_reports} blocked",
        )
    else:
        summary.write(route_section_row, 6, "0 ready  |  0 blocked", po_summary)

    if report_mode == UNIFORM_SIZING_EVENT:
        summary.merge_range(route_header_row, 0, route_header_row, 2, "Vendor", po_header)
        summary.merge_range(route_header_row, 3, route_header_row, 5, "Report Type", po_header)
        summary.merge_range(route_header_row, 6, route_header_row, 7, "PO Number", po_header)
        summary.merge_range(route_header_row, 8, route_header_row, 9, "Status", po_header)
    else:
        summary.merge_range(route_header_row, 0, route_header_row, 3, "Vendor", po_header)
        summary.merge_range(route_header_row, 4, route_header_row, 6, "PO Number", po_header)
        summary.merge_range(route_header_row, 7, route_header_row, 9, "Status", po_header)
    summary.set_row(route_header_row, 30)

    for r, values in enumerate(dashboard_routes.itertuples(index=False, name=None), start=route_data_start):
        vendor_value = clean_text(values[dashboard_routes.columns.get_loc("Vendor")])
        report_value = clean_text(values[dashboard_routes.columns.get_loc("Report Type")])
        po_value = clean_text(values[dashboard_routes.columns.get_loc("PO Number")])
        status_value = clean_text(values[dashboard_routes.columns.get_loc("Status")])
        excel_row = r + 1

        if vendor_value == "Needs Vendor Assignment":
            status_formula = None
            status_fallback = "Needs Review"
        else:
            vendor_escaped = vendor_value.replace('"', '""')
            vendor_condition = f"'Review & Edit'!$O:$O,\"{vendor_escaped}\""
            product_master_guard = f"COUNTIF('Product Master'!$D:$D,\"{vendor_escaped}\")"
            report_lower = report_value.casefold()
            if "screen" in report_lower:
                review_guard = f"SUM(COUNTIFS('Review & Edit'!$B:$B,\"Needs Review\",'Review & Edit'!$U:$U,\"<>No\",{vendor_condition},'Review & Edit'!$P:$P,{{\"Screen Print\",\"Screen Printing\"}}))"
            elif "embroider" in report_lower:
                review_guard = f"SUM(COUNTIFS('Review & Edit'!$B:$B,\"Needs Review\",'Review & Edit'!$U:$U,\"<>No\",{vendor_condition},'Review & Edit'!$P:$P,{{\"Embroidery\",\"Blank Garment*\"}}))"
            elif "blank" in report_lower:
                review_guard = f"COUNTIFS('Review & Edit'!$B:$B,\"Needs Review\",'Review & Edit'!$U:$U,\"<>No\",{vendor_condition},'Review & Edit'!$P:$P,\"Blank Garment*\")"
            else:
                review_guard = f"COUNTIFS('Review & Edit'!$B:$B,\"Needs Review\",'Review & Edit'!$U:$U,\"<>No\",{vendor_condition})"
            status_formula = f'=IF(OR({product_master_guard}>0,{review_guard}>0),"Needs Review","Ready")'
            status_fallback = status_value

        if report_mode == UNIFORM_SIZING_EVENT:
            summary.merge_range(r, 0, r, 2, vendor_value, po_vendor)
            summary.merge_range(r, 3, r, 5, report_value, po_text)
            summary.merge_range(r, 6, r, 7, po_value, po_number_fmt)
            summary.merge_range(r, 8, r, 9, "", po_status_fmt)
            status_cell = f"I{excel_row}"
        else:
            summary.merge_range(r, 0, r, 3, vendor_value, po_vendor)
            summary.merge_range(r, 4, r, 6, po_value, po_number_fmt)
            summary.merge_range(r, 7, r, 9, "", po_status_fmt)
            status_cell = f"H{excel_row}"

        if status_formula:
            summary.write_formula(status_cell, status_formula, po_status_fmt, status_fallback)
        else:
            summary.write(status_cell, status_fallback, po_status_fmt)
        summary.set_row(r, 34)

        status_letter = "I" if report_mode == UNIFORM_SIZING_EVENT else "H"
        summary.conditional_format(r, 0, r, 9, {
            "type": "formula", "criteria": f'=${status_letter}{excel_row}="Ready"',
            "format": workbook.add_format({"bg_color": GREEN_PALE}),
        })
        summary.conditional_format(r, 0, r, 9, {
            "type": "formula", "criteria": f'=${status_letter}{excel_row}="Needs Review"',
            "format": workbook.add_format({"bg_color": RED_PALE}),
        })
        status_first_col = 8 if report_mode == UNIFORM_SIZING_EVENT else 7
        status_last_col = 9
        summary.conditional_format(r, status_first_col, r, status_last_col, {
            "type": "formula", "criteria": f'=${status_letter}{excel_row}="Ready"',
            "format": po_status_ready_badge,
        })
        summary.conditional_format(r, status_first_col, r, status_last_col, {
            "type": "formula", "criteria": f'=${status_letter}{excel_row}="Needs Review"',
            "format": po_status_review_badge,
        })

    # O1 is an independent live guard. Even if the deduplicated decision counter
    # reaches zero, the Dashboard cannot say Review Complete while any vendor
    # purchase-order route still displays Needs Review.
    if len(dashboard_routes):
        first_excel = route_data_start + 1
        last_excel = route_data_end + 1
        summary.write_formula(
            "O1",
            f'=COUNTIF(${status_col_letter}${first_excel}:${status_col_letter}${last_excel},"Needs Review")',
            None,
            review_reports,
        )
    else:
        summary.write_number("O1", 0)

    # Keep the table visible while the user scrolls through longer vendor lists.
    summary.freeze_panes(route_header_row + 1, 0)
    # Product Master is a deduplicated permanent setup list, not a copy
    # of every order line.
    if not detail.empty and permanent_mask.any():
        permanent_source = detail[permanent_mask].copy()
        permanent_source["_Product Master Decision Key"] = permanent_source.apply(_product_master_decision_key, axis=1)
        grouped_permanent = []
        for decision_key, group in permanent_source.groupby("_Product Master Decision Key", dropna=False, sort=False):
            colors = ", ".join(dict.fromkeys(
                clean_text(value) for value in group["Garment Color"] if clean_text(value)
            ))
            missing_information = "; ".join(dict.fromkeys(
                part.strip()
                for value in group["Permanent Review Reason"]
                for part in clean_text(value).split(";")
                if part.strip()
            ))
            style = clean_text(group["Product #"].iloc[0])
            unknown = any("product not found in product master" in clean_text(value).casefold() for value in group["Permanent Review Reason"])
            action = (
                f"Add {style or 'this product'} to Product Master and complete: {missing_information}."
                if unknown
                else f"Update {style or 'this product'} in Product Master: {missing_information}."
            )
            grouped_permanent.append({
                "Product #": style,
                "Description": clean_text(group["Description"].iloc[0]),
                "Garment Color(s)": colors,
                "Current Vendor": clean_text(group["Purchase Vendor"].iloc[0]),
                "Missing Information": missing_information,
                "Action Required": action,
                "Fix In": "Product Master",
                "Affected Orders": ", ".join(dict.fromkeys(clean_text(value) for value in group["Order Number"] if clean_text(value))),
                "Total Qty": int(group["Quantity"].sum()),
                "Source": clean_text(group["Source"].iloc[0]),
                "Affected Line IDs": ",".join(clean_text(value) for value in group["Line ID"] if clean_text(value)),
                "Decision Key": clean_text(decision_key),
            })
        product_master_needed = pd.DataFrame(grouped_permanent)
    else:
        product_master_needed = pd.DataFrame(columns=[
            "Product #", "Description", "Garment Color(s)", "Current Vendor",
            "Missing Information", "Action Required", "Fix In", "Affected Orders",
            "Total Qty", "Source", "Affected Line IDs", "Decision Key",
        ])
    pm_ws = _write_dataframe_sheet(
        workbook,
        "Product Master",
        "PRODUCT MASTER - PERMANENT PRODUCT SETUP",
        product_master_needed,
        fmt,
        note=(
            "Make these reusable corrections in Product Master, not on the current order. Open Orchid Purchase Manager, update or add the listed product, then create a fresh Purchase Review so the permanent settings are applied. "
            "Review & Edit is reserved for order-specific corrections such as missing size, color, product number, or one-time customer decisions."
        ),
    )
    pm_ws.set_tab_color(PURPLE)
    if len(product_master_needed.columns) >= 12:
        pm_ws.set_column(10, 11, None, None, {"hidden": True})
    pm_ws.set_column(0, 0, 15)
    pm_ws.set_column(1, 1, 38)
    pm_ws.set_column(2, 2, 24)
    pm_ws.set_column(3, 4, 24)
    pm_ws.set_column(5, 5, 62)
    pm_ws.set_column(6, 6, 18)
    pm_ws.set_column(7, 7, 24)

    review_columns = [
        "Line ID", "Review Status", "Fix In", "Action Required", "Resolution", "Review Reason",
        "Purchase Instructions", "Note Action", "Original Product #", "Product #", "Description",
        "Garment Color", "Size", "Quantity", "Purchase Vendor", "Decoration Type", "Decoration Location", "Decoration Placement Instructions", "Decoration Color",
        "Employee Name", "Company", "Order Number", "Include", "Do Not Outsource", "Decision Key",
        "Decoration Decision",
    ]
    if not detail.empty:
        review_edit = detail[exception_mask].copy()
        review_edit["Original Product #"] = review_edit["Product #"]
        review_edit["Review Reason"] = review_edit["Event Review Reason"]
        review_edit["Fix In"] = "Purchase Review"
        review_edit["Action Required"] = review_edit.apply(
            lambda row: (
                "Review the full decoration note, choose an explicit order-specific decoration decision, then save."
                if "customer decision required" in clean_text(row.get("Event Review Reason", "")).casefold()
                else (
                    "Correct the product number for this order, or mark it as a one-time manual item."
                    if "missing product number" in clean_text(row.get("Event Review Reason", "")).casefold()
                    else f"Correct this order in Purchase Review: {clean_text(row.get('Event Review Reason', ''))}."
                )
            ),
            axis=1,
        )
        review_edit["Resolution"] = review_edit.apply(
            lambda row: (
                "Choose Decoration Override"
                if "customer decision required" in clean_text(row.get("Event Review Reason", "")).casefold()
                else ("Correct Product Number" if "missing product number" in clean_text(row.get("Event Review Reason", "")).casefold() else "Complete Missing Information")
            ),
            axis=1,
        )
        review_edit = review_edit[review_columns]
    else:
        review_edit = pd.DataFrame(columns=review_columns)
    review_ws = _write_dataframe_sheet(
        workbook,
        "Review & Edit",
        "PURCHASE REVIEW - CURRENT-ORDER DECISIONS",
        review_edit,
        fmt,
        note=(
            "Work only on the decisions shown here. Correct the highlighted order information, then change Review Status to Ready. "
            "Permanent vendor and decoration setup stays in Product Master. Excel remains the detailed, editable review workspace."
        ),
    )
    review_ws.set_tab_color(PURPLE)
    review_ws.set_zoom(82)
    # Keep the decision, action, and resolution columns visible while scrolling
    # through the detailed Excel review. Technical IDs remain hidden.
    review_ws.freeze_panes(3, 5)
    review_ws.set_column(0, 0, None, None, {"hidden": True})
    review_count = max(len(review_edit), 50)
    rstart, rend = 3, 3 + review_count - 1
    review_col = {name: index for index, name in enumerate(review_columns)}
    review_ws.data_validation(rstart, review_col["Review Status"], rend, review_col["Review Status"], {"validate": "list", "source": ["Needs Review", "Ready"]})
    review_ws.data_validation(rstart, review_col["Resolution"], rend, review_col["Resolution"], {
        "validate": "list",
        "source": ["Correct Product Number", "One-Time Manual Item", "Complete Missing Information", "Choose Decoration Override"],
    })
    if vendors:
        review_ws.data_validation(rstart, review_col["Purchase Vendor"], rend, review_col["Purchase Vendor"], {"validate": "list", "source": "=Settings!$A$2:$A$%d" % (len(vendors) + 1)})
    review_ws.data_validation(rstart, review_col["Decoration Type"], rend, review_col["Decoration Type"], {"validate": "list", "source": ["Embroidery", "Screen Print", SEW_ON_PATCH_LABEL, HEMMING_ALTERATION_LABEL, BLANK_DECORATION_LABEL]})
    review_ws.data_validation(rstart, review_col["Decoration Location"], rend, review_col["Decoration Location"], {
        "validate": "list", "source": DECORATION_LOCATIONS, "error_type": "warning",
        "error_title": "Custom Location", "error_message": "Choose a standard location or type the exact custom location."
    })
    review_ws.data_validation(rstart, review_col["Include"], rend, review_col["Include"], {"validate": "list", "source": ["Yes", "No"]})
    review_ws.data_validation(rstart, review_col["Do Not Outsource"], rend, review_col["Do Not Outsource"], {"validate": "list", "source": ["No", "Yes"]})
    review_ws.data_validation(rstart, review_col["Decoration Decision"], rend, review_col["Decoration Decision"], {
        "validate": "list",
        "source": ["Follow Note as Written", "Keep Product Master Default", "No Decoration", "Embroidery", "Screen Print", "Sew On Patch", "Hemming / Alteration", "Do Not Outsource — Ship to Orchid"],
    })
    review_ws.set_column(review_col["Line ID"], review_col["Line ID"], 11)
    review_ws.set_column(review_col["Review Status"], review_col["Review Status"], 14)
    review_ws.set_column(review_col["Fix In"], review_col["Fix In"], 18)
    review_ws.set_column(review_col["Action Required"], review_col["Action Required"], 52)
    review_ws.set_column(review_col["Resolution"], review_col["Resolution"], 23)
    review_ws.set_column(review_col["Review Reason"], review_col["Review Reason"], 38)
    review_ws.set_column(review_col["Purchase Instructions"], review_col["Note Action"], 24)
    review_ws.set_column(review_col["Original Product #"], review_col["Product #"], 15, workbook.add_format({"num_format": "@", "font_color": TEXT, "border": 1, "border_color": BORDER, "valign": "top"}))
    review_ws.set_column(review_col["Description"], review_col["Description"], 34)
    review_ws.set_column(review_col["Garment Color"], review_col["Size"], 16)
    review_ws.set_column(review_col["Quantity"], review_col["Quantity"], 10)
    review_ws.set_column(review_col["Purchase Vendor"], review_col["Decoration Color"], 18)
    review_ws.set_column(review_col["Employee Name"], review_col["Order Number"], 22)
    review_ws.set_column(review_col["Include"], review_col["Include"], 10)
    review_ws.set_column(review_col["Do Not Outsource"], review_col["Do Not Outsource"], None, None, {"hidden": True})
    review_ws.set_column(review_col["Decision Key"], review_col["Decision Key"], None, None, {"hidden": True})
    review_ws.set_column(review_col["Decoration Decision"], review_col["Decoration Decision"], 28)
    review_ws.conditional_format(rstart, review_col["Review Status"], rend, review_col["Review Status"], {
        "type": "text", "criteria": "containing", "value": "Needs Review",
        "format": workbook.add_format({"bg_color": RED_PALE, "font_color": "#9C0006"}),
    })
    review_ws.conditional_format(rstart, review_col["Review Status"], rend, review_col["Review Status"], {
        "type": "text", "criteria": "containing", "value": "Ready",
        "format": workbook.add_format({"bg_color": GREEN_PALE, "font_color": "#006100"}),
    })
    review_ws.conditional_format(rstart, review_col["Note Action"], rend, review_col["Note Action"], {
        "type": "text", "criteria": "containing", "value": "Decision Required",
        "format": workbook.add_format({"bg_color": YELLOW}),
    })
    # Visual priority cues without taking away Excel's editing flexibility.
    issue_red = workbook.add_format({"bg_color": "#FCE8E6", "font_color": "#9C0006", "text_wrap": True})
    issue_orange = workbook.add_format({"bg_color": "#FFF4E8", "font_color": "#8A4B00", "text_wrap": True})
    issue_yellow = workbook.add_format({"bg_color": YELLOW, "font_color": TEXT, "text_wrap": True})
    review_ws.conditional_format(rstart, review_col["Review Reason"], rend, review_col["Review Reason"], {
        "type": "text", "criteria": "containing", "value": "not found in Product Master", "format": issue_red,
    })
    for issue_text in ("Missing size", "Missing garment color", "Missing product number"):
        review_ws.conditional_format(rstart, review_col["Review Reason"], rend, review_col["Review Reason"], {
            "type": "text", "criteria": "containing", "value": issue_text, "format": issue_orange,
        })
    review_ws.conditional_format(rstart, review_col["Review Reason"], rend, review_col["Review Reason"], {
        "type": "text", "criteria": "containing", "value": "Customer decision", "format": issue_yellow,
    })
    editable_hint = workbook.add_format({"bg_color": "#FFFBEA"})
    for edit_name in (
        "Review Status", "Resolution", "Product #", "Garment Color", "Size", "Quantity",
        "Purchase Vendor", "Decoration Type", "Decoration Location", "Decoration Placement Instructions", "Decoration Color",
        "Include", "Do Not Outsource", "Decoration Decision",
    ):
        edit_col = review_col[edit_name]
        review_ws.conditional_format(rstart, edit_col, rend, edit_col, {
            "type": "formula", "criteria": "=TRUE", "format": editable_hint,
        })
    # Status colors stay authoritative over the edit hint.
    review_ws.conditional_format(rstart, review_col["Review Status"], rend, review_col["Review Status"], {
        "type": "text", "criteria": "containing", "value": "Needs Review",
        "format": workbook.add_format({"bg_color": RED_PALE, "font_color": "#9C0006", "bold": True}),
    })
    review_ws.conditional_format(rstart, review_col["Review Status"], rend, review_col["Review Status"], {
        "type": "text", "criteria": "containing", "value": "Ready",
        "format": workbook.add_format({"bg_color": GREEN_PALE, "font_color": "#006100", "bold": True}),
    })

    inferred_columns = [
        "Line ID", "Purchase Vendor", "Decoration Type", "Decoration Location", "Decoration Placement Instructions", "Decoration Color",
        "Product #", "Description", "Garment Color", "Size", "Quantity",
        "Employee Name", "Company", "Order Number", "Decoration Source", "Include", "Do Not Outsource",
    ]
    inferred_routing = (
        detail[detail["Decoration Source"].map(clean_text).ne("")][inferred_columns].copy()
        if not detail.empty else pd.DataFrame(columns=inferred_columns)
    )
    inferred_ws = _write_dataframe_sheet(
        workbook,
        "Inferred Routing",
        "INFERRED DECORATION ROUTING — EDITABLE",
        inferred_routing,
        fmt,
        note=(
            "These decoration assignments were inferred from the embroidery and screen-print quantities on each Shopify order. "
            "They do not count as blockers. Review them when desired and change Purchase Vendor, Decoration Type, Decoration Location, Placement Instructions, or Decoration Color before generating final purchase orders."
        ),
    )
    inferred_count = max(len(inferred_routing), 50)
    istart, iend = 3, 3 + inferred_count - 1
    inferred_col = {name: index for index, name in enumerate(inferred_columns)}
    if vendors:
        inferred_ws.data_validation(istart, inferred_col["Purchase Vendor"], iend, inferred_col["Purchase Vendor"], {"validate": "list", "source": "=Settings!$A$2:$A$%d" % (len(vendors) + 1)})
    inferred_ws.data_validation(istart, inferred_col["Decoration Type"], iend, inferred_col["Decoration Type"], {"validate": "list", "source": ["Embroidery", "Screen Print", SEW_ON_PATCH_LABEL, HEMMING_ALTERATION_LABEL, BLANK_DECORATION_LABEL]})
    inferred_ws.data_validation(istart, inferred_col["Decoration Location"], iend, inferred_col["Decoration Location"], {
        "validate": "list", "source": DECORATION_LOCATIONS, "error_type": "warning",
        "error_title": "Custom Location", "error_message": "Choose a standard location or type the exact custom location."
    })
    inferred_ws.data_validation(istart, inferred_col["Include"], iend, inferred_col["Include"], {"validate": "list", "source": ["Yes", "No"]})
    inferred_ws.data_validation(istart, inferred_col["Do Not Outsource"], iend, inferred_col["Do Not Outsource"], {"validate": "list", "source": ["No", "Yes"]})
    inferred_ws.set_column(inferred_col["Line ID"], inferred_col["Line ID"], 11)
    inferred_ws.set_column(inferred_col["Purchase Vendor"], inferred_col["Decoration Color"], 18)
    inferred_ws.set_column(inferred_col["Product #"], inferred_col["Product #"], 15, workbook.add_format({"num_format": "@", "font_color": TEXT, "border": 1, "border_color": BORDER, "valign": "top"}))
    inferred_ws.set_column(inferred_col["Description"], inferred_col["Description"], 36)
    inferred_ws.set_column(inferred_col["Garment Color"], inferred_col["Size"], 16)
    inferred_ws.set_column(inferred_col["Quantity"], inferred_col["Quantity"], 10)
    inferred_ws.set_column(inferred_col["Employee Name"], inferred_col["Order Number"], 22)
    inferred_ws.set_column(inferred_col["Decoration Source"], inferred_col["Decoration Source"], 38)
    inferred_ws.set_column(inferred_col["Include"], inferred_col["Include"], 10)
    inferred_ws.set_column(inferred_col["Do Not Outsource"], inferred_col["Do Not Outsource"], 18)
    # Technical routing detail remains available to the app but is hidden from
    # normal users. It can still be unhidden in Excel for troubleshooting.
    inferred_ws.hide()

    all_ws = _write_dataframe_sheet(
        workbook, "All PO Lines", "ALL PURCHASE ORDER LINES — SOURCE", detail, fmt,
        note="Complete source data used by the app. Use Review & Edit for normal corrections."
    )
    _add_edit_controls(workbook, all_ws, max(len(detail), 1), {"vendors": vendors})
    all_ws.hide()

    if report_mode == UNIFORM_SIZING_EVENT:
        employee_totals = _build_employee_totals(raw)
        employee_ws = _write_dataframe_sheet(
            workbook,
            "Employee Totals",
            "EMPLOYEE ORDER TOTALS — EDITABLE",
            employee_totals,
            fmt,
            note=(
                "Order Total is the original Shopify amount. Enter any adjustment in Discount, such as $3.00 for a $703.00 order. "
                "Total updates automatically to show the amount that should appear on the customer report or invoice."
            ),
        )
        data_start = 3
        data_end = data_start + max(len(employee_totals), 1) - 1
        # Total = Order Total - Discount.
        for row in range(data_start, data_start + len(employee_totals)):
            employee_ws.write_formula(row, 5, f"=D{row+1}-E{row+1}", fmt["currency"])
        grand_row = data_start + len(employee_totals) + 1
        employee_ws.write(grand_row, 4, "Grand Total", fmt["label"])
        first_excel = data_start + 1
        last_excel = data_start + len(employee_totals)
        if len(employee_totals):
            employee_ws.write_formula(grand_row, 5, f"=SUM(F{first_excel}:F{last_excel})", fmt["currency"])
        else:
            employee_ws.write_number(grand_row, 5, 0, fmt["currency"])
        # Employee Totals are edited from Mission Control in the desktop app.
        # Keep the worksheet available to the app and for troubleshooting, but
        # remove it from the normal employee-facing Excel workflow.
        employee_ws.hide()

    # Review & Edit is the single visible exception worksheet. All PO Lines is hidden as the full source.

    manual = pd.DataFrame([{column: "" for column in ALL_COLUMNS} for _ in range(50)])
    manual["Include"] = "No"
    manual["Do Not Outsource"] = "No"
    manual["Review Status"] = "Needs Review"
    manual_ws = _write_dataframe_sheet(
        workbook,
        "Manual Items",
        "MANUAL ITEMS — EDITABLE",
        manual,
        fmt,
        note="Add products that were missed or need to be ordered manually. Set Include to Yes when complete.",
    )
    _add_edit_controls(workbook, manual_ws, 50, {"vendors": vendors})
    manual_ws.hide()

    excluded_ws = _write_dataframe_sheet(
        workbook,
        "Excluded Services",
        "EXCLUDED SHOPIFY SERVICE / DECORATION LINES",
        excluded,
        fmt,
        note="These Shopify service and decoration-fee lines are not ordered as garments.",
    )
    excluded_ws.hide()

    raw_for_excel = raw.copy()
    if len(raw_for_excel.columns) > 60:
        raw_for_excel = raw_for_excel.iloc[:, :60]
    raw_ws = _write_dataframe_sheet(
        workbook,
        "Event Raw Data",
        "ORIGINAL SHOPIFY EXPORT — REFERENCE",
        raw_for_excel,
        fmt,
        note="Reference copy of the imported Shopify data.",
    )
    raw_ws.hide()

    product_master_ws = _write_dataframe_sheet(
        workbook,
        "Product Master Data",
        "PRODUCT MASTER — REFERENCE COPY",
        master,
        fmt,
        note="Update permanent assignments in Orchid Purchase Manager Product Master, not here.",
    )
    product_master_ws.hide()

    # Vendor preview tabs mirror the familiar Excel packet. All edits still belong on All PO Lines.
    if not detail.empty:
        for vendor in sorted({v for v in detail["Purchase Vendor"].map(lambda value: _canonical_vendor(value, vendor_registry)) if v}, key=str.casefold):
            subset = detail[detail["Purchase Vendor"].eq(vendor)].copy()
            preview_columns = [
                "Product #", "Garment Color", "Size", "Quantity",
                "Company", "Employee Name", "Decoration Type", "Decoration Location", "Decoration Placement Instructions", "Decoration Color",
                "Order Number", "Purchase Instructions",
            ]
            if vendor.casefold() == "tru-spec":
                preview_columns.insert(1, "Description")
            preview = subset[preview_columns].rename(columns={"Garment Color": "Color", "Purchase Instructions": "Purchase Instructions"})
            def _preview_report_type(value):
                value = clean_text(value)
                if report_mode != UNIFORM_SIZING_EVENT:
                    return "Combined Vendor Order"
                if "screen" in value.casefold():
                    return "Screen Printing"
                if is_blank_decoration(value):
                    return "Blank Garments"
                return value or "Needs Routing"
            vendor_ws = _write_dataframe_sheet(
                workbook,
                vendor,
                f"{vendor.upper()} PURCHASE ORDER PREVIEW",
                preview,
                fmt,
                note="Preview only. Make corrections on Review & Edit before finalizing.",
            )
            vendor_ws.hide()

    settings = workbook.add_worksheet("Settings")
    settings.write(0, 0, "Vendors", fmt["header"])
    for r, vendor in enumerate(vendors, start=1):
        settings.write(r, 0, vendor)
    settings.write(0, 1, "Decoration Types", fmt["header"])
    for r, value in enumerate(["Embroidery", "Screen Print", SEW_ON_PATCH_LABEL, HEMMING_ALTERATION_LABEL, BLANK_DECORATION_LABEL], start=1):
        settings.write(r, 1, value)
    settings.write(0, 2, "Include", fmt["header"])
    settings.write(1, 2, "Yes")
    settings.write(2, 2, "No")
    settings.write(0, 3, "Purchase Order Mode", fmt["header"])
    settings.write(1, 3, report_mode)
    settings.write(0, 4, "Available Modes", fmt["header"])
    settings.write(1, 4, UNIFORM_SIZING_EVENT)
    settings.write(2, 4, GENERAL_SALES_PERIOD)
    settings.write(0, 5, "Event Name", fmt["header"])
    settings.write(1, 5, event_name)
    settings.write(0, 6, "Decoration Locations", fmt["header"])
    for r, value in enumerate(DECORATION_LOCATIONS, start=1):
        settings.write(r, 6, value)
    settings.write(0, 7, "Decoration Fulfillment", fmt["header"])
    settings.write(1, 7, decoration_fulfillment)
    settings.hide()

    system_info = workbook.add_worksheet("System Info")
    system_info.write(0, 0, "Product Master Path")
    system_info.write(0, 1, str(product_master_path))
    system_info.write(1, 0, "Product Master Modified")
    system_info.write(1, 1, str(master_signature.get("modified", "")))
    system_info.write(2, 0, "Product Master SHA256")
    system_info.write(2, 1, str(master_signature.get("sha256", "")))
    system_info.write(3, 0, "Product Master Records")
    system_info.write(3, 1, int(master_signature.get("records", 0)))
    system_info.write(4, 0, "Product Master Styles")
    system_info.write(4, 1, int(master_signature.get("styles", 0)))
    system_info.write(5, 0, "Workbook Generated")
    system_info.write(5, 1, now.strftime("%Y-%m-%d %I:%M:%S %p"))
    system_info.hide()

    workbook.close()

    if regenerate_helper:
        request_payload = {
            "source_csv": str(Path(shopify_csv_path).expanduser().resolve()),
            "report_mode": report_mode,
            "event_name": event_name,
            "decoration_fulfillment": decoration_fulfillment,
            "previous_workbook": str(output_path.resolve()),
            "created_at": now.isoformat(),
        }
        regeneration_request_path.write_text(json.dumps(request_payload, indent=2), encoding="utf-8")

    return {
        "output_path": output_path,
        "output_dir": output_path.parent,
        "regeneration_request": regeneration_request_path if regenerate_helper else None,
        "lines": len(detail),
        "quantity": int(detail["Quantity"].sum()) if not detail.empty else 0,
        "review_lines": len(review),
        "review_decisions": count_unique_decisions(review),
        "product_master_decisions": int(len(product_master_needed)),
        "purchase_review_decisions": int(exception_count),
        "inferred_routing_lines": int(detail["Decoration Source"].map(clean_text).ne("").sum()) if not detail.empty else 0,
        "notes_lines": int(detail["Purchase Instructions"].map(bool).sum()) if not detail.empty else 0,
        "purchasing_rules_needed": int(len(product_master_needed)),
        "missing_product_numbers": list(dict.fromkeys(
            clean_text(value)
            for value in (
                detail.loc[permanent_mask & detail["Master Match"].eq("No"), "Product #"].tolist()
                if not detail.empty and permanent_mask.any() else []
            )
            if clean_text(value)
        )),
        "product_master_signature": master_signature,
        "report_mode": report_mode,
        "event_name": event_name,
        "decoration_fulfillment": decoration_fulfillment,
    }
