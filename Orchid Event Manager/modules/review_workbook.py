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
    LEFT_CHEST,
    normalize_decoration_location,
)
from modules.note_rules import (
    alteration_note_only,
    always_embroidery_rain_style,
    blanket_decoration_note_requires_review,
    combine_purchase_instructions,
    decision_note_requires_review,
    decision_note_targets_line,
    decision_target_styles,
    decoration_note_changes_process,
    decoration_note_requires_review,
    decoration_note_targets_line,
    decoration_note_is_global,
    decoration_target_families,
    decoration_target_styles,
    note_action_label,
    note_requires_review,
    note_requires_review_for_line,
    screen_print_note_requires_review,
    waterproof_apparel_requires_review,
    waterproof_style_number,
    waterproof_style_requires_review,
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
    decoration_charge_label, is_final_sale_accounting_product,
    is_in_house_decoration, is_in_house_service_product, is_internal_service_style,
)
from modules.report_modes import GENERAL_SALES_PERIOD, UNIFORM_SIZING_EVENT, normalize_report_mode
from modules.decoration_fulfillment import (
    STANDARD_ORCHID_WORKFLOW, is_entire_order_outsourced, normalize_decoration_fulfillment,
)
from modules.master_sync import product_master_signature
from modules.packet_lock import assert_live_product_master, create_packet_lock
from modules.source_identity import manual_entry_correction_signature, strict_legacy_signature
from modules.outsource_rules import (
    ROUTING_POLICY_VERSION,
    is_boots_product,
    resolve_never_outsource,
    vendor_never_outsource,
)
from modules.purchase_vendor_rules import (
    INTERNAL_VENDOR_REVIEW_REASON,
    is_internal_purchase_vendor,
)
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
    INTERNAL_VENDOR_REVIEW_REASON.casefold(),
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
    "Vendor Product #",
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
    "Product Aliases",
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
    "Source ID",
    "Source Occurrence",
    "Original Quantity",
    "Original Order Number",
    "Original Company",
    "Original Employee Name",
    "Original Order Date",
    "Original Shopify SKU",
    "Quantity Override Confirmed",
]


def _line_note_context(row: pd.Series) -> str:
    return " ".join(filter(None, [
        clean_text(row.get("Product Name", row.get("Description", ""))),
        clean_text(row.get("Original Line Item", row.get("Original Shopify Line", ""))),
        clean_text(row.get("Product Category", "")),
        clean_text(row.get("Product Aliases", "")),
        clean_text(row.get("Style Number", row.get("Product #", ""))),
    ]))


def decoration_note_prompts_enabled(
    report_mode: str = GENERAL_SALES_PERIOD,
    decoration_fulfillment: str = STANDARD_ORCHID_WORKFLOW,
) -> bool:
    """Prompt for decoration notes only in workflows that require up-front decisions."""
    return (
        normalize_report_mode(report_mode) == UNIFORM_SIZING_EVENT
        or is_entire_order_outsourced(decoration_fulfillment)
    )


def waterproof_decoration_prompts_enabled(
    report_mode: str = GENERAL_SALES_PERIOD,
    decoration_fulfillment: str = STANDARD_ORCHID_WORKFLOW,
) -> bool:
    """Prompt whenever Product Master identifies waterproof or rainwear apparel.

    Waterproof is a permanent product characteristic, not merely an order-note
    condition. A saved Product Alias such as ``Waterproof`` or ``Rain Jacket``
    therefore requires an explicit order-level decoration decision in every
    workflow. Permanent always-embroidery exceptions remain excluded later.
    """
    return True



def _row_permanently_stays_at_orchid(row: pd.Series) -> bool:
    return resolve_never_outsource(
        row.get("Never Outsource", ""),
        row.get("Product Name", row.get("Description", "")),
        row.get("Product Category", ""),
        row.get("Style Number", row.get("Product #", "")),
    )


def _note_prompt_applicable_to_row(row: pd.Series, note: object, allow_decoration_prompts: bool) -> bool:
    """Return whether a note creates a decision for this exact purchase line."""
    if alteration_note_only(note):
        return False
    if decision_note_requires_review(note):
        # A do-not-outsource instruction is already satisfied by a permanent
        # in-house route and does not need the user to acknowledge it again.
        lowered = clean_text(note).casefold()
        if _row_permanently_stays_at_orchid(row) and any(term in lowered for term in (
            "do not outsource", "don't outsource", "ship to orchid", "send to orchid",
        )):
            return False
        return True
    if not allow_decoration_prompts or not decoration_note_requires_review(note):
        return False
    # Operational logo/thread/personalization notes on products already routed
    # to Orchid belong on the decoration report, not in Purchase Review. A real
    # process change (screen print, no decoration, etc.) still requires review.
    if _row_permanently_stays_at_orchid(row) and not decoration_note_changes_process(note):
        return False
    return True

def _mandatory_note_review_flags(
    frame: pd.DataFrame,
    allow_decoration_prompts: bool = True,
) -> pd.Series:
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
            standard_workflow_override = (
                blanket_decoration_note_requires_review(line_note)
                or screen_print_note_requires_review(line_note)
            )
            line_prompts_enabled = allow_decoration_prompts or standard_workflow_override
            if _note_prompt_applicable_to_row(row, line_note, line_prompts_enabled):
                flags.at[index] = True

        if decision_note_requires_review(order_note):
            style_targets = decision_target_styles(order_note)
            family_targets = decoration_target_families(order_note)
            if style_targets or family_targets:
                matches = [
                    index for index, row in group.iterrows()
                    if decision_note_targets_line(order_note, _line_note_context(row))
                    and _note_prompt_applicable_to_row(row, order_note, True)
                ]
                if matches:
                    flags.loc[matches] = True
            else:
                matches = [index for index, row in group.iterrows() if _note_prompt_applicable_to_row(row, order_note, True)]
                if matches:
                    flags.loc[matches] = True
            continue

        standard_workflow_override = (
            blanket_decoration_note_requires_review(order_note)
            or screen_print_note_requires_review(order_note)
        )
        if (
            not decoration_note_requires_review(order_note)
            or (not allow_decoration_prompts and not standard_workflow_override)
        ):
            continue

        style_targets = decoration_target_styles(order_note)
        family_targets = decoration_target_families(order_note)
        is_global = decoration_note_is_global(order_note)
        if style_targets or family_targets or is_global:
            matches = [
                index for index, row in group.iterrows()
                if decoration_note_targets_line(order_note, _line_note_context(row))
                and _note_prompt_applicable_to_row(row, order_note, True)
            ]
        else:
            # One consolidated prompt is enough for an ambiguous note, but choose
            # an outsource-eligible line. If every line already stays at Orchid,
            # preserve the note on reports without creating an empty decision.
            matches = [
                index for index, row in group.iterrows()
                if _note_prompt_applicable_to_row(row, order_note, True)
            ][:1]
        if matches:
            flags.loc[matches] = True
    return flags


def _boolean_mask(values: object, index: pd.Index) -> pd.Series:
    """Return a real boolean mask, including for an empty StringDtype Series.

    pandas 3 preserves StringDtype for some empty ``map``/``apply`` results.
    Bitwise mask operations on those empty string arrays raise instead of
    producing an empty boolean result.  Purchase Review must still build an
    Employee Totals workbook for an event containing only a Final Sale Boot or
    other non-purchase line, so normalize every dynamic mask explicitly.
    """
    if isinstance(values, pd.Series):
        values = values.reindex(index).tolist()
    elif values is None:
        values = []
    else:
        values = list(values)
    normalized: list[bool] = []
    for value in values:
        try:
            normalized.append(False if pd.isna(value) else bool(value))
        except (TypeError, ValueError):
            normalized.append(False)
    if len(normalized) < len(index):
        normalized.extend([False] * (len(index) - len(normalized)))
    return pd.Series(normalized[:len(index)], index=index, dtype=bool)


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
    shopify_csv_path: Path,
    product_master_path: Path,
    decoration_fulfillment: str = STANDARD_ORCHID_WORKFLOW,
    report_mode: str = GENERAL_SALES_PERIOD,
    normalized_orders: pd.DataFrame | None = None,
    parsed_orders: pd.DataFrame | None = None,
):
    raw = normalized_orders.copy(deep=True) if normalized_orders is not None else normalize_order_export(shopify_csv_path)
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

    parsed = (
        parsed_orders.copy(deep=True)
        if parsed_orders is not None
        else parse_shopify_orders(shopify_csv_path, product_master_path, normalized_orders=raw).copy()
    )
    master = load_extended_master(product_master_path)

    for column in [
        "Order Number", "Order Date", "Company", "Employee Name",
        "Shopify Order Notes", "Shopify Line Notes", "Original Line Item", "Product Name",
        "Style Number", "Garment Color", "Size", "Needs Review",
        "Parser Source", "Parse Confidence", "Detected Vendor",
        "Source ID", "Source Occurrence", "Original Quantity", "Original Order Number",
        "Original Company", "Original Employee Name", "Original Order Date", "Original Shopify SKU",
    ]:
        if column not in parsed.columns:
            parsed[column] = ""
        parsed[column] = parsed[column].fillna("").map(clean_text)

    if "Quantity" not in parsed.columns:
        parsed["Quantity"] = 0
    parsed["Quantity"] = pd.to_numeric(parsed["Quantity"], errors="coerce").fillna(0).astype(int)
    if "Original Quantity" not in parsed.columns:
        parsed["Original Quantity"] = parsed["Quantity"]
    parsed["Original Quantity"] = pd.to_numeric(parsed["Original Quantity"], errors="coerce").fillna(parsed["Quantity"]).astype(int)
    if "Source Occurrence" not in parsed.columns:
        parsed["Source Occurrence"] = 1
    parsed["Source Occurrence"] = pd.to_numeric(parsed["Source Occurrence"], errors="coerce").fillna(1).astype(int)

    # Defense in depth: service/charge rows must stay out of Purchase Review even
    # when parsed_orders was supplied by a cached or older parser. Product 750 is
    # Orchid's embroidery-charge code and is retained only in the raw/excluded data.
    if not parsed.empty:
        service_mask = parsed.apply(
            lambda row: is_decoration_service(
                row.get("Original Line Item", ""),
                row.get("Style Number", ""),
                row.get("Product Name", ""),
                row.get("Garment Color", ""),
            ),
            axis=1,
        )
        parsed = parsed.loc[~service_mask].copy()

    resolved_records = []
    resolution_cache = {}
    for _, source in parsed.iterrows():
        resolution_key = tuple(
            clean_text(source.get(field, "")).casefold()
            for field in ("Style Number", "Product Name", "Garment Color")
        )
        result = resolution_cache.get(resolution_key)
        if result is None:
            result = resolve_product(
                source.get("Style Number", ""),
                source.get("Product Name", ""),
                source.get("Garment Color", ""),
                master,
                master_prepared=True,
            )
            resolution_cache[resolution_key] = result
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
        "Style Number", "Vendor Product #", "Garment Color", "Size", "Master Match",
        "Parser Source", "Parse Confidence", "Detected Vendor", "Resolver Issue",
        "Product Category", "Product Aliases", "Requires Size", "Requires Color", "Requires Decoration", "Never Outsource",
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
        # Manual entries sometimes use the color field only as a vendor/size note,
        # for example "Portwest Size 38". After extracting 38, do not retain the
        # vendor name as a false garment color.
        vendor_only = merged.apply(
            lambda row: bool(
                clean_text(row.get("Garment Color", ""))
                and clean_text(row.get("Garment Color", "")).casefold()
                in {
                    clean_text(row.get("Vendor", "")).casefold(),
                    clean_text(row.get("Detected Vendor", "")).casefold(),
                }
            ),
            axis=1,
        )
        merged.loc[vendor_only, "Garment Color"] = ""

    # Decoration service lines are excluded from purchasing, but their quantities
    # are valuable routing data. Use them to classify product lines before asking
    # the user to make repetitive style-by-style decisions.
    merged = infer_order_decorations(merged, service_totals_from_raw(raw))

    # These frequently ordered Carhartt Rain Defender styles are known, safe
    # embroidery routes. Preserve any saved Product Master placement and fill a
    # missing legacy placement with Left Chest. An explicit contradictory order
    # note (for example, "Do not embroider CT100615") is still reviewed below.
    always_embroidery_mask = _boolean_mask(
        merged["Style Number"].map(always_embroidery_rain_style), merged.index
    )
    if always_embroidery_mask.any():
        merged.loc[always_embroidery_mask, "Decoration Type"] = "Embroidery"
        missing_location = always_embroidery_mask & _boolean_mask(
            merged["Decoration Location"].map(lambda value: not clean_text(value)), merged.index
        )
        merged.loc[missing_location, "Decoration Location"] = LEFT_CHEST
        merged.loc[always_embroidery_mask, "Requires Decoration"] = "Yes"

    allow_decoration_prompts = decoration_note_prompts_enabled(report_mode, decoration_fulfillment)
    merged["_Mandatory Note Review"] = _boolean_mask(
        _mandatory_note_review_flags(merged, allow_decoration_prompts=allow_decoration_prompts),
        merged.index,
    )
    outsourced_sizing_event = waterproof_decoration_prompts_enabled(
        report_mode, decoration_fulfillment
    )
    # A boot may truthfully contain the word "Waterproof" (for example Georgia
    # Boot Muddog G5594), but boots are never decorated. Exclude the permanent
    # Boots category before evaluating rainwear decoration prompts.
    boots_mask = _boolean_mask(
        merged.apply(
            lambda row: is_boots_product(
                row.get("Product Name", ""), row.get("Product Category", ""), row.get("Style Number", "")
            ),
            axis=1,
        ),
        merged.index,
    )
    forced_waterproof_style_mask = _boolean_mask(
        merged["Style Number"].map(waterproof_style_requires_review), merged.index
    ) & ~boots_mask
    if outsourced_sizing_event:
        waterproof_context = merged.apply(
            lambda row: " ".join(filter(None, [
                clean_text(row.get("Product Name", "")),
                clean_text(row.get("Original Line Item", "")),
                clean_text(row.get("Product Category", "")),
                clean_text(row.get("Product Aliases", "")),
                clean_text(row.get("Style Number", "")),
            ])),
            axis=1,
        )
        merged["_Waterproof Decoration Review"] = (
            forced_waterproof_style_mask
            | (
                _boolean_mask(
                    waterproof_context.map(waterproof_apparel_requires_review), merged.index
                )
                & ~always_embroidery_mask
                & ~boots_mask
            )
        )
    else:
        # These exact waterproof styles are always reviewed, even in Standard
        # Orchid Workflow and general sales reports.
        merged["_Waterproof Decoration Review"] = forced_waterproof_style_mask
    merged["_Mandatory Note Review"] = (
        _boolean_mask(merged["_Mandatory Note Review"], merged.index)
        | _boolean_mask(merged["_Waterproof Decoration Review"], merged.index)
    )

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
        if always_embroidery_rain_style(style):
            deco_type = "Embroidery"
            deco_location_raw = deco_location_raw or LEFT_CHEST
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
        # "Color is not required" means a blank Shopify color must not block
        # purchasing.  It does *not* mean a valid color selected by the
        # customer should be discarded.  This distinction is essential for
        # boots: some styles have no color option, while A64CQ, for example,
        # must keep Black Full Grain and Medium Brown Full Grain on separate
        # PO lines.  Shopify variant parsing already prevents size/width text
        # from reaching this field as a false color.
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
        elif is_internal_purchase_vendor(vendor):
            # “Orchid Uniforms & Apparel” is where these non-outsourced boots
            # will be received. It is not the supplier they are ordered from.
            reasons.append(INTERNAL_VENDOR_REVIEW_REASON)
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
        waterproof_decoration_review = bool(row.get("_Waterproof Decoration Review", False))
        if waterproof_decoration_review:
            forced_waterproof_style = waterproof_style_number(style)
            waterproof_prompt = (
                f"Waterproof style {forced_waterproof_style}: Confirm the decoration method before purchasing."
                if forced_waterproof_style
                else "Waterproof/rainwear item: Decoration Optional."
            )
            purchase_instructions = (
                f"{purchase_instructions} | {waterproof_prompt}"
                if purchase_instructions else waterproof_prompt
            )
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
                "Yes" if (
                    resolve_never_outsource(
                        row.get("Never Outsource", ""), description,
                        purchase_rules["Product Category"], style,
                    )
                    or vendor_never_outsource(vendor)
                ) else "No"
            ),
            "Purchase Vendor": vendor,
            "Decoration Type": deco_type,
            "Decoration Location": deco_location,
            "Decoration Placement Instructions": placement_instructions,
            "Decoration Color": deco_color,
            "Product #": style,
            "Vendor Product #": clean_text(row.get("Vendor Product #", "")),
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
            "Product Aliases": clean_text(row.get("Product Aliases", "")),
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
            "Source ID": clean_text(row.get("Source ID", "")),
            "Source Occurrence": int(row.get("Source Occurrence", 1) or 1),
            "Original Quantity": int(row.get("Original Quantity", qty) or qty),
            "Original Order Number": clean_text(row.get("Original Order Number", row.get("Order Number", ""))),
            "Original Company": clean_text(row.get("Original Company", row.get("Company", ""))),
            "Original Employee Name": clean_text(row.get("Original Employee Name", row.get("Employee Name", ""))),
            "Original Order Date": clean_text(row.get("Original Order Date", row.get("Order Date", ""))),
            "Original Shopify SKU": clean_text(row.get("Original Shopify SKU", row.get("Shopify SKU", ""))),
            "Quantity Override Confirmed": "No",
        })

    detail = pd.DataFrame(rows, columns=ALL_COLUMNS)
    if not detail.empty:
        detail["Decision Key"] = detail.apply(unresolved_decision_key, axis=1)
    review = detail[detail["Review Status"].eq("Needs Review")].copy() if not detail.empty else pd.DataFrame(columns=ALL_COLUMNS)

    excluded_rows = []
    for _, row in raw.iterrows():
        item = clean_text(row.get("Lineitem name", ""))
        variant_values = (
            row.get("Variant Color", ""), row.get("Variant Size", ""),
            row.get("Variant Option 1", ""), row.get("Variant Option 2", ""),
            row.get("Variant Option 3", ""), row.get("Variant Title", ""),
        )
        if item and is_decoration_service(item, row.get("Lineitem sku", ""), *variant_values):
            qty = pd.to_numeric(row.get("Lineitem quantity", 0), errors="coerce")
            excluded_rows.append({
                "Order Number": clean_text(row.get("Name", "")),
                "Company": clean_text(row.get("Billing Company", "")),
                "Employee Name": clean_text(row.get("Billing Name", "")),
                "Service / Fee": decoration_charge_label(*variant_values, item),
                "Quantity": 0 if pd.isna(qty) else int(qty),
                "Shopify Order Notes": clean_text(row.get("Notes", "")),
            })
    excluded = pd.DataFrame(
        excluded_rows,
        columns=["Order Number", "Company", "Employee Name", "Service / Fee", "Quantity", "Shopify Order Notes"],
    )
    return detail, review, excluded, raw


def _revalidate_detail(
    detail: pd.DataFrame,
    allow_decoration_prompts: bool = True,
    allow_waterproof_prompts: bool = False,
) -> pd.DataFrame:
    """Recalculate blockers after carried-forward Review & Edit corrections.

    Product Master changes remain authoritative, while user-entered order-specific
    corrections such as size, color, product number, and include/exclude choices
    survive a Purchase Review regeneration.
    """
    if detail.empty:
        return detail.copy()
    result = detail.copy()
    scoped_note_flags = _boolean_mask(
        _mandatory_note_review_flags(result, allow_decoration_prompts=allow_decoration_prompts),
        result.index,
    )
    # Known waterproof styles always require an explicit decoration decision,
    # regardless of workflow mode. This is an event-level confirmation and does
    # not change the permanent Product Master decoration default.
    boots_mask = _boolean_mask(
        result.apply(
            lambda row: is_boots_product(
                row.get("Description", row.get("Product Name", "")),
                row.get("Product Category", ""),
                row.get("Product #", row.get("Style Number", "")),
            ),
            axis=1,
        ),
        result.index,
    )
    waterproof_style_mask = _boolean_mask(
        result.apply(
            lambda row: waterproof_style_requires_review(
                row.get("Product #", row.get("Style Number", ""))
            ),
            axis=1,
        ),
        result.index,
    ) & ~boots_mask
    scoped_note_flags = scoped_note_flags | waterproof_style_mask
    if allow_waterproof_prompts:
        waterproof_context = result.apply(
            lambda row: " ".join(filter(None, [
                clean_text(row.get("Description", row.get("Product Name", ""))),
                clean_text(row.get("Original Shopify Line", row.get("Original Line Item", ""))),
                clean_text(row.get("Product Category", "")),
                clean_text(row.get("Product Aliases", "")),
                clean_text(row.get("Product #", row.get("Style Number", ""))),
            ])),
            axis=1,
        )
        permanent_embroidery_mask = _boolean_mask(
            result.apply(
                lambda row: always_embroidery_rain_style(
                    row.get("Product #", row.get("Style Number", ""))
                ),
                axis=1,
            ),
            result.index,
        )
        scoped_note_flags = scoped_note_flags | (
            _boolean_mask(
                waterproof_context.map(waterproof_apparel_requires_review), result.index
            )
            & ~permanent_embroidery_mask
            & ~boots_mask
        )
    for index, row in result.iterrows():
        style = clean_text(row.get("Product #", ""))
        description = clean_text(row.get("Description", ""))
        color = clean_text(row.get("Garment Color", ""))
        size = normalize_size(row.get("Size", ""))
        vendor = clean_text(row.get("Purchase Vendor", ""))
        deco_type = clean_text(row.get("Decoration Type", ""))
        deco_location_raw = clean_text(row.get("Decoration Location", ""))
        if always_embroidery_rain_style(style):
            deco_type = "Embroidery"
            deco_location_raw = deco_location_raw or LEFT_CHEST
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
        # Optional color values remain visible and continue to separate
        # purchasing rows.  Only the *requirement* is optional here; do not
        # erase an actual Shopify color selection.
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
        is_boots = is_boots_product(description, row.get("Product Category", ""), style)
        forced_waterproof_style = "" if is_boots else waterproof_style_number(style)
        waterproof_decoration_review = bool(
            forced_waterproof_style
            or (
                allow_waterproof_prompts
                and not is_boots
                and not always_embroidery_rain_style(style)
                and waterproof_apparel_requires_review(" ".join(filter(None, [
                    description, clean_text(row.get("Original Shopify Line", "")),
                    clean_text(row.get("Product Category", "")),
                    clean_text(row.get("Product Aliases", "")), style,
                ])))
            )
        )
        if forced_waterproof_style:
            waterproof_prompt = (
                f"Waterproof style {forced_waterproof_style}: Confirm the decoration method before purchasing."
            )
        else:
            waterproof_prompt = "Waterproof/rainwear item: Decoration Optional."
        if waterproof_decoration_review and waterproof_prompt.casefold() not in purchase_instructions.casefold():
            purchase_instructions = (
                f"{purchase_instructions} | {waterproof_prompt}"
                if purchase_instructions else waterproof_prompt
            )
        mandatory_note_review = bool(scoped_note_flags.at[index])
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


def _apply_previous_review_edits(
    detail: pd.DataFrame,
    previous_review_path: Path | None,
    allow_decoration_prompts: bool = True,
    allow_waterproof_prompts: bool = False,
) -> pd.DataFrame:
    """Carry prior review choices only onto the exact same Shopify source line.

    Professional 4.8.96 and later persist an immutable Source ID on every garment line.
    Generated Line IDs are never used as the primary replay key.  Older packets
    without Source IDs receive a conservative compatibility path: the original
    order/customer/date/line/notes signature must be unique in both workbooks,
    and identity-sensitive fields such as quantity and customer are not copied.
    """
    if detail.empty or not previous_review_path:
        return detail
    previous = Path(previous_review_path).expanduser()
    if not previous.exists():
        return detail
    try:
        from modules.xlsx_reader import read_table
        edits = read_table(previous, "Review & Edit", {"Line ID"})
        previous_source_rows = read_table(previous, "All PO Lines", {"Line ID"})
    except Exception:
        return detail
    if not edits:
        return detail

    result = detail.copy()
    previous_source_by_line = {
        clean_text(row.get("Line ID", "")): row
        for row in previous_source_rows
        if clean_text(row.get("Line ID", ""))
    }

    new_by_source_id: dict[str, list[object]] = {}
    new_by_legacy_signature: dict[str, list[object]] = {}
    new_by_manual_correction_signature: dict[str, list[object]] = {}
    for index, row in result.iterrows():
        source_id = clean_text(row.get("Source ID", ""))
        if source_id:
            new_by_source_id.setdefault(source_id, []).append(index)
        legacy = strict_legacy_signature(row)
        if legacy:
            new_by_legacy_signature.setdefault(legacy, []).append(index)
        correction = manual_entry_correction_signature(row)
        if correction:
            new_by_manual_correction_signature.setdefault(correction, []).append(index)

    previous_legacy_counts: dict[str, int] = {}
    previous_manual_correction_counts: dict[str, int] = {}
    for row in previous_source_rows:
        legacy = strict_legacy_signature(row)
        if legacy:
            previous_legacy_counts[legacy] = previous_legacy_counts.get(legacy, 0) + 1
        correction = manual_entry_correction_signature(row)
        if correction:
            previous_manual_correction_counts[correction] = previous_manual_correction_counts.get(correction, 0) + 1

    # Product Master routing is calculated before prior event edits are carried
    # over. Keep that fresh result so an older hidden Do Not Outsource value
    # cannot undo a newer Never Outsource setting.
    current_do_not_outsource = result["Do Not Outsource"].copy()
    carry_fields = {
        "Include", "Product #", "Description", "Garment Color", "Size", "Quantity",
        "Company", "Employee Name", "Order Number", "Decoration Decision",
        "Decoration Location", "Decoration Placement Instructions",
        "Quantity Override Confirmed",
    }
    permanent_override_fields = {
        "Purchase Vendor", "Decoration Type", "Decoration Location",
        "Decoration Placement Instructions", "Decoration Color",
    }
    identity_sensitive_fields = {
        "Quantity", "Company", "Employee Name", "Order Number", "Quantity Override Confirmed",
    }
    # If the Shopify product title/style itself was corrected after a review,
    # retain only order-specific choices. Product identity, catalog defaults,
    # and routing must be rebuilt from the corrected source and live Master.
    source_correction_fields = {
        "Include", "Quantity", "Company", "Employee Name", "Order Number",
        "Decoration Decision", "Decoration Location",
        "Decoration Placement Instructions", "Quantity Override Confirmed",
    }
    # RC13 and earlier could parse ``3X TL`` as color ``TL`` and size ``3XL``.
    # Do not replay those stale parser-owned color/size cells from an earlier
    # Review & Edit sheet after the parser has recovered the explicit Tall size.
    repaired_tall_indexes = {
        index
        for index, row in result.iterrows()
        if re.search(r"\b[2-8]X(?:L)?\s+TL\b", clean_text(row.get("Original Shopify Line", "")), re.I)
        and normalize_size(row.get("Size", "")).endswith("XLT")
    }
    # A regenerated event must also replace a stale shoe-width value from an
    # older Review & Edit sheet. Earlier builds could retain ``Regular`` as
    # the Size for Shopify's ``11 / Regular`` boot option. The current parser
    # now has the authoritative *size and width* in the fresh source line, so
    # never replay an old Size or Garment Color over that repair.
    repaired_footwear_indexes = {
        index
        for index, row in result.iterrows()
        if bool(re.fullmatch(
            r"(?:[4-9]|1[0-6])(?:\.5)?\s*/\s*"
            r"(?:medium|narrow|wide|extra wide|ee|eee)",
            normalize_size(row.get("Size", "")),
            re.I,
        ))
    }
    previously_ready_indexes: set[object] = set()
    explicit_do_not_outsource_indexes: set[object] = set()

    for edit in edits:
        previous_line_id = clean_text(edit.get("Line ID", ""))
        source_row = previous_source_by_line.get(previous_line_id, {})
        old_source_id = clean_text(edit.get("Source ID", "")) or clean_text(source_row.get("Source ID", ""))
        verified_source_id = False
        corrected_source_match = False
        index = None

        if old_source_id:
            candidates = new_by_source_id.get(old_source_id, [])
            if len(candidates) == 1:
                index = candidates[0]
                verified_source_id = True
        else:
            legacy = strict_legacy_signature(source_row or edit)
            candidates = new_by_legacy_signature.get(legacy, []) if legacy else []
            # Never guess among duplicate legacy rows.  A packet without Source
            # IDs can safely retain non-identity edits only when the match is
            # unique on both sides.
            if legacy and previous_legacy_counts.get(legacy, 0) == 1 and len(candidates) == 1:
                index = candidates[0]

        # A user can correct a manually entered Shopify style/title after a
        # completed review. The new Source ID is rightly different, but a
        # unique match on the remaining immutable order fields can retain the
        # explicit order decision. Never copy old product identity through this
        # path; the corrected source and Product Master remain authoritative.
        if index is None:
            correction = manual_entry_correction_signature(source_row or edit)
            candidates = new_by_manual_correction_signature.get(correction, []) if correction else []
            if (
                correction
                and previous_manual_correction_counts.get(correction, 0) == 1
                and len(candidates) == 1
            ):
                index = candidates[0]
                corrected_source_match = True

        if index is None:
            continue

        was_ready = clean_text(edit.get("Review Status", "")).casefold() == "ready"
        if was_ready:
            previously_ready_indexes.add(index)

        fields = set(source_correction_fields) if corrected_source_match else set(carry_fields)
        if was_ready and not corrected_source_match:
            fields.update(permanent_override_fields)
        if not verified_source_id and not corrected_source_match:
            fields.difference_update(identity_sensitive_fields)

        # ``clean_text`` intentionally turns typographic dashes into a normal
        # hyphen for Excel/PDF safety. Compare the normalized value so a saved
        # "Do Not Outsource — Ship to Orchid" decision survives regeneration.
        if clean_text(edit.get("Decoration Decision", "")).casefold() == "do not outsource - ship to orchid":
            explicit_do_not_outsource_indexes.add(index)

        for field in fields:
            if field not in edit:
                continue
            if index in (repaired_tall_indexes | repaired_footwear_indexes) and field in {"Garment Color", "Size"}:
                continue
            # A regenerated workbook must pick up a corrected Product Master
            # description.  Review & Edit contains the original value for
            # every displayed row, so replaying it blindly can freeze a stale
            # catalog error forever.  Preserve Description only when the user
            # actually changed it from the prior All PO Lines baseline.
            if (
                field == "Description"
                and source_row
                and clean_text(edit.get("Description", ""))
                == clean_text(source_row.get("Description", ""))
            ):
                continue
            value = edit.get(field, "")
            if field == "Quantity":
                try:
                    value = int(float(value or 0))
                except (TypeError, ValueError):
                    value = 0
            if field in {"Include", "Do Not Outsource", "Quantity", "Quantity Override Confirmed"} or clean_text(value):
                result.at[index, field] = value

    result["Do Not Outsource"] = current_do_not_outsource
    for index in explicit_do_not_outsource_indexes:
        result.at[index, "Do Not Outsource"] = "Yes"

    result = _revalidate_detail(
        result,
        allow_decoration_prompts=allow_decoration_prompts,
        allow_waterproof_prompts=allow_waterproof_prompts,
    )

    # Revalidation intentionally recreates hard blockers. Preserve a user's
    # completed manual decision only when no hard blocker remains.
    manual_only_reason = "customer decision required"
    for index in previously_ready_indexes:
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
        valid_decisions = {
            "follow note as written", "keep product master default", "no decoration",
            "embroidery", "screen print", "sew on patch", "hemming / alteration",
            "do not outsource - ship to orchid",
        }
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
        "Product Category": 22, "Product Aliases": 28, "Requires Size": 14, "Requires Color": 14, "Requires Decoration": 18,
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

def _employee_order_key(value: object) -> str:
    return clean_text(value).lstrip("#").casefold()


def _build_employee_totals(
    raw: pd.DataFrame,
    included_order_keys: set[str] | None = None,
) -> pd.DataFrame:
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
        if (
            included_order_keys is not None
            and _employee_order_key(order_number) not in included_order_keys
        ):
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
    active_packet_id: str = "",
    decoration_fulfillment: str = STANDARD_ORCHID_WORKFLOW,
    regenerate_helper: Path | None = None,
    previous_review_path: Path | None = None,
    normalized_orders: pd.DataFrame | None = None,
    parsed_orders: pd.DataFrame | None = None,
) -> dict:
    report_mode = normalize_report_mode(report_mode)
    decoration_fulfillment = normalize_decoration_fulfillment(decoration_fulfillment)
    # A regeneration must keep the fulfillment route selected when the event was
    # created. Older saved state files can omit this field and otherwise fall
    # back to Standard Orchid Workflow, silently removing embroidery from an
    # Entire Order Outsourced report. The previous workbook is authoritative.
    if previous_review_path:
        previous_path = Path(previous_review_path).expanduser()
        if previous_path.exists():
            from modules.xlsx_reader import load_decoration_fulfillment

            decoration_fulfillment = load_decoration_fulfillment(previous_path)
    event_name = clean_text(event_name) if report_mode == UNIFORM_SIZING_EVENT else ""
    product_master_path = assert_live_product_master(Path(product_master_path).expanduser().resolve())
    master_signature = product_master_signature(product_master_path)
    # Always reload the active Product Master from disk immediately before building.
    detail, review, excluded, raw = _build_review_data(
        Path(shopify_csv_path), product_master_path, decoration_fulfillment,
        report_mode=report_mode,
        normalized_orders=normalized_orders, parsed_orders=parsed_orders,
    )
    now = datetime.now()
    stamp = now.strftime("%Y%m%d_%H%M%S_%f")
    output_dir = Path(reports_root) / "Review Workbooks"
    output_dir.mkdir(parents=True, exist_ok=True)
    event_part = f"__{safe_filename(event_name)}" if event_name else ""
    output_path = output_dir / f"Orchid_Purchase_Review{event_part}__{stamp}.xlsx"
    regeneration_request_path = output_path.with_suffix(".orchidregen")

    # Freeze the selected order CSV, live Product Master, Never Outsource
    # overrides, and immutable Shopify source ledger before any prior review
    # decisions are replayed. The final PDF stage must verify this exact bundle.
    packet_lock = create_packet_lock(
        source_csv_path=Path(shopify_csv_path),
        product_master_path=product_master_path,
        detail=detail,
        report_mode=report_mode,
        event_name=event_name,
        decoration_fulfillment=decoration_fulfillment,
    )

    detail = _apply_previous_review_edits(
        detail,
        previous_review_path,
        allow_decoration_prompts=decoration_note_prompts_enabled(report_mode, decoration_fulfillment),
        allow_waterproof_prompts=waterproof_decoration_prompts_enabled(report_mode, decoration_fulfillment),
    )

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
        "Decoration Decision", "Source ID", "Source Occurrence", "Original Quantity",
        "Original Order Number", "Original Company", "Original Employee Name", "Original Order Date",
        "Original Shopify SKU", "Quantity Override Confirmed",
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
    for technical_name in (
        "Source ID", "Source Occurrence", "Original Quantity", "Original Order Number",
        "Original Company", "Original Employee Name", "Original Order Date", "Original Shopify SKU",
        "Quantity Override Confirmed",
    ):
        review_ws.set_column(review_col[technical_name], review_col[technical_name], None, None, {"hidden": True})
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
        included_employee_order_keys = {
            _employee_order_key(row.get("Order Number", ""))
            for _, row in detail.iterrows()
            if clean_text(row.get("Include", "Yes")).casefold() in {"yes", "y", "true", "1", "include"}
            and pd.to_numeric(row.get("Quantity", 0), errors="coerce") > 0
            and _employee_order_key(row.get("Order Number", ""))
        }
        # Final Sale Boot is a stock-sale accounting line, not a purchase line.
        # It must still remain in Employee Totals so the $50 is visible for the
        # employee/event budget even when an order contains no other item.
        included_employee_order_keys.update({
            _employee_order_key(row.get("Name", ""))
            for _, row in raw.iterrows()
            if is_final_sale_accounting_product(row.get("Lineitem name", ""))
            and _employee_order_key(row.get("Name", ""))
        })
        employee_totals = _build_employee_totals(raw, included_employee_order_keys)
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
    system_info.write(5, 0, "Routing Policy Version")
    system_info.write(5, 1, ROUTING_POLICY_VERSION)
    system_info.write(6, 0, "Workbook Generated")
    system_info.write(6, 1, now.strftime("%Y-%m-%d %I:%M:%S %p"))
    system_info.write(7, 0, "Decoration Fulfillment")
    system_info.write(7, 1, decoration_fulfillment)
    system_info.write(8, 0, "Purchase Order Mode")
    system_info.write(8, 1, report_mode)
    system_info.write(9, 0, "Event Name")
    system_info.write(9, 1, event_name)
    # The source hash proves which CSV was used.  The event ID additionally
    # proves which *event instance* owns this review, so a workbook from an
    # earlier run of the same-named event cannot be reopened as the current one.
    system_info.write(10, 0, "Active Packet ID")
    system_info.write(10, 1, clean_text(active_packet_id))
    system_info.write(11, 0, "Packet Lock ID")
    system_info.write(11, 1, str(packet_lock.get("lock_id", "")))
    system_info.write(12, 0, "Packet Lock Manifest Path")
    system_info.write(12, 1, str(packet_lock.get("manifest_path", "")))
    system_info.write(13, 0, "Packet Lock Manifest SHA256")
    system_info.write(13, 1, str(packet_lock.get("manifest_sha256", "")))
    system_info.write(14, 0, "Locked Source Lines")
    system_info.write(14, 1, int(packet_lock.get("source_line_count", 0)))
    system_info.write(15, 0, "Locked Source Quantity")
    system_info.write(15, 1, int(packet_lock.get("source_quantity", 0)))
    system_info.write(16, 0, "Product Master Routing SHA256")
    system_info.write(16, 1, str(master_signature.get("routing_sha256", "")))
    system_info.write(17, 0, "Never Outsource Override SHA256")
    system_info.write(17, 1, str(master_signature.get("override_sha256", "")))
    system_info.hide()

    workbook.close()

    if regenerate_helper:
        request_payload = {
            "source_csv": str(Path(shopify_csv_path).expanduser().resolve()),
            "report_mode": report_mode,
            "event_name": event_name,
            "active_packet_id": clean_text(active_packet_id),
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
        "packet_lock_id": packet_lock.get("lock_id", ""),
        "packet_lock_manifest": packet_lock.get("manifest_path", ""),
        "packet_lock_manifest_sha256": packet_lock.get("manifest_sha256", ""),
        "report_mode": report_mode,
        "event_name": event_name,
        "decoration_fulfillment": decoration_fulfillment,
    }
