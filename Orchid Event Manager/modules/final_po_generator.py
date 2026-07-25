from __future__ import annotations

from collections import OrderedDict
from datetime import datetime
from io import BytesIO
from pathlib import Path
import json
import shutil
import sys

import pandas as pd
from PIL import Image as PILImage
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import (
    BaseDocTemplate, CondPageBreak, Frame, Image, NextPageTemplate, PageBreak, PageTemplate,
    Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle,
)
from reportlab.graphics.shapes import Drawing, Rect

from modules.purchase_order_generator import clean_text, normalize_size, safe_filename, size_sort_key, split_embedded_size
from modules.blank_garment_rules import BLANK_DECORATION_LABEL, is_blank_decoration
from modules.decoration_locations import (
    LEFT_CHEST,
    OTHER_CUSTOM,
    is_custom_location,
    location_sort_key,
    normalize_decoration_location,
)
from modules.purchase_rules import normalize_bool, row_rules
from modules.report_modes import GENERAL_SALES_PERIOD, UNIFORM_SIZING_EVENT, normalize_report_mode
from modules.xlsx_reader import (
    load_decoration_fulfillment, load_event_name, load_po_numbers,
    load_report_mode, load_review_lines,
)
from modules.note_rules import (
    decoration_note_requires_review, decoration_note_targets_line,
    decoration_target_styles, extract_purchase_instructions,
)
from modules.employee_totals import load_employee_totals, employee_totals_summary
from modules.internal_services import (
    is_in_house_decoration, is_in_house_service_product, is_internal_service_style,
    vendor_po_decoration_type,
)
from modules.decoration_fulfillment import is_entire_order_outsourced, is_outsourced_decoration
from modules.outsource_rules import vendor_never_outsource
from modules.product_resolver import resolve_product
from modules.pdf_page_numbers import PageNumberCanvas
from modules.job_logo_image import remove_outer_near_white_background
from modules.packet_lock import copy_lock_bundle, sha256_file, verify_packet_lock

PURPLE = colors.HexColor("#5B2AA8")
PURPLE_DARK = colors.HexColor("#3D176F")
PURPLE_LIGHT = colors.HexColor("#F2ECFB")
BORDER = colors.HexColor("#D7C8EE")
TEXT = colors.HexColor("#201A2D")
MUTED = colors.HexColor("#6F667A")
ROW_ALT = colors.HexColor("#FAF7FE")
YELLOW = colors.HexColor("#FFF2CC")


def _resource_path(*parts: str) -> Path:
    project_root = Path(__file__).resolve().parent.parent
    base = Path(getattr(sys, "_MEIPASS", project_root))
    return base.joinpath(*parts)


def _official_logo_flowable(max_width: float = 1.45 * inch, max_height: float = 0.62 * inch):
    """Return the official Orchid logo sized proportionally for PDF headers."""
    path = _resource_path("assets", "orchid_logo.png")
    if not path.exists():
        return Spacer(1, 1)
    try:
        image = Image(str(path))
        width, height = float(image.imageWidth), float(image.imageHeight)
        scale = min(max_width / width, max_height / height)
        image.drawWidth = width * scale
        image.drawHeight = height * scale
        image.hAlign = "CENTER"
        return image
    except Exception:
        return Spacer(1, 1)


def _job_logo_flowable(path: Path, max_width: float, max_height: float):
    """Return a PDF logo with blank source-canvas padding removed when safe.

    Customer artwork commonly arrives as a logo centered in a much larger white
    or transparent canvas.  The source file must remain untouched, but using
    that empty canvas for PDF sizing makes the visible mark look too small.
    This prepares an in-memory PNG only for the report and falls back to the
    original image if Pillow cannot safely process it.
    """
    source_path = Path(path)
    try:
        with PILImage.open(source_path) as source:
            image = source.convert("RGBA")
            image, _removed_white_canvas = remove_outer_near_white_background(image)
            width, height = image.size
            full_area = max(width * height, 1)
            boxes = []

            # Transparent outer padding is the most reliable crop boundary.
            alpha_box = image.getchannel("A").getbbox()
            if alpha_box:
                alpha_area = max((alpha_box[2] - alpha_box[0]) * (alpha_box[3] - alpha_box[1]), 0)
                if alpha_area < full_area * 0.98:
                    boxes.append(alpha_box)

            # Also support opaque files with a white (or nearly white) canvas.
            # The threshold leaves light logo colors alone while ignoring the
            # white export margin around a typical embroidery or screen-print
            # design file.
            white_background = PILImage.new("RGBA", image.size, (255, 255, 255, 255))
            flattened = PILImage.alpha_composite(white_background, image).convert("L")
            ink_mask = flattened.point(lambda value: 255 if value < 245 else 0)
            ink_box = ink_mask.getbbox()
            if ink_box:
                ink_area = max((ink_box[2] - ink_box[0]) * (ink_box[3] - ink_box[1]), 0)
                if ink_area < full_area * 0.98:
                    boxes.append(ink_box)

            if boxes:
                left = min(box[0] for box in boxes)
                top = min(box[1] for box in boxes)
                right = max(box[2] for box in boxes)
                bottom = max(box[3] for box in boxes)
                padding = max(2, round(max(width, height) * 0.035))
                crop_box = (
                    max(0, left - padding), max(0, top - padding),
                    min(width, right + padding), min(height, bottom + padding),
                )
                crop_area = max((crop_box[2] - crop_box[0]) * (crop_box[3] - crop_box[1]), 0)
                if crop_area < full_area * 0.98:
                    image = image.crop(crop_box)

            buffer = BytesIO()
            image.save(buffer, format="PNG")
            buffer.seek(0)
            logo = Image(buffer)
            # ReportLab reads the stream later while building the document.
            # Retain it on the flowable for that entire build.
            logo._orchid_job_logo_buffer = buffer
    except Exception:
        logo = Image(str(source_path))

    width, height = float(logo.imageWidth), float(logo.imageHeight)
    scale = min(max_width / width, max_height / height)
    logo.drawWidth = width * scale
    logo.drawHeight = height * scale
    logo.hAlign = "CENTER"
    return logo

def _event_logo_path(workbook_path: Path, event_name: str = "") -> Path | None:
    """Find an event-specific outsourced job logo stored beside the workbook."""
    stem = Path(workbook_path).stem
    parent = Path(workbook_path).parent
    for suffix in (".png", ".jpg", ".jpeg", ".webp"):
        if event_name:
            stable = parent / f"{safe_filename(event_name)}__Outsourced_Job_Logo{suffix}"
            if stable.exists():
                return stable
        candidate = parent / f"{stem}__Outsourced_Job_Logo{suffix}"
        if candidate.exists():
            return candidate
    return None

DETAIL_COLUMNS = [
    "Source ID", "Line ID", "Original Quantity", "Original Order Number", "Original Company",
    "Original Employee Name", "Quantity Override Confirmed",
    "Vendor", "Decoration Type", "Decoration Location", "Decoration Placement Instructions", "Decoration Color", "Style Number",
    "Product Name", "Garment Color", "Size", "Quantity", "Company",
    "Employee Name", "Order Number", "Purchase Instructions",
    "Operational Decoration Type", "Operational Decoration Location",
    "Operational Placement Instructions", "Operational Decoration Color",
    "Shopify Order Notes", "Shopify Line Notes", "Do Not Outsource",
]

COMPACT_COLUMNS = [
    "Decoration Type", "Decoration Location", "Decoration Placement Instructions", "Decoration Color", "Style Number", "Product Name",
    "Garment Color", "Size", "Quantity", "Customer(s)", "Purchase Instructions",
]


def _yes(value) -> bool:
    return clean_text(value).casefold() in {"yes", "y", "true", "1", "include"}


def _quantity(value) -> int:
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return 0


def _display_decoration_type(value: object) -> str:
    text = clean_text(value)
    lowered = text.casefold()
    if "screen" in lowered:
        return "Screen Printing"
    if lowered == "embroidery":
        return "Embroidery"
    if is_blank_decoration(text) or not text:
        return "Blank Garments"
    return text



def _canonical_report_type(value: object) -> str:
    text = clean_text(value).casefold()
    if "screen" in text:
        return "screen printing"
    if "embroider" in text:
        return "embroidery"
    if "blank" in text or text in {"none", "no decoration"}:
        return "blank garments"
    if "combined" in text or "general sales" in text:
        return "combined vendor order"
    return text

def _decoration_type_sort(value: object) -> tuple[int, str]:
    text = clean_text(value).casefold()
    if text == "embroidery":
        return (0, text)
    if "screen" in text:
        return (1, text)
    if is_blank_decoration(text):
        return (2, text)
    return (3, text)


def _color_section_label(decoration_type: object, decoration_color: object) -> str:
    deco = clean_text(decoration_type)
    color = clean_text(decoration_color)
    if is_blank_decoration(deco):
        return "Blank Garments"
    label = "Ink" if "screen" in deco.casefold() else "Thread"
    return f"{color or 'Unassigned'} {label}"


def _do_not_outsource(value: object) -> bool:
    return clean_text(value).casefold() in {"yes", "y", "true", "1"}


def _master_requires_never_outsource(
    master: pd.DataFrame | None,
    style: object,
    product: object,
    garment_color: object,
) -> bool:
    """Return the live Product Master routing decision for one garment.

    A reopened archive keeps its historical review workbook so users can add a
    logo or regenerate reports without rebuilding the order.  The workbook's
    saved routing value can therefore be older than the Product Master.  Final
    reports must still honor a current *Never Outsource* selection, so consult
    the live Master at report-generation time without modifying the archive.
    """
    if master is None or master.empty:
        return False
    try:
        resolution = resolve_product(
            style,
            product,
            garment_color,
            master,
            master_prepared=True,
        )
    except Exception:
        # A malformed or unavailable live Master must not prevent a user from
        # opening an older archive; retain its saved routing in that rare case.
        return False
    return bool(resolution.matched and resolution.never_outsource)


def _prepare_lines(workbook_path: Path, decoration_fulfillment: str = ""):
    records = load_review_lines(workbook_path)
    # Final reports must be reproducible from the reviewed workbook. Never
    # silently reread a different live Product Master here: doing so can reroute
    # a completed order after Purchase Review. Active events already regenerate
    # the review when Product Master changes; archived packets retain their
    # frozen, reviewed Do Not Outsource decisions.
    live_master = pd.DataFrame()
    ready_rows = []
    review_rows = []
    non_included_rows = []

    for record in records:
        if not _yes(record.get("Include", "")):
            continue
        vendor = clean_text(record.get("Purchase Vendor", ""))
        original_deco_type = clean_text(record.get("Decoration Type", ""))
        original_deco_location = clean_text(record.get("Decoration Location", ""))
        original_placement = clean_text(record.get("Decoration Placement Instructions", ""))
        original_deco_color = clean_text(record.get("Decoration Color", ""))
        deco_type = original_deco_type
        deco_location_raw = original_deco_location
        placement_instructions = original_placement
        deco_color = original_deco_color
        style = clean_text(record.get("Product #", ""))
        product = clean_text(record.get("Description", ""))
        original_line = clean_text(record.get("Original Shopify Line", ""))
        garment_color = clean_text(record.get("Garment Color", record.get("Color", "")))
        do_not_outsource = (
            _do_not_outsource(record.get("Do Not Outsource", "No"))
            # Product-level routing in the live Master overrides a stale value
            # saved inside an archived review workbook.
            or _master_requires_never_outsource(live_master, style, product, garment_color)
            # Vendor-level Never Outsource defaults are meaningful in both
            # workflows. A Berne screen-print line, for example, must stay at
            # Orchid even when the event otherwise uses standard routing.
            or vendor_never_outsource(vendor)
        )
        if is_internal_service_style(style) or is_in_house_service_product(
            product, original_line, deco_type, style_number=style, garment_color=garment_color
        ):
            # Service-only fees are internal Orchid work and never belong on a
            # garment vendor purchase order or manual garment purchase list.
            continue
        if is_in_house_decoration(deco_type) and not do_not_outsource:
            # A real garment with in-house work must still be purchased, but the
            # vendor should receive it as a blank garment without internal notes.
            deco_type = vendor_po_decoration_type(deco_type)
            deco_location_raw = ""
            placement_instructions = ""
            deco_color = ""
        garment_color, size = split_embedded_size(garment_color, record.get("Size", ""))
        qty = _quantity(record.get("Quantity", 0))
        rules = row_rules({
            "Product Name": product,
            "Style Number": style,
            "Product Category": record.get("Product Category", ""),
            "Requires Size": record.get("Requires Size", ""),
            "Requires Color": record.get("Requires Color", ""),
            "Requires Decoration": record.get("Requires Decoration", ""),
            "Decoration Type": deco_type,
        })
        requires_size = normalize_bool(rules["Requires Size"], True)
        requires_color = normalize_bool(rules["Requires Color"], True)
        requires_decoration = normalize_bool(rules["Requires Decoration"], True)
        deco_location = normalize_decoration_location(
            deco_location_raw, deco_type, requires_decoration=requires_decoration
        )
        if not requires_decoration and not deco_type:
            deco_type = BLANK_DECORATION_LABEL
            deco_color = ""

        reasons = []
        if not vendor:
            reasons.append("Missing purchase vendor")
        if requires_decoration and not deco_type:
            reasons.append("Missing decoration type")
        if requires_decoration and deco_type and not is_blank_decoration(deco_type) and not deco_color:
            reasons.append("Missing decoration color")
        if requires_decoration and deco_location == OTHER_CUSTOM:
            reasons.append("Custom decoration location is incomplete")
        if not style and not product:
            reasons.append("Missing product number and description")
        if requires_color and not garment_color:
            reasons.append("Missing garment color")
        if requires_size and not size:
            reasons.append("Missing size")
        if qty <= 0:
            reasons.append("Quantity must be greater than zero")

        row = {
            "Source ID": clean_text(record.get("Source ID", "")),
            "Line ID": clean_text(record.get("Line ID", "")),
            "Original Quantity": _quantity(record.get("Original Quantity", record.get("Quantity", 0))),
            "Original Order Number": clean_text(record.get("Original Order Number", record.get("Order Number", ""))),
            "Original Company": clean_text(record.get("Original Company", record.get("Company", ""))),
            "Original Employee Name": clean_text(record.get("Original Employee Name", record.get("Employee Name", ""))),
            "Quantity Override Confirmed": clean_text(record.get("Quantity Override Confirmed", "No")) or "No",
            "Vendor": vendor,
            "Decoration Type": deco_type,
            "Decoration Location": deco_location,
            "Decoration Placement Instructions": placement_instructions,
            "Decoration Color": deco_color,
            "Style Number": style,
            "Product Name": product,
            "Garment Color": garment_color,
            "Size": size,
            "Quantity": qty,
            "Company": clean_text(record.get("Company", "")),
            "Employee Name": clean_text(record.get("Employee Name", "")),
            "Order Number": clean_text(record.get("Order Number", "")),
            "Purchase Instructions": (
                clean_text(record.get("Purchase Instructions", ""))
                or extract_purchase_instructions(record.get("Shopify Order Notes", ""))
            ),
            "Operational Decoration Type": original_deco_type or deco_type,
            "Operational Decoration Location": normalize_decoration_location(
                original_deco_location, original_deco_type or deco_type,
                requires_decoration=not is_blank_decoration(original_deco_type or deco_type),
            ),
            "Operational Placement Instructions": original_placement,
            "Operational Decoration Color": original_deco_color,
            "Shopify Order Notes": clean_text(record.get("Shopify Order Notes", "")),
            "Shopify Line Notes": clean_text(record.get("Shopify Line Notes", "")),
            "Do Not Outsource": "Yes" if do_not_outsource else "No",
        }
        route_to_orchid = bool(
            do_not_outsource
            or (
                is_entire_order_outsourced(decoration_fulfillment)
                and not is_outsourced_decoration(original_deco_type, decoration_fulfillment)
            )
        )
        if reasons:
            review_rows.append({**row, "Review Reason": "; ".join(reasons)})
        elif route_to_orchid:
            # During outsourced packets every line that is not actually going to
            # the decorator belongs on the consolidated Ship-to-Orchid order.
            row["Decoration Type"] = original_deco_type
            row["Decoration Location"] = normalize_decoration_location(
                original_deco_location, original_deco_type,
                requires_decoration=not is_blank_decoration(original_deco_type),
            )
            row["Decoration Placement Instructions"] = original_placement
            row["Decoration Color"] = original_deco_color
            row["Do Not Outsource"] = "Yes"
            non_included_rows.append(row)
        else:
            if is_in_house_decoration(original_deco_type):
                row["Purchase Instructions"] = ""
            ready_rows.append(row)

    return (
        pd.DataFrame(ready_rows, columns=DETAIL_COLUMNS),
        pd.DataFrame(review_rows, columns=DETAIL_COLUMNS + ["Review Reason"]),
        pd.DataFrame(non_included_rows, columns=DETAIL_COLUMNS),
    )



def _output_identity(row: dict | pd.Series) -> str:
    source_id = clean_text(row.get("Source ID", ""))
    if source_id:
        return source_id
    line_id = clean_text(row.get("Line ID", ""))
    return f"LEGACY-LINE:{line_id}" if line_id else ""


def _validate_final_output_integrity(
    records: list[dict[str, object]],
    ready: pd.DataFrame,
    review: pd.DataFrame,
    non_included: pd.DataFrame,
) -> dict[str, int]:
    """Block report generation when a garment cannot be traced exactly once.

    This guard runs before any vendor or Non-Included PDF is written.  It proves
    that every included garment source line appears in exactly one output route,
    that Source IDs are unique, and that quantity/customer identity was not
    silently changed. A deliberate quantity edit is allowed only when the app
    recorded Quantity Override Confirmed.
    """
    expected: dict[str, dict[str, object]] = {}
    errors: list[str] = []
    for record in records:
        if not _yes(record.get("Include", "")):
            continue
        style = clean_text(record.get("Product #", ""))
        product = clean_text(record.get("Description", ""))
        original_line = clean_text(record.get("Original Shopify Line", ""))
        deco_type = clean_text(record.get("Decoration Type", ""))
        garment_color = clean_text(record.get("Garment Color", ""))
        if is_internal_service_style(style) or is_in_house_service_product(
            product, original_line, deco_type, style_number=style, garment_color=garment_color
        ):
            continue
        identity = _output_identity(record)
        if not identity:
            errors.append(f"A garment line has no Source ID or Line ID: {style or product or 'unknown product'}")
            continue
        if identity in expected:
            errors.append(f"Duplicate source identity in Purchase Review: {identity}")
            continue
        expected[identity] = record

        source_id = clean_text(record.get("Source ID", ""))
        if source_id:
            comparisons = (
                ("order number", record.get("Original Order Number", ""), record.get("Order Number", "")),
                ("company", record.get("Original Company", ""), record.get("Company", "")),
                ("employee", record.get("Original Employee Name", ""), record.get("Employee Name", "")),
            )
            for label, original, current in comparisons:
                original_text = clean_text(original)
                current_text = clean_text(current)
                if original_text and original_text != current_text:
                    errors.append(f"{identity}: {label} no longer matches its imported Shopify source.")
            original_qty = _quantity(record.get("Original Quantity", 0))
            current_qty = _quantity(record.get("Quantity", 0))
            confirmed = _yes(record.get("Quantity Override Confirmed", "No"))
            if original_qty > 0 and current_qty != original_qty and not confirmed:
                errors.append(
                    f"{identity}: quantity changed from {original_qty} to {current_qty} without an explicit override."
                )

    routed: dict[str, str] = {}
    routed_qty: dict[str, int] = {}
    for route_name, frame in (("Vendor Purchase Order", ready), ("Purchase Review", review), ("Non-Included", non_included)):
        if frame.empty:
            continue
        for _, row in frame.iterrows():
            identity = _output_identity(row)
            if not identity:
                errors.append(f"{route_name} contains a line without a source identity.")
                continue
            if identity in routed:
                errors.append(f"{identity} appears in both {routed[identity]} and {route_name}.")
                continue
            routed[identity] = route_name
            routed_qty[identity] = _quantity(row.get("Quantity", 0))

    missing = sorted(set(expected) - set(routed))
    unexpected = sorted(set(routed) - set(expected))
    if missing:
        errors.append(f"{len(missing)} included garment line(s) are missing from every purchasing output.")
    if unexpected:
        errors.append(f"{len(unexpected)} purchasing output line(s) do not belong to an included garment source.")
    for identity in sorted(set(expected) & set(routed)):
        current_qty = _quantity(expected[identity].get("Quantity", 0))
        if routed_qty.get(identity, 0) != current_qty:
            errors.append(
                f"{identity}: output quantity {routed_qty.get(identity, 0)} does not match review quantity {current_qty}."
            )

    if errors:
        preview = "\n".join(f"- {message}" for message in errors[:12])
        if len(errors) > 12:
            preview += f"\n- Plus {len(errors) - 12} additional integrity error(s)."
        raise RuntimeError(
            "Purchase-order integrity check failed. No PDFs were created.\n\n" + preview
        )


    return {
        "source_lines": len(expected),
        "vendor_lines": len(ready),
        "review_lines": len(review),
        "non_included_lines": len(non_included),
        "vendor_quantity": int(ready["Quantity"].sum()) if not ready.empty else 0,
        "non_included_quantity": int(non_included["Quantity"].sum()) if not non_included.empty else 0,
    }


def _money(value: object) -> float:
    try:
        if value is None or value == "":
            return 0.0
        if isinstance(value, str):
            value = value.replace("$", "").replace(",", "").strip()
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def _build_employee_totals_pdf(records: list[dict], pdf_path: Path, event_name: str) -> None:
    doc = SimpleDocTemplate(
        str(pdf_path), pagesize=landscape(letter), rightMargin=0.42 * inch,
        leftMargin=0.42 * inch, topMargin=0.38 * inch, bottomMargin=0.42 * inch,
        title=f"Employee Totals - {event_name}", author="Orchid Uniforms & Apparel",
    )
    styles = getSampleStyleSheet()
    title = ParagraphStyle("EmpTitle", parent=styles["Title"], fontName="Helvetica-Bold", fontSize=18, leading=21, textColor=PURPLE_DARK)
    sub = ParagraphStyle("EmpSub", parent=styles["Normal"], fontSize=8.5, leading=10, textColor=MUTED)
    body = ParagraphStyle("EmpBody", parent=styles["Normal"], fontSize=7.2, leading=8.5, textColor=TEXT)
    body_right = ParagraphStyle("EmpBodyRight", parent=body, alignment=TA_RIGHT)
    header = ParagraphStyle("EmpHeader", parent=body, fontName="Helvetica-Bold", textColor=colors.white)
    story = [_paragraph("EMPLOYEE TOTALS", title), _paragraph(event_name or "Uniform Sizing Event", sub), Spacer(1, 8)]
    rows = [[
        _paragraph("Company / Department", header), _paragraph("Employee", header),
        _paragraph("Order Number(s)", header), _paragraph("Order Total", header),
        _paragraph("Discount", header), _paragraph("Total", header),
    ]]
    for record in records:
        rows.append([
            _paragraph(record.get("company", ""), body),
            _paragraph(record.get("employee", ""), body),
            _paragraph(record.get("order_numbers", ""), body),
            _paragraph(f"${_money(record.get('order_total')):,.2f}", body_right),
            _paragraph(f"${_money(record.get('discount')):,.2f}", body_right),
            _paragraph(f"${_money(record.get('total')):,.2f}", body_right),
        ])
    grand_total = sum(_money(record.get("total")) for record in records)
    rows.append([
        "", "", "", "", _paragraph("Grand Total", header),
        _paragraph(f"${grand_total:,.2f}", ParagraphStyle("EmpGrand", parent=header, alignment=TA_RIGHT)),
    ])
    table = Table(rows, colWidths=[2.15*inch, 1.75*inch, 1.35*inch, 1.15*inch, 1.05*inch, 1.15*inch], repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), PURPLE), ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("ROWBACKGROUNDS", (0,1), (-1,-2), [colors.white, ROW_ALT]),
        ("BACKGROUND", (4,-1), (-1,-1), PURPLE_DARK), ("TEXTCOLOR", (4,-1), (-1,-1), colors.white),
        ("BOX", (0,0), (-1,-1), 0.7, BORDER), ("INNERGRID", (0,0), (-1,-1), 0.3, BORDER),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"), ("ALIGN", (3,1), (-1,-1), "RIGHT"),
        ("LEFTPADDING", (0,0), (-1,-1), 5), ("RIGHTPADDING", (0,0), (-1,-1), 5),
        ("TOPPADDING", (0,0), (-1,-1), 4), ("BOTTOMPADDING", (0,0), (-1,-1), 4),
        ("SPAN", (0,-1), (3,-1)),
    ]))
    story.append(table)
    doc.build(story, canvasmaker=PageNumberCanvas)


def _event_summary_vendor_counts(
    group: pd.DataFrame,
    decoration_fulfillment: str,
) -> tuple[int, int, int, int]:
    """Return mutually exclusive physical-garment counts for Event Summary.

    Every garment belongs in exactly one operational category: Blank / No
    Decoration, Embroidery In House, Outsourced Embroidery, or Outsourced
    Screen Printing.  The two outsourced categories use the same final
    per-line routing rule as the Outsourced Decoration Job Report, including
    Product Master Never Outsource exceptions.
    """
    blank_garments = 0
    embroidery_in_house = 0
    outsourced_embroidery = 0
    outsourced_screen_printing = 0

    for _, row in group.iterrows():
        decoration_type = (
            clean_text(row.get("Operational Decoration Type", ""))
            or clean_text(row.get("Decoration Type", ""))
        )
        decoration_key = decoration_type.casefold()
        quantity = _quantity(row.get("Quantity", 0))
        if quantity <= 0:
            continue
        if not decoration_key or is_blank_decoration(decoration_type):
            blank_garments += quantity
            continue

        # Product Master uses the normal label "Embroidery".  Match the shared
        # "embroid" root so that label and variations such as "Embroidered"
        # are all classified consistently.
        is_embroidery = "embroid" in decoration_key
        is_screen_printing = "screen" in decoration_key
        outsourced = is_outsourced_decoration(
            decoration_type,
            decoration_fulfillment,
            do_not_outsource=_do_not_outsource(row.get("Do Not Outsource", "No")),
        )
        if is_embroidery:
            if outsourced:
                outsourced_embroidery += quantity
            else:
                embroidery_in_house += quantity
        elif is_screen_printing:
            if outsourced:
                outsourced_screen_printing += quantity
            else:
                # A rare screen-printing line marked Never Outsource does not
                # create work for the outside decorator.  Keep the summary
                # concise by treating it as Blank / No Decoration for this
                # scheduling view, as requested by Orchid.
                blank_garments += quantity
        else:
            # A non-embroidery/non-screen service does not require an outside
            # decoration schedule, so it belongs in Blank / No Decoration.
            blank_garments += quantity
    return (
        blank_garments,
        embroidery_in_house,
        outsourced_embroidery,
        outsourced_screen_printing,
    )


def _build_event_summary_pdf(
    records: list[dict],
    ready: pd.DataFrame,
    pdf_path: Path,
    event_name: str,
    decoration_fulfillment: str = "",
) -> None:
    summary = employee_totals_summary(records)
    doc = SimpleDocTemplate(
        str(pdf_path), pagesize=letter, rightMargin=0.55*inch, leftMargin=0.55*inch,
        topMargin=0.48*inch, bottomMargin=0.48*inch,
        title=f"Event Summary - {event_name}", author="Orchid Uniforms & Apparel",
    )
    styles = getSampleStyleSheet()
    title = ParagraphStyle("EvtTitle", parent=styles["Title"], fontName="Helvetica-Bold", fontSize=20, leading=23, textColor=PURPLE_DARK)
    section = ParagraphStyle("EvtSection", parent=styles["Heading2"], fontName="Helvetica-Bold", fontSize=12, leading=14, textColor=PURPLE_DARK)
    body = ParagraphStyle("EvtBody", parent=styles["Normal"], fontSize=9, leading=11, textColor=TEXT)
    body_right = ParagraphStyle("EvtRight", parent=body, alignment=TA_RIGHT)
    header = ParagraphStyle("EvtHeader", parent=body, fontName="Helvetica-Bold", fontSize=7.6, leading=8.8, textColor=colors.white)
    column_header = ParagraphStyle(
        "EvtColumnHeader", parent=header, fontSize=7.25, leading=8.5, alignment=TA_CENTER,
    )
    overview_header = ParagraphStyle(
        "EvtOverviewHeader", parent=header, fontSize=8.6, leading=10,
        alignment=TA_CENTER,
    )
    overview_value = ParagraphStyle(
        "EvtOverviewValue", parent=body, fontName="Helvetica-Bold", fontSize=19,
        leading=22, textColor=PURPLE_DARK, alignment=TA_CENTER,
    )
    schedule_note = ParagraphStyle(
        "EvtScheduleNote", parent=body, fontSize=8.4, leading=10, textColor=MUTED, alignment=TA_CENTER,
    )
    story = [_paragraph("EVENT SUMMARY", title), _paragraph(event_name or "Uniform Sizing Event", body), Spacer(1, 10)]
    overview = Table([
        [
            _paragraph("NUMBER OF EMPLOYEES", overview_header),
            _paragraph("FINAL TOTAL AFTER DISCOUNTS", overview_header),
        ],
        [
            _paragraph(str(int(summary["employee_count"])), overview_value),
            _paragraph(f"${float(summary['grand_total']):,.2f}", overview_value),
        ],
    ], colWidths=[3.2*inch, 3.7*inch])
    overview.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), PURPLE), ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("BACKGROUND", (0,1), (-1,1), PURPLE_LIGHT), ("BOX", (0,0), (-1,-1), 0.8, BORDER),
        ("INNERGRID", (0,0), (-1,-1), 0.4, BORDER), ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("ALIGN", (0,0), (-1,-1), "CENTER"), ("TOPPADDING", (0,0), (-1,-1), 8), ("BOTTOMPADDING", (0,0), (-1,-1), 8),
    ]))
    if ready.empty:
        outsourced_embroidery_total = 0
        outsourced_screen_printing_total = 0
    else:
        outsourced_embroidery_total = 0
        outsourced_screen_printing_total = 0
        for _, group in ready.groupby("Vendor", sort=False):
            _, _, outsourced_embroidery, outsourced_screen_printing = _event_summary_vendor_counts(
                group, decoration_fulfillment
            )
            outsourced_embroidery_total += outsourced_embroidery
            outsourced_screen_printing_total += outsourced_screen_printing
    total_outsourced = outsourced_embroidery_total + outsourced_screen_printing_total
    outsourced_header = ParagraphStyle(
        "EvtOutsourcedHeader", parent=body, fontName="Helvetica-Bold", fontSize=8.6,
        leading=10, textColor=colors.white, alignment=TA_CENTER,
    )
    outsourced_value = ParagraphStyle(
        "EvtOutsourcedValue", parent=body, fontName="Helvetica-Bold", fontSize=19,
        leading=22, textColor=PURPLE_DARK, alignment=TA_CENTER,
    )
    outsourced_table = Table([
        [
            _paragraph("OUTSOURCED EMBROIDERY PIECES", outsourced_header),
            _paragraph("OUTSOURCED SCREEN-PRINTING PIECES", outsourced_header),
            _paragraph("TOTAL OUTSOURCED PIECES", outsourced_header),
        ],
        [
            _paragraph(str(outsourced_embroidery_total), outsourced_value),
            _paragraph(str(outsourced_screen_printing_total), outsourced_value),
            _paragraph(str(total_outsourced), outsourced_value),
        ],
    ], colWidths=[2.3 * inch, 2.3 * inch, 2.3 * inch])
    outsourced_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), PURPLE),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("BACKGROUND", (0, 1), (-1, 1), PURPLE_LIGHT),
        ("BOX", (0, 0), (-1, -1), 0.8, BORDER),
        ("INNERGRID", (0, 0), (-1, -1), 0.4, BORDER),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
    ]))
    story += [
        overview,
        Spacer(1, 13),
        _paragraph("OUTSOURCED DECORATION - SCHEDULING TOTALS", section),
        Spacer(1, 2),
        _paragraph("Give these two quantities to your outside decorator before purchase orders are placed.", schedule_note),
        Spacer(1, 6),
        outsourced_table,
        Spacer(1, 6),
        _paragraph(
            "These counts match the Outsourced Decoration Job Report. Product Master Never Outsource items are excluded.", schedule_note,
        ),
        Spacer(1, 12),
        _paragraph("VENDOR GARMENT BREAKDOWN", section),
        Spacer(1, 5),
    ]
    vendor_rows = [[
        _paragraph("Vendor", column_header),
        _paragraph("Total Garments", column_header),
        _paragraph("Embroidery In House", column_header),
        _paragraph("Blank / No Decoration", column_header),
        _paragraph("Outsourced Screen Printing", column_header),
        _paragraph("Outsourced Embroidery", column_header),
    ]]
    if ready.empty:
        grouped=[]
    else:
        grouped=[]
        for vendor, group in ready.groupby("Vendor", sort=True):
            blank, embroidery_in_house, outsourced_embroidery, outsourced_screen_printing = _event_summary_vendor_counts(
                group, decoration_fulfillment
            )
            total = int(group["Quantity"].sum())
            grouped.append((
                clean_text(vendor), total, embroidery_in_house, blank,
                outsourced_screen_printing, outsourced_embroidery,
            ))
    for vendor, total, embroidery_in_house, blank, outsourced_screen_printing, outsourced_embroidery in grouped:
        vendor_rows.append([
            _paragraph(vendor, body),
            _paragraph(str(total), body_right),
            _paragraph(str(embroidery_in_house), body_right),
            _paragraph(str(blank), body_right),
            _paragraph(str(outsourced_screen_printing), body_right),
            _paragraph(str(outsourced_embroidery), body_right),
        ])
    vendor_rows.append([
        _paragraph("TOTAL", header),
        _paragraph(str(sum(row[1] for row in grouped)), ParagraphStyle("Tot1", parent=header, alignment=TA_RIGHT)),
        _paragraph(str(sum(row[2] for row in grouped)), ParagraphStyle("Tot2", parent=header, alignment=TA_RIGHT)),
        _paragraph(str(sum(row[3] for row in grouped)), ParagraphStyle("Tot3", parent=header, alignment=TA_RIGHT)),
        _paragraph(str(sum(row[4] for row in grouped)), ParagraphStyle("Tot4", parent=header, alignment=TA_RIGHT)),
        _paragraph(str(sum(row[5] for row in grouped)), ParagraphStyle("Tot5", parent=header, alignment=TA_RIGHT)),
    ])
    table = Table(vendor_rows, colWidths=[2.10*inch, 0.75*inch, 0.95*inch, 1.00*inch, 1.20*inch, 1.15*inch], repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), PURPLE), ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("ROWBACKGROUNDS", (0,1), (-1,-2), [colors.white, ROW_ALT]),
        ("BACKGROUND", (0,-1), (-1,-1), PURPLE_DARK), ("TEXTCOLOR", (0,-1), (-1,-1), colors.white),
        ("BOX", (0,0), (-1,-1), 0.7, BORDER), ("INNERGRID", (0,0), (-1,-1), 0.3, BORDER),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"), ("ALIGN", (1,1), (-1,-1), "RIGHT"),
        ("LEFTPADDING", (0,0), (-1,-1), 6), ("RIGHTPADDING", (0,0), (-1,-1), 6),
        ("TOPPADDING", (0,0), (-1,0), 7), ("BOTTOMPADDING", (0,0), (-1,0), 7),
        ("TOPPADDING", (0,1), (-1,-1), 5), ("BOTTOMPADDING", (0,1), (-1,-1), 5),
    ]))
    story.append(table)
    story += [
        Spacer(1, 7),
        _paragraph(
            "Every garment is counted exactly once: Total Garments = Embroidery In House + Blank / No Decoration + "
            "Outsourced Screen Printing + Outsourced Embroidery. The two outsourced columns match the Outsourced "
            "Decoration Job Report.",
            body,
        ),
    ]
    doc.build(story, canvasmaker=PageNumberCanvas)

def _paragraph(value, style):
    text = clean_text(value).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    return Paragraph(text or " ", style)


def _customer_label(row: pd.Series, include_company: bool) -> str:
    company = clean_text(row.get("Company", ""))
    employee = clean_text(row.get("Employee Name", ""))
    order = clean_text(row.get("Order Number", ""))
    if employee and company and include_company:
        return f"{company} — {employee}"
    if employee:
        return employee
    if company:
        return company
    if order:
        return order
    return "Unspecified"


def _aggregate_rows(detail: pd.DataFrame) -> pd.DataFrame:
    if detail.empty:
        return pd.DataFrame(columns=COMPACT_COLUMNS)

    companies = {clean_text(value) for value in detail["Company"] if clean_text(value)}
    include_company = len(companies) > 1
    group_columns = [
        "Decoration Type", "Decoration Location", "Decoration Placement Instructions", "Decoration Color", "Style Number", "Product Name",
        "Garment Color", "Size",
    ]
    rows = []
    for keys, group in detail.groupby(group_columns, dropna=False, sort=False):
        deco_type, deco_location, placement_instructions, deco_color, style, product, garment_color, size = keys
        customers: OrderedDict[str, int] = OrderedDict()
        notes: OrderedDict[str, None] = OrderedDict()
        for _, source in group.iterrows():
            label = _customer_label(source, include_company)
            customers[label] = customers.get(label, 0) + int(source["Quantity"])
            note = clean_text(source.get("Purchase Instructions", ""))
            placement = clean_text(source.get("Decoration Placement Instructions", "")) or clean_text(placement_instructions)
            combined = "; ".join(filter(None, [f"Placement: {placement}" if placement else "", note]))
            if combined:
                notes[f"{label}: {combined}"] = None
        customer_text = "; ".join(
            f"{name} ({qty})" if qty > 1 else name
            for name, qty in customers.items()
        )
        rows.append({
            "Decoration Type": clean_text(deco_type),
            "Decoration Location": clean_text(deco_location),
            "Decoration Placement Instructions": clean_text(placement_instructions),
            "Decoration Color": clean_text(deco_color),
            "Style Number": clean_text(style),
            "Product Name": clean_text(product),
            "Garment Color": clean_text(garment_color),
            "Size": clean_text(size),
            "Quantity": int(group["Quantity"].sum()),
            "Customer(s)": customer_text,
            "Purchase Instructions": "; ".join(notes.keys()),
        })

    rows.sort(key=lambda row: (
        _decoration_type_sort(row["Decoration Type"]),
        location_sort_key(row["Decoration Location"]),
        row["Decoration Color"].casefold(),
        row["Style Number"].casefold(),
        row["Product Name"].casefold(),
        row["Garment Color"].casefold(),
        size_sort_key(row["Size"]),
    ))
    return pd.DataFrame(rows, columns=COMPACT_COLUMNS)


def _section_heading(text: str, style, dark: bool = False) -> Table:
    table = Table([[_paragraph(text, style)]], colWidths=[10.15 * inch])
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), PURPLE_DARK if dark else PURPLE_LIGHT),
        ("TEXTCOLOR", (0, 0), (-1, -1), colors.white if dark else PURPLE_DARK),
        ("BOX", (0, 0), (-1, -1), 0.7, PURPLE_DARK if dark else BORDER),
        ("LEFTPADDING", (0, 0), (-1, -1), 7),
        ("RIGHTPADDING", (0, 0), (-1, -1), 7),
        ("TOPPADDING", (0, 0), (-1, -1), 9 if dark else 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 9 if dark else 6),
    ]))
    return table


def _compact_table(section_detail: pd.DataFrame, body_style, qty_style, header_style) -> Table:
    """Concise vendor-order table used only for placing the garment order."""
    group_columns = ["Style Number", "Product Name", "Garment Color", "Size"]
    rows = [[
        _paragraph("Product #", header_style),
        _paragraph("Description", header_style),
        _paragraph("Garment Color", header_style),
        _paragraph("Size", header_style),
        _paragraph("Quantity", header_style),
    ]]
    grouped = (
        section_detail.groupby(group_columns, dropna=False, sort=False)["Quantity"]
        .sum()
        .reset_index()
    )
    grouped = grouped.sort_values(
        by=["Style Number", "Product Name", "Garment Color", "Size"],
        key=lambda column: column.map(lambda value: size_sort_key(value) if column.name == "Size" else clean_text(value).casefold()),
        kind="stable",
    )
    for _, row in grouped.iterrows():
        rows.append([
            _paragraph(row["Style Number"], body_style),
            _paragraph(row["Product Name"], body_style),
            _paragraph(row["Garment Color"], body_style),
            _paragraph(row["Size"], body_style),
            _paragraph(str(int(row["Quantity"])), qty_style),
        ])

    table = Table(
        rows,
        colWidths=[1.20 * inch, 4.65 * inch, 1.75 * inch, 1.00 * inch, 1.55 * inch],
        repeatRows=1,
        splitByRow=1,
    )
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), PURPLE),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, ROW_ALT]),
        ("BOX", (0, 0), (-1, -1), 0.8, BORDER),
        ("INNERGRID", (0, 0), (-1, -1), 0.35, BORDER),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (3, 1), (4, -1), "CENTER"),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, 0), 5),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 5),
        ("TOPPADDING", (0, 1), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 1), (-1, -1), 4),
    ]))
    return table

def _build_pdf(
    report_detail: pd.DataFrame,
    pdf_path: Path,
    po_number: str,
    vendor: str,
    report_mode: str,
    decoration_type: str = "",
    event_name: str = "",
):
    report_mode = normalize_report_mode(report_mode)
    page_size = landscape(letter)
    page_width, page_height = page_size
    left_margin = 0.38 * inch
    right_margin = 0.38 * inch
    bottom_margin = 0.42 * inch
    header_reserve = 1.76 * inch
    frame_width = page_width - left_margin - right_margin
    frame_height = page_height - bottom_margin - header_reserve

    styles = getSampleStyleSheet()
    label_style = ParagraphStyle(
        "LabelO", parent=styles["Normal"], fontName="Helvetica-Bold",
        fontSize=8.4, leading=9.6, textColor=MUTED,
    )
    value_style = ParagraphStyle(
        "ValueO", parent=styles["Normal"], fontName="Helvetica-Bold",
        fontSize=10.4, leading=12.0, textColor=TEXT,
    )
    body_style = ParagraphStyle(
        "BodyO", parent=styles["Normal"], fontSize=9.4,
        leading=10.6, textColor=TEXT,
    )
    qty_style = ParagraphStyle(
        "QtyO", parent=body_style, fontName="Helvetica-Bold",
        fontSize=10.8, leading=11.4, alignment=TA_CENTER, textColor=PURPLE_DARK,
    )
    header_style = ParagraphStyle(
        "TableHeaderO", parent=body_style, fontName="Helvetica-Bold",
        fontSize=9.3, leading=10.4, textColor=colors.white, alignment=TA_LEFT,
    )
    section_style = ParagraphStyle(
        "SectionO", parent=body_style, fontName="Helvetica-Bold",
        fontSize=12.8, leading=14.8, textColor=PURPLE_DARK, alignment=TA_LEFT,
    )
    section_dark_style = ParagraphStyle(
        "SectionDarkO", parent=section_style, fontSize=21,
        leading=23.5, textColor=colors.white, alignment=TA_CENTER,
    )

    decoration_types = sorted(
        {clean_text(value) for value in report_detail["Decoration Type"]},
        key=_decoration_type_sort,
    )
    if not decoration_types:
        decoration_types = [BLANK_DECORATION_LABEL]

    order_date = datetime.now().strftime("%B %d, %Y")
    logo_path = _resource_path("assets", "orchid_logo.png")

    def _draw_page_header(canvas, doc_obj, section_label: str):
        canvas.saveState()
        if logo_path.exists():
            canvas.drawImage(
                str(logo_path), left_margin, page_height - 0.95 * inch,
                width=2.05 * inch, height=0.79 * inch,
                preserveAspectRatio=True, mask="auto",
            )
        else:
            canvas.setFont("Helvetica-Bold", 17)
            canvas.setFillColor(PURPLE_DARK)
            canvas.drawString(left_margin, page_height - 0.52 * inch, "ORCHID UNIFORMS & APPAREL")

        canvas.setFillColor(PURPLE_DARK)
        title_text = f"{vendor.upper()} PURCHASE ORDER"
        title_size = 19 if len(title_text) <= 28 else 16.5
        canvas.setFont("Helvetica-Bold", title_size)
        canvas.drawCentredString(page_width / 2, page_height - 0.46 * inch, title_text)
        canvas.setStrokeColor(PURPLE)
        canvas.setLineWidth(1.2)
        canvas.line(left_margin, page_height - 0.99 * inch, page_width - right_margin, page_height - 0.99 * inch)

        context_label = "EVENT" if report_mode == UNIFORM_SIZING_EVENT else "ORDER TYPE"
        context_value = event_name or report_mode
        info = Table([
            [
                _paragraph(context_label, label_style),
                _paragraph("PO NUMBER", label_style),
                _paragraph("ORDER DATE", label_style),
            ],
            [
                _paragraph(context_value, value_style),
                _paragraph(po_number, value_style),
                _paragraph(order_date, value_style),
            ],
        ], colWidths=[4.38 * inch, 3.04 * inch, 2.73 * inch])
        info.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), PURPLE_LIGHT),
            ("BOX", (0, 0), (-1, -1), 0.7, BORDER),
            ("INNERGRID", (0, 0), (-1, -1), 0.35, BORDER),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 5),
            ("RIGHTPADDING", (0, 0), (-1, -1), 5),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        info_width, info_height = info.wrap(frame_width, 0.62 * inch)
        info.drawOn(canvas, left_margin, page_height - 1.04 * inch - info_height)

        canvas.setFont("Helvetica", 7.5)
        canvas.setFillColor(MUTED)
        canvas.drawString(left_margin, 0.20 * inch, f"Orchid Uniforms & Apparel  |  {vendor}  |  {section_label}")
        canvas.drawRightString(page_width - right_margin, 0.20 * inch, f"PO {po_number}")
        canvas.restoreState()

    doc = BaseDocTemplate(
        str(pdf_path),
        pagesize=page_size,
        rightMargin=right_margin,
        leftMargin=left_margin,
        topMargin=header_reserve,
        bottomMargin=bottom_margin,
        title=f"Purchase Order {po_number}",
        author="Orchid Uniforms & Apparel",
    )

    templates = []
    template_ids = []
    group_details = []
    group_index = 0
    for deco_type in decoration_types:
        type_detail = report_detail[
            report_detail["Decoration Type"].map(clean_text).eq(deco_type)
        ].copy()
        if is_blank_decoration(deco_type):
            locations = [""]
        else:
            type_detail["Decoration Location"] = type_detail.apply(
                lambda row: normalize_decoration_location(
                    row.get("Decoration Location", ""), row.get("Decoration Type", "")
                ),
                axis=1,
            )
            locations = sorted(
                {clean_text(value) or LEFT_CHEST for value in type_detail["Decoration Location"]},
                key=location_sort_key,
            )
        for location in locations:
            location_detail = type_detail if not location else type_detail[
                type_detail["Decoration Location"].map(clean_text).eq(location)
            ]
            base_label = _display_decoration_type(deco_type)
            section_label = base_label if not location else f"{base_label} — {location}"
            template_id = f"decoration_group_{group_index}"
            frame = Frame(
                left_margin, bottom_margin, frame_width, frame_height,
                leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0,
                id=f"frame_{group_index}",
            )
            on_page = lambda canvas, doc_obj, label=section_label: _draw_page_header(canvas, doc_obj, label)
            templates.append(PageTemplate(id=template_id, frames=[frame], onPage=on_page, pagesize=page_size))
            template_ids.append(template_id)
            group_details.append((deco_type, location, location_detail, section_label))
            group_index += 1
    doc.addPageTemplates(templates)

    story = []
    for index, (deco_type, location, type_detail, section_label) in enumerate(group_details):
        if index > 0:
            story.extend([NextPageTemplate(template_ids[index]), PageBreak()])

        story.append(_section_heading(section_label.upper(), section_dark_style, dark=True))
        if location and is_custom_location(location):
            story.append(Spacer(1, 5))
            story.append(_section_heading(
                f"CUSTOM LOCATION — VERIFY PLACEMENT: {location}", section_style
            ))
        story.append(Spacer(1, 7))

        if is_blank_decoration(deco_type):
            story.append(_compact_table(type_detail, body_style, qty_style, header_style))
            story.append(Spacer(1, 10))
            continue

        colors_in_type = sorted(
            {clean_text(value) for value in type_detail["Decoration Color"]},
            key=str.casefold,
        )
        for color_index, color in enumerate(colors_in_type):
            section = type_detail[
                type_detail["Decoration Color"].map(clean_text).eq(color)
            ]
            if color_index > 0:
                story.append(Spacer(1, 6))
            story.append(CondPageBreak(0.82 * inch))
            story.append(_section_heading(_color_section_label(deco_type, color), section_style))
            story.append(Spacer(1, 4))
            story.append(_compact_table(section, body_style, qty_style, header_style))
            story.append(Spacer(1, 5))

    unique_items = int(
        report_detail[["Style Number", "Product Name", "Garment Color", "Size"]]
        .drop_duplicates()
        .shape[0]
    )
    order_sections = int(
        report_detail[["Decoration Type", "Decoration Location", "Decoration Color"]]
        .drop_duplicates()
        .shape[0]
    )
    summary = Table([
        [_paragraph("TOTAL GARMENTS", label_style), _paragraph("UNIQUE ITEMS", label_style), _paragraph("ORDER SECTIONS", label_style)],
        [_paragraph(str(int(report_detail["Quantity"].sum())), value_style), _paragraph(str(unique_items), value_style), _paragraph(str(order_sections), value_style)],
    ], colWidths=[3.38 * inch, 3.38 * inch, 3.39 * inch])
    summary.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), PURPLE_LIGHT),
        ("BOX", (0, 0), (-1, -1), 0.7, BORDER),
        ("INNERGRID", (0, 0), (-1, -1), 0.35, BORDER),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    story.append(summary)

    doc.build(story, canvasmaker=PageNumberCanvas)


def _build_non_included_items_pdf(detail: pd.DataFrame, pdf_path: Path, event_name: str) -> None:
    """Create one multi-vendor internal purchase document for Ship-to-Orchid exceptions."""
    doc = SimpleDocTemplate(
        str(pdf_path), pagesize=landscape(letter), rightMargin=0.34 * inch,
        leftMargin=0.34 * inch, topMargin=0.32 * inch, bottomMargin=0.38 * inch,
        title="Non-Included Items - Internal Purchase Order",
        author="Orchid Uniforms & Apparel",
    )
    styles = getSampleStyleSheet()
    title = ParagraphStyle(
        "NonIncludedTitle", parent=styles["Title"], fontName="Helvetica-Bold",
        fontSize=18, leading=21, textColor=PURPLE_DARK, alignment=TA_CENTER,
    )
    subtitle = ParagraphStyle(
        "NonIncludedSub", parent=styles["Normal"], fontName="Helvetica-Bold",
        fontSize=9.5, leading=12, textColor=PURPLE, alignment=TA_CENTER,
    )
    note = ParagraphStyle(
        "NonIncludedNote", parent=styles["Normal"], fontSize=8.2, leading=10.2,
        textColor=TEXT, alignment=TA_LEFT,
    )
    vendor_style = ParagraphStyle(
        "NonIncludedVendor", parent=styles["Heading2"], fontName="Helvetica-Bold",
        fontSize=12.5, leading=15, textColor=colors.white,
    )
    header_style = ParagraphStyle(
        "NonIncludedHeader", parent=styles["Normal"], fontName="Helvetica-Bold",
        fontSize=6.8, leading=8, textColor=colors.white,
    )
    body_style = ParagraphStyle(
        "NonIncludedBody", parent=styles["Normal"], fontSize=6.7, leading=8.2, textColor=TEXT,
    )
    qty_style = ParagraphStyle(
        "NonIncludedQty", parent=body_style, fontName="Helvetica-Bold", alignment=TA_CENTER,
    )

    context = event_name or "Current Purchase Packet"
    story = [
        _official_logo_flowable(), Spacer(1, 3),
        _paragraph("NON-INCLUDED ITEMS", title),
        _paragraph(context, subtitle),
        _paragraph("INTERNAL PURCHASE ORDER — ORDER MANUALLY AND SHIP TO ORCHID", subtitle),
        Spacer(1, 9),
    ]

    for vendor_value, vendor_detail in detail.groupby("Vendor", dropna=False, sort=True):
        vendor = clean_text(vendor_value) or "Vendor Not Assigned"
        story.append(CondPageBreak(0.96 * inch))
        story.append(_section_heading(vendor, vendor_style, dark=True))
        story.append(Spacer(1, 4))
        story.append(_compact_table(vendor_detail, body_style, qty_style, header_style))
        vendor_total = Table([
            [_paragraph("VENDOR TOTAL", header_style), _paragraph(str(int(vendor_detail["Quantity"].sum())), header_style)]
        ], colWidths=[8.60 * inch, 1.55 * inch], hAlign="LEFT")
        vendor_total.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), PURPLE_DARK),
            ("TEXTCOLOR", (0, 0), (-1, -1), colors.white),
            ("ALIGN", (1, 0), (1, 0), "CENTER"),
            ("BOX", (0, 0), (-1, -1), 0.7, BORDER),
            ("LEFTPADDING", (0, 0), (-1, -1), 5),
            ("RIGHTPADDING", (0, 0), (-1, -1), 5),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ]))
        story.append(vendor_total)
        story.append(Spacer(1, 10))

    doc.build(story, canvasmaker=PageNumberCanvas)



def _operational_value(row: pd.Series, operational_field: str, fallback_field: str) -> str:
    return clean_text(row.get(operational_field, "")) or clean_text(row.get(fallback_field, ""))


def _note_key(value: object) -> str:
    """Normalize an instruction for reliable duplicate comparison."""
    import re

    text = clean_text(value).casefold()
    text = re.sub(r"[^a-z0-9$%]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _remove_known_phrase(text: str, phrase: str) -> str:
    """Remove a previously emitted phrase while preserving the remaining note."""
    import re

    words = re.findall(r"[A-Za-z0-9$%]+", clean_text(phrase))
    if not words:
        return text
    pattern = r"(?i)(?<![A-Za-z0-9])" + r"[\s\W_]+".join(re.escape(word) for word in words) + r"(?![A-Za-z0-9])"
    cleaned = re.sub(pattern, " ", text)
    cleaned = re.sub(r"\s+([,.;:])", r"\1", cleaned)
    cleaned = re.sub(r"^[\s,.;:|/-]+|[\s,.;:|/-]+$", "", cleaned)
    return re.sub(r"\s+", " ", cleaned).strip()


def _deduplicated_text(*values: object) -> str:
    """Combine notes while removing repeated sentences and embedded phrases.

    Shopify order notes, line notes, and placement instructions can contain the
    same wording more than once. This keeps each unique instruction once within
    a report row while leaving identical instructions on other product rows.
    """
    import re

    result: list[str] = []
    keys: set[str] = set()
    for value in values:
        raw = clean_text(value)
        if not raw:
            continue
        pieces = [
            clean_text(piece)
            for piece in re.split(r"(?:\r?\n)+|[;|]+|(?<=[.!?])\s+(?=[A-Z0-9])", raw)
            if clean_text(piece)
        ]
        for piece in pieces:
            candidate = piece
            # Remove phrases that were already emitted from a larger combined note.
            for existing in result:
                existing_key = _note_key(existing)
                candidate_key = _note_key(candidate)
                if not candidate_key:
                    break
                if candidate_key == existing_key or candidate_key in existing_key:
                    candidate = ""
                    break
                if existing_key and existing_key in candidate_key:
                    candidate = _remove_known_phrase(candidate, existing)
            candidate = clean_text(candidate)
            key = _note_key(candidate)
            if candidate and key and key not in keys:
                result.append(candidate)
                keys.add(key)
    return "; ".join(result)


def _empty_checkbox() -> Drawing:
    """Return a true vector checkbox that prints sharply at any PDF zoom."""
    drawing = Drawing(12, 12)
    drawing.add(Rect(1.5, 1.5, 9, 9, strokeColor=PURPLE_DARK, fillColor=None, strokeWidth=1.0))
    return drawing


def _strip_pricing_fragments(value: object) -> str:
    """Remove all pricing/budget language from operational notes, regardless of spacing."""
    import re
    raw = clean_text(value)
    if not raw:
        return ""
    # Insert boundaries between compact amounts/words (for example $700Allowance).
    normalized = re.sub(r"(?<=\d)(?=[A-Za-z])", " ", raw)
    normalized = re.sub(r"(?<=[A-Za-z])(?=\$?\d)", " ", normalized)
    pieces = [clean_text(p) for p in re.split(r"(?:\r?\n)+|[;|]+|(?<=[.!?])\s+(?=[A-Z0-9$])", normalized) if clean_text(p)]
    pricing_terms = re.compile(r"(?i)\b(?:allowance|discount(?:ed)?|budget|overage|city\s*total|order\s*total|customer\s*total|price\s*adjustment|pricing\s*adjustment)\b")
    money = re.compile(r"\$\s*[\d,.]+")
    kept=[]
    for piece in pieces:
        if pricing_terms.search(piece):
            # Preserve any operational text before the pricing fragment.
            cut = pricing_terms.search(piece).start()
            candidate = piece[:cut].strip(" ,.;:-")
            candidate = money.sub("", candidate).strip(" ,.;:-")
            if candidate:
                kept.append(candidate)
            continue
        kept.append(piece)
    return _deduplicated_text(*kept)


def _line_context(row: pd.Series) -> str:
    """Text used to decide whether an order-level instruction belongs to this item."""
    return " ".join(filter(None, [
        clean_text(row.get("Style Number", "")),
        clean_text(row.get("Product Name", "")),
        clean_text(row.get("Product Aliases", "")),
        clean_text(row.get("Garment Color", "")),
        clean_text(row.get("Size", "")),
    ]))


def _obvious_family_match(note: str, row: pd.Series) -> bool:
    """Suppress instructions that clearly name a different garment family.

    This intentionally handles only strong, obvious wording. Ambiguous notes are
    preserved for review rather than guessed away.
    """
    import re

    note_text = clean_text(note).casefold()
    product_text = " ".join(filter(None, [
        clean_text(row.get("Product Name", "")),
        clean_text(row.get("Style Number", "")),
    ])).casefold()
    if not note_text:
        return False

    headwear = bool(re.search(r"\b(?:hat|hats|cap|caps|beanie|beanies|boonie|boonies|visor|visors|headwear)\b", product_text))
    shorts = bool(re.search(r"\bshorts?\b", product_text))
    long_bottom = bool(re.search(r"\b(?:pant|pants|jean|jeans|trouser|trousers|bib|bibs|overall|overalls)\b", product_text))
    outerwear = bool(re.search(r"\b(?:jacket|jackets|coat|coats|parka|parkas|outerwear|raincoat|raincoats|softshell|soft shell|shell)\b", product_text))
    shirt = bool(re.search(r"\b(?:shirt|shirts|tee|tees|t-shirt|t-shirts|polo|polos|blouse|blouses|woven|wovens|quarter zip|1/4 zip|half zip)\b", product_text))

    # Alteration/inseam notes belong to full-length bottoms. Shorts are allowed
    # only when the note explicitly says shorts.
    if re.search(r"\b(?:hem|hemmed|hemming|inseam)\b", note_text):
        if shorts:
            return bool(re.search(r"\bshorts?\b", note_text))
        return long_bottom

    named_groups = []
    if re.search(r"\b(?:jacket|jackets|coat|coats|parka|parkas|outerwear|raincoat|raincoats|waterproof jacket|waterproof jackets)\b", note_text):
        named_groups.append(outerwear)
    if re.search(r"\b(?:pant|pants|jean|jeans|trouser|trousers|bib|bibs|overall|overalls)\b", note_text):
        named_groups.append(long_bottom)
    if re.search(r"\bshorts?\b", note_text):
        named_groups.append(shorts)
    if re.search(r"\b(?:hat|hats|cap|caps|beanie|beanies|boonie|boonies|visor|visors|headwear)\b", note_text):
        named_groups.append(headwear)
    if re.search(r"\b(?:shirt|shirts|tee|tees|t-shirt|t-shirts|polo|polos|blouse|blouses|quarter zip|1/4 zip|half zip)\b", note_text):
        named_groups.append(shirt)
    if named_groups:
        return any(named_groups)
    return True


def _targeted_operational_text(value: object, row: pd.Series, *, line_specific: bool = False) -> str:
    """Return only instruction fragments that reasonably apply to this row."""
    import re

    cleaned = _strip_pricing_fragments(value)
    if not cleaned:
        return ""
    pieces = [
        clean_text(piece)
        for piece in re.split(r"(?:\r?\n)+|[;|]+|(?<=[.!?])\s+(?=[A-Z0-9$])", cleaned)
        if clean_text(piece)
    ]
    context = _line_context(row)
    current_style = re.sub(r"\s+", "", clean_text(row.get("Style Number", ""))).upper()
    kept: list[str] = []
    for piece in pieces:
        if line_specific:
            kept.append(piece)
            continue

        explicit_styles = decoration_target_styles(piece)
        if explicit_styles and current_style not in explicit_styles:
            continue
        if decoration_note_requires_review(piece):
            if not decoration_note_targets_line(piece, context):
                continue
        elif not _obvious_family_match(piece, row):
            continue
        kept.append(piece)
    return _deduplicated_text(*kept)


def _receiving_notes(row: pd.Series) -> str:
    return _deduplicated_text(
        _targeted_operational_text(
            _operational_value(row, "Operational Placement Instructions", "Decoration Placement Instructions"),
            row,
            line_specific=True,
        ),
        _targeted_operational_text(row.get("Purchase Instructions", ""), row),
        _targeted_operational_text(row.get("Shopify Line Notes", ""), row, line_specific=True),
        _targeted_operational_text(row.get("Shopify Order Notes", ""), row),
    )


def _decorator_instructions(row: pd.Series) -> str:
    placement = _operational_value(
        row, "Operational Placement Instructions", "Decoration Placement Instructions"
    )
    purchase_note = clean_text(row.get("Purchase Instructions", ""))
    decoration_type = _operational_value(
        row, "Operational Decoration Type", "Decoration Type"
    ).casefold()
    sensitive_values = [
        clean_text(row.get("Employee Name", "")),
        clean_text(row.get("Company", "")),
        clean_text(row.get("Order Number", "")),
    ]
    note_key = purchase_note.casefold()
    internal_only_terms = (
        "allowance", "discount", "budget", "price", "hem ", "hemming",
        "inseam", "name on", "personalization", "new hire package",
        "ship to orchid", "do not outsource", "backorder", "shortage",
    )
    if purchase_note and any(
        value and value.casefold() in note_key
        for value in sensitive_values
    ):
        purchase_note = ""
    elif purchase_note and any(term in note_key for term in internal_only_terms):
        purchase_note = ""
    elif purchase_note:
        common_terms = (
            "logo", "artwork", "placement", "left chest", "right chest",
            "left sleeve", "right sleeve", "front", "back", "yoke",
        )
        if "screen" in decoration_type:
            relevant = any(term in note_key for term in common_terms + ("screen", "print", "ink"))
            conflicting = "embroider" in note_key or "thread" in note_key
        elif "embroider" in decoration_type:
            relevant = any(term in note_key for term in common_terms + ("embroider", "embroidery", "thread", "hat", "beanie"))
            conflicting = "screen print" in note_key or " ink" in note_key
        else:
            relevant = any(term in note_key for term in common_terms)
            conflicting = False
        if not relevant or conflicting:
            purchase_note = ""
    return _deduplicated_text(placement, purchase_note)


def _build_in_house_receiving_report(
    detail: pd.DataFrame,
    pdf_path: Path,
    event_name: str,
    fulfillment_mode: str,
) -> None:
    doc = SimpleDocTemplate(
        str(pdf_path), pagesize=landscape(letter), rightMargin=0.28 * inch,
        leftMargin=0.28 * inch, topMargin=0.28 * inch, bottomMargin=0.34 * inch,
        title="In-House Receiving and Decoration Report",
        author="Orchid Uniforms & Apparel",
    )
    styles = getSampleStyleSheet()
    title = ParagraphStyle(
        "InHouseTitle", parent=styles["Title"], fontName="Helvetica-Bold",
        fontSize=18, leading=21, textColor=PURPLE_DARK, alignment=TA_CENTER,
    )
    subtitle = ParagraphStyle(
        "InHouseSub", parent=styles["Normal"], fontName="Helvetica-Bold",
        fontSize=9.2, leading=11.5, textColor=PURPLE, alignment=TA_CENTER,
    )
    note = ParagraphStyle(
        "InHouseNote", parent=styles["Normal"], fontSize=7.4, leading=9.0,
        textColor=TEXT, alignment=TA_LEFT,
    )
    vendor_style = ParagraphStyle(
        "InHouseVendor", parent=styles["Heading2"], fontName="Helvetica-Bold",
        fontSize=12.3, leading=14.5, textColor=colors.white,
    )
    header_style = ParagraphStyle(
        "InHouseHeader", parent=styles["Normal"], fontName="Helvetica-Bold",
        fontSize=5.9, leading=7.0, textColor=colors.white, alignment=TA_CENTER,
    )
    body_style = ParagraphStyle(
        "InHouseBody", parent=styles["Normal"], fontSize=6.0, leading=7.2,
        textColor=TEXT,
    )
    center_style = ParagraphStyle(
        "InHouseCenter", parent=body_style, fontName="Helvetica-Bold", alignment=TA_CENTER,
    )

    context = event_name or "Current Purchase Packet"
    story = [
        _official_logo_flowable(), Spacer(1, 3),
        _paragraph("IN-HOUSE RECEIVING & DECORATION REPORT", title),
        _paragraph(context, subtitle),
        _paragraph("Internal Receiving & Decoration Worksheet", subtitle),
        Spacer(1, 8),
    ]

    sort_columns = ["Vendor", "Assigned PO Number", "Style Number", "Garment Color", "Size", "Employee Name", "Order Number"]
    detail = detail.copy()
    for column in sort_columns:
        if column not in detail.columns:
            detail[column] = ""
    detail = detail.sort_values(
        by=sort_columns,
        key=lambda column: column.map(
            lambda value: size_sort_key(value) if column.name == "Size" else clean_text(value).casefold()
        ),
        kind="stable",
    )

    for (vendor_value, po_value), group in detail.groupby(
        ["Vendor", "Assigned PO Number"], dropna=False, sort=False
    ):
        vendor = clean_text(vendor_value) or "Vendor Not Assigned"
        po_number = clean_text(po_value) or "PO Not Assigned"
        story.append(CondPageBreak(1.02 * inch))
        story.append(_section_heading(f"{vendor}  —  {po_number}  —  {int(group["Quantity"].sum())} PIECES", vendor_style, dark=True))
        story.append(Spacer(1, 4))
        rows = [[
            _paragraph("Product #", header_style),
            _paragraph("Description", header_style),
            _paragraph("Garment Color", header_style),
            _paragraph("Size", header_style),
            _paragraph("Qty", header_style),
            _paragraph("Employee / Order", header_style),
            _paragraph("Decoration / Location", header_style),
            _paragraph("Thread / Ink", header_style),
            _paragraph("Notes / Instructions", header_style),
            _paragraph("Received", header_style),
        ]]
        for _, row in group.iterrows():
            employee_order = _deduplicated_text(
                _customer_label(row, include_company=True),
                f"Order {clean_text(row.get('Order Number', ''))}" if clean_text(row.get("Order Number", "")) else "",
            )
            decoration_type = _operational_value(row, "Operational Decoration Type", "Decoration Type")
            decoration_location = _operational_value(row, "Operational Decoration Location", "Decoration Location")
            decoration = " — ".join(filter(None, [
                _display_decoration_type(decoration_type), decoration_location,
            ]))
            decoration_color = _operational_value(row, "Operational Decoration Color", "Decoration Color")
            rows.append([
                _paragraph(row.get("Style Number", ""), body_style),
                _paragraph(row.get("Product Name", ""), body_style),
                _paragraph(row.get("Garment Color", ""), body_style),
                _paragraph(row.get("Size", ""), center_style),
                _paragraph(str(int(row.get("Quantity", 0))), center_style),
                _paragraph(employee_order, body_style),
                _paragraph(decoration, body_style),
                _paragraph(decoration_color, body_style),
                _paragraph(_receiving_notes(row), body_style),
                _empty_checkbox(),
            ])
        table = Table(
            rows,
            colWidths=[
                0.72*inch, 1.35*inch, 0.78*inch, 0.45*inch, 0.38*inch,
                1.35*inch, 1.20*inch, 0.65*inch, 2.46*inch, 0.55*inch,
            ],
            repeatRows=1, splitByRow=1, hAlign="LEFT",
        )
        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), PURPLE),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, ROW_ALT]),
            ("BOX", (0, 0), (-1, -1), 0.7, BORDER),
            ("INNERGRID", (0, 0), (-1, -1), 0.3, BORDER),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("ALIGN", (3, 1), (4, -1), "CENTER"),
            ("ALIGN", (9, 0), (9, -1), "CENTER"),
            ("VALIGN", (9, 1), (9, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 2.5),
            ("RIGHTPADDING", (0, 0), (-1, -1), 2.5),
            ("TOPPADDING", (0, 0), (-1, -1), 3.2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3.2),
        ]))
        story.append(table)
        story.append(Spacer(1, 9))

    doc.build(story, canvasmaker=PageNumberCanvas)


def _build_outsourced_decoration_report(
    detail: pd.DataFrame,
    pdf_path: Path,
    event_name: str,
    fulfillment_mode: str,
    job_logo_path: Path | None = None,
    outsourced_job_name: str = "",
) -> None:
    doc = SimpleDocTemplate(
        str(pdf_path), pagesize=landscape(letter), rightMargin=0.34 * inch,
        leftMargin=0.34 * inch, topMargin=0.30 * inch, bottomMargin=0.36 * inch,
        title="Outsourced Decoration Job Report",
        author="Orchid Uniforms & Apparel",
    )
    styles = getSampleStyleSheet()
    # The outsourced decoration report is a production packet. Keep Orchid's
    # own identity compact in the left corner, reserve the middle for the job
    # name and customer/job logo, and keep the outside-decorator destination
    # directly under Orchid's logo.
    job_name_style = ParagraphStyle(
        "OutJobName", parent=styles["Title"], fontName="Helvetica-Bold",
        fontSize=22, leading=25, textColor=PURPLE_DARK, alignment=TA_CENTER,
        spaceAfter=2,
    )
    outsource_label_style = ParagraphStyle(
        "OutsourceLabel", parent=styles["Normal"], fontName="Helvetica-Bold",
        fontSize=8.6, leading=10.5, textColor=MUTED, alignment=TA_LEFT,
    )
    outsource_vendor_style = ParagraphStyle(
        "OutsourceVendor", parent=styles["Normal"], fontName="Helvetica-Bold",
        fontSize=11.5, leading=13.9, textColor=PURPLE_DARK, alignment=TA_LEFT,
    )
    note = ParagraphStyle(
        "OutNote", parent=styles["Normal"], fontSize=8.0, leading=9.8,
        textColor=TEXT, alignment=TA_LEFT,
    )
    route_style = ParagraphStyle(
        "OutRoute", parent=styles["Heading2"], fontName="Helvetica-Bold",
        fontSize=12.8, leading=15.0, textColor=colors.white,
    )
    header_style = ParagraphStyle(
        "OutHeader", parent=styles["Normal"], fontName="Helvetica-Bold",
        fontSize=7.0, leading=8.2, textColor=colors.white,
    )
    body_style = ParagraphStyle(
        "OutBody", parent=styles["Normal"], fontSize=7.2, leading=8.6,
        textColor=TEXT,
    )
    qty_style = ParagraphStyle(
        "OutQty", parent=body_style, fontName="Helvetica-Bold", alignment=TA_CENTER,
    )

    # The manually entered job/logo name is the production name recognized by
    # the outside decorator (for example, “Utilities Department”).  Keep the
    # event/order name as a sensible fallback for older packets.
    context = clean_text(outsourced_job_name) or event_name or "Current Purchase Packet"
    # Enlarge the compact Orchid routing block by 20% while retaining its
    # supporting role at the left edge of the production header.
    orchid_logo = _official_logo_flowable(max_width=1.39 * inch, max_height=0.60 * inch)
    try:
        orchid_logo.hAlign = "LEFT"
    except Exception:
        pass
    left_header = [
        orchid_logo,
        Spacer(1, 5),
        _paragraph("OUTSOURCED TO:", outsource_label_style),
        _paragraph("Stitch N Print", outsource_vendor_style),
    ]
    center_header = [
        _paragraph(context, job_name_style),
        Spacer(1, 11),
    ]
    if job_logo_path and Path(job_logo_path).exists():
        try:
            # Crop empty transparent/white export canvas in memory first, then
            # use the existing large header bounds for the visible artwork.
            logo = _job_logo_flowable(Path(job_logo_path), 4.76 * inch, 2.03 * inch)
            center_header.append(logo)
        except Exception:
            pass
    header = Table(
        [[left_header, center_header, Spacer(1, 1)]],
        colWidths=[1.65 * inch, 7.02 * inch, 1.65 * inch],
        hAlign="LEFT",
    )
    header.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("ALIGN", (0, 0), (0, 0), "LEFT"),
        ("ALIGN", (1, 0), (1, 0), "CENTER"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))
    story = [header]
    story.append(Spacer(1, 8))

    detail = detail.copy()
    detail["Report Decoration Type"] = detail.apply(
        lambda row: _operational_value(row, "Operational Decoration Type", "Decoration Type"), axis=1
    )
    detail["Report Decoration Location"] = detail.apply(
        lambda row: _operational_value(row, "Operational Decoration Location", "Decoration Location"), axis=1
    )
    detail["Report Decoration Color"] = detail.apply(
        lambda row: _operational_value(row, "Operational Decoration Color", "Decoration Color"), axis=1
    )
    route_columns = ["Report Decoration Type", "Report Decoration Location", "Report Decoration Color"]
    routes = list(detail.groupby(route_columns, dropna=False, sort=False))
    routes.sort(key=lambda item: (
        _decoration_type_sort(item[0][0]),
        location_sort_key(item[0][1]),
        clean_text(item[0][2]).casefold(),
    ))

    for keys, group in routes:
        deco_type, location, color_value = keys
        route_label = " — ".join(filter(None, [
            _display_decoration_type(deco_type), clean_text(location),
            _color_section_label(deco_type, color_value),
        ]))
        story.append(CondPageBreak(0.96 * inch))
        story.append(_section_heading(route_label.upper(), route_style, dark=True))
        story.append(Spacer(1, 4))
        aggregate_columns = [
            "Style Number", "Product Name", "Garment Color", "Size",
            "Report Decoration Type", "Report Decoration Location",
            "Report Decoration Color",
        ]
        compact = group.groupby(aggregate_columns, dropna=False, sort=False)["Quantity"].sum().reset_index()
        compact = compact.sort_values(
            by=["Style Number", "Product Name", "Garment Color", "Size"],
            key=lambda column: column.map(
                lambda value: size_sort_key(value) if column.name == "Size" else clean_text(value).casefold()
            ),
            kind="stable",
        )
        rows = [[
            _paragraph("Product #", header_style),
            _paragraph("Description", header_style),
            _paragraph("Garment Color", header_style),
            _paragraph("Size", header_style),
            _paragraph("Qty", header_style),
            _paragraph("Decoration", header_style),
            _paragraph("Location", header_style),
            _paragraph("Thread / Ink", header_style),
        ]]
        for _, row in compact.iterrows():
            rows.append([
                _paragraph(row.get("Style Number", ""), body_style),
                _paragraph(row.get("Product Name", ""), body_style),
                _paragraph(row.get("Garment Color", ""), body_style),
                _paragraph(row.get("Size", ""), body_style),
                _paragraph(str(int(row.get("Quantity", 0))), qty_style),
                _paragraph(_display_decoration_type(row.get("Report Decoration Type", "")), body_style),
                _paragraph(row.get("Report Decoration Location", ""), body_style),
                _paragraph(row.get("Report Decoration Color", ""), body_style),
            ])
        rows.append([
            _paragraph("ROUTE TOTAL", header_style), "", "", "",
            _paragraph(str(int(group["Quantity"].sum())), header_style),
            "", "", "", "",
        ])
        table = Table(
            rows,
            colWidths=[1.00*inch, 3.10*inch, 1.35*inch, 0.70*inch, 0.55*inch, 1.30*inch, 1.35*inch, 1.15*inch],
            repeatRows=1, splitByRow=1, hAlign="LEFT",
        )
        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), PURPLE),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("ROWBACKGROUNDS", (0, 1), (-1, -2), [colors.white, ROW_ALT]),
            ("BACKGROUND", (0, -1), (-1, -1), PURPLE_DARK),
            ("TEXTCOLOR", (0, -1), (-1, -1), colors.white),
            ("SPAN", (0, -1), (3, -1)),
            ("SPAN", (5, -1), (7, -1)),
            ("BOX", (0, 0), (-1, -1), 0.7, BORDER),
            ("INNERGRID", (0, 0), (-1, -2), 0.3, BORDER),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("ALIGN", (3, 1), (4, -1), "CENTER"),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        story.append(table)
        story.append(Spacer(1, 9))

    screen_total = int(detail.loc[
        detail["Report Decoration Type"].map(clean_text).str.casefold().str.contains("screen", na=False),
        "Quantity",
    ].sum())
    embroidery_total = int(detail.loc[
        detail["Report Decoration Type"].map(clean_text).str.casefold().str.contains("embroider", na=False),
        "Quantity",
    ].sum())
    outsourced_total = int(detail["Quantity"].sum())
    summary_label = ParagraphStyle(
        "OutSummaryLabel", parent=body_style, fontName="Helvetica-Bold",
        fontSize=8.2, leading=9.6, textColor=PURPLE_DARK, alignment=TA_CENTER,
    )
    summary_value = ParagraphStyle(
        "OutSummaryValue", parent=summary_label, fontSize=15, leading=17,
    )
    story.append(Spacer(1, 8))
    totals_table = Table([
        [_paragraph("OUTSOURCED PIECE TOTALS", route_style), "", ""],
        [
            _paragraph("SCREEN-PRINTED PIECES", summary_label),
            _paragraph("EMBROIDERED PIECES", summary_label),
            _paragraph("TOTAL OUTSOURCED PIECES", summary_label),
        ],
        [
            _paragraph(str(screen_total), summary_value),
            _paragraph(str(embroidery_total), summary_value),
            _paragraph(str(outsourced_total), summary_value),
        ],
    ], colWidths=[3.38 * inch, 3.38 * inch, 3.39 * inch], splitByRow=0)
    totals_table.setStyle(TableStyle([
        ("SPAN", (0, 0), (-1, 0)),
        ("BACKGROUND", (0, 0), (-1, 0), PURPLE_DARK),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("BACKGROUND", (0, 1), (-1, 1), PURPLE_LIGHT),
        ("BOX", (0, 0), (-1, -1), 0.7, BORDER),
        ("INNERGRID", (0, 1), (-1, -1), 0.35, BORDER),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
    ]))
    story.append(totals_table)

    doc.build(story, canvasmaker=PageNumberCanvas)

def generate_employee_totals_pdf(review_workbook_path: Path, output_path: Path | None = None) -> Path:
    """Generate the current Employee Totals PDF without completing Purchase Review."""
    workbook_path = Path(review_workbook_path)
    event_name = clean_text(load_event_name(workbook_path)) or "Uniform Sizing Event"
    records = load_employee_totals(workbook_path)
    if output_path is None:
        output_path = workbook_path.parent / f"{safe_filename(event_name)}__Employee_Totals.pdf"
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    _build_employee_totals_pdf(records, output_path, event_name)
    return output_path


def generate_final_purchase_orders(
    review_workbook_path: Path,
    reports_root: Path,
    po_number_overrides: dict[tuple[str, str], str] | None = None,
    job_logo_path: Path | None = None,
    outsourced_job_name: str = "",
    allow_historical_lock: bool = False,
) -> dict:
    workbook_path = Path(review_workbook_path)
    report_mode = normalize_report_mode(load_report_mode(workbook_path))
    event_name = clean_text(load_event_name(workbook_path))
    decoration_fulfillment = load_decoration_fulfillment(workbook_path)
    selected_job_logo = Path(job_logo_path) if job_logo_path and Path(job_logo_path).exists() else _event_logo_path(workbook_path, event_name)

    # Protected 4.9.0 preflight: verify the locked source CSV, live Product
    # Master snapshot, Never Outsource overrides, and immutable source ledger
    # before any report directory or PDF is created. Archived report-only
    # reopens may use their historical lock; active events must still match the
    # current live routing files exactly.
    review_records = load_review_lines(workbook_path)
    packet_lock_summary = verify_packet_lock(
        workbook_path, review_records,
        require_current_live_inputs=not bool(allow_historical_lock),
    )
    ready, review, non_included = _prepare_lines(workbook_path, decoration_fulfillment)
    integrity_summary = _validate_final_output_integrity(
        review_records, ready, review, non_included
    )
    po_numbers = load_po_numbers(workbook_path)
    if po_number_overrides:
        for key, value in po_number_overrides.items():
            cleaned = clean_text(value)
            if cleaned:
                po_numbers[(clean_text(key[0]).casefold(), _canonical_report_type(key[1]))] = cleaned
    now = datetime.now()
    stamp = now.strftime("%Y%m%d_%H%M%S_%f")
    visible_name = safe_filename(event_name) if event_name else f"General Sales Period {now.strftime('%Y-%m-%d')}"
    purchase_orders_root = Path(reports_root) / "Purchase Orders"
    purchase_orders_root.mkdir(parents=True, exist_ok=True)
    final_output_dir = purchase_orders_root / f"{visible_name} - Purchase Orders__{stamp}"
    # Build into a hidden staging folder. A folder with the final visible name
    # appears only after every PDF and release-manifest hash has succeeded.
    output_dir = purchase_orders_root / f".building-{stamp}"
    output_dir.mkdir(parents=True, exist_ok=False)

    integrity_path = output_dir / "PROTECTED_ORDER_RELEASE_CHECK.txt"
    integrity_path.write_text(
        "PROTECTED PURCHASE ORDER RELEASE CHECK - PASS\n"
        f"Packet Lock ID: {packet_lock_summary['lock_id']}\n"
        f"Packet Lock SHA256: {packet_lock_summary['manifest_sha256']}\n"
        f"Locked source lines: {packet_lock_summary['source_lines']}\n"
        f"Locked source quantity: {packet_lock_summary['source_quantity']}\n"
        f"Source garment lines: {integrity_summary['source_lines']}\n"
        f"Vendor purchase lines: {integrity_summary['vendor_lines']}\n"
        f"Vendor purchase quantity: {integrity_summary['vendor_quantity']}\n"
        f"Non-Included source lines: {integrity_summary['non_included_lines']}\n"
        f"Non-Included quantity: {integrity_summary['non_included_quantity']}\n"
        f"Purchase Review blocked lines: {integrity_summary['review_lines']}\n"
        "Every included garment was traced to exactly one purchasing route.\n",
        encoding="utf-8",
    )
    created_files = [integrity_path]
    pdf_files = []
    summary_rows = []
    non_included_pdf = None
    po_number_by_vendor: dict[str, str] = {}

    if not ready.empty:
        # One consolidated purchase order PDF per vendor. Decoration types and
        # colors remain clearly separated inside that single vendor document.
        report_number = 1
        for vendor_value, report in ready.groupby("Vendor", dropna=False, sort=True):
            vendor = clean_text(vendor_value)
            report_name = "Consolidated Vendor Order" if report_mode == UNIFORM_SIZING_EVENT else "General Sales"
            if report_mode == UNIFORM_SIZING_EVENT:
                name_parts = [safe_filename(event_name)] if event_name else []
                name_parts.extend([safe_filename(vendor), "Purchase_Order"])
                base_name = "__".join(name_parts)
            else:
                base_name = "__".join([safe_filename(vendor), "Purchase_Order"])

            # Prefer a dedicated combined-vendor PO number. For compatibility
            # with existing workbooks, fall back to the first entered PO number
            # for this vendor before creating an automatic number.
            po_number = clean_text(po_numbers.get((vendor.casefold(), "combined vendor order"), ""))
            if not po_number:
                for fallback_type in ("embroidery", "screen printing", "blank garments", "general sales"):
                    po_number = clean_text(po_numbers.get((vendor.casefold(), fallback_type), ""))
                    if po_number:
                        break
            po_number = po_number or f"PO-{now.strftime('%Y%m%d')}-{report_number:02d}"
            po_number_by_vendor[vendor.casefold()] = po_number
            pdf_path = output_dir / f"{safe_filename(po_number)}__{base_name}.pdf"
            _build_pdf(report, pdf_path, po_number, vendor, report_mode, "", event_name=event_name)
            pdf_files.append(pdf_path)
            created_files.append(pdf_path)

            compact = _aggregate_rows(report)
            section_count = int(report[["Decoration Type", "Decoration Location", "Decoration Color"]].drop_duplicates().shape[0])
            summary_rows.append({
                "Vendor": vendor, "Report": report_name, "Purchase Order Mode": report_mode,
                "Event Name": event_name, "Decoration Sections": section_count,
                "Unique Lines": len(compact), "Total Quantity": int(report["Quantity"].sum()),
                "PO Number": po_number,
            })
            report_number += 1

    if not non_included.empty:
        non_included_pdf = output_dir / "Non-Included_Items__Internal_Purchase_Order.pdf"
        _build_non_included_items_pdf(non_included, non_included_pdf, event_name)
        pdf_files.append(non_included_pdf)
        created_files.append(non_included_pdf)
        summary_rows.append({
            "Vendor": "MULTIPLE VENDORS",
            "Report": "Non-Included Items",
            "Purchase Order Mode": "Internal Manual Purchase",
            "Event Name": event_name,
            "Decoration Sections": int(non_included[["Decoration Type", "Decoration Location", "Decoration Color"]].drop_duplicates().shape[0]),
            "Unique Lines": len(_aggregate_rows(non_included)),
            "Total Quantity": int(non_included["Quantity"].sum()),
            "PO Number": "INTERNAL",
        })

    # Decoration and receiving reports are operational documents, separate from
    # the concise vendor purchase orders above.
    in_house_receiving_pdf = None
    outsourced_decoration_pdf = None
    operational_frames = []
    if not ready.empty:
        ready_operational = ready.copy()
        ready_operational["Assigned PO Number"] = ready_operational["Vendor"].map(
            lambda value: po_number_by_vendor.get(clean_text(value).casefold(), "PO Not Assigned")
        )
        ready_operational["Manual Purchase"] = "No"
        operational_frames.append(ready_operational)
    if not non_included.empty:
        manual_operational = non_included.copy()
        manual_operational["Assigned PO Number"] = "MANUAL / NON-INCLUDED"
        manual_operational["Manual Purchase"] = "Yes"
        operational_frames.append(manual_operational)

    if operational_frames:
        operational = pd.concat(operational_frames, ignore_index=True, sort=False)
        outsourced_mask = operational.apply(
            lambda row: is_outsourced_decoration(
                _operational_value(row, "Operational Decoration Type", "Decoration Type"),
                decoration_fulfillment,
                do_not_outsource=_do_not_outsource(row.get("Do Not Outsource", "No")),
            ),
            axis=1,
        )
        outsourced_detail = operational[outsourced_mask].copy()
        in_house_detail = operational[~outsourced_mask].copy()

        if not in_house_detail.empty:
            in_house_receiving_pdf = output_dir / "In-House_Receiving_and_Decoration_Report.pdf"
            _build_in_house_receiving_report(
                in_house_detail, in_house_receiving_pdf, event_name, decoration_fulfillment
            )
            created_files.append(in_house_receiving_pdf)

        if not outsourced_detail.empty:
            outsourced_decoration_pdf = output_dir / "Outsourced_Decoration_Job_Report.pdf"
            _build_outsourced_decoration_report(
                outsourced_detail, outsourced_decoration_pdf, event_name, decoration_fulfillment,
                job_logo_path=selected_job_logo,
                outsourced_job_name=outsourced_job_name,
            )
            created_files.append(outsourced_decoration_pdf)

    # Uniform Sizing Events also receive Employee Totals and a concise Event
    # Summary. Daily All Orders runs intentionally omit these event-only reports.
    employee_totals_pdf = None
    event_summary_pdf = None
    if report_mode == UNIFORM_SIZING_EVENT:
        employee_records = load_employee_totals(workbook_path)
        employee_totals_pdf = output_dir / f"{safe_filename(event_name or 'Event')}__Employee_Totals.pdf"
        event_summary_pdf = output_dir / f"{safe_filename(event_name or 'Event')}__Event_Summary.pdf"
        _build_employee_totals_pdf(employee_records, employee_totals_pdf, event_name)
        event_ready = pd.concat([ready, non_included], ignore_index=True) if not non_included.empty else ready
        _build_event_summary_pdf(
            employee_records,
            event_ready,
            event_summary_pdf,
            event_name,
            decoration_fulfillment,
        )
        created_files.extend([employee_totals_pdf, event_summary_pdf])

    summary = pd.DataFrame(summary_rows)

    # Include the exact locked inputs with the released reports, then hash every
    # output file. This makes the completed ordering folder self-auditing and
    # prevents reports from being silently mixed with a different source or
    # Product Master later.
    copied_lock_manifest = copy_lock_bundle(packet_lock_summary["manifest_path"], output_dir)
    created_files.append(copied_lock_manifest)
    release_manifest_path = output_dir / "FINAL_RELEASE_MANIFEST.json"
    release_files = []
    for path in sorted((item for item in output_dir.rglob("*") if item.is_file()), key=lambda value: str(value)):
        release_files.append({
            "path": str(path.relative_to(output_dir)),
            "sha256": sha256_file(path),
            "size": path.stat().st_size,
        })
    release_manifest = {
        "release_status": "PASS",
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "event_name": event_name,
        "report_mode": report_mode,
        "decoration_fulfillment": decoration_fulfillment,
        "packet_lock_id": packet_lock_summary["lock_id"],
        "packet_lock_manifest_sha256": packet_lock_summary["manifest_sha256"],
        "source_lines": integrity_summary["source_lines"],
        "vendor_quantity": integrity_summary["vendor_quantity"],
        "non_included_quantity": integrity_summary["non_included_quantity"],
        "blocked_review_lines": integrity_summary["review_lines"],
        "files": release_files,
    }
    release_manifest_path.write_text(
        json.dumps(release_manifest, indent=2, sort_keys=True) + "\n", encoding="utf-8"
    )
    created_files.append(release_manifest_path)

    if final_output_dir.exists():
        raise RuntimeError(f"Final report folder already exists: {final_output_dir}")
    output_dir.rename(final_output_dir)

    def released(path):
        if not path:
            return None
        path = Path(path)
        try:
            relative = path.relative_to(output_dir)
        except ValueError:
            return path
        return final_output_dir / relative

    integrity_path = released(integrity_path)
    pdf_files = [released(path) for path in pdf_files]
    created_files = [released(path) for path in created_files]
    non_included_pdf = released(non_included_pdf)
    in_house_receiving_pdf = released(in_house_receiving_pdf)
    outsourced_decoration_pdf = released(outsourced_decoration_pdf)
    employee_totals_pdf = released(employee_totals_pdf)
    event_summary_pdf = released(event_summary_pdf)
    release_manifest_path = released(release_manifest_path)
    output_dir = final_output_dir

    return {
        "integrity_summary": integrity_summary,
        "packet_lock_summary": packet_lock_summary,
        "release_manifest_path": release_manifest_path,
        "integrity_path": integrity_path,
        "output_dir": output_dir,
        "report_mode": report_mode,
        "event_name": event_name,
        "decoration_fulfillment": decoration_fulfillment,
        "report_count": len(summary),
        "decoration_report_count": int(bool(in_house_receiving_pdf)) + int(bool(outsourced_decoration_pdf)),
        "routes": len(summary),
        "ready_lines": len(ready),
        "ready_quantity": int(ready["Quantity"].sum()) if not ready.empty else 0,
        "non_included_lines": len(non_included),
        "non_included_quantity": int(non_included["Quantity"].sum()) if not non_included.empty else 0,
        "non_included_pdf": non_included_pdf,
        "review_lines": len(review),
        "pdf_files": pdf_files,
        "created_files": created_files,
        "employee_totals_pdf": employee_totals_pdf,
        "event_summary_pdf": event_summary_pdf,
        "in_house_receiving_pdf": in_house_receiving_pdf,
        "outsourced_decoration_pdf": outsourced_decoration_pdf,
    }
