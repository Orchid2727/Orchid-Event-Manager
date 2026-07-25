from __future__ import annotations

from datetime import datetime
from pathlib import Path
import re

import pandas as pd

from modules.blank_garment_rules import is_blank_decoration
from modules.decoration_locations import normalize_decoration_location
from modules.purchase_rules import apply_purchase_rule_defaults, normalize_bool, row_rules

from modules.shopify_parser import parse_shopify_orders
from modules.routing_rules import apply_routing_overrides
from modules.pdf_page_numbers import PageNumberCanvas

try:
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
    from reportlab.lib.pagesizes import landscape, letter
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import inch
    from reportlab.platypus import (
        KeepTogether,
        PageBreak,
        Paragraph,
        SimpleDocTemplate,
        Spacer,
        Table,
        TableStyle,
    )
except ImportError as exc:  # pragma: no cover - handled at runtime on the user's Mac
    raise ImportError(
        "PDF purchase orders require ReportLab. Install it once with: "
        "/Library/Frameworks/Python.framework/Versions/3.13/bin/python3 -m pip install reportlab"
    ) from exc


MASTER_COLUMNS = [
    "Product Name",
    "Style Number",
    "Garment Color",
    "Vendor",
    "Decoration Type",
    "Decoration Location",
    "Decoration Placement Instructions",
    "Decoration Color",
    "Product Category",
    "Requires Size",
    "Requires Color",
    "Requires Decoration",
]

OUTPUT_COLUMNS = [
    "Vendor",
    "Decoration Type",
    "Decoration Location",
    "Decoration Placement Instructions",
    "Decoration Color",
    "Style Number",
    "Product Name",
    "Garment Color",
    "Size",
    "Quantity",
]

PURPLE = colors.HexColor("#5B2AA8")
PURPLE_DARK = colors.HexColor("#3D176F")
PURPLE_LIGHT = colors.HexColor("#F2ECFB")
BORDER = colors.HexColor("#D7C8EE")
TEXT = colors.HexColor("#201A2D")
MUTED = colors.HexColor("#6F667A")
ROW_ALT = colors.HexColor("#FAF7FE")

SIZE_ORDER = [
    "OSFA", "YXS", "YS", "YM", "YL", "YXL",
    "XXS", "XS", "S", "M", "L", "XL",
    "2XL", "3XL", "4XL", "5XL", "6XL", "7XL", "8XL",
    "LT", "XLT", "2XLT", "3XLT", "4XLT", "5XLT", "6XLT",
]


def clean_text(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value)
    # Normalize punctuation that may arrive from legacy Shopify/Excel exports
    # as Windows-1252 control characters. ReportLab's default fonts cannot
    # render those bytes and otherwise show a black replacement square.
    replacements = {
        "\x91": "'", "\x92": "'", "‘": "'", "’": "'",
        "\x93": '"', "\x94": '"', "“": '"', "”": '"',
        "–": "-", "—": "-", "…": "...", "\xa0": " ",
    }
    for source, replacement in replacements.items():
        text = text.replace(source, replacement)
    return text.strip()


def normalize_style(value) -> str:
    return re.sub(r"\s+", "", clean_text(value)).upper()


def normalize_color(value) -> str:
    return re.sub(r"\s+", " ", clean_text(value)).strip().casefold()


def canonical_product(value) -> str:
    name = re.sub(r"\s+", " ", clean_text(value)).strip()
    return re.sub(r"[\s\.\-–—:|/]+$", "", name).strip().casefold()


def identity_key(style, product, color) -> str:
    style_key = normalize_style(style)
    color_key = normalize_color(color)
    # The temporary Tru-Spec 24/7 number is shared by several different pant
    # families. Include the product description so they remain distinct.
    if style_key == "24/7":
        return f"style:{style_key}|product:{canonical_product(product)}|color:{color_key}"
    return f"style:{style_key}|color:{color_key}" if style_key else f"product:{canonical_product(product)}|color:{color_key}"


def normalize_size(value) -> str:
    raw = clean_text(value).upper()
    size = re.sub(r"\s+", "", raw)
    if re.fullmatch(r"\d{2}[X/]\d{2}", size):
        return size.replace("/", "x").replace("X", "x")
    tall_match = re.fullmatch(r"(L|XL|2X|2XL|3X|3XL|4X|4XL|5X|5XL|6X|6XL)(?:TALL|TL|T)", size)
    if tall_match:
        base = tall_match.group(1).replace("XL", "X")
        return {"L": "LT", "X": "XLT", "2X": "2XLT", "3X": "3XLT", "4X": "4XLT", "5X": "5XLT", "6X": "6XLT"}.get(base, size)
    aliases = {
        "2X": "2XL",
        "XXL": "2XL",
        "3X": "3XL",
        "XXXL": "3XL",
        "4X": "4XL",
        "5X": "5XL",
        "6X": "6XL",
        "ONESIZE": "OSFA",
    }
    return aliases.get(size, size)


def size_sort_key(value) -> tuple[int, str]:
    size = normalize_size(value)
    try:
        return (SIZE_ORDER.index(size), size)
    except ValueError:
        return (len(SIZE_ORDER), size)



def split_embedded_size(garment_color: str, size: str) -> tuple[str, str]:
    """Recover a clear size embedded in a manually entered color field.

    Existing Size values remain authoritative. Recognized examples include
    ``Portwest Size 38``, ``Black 2X Tall``, ``Navy XL``, and ``Khaki 38x32``.
    Numeric values without a size label are accepted only as waist/inseam pairs,
    preventing prices, quantities, and style numbers from being guessed as sizes.
    """
    color = clean_text(garment_color)
    current_size = normalize_size(size)
    if current_size or not color:
        return color, current_size

    token_map = {
        "XXS": "XXS", "XS": "XS", "S": "S", "SMALL": "S",
        "M": "M", "MEDIUM": "M", "L": "L", "LARGE": "L",
        "XL": "XL", "X-LARGE": "XL", "2XL": "2XL", "2X": "2XL",
        "XXL": "2XL", "3XL": "3XL", "3X": "3XL", "4XL": "4XL",
        "4X": "4XL", "5XL": "5XL", "5X": "5XL", "6XL": "6XL",
        "6X": "6XL", "OSFA": "OSFA", "ONE SIZE": "OSFA",
        "LT": "LT", "L TALL": "LT", "L T": "LT",
        "XLT": "XLT", "XL TALL": "XLT", "XL T": "XLT",
        "2XLT": "2XLT", "2XL TALL": "2XLT", "2X TALL": "2XLT", "2XL T": "2XLT", "2X TL": "2XLT",
        "3XLT": "3XLT", "3XL TALL": "3XLT", "3X TALL": "3XLT", "3XL T": "3XLT", "3X TL": "3XLT",
        "4XLT": "4XLT", "4XL TALL": "4XLT", "4X TALL": "4XLT", "4XL T": "4XLT", "4X TL": "4XLT",
    }

    size_value = (
        r"(?:\d{2}\s*[xX/]\s*\d{2}|\d{1,3}|XXS|XS|S|SMALL|M|MEDIUM|L|LARGE|"
        r"XL|X-LARGE|2XL|2X|XXL|3XL|3X|XXXL|4XL|4X|5XL|5X|6XL|6X|"
        r"LT|L\s+(?:TALL|TL|T)|XLT|XL\s+(?:TALL|TL|T)|2XLT|2XL?\s+(?:TALL|TL|T)|"
        r"3XLT|3XL?\s+(?:TALL|TL|T)|4XLT|4XL?\s+(?:TALL|TL|T)|OSFA|ONE\s+SIZE)"
    )
    labelled = re.search(
        rf"(?i)(?<![A-Za-z0-9])(?:SIZE|SZ)\s*[:#-]?\s*(?P<size>{size_value})(?![A-Za-z0-9])",
        color,
    )
    if labelled:
        normalized = normalize_size(labelled.group("size"))
        remaining = (color[:labelled.start()] + " " + color[labelled.end():]).strip()
        remaining = re.sub(r"\s{2,}", " ", remaining).strip(" -/:,|")
        return remaining, normalized

    # A waist/inseam pair is unambiguous even when it trails the color.
    waist = re.search(r"(?i)(?<!\d)(?P<size>\d{2}\s*[xX/]\s*\d{2})(?!\d)\s*$", color)
    if waist:
        remaining = color[:waist.start()].strip(" -/:,|")
        return remaining, normalize_size(waist.group("size"))

    # Text sizes are safe at either edge. Sort longest first so "2X Tall" is
    # captured before "2X".
    for token in sorted(token_map, key=len, reverse=True):
        escaped = re.escape(token).replace(r"\ ", r"\s+")
        leading = re.match(rf"(?i)^({escaped})(?:\s+|$)", color)
        if leading:
            remaining = color[leading.end():].strip(" -/:,|")
            return remaining, token_map[token]
        trailing = re.search(rf"(?i)(?:^|\s)({escaped})$", color)
        if trailing:
            remaining = color[:trailing.start()].strip(" -/:,|")
            return remaining, token_map[token]
    return color, current_size

def safe_filename(value) -> str:
    value = clean_text(value) or "Unassigned"
    value = re.sub(r"[^A-Za-z0-9._-]+", "_", value)
    return value.strip("_") or "Unassigned"


def load_master(product_master_path: Path) -> pd.DataFrame:
    master = pd.read_csv(product_master_path, dtype=str).fillna("")
    for column in MASTER_COLUMNS:
        if column not in master.columns:
            master[column] = ""
    master = apply_purchase_rule_defaults(master[MASTER_COLUMNS].copy())
    master = apply_routing_overrides(master)
    master["_identity_key"] = master.apply(lambda row: identity_key(row["Style Number"], row["Product Name"], row["Garment Color"]), axis=1)
    master["_score"] = (
        master["Vendor"].map(lambda x: bool(clean_text(x))).astype(int) * 4
        + master["Decoration Type"].map(lambda x: bool(clean_text(x))).astype(int) * 2
        + master["Decoration Color"].map(lambda x: bool(clean_text(x))).astype(int)
    )
    master = master.sort_values("_score", ascending=False)
    return master.drop_duplicates(["_identity_key"], keep="first")


def build_purchase_order_data(
    shopify_csv_path: Path,
    product_master_path: Path,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    parsed = parse_shopify_orders(shopify_csv_path, product_master_path)
    master = load_master(product_master_path)

    parsed = parsed.copy()
    parsed["_identity_key"] = parsed.apply(lambda row: identity_key(row["Style Number"], row["Product Name"], row["Garment Color"]), axis=1)

    merged = parsed.merge(
        master,
        on=["_identity_key"],
        how="left",
        suffixes=("", "_master"),
    )

    for column in ["Vendor", "Decoration Type", "Decoration Location", "Decoration Placement Instructions", "Decoration Color"]:
        merged[column] = merged[column].fillna("").map(clean_text)

    merged = apply_routing_overrides(merged)

    for column in ["Style Number", "Product Name", "Garment Color", "Size", "Company"]:
        if column not in merged.columns:
            merged[column] = ""
        merged[column] = merged[column].map(clean_text)

    merged["Size"] = merged["Size"].map(normalize_size)
    corrected = merged.apply(
        lambda row: split_embedded_size(row["Garment Color"], row["Size"]),
        axis=1,
        result_type="expand",
    )
    corrected.columns = ["Garment Color", "Size"]
    merged[["Garment Color", "Size"]] = corrected
    vendor_only = merged.apply(
        lambda row: bool(
            clean_text(row.get("Garment Color", ""))
            and clean_text(row.get("Garment Color", "")).casefold()
            == clean_text(row.get("Vendor", "")).casefold()
        ),
        axis=1,
    )
    merged.loc[vendor_only, "Garment Color"] = ""
    merged["Quantity"] = pd.to_numeric(
        merged["Quantity"], errors="coerce"
    ).fillna(0).astype(int)

    missing_reason = []
    for _, row in merged.iterrows():
        reasons = []
        if clean_text(row.get("Needs Review", "")):
            reasons.append(clean_text(row["Needs Review"]))
        rules = row_rules(row)
        requires_color = normalize_bool(rules["Requires Color"], True)
        requires_decoration = normalize_bool(rules["Requires Decoration"], True)
        if not clean_text(row["Style Number"]):
            reasons.append("Missing style number")
        if requires_color and not clean_text(row["Garment Color"]):
            reasons.append("Missing garment color")
        if not clean_text(row["Vendor"]):
            reasons.append("Vendor not assigned in Product Master")
        if requires_decoration and not clean_text(row["Decoration Type"]):
            reasons.append("Decoration type not assigned in Product Master")
        decoration_type = clean_text(row["Decoration Type"]).casefold()
        merged.at[row.name, "Decoration Location"] = normalize_decoration_location(
            row.get("Decoration Location", ""), row.get("Decoration Type", ""),
            requires_decoration=requires_decoration,
        )
        if requires_decoration and decoration_type and not is_blank_decoration(decoration_type) and not clean_text(row["Decoration Color"]):
            reasons.append("Decoration color not assigned in Product Master")
        missing_reason.append("; ".join(dict.fromkeys(reasons)))

    merged["Review Reason"] = missing_reason
    ready_detail = merged[merged["Review Reason"].eq("")].copy()
    review = merged[merged["Review Reason"].ne("")].copy()

    if ready_detail.empty:
        grouped = pd.DataFrame(columns=OUTPUT_COLUMNS)
    else:
        grouped = (
            ready_detail.groupby(
                [
                    "Vendor",
                    "Decoration Type",
                    "Decoration Location",
                    "Decoration Color",
                    "Style Number",
                    "Product Name",
                    "Garment Color",
                    "Size",
                ],
                dropna=False,
                as_index=False,
            )["Quantity"]
            .sum()
            .reset_index(drop=True)
        )
        grouped["_size_sort"] = grouped["Size"].map(size_sort_key)
        grouped = grouped.sort_values(
            [
                "Vendor",
                "Decoration Type",
                "Decoration Location",
                "Decoration Color",
                "Style Number",
                "Garment Color",
                "_size_sort",
            ],
            key=lambda series: series.astype(str).str.casefold()
            if series.name != "_size_sort"
            else series,
        ).drop(columns="_size_sort").reset_index(drop=True)

    review_columns = [
        "Order Number",
        "Company",
        "Original Line Item",
        "Product Name",
        "Style Number",
        "Garment Color",
        "Size",
        "Quantity",
        "Review Reason",
    ]
    review = (
        review[review_columns].copy()
        if not review.empty
        else pd.DataFrame(columns=review_columns)
    )

    summary_rows = []
    if not grouped.empty:
        for (vendor, decoration_type, decoration_location, decoration_color), group in grouped.groupby(
            ["Vendor", "Decoration Type", "Decoration Location", "Decoration Color"], dropna=False
        ):
            summary_rows.append(
                {
                    "Vendor": vendor,
                    "Decoration Type": decoration_type,
                    "Decoration Location": decoration_location,
                    "Decoration Color": decoration_color,
                    "Unique Lines": len(group),
                    "Total Quantity": int(group["Quantity"].sum()),
                }
            )

    summary = pd.DataFrame(
        summary_rows,
        columns=[
            "Vendor",
            "Decoration Type",
            "Decoration Location",
            "Decoration Color",
            "Unique Lines",
            "Total Quantity",
        ],
    )
    return grouped, review, summary, ready_detail


def _route_title(decoration_type: str, decoration_location: str, decoration_color: str) -> str:
    decoration_type = clean_text(decoration_type)
    decoration_location = normalize_decoration_location(decoration_location, decoration_type)
    decoration_color = clean_text(decoration_color)
    if is_blank_decoration(decoration_type):
        return "Blank Garments"
    color_label = "Ink" if "screen" in decoration_type.casefold() else "Thread"
    return f"{decoration_type} - {decoration_location} - {decoration_color} {color_label}"


def _build_pdf(
    route_group: pd.DataFrame,
    route_detail: pd.DataFrame,
    pdf_path: Path,
    po_number: str,
    vendor: str,
    decoration_type: str,
    decoration_location: str,
    decoration_color: str,
) -> None:
    doc = SimpleDocTemplate(
        str(pdf_path),
        pagesize=landscape(letter),
        rightMargin=0.4 * inch,
        leftMargin=0.4 * inch,
        topMargin=0.38 * inch,
        bottomMargin=0.42 * inch,
        title=f"Purchase Order {po_number}",
        author="Orchid Uniforms & Apparel",
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "OrchidTitle",
        parent=styles["Title"],
        fontName="Helvetica-Bold",
        fontSize=20,
        leading=22,
        textColor=PURPLE_DARK,
        alignment=TA_LEFT,
        spaceAfter=2,
    )
    subtitle_style = ParagraphStyle(
        "OrchidSubtitle",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=8.5,
        textColor=MUTED,
        leading=10,
    )
    label_style = ParagraphStyle(
        "Label",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=8,
        textColor=MUTED,
        leading=10,
    )
    value_style = ParagraphStyle(
        "Value",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=10,
        textColor=TEXT,
        leading=12,
    )
    section_style = ParagraphStyle(
        "Section",
        parent=styles["Heading2"],
        fontName="Helvetica-Bold",
        fontSize=10,
        textColor=PURPLE_DARK,
        leading=12,
        spaceBefore=5,
        spaceAfter=4,
    )
    small_style = ParagraphStyle(
        "Small",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=7.3,
        leading=9,
        textColor=TEXT,
    )
    small_right = ParagraphStyle(
        "SmallRight",
        parent=small_style,
        alignment=TA_RIGHT,
    )

    story = []

    header_left = [
        Paragraph("ORCHID UNIFORMS &amp; APPAREL", title_style),
        Paragraph(
            "1304 Cornell Parkway, Suite B | Oklahoma City, OK 73108 | (405) 947-2388",
            subtitle_style,
        ),
    ]
    header_right = Table(
        [
            [Paragraph("PURCHASE ORDER", ParagraphStyle(
                "POHeading", parent=title_style, alignment=TA_RIGHT, fontSize=16
            ))],
            [Paragraph(po_number, ParagraphStyle(
                "PONumber", parent=value_style, alignment=TA_RIGHT, textColor=PURPLE
            ))],
        ],
        colWidths=[2.7 * inch],
    )
    header = Table([[header_left, header_right]], colWidths=[7.4 * inch, 2.7 * inch])
    header.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LINEBELOW", (0, 0), (-1, -1), 1.2, PURPLE),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
    ]))
    story.append(header)
    story.append(Spacer(1, 8))

    order_date = datetime.now().strftime("%B %d, %Y")
    route_text = _route_title(decoration_type, decoration_location, decoration_color)
    total_qty = int(route_group["Quantity"].sum())

    info = Table(
        [
            [
                Paragraph("VENDOR", label_style),
                Paragraph("ROUTE", label_style),
                Paragraph("ORDER DATE", label_style),
                Paragraph("TOTAL GARMENTS", label_style),
            ],
            [
                Paragraph(clean_text(vendor), value_style),
                Paragraph(route_text, value_style),
                Paragraph(order_date, value_style),
                Paragraph(str(total_qty), value_style),
            ],
        ],
        colWidths=[2.4 * inch, 3.5 * inch, 2.2 * inch, 2.1 * inch],
    )
    info.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), PURPLE_LIGHT),
        ("BOX", (0, 0), (-1, -1), 0.7, BORDER),
        ("INNERGRID", (0, 0), (-1, -1), 0.4, BORDER),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 7),
        ("RIGHTPADDING", (0, 0), (-1, -1), 7),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    story.append(info)
    story.append(Spacer(1, 10))

    for (style, product, garment_color), item_group in route_group.groupby(
        ["Style Number", "Product Name", "Garment Color"], dropna=False, sort=False
    ):
        sizes = sorted(
            [clean_text(s) for s in item_group["Size"].unique() if clean_text(s)],
            key=size_sort_key,
        )
        if not sizes:
            sizes = ["UNSPECIFIED"]

        quantity_map = {
            clean_text(row["Size"]): int(row["Quantity"])
            for _, row in item_group.iterrows()
        }

        available_width = 10.1 * inch
        fixed_widths = [0.9 * inch, 2.7 * inch, 1.55 * inch]
        remaining = available_width - sum(fixed_widths) - 0.75 * inch
        size_width = max(0.42 * inch, min(0.65 * inch, remaining / max(len(sizes), 1)))
        col_widths = fixed_widths + [size_width] * len(sizes) + [0.75 * inch]

        header_row = ["Style", "Description", "Garment Color"] + sizes + ["Total"]
        total = int(item_group["Quantity"].sum())
        data_row = [
            clean_text(style),
            clean_text(product),
            clean_text(garment_color),
        ] + [quantity_map.get(size, 0) for size in sizes] + [total]

        matrix = Table([header_row, data_row], colWidths=col_widths, repeatRows=1)
        matrix.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), PURPLE),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, 0), 7.3),
            ("FONTNAME", (0, 1), (-1, 1), "Helvetica"),
            ("FONTSIZE", (0, 1), (-1, 1), 7.4),
            ("BACKGROUND", (0, 1), (-1, 1), colors.white),
            ("BOX", (0, 0), (-1, -1), 0.7, BORDER),
            ("INNERGRID", (0, 0), (-1, -1), 0.35, BORDER),
            ("ALIGN", (3, 0), (-1, -1), "CENTER"),
            ("ALIGN", (-1, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ("FONTNAME", (-1, 1), (-1, 1), "Helvetica-Bold"),
            ("TEXTCOLOR", (-1, 1), (-1, 1), PURPLE_DARK),
        ]))

        detail_item = route_detail[
            (route_detail["Style Number"].map(clean_text) == clean_text(style))
            & (route_detail["Product Name"].map(clean_text) == clean_text(product))
            & (route_detail["Garment Color"].map(clean_text) == clean_text(garment_color))
        ].copy()

        breakdown_rows = [[
            Paragraph("Company / Customer", label_style),
            Paragraph("Quantity", label_style),
        ]]
        if not detail_item.empty:
            detail_item["Company Label"] = detail_item.apply(
                lambda row: clean_text(row.get("Company", ""))
                or (f"Order {clean_text(row.get('Order Number', ''))}" if clean_text(row.get("Order Number", "")) else "Unspecified"),
                axis=1,
            )
            company_group = (
                detail_item.groupby("Company Label", as_index=False)["Quantity"]
                .sum()
                .sort_values("Company Label", key=lambda s: s.astype(str).str.casefold())
            )
            for _, company_row in company_group.iterrows():
                breakdown_rows.append([
                    Paragraph(clean_text(company_row["Company Label"]), small_style),
                    Paragraph(str(int(company_row["Quantity"])), small_right),
                ])

        breakdown = Table(breakdown_rows, colWidths=[3.5 * inch, 0.8 * inch])
        breakdown.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), PURPLE_LIGHT),
            ("BOX", (0, 0), (-1, -1), 0.5, BORDER),
            ("INNERGRID", (0, 0), (-1, -1), 0.3, BORDER),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 5),
            ("RIGHTPADDING", (0, 0), (-1, -1), 5),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, ROW_ALT]),
        ]))

        block = [
            matrix,
            Spacer(1, 4),
            Paragraph("Customer Breakdown", section_style),
            breakdown,
            Spacer(1, 10),
        ]
        story.append(KeepTogether(block))

    notes = Table(
        [[
            Paragraph("NOTES", label_style),
            Paragraph(
                "Verify style, garment color, size totals, and decoration route before placing the order.",
                small_style,
            ),
        ]],
        colWidths=[0.8 * inch, 9.3 * inch],
    )
    notes.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), PURPLE_LIGHT),
        ("BOX", (0, 0), (-1, -1), 0.6, BORDER),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 7),
        ("RIGHTPADDING", (0, 0), (-1, -1), 7),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    story.append(notes)

    def add_page_footer(canvas, doc_obj):
        canvas.saveState()
        canvas.setFont("Helvetica", 7.5)
        canvas.setFillColor(MUTED)
        canvas.drawString(0.45 * inch, 0.22 * inch, f"Orchid Uniforms & Apparel - {po_number}")
        canvas.restoreState()

    doc.build(
        story,
        onFirstPage=add_page_footer,
        onLaterPages=add_page_footer,
        canvasmaker=PageNumberCanvas,
    )


def generate_purchase_orders(
    shopify_csv_path: Path,
    product_master_path: Path,
    reports_root: Path,
) -> dict:
    grouped, review, summary, ready_detail = build_purchase_order_data(
        Path(shopify_csv_path), Path(product_master_path)
    )

    now = datetime.now()
    stamp = now.strftime("%Y%m%d_%H%M%S")
    output_dir = Path(reports_root) / "Purchase Orders" / f"Legacy__{stamp}"
    output_dir.mkdir(parents=True, exist_ok=True)

    created_files = []
    pdf_files = []

    summary_path = output_dir / "Purchase_Order_Summary.csv"
    summary.to_csv(summary_path, index=False)
    created_files.append(summary_path)

    if not review.empty:
        review_path = output_dir / "Review_Needed.csv"
        review.to_csv(review_path, index=False)
        created_files.append(review_path)

    if not grouped.empty:
        route_number = 1
        for (vendor, decoration_type, decoration_location, decoration_color), group in grouped.groupby(
            ["Vendor", "Decoration Type", "Decoration Location", "Decoration Color"],
            dropna=False,
            sort=True,
        ):
            parts = [safe_filename(vendor), safe_filename(decoration_type)]
            if not is_blank_decoration(decoration_type):
                parts.extend([safe_filename(decoration_location), safe_filename(decoration_color)])
            base_name = "__".join(parts)

            csv_path = output_dir / f"{base_name}.csv"
            group[OUTPUT_COLUMNS].to_csv(csv_path, index=False)
            created_files.append(csv_path)

            route_detail = ready_detail[
                (ready_detail["Vendor"].map(clean_text) == clean_text(vendor))
                & (ready_detail["Decoration Type"].map(clean_text) == clean_text(decoration_type))
                & (ready_detail["Decoration Location"].map(clean_text) == clean_text(decoration_location))
                & (ready_detail["Decoration Color"].map(clean_text) == clean_text(decoration_color))
            ].copy()

            po_number = f"PO-{now.strftime('%Y%m%d')}-{route_number:02d}"
            pdf_path = output_dir / f"{po_number}__{base_name}.pdf"
            _build_pdf(
                group,
                route_detail,
                pdf_path,
                po_number,
                clean_text(vendor),
                clean_text(decoration_type),
                clean_text(decoration_location),
                clean_text(decoration_color),
            )
            created_files.append(pdf_path)
            pdf_files.append(pdf_path)
            route_number += 1

    return {
        "output_dir": output_dir,
        "files": created_files,
        "pdf_files": pdf_files,
        "ready_lines": len(grouped),
        "ready_quantity": int(grouped["Quantity"].sum()) if not grouped.empty else 0,
        "review_lines": len(review),
        "routes": len(summary),
    }
