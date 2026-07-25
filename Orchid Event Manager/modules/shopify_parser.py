from __future__ import annotations

import re
from pathlib import Path
import pandas as pd

from modules.routing_rules import apply_parsed_product_overrides
from modules.smart_parser import load_master_index, parse_manual_custom_item
from modules.internal_services import infer_in_house_decoration, is_decoration_charge, is_internal_service_style
from modules.source_identity import build_source_id, source_base_signature

SIZE_ALIASES = {
    "XXS": "XXS", "XS": "XS", "XSMALL": "XS", "EXTRA SMALL": "XS",
    "S": "S", "SMALL": "S", "M": "M", "MED": "M", "MEDIUM": "M",
    "L": "L", "LARGE": "L", "XL": "XL", "X-LARGE": "XL", "XLARGE": "XL",
    "XXL": "2XL", "2XL": "2XL", "2X": "2XL", "XXXL": "3XL",
    "3XL": "3XL", "3X": "3XL", "4XL": "4XL", "4X": "4XL",
    "5XL": "5XL", "5X": "5XL", "6XL": "6XL", "6X": "6XL",
    "7XL": "7XL", "7X": "7XL", "8XL": "8XL", "8X": "8XL",
    "LT": "LT", "L TALL": "LT", "L T": "LT",
    "XLT": "XLT", "XL TALL": "XLT", "XL T": "XLT",
    "2XLT": "2XLT", "2XTL": "2XLT", "2XL TALL": "2XLT", "2XL T": "2XLT", "2XLTL": "2XLT",
    "3XLT": "3XLT", "3XTL": "3XLT", "3XL TALL": "3XLT", "3XL T": "3XLT", "3XLTL": "3XLT",
    "4XLT": "4XLT", "4XTL": "4XLT", "4XL TALL": "4XLT", "4XL T": "4XLT", "4XLTL": "4XLT",
    "5XLT": "5XLT", "5XTL": "5XLT", "5XL TALL": "5XLT", "5XL T": "5XLT", "5XLTL": "5XLT",
    "6XLT": "6XLT", "6XTL": "6XLT", "6XL TALL": "6XLT", "6XL T": "6XLT", "6XLTL": "6XLT",
    "OSFA": "OSFA", "OSFM": "OSFM",
    "ONE SIZE": "OSFA", "ONESIZE": "OSFA", "S/M": "S/M",
    "SM": "S/M", "M/L": "M/L", "L/XL": "L/XL", "UH": "UH",
}

SERVICE_TERMS = [
    "add embroidered logo", "embroidered logo", "embroidery fee",
    "screen print", "screen-print", "screen printing", "hemming", "hem service",
    "alteration", "alterations", "inseam alteration", "one color screen-printed logo",
    "sew on patch", "sew-on patch", "flag patches on", "patch application",
    "name/title/id", "name title id",
    "embroidery - left chest logos", "embroidery left chest logos",
]

FORBIDDEN_STYLE_WORDS = {
    "BLACK", "WHITE", "NAVY", "KHAKI", "SILVER", "GREY", "GRAY",
    "RED", "ROYAL", "ORANGE", "MAROON", "GREEN", "BLUE", "CHARCOAL",
    "COYOTE", "OLIVE", "BROWN", "TAN", "GOLD", "PURPLE", "PINK",
    "YELLOW", "FOSSIL", "SAGE", "GRILL", "CARBON", "MOSS", "STONE",
    "TEMP", "HAT", "PANT", "SHIRT", "POLO", "JACKET", "VEST",
    "CAP", "TEE",
}

PRODUCT_WORDS = {
    "shirt", "pant", "pants", "jean", "cap", "hat", "jacket", "vest",
    "polo", "tee", "t-shirt", "sweatshirt", "hoodie", "coverall",
    "pullover", "fleece", "parka", "shell", "short", "shorts", "men's",
    "womens", "women's", "mens", "workwear", "uniform", "top", "coat",
}

PLACEHOLDER_VARIANT_VALUES = {"temp", "temporary", "tbd", "placeholder", "default title"}


def valid_variant_value(value: object) -> str:
    cleaned = clean(value)
    return "" if cleaned.casefold() in PLACEHOLDER_VARIANT_VALUES else cleaned


def extract_variant_color_size(option_1: object, option_2: object, option_3: object, variant_title: object) -> tuple[str, str]:
    options = [valid_variant_value(value) for value in (option_1, option_2, option_3)]
    options = [value for value in options if value]
    color = ""
    size = ""

    # Waist and inseam may be stored as adjacent option values.
    if len(options) >= 3 and re.fullmatch(r"\d{2}", options[1]) and re.fullmatch(r"\d{2}|UH", options[2], re.I):
        color = options[0]
        size = f"{options[1]}x{options[2]}" if options[2].upper() != "UH" else f"{options[1]}/UH"
    else:
        for value in options:
            if not size and is_size(value):
                size = normalize_size(value)
            elif not color:
                color = value

    title = valid_variant_value(variant_title)
    if title and (not color or not size):
        pieces = [valid_variant_value(piece) for piece in title.split("/")]
        pieces = [piece for piece in pieces if piece]
        size_positions = [index for index, piece in enumerate(pieces) if is_size(piece)]
        if size_positions and not size:
            size = normalize_size(pieces[size_positions[0]])
        if not color:
            non_sizes = [piece for piece in pieces if not is_size(piece)]
            if non_sizes:
                color = " / ".join(non_sizes)

    return valid_variant_value(color), normalize_size(size)


def clean(value) -> str:
    if value is None or pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value).strip())


def canonical_product(value) -> str:
    return re.sub(r"[\s\.\-–—:|/]+$", "", clean(value)).strip()


def norm_style(value) -> str:
    return re.sub(r"\s+", "", clean(value)).upper()


def normalize_size(value) -> str:
    raw = clean(value)
    upper = raw.upper()
    compact = re.sub(r"\s+", "", upper)
    tall_match = re.fullmatch(r"(L|XL|[2-8]X(?:L)?)(?:\s*(?:TALL|TL|T))", upper)
    if tall_match:
        base = tall_match.group(1).replace("XL", "X")
        return {"L": "LT", "X": "XLT", **{f"{number}X": f"{number}XLT" for number in range(2, 9)}}.get(base, compact)
    if upper in SIZE_ALIASES:
        return SIZE_ALIASES[upper]
    if compact in SIZE_ALIASES:
        return SIZE_ALIASES[compact]
    match = re.fullmatch(r"(\d{2})\s*[xX]\s*(\d{2})", raw)
    if match:
        return f"{match.group(1)}x{match.group(2)}"
    if re.fullmatch(r"\d{4}", compact):
        waist, inseam = int(compact[:2]), int(compact[2:])
        if 24 <= waist <= 60 and 24 <= inseam <= 40:
            return f"{waist}x{inseam}"
    return raw


def is_size(value) -> bool:
    normalized = normalize_size(value)
    if normalized.upper() in set(SIZE_ALIASES.values()):
        return True
    return bool(
        re.fullmatch(r"\d{2}x\d{2}", normalized, re.I)
        or re.fullmatch(r"\d{2}/UH", normalized, re.I)
    )


def valid_style(value) -> bool:
    raw = clean(value)
    if not raw or re.search(r"\s", raw):
        return False
    style = norm_style(raw)
    if len(style) > 18 or is_size(style) or style in FORBIDDEN_STYLE_WORDS:
        return False
    if not re.fullmatch(r"[A-Z0-9-]+", style):
        return False
    if style.isdigit():
        return len(style) >= 3
    return any(character.isdigit() for character in style)


def explicit_style_segment(value: object, *, has_product_context: bool = False) -> bool:
    """Return True for an explicit style token in a product-title position.

    Four-digit numeric garment styles such as 5028, 5326, and 5328 can also
    look like compact waist/inseam sizes. They are accepted as styles only when
    the surrounding title provides product context, rather than everywhere a
    size may appear.
    """
    raw = clean(value)
    if valid_style(raw):
        return True
    return bool(has_product_context and re.fullmatch(r"\d{3,6}", raw))


def _lineitem_has_internal_service_style(lineitem_name: object) -> bool:
    text = clean(lineitem_name).upper()
    if not text:
        return False
    # Product 750 is Orchid's embroidery-charge code. Require a standalone token
    # so a garment style such as 7500 is not accidentally removed.
    return any(
        re.search(rf"(?<![A-Z0-9]){re.escape(code)}(?![A-Z0-9])", text)
        for code in {"750"}
    )


def is_decoration_service(lineitem_name, style_number: object = "", *variant_values: object) -> bool:
    name = clean(lineitem_name).lower()
    return (
        is_internal_service_style(style_number)
        or _lineitem_has_internal_service_style(lineitem_name)
        or is_decoration_charge(lineitem_name, *variant_values)
        or any(term in name for term in SERVICE_TERMS)
        or bool(infer_in_house_decoration(name))
    )


def split_spaced_slashes(text: str) -> list[str]:
    return [clean(part) for part in re.split(r"\s+/\s+", clean(text))]


def split_hyphens(text: str) -> list[str]:
    # Split delimiter hyphens only when at least one side has whitespace.
    # This preserves hyphens inside true style numbers.
    return [
        clean(part)
        for part in re.split(r"(?:\s+-\s*|\s*-\s+)", clean(text))
        if clean(part)
    ]


def looks_like_product_description(value: str) -> bool:
    words = clean(value).lower().split()
    return len(words) >= 4 or any(word.strip("®™.,()") in PRODUCT_WORDS for word in words)


def trailing_style(segment: str):
    segment = clean(segment)
    match = re.search(r"(?<![A-Z0-9])([A-Z0-9][A-Z0-9-]{1,17})$", segment, re.I)
    if not match:
        return None
    prefix = canonical_product(segment[: match.start(1)])
    if not explicit_style_segment(match.group(1), has_product_context=bool(re.search(r"[A-Za-z]", prefix))):
        return None
    return norm_style(match.group(1)), prefix


def compact_style_size_color(text: str):
    size_pattern = (
        r"XL\s*Tall|2XL\s*Tall|3XL\s*Tall|One\s*Size|"
        r"XXS|XS|XXL|[2-8]XL|XLT|2XLT|3XLT|OSFA|OSFM|"
        r"Small|Medium|Large|S|M|L|XL"
    )
    match = re.match(
        rf"^([A-Z0-9-]{{2,18}})\s+({size_pattern})\s+(.+)$",
        clean(text),
        re.I,
    )
    if not match or not valid_style(match.group(1)):
        return None
    return {
        "Product Name": norm_style(match.group(1)),
        "Style Number": norm_style(match.group(1)),
        "Garment Color": clean(match.group(3)),
        "Size": normalize_size(match.group(2)),
        "Needs Review": "",
    }



def split_trailing_size(value: str) -> tuple[str, str]:
    text = clean(value)
    size_pattern = (
        r"(?:One\s*Size|XXS|XS|XXL|[2-8]XL|(?:L|XL|[2-8]X(?:L)?)\s+(?:Tall|TL|T)|"
        r"XLT|[2-8]XLT|OSFA|OSFM|"
        r"S/M|M/L|L/XL|Small|Medium|Large|S|M|L|XL|"
        r"\d{2}\s*[xX]\s*\d{2}|\d{4}|\d{2}/UH)"
    )
    match = re.match(rf"^(.*?)\s+({size_pattern})$", text, re.I)
    if not match or not is_size(match.group(2)):
        return text, ""
    return clean(match.group(1)), normalize_size(match.group(2))


def parse_lineitem_name(lineitem_name, known_styles=None):
    original = clean(lineitem_name)
    if not original:
        return {
            "Product Name": "", "Style Number": "", "Garment Color": "",
            "Size": "", "Needs Review": "Line-item name is blank",
        }

    compact = compact_style_size_color(original)
    if compact:
        return compact

    slash_parts = split_spaced_slashes(original)
    core = slash_parts[0]
    tail = slash_parts[1:]
    parts = split_hyphens(core)
    size = ""
    color_from_tail = ""

    # STYLE/Product - waist / inseam / optional color
    if (
        len(parts) >= 2
        and tail
        and re.fullmatch(r"\d{2}", parts[-1])
        and re.fullmatch(r"\d{2}", tail[0])
    ):
        size = f"{parts[-1]}x{tail[0]}"
        parts = parts[:-1]
        if len(tail) > 1:
            color_from_tail = " / ".join(tail[1:])
    elif len(tail) >= 2 and re.fullmatch(r"\d{2}", tail[0]):
        if re.fullmatch(r"\d{2}", tail[1]):
            size = f"{tail[0]}x{tail[1]}"
            if len(tail) > 2:
                color_from_tail = " / ".join(tail[2:])
        elif tail[1].upper() == "UH":
            size = f"{tail[0]}/UH"
            if len(tail) > 2:
                color_from_tail = " / ".join(tail[2:])
        else:
            size = normalize_size(tail[-1])
            color_from_tail = " / ".join(tail[:-1])
    elif len(tail) == 1:
        size = normalize_size(tail[0])
    elif tail:
        size = normalize_size(tail[-1])
        color_from_tail = " / ".join(tail[:-1])

    # Last hyphen segment can be a size: 1104 - Ascent - LE Green - 36x32.
    if parts and is_size(parts[-1]):
        if not size:
            size = normalize_size(parts[-1])
        parts = parts[:-1]

    # STYLE - Size - Color.
    if len(parts) >= 3 and valid_style(parts[0]) and is_size(parts[1]):
        return {
            "Product Name": norm_style(parts[0]),
            "Style Number": norm_style(parts[0]),
            "Garment Color": clean(" - ".join(parts[2:])),
            "Size": normalize_size(parts[1]),
            "Needs Review": "",
        }

    # Full hyphen segment style, such as Product - 212477 - Color.
    style_index = next((
        i for i, part in enumerate(parts)
        if explicit_style_segment(
            part,
            has_product_context=bool(
                (i > 0 and any(re.search(r"[A-Za-z]", previous) for previous in parts[:i]))
                or (i == 0 and len(parts) > 1 and any(re.search(r"[A-Za-z]", later) for later in parts[1:]))
            ),
        )
    ), None)
    if style_index is not None:
        style = norm_style(parts[style_index])
        before = parts[:style_index]
        after = parts[style_index + 1:]

        if style_index == 0:
            if len(after) >= 2:
                product = canonical_product(" - ".join(after[:-1]))
                color = clean(after[-1])
            elif len(after) == 1:
                if color_from_tail:
                    product = canonical_product(after[0])
                    color = color_from_tail
                elif looks_like_product_description(after[0]):
                    product = canonical_product(after[0])
                    color = ""
                else:
                    product = style
                    color = clean(after[0])
            else:
                product, color = style, ""
        else:
            product = canonical_product(" - ".join(before))
            after = [part for part in after if part.upper() != "TEMP"]
            color = clean(" - ".join(after))

        if color_from_tail:
            color = color_from_tail
        needs = []
        if not color:
            needs.append("Garment color could not be determined")
        if not size:
            needs.append("Size could not be determined")
        return {
            "Product Name": product or style,
            "Style Number": style,
            "Garment Color": color,
            "Size": size,
            "Needs Review": "; ".join(needs),
        }

    # Description ending in a style: Port Authority ... CP90 - Color.
    # Search every description segment because brand names may contain their own
    # delimiter hyphen before the segment that ends with the true style number.
    if parts:
        for part_index, part in enumerate(parts):
            found = trailing_style(part)
            if not found:
                continue
            style, product_tail = found
            product_parts = [*parts[:part_index], product_tail]
            product = canonical_product(" - ".join(value for value in product_parts if clean(value)))
            color = clean(" - ".join(parts[part_index + 1:]))
            if color_from_tail:
                color = color_from_tail
            needs = []
            if not color:
                needs.append("Garment color could not be determined")
            if not size:
                needs.append("Size could not be determined")
            return {
                "Product Name": product or style,
                "Style Number": style,
                "Garment Color": color,
                "Size": size,
                "Needs Review": "; ".join(needs),
            }

    # Style first with no hyphen: 6572 Navy, MCK01341 Tour Blue Heather 2XL,
    # or 02MCWST 32 x 30.
    first_match = re.match(r"^([A-Z0-9-]{2,18})\s+(.+)$", core, re.I)
    if first_match and explicit_style_segment(
        first_match.group(1),
        has_product_context=bool(re.search(r"[A-Za-z]", first_match.group(2))),
    ):
        style = norm_style(first_match.group(1))
        rest = clean(first_match.group(2))
        if is_size(rest):
            return {
                "Product Name": style,
                "Style Number": style,
                "Garment Color": "",
                "Size": normalize_size(rest),
                "Needs Review": "Garment color could not be determined",
            }
        rest_color, trailing_size = split_trailing_size(rest)
        resolved_size = size or trailing_size
        resolved_color = color_from_tail or rest_color
        needs = []
        if not resolved_color:
            needs.append("Garment color could not be determined")
        if not resolved_size:
            needs.append("Size could not be determined")
        return {
            "Product Name": style,
            "Style Number": style,
            "Garment Color": resolved_color,
            "Size": resolved_size,
            "Needs Review": "; ".join(needs),
        }

    # No reliable style number. Keep separate by product name, never by color/size.
    if len(parts) >= 2:
        product = canonical_product(" - ".join(parts[:-1]))
        color = clean(parts[-1])
    else:
        product = canonical_product(core)
        color = ""
    if color_from_tail:
        color = color_from_tail

    needs = ["Style number could not be determined"]
    if not color:
        needs.append("Garment color could not be determined")
    if not size:
        needs.append("Size could not be determined")
    return {
        "Product Name": product,
        "Style Number": "",
        "Garment Color": color,
        "Size": size,
        "Needs Review": "; ".join(needs),
    }


def clean_text(value):
    return clean(value)


def normalize_style(value):
    return norm_style(value)


def canonical_product_name(value):
    return canonical_product(value)


def load_known_styles(product_master_path):
    try:
        master = pd.read_csv(product_master_path, dtype=str).fillna("")
    except Exception:
        return []
    return sorted(
        {norm_style(value) for value in master.get("Style Number", []) if valid_style(value)},
        key=len,
        reverse=True,
    )



class OrderExportError(ValueError):
    """Raised when an uploaded order export cannot be mapped safely."""


def _coalesce_columns(frame: pd.DataFrame, aliases: list[str]) -> pd.Series:
    result = pd.Series([""] * len(frame), index=frame.index, dtype="object")
    for alias in aliases:
        if alias not in frame.columns:
            continue
        candidate = frame[alias].fillna("").astype(str).map(clean)
        result = result.where(result.astype(str).str.strip().ne(""), candidate)
    return result.fillna("")


def _has_any_column(frame: pd.DataFrame, aliases: list[str]) -> bool:
    return any(alias in frame.columns for alias in aliases)


def _joined_customer_name(frame: pd.DataFrame) -> pd.Series:
    first = _coalesce_columns(frame, [
        "Customer / Address / First name",
        "Customer / First name",
        "Billing address / First name",
    ])
    last = _coalesce_columns(frame, [
        "Customer / Address / Last name",
        "Customer / Last name",
        "Billing address / Last name",
    ])
    return (first + " " + last).map(clean)


def _build_report_toaster_line_name(frame: pd.DataFrame) -> pd.Series:
    product = _coalesce_columns(frame, ["Line item / Product title", "Line item / Title"])
    variant = _coalesce_columns(frame, ["Line item / Variant / Title", "Line item / Variant title"])
    option_1 = _coalesce_columns(frame, ["Line item / Variant / Option 1"])
    option_2 = _coalesce_columns(frame, ["Line item / Variant / Option 2"])
    option_3 = _coalesce_columns(frame, ["Line item / Variant / Option 3"])

    names: list[str] = []
    for index in frame.index:
        title = clean(product.loc[index])
        variant_title = clean(variant.loc[index])
        options = [clean(option_1.loc[index]), clean(option_2.loc[index]), clean(option_3.loc[index])]
        options = [value for value in options if value and value.casefold() != "default title"]
        if not variant_title or variant_title.casefold() == "default title":
            variant_title = " / ".join(options)
        if variant_title and variant_title.casefold() != "default title":
            names.append(clean(f"{title} - {variant_title}"))
        else:
            names.append(title)
    return pd.Series(names, index=frame.index, dtype="object")


def normalize_order_export(csv_path) -> pd.DataFrame:
    """Return either Shopify or Report Toaster data in Shopify's canonical columns.

    Report Toaster is preferred when edited/removed Shopify line items must be
    excluded. Its Net quantity is mapped to Lineitem quantity and rows with a net
    quantity of zero are removed before purchasing data is created.
    """
    csv_path = Path(csv_path)
    try:
        source = pd.read_csv(csv_path, dtype=str).fillna("")
    except Exception as error:
        raise OrderExportError(f"The CSV could not be read: {error}") from error

    if source.empty:
        raise OrderExportError("The selected CSV contains no order rows.")

    is_shopify = "Name" in source.columns and "Lineitem name" in source.columns
    is_report_toaster = "Order name" in source.columns and _has_any_column(
        source, ["Line item / Product title", "Line item / Title"]
    )

    if not is_shopify and not is_report_toaster:
        expected = (
            "Shopify: Name, Lineitem name, Lineitem quantity\n"
            "Report Toaster: Order name, Line item / Product title, and "
            "Line item / Net quantity (or Current quantity)"
        )
        raise OrderExportError(
            "This CSV is not a recognized Shopify or Report Toaster order export.\n\n"
            f"Expected columns include:\n{expected}"
        )

    if is_shopify:
        if "Lineitem quantity" not in source.columns:
            raise OrderExportError(
                "The Shopify CSV is missing the Lineitem quantity column."
            )
        result = source.copy()
        result["Import Source"] = "Shopify"
        # Standard Shopify order exports often populate order-level fields only
        # on the first line. Fill them within each order for reliable downstream use.
        result["Name"] = result["Name"].replace("", pd.NA).ffill().fillna("")
        for column in ["Billing Company", "Billing Name", "Notes", "Created at", "Total", "Subtotal"]:
            if column in result.columns:
                result[column] = (
                    result[column].replace("", pd.NA)
                    .groupby(result["Name"], dropna=False)
                    .transform(lambda values: values.ffill().bfill())
                    .fillna("")
                )
        result["Line Notes"] = _coalesce_columns(result, [
            "Lineitem properties", "Lineitem property", "Lineitem note", "Line item note",
            "Line item properties", "Line item / Properties", "Line item / Note",
            "Line item / Custom attributes", "Lineitem custom attributes",
        ])
        variant_title = _coalesce_columns(result, [
            "Lineitem variant title", "Lineitem variant", "Variant title",
            "Line item / Variant / Title", "Line item / Variant title",
        ])
        option_1 = _coalesce_columns(result, [
            "Lineitem option1", "Lineitem option 1", "Option1 Value", "Variant Option 1",
        ])
        option_2 = _coalesce_columns(result, [
            "Lineitem option2", "Lineitem option 2", "Option2 Value", "Variant Option 2",
        ])
        option_3 = _coalesce_columns(result, [
            "Lineitem option3", "Lineitem option 3", "Option3 Value", "Variant Option 3",
        ])
        colors, sizes = [], []
        for index in result.index:
            color, size = extract_variant_color_size(
                option_1.loc[index], option_2.loc[index], option_3.loc[index], variant_title.loc[index]
            )
            colors.append(color)
            sizes.append(size)
        result["Variant Color"] = colors
        result["Variant Size"] = sizes
        return result

    quantity_aliases = [
        "Line item / Net quantity",
        "Line item / Current quantity",
        "Line item / Quantity",
    ]
    if not _has_any_column(source, quantity_aliases):
        raise OrderExportError(
            "The Report Toaster CSV is missing Net quantity, Current quantity, and Quantity."
        )

    result = pd.DataFrame(index=source.index)
    result["Name"] = _coalesce_columns(source, ["Order name"])
    result["Created at"] = _coalesce_columns(source, ["Created at", "Order created at"])
    result["Billing Company"] = _coalesce_columns(source, [
        "Billing address / Company",
        "Customer / Address / Company",
        "Company",
    ])
    result["Billing Name"] = _joined_customer_name(source)
    result["Notes"] = _coalesce_columns(source, ["Note", "Notes"])
    result["Line Notes"] = _coalesce_columns(source, [
        "Line item / Properties", "Line item / Note", "Line item / Custom attributes",
        "Line item properties", "Lineitem properties", "Lineitem note",
    ])
    result["Lineitem name"] = _build_report_toaster_line_name(source)
    result["Lineitem quantity"] = _coalesce_columns(source, quantity_aliases)
    result["Lineitem price"] = _coalesce_columns(source, ["Line item / Price"])
    result["Lineitem discount"] = _coalesce_columns(source, ["Line item / Discounts"])
    result["Lineitem sku"] = _coalesce_columns(source, [
        "Line item / Variant / SKU",
        "Line item / SKU",
    ])
    result["Vendor"] = _coalesce_columns(source, [
        "Line item / Product / Vendor",
        "Line item / Vendor",
    ])
    result["Variant Title"] = _coalesce_columns(source, [
        "Line item / Variant / Title",
        "Line item / Variant title",
    ])
    result["Variant Option 1"] = _coalesce_columns(source, ["Line item / Variant / Option 1"])
    result["Variant Option 2"] = _coalesce_columns(source, ["Line item / Variant / Option 2"])
    result["Variant Option 3"] = _coalesce_columns(source, ["Line item / Variant / Option 3"])
    result["Lineitem net sales"] = _coalesce_columns(source, ["Line item / Net sales"])
    result["Import Source"] = "Report Toaster"

    result["Name"] = result["Name"].replace("", pd.NA).ffill().fillna("")
    for column in ["Billing Company", "Billing Name", "Notes", "Created at"]:
        result[column] = (
            result[column].replace("", pd.NA)
            .groupby(result["Name"], dropna=False)
            .transform(lambda values: values.ffill().bfill())
            .fillna("")
        )

    quantity = pd.to_numeric(result["Lineitem quantity"], errors="coerce").fillna(0)
    original_quantity = pd.to_numeric(
        _coalesce_columns(source, ["Line item / Quantity"]), errors="coerce"
    ).fillna(quantity)
    quantity_adjustments = int((original_quantity != quantity).sum())
    result["Lineitem quantity"] = quantity.astype(int)
    # Net/current quantity is the key protection against ordering removed items.
    result = result[result["Lineitem quantity"] > 0].copy()

    net_sales_raw = pd.to_numeric(result["Lineitem net sales"], errors="coerce")
    line_price = pd.to_numeric(result["Lineitem price"], errors="coerce").fillna(0.0)
    line_discount = pd.to_numeric(result["Lineitem discount"], errors="coerce").fillna(0.0)
    computed_net_sales = line_price * pd.to_numeric(result["Lineitem quantity"], errors="coerce").fillna(0) - line_discount
    net_sales = net_sales_raw.where(net_sales_raw.notna(), computed_net_sales).fillna(0.0)
    result["Lineitem net sales"] = net_sales
    order_totals = result.groupby("Name", dropna=False)["Lineitem net sales"].transform("sum")
    result["Total"] = order_totals
    result["Subtotal"] = order_totals

    colors: list[str] = []
    sizes: list[str] = []
    for _, row in result.iterrows():
        color, size = extract_variant_color_size(
            row.get("Variant Option 1", ""), row.get("Variant Option 2", ""),
            row.get("Variant Option 3", ""), row.get("Variant Title", ""),
        )
        colors.append(color)
        sizes.append(size)

    result["Variant Color"] = colors
    result["Variant Size"] = sizes
    result.attrs["import_source"] = "Report Toaster"
    result.attrs["removed_rows_excluded"] = int((quantity <= 0).sum())
    result.attrs["quantity_adjustments_applied"] = quantity_adjustments
    return result.reset_index(drop=True)


def _review_reason_after_variant_override(parsed: dict) -> str:
    reasons: list[str] = []
    if not clean(parsed.get("Style Number", "")):
        reasons.append("Style number could not be determined")
    if not clean(parsed.get("Garment Color", "")):
        reasons.append("Garment color could not be determined")
    if not clean(parsed.get("Size", "")):
        reasons.append("Size could not be determined")
    return "; ".join(reasons)


def parse_shopify_orders(csv_path, product_master_path=None, normalized_orders: pd.DataFrame | None = None):
    orders = normalized_orders.copy(deep=True) if normalized_orders is not None else normalize_order_export(csv_path)
    master_index = load_master_index(product_master_path)
    for column in ["Billing Company", "Billing Name", "Notes", "Created at"]:
        if column in orders.columns:
            orders[column] = (
                orders[column].replace("", pd.NA)
                .groupby(orders["Name"], dropna=False)
                .transform(lambda values: values.ffill().bfill())
                .fillna("")
            )

    rows = []
    source_occurrences: dict[str, int] = {}
    for _, row in orders.iterrows():
        lineitem_name = clean_text(row.get("Lineitem name", ""))
        variant_values = (
            row.get("Variant Color", ""), row.get("Variant Size", ""),
            row.get("Variant Option 1", ""), row.get("Variant Option 2", ""),
            row.get("Variant Option 3", ""), row.get("Variant Title", ""),
        )
        if not lineitem_name or is_decoration_service(
            lineitem_name, row.get("Lineitem sku", ""), *variant_values
        ):
            continue

        parsed = parse_lineitem_name(lineitem_name)
        direct_color = clean_text(row.get("Variant Color", ""))
        direct_size = normalize_size(row.get("Variant Size", ""))
        if direct_color:
            parsed["Garment Color"] = direct_color
        if direct_size:
            parsed["Size"] = direct_size
        if direct_color or direct_size:
            parsed["Needs Review"] = _review_reason_after_variant_override(parsed)

        smart = parse_manual_custom_item(
            lineitem_name,
            master_index,
            vendor_hint=row.get("Vendor", ""),
            sku_hint=row.get("Lineitem sku", ""),
        )
        parser_source = "Report Toaster Product" if clean_text(row.get("Import Source", "")) == "Report Toaster" else "Shopify Product"
        parse_confidence = "High" if not parsed.get("Needs Review", "") else "Medium"
        detected_vendor = clean_text(row.get("Vendor", ""))

        if smart:
            normal_missing = sum(
                not clean_text(parsed.get(field, ""))
                for field in ["Style Number", "Garment Color", "Size"]
            )
            smart_found = sum(
                bool(clean_text(smart.get(field, "")))
                for field in ["Style Number", "Garment Color", "Size"]
            )
            explicit_custom_shape = bool(re.search(r"(?i)\bsize\s*\d{2}\s*[x×/]\s*\d{2}\b", lineitem_name))
            if explicit_custom_shape or smart_found > (3 - normal_missing):
                parsed = smart
                if direct_color:
                    parsed["Garment Color"] = direct_color
                if direct_size:
                    parsed["Size"] = direct_size
                parsed["Needs Review"] = _review_reason_after_variant_override(parsed)
                parser_source = smart.get("Parser Source", "Smart Custom Item")
                parse_confidence = smart.get("Parse Confidence", "Medium")
                detected_vendor = smart.get("Detected Vendor", detected_vendor)

        quantity = pd.to_numeric(row.get("Lineitem quantity", 0), errors="coerce")
        quantity = 0 if pd.isna(quantity) else int(quantity)
        if quantity <= 0:
            continue
        immutable_source = {
            "Name": clean_text(row.get("Name", "")),
            "Created at": clean_text(row.get("Created at", "")),
            "Billing Company": clean_text(row.get("Billing Company", "")),
            "Billing Name": clean_text(row.get("Billing Name", "")),
            "Notes": clean_text(row.get("Notes", "")),
            "Line Notes": clean_text(row.get("Line Notes", "")),
            "Lineitem name": lineitem_name,
            "Lineitem quantity": quantity,
            "Lineitem sku": clean_text(row.get("Lineitem sku", "")),
            "Variant Title": clean_text(row.get("Variant Title", "")),
            "Variant Option 1": clean_text(row.get("Variant Option 1", "")),
            "Variant Option 2": clean_text(row.get("Variant Option 2", "")),
            "Variant Option 3": clean_text(row.get("Variant Option 3", "")),
            "Variant Color": clean_text(row.get("Variant Color", "")),
            "Variant Size": clean_text(row.get("Variant Size", "")),
        }
        base_signature = source_base_signature(immutable_source)
        source_occurrence = source_occurrences.get(base_signature, 0) + 1
        source_occurrences[base_signature] = source_occurrence
        source_id = build_source_id(immutable_source, source_occurrence)

        parsed_record = apply_parsed_product_overrides({
            "Source ID": source_id,
            "Source Occurrence": source_occurrence,
            "Original Quantity": quantity,
            "Original Order Number": clean_text(row.get("Name", "")),
            "Original Order Date": clean_text(row.get("Created at", "")),
            "Original Company": clean_text(row.get("Billing Company", "")),
            "Original Employee Name": clean_text(row.get("Billing Name", "")),
            "Original Shopify SKU": clean_text(row.get("Lineitem sku", "")),
            "Original Variant Title": clean_text(row.get("Variant Title", "")),
            "Original Variant Option 1": clean_text(row.get("Variant Option 1", "")),
            "Original Variant Option 2": clean_text(row.get("Variant Option 2", "")),
            "Original Variant Option 3": clean_text(row.get("Variant Option 3", "")),
            "Original Variant Color": clean_text(row.get("Variant Color", "")),
            "Original Variant Size": clean_text(row.get("Variant Size", "")),
            "Order Number": clean_text(row.get("Name", "")),
            "Order Date": clean_text(row.get("Created at", "")),
            "Company": clean_text(row.get("Billing Company", "")),
            "Employee Name": clean_text(row.get("Billing Name", "")),
            "Shopify Order Notes": clean_text(row.get("Notes", "")),
            "Shopify Line Notes": clean_text(row.get("Line Notes", "")),
            "Original Line Item": lineitem_name,
            "Product Name": parsed.get("Product Name", ""),
            "Style Number": parsed.get("Style Number", ""),
            "Garment Color": parsed.get("Garment Color", ""),
            "Size": parsed.get("Size", ""),
            "Quantity": quantity,
            "Needs Review": parsed.get("Needs Review", ""),
            "Parser Source": parser_source,
            "Parse Confidence": parse_confidence,
            "Detected Vendor": detected_vendor,
            "Shopify SKU": clean_text(row.get("Lineitem sku", "")),
        })
        rows.append(parsed_record)
    result = pd.DataFrame(rows)
    result.attrs["import_source"] = orders.attrs.get("import_source", clean_text(orders.get("Import Source", pd.Series([""])).iloc[0] if not orders.empty else ""))
    result.attrs["removed_rows_excluded"] = int(orders.attrs.get("removed_rows_excluded", 0) or 0)
    result.attrs["quantity_adjustments_applied"] = int(orders.attrs.get("quantity_adjustments_applied", 0) or 0)
    return result
