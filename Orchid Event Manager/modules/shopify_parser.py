import re
from pathlib import Path
import pandas as pd

STYLE_PATTERN = re.compile(
    r"\b(?=[A-Z0-9-]{2,18}\b)(?=[A-Z0-9-]*[A-Z])(?=[A-Z0-9-]*\d)[A-Z0-9-]+\b",
    re.IGNORECASE,
)


def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_space(value):
    return re.sub(r"\s+", " ", clean_text(value)).strip()


def normalize_style(value):
    return re.sub(r"\s+", "", clean_text(value)).upper()


def canonical_product_name(value):
    name = normalize_space(value)
    return re.sub(r"[\s\.\-–—:|/]+$", "", name).strip()


def valid_known_style(value):
    style = normalize_style(value)
    if not style or len(style) > 18:
        return False
    if not re.fullmatch(r"[A-Z0-9-]+", style):
        return False
    if not any(character.isdigit() for character in style):
        return False
    return True


def load_known_styles(product_master_path):
    product_master_path = Path(product_master_path)
    if not product_master_path.exists():
        return []
    try:
        master = pd.read_csv(product_master_path, dtype=str).fillna("")
    except Exception:
        return []
    if "Style Number" not in master.columns:
        return []
    styles = {
        normalize_style(value)
        for value in master["Style Number"]
        if valid_known_style(value)
    }
    return sorted(styles, key=len, reverse=True)


def find_known_style(text, known_styles):
    upper_text = text.upper()
    for style in known_styles:
        pattern = re.compile(
            rf"(?<![A-Z0-9]){re.escape(style)}(?![A-Z0-9])",
            re.IGNORECASE,
        )
        match = pattern.search(upper_text)
        if match:
            return match.group(0), match.start(), match.end()
    return None


def find_style_candidate(text):
    matches = list(STYLE_PATTERN.finditer(text))
    if not matches:
        return None
    match = matches[-1]
    return match.group(0), match.start(), match.end()


def split_size(text):
    text = clean_text(text)
    if " / " in text:
        before_size, size = text.rsplit(" / ", 1)
        return before_size.strip(), size.strip()
    return text, ""


def clean_segment(text):
    return clean_text(
        re.sub(r"^[\s\-–—:|/\.]+|[\s\-–—:|/\.]+$", "", clean_text(text))
    )


def fallback_dash_parser(text_before_size, size):
    parts = [clean_text(part) for part in text_before_size.split(" - ")]
    if len(parts) >= 3:
        return {
            "Product Name": canonical_product_name(parts[0]),
            "Style Number": normalize_style(parts[1]),
            "Garment Color": normalize_space(" - ".join(parts[2:])),
            "Size": size,
            "Needs Review": "",
        }
    return {
        "Product Name": canonical_product_name(text_before_size),
        "Style Number": "",
        "Garment Color": "",
        "Size": size,
        "Needs Review": "Style number and garment color could not be determined",
    }


def parse_lineitem_name(lineitem_name, known_styles=None):
    known_styles = known_styles or []
    original_text = clean_text(lineitem_name)
    if not original_text:
        return {
            "Product Name": "",
            "Style Number": "",
            "Garment Color": "",
            "Size": "",
            "Needs Review": "Line-item name is blank",
        }

    text_before_size, size = split_size(original_text)
    style_match = find_known_style(text_before_size, known_styles)
    if style_match is None:
        style_match = find_style_candidate(text_before_size)
    if style_match is None:
        return fallback_dash_parser(text_before_size, size)

    style_number, style_start, style_end = style_match
    product_name = canonical_product_name(clean_segment(text_before_size[:style_start]))
    garment_color = normalize_space(clean_segment(text_before_size[style_end:]))
    needs_review = "" if garment_color else "Garment color could not be determined"

    return {
        "Product Name": product_name,
        "Style Number": normalize_style(style_number),
        "Garment Color": garment_color,
        "Size": size,
        "Needs Review": needs_review,
    }


def is_decoration_service(lineitem_name):
    name = clean_text(lineitem_name).lower()
    service_terms = [
        "add embroidered logo",
        "embroidered logo",
        "embroidery fee",
        "screen print",
        "screen-print",
        "screen printing",
        "hemming",
        "alteration",
    ]
    return any(term in name for term in service_terms)


def parse_shopify_orders(csv_path, product_master_path=None):
    orders = pd.read_csv(csv_path, dtype=str).fillna("")
    known_styles = load_known_styles(product_master_path) if product_master_path else []

    for column in ["Name", "Billing Company"]:
        if column in orders.columns:
            orders[column] = orders[column].ffill()

    parsed_rows = []
    for _, row in orders.iterrows():
        lineitem_name = clean_text(row.get("Lineitem name", ""))
        if not lineitem_name or is_decoration_service(lineitem_name):
            continue

        parsed = parse_lineitem_name(lineitem_name, known_styles)
        quantity = pd.to_numeric(row.get("Lineitem quantity", 0), errors="coerce")
        if pd.isna(quantity):
            quantity = 0

        parsed_rows.append(
            {
                "Order Number": clean_text(row.get("Name", "")),
                "Company": clean_text(row.get("Billing Company", "")),
                "Original Line Item": lineitem_name,
                "Product Name": parsed["Product Name"],
                "Style Number": parsed["Style Number"],
                "Garment Color": parsed["Garment Color"],
                "Size": parsed["Size"],
                "Quantity": int(quantity),
                "Needs Review": parsed["Needs Review"],
            }
        )

    return pd.DataFrame(parsed_rows)
