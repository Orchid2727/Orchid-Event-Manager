from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

import pandas as pd


VENDOR_ALIASES = {
    "tru spec": "Tru-Spec",
    "tru-spec": "Tru-Spec",
    "truspec": "Tru-Spec",
    "wrangler": "Wrangler",
    "sanmar": "SanMar",
    "s&s": "S&S Activewear",
    "s & s": "S&S Activewear",
    "ss activewear": "S&S Activewear",
    "cutter buck": "Cutter & Buck",
    "cutter & buck": "Cutter & Buck",
    "outdoor cap": "Outdoor Cap",
    "burnside": "Burnside",
    "vf": "VF",
    "richardson": "Richardson",
}

COMMON_COLORS = (
    "Antique Indigo", "Safety Yellow/Black", "Shady Blue Dark Heather",
    "Flat Dark Earth", "Bungee Cord", "Deep Navy", "Neon Orange",
    "Heather Grey/Navy", "Heather Gray/Navy", "Light Blue", "Safety Green",
    "Dark Khaki", "Blue Nights", "Pastel Mint", "Pastel Blue",
    "Steel", "Scarlet", "Royal", "Graphite", "Charcoal", "Silver",
    "Khaki", "Black", "White", "Navy", "Maroon", "Orange", "Red",
    "Caviar", "Morel", "Saddle", "Grey", "Gray", "Ash", "Sand",
    "LE Green", "Brown", "Purple", "Pink", "Gold", "Green", "Blue",
)

GARMENT_WORDS = {
    "pant", "pants", "jean", "jeans", "shirt", "tee", "t-shirt", "polo",
    "jacket", "coat", "vest", "cap", "hat", "coverall", "short", "shorts",
    "hoodie", "sweatshirt", "pullover", "fleece", "shell", "uniform",
}

SIZE_PATTERN = re.compile(
    r"(?i)(?:\bsize\s*)?(\d{2})\s*[x×/]\s*(\d{2})\b|"
    r"\b((?:L|XL|[2-8]X(?:L)?)\s+(?:TALL|TL|T)|XXS|XS|S|M|L|XL|XXL|"
    r"2XL|3XL|4XL|5XL|6XL|7XL|8XL|LT|XLT|[2-8]XLT|OSFA|OSFM|ONE\s*SIZE|S/M|L/XL)\b"
)

STYLE_TOKEN_PATTERN = re.compile(r"(?i)^[A-Z0-9][A-Z0-9/-]{1,17}$")


def clean(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value).strip())


def normalize_style(value: object) -> str:
    return re.sub(r"\s+", "", clean(value)).upper()


def normalize_size(value: object) -> str:
    raw = clean(value).upper().replace("×", "X")
    raw = re.sub(r"\s+", "", raw)
    aliases = {
        "XXL": "2XL", "2X": "2XL", "3X": "3XL", "4X": "4XL",
        "5X": "5XL", "6X": "6XL", "ONESIZE": "OSFA",
    }
    if re.fullmatch(r"\d{2}[X/]\d{2}", raw):
        return raw.replace("/", "x").replace("X", "x")
    tall_match = re.fullmatch(r"(L|XL|[2-8]X(?:L)?)(?:TALL|TL|T)", raw)
    if tall_match:
        base = tall_match.group(1).replace("XL", "X")
        return {"L": "LT", "X": "XLT", **{f"{value}X": f"{value}XLT" for value in range(2, 9)}}.get(base, raw)
    return aliases.get(raw, raw)


def _normalized_phrase(value: object) -> str:
    text = clean(value).casefold()
    text = text.replace("®", " ").replace("™", " ")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _contains_phrase(text: str, phrase: str) -> bool:
    return f" {_normalized_phrase(phrase)} " in f" {_normalized_phrase(text)} "


@dataclass
class MasterIndex:
    frame: pd.DataFrame
    styles: set[str]
    by_style: dict[str, pd.DataFrame]
    colors: list[str]


_MASTER_INDEX_CACHE_KEY: tuple[str, int, int] | None = None
_MASTER_INDEX_CACHE_VALUE: MasterIndex | None = None


def load_master_index(path: Path | None) -> MasterIndex:
    global _MASTER_INDEX_CACHE_KEY, _MASTER_INDEX_CACHE_VALUE
    columns = ["Product Name", "Style Number", "Garment Color", "Vendor", "Product Aliases", "Vendor Color Code", "Color Aliases"]
    resolved = Path(path).resolve() if path else None
    if resolved and resolved.exists():
        stat = resolved.stat()
        cache_key = (str(resolved), int(stat.st_mtime_ns), int(stat.st_size))
    else:
        cache_key = (str(resolved or ""), 0, 0)
    if cache_key == _MASTER_INDEX_CACHE_KEY and _MASTER_INDEX_CACHE_VALUE is not None:
        return _MASTER_INDEX_CACHE_VALUE
    if not path or not Path(path).exists():
        frame = pd.DataFrame(columns=columns)
    else:
        try:
            frame = pd.read_csv(path, dtype=str).fillna("")
        except Exception:
            frame = pd.DataFrame(columns=columns)
    for column in columns:
        if column not in frame.columns:
            frame[column] = ""
    frame = frame[columns].copy()
    frame["_style"] = frame["Style Number"].map(normalize_style)
    styles = {value for value in frame["_style"] if value}
    by_style = {style: group.copy() for style, group in frame.groupby("_style", sort=False) if style}
    alias_colors = set()
    for _, row in frame.iterrows():
        alias_colors.update(clean(value) for value in [row.get("Garment Color", ""), row.get("Vendor Color Code", "")] if clean(value))
        alias_colors.update(clean(value) for value in re.split(r"[|;,\n]+", clean(row.get("Color Aliases", ""))) if clean(value))
    colors = sorted(
        alias_colors | set(COMMON_COLORS),
        key=lambda value: (-len(_normalized_phrase(value)), value.casefold()),
    )
    result = MasterIndex(frame=frame, styles=styles, by_style=by_style, colors=colors)
    _MASTER_INDEX_CACHE_KEY = cache_key
    _MASTER_INDEX_CACHE_VALUE = result
    return result


def _extract_size(text: str) -> tuple[str, str]:
    matches = list(SIZE_PATTERN.finditer(text))
    if not matches:
        return "", text
    match = matches[-1]
    if match.group(1) and match.group(2):
        size = f"{match.group(1)}x{match.group(2)}"
    else:
        size = normalize_size(match.group(3) or "")
    remainder = clean(text[:match.start()] + " " + text[match.end():])
    remainder = re.sub(r"(?i)\bsize\b", " ", remainder)
    return size, clean(remainder)


def _style_token_candidates(text: str) -> list[str]:
    candidates = []
    for raw in re.split(r"\s+", clean(text)):
        token = raw.strip(".,;:()[]{}")
        if not STYLE_TOKEN_PATTERN.fullmatch(token):
            continue
        normalized = normalize_style(token)
        if re.fullmatch(r"\d{2}", normalized):
            continue
        if normalized in {"SIZE", "TEMP"}:
            continue
        candidates.append(normalized)
    return candidates


def _detect_vendor(text: str, master_rows: pd.DataFrame | None, vendor_hint: str = "") -> str:
    if master_rows is not None and not master_rows.empty:
        vendors = [clean(value) for value in master_rows["Vendor"] if clean(value)]
        if vendors:
            return vendors[0]
    normalized = _normalized_phrase(" ".join([vendor_hint, text]))
    for alias, vendor in sorted(VENDOR_ALIASES.items(), key=lambda item: -len(item[0])):
        if f" {alias} " in f" {normalized} ":
            return vendor
    return clean(vendor_hint)


def _detect_color(text: str, master_rows: pd.DataFrame | None, all_colors: list[str]) -> str:
    candidates = []
    display_by_alias = {}
    if master_rows is not None and not master_rows.empty:
        for _, row in master_rows.iterrows():
            display = clean(row.get("Garment Color", "")) or clean(row.get("Vendor Color Code", ""))
            row_terms = [clean(row.get("Garment Color", "")), clean(row.get("Vendor Color Code", ""))]
            row_terms.extend(clean(value) for value in re.split(r"[|;,\n]+", clean(row.get("Color Aliases", ""))) if clean(value))
            for term in row_terms:
                if term:
                    candidates.append(term)
                    display_by_alias[_normalized_phrase(term)] = display or term
    candidates.extend(all_colors)
    seen = set()
    for color in sorted(candidates, key=lambda value: -len(_normalized_phrase(value))):
        key = _normalized_phrase(color)
        if not key or key in seen:
            continue
        seen.add(key)
        if _contains_phrase(text, color):
            return display_by_alias.get(key, color)
    return ""

def _remove_phrase(text: str, phrase: str) -> str:
    if not phrase:
        return clean(text)
    pattern = re.compile(re.escape(phrase), re.I)
    return clean(pattern.sub(" ", text, count=1))


def parse_manual_custom_item(
    lineitem_name: object,
    master_index: MasterIndex,
    vendor_hint: object = "",
    sku_hint: object = "",
) -> dict | None:
    original = clean(lineitem_name)
    if not original:
        return None

    size, without_size = _extract_size(original)
    tokens = _style_token_candidates(without_size)
    style = ""
    for token in tokens:
        if token in master_index.styles:
            style = token
            break
    if not style and tokens:
        first_original_token = normalize_style(original.split()[0]) if original.split() else ""
        if tokens[0] == first_original_token:
            style = tokens[0]

    normalized = _normalized_phrase(original)
    garment_signal = any(f" {word} " in f" {normalized} " for word in GARMENT_WORDS)
    vendor_signal = any(f" {alias} " in f" {normalized} " for alias in VENDOR_ALIASES)
    manual_shape = bool(style and size and (garment_signal or vendor_signal))
    known_style_shape = bool(style in master_index.styles and (size or garment_signal))
    if not (manual_shape or known_style_shape):
        return None

    master_rows = master_index.by_style.get(style)
    vendor = _detect_vendor(original, master_rows, clean(vendor_hint))
    color = _detect_color(original, master_rows, master_index.colors)

    description = without_size
    if style:
        description = re.sub(rf"(?i)^\s*{re.escape(style)}\b", " ", description, count=1)
    description = re.sub(r"(?i)\bsize\b", " ", description)
    description = _remove_phrase(description, color)
    if master_rows is not None and not master_rows.empty:
        for alias in sorted(VENDOR_ALIASES, key=len, reverse=True):
            description = re.sub(rf"(?i)\b{re.escape(alias)}\b", " ", description)
    description = clean(re.sub(r"\s+", " ", description).strip(" -/,"))

    if master_rows is not None and not master_rows.empty:
        products = [clean(value) for value in master_rows["Product Name"] if clean(value)]
        if products:
            description = products[0]

    if not description:
        description = original

    fields_found = sum(bool(value) for value in [style, color, size, vendor])
    confidence = "High" if fields_found >= 4 else ("Medium" if fields_found >= 3 else "Low")
    return {
        "Product Name": description,
        "Style Number": style,
        "Garment Color": color,
        "Size": size,
        "Detected Vendor": vendor,
        "Parser Source": "Smart Custom Item",
        "Parse Confidence": confidence,
        "Needs Review": "" if confidence == "High" else "Custom item needs confirmation",
    }
