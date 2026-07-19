from __future__ import annotations

import re

from modules.blank_garment_rules import is_blank_decoration
from modules.internal_services import is_in_house_decoration

LEFT_CHEST = "Left Chest"
LEFT_SLEEVE = "Left Sleeve"
LEFT_PANEL = "Left Panel"
RIGHT_PANEL = "Right Panel"
HAT = "Hat"
HAT_LEFT_SIDE = "Hat - Left Side"
HAT_RIGHT_SIDE = "Hat - Right Side"
HAT_BACK = "Hat - Back"
BEANIE = "Beanie"
BACK_YOKE = "Back Yoke"
OTHER_CUSTOM = "Other / Custom"
NOT_APPLICABLE_NO_DECORATION = "Not Applicable — No Decoration"
NOT_APPLICABLE_IN_HOUSE = "Not Applicable — In-House Service"

STANDARD_DECORATION_LOCATIONS = [
    LEFT_CHEST,
    LEFT_SLEEVE,
    LEFT_PANEL,
    RIGHT_PANEL,
    HAT,
    HAT_LEFT_SIDE,
    HAT_RIGHT_SIDE,
    HAT_BACK,
    BEANIE,
    BACK_YOKE,
]
DECORATION_LOCATIONS = STANDARD_DECORATION_LOCATIONS + [OTHER_CUSTOM]

_ALIASES = {
    "left chest": LEFT_CHEST,
    "chest": LEFT_CHEST,
    "lc": LEFT_CHEST,
    "left sleeve": LEFT_SLEEVE,
    "sleeve": LEFT_SLEEVE,
    "ls": LEFT_SLEEVE,
    "left panel": LEFT_PANEL,
    "right panel": RIGHT_PANEL,
    "hat": HAT,
    "cap": HAT,
    "standard hat": HAT,
    "hat left side": HAT_LEFT_SIDE,
    "hat - left side": HAT_LEFT_SIDE,
    "left side of hat": HAT_LEFT_SIDE,
    "hat right side": HAT_RIGHT_SIDE,
    "hat - right side": HAT_RIGHT_SIDE,
    "right side of hat": HAT_RIGHT_SIDE,
    "hat back": HAT_BACK,
    "hat - back": HAT_BACK,
    "back of hat": HAT_BACK,
    "back hat": HAT_BACK,
    "beanie": BEANIE,
    "beanies": BEANIE,
    "back yoke": BACK_YOKE,
    "upper back": BACK_YOKE,
    "other": OTHER_CUSTOM,
    "custom": OTHER_CUSTOM,
    "other / custom": OTHER_CUSTOM,
    NOT_APPLICABLE_NO_DECORATION.casefold(): "",
    NOT_APPLICABLE_IN_HOUSE.casefold(): "",
}


def _space(value: object) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip())


def _canonical_custom(value: str) -> str:
    """Normalize custom placement text so identical routes group together."""
    value = _space(value)
    if not value:
        return ""
    words = []
    for word in value.split(" "):
        if len(word) <= 3 and word.isupper():
            words.append(word)
        else:
            words.append(word.title())
    return " ".join(words)


def clean_location(value: object) -> str:
    text = _space(value)
    if not text:
        return ""
    standard = _ALIASES.get(text.casefold())
    if standard is not None:
        return standard
    return _canonical_custom(text)


def is_standard_location(value: object) -> bool:
    return clean_location(value) in STANDARD_DECORATION_LOCATIONS


def is_custom_location(value: object) -> bool:
    location = clean_location(value)
    return bool(location and location not in STANDARD_DECORATION_LOCATIONS and location != OTHER_CUSTOM)


def location_choice_and_custom(value: object, decoration_type: object = "") -> tuple[str, str]:
    """Return the dropdown choice and custom text for a stored location."""
    if is_blank_decoration(decoration_type):
        return NOT_APPLICABLE_NO_DECORATION, ""
    if is_in_house_decoration(decoration_type):
        return NOT_APPLICABLE_IN_HOUSE, ""
    location = clean_location(value)
    if location in STANDARD_DECORATION_LOCATIONS:
        return location, ""
    if location and location != OTHER_CUSTOM:
        return OTHER_CUSTOM, location
    return LEFT_CHEST, ""


def resolve_location_choice(choice: object, custom_value: object = "") -> str:
    """Resolve a UI dropdown choice into the exact structured routing value."""
    if _space(choice) in {NOT_APPLICABLE_NO_DECORATION, NOT_APPLICABLE_IN_HOUSE}:
        return ""
    selected = clean_location(choice)
    if selected == OTHER_CUSTOM:
        custom = clean_location(custom_value)
        return "" if custom == OTHER_CUSTOM else custom
    return selected


def default_decoration_location(
    product_name: object = "",
    category: object = "",
    decoration_type: object = "",
) -> str:
    """Return Orchid's default route for a decorated product.

    Standard caps/hats use Hat, beanies use Beanie, and other decorated apparel
    defaults to Left Chest. Manual saved locations still take precedence.
    """
    if is_blank_decoration(decoration_type) or is_in_house_decoration(decoration_type):
        return ""
    text = f"{_space(product_name)} {_space(category)}".casefold()
    if re.search(r"\bbeanie(?:s)?\b|\bknit\s+cap\b|\bwatch\s+cap\b", text):
        return BEANIE
    if re.search(r"\b(hat|hats|cap|caps|visor|headwear|boonie|booney)\b", text) or _space(category) == "Hats / Headwear":
        return HAT
    return LEFT_CHEST


def normalize_decoration_location(
    value: object,
    decoration_type: object = "",
    *,
    requires_decoration: bool = True,
    product_name: object = "",
    category: object = "",
) -> str:
    """Return Orchid's canonical placement for one purchase line."""
    if not requires_decoration or is_blank_decoration(decoration_type) or is_in_house_decoration(decoration_type):
        return ""
    location = clean_location(value)
    if not location or location == OTHER_CUSTOM:
        return default_decoration_location(product_name, category, decoration_type)
    return location


def location_sort_key(value: object) -> tuple[int, str]:
    location = clean_location(value)
    order = {name: index for index, name in enumerate(STANDARD_DECORATION_LOCATIONS)}
    return (order.get(location, len(order)), location.casefold())


def infer_decoration_location_from_note(value: object) -> str:
    text = _space(value).casefold()
    if not text:
        return ""
    if re.search(r"\bhat\b.*\bleft\b|\bleft\b.*\bhat\b", text):
        return HAT_LEFT_SIDE
    if re.search(r"\bhat\b.*\bright\b|\bright\b.*\bhat\b", text):
        return HAT_RIGHT_SIDE
    if re.search(r"\bhat\b.*\bback\b|\bback\b.*\bhat\b", text):
        return HAT_BACK
    if re.search(r"\bbeanie(?:s)?\b|\bknit\s+cap\b|\bwatch\s+cap\b", text):
        return BEANIE
    if re.search(r"\bleft[ -]?sleeve\b|\bon (?:the )?left sleeve\b", text):
        return LEFT_SLEEVE
    if re.search(r"\bleft[ -]?chest\b|\bon (?:the )?left chest\b", text):
        return LEFT_CHEST
    if re.search(r"\bleft[ -]?panel\b", text):
        return LEFT_PANEL
    if re.search(r"\bright[ -]?panel\b", text):
        return RIGHT_PANEL
    if re.search(r"\bback[ -]?yoke\b|\bupper[ -]?back\b", text):
        return BACK_YOKE
    if re.search(r"\bhat(?:s)?\b|\bcap(?:s)?\b", text):
        return HAT
    if re.search(r"\bsleeve\b", text):
        return LEFT_SLEEVE
    if re.search(r"\bchest\b", text):
        return LEFT_CHEST
    return ""
