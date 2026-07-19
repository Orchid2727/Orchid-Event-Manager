from __future__ import annotations

import re


def clean_note(value: object) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip())


_MONEY = r"\$?\s*-?\s*\d[\d,]*(?:\.\d{1,2})?"
_FINANCIAL_PATTERNS = (
    rf"\b(?:employee\s+)?allowance\b\s*(?:is|of|:|=|-)?\s*{_MONEY}",
    rf"\b(?:employee\s+)?budget\b\s*(?:is|of|:|=|-)?\s*{_MONEY}",
    rf"\bdiscount(?:ed)?\b\s*(?:is|of|:|=|-)?\s*{_MONEY}",
    rf"\b(?:city|company|employee|approved|order)\s+total\b\s*(?:is|of|:|=|-)?\s*{_MONEY}",
    rf"\bbalance\b\s*(?:is|of|:|=|-)?\s*{_MONEY}",
)

# Notes that change the purchasing decision or routing should remain in the work queue.
DECISION_TERMS = (
    "instead of", "change to", "replace with", "substitute", "if unavailable",
    "do not outsource", "don't outsource", "ship to orchid", "send to orchid",
    "if out of stock", "do not order", "do not purchase", "hold order",
    "rush", "ship separately", "drop ship", "dropship", "call before",
    "confirm before", "different vendor", "special order", "backorder",
    "cancel", "remove item", "exchange", "return",
)

# Decoration wording is always a mandatory order-specific decision. Standalone
# EMB is intentionally bounded so words such as "member" do not trigger review.
_DECORATION_TRIGGER = re.compile(
    r"(?i)(?:\bembroider(?:y|ed|ing|s)?\b|(?<![A-Za-z0-9])EMB(?![A-Za-z0-9]))"
)

# A note can materially change decoration even when it never says EMB or
# embroidery. Placement, logo, screen-print, personalization, and alteration
# wording must stop in Purchase Review before it is printed on a vendor PO.
_DECORATION_INSTRUCTION_TRIGGER = re.compile(
    r"(?i)(?:"
    r"\bscreen[- ]?print(?:ed|ing)?\b|"
    r"\b(?:logo|logos|seal|patch|monogram|initials?)\b|"
    r"\b(?:left|right)\s+chest\b|"
    r"\b(?:on|at)\s+(?:the\s+)?(?:front|back|sleeve|sleeves)\b|"
    r"\b(?:front|back|sleeve|sleeves)\s+(?:logo|placement|location|only)\b|"
    r"\b(?:ink|thread)\s*(?:color|colour)?\b|"
    r"\bheat\s*press\b|\bvinyl\b|"
    r"\bpersonaliz(?:e|ed|es|ing|ation)\b|"
    r"\bname\s+on\b|"
    r"\bhem(?:med|ming)?\b|\binseam\b|"
    r"\balter(?:ed|ing|ation|ations)?\b|"
    r"\bdo\s+not\s+decorate\b|\bno\s+decoration\b|\bblank\s+garment\b"
    r")"
)

# General wording such as "All logos on sleeve" applies to the whole order,
# even if the same note also contains a special hat-only instruction.
_GLOBAL_DECORATION_SCOPE = re.compile(
    r"(?i)(?:\ball\s+(?:logos?|decoration|embroidery)\b|"
    r"\b(?:logos?|decoration)\s+on\s+(?:the\s+)?(?:front|back|sleeve|sleeves)\b)"
)

# These instructions matter to purchasing or decoration. Decoration instructions
# are blocking; the remaining terms are visible context unless a decision term is
# also present.
INSTRUCTION_TERMS = (
    "emb ", "embroider", "embroidery", "screen print", "screen-print", "screenprint",
    "ink", "thread", "logo", "left chest", "right chest", "sleeve",
    "name on", "initials", "monogram", "patch", "heat press", "vinyl",
    "do not decorate", "no decoration", "no emb", "blank garment",
    "do not outsource", "don't outsource", "ship to orchid", "send to orchid",
    "personalize", "personalization", "hat", "hats",
)

# Garment families used to route an order-level decoration note to the intended
# line when wording such as "No EMB on Rain Jacket" or "Embroider ... on hats"
# is present.  If no family is named, the note remains mandatory on every line so
# the instruction cannot be missed.
_TARGET_FAMILIES: dict[str, tuple[str, ...]] = {
    "rain_jacket": (
        "rain jacket", "rain jackets", "waterproof jacket", "waterproof jackets",
        "raincoat", "raincoats", "shell jacket", "shell jackets",
    ),
    "jacket": ("jacket", "jackets", "coat", "coats", "parka", "parkas", "outerwear"),
    "hat": ("hat", "hats", "cap", "caps", "beanie", "beanies", "boonie", "boonies", "headwear", "visor", "visors"),
    "shirt": ("shirt", "shirts", "tee", "tees", "t-shirt", "t-shirts", "t shirt", "t shirts"),
    "polo": ("polo", "polos"),
    "pant": ("pant", "pants", "jean", "jeans", "trouser", "trousers", "bib", "bibs", "overall", "overalls"),
    "short": ("short", "shorts"),
    "vest": ("vest", "vests"),
    "bag": ("bag", "bags", "backpack", "backpacks", "duffel", "duffels", "tote", "totes"),
    "hoodie": ("hoodie", "hoodies", "sweatshirt", "sweatshirts", "fleece"),
}


def extract_purchase_instructions(value: object) -> str:
    """Remove financial clauses while preserving purchasing/decoration instructions.

    Example:
        Allowance $700 Discount $2.04 EMB Mr. Lusk on hat
    becomes:
        EMB Mr. Lusk on hat

    The original Shopify note remains available in hidden workbook data.
    """
    note = clean_note(value)
    if not note:
        return ""

    cleaned = note
    for pattern in _FINANCIAL_PATTERNS:
        cleaned = re.sub(pattern, " ", cleaned, flags=re.IGNORECASE)

    # Remove separators left behind by stripped financial clauses without
    # damaging useful punctuation inside the remaining instruction.
    cleaned = re.sub(r"(?:^|\s)[|;,:-]+(?=\s|$)", " ", cleaned)
    cleaned = re.sub(r"\s+-\s+", " ", cleaned)
    cleaned = re.sub(r"\s{2,}", " ", cleaned).strip(" \t\r\n-|;,:.")

    # A note that consisted only of financial wording may leave bare labels.
    if cleaned.casefold() in {
        "allowance", "discount", "discounted", "city total", "company total",
        "employee total", "approved total", "budget", "balance",
    }:
        return ""
    return cleaned


def combine_purchase_instructions(order_note: object, line_note: object = "") -> str:
    """Return readable order/line instructions without duplicating identical text."""
    order_text = extract_purchase_instructions(order_note)
    line_text = extract_purchase_instructions(line_note)
    if order_text and line_text and order_text.casefold() == line_text.casefold():
        return order_text
    parts: list[str] = []
    if order_text:
        parts.append(f"Order note: {order_text}" if line_text else order_text)
    if line_text:
        parts.append(f"Line note: {line_text}" if order_text else line_text)
    return " | ".join(parts)


def decoration_note_requires_review(value: object) -> bool:
    note = extract_purchase_instructions(value)
    return bool(note and (_DECORATION_TRIGGER.search(note) or _DECORATION_INSTRUCTION_TRIGGER.search(note)))


def decision_note_requires_review(value: object) -> bool:
    note = extract_purchase_instructions(value).casefold()
    return bool(note and any(term in note for term in DECISION_TERMS))




def decoration_target_styles(value: object) -> set[str]:
    """Extract explicit product/style references from an instruction note.

    Examples: ``Name on NEA201 Moore`` and ``Logo on style 112``. The
    returned keys are normalized so they can be matched against the current
    order lines without showing the instruction on unrelated products.
    """
    note = extract_purchase_instructions(value)
    if not note:
        return set()
    candidates: set[str] = set()
    patterns = (
        r"(?i)\b(?:style|item|product|sku)\s*#?\s*([A-Z0-9][A-Z0-9-]{1,19})\b",
        r"(?i)\b(?:name|logo|embroider|embroidery|emb|patch)\s+on\s+([A-Z0-9][A-Z0-9-]{1,19})\b",
        r"(?i)\bon\s+([A-Z]{1,8}\d[A-Z0-9-]{1,18}|\d{3,6})\b",
    )
    for pattern in patterns:
        for raw in re.findall(pattern, note):
            key = re.sub(r"\s+", "", str(raw)).upper().strip("-.,:;")
            if not key:
                continue
            if (re.search(r"[A-Z]", key) and re.search(r"\d", key)) or re.fullmatch(r"\d{3,6}", key):
                candidates.add(key)
    return candidates


def decoration_note_is_global(value: object) -> bool:
    note = extract_purchase_instructions(value)
    return bool(note and _GLOBAL_DECORATION_SCOPE.search(note))

def decoration_target_families(value: object) -> set[str]:
    note = extract_purchase_instructions(value).casefold()
    if not note:
        return set()
    if _GLOBAL_DECORATION_SCOPE.search(note):
        return set()
    families: set[str] = set()
    # Prefer the more specific rain-jacket family over the broad jacket family.
    if any(term in note for term in _TARGET_FAMILIES["rain_jacket"]):
        families.add("rain_jacket")
    for family, terms in _TARGET_FAMILIES.items():
        if family == "rain_jacket":
            continue
        if any(re.search(rf"(?<![a-z0-9]){re.escape(term)}(?![a-z0-9])", note) for term in terms):
            families.add(family)
    if "rain_jacket" in families:
        families.discard("jacket")
    return families


def decoration_note_targets_line(value: object, line_text: object) -> bool:
    """Return True when an order-level decoration note applies to this line."""
    if not decoration_note_requires_review(value):
        return False
    text_raw = clean_note(line_text)
    text = text_raw.casefold()
    if not text:
        return False

    style_targets = decoration_target_styles(value)
    if style_targets:
        line_tokens = {
            re.sub(r"\s+", "", token).upper()
            for token in re.findall(r"(?i)\b[A-Z0-9][A-Z0-9-]{1,29}\b", text_raw)
        }
        return bool(style_targets & line_tokens)

    families = decoration_target_families(value)
    if not families:
        return True
    for family in families:
        if any(re.search(rf"(?<![a-z0-9]){re.escape(term)}(?![a-z0-9])", text) for term in _TARGET_FAMILIES[family]):
            return True
    return False


def note_requires_review(value: object) -> bool:
    note = extract_purchase_instructions(value)
    if not note:
        return False
    return decoration_note_requires_review(note) or decision_note_requires_review(note)


def note_requires_review_for_line(
    order_note: object,
    line_note: object = "",
    line_text: object = "",
) -> bool:
    """Evaluate order and line notes for one purchase line.

    Line-level notes always apply to their own line.  Non-decoration decision terms
    in an order note retain the historic all-line behavior.  Decoration notes are
    routed to a named garment family when possible.
    """
    if note_requires_review(line_note):
        return True
    if decision_note_requires_review(order_note):
        return True
    if not decoration_note_requires_review(order_note):
        return False
    # Ambiguous order-level decoration notes are represented by one consolidated
    # decision in review_workbook._mandatory_note_review_flags. Do not recreate
    # the same prompt on every line during revalidation.
    if not decoration_target_styles(order_note) and not decoration_target_families(order_note) and not decoration_note_is_global(order_note):
        return False
    return decoration_note_targets_line(order_note, line_text)


def note_action_label(value: object) -> str:
    note = extract_purchase_instructions(value)
    if not note:
        return ""
    if decoration_note_requires_review(note):
        return "Instruction Decision Required"
    if decision_note_requires_review(note):
        return "Decision Required"
    lowered = note.casefold()
    if any(term in lowered for term in INSTRUCTION_TERMS):
        return "Purchase Instruction"
    return "Informational"


def decoration_instruction_recommendation(value: object, current_decoration_type: object = "") -> tuple[str, str]:
    """Recommend the safest explicit decision for a decoration instruction.

    The recommendation is advisory only. It distinguishes a note that changes
    placement/personalization from one that changes the actual decoration process.
    """
    note = extract_purchase_instructions(value)
    lowered = note.casefold()
    current = clean_note(current_decoration_type).casefold()
    if not note:
        return "", ""

    if re.search(r"\b(?:no|do not|don't)\s+(?:emb|embroider|embroidery|decorate)\b|\bno decoration\b|\bblank garment\b", lowered):
        return "No Decoration", "The note explicitly removes decoration from this item."

    mentions_screen = bool(re.search(r"\bscreen[- ]?print(?:ed|ing)?\b", lowered))
    mentions_embroidery = bool(_DECORATION_TRIGGER.search(note))

    if mentions_screen:
        if "screen" in current:
            return "Follow Note as Written", "Screen printing is already the saved decoration type; the note adds an instruction rather than changing the process."
        return "Screen Print", "The note explicitly sets the decoration process to screen printing."

    if mentions_embroidery:
        if "embroider" in current:
            return "Follow Note as Written", "Embroidery is already the saved decoration type; follow the note for placement or personalization."
        return "Embroidery", "The note explicitly sets the decoration process to embroidery."

    if decoration_note_requires_review(note):
        return "Follow Note as Written", "The note changes placement, logo wording, personalization, or another instruction while keeping the saved decoration type."

    return "", ""
