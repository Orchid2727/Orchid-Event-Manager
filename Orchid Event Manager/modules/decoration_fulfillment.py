from __future__ import annotations

from typing import Any

from modules.blank_garment_rules import is_blank_decoration
from modules.internal_services import is_in_house_decoration

STANDARD_ORCHID_WORKFLOW = "Standard Orchid Workflow"
ENTIRE_ORDER_OUTSOURCED = "Entire Order Outsourced"


def clean(value: Any) -> str:
    return " ".join(str(value or "").split())


def normalize_decoration_fulfillment(value: Any) -> str:
    text = clean(value).casefold()
    if text in {
        ENTIRE_ORDER_OUTSOURCED.casefold(),
        "all outsourced",
        "fully outsourced",
        "outsourced",
        "entire event outsourced",
    }:
        return ENTIRE_ORDER_OUTSOURCED
    return STANDARD_ORCHID_WORKFLOW


def is_entire_order_outsourced(value: Any) -> bool:
    return normalize_decoration_fulfillment(value) == ENTIRE_ORDER_OUTSOURCED


def is_outsourced_decoration(
    decoration_type: Any,
    fulfillment_mode: Any,
    *,
    do_not_outsource: bool = False,
) -> bool:
    """Return True when a physical garment belongs on the outside-decorator report.

    Orchid's standard workflow decorates embroidery in house and outsources screen
    printing. An Entire Order Outsourced packet sends every decorated garment to
    the outside decorator unless the line-level Do Not Outsource exception is set.
    Blank garments and explicitly in-house service routes are never outsourced.
    """
    if do_not_outsource:
        return False
    decoration = clean(decoration_type)
    if not decoration or is_blank_decoration(decoration) or is_in_house_decoration(decoration):
        return False
    if is_entire_order_outsourced(fulfillment_mode):
        return True
    return "screen" in decoration.casefold()
