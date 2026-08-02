from __future__ import annotations

UNIFORM_SIZING_EVENT = "Uniform Sizing Event"
GENERAL_SALES_PERIOD = "General Sales Period"
VALID_REPORT_MODES = (UNIFORM_SIZING_EVENT, GENERAL_SALES_PERIOD)


def normalize_report_mode(value: object) -> str:
    text = " ".join(str(value or "").strip().split())
    for mode in VALID_REPORT_MODES:
        if text.casefold() == mode.casefold():
            return mode
    return GENERAL_SALES_PERIOD
