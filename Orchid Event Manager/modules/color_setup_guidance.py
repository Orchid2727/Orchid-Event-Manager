from __future__ import annotations

import re
from typing import Any

import pandas as pd

from modules.blank_garment_rules import is_blank_decoration
from modules.internal_services import is_in_house_decoration
from modules.purchase_rules import normalize_bool, row_rules


def clean_text(value: Any) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return re.sub(r"\s+", " ", str(value).strip())


def is_setup_required(value: Any) -> bool:
    """Return whether a row was added by a new Product Master candidate sync."""
    return clean_text(value).casefold() in {"yes", "y", "true", "1"}


def is_new_color_setup_row(row: pd.Series | dict[str, Any]) -> bool:
    """Return True only for an actual newly imported purchasing color.

    Product Master uses ``Setup Required`` for both a wholly new style and a
    newly observed color on an already configured style. A blank color row is
    only a product placeholder, though, and must not be presented as a new
    color decision.
    """
    return bool(
        is_setup_required(row.get("Setup Required", ""))
        and clean_text(row.get("Garment Color", ""))
    )


_GUIDANCE_COLUMNS = ("Garment Color", "Setup Required", "Decoration Type", "Decoration Color")


def _with_guidance_columns(rows: pd.DataFrame) -> pd.DataFrame:
    """Return a copy that is safe to inspect for color-setup guidance.

    Early Product Master files and interrupted migrations can legitimately be
    missing one or more recently added columns.  Guidance is informational; it
    must never prevent the editor from opening that catalog.
    """
    working = rows.copy()
    for column in _GUIDANCE_COLUMNS:
        if column not in working.columns:
            working[column] = ""
    return working


def _column_values(rows: pd.DataFrame, column: str) -> pd.Series:
    """Read a guidance column without assuming every legacy frame has it."""
    if column in rows.columns:
        return rows[column]
    return pd.Series("", index=rows.index, dtype=object)


def new_color_setup_guidance(rows: pd.DataFrame | None) -> dict[str, object]:
    """Summarize the color-specific work still required for one style.

    Existing color rows retain their saved Thread / Ink choice. This lets the
    editor call out only the purchasing colors added by the current import,
    instead of making a high-color-count style look as if every color requires
    another decision.
    """
    if rows is None or rows.empty:
        return {
            "new_count": 0,
            "existing_count": 0,
            "new_color_names": [],
            "thread_ink_color_names": [],
        }

    working = _with_guidance_columns(rows)
    actual_colors = working[working["Garment Color"].map(clean_text).ne("")]
    # A service/add-on style can legitimately have no purchasing-color rows.
    # Do not run DataFrame.apply() on that empty frame: with current pandas
    # versions its empty result can be interpreted as a column selector and
    # discard the guidance columns, which then raised KeyError("Decoration
    # Type") during the post-save progress refresh.
    if actual_colors.empty:
        return {
            "new_count": 0,
            "existing_count": 0,
            "new_color_names": [],
            "thread_ink_color_names": [],
        }
    new_rows = actual_colors[actual_colors.apply(is_new_color_setup_row, axis=1)]
    rules = row_rules(working.iloc[0])
    requires_decoration = normalize_bool(rules.get("Requires Decoration", ""), True)
    thread_ink_rows = new_rows.iloc[0:0]
    if requires_decoration:
        thread_ink_rows = new_rows[
            new_rows["Decoration Type"].map(
                lambda value: not is_blank_decoration(value) and not is_in_house_decoration(value)
            )
            & new_rows["Decoration Color"].map(clean_text).eq("")
        ]
    return {
        "new_count": len(new_rows),
        "existing_count": max(0, len(actual_colors) - len(new_rows)),
        "new_color_names": [clean_text(value) for value in _column_values(new_rows, "Garment Color")],
        "thread_ink_color_names": [clean_text(value) for value in _column_values(thread_ink_rows, "Garment Color")],
    }


def new_color_setup_message(rows: pd.DataFrame | None) -> tuple[str, bool]:
    """Build the concise Product Master guidance shown above purchasing colors."""
    guidance = new_color_setup_guidance(rows)
    new_count = int(guidance["new_count"])
    if not new_count:
        return "All saved purchasing colors and Thread / Ink settings are retained below.", False

    thread_ink_names = list(guidance["thread_ink_color_names"])
    new_color_names = list(guidance["new_color_names"])
    names = thread_ink_names or new_color_names
    preview = ", ".join(names[:3])
    if len(names) > 3:
        preview += f" +{len(names) - 3} more"
    verb = "NEEDS" if new_count == 1 else "NEED"
    action = "THREAD / INK" if thread_ink_names else "SETUP"
    text = f"{new_count} NEW COLOR{'S' if new_count != 1 else ''} {verb} {action}: {preview}"
    existing_count = int(guidance["existing_count"])
    if existing_count:
        text += f"  •  {existing_count} saved color setting{'s' if existing_count != 1 else ''} unchanged"
    return text, True


def prioritize_new_color_setup_rows(rows: pd.DataFrame) -> pd.DataFrame:
    """Show newly imported colors first while preserving alphabetical order otherwise."""
    if rows is None or rows.empty:
        return rows
    working = _with_guidance_columns(rows)
    pending_mask = working.apply(is_new_color_setup_row, axis=1)
    pending = working.loc[pending_mask]
    existing = working.loc[~pending_mask]
    sort_colors = lambda frame: frame.sort_values(
        by="Garment Color", key=lambda series: series.astype(str).str.casefold()
    )
    return pd.concat([sort_colors(pending), sort_colors(existing)])
