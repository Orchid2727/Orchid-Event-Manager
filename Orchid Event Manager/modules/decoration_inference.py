from __future__ import annotations

from dataclasses import dataclass
import re
from typing import Iterable

import pandas as pd

from modules.blank_garment_rules import BLANK_DECORATION_LABEL, is_blank_decoration


EMBROIDERY = "Embroidery"
SCREEN_PRINT = "Screen Print"


DARK_COLOR_TERMS = {
    "black", "navy", "charcoal", "maroon", "burgundy", "forest", "green",
    "royal", "blue", "purple", "brown", "coyote", "loden", "olive", "red",
    "carbon", "graphite", "shadow", "deep", "dark", "bottle", "cardinal",
    "saddle", "moss", "grill", "team blue", "blue nights",
}
LIGHT_COLOR_TERMS = {
    "white", "ash", "silver", "grey", "gray", "khaki", "tan", "stone",
    "fossil", "sand", "yellow", "orange", "safety", "neon", "pastel",
    "athletic oxford", "light", "university blue", "sage", "mint", "fog",
}


@dataclass(frozen=True)
class ServiceTotals:
    embroidery: int = 0
    screen_print: int = 0


def clean(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value).strip())


def canonical_decoration(value: object) -> str:
    text = clean(value).casefold()
    if "embroider" in text:
        return EMBROIDERY
    if "screen" in text:
        return SCREEN_PRINT
    if is_blank_decoration(text):
        return BLANK_DECORATION_LABEL
    return clean(value)


def service_totals_from_raw(raw: pd.DataFrame) -> dict[str, ServiceTotals]:
    """Return paid decoration-piece totals by Shopify order number.

    Shopify stores embroidery and screen-print charges as separate line items. Those
    quantities are the strongest available signal for how many garments in an order
    receive each decoration method, so the review engine uses them before asking the
    user to classify every product row manually.
    """
    if raw.empty:
        return {}
    frame = raw.copy()
    for column in ["Name", "Lineitem name", "Lineitem quantity"]:
        if column not in frame.columns:
            frame[column] = ""
    frame["Name"] = frame["Name"].replace("", pd.NA).ffill().fillna("")
    frame["Lineitem name"] = frame["Lineitem name"].fillna("").map(clean)
    frame["Lineitem quantity"] = pd.to_numeric(
        frame["Lineitem quantity"], errors="coerce"
    ).fillna(0).astype(int)

    result: dict[str, ServiceTotals] = {}
    for order, group in frame.groupby("Name", sort=False, dropna=False):
        embroidery = 0
        screen = 0
        for _, row in group.iterrows():
            name = clean(row.get("Lineitem name", "")).casefold()
            qty = max(int(row.get("Lineitem quantity", 0) or 0), 0)
            if "embroider" in name:
                embroidery += qty
            elif "screen" in name and "print" in name:
                screen += qty
        result[clean(order)] = ServiceTotals(embroidery=embroidery, screen_print=screen)
    return result


def _is_dark_color(color: object) -> bool:
    text = clean(color).casefold()
    if not text:
        return False
    first = re.split(r"[/|-]", text, maxsplit=1)[0].strip()
    dark_prefixes = {
        "black", "navy", "charcoal", "maroon", "burgundy", "forest",
        "royal", "purple", "brown", "coyote", "loden", "olive",
        "carbon", "graphite", "shadow", "deep", "dark", "bottle",
        "cardinal", "saddle", "moss", "grill",
    }
    if any(term in text for term in LIGHT_COLOR_TERMS):
        # Safety Green, Neon Orange, Light Blue, and similar colors need dark
        # decoration. A clearly dark first component such as Black/Silver still
        # counts as a dark garment.
        return any(first.startswith(term) for term in dark_prefixes)
    return any(term in text for term in DARK_COLOR_TERMS)


def infer_decoration_color(decoration_type: object, garment_color: object) -> str:
    """Apply Orchid's normal one-color contrast rule when no saved mapping exists.

    The result is intentionally conservative and remains editable in Purchase Review.
    Screen print uses white on dark garments and black on light garments. Embroidery
    uses silver on dark garments and black on light garments.
    """
    decoration = canonical_decoration(decoration_type)
    color = clean(garment_color)
    if not decoration or decoration == BLANK_DECORATION_LABEL or not color:
        return ""
    dark = _is_dark_color(color)
    if decoration == SCREEN_PRINT:
        return "White" if dark else "Black"
    if decoration == EMBROIDERY:
        return "Silver" if dark else "Black"
    return ""



def _is_lower_body_row(row: pd.Series) -> bool:
    category = clean(row.get("Product Category", "")).casefold()
    description = clean(row.get("Product Name", row.get("Description", ""))).casefold()
    size = clean(row.get("Size", ""))
    if "pants / jeans / shorts" in category:
        return True
    if re.search(r"\b(pant|pants|jean|jeans|trouser|trousers|shorts|slacks)\b", description):
        return True
    # Waist-by-inseam sizes are a reliable signal for manual pants such as
    # "1104 Ascent 36x32" even when the description does not say pant.
    upper_body = re.search(r"\b(shirt|tee|polo|jacket|coat|vest|cap|hat|hoodie|fleece)\b", description)
    return bool(re.fullmatch(r"\d{2}x\d{2}", size, re.I) and not upper_body)

def _preference(row: pd.Series, assignment: str) -> int:
    category = clean(row.get("Product Category", "")).casefold()
    description = clean(row.get("Product Name", row.get("Description", ""))).casefold()
    text = f"{category} {description}"

    # Strong garment-shape preferences. These are scoring hints only; the paid
    # service quantities must still balance exactly before a rule is applied.
    if re.search(r"\b(pant|pants|jean|jeans|trouser|trousers|shorts|slacks)\b", text):
        return 20 if assignment == BLANK_DECORATION_LABEL else -20
    if re.search(r"\b(tee|t-shirt|t shirt|hooded tee|long sleeve tee|short sleeve tee)\b", text):
        return 12 if assignment == SCREEN_PRINT else (3 if assignment == EMBROIDERY else -8)
    if re.search(r"\b(cap|hat|beanie|boonie|booney|visor|headwear)\b", text):
        return 12 if assignment == EMBROIDERY else (-2 if assignment == SCREEN_PRINT else -8)
    if re.search(r"\b(polo|woven|button[- ]?down|dress shirt|work shirt|blouse)\b", text):
        return 10 if assignment == EMBROIDERY else (2 if assignment == SCREEN_PRINT else -7)
    if re.search(r"\b(jacket|coat|parka|vest|fleece|soft shell|softshell|pullover|quarter[- ]?zip|hoodie|sweatshirt)\b", text):
        return 9 if assignment == EMBROIDERY else (2 if assignment == SCREEN_PRINT else -6)
    if re.search(r"\b(bag|backpack|duffel|tote|messenger)\b", text):
        return 9 if assignment == EMBROIDERY else (1 if assignment == SCREEN_PRINT else -5)
    if "safety workwear" in category:
        return 8 if assignment == SCREEN_PRINT else (4 if assignment == EMBROIDERY else -6)
    if "t-shirts" in category:
        return 8 if assignment == SCREEN_PRINT else (3 if assignment == EMBROIDERY else -5)
    if "hats" in category:
        return 8 if assignment == EMBROIDERY else (0 if assignment == SCREEN_PRINT else -5)
    if "outerwear" in category or "polos" in category:
        return 7 if assignment == EMBROIDERY else (2 if assignment == SCREEN_PRINT else -4)
    if "pants" in category:
        return 12 if assignment == BLANK_DECORATION_LABEL else -12
    if "other apparel" in category:
        return 4 if assignment in {EMBROIDERY, SCREEN_PRINT} else -2
    return 1 if assignment in {EMBROIDERY, SCREEN_PRINT} else 0


def _group_key(row: pd.Series) -> str:
    style = re.sub(r"\s+", "", clean(row.get("Style Number", row.get("Product #", "")))).upper()
    description = clean(row.get("Product Name", row.get("Description", ""))).casefold()
    return f"{style}|{description}" if style else f"description:{description}"


def _best_exact_assignment(rows: pd.DataFrame, embroidery_target: int, screen_target: int) -> dict[int, str] | None:
    """Choose a confident whole-product allocation using dynamic programming.

    All colors and sizes of the same style/description are assigned together. This
    prevents the engine from arbitrarily embroidering one color while marking a
    second color of the same product blank merely to make quantities balance.
    """
    if rows.empty:
        return {} if embroidery_target == 0 and screen_target == 0 else None
    if embroidery_target < 0 or screen_target < 0:
        return None

    grouped: list[tuple[list[int], pd.Series, int]] = []
    keyed = rows.copy()
    keyed["_allocation_group"] = keyed.apply(_group_key, axis=1)
    for _, group in keyed.groupby("_allocation_group", sort=False, dropna=False):
        indices = [int(value) for value in group.index]
        qty = int(pd.to_numeric(group["Quantity"], errors="coerce").fillna(0).sum())
        grouped.append((indices, group.iloc[0], max(qty, 0)))

    # Keep the two best distinct assignments for each quantity state so an
    # ambiguous tie can be left for review rather than silently guessed.
    states: dict[tuple[int, int], list[tuple[int, dict[int, str]]]] = {(0, 0): [(0, {})]}
    for group_indices, sample_row, qty in grouped:
        next_states: dict[tuple[int, int], list[tuple[int, dict[int, str]]]] = {}
        for (used_e, used_s), candidates in states.items():
            for score, assignment in candidates:
                for choice in (EMBROIDERY, SCREEN_PRINT, BLANK_DECORATION_LABEL):
                    add_e = qty if choice == EMBROIDERY else 0
                    add_s = qty if choice == SCREEN_PRINT else 0
                    new_e, new_s = used_e + add_e, used_s + add_s
                    if new_e > embroidery_target or new_s > screen_target:
                        continue
                    new_score = score + _preference(sample_row, choice)
                    new_assignment = dict(assignment)
                    for index in group_indices:
                        new_assignment[index] = choice
                    key = (new_e, new_s)
                    bucket = next_states.setdefault(key, [])
                    signature = tuple(sorted(new_assignment.items()))
                    if any(tuple(sorted(existing.items())) == signature for _, existing in bucket):
                        continue
                    bucket.append((new_score, new_assignment))
                    bucket.sort(key=lambda value: value[0], reverse=True)
                    del bucket[2:]
        states = next_states
        if not states:
            return None

    finalists = states.get((embroidery_target, screen_target), [])
    if not finalists:
        return None
    finalists.sort(key=lambda value: value[0], reverse=True)
    if len(finalists) > 1 and finalists[0][0] - finalists[1][0] < 4:
        return None
    return finalists[0][1]


def infer_order_decorations(
    frame: pd.DataFrame,
    service_totals: dict[str, ServiceTotals],
) -> pd.DataFrame:
    """Fill missing decoration methods from the order's paid service quantities.

    Existing saved Product Master mappings remain authoritative. The function only
    fills blank decoration methods, and only when an exact, whole-line allocation is
    possible. Orders without any decoration service are treated as blank garments.
    """
    if frame.empty:
        return frame.copy()
    result = frame.copy()
    for column in [
        "Order Number", "Decoration Type", "Decoration Color", "Garment Color",
        "Quantity", "Product Category", "Product Name", "Decoration Source",
    ]:
        if column not in result.columns:
            result[column] = "" if column != "Quantity" else 0
    result["Quantity"] = pd.to_numeric(result["Quantity"], errors="coerce").fillna(0).astype(int)
    result["Decoration Type"] = result["Decoration Type"].map(canonical_decoration)
    result["Decoration Source"] = result["Decoration Source"].fillna("").map(clean)

    for order, indices in result.groupby("Order Number", sort=False).groups.items():
        order_key = clean(order)
        totals = service_totals.get(order_key, ServiceTotals())
        group = result.loc[list(indices)].copy()

        # Manual pants frequently arrive with only a style name and a waist x
        # inseam size. Classify those as blank before reconciling paid decoration
        # quantities, so the engine does not ask the user to decorate pants.
        for index, row in group.iterrows():
            if not clean(row.get("Decoration Type", "")) and _is_lower_body_row(row):
                result.at[index, "Decoration Type"] = BLANK_DECORATION_LABEL
                result.at[index, "Decoration Color"] = ""
                result.at[index, "Decoration Source"] = "Blank garment inferred from product and size"
        group = result.loc[list(indices)].copy()
        unknown_mask = group["Decoration Type"].map(clean).eq("")
        if totals.embroidery == 0 and totals.screen_print == 0:
            for index in group.index[unknown_mask]:
                result.at[index, "Decoration Type"] = BLANK_DECORATION_LABEL
                result.at[index, "Decoration Color"] = ""
                result.at[index, "Decoration Source"] = "No decoration service on Shopify order"
            continue

        known_e = int(group.loc[group["Decoration Type"].eq(EMBROIDERY), "Quantity"].sum())
        known_s = int(group.loc[group["Decoration Type"].eq(SCREEN_PRINT), "Quantity"].sum())
        remaining_e = totals.embroidery - known_e
        remaining_s = totals.screen_print - known_s
        unknown = group.loc[unknown_mask].copy()
        assignment = _best_exact_assignment(unknown, remaining_e, remaining_s)
        if assignment is None:
            continue

        for index, decoration_type in assignment.items():
            result.at[index, "Decoration Type"] = decoration_type
            if decoration_type == BLANK_DECORATION_LABEL:
                result.at[index, "Decoration Color"] = ""
            elif not clean(result.at[index, "Decoration Color"]):
                result.at[index, "Decoration Color"] = infer_decoration_color(
                    decoration_type, result.at[index, "Garment Color"]
                )
            result.at[index, "Decoration Source"] = "Inferred from Shopify decoration quantities"

    # Fill contrast colors for saved or inferred decoration methods only when no
    # explicit Product Master color exists.
    for index, row in result.iterrows():
        decoration = canonical_decoration(row.get("Decoration Type", ""))
        if decoration in {EMBROIDERY, SCREEN_PRINT} and not clean(row.get("Decoration Color", "")):
            inferred = infer_decoration_color(decoration, row.get("Garment Color", ""))
            if inferred:
                result.at[index, "Decoration Color"] = inferred
                if not clean(result.at[index, "Decoration Source"]):
                    result.at[index, "Decoration Source"] = "Contrast color inferred from garment color"
    return result


def unresolved_decision_key(row: pd.Series) -> str:
    """Return a stable key so repeated size/color rows count as one decision."""
    style = re.sub(r"\s+", "", clean(row.get("Product #", row.get("Style Number", "")))).upper()
    description = clean(row.get("Description", row.get("Product Name", ""))).casefold()
    reason = clean(row.get("Review Reason", "")).casefold()
    # Missing garment color/size is usually specific to a Shopify line, while vendor
    # and decoration setup are permanent style-level decisions.
    if "missing garment color" in reason or "missing size" in reason or "customer decision" in reason:
        order = clean(row.get("Order Number", ""))
        line_id = clean(row.get("Line ID", ""))
        return f"order:{order}|line:{line_id}"
    product_key = style or description or clean(row.get("Line ID", ""))
    return f"product:{product_key}|{reason}"


def count_unique_decisions(frame: pd.DataFrame) -> int:
    if frame.empty:
        return 0
    return len({unresolved_decision_key(row) for _, row in frame.iterrows()})
