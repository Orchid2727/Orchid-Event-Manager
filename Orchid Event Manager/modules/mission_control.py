from __future__ import annotations

from datetime import datetime
import json
from pathlib import Path
from typing import Any
import re

from modules.blank_garment_rules import BLANK_DECORATION_LABEL, is_blank_decoration
from modules.decoration_locations import OTHER_CUSTOM, normalize_decoration_location
from modules.report_modes import GENERAL_SALES_PERIOD, UNIFORM_SIZING_EVENT
from modules.purchase_order_generator import normalize_size, split_embedded_size
from modules.purchase_rules import normalize_bool, row_rules
from modules.product_resolver import load_extended_master, resolve_product, normalize_style
from modules.employee_totals import employee_totals_summary, load_employee_totals
from modules.internal_services import (
    is_in_house_decoration, is_in_house_service_product, is_internal_service_style,
)
from modules.xlsx_reader import (
    load_event_name,
    load_report_mode,
    load_review_lines,
    read_table,
)


def clean(value: object) -> str:
    return " ".join(str(value or "").split())


def canonical_report_type(value: object) -> str:
    text = clean(value).casefold()
    if "screen" in text:
        return "screen printing"
    if "embroider" in text:
        return "embroidery"
    if "blank" in text or text in {"none", "no decoration"}:
        return "blank garments"
    if "combined" in text or not text:
        return "combined vendor order"
    return text


def route_key(vendor: object, report_type: object) -> str:
    return f"{clean(vendor).casefold()}|||{canonical_report_type(report_type)}"


def _state_file(data_root: Path) -> Path:
    return Path(data_root) / "mission_control_state.json"


def _read_state(data_root: Path) -> dict[str, Any]:
    path = _state_file(data_root)
    if not path.exists():
        return {"workbooks": {}}
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
        if not isinstance(payload, dict):
            return {"workbooks": {}}
        payload.setdefault("workbooks", {})
        return payload
    except Exception:
        return {"workbooks": {}}


def load_po_overrides(data_root: Path, workbook_path: Path) -> dict[tuple[str, str], str]:
    payload = _read_state(data_root)
    workbook_key = str(Path(workbook_path).resolve())
    stored = payload.get("workbooks", {}).get(workbook_key, {}).get("po_numbers", {})
    result: dict[tuple[str, str], str] = {}
    if isinstance(stored, dict):
        for key, value in stored.items():
            if "|||" not in str(key):
                continue
            vendor, report = str(key).split("|||", 1)
            number = clean(value)
            if number:
                result[(vendor, report)] = number
    return result


def save_po_overrides(
    data_root: Path,
    workbook_path: Path,
    po_numbers: dict[tuple[str, str], str],
) -> Path:
    root = Path(data_root)
    root.mkdir(parents=True, exist_ok=True)
    state = _read_state(root)
    workbook_key = str(Path(workbook_path).resolve())
    encoded = {
        f"{clean(vendor).casefold()}|||{canonical_report_type(report)}": clean(value)
        for (vendor, report), value in po_numbers.items()
        if clean(value)
    }
    state.setdefault("workbooks", {})[workbook_key] = {
        "po_numbers": encoded,
        "updated": datetime.now().isoformat(timespec="seconds"),
    }
    path = _state_file(root)
    path.write_text(json.dumps(state, indent=2, sort_keys=True), encoding="utf-8")
    return path


def _record_is_service_only(record: dict[str, object]) -> bool:
    product_number = clean(record.get("Product #", record.get("Style Number", "")))
    description = clean(record.get("Description", record.get("Product Name", "")))
    garment_color = clean(record.get("Garment Color", record.get("Color", "")))
    decoration_type = clean(record.get("Decoration Type", ""))
    return bool(
        is_internal_service_style(product_number)
        or is_in_house_service_product(
            description,
            record.get("Original Shopify Line", record.get("Original Line Item", "")),
            decoration_type,
            style_number=product_number,
            garment_color=garment_color,
        )
    )



def _split_issue_colors(value: object) -> list[str]:
    text = clean(value)
    if not text:
        return []
    return list(dict.fromkeys(
        clean(part) for part in re.split(r"\s*[,;|]\s*", text) if clean(part)
    ))


def _product_master_issue_is_resolved(
    issue: dict[str, object],
    master_path: Path | None,
    master=None,
) -> bool:
    """Return True when a workbook's permanent warning is stale.

    Purchase Review workbooks are snapshots. Product Master may be corrected later,
    so reopening an event must verify the exact ordered color against the live master
    instead of blocking on old warning text or an unfinished unrelated color row.
    """
    if master_path is None:
        return False
    candidate_path = Path(master_path)
    if not candidate_path.exists():
        return False
    try:
        frame = master if master is not None else load_extended_master(candidate_path)
    except Exception:
        return False
    style = clean(issue.get("product", issue.get("Product #", "")))
    description = clean(issue.get("description", issue.get("Description", "")))
    colors = _split_issue_colors(
        issue.get("garment_colors", issue.get("Garment Color(s)", issue.get("Garment Color", "")))
    ) or [""]
    reason = clean(issue.get("reason", issue.get("Missing Information", ""))).casefold()

    # A true unknown style remains permanent setup work.
    if style:
        style_rows = frame[frame["Style Number"].map(normalize_style).eq(normalize_style(style))]
    else:
        style_rows = frame.iloc[0:0]
    if style and style_rows.empty:
        return False

    for color in colors:
        result = resolve_product(style, description, color, frame, master_prepared=True)
        if not result.matched:
            return False
        issue_text = clean(result.issue).casefold()
        if "setup required" in issue_text:
            return False
        if not result.vendor:
            return False
        if result.requires_decoration:
            if not result.decoration_type:
                return False
            if (
                not is_blank_decoration(result.decoration_type)
                and not is_in_house_decoration(result.decoration_type)
                and not result.decoration_color
            ):
                return False
            if normalize_decoration_location(
                result.decoration_location,
                result.decoration_type,
                requires_decoration=True,
                product_name=result.product_name,
                category=result.product_category,
            ) == OTHER_CUSTOM:
                return False
        if result.requires_color and color and not result.garment_color:
            return False

    # Only suppress a permanent queue item when the live master now satisfies it.
    return bool(reason)


def _filter_live_product_master_issues(
    issues: list[dict[str, str]],
    master_path: Path | None,
) -> tuple[list[dict[str, str]], set[str]]:
    if master_path is None:
        return issues, set()
    candidate_path = Path(master_path)
    if not candidate_path.exists():
        return issues, set()
    try:
        master = load_extended_master(candidate_path)
    except Exception:
        return issues, set()
    kept: list[dict[str, str]] = []
    resolved_line_ids: set[str] = set()
    for issue in issues:
        is_master = clean(issue.get("fix_in", issue.get("source", ""))).casefold() == "product master"
        if is_master and _product_master_issue_is_resolved(issue, candidate_path, master=master):
            for value in clean(issue.get("affected_line_ids", issue.get("line_id", ""))).split(","):
                if clean(value):
                    resolved_line_ids.add(clean(value))
            continue
        kept.append(issue)
    return kept, resolved_line_ids


def _effective_review_blockers(record: dict[str, object]) -> list[str]:
    """Return live blockers without trusting an older row's stale reason text."""
    candidate = dict(record)
    candidate["Review Status"] = "Ready"
    try:
        blockers = _line_block_reasons(candidate)
    except Exception:
        return [clean(record.get("Review Reason", "")) or "Needs review"]
    return [
        "Customer decision required" if reason == "Customer instruction decision required" else reason
        for reason in blockers
    ]


def _record_is_effectively_complete(record: dict[str, object]) -> bool:
    """Recognize a saved decision even when an older build missed the status flag."""
    return not _effective_review_blockers(record)


def _review_queue(path: Path, review_rows: list[dict[str, object]] | None = None) -> tuple[list[dict[str, str]], set[str]]:
    """Read permanent and order-specific queues as separate decisions.

    The same order line can need both a reusable Product Master correction and an
    event-only correction. Distinct keys preserve both, while the app still blocks
    Purchase Review until the permanent setup is finished.
    """
    issues_by_key: dict[str, dict[str, str]] = {}
    unresolved_ids: set[str] = set()

    permanent_rows = []
    for sheet_name in ("Product Master", "Product Master Needed", "Purchasing Rules Needed"):
        try:
            permanent_rows = read_table(path, sheet_name, {"Product #", "Missing Information"})
            break
        except Exception:
            continue
    for record in permanent_rows:
        decision_key = clean(record.get("Decision Key", "")) or (
            "catalog:" + clean(record.get("Product #", "")) + "|" + clean(record.get("Missing Information", ""))
        )
        line_ids = [value.strip() for value in clean(record.get("Affected Line IDs", "")).split(",") if value.strip()]
        unresolved_ids.update(line_ids)
        issues_by_key[f"master:{decision_key}"] = {
            "line_id": line_ids[0] if line_ids else "",
            "affected_line_ids": ",".join(line_ids),
            "order": clean(record.get("Affected Orders", "")),
            "employee": "",
            "product": clean(record.get("Product #", "")),
            "description": clean(record.get("Description", "")),
            "vendor": clean(record.get("Current Vendor", record.get("Suggested Vendor", ""))),
            "garment_colors": clean(record.get("Garment Color(s)", record.get("Garment Color", ""))),
            "total_qty": clean(record.get("Total Qty", "")),
            "reason": clean(record.get("Missing Information", "")) or "Product Master detail needed",
            "instructions": clean(record.get("Action Required", "")) or "Update this product in Product Master, then recreate Purchase Review.",
            "source": "Product Master",
            "fix_in": "Product Master",
            "decision_key": decision_key,
        }

    if review_rows is None:
        try:
            rows = read_table(path, "Review & Edit", {"Line ID", "Review Status"})
        except Exception:
            rows = []
    else:
        rows = review_rows
    for record in rows:
        line_id = clean(record.get("Line ID", ""))
        if not line_id:
            continue
        if _record_is_service_only(record):
            continue
        include = clean(record.get("Include", "Yes")).casefold()
        if include not in {"yes", "y", "true", "1", "include"}:
            continue
        recovered_color, recovered_size = split_embedded_size(
            clean(record.get("Garment Color", "")), clean(record.get("Size", ""))
        )
        if recovered_color != clean(record.get("Garment Color", "")) or recovered_size != normalize_size(record.get("Size", "")):
            record = dict(record)
            record["Garment Color"] = recovered_color
            record["Size"] = recovered_size
        status = clean(record.get("Review Status", ""))
        resolution = clean(record.get("Resolution", "")).casefold()
        completed_resolution = resolution == "completed in app" or resolution.startswith("instruction:")
        live_blockers = _effective_review_blockers(record)
        if status.casefold() == "ready" or completed_resolution or not live_blockers:
            continue
        unresolved_ids.add(line_id)
        decision_key = clean(record.get("Decision Key", "")) or f"line:{line_id}"
        live_reason = "; ".join(live_blockers)
        original_reason = clean(record.get("Review Reason", ""))
        action_required = clean(record.get("Action Required", ""))
        if original_reason.casefold() != live_reason.casefold():
            action_required = f"Correct this order in Purchase Review: {live_reason}."
        issues_by_key[f"review:{decision_key}"] = {
            "line_id": line_id,
            "source_id": clean(record.get("Source ID", "")),
            "order": clean(record.get("Order Number", "")),
            "employee": clean(record.get("Employee Name", "")),
            "product": clean(record.get("Product #", "")),
            "description": clean(record.get("Description", "")),
            "vendor": clean(record.get("Purchase Vendor", "")),
            "reason": live_reason or "Needs review",
            "instructions": action_required or clean(record.get("Purchase Instructions", "")),
            "source": "Purchase Review",
            "fix_in": clean(record.get("Fix In", "")) or "Purchase Review",
            "decision_key": decision_key,
        }

    return list(issues_by_key.values()), unresolved_ids

def load_purchase_review_snapshot(workbook_path: Path, product_master_path: Path | None = None) -> dict[str, Any]:
    """Load the same authoritative review queue used by final PO preflight.

    Earlier lightweight loading inspected only the visible ``Review & Edit``
    worksheet.  A source line could therefore remain unresolved on hidden
    ``All PO Lines`` while Purchase Review incorrectly displayed zero decisions.
    This loader now takes its issue queue from the full mission-control
    validator, then builds the editable value cache from both worksheets.  The
    UI and final release gate therefore cannot disagree about required review.
    """
    path = Path(workbook_path)
    snapshot: dict[str, Any] = {
        "workbook": path,
        "issues": [],
        "routes": [],
        "review_count": 0,
        "purchase_review_count": 0,
        "product_master_review_count": 0,
        "blocked_route_count": 0,
        "workflow_blocked": False,
        "review_values": {},
        "review_completed": 0,
        "review_remaining": 0,
        "updated": path.stat().st_mtime if path.exists() else 0,
        "lightweight": True,
    }
    if not path.exists():
        return snapshot

    # The full validator is the single source of truth for unresolved decisions.
    # This intentionally favors correctness over the former lightweight shortcut.
    authoritative = load_mission_control_snapshot(path, None, product_master_path)
    issues = [dict(issue) for issue in authoritative.get("issues", [])]

    try:
        review_rows = read_table(path, "Review & Edit", {"Line ID", "Review Status"})
    except Exception:
        review_rows = []
    try:
        all_rows = load_review_lines(path)
    except Exception:
        all_rows = []

    review_source_ids = {
        clean(record.get("Source ID", "")) for record in review_rows
        if clean(record.get("Source ID", ""))
    }
    for issue in issues:
        source_id = clean(issue.get("source_id", ""))
        if (
            source_id
            and clean(issue.get("fix_in", issue.get("source", ""))).casefold() != "product master"
            and source_id not in review_source_ids
        ):
            # The decision exists only in All PO Lines.  The app can still show
            # and save it safely by updating the immutable source row directly.
            issue["source_only"] = True

    snapshot["issues"] = issues
    product_issues = [
        issue for issue in issues
        if clean(issue.get("fix_in", issue.get("source", ""))).casefold() == "product master"
    ]
    review_issues = [
        issue for issue in issues
        if clean(issue.get("fix_in", issue.get("source", ""))).casefold() != "product master"
    ]
    snapshot["product_master_review_count"] = len(product_issues)
    snapshot["review_count"] = len(review_issues)
    snapshot["purchase_review_count"] = len(review_issues)
    snapshot["blocked_route_count"] = int(authoritative.get("blocked_route_count", 0) or 0)
    snapshot["workflow_blocked"] = bool(issues or snapshot["blocked_route_count"])

    values_by_key: dict[str, dict[str, object]] = {}
    included_decisions: dict[str, list[str]] = {}

    # All PO Lines contains the immutable event source and therefore supplies
    # values for source-only decisions.  Visible Review & Edit rows are applied
    # afterward so any user-facing edits remain authoritative.
    combined_rows = list(all_rows) + list(review_rows)
    for row_number, record in enumerate(combined_rows, start=1):
        line_id = clean(record.get("Line ID", ""))
        source_id = clean(record.get("Source ID", ""))
        decision_key = clean(record.get("Decision Key", "")) or (
            f"line:{line_id}" if line_id else (f"source:{source_id}" if source_id else f"row:{row_number}")
        )
        if _record_is_service_only(record):
            continue
        include = clean(record.get("Include", "Yes")).casefold()
        if include not in {"yes", "y", "true", "1", "include"}:
            continue
        values = dict(record)
        recovered_color, recovered_size = split_embedded_size(
            clean(values.get("Garment Color", "")), clean(values.get("Size", ""))
        )
        values["Garment Color"] = recovered_color
        values["Size"] = recovered_size
        values_by_key[decision_key] = values
        if line_id:
            values_by_key[line_id] = values
        if source_id:
            values_by_key[source_id] = values

        status = clean(record.get("Review Status", "")).casefold()
        resolution = clean(record.get("Resolution", "")).casefold()
        if status != "ready" and (
            resolution == "completed in app"
            or resolution.startswith("instruction:")
            or _record_is_effectively_complete(values)
        ):
            status = "ready"
        included_decisions.setdefault(decision_key, []).append(status)

    completed = sum(
        1 for statuses in included_decisions.values()
        if statuses and all(status == "ready" for status in statuses)
    )
    snapshot["review_values"] = values_by_key
    snapshot["review_completed"] = completed
    snapshot["review_remaining"] = len(review_issues)
    return snapshot



def _dashboard_po_numbers(path: Path) -> dict[tuple[str, str], str]:
    try:
        rows = read_table(path, "Dashboard", {"Vendor", "PO Number"})
    except Exception:
        return {}
    output: dict[tuple[str, str], str] = {}
    for record in rows:
        vendor = clean(record.get("Vendor", ""))
        if not vendor or vendor.casefold() in {"needs vendor assignment", "unassigned"}:
            continue
        report = clean(record.get("Report Type", record.get("Report", ""))) or "Combined Vendor Order"
        po_number = clean(record.get("PO Number", ""))
        if po_number:
            output[(vendor.casefold(), canonical_report_type(report))] = po_number
    return output


def _quantity(value: object) -> int:
    try:
        return max(int(float(value)), 0)
    except (TypeError, ValueError):
        return 0


def _line_block_reasons(record: dict[str, object]) -> list[str]:
    """Return the same purchasing blockers enforced by final PO generation.

    Review Status is still honored for customer decisions, but changing a status
    to Ready cannot hide missing vendor, decoration, product, color, size, or
    quantity information. This keeps Mission Control synchronized with the
    workbook and the final purchase-order generator.
    """
    include = clean(record.get("Include", "Yes")).casefold()
    if include not in {"yes", "y", "true", "1", "include"}:
        return []

    vendor = clean(record.get("Purchase Vendor", ""))
    decoration_type = clean(record.get("Decoration Type", ""))
    decoration_color = clean(record.get("Decoration Color", ""))
    decoration_location = normalize_decoration_location(
        record.get("Decoration Location", ""), decoration_type,
        requires_decoration=not is_blank_decoration(decoration_type) and not is_in_house_decoration(decoration_type),
        product_name=record.get("Description", ""), category=record.get("Product Category", ""),
    )
    product_number = clean(record.get("Product #", ""))
    description = clean(record.get("Description", ""))
    garment_color = clean(record.get("Garment Color", record.get("Color", "")))
    garment_color, size = split_embedded_size(garment_color, record.get("Size", ""))
    quantity = _quantity(record.get("Quantity", 0))

    # Product 750 and other registered service/charge rows are not products.
    if _record_is_service_only(record):
        return []

    rules = row_rules({
        "Product Name": description,
        "Style Number": product_number,
        "Product Category": record.get("Product Category", ""),
        "Requires Size": record.get("Requires Size", ""),
        "Requires Color": record.get("Requires Color", ""),
        "Requires Decoration": record.get("Requires Decoration", ""),
        "Decoration Type": decoration_type,
    })
    requires_size = normalize_bool(rules["Requires Size"], True)
    requires_color = normalize_bool(rules["Requires Color"], True)
    requires_decoration = normalize_bool(rules["Requires Decoration"], True)
    if not requires_decoration and not decoration_type:
        decoration_type = BLANK_DECORATION_LABEL
        decoration_color = ""

    reasons: list[str] = []
    review_status = clean(record.get("Review Status", ""))
    if review_status and review_status.casefold() != "ready":
        reasons.append(clean(record.get("Review Reason", "")) or "Review status is not Ready")
    mandatory_note_review = normalize_bool(record.get("Mandatory Note Review", ""), False)
    decoration_decision = clean(record.get("Decoration Decision", ""))
    decision_key = decoration_decision.casefold().replace("—", "-").replace("–", "-")
    valid_decisions = {
        "follow note as written",
        "keep product master default",
        "no decoration",
        "embroidery",
        "screen print",
        "sew on patch",
        "hemming / alteration",
        "do not outsource - ship to orchid",
    }
    if mandatory_note_review and decision_key not in valid_decisions:
        reasons.append("Customer instruction decision required")
    if not vendor:
        reasons.append("Missing purchase vendor")
    if requires_decoration and not decoration_type:
        reasons.append("Missing decoration type")
    if requires_decoration and decoration_type and not is_blank_decoration(decoration_type) and not is_in_house_decoration(decoration_type) and not decoration_color:
        reasons.append("Missing decoration color")
    if requires_decoration and decoration_location == OTHER_CUSTOM:
        reasons.append("Custom decoration location is incomplete")
    if not product_number and not description:
        reasons.append("Missing product number and description")
    if requires_color and not garment_color:
        reasons.append("Missing garment color")
    if requires_size and not size:
        reasons.append("Missing size")
    if quantity <= 0:
        reasons.append("Quantity must be greater than zero")

    return list(dict.fromkeys(reason for reason in reasons if reason))


def line_block_reasons(record: dict[str, object]) -> list[str]:
    """Public purchasing preflight used by the app and final PO workflow.

    Keeping this as the single validation entry point prevents Save & Next,
    Purchase Review, and Step 4 from disagreeing about whether a line is ready.
    """
    return _line_block_reasons(record)


def _line_route_type(decoration: object, report_mode: str) -> str:
    if report_mode != UNIFORM_SIZING_EVENT:
        return "Combined Vendor Order"
    text = clean(decoration)
    lowered = text.casefold()
    if "screen" in lowered:
        return "Screen Printing"
    if is_blank_decoration(text) or is_in_house_decoration(text):
        return "Blank Garments"
    if "embroider" in lowered:
        return "Embroidery"
    return "Needs Routing"


def _live_routes(
    path: Path,
    report_mode: str,
    unresolved_ids: set[str],
    data_root: Path | None,
    lines: list[dict[str, object]] | None = None,
) -> list[dict[str, str]]:
    if lines is None:
        try:
            lines = load_review_lines(path)
        except Exception:
            lines = []

    prepared: list[dict[str, object]] = []
    for record in lines:
        include = clean(record.get("Include", "Yes")).casefold()
        if include not in {"yes", "y", "true", "1", "include"}:
            continue
        if _record_is_service_only(record):
            continue
        vendor = clean(record.get("Purchase Vendor", ""))
        # Missing-vendor lines remain in the Review queue; they are not a fake
        # purchase-order route and therefore must not request a PO number.
        if not vendor:
            continue
        report = _line_route_type(record.get("Decoration Type", ""), report_mode)
        prepared.append({
            "line_id": clean(record.get("Line ID", "")),
            "vendor": vendor,
            "report": report,
            "quantity": _quantity(record.get("Quantity", 0)),
        })

    if report_mode == UNIFORM_SIZING_EVENT:
        embroidery_vendors = {
            clean(row["vendor"]).casefold()
            for row in prepared
            if clean(row["report"]).casefold() == "embroidery"
        }
        for row in prepared:
            if (
                clean(row["report"]).casefold() == "blank garments"
                and clean(row["vendor"]).casefold() in embroidery_vendors
            ):
                row["report"] = "Embroidery"

    # Final output is one consolidated PDF per vendor, so Mission Control asks
    # for one native PO number per vendor rather than one number per decoration.
    grouped: dict[str, dict[str, object]] = {}
    for row in prepared:
        vendor = clean(row["vendor"])
        key = vendor.casefold()
        entry = grouped.setdefault(key, {
            "vendor": vendor,
            "report_type": "Consolidated Vendor Order",
            "pieces": 0,
            "needs_review": False,
            "key": route_key(vendor, "Combined Vendor Order"),
        })
        entry["pieces"] = int(entry["pieces"]) + int(row["quantity"])
        if clean(row["line_id"]) in unresolved_ids or clean(row["report"]).casefold() == "needs routing":
            entry["needs_review"] = True

    workbook_numbers = _dashboard_po_numbers(path)
    overrides = load_po_overrides(data_root, path) if data_root is not None else {}
    routes: list[dict[str, str]] = []
    for entry in grouped.values():
        vendor = clean(entry["vendor"])
        combined_lookup = (vendor.casefold(), "combined vendor order")
        po_number = clean(overrides.get(combined_lookup, "")) or clean(workbook_numbers.get(combined_lookup, ""))
        if not po_number:
            # Preserve compatibility with older workbooks that stored separate
            # decoration-route PO numbers. Reuse the first number found.
            for source in (overrides, workbook_numbers):
                for (stored_vendor, _stored_report), stored_number in source.items():
                    if clean(stored_vendor).casefold() == vendor.casefold() and clean(stored_number):
                        po_number = clean(stored_number)
                        break
                if po_number:
                    break
        routes.append({
            "vendor": vendor,
            "report_type": "Consolidated Vendor Order",
            "po_number": po_number,
            "status": "Needs Review" if bool(entry["needs_review"]) else "Ready",
            "pieces": str(entry["pieces"]),
            "key": clean(entry["key"]),
        })
    routes.sort(key=lambda row: row["vendor"].casefold())
    return routes


def load_mission_control_snapshot(
    workbook_path: Path,
    data_root: Path | None = None,
    product_master_path: Path | None = None,
) -> dict[str, Any]:
    path = Path(workbook_path)
    snapshot: dict[str, Any] = {
        "workbook": path,
        "event_name": "",
        "report_mode": GENERAL_SALES_PERIOD,
        "issues": [],
        "routes": [],
        "line_count": 0,
        "ready_count": 0,
        "review_count": 0,
        "employee_totals": [],
        "employee_count": 0,
        "employee_grand_total": 0.0,
        "updated": path.stat().st_mtime if path.exists() else 0,
    }
    if not path.exists():
        return snapshot

    snapshot["event_name"] = clean(load_event_name(path))
    snapshot["report_mode"] = clean(load_report_mode(path)) or GENERAL_SALES_PERIOD
    employee_totals = load_employee_totals(path)
    employee_summary = employee_totals_summary(employee_totals)
    snapshot["employee_totals"] = employee_totals
    snapshot["employee_count"] = int(employee_summary["employee_count"])
    snapshot["employee_grand_total"] = float(employee_summary["grand_total"])

    issues, unresolved_ids = _review_queue(path)
    issues_by_key: dict[str, dict[str, str]] = {}
    for issue in issues:
        fix_in = clean(issue.get("fix_in", issue.get("source", "issue"))).casefold()
        prefix = "master" if fix_in == "product master" else "review"
        decision_key = clean(issue.get("decision_key", "")) or f"line:{clean(issue.get('line_id', ''))}"
        issues_by_key[f"{prefix}:{decision_key}"] = issue
    blocked_line_ids = set(unresolved_ids)
    try:
        all_lines = load_review_lines(path)
    except Exception:
        all_lines = []

    included_lines = []
    for record in all_lines:
        include = clean(record.get("Include", "Yes")).casefold()
        if include not in {"yes", "y", "true", "1", "include"}:
            continue
        included_lines.append(record)
        line_id = clean(record.get("Line ID", ""))

        # A product that did not match Product Master is permanent setup work,
        # even when a user has corrected the product number on Review & Edit.
        # This closes the legacy loophole where unknown styles could advance as
        # order-only decisions and inherit vendor/decoration data incorrectly.
        product_number = clean(record.get("Product #", ""))
        permanent_reason = clean(record.get("Permanent Review Reason", ""))
        master_match = clean(record.get("Master Match", "")).casefold()
        source_kind = clean(record.get("Source", "")).casefold()
        unknown_master = bool(
            product_number
            and master_match == "no"
            and source_kind != "smart custom item"
        )
        if unknown_master and "product not found in product master" not in permanent_reason.casefold():
            permanent_reason = "; ".join(filter(None, [
                "Product not found in Product Master", permanent_reason
            ]))
        if permanent_reason:
            if line_id:
                blocked_line_ids.add(line_id)
            decision_key = f"product:{product_number or clean(record.get('Description', '')).casefold() or line_id}"
            issues_by_key.setdefault(f"master:{decision_key}", {
                "line_id": line_id,
                "source_id": clean(record.get("Source ID", "")),
                "affected_line_ids": line_id,
                "order": clean(record.get("Order Number", "")),
                "employee": clean(record.get("Employee Name", "")),
                "product": product_number,
                "description": clean(record.get("Description", "")),
                "vendor": clean(record.get("Purchase Vendor", "")),
                "garment_colors": clean(record.get("Garment Color", "")),
                "total_qty": str(_quantity(record.get("Quantity", 0))),
                "reason": permanent_reason,
                "instructions": (
                    f"Add {product_number or 'this product'} to Product Master and complete the permanent purchasing setup."
                    if unknown_master else
                    f"Complete the permanent Product Master setup for {product_number or 'this product'}."
                ),
                "source": "Product Master",
                "fix_in": "Product Master",
                "decision_key": decision_key,
            })

        reasons = _line_block_reasons(record)
        if not reasons:
            continue
        if line_id:
            blocked_line_ids.add(line_id)
        # Visible Product Master and Review & Edit rows are authoritative. Add a
        # validation issue only when a user marked a line Ready while a required
        # field is still missing, rather than duplicating an existing queue item.
        if line_id and line_id in unresolved_ids:
            continue
        decision_key = clean(record.get("Decision Key", "")) or f"line:{line_id or len(included_lines)}"
        issues_by_key.setdefault(f"validation:{decision_key}", {
            "line_id": line_id,
            "source_id": clean(record.get("Source ID", "")),
            "order": clean(record.get("Order Number", "")),
            "employee": clean(record.get("Employee Name", "")),
            "product": clean(record.get("Product #", "")),
            "description": clean(record.get("Description", "")),
            "vendor": clean(record.get("Purchase Vendor", "")),
            "reason": "; ".join(reasons),
            "instructions": clean(record.get("Purchase Instructions", "")),
            "source": "Purchase Review validation",
            "fix_in": "Purchase Review",
            "decision_key": decision_key,
        })

    issues = list(issues_by_key.values())
    issues, resolved_master_lines = _filter_live_product_master_issues(issues, product_master_path)
    blocked_line_ids.difference_update(resolved_master_lines)
    routes = _live_routes(path, snapshot["report_mode"], blocked_line_ids, data_root, lines=all_lines)
    blocked_routes = [route for route in routes if clean(route.get("status", "")).casefold() != "ready"]
    line_count = len(included_lines)
    snapshot["issues"] = issues
    snapshot["product_master_count"] = sum(1 for issue in issues if clean(issue.get("fix_in", "")).casefold() == "product master")
    snapshot["purchase_review_count"] = sum(1 for issue in issues if clean(issue.get("fix_in", "")).casefold() != "product master")
    snapshot["review_count"] = len(issues)
    snapshot["line_count"] = line_count
    snapshot["ready_count"] = max(line_count - len(blocked_line_ids), 0)
    snapshot["routes"] = routes
    snapshot["blocked_route_count"] = len(blocked_routes)
    snapshot["workflow_blocked"] = bool(issues or blocked_routes)
    return snapshot
