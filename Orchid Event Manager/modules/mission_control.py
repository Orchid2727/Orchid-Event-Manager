from __future__ import annotations

from datetime import datetime
import json
from pathlib import Path
from typing import Any

from modules.blank_garment_rules import BLANK_DECORATION_LABEL, is_blank_decoration
from modules.decoration_locations import OTHER_CUSTOM, normalize_decoration_location
from modules.report_modes import GENERAL_SALES_PERIOD, UNIFORM_SIZING_EVENT
from modules.purchase_order_generator import normalize_size
from modules.purchase_rules import normalize_bool, row_rules
from modules.employee_totals import employee_totals_summary, load_employee_totals
from modules.internal_services import is_in_house_decoration, is_in_house_service_product
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


def _review_queue(path: Path) -> tuple[list[dict[str, str]], set[str]]:
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

    try:
        rows = read_table(path, "Review & Edit", {"Line ID", "Review Status"})
    except Exception:
        rows = []
    for record in rows:
        line_id = clean(record.get("Line ID", ""))
        if not line_id:
            continue
        include = clean(record.get("Include", "Yes")).casefold()
        if include not in {"yes", "y", "true", "1", "include"}:
            continue
        status = clean(record.get("Review Status", ""))
        if status.casefold() == "ready":
            continue
        unresolved_ids.add(line_id)
        decision_key = clean(record.get("Decision Key", "")) or f"line:{line_id}"
        issues_by_key[f"review:{decision_key}"] = {
            "line_id": line_id,
            "order": clean(record.get("Order Number", "")),
            "employee": clean(record.get("Employee Name", "")),
            "product": clean(record.get("Product #", "")),
            "description": clean(record.get("Description", "")),
            "vendor": clean(record.get("Purchase Vendor", "")),
            "reason": clean(record.get("Review Reason", "")) or "Needs review",
            "instructions": clean(record.get("Action Required", "")) or clean(record.get("Purchase Instructions", "")),
            "source": "Purchase Review",
            "fix_in": clean(record.get("Fix In", "")) or "Purchase Review",
            "decision_key": decision_key,
        }

    return list(issues_by_key.values()), unresolved_ids

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
    size = normalize_size(record.get("Size", ""))
    quantity = _quantity(record.get("Quantity", 0))

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
) -> list[dict[str, str]]:
    try:
        lines = load_review_lines(path)
    except Exception:
        lines = []

    prepared: list[dict[str, object]] = []
    for record in lines:
        include = clean(record.get("Include", "Yes")).casefold()
        if include not in {"yes", "y", "true", "1", "include"}:
            continue
        if clean(record.get("Do Not Outsource", "No")).casefold() in {"yes", "y", "true", "1"}:
            # This rare line-level exception is purchased manually and appears
            # on the consolidated Non-Included Items internal document.
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


def load_mission_control_snapshot(workbook_path: Path, data_root: Path | None = None) -> dict[str, Any]:
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
    routes = _live_routes(path, snapshot["report_mode"], blocked_line_ids, data_root)
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
