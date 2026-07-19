from __future__ import annotations

from collections import OrderedDict
from pathlib import Path
import re
import zipfile
import xml.etree.ElementTree as ET

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"

_SHEET_ROWS_CACHE: OrderedDict[tuple[str, int, int, str], tuple[tuple[object, ...], ...]] = OrderedDict()
_SHEET_ROWS_CACHE_LIMIT = 12


def _cache_rows(key: tuple[str, int, int, str], rows: list[list[object]]) -> None:
    _SHEET_ROWS_CACHE[key] = tuple(tuple(row) for row in rows)
    _SHEET_ROWS_CACHE.move_to_end(key)
    while len(_SHEET_ROWS_CACHE) > _SHEET_ROWS_CACHE_LIMIT:
        _SHEET_ROWS_CACHE.popitem(last=False)


def _column_index(cell_reference: str) -> int:
    letters = re.match(r"[A-Z]+", cell_reference or "")
    if not letters:
        return 0
    value = 0
    for character in letters.group(0):
        value = value * 26 + (ord(character) - 64)
    return value - 1


def _shared_strings(archive: zipfile.ZipFile) -> list[str]:
    try:
        root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    values = []
    for item in root.findall(f"{{{NS_MAIN}}}si"):
        texts = [node.text or "" for node in item.iter(f"{{{NS_MAIN}}}t")]
        values.append("".join(texts))
    return values


def _sheet_paths(archive: zipfile.ZipFile) -> dict[str, str]:
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    rel_targets = {
        rel.attrib.get("Id", ""): rel.attrib.get("Target", "")
        for rel in relationships.findall(f"{{{NS_PKG_REL}}}Relationship")
    }
    result = {}
    sheets = workbook.find(f"{{{NS_MAIN}}}sheets")
    if sheets is None:
        return result
    for sheet in sheets:
        name = sheet.attrib.get("name", "")
        rel_id = sheet.attrib.get(f"{{{NS_REL}}}id", "")
        target = rel_targets.get(rel_id, "")
        if target.startswith("/"):
            target = target.lstrip("/")
        elif not target.startswith("xl/"):
            target = "xl/" + target.lstrip("/")
        result[name] = target
    return result


def read_sheet_rows(xlsx_path: Path, sheet_name: str) -> list[list[object]]:
    path = Path(xlsx_path)
    stat = path.stat()
    cache_key = (str(path.resolve()), int(stat.st_mtime_ns), int(stat.st_size), str(sheet_name))
    cached = _SHEET_ROWS_CACHE.get(cache_key)
    if cached is not None:
        _SHEET_ROWS_CACHE.move_to_end(cache_key)
        return [list(row) for row in cached]
    with zipfile.ZipFile(path, "r") as archive:
        sheets = _sheet_paths(archive)
        if sheet_name not in sheets:
            raise ValueError(f"Workbook does not contain a '{sheet_name}' sheet.")
        shared = _shared_strings(archive)
        root = ET.fromstring(archive.read(sheets[sheet_name]))
        sheet_data = root.find(f"{{{NS_MAIN}}}sheetData")
        if sheet_data is None:
            _cache_rows(cache_key, [])
            return []
        output = []
        for row in sheet_data.findall(f"{{{NS_MAIN}}}row"):
            values: dict[int, object] = {}
            for cell in row.findall(f"{{{NS_MAIN}}}c"):
                index = _column_index(cell.attrib.get("r", ""))
                cell_type = cell.attrib.get("t", "")
                value_node = cell.find(f"{{{NS_MAIN}}}v")
                if cell_type == "inlineStr":
                    texts = [node.text or "" for node in cell.iter(f"{{{NS_MAIN}}}t")]
                    value: object = "".join(texts)
                elif value_node is None:
                    value = ""
                else:
                    raw = value_node.text or ""
                    if cell_type == "s":
                        try:
                            value = shared[int(raw)]
                        except (ValueError, IndexError):
                            value = raw
                    elif cell_type == "b":
                        value = raw == "1"
                    elif cell_type in {"str", "e"}:
                        value = raw
                    else:
                        try:
                            number = float(raw)
                            value = int(number) if number.is_integer() else number
                        except ValueError:
                            value = raw
                values[index] = value
            if values:
                row_values = [""] * (max(values) + 1)
                for index, value in values.items():
                    row_values[index] = value
                output.append(row_values)
            else:
                output.append([])
        _cache_rows(cache_key, output)
        return output


def read_table(xlsx_path: Path, sheet_name: str, required_headers: set[str]) -> list[dict[str, object]]:
    rows = read_sheet_rows(xlsx_path, sheet_name)
    header_index = None
    headers: list[str] = []
    normalized_required = {str(value).strip().casefold() for value in required_headers}
    for index, row in enumerate(rows):
        candidate = [str(value).strip() if value is not None else "" for value in row]
        normalized = {value.casefold() for value in candidate if value}
        if normalized_required.issubset(normalized):
            header_index = index
            headers = candidate
            break
    if header_index is None:
        required = ", ".join(sorted(required_headers))
        raise ValueError(f"Could not find the expected headers on '{sheet_name}': {required}")

    records = []
    for row in rows[header_index + 1:]:
        values = list(row) + [""] * max(0, len(headers) - len(row))
        record = {headers[i]: values[i] for i in range(len(headers)) if headers[i]}
        if any(str(value).strip() for value in record.values() if value is not None):
            records.append(record)
    return records


def load_review_lines(xlsx_path: Path) -> list[dict[str, object]]:
    required = {"Include", "Purchase Vendor", "Product #", "Description", "Quantity"}
    records = read_table(xlsx_path, "All PO Lines", required)

    editable_fields = {
        "Include", "Do Not Outsource", "Purchase Vendor", "Decoration Type", "Decoration Location", "Decoration Placement Instructions", "Decoration Color",
        "Product #", "Description", "Garment Color", "Size", "Quantity",
        "Company", "Employee Name", "Order Number", "Shopify Order Notes",
        "Purchase Instructions", "Review Status", "Review Reason", "Resolution",
        "Decoration Decision",
    }
    by_id = {
        str(record.get("Line ID", "") or "").strip(): record
        for record in records
        if str(record.get("Line ID", "") or "").strip()
    }

    def apply_edits(edits: list[dict[str, object]], *, blank_values_are_authoritative: bool) -> None:
        for edit in edits:
            line_id = str(edit.get("Line ID", "") or "").strip()
            target = by_id.get(line_id)
            if target is None:
                continue
            for field in editable_fields:
                if field not in edit:
                    continue
                value = edit[field]
                if not blank_values_are_authoritative and not str(value or "").strip():
                    continue
                target[field] = value

    # Inferred Routing and legacy permanent surfaces are suggestions. Apply them
    # first and ignore blank cells so they cannot erase a later user correction.
    try:
        inferred_edits = read_table(xlsx_path, "Inferred Routing", {"Line ID"})
    except ValueError:
        inferred_edits = []
    apply_edits(inferred_edits, blank_values_are_authoritative=False)

    permanent_edits = []
    for sheet_name in ("Product Master", "Product Master Needed", "Purchasing Rules Needed"):
        try:
            permanent_edits = read_table(xlsx_path, sheet_name, {"Line ID"})
            break
        except ValueError:
            continue
    apply_edits(permanent_edits, blank_values_are_authoritative=False)

    # Review & Edit is the authoritative current-order surface and must be last.
    # This fixes the Step 3/Step 4 mismatch where a hidden Inferred Routing row
    # could overwrite 5028/5326 back to blank after Save & Next showed Ready.
    try:
        review_edits = read_table(xlsx_path, "Review & Edit", {"Line ID"})
    except ValueError:
        review_edits = []
    apply_edits(review_edits, blank_values_are_authoritative=True)

    try:
        manual = read_table(xlsx_path, "Manual Items", required)
    except ValueError:
        manual = []
    return records + manual


def load_report_mode(xlsx_path: Path) -> str:
    from modules.report_modes import GENERAL_SALES_PERIOD, normalize_report_mode

    try:
        rows = read_sheet_rows(xlsx_path, "Settings")
    except (ValueError, KeyError, zipfile.BadZipFile):
        return GENERAL_SALES_PERIOD

    for row_index, row in enumerate(rows):
        for column_index, value in enumerate(row):
            if str(value or "").strip().casefold() == "purchase order mode":
                if row_index + 1 < len(rows) and column_index < len(rows[row_index + 1]):
                    return normalize_report_mode(rows[row_index + 1][column_index])
    return GENERAL_SALES_PERIOD



def load_decoration_fulfillment(xlsx_path: Path) -> str:
    from modules.decoration_fulfillment import STANDARD_ORCHID_WORKFLOW, normalize_decoration_fulfillment

    try:
        rows = read_sheet_rows(xlsx_path, "Settings")
    except (ValueError, KeyError, zipfile.BadZipFile):
        return STANDARD_ORCHID_WORKFLOW

    for row_index, row in enumerate(rows):
        for column_index, value in enumerate(row):
            if str(value or "").strip().casefold() == "decoration fulfillment":
                if row_index + 1 < len(rows) and column_index < len(rows[row_index + 1]):
                    return normalize_decoration_fulfillment(rows[row_index + 1][column_index])
    return STANDARD_ORCHID_WORKFLOW

def load_event_name(xlsx_path: Path) -> str:
    try:
        rows = read_sheet_rows(xlsx_path, "Settings")
    except (ValueError, KeyError, zipfile.BadZipFile):
        return ""

    for row_index, row in enumerate(rows):
        for column_index, value in enumerate(row):
            if str(value or "").strip().casefold() == "event name":
                if row_index + 1 < len(rows) and column_index < len(rows[row_index + 1]):
                    return " ".join(str(rows[row_index + 1][column_index] or "").split())
    return ""



def _canonical_report_type(value: object) -> str:
    text = " ".join(str(value or "").split()).casefold()
    if "screen" in text:
        return "screen printing"
    if "embroider" in text:
        return "embroidery"
    if "blank" in text or text in {"none", "no decoration"}:
        return "blank garments"
    if "combined" in text:
        return "combined vendor order"
    return text

def load_po_numbers(xlsx_path: Path) -> dict[tuple[str, str], str]:
    try:
        records = read_table(xlsx_path, "Dashboard", {"Vendor", "PO Number"})
    except (ValueError, KeyError, zipfile.BadZipFile):
        try:
            records = read_table(xlsx_path, "Summary", {"Vendor", "PO Number"})
        except (ValueError, KeyError, zipfile.BadZipFile):
            return {}
    result: dict[tuple[str, str], str] = {}
    for record in records:
        vendor = " ".join(str(record.get("Vendor", "") or "").split())
        report = " ".join(str(record.get("Report Type", record.get("Report", "")) or "").split())
        if not report:
            report = "Combined Vendor Order"
        po_number = " ".join(str(record.get("PO Number", "") or "").split())
        if vendor and po_number:
            result[(vendor.casefold(), _canonical_report_type(report))] = po_number
    return result
