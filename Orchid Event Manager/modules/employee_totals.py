from __future__ import annotations

from datetime import datetime
from pathlib import Path
import shutil
from typing import Any

from openpyxl import load_workbook


SHEET_NAME = "Employee Totals"
HEADERS = [
    "Company / Department",
    "Employee Name",
    "Order Number(s)",
    "Order Total",
    "Discount",
    "Total",
]


def clean(value: object) -> str:
    return " ".join(str(value or "").split())


def money(value: object) -> float:
    try:
        if value is None or value == "":
            return 0.0
        if isinstance(value, str):
            value = value.replace("$", "").replace(",", "").strip()
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def _find_header_row(sheet) -> tuple[int, dict[str, int]]:
    wanted = {header.casefold(): header for header in HEADERS}
    for row in range(1, min(sheet.max_row, 40) + 1):
        columns: dict[str, int] = {}
        for column in range(1, min(sheet.max_column, 20) + 1):
            value = clean(sheet.cell(row=row, column=column).value)
            if value.casefold() in wanted:
                columns[wanted[value.casefold()]] = column
        if all(header in columns for header in HEADERS):
            return row, columns
    raise ValueError("Could not find the Employee Totals table headers.")


def load_employee_totals(workbook_path: Path) -> list[dict[str, Any]]:
    path = Path(workbook_path)
    if not path.exists():
        return []
    workbook = load_workbook(path, data_only=False, read_only=True)
    try:
        if SHEET_NAME not in workbook.sheetnames:
            return []
        sheet = workbook[SHEET_NAME]
        header_row, columns = _find_header_row(sheet)
        records: list[dict[str, Any]] = []
        for row in range(header_row + 1, sheet.max_row + 1):
            employee = clean(sheet.cell(row=row, column=columns["Employee Name"]).value)
            company = clean(sheet.cell(row=row, column=columns["Company / Department"]).value)
            order_numbers = clean(sheet.cell(row=row, column=columns["Order Number(s)"]).value)
            label = clean(sheet.cell(row=row, column=columns["Discount"]).value)
            if label.casefold() == "grand total":
                break
            if not any((employee, company, order_numbers)):
                continue
            order_total = money(sheet.cell(row=row, column=columns["Order Total"]).value)
            discount = money(sheet.cell(row=row, column=columns["Discount"]).value)
            records.append({
                "row": row,
                "company": company,
                "employee": employee,
                "order_numbers": order_numbers,
                "order_total": order_total,
                "discount": discount,
                "total": order_total - discount,
            })
        return records
    finally:
        workbook.close()


def save_employee_discounts(workbook_path: Path, discounts_by_row: dict[int, float]) -> Path:
    path = Path(workbook_path)
    if not path.exists():
        raise FileNotFoundError(path)
    excel_lock = path.parent / f"~${path.name}"
    if excel_lock.exists():
        raise PermissionError(f"Excel is currently using {path.name}")

    backup_dir = path.parent / "Backups"
    backup_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = backup_dir / f"{path.stem}__before_employee_totals_{stamp}{path.suffix}"
    shutil.copy2(path, backup)

    workbook = load_workbook(path)
    try:
        if SHEET_NAME not in workbook.sheetnames:
            raise ValueError("This Purchase Review does not contain an Employee Totals sheet.")
        sheet = workbook[SHEET_NAME]
        header_row, columns = _find_header_row(sheet)
        order_col = columns["Order Total"]
        discount_col = columns["Discount"]
        total_col = columns["Total"]

        last_data_row = header_row
        for row in range(header_row + 1, sheet.max_row + 1):
            label = clean(sheet.cell(row=row, column=discount_col).value)
            if label.casefold() == "grand total":
                break
            employee = clean(sheet.cell(row=row, column=columns["Employee Name"]).value)
            company = clean(sheet.cell(row=row, column=columns["Company / Department"]).value)
            order_numbers = clean(sheet.cell(row=row, column=columns["Order Number(s)"]).value)
            if not any((employee, company, order_numbers)):
                continue
            last_data_row = row
            discount = max(money(discounts_by_row.get(row, sheet.cell(row=row, column=discount_col).value)), 0.0)
            sheet.cell(row=row, column=discount_col).value = discount
            sheet.cell(row=row, column=discount_col).number_format = "$#,##0.00"
            order_letter = sheet.cell(row=row, column=order_col).column_letter
            discount_letter = sheet.cell(row=row, column=discount_col).column_letter
            sheet.cell(row=row, column=total_col).value = f"={order_letter}{row}-{discount_letter}{row}"
            sheet.cell(row=row, column=total_col).number_format = "$#,##0.00"

        grand_row = None
        for row in range(header_row + 1, sheet.max_row + 3):
            if clean(sheet.cell(row=row, column=discount_col).value).casefold() == "grand total":
                grand_row = row
                break
        if grand_row is None:
            grand_row = last_data_row + 2
            sheet.cell(row=grand_row, column=discount_col).value = "Grand Total"
        total_letter = sheet.cell(row=grand_row, column=total_col).column_letter
        first_data_row = header_row + 1
        if last_data_row >= first_data_row:
            sheet.cell(row=grand_row, column=total_col).value = f"=SUM({total_letter}{first_data_row}:{total_letter}{last_data_row})"
        else:
            sheet.cell(row=grand_row, column=total_col).value = 0
        sheet.cell(row=grand_row, column=total_col).number_format = "$#,##0.00"

        # The desktop dashboard is now the normal Employee Totals interface.
        sheet.sheet_state = "hidden"
        calculation = getattr(workbook, "calculation", None)
        if calculation is not None:
            calculation.fullCalcOnLoad = True
            calculation.forceFullCalc = True
            calculation.calcMode = "auto"
        workbook.save(path)
    except Exception:
        # Preserve the user's original workbook if saving fails partway through.
        workbook.close()
        shutil.copy2(backup, path)
        raise
    finally:
        try:
            workbook.close()
        except Exception:
            pass
    return backup


def employee_totals_summary(records: list[dict[str, Any]]) -> dict[str, float | int]:
    return {
        "employee_count": len(records),
        "grand_total": round(sum(money(row.get("total", 0)) for row in records), 2),
    }
