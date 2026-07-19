from __future__ import annotations

from datetime import datetime
from pathlib import Path
import shutil

from modules.purchase_order_generator import safe_filename


def archive_completed_event(*, reports_root: Path, event_name: str, review_workbook: Path, purchase_order_dir: Path, source_csv: Path | None = None) -> Path:
    stamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    label = safe_filename(event_name or "General Sales Period")
    archive_dir = reports_root / "Event Archive" / f"{label} - {stamp}"
    archive_dir.mkdir(parents=True, exist_ok=True)

    shutil.copy2(review_workbook, archive_dir / review_workbook.name)
    if source_csv and source_csv.exists():
        shutil.copy2(source_csv, archive_dir / source_csv.name)

    po_target = archive_dir / "Purchase Orders"
    po_target.mkdir(parents=True, exist_ok=True)
    for pdf in purchase_order_dir.glob("*.pdf"):
        shutil.copy2(pdf, po_target / pdf.name)

    return archive_dir
