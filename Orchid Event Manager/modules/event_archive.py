from __future__ import annotations

from datetime import datetime
from pathlib import Path
import hashlib
import json
import shutil

from modules.purchase_order_generator import safe_filename


def validate_release_packet(packet_dir: Path, *, archived: bool = False) -> tuple[bool, str]:
    """Verify the protected release manifest before treating a packet as ready.

    Final packets keep PDFs at their root. Event Archive stores the same PDFs in
    its ``Purchase Orders`` subfolder, so archived validation remaps only those
    entries and verifies the original manifest's size and SHA-256 values.
    """
    root = Path(packet_dir)
    manifest_path = root / "FINAL_RELEASE_MANIFEST.json"
    if not manifest_path.is_file():
        return False, "missing FINAL_RELEASE_MANIFEST.json"
    try:
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    except Exception as error:
        return False, f"cannot read final-release manifest ({error})"
    if manifest.get("release_status") != "PASS":
        return False, "final-release manifest is not marked PASS"
    entries = manifest.get("files")
    if not isinstance(entries, list) or not entries:
        return False, "final-release manifest has no file inventory"
    for entry in entries:
        relative = Path(str(entry.get("path", "")))
        if not relative.parts or relative.is_absolute() or ".." in relative.parts:
            return False, "final-release manifest contains an unsafe file path"
        candidate = root / relative
        if archived and candidate.suffix.casefold() == ".pdf":
            candidate = root / "Purchase Orders" / candidate.name
        if not candidate.is_file():
            return False, f"missing released file: {relative}"
        if candidate.stat().st_size != int(entry.get("size", -1)):
            return False, f"size mismatch: {relative}"
        digest = hashlib.sha256(candidate.read_bytes()).hexdigest()
        if digest != str(entry.get("sha256", "")):
            return False, f"hash mismatch: {relative}"
    return True, "verified"


def archive_completed_event(
    *,
    reports_root: Path,
    event_name: str,
    review_workbook: Path,
    purchase_order_dir: Path,
    source_csv: Path | None = None,
    job_logo_path: Path | None = None,
    job_logo_metadata_path: Path | None = None,
) -> Path:
    stamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    label = safe_filename(event_name or "General Sales Period")
    archive_dir = reports_root / "Event Archive" / f"{label} - {stamp}"
    archive_dir.mkdir(parents=True, exist_ok=True)

    shutil.copy2(review_workbook, archive_dir / review_workbook.name)
    # A blank saved source value becomes Path(".") if a legacy state file is
    # restored without validation. Only copy actual source files; never ask
    # shutil to copy a directory into an archive.
    if source_csv and source_csv.is_file():
        shutil.copy2(source_csv, archive_dir / source_csv.name)

    # Keep the event's job logo with the completed packet as well. That makes a
    # later report-only reopening self-contained instead of depending on the
    # original working directory still being present.
    logo_candidates = [Path(job_logo_path)] if job_logo_path else []
    for stem in (
        f"{safe_filename(event_name or 'Current Event')}__Outsourced_Job_Logo",
        f"{review_workbook.stem}__Outsourced_Job_Logo",
    ):
        logo_candidates.extend(review_workbook.parent / f"{stem}{suffix}" for suffix in (".png", ".jpg", ".jpeg", ".webp"))
    for candidate in logo_candidates:
        if candidate.exists():
            shutil.copy2(candidate, archive_dir / candidate.name)
            break

    # The optional production name is stored beside the image so an archived
    # packet can later recreate the exact outsourced-report header.  It is a
    # separate file because it must remain editable without altering the image.
    metadata_candidates = [Path(job_logo_metadata_path)] if job_logo_metadata_path else []
    for stem in (
        f"{safe_filename(event_name or 'Current Event')}__Outsourced_Job_Logo_Metadata",
        f"{review_workbook.stem}__Outsourced_Job_Logo_Metadata",
    ):
        metadata_candidates.append(review_workbook.parent / f"{stem}.json")
    copied_metadata = None
    for candidate in metadata_candidates:
        if candidate.exists():
            shutil.copy2(candidate, archive_dir / candidate.name)
            copied_metadata = candidate
            break

    # Per-color cover artwork is stored beside the event workbook and listed
    # in the same metadata file. Copy only those named entries, never every
    # image in the working directory, so an archived cover remains complete
    # and self-contained when it is reopened later.
    if copied_metadata:
        try:
            artwork = json.loads(copied_metadata.read_text(encoding="utf-8")).get("artwork", {})
        except (OSError, ValueError, TypeError, AttributeError):
            artwork = {}
        if isinstance(artwork, dict):
            for entry in artwork.values():
                filename = str(entry.get("filename", "")).strip() if isinstance(entry, dict) else ""
                if not filename or Path(filename).name != filename:
                    continue
                candidate = review_workbook.parent / filename
                if candidate.is_file():
                    shutil.copy2(candidate, archive_dir / filename)

    po_target = archive_dir / "Purchase Orders"
    po_target.mkdir(parents=True, exist_ok=True)
    for pdf in purchase_order_dir.glob("*.pdf"):
        shutil.copy2(pdf, po_target / pdf.name)

    # Preserve the protected packet proof with the archive. The final report
    # folder contains the locked source CSV, Product Master, Never Outsource
    # overrides, source ledger, and all release hashes.
    lock_source = purchase_order_dir / "Packet Lock"
    if lock_source.is_dir():
        shutil.copytree(lock_source, archive_dir / "Packet Lock", dirs_exist_ok=True)
    for proof_name in ("PROTECTED_ORDER_RELEASE_CHECK.txt", "FINAL_RELEASE_MANIFEST.json"):
        proof = purchase_order_dir / proof_name
        if proof.is_file():
            shutil.copy2(proof, archive_dir / proof.name)

    return archive_dir
