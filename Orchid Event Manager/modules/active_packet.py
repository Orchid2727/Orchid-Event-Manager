from __future__ import annotations

"""Identity checks for Orchid's one active purchase packet.

The app retains prior workbooks and completed packets intentionally.  They are
valuable records, but they must never become the active event simply because a
file happens to be newer than another file.  This module binds the active state
to one selected CSV and one protected Purchase Review workbook.
"""

from hashlib import sha256
import json
from pathlib import Path
import uuid

from modules.xlsx_reader import load_system_info_value


def source_signature(path: Path | str | None) -> str:
    """Return a stable content hash for an order export, or blank when absent."""
    if not path:
        return ""
    candidate = Path(path).expanduser()
    if not candidate.is_file():
        return ""
    digest = sha256()
    with candidate.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def new_active_packet_state(source_csv: Path | str) -> dict:
    """Create the minimum durable state for a newly selected replacement CSV."""
    source = Path(source_csv).expanduser().resolve()
    signature = source_signature(source)
    if not signature:
        raise FileNotFoundError(f"The selected order CSV was not found: {source}")
    return {
        "active_packet_id": f"EVENT-{uuid.uuid4().hex.upper()}",
        "source_csv": str(source),
        "source_signature": signature,
    }


def _workbook_lock_details(workbook_path: Path | str) -> dict:
    workbook = Path(workbook_path).expanduser().resolve()
    active_packet_id = load_system_info_value(workbook, "Active Packet ID")
    lock_id = load_system_info_value(workbook, "Packet Lock ID")
    manifest_text = load_system_info_value(workbook, "Packet Lock Manifest Path")
    manifest_hash = load_system_info_value(workbook, "Packet Lock Manifest SHA256")
    if not active_packet_id or not lock_id or not manifest_text or not manifest_hash:
        raise ValueError("Purchase Review is missing its protected packet identity. Regenerate Purchase Review.")
    manifest_path = Path(manifest_text).expanduser()
    if not manifest_path.is_file():
        raise FileNotFoundError("The Purchase Review's protected packet manifest is missing. Regenerate Purchase Review.")
    actual_manifest_hash = source_signature(manifest_path)
    if actual_manifest_hash != manifest_hash:
        raise ValueError("The Purchase Review's protected packet manifest no longer matches. Regenerate Purchase Review.")
    try:
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        source_info = manifest.get("files", {}).get("source", {})
        locked_source_hash = str(source_info.get("sha256", "") or "").strip()
    except (OSError, ValueError, TypeError) as exc:
        raise ValueError("The Purchase Review's protected packet manifest could not be read.") from exc
    if not locked_source_hash:
        raise ValueError("The Purchase Review's protected packet manifest has no source CSV identity.")
    return {
        "active_review_workbook": str(workbook),
        "active_review_packet_id": active_packet_id,
        "active_review_packet_lock_id": lock_id,
        "active_review_lock_manifest_sha256": manifest_hash,
        "active_review_source_signature": locked_source_hash,
    }


def bind_active_workbook(state: dict, workbook_path: Path | str) -> dict:
    """Attach one protected review workbook to an already selected CSV packet."""
    result = dict(state or {})
    details = _workbook_lock_details(workbook_path)
    expected_source = str(result.get("source_signature", "") or "").strip()
    if not expected_source:
        raise ValueError("The active packet is missing its CSV identity.")
    if details["active_review_packet_id"] != str(result.get("active_packet_id", "") or "").strip():
        raise ValueError(
            "The generated Purchase Review belongs to a different event packet and was blocked. "
            "No old event data was attached."
        )
    if details["active_review_source_signature"] != expected_source:
        raise ValueError(
            "The generated Purchase Review belongs to a different CSV and was blocked. "
            "No old event data was attached."
        )
    result.update(details)
    return result


def active_workbook_matches_state(state: dict, workbook_path: Path | str | None) -> tuple[bool, str]:
    """Return whether a workbook is exactly the active packet's protected review."""
    if not isinstance(state, dict):
        return False, "No active packet state exists."
    active_id = str(state.get("active_packet_id", "") or "").strip()
    saved_workbook = str(state.get("active_review_workbook", "") or "").strip()
    expected_source = str(state.get("source_signature", "") or "").strip()
    if not active_id or not saved_workbook or not expected_source:
        return False, "The active packet identity is incomplete."
    source_csv = str(state.get("source_csv", "") or "").strip()
    if not source_csv or source_signature(source_csv) != expected_source:
        return False, "The selected order CSV has changed or is unavailable."
    if not workbook_path:
        return False, "No active Purchase Review is selected."
    workbook = Path(workbook_path).expanduser()
    if not workbook.is_file():
        return False, "The active Purchase Review file is missing."
    try:
        if workbook.resolve() != Path(saved_workbook).expanduser().resolve():
            return False, "This Purchase Review belongs to a different event."
        details = _workbook_lock_details(workbook)
    except (OSError, ValueError) as exc:
        return False, str(exc)
    if details["active_review_packet_lock_id"] != str(state.get("active_review_packet_lock_id", "") or ""):
        return False, "This Purchase Review has a different protected packet lock."
    if details["active_review_packet_id"] != active_id:
        return False, "This Purchase Review belongs to a different event packet."
    if details["active_review_lock_manifest_sha256"] != str(state.get("active_review_lock_manifest_sha256", "") or ""):
        return False, "This Purchase Review has a different protected packet manifest."
    if details["active_review_source_signature"] != expected_source:
        return False, "This Purchase Review was created from a different CSV."
    return True, ""
