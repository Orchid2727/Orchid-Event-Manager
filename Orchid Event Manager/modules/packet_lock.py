from __future__ import annotations

"""Protected purchase-packet input locks and final release verification.

A packet lock is created at Purchase Review generation time from three immutable
inputs: the selected order CSV, the live Product Master, and the optional
Never-Outsource override file. A separate source ledger records the imported
Shopify identity for every source line. Final PDFs are blocked unless the
workbook, the lock bundle, and (for an active event) the current live routing
files still agree.
"""

from datetime import datetime
from hashlib import sha256
import json
import os
from pathlib import Path
import re
import shutil
import uuid
from typing import Any, Iterable, Mapping

import pandas as pd

from modules.master_sync import live_product_master_path, product_master_signature
from modules.outsource_rules import ROUTING_POLICY_VERSION, vendor_never_outsource
from modules.paths import data_dir
from modules.product_resolver import load_extended_master, normalize_style, resolve_product
from modules.xlsx_reader import load_decoration_fulfillment, load_report_mode, load_system_info_value

LOCK_SCHEMA_VERSION = "1"
LOCK_ROOT_NAME = "packet_locks"
MANIFEST_NAME = "packet_lock_manifest.json"
LEDGER_NAME = "source_ledger.csv"
MASTER_NAME = "product_master_locked.csv"
SOURCE_NAME = "source_orders_locked.csv"
OVERRIDES_NAME = "never_outsource_overrides_locked.json"

IMMUTABLE_COLUMNS = [
    "Source ID",
    "Source Occurrence",
    "Original Quantity",
    "Original Order Number",
    "Original Order Date",
    "Original Company",
    "Original Employee Name",
    "Original Shopify Line",
    "Original Shopify SKU",
    "Shopify Order Notes",
    "Shopify Line Notes",
    "Original Variant Title",
    "Original Variant Option 1",
    "Original Variant Option 2",
    "Original Variant Option 3",
    "Original Variant Color",
    "Original Variant Size",
]


def _clean(value: object) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return re.sub(r"\s+", " ", str(value).strip())


def _canonical_quantity(value: object) -> int:
    try:
        return int(float(str(value).replace(",", "").strip()))
    except (TypeError, ValueError):
        return 0


def sha256_file(path: Path) -> str:
    path = Path(path)
    digest = sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _override_path(master_path: Path) -> Path:
    return Path(master_path).parent / "never_outsource_overrides.json"


def _load_overrides(path: Path) -> dict[str, bool]:
    if not Path(path).exists():
        return {}
    try:
        payload = json.loads(Path(path).read_text(encoding="utf-8"))
    except Exception as exc:
        raise RuntimeError(f"Never Outsource override file could not be read: {path}\n{exc}") from exc
    if not isinstance(payload, dict):
        raise RuntimeError(f"Never Outsource override file is not a JSON object: {path}")
    return {str(key).strip(): bool(value) for key, value in payload.items() if str(key).strip()}


def _style_override(overrides: Mapping[str, bool], style: object, product: object = "") -> bool:
    style_key = normalize_style(style)
    if style_key and overrides.get(f"style:{style_key}", False):
        return True
    product_key = re.sub(r"\s+", " ", _clean(product).casefold()).strip()
    return bool(product_key and overrides.get(f"product:{product_key}", False))


def assert_live_product_master(master_path: Path) -> Path:
    """Reject accidental test/snapshot masters in normal app operation."""
    requested = Path(master_path).expanduser().resolve()
    live = Path(live_product_master_path()).expanduser().resolve()
    allow_test = os.environ.get("ORCHID_ALLOW_TEST_MASTER", "").strip().casefold() in {"1", "yes", "true"}
    if requested != live and not allow_test:
        raise RuntimeError(
            "Protected build blocked Purchase Review creation because the supplied Product Master is not the live "
            "Orchid Product Master.\n\n"
            f"Expected: {live}\nReceived: {requested}\n\n"
            "No workbook was created."
        )
    return requested


def _ledger_frame(detail: pd.DataFrame) -> pd.DataFrame:
    result = pd.DataFrame()
    for column in IMMUTABLE_COLUMNS:
        if column in detail.columns:
            result[column] = detail[column]
        else:
            result[column] = ""
    result = result.fillna("")
    if "Source ID" in result.columns:
        result["Source ID"] = result["Source ID"].map(_clean)
        result = result[result["Source ID"].ne("")].copy()
    if "Original Quantity" in result.columns:
        result["Original Quantity"] = result["Original Quantity"].map(_canonical_quantity)
    result = result.drop_duplicates(subset=["Source ID"], keep=False)
    if len(result) != int(detail.get("Source ID", pd.Series(dtype=str)).map(_clean).ne("").sum()):
        raise RuntimeError("Protected packet lock could not be created because Source IDs are missing or duplicated.")
    return result.sort_values("Source ID", kind="stable").reset_index(drop=True)


def create_packet_lock(
    *,
    source_csv_path: Path,
    product_master_path: Path,
    detail: pd.DataFrame,
    report_mode: str,
    event_name: str,
    decoration_fulfillment: str,
) -> dict[str, Any]:
    master_path = assert_live_product_master(product_master_path)
    source_path = Path(source_csv_path).expanduser().resolve()
    if not source_path.is_file():
        raise FileNotFoundError(f"The selected order CSV was not found: {source_path}")
    if not master_path.is_file():
        raise FileNotFoundError(f"The live Product Master was not found: {master_path}")

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    lock_id = f"LOCK-{stamp}-{uuid.uuid4().hex[:10].upper()}"
    lock_dir = data_dir() / LOCK_ROOT_NAME / lock_id
    lock_dir.mkdir(parents=True, exist_ok=False)

    locked_source = lock_dir / SOURCE_NAME
    locked_master = lock_dir / MASTER_NAME
    locked_overrides = lock_dir / OVERRIDES_NAME
    ledger_path = lock_dir / LEDGER_NAME
    manifest_path = lock_dir / MANIFEST_NAME

    shutil.copy2(source_path, locked_source)
    shutil.copy2(master_path, locked_master)
    override_source = _override_path(master_path)
    if override_source.is_file():
        shutil.copy2(override_source, locked_overrides)
    else:
        locked_overrides.write_text("{}\n", encoding="utf-8")

    ledger = _ledger_frame(detail)
    ledger.to_csv(ledger_path, index=False)

    manifest = {
        "schema_version": LOCK_SCHEMA_VERSION,
        "lock_id": lock_id,
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "routing_policy_version": ROUTING_POLICY_VERSION,
        "report_mode": _clean(report_mode),
        "event_name": _clean(event_name),
        "decoration_fulfillment": _clean(decoration_fulfillment),
        "source_original_path": str(source_path),
        "product_master_original_path": str(master_path),
        "override_original_path": str(override_source),
        "files": {
            "source": {"name": SOURCE_NAME, "sha256": sha256_file(locked_source), "size": locked_source.stat().st_size},
            "product_master": {"name": MASTER_NAME, "sha256": sha256_file(locked_master), "size": locked_master.stat().st_size},
            "never_outsource_overrides": {"name": OVERRIDES_NAME, "sha256": sha256_file(locked_overrides), "size": locked_overrides.stat().st_size},
            "source_ledger": {"name": LEDGER_NAME, "sha256": sha256_file(ledger_path), "size": ledger_path.stat().st_size},
        },
        "source_line_count": int(len(ledger)),
        "source_quantity": int(ledger["Original Quantity"].sum()) if not ledger.empty else 0,
    }
    manifest_path.write_text(json.dumps(manifest, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    return {
        "lock_id": lock_id,
        "lock_dir": lock_dir,
        "manifest_path": manifest_path,
        "manifest_sha256": sha256_file(manifest_path),
        "source_line_count": manifest["source_line_count"],
        "source_quantity": manifest["source_quantity"],
    }


def _resolve_manifest(workbook_path: Path) -> tuple[Path, str, str]:
    lock_id = load_system_info_value(workbook_path, "Packet Lock ID")
    saved_path = load_system_info_value(workbook_path, "Packet Lock Manifest Path")
    saved_hash = load_system_info_value(workbook_path, "Packet Lock Manifest SHA256")
    candidates: list[Path] = []
    if saved_path:
        candidates.append(Path(saved_path).expanduser())
    candidates.append(Path(workbook_path).parent / "Packet Lock" / MANIFEST_NAME)
    if lock_id:
        candidates.append(data_dir() / LOCK_ROOT_NAME / lock_id / MANIFEST_NAME)
    for candidate in candidates:
        if candidate.is_file():
            return candidate.resolve(), lock_id, saved_hash
    raise RuntimeError(
        "Protected packet lock is missing. No PDFs were created.\n\n"
        "Regenerate Purchase Review from the current order CSV and live Product Master."
    )


def _manifest_file(manifest_path: Path, manifest: Mapping[str, Any], key: str) -> Path:
    info = manifest.get("files", {}).get(key, {}) if isinstance(manifest.get("files", {}), dict) else {}
    name = str(info.get("name", "") or "").strip()
    if not name:
        raise RuntimeError(f"Protected packet manifest is missing its {key} entry.")
    return manifest_path.parent / name


def _verify_file_entry(manifest_path: Path, manifest: Mapping[str, Any], key: str) -> Path:
    path = _manifest_file(manifest_path, manifest, key)
    info = manifest.get("files", {}).get(key, {})
    expected_hash = str(info.get("sha256", "") or "").strip()
    if not path.is_file():
        raise RuntimeError(f"Protected packet file is missing: {path}")
    actual_hash = sha256_file(path)
    if not expected_hash or actual_hash != expected_hash:
        raise RuntimeError(f"Protected packet file failed its hash check: {path.name}")
    return path


def _record_map(records: Iterable[Mapping[str, Any]]) -> dict[str, Mapping[str, Any]]:
    result: dict[str, Mapping[str, Any]] = {}
    duplicates: list[str] = []
    for row in records:
        source_id = _clean(row.get("Source ID", ""))
        if not source_id:
            continue
        if source_id in result:
            duplicates.append(source_id)
        result[source_id] = row
    if duplicates:
        raise RuntimeError(f"Purchase Review contains duplicate Source IDs: {', '.join(duplicates[:8])}")
    return result


def _same_text(left: object, right: object) -> bool:
    return _clean(left).casefold() == _clean(right).casefold()


def verify_packet_lock(
    workbook_path: Path,
    records: list[dict[str, object]],
    *,
    require_current_live_inputs: bool = True,
) -> dict[str, Any]:
    workbook_path = Path(workbook_path)
    manifest_path, saved_lock_id, saved_manifest_hash = _resolve_manifest(workbook_path)
    if saved_manifest_hash and sha256_file(manifest_path) != saved_manifest_hash:
        raise RuntimeError("Protected packet manifest no longer matches the Purchase Review workbook. No PDFs were created.")
    try:
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    except Exception as exc:
        raise RuntimeError(f"Protected packet manifest could not be read: {manifest_path}\n{exc}") from exc
    if str(manifest.get("schema_version", "")) != LOCK_SCHEMA_VERSION:
        raise RuntimeError("Protected packet lock uses an unsupported schema. Regenerate Purchase Review.")
    lock_id = str(manifest.get("lock_id", "") or "")
    if saved_lock_id and lock_id != saved_lock_id:
        raise RuntimeError("Protected packet lock ID does not match the Purchase Review workbook.")

    source_path = _verify_file_entry(manifest_path, manifest, "source")
    master_path = _verify_file_entry(manifest_path, manifest, "product_master")
    overrides_path = _verify_file_entry(manifest_path, manifest, "never_outsource_overrides")
    ledger_path = _verify_file_entry(manifest_path, manifest, "source_ledger")

    if require_current_live_inputs:
        current_master = Path(live_product_master_path()).expanduser().resolve()
        current_overrides = _override_path(current_master)
        current_master_hash = str(product_master_signature(current_master).get("sha256", ""))
        locked_master_hash = str(manifest["files"]["product_master"]["sha256"])
        if not current_master_hash or current_master_hash != locked_master_hash:
            raise RuntimeError(
                "The live Product Master changed after this Purchase Review was created. No PDFs were created.\n\n"
                "Regenerate Purchase Review so every vendor and Never Outsource route is recalculated from the current live master."
            )
        current_override_hash = sha256_file(current_overrides) if current_overrides.is_file() else sha256(b"{}\n").hexdigest()
        locked_override_hash = str(manifest["files"]["never_outsource_overrides"]["sha256"])
        if current_override_hash != locked_override_hash:
            raise RuntimeError(
                "The Never Outsource override file changed after this Purchase Review was created. No PDFs were created.\n\n"
                "Regenerate Purchase Review before creating purchase orders."
            )

    if _clean(load_report_mode(workbook_path)).casefold() != _clean(manifest.get("report_mode", "")).casefold():
        raise RuntimeError("Purchase Order mode does not match the protected packet lock.")
    if _clean(load_decoration_fulfillment(workbook_path)).casefold() != _clean(manifest.get("decoration_fulfillment", "")).casefold():
        raise RuntimeError("Decoration workflow does not match the protected packet lock.")

    ledger = pd.read_csv(ledger_path, dtype=str).fillna("")
    ledger_map = {str(row["Source ID"]).strip(): row for _, row in ledger.iterrows() if str(row.get("Source ID", "")).strip()}
    workbook_map = _record_map(records)
    missing = sorted(set(ledger_map) - set(workbook_map))
    unexpected = sorted(set(workbook_map) - set(ledger_map))
    errors: list[str] = []
    if missing:
        errors.append(f"{len(missing)} locked Shopify source line(s) are missing from Purchase Review.")
    if unexpected:
        errors.append(f"{len(unexpected)} Purchase Review Source ID(s) are not in the locked order source.")

    for source_id in sorted(set(ledger_map) & set(workbook_map)):
        locked = ledger_map[source_id]
        current = workbook_map[source_id]
        for column in IMMUTABLE_COLUMNS:
            if column == "Original Quantity":
                if _canonical_quantity(locked.get(column, 0)) != _canonical_quantity(current.get(column, 0)):
                    errors.append(f"{source_id}: original quantity changed.")
            elif not _same_text(locked.get(column, ""), current.get(column, "")):
                errors.append(f"{source_id}: immutable field changed: {column}.")

    master = load_extended_master(master_path)
    overrides = _load_overrides(overrides_path)
    never_outsource_violations: list[str] = []
    for source_id, row in workbook_map.items():
        if _clean(row.get("Include", "")).casefold() not in {"yes", "y", "true", "1", "include"}:
            continue
        style = _clean(row.get("Product #", ""))
        product = _clean(row.get("Description", ""))
        color = _clean(row.get("Garment Color", row.get("Color", "")))
        vendor = _clean(row.get("Purchase Vendor", ""))
        try:
            resolution = resolve_product(style, product, color, master, master_prepared=True)
            locked_never_outsource = bool(resolution.matched and resolution.never_outsource)
        except Exception:
            locked_never_outsource = False
        mandatory = locked_never_outsource or _style_override(overrides, style, product) or vendor_never_outsource(vendor)
        current_flag = _clean(row.get("Do Not Outsource", "")).casefold() in {"yes", "y", "true", "1"}
        if mandatory and not current_flag:
            never_outsource_violations.append(f"{source_id}: {style or product or 'unknown product'}")
    if never_outsource_violations:
        errors.append(
            f"{len(never_outsource_violations)} line(s) violate the locked Never Outsource routing, including "
            + ", ".join(never_outsource_violations[:6])
        )

    if errors:
        preview = "\n".join(f"- {item}" for item in errors[:14])
        if len(errors) > 14:
            preview += f"\n- Plus {len(errors) - 14} additional error(s)."
        raise RuntimeError("Protected packet verification failed. No PDFs were created.\n\n" + preview)

    return {
        "lock_id": lock_id,
        "manifest_path": manifest_path,
        "manifest_sha256": sha256_file(manifest_path),
        "source_path": source_path,
        "master_path": master_path,
        "overrides_path": overrides_path,
        "ledger_path": ledger_path,
        "source_lines": len(ledger_map),
        "source_quantity": sum(_canonical_quantity(row.get("Original Quantity", 0)) for row in ledger_map.values()),
        "never_outsource_violations": 0,
        "require_current_live_inputs": require_current_live_inputs,
    }


def copy_lock_bundle(manifest_path: Path, destination: Path) -> Path:
    """Copy the verified lock bundle into a final report or archive folder."""
    manifest_path = Path(manifest_path)
    destination = Path(destination) / "Packet Lock"
    destination.mkdir(parents=True, exist_ok=True)
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    for key in ("source", "product_master", "never_outsource_overrides", "source_ledger"):
        source = _verify_file_entry(manifest_path, manifest, key)
        shutil.copy2(source, destination / source.name)
    shutil.copy2(manifest_path, destination / MANIFEST_NAME)
    return destination / MANIFEST_NAME
