from __future__ import annotations

from pathlib import Path
import json
import hashlib
import os
import subprocess
import shutil
import sys
import time

import customtkinter as ctk
import pandas as pd
from openpyxl import load_workbook
from PIL import Image as PILImage, ImageDraw, ImageFont
from tkinter import filedialog, messagebox, simpledialog

from modules.shopify_parser import normalize_order_export, parse_shopify_orders
from modules.catalog_manager import (
    backup_file,
    clean_product_master,
    restore_product_master,
    restore_seed_product_master,
    save_product_candidate_list,
)
from modules.purchase_order_generator import safe_filename
from modules.paths import (
    data_dir,
    legacy_data_dir,
    legacy_product_master_path,
    product_master_path,
    product_master_update_marker_path,
    reports_dir,
    v12_data_dir,
)
from modules.master_sync import live_product_master_path, product_master_signature
from modules.report_modes import GENERAL_SALES_PERIOD, UNIFORM_SIZING_EVENT
from modules.decoration_fulfillment import (
    STANDARD_ORCHID_WORKFLOW, ENTIRE_ORDER_OUTSOURCED, normalize_decoration_fulfillment,
)
from modules.purchase_rules import apply_purchase_rule_defaults, normalize_bool, row_rules
from modules.routing_rules import enforce_permanent_vendor_overrides
from modules.product_intelligence import apply_product_intelligence
from modules.blank_garment_rules import BLANK_DECORATION_LABEL, is_blank_decoration
from modules.decoration_locations import (
    DECORATION_LOCATIONS,
    STANDARD_DECORATION_LOCATIONS,
    LEFT_CHEST,
    OTHER_CUSTOM,
    NOT_APPLICABLE_IN_HOUSE,
    NOT_APPLICABLE_NO_DECORATION,
    default_decoration_location,
    infer_decoration_location_from_note,
    location_choice_and_custom,
    normalize_decoration_location,
    resolve_location_choice,
)
from modules.event_archive import archive_completed_event
from modules.employee_totals import save_employee_discounts
from modules.xlsx_reader import load_review_lines
from modules.note_rules import decoration_note_requires_review, decoration_instruction_recommendation
from modules.internal_services import (
    HEMMING_ALTERATION_LABEL, SEW_ON_PATCH_LABEL,
    is_in_house_decoration, is_in_house_service_product,
)
from modules.product_candidate_sync import sync_review_product_candidates
from modules.scroll_support import install_native_scroll_support, register_scrollable
from modules.mission_control import (
    canonical_report_type,
    load_mission_control_snapshot,
    line_block_reasons,
    load_po_overrides,
    save_po_overrides,
)


PROJECT = Path(__file__).resolve().parent


def resource_path(*parts: str) -> Path:
    base = Path(getattr(sys, "_MEIPASS", PROJECT))
    return base.joinpath(*parts)

DATA = data_dir()
MASTER = live_product_master_path()  # Display only; actions resolve the live path again.
REPORTS = reports_dir()
PRODUCT_MASTER_PID_FILE = DATA / "product_master.pid"
PRODUCT_MASTER_REQUEST_FILE = DATA / "product_master_open_request.json"
REVIEW_STATE_FILE = DATA / "current_review_state.json"
REGENERATE_HELPER = DATA / "Regenerate Purchase Review.command"
ACTIVE_PACKET_RESET_FILE = DATA / "active_packet_reset.json"
UI_PREFERENCES_FILE = DATA / "ui_preferences.json"

PURPLE = "#5A18C9"
PURPLE_DARK = "#17122F"
PURPLE_LIGHT = "#F2EAFD"
BORDER = "#E0D8EB"
SHADOW = "#E9E4F0"
TEXT = "#14111F"
MUTED = "#6A6375"
BG = "#F3F0F7"
WHITE = "#FFFFFF"
DASHBOARD_WATERMARK_OPACITY = 0.09
SUCCESS = "#16A34A"
WARNING = "#E56A00"
DANGER = "#9C0006"

DO_NOT_OUTSOURCE_DECISION = "Do Not Outsource — Ship to Orchid"
INSTRUCTION_DECISION_OPTIONS = [
    "Choose an instruction decision",
    "Follow Note as Written",
    "Keep Product Master Default",
    "No Decoration",
    "Embroidery",
    "Screen Print",
    "Sew On Patch",
    "Hemming / Alteration",
]
OUTSOURCE_INSTRUCTION_TERMS = (
    "do not outsource", "don't outsource", "ship to orchid", "send to orchid",
)


def _clean(value: object) -> str:
    return " ".join(str(value or "").split())


def _file_stamp(path: Path | None) -> tuple[str, int, int]:
    if not path:
        return ("", 0, 0)
    candidate = Path(path)
    try:
        stat = candidate.stat()
        return (str(candidate.resolve()), int(stat.st_mtime_ns), int(stat.st_size))
    except (FileNotFoundError, OSError):
        return (str(candidate), 0, 0)


def should_show_employee_totals(
    snapshot: dict, workbook_exists: bool, current_mode: str = "", imported_line_count: int = 0
) -> bool:
    """Employee Totals is an event-only sidebar feature."""
    report_mode = str(snapshot.get("report_mode", "") or current_mode or "")
    line_count = int(snapshot.get("line_count", 0) or imported_line_count or 0)
    return bool(workbook_exists and line_count > 0 and report_mode == UNIFORM_SIZING_EVENT)


_PRODUCT_MASTER_HEALTH_CACHE_KEY: tuple[str, int, int] | None = None
_PRODUCT_MASTER_HEALTH_CACHE_VALUE: dict[str, int] | None = None


def _product_master_health() -> dict[str, int]:
    global _PRODUCT_MASTER_HEALTH_CACHE_KEY, _PRODUCT_MASTER_HEALTH_CACHE_VALUE
    result = {"records": 0, "styles": 0, "complete": 0, "incomplete": 0, "percent": 0}
    master_path = live_product_master_path()
    cache_key = _file_stamp(master_path)
    if cache_key == _PRODUCT_MASTER_HEALTH_CACHE_KEY and _PRODUCT_MASTER_HEALTH_CACHE_VALUE is not None:
        return dict(_PRODUCT_MASTER_HEALTH_CACHE_VALUE)
    if not master_path.exists():
        return result
    try:
        frame = pd.read_csv(master_path, dtype=str).fillna("")
        frame = apply_product_intelligence(apply_purchase_rule_defaults(frame))
    except Exception:
        return result
    result["records"] = len(frame)
    if frame.empty:
        return result

    def style_key(row: pd.Series) -> str:
        style = _clean(row.get("Style Number", "")).casefold()
        name = _clean(row.get("Product Name", "")).casefold()
        return f"style:{style}" if style else f"product:{name}"

    complete = 0
    groups = list(frame.groupby(frame.apply(style_key, axis=1), sort=False))
    for _, rows in groups:
        first = rows.iloc[0]
        rules = row_rules(first)
        requires_color = normalize_bool(rules["Requires Color"], True)
        requires_decoration = normalize_bool(rules["Requires Decoration"], True)
        issues = []
        if "Setup Required" in rows.columns and rows["Setup Required"].astype(str).str.strip().str.casefold().isin({"yes", "y", "true", "1"}).any():
            issues.append("setup review")
        if not _clean(first.get("Product Name", "")):
            issues.append("name")
        if not _clean(first.get("Style Number", "")):
            issues.append("style")
        if not _clean(rules.get("Product Category", "")):
            issues.append("category")
        service_only = is_in_house_service_product(
            first.get("Product Name", ""), first.get("Original Line Item", ""), first.get("Decoration Type", "")
        )
        if not service_only and rows.get("Vendor", pd.Series(dtype=str)).astype(str).str.strip().eq("").any():
            issues.append("vendor")
        if requires_color and rows.get("Garment Color", pd.Series(dtype=str)).astype(str).str.strip().eq("").any():
            issues.append("color")
        if rows.get("Decoration Type", pd.Series(dtype=str)).astype(str).str.strip().eq("").any():
            issues.append("decoration")
        if requires_decoration and "Decoration Location" in rows.columns:
            locations = rows["Decoration Location"].astype(str).str.strip()
            if locations.eq("").any() or locations.eq(OTHER_CUSTOM).any():
                issues.append("decoration location")
        if requires_decoration and "Decoration Color" in rows.columns:
            decorated = rows[~rows["Decoration Type"].map(lambda value: is_blank_decoration(value) or is_in_house_decoration(value))]
            if not decorated.empty and decorated["Decoration Color"].astype(str).str.strip().eq("").any():
                issues.append("decoration color")
        if not issues:
            complete += 1

    styles = len(groups)
    result.update({
        "styles": styles,
        "complete": complete,
        "incomplete": max(styles - complete, 0),
        "percent": round((complete / styles) * 100) if styles else 0,
    })
    _PRODUCT_MASTER_HEALTH_CACHE_KEY = cache_key
    _PRODUCT_MASTER_HEALTH_CACHE_VALUE = dict(result)
    return result


class OrchidPurchaseManager(ctk.CTk):
    PAGE_TITLES = {
        "dashboard": ("Dashboard", "Move the current purchase packet through the four purchasing stages."),
        "import": ("Current Event", "Import, replace, regenerate, archive, or discard the active purchase packet."),
        "master": ("Product Master", "Search and maintain permanent vendor, category, color, and decoration rules."),
        "review": ("Purchase Review", "Complete only the order-specific decisions that remain after Product Master setup."),
        "purchase": ("Purchase Orders", "Create final vendor purchase orders from the completed review."),
        "employees": ("Employee Totals", "View employee order totals, enter discounts, and save final event totals."),
        "archive": ("Event Archive", "Open completed events with their order export, review workbook, and final PDFs."),
        "settings": ("Settings", "Diagnostics, catalog maintenance, backups, and storage locations."),
    }

    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("light")
        self.title("Orchid Purchase Manager Professional 4.8.33")
        self.geometry("1400x900")
        self.minsize(1180, 760)
        self.configure(fg_color=BG)
        self.after(80, self._maximize_main_window)

        self.selected_csv: Path | None = None
        self.current_event_name = ""
        self.current_mode = ""
        self.current_decoration_fulfillment = STANDARD_ORCHID_WORKFLOW
        self.last_review_workbook: Path | None = None
        self.last_purchase_order_dir: Path | None = None
        self.imported_line_count = 0
        self.review_needed: int | None = None
        self.active_page = "dashboard"
        self.nav_buttons: dict[str, ctk.CTkButton] = {}
        self.pages: dict[str, ctk.CTkFrame] = {}
        self.po_entry_vars: dict[str, tuple[str, str, ctk.StringVar]] = {}
        self.employee_discount_vars: dict[int, ctk.StringVar] = {}
        self.employee_total_labels: dict[int, ctk.CTkLabel] = {}
        self.employee_records_by_row: dict[int, dict] = {}
        self.last_snapshot: dict = {}
        self.product_master_process: subprocess.Popen | None = None
        self.product_master_launching = False
        self._product_master_was_open = False
        self._candidate_sync_workbook = ""
        self._regenerating_review = False
        self._auto_regenerate_after_id = None
        self.product_master_update_marker = product_master_update_marker_path()
        self._last_product_master_update = (
            self.product_master_update_marker.stat().st_mtime
            if self.product_master_update_marker.exists() else 0.0
        )
        self._launch_choice_shown = False
        self._suppress_empty_import_prompt = False
        self.success_sound_enabled = True
        self._success_animation_running = False
        self._success_animation_skip_requested = False
        self._reduced_motion_override = None
        ui_preferences = self._load_ui_preferences()
        self.dashboard_view = str(ui_preferences.get("dashboard_view", "guided") or "guided")
        if self.dashboard_view not in {"guided", "management"}:
            self.dashboard_view = "guided"
        self.interaction_animations_enabled = bool(ui_preferences.get("interaction_animations_enabled", True))
        self._logo_pulse_running = False
        self._logo_pulse_after_ids = []
        self._outsourced_logo_dialog_pending = False
        self._mission_snapshot_cache_key = None
        self._mission_snapshot_cache: dict = {}
        self._dashboard_refresh_key = None
        self._normalized_import_cache = None
        self._parsed_import_cache = None
        self._import_cache_key = None

        try:
            enforce_permanent_vendor_overrides(live_product_master_path())
        except Exception:
            pass
        self.discover_existing_work()
        self.build_ui()
        self._install_global_click_pulse()
        install_native_scroll_support(
            self,
            [
                frame for frame in (
                    getattr(self, "mc_work_scroll", None),
                    getattr(self, "mc_employee_scroll", None),
                    getattr(self, "settings_scroll", None),
                    getattr(self, "review_decision_frame", None),
                ) if frame is not None
            ],
        )
        self.refresh_dashboard()
        self.show_page("dashboard")
        self.after(650, self._prompt_resume_or_start_new)
        self.after(950, self._prompt_for_csv_if_empty)
        self.after(1100, self._poll_product_master_updates)
        self.after(1300, self._upgrade_active_review_for_decoration_notes)

    def _load_ui_preferences(self) -> dict:
        try:
            if UI_PREFERENCES_FILE.exists():
                payload = json.loads(UI_PREFERENCES_FILE.read_text(encoding="utf-8"))
                return payload if isinstance(payload, dict) else {}
        except Exception:
            pass
        return {}

    def _save_ui_preferences(self) -> None:
        try:
            UI_PREFERENCES_FILE.parent.mkdir(parents=True, exist_ok=True)
            UI_PREFERENCES_FILE.write_text(
                json.dumps({
                    "dashboard_view": getattr(self, "dashboard_view", "guided"),
                    "interaction_animations_enabled": bool(getattr(self, "interaction_animations_enabled", True)),
                }, indent=2, sort_keys=True),
                encoding="utf-8",
            )
        except Exception:
            pass

    def _build_sidebar_logo_frames(self, source: PILImage.Image) -> list[ctk.CTkImage]:
        """Create a fixed-size official logo plus a graceful purple radial pulse."""
        canvas_size = 108
        flower_size = 70
        flower = source.convert("RGBA").resize((flower_size, flower_size), PILImage.Resampling.LANCZOS)
        frames: list[ctk.CTkImage] = []
        purple_shades = [
            (93, 33, 201), (117, 55, 218), (142, 82, 229), (172, 121, 237),
        ]
        for frame_index in range(11):
            canvas = PILImage.new("RGBA", (canvas_size, canvas_size), (0, 0, 0, 0))
            if frame_index:
                draw = ImageDraw.Draw(canvas, "RGBA")
                progress = frame_index / 10.0
                for ring_index, color in enumerate(purple_shades):
                    phase = max(0.0, min(1.0, progress - ring_index * 0.10))
                    radius = 24 + phase * 28 + ring_index * 4
                    alpha = int(86 * (1.0 - progress * 0.72) * (1.0 - ring_index * 0.13))
                    box = (
                        canvas_size / 2 - radius,
                        canvas_size / 2 - radius,
                        canvas_size / 2 + radius,
                        canvas_size / 2 + radius,
                    )
                    draw.ellipse(box, outline=(*color, max(alpha, 0)), width=3)
            x = (canvas_size - flower_size) // 2
            y = (canvas_size - flower_size) // 2
            canvas.alpha_composite(flower, (x, y))
            frames.append(ctk.CTkImage(light_image=canvas, dark_image=canvas, size=(canvas_size, canvas_size)))
        return frames

    def _initialize_sidebar_logo(self, brand: ctk.CTkFrame) -> None:
        self.sidebar_brand_logo_label = None
        self.sidebar_brand_flower_frames = []
        white_flower_path = resource_path("assets", "orchid_flower_white.png")
        if not white_flower_path.exists():
            return
        source = PILImage.open(white_flower_path).convert("RGBA")
        self.sidebar_brand_flower_frames = self._build_sidebar_logo_frames(source)
        self.sidebar_brand_flower = self.sidebar_brand_flower_frames[0]
        self.sidebar_brand_logo_label = ctk.CTkLabel(
            brand, text="", image=self.sidebar_brand_flower,
            width=108, height=108, fg_color="transparent",
        )
        self.sidebar_brand_logo_label.grid(row=0, column=0, pady=(12, 0))

    def _initialize_header_wordmark(self, parent: ctk.CTkFrame) -> None:
        """Place the official Orchid wordmark directly on the purple header."""
        self.sidebar_brand_logo_label = None
        self.sidebar_brand_flower_frames = []
        logo_path = resource_path("assets", "orchid_logo.png")
        if not logo_path.exists():
            return
        source = PILImage.open(logo_path).convert("RGBA")
        # The lower portion of the source contains the retail tagline. The app
        # uses a dedicated PURCHASE MANAGER subtitle instead.
        wordmark = source.crop((0, 0, source.width, int(source.height * 0.81)))
        pixels = wordmark.load()
        for y in range(wordmark.height):
            for x in range(wordmark.width):
                red, green, blue, alpha = pixels[x, y]
                if not alpha:
                    continue
                is_orchid_purple = blue > red and red > green * 1.35
                pixels[x, y] = (red, green, blue, alpha) if is_orchid_purple else (248, 246, 252, alpha)
        bbox = wordmark.getchannel("A").getbbox()
        if bbox:
            wordmark = wordmark.crop(bbox)

        base_width = 560
        base_height = max(1, round(base_width * wordmark.height / wordmark.width))
        frames = []
        for index in range(11):
            # A very small scale bloom preserves the click animation without
            # disturbing the approved header proportions at rest.
            bloom = round(8 * (index / 10.0)) if index else 0
            size = (base_width + bloom, base_height + round(bloom * base_height / base_width))
            frames.append(ctk.CTkImage(light_image=wordmark, dark_image=wordmark, size=size))
        self.sidebar_brand_flower_frames = frames
        self.sidebar_brand_flower = frames[0]
        self.sidebar_brand_logo_label = ctk.CTkLabel(
            parent, text="", image=self.sidebar_brand_flower,
            width=590, height=142, fg_color="transparent",
        )
        self.sidebar_brand_logo_label.place(relx=0.5, rely=0.40, anchor="center")
        ctk.CTkLabel(
            parent, text="PURCHASE MANAGER", text_color="#F8F6FC",
            font=ctk.CTkFont(size=22, weight="bold"),
        ).place(relx=0.5, rely=0.80, anchor="center")

    def _finish_logo_pulse(self) -> None:
        self._logo_pulse_running = False
        for after_id in list(getattr(self, "_logo_pulse_after_ids", [])):
            try:
                self.after_cancel(after_id)
            except Exception:
                pass
        self._logo_pulse_after_ids = []
        label = getattr(self, "sidebar_brand_logo_label", None)
        frames = getattr(self, "sidebar_brand_flower_frames", [])
        if label is not None and frames:
            try:
                label.configure(image=frames[0])
            except Exception:
                pass

    def _play_logo_pulse(self) -> None:
        if not bool(getattr(self, "interaction_animations_enabled", True)):
            return
        if self._prefers_reduced_motion():
            return
        if bool(getattr(self, "_logo_pulse_running", False)):
            self._finish_logo_pulse()
        label = getattr(self, "sidebar_brand_logo_label", None)
        frames = getattr(self, "sidebar_brand_flower_frames", [])
        if label is None or len(frames) < 2:
            return
        self._logo_pulse_running = True
        sequence = list(range(1, len(frames))) + list(range(len(frames) - 2, 0, -1)) + [0]
        self._logo_pulse_after_ids = []
        for position, frame_index in enumerate(sequence):
            def show(index=frame_index, is_last=(position == len(sequence) - 1)):
                if not getattr(self, "_logo_pulse_running", False):
                    return
                try:
                    label.configure(image=frames[index])
                except Exception:
                    self._finish_logo_pulse()
                    return
                if is_last:
                    self._logo_pulse_running = False
                    self._logo_pulse_after_ids = []
            self._logo_pulse_after_ids.append(self.after(position * 36, show))

    def _install_global_click_pulse(self) -> None:
        """Pulse the official Orchid flower for every in-app mouse click."""
        try:
            self.bind_all("<Button-1>", self._on_global_click_pulse, add="+")
        except Exception:
            pass

    def _on_global_click_pulse(self, _event=None) -> None:
        if not bool(getattr(self, "interaction_animations_enabled", True)):
            return
        try:
            self.after_idle(self._play_logo_pulse)
        except Exception:
            self._play_logo_pulse()

    def _set_interaction_animations_enabled(self) -> None:
        enabled = bool(self.interaction_animation_var.get()) if hasattr(self, "interaction_animation_var") else True
        self.interaction_animations_enabled = enabled
        if not enabled:
            self._finish_logo_pulse()
        self._save_ui_preferences()

    def _maximize_main_window(self):
        """Open maximized while remaining a normal macOS window."""
        try:
            self.state("zoomed")
            return
        except Exception:
            pass
        try:
            self.update_idletasks()
            width = max(self.winfo_screenwidth(), self.winfo_width())
            height = max(self.winfo_screenheight() - 70, self.winfo_height())
            self.geometry(f"{width}x{height}+0+0")
        except Exception:
            pass

    def discover_existing_work(self):
        """Find the most recent saved work unless the user intentionally started a new packet."""
        reset_time = 0.0
        try:
            if ACTIVE_PACKET_RESET_FILE.exists():
                reset_time = float(json.loads(ACTIVE_PACKET_RESET_FILE.read_text(encoding="utf-8")).get("reset_time", 0.0))
        except Exception:
            reset_time = 0.0

        review_files = []
        for folder in (REPORTS / "Review Workbooks", REPORTS / "review_workbooks"):
            if folder.exists():
                review_files.extend(
                    item for item in folder.glob("*.xlsx")
                    if item.stat().st_mtime > reset_time
                )
        if review_files:
            self.last_review_workbook = max(review_files, key=lambda item: item.stat().st_mtime)

        purchase_root = REPORTS / "Purchase Orders"
        if purchase_root.exists():
            folders = [
                item for item in purchase_root.iterdir()
                if item.is_dir() and item.stat().st_mtime > reset_time
            ]
            pdfs = [item for item in purchase_root.glob("*.pdf") if item.stat().st_mtime > reset_time]
            if folders:
                self.last_purchase_order_dir = max(folders, key=lambda item: item.stat().st_mtime)
            elif pdfs:
                self.last_purchase_order_dir = purchase_root

        # A completed packet may show its success screen during the session in which
        # PDFs are generated, but it must not reopen as the active packet next time.
        if (
            self.last_review_workbook
            and self.last_review_workbook.exists()
            and self.last_purchase_order_dir
            and self.last_purchase_order_dir.exists()
            and self.last_purchase_order_dir.stat().st_mtime >= self.last_review_workbook.stat().st_mtime
        ):
            try:
                ACTIVE_PACKET_RESET_FILE.parent.mkdir(parents=True, exist_ok=True)
                ACTIVE_PACKET_RESET_FILE.write_text(
                    json.dumps({"reset_time": time.time()}, indent=2),
                    encoding="utf-8",
                )
                if REVIEW_STATE_FILE.exists():
                    REVIEW_STATE_FILE.unlink()
            except Exception:
                pass
            self.last_review_workbook = None
            self.last_purchase_order_dir = None

        # Restore the saved source CSV and event identity for an unfinished
        # review so Current Event and the launch choice describe the packet
        # correctly after reopening a newer app version.
        if self.last_review_workbook and self.last_review_workbook.exists() and REVIEW_STATE_FILE.exists():
            try:
                saved_state = json.loads(REVIEW_STATE_FILE.read_text(encoding="utf-8"))
                source_csv = Path(saved_state.get("source_csv", "")).expanduser()
                if source_csv.exists():
                    self.selected_csv = source_csv
                self.current_event_name = _clean(saved_state.get("event_name", ""))
                self.current_mode = _clean(saved_state.get("report_mode", ""))
                self.current_decoration_fulfillment = normalize_decoration_fulfillment(saved_state.get("decoration_fulfillment", STANDARD_ORCHID_WORKFLOW))
            except Exception:
                pass

    def _active_review_needs_decoration_note_upgrade(self) -> bool:
        if (
            self.current_mode != UNIFORM_SIZING_EVENT
            and normalize_decoration_fulfillment(self.current_decoration_fulfillment) != ENTIRE_ORDER_OUTSOURCED
        ):
            return False
        path = self.last_review_workbook
        if not path or not path.exists() or not REVIEW_STATE_FILE.exists():
            return False
        try:
            workbook = load_workbook(path, read_only=True, data_only=False)
            if "All PO Lines" not in workbook.sheetnames:
                workbook.close()
                return False
            sheet = workbook["All PO Lines"]
            header_row = None
            headers: dict[str, int] = {}
            for row_number in range(1, min(sheet.max_row, 20) + 1):
                values = [str(sheet.cell(row_number, col).value or "").strip() for col in range(1, sheet.max_column + 1)]
                if "Line ID" in values:
                    header_row = row_number
                    headers = {value: idx + 1 for idx, value in enumerate(values) if value}
                    break
            required_headers = {"Decoration Decision", "Mandatory Note Review", "Decoration Location", "Decoration Placement Instructions"}
            if not header_row or not required_headers.issubset(headers):
                workbook.close()
                return True

            # Professional 4.7.4 also upgrades active workbooks to the permanent/order-specific decoration-location schema.
            # Professional 4.4.6 broadened mandatory review beyond EMB wording to
            # placement, logo, screen-print, personalization, and alteration notes.
            # Rebuild only when an existing active row would now require review.
            instruction_col = headers.get("Purchase Instructions") or headers.get("Shopify Order Notes")
            mandatory_col = headers.get("Mandatory Note Review")
            if instruction_col and mandatory_col:
                for row_number in range(header_row + 1, sheet.max_row + 1):
                    note_text = _clean(sheet.cell(row_number, instruction_col).value)
                    mandatory = _clean(sheet.cell(row_number, mandatory_col).value).casefold() in {"yes", "y", "true", "1"}
                    if note_text and decoration_note_requires_review(note_text) and not mandatory:
                        workbook.close()
                        return True
            workbook.close()
            return False
        except Exception:
            return False

    def _upgrade_active_review_for_decoration_notes(self):
        """Rebuild an older active review so no customer instruction can be missed."""
        if self._regenerating_review or not self._active_review_needs_decoration_note_upgrade():
            return
        self.regenerate_current_review(silent=True)
        try:
            self.dashboard_status.configure(
                text="Purchase Review upgraded for expanded decoration locations, faster review navigation, and success sound controls.",
                text_color=SUCCESS,
            )
        except Exception:
            pass

    def _prompt_for_csv_if_empty(self):
        """Open the branded Current Event start page when no packet exists."""
        if self._suppress_empty_import_prompt:
            self._suppress_empty_import_prompt = False
            return
        if self.selected_csv or self.last_review_workbook or self.last_purchase_order_dir:
            return
        self.show_page("import")

    def _prompt_resume_or_start_new(self):
        """Explain unfinished saved work instead of silently landing mid-workflow."""
        if self._launch_choice_shown:
            return
        has_active_packet = bool(
            self.selected_csv
            or (self.last_review_workbook and self.last_review_workbook.exists())
        )
        if not has_active_packet:
            return
        self._launch_choice_shown = True

        dialog = ctk.CTkToplevel(self)
        dialog.title("Unfinished Purchase Packet")
        dialog.transient(self)
        dialog.grab_set()
        dialog.resizable(False, False)
        dialog.configure(fg_color=BG)
        shell = ctk.CTkFrame(dialog, fg_color=WHITE, border_width=1, border_color="#D8CDE6", corner_radius=18)
        shell.pack(fill="both", expand=True, padx=18, pady=18)
        ctk.CTkLabel(
            shell, text="An unfinished purchase packet was found", text_color=PURPLE_DARK,
            font=ctk.CTkFont(size=24, weight="bold"), justify="center", wraplength=500,
        ).pack(padx=28, pady=(26, 8))
        packet_name = self.current_event_name or (
            self.last_review_workbook.stem if self.last_review_workbook else "the saved event"
        )
        ctk.CTkLabel(
            shell,
            text=f"Would you like to resume {packet_name}, or clear the active packet and begin a new event?\n\nProduct Master and archived reports will not be deleted.",
            text_color=MUTED, font=ctk.CTkFont(size=14), justify="center", wraplength=520,
        ).pack(padx=30, pady=(0, 22))
        actions = ctk.CTkFrame(shell, fg_color="transparent")
        actions.pack(padx=24, pady=(0, 26))

        def resume():
            dialog.grab_release()
            dialog.destroy()
            self.show_page("dashboard")

        def start_new():
            dialog.grab_release()
            dialog.destroy()
            self._suppress_empty_import_prompt = True
            self._clear_active_packet_state()
            self.refresh_current_event_page({})
            self.refresh_dashboard()
            self.show_page("dashboard")
            self.after_idle(lambda: (self.refresh_current_event_page({}), self.refresh_dashboard()))

        ctk.CTkButton(
            actions, text="Resume Current Event", command=resume,
            width=210, height=48, fg_color=PURPLE, hover_color=PURPLE_DARK,
            font=ctk.CTkFont(size=14, weight="bold"),
        ).pack(side="left", padx=7)
        ctk.CTkButton(
            actions, text="Start New Event", command=start_new,
            width=190, height=48, fg_color=PURPLE_LIGHT, hover_color="#E6D9F5",
            border_width=1, border_color="#BFA9DD", text_color=PURPLE_DARK,
            font=ctk.CTkFont(size=14, weight="bold"),
        ).pack(side="left", padx=7)
        dialog.protocol("WM_DELETE_WINDOW", resume)
        dialog.update_idletasks()
        width, height = 610, 310
        x = max(self.winfo_rootx() + (self.winfo_width() - width) // 2, 0)
        y = max(self.winfo_rooty() + (self.winfo_height() - height) // 2, 0)
        dialog.geometry(f"{width}x{height}+{x}+{y}")
        dialog.lift()
        dialog.focus_force()

    @staticmethod
    def _friendly_modified(path: Path | None) -> str:
        if not path or not path.exists():
            return ""
        from datetime import datetime
        return datetime.fromtimestamp(path.stat().st_mtime).strftime("%B %-d, %Y at %-I:%M %p")

    @staticmethod
    def _source_file_signature(path: Path | None) -> str:
        if not path or not Path(path).exists():
            return ""
        digest = hashlib.sha256()
        with Path(path).open("rb") as handle:
            for chunk in iter(lambda: handle.read(1024 * 1024), b""):
                digest.update(chunk)
        return digest.hexdigest()

    def _remember_import_frames(self, path: Path, normalized, parsed) -> None:
        self._import_cache_key = (_file_stamp(path), _file_stamp(live_product_master_path()))
        self._normalized_import_cache = normalized.copy(deep=True)
        self._parsed_import_cache = parsed.copy(deep=True)

    def _cached_import_frames(self, path: Path):
        key = (_file_stamp(path), _file_stamp(live_product_master_path()))
        if key != self._import_cache_key:
            return None, None
        normalized = (
            self._normalized_import_cache.copy(deep=True)
            if self._normalized_import_cache is not None else None
        )
        parsed = (
            self._parsed_import_cache.copy(deep=True)
            if self._parsed_import_cache is not None else None
        )
        return normalized, parsed

    def _load_review_state(self) -> dict:
        try:
            if REVIEW_STATE_FILE.exists():
                payload = json.loads(REVIEW_STATE_FILE.read_text(encoding="utf-8"))
                return payload if isinstance(payload, dict) else {}
        except Exception:
            pass
        return {}

    def _update_review_state(self, **updates) -> dict:
        state = self._load_review_state()
        state.update(updates)
        try:
            REVIEW_STATE_FILE.parent.mkdir(parents=True, exist_ok=True)
            REVIEW_STATE_FILE.write_text(json.dumps(state, indent=2), encoding="utf-8")
        except Exception:
            pass
        return state

    def _clear_saved_review_draft(self, issue: dict | None = None) -> None:
        state = self._load_review_state()
        if not state:
            return
        if issue:
            saved_line = _clean(state.get("draft_line_id", ""))
            saved_key = _clean(state.get("draft_decision_key", ""))
            issue_line = _clean(issue.get("line_id", ""))
            issue_key = _clean(issue.get("decision_key", ""))
            if saved_line and issue_line and saved_line != issue_line and (not saved_key or saved_key != issue_key):
                return
        changed = False
        for key in ("draft_line_id", "draft_decision_key", "draft_saved_at"):
            if key in state:
                state.pop(key, None)
                changed = True
        if changed:
            try:
                REVIEW_STATE_FILE.write_text(json.dumps(state, indent=2), encoding="utf-8")
            except Exception:
                pass

    def _saved_review_draft_issue(self, issues: list[dict]) -> dict | None:
        state = self._load_review_state()
        saved_line = _clean(state.get("draft_line_id", ""))
        saved_key = _clean(state.get("draft_decision_key", ""))
        if not saved_line and not saved_key:
            return None
        for issue in issues:
            line_id = _clean(issue.get("line_id", ""))
            decision_key = _clean(issue.get("decision_key", ""))
            if (saved_line and line_id == saved_line) or (saved_key and decision_key == saved_key):
                return issue
        self._clear_saved_review_draft()
        return None

    # ---------- shell ----------
    def build_ui(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.build_sidebar()
        self.build_main_shell()

    @staticmethod
    def _purple_orchid_art(source: PILImage.Image, size: tuple[int, int]) -> PILImage.Image:
        """Return the Orchid flower in a lavender-to-purple treatment with no green."""
        flower = source.convert("RGBA")
        bbox = flower.getchannel("A").getbbox()
        if bbox:
            flower = flower.crop(bbox)
        flower.thumbnail(size, PILImage.Resampling.LANCZOS)
        alpha = flower.getchannel("A")
        detail = flower.convert("L")
        tinted = PILImage.new("RGBA", flower.size, (0, 0, 0, 0))
        height = max(1, flower.height - 1)
        top_color = (202, 126, 247)
        bottom_color = (132, 52, 213)
        pixels = []
        for y in range(flower.height):
            ratio = y / height
            base_color = tuple(
                round(top_color[channel] * (1.0 - ratio) + bottom_color[channel] * ratio)
                for channel in range(3)
            )
            for x in range(flower.width):
                pixel_alpha = alpha.getpixel((x, y))
                source_luma = detail.getpixel((x, y))
                # Keep the original petal and center linework as darker purple
                # while ensuring the entire mark stays in the purple family.
                shade = 0.66 + 0.34 * min(1.0, source_luma / 92.0)
                color = tuple(round(channel * shade) for channel in base_color)
                pixels.append((*color, pixel_alpha))
        tinted.putdata(pixels)
        return tinted

    @staticmethod
    def _sidebar_spotlight_background(width: int, height: int) -> PILImage.Image:
        """Continue the header color family behind Current Event, then fade to navigation."""
        width = max(1, int(width))
        height = max(1, int(height))
        sample_width = min(160, width)
        sample_height = min(132, height)

        base = PILImage.new("RGBA", (sample_width, sample_height), (0, 0, 0, 255))
        base_draw = ImageDraw.Draw(base)
        top = (17, 8, 40)
        bottom = (16, 12, 37)
        for y in range(sample_height):
            ratio = y / max(1, sample_height - 1)
            color = tuple(
                round(top[channel] * (1.0 - ratio) + bottom[channel] * ratio)
                for channel in range(3)
            )
            base_draw.line((0, y, sample_width, y), fill=(*color, 255))

        mask = PILImage.new("L", (sample_width, sample_height), 0)
        mask_pixels = []
        for y in range(sample_height):
            y_ratio = y / max(1, sample_height - 1)
            normalized_y = (y_ratio - 0.26) / 0.42
            vertical_fade = pow(max(0.0, 1.0 - y_ratio), 1.65)
            for x in range(sample_width):
                normalized_x = (x / max(1, sample_width - 1) - 0.50) / 0.48
                distance = normalized_x * normalized_x + normalized_y * normalized_y
                alpha = 132 * pow(2.718281828, -distance * 1.55) * vertical_fade
                mask_pixels.append(round(alpha))
        mask.putdata(mask_pixels)

        spotlight = PILImage.new("RGBA", base.size, (62, 18, 112, 255))
        background = PILImage.composite(spotlight, base, mask)
        return background.resize((width, height), PILImage.Resampling.LANCZOS)

    def build_sidebar(self):
        sidebar = ctk.CTkFrame(self, width=270, corner_radius=0, fg_color="#100C25")
        sidebar.grid(row=0, column=0, sticky="nsew")
        sidebar.grid_propagate(False)
        sidebar.grid_rowconfigure(2, weight=1)

        brand = ctk.CTkFrame(sidebar, fg_color="transparent", corner_radius=0, height=160)
        brand.grid(row=0, column=0, sticky="ew")
        brand.grid_propagate(False)
        brand.grid_columnconfigure(0, weight=1)
        self.sidebar_event_card = brand
        sidebar_spotlight = self._sidebar_spotlight_background(540, 320)
        self.sidebar_spotlight_image = ctk.CTkImage(
            light_image=sidebar_spotlight, dark_image=sidebar_spotlight, size=(270, 160)
        )
        self.sidebar_spotlight_label = ctk.CTkLabel(
            brand, text="", image=self.sidebar_spotlight_image,
            width=270, height=160, fg_color="transparent",
        )
        self.sidebar_spotlight_label.place(x=0, y=0)
        # Compatibility handle retained for state-refresh code; the sidebar is
        # intentionally text-only so the header flower remains the sole brand mark.
        self.sidebar_event_icon = ctk.CTkLabel(
            brand, text="", width=1, height=1, fg_color="transparent",
        )
        ctk.CTkLabel(
            brand, text="CURRENT EVENT", text_color="#CDB7F1",
            font=ctk.CTkFont(size=10, weight="bold"),
        ).grid(row=0, column=0, pady=(30, 4))
        self.sidebar_event_title = ctk.CTkLabel(
            brand, text="No Active Purchase Packet", text_color=WHITE,
            font=ctk.CTkFont(size=15, weight="bold"), justify="center",
            wraplength=232,
        )
        self.sidebar_event_title.grid(row=1, column=0, padx=14)
        self.sidebar_event_meta = ctk.CTkLabel(
            brand, text=time.strftime("%b %d, %Y"), text_color="#CFC4E0",
            font=ctk.CTkFont(size=10),
        )
        self.sidebar_event_meta.grid(row=2, column=0, pady=(5, 0))

        nav = ctk.CTkFrame(sidebar, fg_color="transparent")
        nav.grid(row=1, column=0, rowspan=2, sticky="nsew", padx=14, pady=(4, 0))
        nav.grid_columnconfigure(0, weight=1)
        self.sidebar_page_title = ctk.CTkLabel(nav, text="", width=1, height=1)
        self.sidebar_page_title.grid_remove()

        self.sidebar_home_button = ctk.CTkButton(
            nav, text="⌂   Dashboard", command=lambda: self.show_page("dashboard"), anchor="w",
            height=48, corner_radius=12, fg_color="#6A24D4", hover_color="#7C38E0",
            text_color=WHITE, font=ctk.CTkFont(size=15, weight="bold"),
        )
        self.sidebar_home_button.grid(row=0, column=0, pady=(0, 6), sticky="ew")
        self.nav_buttons["dashboard"] = self.sidebar_home_button

        items = [
            ("import", "▣   Current Event"),
            ("master", "◇   Product Master"),
            ("review", "☑   Purchase Review"),
            ("purchase", "▤   Purchase Orders"),
            ("employees", "▦   Employee Totals"),
            ("job_logo", "◉   Logo"),
        ]
        for row, (page, label) in enumerate(items, start=1):
            button = ctk.CTkButton(
                nav, text=label, command=lambda p=page: self.show_page(p), anchor="w",
                height=47, corner_radius=12, fg_color="transparent", hover_color="#2B214A",
                text_color=WHITE, font=ctk.CTkFont(size=15),
            )
            button.grid(row=row, column=0, pady=3, sticky="ew")
            self.nav_buttons[page] = button
            if page == "employees":
                self.employee_nav_button = button
                button.grid_remove()
            if page == "job_logo":
                self.outsourced_logo_nav_button = button

        divider = ctk.CTkFrame(nav, height=1, fg_color="#40375A", corner_radius=0)
        divider.grid(row=7, column=0, sticky="ew", padx=10, pady=(16, 12))
        for row, page, label in [
            (8, "archive", "▦   Event Archive"),
            (9, "settings", "⚙   Settings & Tools"),
        ]:
            button = ctk.CTkButton(
                nav, text=label, command=lambda p=page: self.show_page(p), anchor="w",
                height=47, corner_radius=12, fg_color="transparent", hover_color="#2B214A",
                text_color=WHITE, font=ctk.CTkFont(size=15),
            )
            button.grid(row=row, column=0, pady=3, sticky="ew")
            self.nav_buttons[page] = button

        footer = ctk.CTkFrame(sidebar, fg_color="transparent")
        footer.grid(row=4, column=0, sticky="ew", padx=16, pady=(10, 18))
        ctk.CTkLabel(
            footer, text="v 4.8.33", text_color="#CFC4E0",
            font=ctk.CTkFont(size=12), justify="left", anchor="w",
        ).pack(anchor="w", padx=8, pady=(0, 10))
        ctk.CTkFrame(footer, height=1, fg_color="#40375A").pack(fill="x", padx=7, pady=(0, 11))
        ctk.CTkButton(
            footer, text="⇥  Exit", height=50, corner_radius=12,
            fg_color="transparent", hover_color="#2B214A", border_width=1,
            border_color="#5A4D78", text_color=WHITE, command=self.destroy,
            font=ctk.CTkFont(size=15, weight="bold"),
        ).pack(fill="x")

    def build_main_shell(self):
        main = ctk.CTkFrame(self, fg_color=BG, corner_radius=0)
        main.grid(row=0, column=1, sticky="nsew")
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(1, weight=1)

        header = ctk.CTkFrame(main, fg_color=BG, corner_radius=0, height=148)
        self.main_header = header
        header.grid(row=0, column=0, sticky="ew")
        header.grid_propagate(False)
        header.grid_columnconfigure(0, weight=1)

        wave_path = resource_path("assets", "purple_wave_header.png")
        self._wave_render_job = None
        self._wave_render_size = (0, 0)
        self._wave_source = None
        if wave_path.exists():
            # Keep the full-resolution source and render a Retina-sized copy for
            # the exact window width. This prevents blur and avoids cropping the
            # complete wave artwork on larger displays.
            self._wave_source = PILImage.open(wave_path).convert("RGBA")
            self.wave_header_label = ctk.CTkLabel(
                header, text="", fg_color="transparent"
            )
            self.wave_header_label.place(x=0, y=0, relwidth=1, relheight=1)
            header.bind("<Configure>", self._schedule_wave_header_resize, add="+")
            self.after(40, self._render_current_wave_header)
        else:
            ctk.CTkFrame(header, fg_color="#5A18C9").place(relx=0, rely=0, relwidth=1, relheight=1)

        self.shell_event_card = self.sidebar_event_card
        self.shell_event_title = self.sidebar_event_title
        self.shell_event_meta = self.sidebar_event_meta
        icon = self.sidebar_event_icon
        event = self.sidebar_event_card

        # Existing refresh code uses these names; all now point to the single shell card.
        self.dashboard_summary = event
        self.current_event_summary = event
        self.dashboard_summary_icon = icon
        self.current_event_summary_icon = icon
        self.dashboard_summary_title = self.shell_event_title
        self.current_event_summary_title = self.shell_event_title
        self.dashboard_date_label = self.shell_event_meta
        self.current_event_date_label = self.shell_event_meta

        # Compatibility labels retained for the existing state-refresh code.
        self.page_title = ctk.CTkLabel(header, text="", width=1, height=1)
        self.page_title.grid_remove()
        self.page_subtitle = ctk.CTkLabel(header, text="", width=1, height=1)
        self.page_subtitle.grid_remove()
        self.header_chip = ctk.CTkLabel(header, text="", width=1, height=1)
        self.header_chip.grid_remove()
        self.header_date = ctk.CTkLabel(header, text="", width=1, height=1)
        self.header_date.grid_remove()

        self.page_container = ctk.CTkFrame(main, fg_color=BG, corner_radius=0)
        self.page_container.grid(row=1, column=0, sticky="nsew")
        self.page_container.grid_rowconfigure(0, weight=1)
        self.page_container.grid_columnconfigure(0, weight=1)

        self.build_dashboard_page()
        self._page_builders = {
            "import": self.build_import_page,
            "master": self.build_master_page,
            "review": self.build_review_page,
            "purchase": self.build_purchase_page,
            "employees": self.build_employee_totals_page,
            "job_logo": self.build_outsourced_job_logo_page,
            "archive": self.build_archive_page,
            "settings": self.build_settings_page,
        }

    def _ensure_page_built(self, page: str) -> None:
        if page in self.pages:
            return
        builder = getattr(self, "_page_builders", {}).get(page)
        if builder is None:
            return
        builder()
        if page in {"import", "master", "review", "purchase", "employees"}:
            self._dashboard_refresh_key = None
        scrollable_names = {
            "review": ("review_decision_frame",),
            "employees": ("mc_employee_scroll",),
            "settings": ("settings_scroll",),
        }
        for name in scrollable_names.get(page, ()):
            register_scrollable(self, getattr(self, name, None))

    def _schedule_wave_header_resize(self, event=None):
        """Debounce high-resolution header rendering while the window resizes."""
        if not getattr(self, "_wave_source", None):
            return
        width = int(getattr(event, "width", 0) or self.main_header.winfo_width())
        height = int(getattr(event, "height", 0) or self.main_header.winfo_height())
        if width < 80 or height < 40:
            return
        if self._wave_render_job is not None:
            try:
                self.after_cancel(self._wave_render_job)
            except Exception:
                pass
        self._wave_render_job = self.after(
            70, lambda w=width, h=height: self._render_wave_header(w, h)
        )

    def _render_current_wave_header(self):
        if not getattr(self, "_wave_source", None):
            return
        width = max(80, int(self.main_header.winfo_width() or 1600))
        height = max(40, int(self.main_header.winfo_height() or 148))
        self._render_wave_header(width, height)

    @staticmethod
    def _spotlight_header_background(width: int, height: int) -> PILImage.Image:
        """Create the approved dark aubergine header with a soft central spotlight."""
        width = max(1, int(width))
        height = max(1, int(height))

        # Build the effect at a small resolution and upscale it smoothly. This
        # keeps window resizing fast while avoiding visible gradient banding.
        sample_width = min(320, width)
        sample_height = min(48, height)
        base = PILImage.new("RGBA", (sample_width, sample_height), (0, 0, 0, 255))
        base_draw = ImageDraw.Draw(base)
        top = (13, 6, 31)
        bottom = (19, 8, 43)
        for y in range(sample_height):
            ratio = y / max(1, sample_height - 1)
            color = tuple(
                round(top[channel] * (1.0 - ratio) + bottom[channel] * ratio)
                for channel in range(3)
            )
            base_draw.line((0, y, sample_width, y), fill=(*color, 255))

        mask = PILImage.new("L", (sample_width, sample_height), 0)
        mask_pixels = []
        for y in range(sample_height):
            normalized_y = (y / max(1, sample_height - 1) - 0.43) / 0.78
            for x in range(sample_width):
                normalized_x = (x / max(1, sample_width - 1) - 0.50) / 0.36
                distance = normalized_x * normalized_x + normalized_y * normalized_y
                mask_pixels.append(round(150 * pow(2.718281828, -distance * 1.65)))
        mask.putdata(mask_pixels)

        spotlight = PILImage.new("RGBA", base.size, (66, 20, 143, 255))
        background = PILImage.composite(spotlight, base, mask)
        return background.resize((width, height), PILImage.Resampling.LANCZOS)

    def _render_wave_header(self, width: int, height: int):
        """Render the approved spotlight header at 2x display resolution."""
        self._wave_render_job = None
        width = max(80, int(width))
        height = max(40, int(height))
        if self._wave_render_size == (width, height):
            return
        try:
            retina_size = (width * 2, height * 2)
            rendered = self._spotlight_header_background(*retina_size)
            pulse_scales = (1.0, 1.008, 1.016, 1.024)
            self.wave_header_pulse_frames = []
            for scale in pulse_scales:
                branded = self._composite_header_branding(rendered.copy(), scale=scale)
                self.wave_header_pulse_frames.append(ctk.CTkImage(
                    light_image=branded, dark_image=branded, size=(width, height)
                ))
            self.wave_header_image = self.wave_header_pulse_frames[0]
            self.wave_header_label.configure(image=self.wave_header_image)
            # Reuse the established interaction-animation controller with the
            # fully composited header frames. No transparent overlay is needed.
            self.sidebar_brand_logo_label = self.wave_header_label
            self.sidebar_brand_flower_frames = self.wave_header_pulse_frames
            self._wave_render_size = (width, height)
        except Exception:
            # The solid purple header remains visible if image rendering fails.
            pass

    @staticmethod
    def _tracked_text_width(draw: ImageDraw.ImageDraw, text: str, font, tracking: int) -> int:
        widths = [draw.textlength(character, font=font) for character in text]
        return round(sum(widths) + tracking * max(len(text) - 1, 0))

    def _composite_header_branding(self, canvas: PILImage.Image, scale: float = 1.0) -> PILImage.Image:
        """Render the approved flower / ORCHID | PURCHASE MANAGER header lockup."""
        logo_path = resource_path("assets", "orchid_logo.png")
        flower_path = resource_path("assets", "orchid_flower_purple.png")
        if not logo_path.exists() or not flower_path.exists():
            return canvas
        source = PILImage.open(logo_path).convert("RGBA")

        # Isolate the thin ORCHID lettering without the retail tagline or flower.
        letters = source.crop((0, 90, source.width, 325))
        bbox = letters.getchannel("A").getbbox()
        if bbox:
            letters = letters.crop(bbox)
        pixels = letters.load()
        for y in range(letters.height):
            for x in range(letters.width):
                red, green, blue, alpha = pixels[x, y]
                if alpha:
                    pixels[x, y] = (248, 246, 252, alpha)

        # Keep branding at the approved 4.8.19 size while later builds improve
        # performance; the taller header continues to frame the workspace.
        branding_height = min(canvas.height, 256)
        target_width = max(1, round(canvas.width * 0.18 * scale))
        target_height = max(1, round(branding_height * 0.29 * scale))
        letters = letters.resize((target_width, target_height), PILImage.Resampling.LANCZOS)
        draw = ImageDraw.Draw(canvas)
        font_size = max(12, round(branding_height * 0.14 * scale))
        font_candidates = (
            "/System/Library/Fonts/Supplemental/Arial Bold.ttf",
            "/System/Library/Fonts/Helvetica.ttc",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        )
        font = None
        for candidate in font_candidates:
            try:
                font = ImageFont.truetype(candidate, font_size)
                break
            except Exception:
                continue
        if font is None:
            font = ImageFont.load_default()
        subtitle = "PURCHASE MANAGER"
        tracking = max(1, round(font_size * 0.24))
        subtitle_width = self._tracked_text_width(draw, subtitle, font, tracking)

        flower_source = PILImage.open(flower_path).convert("RGBA")
        flower_limit = max(1, round(branding_height * 0.42 * scale))
        flower = self._purple_orchid_art(flower_source, (flower_limit, flower_limit))

        flower_gap = round(canvas.width * 0.013 * scale)
        divider_gap_left = round(canvas.width * 0.020 * scale)
        divider_gap_right = round(canvas.width * 0.024 * scale)
        divider_width = max(2, round(canvas.width * 0.0008))
        group_width = (
            flower.width + flower_gap + target_width + divider_gap_left
            + divider_width + divider_gap_right + subtitle_width
        )
        # The right-side PURCHASE MANAGER text makes a mathematically centered
        # lockup read slightly right-heavy. Shift the complete group 16 display
        # pixels left so its optical center aligns with the import/upload icon.
        display_scale = canvas.height / 148
        left = (canvas.width - group_width) // 2 - round(16 * display_scale)
        center_y = round(canvas.height * 0.50)

        flower_y = center_y - flower.height // 2
        canvas.alpha_composite(flower, (left, flower_y))
        letters_left = left + flower.width + flower_gap
        letters_top = center_y - target_height // 2
        canvas.alpha_composite(letters, (letters_left, letters_top))

        divider_x = letters_left + target_width + divider_gap_left
        divider_height = max(1, round(branding_height * 0.46 * scale))
        divider_top = center_y - divider_height // 2
        draw.rounded_rectangle(
            (divider_x, divider_top, divider_x + divider_width, divider_top + divider_height),
            radius=max(1, divider_width // 2), fill=(220, 201, 244, 218),
        )

        x = divider_x + divider_width + divider_gap_right
        text_box = draw.textbbox((0, 0), subtitle, font=font)
        text_height = text_box[3] - text_box[1]
        y = center_y - text_height // 2 - text_box[1]
        for character in subtitle:
            draw.text((x, y), character, font=font, fill=(248, 246, 252, 255))
            x += round(draw.textlength(character, font=font)) + tracking
        return canvas

    def show_page(self, page: str):
        if page == "review" and self.last_review_workbook and self.last_review_workbook.exists():
            try:
                has_product_setup = any(
                    _clean(issue.get("source", "")) == "Product Master"
                    for issue in (self.last_snapshot or self._mission_snapshot()).get("issues", [])
                )
                if has_product_setup:
                    page = "master"
            except Exception:
                pass
        self._ensure_page_built(page)
        if page not in self.pages:
            page = "dashboard"
        self.active_page = page
        for frame in self.pages.values():
            frame.grid_remove()
        self.pages[page].grid(row=0, column=0, sticky="nsew")
        for name, button in self.nav_buttons.items():
            if name == "import":
                button.grid()
            active = name == page
            button.configure(
                fg_color="#6A24D4" if active else "transparent",
                font=ctk.CTkFont(size=15, weight="bold" if active else "normal"),
            )
        active_event = self.current_event_name or "No Active Purchase Packet"
        self.header_chip.configure(text=active_event if len(active_event) <= 34 else active_event[:31] + "...")
        self.main_header.grid()
        self.refresh_dashboard()
        if page == "job_logo":
            self.refresh_outsourced_job_logo_page()

    # ---------- reusable UI ----------

    # ---------- reusable UI ----------
    def new_page(self, name: str) -> ctk.CTkFrame:
        page = ctk.CTkFrame(self.page_container, fg_color=BG, corner_radius=0)
        page.grid_columnconfigure(0, weight=1)
        page.grid_rowconfigure(9, weight=1)
        self.pages[name] = page
        return page

    def stat_card(self, parent, row, column, title, value, subtitle, columns=4):
        parent.grid_columnconfigure(tuple(range(columns)), weight=1)
        shadow = ctk.CTkFrame(parent, fg_color=SHADOW, corner_radius=15, width=1, height=1)
        shadow.grid(row=row, column=column, sticky="nsew", padx=6, pady=(9, 4))
        frame = ctk.CTkFrame(parent, fg_color=WHITE, border_width=1, border_color=BORDER, corner_radius=15)
        frame.grid(row=row, column=column, sticky="nsew", padx=6, pady=(5, 8))
        value_label = ctk.CTkLabel(frame, text=value, text_color=PURPLE,
                                   font=ctk.CTkFont(size=24, weight="bold"), anchor="w")
        value_label.pack(anchor="w", padx=18, pady=(16, 2))
        title_label = ctk.CTkLabel(frame, text=title, text_color=TEXT,
                                   font=ctk.CTkFont(size=13, weight="bold"), anchor="w")
        title_label.pack(anchor="w", padx=18)
        subtitle_label = ctk.CTkLabel(frame, text=subtitle, text_color=MUTED,
                                      font=ctk.CTkFont(size=11), anchor="w", wraplength=230, justify="left")
        subtitle_label.pack(anchor="w", padx=18, pady=(3, 15))
        return value_label, title_label, subtitle_label

    def section_card(self, parent, row, title, body):
        shadow = ctk.CTkFrame(parent, fg_color=SHADOW, corner_radius=16, width=1, height=1)
        shadow.grid(row=row, column=0, sticky="ew", padx=30, pady=(12, 4))
        frame = ctk.CTkFrame(parent, fg_color=WHITE, border_width=1, border_color=BORDER, corner_radius=16)
        frame.grid(row=row, column=0, sticky="ew", padx=30, pady=(8, 8))
        frame.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(frame, text=title, text_color=TEXT,
                     font=ctk.CTkFont(size=20, weight="bold"), anchor="w").grid(
            row=0, column=0, padx=24, pady=(19, 5), sticky="w"
        )
        ctk.CTkLabel(frame, text=body, text_color=MUTED,
                     font=ctk.CTkFont(size=13), wraplength=830,
                     justify="left", anchor="w").grid(
            row=1, column=0, padx=24, pady=(0, 18), sticky="w"
        )
        return frame

    def primary_button(self, parent, text, command, row=2, column=0, padx=(22, 8), sticky="w"):
        button = ctk.CTkButton(parent, text=text, command=command, height=42,
                               fg_color=PURPLE, hover_color=PURPLE_DARK,
                               font=ctk.CTkFont(size=13, weight="bold"))
        button.grid(row=row, column=column, padx=padx, pady=(0, 20), sticky=sticky)
        return button

    def secondary_button(self, parent, text, command, row=2, column=1, padx=(8, 22), sticky="w"):
        button = ctk.CTkButton(parent, text=text, command=command, height=42,
                               fg_color=PURPLE_LIGHT, hover_color="#E5D9F4",
                               border_width=1, border_color=BORDER, text_color=PURPLE_DARK,
                               font=ctk.CTkFont(size=13, weight="bold"))
        button.grid(row=row, column=column, padx=padx, pady=(0, 20), sticky=sticky)
        return button


    def _add_orchid_watermark(
        self, parent, size=(520, 520), relx=0.5, rely=0.56,
        anchor="center", opacity=0.055,
    ):
        """Place a large, centered, high-resolution orchid watermark behind page content."""
        try:
            path = resource_path("assets", "orchid_flower_purple.png")
            if not path.exists():
                return None
            source = PILImage.open(path).convert("RGBA")
            alpha = source.getchannel("A").point(
                lambda value: max(0, min(255, int(value * opacity)))
            )
            source.putalpha(alpha)
            image = ctk.CTkImage(light_image=source, dark_image=source, size=size)
            if not hasattr(self, "_workspace_watermark_images"):
                self._workspace_watermark_images = []
            self._workspace_watermark_images.append(image)
            label = ctk.CTkLabel(parent, text="", image=image, fg_color="transparent")
            label.place(relx=relx, rely=rely, anchor=anchor)
            return label
        except Exception:
            return None

    def _build_event_summary(self, page, row: int):
        """Create the compact current-event card without oversized empty space."""
        shadow = ctk.CTkFrame(page, fg_color=SHADOW, corner_radius=16, width=465, height=82)
        shadow.grid(row=row, column=0, sticky="e", padx=32, pady=(15, 5))
        shadow.grid_propagate(False)
        summary = ctk.CTkFrame(
            page, fg_color=WHITE, border_width=1, border_color=BORDER,
            corner_radius=16, width=465, height=76,
        )
        summary.grid(row=row, column=0, sticky="e", padx=32, pady=(11, 9))
        summary.pack_propagate(False)

        icon = ctk.CTkLabel(
            summary, text="▣", width=46, height=46, corner_radius=11,
            fg_color=PURPLE_LIGHT, text_color=PURPLE,
            font=ctk.CTkFont(size=20, weight="bold"),
        )
        icon.pack(side="left", padx=(13, 10), pady=14)

        text_block = ctk.CTkFrame(summary, fg_color="transparent")
        text_block.pack(side="left", fill="y")
        ctk.CTkLabel(
            text_block, text="CURRENT EVENT", text_color=PURPLE,
            font=ctk.CTkFont(size=9, weight="bold"), anchor="w",
        ).pack(anchor="w", pady=(15, 0))
        title = ctk.CTkLabel(
            text_block, text="No Active Purchase Packet", text_color=TEXT,
            font=ctk.CTkFont(size=13, weight="bold"), anchor="w", width=275,
        )
        title.pack(anchor="w", pady=(1, 0))

        date = ctk.CTkLabel(
            summary, text=time.strftime("%b %d, %Y"), text_color=MUTED,
            font=ctk.CTkFont(size=10, weight="bold"), anchor="e",
        )
        date.pack(side="right", padx=(10, 14))
        ctk.CTkFrame(summary, width=1, height=46, fg_color=BORDER).pack(
            side="right", fill="y", padx=(10, 0), pady=14
        )
        return summary, icon, title, date

    def _build_workflow_tracker(self, page, row: int, storage_attr: str, padx: int = 52):
        """Build the compact four-stage workflow card beneath the wave header."""
        panel = ctk.CTkFrame(
            page, fg_color=WHITE, border_width=1, border_color=BORDER,
            corner_radius=18, width=880, height=140,
        )
        panel.grid(row=row, column=0, sticky="n", padx=min(padx, 12), pady=(20, 14))
        panel.grid_propagate(False)
        stages = ctk.CTkFrame(panel, fg_color="transparent")
        stages.pack(fill="x", padx=26, pady=(20, 10))
        for column in (0, 2, 4, 6):
            stages.grid_columnconfigure(column, weight=0, minsize=132)
        for column in (1, 3, 5):
            stages.grid_columnconfigure(column, weight=1, minsize=48)

        stage_widgets = []
        stage_names = ("Import Orders", "Product Master", "Purchase Review", "Purchase Orders")
        target_pages = ("import", "master", "review", "purchase")
        for index, (label, target_page) in enumerate(zip(stage_names, target_pages)):
            step = ctk.CTkFrame(stages, fg_color="transparent", width=132)
            step.grid(row=0, column=index * 2, sticky="nsew")
            step.grid_columnconfigure(0, weight=1)
            circle = ctk.CTkButton(
                step, text=str(index + 1), width=44, height=44, corner_radius=22,
                fg_color=WHITE, hover_color=PURPLE_LIGHT, text_color=MUTED,
                border_width=1, border_color="#CFC7DA",
                font=ctk.CTkFont(size=14, weight="bold"),
                command=lambda p=target_page: self.show_page(p),
            )
            circle.grid(row=0, column=0, pady=(0, 8))
            title = ctk.CTkLabel(
                step, text=label, text_color=TEXT,
                font=ctk.CTkFont(size=12, weight="bold"),
            )
            title.grid(row=1, column=0)
            body_label = ctk.CTkLabel(step, text="", width=1, height=1)
            body_label.grid_remove()
            status = ctk.CTkLabel(
                step, text="Waiting", text_color=MUTED, fg_color="transparent",
                font=ctk.CTkFont(size=10, weight="bold"),
            )
            status.grid(row=2, column=0, pady=(4, 0))
            connector = None
            if index < 3:
                connector = ctk.CTkFrame(stages, height=2, fg_color="#D7D0E0", corner_radius=2)
                connector.grid(row=0, column=index * 2 + 1, sticky="ew", padx=10, pady=(21, 0))
            stage_widgets.append({
                "card": step, "circle": circle, "title": title,
                "body": body_label, "status": status, "connector": connector,
            })
        setattr(self, storage_attr, stage_widgets)
        return panel

    def build_dashboard_page(self):
        page = self.new_page("dashboard")
        page.grid_rowconfigure(2, weight=1)

        self.dashboard_view_toggle = ctk.CTkSegmentedButton(
            page, values=["Guided View", "Management View"], command=self._set_dashboard_view,
            width=300, height=36, corner_radius=11, fg_color="#EEE9F5",
            selected_color=PURPLE, selected_hover_color="#46109F",
            unselected_color="#EEE9F5", unselected_hover_color="#E3D9EF",
            text_color=TEXT, font=ctk.CTkFont(size=12, weight="bold"),
        )
        self.dashboard_view_toggle.place(relx=0.975, y=10, anchor="ne")
        self.mc_guided_tracker = self._build_workflow_tracker(page, 0, "mc_stage_labels", padx=52)
        self.mc_guided_tracker.grid_configure(pady=(46, 12))

        hero_shadow = ctk.CTkFrame(page, fg_color=SHADOW, corner_radius=20, width=760, height=382)
        self.mc_guided_shadow = hero_shadow
        hero_shadow.grid(row=1, column=0, sticky="n", padx=48, pady=(4, 18))
        hero = ctk.CTkFrame(
            page, fg_color=WHITE, border_width=1, border_color=BORDER,
            corner_radius=20, width=760, height=376,
        )
        hero.grid(row=1, column=0, sticky="n", padx=48, pady=(0, 22))
        hero.grid_propagate(False)
        hero.grid_columnconfigure(0, weight=1)
        self.mc_hero = hero
        self._add_orchid_watermark(
            hero, size=(420, 420), relx=0.5, rely=0.58,
            opacity=DASHBOARD_WATERMARK_OPACITY,
        )

        self.mc_animation_frame = ctk.CTkFrame(hero, fg_color="transparent", height=42)
        self.mc_animation_frame.grid(row=0, column=0, sticky="ew", padx=28, pady=(8, 0))
        self.mc_animation_frame.grid_propagate(False)
        self.success_flower_labels = []
        purple_flower = resource_path("assets", "orchid_flower_purple.png")
        if purple_flower.exists():
            flower_source = PILImage.open(purple_flower).convert("RGBA")
            self.success_flower_image = ctk.CTkImage(light_image=flower_source, dark_image=flower_source, size=(42, 42))
            self.success_bloom_image = ctk.CTkImage(light_image=flower_source, dark_image=flower_source, size=(64, 64))
            for _ in range(7):
                label = ctk.CTkLabel(self.mc_animation_frame, text="", image=self.success_flower_image,
                                     fg_color="transparent", width=44, height=44)
                self.success_flower_labels.append(label)
            self.success_bloom_label = ctk.CTkLabel(self.mc_animation_frame, text="", image=self.success_bloom_image,
                                                    fg_color="transparent", width=66, height=66)
        self.mc_animation_controls = ctk.CTkFrame(self.mc_animation_frame, fg_color="transparent")
        self.mc_sound_toggle = ctk.CTkButton(
            self.mc_animation_controls, text="Sound: On", command=self._toggle_success_sound,
            width=98, height=28, corner_radius=10, fg_color=PURPLE_LIGHT, hover_color="#E6D9F5",
            border_width=1, border_color="#CDBBE5", text_color=PURPLE_DARK,
            font=ctk.CTkFont(size=12, weight="bold"),
        )
        self.mc_sound_toggle.pack(side="left", padx=(0, 6))
        self.mc_skip_animation_button = ctk.CTkButton(
            self.mc_animation_controls, text="Skip Animation", command=self._skip_orchid_success_animation,
            width=122, height=28, corner_radius=10, fg_color="transparent", hover_color=PURPLE_LIGHT,
            border_width=1, border_color="#CDBBE5", text_color=PURPLE_DARK,
            font=ctk.CTkFont(size=12, weight="bold"), state="disabled",
        )
        self.mc_skip_animation_button.pack(side="left")
        self.mc_animation_controls.place(relx=1.0, rely=0.05, anchor="ne")
        self.mc_animation_controls.place_forget()
        self._update_success_control_buttons()

        self.mc_step_label = ctk.CTkLabel(hero, text="", width=1, height=1)
        self.mc_step_label.grid_remove()
        self.mc_banner_icon = ctk.CTkLabel(
            hero, text="↑", width=58, height=58, corner_radius=29,
            fg_color=PURPLE_LIGHT, text_color=PURPLE,
            font=ctk.CTkFont(size=27, weight="bold"),
        )
        self.mc_banner_icon.grid(row=1, column=0, pady=(6, 10))
        self.mc_banner_title = ctk.CTkLabel(
            hero, text="Import Orders", text_color=PURPLE_DARK,
            font=ctk.CTkFont(size=25, weight="bold"), justify="center", wraplength=620,
        )
        self.mc_banner_title.grid(row=2, column=0, padx=50, pady=(0, 8))
        self.mc_banner_subtitle = ctk.CTkLabel(
            hero, text="Choose the Shopify or Report Toaster export to start a new purchase packet.",
            text_color=MUTED, font=ctk.CTkFont(size=14), justify="center", wraplength=650,
        )
        self.mc_banner_subtitle.grid(row=3, column=0, padx=56, pady=(0, 14))
        self.dashboard_status = self.mc_banner_subtitle
        self.mc_event_label = ctk.CTkLabel(
            hero, text="No order CSV selected", text_color=TEXT,
            fg_color="#F3F1F6", corner_radius=14, height=31,
            font=ctk.CTkFont(size=12, weight="bold"), justify="center",
        )
        self.mc_event_label.grid(row=4, column=0, padx=40, pady=(0, 16))
        metrics = ctk.CTkFrame(hero, fg_color="#FAF8FC", corner_radius=14)
        metrics.grid(row=5, column=0, sticky="ew", padx=90, pady=(0, 16))
        metrics.grid_columnconfigure((0, 1), weight=1, uniform="metric")
        self.mc_metrics_frame = metrics
        self.mc_metric_widgets = {}
        for column, (key, caption) in enumerate((("employees", "Employees"), ("total", "Final Total"))):
            metric = ctk.CTkFrame(metrics, fg_color="transparent")
            metric.grid(row=0, column=column, padx=10, pady=12, sticky="nsew")
            value = ctk.CTkLabel(metric, text="—", text_color=PURPLE_DARK, font=ctk.CTkFont(size=23, weight="bold"))
            value.pack(pady=(0, 2))
            ctk.CTkLabel(metric, text=caption, text_color=MUTED, font=ctk.CTkFont(size=10, weight="bold")).pack()
            self.mc_metric_widgets[key] = {"frame": metric, "value": value}
        metrics.grid_remove()
        button_row = ctk.CTkFrame(hero, fg_color="transparent")
        button_row.grid(row=6, column=0, pady=(0, 22))
        self.mc_banner_button = ctk.CTkButton(
            button_row, text="Import Orders", command=lambda: self.show_page("import"),
            width=290, height=52, corner_radius=10, fg_color=PURPLE, hover_color="#46109F",
            font=ctk.CTkFont(size=15, weight="bold"),
        )
        self.mc_banner_button.pack(side="left", padx=7)
        self.mc_banner_secondary_button = ctk.CTkButton(
            button_row, text="Start New Event", command=self.start_new_purchase_packet,
            width=205, height=54, corner_radius=9, fg_color=PURPLE_LIGHT, hover_color="#E6D9F5",
            border_width=1, border_color="#BFA9DD", text_color=PURPLE_DARK,
            font=ctk.CTkFont(size=14, weight="bold"),
        )
        self.mc_banner_secondary_button.pack(side="left", padx=7)
        self.mc_banner_secondary_button.pack_forget()

        # Compatibility controls retained for the existing workflow engine.
        hidden = ctk.CTkFrame(page, fg_color="transparent")
        self.mc_master_card = self._mission_card(hidden, 0, "Product Master", "Permanent product catalog")
        self.mc_master_status = self.mc_master_card["status"]
        self.mc_master_detail = self.mc_master_card["detail"]
        self.mc_master_button = ctk.CTkButton(self.mc_master_card["frame"], text="Open Product Master", command=self.open_product_master)
        self.mc_master_button.grid(row=5, column=0)
        self.mc_review_card = self._mission_card(hidden, 1, "Purchase Review", "Only decisions that require attention")
        self.mc_review_status = self.mc_review_card["status"]
        self.mc_review_detail = self.mc_review_card["detail"]
        self.mc_review_button = ctk.CTkButton(self.mc_review_card["frame"], text="Open Purchase Review", command=self.open_latest_review)
        self.mc_review_button.grid(row=5, column=0)
        self.mc_po_card = self._mission_card(hidden, 2, "Purchase Orders", "PO numbers and final vendor PDFs")
        self.mc_po_status = self.mc_po_card["status"]
        self.mc_po_detail = self.mc_po_card["detail"]
        self.mc_po_card_button = ctk.CTkButton(self.mc_po_card["frame"], text="Edit PO Numbers", command=self.open_po_number_editor)
        self.mc_po_card_button.grid(row=5, column=0)
        self.mc_work_summary = ctk.CTkLabel(hidden, text="")
        self.mc_work_open_button = ctk.CTkButton(hidden, text="Open Review", command=self.open_latest_review)
        self.mc_work_scroll = ctk.CTkScrollableFrame(hidden, width=2, height=2)
        self.mc_work_scroll.grid_columnconfigure(1, weight=1)
        self.mc_action_title = ctk.CTkLabel(hidden, text="")
        self.mc_action_detail = ctk.CTkLabel(hidden, text="")
        self.mc_action_button = ctk.CTkButton(hidden, text="", command=lambda: None)

        self._build_management_dashboard(page)
        self.after_idle(self._apply_dashboard_view)

    def _set_dashboard_view(self, value: str) -> None:
        normalized = "management" if str(value).strip().casefold().startswith("management") else "guided"
        if normalized != getattr(self, "dashboard_view", "guided"):
            self.dashboard_view = normalized
            self._save_ui_preferences()
        self._apply_dashboard_view()
        self.refresh_dashboard()

    def _apply_dashboard_view(self) -> None:
        management = getattr(self, "dashboard_view", "guided") == "management"
        if hasattr(self, "dashboard_view_toggle"):
            self.dashboard_view_toggle.set("Management View" if management else "Guided View")
        guided_widgets = [
            getattr(self, "mc_guided_tracker", None),
            getattr(self, "mc_guided_shadow", None),
            getattr(self, "mc_hero", None),
        ]
        if management:
            for widget in guided_widgets:
                if widget is not None:
                    widget.grid_remove()
            if hasattr(self, "mc_management_frame"):
                self.mc_management_frame.place(relx=0, rely=0, relwidth=1, relheight=1)
        else:
            if hasattr(self, "mc_management_frame"):
                self.mc_management_frame.place_forget()
            for widget in guided_widgets:
                if widget is not None:
                    widget.grid()
        if hasattr(self, "dashboard_view_toggle"):
            self.dashboard_view_toggle.lift()

    def _management_metric_card(self, parent, column: int, title: str, key: str) -> None:
        card = ctk.CTkFrame(parent, fg_color=WHITE, border_width=1, border_color=BORDER, corner_radius=16)
        card.grid(row=0, column=column, sticky="nsew", padx=6, pady=5)
        value = ctk.CTkLabel(card, text="—", text_color=PURPLE, font=ctk.CTkFont(size=24, weight="bold"))
        value.pack(pady=(14, 2))
        ctk.CTkLabel(card, text=title, text_color=MUTED, font=ctk.CTkFont(size=11, weight="bold")).pack(pady=(0, 14))
        self.mc_management_metric_labels[key] = value

    def _management_detail_card(self, parent, row: int, column: int, title: str) -> tuple[ctk.CTkFrame, ctk.CTkLabel]:
        card = ctk.CTkFrame(parent, fg_color=WHITE, border_width=1, border_color=BORDER, corner_radius=17)
        card.grid(row=row, column=column, sticky="nsew", padx=7, pady=7)
        card.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(card, text=title, text_color=PURPLE_DARK, font=ctk.CTkFont(size=16, weight="bold"), anchor="w").grid(
            row=0, column=0, padx=18, pady=(16, 8), sticky="ew"
        )
        detail = ctk.CTkLabel(card, text="", text_color=TEXT, font=ctk.CTkFont(size=12), justify="left", anchor="nw", wraplength=330)
        detail.grid(row=1, column=0, padx=18, pady=(0, 16), sticky="nsew")
        card.grid_rowconfigure(1, weight=1)
        return card, detail

    def _build_management_dashboard(self, page: ctk.CTkFrame) -> None:
        frame = ctk.CTkScrollableFrame(page, fg_color=BG, corner_radius=0)
        self.mc_management_frame = frame
        frame.grid_columnconfigure((0, 1, 2), weight=1, uniform="management")

        heading = ctk.CTkFrame(frame, fg_color="transparent")
        heading.grid(row=0, column=0, columnspan=3, sticky="ew", padx=14, pady=(52, 6))
        heading.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(heading, text="Management Dashboard", text_color=PURPLE_DARK, font=ctk.CTkFont(size=27, weight="bold"), anchor="w").grid(row=0, column=0, sticky="w")
        self.mc_management_heading = ctk.CTkLabel(heading, text="Experienced view • the same live event and workflow", text_color=MUTED, font=ctk.CTkFont(size=12), anchor="w")
        self.mc_management_heading.grid(row=1, column=0, sticky="w", pady=(2, 0))

        metrics = ctk.CTkFrame(frame, fg_color="transparent")
        metrics.grid(row=1, column=0, columnspan=3, sticky="ew", padx=8, pady=(4, 4))
        for column in range(7):
            metrics.grid_columnconfigure(column, weight=1, uniform="management_metric")
        self.mc_management_metric_labels = {}
        for column, (title, key) in enumerate((
            ("Employees", "employees"), ("Line Items", "lines"), ("Ready", "ready"),
            ("Decisions", "decisions"), ("Blocked", "blocked"), ("Vendors", "vendors"),
            ("Ship to Orchid", "manual"),
        )):
            self._management_metric_card(metrics, column, title, key)

        event_card, self.mc_management_event_detail = self._management_detail_card(frame, 2, 0, "Event Summary")
        review_card, self.mc_management_review_detail = self._management_detail_card(frame, 2, 1, "Purchase Review Overview")

        actions = ctk.CTkFrame(frame, fg_color=WHITE, border_width=1, border_color=BORDER, corner_radius=17)
        actions.grid(row=2, column=2, sticky="nsew", padx=7, pady=7)
        actions.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(actions, text="Quick Actions", text_color=PURPLE_DARK, font=ctk.CTkFont(size=16, weight="bold"), anchor="w").grid(row=0, column=0, padx=18, pady=(16, 8), sticky="ew")
        quick_actions = [
            ("Open Purchase Review", lambda: self.show_page("review")),
            ("Regenerate Purchase Review", self.regenerate_current_review),
            ("Employee Totals", lambda: self.show_page("employees")),
            ("In-House Receiving Report", lambda: self._open_generated_report("in-house")),
            ("Outsourced Job Report", lambda: self._open_generated_report("outsourced")),
            ("Non-Included Items", lambda: self._open_generated_report("non-included")),
            ("Purchase Documents", self.open_latest_purchase_orders),
        ]
        self.mc_management_action_buttons = {}
        for row_index, (label, command) in enumerate(quick_actions, start=1):
            button = ctk.CTkButton(
                actions, text=label, command=command, anchor="w", height=36, corner_radius=9,
                fg_color="#F7F3FC", hover_color=PURPLE_LIGHT, border_width=1, border_color="#D8CBE8",
                text_color=PURPLE_DARK, font=ctk.CTkFont(size=12, weight="bold"),
            )
            button.grid(row=row_index, column=0, padx=16, pady=3, sticky="ew")
            self.mc_management_action_buttons[label] = button
        ctk.CTkFrame(actions, fg_color="transparent", height=9).grid(row=len(quick_actions) + 1, column=0)

        recent = ctk.CTkFrame(frame, fg_color=WHITE, border_width=1, border_color=BORDER, corner_radius=17)
        recent.grid(row=3, column=0, columnspan=3, sticky="ew", padx=7, pady=(7, 22))
        recent.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(recent, text="Recent Activity", text_color=PURPLE_DARK, font=ctk.CTkFont(size=16, weight="bold"), anchor="w").grid(row=0, column=0, padx=18, pady=(15, 6), sticky="ew")
        self.mc_management_recent_detail = ctk.CTkLabel(recent, text="No event activity yet.", text_color=TEXT, font=ctk.CTkFont(size=12), justify="left", anchor="w")
        self.mc_management_recent_detail.grid(row=1, column=0, padx=18, pady=(0, 16), sticky="ew")
        frame.place_forget()

    def _open_generated_report(self, report_kind: str) -> None:
        root = self.last_purchase_order_dir if self.last_purchase_order_dir and self.last_purchase_order_dir.exists() else REPORTS / "Purchase Orders"
        if not root.exists():
            messagebox.showinfo("Report Not Generated", "Generate purchase orders and reports first.")
            return
        tokens = {
            "in-house": ("in-house", "receiving"),
            "outsourced": ("outsourced", "decoration"),
            "non-included": ("non-included", "non included"),
        }.get(report_kind, (report_kind,))
        candidates = sorted(root.rglob("*.pdf"), key=lambda path: path.stat().st_mtime, reverse=True)
        match = next((path for path in candidates if any(token in path.name.casefold() for token in tokens)), None)
        if match is None:
            messagebox.showinfo("Report Not Generated", "That report has not been generated for the current event yet.")
            return
        subprocess.run(["open", str(match)], check=False)

    def _management_manual_piece_count(self) -> int:
        workbook = self.last_review_workbook
        if not workbook or not workbook.exists():
            return 0
        try:
            total = 0
            for row in load_review_lines(workbook):
                include = _clean(row.get("Include", "Yes")).casefold()
                manual = _clean(row.get("Do Not Outsource", "No")).casefold()
                if include in {"yes", "y", "true", "1", "include"} and manual in {"yes", "y", "true", "1"}:
                    try:
                        total += max(int(float(row.get("Quantity", 0) or 0)), 0)
                    except (TypeError, ValueError):
                        pass
            return total
        except Exception:
            return 0

    def _management_recent_activity(self) -> str:
        entries = []
        paths = [
            (self.selected_csv, "Orders imported"),
            (live_product_master_path(), "Product Master updated"),
            (self.last_review_workbook, "Purchase Review updated"),
            (self.last_purchase_order_dir, "Purchase documents generated"),
        ]
        for path, label in paths:
            try:
                if path and Path(path).exists():
                    modified = Path(path).stat().st_mtime
                    entries.append((modified, f"{time.strftime('%b %d, %I:%M %p', time.localtime(modified))}  •  {label}"))
            except Exception:
                pass
        entries.sort(key=lambda item: item[0], reverse=True)
        return "\n".join(text for _, text in entries[:5]) or "No event activity yet."

    def _refresh_management_dashboard(self, snapshot: dict, purchase_generated: bool = False) -> None:
        if not hasattr(self, "mc_management_metric_labels"):
            return
        line_count = int(snapshot.get("line_count", self.imported_line_count or 0) or 0)
        ready = int(snapshot.get("ready_count", 0) or 0)
        decisions = int(snapshot.get("purchase_review_count", snapshot.get("review_count", 0)) or 0)
        blocked = int(snapshot.get("blocked_route_count", 0) or 0)
        product_setup = int(snapshot.get("product_master_count", 0) or 0)
        employees = int(snapshot.get("employee_count", 0) or 0)
        routes = list(snapshot.get("routes", []))
        vendors = len({_clean(route.get("vendor", "")).casefold() for route in routes if _clean(route.get("vendor", ""))})
        manual = self._management_manual_piece_count()
        values = {
            "employees": employees, "lines": line_count, "ready": ready, "decisions": decisions,
            "blocked": blocked, "vendors": vendors, "manual": manual,
        }
        for key, value in values.items():
            self.mc_management_metric_labels[key].configure(text=f"{value:,}")

        event_name = self.current_event_name or snapshot.get("event_name") or "No Active Purchase Packet"
        mode = self.current_mode or snapshot.get("report_mode") or "—"
        fulfillment = normalize_decoration_fulfillment(self.current_decoration_fulfillment)
        final_total = float(snapshot.get("employee_grand_total", 0.0) or 0.0)
        po_state = "Generated" if purchase_generated else ("Ready to generate" if line_count and not decisions and not blocked and not product_setup else "In progress")
        self.mc_management_event_detail.configure(text=(
            f"Event:  {event_name}\n"
            f"Order mode:  {mode}\n"
            f"Decoration workflow:  {fulfillment}\n"
            f"Employee total:  ${final_total:,.2f}\n"
            f"Purchase documents:  {po_state}"
        ))
        self.mc_management_review_detail.configure(text=(
            f"Product Master setup:  {product_setup:,}\n"
            f"Purchase Review decisions:  {decisions:,}\n"
            f"Blocked routes:  {blocked:,}\n"
            f"Ready purchase lines:  {ready:,} of {line_count:,}\n"
            f"Ship-to-Orchid exception pieces:  {manual:,}"
        ))
        self.mc_management_recent_detail.configure(text=self._management_recent_activity())
        employee_button = self.mc_management_action_buttons.get("Employee Totals")
        if employee_button is not None:
            employee_button.configure(state="normal" if should_show_employee_totals(snapshot, bool(self.last_review_workbook), self.current_mode, self.imported_line_count) else "disabled")
        report_state = "normal" if self.last_purchase_order_dir and self.last_purchase_order_dir.exists() else "disabled"
        for label in ("In-House Receiving Report", "Outsourced Job Report", "Non-Included Items"):
            button = self.mc_management_action_buttons.get(label)
            if button is not None:
                button.configure(state=report_state)

    def _mission_card(self, parent, column: int, title: str, subtitle: str) -> dict:
        frame = ctk.CTkFrame(parent, fg_color=WHITE, border_width=1, border_color=BORDER, corner_radius=15)
        frame.grid(row=0, column=column, sticky="nsew", padx=7, pady=5)
        frame.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(frame, text=title, text_color=TEXT,
                     font=ctk.CTkFont(size=19, weight="bold")).grid(row=0, column=0, padx=16, pady=(16, 3))
        ctk.CTkLabel(frame, text=subtitle, text_color=MUTED,
                     font=ctk.CTkFont(size=12)).grid(row=1, column=0, padx=16)
        status = ctk.CTkLabel(frame, text="Not started", height=28, corner_radius=14,
                              fg_color="#ECEAF0", text_color=MUTED,
                              font=ctk.CTkFont(size=13, weight="bold"))
        status.grid(row=2, column=0, padx=16, pady=(10, 6))
        detail = ctk.CTkLabel(frame, text="", text_color=TEXT,
                              font=ctk.CTkFont(size=12), justify="center", wraplength=290)
        detail.grid(row=3, column=0, padx=16, pady=(0, 6))
        return {"frame": frame, "status": status, "detail": detail}

    def _highlight_mission_card(self, active: str | None):
        cards = {
            "master": self.mc_master_card["frame"],
            "review": self.mc_review_card["frame"],
            "purchase": self.mc_po_card["frame"],
        }
        for name, frame in cards.items():
            frame.configure(
                border_width=3 if name == active else 1,
                border_color=PURPLE if name == active else BORDER,
            )

    def _sync_current_product_candidates(self) -> dict:
        """Add deduplicated review candidates to the live Product Master once.

        This makes genuinely new Report Toaster/Shopify styles immediately
        available in the Product Master Needs Setup queue without overwriting
        any saved manual assignments.
        """
        workbook = self.last_review_workbook
        if not workbook or not Path(workbook).exists():
            return {"changed": False}
        workbook_key = str(Path(workbook).resolve())
        if self._candidate_sync_workbook == workbook_key:
            return {"changed": False, "cached": True}
        try:
            result = sync_review_product_candidates(Path(workbook), live_product_master_path())
            self._candidate_sync_workbook = workbook_key
            if result.get("changed"):
                try:
                    self.product_master_update_marker.parent.mkdir(parents=True, exist_ok=True)
                    self.product_master_update_marker.write_text(str(time.time()), encoding="utf-8")
                    self._last_product_master_update = self.product_master_update_marker.stat().st_mtime
                except Exception:
                    pass
            return result
        except Exception as error:
            print(f"Product candidate sync warning: {error}", file=sys.stderr)
            return {"changed": False, "error": str(error)}

    def _review_is_stale(self) -> bool:
        try:
            master = live_product_master_path()
            return bool(
                self.last_review_workbook
                and self.last_review_workbook.exists()
                and master.exists()
                and master.stat().st_mtime > self.last_review_workbook.stat().st_mtime
            )
        except Exception:
            return False

    def _poll_product_master_updates(self):
        try:
            existing_pid = self._existing_product_master_pid()
            if existing_pid:
                self._product_master_was_open = True
            elif self._product_master_was_open:
                # Product Master has just closed. Rebuild the review once so all
                # saved setup decisions are applied without asking the employee
                # to hunt for a regenerate action.
                self._product_master_was_open = False
                if self._review_is_stale() and self.last_review_workbook:
                    self.regenerate_current_review(silent=True)

            if self.product_master_update_marker.exists():
                modified = self.product_master_update_marker.stat().st_mtime
                if modified > self._last_product_master_update:
                    self._last_product_master_update = modified
                    self.refresh_dashboard()
                    self._schedule_auto_review_regeneration()
        except Exception:
            pass
        try:
            self.after(900, self._poll_product_master_updates)
        except Exception:
            pass

    def _schedule_auto_review_regeneration(self):
        """Debounce Product Master saves into one silent Purchase Review refresh."""
        if not self.last_review_workbook or not REVIEW_STATE_FILE.exists() or not self._review_is_stale():
            return
        if self._auto_regenerate_after_id:
            try:
                self.after_cancel(self._auto_regenerate_after_id)
            except Exception:
                pass
        self._auto_regenerate_after_id = self.after(700, self._run_auto_review_regeneration)

    def _run_auto_review_regeneration(self):
        self._auto_regenerate_after_id = None
        if self._regenerating_review or not self._review_is_stale():
            return
        self.regenerate_current_review(silent=True)

    def regenerate_current_review(self, silent: bool = False):
        from modules.review_workbook import generate_review_workbook

        if self._regenerating_review:
            return
        self._regenerating_review = True
        try:
            if not REVIEW_STATE_FILE.exists():
                messagebox.showinfo("No Saved Review", "Create a Purchase Review first.")
                return
            state = json.loads(REVIEW_STATE_FILE.read_text(encoding="utf-8"))
            source_csv = Path(state.get("source_csv", "")).expanduser()
            if not source_csv.exists():
                messagebox.showerror("Orders File Missing", f"The saved order CSV can no longer be found:\n{source_csv}")
                return
            result = generate_review_workbook(
                source_csv, live_product_master_path(), REPORTS,
                report_mode=state.get("report_mode") or GENERAL_SALES_PERIOD,
                event_name=state.get("event_name") or "",
                decoration_fulfillment=state.get("decoration_fulfillment") or STANDARD_ORCHID_WORKFLOW,
                regenerate_helper=REGENERATE_HELPER,
                previous_review_path=self.last_review_workbook,
            )
            self.selected_csv = source_csv
            self.last_review_workbook = Path(result["output_path"])
            self._reset_review_navigation()
            self.current_event_name = result.get("event_name") or self.current_event_name or "General Sales Period"
            self.current_mode = result.get("report_mode") or self.current_mode
            self.current_decoration_fulfillment = result.get("decoration_fulfillment") or self.current_decoration_fulfillment
            self.imported_line_count = int(result.get("lines", 0))
            self.review_needed = int(result.get("review_decisions", result.get("review_lines", 0)))
            self._sync_current_product_candidates()
            self.refresh_dashboard()
            if silent:
                self.dashboard_status.configure(
                    text="Purchase Review refreshed automatically after Product Master was saved.",
                    text_color=SUCCESS,
                )
            else:
                messagebox.showinfo(
                    "Purchase Review Refreshed",
                    "The current Product Master changes were applied. Continue the review inside Orchid Purchase Manager.",
                )
        except Exception as error:
            if not silent:
                messagebox.showerror("Unable to Regenerate Purchase Review", str(error))
            else:
                self.dashboard_status.configure(text=f"Unable to refresh Purchase Review: {error}", text_color=DANGER)
        finally:
            self._regenerating_review = False

    def _set_pill(self, widget, text: str, state: str):
        styles = {
            "complete": (SUCCESS, "#EAF8EF"),
            "attention": (WARNING, "#FFF1DC"),
            "pending": (MUTED, "#ECEAF0"),
        }
        color, fill = styles[state]
        widget.configure(text=text, text_color=color, fg_color=fill)

    def _set_stage(self, index: int, state: str) -> None:
        if state == "complete":
            circle_fill, circle_text, border, symbol = SUCCESS, WHITE, SUCCESS, "✓"
            title_color, status_color, status_text = "#166534", "#166534", "Complete"
            enabled = "normal"
        elif state == "active":
            circle_fill, circle_text, border, symbol = PURPLE, WHITE, PURPLE, str(index + 1)
            title_color, status_color, status_text = PURPLE, PURPLE, "Start Here" if index == 0 else "Current"
            enabled = "normal"
        elif state == "attention":
            circle_fill, circle_text, border, symbol = "#FFF1DC", "#9A4D00", WARNING, "!"
            title_color, status_color, status_text = "#9A4D00", "#9A4D00", "Needs Attention"
            enabled = "normal"
        else:
            circle_fill, circle_text, border, symbol = WHITE, MUTED, "#CFC7DA", str(index + 1)
            title_color, status_color, status_text = TEXT, MUTED, "Waiting"
            enabled = "disabled"
        for attr in (
            "mc_stage_labels", "current_event_stage_labels", "master_stage_labels",
            "review_stage_labels", "purchase_stage_labels",
        ):
            tracker = getattr(self, attr, None)
            if not tracker or index >= len(tracker):
                continue
            stage = tracker[index]
            stage["circle"].configure(
                text=symbol, fg_color=circle_fill, hover_color=PURPLE_LIGHT,
                text_color=circle_text, border_color=border, state=enabled,
            )
            stage["title"].configure(text_color=title_color)
            stage["status"].configure(text=status_text, text_color=status_color, fg_color="transparent")
            connector = stage.get("connector")
            if connector is not None:
                connector.configure(fg_color=PURPLE if state == "complete" else "#D7D0E0")

    def _update_dashboard_metrics(self, snapshot: dict, show: bool) -> None:
        if not hasattr(self, "mc_metrics_frame"):
            return
        if not show:
            self.mc_metrics_frame.grid_remove()
            self.mc_hero.configure(height=376)
            self.mc_guided_shadow.configure(height=382)
            return
        self.mc_hero.configure(height=450)
        self.mc_guided_shadow.configure(height=456)
        employee_count = int(snapshot.get("employee_count", 0) or 0)
        final_total = float(snapshot.get("employee_grand_total", 0.0) or 0.0)
        values = {
            "employees": f"{employee_count:,}" if employee_count else "—",
            "total": f"${final_total:,.2f}" if final_total else "—",
        }
        for key, value in values.items():
            self.mc_metric_widgets[key]["value"].configure(text=value)
        self.mc_metrics_frame.grid()

    def _prefers_reduced_motion(self) -> bool:
        env = os.environ.get("ORCHID_REDUCED_MOTION", "").strip().casefold()
        if env in {"1", "true", "yes", "on"}:
            return True
        if env in {"0", "false", "no", "off"}:
            return False
        if getattr(self, "_reduced_motion_override", None) is not None:
            return bool(self._reduced_motion_override)
        if sys.platform == "darwin":
            try:
                value = subprocess.check_output(
                    ["defaults", "read", "com.apple.universalaccess", "reduceMotion"],
                    stderr=subprocess.DEVNULL,
                    text=True,
                    timeout=1.2,
                ).strip()
                return value == "1"
            except Exception:
                return False
        return False

    def _toggle_success_sound(self) -> None:
        self.success_sound_enabled = not bool(getattr(self, "success_sound_enabled", True))
        self._update_success_control_buttons()

    def _update_success_control_buttons(self) -> None:
        if hasattr(self, "mc_sound_toggle"):
            self.mc_sound_toggle.configure(
                text="Sound: On" if bool(getattr(self, "success_sound_enabled", True)) else "Sound: Off"
            )
        if hasattr(self, "mc_skip_animation_button"):
            self.mc_skip_animation_button.configure(
                state="normal" if bool(getattr(self, "_success_animation_running", False)) else "disabled"
            )

    def _play_success_sound(self) -> None:
        if not bool(getattr(self, "success_sound_enabled", True)):
            return
        sound_path = resource_path("assets", "success_chime.wav")
        try:
            if sys.platform == "darwin":
                if sound_path.exists():
                    subprocess.Popen(
                        ["afplay", str(sound_path)],
                        stdout=subprocess.DEVNULL,
                        stderr=subprocess.DEVNULL,
                    )
                else:
                    self.bell()
            elif sys.platform.startswith("win"):
                try:
                    import winsound
                    if sound_path.exists():
                        winsound.PlaySound(str(sound_path), winsound.SND_FILENAME | winsound.SND_ASYNC)
                    else:
                        winsound.MessageBeep()
                except Exception:
                    self.bell()
            else:
                self.bell()
        except Exception:
            try:
                self.bell()
            except Exception:
                pass

    def _finish_success_animation(self) -> None:
        for label in getattr(self, "success_flower_labels", []):
            try:
                label.place_forget()
            except Exception:
                pass
        bloom = getattr(self, "success_bloom_label", None)
        if bloom is not None:
            try:
                bloom.place_forget()
            except Exception:
                pass
        self._success_animation_running = False
        self._update_success_control_buttons()

    def _skip_orchid_success_animation(self) -> None:
        if not bool(getattr(self, "_success_animation_running", False)):
            return
        self._success_animation_skip_requested = True
        self._finish_success_animation()

    def _play_orchid_success_animation(self) -> None:
        labels = getattr(self, "success_flower_labels", [])
        if getattr(self, "_success_animation_running", False):
            return
        self._success_animation_running = True
        self._success_animation_skip_requested = False
        self._update_success_control_buttons()
        for label in labels:
            label.place_forget()
        bloom = getattr(self, "success_bloom_label", None)
        if bloom is not None:
            bloom.place_forget()

        reduced = self._prefers_reduced_motion()
        positions = (0.10, 0.23, 0.36, 0.49, 0.62, 0.75, 0.88)
        walk_delay = 180 if reduced else 320
        visible_ms = 640 if reduced else 1200

        for index, (label, relx) in enumerate(zip(labels, positions)):
            def show_flower(widget=label, x=relx):
                if not getattr(self, "_success_animation_running", False) or getattr(self, "_success_animation_skip_requested", False):
                    return
                widget.place(relx=x, rely=0.5, anchor="center")
                widget.lift()
            def hide_flower(widget=label):
                if getattr(self, "_success_animation_skip_requested", False):
                    return
                widget.place_forget()
            self.after(index * walk_delay, show_flower)
            self.after(index * walk_delay + visible_ms, hide_flower)

        bloom_start = 90 if reduced else len(labels) * walk_delay + 140

        def show_bloom():
            if not getattr(self, "_success_animation_running", False) or getattr(self, "_success_animation_skip_requested", False):
                return
            bloom_widget = getattr(self, "success_bloom_label", None)
            if bloom_widget is not None:
                bloom_widget.place(relx=0.5, rely=0.5, anchor="center")
                bloom_widget.lift()
            self._play_success_sound()

        def hide_bloom():
            if getattr(self, "_success_animation_skip_requested", False):
                return
            bloom_widget = getattr(self, "success_bloom_label", None)
            if bloom_widget is not None:
                bloom_widget.place_forget()
            self._finish_success_animation()

        self.after(bloom_start, show_bloom)
        self.after(bloom_start + (650 if reduced else 1050), hide_bloom)

    def _mission_snapshot(self, force: bool = False) -> dict:
        if not self.last_review_workbook or not self.last_review_workbook.exists():
            self._mission_snapshot_cache_key = ("none",)
            self._mission_snapshot_cache = {
                "issues": [], "routes": [], "review_count": 0, "blocked_route_count": 0,
                "workflow_blocked": False, "line_count": 0, "ready_count": 0,
                "event_name": "", "report_mode": "", "employee_totals": [],
                "employee_count": 0, "employee_grand_total": 0.0,
            }
            return self._mission_snapshot_cache
        cache_key = (
            _file_stamp(self.last_review_workbook),
            _file_stamp(DATA / "mission_control_state.json"),
        )
        if not force and cache_key == self._mission_snapshot_cache_key:
            return self._mission_snapshot_cache
        try:
            snapshot = load_mission_control_snapshot(self.last_review_workbook, DATA)
            self._mission_snapshot_cache_key = cache_key
            self._mission_snapshot_cache = snapshot
            return snapshot
        except Exception as error:
            return {"issues": [], "routes": [], "review_count": 0, "blocked_route_count": 0,
                    "workflow_blocked": False, "line_count": 0, "ready_count": 0,
                    "event_name": "", "report_mode": "", "employee_totals": [],
                    "employee_count": 0, "employee_grand_total": 0.0, "error": str(error)}

    def _clear_children(self, frame):
        for child in frame.winfo_children():
            child.destroy()

    def _configure_dashboard_workspace(self, snapshot: dict) -> bool:
        """Expose Employee Totals as a sidebar page only for sizing events."""
        show_employee_totals = should_show_employee_totals(
            snapshot,
            bool(self.last_review_workbook and self.last_review_workbook.exists()),
            self.current_mode,
            self.imported_line_count,
        )
        if hasattr(self, "employee_nav_button"):
            if show_employee_totals:
                self.employee_nav_button.grid()
            else:
                self.employee_nav_button.grid_remove()
        if hasattr(self, "outsourced_logo_nav_button"):
            # The job-report logo is used by standard outsourced screen-print
            # reports as well as Entire Order Outsourced packets, so keep this
            # setup page permanently available in the sidebar.
            self.outsourced_logo_nav_button.grid()
        return show_employee_totals

    def _render_work_queue(self, snapshot: dict):
        self._clear_children(self.mc_work_scroll)
        issues = list(snapshot.get("issues", []))
        routes = list(snapshot.get("routes", []))
        line_count = int(snapshot.get("line_count", 0) or 0)
        blocked_routes = [
            route for route in routes
            if str(route.get("status", "")).casefold() != "ready"
        ]
        if not self.last_review_workbook or line_count == 0:
            self.mc_work_summary.configure(text="No active purchase packet.")
            ctk.CTkLabel(
                self.mc_work_scroll,
                text="Import Shopify orders and create a Purchase Review to begin.",
                text_color=MUTED, font=ctk.CTkFont(size=12),
            ).grid(row=0, column=0, padx=12, pady=22, sticky="w")
            return

        product_master_issues = [
            issue for issue in issues
            if str(issue.get("fix_in", "")).casefold() == "product master"
        ]
        purchase_review_issues = [
            issue for issue in issues
            if str(issue.get("fix_in", "")).casefold() != "product master"
        ]
        if not product_master_issues and not purchase_review_issues and not blocked_routes:
            self.mc_work_summary.configure(text="No blockers remain. This purchase packet is ready.")
            ctk.CTkLabel(
                self.mc_work_scroll, text="✓  Purchase Review complete", text_color=SUCCESS,
                font=ctk.CTkFont(size=14, weight="bold"),
            ).grid(row=0, column=0, padx=12, pady=22, sticky="w")
            return

        self.mc_work_summary.configure(
            text=(
                f"{len(product_master_issues)} new product setup item(s)  •  "
                f"{len(purchase_review_issues)} order review item(s)  •  "
                f"{len(blocked_routes)} blocked route(s)"
            )
        )
        self.mc_work_scroll.grid_columnconfigure(0, weight=1)
        row_cursor = 0

        def add_section(title: str, subtitle: str, rows: list[dict], kind: str, limit: int = 6):
            nonlocal row_cursor
            if not rows:
                return
            section = ctk.CTkFrame(
                self.mc_work_scroll, fg_color=WHITE, border_width=1,
                border_color="#E5DDEE", corner_radius=10,
            )
            section.grid(row=row_cursor, column=0, sticky="ew", padx=4, pady=(4, 8))
            section.grid_columnconfigure(0, weight=1)
            header = ctk.CTkFrame(section, fg_color=PURPLE_LIGHT, corner_radius=8)
            header.grid(row=0, column=0, sticky="ew", padx=7, pady=7)
            header.grid_columnconfigure(0, weight=1)
            ctk.CTkLabel(
                header, text=f"{title}  ({len(rows)})", text_color=PURPLE_DARK,
                font=ctk.CTkFont(size=13, weight="bold"), anchor="w",
            ).grid(row=0, column=0, padx=10, pady=(7, 1), sticky="w")
            ctk.CTkLabel(
                header, text=subtitle, text_color=MUTED,
                font=ctk.CTkFont(size=10), anchor="w",
            ).grid(row=1, column=0, padx=10, pady=(0, 7), sticky="w")

            for index, item in enumerate(rows[:limit], start=1):
                row_frame = ctk.CTkFrame(section, fg_color="#FCFBFE", corner_radius=8)
                row_frame.grid(row=index, column=0, sticky="ew", padx=7, pady=(0, 6))
                row_frame.grid_columnconfigure(1, weight=1)
                if kind == "master":
                    badge_text, badge_color, badge_fill = "NEW STYLE", WARNING, "#FFF1DC"
                    product = " — ".join(part for part in [item.get("product"), item.get("description")] if part)
                    action_text = item.get("instructions") or item.get("reason") or "Complete Product Master setup"
                    product_number = item.get("product") or ""
                    command = lambda value=product_number: self.open_product_master(search_product=value)
                    button_text = "Open Product"
                elif kind == "route":
                    badge_text, badge_color, badge_fill = "BLOCKED ROUTE", DANGER, "#FDECEC"
                    product = f"{item.get('vendor', 'Unknown vendor')} — {item.get('report_type', 'Needs routing')}"
                    action_text = "Resolve the remaining product or order decision that is blocking this purchase order."
                    command = self.open_latest_review
                    button_text = "Open Review"
                else:
                    badge_text, badge_color, badge_fill = "ORDER REVIEW", WARNING, "#FFF1DC"
                    product = " — ".join(part for part in [item.get("product"), item.get("description")] if part)
                    person = " • ".join(part for part in [item.get("employee"), item.get("order")] if part)
                    if person:
                        product = product + ("\n" if product else "") + person
                    action_text = item.get("instructions") or item.get("reason") or "Complete the order-specific correction"
                    command = self.open_latest_review
                    button_text = "Open Review"
                ctk.CTkLabel(
                    row_frame, text=badge_text, width=96, height=25, corner_radius=12,
                    fg_color=badge_fill, text_color=badge_color,
                    font=ctk.CTkFont(size=9, weight="bold"),
                ).grid(row=0, column=0, rowspan=2, padx=(9, 8), pady=9, sticky="w")
                ctk.CTkLabel(
                    row_frame, text=product or "Unknown item", text_color=TEXT,
                    font=ctk.CTkFont(size=11, weight="bold"), anchor="w", justify="left",
                ).grid(row=0, column=1, padx=4, pady=(7, 1), sticky="ew")
                ctk.CTkLabel(
                    row_frame, text=action_text, text_color=MUTED,
                    font=ctk.CTkFont(size=10), anchor="w", justify="left", wraplength=570,
                ).grid(row=1, column=1, padx=4, pady=(0, 7), sticky="ew")
                ctk.CTkButton(
                    row_frame, text=button_text, width=105, height=30,
                    fg_color=PURPLE_LIGHT, hover_color="#E4D6F5", text_color=PURPLE_DARK,
                    command=command,
                ).grid(row=0, column=2, rowspan=2, padx=9, pady=9, sticky="e")
            if len(rows) > limit:
                ctk.CTkLabel(
                    section, text=f"+ {len(rows) - limit} more", text_color=MUTED,
                    font=ctk.CTkFont(size=10, weight="bold"),
                ).grid(row=limit + 1, column=0, padx=12, pady=(0, 8), sticky="w")
            row_cursor += 1

        add_section(
            "Product Master", "Permanent setup saved for future orders.",
            product_master_issues, "master",
        )
        add_section(
            "Purchase Review", "Corrections that apply only to the current orders.",
            purchase_review_issues, "review",
        )
        add_section(
            "Blocked Purchase Orders", "Routes waiting on a remaining decision.",
            blocked_routes, "route", limit=5,
        )

    @staticmethod
    def _money_text(value: object) -> str:
        try:
            return f"${float(value):,.2f}"
        except (TypeError, ValueError):
            return "$0.00"

    @staticmethod
    def _discount_value(value: object) -> float:
        text = str(value or "").replace("$", "").replace(",", "").strip()
        if not text:
            return 0.0
        amount = float(text)
        if amount < 0:
            raise ValueError("Discounts cannot be negative.")
        return amount

    def _render_employee_totals(self, snapshot: dict):
        self._clear_children(self.mc_employee_scroll)
        self.employee_discount_vars = {}
        self.employee_total_labels = {}
        records = list(snapshot.get("employee_totals", []))
        self.employee_records_by_row = {
            int(record.get("row", index + 1)): dict(record)
            for index, record in enumerate(records)
        }

        report_mode = str(snapshot.get("report_mode", "") or "")
        if not self.last_review_workbook:
            self.mc_employee_summary.configure(
                text="Create a Uniform Sizing Event Purchase Review to see employee totals."
            )
            self.mc_employee_grand_total.configure(text="Grand Total  $0.00")
            self.mc_employee_save_button.configure(state="disabled")
            ctk.CTkLabel(
                self.mc_employee_scroll,
                text="Employee totals will appear here after the Purchase Review is created.",
                text_color=MUTED, font=ctk.CTkFont(size=12), wraplength=360, justify="left",
            ).grid(row=0, column=0, padx=12, pady=22, sticky="w")
            return
        if report_mode != UNIFORM_SIZING_EVENT:
            self.mc_employee_summary.configure(text="Employee totals are available for Uniform Sizing Events.")
            self.mc_employee_grand_total.configure(text="Grand Total  $0.00")
            self.mc_employee_save_button.configure(state="disabled")
            ctk.CTkLabel(
                self.mc_employee_scroll,
                text="This purchase packet is using General Sales Period mode.",
                text_color=MUTED, font=ctk.CTkFont(size=12), wraplength=360, justify="left",
            ).grid(row=0, column=0, padx=12, pady=22, sticky="w")
            return
        if not records:
            self.mc_employee_summary.configure(text="No employee totals were found in this Purchase Review.")
            self.mc_employee_grand_total.configure(text="Grand Total  $0.00")
            self.mc_employee_save_button.configure(state="disabled")
            return

        self.mc_employee_summary.configure(
            text=f"{len(records)} employee{'s' if len(records) != 1 else ''} • Enter discounts below, then save."
        )
        self.mc_employee_save_button.configure(state="normal")
        headers = ["Employee", "Order Total", "Discount", "Total"]
        for column, header in enumerate(headers):
            ctk.CTkLabel(
                self.mc_employee_scroll, text=header, text_color=PURPLE_DARK,
                font=ctk.CTkFont(size=10, weight="bold"), anchor="w",
            ).grid(row=0, column=column, padx=5, pady=(4, 7), sticky="ew")
        self.mc_employee_scroll.grid_columnconfigure(0, weight=1)

        for index, record in enumerate(records, start=1):
            row_number = int(record.get("row", index))
            employee = str(record.get("employee", "") or "").strip()
            company = str(record.get("company", "") or "").strip()
            orders = str(record.get("order_numbers", "") or "").strip()
            primary = employee or company or "Unnamed employee"
            secondary_parts = []
            if company and company.casefold() != primary.casefold():
                secondary_parts.append(company)
            if orders:
                secondary_parts.append(orders)
            display = primary + ("\n" + " • ".join(secondary_parts) if secondary_parts else "")
            ctk.CTkLabel(
                self.mc_employee_scroll, text=display, text_color=TEXT,
                font=ctk.CTkFont(size=10, weight="bold"), anchor="w", justify="left", wraplength=170,
            ).grid(row=index, column=0, padx=5, pady=5, sticky="w")
            ctk.CTkLabel(
                self.mc_employee_scroll, text=self._money_text(record.get("order_total", 0)),
                text_color=TEXT, font=ctk.CTkFont(size=10), anchor="e",
            ).grid(row=index, column=1, padx=5, pady=5, sticky="e")
            discount = float(record.get("discount", 0) or 0)
            variable = ctk.StringVar(value=f"{discount:.2f}")
            self.employee_discount_vars[row_number] = variable
            entry = ctk.CTkEntry(
                self.mc_employee_scroll, textvariable=variable, width=82, height=30,
                border_color="#D9D0E6", fg_color=WHITE, justify="right",
            )
            entry.grid(row=index, column=2, padx=5, pady=5, sticky="e")
            total_label = ctk.CTkLabel(
                self.mc_employee_scroll, text=self._money_text(record.get("total", 0)),
                text_color=PURPLE_DARK, font=ctk.CTkFont(size=10, weight="bold"), anchor="e",
            )
            total_label.grid(row=index, column=3, padx=5, pady=5, sticky="e")
            self.employee_total_labels[row_number] = total_label
            variable.trace_add("write", lambda *_args, row=row_number: self._update_employee_total_row(row))

        self._refresh_employee_grand_total()

    def _update_employee_total_row(self, row_number: int):
        record = self.employee_records_by_row.get(row_number)
        variable = self.employee_discount_vars.get(row_number)
        label = self.employee_total_labels.get(row_number)
        if not record or variable is None or label is None:
            return
        try:
            discount = self._discount_value(variable.get())
            total = float(record.get("order_total", 0) or 0) - discount
            label.configure(text=self._money_text(total), text_color=PURPLE_DARK)
        except (TypeError, ValueError):
            label.configure(text="Invalid", text_color=DANGER)
        self._refresh_employee_grand_total()

    def _refresh_employee_grand_total(self):
        grand_total = 0.0
        valid = True
        for row_number, record in self.employee_records_by_row.items():
            variable = self.employee_discount_vars.get(row_number)
            try:
                discount = self._discount_value(variable.get() if variable else 0)
                grand_total += float(record.get("order_total", 0) or 0) - discount
            except (TypeError, ValueError):
                valid = False
        if valid:
            self.mc_employee_grand_total.configure(
                text=f"Grand Total  {self._money_text(grand_total)}", text_color=PURPLE_DARK,
            )
        else:
            self.mc_employee_grand_total.configure(text="Grand Total  Check discounts", text_color=DANGER)

    def save_dashboard_employee_totals(self):
        if not self.last_review_workbook or not self.last_review_workbook.exists():
            messagebox.showinfo("No Purchase Review", "Create a Purchase Review first.")
            return
        discounts: dict[int, float] = {}
        try:
            for row_number, variable in self.employee_discount_vars.items():
                discounts[row_number] = self._discount_value(variable.get())
        except ValueError as error:
            messagebox.showwarning("Invalid Discount", str(error) or "Enter a valid discount amount.")
            return
        try:
            backup = save_employee_discounts(self.last_review_workbook, discounts)
        except PermissionError:
            messagebox.showwarning(
                "Close the Excel Workbook",
                "Excel is currently using this Purchase Review. Close the workbook, then click Save Totals again.",
            )
            return
        except Exception as error:
            messagebox.showerror("Employee Totals Not Saved", str(error))
            return
        self.refresh_dashboard()
        messagebox.showinfo(
            "Employee Totals Saved",
            f"Saved {len(discounts)} employee discount adjustment(s).\n\nA backup was created at:\n{backup}",
        )

    def _employee_totals_pdf_path(self):
        from modules.final_po_generator import generate_employee_totals_pdf

        if not self.last_review_workbook or not self.last_review_workbook.exists():
            raise FileNotFoundError("Create a Uniform Sizing Event purchase packet first.")
        # Save onscreen discounts first so the shared PDF is current.
        if self.employee_discount_vars:
            discounts = {row: self._discount_value(var.get()) for row, var in self.employee_discount_vars.items()}
            save_employee_discounts(self.last_review_workbook, discounts)
        return generate_employee_totals_pdf(self.last_review_workbook)

    def view_employee_totals_pdf(self):
        try:
            path = self._employee_totals_pdf_path()
            subprocess.run(["open", str(path)], check=False)
        except Exception as error:
            messagebox.showerror("Unable to View Employee Totals", str(error))

    def share_employee_totals_pdf(self):
        try:
            path = self._employee_totals_pdf_path()
            subject = f"Employee Totals — {self.current_event_name or 'Uniform Sizing Event'}"
            script = '''
            on run argv
                set attachmentFile to POSIX file (item 1 of argv)
                set messageSubject to item 2 of argv
                tell application "Mail"
                    activate
                    set newMessage to make new outgoing message with properties {subject:messageSubject, visible:true}
                    tell content of newMessage to make new attachment with properties {file name:attachmentFile} at after the last paragraph
                end tell
            end run
            '''
            subprocess.run(["osascript", "-e", script, str(path), subject], check=True)
        except Exception as error:
            messagebox.showerror("Unable to Share Employee Totals", str(error))

    def _outsourced_job_logo_path(self):
        if not self.last_review_workbook:
            return None
        stem = self.last_review_workbook.stem
        event_key = safe_filename(self.current_event_name or "Current Event")
        for ext in (".png", ".jpg", ".jpeg", ".webp"):
            stable = self.last_review_workbook.parent / f"{event_key}__Outsourced_Job_Logo{ext}"
            if stable.exists():
                return stable
            candidate = self.last_review_workbook.parent / f"{stem}__Outsourced_Job_Logo{ext}"
            if candidate.exists():
                return candidate
        return None

    def build_outsourced_job_logo_page(self):
        page = self.new_page("job_logo")
        card = ctk.CTkFrame(page, fg_color=WHITE, border_width=1, border_color=BORDER, corner_radius=18)
        card.grid(row=0, column=0, sticky="nsew", padx=34, pady=30)
        card.grid_columnconfigure(0, weight=1)
        page.grid_rowconfigure(0, weight=1)
        ctk.CTkLabel(
            card, text="Job Report Logo", text_color=PURPLE_DARK,
            font=ctk.CTkFont(size=28, weight="bold"),
        ).grid(row=0, column=0, padx=30, pady=(34, 6))
        ctk.CTkLabel(
            card,
            text="Upload the customer or job logo that should appear on the Outsourced Decoration Job Report, including standard screen-print workflows.",
            text_color=MUTED, font=ctk.CTkFont(size=14), wraplength=720, justify="center",
        ).grid(row=1, column=0, padx=30, pady=(0, 24))
        self.outsourced_logo_preview = ctk.CTkLabel(
            card, text="No job logo uploaded", width=420, height=230, corner_radius=16,
            fg_color="#F7F3FC", text_color=MUTED, font=ctk.CTkFont(size=16, weight="bold"),
        )
        self.outsourced_logo_preview.grid(row=2, column=0, padx=30, pady=10)
        self.outsourced_logo_status = ctk.CTkLabel(
            card, text="", text_color=TEXT, font=ctk.CTkFont(size=13, weight="bold"),
        )
        self.outsourced_logo_status.grid(row=3, column=0, padx=30, pady=(4, 16))
        buttons = ctk.CTkFrame(card, fg_color="transparent")
        buttons.grid(row=4, column=0, pady=(0, 34))
        self.outsourced_logo_upload_button = ctk.CTkButton(
            buttons, text="Upload / Replace Logo", command=lambda: self.manage_outsourced_job_logo("upload"),
            width=210, height=48, fg_color=PURPLE, hover_color=PURPLE_DARK,
            font=ctk.CTkFont(size=14, weight="bold"),
        )
        self.outsourced_logo_upload_button.pack(side="left", padx=6)
        self.outsourced_logo_preview_button = ctk.CTkButton(
            buttons, text="Open Logo", command=lambda: self.manage_outsourced_job_logo("preview"),
            width=145, height=48, fg_color=PURPLE_LIGHT, hover_color="#E6D9F5",
            border_width=1, border_color="#BFA9DD", text_color=PURPLE_DARK,
        )
        self.outsourced_logo_preview_button.pack(side="left", padx=6)
        self.outsourced_logo_remove_button = ctk.CTkButton(
            buttons, text="Remove Logo", command=lambda: self.manage_outsourced_job_logo("remove"),
            width=145, height=48, fg_color="transparent", hover_color=PURPLE_LIGHT,
            border_width=1, border_color="#BFA9DD", text_color=PURPLE_DARK,
        )
        self.outsourced_logo_remove_button.pack(side="left", padx=6)
        report_buttons = ctk.CTkFrame(card, fg_color="transparent")
        report_buttons.grid(row=5, column=0, pady=(0, 28))
        self.outsourced_logo_regenerate_button = ctk.CTkButton(
            report_buttons, text="Regenerate Purchase Review & Reports",
            command=self.regenerate_review_and_reports,
            width=300, height=46, fg_color=PURPLE, hover_color=PURPLE_DARK,
            font=ctk.CTkFont(size=13, weight="bold"),
        )
        self.outsourced_logo_regenerate_button.pack(side="left", padx=6)
        self.outsourced_logo_open_report_button = ctk.CTkButton(
            report_buttons, text="Open Outsourced Job Report",
            command=lambda: self._open_generated_report("outsourced"),
            width=235, height=46, fg_color=PURPLE_LIGHT, hover_color="#E6D9F5",
            border_width=1, border_color="#BFA9DD", text_color=PURPLE_DARK,
            font=ctk.CTkFont(size=13, weight="bold"),
        )
        self.outsourced_logo_open_report_button.pack(side="left", padx=6)

    def refresh_outsourced_job_logo_page(self):
        if not hasattr(self, "outsourced_logo_status"):
            return
        applicable = bool(self.last_review_workbook and self.last_review_workbook.exists())
        current = self._outsourced_job_logo_path() if applicable else None
        state = "normal" if applicable else "disabled"
        self.outsourced_logo_upload_button.configure(state=state)
        self.outsourced_logo_regenerate_button.configure(state=state)
        if not applicable:
            self.outsourced_logo_preview.configure(image=None, text="Create a Purchase Review first")
            self.outsourced_logo_status.configure(text="Import orders and create the Purchase Review before uploading a logo.", text_color=WARNING)
            self.outsourced_logo_preview_button.configure(state="disabled")
            self.outsourced_logo_remove_button.configure(state="disabled")
            return
        if not current:
            self.outsourced_logo_preview.configure(image=None, text="No job logo uploaded")
            self.outsourced_logo_status.configure(text="Choose a PNG, JPG, JPEG, or WEBP image.", text_color=MUTED)
            self.outsourced_logo_preview_button.configure(state="disabled")
            self.outsourced_logo_remove_button.configure(state="disabled")
            return
        try:
            image = PILImage.open(current).convert("RGBA")
            image.thumbnail((390, 200), PILImage.Resampling.LANCZOS)
            self._outsourced_logo_preview_image = ctk.CTkImage(
                light_image=image, dark_image=image, size=image.size,
            )
            self.outsourced_logo_preview.configure(image=self._outsourced_logo_preview_image, text="")
        except Exception:
            self.outsourced_logo_preview.configure(image=None, text=current.name)
        self.outsourced_logo_status.configure(text=f"Uploaded: {current.name}", text_color=SUCCESS)
        self.outsourced_logo_preview_button.configure(state="normal")
        self.outsourced_logo_remove_button.configure(state="normal")

    def manage_outsourced_job_logo(self, action: str = "upload"):
        if not self.last_review_workbook or not self.last_review_workbook.exists():
            messagebox.showinfo("Job Report Logo", "Create the Purchase Review before uploading the event logo.")
            return
        current = self._outsourced_job_logo_path()
        if action == "preview":
            if current:
                subprocess.run(["open", str(current)], check=False)
            return
        if action == "remove":
            if not current:
                return
            if not messagebox.askyesno("Remove Outsourced Job Logo?", f"Remove {current.name} from this event?"):
                return
            current.unlink(missing_ok=True)
            self.refresh_outsourced_job_logo_page()
            return

        # Opening a native macOS file dialog synchronously from a CTkButton press
        # can leave the mouse press captured by that button. The next click can then
        # be delivered back to Upload / Replace Logo, reopening the chooser even when
        # the user clicked Current Event or another navigation control. Let the button
        # release finish first, then open exactly one chooser.
        if bool(getattr(self, "_outsourced_logo_dialog_pending", False)):
            return
        self._outsourced_logo_dialog_pending = True
        try:
            self.outsourced_logo_upload_button.configure(state="disabled")
        except Exception:
            pass
        self.after(180, self._open_outsourced_job_logo_dialog)

    def _open_outsourced_job_logo_dialog(self):
        try:
            selected = filedialog.askopenfilename(
                parent=self,
                title="Choose Job Report Logo",
                filetypes=[
                    ("Logo images", "*.png *.jpg *.jpeg *.webp"),
                    ("PNG images", "*.png"), ("JPEG images", "*.jpg *.jpeg"),
                ],
            )
            if not selected:
                return
            source = Path(selected)
            if source.suffix.casefold() not in {".png", ".jpg", ".jpeg", ".webp"}:
                messagebox.showwarning(
                    "Unsupported Logo File",
                    "Choose a PNG, JPG, JPEG, or WEBP image. Embroidery machine files such as EXP cannot be placed on the PDF report.",
                )
                return
            event_key = safe_filename(self.current_event_name or "Current Event")
            for ext in (".png", ".jpg", ".jpeg", ".webp"):
                old = self.last_review_workbook.parent / f"{self.last_review_workbook.stem}__Outsourced_Job_Logo{ext}"
                old.unlink(missing_ok=True)
                stable_old = self.last_review_workbook.parent / f"{event_key}__Outsourced_Job_Logo{ext}"
                stable_old.unlink(missing_ok=True)
            target = self.last_review_workbook.parent / f"{event_key}__Outsourced_Job_Logo{source.suffix.lower()}"
            shutil.copy2(source, target)
            self.dashboard_status.configure(text=f"Outsourced job logo saved: {target.name}", text_color=SUCCESS)
        finally:
            self._outsourced_logo_dialog_pending = False
            # Refresh after the native dialog fully releases its modal mouse grab.
            try:
                self.after(120, self.refresh_outsourced_job_logo_page)
            except Exception:
                self.refresh_outsourced_job_logo_page()

    def _render_po_manager(self, snapshot: dict):
        self._clear_children(self.mc_po_scroll)
        self.po_entry_vars = {}
        routes = snapshot.get("routes", [])
        if not routes:
            self.mc_po_summary.configure(text="Create a Purchase Review to manage PO numbers here.")
            ctk.CTkLabel(self.mc_po_scroll, text="No purchase-order routes available yet.", text_color=MUTED,
                         font=ctk.CTkFont(size=12)).grid(row=0, column=0, padx=12, pady=22)
            self.mc_generate_button.configure(state="disabled")
            return
        ready = sum(1 for row in routes if str(row.get("status", "")).casefold() == "ready")
        self.mc_po_summary.configure(text=f"{ready} ready  •  {len(routes) - ready} need review  •  Edit PO numbers directly below.")
        headers = ["Vendor", "Report", "PO Number", "Status"]
        for col, header in enumerate(headers):
            ctk.CTkLabel(self.mc_po_scroll, text=header, text_color=PURPLE_DARK,
                         font=ctk.CTkFont(size=11, weight="bold"), anchor="w").grid(
                row=0, column=col, padx=8, pady=(5, 7), sticky="ew")
        for row_index, route in enumerate(routes, start=1):
            vendor = route.get("vendor", "")
            report = route.get("report_type", "")
            key = route.get("key", "")
            var = ctk.StringVar(value=route.get("po_number", ""))
            self.po_entry_vars[key] = (vendor, report, var)
            ctk.CTkLabel(self.mc_po_scroll, text=vendor, text_color=TEXT,
                         font=ctk.CTkFont(size=11, weight="bold"), anchor="w").grid(
                row=row_index, column=0, padx=8, pady=4, sticky="ew")
            ctk.CTkLabel(self.mc_po_scroll, text=report, text_color=MUTED,
                         font=ctk.CTkFont(size=10), anchor="w").grid(
                row=row_index, column=1, padx=8, pady=4, sticky="ew")
            ctk.CTkEntry(self.mc_po_scroll, textvariable=var, height=32,
                         border_color="#D9D0E6", fg_color=WHITE).grid(
                row=row_index, column=2, padx=8, pady=4, sticky="ew")
            status = route.get("status", "Ready")
            is_ready = str(status).casefold() == "ready"
            ctk.CTkLabel(
                self.mc_po_scroll, text=("✓  " if is_ready else "⚠  ") + status,
                text_color="#166534" if is_ready else "#8A3F00",
                fg_color="#EAF8EF" if is_ready else "#FFF1DC",
                height=28, corner_radius=14,
                font=ctk.CTkFont(size=11, weight="bold"), anchor="center",
            ).grid(row=row_index, column=3, padx=8, pady=4, sticky="ew")
        blockers = int(snapshot.get("review_count", 0))
        blocked_routes = sum(1 for row in routes if str(row.get("status", "")).casefold() != "ready")
        missing_po = any(not str(row.get("po_number", "")).strip() for row in routes)
        self.mc_generate_button.configure(
            state="normal" if blockers == 0 and blocked_routes == 0 and not missing_po else "disabled"
        )

    def refresh_dashboard(self):
        if self.last_review_workbook and self.last_review_workbook.exists():
            self._sync_current_product_candidates()
        health = _product_master_health()
        snapshot = self._mission_snapshot()
        self.last_snapshot = snapshot
        if self.last_review_workbook and self.last_review_workbook.exists():
            self.review_needed = int(snapshot.get("review_count", 0))
            self.imported_line_count = int(snapshot.get("line_count", self.imported_line_count or 0))
            self.current_event_name = snapshot.get("event_name") or self.current_event_name or "General Sales Period"
            self.current_mode = snapshot.get("report_mode") or self.current_mode

        refresh_key = (
            self._mission_snapshot_cache_key,
            _file_stamp(live_product_master_path()),
            _file_stamp(self.selected_csv),
            _file_stamp(self.last_purchase_order_dir),
            self.current_event_name,
            self.current_mode,
            self.current_decoration_fulfillment,
        )
        if refresh_key == self._dashboard_refresh_key:
            self._apply_dashboard_view()
            return

        review_count = int(snapshot.get("review_count", 0))
        product_master_count = int(snapshot.get("product_master_count", 0))
        purchase_review_count = int(snapshot.get("purchase_review_count", review_count))
        routes = list(snapshot.get("routes", []))
        blocked_routes = [
            route for route in routes
            if str(route.get("status", "")).casefold() != "ready"
        ]
        blocked_route_count = len(blocked_routes)
        workflow_blocked = bool(review_count or blocked_route_count)
        missing_routes = [r for r in routes if not str(r.get("po_number", "")).strip()]
        entered = len(routes) - len(missing_routes)
        all_po_ready = bool(routes) and not workflow_blocked
        line_count = int(snapshot.get("line_count", 0) or 0)
        has_active_packet = bool(line_count > 0 or routes or review_count or product_master_count or purchase_review_count)
        has_active_review = bool(self.last_review_workbook and has_active_packet)
        # While permanent setup decisions remain, Step 2 owns the workflow.
        # A review is considered stale only after those setup items are complete.
        review_stale = bool(has_active_review and product_master_count == 0 and self._review_is_stale())
        purchase_generated_for_current = bool(
            self.last_purchase_order_dir
            and self.last_purchase_order_dir.exists()
            and self.last_review_workbook
            and self.last_review_workbook.exists()
            and self.last_purchase_order_dir.stat().st_mtime >= self.last_review_workbook.stat().st_mtime
        )

        import_complete = bool(has_active_review)
        master_complete = bool(has_active_review and product_master_count == 0)
        review_complete = bool(has_active_review and product_master_count == 0 and purchase_review_count == 0 and blocked_route_count == 0)
        po_complete = bool(review_complete and purchase_generated_for_current)
        self._set_stage(0, "complete" if import_complete else "active")
        if not import_complete:
            self._set_stage(1, "pending")
            self._set_stage(2, "pending")
            self._set_stage(3, "pending")
        else:
            self._set_stage(1, "attention" if product_master_count else ("complete" if master_complete else "active"))
            self._set_stage(2, "attention" if purchase_review_count or blocked_route_count else ("complete" if review_complete else "pending"))
            self._set_stage(3, "complete" if po_complete else ("active" if review_complete else "pending"))

        # Product Master owns reusable vendor and decoration setup.
        if product_master_count:
            self._set_pill(
                self.mc_master_status,
                f"{product_master_count} NEW STYLE{'S' if product_master_count != 1 else ''} NEED SETUP",
                "attention",
            )
            self.mc_master_detail.configure(
                text="Complete these new style details once. Orchid will refresh the review when Product Master closes."
            )
        elif review_stale:
            self._set_pill(self.mc_master_status, "UPDATES NOT APPLIED", "attention")
            self.mc_master_detail.configure(text="Product Master changed after this Purchase Review was created")
        else:
            self._set_pill(self.mc_master_status, f"{health['styles']} PRODUCTS CONFIGURED", "complete")
            self.mc_master_detail.configure(text="Permanent vendor and decoration catalog")

        # Purchase Review card is reserved for order-specific decisions.
        if product_master_count:
            self._set_pill(self.mc_review_status, "WAITING FOR PRODUCT MASTER", "pending")
            self.mc_review_detail.configure(text="Step 3 unlocks after permanent product setup is complete")
        elif review_stale:
            self._set_pill(self.mc_review_status, "REFRESH REQUIRED", "attention")
            self.mc_review_detail.configure(text="Regenerate Purchase Review to apply the latest Product Master changes")
        elif not has_active_review:
            self._set_pill(self.mc_review_status, "NOT CREATED", "pending")
            self.mc_review_detail.configure(text="Process the imported orders first")
        elif purchase_review_count:
            self._set_pill(self.mc_review_status, f"{purchase_review_count} ORDER DECISION{'S' if purchase_review_count != 1 else ''} REQUIRED", "attention")
            self.mc_review_detail.configure(
                text=f"{purchase_review_count} order-specific correction{'s' if purchase_review_count != 1 else ''} remain"
            )
        elif blocked_route_count:
            self._set_pill(
                self.mc_review_status,
                f"WAITING ON {blocked_route_count} DECISION{'S' if blocked_route_count != 1 else ''}",
                "attention",
            )
            self.mc_review_detail.configure(text="Open the Work Queue to see where each route must be corrected")
        else:
            self._set_pill(self.mc_review_status, "COMPLETE", "complete")
            self.mc_review_detail.configure(text="No purchasing decisions or blocked routes remain")

        # Purchase Orders card: a simple status and one edit/generate button.
        if product_master_count:
            self._set_pill(self.mc_po_status, "WAITING FOR PRODUCT MASTER", "pending")
            self.mc_po_detail.configure(text="Complete Step 2 before purchase orders can continue")
            self.mc_po_card_button.configure(text="Set Up Products", command=lambda: self.show_page("master"), state="normal")
        elif review_stale:
            self._set_pill(self.mc_po_status, "WAITING FOR REFRESH", "pending")
            self.mc_po_detail.configure(text="Apply Product Master updates before continuing")
            self.mc_po_card_button.configure(text="Regenerate Review", command=self.regenerate_current_review, state="normal")
        elif not routes:
            self._set_pill(self.mc_po_status, "NOT STARTED", "pending")
            self.mc_po_detail.configure(text="Purchase-order routes appear after review is created")
            self.mc_po_card_button.configure(text="Edit PO Numbers", command=self.open_po_number_editor, state="disabled")
        elif blocked_route_count:
            self._set_pill(
                self.mc_po_status,
                f"WAITING ON {blocked_route_count} DECISION{'S' if blocked_route_count != 1 else ''}",
                "attention",
            )
            self.mc_po_detail.configure(text="Resolve the remaining Product Master or order-specific decisions")
            self.mc_po_card_button.configure(text="View Blocked Routes", command=lambda: self.show_page("review"), state="normal")
        else:
            self._set_pill(self.mc_po_status, "READY", "complete")
            self.mc_po_detail.configure(text=f"{len(routes)} vendor purchase order{'s' if len(routes) != 1 else ''} ready")
            self.mc_po_card_button.configure(text="Generate Purchase Orders", command=self.generate_latest_purchase_orders, state="normal")

        # Centered hero and one primary workflow action.
        active_card = None
        completed = False
        if not has_active_packet and not self.selected_csv:
            step_number = 1
            title = "Import Orders"
            detail = "Choose the Shopify or Report Toaster export to start a new purchase packet."
            button_text = "Import Orders"
            command = lambda: self.show_page("import")
        elif not has_active_review:
            step_number = 1
            title = "Process Imported Orders"
            detail = "The order CSV is selected. Process the import so Orchid can check Product Master before Purchase Review."
            button_text = "Process Imported Orders"
            command = self.create_review_workbook
        elif product_master_count:
            step_number = 2
            title = "Complete Product Setup"
            detail = f"{product_master_count} new style{'s' if product_master_count != 1 else ''} need permanent setup before continuing."
            button_text = "Set Up Next Product"
            command = self.open_next_product_setup
            active_card = "master"
        elif review_stale:
            step_number = 2
            title = "Apply Product Master Updates"
            detail = "Regenerate Purchase Review to apply the latest routing and decoration changes."
            button_text = "Regenerate Purchase Review"
            command = self.regenerate_current_review
            active_card = "master"
        elif purchase_review_count or blocked_route_count:
            step_number = 3
            title = "Complete Purchase Review"
            if purchase_review_count:
                detail = f"{purchase_review_count} order-specific decision{'s' if purchase_review_count != 1 else ''} still require attention."
            else:
                detail = f"{blocked_route_count} purchase-order route{'s are' if blocked_route_count != 1 else ' is'} still blocked."
            button_text = "Open Purchase Review"
            command = lambda: self.show_page("review")
            active_card = "review"
        elif not purchase_generated_for_current:
            step_number = 4
            title = "Generate Purchase Orders"
            detail = "Purchase Review is complete. Enter one PO number for each vendor, then create the final PDFs."
            button_text = "Generate Purchase Orders"
            command = self.generate_latest_purchase_orders
            active_card = "purchase"
        else:
            step_number = 4
            completed = True
            title = "Purchase Orders Generated Successfully"
            detail = "The final vendor PDFs have been created and the event has been archived."
            button_text = "Open Purchase Orders"
            command = self.open_latest_purchase_orders
            active_card = "purchase"

        self._highlight_mission_card(active_card)
        self.mc_step_label.configure(text="")
        hero_icons = {1: "↑", 2: "◇", 3: "☑", 4: "▤"}
        processing_import = bool(step_number == 1 and self.selected_csv and not has_active_review)
        hero_icon = "↻" if processing_import else hero_icons.get(step_number, "◆")
        self.mc_banner_icon.configure(text="✓" if completed else hero_icon)
        hero_colors = {
            1: ("#2563EB", "#EAF2FF"),
            2: (PURPLE, PURPLE_LIGHT),
            3: ("#B85C00", "#FFF1DC"),
            4: (SUCCESS, "#EAF8EF"),
        }
        hero_color, hero_fill = hero_colors[step_number]
        if completed:
            hero_color, hero_fill = "#166534", "#EAF8EF"
        self.mc_step_label.configure(text_color=hero_color)
        self.mc_banner_icon.configure(fg_color=hero_fill, text_color=hero_color)
        self.mc_banner_title.configure(text=title)
        self.mc_banner_subtitle.configure(text=detail, text_color=MUTED)
        event_name = self.current_event_name or snapshot.get("event_name") or ""
        line_count_text = int(snapshot.get("line_count", self.imported_line_count or 0) or 0)
        if event_name:
            packet_line = event_name
            if line_count_text and not completed:
                packet_line += f"  •  {line_count_text:,} purchase lines"
            self.mc_event_label.configure(text=packet_line)
            self.mc_event_label.grid()
        elif self.selected_csv:
            self.mc_event_label.configure(text=self.selected_csv.name)
            self.mc_event_label.grid()
        else:
            self.mc_event_label.grid_remove()
        self.mc_banner_button.configure(text=button_text, command=command, state="normal")
        if completed:
            self.mc_banner_secondary_button.pack(side="left", padx=7)
            if hasattr(self, "mc_animation_controls"):
                self.mc_animation_controls.place(relx=1.0, rely=0.05, anchor="ne")
        else:
            self.mc_banner_secondary_button.pack_forget()
            if hasattr(self, "mc_animation_controls"):
                self.mc_animation_controls.place_forget()
            self._success_animation_skip_requested = True
            self._finish_success_animation()
        self._update_success_control_buttons()
        self._update_dashboard_metrics(snapshot, show=completed)
        # Avoid repeating the same primary action in multiple dashboard locations.
        if product_master_count:
            self.mc_master_button.grid_remove()
            self.mc_work_open_button.configure(text="Set Up Next Product", command=self.open_next_product_setup)
        elif review_stale:
            self.mc_master_button.grid()
            self.mc_master_button.configure(text="Open Product Master", command=self.open_product_master, state="normal")
            self.mc_work_open_button.configure(text="Regenerate Review", command=self.regenerate_current_review)
        else:
            self.mc_master_button.grid()
            if purchase_review_count or blocked_route_count:
                self.mc_work_open_button.configure(text="Open Purchase Review", command=self.open_latest_review)
            else:
                self.mc_work_open_button.configure(text="Open Review", command=self.open_latest_review)
        if purchase_review_count or (blocked_route_count and not product_master_count):
            self.mc_review_button.grid_remove()
        else:
            self.mc_review_button.grid()
        self.mc_action_title.configure(text=title.title())
        self.mc_action_detail.configure(text=detail)
        self.mc_action_button.configure(text=button_text, command=command, state="normal")

        show_employee_totals = self._configure_dashboard_workspace(snapshot)
        self._render_work_queue(snapshot)
        if show_employee_totals and hasattr(self, "mc_employee_scroll"):
            self._render_employee_totals(snapshot)
        self._refresh_management_dashboard(snapshot, purchase_generated_for_current)
        self._apply_dashboard_view()

        selected_text = self.selected_csv.name if self.selected_csv else "No order CSV selected"
        self.refresh_master_page()
        if hasattr(self, "import_file_label"):
            self.import_file_label.configure(text=selected_text)
        self.refresh_review_page()
        self.refresh_purchase_page()
        self.refresh_current_event_page(snapshot)
        self._dashboard_refresh_key = refresh_key

    def open_product_master_search(self):
        value = " ".join(self.mc_product_search_var.get().split()) if hasattr(self, "mc_product_search_var") else ""
        self.open_product_master(search_product=value)

    def open_po_number_editor(
        self,
        generate_after_save: bool = False,
        workbook_path: Path | None = None,
    ):
        workbook = Path(workbook_path) if workbook_path else self.last_review_workbook
        if not workbook or not workbook.exists():
            messagebox.showinfo("No Purchase Review", "Create a Purchase Review first.")
            return
        self.last_review_workbook = workbook
        snapshot = load_mission_control_snapshot(workbook, DATA)
        routes = list(snapshot.get("routes", []))
        if not routes:
            messagebox.showinfo("No Purchase Orders", "No vendor purchase orders are available yet.")
            return

        routes.sort(key=lambda route: str(route.get("vendor", "")).casefold())
        window = ctk.CTkToplevel(self)
        window.title("Assign Purchase Order Numbers" if generate_after_save else "Edit Purchase Order Numbers")
        window.geometry("700x600")
        window.minsize(620, 480)
        window.transient(self)
        window.grab_set()
        window.grid_columnconfigure(0, weight=1)
        window.grid_rowconfigure(2, weight=1)

        heading = "Assign Purchase Order Numbers" if generate_after_save else "Edit Purchase Order Numbers"
        ctk.CTkLabel(
            window, text=heading, text_color=TEXT,
            font=ctk.CTkFont(size=23, weight="bold"), anchor="w",
        ).grid(row=0, column=0, padx=24, pady=(22, 4), sticky="ew")
        subtitle = (
            "Enter one PO number for each vendor. These numbers are saved with this Purchase Review "
            "and printed on every page of the vendor PDF."
        )
        ctk.CTkLabel(
            window, text=subtitle, text_color=MUTED,
            font=ctk.CTkFont(size=13), anchor="w", justify="left", wraplength=620,
        ).grid(row=1, column=0, padx=24, pady=(0, 12), sticky="ew")

        scroll = ctk.CTkScrollableFrame(window, fg_color="#FCFBFE", corner_radius=12)
        scroll.grid(row=2, column=0, padx=24, pady=(0, 14), sticky="nsew")
        register_scrollable(self, scroll)
        scroll.grid_columnconfigure(1, weight=1)
        for col, label in enumerate(("Vendor", "PO Number")):
            ctk.CTkLabel(
                scroll, text=label, text_color=PURPLE_DARK,
                font=ctk.CTkFont(size=12, weight="bold"), anchor="w",
            ).grid(row=0, column=col, padx=12, pady=(10, 8), sticky="ew")

        editor_vars = []
        first_missing_entry = None
        for idx, route in enumerate(routes, start=1):
            vendor = str(route.get("vendor", "")).strip()
            original = str(route.get("po_number", "")).strip()
            variable = ctk.StringVar(value=original)
            editor_vars.append((vendor, variable))
            missing = not original
            row_fill = "#FFF8EC" if missing else "transparent"
            ctk.CTkLabel(
                scroll, text=vendor, text_color=TEXT, fg_color=row_fill,
                font=ctk.CTkFont(size=12, weight="bold"), anchor="w",
            ).grid(row=idx, column=0, padx=12, pady=6, sticky="ew")
            entry = ctk.CTkEntry(
                scroll, textvariable=variable, height=38,
                placeholder_text="Enter PO number",
                border_color=WARNING if missing else "#D9D0E6", fg_color=WHITE,
            )
            entry.grid(row=idx, column=1, padx=12, pady=6, sticky="ew")
            if missing and first_missing_entry is None:
                first_missing_entry = entry

        button_bar = ctk.CTkFrame(window, fg_color="transparent")
        button_bar.grid(row=3, column=0, padx=24, pady=(0, 22), sticky="e")
        ctk.CTkButton(
            button_bar, text="Cancel", width=110, height=42,
            fg_color=PURPLE_LIGHT, hover_color="#E4D6F5", text_color=PURPLE_DARK,
            command=window.destroy,
        ).pack(side="left", padx=4)

        def save_numbers():
            values: dict[tuple[str, str], str] = {}
            missing_vendors = []
            for vendor, variable in editor_vars:
                number = " ".join(variable.get().split())
                if not number:
                    missing_vendors.append(vendor)
                else:
                    values[(vendor.casefold(), "combined vendor order")] = number

            if generate_after_save and missing_vendors:
                messagebox.showwarning(
                    "PO Numbers Required",
                    "Enter a PO number for every vendor before generating purchase orders:\n\n"
                    + "\n".join(missing_vendors),
                    parent=window,
                )
                return

            save_po_overrides(DATA, workbook, values)
            window.destroy()
            self.refresh_dashboard()
            if generate_after_save:
                self._run_final_purchase_order_generation(workbook, values)
            elif missing_vendors:
                messagebox.showwarning(
                    "PO Numbers Saved",
                    "The entered PO numbers were saved. Still missing:\n\n"
                    + "\n".join(missing_vendors),
                )
            else:
                messagebox.showinfo("PO Numbers Saved", "The vendor PO numbers were saved with this Purchase Review.")

        button_text = "Save & Generate Purchase Orders" if generate_after_save else "Save PO Numbers"
        ctk.CTkButton(
            button_bar, text=button_text, width=245 if generate_after_save else 170, height=42,
            fg_color=PURPLE, hover_color=PURPLE_DARK,
            font=ctk.CTkFont(size=13, weight="bold"),
            command=save_numbers,
        ).pack(side="left", padx=4)
        if first_missing_entry is not None:
            window.after(250, first_missing_entry.focus_set)

    def save_dashboard_po_numbers(self, quiet: bool = False) -> dict[tuple[str, str], str]:
        if not self.last_review_workbook or not self.last_review_workbook.exists():
            if not quiet:
                messagebox.showinfo("No Purchase Review", "Create a Purchase Review first.")
            return {}
        values: dict[tuple[str, str], str] = {}
        for _key, (vendor, report, variable) in self.po_entry_vars.items():
            number = " ".join(variable.get().split())
            if number:
                values[(vendor.casefold(), canonical_report_type(report))] = number
        save_po_overrides(DATA, self.last_review_workbook, values)
        if not quiet:
            self.dashboard_status.configure(text=f"Saved {len(values)} PO number(s) in Orchid Purchase Manager.", text_color=SUCCESS)
            messagebox.showinfo("PO Numbers Saved", f"Saved {len(values)} PO number(s) for the latest Purchase Review.")
        return values

    def generate_latest_purchase_orders(self):
        if not self.last_review_workbook or not self.last_review_workbook.exists():
            messagebox.showinfo("No Purchase Review", "Create a Purchase Review first.")
            return
        snapshot = self._fresh_snapshot()
        review_count = int(snapshot.get("review_count", 0))
        blocked_route_count = int(snapshot.get("blocked_route_count", 0))
        if review_count or blocked_route_count:
            summary = self._blocker_summary(snapshot)
            messagebox.showwarning(
                "Purchase Review Incomplete",
                "Purchase Review and Step 4 did not pass the final readiness check.\n\n"
                + (summary or f"{review_count} decision(s) and {blocked_route_count} blocked route(s) remain."),
            )
            self.show_page("review")
            return
        self.open_po_number_editor(generate_after_save=True)

    def regenerate_review_and_reports(self):
        """Refresh the active review, then rebuild reports with the saved event logo."""
        if not self.last_review_workbook or not self.last_review_workbook.exists():
            messagebox.showinfo("No Purchase Review", "Create a Purchase Review first.")
            return
        self.regenerate_current_review(silent=True)
        if not self.last_review_workbook or not self.last_review_workbook.exists():
            return
        snapshot = self._fresh_snapshot()
        review_count = int(snapshot.get("review_count", 0))
        blocked_route_count = int(snapshot.get("blocked_route_count", 0))
        if review_count or blocked_route_count:
            messagebox.showwarning(
                "Purchase Review Requires Attention",
                "The Purchase Review was regenerated, but reports cannot be rebuilt until the remaining decisions are completed.\n\n"
                + (self._blocker_summary(snapshot) or f"{review_count} decision(s) and {blocked_route_count} blocked route(s) remain."),
            )
            self.show_page("review")
            return
        overrides = load_po_overrides(DATA, self.last_review_workbook)
        self._run_final_purchase_order_generation(self.last_review_workbook, overrides)

    def guided_stage_card(self, page, step_label: str, title: str, body: str):
        """Create the large white action card used throughout the workflow."""
        page.grid_rowconfigure(2, weight=1)
        shadow = ctk.CTkFrame(page, fg_color=SHADOW, corner_radius=24, height=450)
        shadow.grid(row=1, column=0, sticky="ew", padx=48, pady=(8, 18))
        shadow.grid_propagate(False)
        frame = ctk.CTkFrame(
            page, fg_color=WHITE, border_width=1, border_color=BORDER,
            corner_radius=24, height=444,
        )
        frame.grid(row=1, column=0, sticky="ew", padx=48, pady=(3, 23))
        frame.grid_propagate(False)
        frame.grid_columnconfigure(0, weight=1)
        self._add_orchid_watermark(frame, size=(520, 520), relx=0.5, rely=0.57)

        badge = ctk.CTkLabel(frame, text="", width=1, height=1)
        badge.grid_remove()
        icon_map = {
            "Complete Product Setup": "◇",
            "Complete Purchase Review": "☑",
            "Generate Purchase Orders": "▤",
        }
        icon = ctk.CTkLabel(
            frame, text=icon_map.get(title, "◆"), width=72, height=72, corner_radius=36,
            fg_color=PURPLE_LIGHT, text_color=PURPLE,
            font=ctk.CTkFont(size=28, weight="bold"),
        )
        icon.grid(row=1, column=0, pady=(29, 13))
        title_label = ctk.CTkLabel(
            frame, text=title, text_color=TEXT,
            font=ctk.CTkFont(size=29, weight="bold"), justify="center", wraplength=900,
        )
        title_label.grid(row=2, column=0, padx=50, pady=(0, 8))
        body_label = ctk.CTkLabel(
            frame, text=body, text_color=MUTED,
            font=ctk.CTkFont(size=15), justify="center", wraplength=840,
        )
        body_label.grid(row=3, column=0, padx=56, pady=(0, 15))
        return frame, badge, title_label, body_label

    def build_import_page(self):
        page = self.new_page("import")
        page.grid_rowconfigure(2, weight=1)

        self._build_workflow_tracker(page, 0, "current_event_stage_labels", padx=52)

        card_shadow = ctk.CTkFrame(page, fg_color=SHADOW, corner_radius=20, width=760, height=382)
        card_shadow.grid(row=1, column=0, sticky="n", padx=48, pady=(4, 18))
        card = ctk.CTkFrame(
            page, fg_color=WHITE, border_width=1, border_color=BORDER,
            corner_radius=20, width=760, height=376,
        )
        card.grid(row=1, column=0, sticky="n", padx=48, pady=(0, 22))
        card.grid_propagate(False)
        card.grid_columnconfigure(0, weight=1)
        self.current_event_card = card
        self.current_event_card_shadow = card_shadow
        self._add_orchid_watermark(card, size=(420, 420), relx=0.5, rely=0.58, opacity=0.035)

        self.current_event_icon = ctk.CTkLabel(
            card, text="↥", width=58, height=58, corner_radius=29,
            fg_color=PURPLE_LIGHT, text_color=PURPLE,
            font=ctk.CTkFont(size=25, weight="bold"),
        )
        self.current_event_icon.grid(row=1, column=0, pady=(26, 10))
        self.current_event_title = ctk.CTkLabel(
            card, text="No Active Purchase Packet", text_color=PURPLE_DARK,
            font=ctk.CTkFont(size=25, weight="bold"), justify="center", wraplength=620,
        )
        self.current_event_title.grid(row=2, column=0, padx=50, pady=(0, 8))
        self.current_event_body = ctk.CTkLabel(
            card, text="Import the Shopify or Report Toaster CSV for the orders you want to process. Product Master and saved settings are preserved.",
            text_color=MUTED, font=ctk.CTkFont(size=14), justify="center", wraplength=650,
        )
        self.current_event_body.grid(row=3, column=0, padx=56, pady=(0, 14))
        self.import_file_label = ctk.CTkLabel(
            card, text="No order CSV selected", text_color=TEXT,
            fg_color="#F3F1F6", corner_radius=14, height=31,
            font=ctk.CTkFont(size=12, weight="bold"), justify="center",
        )
        self.import_file_label.grid(row=4, column=0, padx=40, pady=(0, 9))
        self.current_event_details = ctk.CTkLabel(
            card, text="", text_color=MUTED, font=ctk.CTkFont(size=12),
            justify="center", wraplength=650,
        )
        self.current_event_details.grid(row=5, column=0, padx=40, pady=(0, 10))
        buttons = ctk.CTkFrame(card, fg_color="transparent")
        buttons.grid(row=6, column=0, padx=90, pady=(0, 5))
        buttons.grid_columnconfigure((0, 1), weight=1, uniform="event_actions")
        self.current_event_buttons = buttons
        self.current_event_primary = ctk.CTkButton(
            buttons, text="Import Shopify CSV", command=self.choose_csv,
            width=290, height=50, corner_radius=9, fg_color=PURPLE, hover_color=PURPLE_DARK,
            font=ctk.CTkFont(size=15, weight="bold"),
        )
        self.current_event_primary.grid(row=0, column=0, columnspan=2, padx=6, pady=6)
        self.current_event_replace = ctk.CTkButton(
            buttons, text="Replace Imported CSV", command=self.replace_current_csv,
            width=245, height=50, corner_radius=9, fg_color=PURPLE_LIGHT, hover_color="#E6D9F5",
            border_width=1, border_color="#BFA9DD", text_color=PURPLE_DARK,
            font=ctk.CTkFont(size=14, weight="bold"),
        )
        self.current_event_replace.grid(row=0, column=1, padx=6, pady=6)
        self.current_event_regenerate = ctk.CTkButton(
            buttons, text="Regenerate Purchase Review", command=self.regenerate_current_review,
            width=250, height=44, corner_radius=9, fg_color="transparent", hover_color=PURPLE_LIGHT,
            border_width=1, border_color=BORDER, text_color=PURPLE_DARK,
        )
        self.current_event_regenerate.grid(row=1, column=0, columnspan=2, padx=6, pady=6)
        self.current_event_discard = ctk.CTkButton(
            buttons, text="Discard Current Event", command=self.discard_current_event,
            width=230, height=40, corner_radius=9, fg_color="#FFF1F1", hover_color="#FADDDD",
            border_width=1, border_color="#D8A5A5", text_color=DANGER,
            font=ctk.CTkFont(size=13, weight="bold"),
        )
        self.current_event_discard.grid(row=2, column=0, columnspan=2, padx=6, pady=(10, 0))
        self.current_event_archive = ctk.CTkButton(
            card, text="Open Event Archive", command=lambda: self.show_page("archive"),
            width=180, height=32, corner_radius=8, fg_color="transparent",
            hover_color=PURPLE_LIGHT, text_color=PURPLE,
            font=ctk.CTkFont(size=12, weight="bold"),
        )
        self.current_event_archive.grid(row=7, column=0, pady=(0, 16))

    def _set_current_event_stage(self, index: int, state: str, status_text: str) -> None:
        if state == "complete":
            circle_fill, circle_text, border, symbol = SUCCESS, WHITE, SUCCESS, "✓"
            title_color, status_color = "#166534", "#166534"
            enabled = "normal"
        elif state == "attention":
            circle_fill, circle_text, border, symbol = "#FFF1DC", "#9A4D00", WARNING, "!"
            title_color, status_color = "#9A4D00", "#9A4D00"
            enabled = "normal"
        elif state == "active":
            circle_fill, circle_text, border, symbol = PURPLE, WHITE, PURPLE, str(index + 1)
            title_color, status_color = PURPLE, PURPLE
            enabled = "normal"
        else:
            circle_fill, circle_text, border, symbol = WHITE, MUTED, "#CFC7DA", str(index + 1)
            title_color, status_color = TEXT, MUTED
            enabled = "disabled"
        for attr in (
            "mc_stage_labels", "current_event_stage_labels", "master_stage_labels",
            "review_stage_labels", "purchase_stage_labels",
        ):
            tracker = getattr(self, attr, None)
            if not tracker or index >= len(tracker):
                continue
            stage = tracker[index]
            stage["circle"].configure(
                text=symbol, fg_color=circle_fill, hover_color=PURPLE_LIGHT,
                text_color=circle_text, border_color=border, state=enabled,
            )
            stage["title"].configure(text_color=title_color)
            stage["status"].configure(text=status_text, text_color=status_color, fg_color="transparent")
            connector = stage.get("connector")
            if connector is not None:
                connector.configure(fg_color=PURPLE if state == "complete" else "#D7D0E0")

    def _refresh_current_event_progress(self, snapshot: dict | None = None) -> None:
        snapshot = snapshot or {}
        has_review = bool(self.last_review_workbook and self.last_review_workbook.exists())
        has_csv = bool(self.selected_csv) or has_review
        has_purchase_orders = bool(self.last_purchase_order_dir and self.last_purchase_order_dir.exists())
        product_issues = sum(
            1 for issue in snapshot.get("issues", [])
            if _clean(issue.get("fix_in", issue.get("source", ""))).casefold() == "product master"
        )
        review_issues = int(snapshot.get("review_count", 0) or 0)
        blocked_routes = int(snapshot.get("blocked_route_count", 0) or 0)

        if not has_csv:
            states = (
                ("active", "Start Here"), ("pending", "Waiting"),
                ("pending", "Waiting"), ("pending", "Waiting"),
            )
        elif not has_review:
            states = (
                ("active", "CSV Selected"), ("pending", "Waiting"),
                ("pending", "Waiting"), ("pending", "Waiting"),
            )
        elif has_purchase_orders:
            states = tuple(("complete", "Complete") for _ in range(4))
        elif product_issues:
            states = (
                ("complete", "Complete"), ("attention", "Needs Setup"),
                ("pending", "Waiting"), ("pending", "Waiting"),
            )
        elif review_issues or blocked_routes:
            states = (
                ("complete", "Complete"), ("complete", "Complete"),
                ("attention", "Needs Attention"), ("pending", "Waiting"),
            )
        else:
            states = (
                ("complete", "Complete"), ("complete", "Complete"),
                ("complete", "Complete"), ("active", "Ready"),
            )
        for index, (state, status) in enumerate(states):
            self._set_current_event_stage(index, state, status)

    def build_master_page(self):
        page = self.new_page("master")
        self._build_workflow_tracker(page, 0, "master_stage_labels", padx=52)
        card, badge, title, body = self.guided_stage_card(
            page, "", "Complete Product Setup",
            "Complete permanent product information once so it can be reused on every future purchase packet."
        )
        self.master_step_badge = badge
        self.master_state_title = title
        self.master_state_body = body
        self.master_health_large = ctk.CTkLabel(
            card, text="—", text_color=PURPLE,
            font=ctk.CTkFont(size=30, weight="bold"), justify="center"
        )
        self.master_health_large.grid(row=4, column=0, pady=(0, 4))
        self.master_details = ctk.CTkLabel(
            card, text="", text_color=MUTED, font=ctk.CTkFont(size=14),
            justify="center", wraplength=860
        )
        self.master_details.grid(row=5, column=0, padx=44, pady=(0, 10))
        self.master_issue_identity = ctk.CTkLabel(
            card, text="", text_color=TEXT, font=ctk.CTkFont(size=19, weight="bold"),
            justify="center", wraplength=900
        )
        self.master_issue_identity.grid(row=6, column=0, padx=44, pady=(0, 4))
        self.master_issue_reason = ctk.CTkLabel(
            card, text="", text_color=WARNING, font=ctk.CTkFont(size=14, weight="bold"),
            justify="center", wraplength=900
        )
        self.master_issue_reason.grid(row=7, column=0, padx=44, pady=(0, 14))
        buttons = ctk.CTkFrame(card, fg_color="transparent")
        buttons.grid(row=8, column=0, pady=(0, 34))
        self.master_page_open_button = ctk.CTkButton(
            buttons, text="Open Product Master", command=self.open_product_master,
            width=280, height=52, fg_color=PURPLE, hover_color=PURPLE_DARK,
            font=ctk.CTkFont(size=15, weight="bold")
        )
        self.master_page_open_button.pack(side="left", padx=7)
        self.master_page_secondary_button = ctk.CTkButton(
            buttons, text="+ Add New Product", command=self.add_new_product,
            width=200, height=52, fg_color=PURPLE_LIGHT, hover_color="#E6D9F5",
            border_width=1, border_color="#BFA9DD", text_color=PURPLE_DARK,
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.master_page_secondary_button.pack(side="left", padx=7)
        self._active_product_setup_issue = None

    def _product_master_issues(self) -> list[dict]:
        if not self.last_review_workbook or not self.last_review_workbook.exists():
            return []
        snapshot = self.last_snapshot or self._mission_snapshot()
        return [
            issue for issue in snapshot.get("issues", [])
            if _clean(issue.get("source", "")) == "Product Master"
        ]

    def open_next_product_setup(self):
        issues = self._product_master_issues()
        if not issues:
            self.open_product_master()
            return
        issue = issues[0]
        product = _clean(issue.get("product", "")) or _clean(issue.get("description", ""))
        self.open_product_master(setup_product=product)

    def refresh_master_page(self):
        if not hasattr(self, "master_state_title"):
            return
        health = _product_master_health()
        issues = self._product_master_issues()
        if issues:
            issue = issues[0]
            self._active_product_setup_issue = issue
            count = len(issues)
            self.master_step_badge.configure(text="")
            self.master_state_title.configure(text="Complete Product Setup")
            self.master_state_body.configure(
                text="New styles from this order have been added safely to Product Master. Complete each setup item, then Save & Next."
            )
            self.master_health_large.configure(
                text=f"{count} setup decision{'s' if count != 1 else ''} remaining",
                text_color=WARNING,
            )
            colors = _clean(issue.get("garment_colors", ""))
            identity_parts = [
                _clean(issue.get("product", "")),
                _clean(issue.get("description", "")),
            ]
            identity = " — ".join(dict.fromkeys(part for part in identity_parts if part))
            self.master_issue_identity.configure(text=identity or "Next product setup item")
            detail_parts = [_clean(issue.get("reason", ""))]
            if colors:
                detail_parts.append(f"Colors in this order: {colors}")
            self.master_issue_reason.configure(text="  •  ".join(part for part in detail_parts if part))
            self.master_page_open_button.configure(
                text="Set Up This Product", command=self.open_next_product_setup, state="normal"
            )
            self.master_page_secondary_button.configure(
                text="Open Full Product Master", command=self.open_product_master, state="normal"
            )
        elif self.last_review_workbook and self._review_is_stale():
            self._active_product_setup_issue = None
            self.master_state_title.configure(text="Product Setup Saved")
            self.master_state_body.configure(text="Apply the saved Product Master changes to the current purchase packet.")
            self.master_health_large.configure(text="Ready to refresh", text_color=SUCCESS)
            self.master_issue_identity.configure(text="")
            self.master_issue_reason.configure(text="")
            self.master_page_open_button.configure(
                text="Apply Changes to Purchase Review", command=self.regenerate_current_review, state="normal"
            )
            self.master_page_secondary_button.configure(
                text="Open Product Master", command=self.open_product_master, state="normal"
            )
        else:
            self._active_product_setup_issue = None
            self.master_state_title.configure(text="Complete Product Setup")
            self.master_state_body.configure(
                text="Maintain the permanent vendor, category, color, and decoration rules used for every future purchase packet."
            )
            self.master_health_large.configure(text=f"{health['percent']}% Complete", text_color=PURPLE)
            self.master_details.configure(
                text=f"{health['styles']} styles  •  {health['complete']} complete  •  {health['incomplete']} need attention"
            )
            self.master_issue_identity.configure(text="")
            self.master_issue_reason.configure(text="")
            self.master_page_open_button.configure(
                text="Open Product Master", command=self.open_product_master, state="normal"
            )
            self.master_page_secondary_button.configure(
                text="+ Add New Product", command=self.add_new_product, state="normal"
            )

    def build_review_page(self):
        page = self.new_page("review")
        # new_page reserves row 9 for some dashboard layouts. Purchase Review
        # uses row 1 as its expanding workspace, so remove that empty weighted
        # row and give every available vertical pixel to the editor.
        page.grid_rowconfigure(9, weight=0, minsize=0)
        page.grid_rowconfigure(1, weight=1, minsize=360)
        self._build_workflow_tracker(page, 0, "review_stage_labels", padx=34)
        review_shadow = ctk.CTkFrame(page, fg_color=SHADOW, corner_radius=24, width=1, height=1)
        review_shadow.grid(row=1, column=0, sticky="nsew", padx=20, pady=(1, 3))
        shell = ctk.CTkFrame(page, fg_color=WHITE, border_width=1, border_color=BORDER, corner_radius=24)
        shell.grid(row=1, column=0, sticky="nsew", padx=20, pady=(0, 4))
        shell.grid_columnconfigure(0, weight=1)
        shell.grid_rowconfigure(4, weight=1, minsize=320)
        self.review_shell = shell
        self._add_orchid_watermark(shell, size=(520, 520), relx=0.5, rely=0.58)

        self.review_step_label = ctk.CTkLabel(shell, text="", text_color=PURPLE,
                                              font=ctk.CTkFont(size=1))
        self.review_state_title = ctk.CTkLabel(shell, text="Complete Purchase Review", text_color=PURPLE_DARK,
                                               font=ctk.CTkFont(size=27, weight="bold"), justify="center")
        self.review_state_title.grid(row=0, column=0, padx=24, pady=(7, 1))
        self.review_state_body = ctk.CTkLabel(shell, text="Resolve only the order decisions that need attention.",
                                              text_color=MUTED, font=ctk.CTkFont(size=13), justify="center")
        self.review_state_body.grid(row=1, column=0, padx=24, pady=(0, 2))
        self.review_progress = ctk.CTkLabel(shell, text="", text_color=TEXT,
                                            font=ctk.CTkFont(size=13, weight="bold"))
        self.review_progress.grid(row=2, column=0, pady=(0, 1))

        self.review_decision_frame = ctk.CTkScrollableFrame(
            shell, fg_color="#FCFBFE", border_width=1, border_color=BORDER,
            corner_radius=16, scrollbar_button_color="#CDBBE5", scrollbar_button_hover_color=PURPLE,
        )
        self.review_decision_frame.grid(row=4, column=0, sticky="nsew", padx=8, pady=(1, 0))
        self.review_decision_frame.grid_columnconfigure(1, weight=1)

        self.review_identity = ctk.CTkLabel(self.review_decision_frame, text="", text_color=TEXT,
                                            font=ctk.CTkFont(size=18, weight="bold"), anchor="w", justify="left")
        self.review_identity.grid(row=0, column=0, columnspan=2, padx=14, pady=(8, 3), sticky="ew")
        self.review_note_callout = ctk.CTkFrame(
            self.review_decision_frame, fg_color="#FFF6D9", border_width=2,
            border_color="#E2A21A", corner_radius=14,
        )
        self.review_note_callout.grid(row=1, column=0, columnspan=2, padx=14, pady=(3, 6), sticky="ew")
        self.review_note_callout.grid_columnconfigure(0, weight=1)
        self.review_note_callout_title = ctk.CTkLabel(
            self.review_note_callout, text="⚠  CUSTOMER INSTRUCTION — REVIEW REQUIRED",
            text_color="#8A5200", font=ctk.CTkFont(size=15, weight="bold"), anchor="w", justify="left",
        )
        self.review_note_callout_title.grid(row=0, column=0, padx=14, pady=(9, 3), sticky="ew")
        self.review_note_callout_text = ctk.CTkLabel(
            self.review_note_callout, text="", text_color=PURPLE_DARK,
            font=ctk.CTkFont(size=18, weight="bold"), anchor="w", justify="left", wraplength=900,
        )
        self.review_note_callout_text.grid(row=1, column=0, padx=14, pady=(0, 10), sticky="ew")
        self.review_note_callout.grid_remove()

        self.review_reason = ctk.CTkLabel(self.review_decision_frame, text="", text_color=WARNING,
                                          font=ctk.CTkFont(size=14, weight="bold"), anchor="w", justify="left", wraplength=820)
        self.review_reason.grid(row=1, column=0, columnspan=2, padx=14, pady=(0, 3), sticky="ew")
        self.review_instructions = ctk.CTkLabel(self.review_decision_frame, text="", text_color=MUTED,
                                                font=ctk.CTkFont(size=13), anchor="w", justify="left", wraplength=900)
        self.review_instructions.grid(row=2, column=0, columnspan=2, padx=14, pady=(0, 5), sticky="ew")
        self.review_recommendation = ctk.CTkLabel(
            self.review_decision_frame, text="", fg_color="#EFE7FB", text_color=PURPLE_DARK,
            corner_radius=10, font=ctk.CTkFont(size=13, weight="bold"),
            anchor="w", justify="left", wraplength=900,
        )
        self.review_recommendation.grid(row=3, column=0, columnspan=2, padx=14, pady=(0, 6), sticky="ew")
        self.review_recommendation.grid_remove()

        self.review_vars = {
            "Product #": ctk.StringVar(), "Garment Color": ctk.StringVar(), "Size": ctk.StringVar(),
            "Quantity": ctk.StringVar(), "Purchase Vendor": ctk.StringVar(),
            "Decoration Type": ctk.StringVar(), "Decoration Location": ctk.StringVar(value=LEFT_CHEST),
            "Decoration Placement Instructions": ctk.StringVar(), "Decoration Color": ctk.StringVar(),
            "Include": ctk.StringVar(value="Yes"),
            "Do Not Outsource": ctk.StringVar(value="No"),
            "Decoration Decision": ctk.StringVar(value="Choose an instruction decision"),
        }
        self.review_custom_location_var = ctk.StringVar()
        self.review_vars["Decoration Type"].trace_add("write", self._on_review_decoration_type_changed)
        self.review_decoration_decision_label = ctk.CTkLabel(
            self.review_decision_frame, text="Instruction Decision", text_color=PURPLE_DARK,
            font=ctk.CTkFont(size=13, weight="bold"), anchor="w"
        )
        self.review_decoration_decision_label.grid(row=4, column=0, padx=(22, 12), pady=6, sticky="w")
        self.review_decoration_decision_menu = ctk.CTkOptionMenu(
            self.review_decision_frame, variable=self.review_vars["Decoration Decision"],
            values=INSTRUCTION_DECISION_OPTIONS,
            command=self._apply_review_decoration_decision, height=38,
            fg_color=PURPLE, button_color=PURPLE_DARK,
        )
        self.review_decoration_decision_menu.grid(row=4, column=1, padx=(0, 22), pady=6, sticky="ew")

        self.review_entries = {}
        labels = (
            ("Product #", 5), ("Garment Color", 6), ("Size", 7), ("Quantity", 8),
            ("Purchase Vendor", 9), ("Decoration Type", 10), ("Decoration Location", 11),
            ("Decoration Placement Instructions", 13), ("Decoration Color", 14),
        )
        for label, row in labels:
            display_label = "Placement Instructions" if label == "Decoration Placement Instructions" else label
            ctk.CTkLabel(self.review_decision_frame, text=display_label, text_color=TEXT,
                         font=ctk.CTkFont(size=13, weight="bold"), anchor="w").grid(
                row=row, column=0, padx=(14, 10), pady=4, sticky="w")
            if label == "Decoration Location":
                self.review_location_menu = ctk.CTkOptionMenu(
                    self.review_decision_frame, variable=self.review_vars[label],
                    values=DECORATION_LOCATIONS, command=self._apply_review_location_choice,
                    height=36, fg_color=PURPLE, button_color=PURPLE_DARK,
                )
                self.review_location_menu.grid(row=row, column=1, padx=(0, 14), pady=4, sticky="ew")
            elif label == "Decoration Type":
                self.review_decoration_type_menu = ctk.CTkOptionMenu(
                    self.review_decision_frame, variable=self.review_vars[label],
                    values=["Embroidery", "Screen Print", SEW_ON_PATCH_LABEL, HEMMING_ALTERATION_LABEL, BLANK_DECORATION_LABEL],
                    command=lambda _choice: self._on_review_decoration_type_changed(),
                    height=36, fg_color=PURPLE, button_color=PURPLE_DARK,
                )
                self.review_decoration_type_menu.grid(row=row, column=1, padx=(0, 14), pady=4, sticky="ew")
            else:
                entry = ctk.CTkEntry(self.review_decision_frame, textvariable=self.review_vars[label], height=36)
                entry.grid(row=row, column=1, padx=(0, 14), pady=4, sticky="ew")
                self.review_entries[label] = entry

        # The reserved panel gives the expanded location menu a clean area instead
        # of letting it visually cover Include Item. It also holds custom input.
        self.review_location_detail = ctk.CTkFrame(self.review_decision_frame, fg_color="#F7F3FC", corner_radius=10, height=56)
        self.review_location_detail.grid(row=12, column=0, columnspan=2, padx=14, pady=(1, 5), sticky="ew")
        self.review_location_detail.grid_columnconfigure(1, weight=1)
        self.review_location_detail.grid_propagate(False)
        self.review_location_helper = ctk.CTkLabel(
            self.review_location_detail,
            text="Location routes are separated on the vendor PO. Choose Other / Custom for a rare placement.",
            text_color=MUTED, font=ctk.CTkFont(size=12), anchor="w", justify="left", wraplength=850,
        )
        self.review_location_helper.grid(row=0, column=0, columnspan=2, padx=12, pady=(10, 5), sticky="ew")
        self.review_custom_location_label = ctk.CTkLabel(
            self.review_location_detail, text="Custom Location *", text_color=PURPLE_DARK,
            font=ctk.CTkFont(size=13, weight="bold"), anchor="w",
        )
        self.review_custom_location_entry = ctk.CTkEntry(
            self.review_location_detail, textvariable=self.review_custom_location_var, height=38,
            placeholder_text="Example: Right Chest or Center Back",
        )
        self.review_custom_location_label.grid(row=1, column=0, padx=(12, 10), pady=(3, 10), sticky="w")
        self.review_custom_location_entry.grid(row=1, column=1, padx=(0, 12), pady=(3, 10), sticky="ew")
        self.review_custom_location_label.grid_remove()
        self.review_custom_location_entry.grid_remove()

        ctk.CTkLabel(self.review_decision_frame, text="Include Item", text_color=TEXT,
                     font=ctk.CTkFont(size=13, weight="bold"), anchor="w").grid(
            row=15, column=0, padx=(14, 10), pady=4, sticky="w")
        self.review_include_menu = ctk.CTkOptionMenu(
            self.review_decision_frame, variable=self.review_vars["Include"], values=["Yes", "No"], height=38
        )
        self.review_include_menu.grid(row=15, column=1, padx=(0, 14), pady=4, sticky="w")

        # Keep review actions pinned at the bottom of the full-height page.
        # The form itself scrolls independently, so the buttons stay available
        # and no longer consume space inside the editable field list.
        action_bar = ctk.CTkFrame(shell, fg_color="transparent")
        action_bar.grid(row=5, column=0, padx=12, pady=(4, 10), sticky="e")
        self.review_product_master_button = ctk.CTkButton(
            action_bar, text="Open Product Master", command=self.open_product_master, width=180, height=44,
            fg_color=PURPLE_LIGHT, hover_color="#E6D9F5", border_width=1,
            border_color="#BFA9DD", text_color=PURPLE_DARK,
        )
        self.review_product_master_button.pack(side="left", padx=4)
        self.review_exit_button = ctk.CTkButton(
            action_bar, text="Save Progress & Exit", command=self.save_review_progress_and_exit,
            width=175, height=42, fg_color="transparent", hover_color=PURPLE_LIGHT,
            border_width=1, border_color="#BFA9DD", text_color=PURPLE_DARK,
        )
        self.review_exit_button.pack(side="left", padx=4)
        self.review_previous_button = ctk.CTkButton(
            action_bar, text="← Previous Decision", command=self.open_previous_review_decision,
            width=180, height=44, fg_color="transparent", hover_color=PURPLE_LIGHT,
            border_width=1, border_color="#BFA9DD", text_color=PURPLE_DARK, state="disabled",
        )
        self.review_previous_button.pack(side="left", padx=4)
        self.review_search_button = ctk.CTkButton(
            action_bar, text="Search Review", command=self.open_purchase_review_search,
            width=145, height=44, fg_color=PURPLE_LIGHT, hover_color="#E6D9F5",
            border_width=1, border_color="#BFA9DD", text_color=PURPLE_DARK,
        )
        self.review_search_button.pack(side="left", padx=4)
        self.review_export_button = ctk.CTkButton(
            action_bar, text="Export/Open Excel Copy", command=self.open_latest_review, width=190, height=44,
            fg_color="transparent", hover_color=PURPLE_LIGHT, border_width=1, border_color=BORDER, text_color=MUTED,
        )
        self.review_export_button.pack(side="left", padx=4)
        self.review_save_button = ctk.CTkButton(
            action_bar, text="Save & Next", command=self.save_current_review_decision,
            width=160, height=46, fg_color=PURPLE, hover_color=PURPLE_DARK,
            font=ctk.CTkFont(size=14, weight="bold"),
        )
        self.review_save_button.pack(side="left", padx=4)
        self._active_review_issue = None
        self._review_decision_history = []
        self._review_history_index = None
        self._review_search_issue = None
    def _set_review_location_value(self, location: str):
        decoration_type = self.review_vars["Decoration Type"].get()
        normalized = normalize_decoration_location(
            location, decoration_type,
            requires_decoration=not is_blank_decoration(decoration_type) and not is_in_house_decoration(decoration_type),
            product_name=(self._active_review_issue or {}).get("description", ""),
        )
        choice, custom = location_choice_and_custom(normalized, decoration_type)
        self.review_vars["Decoration Location"].set(choice)
        self.review_custom_location_var.set(custom)
        self._apply_review_location_choice(choice)

    def _apply_review_location_choice(self, choice: str):
        choice = _clean(choice)
        is_custom = choice == OTHER_CUSTOM
        if is_custom:
            self.review_location_detail.configure(height=112)
            self.review_location_helper.configure(
                text="Enter the exact custom route below. It will receive its own labeled PO section/page."
            )
            self.review_custom_location_label.grid()
            self.review_custom_location_entry.grid()
            self.review_custom_location_entry.configure(state="normal")
        else:
            self.review_location_detail.configure(height=56)
            self.review_custom_location_var.set("")
            self.review_custom_location_label.grid_remove()
            self.review_custom_location_entry.grid_remove()
            if choice in {NOT_APPLICABLE_NO_DECORATION, NOT_APPLICABLE_IN_HOUSE}:
                self.review_location_helper.configure(text=choice)
            else:
                self.review_location_helper.configure(
                    text="Location routes are separated on the vendor PO. Choose Other / Custom for a rare placement."
                )

    def _resolved_review_location(self, decoration_type: str) -> str:
        if is_blank_decoration(decoration_type) or is_in_house_decoration(decoration_type):
            return ""
        choice = self.review_vars["Decoration Location"].get()
        location = resolve_location_choice(choice, self.review_custom_location_var.get())
        return normalize_decoration_location(
            location, decoration_type, requires_decoration=True,
            product_name=(self._active_review_issue or {}).get("description", ""),
        )

    def _on_review_decoration_type_changed(self, *_):
        if not hasattr(self, "review_location_menu"):
            return
        decoration_type = self.review_vars["Decoration Type"].get()
        in_house = is_in_house_decoration(decoration_type)
        no_decoration = is_blank_decoration(decoration_type)
        disabled = in_house or no_decoration
        if disabled:
            display = NOT_APPLICABLE_IN_HOUSE if in_house else NOT_APPLICABLE_NO_DECORATION
            self.review_vars["Decoration Location"].set(display)
            self.review_custom_location_var.set("")
            self.review_vars["Decoration Placement Instructions"].set("")
            self.review_vars["Decoration Color"].set("")
            self.review_location_menu.configure(state="disabled")
            self.review_custom_location_entry.configure(state="disabled")
            placement = self.review_entries.get("Decoration Placement Instructions")
            color = self.review_entries.get("Decoration Color")
            if placement is not None:
                placement.configure(state="disabled")
            if color is not None:
                color.configure(state="disabled")
            self._apply_review_location_choice(display)
            self._update_instruction_decision_options()
        else:
            current = self.review_vars["Decoration Location"].get().strip()
            if not current or current in {NOT_APPLICABLE_NO_DECORATION, NOT_APPLICABLE_IN_HOUSE}:
                self._set_review_location_value(default_decoration_location(
                    (self._active_review_issue or {}).get("description", ""), "", decoration_type
                ))
            self.review_location_menu.configure(state="normal")
            placement = self.review_entries.get("Decoration Placement Instructions")
            color = self.review_entries.get("Decoration Color")
            if placement is not None:
                placement.configure(state="normal")
            if color is not None:
                color.configure(state="normal")
            self._update_instruction_decision_options()

    def _review_outsourcing_exception_applicable(self, decoration_type: str | None = None) -> bool:
        """Return whether this line can be redirected from an outside decorator.

        In an Entire Order Outsourced packet, every included physical garment is
        normally routed to the outside decorator. The exception must therefore
        remain available even when the user changes the decoration type to
        Blank Garment (No Decoration), because that garment still needs to be
        purchased separately and shipped to Orchid. Internal service-only lines
        do not use this routing exception.
        """
        decoration_value = _clean(
            decoration_type if decoration_type is not None else self.review_vars["Decoration Type"].get()
        )
        decoration_key = decoration_value.casefold()
        entire_order_outsourced = (
            normalize_decoration_fulfillment(self.current_decoration_fulfillment)
            == ENTIRE_ORDER_OUTSOURCED
        )
        if entire_order_outsourced:
            return not is_in_house_decoration(decoration_value)
        return "screen" in decoration_key

    def _review_instruction_decision_options(self, decoration_type: str | None = None) -> list[str]:
        instruction_kind = getattr(self, "_review_instruction_kind", "decoration")
        if instruction_kind == "general":
            return [
                "Choose an instruction decision",
                "Follow Note as Written",
                "Keep Product Master Default",
            ]
        options = list(INSTRUCTION_DECISION_OPTIONS)
        if self._review_outsourcing_exception_applicable(decoration_type):
            options.append(DO_NOT_OUTSOURCE_DECISION)
        return options

    def _update_instruction_decision_options(self):
        if not hasattr(self, "review_decoration_decision_menu"):
            return
        options = self._review_instruction_decision_options()
        self.review_decoration_decision_menu.configure(values=options)
        current = _clean(self.review_vars["Decoration Decision"].get())
        if current not in options:
            if current == DO_NOT_OUTSOURCE_DECISION:
                self.review_vars["Do Not Outsource"].set("No")
            self.review_vars["Decoration Decision"].set("Choose an instruction decision")

    def _save_current_review_draft(self) -> bool:
        issue = getattr(self, "_active_review_issue", None)
        if not issue or not self.last_review_workbook or not self.last_review_workbook.exists():
            return True
        if _clean(issue.get("fix_in", issue.get("source", ""))).casefold() == "product master":
            return True
        values = {field: _clean(variable.get()) for field, variable in self.review_vars.items()}
        if values.get("Decoration Location") == OTHER_CUSTOM and _clean(self.review_custom_location_var.get()):
            values["Decoration Location"] = _clean(self.review_custom_location_var.get())
        target_line = _clean(issue.get("line_id", ""))
        target_key = _clean(issue.get("decision_key", ""))
        try:
            workbook = load_workbook(self.last_review_workbook)
            updated = False
            for sheet_name in ("Review & Edit", "All PO Lines"):
                if sheet_name not in workbook.sheetnames:
                    continue
                sheet = workbook[sheet_name]
                header_row = None
                headers = {}
                for row_number in range(1, min(sheet.max_row, 20) + 1):
                    row_values = [str(sheet.cell(row_number, col).value or "").strip() for col in range(1, sheet.max_column + 1)]
                    if "Line ID" in row_values:
                        header_row = row_number
                        headers = {value: idx + 1 for idx, value in enumerate(row_values) if value}
                        break
                if not header_row:
                    continue
                for field, default_value in (("Do Not Outsource", "No"), ("Decoration Decision", "")):
                    if field not in headers:
                        column = sheet.max_column + 1
                        sheet.cell(header_row, column).value = field
                        for row_number in range(header_row + 1, sheet.max_row + 1):
                            sheet.cell(row_number, column).value = default_value
                        headers[field] = column
                for row_number in range(header_row + 1, sheet.max_row + 1):
                    line_id = _clean(sheet.cell(row_number, headers.get("Line ID", 0)).value) if headers.get("Line ID") else ""
                    decision_key = _clean(sheet.cell(row_number, headers.get("Decision Key", 0)).value) if headers.get("Decision Key") else ""
                    if not ((target_line and line_id == target_line) or (target_key and decision_key == target_key)):
                        continue
                    for field, value in values.items():
                        column = headers.get(field)
                        if column:
                            sheet.cell(row_number, column).value = value
                    updated = True
            if updated:
                workbook.save(self.last_review_workbook)
            workbook.close()
        except PermissionError:
            messagebox.showwarning(
                "Close Excel First",
                "Excel currently has the Purchase Review open. Close it, then click Save Progress & Exit again.",
            )
            return False
        except Exception as error:
            messagebox.showerror("Unable to Save Review Draft", str(error))
            return False
        self._update_review_state(
            draft_line_id=target_line, draft_decision_key=target_key, draft_saved_at=time.time(),
        )
        return True

    def save_review_progress_and_exit(self):
        if not self._save_current_review_draft():
            return
        self.dashboard_status.configure(
            text="Purchase Review draft saved. Reopen Purchase Review to continue from this line.",
            text_color=SUCCESS,
        )
        self.refresh_dashboard()
        self.show_page("dashboard")

    def _apply_review_decoration_decision(self, choice: str):
        """Apply the user's explicit instruction decision to the editable order fields."""
        choice = _clean(choice)
        if choice == DO_NOT_OUTSOURCE_DECISION:
            if not self._review_outsourcing_exception_applicable():
                messagebox.showwarning(
                    "Do Not Outsource Is Not Applicable",
                    "This exception is only available when the current decoration route would normally be outsourced.",
                )
                self.review_vars["Do Not Outsource"].set("No")
                self.review_vars["Decoration Decision"].set("Choose an instruction decision")
                return
            self.review_vars["Include"].set("Yes")
            self.review_vars["Do Not Outsource"].set("Yes")
            return

        # Every other instruction choice follows the event's normal fulfillment route.
        self.review_vars["Do Not Outsource"].set("No")
        if choice == "Keep Product Master Default":
            self.review_vars["Decoration Type"].set(getattr(self, "_review_default_decoration_type", ""))
            self._set_review_location_value(getattr(self, "_review_default_decoration_location", LEFT_CHEST))
            self.review_vars["Decoration Placement Instructions"].set(
                getattr(self, "_review_default_placement_instructions", "")
            )
            self.review_vars["Decoration Color"].set(getattr(self, "_review_default_decoration_color", ""))
        elif choice == "No Decoration":
            self.review_vars["Decoration Type"].set(BLANK_DECORATION_LABEL)
            self.review_vars["Decoration Location"].set("")
            self.review_custom_location_var.set("")
            self.review_vars["Decoration Placement Instructions"].set("")
            self.review_vars["Decoration Color"].set("")
            self._apply_review_location_choice("")
        elif choice in {"Embroidery", "Screen Print"}:
            self.review_vars["Decoration Type"].set(choice)
            if not self.review_vars["Decoration Location"].get().strip():
                self._set_review_location_value(default_decoration_location(
                    (self._active_review_issue or {}).get("description", ""), "", choice
                ))
        elif choice == "Sew On Patch":
            self.review_vars["Decoration Type"].set(SEW_ON_PATCH_LABEL)
        elif choice == "Hemming / Alteration":
            self.review_vars["Decoration Type"].set(HEMMING_ALTERATION_LABEL)
        elif choice == "Follow Note as Written":
            inferred = infer_decoration_location_from_note(getattr(self, "_review_instruction_text", ""))
            if inferred:
                self._set_review_location_value(inferred)

    def _update_previous_button(self):
        if not hasattr(self, "review_previous_button"):
            return
        if self._review_history_index is not None:
            enabled = self._review_history_index > 0
        else:
            enabled = bool(self._review_decision_history)
        self.review_previous_button.configure(state="normal" if enabled else "disabled")

    def _search_purchase_review_records(self, query: str) -> list[dict]:
        if not self.last_review_workbook or not self.last_review_workbook.exists():
            return []
        needle = _clean(query).casefold()
        if not needle:
            return []
        try:
            records = load_review_lines(self.last_review_workbook)
        except Exception:
            return []

        matches = []
        seen = set()
        for record in records:
            line_id = _clean(record.get("Line ID", ""))
            identity = line_id or "|".join(_clean(record.get(field, "")) for field in ("Order Number", "Product #", "Garment Color", "Size"))
            if identity in seen:
                continue
            searchable = " ".join(_clean(record.get(field, "")) for field in (
                "Product #", "Description", "Employee Name", "Company", "Order Number",
                "Garment Color", "Purchase Vendor",
            )).casefold()
            if needle not in searchable:
                continue
            seen.add(identity)
            style = _clean(record.get("Product #", ""))
            exact_style = style.casefold() == needle
            record_copy = dict(record)
            record_copy["_exact_style"] = exact_style
            matches.append(record_copy)
        matches.sort(key=lambda record: (
            0 if record.get("_exact_style") else 1,
            _clean(record.get("Product #", "")).casefold(),
            _clean(record.get("Employee Name", "")).casefold(),
            _clean(record.get("Order Number", "")).casefold(),
        ))
        return matches[:150]

    def open_purchase_review_search(self):
        if not self.last_review_workbook or not self.last_review_workbook.exists():
            messagebox.showinfo("No Purchase Review", "Process imported orders before searching Purchase Review.")
            return

        window = ctk.CTkToplevel(self)
        window.title("Search Purchase Review")
        window.geometry("980x720")
        window.minsize(760, 560)
        window.transient(self)
        window.grab_set()
        window.configure(fg_color=BG)
        window.grid_columnconfigure(0, weight=1)
        window.grid_rowconfigure(3, weight=1)

        ctk.CTkLabel(
            window, text="Search Purchase Review", text_color=PURPLE_DARK,
            font=ctk.CTkFont(size=25, weight="bold"), anchor="w",
        ).grid(row=0, column=0, padx=24, pady=(20, 4), sticky="ew")
        ctk.CTkLabel(
            window,
            text="Search by product/style number, employee, company/department, or order number. Select a result to edit that exact line.",
            text_color=MUTED, font=ctk.CTkFont(size=13), anchor="w", justify="left", wraplength=900,
        ).grid(row=1, column=0, padx=24, pady=(0, 10), sticky="ew")

        search_row = ctk.CTkFrame(window, fg_color="transparent")
        search_row.grid(row=2, column=0, padx=24, pady=(0, 12), sticky="ew")
        search_row.grid_columnconfigure(0, weight=1)
        query_var = ctk.StringVar()
        entry = ctk.CTkEntry(
            search_row, textvariable=query_var, height=42,
            placeholder_text="Example: ST350, J799S, Brock Bennett, or 1513",
        )
        entry.grid(row=0, column=0, padx=(0, 8), sticky="ew")
        results_frame = ctk.CTkScrollableFrame(
            window, fg_color=WHITE, border_width=1, border_color=BORDER, corner_radius=14,
            scrollbar_button_color="#CDBBE5", scrollbar_button_hover_color=PURPLE,
        )
        results_frame.grid(row=3, column=0, padx=24, pady=(0, 20), sticky="nsew")
        results_frame.grid_columnconfigure(0, weight=1)
        register_scrollable(self, results_frame)

        def open_result(record: dict):
            self._review_history_index = None
            self._review_search_issue = {
                "line_id": _clean(record.get("Line ID", "")),
                "decision_key": _clean(record.get("Decision Key", "")),
                "order": f"#{_clean(record.get('Order Number', ''))}" if _clean(record.get("Order Number", "")) else "",
                "employee": _clean(record.get("Employee Name", "")),
                "product": _clean(record.get("Product #", "")),
                "description": _clean(record.get("Description", "")),
                "reason": "Search result selected for correction",
                "instructions": "Update the incorrect order information, then save the line.",
                "fix_in": "Purchase Review",
                "source": "Purchase Review Search",
            }
            window.grab_release()
            window.destroy()
            self.show_page("review")
            self.refresh_review_page()

        def render_results(*_args):
            for child in results_frame.winfo_children():
                child.destroy()
            query = _clean(query_var.get())
            if not query:
                ctk.CTkLabel(
                    results_frame, text="Enter a style, employee, or order number to search.",
                    text_color=MUTED, font=ctk.CTkFont(size=14),
                ).grid(row=0, column=0, padx=20, pady=30)
                return
            matches = self._search_purchase_review_records(query)
            if not matches:
                ctk.CTkLabel(
                    results_frame, text=f"No Purchase Review lines matched ‘{query}’.",
                    text_color=WARNING, font=ctk.CTkFont(size=14, weight="bold"),
                ).grid(row=0, column=0, padx=20, pady=30)
                return
            ctk.CTkLabel(
                results_frame, text=f"{len(matches)} matching line{'s' if len(matches) != 1 else ''}",
                text_color=PURPLE_DARK, font=ctk.CTkFont(size=13, weight="bold"), anchor="w",
            ).grid(row=0, column=0, padx=12, pady=(10, 4), sticky="ew")
            for index, record in enumerate(matches, start=1):
                card = ctk.CTkFrame(
                    results_frame, fg_color="#FCFBFE", border_width=1, border_color=BORDER, corner_radius=12,
                )
                card.grid(row=index, column=0, padx=10, pady=5, sticky="ew")
                card.grid_columnconfigure(0, weight=1)
                style = _clean(record.get("Product #", "")) or "No style"
                description = _clean(record.get("Description", ""))
                employee = _clean(record.get("Employee Name", "")) or "Employee not listed"
                company = _clean(record.get("Company", ""))
                order = _clean(record.get("Order Number", ""))
                color = _clean(record.get("Garment Color", "")) or "No color"
                size = _clean(record.get("Size", "")) or "No size"
                qty = _clean(record.get("Quantity", "")) or "0"
                vendor = _clean(record.get("Purchase Vendor", "")) or "Vendor not assigned"
                decoration = " — ".join(filter(None, [
                    _clean(record.get("Decoration Type", "")),
                    _clean(record.get("Decoration Location", "")),
                ])) or "Decoration not assigned"
                status = _clean(record.get("Review Status", "")) or "Current line"
                ctk.CTkLabel(
                    card, text=f"{style}  •  {description}", text_color=TEXT,
                    font=ctk.CTkFont(size=15, weight="bold"), anchor="w", justify="left",
                ).grid(row=0, column=0, padx=14, pady=(10, 2), sticky="ew")
                ctk.CTkLabel(
                    card,
                    text=f"{employee}{' — ' + company if company else ''}  •  Order #{order or '—'}  •  {color}  •  {size}  •  Qty {qty}",
                    text_color=PURPLE_DARK, font=ctk.CTkFont(size=12), anchor="w", justify="left",
                ).grid(row=1, column=0, padx=14, pady=1, sticky="ew")
                ctk.CTkLabel(
                    card, text=f"{vendor}  •  {decoration}  •  Status: {status}",
                    text_color=MUTED, font=ctk.CTkFont(size=12), anchor="w", justify="left",
                ).grid(row=2, column=0, padx=14, pady=(1, 10), sticky="ew")
                ctk.CTkButton(
                    card, text="Open Line", command=lambda item=dict(record): open_result(item),
                    width=115, height=34, fg_color=PURPLE, hover_color=PURPLE_DARK,
                ).grid(row=0, column=1, rowspan=3, padx=12, pady=10)

        ctk.CTkButton(
            search_row, text="Search", command=render_results, width=120, height=42,
            fg_color=PURPLE, hover_color=PURPLE_DARK,
        ).grid(row=0, column=1)
        entry.bind("<Return>", render_results)
        render_results()
        entry.focus_set()

    def open_previous_review_decision(self):
        """Edit the immediately preceding saved decision without reopening the queue."""
        if not self._review_decision_history:
            return
        if self._review_history_index is None:
            self._review_history_index = len(self._review_decision_history) - 1
        elif self._review_history_index > 0:
            self._review_history_index -= 1
        self.refresh_review_page()

    def _reset_review_navigation(self):
        self._review_decision_history = []
        self._review_history_index = None
        self._review_search_issue = None
        self._update_previous_button()

    def _set_decoration_decision_visibility(self, visible: bool):
        if visible:
            self._update_instruction_decision_options()
            self.review_decoration_decision_label.grid()
            self.review_decoration_decision_menu.grid()
            self.review_decoration_decision_menu.configure(state="normal")
        else:
            self.review_decoration_decision_label.grid_remove()
            self.review_decoration_decision_menu.grid_remove()
            if hasattr(self, "review_note_callout"):
                self._set_review_instruction_callout("", False)

    def _set_review_instruction_callout(self, note_text: str = "", visible: bool = False):
        """Show the customer instruction as the most prominent part of the review."""
        if visible and _clean(note_text):
            self._review_instruction_text = _clean(note_text)
            self.review_reason.grid_remove()
            self.review_note_callout_text.configure(text=_clean(note_text))
            self.review_note_callout.grid()
            self.review_decoration_decision_label.configure(text="Instruction Decision")
        else:
            self._review_instruction_text = ""
            self.review_note_callout.grid_remove()
            self.review_note_callout_text.configure(text="")
            self.review_reason.grid()
            self.review_decoration_decision_label.configure(text="Instruction Decision")


    def _set_review_fields_state(self, state: str):
        for entry in self.review_entries.values():
            entry.configure(state=state)
        if hasattr(self, "review_decoration_type_menu"):
            self.review_decoration_type_menu.configure(state=state)
        if hasattr(self, "review_location_menu"):
            self.review_location_menu.configure(state=state)
        if hasattr(self, "review_custom_location_entry") and self.review_vars["Decoration Location"].get() == OTHER_CUSTOM:
            self.review_custom_location_entry.configure(state=state)
        if state == "normal":
            self._on_review_decoration_type_changed()

    def _review_sheet_context(self):
        if not self.last_review_workbook or not self.last_review_workbook.exists():
            return None, None, None
        workbook = load_workbook(self.last_review_workbook)
        if "Review & Edit" not in workbook.sheetnames:
            workbook.close()
            return None, None, None
        sheet = workbook["Review & Edit"]
        header_row = None
        headers = {}
        for row in range(1, min(sheet.max_row, 20) + 1):
            values = [str(sheet.cell(row, col).value or "").strip() for col in range(1, sheet.max_column + 1)]
            if "Line ID" in values and "Review Status" in values:
                header_row = row
                headers = {value: idx + 1 for idx, value in enumerate(values) if value}
                break
        if not header_row:
            workbook.close()
            return None, None, None
        return workbook, sheet, (header_row, headers)

    def _read_review_line_values(self, issue: dict) -> dict:
        if not self.last_review_workbook or not self.last_review_workbook.exists():
            return {}
        workbook = load_workbook(self.last_review_workbook)
        target_line = _clean(issue.get("line_id", ""))
        target_key = _clean(issue.get("decision_key", ""))
        visible_and_context_fields = list(self.review_vars) + [
            "Purchase Instructions", "Shopify Order Notes", "Shopify Line Notes", "Review Reason",
            "Action Required", "Resolution", "Review Status",
        ]

        def context_for(sheet):
            for row_number in range(1, min(sheet.max_row, 20) + 1):
                row_values = [
                    str(sheet.cell(row_number, col).value or "").strip()
                    for col in range(1, sheet.max_column + 1)
                ]
                if "Line ID" in row_values:
                    return row_number, {value: idx + 1 for idx, value in enumerate(row_values) if value}
            return None, {}

        values = {}
        for sheet_name in ("Review & Edit", "All PO Lines"):
            if sheet_name not in workbook.sheetnames:
                continue
            sheet = workbook[sheet_name]
            header_row, headers = context_for(sheet)
            if not header_row:
                continue
            for row in range(header_row + 1, sheet.max_row + 1):
                line_id = _clean(sheet.cell(row, headers.get("Line ID", 0)).value) if headers.get("Line ID") else ""
                decision_key = _clean(sheet.cell(row, headers.get("Decision Key", 0)).value) if headers.get("Decision Key") else ""
                if (target_key and decision_key == target_key) or (target_line and line_id == target_line):
                    for field in visible_and_context_fields:
                        col = headers.get(field)
                        values[field] = _clean(sheet.cell(row, col).value) if col else ""
                    workbook.close()
                    return values
        workbook.close()
        return values

    def _scroll_review_to_top(self) -> None:
        frame = getattr(self, "review_decision_frame", None)
        if frame is None:
            return
        canvas = getattr(frame, "_parent_canvas", None)
        if canvas is None:
            return
        try:
            canvas.yview_moveto(0)
        except Exception:
            pass

    def _fresh_snapshot(self) -> dict:
        """Reload the workbook so every workflow screen uses the same live state."""
        snapshot = self._mission_snapshot()
        self.last_snapshot = snapshot
        return snapshot

    def _current_purchase_review_issues(self, snapshot: dict | None = None) -> list[dict]:
        snapshot = snapshot or self._fresh_snapshot()
        return [
            issue for issue in snapshot.get("issues", [])
            if _clean(issue.get("fix_in", issue.get("source", ""))).casefold() != "product master"
        ]

    def _blocker_summary(self, snapshot: dict, limit: int = 5) -> str:
        lines: list[str] = []
        for issue in snapshot.get("issues", [])[:limit]:
            identity = " • ".join(filter(None, [
                _clean(issue.get("order", "")),
                _clean(issue.get("employee", "")),
                _clean(issue.get("product", "")) or _clean(issue.get("description", "")),
            ]))
            reason = _clean(issue.get("reason", "")) or "Needs review"
            lines.append(f"• {identity or 'Purchase line'} — {reason}")
        remaining = max(int(snapshot.get("review_count", 0)) - len(lines), 0)
        if remaining:
            lines.append(f"• Plus {remaining} additional decision(s)")
        if not lines and int(snapshot.get("blocked_route_count", 0)):
            for route in snapshot.get("routes", [])[:limit]:
                if _clean(route.get("status", "")).casefold() == "ready":
                    continue
                lines.append(f"• {_clean(route.get('vendor', 'Vendor route'))} — route still blocked")
        return "\n".join(lines)

    def _review_status_counts(self) -> tuple[int, int]:
        if not self.last_review_workbook or not self.last_review_workbook.exists():
            return 0, 0
        workbook, sheet, context = self._review_sheet_context()
        if not workbook:
            return 0, 0
        header_row, headers = context
        decisions: dict[str, list[str]] = {}
        for row in range(header_row + 1, sheet.max_row + 1):
            include = _clean(sheet.cell(row, headers.get("Include", 0)).value) if headers.get("Include") else "Yes"
            if include.casefold() not in {"yes", "y", "true", "1", "include"}:
                continue
            line_id = _clean(sheet.cell(row, headers.get("Line ID", 0)).value) if headers.get("Line ID") else str(row)
            key = _clean(sheet.cell(row, headers.get("Decision Key", 0)).value) if headers.get("Decision Key") else ""
            key = key or f"line:{line_id}"
            status = _clean(sheet.cell(row, headers.get("Review Status", 0)).value) if headers.get("Review Status") else ""
            decisions.setdefault(key, []).append(status.casefold())
        workbook.close()
        total = len(decisions)
        completed = sum(1 for statuses in decisions.values() if statuses and all(status == "ready" for status in statuses))
        return completed, max(total - completed, 0)

    def refresh_review_page(self):
        if not hasattr(self, "review_state_title"):
            return
        self.after_idle(self._scroll_review_to_top)
        if not self.last_review_workbook or not self.last_review_workbook.exists():
            self.review_state_title.configure(text="Process Imported Orders")
            self.review_state_body.configure(text="Import orders first, then process the imported CSV. Orchid will check Product Master before Purchase Review.")
            self.review_progress.configure(text="No active review")
            self.review_identity.configure(text="The imported orders have not been processed yet.")
            self.review_reason.configure(text="")
            self.review_instructions.configure(text="")
            self.review_save_button.configure(text="Process Imported Orders", command=self.create_review_workbook, state="normal")
            self.review_product_master_button.pack_forget()
            self.review_recommendation.grid_remove()
            self._set_decoration_decision_visibility(False)
            self._update_previous_button()
            return
        snapshot = self._fresh_snapshot()
        issues = self._current_purchase_review_issues(snapshot)
        product_issues = [
            issue for issue in snapshot.get("issues", [])
            if _clean(issue.get("fix_in", issue.get("source", ""))).casefold() == "product master"
        ]
        search_issue = getattr(self, "_review_search_issue", None)
        if product_issues and not search_issue:
            issue = product_issues[0]
            self._active_review_issue = issue
            self.review_state_title.configure(text="Product Setup Required")
            self.review_state_body.configure(text="Complete the permanent product setup before order-specific review can continue.")
            self.review_progress.configure(text=f"{len(product_issues)} product setup decision(s)")
            self.review_identity.configure(text=f"{issue.get('product','')}  {issue.get('description','')}")
            self.review_reason.configure(text=issue.get("reason", "Product Master detail needed"))
            self.review_instructions.configure(text=issue.get("instructions", "Open Product Master to complete this item."))
            self._set_review_fields_state("disabled")
            self.review_include_menu.configure(state="disabled")
            self.review_save_button.configure(state="disabled", text="Save & Next")
            self.review_product_master_button.pack(side="left", padx=4)
            self.review_recommendation.grid_remove()
            self._set_decoration_decision_visibility(False)
            self.review_previous_button.configure(state="disabled")
            return
        self.review_product_master_button.pack_forget()
        if not issues and self._review_history_index is None and not search_issue:
            self._active_review_issue = None
            if int(snapshot.get("blocked_route_count", 0)):
                self.review_state_title.configure(text="Purchase Review Still Required")
                self.review_state_body.configure(text="A purchase-order route is still blocked. The same final readiness check used by Step 4 is shown below.")
                self.review_progress.configure(text=f"{int(snapshot.get('blocked_route_count', 0))} blocked route(s)")
                self.review_identity.configure(text="Purchase-order preflight did not pass")
                self.review_reason.configure(text=self._blocker_summary(snapshot) or "A vendor route still needs correction.")
                self.review_instructions.configure(text="Return to the dashboard Work Queue to open the exact blocked route.")
                self._set_review_fields_state("disabled")
                self.review_include_menu.configure(state="disabled")
                self.review_save_button.configure(text="Return to Dashboard", command=lambda: self.show_page("dashboard"), state="normal")
                self.review_recommendation.grid_remove()
                self._set_decoration_decision_visibility(False)
                self._update_previous_button()
                return
            self.review_state_title.configure(text="Purchase Review Complete")
            self.review_state_body.configure(text="No order decisions or blocked purchase-order rows remain. You can continue to Purchase Orders.")
            self.review_progress.configure(text="0 decisions remaining")
            self.review_identity.configure(text="✓ Final purchase-order preflight passed")
            self.review_reason.configure(text="")
            self.review_instructions.configure(text="The dashboard will advance automatically to Step 4.")
            self._set_review_fields_state("disabled")
            self.review_include_menu.configure(state="disabled")
            self.review_save_button.configure(text="Go to Purchase Orders", command=lambda: self.show_page("purchase"), state="normal")
            self.review_recommendation.grid_remove()
            self._set_decoration_decision_visibility(False)
            self._update_previous_button()
            return
        editing_previous = self._review_history_index is not None
        editing_search = bool(search_issue)
        draft_issue = self._saved_review_draft_issue(issues) if not editing_previous and not editing_search else None
        issue = search_issue if editing_search else (
            self._review_decision_history[self._review_history_index] if editing_previous else (draft_issue or issues[0])
        )
        self._active_review_issue = issue
        if editing_search:
            title_text = "Edit Search Result"
            body_text = "Correct this exact order line, then save to return to the normal review queue."
        elif editing_previous:
            title_text = "Edit Previous Decision"
            body_text = "Correct the saved choice, then return to your current place in the review queue."
        else:
            title_text = "Complete Purchase Review"
            body_text = "Correct the highlighted information, then save to continue to the next decision."
        self.review_state_title.configure(text=title_text)
        self.review_state_body.configure(text=body_text)
        completed_count, remaining_count = self._review_status_counts()
        self.review_progress.configure(
            text="Search result selected" if editing_search else (
                f"Previous decision {self._review_history_index + 1} of {len(self._review_decision_history)}"
                if editing_previous else f"{completed_count} completed  •  {remaining_count} remaining"
            )
        )
        identity = "  •  ".join(filter(None, [issue.get("order", ""), issue.get("employee", ""), issue.get("product", ""), issue.get("description", "")]))
        self.review_identity.configure(text=identity or "Current order decision")
        reason_text = issue.get("reason", "Needs review")
        self.review_reason.configure(text=reason_text)
        values = self._read_review_line_values(issue)
        instruction_text = issue.get("instructions", "Complete the missing order information.")
        note_text = values.get("Purchase Instructions") or values.get("Shopify Order Notes")
        is_customer_decision = "customer decision required" in _clean(reason_text).casefold()
        note_key = _clean(note_text).casefold()
        is_outsource_instruction = any(term in note_key for term in OUTSOURCE_INSTRUCTION_TERMS)
        is_decoration_instruction = bool(is_customer_decision and decoration_note_requires_review(note_text))
        is_instruction_decision = bool(is_customer_decision)
        self._review_instruction_kind = (
            "decoration" if is_decoration_instruction else
            "outsource" if is_outsource_instruction else
            "general"
        )
        if is_instruction_decision:
            self._set_review_instruction_callout(note_text, True)
            instruction_text = (
                "Choose an explicit instruction decision below, update any affected order fields, then save."
            )
            if is_decoration_instruction:
                recommendation, recommendation_reason = decoration_instruction_recommendation(
                    note_text, values.get("Decoration Type", "")
                )
            elif is_outsource_instruction and self._review_outsourcing_exception_applicable(values.get("Decoration Type", "")):
                recommendation = DO_NOT_OUTSOURCE_DECISION
                recommendation_reason = "The note directs this one item to Orchid instead of the outside decorator."
            else:
                recommendation = "Follow Note as Written"
                recommendation_reason = "Update the product, size, color, vendor, or other order-specific fields to match the instruction."
            if recommendation:
                self.review_recommendation.configure(
                    text=f"Recommended: {recommendation} — {recommendation_reason}", padx=14, pady=10
                )
                self.review_recommendation.grid()
            else:
                self.review_recommendation.grid_remove()
        else:
            self._review_instruction_kind = ""
            self.review_recommendation.grid_remove()
        self.review_instructions.configure(text=instruction_text)
        self._review_default_decoration_type = values.get("Decoration Type", "")
        self._review_default_decoration_location = normalize_decoration_location(
            values.get("Decoration Location", ""), values.get("Decoration Type", ""),
            requires_decoration=not is_blank_decoration(values.get("Decoration Type", "")) and not is_in_house_decoration(values.get("Decoration Type", "")),
            product_name=issue.get("description", ""),
        )
        self._review_default_placement_instructions = values.get("Decoration Placement Instructions", "")
        self._review_default_decoration_color = values.get("Decoration Color", "")
        actual_location = normalize_decoration_location(
            values.get("Decoration Location", ""), values.get("Decoration Type", ""),
            requires_decoration=not is_blank_decoration(values.get("Decoration Type", "")) and not is_in_house_decoration(values.get("Decoration Type", "")),
            product_name=issue.get("description", ""),
        )
        location_choice, custom_location = location_choice_and_custom(actual_location, values.get("Decoration Type", ""))
        values["Decoration Location"] = location_choice
        for field, variable in self.review_vars.items():
            default = "Yes" if field == "Include" else ("No" if field == "Do Not Outsource" else "")
            if field == "Decoration Decision" and is_instruction_decision:
                default = "Choose an instruction decision"
            variable.set(values.get(field, "") or default)
        if (
            is_instruction_decision
            and _clean(values.get("Do Not Outsource", "No")).casefold() in {"yes", "y", "true", "1"}
        ):
            self.review_vars["Decoration Decision"].set(DO_NOT_OUTSOURCE_DECISION)
        self.review_custom_location_var.set(custom_location)
        self._apply_review_location_choice(location_choice)
        self._set_decoration_decision_visibility(is_instruction_decision)
        self._set_review_fields_state("normal")
        self.review_include_menu.configure(state="normal")
        self.review_save_button.configure(
            text="Save Changes & Return" if (editing_previous or editing_search) else "Save & Next",
            command=self.save_current_review_decision, state="normal"
        )
        self._update_previous_button()

    def save_current_review_decision(self):
        issue = self._active_review_issue
        if not issue or _clean(issue.get("fix_in", issue.get("source", ""))).casefold() == "product master":
            return

        editing_previous = self._review_history_index is not None
        editing_search = getattr(self, "_review_search_issue", None) is not None
        values = {field: _clean(variable.get()) for field, variable in self.review_vars.items()}
        current_values = self._read_review_line_values(issue)
        note_text = current_values.get("Purchase Instructions") or current_values.get("Shopify Order Notes")
        note_key = _clean(note_text).casefold()
        is_outsource_instruction = any(term in note_key for term in OUTSOURCE_INSTRUCTION_TERMS)
        is_decoration_instruction = decoration_note_requires_review(note_text)
        is_instruction_decision = bool(
            "customer decision required" in _clean(issue.get("reason", "")).casefold()
        )
        self._review_instruction_kind = (
            "decoration" if is_decoration_instruction else
            "outsource" if is_outsource_instruction else
            "general"
        )
        choice = values.get("Decoration Decision", "")
        if choice == DO_NOT_OUTSOURCE_DECISION:
            values["Include"] = "Yes"
            values["Do Not Outsource"] = "Yes"
            self.review_vars["Include"].set("Yes")
        elif is_instruction_decision:
            values["Do Not Outsource"] = "No"
        else:
            # Preserve Product Master's automatic Never Outsource route when this
            # decision is about another field such as size, color, or vendor.
            values["Do Not Outsource"] = current_values.get("Do Not Outsource", "No") or "No"
        include_item = values.get("Include", "Yes").casefold() in {"yes", "y", "true", "1", "include"}
        do_not_outsource = values.get("Do Not Outsource", "No").casefold() in {"yes", "y", "true", "1"}
        if not include_item:
            values["Do Not Outsource"] = "No"
            do_not_outsource = False
        elif do_not_outsource and not self._review_outsourcing_exception_applicable(values.get("Decoration Type", "")):
            messagebox.showwarning(
                "Do Not Outsource Is Not Applicable",
                "This exception is only available when the current decoration route would normally be outsourced.",
            )
            self.review_vars["Do Not Outsource"].set("No")
            self.review_vars["Decoration Decision"].set("Choose an instruction decision")
            return
        if is_instruction_decision:
            choice = values.get("Decoration Decision", "")
            valid_choices = set(
                self._review_instruction_decision_options(values.get("Decoration Type", ""))[1:]
            )
            if choice not in valid_choices:
                choices_text = ", ".join(sorted(valid_choices))
                messagebox.showwarning(
                    "Choose an Instruction Decision",
                    f"This customer instruction must be acknowledged. Choose one of these options before saving: {choices_text}.",
                )
                self.review_decoration_decision_menu.focus_set()
                return
            if choice == "No Decoration":
                values["Decoration Type"] = BLANK_DECORATION_LABEL
                values["Decoration Color"] = ""
            elif choice in {"Embroidery", "Screen Print"}:
                values["Decoration Type"] = choice
                if include_item and not values.get("Decoration Color", ""):
                    messagebox.showwarning(
                        "Decoration Color Required",
                        f"Enter the thread/ink color for {choice}, then click Save & Next.",
                    )
                    self.review_entries["Decoration Color"].focus_set()
                    return
            elif choice == "Sew On Patch":
                values["Decoration Type"] = SEW_ON_PATCH_LABEL
                values["Decoration Color"] = ""
            elif choice == "Hemming / Alteration":
                values["Decoration Type"] = HEMMING_ALTERATION_LABEL
                values["Decoration Color"] = ""

        if include_item and not is_blank_decoration(values.get("Decoration Type", "")) and not is_in_house_decoration(values.get("Decoration Type", "")):
            if values.get("Decoration Location") == OTHER_CUSTOM and not _clean(self.review_custom_location_var.get()):
                messagebox.showwarning(
                    "Custom Location Required",
                    "Enter the exact custom decoration location before saving this decision.",
                )
                self.review_custom_location_entry.focus_set()
                return
            values["Decoration Location"] = self._resolved_review_location(values.get("Decoration Type", ""))
        else:
            values["Decoration Location"] = ""
            values["Decoration Placement Instructions"] = ""

        try:
            raw_quantity = values.get("Quantity", "")
            numeric_value = float(raw_quantity or 0)
            if include_item and (numeric_value <= 0 or not numeric_value.is_integer()):
                raise ValueError
            values["Quantity"] = int(numeric_value) if numeric_value.is_integer() else raw_quantity
        except (TypeError, ValueError):
            if include_item:
                messagebox.showwarning("Check Quantity", "Quantity must be a whole number greater than zero.")
                quantity_entry = self.review_entries.get("Quantity")
                if quantity_entry is not None:
                    quantity_entry.focus_set()
                return
            values["Quantity"] = 0

        workbook, sheet, context = self._review_sheet_context()
        if not workbook:
            messagebox.showerror("Unable to Save", "The Purchase Review workbook could not be opened.")
            return
        header_row, headers = context
        target_line = _clean(issue.get("line_id", ""))
        target_key = _clean(issue.get("decision_key", ""))

        def ensure_editable_header(target_sheet, target_header_row, target_headers, field, default_value):
            if target_sheet is None or not target_header_row or field in target_headers:
                return target_headers
            column = target_sheet.max_column + 1
            target_sheet.cell(target_header_row, column).value = field
            for row_number in range(target_header_row + 1, target_sheet.max_row + 1):
                target_sheet.cell(row_number, column).value = default_value
            target_headers[field] = column
            return target_headers

        headers = ensure_editable_header(sheet, header_row, headers, "Do Not Outsource", "No")

        def header_context(target_sheet):
            for row_number in range(1, min(target_sheet.max_row, 20) + 1):
                row_values = [
                    str(target_sheet.cell(row_number, col).value or "").strip()
                    for col in range(1, target_sheet.max_column + 1)
                ]
                if "Line ID" in row_values:
                    return row_number, {value: idx + 1 for idx, value in enumerate(row_values) if value}
            return None, {}

        def matching_rows(target_sheet, target_header_row, target_headers):
            rows = []
            for row_number in range(target_header_row + 1, target_sheet.max_row + 1):
                line_id = _clean(target_sheet.cell(row_number, target_headers.get("Line ID", 0)).value) if target_headers.get("Line ID") else ""
                decision_key = _clean(target_sheet.cell(row_number, target_headers.get("Decision Key", 0)).value) if target_headers.get("Decision Key") else ""
                if (target_line and line_id == target_line) or (target_key and decision_key == target_key):
                    rows.append(row_number)
            return rows

        visible_rows = matching_rows(sheet, header_row, headers)
        source_sheet = workbook["All PO Lines"] if "All PO Lines" in workbook.sheetnames else None
        source_header_row, source_headers = header_context(source_sheet) if source_sheet is not None else (None, {})
        if source_sheet is not None and source_header_row:
            source_headers = ensure_editable_header(source_sheet, source_header_row, source_headers, "Do Not Outsource", "No")
        source_rows = matching_rows(source_sheet, source_header_row, source_headers) if source_sheet is not None and source_header_row else []

        if not visible_rows and not editing_search:
            workbook.close()
            messagebox.showwarning("Decision Not Found", "The matching decision could not be located in Review & Edit.")
            return
        if source_sheet is not None and not source_rows:
            workbook.close()
            messagebox.showerror(
                "Source Line Not Found",
                "Orchid found the visible decision but could not find its matching source line. The decision was not advanced.",
            )
            return

        # Validate the proposed source line with the exact same purchasing rules
        # used by Mission Control and final purchase-order generation. This catches
        # stale labels such as a row saying Missing size when color is actually blank.
        proposed = {}
        if source_rows:
            source_row = source_rows[0]
            proposed = {
                field: source_sheet.cell(source_row, col).value
                for field, col in source_headers.items()
            }
        else:
            visible_row = visible_rows[0]
            proposed = {
                field: sheet.cell(visible_row, col).value
                for field, col in headers.items()
            }
        proposed.update(values)
        proposed["Review Status"] = "Ready"
        blockers = line_block_reasons(proposed)
        if include_item and blockers:
            workbook.close()
            field_by_reason = {
                "missing product number and description": "Product #",
                "missing purchase vendor": "Purchase Vendor",
                "missing decoration type": "Decoration Type",
                "missing decoration color": "Decoration Color",
                "missing garment color": "Garment Color",
                "missing size": "Size",
                "quantity must be greater than zero": "Quantity",
                "custom decoration location is incomplete": "Decoration Location",
            }
            first_field = next((field_by_reason.get(reason.casefold()) for reason in blockers if field_by_reason.get(reason.casefold())), None)
            messagebox.showwarning(
                "Complete Required Information",
                "This line cannot be marked Ready yet:\n\n" + "\n".join(f"• {reason}" for reason in blockers),
            )
            if first_field and self.review_entries.get(first_field) is not None:
                self.review_entries[first_field].focus_set()
            return

        def write_rows(target_sheet, row_numbers, target_headers, *, source=False):
            for row_number in row_numbers:
                for field, value in values.items():
                    col = target_headers.get(field)
                    if col:
                        target_sheet.cell(row_number, col).value = value
                if target_headers.get("Review Status"):
                    target_sheet.cell(row_number, target_headers["Review Status"]).value = "Ready"
                if target_headers.get("Resolution"):
                    resolution = "Completed in App"
                    if is_instruction_decision:
                        resolution = f"Instruction: {values.get('Decoration Decision', '')}"
                    target_sheet.cell(row_number, target_headers["Resolution"]).value = resolution
                if source:
                    permanent_reason = _clean(target_sheet.cell(row_number, target_headers.get("Permanent Review Reason", 0)).value) if target_headers.get("Permanent Review Reason") else ""
                    if target_headers.get("Event Review Reason"):
                        target_sheet.cell(row_number, target_headers["Event Review Reason"]).value = ""
                    if target_headers.get("Review Reason"):
                        target_sheet.cell(row_number, target_headers["Review Reason"]).value = permanent_reason
                    if target_headers.get("Fix In"):
                        target_sheet.cell(row_number, target_headers["Fix In"]).value = "Product Master" if permanent_reason else "Purchase Review"
                    if target_headers.get("Action Required") and not permanent_reason:
                        target_sheet.cell(row_number, target_headers["Action Required"]).value = ""

        write_rows(sheet, visible_rows, headers)
        if source_sheet is not None:
            write_rows(source_sheet, source_rows, source_headers, source=True)

        try:
            workbook.save(self.last_review_workbook)
        except PermissionError:
            workbook.close()
            messagebox.showwarning("Close Excel First", "Excel currently has the Purchase Review open. Close it, then click Save & Next again.")
            return
        workbook.close()
        self._clear_saved_review_draft(issue)

        snapshot = self._fresh_snapshot()
        same_line_issues = [
            item for item in snapshot.get("issues", [])
            if target_line and _clean(item.get("line_id", "")) == target_line
        ]
        if same_line_issues:
            self.refresh_dashboard()
            self.refresh_review_page()
            first = same_line_issues[0]
            if _clean(first.get("fix_in", first.get("source", ""))).casefold() == "product master":
                messagebox.showinfo(
                    "Product Master Setup Required",
                    f"{values.get('Product #') or first.get('product') or 'This product'} is not fully configured in Product Master. Orchid moved it to Step 2 before continuing Purchase Review.",
                )
                self.show_page("master")
            else:
                messagebox.showwarning(
                    "Decision Still Requires Attention",
                    self._blocker_summary(snapshot) or _clean(first.get("reason", "Needs review")),
                )
            return

        if editing_search:
            self._review_search_issue = None
        elif editing_previous:
            self._review_history_index = None
        else:
            history_key = _clean(issue.get("decision_key", "")) or _clean(issue.get("line_id", ""))
            last_key = ""
            if self._review_decision_history:
                last = self._review_decision_history[-1]
                last_key = _clean(last.get("decision_key", "")) or _clean(last.get("line_id", ""))
            if history_key and history_key != last_key:
                self._review_decision_history.append(dict(issue))

        self.refresh_dashboard()
        self.refresh_review_page()
        remaining = self._current_purchase_review_issues(snapshot)
        if not remaining and not int(snapshot.get("blocked_route_count", 0)):
            self.show_page("dashboard")

    def build_purchase_page(self):
        page = self.new_page("purchase")
        self._build_workflow_tracker(page, 0, "purchase_stage_labels", padx=52)
        self.purchase_card, _, self.purchase_state_title, self.purchase_state_body = self.guided_stage_card(
            page, "", "Generate Purchase Orders & Reports",
            "Assign one PO number per vendor, create concise vendor POs, and generate the in-house and outsourced decoration reports."
        )
        self.purchase_folder_label=ctk.CTkLabel(self.purchase_card,text="",text_color=TEXT,font=ctk.CTkFont(size=13,weight="bold"),justify="center",wraplength=820)
        self.purchase_folder_label.grid(row=4,column=0,padx=40,pady=(0,18))
        buttons=ctk.CTkFrame(self.purchase_card,fg_color="transparent")
        buttons.grid(row=6,column=0,pady=(0,34))
        self.purchase_primary_button=ctk.CTkButton(buttons,text="Generate Purchase Orders & Reports",command=self.create_final_purchase_orders,width=275,height=52,fg_color=PURPLE,hover_color=PURPLE_DARK,font=ctk.CTkFont(size=15,weight="bold"))
        self.purchase_primary_button.grid(row=0,column=0,padx=7,pady=5)
        self.purchase_secondary_button=ctk.CTkButton(buttons,text="",width=1,height=1)
        self.purchase_secondary_button.grid_remove()

    def refresh_purchase_page(self):
        if not hasattr(self, "purchase_state_title"):
            return
        if self.last_purchase_order_dir and self.last_purchase_order_dir.exists():
            self.purchase_state_title.configure(text="Purchase Documents Generated")
            self.purchase_state_body.configure(text="The vendor purchase orders and decoration reports are ready, and the event has been archived.")
            self.purchase_folder_label.configure(text=str(self.last_purchase_order_dir))
            self.purchase_primary_button.configure(text="Open Latest Purchase Documents", command=self.open_latest_purchase_orders, state="normal")
            self.purchase_secondary_button.grid_remove()
        elif self.last_review_workbook and self.last_review_workbook.exists():
            snapshot = self._fresh_snapshot()
            if int(snapshot.get("review_count", 0)) or int(snapshot.get("blocked_route_count", 0)):
                self.purchase_state_title.configure(text="Purchase Review Required")
                self.purchase_state_body.configure(
                    text="Step 4 uses the same final preflight as Purchase Review. Resolve the exact items below before generating vendor PDFs."
                )
                self.purchase_folder_label.configure(text=self._blocker_summary(snapshot) or self.last_review_workbook.name)
                self.purchase_primary_button.configure(text="Return to Purchase Review", command=lambda: self.show_page("review"), state="normal")
            else:
                self.purchase_state_title.configure(text="Generate Purchase Orders & Reports")
                self.purchase_state_body.configure(text="Final preflight passed. Click Generate to enter one PO number per vendor and create concise purchase orders plus both decoration reports.")
                self.purchase_folder_label.configure(text=self.last_review_workbook.name)
                self.purchase_primary_button.configure(text="Generate Purchase Orders & Reports", command=self.generate_latest_purchase_orders, state="normal")
            self.purchase_secondary_button.grid_remove()
        else:
            self.purchase_state_title.configure(text="Purchase Review Required")
            self.purchase_state_body.configure(text="Process the imported orders, complete Product Master if needed, and then finish Purchase Review before generating vendor purchase orders.")
            self.purchase_folder_label.configure(text="")
            self.purchase_primary_button.configure(text="Return to Import Orders", command=lambda: self.show_page("import"), state="normal")
            self.purchase_secondary_button.grid_remove()

    def build_employee_totals_page(self):
        page = self.new_page("employees")
        page.grid_rowconfigure(0, weight=1)

        employee_shadow = ctk.CTkFrame(page, fg_color=SHADOW, corner_radius=18, width=1, height=1)
        employee_shadow.grid(row=0, column=0, sticky="nsew", padx=30, pady=(25, 17))
        employee = ctk.CTkFrame(
            page, fg_color=WHITE, border_width=1,
            border_color=BORDER, corner_radius=18,
        )
        employee.grid(row=0, column=0, sticky="nsew", padx=30, pady=(19, 23))
        self.mc_employee_frame = employee
        employee.grid_columnconfigure(0, weight=1)
        employee.grid_rowconfigure(2, weight=1)

        ctk.CTkLabel(
            employee, text="Employee Totals", text_color=TEXT,
            font=ctk.CTkFont(size=24, weight="bold"), anchor="w",
        ).grid(row=0, column=0, padx=22, pady=(20, 3), sticky="w")
        self.mc_employee_summary = ctk.CTkLabel(
            employee,
            text="Create a Uniform Sizing Event Purchase Review to see employee totals.",
            text_color=MUTED, font=ctk.CTkFont(size=13), anchor="w",
            justify="left", wraplength=760,
        )
        self.mc_employee_summary.grid(row=1, column=0, padx=14, pady=(0, 6), sticky="w")

        employee_buttons = ctk.CTkFrame(employee, fg_color="transparent")
        employee_buttons.grid(row=0, column=1, rowspan=2, padx=18, pady=16, sticky="e")
        self.mc_employee_save_button = ctk.CTkButton(
            employee_buttons, text="Save Totals", width=132, height=40,
            fg_color=PURPLE, hover_color=PURPLE_DARK,
            font=ctk.CTkFont(size=13, weight="bold"),
            command=self.save_dashboard_employee_totals, state="disabled",
        )
        self.mc_employee_save_button.pack(side="left", padx=5)
        ctk.CTkButton(
            employee_buttons, text="View PDF", width=105, height=40,
            fg_color=PURPLE_LIGHT, hover_color="#E4D6F5", text_color=PURPLE_DARK,
            command=self.view_employee_totals_pdf,
        ).pack(side="left", padx=5)
        ctk.CTkButton(
            employee_buttons, text="Share / Email", width=125, height=40,
            fg_color=PURPLE_LIGHT, hover_color="#E4D6F5", text_color=PURPLE_DARK,
            command=self.share_employee_totals_pdf,
        ).pack(side="left", padx=5)
        ctk.CTkButton(
            employee_buttons, text="Open Purchase Review", width=158, height=40,
            fg_color=PURPLE_LIGHT, hover_color="#E4D6F5", text_color=PURPLE_DARK,
            command=self.open_latest_review,
        ).pack(side="left", padx=5)

        self.mc_employee_scroll = ctk.CTkScrollableFrame(
            employee, fg_color="#FCFBFE", corner_radius=10,
        )
        self.mc_employee_scroll.grid(row=2, column=0, columnspan=2, padx=18, pady=(0, 10), sticky="nsew")
        self.mc_employee_scroll.grid_columnconfigure(0, weight=1)
        register_scrollable(self, self.mc_employee_scroll)

        self.mc_employee_grand_total = ctk.CTkLabel(
            employee, text="Grand Total  $0.00", text_color=PURPLE_DARK,
            font=ctk.CTkFont(size=17, weight="bold"), anchor="e",
        )
        self.mc_employee_grand_total.grid(row=3, column=0, columnspan=2, padx=22, pady=(2, 18), sticky="e")

    def build_archive_page(self):
        page = self.new_page("archive")
        card = self.section_card(
            page, 0, "Completed Events",
            "Each completed event is automatically stored with its order CSV, Purchase Review workbook, and final purchase-order PDFs."
        )
        self.primary_button(card, "Open Event Archive", self.open_event_archive)

    def build_settings_page(self):
        page = self.new_page("settings")
        scroll = ctk.CTkScrollableFrame(page, fg_color=BG, corner_radius=0)
        self.settings_scroll = scroll
        scroll.grid(row=0, column=0, sticky="nsew")
        scroll.grid_columnconfigure(0, weight=1)
        page.grid_rowconfigure(0, weight=1)
        diagnostics = self.section_card(scroll, 0, "Routing Diagnostics", "Create a workbook explaining how Shopify lines matched Product Master and why information is missing.")
        self.primary_button(diagnostics, "Create Diagnostic Workbook", self.create_routing_diagnostics)

        recovery = self.section_card(
            scroll, 1, "Product Master Backup & Recovery",
            "Create a backup, restore a previous Product Master CSV, or return to the clean catalog included with Professional 4.3.8. The current file is always backed up before a restore."
        )
        recovery.grid_columnconfigure((0, 1, 2), weight=1)
        self.primary_button(recovery, "Back Up Now", self.backup_product_master_now, row=2, column=0, padx=(22, 8), sticky="ew")
        self.secondary_button(recovery, "Restore a Backup", self.restore_product_master_backup, row=2, column=1, padx=8, sticky="ew")
        clean_restore = ctk.CTkButton(
            recovery, text="Restore Clean v10 Master", command=self.restore_clean_product_master,
            height=42, fg_color="#FFF4E8", hover_color="#FBE4CC", border_width=1,
            border_color="#E8B77A", text_color="#8A4B00",
            font=ctk.CTkFont(size=13, weight="bold")
        )
        clean_restore.grid(row=2, column=2, padx=(8, 22), pady=(0, 20), sticky="ew")

        cleanup = self.section_card(scroll, 2, "Catalog Maintenance", "Back up Product Master and remove duplicate records without losing purchasing assignments.")
        self.primary_button(cleanup, "Clean & Backup", self.clean_catalog)

        export = self.section_card(
            scroll, 3, "Optional Excel Export",
            "The normal Purchase Review is completed inside Orchid Purchase Manager. Open the saved Excel workbook only when a detailed offline copy is useful."
        )
        self.primary_button(export, "Open Excel Purchase Review", self.open_latest_review)

        optional = self.section_card(
            scroll, 4, "Safe Product Candidate Preview",
            "Review new product/style numbers found in the selected order export. Professional 4.8.33 does not mass-add order lines to Product Master, so a large import cannot create hundreds of incomplete permanent records."
        )
        self.primary_button(optional, "Create Candidate List", self.process_csv)

        legacy = legacy_product_master_path()
        paths = self.section_card(
            scroll, 5, "Storage Locations",
            f"Active 2.x Product Master file: {live_product_master_path()}\nReports: {REPORTS}\nv12 data (left untouched): {v12_data_dir()}\nOlder data (left untouched): {legacy_data_dir()}"
        )
        self.primary_button(paths, "Open 2.x Data Folder", lambda: subprocess.run(["open", str(DATA)], check=False), row=2)
        if legacy.parent.exists():
            self.secondary_button(paths, "Open Older Data Folder", lambda: subprocess.run(["open", str(legacy.parent)], check=False), row=2)

        motion = self.section_card(
            scroll, 6, "Interface Motion",
            "Use the subtle purple radial pulse around the official Orchid flower when navigating or completing actions. Reduced-motion system preferences are always respected."
        )
        self.interaction_animation_var = ctk.BooleanVar(value=bool(self.interaction_animations_enabled))
        ctk.CTkSwitch(
            motion, text="Enable Orchid interaction animation", variable=self.interaction_animation_var,
            command=self._set_interaction_animations_enabled, progress_color=PURPLE,
            font=ctk.CTkFont(size=13, weight="bold"), text_color=TEXT,
        ).grid(row=2, column=0, padx=22, pady=(0, 20), sticky="w")

        about = self.section_card(
            scroll, 7, "About Orchid Purchase Manager",
            "Professional 4.8.33\nTargets customer instructions to the correct lines, keeps blanket embroidery and every screen-print note in review, and restores Logo access."
        )
        ctk.CTkLabel(
            about, text="Version 4.8.33 Pro", text_color=PURPLE_DARK,
            font=ctk.CTkFont(size=18, weight="bold"), anchor="w",
        ).grid(row=2, column=0, padx=22, pady=(0, 20), sticky="w")

    def refresh_current_event_page(self, snapshot: dict | None = None):
        if not hasattr(self, "current_event_title"):
            return
        snapshot = snapshot or self.last_snapshot or {}
        self._refresh_current_event_progress(snapshot)
        if self.selected_csv or (self.last_review_workbook and self.last_review_workbook.exists()):
            self.current_event_card.configure(height=500)
            self.current_event_card_shadow.configure(height=506)
            event = self.current_event_name or snapshot.get("event_name") or "Active Purchase Packet"
            self.current_event_title.configure(text=event)
            self.current_event_summary_title.configure(text=event)
            self.dashboard_summary_title.configure(text=event)
            self.header_chip.configure(text=event if len(event) <= 34 else event[:31] + "...")
            self.current_event_body.configure(text="Manage the active order import and its Purchase Review from one place.")
            lines = int(snapshot.get("line_count", self.imported_line_count or 0) or 0)
            vendor_count = int(snapshot.get("vendor_count", 0) or 0)
            meta_parts = [time.strftime("%b %d, %Y")]
            if lines:
                meta_parts.append(f"{lines:,} Lines")
            if vendor_count:
                meta_parts.append(f"{vendor_count:,} Vendors")
            self.shell_event_meta.configure(text="  •  ".join(meta_parts))
            details = [f"CSV: {self.selected_csv.name}" if self.selected_csv else f"Review: {self.last_review_workbook.name}"]
            if lines:
                details.append(f"{lines:,} purchase lines")
            if self.selected_csv and self.last_review_workbook and self.last_review_workbook.exists():
                details.append(f"Review: {self.last_review_workbook.name}")
            self.import_file_label.configure(text="  •  ".join(details[:2]))
            self.current_event_details.configure(text="\n".join(details[2:]))
            self.current_event_primary.grid_configure(row=0, column=0, columnspan=1, padx=6, pady=6)
            if self.last_review_workbook and self.last_review_workbook.exists():
                product_setup_count = sum(
                    1 for issue in snapshot.get("issues", [])
                    if _clean(issue.get("fix_in", issue.get("source", ""))).casefold() == "product master"
                )
                if product_setup_count:
                    self.current_event_primary.configure(
                        text="Continue Product Setup", command=lambda: self.show_page("master")
                    )
                elif self._review_is_stale():
                    self.current_event_primary.configure(
                        text="Apply Product Master Updates", command=self.regenerate_current_review
                    )
                else:
                    self.current_event_primary.configure(
                        text="Continue Purchase Review", command=lambda: self.show_page("review")
                    )
            else:
                self.current_event_primary.configure(
                    text="Process Imported Orders", command=self.create_review_workbook
                )
            self.current_event_replace.grid(row=0, column=1, padx=6, pady=6)
            self.current_event_discard.grid()
            if self.last_review_workbook and self.last_review_workbook.exists():
                self.current_event_regenerate.configure(state="normal")
                self.current_event_regenerate.grid()
            else:
                self.current_event_regenerate.grid_remove()
        else:
            self.current_event_card.configure(height=376)
            self.current_event_card_shadow.configure(height=382)
            self.current_event_title.configure(text="No Active Purchase Packet")
            self.current_event_summary_title.configure(text="No Active Purchase Packet")
            self.dashboard_summary_title.configure(text="No Active Purchase Packet")
            self.header_chip.configure(text="No Active Purchase Packet")
            self.current_event_body.configure(text="Import the Shopify or Report Toaster CSV for the orders you want to process. Product Master and saved settings are preserved.")
            self.shell_event_meta.configure(text=time.strftime("%b %d, %Y"))
            self.import_file_label.configure(text="No order CSV selected")
            self.current_event_details.configure(text="")
            self.current_event_primary.grid_configure(row=0, column=0, columnspan=2, padx=6, pady=6)
            self.current_event_primary.configure(text="Import Shopify CSV", command=self.choose_csv)
            self.current_event_replace.grid_remove()
            self.current_event_regenerate.grid_remove()
            self.current_event_discard.grid_remove()

    def replace_current_csv(self):
        if self.selected_csv and not messagebox.askyesno(
            "Replace Imported CSV",
            "Replace the current order import?\n\nThe active Purchase Review and generated purchase orders will be cleared from the workflow. Product Master, settings, and the original CSV file will not be changed.",
        ):
            return
        self._clear_active_packet_state()
        self.choose_csv()

    def _clear_active_packet_state(self):
        try:
            ACTIVE_PACKET_RESET_FILE.parent.mkdir(parents=True, exist_ok=True)
            ACTIVE_PACKET_RESET_FILE.write_text(json.dumps({"reset_time": time.time()}, indent=2), encoding="utf-8")
            if REVIEW_STATE_FILE.exists():
                REVIEW_STATE_FILE.unlink()
        except Exception:
            pass
        self.selected_csv = None
        self.current_event_name = ""
        self.current_mode = ""
        self.last_review_workbook = None
        self.last_purchase_order_dir = None
        self.imported_line_count = 0
        self.review_needed = None
        self.last_snapshot = {}

    def discard_current_event(self):
        if not (self.selected_csv or self.last_review_workbook or self.last_purchase_order_dir):
            messagebox.showinfo("No Current Event", "There is no active purchase packet to discard.")
            return
        if not messagebox.askyesno(
            "Discard Current Event",
            "Discard the active purchase packet?\n\nThis removes the imported CSV from Orchid's current workflow and clears its active review and purchase-order references. It does not delete the original CSV, Product Master, settings, or completed archived events.",
        ):
            return
        self._suppress_empty_import_prompt = True
        self._clear_active_packet_state()
        self.refresh_current_event_page({})
        self.refresh_dashboard()
        self.show_page("dashboard")
        self.after_idle(lambda: (self.refresh_current_event_page({}), self.refresh_dashboard()))

    def start_new_purchase_packet(self):
        if not messagebox.askyesno(
            "Start New Event",
            "Start a new purchase packet?\n\nThe completed event and all generated files will remain safely archived.",
        ):
            return
        self._suppress_empty_import_prompt = True
        self._clear_active_packet_state()
        self.refresh_current_event_page({})
        self.refresh_dashboard()
        self.show_page("dashboard")
        self.after_idle(lambda: (self.refresh_current_event_page({}), self.refresh_dashboard()))

    def open_latest_purchase_orders(self):
        if self.last_purchase_order_dir and self.last_purchase_order_dir.exists():
            subprocess.run(["open", str(self.last_purchase_order_dir)], check=False)
        else:
            self.open_purchase_orders_folder()

    # ---------- actions ----------
    def _show_create_review_dashboard(self) -> None:
        """Return to and fully repaint Step 1 after the macOS file picker closes."""
        if self.selected_csv and not self.last_review_workbook:
            self.refresh_dashboard()
            self.show_page("dashboard")
            try:
                self.update_idletasks()
                self.update()
                self.lift()
                self.focus_force()
            except Exception:
                pass

    def choose_csv(self):
        downloads = Path.home() / "Downloads"
        path = filedialog.askopenfilename(
            title="Import Shopify or Report Toaster Orders (.CSV)",
            initialdir=str(downloads if downloads.exists() else Path.home()),
            filetypes=[("Order CSV files", "*.csv")],
        )
        if not path:
            return
        candidate = Path(path)
        try:
            normalized = normalize_order_export(candidate)
            parsed = parse_shopify_orders(
                candidate, live_product_master_path(), normalized_orders=normalized
            )
            self._remember_import_frames(candidate, normalized, parsed)
        except Exception as error:
            messagebox.showerror(
                "Unable to Import Order CSV",
                f"{error}\n\nNo event was created. Update the export columns and try again.",
            )
            return

        candidate_signature = self._source_file_signature(candidate)
        state = self._load_review_state()
        active_review = bool(self.last_review_workbook and self.last_review_workbook.exists())
        saved_source = Path(state.get("source_csv", "")).expanduser() if state.get("source_csv") else None
        saved_signature = _clean(state.get("source_signature", ""))
        if not saved_signature and saved_source and saved_source.exists():
            saved_signature = self._source_file_signature(saved_source)
        same_source = bool(
            active_review and (
                (candidate_signature and saved_signature and candidate_signature == saved_signature)
                or (saved_source and saved_source.exists() and candidate.resolve() == saved_source.resolve())
            )
        )
        if active_review and same_source:
            resume = messagebox.askyesnocancel(
                "Existing Purchase Review Found",
                "This CSV matches the active purchase packet.\n\n"
                "Yes — Resume Existing Purchase Review and keep all saved corrections.\n"
                "No — Replace it and start the Purchase Review over.\n"
                "Cancel — Leave the current packet unchanged.",
                default=messagebox.YES,
            )
            if resume is None:
                return
            if resume:
                self.selected_csv = candidate
                self.imported_line_count = len(parsed)
                self._update_review_state(source_csv=str(candidate), source_signature=candidate_signature)
                self.dashboard_status.configure(
                    text=f"Resumed existing Purchase Review for {candidate.name}. Saved corrections were preserved.",
                    text_color=SUCCESS,
                )
                self.refresh_dashboard()
                self.show_page("review")
                return
            self._suppress_empty_import_prompt = True
            self._clear_active_packet_state()
        elif active_review:
            replace = messagebox.askyesno(
                "Replace Active Purchase Packet?",
                "A different Purchase Review is already active. Replace it with the selected CSV?\n\n"
                "The existing workbook will remain in Reports, but it will no longer be the active packet.",
                default=messagebox.NO,
            )
            if not replace:
                return
            self._suppress_empty_import_prompt = True
            self._clear_active_packet_state()

        self.selected_csv = candidate
        self.current_event_name = ""
        self.current_mode = ""
        self.last_review_workbook = None
        self.last_purchase_order_dir = None
        self.review_needed = None
        self.imported_line_count = len(parsed)
        source_name = str(parsed.attrs.get("import_source", "Order CSV") or "Order CSV")
        removed = int(parsed.attrs.get("removed_rows_excluded", 0) or 0)
        adjusted = int(parsed.attrs.get("quantity_adjustments_applied", 0) or 0)
        status = f"Selected {self.selected_csv.name} • {source_name} • {self.imported_line_count:,} purchase lines"
        if removed:
            status += f" • {removed:,} removed item row(s) excluded"
        if adjusted:
            status += f" • {adjusted:,} edited quantity row(s) adjusted"
        self.dashboard_status.configure(text=status, text_color=SUCCESS)
        self.refresh_dashboard()
        self._show_create_review_dashboard()
        self.after_idle(self._show_create_review_dashboard)
        self.after(100, self._show_create_review_dashboard)
        self.after(260, self._show_create_review_dashboard)

    def process_csv(self):
        if not self.selected_csv:
            self.choose_csv()
        if not self.selected_csv:
            return
        try:
            normalized, parsed = self._cached_import_frames(self.selected_csv)
            if parsed is None:
                normalized = normalize_order_export(self.selected_csv)
                parsed = parse_shopify_orders(
                    self.selected_csv, live_product_master_path(), normalized_orders=normalized
                )
                self._remember_import_frames(self.selected_csv, normalized, parsed)
            result = save_product_candidate_list(live_product_master_path(), parsed, REPORTS)
            self.imported_line_count = len(parsed)
            self.dashboard_status.configure(
                text=f"Candidate preview created with {result['candidate_count']} genuinely new product/style number(s).",
                text_color=SUCCESS,
            )
            messagebox.showinfo(
                "Product Candidate List Created",
                f"Shopify product lines reviewed: {len(parsed)}\n"
                f"New product/style candidates: {result['candidate_count']}\n\n"
                "No products were added to Product Master. This prevents a large order import from creating hundreds of incomplete permanent records.\n\n"
                f"Saved as:\n{result['output_path']}",
            )
            self.refresh_dashboard()
        except Exception as error:
            messagebox.showerror("Unable to Create Candidate List", str(error))

    @staticmethod
    def _pid_is_running(pid: int) -> bool:
        if pid <= 0:
            return False
        try:
            os.kill(pid, 0)
            return True
        except (ProcessLookupError, PermissionError, OSError):
            return False

    def _existing_product_master_pid(self) -> int:
        try:
            pid = int(PRODUCT_MASTER_PID_FILE.read_text(encoding="utf-8").strip())
        except Exception:
            return 0
        if self._pid_is_running(pid):
            return pid
        try:
            PRODUCT_MASTER_PID_FILE.unlink(missing_ok=True)
        except Exception:
            pass
        return 0

    @staticmethod
    def _activate_process(pid: int) -> None:
        if pid <= 0 or sys.platform != "darwin":
            return
        script = (
            'tell application "System Events" to set frontmost of '
            f'(first process whose unix id is {pid}) to true'
        )
        subprocess.run(
            ["/usr/bin/osascript", "-e", script],
            check=False,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )

    def _send_product_master_request(self, prefill_product: str = "", search_product: str = "", setup_product: str = "") -> None:
        action = "show"
        value = ""
        if prefill_product:
            action, value = "new", str(prefill_product)
        elif setup_product:
            action, value = "setup", str(setup_product)
        elif search_product:
            action, value = "search", str(search_product)
        payload = {"action": action, "value": value, "requested_at": time.time()}
        PRODUCT_MASTER_REQUEST_FILE.parent.mkdir(parents=True, exist_ok=True)
        PRODUCT_MASTER_REQUEST_FILE.write_text(json.dumps(payload), encoding="utf-8")

    def _set_product_master_opening_state(self, opening: bool) -> None:
        self.product_master_launching = opening
        widgets = [
            getattr(self, "mc_master_button", None),
            getattr(self, "master_page_open_button", None),
            getattr(self, "mc_banner_button", None),
            getattr(self, "mc_work_open_button", None),
            getattr(self, "mc_action_button", None),
        ]
        for widget in widgets:
            if widget is None:
                continue
            try:
                current = str(widget.cget("text"))
                if "Product Master" not in current and current != "Opening Product Master…":
                    continue
                widget.configure(
                    text="Opening Product Master…" if opening else "Open Product Master",
                    state="disabled" if opening else "normal",
                )
            except Exception:
                pass

    def _finish_product_master_launch(self, pid: int) -> None:
        if self._pid_is_running(pid):
            self._activate_process(pid)
            self.dashboard_status.configure(
                text="Product Master is open. Repeated clicks will bring the same window forward.",
                text_color=PURPLE,
            )
        self._set_product_master_opening_state(False)
        self.refresh_dashboard()

    def open_product_master(self, prefill_product: str = "", search_product: str = "", setup_product: str = ""):
        try:
            self._sync_current_product_candidates()
            self._send_product_master_request(prefill_product, search_product, setup_product)
            existing_pid = self._existing_product_master_pid()
            if existing_pid:
                self._product_master_was_open = True
                self._activate_process(existing_pid)
                self.after(180, lambda pid=existing_pid: self._activate_process(pid))
                self.dashboard_status.configure(
                    text="Product Master was already open and has been brought to the front.",
                    text_color=PURPLE,
                )
                return

            if self.product_master_launching:
                self.dashboard_status.configure(
                    text="Product Master is still opening. Please wait…",
                    text_color=PURPLE,
                )
                return

            self._set_product_master_opening_state(True)
            command = [sys.executable, "--product-master"] if getattr(sys, "frozen", False) else [sys.executable, str(Path(__file__).resolve()), "--product-master"]
            if prefill_product:
                command.extend(["--new-product", str(prefill_product)])
            elif setup_product:
                command.extend(["--setup-product", str(setup_product)])
            elif search_product:
                command.extend(["--search-product", str(search_product)])
            log_file = open(DATA / "product_master_launch.log", "a", encoding="utf-8")
            process = subprocess.Popen(
                command,
                cwd=str(PROJECT),
                stdout=log_file,
                stderr=log_file,
                start_new_session=True,
            )
            self.product_master_process = process
            self._product_master_was_open = True
            PRODUCT_MASTER_PID_FILE.write_text(str(process.pid), encoding="utf-8")
            self.dashboard_status.configure(text="Opening Product Master…", text_color=PURPLE)
            # Product Master can take a few seconds to load a large catalog. Try
            # activation more than once so it appears in front when it is ready.
            for delay in (500, 1300, 2600, 4200):
                self.after(delay, lambda pid=process.pid: self._activate_process(pid))
            self.after(4600, lambda pid=process.pid: self._finish_product_master_launch(pid))
        except Exception as error:
            self._set_product_master_opening_state(False)
            messagebox.showerror("Unable to Open Product Master", str(error))

    def add_new_product(self):
        value = simpledialog.askstring(
            "Add New Product",
            "Enter the missing product/style number.\n\nExample: 1104",
            parent=self,
        )
        value = " ".join((value or "").split())
        if value:
            self.open_product_master(prefill_product=value)

    def choose_purchase_order_mode(self):
        dialog = ctk.CTkToplevel(self)
        dialog.title("Choose Purchase Order Mode")
        dialog.geometry("620x390")
        dialog.resizable(False, False)
        dialog.configure(fg_color=BG)
        dialog.transient(self)
        dialog.grab_set()
        choice = {"value": None}

        ctk.CTkLabel(dialog, text="Choose Purchase Order Mode", text_color=TEXT,
                     font=ctk.CTkFont(size=24, weight="bold")).pack(anchor="w", padx=28, pady=(26, 6))
        ctk.CTkLabel(dialog, text="This choice controls how final vendor purchase orders are grouped.",
                     text_color=MUTED, font=ctk.CTkFont(size=13)).pack(anchor="w", padx=28, pady=(0, 18))

        def select(value):
            choice["value"] = value
            dialog.destroy()

        uniform = ctk.CTkFrame(dialog, fg_color=WHITE, border_width=1, border_color=BORDER, corner_radius=12)
        uniform.pack(fill="x", padx=28, pady=7)
        ctk.CTkLabel(uniform, text=UNIFORM_SIZING_EVENT, text_color=PURPLE_DARK,
                     font=ctk.CTkFont(size=16, weight="bold"), anchor="w").pack(fill="x", padx=18, pady=(14, 3))
        ctk.CTkLabel(uniform, text="One report per vendor and decoration type. Decoration colors are sections within each report.",
                     text_color=MUTED, font=ctk.CTkFont(size=12), wraplength=520, justify="left", anchor="w").pack(fill="x", padx=18, pady=(0, 10))
        ctk.CTkButton(uniform, text="Use Uniform Sizing Event", fg_color=PURPLE, hover_color=PURPLE_DARK,
                      command=lambda: select(UNIFORM_SIZING_EVENT)).pack(fill="x", padx=18, pady=(0, 14))

        general = ctk.CTkFrame(dialog, fg_color=WHITE, border_width=1, border_color=BORDER, corner_radius=12)
        general.pack(fill="x", padx=28, pady=7)
        ctk.CTkLabel(general, text="All Orders", text_color=PURPLE_DARK,
                     font=ctk.CTkFont(size=16, weight="bold"), anchor="w").pack(fill="x", padx=18, pady=(14, 3))
        ctk.CTkLabel(general, text="One combined report per vendor with decoration types and colors separated inside.",
                     text_color=MUTED, font=ctk.CTkFont(size=12), wraplength=520, justify="left", anchor="w").pack(fill="x", padx=18, pady=(0, 10))
        ctk.CTkButton(general, text="Use All Orders", fg_color=PURPLE, hover_color=PURPLE_DARK,
                      command=lambda: select(GENERAL_SALES_PERIOD)).pack(fill="x", padx=18, pady=(0, 14))

        dialog.protocol("WM_DELETE_WINDOW", dialog.destroy)
        dialog.wait_window()
        report_mode = choice["value"]
        if not report_mode:
            return None
        decoration_fulfillment = self.choose_decoration_fulfillment()
        if not decoration_fulfillment:
            return None
        if report_mode == UNIFORM_SIZING_EVENT:
            event_name = simpledialog.askstring(
                "Name Uniform Sizing Event",
                "Enter the event name that should appear on the workbook and final purchase orders.\n\nExample: Water Quality 2026",
                parent=self,
            )
            event_name = " ".join((event_name or "").split())
            if not event_name:
                messagebox.showwarning("Event Name Required", "Please enter a name for the Uniform Sizing Event.", parent=self)
                return None
            return report_mode, event_name, decoration_fulfillment
        if report_mode:
            event_name = simpledialog.askstring(
                "Name All Orders Period",
                "Optional: enter a helpful name for these daily/company orders.\n\nExample: July 8-12 Orders",
                parent=self,
            )
            return report_mode, " ".join((event_name or "").split()), decoration_fulfillment
        return None

    def choose_decoration_fulfillment(self):
        dialog = ctk.CTkToplevel(self)
        dialog.title("Choose Decoration Workflow")
        dialog.geometry("680x430")
        dialog.resizable(False, False)
        dialog.configure(fg_color=BG)
        dialog.transient(self)
        dialog.grab_set()
        choice = {"value": None}

        ctk.CTkLabel(
            dialog, text="How will this order be decorated?", text_color=TEXT,
            font=ctk.CTkFont(size=24, weight="bold"),
        ).pack(anchor="w", padx=28, pady=(26, 6))
        ctk.CTkLabel(
            dialog,
            text="This controls which garments appear on the in-house receiving report and which appear on the outside-decorator job report.",
            text_color=MUTED, font=ctk.CTkFont(size=13), wraplength=610, justify="left",
        ).pack(anchor="w", padx=28, pady=(0, 18))

        def select(value):
            choice["value"] = value
            dialog.destroy()

        standard = ctk.CTkFrame(dialog, fg_color=WHITE, border_width=1, border_color=BORDER, corner_radius=12)
        standard.pack(fill="x", padx=28, pady=7)
        ctk.CTkLabel(
            standard, text=STANDARD_ORCHID_WORKFLOW, text_color=PURPLE_DARK,
            font=ctk.CTkFont(size=16, weight="bold"), anchor="w",
        ).pack(fill="x", padx=18, pady=(14, 3))
        ctk.CTkLabel(
            standard,
            text="Normal Orchid workflow: embroidery and blank garments come to Orchid; screen printing is sent to the outside decorator.",
            text_color=MUTED, font=ctk.CTkFont(size=12), wraplength=580, justify="left", anchor="w",
        ).pack(fill="x", padx=18, pady=(0, 10))
        ctk.CTkButton(
            standard, text="Use Standard Orchid Workflow", fg_color=PURPLE, hover_color=PURPLE_DARK,
            command=lambda: select(STANDARD_ORCHID_WORKFLOW),
        ).pack(fill="x", padx=18, pady=(0, 14))

        outsourced = ctk.CTkFrame(dialog, fg_color=WHITE, border_width=1, border_color=BORDER, corner_radius=12)
        outsourced.pack(fill="x", padx=28, pady=7)
        ctk.CTkLabel(
            outsourced, text=ENTIRE_ORDER_OUTSOURCED, text_color=PURPLE_DARK,
            font=ctk.CTkFont(size=16, weight="bold"), anchor="w",
        ).pack(fill="x", padx=18, pady=(14, 3))
        ctk.CTkLabel(
            outsourced,
            text="All decorated garments are sent to the outside decorator. Use Do Not Outsource only for rare lines that must come to Orchid.",
            text_color=MUTED, font=ctk.CTkFont(size=12), wraplength=580, justify="left", anchor="w",
        ).pack(fill="x", padx=18, pady=(0, 10))
        ctk.CTkButton(
            outsourced, text="Use Entire Order Outsourced", fg_color=PURPLE, hover_color=PURPLE_DARK,
            command=lambda: select(ENTIRE_ORDER_OUTSOURCED),
        ).pack(fill="x", padx=18, pady=(0, 14))

        dialog.protocol("WM_DELETE_WINDOW", dialog.destroy)
        dialog.wait_window()
        return choice["value"]

    def _save_review_state(self, report_mode: str, event_name: str, decoration_fulfillment: str):
        state = {
            "source_csv": str(self.selected_csv or ""),
            "source_signature": self._source_file_signature(self.selected_csv),
            "report_mode": report_mode,
            "event_name": event_name,
            "decoration_fulfillment": normalize_decoration_fulfillment(decoration_fulfillment),
            "app_source": str(Path(__file__).resolve()),
        }
        REVIEW_STATE_FILE.write_text(json.dumps(state, indent=2), encoding="utf-8")

    def create_review_workbook(self):
        from modules.review_workbook import generate_review_workbook

        if not self.selected_csv or not self.selected_csv.exists():
            messagebox.showinfo(
                "Import Orders First",
                "No Shopify orders have been imported yet.\n\nGo to Import Orders and select the Shopify Orders CSV before processing the import.",
            )
            self.show_page("import")
            return
        selection = self.choose_purchase_order_mode()
        if not selection:
            return
        report_mode, event_name, decoration_fulfillment = selection
        try:
            master_path = live_product_master_path()
            before_signature = product_master_signature(master_path)
            normalized_orders, parsed_orders = self._cached_import_frames(self.selected_csv)
            self._save_review_state(report_mode, event_name, decoration_fulfillment)
            result = generate_review_workbook(self.selected_csv, master_path, REPORTS,
                                              report_mode=report_mode, event_name=event_name,
                                              decoration_fulfillment=decoration_fulfillment,
                                              regenerate_helper=REGENERATE_HELPER,
                                              normalized_orders=normalized_orders,
                                              parsed_orders=parsed_orders)
            self.current_event_name = result["event_name"] or "General Sales Period"
            self.current_mode = result["report_mode"]
            self.current_decoration_fulfillment = result.get("decoration_fulfillment", decoration_fulfillment)
            self.last_review_workbook = Path(result["output_path"])
            self._candidate_sync_workbook = ""
            candidate_sync = self._sync_current_product_candidates()
            # Candidate records must exist before the final workflow decision is
            # made. Rebuild once so Step 2 sees every unknown/incomplete style and
            # Purchase Review cannot bypass Product Master.
            if candidate_sync.get("changed"):
                first_workbook = self.last_review_workbook
                result = generate_review_workbook(
                    self.selected_csv, live_product_master_path(), REPORTS,
                    report_mode=report_mode, event_name=event_name,
                    decoration_fulfillment=decoration_fulfillment,
                    regenerate_helper=REGENERATE_HELPER, previous_review_path=first_workbook,
                    normalized_orders=normalized_orders,
                )
                self.last_review_workbook = Path(result["output_path"])
                self._candidate_sync_workbook = str(self.last_review_workbook.resolve())
                try:
                    if first_workbook.exists() and first_workbook != self.last_review_workbook:
                        first_workbook.unlink()
                except Exception:
                    pass
            self._reset_review_navigation()
            self.imported_line_count = int(result["lines"])
            self.review_needed = int(result.get("review_decisions", result["review_lines"]))
            missing_products = result.get("missing_product_numbers", [])
            status_text = f"Created {result['report_mode']} workbook with {result['lines']} purchase lines."
            if missing_products:
                status_text += f" Added {len(missing_products)} unknown product number(s) to the Product Master setup queue."
            self.dashboard_status.configure(
                text=status_text,
                text_color=WARNING if missing_products else SUCCESS,
            )
            messagebox.showinfo(
                "Imported Orders Processed",
                f"Purchase order mode: {result['report_mode']}\n"
                f"Decoration workflow: {result.get('decoration_fulfillment', STANDARD_ORCHID_WORKFLOW)}\n"
                + (f"Event name: {result['event_name']}\n" if result['event_name'] else "")
                + f"Purchase lines: {result['lines']}\n"
                f"Total quantity: {result['quantity']}\n"
                f"Product Master products needing setup: {result.get('product_master_decisions', 0)}\n"
                f"Order-specific Purchase Review decisions: {result.get('purchase_review_decisions', 0)}\n"
                f"Total affected order rows: {result['review_lines']}\n"
                f"Automatically routed rows: {result.get('inferred_routing_lines', 0)}\n"
                f"Rows with purchase instructions: {result['notes_lines']}\n"
                f"Product Master reloaded: {before_signature['modified'] or 'just now'}\n"
                + (f"Unknown product numbers to correct or resolve: {', '.join(result.get('missing_product_numbers', []))}\n" if result.get('missing_product_numbers') else "")
                + (f"\nNew Product Master candidates added to the setup queue: {candidate_sync.get('added_styles', 0)}\n" if candidate_sync.get("added_styles") else "")
                + "\nContinue inside Orchid Purchase Manager. Excel is no longer required for the normal workflow.\n\n"
                + f"An optional Excel copy was saved as:\n{result['output_path']}",
            )
            self.refresh_dashboard()
            snapshot = self._fresh_snapshot()
            product_setup_count = sum(
                1 for issue in snapshot.get("issues", [])
                if _clean(issue.get("fix_in", issue.get("source", ""))).casefold() == "product master"
            )
            if product_setup_count or int(result.get("product_master_decisions", 0)):
                self.show_page("master")
            else:
                # Step 3 is available only after Step 2 has no unresolved permanent setup.
                self.show_page("review")
        except Exception as error:
            messagebox.showerror("Unable to Create Review Workbook", str(error))

    def create_routing_diagnostics(self):
        from modules.routing_diagnostics import generate_routing_diagnostic_workbook

        if not self.selected_csv:
            self.choose_csv()
        if not self.selected_csv:
            return
        try:
            result = generate_routing_diagnostic_workbook(self.selected_csv, live_product_master_path(), REPORTS)
            self.dashboard_status.configure(
                text=f"Diagnostic complete: {result['complete_matches']} complete matches; {result['issue_lines']} require attention.",
                text_color=SUCCESS,
            )
            messagebox.showinfo(
                "Routing Diagnostic Workbook Created",
                f"Product lines analyzed: {result['total_lines']}\n"
                f"Exact complete matches: {result['complete_matches']}\n"
                f"No Product Master match: {result['new_products']}\n"
                f"Order blockers: {result['mismatches']}\n"
                f"Matched records missing routing: {result['incomplete_matches']}\n\n"
                f"Saved as:\n{result['output_path']}",
            )
            subprocess.run(["open", str(result["output_path"])], check=False)
        except Exception as error:
            messagebox.showerror("Unable to Create Routing Diagnostics", str(error))

    def _run_final_purchase_order_generation(self, workbook_path: Path, po_overrides=None):
        from modules.final_po_generator import generate_final_purchase_orders

        try:
            overrides = po_overrides if po_overrides is not None else load_po_overrides(DATA, workbook_path)
            result = generate_final_purchase_orders(Path(workbook_path), REPORTS, po_number_overrides=overrides)
            self.last_review_workbook = Path(workbook_path)
            self.last_purchase_order_dir = Path(result["output_dir"])
            self.current_event_name = result["event_name"] or self.current_event_name or "General Sales Period"
            self.current_mode = result["report_mode"]
            self.review_needed = int(result["review_lines"])
            archive_dir = archive_completed_event(
                reports_root=REPORTS,
                event_name=self.current_event_name,
                review_workbook=Path(workbook_path),
                purchase_order_dir=self.last_purchase_order_dir,
                source_csv=self.selected_csv,
            )
            self.dashboard_status.configure(
                text=(
                    f"Created {result['report_count']} purchase-order document(s), "
                    f"plus {result.get('decoration_report_count', 0)} decoration report(s)."
                ),
                text_color=SUCCESS,
            )
            self.refresh_dashboard()
            self.show_page("dashboard")
            self.after(220, self._play_orchid_success_animation)
        except Exception as error:
            messagebox.showerror("Unable to Create Final Purchase Orders", str(error))

    def create_final_purchase_orders(self):
        initial = REPORTS / "Review Workbooks"
        if not initial.exists():
            initial = REPORTS / "review_workbooks"
        path = filedialog.askopenfilename(
            title="Choose Edited Orchid Purchase Review Workbook",
            initialdir=str(initial if initial.exists() else REPORTS),
            filetypes=[("Excel workbooks", "*.xlsx")],
        )
        if not path:
            return
        workbook = Path(path)
        snapshot = load_mission_control_snapshot(workbook, DATA)
        review_count = int(snapshot.get("review_count", 0))
        blocked_route_count = int(snapshot.get("blocked_route_count", 0))
        if review_count or blocked_route_count:
            messagebox.showwarning(
                "Purchase Review Incomplete",
                "The selected workbook did not pass the final readiness check.\n\n"
                + (self._blocker_summary(snapshot) or f"{review_count} decision(s) and {blocked_route_count} blocked route(s) remain."),
            )
            return
        self.last_review_workbook = workbook
        self.open_po_number_editor(generate_after_save=True, workbook_path=workbook)

    def open_event_archive(self):
        folder = REPORTS / "Event Archive"
        folder.mkdir(parents=True, exist_ok=True)
        subprocess.run(["open", str(folder)], check=False)

    def backup_product_master_now(self):
        try:
            backup = backup_file(live_product_master_path(), label="manual")
            if not backup:
                messagebox.showinfo("Product Master Backup", "There is no active Product Master file to back up.")
                return
            messagebox.showinfo("Product Master Backed Up", f"Backup saved as:\n{backup}")
        except Exception as error:
            messagebox.showerror("Unable to Back Up Product Master", str(error))

    def restore_product_master_backup(self):
        initial = DATA / "backups"
        initial.mkdir(parents=True, exist_ok=True)
        selected = filedialog.askopenfilename(
            title="Choose Product Master Backup",
            initialdir=str(initial),
            filetypes=[("CSV files", "*.csv")],
        )
        if not selected:
            return
        if not messagebox.askyesno(
            "Restore Product Master?",
            "The current Product Master will be backed up, then replaced with the selected CSV. Continue?",
            parent=self,
        ):
            return
        try:
            result = restore_product_master(live_product_master_path(), Path(selected))
            messagebox.showinfo(
                "Product Master Restored",
                f"Restored {result['records']} records across {result['styles']} products/styles.\n\n"
                f"Previous active file backed up as:\n{result['backup'] or 'No previous file'}",
            )
            self.refresh_dashboard()
        except Exception as error:
            messagebox.showerror("Unable to Restore Product Master", str(error))

    def restore_clean_product_master(self):
        if not messagebox.askyesno(
            "Restore Clean Product Master?",
            "This will replace the active 2.x Product Master with the clean 166-record Product Master included with the original v10.0 package. The current file will be backed up first. Continue?",
            parent=self,
        ):
            return
        try:
            result = restore_seed_product_master(live_product_master_path())
            messagebox.showinfo(
                "Clean Product Master Restored",
                f"Restored {result['records']} records across {result['styles']} products/styles.\n\n"
                f"Previous active file backed up as:\n{result['backup'] or 'No previous file'}",
            )
            self.refresh_dashboard()
        except Exception as error:
            messagebox.showerror("Unable to Restore Clean Product Master", str(error))

    def clean_catalog(self):
        try:
            result = clean_product_master(live_product_master_path(), make_backup=True)
            self.dashboard_status.configure(text=f"Product Master cleaned. {result['after']} unique records remain.", text_color=SUCCESS)
            backup_text = str(result["backup"]) if result.get("backup") else "No backup was needed."
            messagebox.showinfo(
                "Product Master Cleaned",
                f"Rows before: {result['before']}\n"
                f"Rows after: {result['after']}\n"
                f"Duplicates removed: {result['removed']}\n\n"
                f"Backup saved as:\n{backup_text}",
            )
            self.refresh_dashboard()
        except Exception as error:
            messagebox.showerror("Unable to Clean Product Master", str(error))

    def open_latest_review(self):
        if not self.last_review_workbook or not self.last_review_workbook.exists():
            self.discover_existing_work()
        if self.last_review_workbook and self.last_review_workbook.exists():
            subprocess.run(["open", str(self.last_review_workbook)], check=False)
        else:
            messagebox.showinfo("No Purchase Review", "Create a Purchase Review first.")

    def open_review_folder(self):
        folder = REPORTS / "Review Workbooks"
        if not folder.exists():
            folder = REPORTS / "review_workbooks"
        folder.mkdir(parents=True, exist_ok=True)
        subprocess.run(["open", str(folder)], check=False)

    def open_purchase_orders_folder(self):
        folder = REPORTS / "Purchase Orders"
        folder.mkdir(parents=True, exist_ok=True)
        subprocess.run(["open", str(folder)], check=False)


def _save_and_close_active_excel_workbook() -> None:
    """Best-effort save/close for the workbook that launched regeneration."""
    if sys.platform != "darwin":
        return
    script = (
        'tell application "Microsoft Excel"\n'
        'if it is running then\n'
        'try\n'
        'save active workbook\n'
        'close active workbook saving yes\n'
        'end try\n'
        'end if\n'
        'end tell'
    )
    subprocess.run(["/usr/bin/osascript", "-e", script], check=False,
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    time.sleep(0.8)


def _show_regeneration_error(message: str) -> None:
    if sys.platform == "darwin":
        safe = str(message).replace('"', "'")
        subprocess.run([
            "/usr/bin/osascript", "-e",
            f'display alert "Purchase Review could not be regenerated" message "{safe}" as critical',
        ], check=False)
    else:
        print(message, file=sys.stderr)


def regenerate_review_from_saved_state(request_path: Path | None = None):
    """Regenerate from a workbook request file or the app's saved state."""
    from modules.review_workbook import generate_review_workbook

    try:
        if request_path:
            request = Path(request_path).expanduser().resolve()
            if not request.exists():
                raise FileNotFoundError(f"Regeneration request not found: {request}")
            state = json.loads(request.read_text(encoding="utf-8"))
        else:
            if not REVIEW_STATE_FILE.exists():
                raise FileNotFoundError("No saved Purchase Review state was found.")
            state = json.loads(REVIEW_STATE_FILE.read_text(encoding="utf-8"))

        source_csv = Path(state.get("source_csv", "")).expanduser()
        if not source_csv.exists():
            raise FileNotFoundError(f"The saved order CSV no longer exists: {source_csv}")
        previous_workbook_text = state.get("previous_workbook", "")
        previous_workbook = Path(previous_workbook_text).expanduser() if previous_workbook_text else None

        _save_and_close_active_excel_workbook()
        result = generate_review_workbook(
            source_csv, live_product_master_path(), REPORTS,
            report_mode=state.get("report_mode") or GENERAL_SALES_PERIOD,
            event_name=state.get("event_name") or "",
            decoration_fulfillment=state.get("decoration_fulfillment") or STANDARD_ORCHID_WORKFLOW,
            regenerate_helper=REGENERATE_HELPER,
            previous_review_path=previous_workbook,
        )
        subprocess.run(["open", str(result["output_path"])], check=False)
    except Exception as error:
        _show_regeneration_error(str(error))
        raise SystemExit(1) from error


if __name__ == "__main__":
    regeneration_requests = [Path(arg) for arg in sys.argv[1:] if str(arg).casefold().endswith(".orchidregen")]
    if regeneration_requests:
        regenerate_review_from_saved_state(regeneration_requests[0])
    elif "--regenerate-review" in sys.argv:
        regenerate_review_from_saved_state()
    elif "--product-master" in sys.argv:
        from modules.product_master_editor import ProductMasterV2
        prefill = ""
        if "--new-product" in sys.argv:
            try:
                prefill = sys.argv[sys.argv.index("--new-product") + 1]
            except (ValueError, IndexError):
                prefill = ""
        search = ""
        if "--search-product" in sys.argv:
            try:
                search = sys.argv[sys.argv.index("--search-product") + 1]
            except (ValueError, IndexError):
                search = ""
        setup = ""
        if "--setup-product" in sys.argv:
            try:
                setup = sys.argv[sys.argv.index("--setup-product") + 1]
            except (ValueError, IndexError):
                setup = ""
        ProductMasterV2(prefill_product=prefill, search_product=search, setup_product=setup).mainloop()
    else:
        OrchidPurchaseManager().mainloop()
