from __future__ import annotations

import re
import shutil
from datetime import datetime
from pathlib import Path
import json
import os
import subprocess
import sys
from tkinter import messagebox, simpledialog
from PIL import Image as PILImage

import customtkinter as ctk
import pandas as pd

from modules.paths import product_master_path, product_master_update_marker_path
from modules.outsource_rules import (
    NEVER_OUTSOURCE_COLUMN, apply_never_outsource_defaults, default_never_outsource,
)
from modules.product_resolver import default_product_id
from modules.blank_garment_rules import (
    BLANK_DECORATION_LABEL,
    apply_blank_garment_defaults,
    is_blank_decoration,
    is_blank_garment_product,
    normalize_decoration_type,
)
from modules.scroll_support import install_native_scroll_support
from modules.decoration_locations import (
    DECORATION_LOCATIONS,
    STANDARD_DECORATION_LOCATIONS,
    LEFT_CHEST,
    OTHER_CUSTOM,
    NOT_APPLICABLE_IN_HOUSE,
    NOT_APPLICABLE_NO_DECORATION,
    default_decoration_location,
    location_choice_and_custom,
    normalize_decoration_location,
    resolve_location_choice,
)
from modules.product_intelligence import apply_product_intelligence, infer_default_decoration
from modules.routing_rules import preferred_vendor_for_brand_text
from modules.internal_services import (
    HEMMING_ALTERATION_LABEL, SEW_ON_PATCH_LABEL,
    infer_in_house_decoration, is_in_house_decoration, is_in_house_service_product,
)
from modules.purchase_rules import (
    CATEGORY_OPTIONS,
    RULE_COLUMNS,
    apply_purchase_rule_defaults,
    bool_text,
    category_defaults,
    infer_category,
    normalize_bool,
    row_rules,
)

PRODUCT_MASTER_PATH = product_master_path()
PRODUCT_MASTER_PID_FILE = PRODUCT_MASTER_PATH.parent / "product_master.pid"
PRODUCT_MASTER_REQUEST_FILE = PRODUCT_MASTER_PATH.parent / "product_master_open_request.json"
PRODUCT_MASTER_UPDATE_MARKER = product_master_update_marker_path()
COLUMNS = [
    "Product Name",
    "Style Number",
    "Garment Color",
    "Vendor",
    "Decoration Type",
    "Decoration Location",
    "Decoration Placement Instructions",
    "Decoration Color",
    "Product ID",
    "Product Aliases",
    "Vendor Color Code",
    "Color Aliases",
    "Product Category",
    "Requires Size",
    "Requires Color",
    "Requires Decoration",
    NEVER_OUTSOURCE_COLUMN,
    "Setup Required",
]

WORKBOOK_VENDORS = [
    "Burnside",
    "Cutter & Buck",
    "Outdoor Cap",
    "Richardson",
    "S&S Activewear",
    "SanMar",
    "Tru-Spec",
    "VF",
    "Wrangler",
]
PREFERRED_VENDOR_CASE = {vendor.casefold(): vendor for vendor in WORKBOOK_VENDORS}
DECORATION_TYPES = ["", "Embroidery", "Screen Print", SEW_ON_PATCH_LABEL, HEMMING_ALTERATION_LABEL, BLANK_DECORATION_LABEL]
DECORATION_COLORS = ["", "White", "Black", "Navy", "Red", "Royal", "Gold", "Silver", "Gray"]

PURPLE = "#5B2AA8"
PURPLE_DARK = "#3D176F"
PURPLE_LIGHT = "#F2ECFB"
PURPLE_BORDER = "#D7C8EE"
TEXT_DARK = "#201A2D"
TEXT_MUTED = "#6F667A"
CARD_BG = "#FFFFFF"
WINDOW_BG = "#F7F5FA"
ROW_ALT = "#FAF7FE"
SUCCESS = "#2D7A46"
WARNING = "#B26A00"
DANGER = "#B42318"
MAX_DASHBOARD_ROWS = 60


def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_space(value):
    return re.sub(r"\s+", " ", clean_text(value)).strip()


def canonical_vendor_name(value):
    """Return one consistent spelling for a vendor, ignoring case and spacing."""
    value = normalize_space(value)
    if not value:
        return ""
    return PREFERRED_VENDOR_CASE.get(value.casefold(), value)


def append_vendor_option(options, value):
    value = canonical_vendor_name(value)
    if value and not any(existing.casefold() == value.casefold() for existing in options):
        options.append(value)
    return value


def canonicalize_vendor_values(values):
    """Normalize all vendor spellings case-insensitively, including custom vendors."""
    registry = dict(PREFERRED_VENDOR_CASE)
    result = []
    for value in values:
        value = canonical_vendor_name(value)
        if not value:
            result.append("")
            continue
        key = value.casefold()
        canonical = registry.get(key)
        if canonical is None:
            registry[key] = value
            canonical = value
        result.append(canonical)
    return result


def normalize_style(value):
    return re.sub(r"\s+", "", clean_text(value)).upper()


def normalize_color(value):
    return normalize_space(value).casefold()


def normalize_search_text(value):
    """Normalize punctuation and spacing so style/brand searches are forgiving."""
    value = normalize_space(value).casefold()
    value = value.replace("™", " ").replace("®", " ").replace("+", " and ").replace("&", " and ")
    return re.sub(r"[^a-z0-9]+", " ", value).strip()


def search_tokens(value):
    return [token for token in normalize_search_text(value).split() if token]


def search_match_score(rows, query):
    """Return a lower-is-better search score, or None when the row does not match."""
    tokens = search_tokens(query)
    if not tokens:
        return 50
    style = normalize_search_text(first_nonblank(rows.get("Style Number", [])))
    product_id = normalize_search_text(first_nonblank(rows.get("Product ID", [])))
    name = normalize_search_text(best_product_name(rows.get("Product Name", [])))
    vendor = normalize_search_text(first_nonblank(rows.get("Vendor", [])))
    combined = normalize_search_text(" ".join(clean_text(value) for value in rows[COLUMNS].to_numpy().flatten()))
    if not all(token in combined for token in tokens):
        return None
    normalized_query = " ".join(tokens)
    if normalized_query == style:
        return 0
    if normalized_query == product_id:
        return 1
    if style.startswith(normalized_query):
        return 2
    if product_id.startswith(normalized_query):
        return 3
    if name.startswith(normalized_query):
        return 4
    if normalized_query in name:
        return 5
    if normalized_query in vendor:
        return 6
    return 10


def canonical_product_name(value):
    name = normalize_space(value)
    return re.sub(r"[\s\.\-–—:|/]+$", "", name).strip()


def first_nonblank(values):
    for value in values:
        value = clean_text(value)
        if value:
            return value
    return ""


def best_product_name(values):
    cleaned = [canonical_product_name(value) for value in values]
    cleaned = [value for value in cleaned if value]
    if not cleaned:
        return ""
    counts = {}
    for value in cleaned:
        counts[value] = counts.get(value, 0) + 1
    return sorted(counts, key=lambda value: (-counts[value], len(value), value.casefold()))[0]


def catalog_key(row):
    style = normalize_style(row["Style Number"])
    color = normalize_color(row["Garment Color"])
    if style:
        return f"style:{style}|color:{color}"
    return f"product:{canonical_product_name(row['Product Name']).casefold()}|color:{color}"


def style_group_key(row):
    style = normalize_style(row["Style Number"])
    if style:
        return f"style:{style}"
    return f"product:{canonical_product_name(row['Product Name']).casefold()}"


def clean_and_deduplicate_master(master):
    master = master.copy()
    for column in COLUMNS:
        if column not in master.columns:
            master[column] = ""
    master = master[COLUMNS].fillna("")
    for column in COLUMNS:
        master[column] = master[column].map(clean_text)
    master["Product Name"] = master["Product Name"].map(canonical_product_name)
    master["Style Number"] = master["Style Number"].map(normalize_style)
    master["Garment Color"] = master["Garment Color"].map(normalize_space)
    # Repair duplicate spellings such as SanMar / Sanmar before saving or routing.
    master["Vendor"] = canonicalize_vendor_values(master["Vendor"])
    if master.empty:
        return pd.DataFrame(columns=COLUMNS)

    master["_key"] = master.apply(catalog_key, axis=1)
    rows = []
    for _, group in master.groupby("_key", sort=False, dropna=False):
        rows.append(
            {
                "Product Name": best_product_name(group["Product Name"]),
                "Style Number": first_nonblank(group["Style Number"]),
                "Garment Color": first_nonblank(group["Garment Color"]),
                "Vendor": first_nonblank(group["Vendor"]),
                "Decoration Type": first_nonblank(group["Decoration Type"]),
                "Decoration Location": first_nonblank(group["Decoration Location"]),
                "Decoration Placement Instructions": first_nonblank(group["Decoration Placement Instructions"]),
                "Decoration Color": first_nonblank(group["Decoration Color"]),
                "Product ID": first_nonblank(group["Product ID"]),
                "Product Aliases": first_nonblank(group["Product Aliases"]),
                "Vendor Color Code": first_nonblank(group["Vendor Color Code"]),
                "Color Aliases": first_nonblank(group["Color Aliases"]),
                "Product Category": first_nonblank(group["Product Category"]),
                "Requires Size": first_nonblank(group["Requires Size"]),
                "Requires Color": first_nonblank(group["Requires Color"]),
                "Requires Decoration": first_nonblank(group["Requires Decoration"]),
                NEVER_OUTSOURCE_COLUMN: first_nonblank(group[NEVER_OUTSOURCE_COLUMN]),
            }
        )
    result = pd.DataFrame(rows, columns=COLUMNS)
    if not result.empty:
        missing_id = result["Product ID"].astype(str).str.strip().eq("")
        result.loc[missing_id, "Product ID"] = result.loc[missing_id].apply(
            lambda row: default_product_id(row.get("Style Number", ""), row.get("Product Name", "")),
            axis=1,
        )
    result = apply_never_outsource_defaults(apply_product_intelligence(apply_purchase_rule_defaults(apply_blank_garment_defaults(result))))
    return result.sort_values(
        by=["Style Number", "Product Name", "Garment Color"],
        key=lambda series: series.astype(str).str.casefold(),
    ).reset_index(drop=True)


def load_product_master():
    PRODUCT_MASTER_PATH.parent.mkdir(parents=True, exist_ok=True)
    if not PRODUCT_MASTER_PATH.exists():
        pd.DataFrame(columns=COLUMNS).to_csv(PRODUCT_MASTER_PATH, index=False)
    try:
        master = pd.read_csv(PRODUCT_MASTER_PATH, dtype=str).fillna("")
    except pd.errors.EmptyDataError:
        master = pd.DataFrame(columns=COLUMNS)
    except Exception as error:
        messagebox.showerror("Unable to Open Product Master", str(error))
        return pd.DataFrame(columns=COLUMNS)
    cleaned = clean_and_deduplicate_master(master)
    comparable_original = master.copy()
    for column in COLUMNS:
        if column not in comparable_original.columns:
            comparable_original[column] = ""
    comparable_original = comparable_original[COLUMNS].fillna("")
    for column in COLUMNS:
        comparable_original[column] = comparable_original[column].map(clean_text)
    changed = not comparable_original.reset_index(drop=True).equals(cleaned.reset_index(drop=True))
    if changed and PRODUCT_MASTER_PATH.exists():
        backup_dir = PRODUCT_MASTER_PATH.parent / "backups"
        backup_dir.mkdir(parents=True, exist_ok=True)
        migration_backup = backup_dir / f"product_master_before_v50_purchase_rules_{datetime.now():%Y%m%d_%H%M%S}.csv"
        shutil.copy2(PRODUCT_MASTER_PATH, migration_backup)
    cleaned.to_csv(PRODUCT_MASTER_PATH, index=False)
    return cleaned


def save_product_master(master):
    try:
        cleaned = clean_and_deduplicate_master(master)
        cleaned.to_csv(PRODUCT_MASTER_PATH, index=False)
        return cleaned
    except Exception as error:
        messagebox.showerror("Unable to Save Product Master", str(error))
        return None


class ProductMasterV2(ctk.CTk):
    """Product Master editor. The class name remains compatible with older app versions."""

    def __init__(self, prefill_product: str = "", search_product: str = "", setup_product: str = ""):
        super().__init__()
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")
        self.title("Orchid Purchase Manager - Product Master 4.8.19")
        self.geometry("1320x930")
        self.minsize(1120, 780)
        self.configure(fg_color=WINDOW_BG)
        self.after(80, self._maximize_window)
        self._last_open_request = 0.0
        try:
            PRODUCT_MASTER_PID_FILE.parent.mkdir(parents=True, exist_ok=True)
            PRODUCT_MASTER_PID_FILE.write_text(str(os.getpid()), encoding="utf-8")
        except Exception:
            pass

        self.master_data = load_product_master()
        self.session_start_complete = 0
        self.session_start_total = 0
        self.session_saved_styles = set()
        self.session_backup_path = None
        self._save_toast_after_id = None

        self.vendor_options = list(WORKBOOK_VENDORS)
        for vendor in self.master_data.get("Vendor", pd.Series(dtype=str)):
            append_vendor_option(self.vendor_options, vendor)

        self.all_style_keys = []
        self.filtered_style_keys = []
        self.style_position = 0
        self.color_controls = []
        self.current_style_key = None

        self.search_var = ctk.StringVar()
        self.catalog_filter_var = ctk.StringVar(value="All Products")
        self.incomplete_only_var = ctk.BooleanVar(value=False)  # backward-compatible internal flag
        self.dashboard_search_var = ctk.StringVar()
        self.dashboard_vendor_filter_var = ctk.StringVar(value="All Vendors")
        self.vendor_var = ctk.StringVar()
        self.decoration_type_var = ctk.StringVar()
        self.decoration_location_var = ctk.StringVar(value=LEFT_CHEST)
        self.custom_decoration_location_var = ctk.StringVar()
        self.decoration_placement_instructions_var = ctk.StringVar()
        self.product_name_var = ctk.StringVar()
        self.style_number_var = ctk.StringVar()
        self.product_id_var = ctk.StringVar()
        self.product_aliases_var = ctk.StringVar()
        self.category_var = ctk.StringVar()
        self.requires_size_var = ctk.BooleanVar(value=True)
        self.requires_color_var = ctk.BooleanVar(value=True)
        self.requires_decoration_var = ctk.BooleanVar(value=True)
        self.never_outsource_var = ctk.BooleanVar(value=False)
        self.advanced_visible = False
        self.apply_color_var = ctk.StringVar(value="")

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.build_style_list()
        self.session_start_total = len(self.all_style_keys)
        self.session_start_complete = self.completed_style_count()

        self.dashboard_frame = ctk.CTkFrame(self, fg_color=WINDOW_BG, corner_radius=0)
        self.editor_frame = ctk.CTkFrame(self, fg_color=WINDOW_BG, corner_radius=0)
        self.dashboard_frame.grid(row=0, column=0, sticky="nsew")
        self.editor_frame.grid(row=0, column=0, sticky="nsew")
        self.build_dashboard()
        self.build_editor()
        install_native_scroll_support(
            self,
            [self.queue_scroll, self.vendor_stats_frame, self.left_scroll, self.colors_scroll],
        )
        self.product_name_var.trace_add("write", self.on_product_name_changed)
        self.decoration_type_var.trace_add("write", self.on_decoration_type_changed)
        self.show_dashboard()
        if clean_text(prefill_product):
            self.after(150, lambda value=clean_text(prefill_product): self.add_new_product(value, prompt=False))
        elif clean_text(setup_product):
            self.after(150, lambda value=clean_text(setup_product): self._open_setup_product(value))
        elif clean_text(search_product):
            self.after(150, lambda value=clean_text(search_product): self._open_search_result(value))
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.after(40, self._bring_to_front)
        self.after(300, self._poll_open_requests)

    def _maximize_window(self):
        """Open at the largest normal window size without entering macOS full-screen mode."""
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

    def _bring_to_front(self):
        try:
            self.deiconify()
            self.lift()
            self.attributes("-topmost", True)
            self.focus_force()
            self.after(700, lambda: self.attributes("-topmost", False))
        except Exception:
            pass
        if sys.platform == "darwin":
            script = (
                'tell application "System Events" to set frontmost of '
                f'(first process whose unix id is {os.getpid()}) to true'
            )
            subprocess.run(
                ["/usr/bin/osascript", "-e", script], check=False,
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            )

    def _poll_open_requests(self):
        try:
            if PRODUCT_MASTER_REQUEST_FILE.exists():
                modified = PRODUCT_MASTER_REQUEST_FILE.stat().st_mtime
                if modified > self._last_open_request:
                    self._last_open_request = modified
                    payload = json.loads(PRODUCT_MASTER_REQUEST_FILE.read_text(encoding="utf-8"))
                    PRODUCT_MASTER_REQUEST_FILE.unlink(missing_ok=True)
                    self._bring_to_front()
                    action = str(payload.get("action", "show"))
                    value = clean_text(payload.get("value", ""))
                    if action == "new" and value:
                        self.add_new_product(value, prompt=False)
                    elif action == "setup" and value:
                        self._open_setup_product(value)
                    elif action == "search" and value:
                        self._open_search_result(value)
        except Exception:
            pass
        try:
            self.after(300, self._poll_open_requests)
        except Exception:
            pass

    # ---------- data and health ----------
    def build_style_list(self):
        """Build an index once so large Product Masters do not require repeated full-file scans."""
        self.all_style_keys = []
        self.style_rows_cache = {}
        for index, row in self.master_data.iterrows():
            key = style_group_key(row)
            if key not in self.style_rows_cache:
                self.style_rows_cache[key] = []
                self.all_style_keys.append(key)
            self.style_rows_cache[key].append(index)

    def get_style_rows(self, style_key):
        if self.master_data.empty:
            return self.master_data.copy()
        indices = self.style_rows_cache.get(style_key, [])
        if not indices:
            return self.master_data.iloc[0:0].copy()
        return self.master_data.loc[indices].copy()

    def style_issues(self, rows):
        issues = []
        if rows.empty:
            return ["No product record"]
        if "Setup Required" in rows.columns and rows["Setup Required"].astype(str).str.strip().str.casefold().isin({"yes", "y", "true", "1"}).any():
            issues.append("New product setup must be reviewed and saved")
        if not first_nonblank(rows["Product Name"]):
            issues.append("Missing product name")
        if not first_nonblank(rows["Style Number"]):
            issues.append("Missing product number")
        if not first_nonblank(rows["Product Category"]):
            issues.append("Missing product category")
        rules = row_rules(rows.iloc[0])
        requires_color = normalize_bool(rules["Requires Color"], True)
        requires_decoration = normalize_bool(rules["Requires Decoration"], True)
        if requires_color and rows["Garment Color"].astype(str).str.strip().eq("").any():
            issues.append("Missing purchasing color")
        decoration_type = first_nonblank(rows["Decoration Type"])
        service_only = is_in_house_service_product(
            first_nonblank(rows["Product Name"]),
            decoration_type=decoration_type,
        )
        if not service_only and rows["Vendor"].astype(str).str.strip().eq("").any():
            issues.append("Missing purchase vendor")
        if rows["Decoration Type"].astype(str).str.strip().eq("").any():
            issues.append("Missing decoration type")
        decorated = rows[rows["Decoration Type"].map(lambda value: not is_blank_decoration(value) and not is_in_house_decoration(value))]
        if requires_decoration and not decorated.empty:
            locations = decorated["Decoration Location"].astype(str).str.strip()
            if locations.eq("").any() or locations.eq(OTHER_CUSTOM).any():
                issues.append("Missing decoration location")
            if decorated["Decoration Color"].astype(str).str.strip().eq("").any():
                issues.append("Missing thread/ink color")
        return issues

    def style_is_complete(self, rows):
        return not self.style_issues(rows)

    def completed_style_count(self):
        return sum(1 for key in self.all_style_keys if self.style_is_complete(self.get_style_rows(key)))

    def incomplete_style_keys(self):
        return [key for key in self.all_style_keys if not self.style_is_complete(self.get_style_rows(key))]

    def vendor_style_counts(self):
        counts = {}
        for key in self.all_style_keys:
            vendor = first_nonblank(self.get_style_rows(key)["Vendor"]) or "Unassigned"
            counts[vendor] = counts.get(vendor, 0) + 1
        return sorted(counts.items(), key=lambda item: (-item[1], item[0].casefold()))

    def style_display(self, key):
        rows = self.get_style_rows(key)
        style = first_nonblank(rows["Style Number"])
        name = best_product_name(rows["Product Name"])
        if style and name:
            return f"{style} — {name}"
        return style or name or "Unnamed product"

    # ---------- dashboard ----------
    def _open_search_result(self, value: str):
        value = clean_text(value)
        if not value:
            return
        self.dashboard_search_var.set(value)
        self.refresh_dashboard()
        matches = self.dashboard_matching_style_keys()
        if matches:
            self.open_style(matches[0])

    def _open_setup_product(self, value: str):
        """Open the requested candidate inside the Needs Setup queue."""
        value = clean_text(value)
        if not value:
            return
        self.reload_live_file(silent=True)
        self.catalog_filter_var.set("Needs Setup")
        self.incomplete_only_var.set(True)
        self.search_var.set("")
        self.dashboard_search_var.set(value)
        matches = self.dashboard_matching_style_keys()
        if not matches:
            self.add_new_product(value, prompt=False)
            return
        target = matches[0]
        self.show_editor(target_key=target)
        self.set_catalog_filter("Needs Setup")
        if target in self.filtered_style_keys:
            self.style_position = self.filtered_style_keys.index(target)
            self.show_current_style()

    def build_dashboard(self):
        frame = self.dashboard_frame
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(3, weight=1)

        header = ctk.CTkFrame(frame, fg_color=CARD_BG, corner_radius=0, height=104)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_columnconfigure(1, weight=1)
        flower_path = Path(__file__).resolve().parent.parent / "assets" / "orchid_flower_purple.png"
        if flower_path.exists():
            flower_source = PILImage.open(flower_path).convert("RGBA")
            self.product_master_flower_image = ctk.CTkImage(light_image=flower_source, dark_image=flower_source, size=(62, 62))
            ctk.CTkLabel(header, text="", image=self.product_master_flower_image, width=70).grid(row=0, column=0, rowspan=2, padx=(28, 8), pady=18)
        else:
            # The branded flower asset ships with every build. Avoid substituting
            # a generic glyph if an installation is incomplete.
            ctk.CTkLabel(header, text="", width=70).grid(
                row=0, column=0, rowspan=2, padx=(28, 8), pady=18
            )
        ctk.CTkLabel(header, text="Product Master", text_color=TEXT_DARK, font=ctk.CTkFont(size=31, weight="bold"), anchor="w").grid(row=0, column=1, sticky="sw", pady=(18, 0))
        ctk.CTkLabel(header, text="Define where products are purchased and what information is required to place an order.", text_color=TEXT_MUTED, font=ctk.CTkFont(size=14), anchor="w").grid(row=1, column=1, sticky="nw", pady=(2, 16))
        ctk.CTkButton(header, text="+ Add New Product", command=self.add_new_product, width=155, height=42, fg_color="transparent", hover_color=PURPLE_LIGHT, border_width=1, border_color=PURPLE, text_color=PURPLE, font=ctk.CTkFont(size=13, weight="bold")).grid(row=0, column=2, rowspan=2, padx=(8, 6), pady=28)
        ctk.CTkButton(header, text="Refresh", command=self.reload_live_file, width=145, height=42, fg_color="transparent", hover_color=PURPLE_LIGHT, border_width=1, border_color=PURPLE, text_color=PURPLE, font=ctk.CTkFont(size=13, weight="bold")).grid(row=0, column=3, rowspan=2, padx=6, pady=28)
        ctk.CTkButton(header, text="Browse Catalog  ▶", command=self.browse_all, width=205, height=44, fg_color=PURPLE, hover_color=PURPLE_DARK, font=ctk.CTkFont(size=14, weight="bold")).grid(row=0, column=4, rowspan=2, padx=(6, 22), pady=28)

        self.dashboard_stats = ctk.CTkFrame(frame, fg_color="transparent")
        self.dashboard_stats.grid(row=1, column=0, sticky="ew", padx=20, pady=(16, 8))
        for column in range(4):
            self.dashboard_stats.grid_columnconfigure(column, weight=1)
        self.total_card = self.create_stat_card(self.dashboard_stats, 0, "Total Styles", "0")
        self.complete_card = self.create_stat_card(self.dashboard_stats, 1, "Fully Configured", "0")
        self.incomplete_card = self.create_stat_card(self.dashboard_stats, 2, "Vendors", "0")
        self.percent_card = self.create_stat_card(self.dashboard_stats, 3, "Catalog Status", "Ready")

        health = ctk.CTkFrame(frame, fg_color=CARD_BG, border_width=1, border_color=PURPLE_BORDER, corner_radius=12)
        health.grid(row=2, column=0, sticky="ew", padx=20, pady=8)
        health.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(health, text="Product Master Progress", text_color=TEXT_DARK, font=ctk.CTkFont(size=15, weight="bold")).grid(row=0, column=0, padx=(20, 14), pady=18)
        self.dashboard_progress = ctk.CTkProgressBar(health, progress_color=PURPLE, fg_color="#EDE8F3")
        self.dashboard_progress.grid(row=0, column=1, sticky="ew", padx=10, pady=18)
        self.dashboard_health_text = ctk.CTkLabel(health, text="", text_color=TEXT_MUTED, font=ctk.CTkFont(size=13))
        self.dashboard_health_text.grid(row=0, column=2, padx=(14, 20), pady=18)

        body = ctk.CTkFrame(frame, fg_color="transparent")
        body.grid(row=3, column=0, sticky="nsew", padx=20, pady=(8, 16))
        body.grid_columnconfigure(0, weight=2)
        body.grid_columnconfigure(1, weight=1)
        body.grid_rowconfigure(0, weight=1)

        queue_card = ctk.CTkFrame(body, fg_color=CARD_BG, border_width=1, border_color=PURPLE_BORDER, corner_radius=12)
        queue_card.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        queue_card.grid_columnconfigure(0, weight=1)
        queue_card.grid_rowconfigure(3, weight=1)
        ctk.CTkLabel(queue_card, text="Product Catalog", text_color=PURPLE, font=ctk.CTkFont(size=18, weight="bold"), anchor="w").grid(row=0, column=0, sticky="ew", padx=20, pady=(18, 4))
        ctk.CTkLabel(queue_card, text="Search or browse saved product master details. Only gaps affecting the current import appear in Purchase Review.", text_color=TEXT_MUTED, font=ctk.CTkFont(size=12), anchor="w").grid(row=1, column=0, sticky="ew", padx=20, pady=(0, 8))
        search_row = ctk.CTkFrame(queue_card, fg_color="transparent")
        search_row.grid(row=2, column=0, sticky="ew", padx=20, pady=(0, 10))
        search_row.grid_columnconfigure(0, weight=1)
        self.dashboard_search_entry = ctk.CTkEntry(search_row, textvariable=self.dashboard_search_var, placeholder_text="Search product #, ID, alias, description, vendor, or color...", height=38, border_color=PURPLE_BORDER, fg_color="#FFFFFF", text_color=TEXT_DARK)
        self.dashboard_search_entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        self.dashboard_search_entry.bind("<Return>", lambda _event: self.open_first_dashboard_match())
        dashboard_vendor_values = ["All Vendors"] + self.vendor_options + ["Unassigned"]
        self.dashboard_vendor_filter = ctk.CTkComboBox(search_row, variable=self.dashboard_vendor_filter_var, values=list(dict.fromkeys(dashboard_vendor_values)), width=170, height=38, command=lambda _value: self.refresh_dashboard(), border_color=PURPLE_BORDER, button_color="#EAE3F5", button_hover_color="#DED2EF", fg_color="#FFFFFF", text_color=TEXT_DARK)
        self.dashboard_vendor_filter.grid(row=0, column=1, padx=(8, 0))
        self.dashboard_search_results = ctk.CTkLabel(
            search_row, text="", text_color=TEXT_MUTED, font=ctk.CTkFont(size=11), anchor="w"
        )
        self.dashboard_search_results.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(5, 0))
        self.dashboard_search_var.trace_add("write", lambda *_: self.refresh_dashboard())
        self.queue_scroll = ctk.CTkScrollableFrame(queue_card, fg_color="#FFFFFF", corner_radius=8, border_width=1, border_color=PURPLE_BORDER)
        self.queue_scroll.grid(row=3, column=0, sticky="nsew", padx=20, pady=(0, 18))
        self.queue_scroll.grid_columnconfigure(0, weight=1)

        vendor_card = ctk.CTkFrame(body, fg_color=CARD_BG, border_width=1, border_color=PURPLE_BORDER, corner_radius=12)
        vendor_card.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        vendor_card.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(vendor_card, text="Styles by Vendor", text_color=PURPLE, font=ctk.CTkFont(size=18, weight="bold"), anchor="w").grid(row=0, column=0, sticky="ew", padx=20, pady=(18, 4))
        self.vendor_stats_frame = ctk.CTkScrollableFrame(vendor_card, fg_color="#FFFFFF", corner_radius=8, border_width=1, border_color=PURPLE_BORDER)
        self.vendor_stats_frame.grid(row=1, column=0, sticky="nsew", padx=20, pady=(10, 12))
        self.vendor_stats_frame.grid_columnconfigure(0, weight=1)
        vendor_card.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(vendor_card, text="Click a vendor to filter the Product Catalog.", text_color=TEXT_MUTED, font=ctk.CTkFont(size=11), justify="left", wraplength=330).grid(row=2, column=0, sticky="ew", padx=20, pady=(0, 18))

    def create_stat_card(self, parent, column, title, value):
        card = ctk.CTkFrame(parent, fg_color=CARD_BG, border_width=1, border_color=PURPLE_BORDER, corner_radius=12)
        card.grid(row=0, column=column, sticky="ew", padx=(0 if column == 0 else 7, 0 if column == 3 else 7))
        ctk.CTkLabel(card, text=title, text_color=TEXT_MUTED, font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w", padx=18, pady=(15, 3))
        label = ctk.CTkLabel(card, text=value, text_color=TEXT_DARK, font=ctk.CTkFont(size=27, weight="bold"))
        label.pack(anchor="w", padx=18, pady=(0, 15))
        return label

    def refresh_dashboard(self):
        self.build_style_list()
        total = len(self.all_style_keys)
        complete = self.completed_style_count()
        incomplete = total - complete
        percent = complete / total if total else 0
        assigned_vendors = {first_nonblank(self.get_style_rows(key)["Vendor"]) for key in self.all_style_keys}
        assigned_vendors.discard("")
        self.total_card.configure(text=str(total))
        self.complete_card.configure(text=str(complete), text_color=SUCCESS)
        self.incomplete_card.configure(text=str(len(assigned_vendors)), text_color=PURPLE)
        self.percent_card.configure(
            text="Ready" if not incomplete else "Needs Setup",
            text_color=SUCCESS if not incomplete else WARNING,
        )
        self.dashboard_progress.set(percent)
        if incomplete:
            self.dashboard_health_text.configure(text=f"{complete} complete • {incomplete} need attention • {total} total styles • {round(percent * 100)}%")
        else:
            self.dashboard_health_text.configure(text=f"{total} complete • Product Master ready • 100%")
        if hasattr(self, "dashboard_vendor_filter"):
            values = ["All Vendors"] + self.vendor_options + ["Unassigned"]
            self.dashboard_vendor_filter.configure(values=list(dict.fromkeys(values)))
        if hasattr(self, "filter_buttons"):
            self.filter_buttons["All Products"].configure(text=f"All Products ({total})")
            self.filter_buttons["Needs Setup"].configure(text=f"Needs Setup ({incomplete})")
            self.filter_buttons["Complete"].configure(text=f"Complete ({complete})")
            self._refresh_filter_button_styles()

        for widget in self.queue_scroll.winfo_children():
            widget.destroy()
        display_keys = self.dashboard_matching_style_keys()
        searching = bool(self.dashboard_search_var.get().strip()) or self.dashboard_vendor_filter_var.get() != "All Vendors"
        if hasattr(self, "dashboard_search_results"):
            label = f"{len(display_keys)} result{'s' if len(display_keys) != 1 else ''}"
            if searching:
                label += " — press Return to open the best match"
            self.dashboard_search_results.configure(text=label)
        if not display_keys:
            search_value = self.dashboard_search_var.get().strip()
            if searching and search_value:
                ctk.CTkLabel(self.queue_scroll, text=f"{search_value} is not in Product Master", text_color=WARNING, font=ctk.CTkFont(size=18, weight="bold")).grid(row=0, column=0, padx=20, pady=(34, 8))
                ctk.CTkLabel(self.queue_scroll, text="Add it as a new permanent product master detail, then complete the vendor, color, size requirement, and decoration information.", text_color=TEXT_MUTED, font=ctk.CTkFont(size=12), wraplength=590).grid(row=1, column=0, padx=20, pady=(0, 14))
                ctk.CTkButton(self.queue_scroll, text=f"+ Add {search_value} as New Product", command=lambda value=search_value: self.add_new_product(value, prompt=False), width=235, height=38, fg_color=PURPLE, hover_color=PURPLE_DARK, font=ctk.CTkFont(size=13, weight="bold")).grid(row=2, column=0, padx=20, pady=(0, 34))
            else:
                message = "No products match this search" if searching else "No products are saved in the catalog yet"
                color = TEXT_MUTED
                ctk.CTkLabel(self.queue_scroll, text=message, text_color=color, font=ctk.CTkFont(size=18, weight="bold")).grid(row=0, column=0, padx=20, pady=40)
        else:
            visible_keys = display_keys[:MAX_DASHBOARD_ROWS]
            row_offset = 0
            if len(display_keys) > MAX_DASHBOARD_ROWS:
                ctk.CTkLabel(
                    self.queue_scroll,
                    text=f"Showing the first {MAX_DASHBOARD_ROWS} of {len(display_keys)} products. Use search or the vendor filter to find a specific style without loading every card at once.",
                    text_color=WARNING,
                    font=ctk.CTkFont(size=11, weight="bold"),
                    justify="left",
                    wraplength=590,
                ).grid(row=0, column=0, sticky="ew", padx=14, pady=(12, 8))
                row_offset = 1
            for row_index, key in enumerate(visible_keys, start=row_offset):
                rows = self.get_style_rows(key)
                item = ctk.CTkFrame(self.queue_scroll, fg_color=ROW_ALT if row_index % 2 else "#FFFFFF", corner_radius=8)
                item.grid(row=row_index, column=0, sticky="ew", padx=4, pady=3)
                item.grid_columnconfigure(0, weight=1)
                ctk.CTkLabel(item, text=self.style_display(key), text_color=TEXT_DARK, font=ctk.CTkFont(size=13, weight="bold"), anchor="w", wraplength=500).grid(row=0, column=0, sticky="ew", padx=14, pady=(10, 2))
                issues = self.style_issues(rows)
                has_vendor = bool(first_nonblank(rows["Vendor"]))
                has_decoration = bool(first_nonblank(rows["Decoration Type"]))
                if not issues:
                    badge_text, badge_text_color, badge_fill = "COMPLETE", SUCCESS, "#EAF8EF"
                    detail_text = "All required purchasing details are configured"
                    button_text = "Edit"
                elif has_vendor or has_decoration:
                    badge_text, badge_text_color, badge_fill = "PARTIAL", WARNING, "#FFF1DC"
                    detail_text = " • ".join(issues[:2]) + (f" • +{len(issues) - 2} more" if len(issues) > 2 else "")
                    button_text = "Continue"
                else:
                    badge_text, badge_text_color, badge_fill = "NEEDS SETUP", DANGER, "#FDECEC"
                    detail_text = " • ".join(issues[:2]) + (f" • +{len(issues) - 2} more" if len(issues) > 2 else "")
                    button_text = "Set Up"
                ctk.CTkLabel(item, text=detail_text, text_color=TEXT_MUTED, font=ctk.CTkFont(size=11), anchor="w", wraplength=500).grid(row=1, column=0, sticky="ew", padx=14, pady=(0, 10))
                ctk.CTkLabel(item, text=badge_text, height=25, corner_radius=12, fg_color=badge_fill, text_color=badge_text_color, font=ctk.CTkFont(size=10, weight="bold")).grid(row=0, column=1, rowspan=2, padx=(8, 4), pady=10)
                ctk.CTkButton(item, text=button_text, command=lambda target=key: self.open_style(target), width=82, height=32, fg_color="transparent", hover_color=PURPLE_LIGHT, border_width=1, border_color=PURPLE, text_color=PURPLE).grid(row=0, column=2, rowspan=2, padx=(4, 12), pady=10)

        for widget in self.vendor_stats_frame.winfo_children():
            widget.destroy()
        vendor_rows = [("All Vendors", len(self.all_style_keys))] + self.vendor_style_counts()
        for index, (vendor, count) in enumerate(vendor_rows):
            selected = self.dashboard_vendor_filter_var.get().casefold() == vendor.casefold()
            row_fill = PURPLE_LIGHT if selected else (ROW_ALT if index % 2 else "#FFFFFF")
            row = ctk.CTkFrame(self.vendor_stats_frame, fg_color=row_fill, corner_radius=6)
            row.grid(row=index, column=0, sticky="ew", pady=1)
            row.grid_columnconfigure(0, weight=1)
            ctk.CTkButton(
                row,
                text=vendor,
                command=lambda value=vendor: self.filter_dashboard_by_vendor(value),
                fg_color="transparent",
                hover_color=PURPLE_LIGHT,
                text_color=PURPLE_DARK if selected else TEXT_DARK,
                anchor="w",
                height=32,
                font=ctk.CTkFont(size=12, weight="bold" if selected else "normal"),
            ).grid(row=0, column=0, sticky="ew", padx=(4, 0), pady=2)
            ctk.CTkButton(
                row,
                text=str(count),
                command=lambda value=vendor: self.filter_dashboard_by_vendor(value),
                width=48,
                height=28,
                fg_color="transparent",
                hover_color=PURPLE_LIGHT,
                text_color=PURPLE,
                font=ctk.CTkFont(size=12, weight="bold"),
            ).grid(row=0, column=1, padx=(0, 4), pady=2)

    def filter_dashboard_by_vendor(self, vendor: str):
        """Filter the Product Catalog when a vendor summary row is clicked."""
        self.dashboard_vendor_filter_var.set(vendor or "All Vendors")
        self.refresh_dashboard()

    def dashboard_matching_style_keys(self):
        query = self.dashboard_search_var.get().strip()
        vendor_filter = self.dashboard_vendor_filter_var.get().strip() or "All Vendors"
        ranked = []
        for position, key in enumerate(self.all_style_keys):
            rows = self.get_style_rows(key)
            vendor = first_nonblank(rows["Vendor"]) or "Unassigned"
            if vendor_filter != "All Vendors" and vendor.casefold() != vendor_filter.casefold():
                continue
            score = search_match_score(rows, query)
            if score is None:
                continue
            ranked.append((score, position, key))
        ranked.sort(key=lambda item: (item[0], item[1]))
        return [key for _, _, key in ranked]

    def open_first_dashboard_match(self):
        matches = self.dashboard_matching_style_keys()
        if matches:
            self.open_style(matches[0])

    def show_dashboard(self):
        """Return reliably from Style Editor to the Product Catalog dashboard."""
        self.editor_frame.grid_remove()
        self.dashboard_frame.grid(sticky="nsew")
        self.dashboard_frame.tkraise()
        self.reload_live_file(silent=True)
        self.refresh_dashboard()
        self.update_idletasks()

    def reload_live_file(self, silent: bool = False):
        """Reload the active Product Master from disk so this window cannot show stale data."""
        self.master_data = load_product_master()
        self.vendor_options = list(WORKBOOK_VENDORS)
        for vendor in self.master_data.get("Vendor", pd.Series(dtype=str)):
            append_vendor_option(self.vendor_options, vendor)
        if hasattr(self, "dashboard_vendor_filter"):
            self.dashboard_vendor_filter.configure(values=list(dict.fromkeys(["All Vendors"] + self.vendor_options + ["Unassigned"])))
        if hasattr(self, "vendor_combo"):
            self.vendor_combo.configure(values=self.vendor_options)
        self.build_style_list()
        self.refresh_dashboard()
        if not silent:
            messagebox.showinfo("Product Master Reloaded", f"Reloaded the live Product Master catalog from:\n{PRODUCT_MASTER_PATH}")

    def add_vendor_option(self):
        """Add and select a vendor without leaving the Style Editor."""
        value = simpledialog.askstring(
            "Add Purchase Vendor",
            "Enter the vendor name exactly as it should appear on purchase orders.",
            parent=self,
        )
        value = canonical_vendor_name(value or "")
        if not value:
            return
        existing = next((vendor for vendor in self.vendor_options if vendor.casefold() == value.casefold()), "")
        selected = existing or value
        if not existing:
            append_vendor_option(self.vendor_options, value)
            if hasattr(self, "vendor_combo"):
                self.vendor_combo.configure(values=self.vendor_options)
            if hasattr(self, "dashboard_vendor_filter"):
                self.dashboard_vendor_filter.configure(
                    values=list(dict.fromkeys(["All Vendors"] + self.vendor_options + ["Unassigned"]))
                )
        self.vendor_var.set(selected)
        self.status_label.configure(text=f"Vendor {selected} added and selected. Save this product to keep it in future dropdowns.")

    def add_new_product(self, seed: str | None = None, prompt: bool = True):
        """Create a new Product Master record and open it for completion."""
        if prompt:
            seed = simpledialog.askstring(
                "Add New Product",
                "Enter the product/style number or a product description.\n\nExample: 1104",
                parent=self,
            )
        seed = normalize_space(seed or "")
        if not seed:
            return

        style_candidate = normalize_style(seed) if re.fullmatch(r"[A-Za-z0-9._/-]+", seed) else ""
        if style_candidate:
            existing_key = f"style:{style_candidate}"
            if existing_key in self.all_style_keys:
                self.open_style(existing_key)
                messagebox.showinfo("Product Already Exists", f"{style_candidate} is already in Product Master.")
                return
            product_name = ""
        else:
            product_name = canonical_product_name(seed)
            existing_key = f"product:{product_name.casefold()}"
            if existing_key in self.all_style_keys:
                self.open_style(existing_key)
                messagebox.showinfo("Product Already Exists", f"{product_name} is already in Product Master.")
                return

        category = infer_category(product_name, style_candidate)
        defaults = category_defaults(category, product_name, style_candidate)
        row = {
            "Product Name": product_name,
            "Style Number": style_candidate,
            "Garment Color": "",
            "Vendor": "",
            "Decoration Type": "",
            "Decoration Location": "",
            "Decoration Placement Instructions": "",
            "Decoration Color": "",
            "Product ID": default_product_id(style_candidate, product_name),
            "Product Aliases": "",
            "Vendor Color Code": "",
            "Color Aliases": "",
            "Product Category": category,
            "Requires Size": bool_text(defaults["requires_size"]),
            "Requires Color": bool_text(defaults["requires_color"]),
            "Requires Decoration": bool_text(defaults["requires_decoration"]),
            NEVER_OUTSOURCE_COLUMN: bool_text(default_never_outsource(product_name, category, style_candidate)),
            "Setup Required": "Yes",
        }
        self.master_data = pd.concat([self.master_data, pd.DataFrame([row], columns=COLUMNS)], ignore_index=True)
        self.build_style_list()
        new_key = style_group_key(pd.Series(row))
        self.dashboard_search_var.set("")
        self.search_var.set("")
        self.incomplete_only_var.set(False)
        self.show_editor(target_key=new_key)
        if style_candidate and not product_name:
            self.category_var.set("")
        self.status_label.configure(text="New product created. Complete the product name, category, vendor, color, and decoration rules, then click Save.")

    def start_incomplete_queue(self):
        incomplete = self.incomplete_style_keys()
        if not incomplete:
            messagebox.showinfo("Product Master", "Every saved style currently has all optional purchasing fields completed.")
            return
        self.search_var.set("")
        self.incomplete_only_var.set(True)
        self.show_editor(target_key=incomplete[0])

    def open_style(self, key):
        self.search_var.set("")
        self.incomplete_only_var.set(False)
        self.show_editor(target_key=key)

    # ---------- editor ----------
    def build_editor(self):
        frame = self.editor_frame
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(4, weight=1)
        self.build_editor_header(frame)
        self.build_filter_bar(frame)
        self.build_style_banner(frame)
        self.build_content_area(frame)
        self.build_action_bar(frame)
        self.build_footer(frame)
        self.search_var.trace_add("write", lambda *_: self.apply_filters(reset_position=True))

    def build_editor_header(self, parent):
        header = ctk.CTkFrame(parent, fg_color=CARD_BG, corner_radius=0, height=100)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_columnconfigure(1, weight=1)
        ctk.CTkButton(header, text="← Product Catalog", command=self.show_dashboard, width=165, height=38, fg_color="transparent", hover_color=PURPLE_LIGHT, border_width=1, border_color=PURPLE, text_color=PURPLE).grid(row=0, column=0, rowspan=2, padx=24, pady=28)
        ctk.CTkLabel(header, text="Product Master — Style Editor", text_color=TEXT_DARK, font=ctk.CTkFont(size=29, weight="bold"), anchor="w").grid(row=0, column=1, sticky="sw", pady=(18, 0))
        ctk.CTkLabel(header, text="Permanent product information entered here is reused for future Shopify imports.", text_color=TEXT_MUTED, font=ctk.CTkFont(size=13), anchor="w").grid(row=1, column=1, sticky="nw", pady=(2, 16))
        progress_card = ctk.CTkFrame(header, fg_color=CARD_BG, border_width=1, border_color=PURPLE_BORDER, corner_radius=12, width=340, height=64)
        progress_card.grid(row=0, column=2, rowspan=2, padx=24, pady=18)
        self.progress_text = ctk.CTkLabel(progress_card, text="", text_color=TEXT_DARK, font=ctk.CTkFont(size=13, weight="bold"))
        self.progress_text.pack(padx=18, pady=(12, 4))
        self.progress_bar = ctk.CTkProgressBar(progress_card, width=260, progress_color=PURPLE, fg_color="#EDE8F3")
        self.progress_bar.pack(padx=18, pady=(0, 12))

    def set_catalog_filter(self, value: str, open_first: bool = False):
        self.catalog_filter_var.set(value)
        self.incomplete_only_var.set(value == "Needs Setup")
        self._refresh_filter_button_styles()
        self.apply_filters(reset_position=True)
        if open_first and self.filtered_style_keys:
            self.style_position = 0
            self.show_current_style()

    def _refresh_filter_button_styles(self):
        selected = self.catalog_filter_var.get()
        for value, button in getattr(self, "filter_buttons", {}).items():
            active = value == selected
            button.configure(
                fg_color=PURPLE if active else PURPLE_LIGHT,
                hover_color=PURPLE_DARK if active else "#E4D6F5",
                text_color="#FFFFFF" if active else TEXT_DARK,
            )

    def continue_setup(self):
        incomplete = self.incomplete_style_keys()
        if not incomplete:
            messagebox.showinfo("Product Master Complete", "Every saved style is fully configured.")
            return
        self.set_catalog_filter("Needs Setup")
        self.show_editor(target_key=incomplete[0])

    def build_filter_bar(self, parent):
        bar = ctk.CTkFrame(parent, fg_color=CARD_BG, border_width=1, border_color=PURPLE_BORDER, corner_radius=12)
        bar.grid(row=1, column=0, sticky="ew", padx=20, pady=(14, 8))
        bar.grid_columnconfigure(0, weight=1)
        self.search_entry = ctk.CTkEntry(bar, textvariable=self.search_var, placeholder_text="Search by product number, ID, alias, description, vendor, or color...", height=42, border_color=PURPLE_BORDER, fg_color="#FFFFFF", text_color=TEXT_DARK, placeholder_text_color="#9A92A4")
        self.search_entry.grid(row=0, column=0, sticky="ew", padx=(18, 12), pady=14)
        self.search_entry.bind("<Return>", lambda _event: self.open_first_editor_match())

        filter_frame = ctk.CTkFrame(bar, fg_color="transparent")
        filter_frame.grid(row=0, column=1, padx=8, pady=14)
        self.filter_buttons = {}
        total = len(self.all_style_keys)
        incomplete = len(self.incomplete_style_keys())
        complete = total - incomplete
        labels = {"All Products": f"All Products ({total})", "Needs Setup": f"Needs Setup ({incomplete})", "Complete": f"Complete ({complete})"}
        for col, value in enumerate(("All Products", "Needs Setup", "Complete")):
            button = ctk.CTkButton(
                filter_frame, text=labels[value], width=118 if value != "Needs Setup" else 132, height=40,
                command=lambda v=value: self.set_catalog_filter(v, open_first=(v == "Needs Setup")),
                corner_radius=7, border_width=1, border_color=PURPLE_BORDER,
                font=ctk.CTkFont(size=12, weight="bold"),
            )
            button.grid(row=0, column=col, padx=2)
            self.filter_buttons[value] = button
        self._refresh_filter_button_styles()

        ctk.CTkButton(bar, text="Continue Setup ▶", command=self.continue_setup, width=145, height=42, fg_color=PURPLE, hover_color=PURPLE_DARK, text_color="#FFFFFF", font=ctk.CTkFont(size=12, weight="bold")).grid(row=0, column=2, padx=(8, 4), pady=14)
        ctk.CTkButton(bar, text="Clear Filters", command=self.browse_all, width=110, height=42, fg_color="transparent", hover_color=PURPLE_LIGHT, border_width=1, border_color=PURPLE, text_color=PURPLE).grid(row=0, column=3, padx=(4, 18), pady=14)

    def build_style_banner(self, parent):
        banner = ctk.CTkFrame(parent, fg_color=CARD_BG, border_width=1, border_color=PURPLE_BORDER, corner_radius=12)
        banner.grid(row=2, column=0, sticky="ew", padx=20, pady=8)
        banner.grid_columnconfigure(1, weight=1)
        ctk.CTkButton(banner, text="◀ Previous", command=self.previous_style, width=125, height=38, fg_color="transparent", hover_color=PURPLE_LIGHT, border_width=1, border_color=PURPLE, text_color=PURPLE).grid(row=0, column=0, rowspan=3, padx=20, pady=16)
        self.style_count_label = ctk.CTkLabel(banner, text="", text_color=TEXT_MUTED, font=ctk.CTkFont(size=12, weight="bold"))
        self.style_count_label.grid(row=0, column=1, pady=(12, 0))
        self.product_name_label = ctk.CTkLabel(banner, text="", text_color=TEXT_DARK, font=ctk.CTkFont(size=22, weight="bold"), wraplength=760)
        self.product_name_label.grid(row=1, column=1, pady=(2, 0))
        self.issue_label = ctk.CTkLabel(banner, text="", text_color=DANGER, font=ctk.CTkFont(size=12), wraplength=760)
        self.issue_label.grid(row=2, column=1, pady=(2, 12))
        ctk.CTkButton(banner, text="Next ▶", command=self.next_style, width=125, height=38, fg_color="transparent", hover_color=PURPLE_LIGHT, border_width=1, border_color=PURPLE, text_color=PURPLE).grid(row=0, column=2, rowspan=3, padx=20, pady=16)

    def build_content_area(self, parent):
        content = ctk.CTkFrame(parent, fg_color="transparent")
        content.grid(row=4, column=0, sticky="nsew", padx=20, pady=8)
        content.grid_columnconfigure(0, weight=0)
        content.grid_columnconfigure(1, weight=1)
        content.grid_rowconfigure(0, weight=1)
        self.build_left_panel(content)
        self.build_color_panel(content)

    def build_left_panel(self, parent):
        left = ctk.CTkScrollableFrame(parent, fg_color="transparent", width=390, corner_radius=0)
        left.grid(row=0, column=0, sticky="nsw", padx=(0, 10))
        left.grid_columnconfigure(0, weight=1)
        self.left_scroll = left

        info = ctk.CTkFrame(left, fg_color=CARD_BG, border_width=1, border_color=PURPLE_BORDER, corner_radius=12, width=370)
        info.grid(row=0, column=0, sticky="ew")
        info.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(info, text="① Purchasing Information", text_color=PURPLE, font=ctk.CTkFont(size=17, weight="bold"), anchor="w").grid(row=0, column=0, sticky="ew", padx=20, pady=(16, 10))

        self.add_labeled_entry(info, 1, "Product Number *", self.style_number_var, "Example: K500, 1062, ST660")
        self.add_labeled_entry(info, 4, "Product Description *", self.product_name_var, "Permanent catalog description")

        ctk.CTkLabel(info, text="Purchase Vendor *", text_color=TEXT_DARK, font=ctk.CTkFont(size=12, weight="bold"), anchor="w").grid(row=7, column=0, sticky="ew", padx=20, pady=(4, 4))
        vendor_row = ctk.CTkFrame(info, fg_color="transparent")
        vendor_row.grid(row=8, column=0, sticky="ew", padx=20)
        vendor_row.grid_columnconfigure(0, weight=1)
        self.vendor_combo = ctk.CTkComboBox(vendor_row, variable=self.vendor_var, values=self.vendor_options, height=38, border_color=PURPLE_BORDER, button_color="#EAE3F5", button_hover_color="#DED2EF", fg_color="#FFFFFF", text_color=TEXT_DARK, dropdown_fg_color="#FFFFFF", dropdown_text_color=TEXT_DARK)
        self.vendor_combo.grid(row=0, column=0, sticky="ew", padx=(0, 7))
        ctk.CTkButton(
            vendor_row, text="+ Add Vendor", command=self.add_vendor_option,
            width=108, height=38, fg_color="transparent", hover_color=PURPLE_LIGHT,
            border_width=1, border_color=PURPLE, text_color=PURPLE,
            font=ctk.CTkFont(size=11, weight="bold"),
        ).grid(row=0, column=1)
        ctk.CTkLabel(info, text="Choose a saved vendor, type a name, or use Add Vendor. Saving this product keeps the vendor in future dropdowns.", text_color=TEXT_MUTED, font=ctk.CTkFont(size=10), anchor="w", wraplength=330).grid(row=9, column=0, sticky="ew", padx=20, pady=(3, 8))

        ctk.CTkLabel(info, text="Product Category *", text_color=TEXT_DARK, font=ctk.CTkFont(size=12, weight="bold"), anchor="w").grid(row=10, column=0, sticky="ew", padx=20, pady=(4, 4))
        self.category_combo = ctk.CTkComboBox(info, variable=self.category_var, values=CATEGORY_OPTIONS, height=38, command=self.on_category_selected, border_color=PURPLE_BORDER, button_color="#EAE3F5", button_hover_color="#DED2EF", fg_color="#FFFFFF", text_color=TEXT_DARK, dropdown_fg_color="#FFFFFF", dropdown_text_color=TEXT_DARK)
        self.category_combo.grid(row=11, column=0, sticky="ew", padx=20)
        ctk.CTkLabel(info, text="Category supplies sensible defaults for size, color, and decoration requirements.", text_color=TEXT_MUTED, font=ctk.CTkFont(size=10), anchor="w", wraplength=320).grid(row=12, column=0, sticky="ew", padx=20, pady=(3, 8))

        ctk.CTkLabel(info, text="Decoration Type *", text_color=TEXT_DARK, font=ctk.CTkFont(size=12, weight="bold"), anchor="w").grid(row=13, column=0, sticky="ew", padx=20, pady=(4, 4))
        self.decoration_combo = ctk.CTkComboBox(info, variable=self.decoration_type_var, values=DECORATION_TYPES, height=38, border_color=PURPLE_BORDER, button_color="#EAE3F5", button_hover_color="#DED2EF", fg_color="#FFFFFF", text_color=TEXT_DARK, dropdown_fg_color="#FFFFFF", dropdown_text_color=TEXT_DARK)
        self.decoration_combo.grid(row=14, column=0, sticky="ew", padx=20, pady=(0, 8))
        ctk.CTkLabel(info, text="Default Decoration Location *", text_color=TEXT_DARK, font=ctk.CTkFont(size=12, weight="bold"), anchor="w").grid(row=15, column=0, sticky="ew", padx=20, pady=(4, 4))
        self.decoration_location_combo = ctk.CTkComboBox(
            info, variable=self.decoration_location_var, values=DECORATION_LOCATIONS, height=38,
            command=self.on_decoration_location_selected,
            border_color=PURPLE_BORDER, button_color="#EAE3F5", button_hover_color="#DED2EF",
            fg_color="#FFFFFF", text_color=TEXT_DARK, dropdown_fg_color="#FFFFFF", dropdown_text_color=TEXT_DARK,
        )
        self.decoration_location_combo.grid(row=16, column=0, sticky="ew", padx=20, pady=(0, 4))
        ctk.CTkLabel(
            info, text="Choose a standard route or Other / Custom. Every exact location is separated on the PO.",
            text_color=TEXT_MUTED, font=ctk.CTkFont(size=10), anchor="w", wraplength=330,
        ).grid(row=17, column=0, sticky="ew", padx=20, pady=(0, 8))
        self.custom_location_label = ctk.CTkLabel(
            info, text="Custom Location *", text_color=TEXT_DARK,
            font=ctk.CTkFont(size=12, weight="bold"), anchor="w",
        )
        self.custom_location_label.grid(row=18, column=0, sticky="ew", padx=20, pady=(4, 4))
        self.custom_location_entry = ctk.CTkEntry(
            info, textvariable=self.custom_decoration_location_var, height=38,
            placeholder_text="Example: Right Chest or Center Back",
        )
        self.custom_location_entry.grid(row=19, column=0, sticky="ew", padx=20, pady=(0, 8))
        ctk.CTkLabel(
            info, text="Placement Instructions (optional)", text_color=TEXT_DARK,
            font=ctk.CTkFont(size=12, weight="bold"), anchor="w",
        ).grid(row=20, column=0, sticky="ew", padx=20, pady=(4, 4))
        self.placement_instructions_entry = ctk.CTkEntry(
            info, textvariable=self.decoration_placement_instructions_var, height=38,
            placeholder_text="Example: 2 inches below shoulder seam",
        )
        self.placement_instructions_entry.grid(row=21, column=0, sticky="ew", padx=20, pady=(0, 14))
        self.on_decoration_location_selected(self.decoration_location_var.get())

        rules = ctk.CTkFrame(left, fg_color=CARD_BG, border_width=1, border_color=PURPLE_BORDER, corner_radius=12)
        rules.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        rules.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(rules, text="② Purchase Requirements", text_color=PURPLE, font=ctk.CTkFont(size=17, weight="bold"), anchor="w").grid(row=0, column=0, sticky="ew", padx=18, pady=(14, 6))
        ctk.CTkLabel(rules, text="Orchid only flags information that is required for this product.", text_color=TEXT_MUTED, font=ctk.CTkFont(size=10), anchor="w", wraplength=320).grid(row=1, column=0, sticky="ew", padx=18, pady=(0, 8))
        self.requires_size_switch = ctk.CTkSwitch(rules, text="Size is required", variable=self.requires_size_var, onvalue=True, offvalue=False, progress_color=PURPLE)
        self.requires_size_switch.grid(row=2, column=0, sticky="w", padx=18, pady=5)
        self.requires_color_switch = ctk.CTkSwitch(rules, text="Color is required", variable=self.requires_color_var, onvalue=True, offvalue=False, progress_color=PURPLE)
        self.requires_color_switch.grid(row=3, column=0, sticky="w", padx=18, pady=5)
        self.requires_decoration_switch = ctk.CTkSwitch(rules, text="Decoration is required", variable=self.requires_decoration_var, onvalue=True, offvalue=False, progress_color=PURPLE, command=self.on_requires_decoration_changed)
        self.requires_decoration_switch.grid(row=4, column=0, sticky="w", padx=18, pady=5)
        self.never_outsource_switch = ctk.CTkSwitch(
            rules, text="Never outsource — ship to Orchid", variable=self.never_outsource_var,
            onvalue=True, offvalue=False, progress_color=PURPLE,
        )
        self.never_outsource_switch.grid(row=5, column=0, sticky="w", padx=18, pady=(5, 3))
        ctk.CTkLabel(
            rules, text="Headwear defaults to Yes. Turn off only for a rare product that may be outsourced.",
            text_color=TEXT_MUTED, font=ctk.CTkFont(size=10), anchor="w", wraplength=320,
        ).grid(row=6, column=0, sticky="ew", padx=18, pady=(0, 14))

        self.advanced_button = ctk.CTkButton(left, text="Show Advanced Matching Fields", command=self.toggle_advanced_fields, height=36, fg_color="transparent", hover_color=PURPLE_LIGHT, border_width=1, border_color=PURPLE, text_color=PURPLE)
        self.advanced_button.grid(row=2, column=0, sticky="ew", pady=(12, 0))
        self.advanced_frame = ctk.CTkFrame(left, fg_color=CARD_BG, border_width=1, border_color=PURPLE_BORDER, corner_radius=12)
        self.advanced_frame.grid(row=3, column=0, sticky="ew", pady=(8, 0))
        self.advanced_frame.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(self.advanced_frame, text="Advanced Matching", text_color=PURPLE, font=ctk.CTkFont(size=16, weight="bold"), anchor="w").grid(row=0, column=0, sticky="ew", padx=18, pady=(14, 6))
        self.add_labeled_entry(self.advanced_frame, 1, "Product ID", self.product_id_var, "Generated internal identity; usually leave unchanged")
        self.add_labeled_entry(self.advanced_frame, 4, "Product Aliases", self.product_aliases_var, "Optional alternate names separated by commas")
        self.advanced_frame.grid_remove()

        tools = ctk.CTkFrame(left, fg_color=PURPLE_LIGHT, border_width=1, border_color=PURPLE_BORDER, corner_radius=12)
        tools.grid(row=4, column=0, sticky="ew", pady=(12, 8))
        tools.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(tools, text="Quick Tools", text_color=PURPLE, font=ctk.CTkFont(size=16, weight="bold"), anchor="w").grid(row=0, column=0, sticky="ew", padx=18, pady=(14, 8))
        self.apply_color_combo = ctk.CTkComboBox(tools, variable=self.apply_color_var, values=DECORATION_COLORS, height=36, border_color=PURPLE_BORDER, button_color="#EAE3F5", button_hover_color="#DED2EF", fg_color="#FFFFFF", text_color=TEXT_DARK)
        self.apply_color_combo.grid(row=1, column=0, sticky="ew", padx=18)
        ctk.CTkButton(tools, text="Apply Thread/Ink Color to All", command=self.apply_color_to_all, height=36, fg_color=PURPLE, hover_color=PURPLE_DARK).grid(row=2, column=0, sticky="ew", padx=18, pady=(8, 6))
        ctk.CTkButton(tools, text="Copy First Completed Color to All", command=self.copy_first_color_to_all, height=34, fg_color="transparent", hover_color="#E7DCF6", border_width=1, border_color=PURPLE, text_color=PURPLE).grid(row=3, column=0, sticky="ew", padx=18, pady=(0, 14))

    def add_labeled_entry(self, parent, row, label, variable, helper):
        ctk.CTkLabel(parent, text=label, text_color=TEXT_DARK, font=ctk.CTkFont(size=12, weight="bold"), anchor="w").grid(row=row, column=0, sticky="ew", padx=20, pady=(3, 4))
        ctk.CTkEntry(parent, textvariable=variable, height=38, border_color=PURPLE_BORDER, fg_color="#FFFFFF", text_color=TEXT_DARK).grid(row=row + 1, column=0, sticky="ew", padx=20)
        ctk.CTkLabel(parent, text=helper, text_color=TEXT_MUTED, font=ctk.CTkFont(size=10), anchor="w").grid(row=row + 2, column=0, sticky="ew", padx=20, pady=(3, 8))

    def build_color_panel(self, parent):
        panel = ctk.CTkFrame(parent, fg_color=CARD_BG, border_width=1, border_color=PURPLE_BORDER, corner_radius=12)
        panel.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        panel.grid_columnconfigure(0, weight=1)
        panel.grid_rowconfigure(3, weight=1)
        title_row = ctk.CTkFrame(panel, fg_color="transparent")
        title_row.grid(row=0, column=0, sticky="ew", padx=22, pady=(16, 4))
        title_row.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(title_row, text="③ Purchasing Colors and Decoration", text_color=PURPLE, font=ctk.CTkFont(size=17, weight="bold"), anchor="w").grid(row=0, column=0, sticky="ew")
        ctk.CTkButton(title_row, text="+ Add Garment Color", command=self.add_color_row, width=150, height=34, fg_color="transparent", hover_color=PURPLE_LIGHT, border_width=1, border_color=PURPLE, text_color=PURPLE).grid(row=0, column=1)
        ctk.CTkLabel(panel, text="Use the purchasing color sent to the vendor. Add vendor color codes and Shopify aliases when names differ. Thread/ink color is not required for Blank Garment (No Decoration).", text_color=TEXT_MUTED, font=ctk.CTkFont(size=12), anchor="w").grid(row=1, column=0, sticky="ew", padx=22, pady=(0, 10))

        table_header = ctk.CTkFrame(panel, fg_color=PURPLE_LIGHT, border_width=1, border_color=PURPLE_BORDER, corner_radius=8, height=44)
        table_header.grid(row=2, column=0, sticky="ew", padx=22)
        table_header.grid_columnconfigure(1, weight=2)
        table_header.grid_columnconfigure(2, weight=1)
        table_header.grid_columnconfigure(3, weight=2)
        table_header.grid_columnconfigure(4, weight=1)
        table_header.grid_columnconfigure(5, weight=0)
        ctk.CTkLabel(table_header, text="#", text_color=PURPLE, font=ctk.CTkFont(size=12, weight="bold"), width=38).grid(row=0, column=0, padx=(8, 0), pady=10)
        ctk.CTkLabel(table_header, text="Purchasing Color", text_color=PURPLE, font=ctk.CTkFont(size=12, weight="bold"), anchor="w").grid(row=0, column=1, sticky="ew", padx=8, pady=10)
        ctk.CTkLabel(table_header, text="Vendor Color Code (optional)", text_color=PURPLE, font=ctk.CTkFont(size=12, weight="bold"), anchor="w").grid(row=0, column=2, sticky="ew", padx=8, pady=10)
        ctk.CTkLabel(table_header, text="Shopify Color Aliases", text_color=PURPLE, font=ctk.CTkFont(size=12, weight="bold"), anchor="w").grid(row=0, column=3, sticky="ew", padx=8, pady=10)
        ctk.CTkLabel(table_header, text="Thread / Ink", text_color=PURPLE, font=ctk.CTkFont(size=12, weight="bold"), anchor="w").grid(row=0, column=4, sticky="ew", padx=8, pady=10)
        ctk.CTkLabel(table_header, text="Remove", text_color=PURPLE, font=ctk.CTkFont(size=12, weight="bold"), width=68).grid(row=0, column=5, padx=(4, 10), pady=10)

        self.colors_scroll = ctk.CTkScrollableFrame(panel, fg_color="#FFFFFF", corner_radius=8, border_width=1, border_color=PURPLE_BORDER)
        self.colors_scroll.grid(row=3, column=0, sticky="nsew", padx=22, pady=(6, 18))
        self.colors_scroll.grid_columnconfigure(0, weight=1)

    def build_action_bar(self, parent):
        actions = ctk.CTkFrame(parent, fg_color="transparent")
        actions.grid(row=5, column=0, sticky="ew", padx=20, pady=(4, 10))
        for column in range(4):
            actions.grid_columnconfigure(column, weight=1)
        kwargs = {"height": 48, "fg_color": PURPLE, "hover_color": PURPLE_DARK, "font": ctk.CTkFont(size=14, weight="bold")}
        ctk.CTkButton(actions, text="◀ Previous", command=self.previous_style, **kwargs).grid(row=0, column=0, sticky="ew", padx=(0, 7))
        self.save_button = ctk.CTkButton(actions, text="Save", command=self.save_current_style, **kwargs)
        self.save_button.grid(row=0, column=1, sticky="ew", padx=7)
        self.save_next_button = ctk.CTkButton(actions, text="Save & Next Setup Item", command=self.save_and_next_style, **kwargs)
        self.save_next_button.grid(row=0, column=2, sticky="ew", padx=7)
        ctk.CTkButton(actions, text="Next ▶", command=self.next_style, **kwargs).grid(row=0, column=3, sticky="ew", padx=(7, 0))

    def build_footer(self, parent):
        footer = ctk.CTkFrame(parent, fg_color=PURPLE_DARK, corner_radius=0, height=64)
        footer.grid(row=6, column=0, sticky="ew")
        footer.grid_columnconfigure(1, weight=1)
        footer_brand = ctk.CTkFrame(footer, fg_color="transparent")
        footer_brand.grid(row=0, column=0, padx=26, pady=12, sticky="w")
        footer_flower = Path(__file__).resolve().parent.parent / "assets" / "orchid_flower_white.png"
        if footer_flower.exists():
            source = PILImage.open(footer_flower).convert("RGBA")
            self.footer_flower_image = ctk.CTkImage(light_image=source, dark_image=source, size=(28, 28))
            ctk.CTkLabel(footer_brand, text="", image=self.footer_flower_image, width=32).pack(side="left", padx=(0, 7))
        ctk.CTkLabel(footer_brand, text="ORCHID PURCHASING SYSTEM", text_color="#FFFFFF", font=ctk.CTkFont(size=14, weight="bold")).pack(side="left")
        self.status_label = ctk.CTkLabel(footer, text="", text_color="#FFFFFF", font=ctk.CTkFont(size=12, weight="bold"))
        self.status_label.grid(row=0, column=1, pady=18)
        self.footer_record_label = ctk.CTkLabel(footer, text="", text_color="#FFFFFF", font=ctk.CTkFont(size=11))
        self.footer_record_label.grid(row=0, column=2, padx=26, pady=18, sticky="e")

    # ---------- editor behavior ----------
    def show_editor(self, target_key=None):
        self.dashboard_frame.grid_remove()
        self.editor_frame.grid()
        self.apply_filters(reset_position=True, target_key=target_key)

    def browse_all(self):
        self.search_var.set("")
        self.incomplete_only_var.set(False)
        self.catalog_filter_var.set("All Products")
        self._refresh_filter_button_styles()
        self.apply_filters(reset_position=True)

    def get_matching_style_keys(self):
        query = self.search_var.get().strip()
        catalog_filter = self.catalog_filter_var.get() if hasattr(self, "catalog_filter_var") else ("Needs Setup" if self.incomplete_only_var.get() else "All Products")
        ranked = []
        for position, key in enumerate(self.all_style_keys):
            rows = self.get_style_rows(key)
            is_complete = self.style_is_complete(rows)
            if catalog_filter == "Needs Setup" and is_complete:
                continue
            if catalog_filter == "Complete" and not is_complete:
                continue
            score = search_match_score(rows, query)
            if score is None:
                continue
            ranked.append((score, position, key))
        ranked.sort(key=lambda item: (item[0], item[1]))
        return [key for _, _, key in ranked]

    def open_first_editor_match(self):
        matches = self.get_matching_style_keys()
        if matches:
            self.show_editor(target_key=matches[0])
        else:
            self.status_label.configure(text="No Product Master records match that search.")

    def apply_filters(self, reset_position=True, target_key=None):
        self.filtered_style_keys = self.get_matching_style_keys()
        if target_key and target_key in self.filtered_style_keys:
            self.style_position = self.filtered_style_keys.index(target_key)
        elif reset_position:
            self.style_position = 0
        self.refresh_progress()
        self.show_current_style()

    def refresh_progress(self):
        total = len(self.all_style_keys)
        complete = self.completed_style_count()
        remaining = total - complete
        percent = complete / total if total else 0
        self.progress_text.configure(text=f"{total} saved styles • {complete} fully configured")
        self.progress_bar.set(percent)

    def _scroll_to_top(self):
        """Reset both editor panels for the newly displayed product."""
        for frame in (getattr(self, "left_scroll", None), getattr(self, "colors_scroll", None)):
            canvas = getattr(frame, "_parent_canvas", None) if frame is not None else None
            if canvas is not None:
                try:
                    canvas.yview_moveto(0.0)
                except Exception:
                    pass

    def _scroll_colors_to_bottom(self):
        canvas = getattr(getattr(self, "colors_scroll", None), "_parent_canvas", None)
        if canvas is not None:
            try:
                canvas.yview_moveto(1.0)
            except Exception:
                pass

    def clear_color_rows(self):
        for widget in self.colors_scroll.winfo_children():
            widget.destroy()
        self.color_controls = []

    def show_current_style(self):
        self.clear_color_rows()
        self.status_label.configure(text="")
        if not self.filtered_style_keys:
            self.current_style_key = None
            self.style_count_label.configure(text="No matching styles")
            self.product_name_label.configure(text="")
            self.issue_label.configure(text="")
            self.footer_record_label.configure(text="No records")
            self.product_name_var.set("")
            self.style_number_var.set("")
            self.vendor_var.set("")
            self.decoration_type_var.set("")
            self.decoration_location_var.set(LEFT_CHEST)
            self.custom_decoration_location_var.set("")
            self.decoration_placement_instructions_var.set("")
            self.product_id_var.set("")
            self.product_aliases_var.set("")
            self.category_var.set("")
            self.requires_size_var.set(True)
            self.requires_color_var.set(True)
            self.requires_decoration_var.set(True)
            self.never_outsource_var.set(False)
            return

        self.style_position = max(0, min(self.style_position, len(self.filtered_style_keys) - 1))
        key = self.filtered_style_keys[self.style_position]
        self.current_style_key = key
        rows = self.get_style_rows(key)
        issues = self.style_issues(rows)
        self.style_count_label.configure(text=f"Product {self.style_position + 1} of {len(self.filtered_style_keys)} in this view")
        self.footer_record_label.configure(text=f"{len(self.all_style_keys)} saved styles")
        display_name = self.style_display(key)
        self.product_name_label.configure(text=display_name)
        self.issue_label.configure(
            text="Fully configured" if not issues else "Additional setup available: " + " • ".join(issues),
            text_color=SUCCESS if not issues else TEXT_MUTED,
        )

        self.product_name_var.set(best_product_name(rows["Product Name"]))
        self.style_number_var.set(first_nonblank(rows["Style Number"]))
        self.product_id_var.set(first_nonblank(rows["Product ID"]) or default_product_id(first_nonblank(rows["Style Number"]), best_product_name(rows["Product Name"])))
        self.product_aliases_var.set(first_nonblank(rows["Product Aliases"]))
        rules = row_rules(rows.iloc[0])
        self.category_var.set(rules["Product Category"])
        self.requires_size_var.set(normalize_bool(rules["Requires Size"], True))
        self.requires_color_var.set(normalize_bool(rules["Requires Color"], True))
        self.requires_decoration_var.set(normalize_bool(rules["Requires Decoration"], True))
        self.never_outsource_var.set(normalize_bool(first_nonblank(rows[NEVER_OUTSOURCE_COLUMN]), default_never_outsource(
            self.product_name_var.get(), self.category_var.get(), self.style_number_var.get()
        )))
        vendors = list(dict.fromkeys(clean_text(v) for v in rows["Vendor"] if clean_text(v)))
        decorations = list(dict.fromkeys(clean_text(v) for v in rows["Decoration Type"] if clean_text(v)))
        locations = list(dict.fromkeys(clean_text(v) for v in rows["Decoration Location"] if clean_text(v)))
        self.vendor_var.set(vendors[0] if len(vendors) == 1 else "")
        self.decoration_type_var.set(decorations[0] if len(decorations) == 1 else "")
        stored_location = normalize_decoration_location(
            locations[0] if len(locations) == 1 else "",
            self.decoration_type_var.get(),
            requires_decoration=self.requires_decoration_var.get(),
            product_name=self.product_name_var.get(), category=self.category_var.get(),
        )
        location_choice, custom_location = location_choice_and_custom(stored_location, self.decoration_type_var.get())
        self.decoration_location_var.set(location_choice)
        self.custom_decoration_location_var.set(custom_location)
        instructions = list(dict.fromkeys(
            clean_text(v) for v in rows.get("Decoration Placement Instructions", pd.Series(dtype=str)) if clean_text(v)
        ))
        self.decoration_placement_instructions_var.set(instructions[0] if len(instructions) == 1 else "")
        self.on_decoration_location_selected(location_choice)

        # Apply visible defaults immediately for any legacy/incomplete record that
        # predates the Product Intelligence Engine. These remain editable.
        detected_brand, preferred_vendor = preferred_vendor_for_brand_text(
            f"{self.product_name_var.get()} {self.style_number_var.get()}"
        )
        if not self.vendor_var.get().strip() and preferred_vendor:
            self.vendor_var.set(preferred_vendor)
        if not self.decoration_type_var.get().strip():
            inferred_decoration = infer_default_decoration(
                self.category_var.get(), self.product_name_var.get(), self.style_number_var.get()
            )
            if inferred_decoration:
                self.decoration_type_var.set(inferred_decoration)

        sorted_rows = rows.sort_values(by="Garment Color", key=lambda series: series.astype(str).str.casefold())
        for display_index, (data_index, row) in enumerate(sorted_rows.iterrows(), start=1):
            self.render_color_row(display_index, data_index, row)
        self.update_decoration_color_state()
        self.after_idle(self._scroll_to_top)

    def render_color_row(self, display_index, data_index, row):
        row_color = ROW_ALT if display_index % 2 == 0 else "#FFFFFF"
        frame = ctk.CTkFrame(self.colors_scroll, fg_color=row_color, corner_radius=0, height=58)
        frame.grid(row=display_index - 1, column=0, sticky="ew")
        frame.grid_columnconfigure(1, weight=2)
        frame.grid_columnconfigure(2, weight=1)
        frame.grid_columnconfigure(3, weight=2)
        frame.grid_columnconfigure(4, weight=1)
        frame.grid_columnconfigure(5, weight=0)
        ctk.CTkLabel(frame, text=str(display_index), text_color="#FFFFFF", fg_color=PURPLE, width=28, height=28, corner_radius=14, font=ctk.CTkFont(size=11, weight="bold")).grid(row=0, column=0, padx=(10, 8), pady=14)
        garment_var = ctk.StringVar(value=clean_text(row.get("Garment Color", "")))
        vendor_code_var = ctk.StringVar(value=clean_text(row.get("Vendor Color Code", "")))
        aliases_var = ctk.StringVar(value=clean_text(row.get("Color Aliases", "")))
        decoration_var = ctk.StringVar(value=clean_text(row.get("Decoration Color", "")))
        garment_entry = ctk.CTkEntry(frame, textvariable=garment_var, height=36, border_color=PURPLE_BORDER, fg_color="#FFFFFF", text_color=TEXT_DARK)
        garment_entry.grid(row=0, column=1, sticky="ew", padx=5, pady=10)
        ctk.CTkEntry(frame, textvariable=vendor_code_var, height=36, border_color=PURPLE_BORDER, fg_color="#FFFFFF", text_color=TEXT_DARK).grid(row=0, column=2, sticky="ew", padx=5, pady=10)
        ctk.CTkEntry(frame, textvariable=aliases_var, placeholder_text="e.g. Dark Indigo, Navy", height=36, border_color=PURPLE_BORDER, fg_color="#FFFFFF", text_color=TEXT_DARK).grid(row=0, column=3, sticky="ew", padx=5, pady=10)
        decoration_combo = ctk.CTkComboBox(frame, variable=decoration_var, values=DECORATION_COLORS, height=36, border_color=PURPLE_BORDER, button_color="#EAE3F5", button_hover_color="#DED2EF", fg_color="#FFFFFF", text_color=TEXT_DARK, dropdown_fg_color="#FFFFFF", dropdown_text_color=TEXT_DARK)
        decoration_combo.grid(row=0, column=4, sticky="ew", padx=5, pady=10)
        delete_button = ctk.CTkButton(
            frame, text="Delete", width=62, height=32,
            fg_color="#FFF1F0", hover_color="#FDE2E0",
            border_width=1, border_color="#E5A09A", text_color=DANGER,
            command=lambda index=data_index: self.delete_color_row(index),
        )
        delete_button.grid(row=0, column=5, padx=(4, 10), pady=10)
        self.color_controls.append({
            "index": data_index,
            "garment_var": garment_var,
            "garment_entry": garment_entry,
            "vendor_code_var": vendor_code_var,
            "aliases_var": aliases_var,
            "decoration_var": decoration_var,
            "decoration_combo": decoration_combo,
            "delete_button": delete_button,
        })

    def toggle_advanced_fields(self):
        self.advanced_visible = not self.advanced_visible
        if self.advanced_visible:
            self.advanced_frame.grid()
            self.advanced_button.configure(text="Hide Advanced Matching Fields")
        else:
            self.advanced_frame.grid_remove()
            self.advanced_button.configure(text="Show Advanced Matching Fields")

    def on_category_selected(self, selected_category):
        defaults = category_defaults(selected_category, self.product_name_var.get(), self.style_number_var.get())
        self.requires_size_var.set(defaults["requires_size"])
        self.requires_color_var.set(defaults["requires_color"])
        self.requires_decoration_var.set(defaults["requires_decoration"])
        self.never_outsource_var.set(default_never_outsource(
            self.product_name_var.get(), self.category_var.get(), self.style_number_var.get()
        ))
        if not defaults["requires_decoration"]:
            self.decoration_type_var.set(BLANK_DECORATION_LABEL)
        elif is_blank_decoration(self.decoration_type_var.get()):
            self.decoration_type_var.set("")
        if not self.decoration_type_var.get().strip():
            inferred_decoration = infer_default_decoration(
                selected_category, self.product_name_var.get(), self.style_number_var.get()
            )
            if inferred_decoration:
                self.decoration_type_var.set(inferred_decoration)
        self.update_decoration_color_state()
        self.status_label.configure(text=f"Applied {selected_category} purchase defaults.")

    def on_decoration_location_selected(self, choice: str):
        is_custom = clean_text(choice) == OTHER_CUSTOM
        if hasattr(self, "custom_location_entry"):
            self.custom_location_entry.configure(state="normal" if is_custom else "disabled")
        if hasattr(self, "custom_location_label"):
            self.custom_location_label.configure(text_color=PURPLE if is_custom else TEXT_MUTED)
        if not is_custom:
            self.custom_decoration_location_var.set("")

    def resolved_decoration_location(self, decoration_type: str, requires_decoration: bool) -> str:
        if is_blank_decoration(decoration_type) or is_in_house_decoration(decoration_type):
            return ""
        choice = self.decoration_location_var.get()
        location = resolve_location_choice(choice, self.custom_decoration_location_var.get())
        return normalize_decoration_location(
            location, decoration_type, requires_decoration=requires_decoration,
            product_name=self.product_name_var.get(), category=self.category_var.get(),
        )

    def on_requires_decoration_changed(self):
        if not self.requires_decoration_var.get() and not is_in_house_decoration(self.decoration_type_var.get()):
            self.decoration_type_var.set(BLANK_DECORATION_LABEL)
        elif self.requires_decoration_var.get() and is_blank_decoration(self.decoration_type_var.get()):
            self.decoration_type_var.set("")
        self.update_decoration_color_state()

    def on_product_name_changed(self, *_):
        if not self.category_var.get().strip():
            self.category_var.set(infer_category(self.product_name_var.get(), self.style_number_var.get()))
        inferred_service = infer_in_house_decoration(self.product_name_var.get())
        if inferred_service and not self.decoration_type_var.get().strip():
            self.decoration_type_var.set(inferred_service)
            self.requires_size_var.set(False)
            self.requires_color_var.set(False)
            self.requires_decoration_var.set(False)
        elif is_blank_garment_product(self.product_name_var.get(), style_number=self.style_number_var.get()) and not self.decoration_type_var.get().strip():
            self.decoration_type_var.set(BLANK_DECORATION_LABEL)

    def on_decoration_type_changed(self, *_):
        self.update_decoration_color_state()

    def update_decoration_color_state(self):
        decoration_type = self.decoration_type_var.get()
        in_house = is_in_house_decoration(decoration_type)
        none_selected = is_blank_decoration(decoration_type) or in_house or not self.requires_decoration_var.get()
        if in_house:
            self.requires_decoration_var.set(False)
        if hasattr(self, "decoration_location_combo"):
            if none_selected:
                display = NOT_APPLICABLE_IN_HOUSE if in_house else NOT_APPLICABLE_NO_DECORATION
                self.decoration_location_var.set(display)
                self.custom_decoration_location_var.set("")
                self.decoration_location_combo.configure(state="disabled")
                if hasattr(self, "custom_location_entry"):
                    self.custom_location_entry.configure(state="disabled")
                if hasattr(self, "placement_instructions_entry"):
                    self.placement_instructions_entry.configure(state="disabled")
                    self.decoration_placement_instructions_var.set("")
            else:
                current = self.decoration_location_var.get().strip()
                if not current or current in {NOT_APPLICABLE_IN_HOUSE, NOT_APPLICABLE_NO_DECORATION}:
                    self.decoration_location_var.set(default_decoration_location(
                        self.product_name_var.get(), self.category_var.get(), decoration_type
                    ))
                self.decoration_location_combo.configure(state="normal")
                if hasattr(self, "placement_instructions_entry"):
                    self.placement_instructions_entry.configure(state="normal")
                self.on_decoration_location_selected(self.decoration_location_var.get())
        for control in self.color_controls:
            combo = control.get("decoration_combo")
            if none_selected:
                control["decoration_var"].set("")
                if combo is not None:
                    combo.configure(state="disabled")
            elif combo is not None:
                combo.configure(state="normal")
        if hasattr(self, "apply_color_combo"):
            self.apply_color_combo.configure(state="disabled" if none_selected else "normal")

    def delete_color_row(self, data_index):
        rows = self.get_style_rows(self.current_style_key) if self.current_style_key else pd.DataFrame()
        if len(rows) <= 1:
            messagebox.showinfo("Color Required", "A product must keep at least one garment-color row.")
            return
        if data_index not in self.master_data.index:
            return
        color = clean_text(self.master_data.at[data_index, "Garment Color"]) or "blank color row"
        if not messagebox.askyesno("Delete Garment Color", f"Delete {color!r} from this product?"):
            return
        self.ensure_session_backup()
        current_key = self.current_style_key
        self.master_data = self.master_data.drop(index=data_index).reset_index(drop=True)
        saved = save_product_master(self.master_data)
        if saved is None:
            return
        self.master_data = saved
        self.build_style_list()
        self.filtered_style_keys = self.get_matching_style_keys()
        if current_key in self.filtered_style_keys:
            self.style_position = self.filtered_style_keys.index(current_key)
        elif self.filtered_style_keys:
            self.style_position = min(self.style_position, len(self.filtered_style_keys) - 1)
        self.show_current_style()
        self.refresh_progress()
        self.status_label.configure(text=f"Deleted {color}.")

    def add_color_row(self):
        if not self.current_style_key:
            return
        new_row = {
            "Product Name": self.product_name_var.get().strip(),
            "Style Number": self.style_number_var.get().strip(),
            "Garment Color": "",
            "Vendor": self.vendor_var.get().strip(),
            "Decoration Type": self.decoration_type_var.get().strip(),
            "Decoration Location": self.resolved_decoration_location(
                self.decoration_type_var.get(), self.requires_decoration_var.get()
            ),
            "Decoration Placement Instructions": normalize_space(self.decoration_placement_instructions_var.get()),
            "Decoration Color": "",
            "Product ID": self.product_id_var.get().strip() or default_product_id(self.style_number_var.get(), self.product_name_var.get()),
            "Product Aliases": self.product_aliases_var.get().strip(),
            "Vendor Color Code": "",
            "Color Aliases": "",
            "Product Category": self.category_var.get().strip() or infer_category(self.product_name_var.get(), self.style_number_var.get()),
            "Requires Size": bool_text(self.requires_size_var.get()),
            "Requires Color": bool_text(self.requires_color_var.get()),
            "Requires Decoration": bool_text(self.requires_decoration_var.get()),
            NEVER_OUTSOURCE_COLUMN: bool_text(self.never_outsource_var.get()),
            "Setup Required": "Yes",
        }
        self.master_data = pd.concat([self.master_data, pd.DataFrame([new_row])], ignore_index=True)
        self.build_style_list()
        new_key = style_group_key(pd.Series(new_row))
        if new_key not in self.filtered_style_keys:
            self.filtered_style_keys.append(new_key)
        self.current_style_key = new_key
        self.style_position = self.filtered_style_keys.index(new_key)
        self.show_current_style()
        def focus_new_row():
            self._scroll_colors_to_bottom()
            blank_controls = [control for control in self.color_controls if control["index"] == self.master_data.index[-1]]
            if blank_controls:
                blank_controls[0]["garment_entry"].focus_set()
        self.after(120, focus_new_row)
        self.status_label.configure(text="Blank garment-color row added. Complete it before saving.")

    def apply_color_to_all(self):
        value = self.apply_color_var.get().strip()
        for control in self.color_controls:
            control["decoration_var"].set(value)
        self.status_label.configure(text=f"Applied {value or 'blank'} to all garment colors.")

    def copy_first_color_to_all(self):
        value = next((control["decoration_var"].get().strip() for control in self.color_controls if control["decoration_var"].get().strip()), "")
        if not value:
            messagebox.showinfo("No Completed Color", "Enter one thread or ink color first, then use this button.")
            return
        self.apply_color_var.set(value)
        self.apply_color_to_all()

    def ensure_session_backup(self):
        if self.session_backup_path or not PRODUCT_MASTER_PATH.exists():
            return
        backup_dir = PRODUCT_MASTER_PATH.parent / "backups"
        backup_dir.mkdir(parents=True, exist_ok=True)
        self.session_backup_path = backup_dir / f"product_master_before_pro_session_{datetime.now():%Y%m%d_%H%M%S}.csv"
        shutil.copy2(PRODUCT_MASTER_PATH, self.session_backup_path)

    def _mark_catalog_updated(self, style_name: str = ""):
        try:
            payload = {
                "saved_at": datetime.now().isoformat(timespec="seconds"),
                "style": style_name,
                "complete": self.completed_style_count(),
                "total": len(self.all_style_keys),
            }
            PRODUCT_MASTER_UPDATE_MARKER.parent.mkdir(parents=True, exist_ok=True)
            temporary = PRODUCT_MASTER_UPDATE_MARKER.with_suffix(".tmp")
            temporary.write_text(json.dumps(payload, indent=2), encoding="utf-8")
            temporary.replace(PRODUCT_MASTER_UPDATE_MARKER)
        except Exception:
            pass

    def _show_save_confirmation(self, message: str, complete: bool):
        if not hasattr(self, "save_toast"):
            self.save_toast = ctk.CTkLabel(
                self.editor_frame, text="", height=38, corner_radius=19,
                fg_color="#EAF8EF", text_color=SUCCESS,
                font=ctk.CTkFont(size=13, weight="bold"),
            )
        if self._save_toast_after_id:
            try:
                self.after_cancel(self._save_toast_after_id)
            except Exception:
                pass
        self.save_toast.configure(
            text=message,
            fg_color="#EAF8EF" if complete else "#FFF1DC",
            text_color=SUCCESS if complete else WARNING,
        )
        self.save_toast.place(relx=0.5, y=112, anchor="n")
        self.save_toast.lift()
        self._save_toast_after_id = self.after(1900, self.save_toast.place_forget)

    def save_current_style(self):
        if not self.current_style_key:
            return False
        rows = self.get_style_rows(self.current_style_key)
        if rows.empty:
            return False
        self.ensure_session_backup()

        product_name = canonical_product_name(self.product_name_var.get())
        style_number = normalize_style(self.style_number_var.get())
        vendor = canonical_vendor_name(self.vendor_var.get())
        vendor = next(
            (existing for existing in self.vendor_options if existing.casefold() == vendor.casefold()),
            vendor,
        )
        product_id = self.product_id_var.get().strip() or default_product_id(style_number, product_name)
        product_aliases = normalize_space(self.product_aliases_var.get())
        category = self.category_var.get().strip() or infer_category(product_name, style_number)
        requires_size = self.requires_size_var.get()
        requires_color = self.requires_color_var.get()
        requires_decoration = self.requires_decoration_var.get()
        decoration_type = self.decoration_type_var.get().strip()
        if is_in_house_decoration(decoration_type):
            requires_decoration = False
            self.requires_decoration_var.set(False)
        elif not requires_decoration:
            decoration_type = BLANK_DECORATION_LABEL
            self.decoration_type_var.set(BLANK_DECORATION_LABEL)
        if is_blank_garment_product(product_name, style_number=style_number) and not decoration_type:
            decoration_type = BLANK_DECORATION_LABEL
            self.decoration_type_var.set(BLANK_DECORATION_LABEL)
        if requires_decoration and not is_blank_decoration(decoration_type) and not is_in_house_decoration(decoration_type) and self.decoration_location_var.get() == OTHER_CUSTOM and not normalize_space(self.custom_decoration_location_var.get()):
            messagebox.showwarning(
                "Custom Location Required",
                "Enter the custom decoration location before saving this product.",
            )
            self.custom_location_entry.focus_set()
            return False
        decoration_location = self.resolved_decoration_location(decoration_type, requires_decoration)
        placement_instructions = normalize_space(self.decoration_placement_instructions_var.get())
        if not requires_decoration or is_blank_decoration(decoration_type) or is_in_house_decoration(decoration_type):
            placement_instructions = ""
            self.decoration_placement_instructions_var.set("")
        choice, custom = location_choice_and_custom(decoration_location, decoration_type)
        self.decoration_location_var.set(choice)
        self.custom_decoration_location_var.set(custom)

        for index in rows.index:
            self.master_data.at[index, "Product Name"] = product_name
            self.master_data.at[index, "Style Number"] = style_number
            self.master_data.at[index, "Vendor"] = vendor
            self.master_data.at[index, "Product ID"] = product_id
            self.master_data.at[index, "Product Aliases"] = product_aliases
            self.master_data.at[index, "Product Category"] = category
            self.master_data.at[index, "Requires Size"] = bool_text(requires_size)
            self.master_data.at[index, "Requires Color"] = bool_text(requires_color)
            self.master_data.at[index, "Requires Decoration"] = bool_text(requires_decoration)
            self.master_data.at[index, NEVER_OUTSOURCE_COLUMN] = bool_text(self.never_outsource_var.get())
            if "Setup Required" in self.master_data.columns:
                self.master_data.at[index, "Setup Required"] = "No"
            self.master_data.at[index, "Decoration Type"] = normalize_decoration_type(decoration_type)
            self.master_data.at[index, "Decoration Location"] = normalize_decoration_location(
                decoration_location, decoration_type, requires_decoration=requires_decoration
            )
            self.master_data.at[index, "Decoration Placement Instructions"] = placement_instructions
        for control in self.color_controls:
            index = control["index"]
            if index in self.master_data.index:
                self.master_data.at[index, "Garment Color"] = normalize_space(control["garment_var"].get())
                self.master_data.at[index, "Vendor Color Code"] = normalize_space(control["vendor_code_var"].get())
                self.master_data.at[index, "Color Aliases"] = normalize_space(control["aliases_var"].get())
                self.master_data.at[index, "Decoration Color"] = "" if (is_blank_decoration(decoration_type) or is_in_house_decoration(decoration_type)) else control["decoration_var"].get().strip()

        saved = save_product_master(self.master_data)
        if saved is None:
            return False
        self.master_data = saved
        if vendor and not any(item.casefold() == vendor.casefold() for item in self.vendor_options):
            self.vendor_options.append(vendor)
            self.vendor_combo.configure(values=self.vendor_options)
            if hasattr(self, "dashboard_vendor_filter"):
                self.dashboard_vendor_filter.configure(
                    values=list(dict.fromkeys(["All Vendors"] + self.vendor_options + ["Unassigned"]))
                )
        self.build_style_list()
        new_key = f"style:{style_number}" if style_number else f"product:{product_name.casefold()}"
        self.current_style_key = new_key if new_key in self.all_style_keys else None
        self.session_saved_styles.add(style_number or product_name or self.current_style_key or "Unknown")
        self.refresh_progress()

        current_rows = self.get_style_rows(new_key) if new_key in self.all_style_keys else pd.DataFrame(columns=COLUMNS)
        issues = self.style_issues(current_rows)
        saved_name = style_number or product_name or "Product"
        self._mark_catalog_updated(saved_name)
        if issues:
            message = "Saved — additional setup available: " + ", ".join(issues)
            self.status_label.configure(text=message)
            self._show_save_confirmation(f"Saved {saved_name} — setup still needed", complete=False)
        else:
            self.status_label.configure(text="Saved ✓ Product Master record is fully configured.")
            self._show_save_confirmation(f"✓ {saved_name} saved", complete=True)
        return True

    def save_and_next_style(self):
        current_key = self.current_style_key
        try:
            global_position = self.all_style_keys.index(current_key)
        except (ValueError, TypeError):
            global_position = -1
        if not self.save_current_style():
            return
        self.filtered_style_keys = self.get_matching_style_keys()
        if not self.filtered_style_keys:
            self.show_dashboard()
            messagebox.showinfo("Setup View Complete", "There are no additional setup items remaining in the current view.")
            return
        next_key = next((key for key in self.all_style_keys[global_position + 1 :] if key in self.filtered_style_keys), None)
        if next_key is None:
            next_key = self.filtered_style_keys[0]
        self.style_position = self.filtered_style_keys.index(next_key)
        self.show_current_style()
        self.after(120, self._scroll_to_top)
        self.refresh_progress()

    def previous_style(self):
        if self.filtered_style_keys and self.style_position > 0:
            self.style_position -= 1
            self.show_current_style()

    def next_style(self):
        if self.filtered_style_keys and self.style_position < len(self.filtered_style_keys) - 1:
            self.style_position += 1
            self.show_current_style()

    def on_close(self):
        try:
            if PRODUCT_MASTER_PID_FILE.exists() and PRODUCT_MASTER_PID_FILE.read_text(encoding="utf-8").strip() == str(os.getpid()):
                PRODUCT_MASTER_PID_FILE.unlink(missing_ok=True)
            PRODUCT_MASTER_REQUEST_FILE.unlink(missing_ok=True)
        except Exception:
            pass
        if self.session_saved_styles:
            self.build_style_list()
            complete_now = self.completed_style_count()
            gained = max(0, complete_now - self.session_start_complete)
            messagebox.showinfo(
                "Product Master Session",
                f"Styles saved this session: {len(self.session_saved_styles)}\n"
                f"Newly completed styles: {gained}\n"
                f"Product Master health: {complete_now} of {len(self.all_style_keys)} complete\n\n"
                f"Active file:\n{PRODUCT_MASTER_PATH}",
            )
        self.destroy()


if __name__ == "__main__":
    import sys
    prefill = ""
    search = ""
    setup = ""
    if "--new-product" in sys.argv:
        try:
            prefill = sys.argv[sys.argv.index("--new-product") + 1]
        except (ValueError, IndexError):
            prefill = ""
    if "--search-product" in sys.argv:
        try:
            search = sys.argv[sys.argv.index("--search-product") + 1]
        except (ValueError, IndexError):
            search = ""
    if "--setup-product" in sys.argv:
        try:
            setup = sys.argv[sys.argv.index("--setup-product") + 1]
        except (ValueError, IndexError):
            setup = ""
    ProductMasterV2(prefill_product=prefill, search_product=search, setup_product=setup).mainloop()
