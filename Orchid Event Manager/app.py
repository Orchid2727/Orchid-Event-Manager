from __future__ import annotations

from pathlib import Path
import subprocess
import sys

import customtkinter as ctk
from tkinter import filedialog, messagebox

from modules.shopify_parser import parse_shopify_orders
from modules.catalog_manager import clean_product_master, merge_parsed_products
from modules.purchase_order_generator import generate_purchase_orders


PROJECT = Path(__file__).resolve().parent
DATA = PROJECT / "data"
MASTER = DATA / "product_master.csv"

PURPLE = "#5B2AA8"
PURPLE_DARK = "#3D176F"
PURPLE_LIGHT = "#F2ECFB"
BORDER = "#D7C8EE"
TEXT = "#201A2D"
MUTED = "#6F667A"
BG = "#F7F5FA"
WHITE = "#FFFFFF"
SUCCESS = "#2D7A46"


class OrchidEventManager(ctk.CTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("light")
        self.title("Orchid Event Manager v1.4")
        self.geometry("1240x820")
        self.minsize(1100, 720)
        self.configure(fg_color=BG)
        self.selected_csv: Path | None = None
        self.build_ui()
        self.refresh_dashboard()

    def build_ui(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.build_sidebar()
        self.build_main_area()

    def build_sidebar(self):
        sidebar = ctk.CTkFrame(self, width=250, corner_radius=0, fg_color=PURPLE_DARK)
        sidebar.grid(row=0, column=0, sticky="nsew")
        sidebar.grid_propagate(False)
        sidebar.grid_rowconfigure(8, weight=1)

        ctk.CTkLabel(sidebar, text="✿", text_color=WHITE,
                     font=ctk.CTkFont(size=44, weight="bold")).grid(
            row=0, column=0, padx=24, pady=(28, 4), sticky="w"
        )
        ctk.CTkLabel(sidebar, text="ORCHID", text_color=WHITE,
                     font=ctk.CTkFont(size=24, weight="bold"), anchor="w").grid(
            row=1, column=0, padx=24, sticky="w"
        )
        ctk.CTkLabel(sidebar, text="EVENT MANAGER", text_color="#D9C8F2",
                     font=ctk.CTkFont(size=11, weight="bold"), anchor="w").grid(
            row=2, column=0, padx=24, pady=(0, 22), sticky="w"
        )

        self.nav_button(sidebar, 3, "⌂  Dashboard", self.show_dashboard, True)
        self.nav_button(sidebar, 4, "⇩  Shopify Import", self.choose_csv)
        self.nav_button(sidebar, 5, "▣  Product Master", self.open_product_master)
        self.nav_button(sidebar, 6, "▤  Catalog Maintenance", self.clean_catalog)

        ctk.CTkLabel(sidebar, text="COMING SOON", text_color="#BCA7DD",
                     font=ctk.CTkFont(size=10, weight="bold"), anchor="w").grid(
            row=7, column=0, padx=24, pady=(24, 8), sticky="w"
        )

        coming_soon = ctk.CTkLabel(
            sidebar,
            text="□  Purchase Orders\n\n□  Embroidery Queue\n\n□  Screen Print Queue\n\n□  Reports",
            text_color="#AFA4BD",
            font=ctk.CTkFont(size=13),
            justify="left",
            anchor="nw",
        )
        coming_soon.grid(row=8, column=0, padx=28, pady=4, sticky="nw")

        ctk.CTkLabel(sidebar, text="v1.4 PDF Purchase Orders", text_color="#D9C8F2",
                     font=ctk.CTkFont(size=11, weight="bold")).grid(
            row=9, column=0, padx=24, pady=(8, 16), sticky="sw"
        )

        ctk.CTkButton(
            sidebar, text="Exit", width=200, height=40,
            fg_color="transparent", hover_color=PURPLE,
            border_width=1, border_color="#BCA7DD", text_color=WHITE,
            command=self.destroy,
        ).grid(row=10, column=0, padx=24, pady=(0, 24), sticky="sw")

    def nav_button(self, parent, row, text, command, active=False):
        ctk.CTkButton(
            parent, text=text, command=command, anchor="w",
            width=205, height=42, corner_radius=8,
            fg_color=PURPLE if active else "transparent",
            hover_color=PURPLE, text_color=WHITE,
            font=ctk.CTkFont(size=13, weight="bold" if active else "normal"),
        ).grid(row=row, column=0, padx=22, pady=4, sticky="ew")

    def build_main_area(self):
        main = ctk.CTkFrame(self, fg_color=BG, corner_radius=0)
        main.grid(row=0, column=1, sticky="nsew")
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(3, weight=1)

        header = ctk.CTkFrame(main, fg_color=WHITE, corner_radius=0, height=95)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(header, text="Dashboard", text_color=TEXT,
                     font=ctk.CTkFont(size=30, weight="bold"), anchor="w").grid(
            row=0, column=0, padx=28, pady=(20, 0), sticky="w"
        )
        ctk.CTkLabel(
            header,
            text="Import orders, maintain your catalog, and prepare for purchase order automation.",
            text_color=MUTED, font=ctk.CTkFont(size=13), anchor="w",
        ).grid(row=1, column=0, padx=28, pady=(2, 18), sticky="w")

        self.status = ctk.CTkLabel(
            main, text="Ready", text_color=MUTED,
            font=ctk.CTkFont(size=13), anchor="w"
        )
        self.status.grid(row=1, column=0, sticky="ew", padx=28, pady=(14, 2))

        summary = ctk.CTkFrame(main, fg_color="transparent")
        summary.grid(row=2, column=0, sticky="ew", padx=28, pady=(10, 6))
        summary.grid_columnconfigure((0, 1, 2, 3), weight=1)

        self.total_records_card = self.stat_card(summary, 0, "Product Records", "0", "Current catalog rows")
        self.complete_styles_card = self.stat_card(summary, 1, "Completed Styles", "—", "From Product Master")
        self.selected_file_card = self.stat_card(summary, 2, "Selected CSV", "None", "Choose a Shopify export")
        self.review_card = self.stat_card(summary, 3, "Review Needed", "—", "After processing")

        content = ctk.CTkFrame(main, fg_color="transparent")
        content.grid(row=3, column=0, sticky="nsew", padx=28, pady=(8, 20))
        content.grid_columnconfigure((0, 1, 2), weight=1)
        content.grid_rowconfigure((0, 1), weight=1)

        self.action_card(content, 0, 0, "1", "Import Shopify Orders",
                         "Choose a Shopify CSV file from your Mac.",
                         "Choose CSV", self.choose_csv)
        self.action_card(content, 0, 1, "2", "Process Orders",
                         "Parse the selected CSV and update the Product Master safely.",
                         "Process Selected CSV", self.process_csv)
        self.action_card(content, 0, 2, "3", "Product Master",
                         "Open the Style Mode editor to review vendor and decoration rules.",
                         "Open Product Master", self.open_product_master)
        self.action_card(content, 1, 0, "4", "Catalog Maintenance",
                         "Create a backup and remove duplicate catalog records.",
                         "Clean & Backup", self.clean_catalog)
        self.action_card(content, 1, 1, "5", "Purchase Orders",
                         "Generate grouped CSV reports and polished PDF purchase orders.",
                         "Generate Purchase Orders", self.generate_purchase_orders)
        self.future_card(content, 1, 2, "Production Queues",
                         "Embroidery and screen print workflow tools are coming next.")

    def stat_card(self, parent, column, title, value, subtitle):
        frame = ctk.CTkFrame(parent, fg_color=WHITE, border_width=1,
                             border_color=BORDER, corner_radius=12)
        frame.grid(row=0, column=column, sticky="ew", padx=6, pady=4)
        value_label = ctk.CTkLabel(
            frame, text=value, text_color=PURPLE,
            font=ctk.CTkFont(size=24, weight="bold")
        )
        value_label.pack(anchor="w", padx=18, pady=(15, 2))
        ctk.CTkLabel(frame, text=title, text_color=TEXT,
                     font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w", padx=18)
        ctk.CTkLabel(frame, text=subtitle, text_color=MUTED,
                     font=ctk.CTkFont(size=10)).pack(anchor="w", padx=18, pady=(2, 14))
        return value_label

    def action_card(self, parent, row, col, number, title, body, button_text, command):
        frame = ctk.CTkFrame(parent, fg_color=WHITE, border_width=1,
                             border_color=BORDER, corner_radius=14)
        frame.grid(row=row, column=col, sticky="nsew", padx=7, pady=7)
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(2, weight=1)

        ctk.CTkLabel(frame, text=number, width=36, height=36, corner_radius=18,
                     fg_color=PURPLE, text_color=WHITE,
                     font=ctk.CTkFont(size=14, weight="bold")).grid(
            row=0, column=0, padx=18, pady=(18, 8), sticky="w"
        )
        ctk.CTkLabel(frame, text=title, text_color=TEXT,
                     font=ctk.CTkFont(size=17, weight="bold"), anchor="w").grid(
            row=1, column=0, padx=18, sticky="ew"
        )
        ctk.CTkLabel(frame, text=body, text_color=MUTED,
                     font=ctk.CTkFont(size=12), wraplength=260,
                     justify="left", anchor="nw").grid(
            row=2, column=0, padx=18, pady=(8, 14), sticky="nsew"
        )
        ctk.CTkButton(frame, text=button_text, height=40,
                      fg_color=PURPLE, hover_color=PURPLE_DARK,
                      font=ctk.CTkFont(size=12, weight="bold"),
                      command=command).grid(
            row=3, column=0, sticky="ew", padx=18, pady=(0, 18)
        )

    def future_card(self, parent, row, col, title, body):
        frame = ctk.CTkFrame(parent, fg_color=PURPLE_LIGHT, border_width=1,
                             border_color=BORDER, corner_radius=14)
        frame.grid(row=row, column=col, sticky="nsew", padx=7, pady=7)
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(2, weight=1)

        ctk.CTkLabel(frame, text="COMING SOON", text_color=PURPLE,
                     font=ctk.CTkFont(size=10, weight="bold")).grid(
            row=0, column=0, padx=18, pady=(18, 8), sticky="w"
        )
        ctk.CTkLabel(frame, text=title, text_color=TEXT,
                     font=ctk.CTkFont(size=17, weight="bold"), anchor="w").grid(
            row=1, column=0, padx=18, sticky="ew"
        )
        ctk.CTkLabel(frame, text=body, text_color=MUTED,
                     font=ctk.CTkFont(size=12), wraplength=260,
                     justify="left", anchor="nw").grid(
            row=2, column=0, padx=18, pady=(8, 14), sticky="nsew"
        )
        ctk.CTkButton(frame, text="Not Available Yet", height=40,
                      fg_color="#D6CCE2", hover_color="#D6CCE2",
                      text_color="#8B8295", state="disabled").grid(
            row=3, column=0, sticky="ew", padx=18, pady=(0, 18)
        )

    def show_dashboard(self):
        self.status.configure(text="Dashboard ready.", text_color=MUTED)

    def refresh_dashboard(self):
        total_records = 0
        if MASTER.exists():
            try:
                import pandas as pd
                total_records = len(pd.read_csv(MASTER, dtype=str).fillna(""))
            except Exception:
                total_records = 0
        self.total_records_card.configure(text=str(total_records))
        self.selected_file_card.configure(
            text=self.selected_csv.name if self.selected_csv else "None"
        )

    def choose_csv(self):
        path = filedialog.askopenfilename(
            title="Choose Shopify CSV",
            filetypes=[("CSV files", "*.csv")]
        )
        if path:
            self.selected_csv = Path(path)
            self.selected_file_card.configure(text=self.selected_csv.name)
            self.status.configure(
                text=f"Selected: {self.selected_csv.name}",
                text_color=PURPLE
            )

    def process_csv(self):
        if not self.selected_csv:
            self.choose_csv()
        if not self.selected_csv:
            return
        try:
            parsed = parse_shopify_orders(self.selected_csv, MASTER)
            result = merge_parsed_products(MASTER, parsed)
            review_count = (
                int(parsed["Needs Review"].astype(str).str.strip().ne("").sum())
                if not parsed.empty else 0
            )
            self.review_card.configure(text=str(review_count))
            self.total_records_card.configure(text=str(result["final_rows"]))
            self.status.configure(
                text=f"Processed {len(parsed)} product lines. Product Master now has {result['final_rows']} unique records.",
                text_color=SUCCESS
            )
            messagebox.showinfo(
                "Shopify Import Complete",
                f"Parsed product lines: {len(parsed)}\n"
                f"Unique Product Master records: {result['final_rows']}\n"
                f"Lines needing review: {review_count}\n\n"
                "Existing vendor and decoration assignments were preserved."
            )
        except Exception as error:
            messagebox.showerror("Unable to Process CSV", str(error))

    def open_product_master(self):
        editor = PROJECT / "modules" / "product_master_editor.py"
        try:
            subprocess.Popen([sys.executable, str(editor)], cwd=str(PROJECT))
            self.status.configure(text="Product Master opened.", text_color=PURPLE)
        except Exception as error:
            messagebox.showerror("Unable to Open Product Master", str(error))

    def generate_purchase_orders(self):
        if not self.selected_csv:
            self.choose_csv()
        if not self.selected_csv:
            return
        try:
            result = generate_purchase_orders(
                self.selected_csv, MASTER, PROJECT / "reports"
            )
            self.status.configure(
                text=(
                    f"Created {result['routes']} purchase-order route(s). "
                    f"{result['review_lines']} line(s) need review."
                ),
                text_color=SUCCESS,
            )
            messagebox.showinfo(
                "Purchase Orders Generated",
                f"Ready quantity: {result['ready_quantity']}\n"
                f"Purchase-order routes: {result['routes']}\n"
                f"PDF purchase orders: {len(result.get('pdf_files', []))}\n"
                f"Lines needing review: {result['review_lines']}\n\n"
                f"Saved in:\n{result['output_dir']}",
            )
            subprocess.run(["open", str(result["output_dir"])], check=False)
        except Exception as error:
            messagebox.showerror("Unable to Generate Purchase Orders", str(error))

    def clean_catalog(self):
        try:
            result = clean_product_master(MASTER, make_backup=True)
            self.total_records_card.configure(text=str(result["after"]))
            self.status.configure(
                text=f"Catalog cleaned. Removed {result['removed']} duplicate record(s).",
                text_color=SUCCESS
            )
            backup_text = str(result["backup"]) if result["backup"] else "No backup was needed."
            messagebox.showinfo(
                "Catalog Cleanup Complete",
                f"Rows before: {result['before']}\n"
                f"Rows after: {result['after']}\n"
                f"Duplicates removed: {result['removed']}\n\n"
                f"Backup: {backup_text}"
            )
        except Exception as error:
            messagebox.showerror("Unable to Clean Catalog", str(error))


if __name__ == "__main__":
    DATA.mkdir(parents=True, exist_ok=True)
    OrchidEventManager().mainloop()
