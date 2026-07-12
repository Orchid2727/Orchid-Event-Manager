from __future__ import annotations

from pathlib import Path
import subprocess
import sys
import customtkinter as ctk
from tkinter import filedialog, messagebox

from modules.shopify_parser import parse_shopify_orders
from modules.catalog_manager import clean_product_master, merge_parsed_products

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


class OrchidEventManager(ctk.CTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("light")
        self.title("Orchid Event Manager v1.1")
        self.geometry("900x660")
        self.minsize(820, 600)
        self.configure(fg_color=BG)
        self.selected_csv: Path | None = None
        self.build_ui()

    def build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        header = ctk.CTkFrame(self, fg_color="white", corner_radius=0)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(header, text="✿", text_color=PURPLE,
                     font=ctk.CTkFont(size=52, weight="bold")).grid(row=0, column=0, rowspan=2, padx=(30,12), pady=18)
        ctk.CTkLabel(header, text="Orchid Event Manager",
                     text_color=TEXT, font=ctk.CTkFont(size=31, weight="bold"), anchor="w").grid(row=0, column=1, sticky="sw", pady=(18,0))
        ctk.CTkLabel(header, text="Shopify order processing and product catalog management",
                     text_color=MUTED, font=ctk.CTkFont(size=14), anchor="w").grid(row=1, column=1, sticky="nw", pady=(2,18))
        ctk.CTkLabel(header, text="v1.1", text_color=PURPLE,
                     font=ctk.CTkFont(size=14, weight="bold")).grid(row=0, column=2, rowspan=2, padx=30)

        self.status = ctk.CTkLabel(self, text="Ready", text_color=MUTED, font=ctk.CTkFont(size=13))
        self.status.grid(row=1, column=0, sticky="ew", padx=28, pady=(16,4))

        cards = ctk.CTkFrame(self, fg_color="transparent")
        cards.grid(row=2, column=0, sticky="nsew", padx=28, pady=14)
        cards.grid_columnconfigure((0,1), weight=1)
        cards.grid_rowconfigure((0,1), weight=1)

        self.card(cards, 0, 0, "1", "Import Shopify Orders",
                  "Choose a Shopify CSV, parse products, and add new styles and colors to the permanent Product Master.",
                  "Choose CSV", self.choose_csv)
        self.card(cards, 0, 1, "2", "Process Selected CSV",
                  "Update the Product Master while preserving all vendor, decoration type, and decoration color assignments.",
                  "Process Orders", self.process_csv)
        self.card(cards, 1, 0, "3", "Product Master",
                  "Open the polished Style Mode editor to complete vendor and decoration information.",
                  "Open Product Master", self.open_product_master)
        self.card(cards, 1, 1, "4", "Catalog Maintenance",
                  "Create a backup, merge duplicate style/color records, and standardize product names safely.",
                  "Clean & Backup", self.clean_catalog)

        footer = ctk.CTkFrame(self, fg_color=PURPLE_DARK, corner_radius=0, height=62)
        footer.grid(row=3, column=0, sticky="ew")
        footer.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(footer, text="✿  O R C H I D", text_color="white",
                     font=ctk.CTkFont(size=15, weight="bold")).grid(row=0, column=0, padx=28, pady=18)
        ctk.CTkLabel(footer, text="Product Master • Shopify Import • Automatic Backups",
                     text_color="white", font=ctk.CTkFont(size=12)).grid(row=0, column=1, pady=18)
        ctk.CTkButton(footer, text="Exit", width=90, fg_color="transparent", border_width=1,
                      border_color="white", hover_color=PURPLE, command=self.destroy).grid(row=0, column=2, padx=28, pady=12)

    def card(self, parent, row, col, number, title, body, button_text, command):
        frame = ctk.CTkFrame(parent, fg_color="white", border_width=1, border_color=BORDER, corner_radius=14)
        frame.grid(row=row, column=col, sticky="nsew", padx=9, pady=9)
        frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(frame, text=number, width=38, height=38, corner_radius=19,
                     fg_color=PURPLE, text_color="white", font=ctk.CTkFont(size=15, weight="bold")).grid(row=0, column=0, padx=(20,12), pady=(20,8))
        ctk.CTkLabel(frame, text=title, text_color=TEXT,
                     font=ctk.CTkFont(size=18, weight="bold"), anchor="w").grid(row=0, column=1, sticky="ew", padx=(0,20), pady=(20,8))
        ctk.CTkLabel(frame, text=body, text_color=MUTED, wraplength=310,
                     justify="left", anchor="nw", font=ctk.CTkFont(size=12)).grid(row=1, column=0, columnspan=2, sticky="nsew", padx=20, pady=8)
        ctk.CTkButton(frame, text=button_text, height=42, fg_color=PURPLE,
                      hover_color=PURPLE_DARK, font=ctk.CTkFont(size=13, weight="bold"),
                      command=command).grid(row=2, column=0, columnspan=2, sticky="ew", padx=20, pady=(10,20))

    def choose_csv(self):
        path = filedialog.askopenfilename(title="Choose Shopify CSV", filetypes=[("CSV files", "*.csv")])
        if path:
            self.selected_csv = Path(path)
            self.status.configure(text=f"Selected: {self.selected_csv.name}", text_color=PURPLE)

    def process_csv(self):
        if not self.selected_csv:
            self.choose_csv()
        if not self.selected_csv:
            return
        try:
            parsed = parse_shopify_orders(self.selected_csv, MASTER)
            result = merge_parsed_products(MASTER, parsed)
            review_count = int(parsed["Needs Review"].astype(str).str.strip().ne("").sum()) if not parsed.empty else 0
            self.status.configure(text=f"Processed {len(parsed)} product lines. Product Master now has {result['final_rows']} unique records.", text_color="#2D7A46")
            messagebox.showinfo("Shopify Import Complete",
                                f"Parsed product lines: {len(parsed)}\n"
                                f"Unique Product Master records: {result['final_rows']}\n"
                                f"Lines needing review: {review_count}\n\n"
                                "Existing vendor and decoration assignments were preserved.")
        except Exception as error:
            messagebox.showerror("Unable to Process CSV", str(error))

    def open_product_master(self):
        editor = PROJECT / "modules" / "product_master_editor.py"
        try:
            subprocess.Popen([sys.executable, str(editor)], cwd=str(PROJECT))
            self.status.configure(text="Product Master opened.", text_color=PURPLE)
        except Exception as error:
            messagebox.showerror("Unable to Open Product Master", str(error))

    def clean_catalog(self):
        try:
            result = clean_product_master(MASTER, make_backup=True)
            self.status.configure(text=f"Catalog cleaned. Removed {result['removed']} duplicate record(s).", text_color="#2D7A46")
            backup_text = str(result["backup"]) if result["backup"] else "No backup was needed."
            messagebox.showinfo("Catalog Cleanup Complete",
                                f"Rows before: {result['before']}\n"
                                f"Rows after: {result['after']}\n"
                                f"Duplicates removed: {result['removed']}\n\n"
                                f"Backup: {backup_text}")
        except Exception as error:
            messagebox.showerror("Unable to Clean Catalog", str(error))


if __name__ == "__main__":
    DATA.mkdir(parents=True, exist_ok=True)
    OrchidEventManager().mainloop()
