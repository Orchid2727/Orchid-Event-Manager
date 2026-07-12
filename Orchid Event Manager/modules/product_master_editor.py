import customtkinter as ctk
from pathlib import Path
from tkinter import messagebox
import re
import pandas as pd

PROJECT_FOLDER = Path(__file__).resolve().parent.parent
PRODUCT_MASTER_PATH = PROJECT_FOLDER / 'data' / 'product_master.csv'

COLUMNS = ['Product Name','Style Number','Garment Color','Vendor','Decoration Type','Decoration Color']
WORKBOOK_VENDORS = ['Burnside','Cutter & Buck','Outdoor Cap','Richardson','S&S Activewear','SanMar','Tru-Spec','VF','Wrangler']
DECORATION_TYPES = ['', 'Embroidery', 'Screen Print', 'None']
DECORATION_COLORS = ['', 'White', 'Black', 'Navy', 'Red', 'Royal', 'Gold', 'Silver', 'Gray']

PURPLE = '#5B2AA8'
PURPLE_DARK = '#3D176F'
PURPLE_LIGHT = '#F2ECFB'
PURPLE_BORDER = '#D7C8EE'
TEXT_DARK = '#201A2D'
TEXT_MUTED = '#6F667A'
CARD_BG = '#FFFFFF'
WINDOW_BG = '#F7F5FA'
ROW_ALT = '#FAF7FE'
SUCCESS = '#2D7A46'


def clean_text(value):
    if pd.isna(value):
        return ''
    return str(value).strip()


def normalize_space(value):
    return re.sub(r'\s+', ' ', clean_text(value)).strip()


def normalize_style(value):
    return re.sub(r'\s+', '', clean_text(value)).upper()


def normalize_color(value):
    return normalize_space(value).casefold()


def canonical_product_name(value):
    name = normalize_space(value)
    # Old imports sometimes left punctuation immediately after the style number.
    # Removing trailing separators lets names such as "Knit Cap" and "Knit Cap."
    # merge into the same catalog item.
    name = re.sub(r'[\s\.\-–—:|/]+$', '', name).strip()
    return name


def first_nonblank(values):
    for value in values:
        value = clean_text(value)
        if value:
            return value
    return ''


def best_product_name(values):
    cleaned = [canonical_product_name(value) for value in values]
    cleaned = [value for value in cleaned if value]
    if not cleaned:
        return ''
    # Prefer the most common cleaned spelling; use the shortest spelling as a tie-breaker.
    counts = {}
    for value in cleaned:
        counts[value] = counts.get(value, 0) + 1
    return sorted(counts, key=lambda value: (-counts[value], len(value), value.casefold()))[0]


def catalog_key(row):
    style = normalize_style(row['Style Number'])
    color = normalize_color(row['Garment Color'])
    if style:
        # Style number + garment color is the permanent catalog identity.
        # Product-name punctuation must never create a duplicate record.
        return f'style:{style}|color:{color}'
    return f"product:{canonical_product_name(row['Product Name']).casefold()}|color:{color}"


def style_group_key(row):
    style = normalize_style(row['Style Number'])
    if style:
        return f'style:{style}'
    return f"product:{canonical_product_name(row['Product Name']).casefold()}"


def clean_and_deduplicate_master(master):
    for column in COLUMNS:
        if column not in master.columns:
            master[column] = ''

    master = master[COLUMNS].fillna('').copy()
    for column in COLUMNS:
        master[column] = master[column].apply(clean_text)

    master['Product Name'] = master['Product Name'].apply(canonical_product_name)
    master['Style Number'] = master['Style Number'].apply(normalize_style)
    master['Garment Color'] = master['Garment Color'].apply(normalize_space)

    if master.empty:
        return pd.DataFrame(columns=COLUMNS)

    master['_key'] = master.apply(catalog_key, axis=1)
    rows = []

    for _, group in master.groupby('_key', sort=False, dropna=False):
        rows.append({
            'Product Name': best_product_name(group['Product Name']),
            'Style Number': first_nonblank(group['Style Number']),
            'Garment Color': first_nonblank(group['Garment Color']),
            'Vendor': first_nonblank(group['Vendor']),
            'Decoration Type': first_nonblank(group['Decoration Type']),
            'Decoration Color': first_nonblank(group['Decoration Color']),
        })

    result = pd.DataFrame(rows, columns=COLUMNS)
    return result.sort_values(
        by=['Style Number', 'Product Name', 'Garment Color'],
        key=lambda series: series.astype(str).str.casefold(),
    ).reset_index(drop=True)


def _data_signature(master):
    comparable = master[COLUMNS].fillna('').astype(str).copy()
    for column in COLUMNS:
        comparable[column] = comparable[column].map(clean_text)
    return comparable.to_csv(index=False)


def create_cleanup_backup():
    from datetime import datetime
    import shutil

    backup_folder = PRODUCT_MASTER_PATH.parent / 'backups'
    backup_folder.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    backup_path = backup_folder / f'product_master_before_v1_1_cleanup_{timestamp}.csv'
    shutil.copy2(PRODUCT_MASTER_PATH, backup_path)
    return backup_path


def load_product_master():
    PRODUCT_MASTER_PATH.parent.mkdir(parents=True, exist_ok=True)
    if not PRODUCT_MASTER_PATH.exists():
        empty = pd.DataFrame(columns=COLUMNS)
        empty.to_csv(PRODUCT_MASTER_PATH, index=False)
        return empty

    try:
        master = pd.read_csv(PRODUCT_MASTER_PATH, dtype=str).fillna('')
    except pd.errors.EmptyDataError:
        master = pd.DataFrame(columns=COLUMNS)
    except Exception as error:
        messagebox.showerror('Unable to Open Product Master', str(error))
        return pd.DataFrame(columns=COLUMNS)

    for column in COLUMNS:
        if column not in master.columns:
            master[column] = ''
    original = master[COLUMNS].copy()
    cleaned = clean_and_deduplicate_master(original)

    if _data_signature(original) != _data_signature(cleaned):
        backup_path = create_cleanup_backup()
        cleaned.to_csv(PRODUCT_MASTER_PATH, index=False)
        messagebox.showinfo(
            'Product Master v1.1 Cleanup Complete',
            f'The catalog was cleaned from {len(original)} rows to {len(cleaned)} unique product/color records.\n\n'
            f'Your original file was backed up here:\n{backup_path}',
        )
    else:
        cleaned.to_csv(PRODUCT_MASTER_PATH, index=False)

    return cleaned


def save_product_master(master):
    try:
        cleaned = clean_and_deduplicate_master(master)
        cleaned.to_csv(PRODUCT_MASTER_PATH, index=False)
        return cleaned
    except Exception as error:
        messagebox.showerror('Unable to Save Product Master', str(error))
        return None


class ProductMasterV2(ctk.CTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode('light')
        ctk.set_default_color_theme('blue')
        self.title('Orchid Event Manager - Product Master')
        self.geometry('1220x900')
        self.minsize(1080, 760)
        self.configure(fg_color=WINDOW_BG)

        self.master_data = load_product_master()
        self.vendor_options = list(WORKBOOK_VENDORS)
        for vendor in self.master_data.get('Vendor', pd.Series(dtype=str)):
            vendor = clean_text(vendor)
            if vendor and vendor not in self.vendor_options:
                self.vendor_options.append(vendor)

        self.all_style_keys = []
        self.filtered_style_keys = []
        self.style_position = 0
        self.color_controls = []

        self.search_var = ctk.StringVar()
        self.incomplete_only_var = ctk.BooleanVar(value=True)
        self.vendor_var = ctk.StringVar()
        self.decoration_type_var = ctk.StringVar()

        self.build_style_list()
        self.build_ui()
        self.apply_filters()

    def build_style_list(self):
        self.all_style_keys = []
        seen = set()
        for _, row in self.master_data.iterrows():
            key = style_group_key(row)
            if key not in seen:
                seen.add(key)
                self.all_style_keys.append(key)

    def get_style_rows(self, style_key):
        if self.master_data.empty:
            return self.master_data.copy()
        mask = self.master_data.apply(lambda row: style_group_key(row) == style_key, axis=1)
        return self.master_data.loc[mask].copy()

    def style_is_complete(self, rows):
        if rows.empty:
            return False
        vendor_complete = rows['Vendor'].astype(str).str.strip().ne('').all()
        decoration_complete = rows['Decoration Type'].astype(str).str.strip().ne('').all()
        if not (vendor_complete and decoration_complete):
            return False
        decorated_rows = rows[rows['Decoration Type'].astype(str).str.strip().str.lower().ne('none')]
        if decorated_rows.empty:
            return True
        return decorated_rows['Decoration Color'].astype(str).str.strip().ne('').all()

    def completed_style_count(self):
        return sum(1 for key in self.all_style_keys if self.style_is_complete(self.get_style_rows(key)))

    def build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(4, weight=1)
        self.build_header()
        self.build_filter_bar()
        self.build_style_banner()
        self.build_content_area()
        self.build_action_bar()
        self.build_footer()
        self.search_var.trace_add('write', lambda *_: self.apply_filters())

    def build_header(self):
        header = ctk.CTkFrame(self, fg_color=CARD_BG, corner_radius=0, border_width=0, height=105)
        header.grid(row=0, column=0, sticky='ew')
        header.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(header, text='✿', text_color=PURPLE, font=ctk.CTkFont(size=48, weight='bold'), width=70).grid(row=0, column=0, rowspan=2, padx=(28,8), pady=18)
        ctk.CTkLabel(header, text='Product Master - Style Mode', text_color=TEXT_DARK, font=ctk.CTkFont(size=30, weight='bold'), anchor='w').grid(row=0, column=1, sticky='sw', pady=(20,0))
        ctk.CTkLabel(header, text='Assign vendor and decoration type once, then review each garment color.', text_color=TEXT_MUTED, font=ctk.CTkFont(size=14), anchor='w').grid(row=1, column=1, sticky='nw', pady=(2,18))

        progress_card = ctk.CTkFrame(header, fg_color=CARD_BG, border_width=1, border_color=PURPLE_BORDER, corner_radius=12, width=360, height=66)
        progress_card.grid(row=0, column=2, rowspan=2, padx=28, pady=18, sticky='e')
        progress_card.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(progress_card, text='◉', text_color=PURPLE, font=ctk.CTkFont(size=22, weight='bold')).grid(row=0, column=0, rowspan=2, padx=(16,10), pady=12)
        ctk.CTkLabel(progress_card, text='Progress', text_color=TEXT_DARK, font=ctk.CTkFont(size=13, weight='bold'), anchor='w').grid(row=0, column=1, sticky='sw', pady=(10,0))
        self.progress_text = ctk.CTkLabel(progress_card, text='', text_color=TEXT_MUTED, font=ctk.CTkFont(size=12), anchor='w')
        self.progress_text.grid(row=1, column=1, sticky='nw', pady=(0,10))
        self.progress_bar = ctk.CTkProgressBar(progress_card, width=120, progress_color=PURPLE, fg_color='#EDE8F3')
        self.progress_bar.grid(row=0, column=2, rowspan=2, padx=(12,16), pady=22)
        self.progress_bar.set(0)

    def build_filter_bar(self):
        bar = ctk.CTkFrame(self, fg_color=CARD_BG, border_width=1, border_color=PURPLE_BORDER, corner_radius=12)
        bar.grid(row=1, column=0, sticky='ew', padx=20, pady=(14,8))
        bar.grid_columnconfigure(0, weight=1)
        self.search_entry = ctk.CTkEntry(bar, textvariable=self.search_var, placeholder_text='Search styles by number, product name, vendor, or color...', height=42, border_color=PURPLE_BORDER, fg_color='#FFFFFF', text_color=TEXT_DARK, placeholder_text_color='#9A92A4')
        self.search_entry.grid(row=0, column=0, sticky='ew', padx=(18,12), pady=14)
        ctk.CTkCheckBox(bar, text='Incomplete Only', variable=self.incomplete_only_var, command=self.apply_filters, fg_color=PURPLE, hover_color=PURPLE_DARK, border_color=PURPLE, text_color=TEXT_DARK).grid(row=0, column=1, padx=10, pady=14)
        ctk.CTkButton(bar, text='Clear Filters', command=self.clear_filters, width=130, height=42, fg_color='transparent', hover_color=PURPLE_LIGHT, border_width=1, border_color=PURPLE, text_color=PURPLE).grid(row=0, column=2, padx=(10,18), pady=14)

    def build_style_banner(self):
        banner = ctk.CTkFrame(self, fg_color=CARD_BG, border_width=1, border_color=PURPLE_BORDER, corner_radius=12)
        banner.grid(row=2, column=0, sticky='ew', padx=20, pady=8)
        banner.grid_columnconfigure(1, weight=1)
        ctk.CTkButton(banner, text='◀  Previous Style', command=self.previous_style, width=155, height=40, fg_color='transparent', hover_color=PURPLE_LIGHT, border_width=1, border_color=PURPLE, text_color=PURPLE).grid(row=0, column=0, rowspan=3, padx=24, pady=18)
        self.style_count_label = ctk.CTkLabel(banner, text='', text_color=TEXT_DARK, font=ctk.CTkFont(size=13, weight='bold'))
        self.style_count_label.grid(row=0, column=1, pady=(14,0))
        self.product_name_label = ctk.CTkLabel(banner, text='', text_color=TEXT_DARK, font=ctk.CTkFont(size=24, weight='bold'), wraplength=700)
        self.product_name_label.grid(row=1, column=1, pady=(2,0))
        self.style_number_label = ctk.CTkLabel(banner, text='', text_color=TEXT_MUTED, font=ctk.CTkFont(size=14))
        self.style_number_label.grid(row=2, column=1, pady=(0,14))
        ctk.CTkButton(banner, text='Next Style  ▶', command=self.next_style, width=155, height=40, fg_color='transparent', hover_color=PURPLE_LIGHT, border_width=1, border_color=PURPLE, text_color=PURPLE).grid(row=0, column=2, rowspan=3, padx=24, pady=18)

    def build_content_area(self):
        content = ctk.CTkFrame(self, fg_color='transparent')
        content.grid(row=4, column=0, sticky='nsew', padx=20, pady=8)
        content.grid_columnconfigure(0, weight=0)
        content.grid_columnconfigure(1, weight=1)
        content.grid_rowconfigure(0, weight=1)
        self.build_left_panel(content)
        self.build_color_panel(content)

    def build_left_panel(self, parent):
        left = ctk.CTkFrame(parent, fg_color='transparent', width=320)
        left.grid(row=0, column=0, sticky='nsw', padx=(0,10))
        left.grid_columnconfigure(0, weight=1)

        info = ctk.CTkFrame(left, fg_color=CARD_BG, border_width=1, border_color=PURPLE_BORDER, corner_radius=12, width=320)
        info.grid(row=0, column=0, sticky='ew')
        info.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(info, text='①  Style Information', text_color=PURPLE, font=ctk.CTkFont(size=17, weight='bold'), anchor='w').grid(row=0, column=0, sticky='ew', padx=20, pady=(18,12))
        ctk.CTkLabel(info, text='Vendor *', text_color=TEXT_DARK, font=ctk.CTkFont(size=13, weight='bold'), anchor='w').grid(row=1, column=0, sticky='ew', padx=20, pady=(4,5))
        self.vendor_combo = ctk.CTkComboBox(info, variable=self.vendor_var, values=self.vendor_options, height=40, border_color=PURPLE_BORDER, button_color='#EAE3F5', button_hover_color='#DED2EF', fg_color='#FFFFFF', text_color=TEXT_DARK, dropdown_fg_color='#FFFFFF', dropdown_text_color=TEXT_DARK)
        self.vendor_combo.grid(row=2, column=0, sticky='ew', padx=20)
        ctk.CTkLabel(info, text='Select the primary vendor for this style.', text_color=TEXT_MUTED, font=ctk.CTkFont(size=11), anchor='w').grid(row=3, column=0, sticky='ew', padx=20, pady=(5,14))
        ctk.CTkLabel(info, text='Decoration Type *', text_color=TEXT_DARK, font=ctk.CTkFont(size=13, weight='bold'), anchor='w').grid(row=4, column=0, sticky='ew', padx=20, pady=(2,5))
        self.decoration_combo = ctk.CTkComboBox(info, variable=self.decoration_type_var, values=DECORATION_TYPES, height=40, border_color=PURPLE_BORDER, button_color='#EAE3F5', button_hover_color='#DED2EF', fg_color='#FFFFFF', text_color=TEXT_DARK, dropdown_fg_color='#FFFFFF', dropdown_text_color=TEXT_DARK)
        self.decoration_combo.grid(row=5, column=0, sticky='ew', padx=20)
        ctk.CTkLabel(info, text='Select the decoration type for this style.', text_color=TEXT_MUTED, font=ctk.CTkFont(size=11), anchor='w').grid(row=6, column=0, sticky='ew', padx=20, pady=(5,18))

        tips = ctk.CTkFrame(left, fg_color=PURPLE_LIGHT, border_width=1, border_color=PURPLE_BORDER, corner_radius=12, width=320)
        tips.grid(row=1, column=0, sticky='ew', pady=(12,0))
        ctk.CTkLabel(tips, text='💡  Tips', text_color=PURPLE, font=ctk.CTkFont(size=16, weight='bold'), anchor='w').pack(fill='x', padx=20, pady=(16,8))
        ctk.CTkLabel(tips, text='Use Incomplete Only to focus on styles that still need vendor, decoration type, or color assignments.\n\nSave your progress often. You can return to any style at any time.', text_color=TEXT_DARK, justify='left', anchor='w', wraplength=270, font=ctk.CTkFont(size=12)).pack(fill='x', padx=20, pady=(0,18))

    def build_color_panel(self, parent):
        panel = ctk.CTkFrame(parent, fg_color=CARD_BG, border_width=1, border_color=PURPLE_BORDER, corner_radius=12)
        panel.grid(row=0, column=1, sticky='nsew', padx=(10,0))
        panel.grid_columnconfigure(0, weight=1)
        panel.grid_rowconfigure(3, weight=1)
        ctk.CTkLabel(panel, text='②  Garment Colors and Decoration Colors', text_color=PURPLE, font=ctk.CTkFont(size=17, weight='bold'), anchor='w').grid(row=0, column=0, sticky='ew', padx=22, pady=(18,6))
        ctk.CTkLabel(panel, text='For each garment color, specify the decoration color (thread or ink) to be used.', text_color=TEXT_MUTED, font=ctk.CTkFont(size=12), anchor='w').grid(row=1, column=0, sticky='ew', padx=22, pady=(0,12))

        table_header = ctk.CTkFrame(panel, fg_color=PURPLE_LIGHT, border_width=1, border_color=PURPLE_BORDER, corner_radius=8, height=44)
        table_header.grid(row=2, column=0, sticky='ew', padx=22)
        table_header.grid_columnconfigure(1, weight=1)
        table_header.grid_columnconfigure(2, weight=1)
        ctk.CTkLabel(table_header, text='#', text_color=PURPLE, font=ctk.CTkFont(size=12, weight='bold'), width=44).grid(row=0, column=0, padx=(8,0), pady=10)
        ctk.CTkLabel(table_header, text='Garment Color', text_color=PURPLE, font=ctk.CTkFont(size=12, weight='bold'), anchor='w').grid(row=0, column=1, sticky='ew', padx=8, pady=10)
        ctk.CTkLabel(table_header, text='Decoration Color (Thread / Ink)', text_color=PURPLE, font=ctk.CTkFont(size=12, weight='bold'), anchor='w').grid(row=0, column=2, sticky='ew', padx=8, pady=10)

        self.colors_scroll = ctk.CTkScrollableFrame(panel, fg_color='#FFFFFF', corner_radius=8, border_width=1, border_color=PURPLE_BORDER)
        self.colors_scroll.grid(row=3, column=0, sticky='nsew', padx=22, pady=(6,18))
        self.colors_scroll.grid_columnconfigure(1, weight=1)
        self.colors_scroll.grid_columnconfigure(2, weight=1)

    def build_action_bar(self):
        actions = ctk.CTkFrame(self, fg_color='transparent')
        actions.grid(row=5, column=0, sticky='ew', padx=20, pady=(4,10))
        for column in range(4):
            actions.grid_columnconfigure(column, weight=1)
        kwargs = {'height':50,'fg_color':PURPLE,'hover_color':PURPLE_DARK,'font':ctk.CTkFont(size=14, weight='bold')}
        ctk.CTkButton(actions, text='◀◀  Previous Style', command=self.previous_style, **kwargs).grid(row=0, column=0, sticky='ew', padx=(0,7))
        ctk.CTkButton(actions, text='▣  Save Style', command=self.save_current_style, **kwargs).grid(row=0, column=1, sticky='ew', padx=7)
        ctk.CTkButton(actions, text='▣  Save & Next Style', command=self.save_and_next_style, **kwargs).grid(row=0, column=2, sticky='ew', padx=7)
        ctk.CTkButton(actions, text='Next Style  ▶▶', command=self.next_style, **kwargs).grid(row=0, column=3, sticky='ew', padx=(7,0))

    def build_footer(self):
        footer = ctk.CTkFrame(self, fg_color=PURPLE_DARK, corner_radius=0, height=72)
        footer.grid(row=6, column=0, sticky='ew')
        footer.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(footer, text='✿  O R C H I D\n     E V E N T  M A N A G E R', text_color='#FFFFFF', justify='left', font=ctk.CTkFont(size=14, weight='bold')).grid(row=0, column=0, padx=28, pady=14, sticky='w')
        ctk.CTkLabel(footer, text='Product Master - Style Mode\nKeep your product data accurate and consistent.', text_color='#FFFFFF', justify='center', font=ctk.CTkFont(size=12)).grid(row=0, column=1, pady=14)
        self.footer_record_label = ctk.CTkLabel(footer, text='', text_color='#FFFFFF', font=ctk.CTkFont(size=12))
        self.footer_record_label.grid(row=0, column=2, padx=28, pady=14, sticky='e')
        self.status_label = ctk.CTkLabel(self, text='', text_color=SUCCESS, font=ctk.CTkFont(size=12, weight='bold'))
        self.status_label.place(relx=0.5, rely=0.925, anchor='center')

    def clear_filters(self):
        self.search_var.set('')
        self.incomplete_only_var.set(False)
        self.apply_filters()

    def apply_filters(self):
        search_text = self.search_var.get().strip().lower()
        incomplete_only = self.incomplete_only_var.get()
        matched = []
        for style_key in self.all_style_keys:
            rows = self.get_style_rows(style_key)
            if incomplete_only and self.style_is_complete(rows):
                continue
            combined = ' '.join(clean_text(value) for value in rows[COLUMNS].to_numpy().flatten()).lower()
            if search_text and search_text not in combined:
                continue
            matched.append(style_key)
        self.filtered_style_keys = matched
        self.style_position = 0
        self.refresh_progress()
        self.show_current_style()

    def refresh_progress(self):
        total = len(self.all_style_keys)
        complete = self.completed_style_count()
        percent = (complete / total) if total else 0
        self.progress_text.configure(text=f'{complete} of {total} complete   {round(percent*100)}%')
        self.progress_bar.set(percent)

    def clear_color_rows(self):
        for widget in self.colors_scroll.winfo_children():
            widget.destroy()
        self.color_controls = []

    def show_current_style(self):
        self.clear_color_rows()
        self.status_label.configure(text='')
        if not self.filtered_style_keys:
            self.style_count_label.configure(text='No matching styles')
            self.product_name_label.configure(text='')
            self.style_number_label.configure(text='')
            self.footer_record_label.configure(text='No records')
            self.vendor_var.set('')
            self.decoration_type_var.set('')
            return

        self.style_position = max(0, min(self.style_position, len(self.filtered_style_keys)-1))
        style_key = self.filtered_style_keys[self.style_position]
        rows = self.get_style_rows(style_key)
        self.style_count_label.configure(text=f'Style {self.style_position+1} of {len(self.filtered_style_keys)}')
        self.footer_record_label.configure(text=f'Record {self.style_position+1} of {len(self.filtered_style_keys)}')
        product_names = [clean_text(value) for value in rows['Product Name'].drop_duplicates().tolist() if clean_text(value)]
        self.product_name_label.configure(text=' / '.join(product_names))
        style_number = first_nonblank(rows['Style Number'])
        self.style_number_label.configure(text=f"Style Number: {style_number or 'Not supplied'}")

        vendors = [clean_text(value) for value in rows['Vendor'].tolist() if clean_text(value)]
        unique_vendors = list(dict.fromkeys(vendors))
        self.vendor_var.set(unique_vendors[0] if len(unique_vendors)==1 else '')

        decorations = [clean_text(value) for value in rows['Decoration Type'].tolist() if clean_text(value)]
        unique_decorations = list(dict.fromkeys(decorations))
        self.decoration_type_var.set(unique_decorations[0] if len(unique_decorations)==1 else '')

        sorted_rows = rows.sort_values(by='Garment Color', key=lambda series: series.astype(str).str.lower())
        for display_index, (data_index, row) in enumerate(sorted_rows.iterrows(), start=1):
            row_color = ROW_ALT if display_index % 2 == 0 else '#FFFFFF'
            row_frame = ctk.CTkFrame(self.colors_scroll, fg_color=row_color, corner_radius=0, height=52)
            row_frame.grid(row=display_index-1, column=0, columnspan=3, sticky='ew', padx=0, pady=0)
            row_frame.grid_columnconfigure(1, weight=1)
            row_frame.grid_columnconfigure(2, weight=1)
            ctk.CTkLabel(row_frame, text=str(display_index), text_color='#FFFFFF', fg_color=PURPLE, width=28, height=28, corner_radius=14, font=ctk.CTkFont(size=11, weight='bold')).grid(row=0, column=0, padx=(12,10), pady=12)
            ctk.CTkLabel(row_frame, text=clean_text(row['Garment Color']) or 'No garment color', text_color=TEXT_DARK, font=ctk.CTkFont(size=12), anchor='w').grid(row=0, column=1, sticky='ew', padx=6, pady=10)
            color_var = ctk.StringVar(value=clean_text(row['Decoration Color']))
            ctk.CTkComboBox(row_frame, variable=color_var, values=DECORATION_COLORS, height=36, border_color=PURPLE_BORDER, button_color='#EAE3F5', button_hover_color='#DED2EF', fg_color='#FFFFFF', text_color=TEXT_DARK, dropdown_fg_color='#FFFFFF', dropdown_text_color=TEXT_DARK).grid(row=0, column=2, sticky='ew', padx=(8,14), pady=8)
            self.color_controls.append({'index':data_index, 'decoration_color_var':color_var})

    def save_current_style(self):
        if not self.filtered_style_keys:
            return False
        style_key = self.filtered_style_keys[self.style_position]
        rows = self.get_style_rows(style_key)
        vendor = self.vendor_var.get().strip()
        decoration_type = self.decoration_type_var.get().strip()
        for index in rows.index:
            self.master_data.at[index, 'Vendor'] = vendor
            self.master_data.at[index, 'Decoration Type'] = decoration_type
        for item in self.color_controls:
            self.master_data.at[item['index'], 'Decoration Color'] = item['decoration_color_var'].get().strip()
        saved = save_product_master(self.master_data)
        if saved is None:
            return False
        self.master_data = saved
        if vendor and vendor not in self.vendor_options:
            self.vendor_options.append(vendor)
            self.vendor_combo.configure(values=self.vendor_options)
        self.build_style_list()
        self.refresh_progress()
        self.status_label.configure(text='Style saved successfully.')
        return True

    def save_and_next_style(self):
        current_key = self.filtered_style_keys[self.style_position] if self.filtered_style_keys else None
        if not self.save_current_style():
            return
        self.apply_filters()
        if current_key in self.filtered_style_keys:
            current_index = self.filtered_style_keys.index(current_key)
            self.style_position = min(current_index+1, len(self.filtered_style_keys)-1)
        else:
            self.style_position = min(self.style_position, max(len(self.filtered_style_keys)-1, 0))
        self.show_current_style()

    def previous_style(self):
        if self.filtered_style_keys and self.style_position > 0:
            self.style_position -= 1
            self.show_current_style()

    def next_style(self):
        if self.filtered_style_keys and self.style_position < len(self.filtered_style_keys)-1:
            self.style_position += 1
            self.show_current_style()


if __name__ == '__main__':
    app = ProductMasterV2()
    app.mainloop()
