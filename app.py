import os
import subprocess
import sys
if sys.stdout.encoding.lower() != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

# ── Hide Console Terminal on Windows ───────────────────────────────────────────
def hide_console():
    import ctypes
    if os.name == 'nt':
        try:
            hwnd = ctypes.windll.kernel32.GetConsoleWindow()
            if hwnd != 0:
                ctypes.windll.user32.ShowWindow(hwnd, 0)
        except Exception:
            pass

hide_console()

import threading
import tkinter as tk
from tkinter import filedialog, scrolledtext
from tkinter import ttk
import tkinter.font as tkfont
from datetime import datetime
import unicodedata

# Try importing tkinterdnd2 for native drag-and-drop
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    _BASE = TkinterDnD.Tk
    _DND_AVAILABLE = True
except ImportError:
    _BASE = tk.Tk
    _DND_AVAILABLE = False

# Import logic module
from generate_contracts import generate, get_contract_preview, get_responsible_persons
from data_config import default_fields, default_pdf_engine, default_template_path, load_data_settings, save_data_settings

# ── Color palette (Luxury Gold & Black) ──────────────────────────────────────────
BG_DARK       = "#0a0a0a"  # Obsidian black
BG_CARD       = "#141414"  # Premium dark gray
BG_DROP       = "#1b1a17"  # Deep amber-tinted black for drop zone
BG_DROP_HOVER = "#26231d"  # Warm hover state for drop zone
ACCENT        = "#c5a059"  # Muted luxury gold
ACCENT_GLOW   = "#e5c17d"  # Warm metallic gold glow
SUCCESS       = "#22c55e"  # Emerald green for success
WARNING       = "#f59e0b"  # Amber warning
ERROR_CLR     = "#ef4444"  # Crimson red for errors
TEXT_PRIMARY  = "#f5f5f7"  # Warm off-white
TEXT_MUTED    = "#8e8e93"  # Cool elegant gray
BORDER        = "#2b251b"  # Sophisticated gold-tinted border
LOG_BG        = "#050505"  # Absolute obsidian black for console logs

# ── Fonts ──────────────────────────────────────────────────────────────────────
FONT_TITLE   = ("Segoe UI", 22, "bold")
FONT_SUB     = ("Segoe UI", 10)
FONT_LABEL   = ("Segoe UI", 11, "bold")
FONT_BTN     = ("Segoe UI", 11, "bold")
FONT_LOG     = ("Consolas", 9)
FONT_PATH    = ("Segoe UI", 9)
FONT_DROP    = ("Segoe UI", 13)
FONT_DROP_SM = ("Segoe UI", 9)


class App(_BASE):
    def __init__(self):
        super().__init__()

        self.title("Contract Generator")
        self.configure(bg=BG_DARK)
        self.resizable(True, True)
        self.minsize(860, 720)

        # Set window icon if icon.ico exists
        try:
            if getattr(sys, 'frozen', False):
                icon_path = os.path.join(sys._MEIPASS, "icon.ico")
            else:
                icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icon.ico")
            
            if os.path.exists(icon_path):
                self.iconbitmap(icon_path)
        except Exception:
            pass

        # Center window
        self.geometry("980x780")
        self.update_idletasks()
        x = (self.winfo_screenwidth() - 980) // 2
        y = (self.winfo_screenheight() - 780) // 2
        self.geometry(f"980x780+{x}+{y}")

        self.excel_path = tk.StringVar(value="")
        self.company_search = tk.StringVar(value="")
        self.preview_rows = []
        self.filtered_preview_rows = []
        self.selected_row_ids = set()
        self.data_settings = load_data_settings()
        self.data_fields = self.data_settings["fields"]
        self.template_path = tk.StringVar(value=self.data_settings["template_path"])
        self.pdf_engine = tk.StringVar(value=self.data_settings.get("pdf_engine", default_pdf_engine()))
        self._init_template_path()
        self.cancel_event = threading.Event()
        self.is_running  = False

        self._build_ui()
        self.company_search.trace_add("write", lambda *_: self._on_company_search_changed())
        self._setup_drag_drop()

    def _resolve_app_path(self, path: str) -> str:
        if not path:
            return ""
        if os.path.isabs(path):
            return path
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), path)

    def _init_template_path(self):
        configured = self.data_settings.get("template_path", default_template_path())
        default_path = default_template_path()
        chosen = configured or default_path
        if os.path.exists(self._resolve_app_path(chosen)):
            self.template_path.set(chosen)
            return
        if chosen != default_path and os.path.exists(self._resolve_app_path(default_path)):
            self.template_path.set(default_path)
            save_data_settings(self.data_fields, default_path)
            return
        self.template_path.set("")

    # ── UI Layout ──────────────────────────────────────────────────────────────
    def _build_ui(self):
        # ── Header ──
        header = tk.Frame(self, bg=BG_DARK, pady=20)
        header.pack(fill="x", padx=30)

        tk.Label(
            header, text="Tạo Hợp Đồng Nguyên Tắc",
            font=FONT_TITLE, bg=BG_DARK, fg=TEXT_PRIMARY
        ).pack(anchor="w")

        # ── Drop zone ──
        drop_frame = tk.Frame(self, bg=BG_DARK, padx=30)
        drop_frame.pack(fill="x", pady=(0, 10))

        self.template_drop_zone = tk.Frame(
            drop_frame, bg=BG_DROP,
            highlightbackground=BORDER, highlightthickness=2,
            cursor="hand2"
        )
        self.template_drop_zone.pack(side="left", fill="both", expand=True, ipady=1, padx=(0, 6))

        self.template_drop_icon = tk.Label(
            self.template_drop_zone, text="Mẫu HĐ", font=("Segoe UI", 11, "bold"),
            bg=BG_DROP, fg=ACCENT
        )
        self.template_drop_icon.pack(pady=(2, 0))

        self.template_drop_label = tk.Label(
            self.template_drop_zone,
            text="Kéo & Thả file .docx/.doc vào đây" if not self.template_path.get().strip() else os.path.basename(self.template_path.get()),
            font=FONT_SUB, bg=BG_DROP,
            fg=TEXT_PRIMARY if not self.template_path.get().strip() else SUCCESS
        )
        self.template_drop_label.pack()

        tk.Label(
            self.template_drop_zone, text="hoặc click để chọn file",
            font=FONT_DROP_SM, bg=BG_DROP, fg=TEXT_MUTED
        ).pack(pady=(0, 2))

        for widget in [self.template_drop_zone, self.template_drop_icon, self.template_drop_label]:
            widget.bind("<Button-1>", lambda e: self._browse_template())
            widget.bind("<Enter>",    lambda e: self._template_drop_hover(True))
            widget.bind("<Leave>",    lambda e: self._template_drop_hover(False))

        self.drop_zone = tk.Frame(
            drop_frame, bg=BG_DROP,
            highlightbackground=BORDER, highlightthickness=2,
            cursor="hand2"
        )
        self.drop_zone.pack(side="right", fill="both", expand=True, ipady=1, padx=(6, 0))

        self.drop_icon = tk.Label(
            self.drop_zone, text="Dữ liệu Excel", font=("Segoe UI", 14, "bold"),
            bg=BG_DROP, fg=ACCENT
        )
        self.drop_icon.config(font=("Segoe UI", 11, "bold"))
        self.drop_icon.pack(pady=(2, 0))

        self.drop_label = tk.Label(
            self.drop_zone,
            text="Kéo & Thả file .xlsx vào đây",
            font=FONT_SUB, bg=BG_DROP, fg=TEXT_PRIMARY
        )
        self.drop_label.pack()

        tk.Label(
            self.drop_zone, text="hoặc click để chọn file",
            font=FONT_DROP_SM, bg=BG_DROP, fg=TEXT_MUTED
        ).pack(pady=(0, 2))

        # Bind click events for all children
        for widget in [self.drop_zone, self.drop_icon, self.drop_label]:
            widget.bind("<Button-1>", lambda e: self._browse_file())
            widget.bind("<Enter>",    lambda e: self._drop_hover(True))
            widget.bind("<Leave>",    lambda e: self._drop_hover(False))

        # ── Selected file path display ──
        path_frame = tk.Frame(self, bg=BG_DARK, padx=30)
        path_frame.pack_forget()

        self.path_label = tk.Label(
            path_frame, textvariable=self.excel_path,
            font=FONT_PATH, bg=BG_DARK, fg=TEXT_MUTED,
            anchor="w", wraplength=760
        )
        self.path_label.pack_forget()

        # ── Responsible person selector ──
        self.selector_frame = tk.Frame(self, bg=BG_DARK, padx=30)
        # The sale filter is displayed in the preview header below.

        tk.Label(
            self.selector_frame, text="Sale phụ trách",
            font=FONT_SUB, bg=BG_DARK, fg=TEXT_MUTED
        ).pack(anchor="w", pady=(0, 4))

        self.responsible_person = tk.StringVar(value="")
        self.responsible_person.trace_add("write", lambda *_: self._on_responsible_person_changed())
        self.person_menu = None

        # ── Excel preview with row selection ──
        self.preview_frame = tk.Frame(self, bg=BG_DARK, padx=30)
        self.preview_frame.pack(fill="both", expand=False, pady=(0, 8))

        preview_header = tk.Frame(self.preview_frame, bg=BG_DARK)
        preview_header.pack(fill="x", pady=(0, 4))

        tk.Label(
            preview_header, text="Danh sách thông tin",
            font=FONT_LABEL, bg=BG_DARK, fg=TEXT_PRIMARY
        ).pack(side="left")

        self.preview_count_label = tk.Label(
            preview_header, text="Chưa có dữ liệu",
            font=FONT_DROP_SM, bg=BG_DARK, fg=TEXT_MUTED
        )
        self.preview_count_label.pack(side="left", padx=(10, 0))

        self.select_none_btn = tk.Button(
            preview_header,
            text="Bỏ chọn",
            font=FONT_DROP_SM,
            bg=BG_CARD, fg=TEXT_MUTED,
            activebackground=BORDER, activeforeground=TEXT_PRIMARY,
            relief="flat", bd=0,
            padx=10, pady=4,
            cursor="hand2",
            command=lambda: self._set_all_preview_rows(False)
        )
        self.select_none_btn.pack(side="left", padx=(14, 0))

        self.select_all_btn = tk.Button(
            preview_header,
            text="Chọn tất cả",
            font=FONT_DROP_SM,
            bg=BG_CARD, fg=TEXT_MUTED,
            activebackground=BORDER, activeforeground=TEXT_PRIMARY,
            relief="flat", bd=0,
            padx=10, pady=4,
            cursor="hand2",
            command=lambda: self._set_all_preview_rows(True)
        )
        self.select_all_btn.pack(side="left", padx=(8, 0))

        self.right_filter_frame = tk.Frame(preview_header, bg=BG_DARK)
        self.right_filter_frame.pack(side="right", padx=(0, 8))

        self.search_frame = tk.Frame(self.right_filter_frame, bg=BG_DARK)
        self.search_frame.pack_forget()

        tk.Label(
            self.search_frame, text="Tìm công ty",
            font=FONT_DROP_SM, bg=BG_DARK, fg=TEXT_MUTED
        ).pack(side="left", padx=(0, 8))

        self.search_box = tk.Frame(
            self.search_frame,
            bg=BG_CARD,
            highlightbackground=BORDER,
            highlightcolor=ACCENT,
            highlightthickness=1,
        )
        self.search_box.pack(side="left")

        self.search_entry = tk.Entry(
            self.search_box,
            textvariable=self.company_search,
            font=FONT_PATH,
            bg=BG_CARD,
            fg=TEXT_PRIMARY,
            insertbackground=ACCENT,
            relief="flat",
            bd=0,
            width=20,
        )
        self.search_entry.pack(side="left", ipady=4)

        self.clear_search_btn = tk.Button(
            self.search_box,
            text="x",
            font=("Segoe UI", 10, "bold"),
            bg=BG_CARD,
            fg=ACCENT_GLOW,
            activebackground=BORDER,
            activeforeground=TEXT_PRIMARY,
            relief="flat",
            bd=0,
            padx=6,
            pady=1,
            cursor="hand2",
            command=self._clear_company_search,
        )
        self.clear_search_btn.pack_forget()

        self.preview_filter_frame = tk.Frame(self.right_filter_frame, bg=BG_DARK)
        self.preview_filter_frame.pack(side="left", padx=(0, 12))

        tk.Label(
            self.preview_filter_frame, text="Sale phụ trách",
            font=FONT_DROP_SM, bg=BG_DARK, fg=TEXT_MUTED
        ).pack(side="left", padx=(0, 8))

        self.person_box = tk.Frame(
            self.preview_filter_frame,
            bg=BG_CARD,
            highlightbackground=BORDER,
            highlightcolor=ACCENT,
            highlightthickness=1,
        )
        self.person_box.pack(side="left")

        self._set_person_options(["Tất cả"])
        self.search_frame.pack(side="left")

        unused_select_all_btn = tk.Button(
            preview_header,
            text="Chọn tất cả",
            font=FONT_DROP_SM,
            bg=BG_CARD, fg=TEXT_MUTED,
            activebackground=BORDER, activeforeground=TEXT_PRIMARY,
            relief="flat", bd=0,
            padx=10, pady=4,
            cursor="hand2",
            command=lambda: self._set_all_preview_rows(True)
        )
        unused_select_all_btn.pack_forget()

        unused_select_none_btn = tk.Button(
            preview_header,
            text="Bỏ chọn",
            font=FONT_DROP_SM,
            bg=BG_CARD, fg=TEXT_MUTED,
            activebackground=BORDER, activeforeground=TEXT_PRIMARY,
            relief="flat", bd=0,
            padx=10, pady=4,
            cursor="hand2",
            command=lambda: self._set_all_preview_rows(False)
        )
        unused_select_none_btn.pack_forget()

        preview_table_frame = tk.Frame(
            self.preview_frame,
            bg=BG_CARD,
            highlightbackground=BORDER,
            highlightthickness=1
        )
        preview_table_frame.pack(fill="x")

        style = ttk.Style()
        style.theme_use("default")
        style.configure(
            "ContractPreview.Treeview",
            background=BG_CARD,
            foreground=TEXT_PRIMARY,
            fieldbackground=BG_CARD,
            rowheight=28,
            borderwidth=0,
            font=FONT_PATH,
        )
        style.configure(
            "ContractPreview.Treeview.Heading",
            background=BG_DROP,
            foreground=ACCENT_GLOW,
            relief="flat",
            font=FONT_DROP_SM,
        )
        style.map(
            "ContractPreview.Treeview.Heading",
            background=[("active", BG_DROP), ("pressed", BG_DROP)],
            foreground=[("active", ACCENT_GLOW), ("pressed", ACCENT_GLOW)],
            relief=[("active", "flat"), ("pressed", "flat")],
        )
        style.map("ContractPreview.Treeview", background=[("selected", "#2a2418")])

        self.preview_tree = ttk.Treeview(
            preview_table_frame,
            columns=self._preview_columns(),
            show="headings",
            height=6,
            style="ContractPreview.Treeview",
            selectmode="none"
        )
        self._configure_preview_columns()

        preview_scroll = ttk.Scrollbar(preview_table_frame, orient="vertical", command=self.preview_tree.yview)
        self.preview_tree.configure(yscrollcommand=preview_scroll.set)
        self.preview_tree.pack(side="left", fill="x", expand=True)
        preview_scroll.pack(side="right", fill="y")
        preview_x_scroll = ttk.Scrollbar(self.preview_frame, orient="horizontal", command=self.preview_tree.xview)
        self.preview_tree.configure(xscrollcommand=preview_x_scroll.set)
        preview_x_scroll.pack(fill="x")
        self.preview_tree.tag_configure("normal", background=BG_CARD, foreground=TEXT_PRIMARY)
        self.preview_tree.tag_configure("hover", background="#24211a", foreground=TEXT_PRIMARY)
        self.preview_tree.bind("<Button-1>", self._on_preview_click)
        self.preview_tree.bind("<Motion>", self._on_preview_motion)
        self.preview_tree.bind("<Leave>", self._on_preview_leave)

        # ── Run button ──
        btn_frame = tk.Frame(self, bg=BG_DARK, padx=30, pady=12)
        btn_frame.pack(fill="x")

        self.docx_btn = tk.Button(
            btn_frame,
            text="Tạo HĐ Docx",
            font=FONT_BTN,
            bg=ACCENT, fg="#0a0a0a",
            activebackground=ACCENT_GLOW, activeforeground="#0a0a0a",
            relief="flat", bd=0,
            padx=24, pady=10,
            cursor="hand2",
            command=lambda: self._run("docx")
        )
        self.docx_btn.pack(side="left")

        self.pdf_btn = tk.Button(
            btn_frame,
            text="Tạo HĐ PDF",
            font=FONT_BTN,
            bg=ACCENT, fg="#0a0a0a",
            activebackground=ACCENT_GLOW, activeforeground="#0a0a0a",
            relief="flat", bd=0,
            padx=24, pady=10,
            cursor="hand2",
            command=lambda: self._run("pdf")
        )
        self.pdf_btn.pack(side="left", padx=(10, 0))

        self.cancel_btn = tk.Button(
            btn_frame,
            text="Hủy",
            font=FONT_BTN,
            bg=BG_CARD, fg=TEXT_MUTED,
            activebackground=ERROR_CLR, activeforeground=TEXT_PRIMARY,
            relief="flat", bd=0,
            padx=18, pady=10,
            cursor="hand2",
            state="disabled",
            command=self._cancel_run
        )
        self.cancel_btn.pack(side="left", padx=(10, 0))

        self.clear_btn = tk.Button(
            btn_frame,
            text="Xoá Log",
            font=FONT_BTN,
            bg=BG_CARD, fg=TEXT_MUTED,
            activebackground=BORDER, activeforeground=TEXT_PRIMARY,
            relief="flat", bd=0,
            padx=18, pady=10,
            cursor="hand2",
            command=self._clear_log
        )
        self.clear_btn.pack(side="left", padx=(10, 0))

        self.settings_btn = tk.Button(
            btn_frame,
            text="Thiết lập",
            font=FONT_BTN,
            bg=BG_CARD, fg=TEXT_MUTED,
            activebackground=BORDER, activeforeground=TEXT_PRIMARY,
            relief="flat", bd=0,
            padx=18, pady=10,
            cursor="hand2",
            command=self._open_data_settings
        )
        self.settings_btn.pack(side="left", padx=(10, 0))

        self.open_btn = tk.Button(
            btn_frame,
            text="Mở thư mục Hợp Đồng",
            font=FONT_BTN,
            bg=BG_CARD, fg=TEXT_MUTED,
            activebackground=BORDER, activeforeground=TEXT_PRIMARY,
            relief="flat", bd=0,
            padx=18, pady=10,
            cursor="hand2",
            command=self._open_contracts
        )
        self.open_btn.pack(side="right")

        # ── Divider ──
        tk.Frame(self, bg=BORDER, height=1).pack(fill="x", padx=30, pady=(0, 0))

        # ── Log panel ──
        log_header = tk.Frame(self, bg=BG_DARK, padx=30, pady=8)
        log_header.pack(fill="x")
        tk.Label(
            log_header, text="Log",
            font=FONT_LABEL, bg=BG_DARK, fg=TEXT_PRIMARY
        ).pack(side="left")

        self.status_label = tk.Label(
            log_header, text="Sẵn sàng",
            font=FONT_DROP_SM, bg=BG_DARK, fg=TEXT_MUTED
        )
        self.status_label.pack(side="right")

        # ── Footer ──
        import webbrowser
        footer_frame = tk.Frame(self, bg=BG_DARK)
        footer_frame.pack(side="bottom", anchor="w", padx=30, pady=(8, 12))

        lbl_created_by = tk.Label(
            footer_frame, text="Tạo bởi",
            font=("Segoe UI", 9), bg=BG_DARK, fg=TEXT_MUTED
        )
        lbl_created_by.pack(side="left")

        lbl_link = tk.Label(
            footer_frame, text="mhqb365.com",
            font=("Segoe UI", 9), bg=BG_DARK, fg=TEXT_MUTED,
            cursor="hand2"
        )
        lbl_link.pack(side="left")

        lbl_link.bind("<Button-1>", lambda e: webbrowser.open("https://mhqb365.com"))
        lbl_link.bind("<Enter>",    lambda e: lbl_link.config(fg=ACCENT, font=("Segoe UI", 9, "underline")))
        lbl_link.bind("<Leave>",    lambda e: lbl_link.config(fg=TEXT_MUTED, font=("Segoe UI", 9)))

        log_frame = tk.Frame(self, bg=BG_DARK, padx=30)
        log_frame.pack(fill="both", expand=True, pady=(0, 0))

        self.log_box = scrolledtext.ScrolledText(
            log_frame,
            bg=LOG_BG, fg=TEXT_PRIMARY,
            font=FONT_LOG,
            relief="flat",
            bd=0,
            wrap="word",
            state="disabled",
            highlightbackground=BORDER,
            highlightthickness=1,
            insertbackground=ACCENT,
        )
        self.log_box.pack(fill="both", expand=True)

        # Log color tags
        self.log_box.tag_config("success", foreground=SUCCESS)
        self.log_box.tag_config("warning", foreground=WARNING)
        self.log_box.tag_config("error",   foreground=ERROR_CLR)
        self.log_box.tag_config("accent",  foreground=ACCENT_GLOW)
        self.log_box.tag_config("muted",   foreground=TEXT_MUTED)
        self.log_box.tag_config("time",    foreground="#4b5563")

        self._log("Kéo file .xlsx vào khung Dữ liệu Excel. Tạo HĐ Docx trước mới tạo được HĐ PDF.", "muted")

    # ── Drag-and-drop setup ────────────────────────────────────────────────────
    def _preview_columns(self):
        return ("checked", *[field["template_name"] for field in self.data_fields])

    def _configure_preview_columns(self):
        self.preview_tree["columns"] = self._preview_columns()
        self.preview_tree.heading("checked", text="Chọn", anchor="w")
        self.preview_tree.column("checked", width=58, minwidth=58, stretch=False, anchor="center")

        for field in self.data_fields:
            column_id = field["template_name"]
            self.preview_tree.heading(column_id, text=field["display_name"], anchor="w")
            width = 300 if field.get("role") == "company" else 160
            minwidth = 220 if field.get("role") == "company" else 110
            self.preview_tree.column(column_id, width=width, minwidth=minwidth, stretch=False)

    def _set_person_options(self, options):
        if getattr(self, "person_menu", None) is not None:
            self.person_menu.destroy()

        options = [option for option in options if option]
        if not options:
            options = ["Tất cả"]
        if self.responsible_person.get() not in options:
            self.responsible_person.set(options[0])

        self.person_menu = tk.Menubutton(
            self.person_box,
            textvariable=self.responsible_person,
            font=FONT_PATH,
            bg=BG_CARD,
            fg=TEXT_PRIMARY,
            activebackground=BG_CARD,
            activeforeground=TEXT_PRIMARY,
            relief="flat",
            bd=0,
            highlightthickness=0,
            anchor="w",
            width=22,
            padx=8,
            pady=4,
            cursor="hand2",
        )
        menu = tk.Menu(
            self.person_menu,
            tearoff=0,
            bg=BG_CARD,
            fg=TEXT_PRIMARY,
            activebackground=ACCENT,
            activeforeground="#0a0a0a",
            relief="flat",
            bd=0,
            font=FONT_SUB,
        )
        for option in options:
            menu.add_command(
                label=option,
                command=lambda value=option: self.responsible_person.set(value),
            )
        self.person_menu["menu"] = menu
        self.person_menu.pack(side="left", fill="x")

    def _open_data_settings(self):
        if self.is_running:
            self._log("Vui lòng đợi quá trình tạo hợp đồng hoàn tất trước khi mở thiết lập.", "warning")
            return

        window = tk.Toplevel(self)
        window.title("Thiết lập")
        window.configure(bg=BG_DARK)
        window.transient(self)
        window.grab_set()
        window.geometry("820x520")
        window.minsize(820, 520)

        tk.Label(window, text="Thiết lập", font=FONT_LABEL, bg=BG_DARK, fg=TEXT_PRIMARY).pack(anchor="w", padx=18, pady=(16, 8))

        pdf_settings = tk.Frame(window, bg=BG_DARK)
        pdf_settings.pack(fill="x", padx=18, pady=(0, 12))

        tk.Label(pdf_settings, text="Công cụ xuất PDF", font=FONT_DROP_SM, bg=BG_DARK, fg=TEXT_MUTED).pack(side="left", padx=(0, 10))
        working_pdf_engine = tk.StringVar(value=self.pdf_engine.get())
        pdf_engine_label = tk.StringVar(value="LibreOffice" if working_pdf_engine.get() == "libreoffice" else "Microsoft Word")
        pdf_engine_menu = tk.Menubutton(
            pdf_settings,
            textvariable=pdf_engine_label,
            font=FONT_PATH,
            bg=BG_CARD,
            fg=TEXT_PRIMARY,
            activebackground=BG_CARD,
            activeforeground=TEXT_PRIMARY,
            relief="flat",
            bd=0,
            highlightthickness=0,
            anchor="w",
            width=18,
            padx=8,
            pady=5,
            cursor="hand2",
        )
        pdf_menu = tk.Menu(
            pdf_engine_menu,
            tearoff=0,
            bg=BG_CARD,
            fg=TEXT_PRIMARY,
            activebackground=ACCENT,
            activeforeground="#0a0a0a",
            relief="flat",
            bd=0,
            font=FONT_SUB,
        )

        def set_pdf_engine(value, label):
            working_pdf_engine.set(value)
            pdf_engine_label.set(label)

        pdf_menu.add_command(label="LibreOffice", command=lambda: set_pdf_engine("libreoffice", "LibreOffice"))
        pdf_menu.add_command(label="Microsoft Word", command=lambda: set_pdf_engine("word", "Microsoft Word"))
        pdf_engine_menu["menu"] = pdf_menu
        pdf_engine_menu.pack(side="left")

        table_frame = tk.Frame(window, bg=BG_CARD, highlightbackground=BORDER, highlightthickness=1, height=310)
        table_frame.pack(fill="x", expand=False, padx=18, pady=(0, 14))
        table_frame.pack_propagate(False)

        tree = ttk.Treeview(table_frame, columns=("template_name", "excel_column", "display_name"), show="headings", height=10, style="ContractPreview.Treeview", selectmode="browse")
        tree.heading("template_name", text="Docx template", anchor="w")
        tree.heading("excel_column", text="Cột Excel", anchor="w")
        tree.heading("display_name", text="Tên hiển thị trên app", anchor="w")
        tree.column("template_name", width=330, minwidth=220, stretch=True)
        tree.column("excel_column", width=120, minwidth=90, stretch=False, anchor="w")
        tree.column("display_name", width=340, minwidth=220, stretch=True)
        tree.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")

        editor = tk.Frame(window, bg=BG_DARK)
        editor.pack(fill="x", padx=18, pady=(0, 12))

        field_template_var = tk.StringVar()
        column_var = tk.StringVar()
        display_var = tk.StringVar()

        def make_entry(label, var, width):
            box = tk.Frame(editor, bg=BG_DARK)
            box.pack(side="left", padx=(0, 10), fill="x", expand=True)
            tk.Label(box, text=label, font=FONT_DROP_SM, bg=BG_DARK, fg=TEXT_MUTED).pack(anchor="w")
            entry = tk.Entry(box, textvariable=var, font=FONT_PATH, bg=BG_CARD, fg=TEXT_PRIMARY, insertbackground=ACCENT, relief="flat", bd=0, width=width)
            entry.pack(fill="x", ipady=5)
            return entry

        make_entry("Template", field_template_var, 24)
        make_entry("Cột", column_var, 8)
        make_entry("Header", display_var, 20)

        working_fields = [dict(field) for field in self.data_fields]

        def reload_tree():
            tree.delete(*tree.get_children())
            for idx, field in enumerate(working_fields):
                tree.insert("", "end", iid=str(idx), values=(f"{{{{{field['template_name']}}}}}", field["excel_column"], field["display_name"]))

        def selected_index():
            selection = tree.selection()
            return int(selection[0]) if selection else None

        def fill_editor(_event=None):
            idx = selected_index()
            if idx is None:
                return
            field = working_fields[idx]
            field_template_var.set(field["template_name"])
            column_var.set(field["excel_column"])
            display_var.set(field["display_name"])

        def apply_editor():
            idx = selected_index()
            if idx is None:
                return
            name = field_template_var.get().strip().strip("{} ")
            column = column_var.get().strip().upper()
            display = display_var.get().strip()
            if not name or not column:
                self._log("Template và cột Excel không được để trống.", "warning")
                return
            working_fields[idx].update({"template_name": name, "excel_column": column, "display_name": display or name})
            reload_tree()
            tree.selection_set(str(idx))

        def add_field():
            working_fields.append({"template_name": "truong_moi", "excel_column": "A", "display_name": "Trường mới", "role": ""})
            reload_tree()
            tree.selection_set(str(len(working_fields) - 1))
            fill_editor()

        def remove_field():
            idx = selected_index()
            if idx is None:
                return
            del working_fields[idx]
            reload_tree()
            if working_fields:
                tree.selection_set(str(min(idx, len(working_fields) - 1)))
                fill_editor()

        def reset_defaults():
            working_fields[:] = default_fields()
            reload_tree()
            if working_fields:
                tree.selection_set("0")
                fill_editor()

        def save_and_close():
            apply_editor()
            self.pdf_engine.set(working_pdf_engine.get())
            self.data_fields = save_data_settings(working_fields, self.template_path.get().strip(), self.pdf_engine.get())
            self._configure_preview_columns()
            if self.excel_path.get().strip():
                self._load_responsible_persons(self.excel_path.get())
                self._load_contract_preview(self.excel_path.get())
            self._log("Đã lưu thiết lập dữ liệu vào data_fields.json", "success")
            self._log(f"Công cụ xuất PDF mặc định: {'LibreOffice' if self.pdf_engine.get() == 'libreoffice' else 'Microsoft Word'}", "accent")
            window.destroy()

        tree.bind("<<TreeviewSelect>>", fill_editor)
        reload_tree()
        if working_fields:
            tree.selection_set("0")
            fill_editor()

        buttons = tk.Frame(window, bg=BG_DARK)
        buttons.pack(fill="x", padx=18, pady=(0, 16))

        for text, command in (("Cập nhật dòng", apply_editor), ("Thêm", add_field), ("Xóa", remove_field), ("Mặc định", reset_defaults)):
            tk.Button(buttons, text=text, font=FONT_DROP_SM, bg=BG_CARD, fg=TEXT_MUTED, activebackground=BORDER, activeforeground=TEXT_PRIMARY, relief="flat", bd=0, padx=12, pady=7, cursor="hand2", command=command).pack(side="left", padx=(0, 8))

        tk.Button(buttons, text="Lưu", font=FONT_BTN, bg=ACCENT, fg="#0a0a0a", activebackground=ACCENT_GLOW, activeforeground="#0a0a0a", relief="flat", bd=0, padx=22, pady=8, cursor="hand2", command=save_and_close).pack(side="right")

    def _setup_drag_drop(self):
        if _DND_AVAILABLE:
            try:
                self.drop_zone.drop_target_register(DND_FILES)
                self.drop_zone.dnd_bind('<<Drop>>', self._on_drop)
                self.template_drop_zone.drop_target_register(DND_FILES)
                self.template_drop_zone.dnd_bind('<<Drop>>', self._on_template_drop)
                # self._log("Drag-and-drop đã sẵn sàng.", "muted")
            except Exception as e:
                self._log(f"Drag-drop lỗi: {e}. Dùng nút chọn file.", "warning")
        else:
            self._log("Kéo-thả không khả dụng (pip install tkinterdnd2). Dùng nút chọn file.", "warning")

    def _on_drop(self, event):
        """Handle file drop event."""
        if self.is_running:
            self._log("Vui lòng đợi quá trình tạo hợp đồng hoàn tất trước khi chọn file khác.", "warning")
            return
        raw = event.data.strip()
        # Path may be wrapped in braces on Windows: {C:/path/to/file.xlsx}
        path = raw.strip('{}').split('} {')[0].strip('{}')
        if path.lower().endswith('.xlsx') or path.lower().endswith('.xls'):
            self._set_file(path)
        else:
            self._log(f"Chỉ chấp nhận file .xlsx hoặc .xls, bạn thả: {os.path.basename(path)}", "warning")

    def _on_template_drop(self, event):
        if self.is_running:
            self._log("Vui lòng đợi quá trình tạo hợp đồng hoàn tất trước khi chọn template khác.", "warning")
            return
        raw = event.data.strip()
        path = raw.strip('{}').split('} {')[0].strip('{}')
        if path.lower().endswith(('.docx', '.doc')):
            self._set_template_file(path)
        else:
            self._log(f"Chỉ chấp nhận file .docx hoặc .doc, bạn thả: {os.path.basename(path)}", "warning")

    # ── Helpers ────────────────────────────────────────────────────────────────
    def _drop_hover(self, entering: bool):
        if self.is_running:
            return
        color = BG_DROP_HOVER if entering else BG_DROP
        border = ACCENT if entering else BORDER
        self.drop_zone.config(bg=color, highlightbackground=border)
        self.drop_icon.config(bg=color)
        self.drop_label.config(bg=color)
        for w in self.drop_zone.winfo_children():
            try:
                w.config(bg=color)
            except Exception:
                pass

    def _template_drop_hover(self, entering: bool):
        if self.is_running:
            return
        color = BG_DROP_HOVER if entering else BG_DROP
        border = ACCENT if entering else BORDER
        self.template_drop_zone.config(bg=color, highlightbackground=border)
        self.template_drop_icon.config(bg=color)
        self.template_drop_label.config(bg=color)
        for w in self.template_drop_zone.winfo_children():
            try:
                w.config(bg=color)
            except Exception:
                pass

    def _browse_file(self):
        if self.is_running:
            self._log("Vui lòng đợi quá trình tạo hợp đồng hoàn tất trước khi chọn file khác.", "warning")
            return
        path = filedialog.askopenfilename(
            title="Chọn file Excel",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if path:
            self._set_file(path)

    def _browse_template(self):
        if self.is_running:
            self._log("Vui lòng đợi quá trình tạo hợp đồng hoàn tất trước khi chọn template khác.", "warning")
            return
        current = self.template_path.get().strip()
        initial_dir = os.path.dirname(self._resolve_app_path(current)) if current else os.path.join(os.path.dirname(os.path.abspath(__file__)), "templates")
        path = filedialog.askopenfilename(
            title="Chọn file mẫu HĐ",
            initialdir=initial_dir if os.path.exists(initial_dir) else None,
            filetypes=[("Mẫu HĐ", "*.docx *.doc"), ("All files", "*.*")]
        )
        if path:
            self._set_template_file(path)

    def _set_file(self, path: str):
        self.excel_path.set(path)
        filename = os.path.basename(path)
        self.drop_label.config(text=filename, fg=SUCCESS)
        self._log(f"File được chọn: {path}", "accent")
        
        # Load responsible persons and contract rows from Excel
        self.after(100, lambda: self._load_responsible_persons(self.excel_path.get()))
        self.after(120, lambda: self._load_contract_preview(self.excel_path.get()))

    def _set_template_file(self, path: str):
        self.template_path.set(path)
        self.template_drop_label.config(text=os.path.basename(path), fg=SUCCESS)
        save_data_settings(self.data_fields, path)
        self._log(f"Mẫu HĐ được chọn: {path}", "accent")

    def _load_responsible_persons(self, excel_path: str):
        """Load responsible persons from Excel into the sale selector."""
        try:
            if not os.path.exists(excel_path):
                return
            
            persons = get_responsible_persons(excel_path, log=lambda x: None)
            if persons:
                person_list = sorted([p for p in persons.keys() if p.strip()])  # Filter empty strings
                # Recreate sale selector with new values.
                self.responsible_person.set("Tất cả")
                self._set_person_options(["Tất cả", *person_list])
                self._log(f"Tìm thấy {len(person_list)} sale phụ trách", "success")
        except Exception as e:
            self._log(f"Lỗi tải danh sách sale phụ trách: {e}", "warning")

    def _load_contract_preview(self, excel_path: str):
        """Load contract rows from Excel into the preview table."""
        try:
            if not os.path.exists(excel_path):
                return

            self.preview_tree.delete(*self.preview_tree.get_children())
            self.preview_rows = get_contract_preview(excel_path, log=lambda x: None)
            self.selected_row_ids = {row["row_index"] for row in self.preview_rows}
            self._refresh_preview_table()

            if self.preview_rows:
                self._log(f"Đã tải {len(self.preview_rows)} dòng thông tin hợp đồng", "success")
            else:
                self._log("Không tìm thấy dòng hợp đồng hợp lệ trong file Excel.", "warning")
        except Exception as e:
            self._log(f"Lỗi tải danh sách thông tin: {e}", "warning")

    def _is_all_responsible_person_selected(self) -> bool:
        selected_person = self.responsible_person.get().strip()
        normalized = selected_person.casefold().replace("(", "").replace(")", "").strip()
        return not selected_person or normalized in {"tất cả", "tat ca"}

    def _get_filtered_preview_rows(self):
        if self._is_all_responsible_person_selected():
            rows = list(self.preview_rows)
        else:
            selected_person = self.responsible_person.get().strip()
            rows = [
                row for row in self.preview_rows
                if row.get("_responsible_person", "") == selected_person
            ]

        keyword = self._normalize_search_text(self.company_search.get())
        if not keyword:
            return rows

        return [
            row for row in rows
            if keyword in self._normalize_search_text(row.get("_company", ""))
        ]

    def _normalize_search_text(self, text: str) -> str:
        normalized = unicodedata.normalize("NFKD", text or "")
        no_marks = "".join(
            char for char in normalized
            if unicodedata.category(char) != "Mn"
        )
        return no_marks.casefold().strip()

    def _refresh_preview_table(self):
        self.preview_tree.delete(*self.preview_tree.get_children())
        self.filtered_preview_rows = self._get_filtered_preview_rows()

        for row in self.filtered_preview_rows:
            row_id = str(row["row_index"])
            self.preview_tree.insert(
                "",
                "end",
                iid=row_id,
                tags=("normal",),
                values=(
                    self._checkbox_text(row["row_index"]),
                    *[row.get(field["template_name"], "") for field in self.data_fields],
                )
            )

        self._update_preview_count()
        self.hovered_preview_item = None

    def _on_responsible_person_changed(self, *_):
        if self.is_running or not hasattr(self, "preview_tree"):
            return
        self._refresh_preview_table()

    def _on_company_search_changed(self, *_):
        if self.is_running or not hasattr(self, "preview_tree"):
            return
        if self.company_search.get():
            if not self.clear_search_btn.winfo_ismapped():
                self.clear_search_btn.pack(side="left", padx=(2, 3))
        else:
            if self.clear_search_btn.winfo_ismapped():
                self.clear_search_btn.pack_forget()
        self._refresh_preview_table()

    def _clear_company_search(self):
        self.company_search.set("")
        self.search_entry.focus_set()

    def _checkbox_text(self, row_index: int) -> str:
        return "☑" if row_index in self.selected_row_ids else "☐"

    def _on_preview_click(self, event):
        if self.is_running:
            return "break"

        region = self.preview_tree.identify("region", event.x, event.y)
        if region != "cell":
            return None

        item_id = self.preview_tree.identify_row(event.y)
        if not item_id:
            return None

        row_index = int(item_id)
        if row_index in self.selected_row_ids:
            self.selected_row_ids.remove(row_index)
        else:
            self.selected_row_ids.add(row_index)

        values = list(self.preview_tree.item(item_id, "values"))
        values[0] = self._checkbox_text(row_index)
        self.preview_tree.item(item_id, values=values)
        self._update_preview_count()
        return "break"

    def _on_preview_motion(self, event):
        if self.preview_tree.identify("region", event.x, event.y) != "cell":
            previous_item = getattr(self, "hovered_preview_item", None)
            if previous_item and self.preview_tree.exists(previous_item):
                self.preview_tree.item(previous_item, tags=("normal",))
            self.hovered_preview_item = None
            return

        item_id = self.preview_tree.identify_row(event.y)
        if item_id == getattr(self, "hovered_preview_item", None):
            return

        previous_item = getattr(self, "hovered_preview_item", None)
        if previous_item and self.preview_tree.exists(previous_item):
            self.preview_tree.item(previous_item, tags=("normal",))

        self.hovered_preview_item = item_id
        if item_id and self.preview_tree.exists(item_id):
            self.preview_tree.item(item_id, tags=("hover",))

    def _on_preview_leave(self, _event):
        previous_item = getattr(self, "hovered_preview_item", None)
        if previous_item and self.preview_tree.exists(previous_item):
            self.preview_tree.item(previous_item, tags=("normal",))
        self.hovered_preview_item = None

    def _set_all_preview_rows(self, selected: bool):
        if self.is_running:
            return

        visible_row_ids = {row["row_index"] for row in self.filtered_preview_rows}
        if selected:
            self.selected_row_ids.update(visible_row_ids)
        else:
            self.selected_row_ids.difference_update(visible_row_ids)

        for item_id in self.preview_tree.get_children():
            row_index = int(item_id)
            values = list(self.preview_tree.item(item_id, "values"))
            values[0] = self._checkbox_text(row_index)
            self.preview_tree.item(item_id, values=values)

        self._update_preview_count()

    def _update_preview_count(self):
        total = len(self.filtered_preview_rows)
        selected = len({
            row["row_index"] for row in self.filtered_preview_rows
            if row["row_index"] in self.selected_row_ids
        })
        if total:
            self.preview_count_label.config(text=f"Đã chọn {selected}/{total}", fg=ACCENT_GLOW if selected else WARNING)
        else:
            self.preview_count_label.config(text="Chưa có dữ liệu", fg=TEXT_MUTED)

    def _log(self, message: str, tag: str = ""):
        """Append a message to the log box (thread-safe)."""
        def _append():
            self.log_box.config(state="normal")
            lines = str(message).splitlines() or [""]
            for line in lines:
                time_str = datetime.now().strftime("%H:%M:%S")
                self.log_box.insert("end", f"[{time_str}] ", "time")
                self.log_box.insert("end", line + "\n", tag if tag else "")
            self.log_box.see("end")
            self.log_box.config(state="disabled")
        self.after(0, _append)

    def _smart_log(self, message: str):
        """Auto-detect log level from plain-text message prefixes."""
        normalized = message.strip()
        if normalized.startswith(("Lỗi", "CẢNH BÁO", "Không có file DOCX", "Không thể")):
            tag = "error"
        elif normalized.startswith(("Cảnh báo", "Chưa", "LibreOffice không", "Không tìm thấy", "Đã hủy")):
            tag = "warning"
        elif normalized.startswith(("Đã", "Tìm thấy", "Hoàn tất")):
            tag = "success"
        elif normalized.startswith(("Đang", "Bắt đầu", "DOCX:", "PDF:", "Công cụ xuất PDF")):
            tag = "accent"
        else:
            tag = ""
        self._log(message, tag)

    def _clear_log(self):
        self.log_box.config(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.config(state="disabled")

    def _open_contracts(self):
        folder = os.path.join(os.path.dirname(os.path.abspath(__file__)), "contracts")
        if os.path.exists(folder):
            if sys.platform == "darwin":
                subprocess.Popen(["open", folder])
            elif os.name == "nt":
                os.startfile(folder)
            else:
                subprocess.Popen(["xdg-open", folder])
        else:
            self._log("Thư mục 'contracts' chưa tồn tại. Hãy tạo hợp đồng trước.", "warning")

    def _show_responsible_persons(self):
        excel = self.excel_path.get().strip()
        if not excel:
            self._log("Chưa chọn file Excel! Vui lòng kéo thả hoặc nhấn chọn file.", "warning")
            return
        if not os.path.exists(excel):
            self._log(f"Không tìm thấy file: {excel}", "error")
            return

        self._log("-" * 55, "muted")
        self._log("Lấy danh sách sale phụ trách...", "accent")
        
        persons = get_responsible_persons(excel, log=self._smart_log)
        
        if not persons:
            self._log("Không tìm thấy sale phụ trách nào trong file.", "warning")
            return
        
        self._log(f"Tìm thấy {len(persons)} sale phụ trách:\n", "success")
        
        for idx, (person, companies) in enumerate(sorted(persons.items()), 1):
            unique_companies = list(set(companies))
            self._log(f"  {idx}. {person} - Quản lý {len(unique_companies)} công ty", "")
            for company in sorted(unique_companies)[:5]:  # Show first 5 companies
                self._log(f"     - {company}", "muted")
            if len(unique_companies) > 5:
                self._log(f"     ... và {len(unique_companies) - 5} công ty khác", "muted")

    # ── Run ────────────────────────────────────────────────────────────────────
    def _cancel_run(self):
        if not self.is_running:
            return
        self.cancel_event.set()
        self.cancel_btn.config(state="disabled", bg="#3a1f1f", fg=TEXT_MUTED)
        self.status_label.config(text="Đang hủy...", fg=WARNING)
        self._log("Đã yêu cầu hủy. Sẽ dừng sau tác vụ hiện tại.", "warning")

    def _run(self, output_format="docx"):
        if self.is_running:
            return
        excel = self.excel_path.get().strip()
        if not excel:
            self._log("Chưa chọn file Excel! Vui lòng kéo thả hoặc nhấn chọn file.", "warning")
            return
        if not os.path.exists(excel):
            self._log(f"Không tìm thấy file: {excel}", "error")
            return
        template = self.template_path.get().strip()
        if not template:
            self._log("Chưa chọn file mẫu HĐ! Vui lòng kéo thả hoặc click chọn template.", "warning")
            return
        if not os.path.exists(self._resolve_app_path(template)):
            self._log(f"Không tìm thấy file template: {template}", "error")
            return

        selected_row_ids = None
        if self.preview_rows:
            selected_row_ids = [
                row["row_index"] for row in self.filtered_preview_rows
                if row["row_index"] in self.selected_row_ids
            ]

        if self.preview_rows and not selected_row_ids:
            self._log("Chưa chọn thông tin nào để xuất hợp đồng.", "warning")
            return

        self.is_running = True
        self.cancel_event.clear()
        active_button = self.docx_btn if output_format == "docx" else self.pdf_btn
        active_text = "Đang tạo Docx..." if output_format == "docx" else "Đang tạo PDF..."

        for btn in (self.docx_btn, self.pdf_btn):
            btn.config(state="disabled", bg="#504124", fg="#8e8e93")
        self.cancel_btn.config(state="normal", bg=ERROR_CLR, fg=TEXT_PRIMARY)
        self.person_menu.config(state="disabled")
        self.select_all_btn.config(state="disabled")
        self.select_none_btn.config(state="disabled")
        active_button.config(text=active_text)

        self.status_label.config(text="Đang tạo...", fg=WARNING)
        self._log("-" * 55, "muted")
        if output_format == "pdf":
            self._log(f"Công cụ xuất PDF đã chọn: {'LibreOffice' if self.pdf_engine.get() == 'libreoffice' else 'Microsoft Word'}", "accent")
        
        selected_person = self.responsible_person.get().strip()
        if not self._is_all_responsible_person_selected():
            self._log(f"Chuẩn bị xử lý file: {os.path.basename(excel)} (Sale phụ trách: {selected_person})", "accent")
            # Pass actual person name
            thread = threading.Thread(target=self._run_worker, args=(excel, output_format, selected_person, selected_row_ids), daemon=True)
        else:
            self._log(f"Chuẩn bị xử lý file: {os.path.basename(excel)} Tất cả sale phụ trách", "accent")
            # Pass None to generate all
            thread = threading.Thread(target=self._run_worker, args=(excel, output_format, "", selected_row_ids), daemon=True)

        # Run in background thread to keep UI responsive
        thread.start()

    def _run_worker(self, excel_path: str, output_format: str, responsible_person: str = "", selected_rows=None):
        try:
            success = generate(
                excel_path,
                log=self._smart_log,
                output_format=output_format,
                responsible_person=responsible_person if responsible_person else None,
                selected_rows=selected_rows,
                template_path=self.template_path.get().strip(),
                cancel_event=self.cancel_event,
                pdf_engine=self.pdf_engine.get(),
            )
            if success == "cancelled":
                self.after(0, lambda: self.status_label.config(text="Đã hủy", fg=WARNING))
            elif success:
                self.after(0, lambda: self.status_label.config(text="Hoàn tất", fg=SUCCESS))
            else:
                self.after(0, lambda: self.status_label.config(text="Có lỗi xảy ra", fg=WARNING))
        except Exception as e:
            self._log(f"Lỗi không mong muốn: {e}", "error")
            self.after(0, lambda: self.status_label.config(text="Lỗi", fg=ERROR_CLR))
        finally:
            self.is_running = False
            self.after(0, lambda: [
                self.docx_btn.config(state="normal", text="Tạo HĐ Docx", bg=ACCENT, fg="#0a0a0a"),
                self.pdf_btn.config(state="normal", text="Tạo HĐ PDF", bg=ACCENT, fg="#0a0a0a"),
                self.cancel_btn.config(state="disabled", bg=BG_CARD, fg=TEXT_MUTED),
                self.person_menu.config(state="normal"),
                self.select_all_btn.config(state="normal"),
                self.select_none_btn.config(state="normal")
            ])


if __name__ == "__main__":
    app = App()
    app.mainloop()
