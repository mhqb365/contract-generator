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

# Try importing tkinterdnd2 for native drag-and-drop
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    _BASE = TkinterDnD.Tk
    _DND_AVAILABLE = True
except ImportError:
    _BASE = tk.Tk
    _DND_AVAILABLE = False

# Import logic module
from generate_contracts import generate, get_responsible_persons

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
        self.minsize(700, 600)

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
        self.geometry("820x680")
        self.update_idletasks()
        x = (self.winfo_screenwidth() - 820) // 2
        y = (self.winfo_screenheight() - 680) // 2
        self.geometry(f"820x680+{x}+{y}")

        self.excel_path = tk.StringVar(value="")
        self.is_running  = False

        self._build_ui()
        self._setup_drag_drop()

    # ── UI Layout ──────────────────────────────────────────────────────────────
    def _build_ui(self):
        # ── Header ──
        header = tk.Frame(self, bg=BG_DARK, pady=20)
        header.pack(fill="x", padx=30)

        tk.Label(
            header, text="Tạo Hợp Đồng Nguyên Tắc",
            font=FONT_TITLE, bg=BG_DARK, fg=TEXT_PRIMARY
        ).pack(anchor="w")
        tk.Label(
            header,
            text="Kéo thả file Excel thông tin vào khung bên dưới hoặc click vào để chọn file. Sau đó bấm nút Tạo HĐ",
            font=FONT_SUB, bg=BG_DARK, fg=TEXT_MUTED
        ).pack(anchor="w", pady=(2, 0))

        # ── Drop zone ──
        drop_frame = tk.Frame(self, bg=BG_DARK, padx=30)
        drop_frame.pack(fill="x")

        self.drop_zone = tk.Frame(
            drop_frame, bg=BG_DROP,
            highlightbackground=BORDER, highlightthickness=2,
            cursor="hand2"
        )
        self.drop_zone.pack(fill="x", ipady=10)

        self.drop_icon = tk.Label(
            self.drop_zone, text="Chọn file", font=("Segoe UI", 14, "bold"),
            bg=BG_DROP, fg=ACCENT
        )
        self.drop_icon.pack(pady=(10, 3))

        self.drop_label = tk.Label(
            self.drop_zone,
            text="Kéo & Thả file .xlsx vào đây",
            font=FONT_DROP, bg=BG_DROP, fg=TEXT_PRIMARY
        )
        self.drop_label.pack()

        tk.Label(
            self.drop_zone, text="hoặc click để chọn file",
            font=FONT_DROP_SM, bg=BG_DROP, fg=TEXT_MUTED
        ).pack(pady=(2, 10))

        # Bind click events for all children
        for widget in [self.drop_zone, self.drop_icon, self.drop_label]:
            widget.bind("<Button-1>", lambda e: self._browse_file())
            widget.bind("<Enter>",    lambda e: self._drop_hover(True))
            widget.bind("<Leave>",    lambda e: self._drop_hover(False))

        # ── Selected file path display ──
        path_frame = tk.Frame(self, bg=BG_DARK, padx=30)
        path_frame.pack(fill="x", pady=(8, 0))

        self.path_label = tk.Label(
            path_frame, textvariable=self.excel_path,
            font=FONT_PATH, bg=BG_DARK, fg=TEXT_MUTED,
            anchor="w", wraplength=760
        )
        self.path_label.pack(fill="x")

        # ── Responsible person selector ──
        self.selector_frame = tk.Frame(self, bg=BG_DARK, padx=30)
        self.selector_frame.pack(fill="x", pady=(8, 12))

        tk.Label(
            self.selector_frame, text="Sale phụ trách",
            font=FONT_SUB, bg=BG_DARK, fg=TEXT_MUTED
        ).pack(anchor="w", pady=(0, 4))

        self.responsible_person = tk.StringVar(value="")
        self.person_menu = tk.OptionMenu(
            self.selector_frame,
            self.responsible_person,
            "(Tất cả)"
        )
        self.person_menu.config(
            bg=BG_CARD, fg=ACCENT_GLOW, font=FONT_SUB,
            activebackground=ACCENT, activeforeground="#0a0a0a",
            relief="flat", bd=0, anchor="w", indicatoron=False
        )
        self.person_menu["menu"].config(bg=BG_CARD, fg=ACCENT_GLOW, font=FONT_SUB)
        self.person_menu.pack(fill="x", padx=0)

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

        self._log("Kéo file Excel vào khung phía trên. Tạo Docx trước mới tạo được PDF.", "muted")

    # ── Drag-and-drop setup ────────────────────────────────────────────────────
    def _setup_drag_drop(self):
        if _DND_AVAILABLE:
            try:
                self.drop_zone.drop_target_register(DND_FILES)
                self.drop_zone.dnd_bind('<<Drop>>', self._on_drop)
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

    def _set_file(self, path: str):
        self.excel_path.set(path)
        filename = os.path.basename(path)
        self.drop_label.config(text=filename, fg=SUCCESS)
        self.path_label.config(fg=TEXT_MUTED)
        self._log(f"File được chọn: {path}", "accent")
        
        # Load responsible persons into combobox
        self.after(100, lambda: self._load_responsible_persons(self.excel_path.get()))

    def _load_responsible_persons(self, excel_path: str):
        """Load responsible persons from Excel into OptionMenu."""
        try:
            if not os.path.exists(excel_path):
                return
            
            persons = get_responsible_persons(excel_path, log=lambda x: None)
            if persons:
                person_list = sorted([p for p in persons.keys() if p.strip()])  # Filter empty strings
                # Recreate OptionMenu with new values
                self.person_menu.destroy()
                self.responsible_person.set("(Tất cả)")
                self.person_menu = tk.OptionMenu(
                    self.selector_frame,
                    self.responsible_person,
                    "(Tất cả)",
                    *person_list
                )
                self.person_menu.config(
                    bg=BG_CARD, fg=ACCENT_GLOW, font=FONT_SUB,
                    activebackground=ACCENT, activeforeground="#0a0a0a",
                    relief="flat", bd=0, anchor="w", indicatoron=False
                )
                self.person_menu["menu"].config(bg=BG_CARD, fg=ACCENT_GLOW, font=FONT_SUB)
                self.person_menu.pack(fill="x", padx=0)
                self._log(f"Tìm thấy {len(person_list)} sale phụ trách", "success")
        except Exception as e:
            self._log(f"Lỗi tải danh sách sale phụ trách: {e}", "warning")

    def _log(self, message: str, tag: str = ""):
        """Append a message to the log box (thread-safe)."""
        def _append():
            self.log_box.config(state="normal")
            time_str = datetime.now().strftime("%H:%M:%S")
            self.log_box.insert("end", f"[{time_str}] ", "time")
            self.log_box.insert("end", message + "\n", tag if tag else "")
            self.log_box.see("end")
            self.log_box.config(state="disabled")
        self.after(0, _append)

    def _smart_log(self, message: str):
        """Auto-detect log level from plain-text message prefixes."""
        normalized = message.strip()
        if normalized.startswith(("Lỗi", "CẢNH BÁO", "Không có file DOCX", "Không thể")):
            tag = "error"
        elif normalized.startswith(("Cảnh báo", "Chưa", "LibreOffice không", "Không tìm thấy")):
            tag = "warning"
        elif normalized.startswith(("Đã", "Tìm thấy", "Hoàn tất")):
            tag = "success"
        elif normalized.startswith(("Đang", "Bắt đầu", "DOCX:", "PDF:")):
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

        self.is_running = True
        active_button = self.docx_btn if output_format == "docx" else self.pdf_btn
        active_text = "Đang tạo Docx..." if output_format == "docx" else "Đang tạo PDF..."

        for btn in (self.docx_btn, self.pdf_btn):
            btn.config(state="disabled", bg="#504124", fg="#8e8e93")
        self.person_menu.config(state="disabled")
        active_button.config(text=active_text)

        self.status_label.config(text="Đang tạo...", fg=WARNING)
        self._log("-" * 55, "muted")
        
        selected_person = self.responsible_person.get().strip()
        if selected_person and selected_person != "(Tất cả)":
            self._log(f"Bắt đầu xử lý file: {os.path.basename(excel)} (Sale phụ trách: {selected_person})", "accent")
            # Pass actual person name
            thread = threading.Thread(target=self._run_worker, args=(excel, output_format, selected_person), daemon=True)
        else:
            self._log(f"Bắt đầu xử lý file: {os.path.basename(excel)} (Tất cả sale phụ trách)", "accent")
            # Pass None to generate all
            thread = threading.Thread(target=self._run_worker, args=(excel, output_format, ""), daemon=True)

        # Run in background thread to keep UI responsive
        thread.start()

    def _run_worker(self, excel_path: str, output_format: str, responsible_person: str = ""):
        try:
            success = generate(excel_path, log=self._smart_log, output_format=output_format, responsible_person=responsible_person if responsible_person else None)
            if success:
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
                self.person_menu.config(state="normal")
            ])


if __name__ == "__main__":
    app = App()
    app.mainloop()
