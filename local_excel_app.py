from __future__ import annotations
import sys
from pathlib import Path
from tkinter import messagebox

try:
    import customtkinter as ctk
except ImportError:
    import tkinter as tk
    
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror(
        "Thiếu thư viện hệ thống",
        "Chưa cài đặt thư viện giao diện CustomTkinter. Vui lòng chạy lệnh cài đặt:\n\n"
        "pip install -r requirements.txt\n\n"
        "Sau đó chạy lại phần mềm."
    )
    sys.exit(1)

import os
import threading
import traceback
import webbrowser
import json
import re
from tkinter import filedialog, scrolledtext, StringVar

from local_excel_runner import DEFAULT_WORKBOOK_PATH, inspect_workbook, run_local_refresh

WINDOW_TITLE = "HR-NEXUS Local Excel Runner"
GUIDE_PATH = Path(__file__).resolve().parent / "LOCAL_EXCEL_USAGE.md"

# Configure CustomTkinter
ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")

class SettingsWindow(ctk.CTkToplevel):
    def __init__(self, parent, current_id: str, on_save_callback):
        super().__init__(parent)
        self.title("Cấu hình Google Sheets")
        self.geometry("500x340")
        self.resizable(False, False)
        
        # Make modal
        self.transient(parent)
        self.grab_set()
        
        self.on_save_callback = on_save_callback
        
        # GUI
        self._build_ui(current_id)
        
    def _build_ui(self, current_id):
        frame = ctk.CTkFrame(self, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Email Display
        ctk.CTkLabel(frame, text="Email Service Account can share (Editor):", font=ctk.CTkFont(weight="bold")).pack(anchor="w")
        
        email_frame = ctk.CTkFrame(frame, fg_color="transparent")
        email_frame.pack(fill="x", pady=(5, 15))
        
        self.email_entry = ctk.CTkEntry(email_frame, width=350)
        self.email_entry.pack(side="left", fill="x", expand=True)
        self.email_entry.insert(0, "hr-nexus-sync@[YOUR_PROJECT_ID].iam.gserviceaccount.com")
        self.email_entry.configure(state="readonly")
        
        copy_btn = ctk.CTkButton(email_frame, text="Copy", width=60, command=self._copy_email)
        copy_btn.pack(side="left", padx=(10, 0))
        
        # ID Input
        ctk.CTkLabel(frame, text="Spreadsheet ID hoặc dán nguyên đường dẫn URL:", font=ctk.CTkFont(weight="bold")).pack(anchor="w")
        
        self.id_entry = ctk.CTkEntry(frame)
        self.id_entry.pack(fill="x", pady=(5, 20))
        self.id_entry.insert(0, current_id)
        
        # Buttons
        btn_frame = ctk.CTkFrame(frame, fg_color="transparent")
        btn_frame.pack(fill="x", pady=(10, 0))
        
        save_btn = ctk.CTkButton(btn_frame, text="Lưu cấu hình", command=self._save)
        save_btn.pack(side="left", expand=True, padx=(0, 5))
        
        cancel_btn = ctk.CTkButton(btn_frame, text="Hủy", fg_color="gray", hover_color="#555555", command=self.destroy)
        cancel_btn.pack(side="left", expand=True, padx=(5, 0))
        
    def _copy_email(self):
        self.clipboard_clear()
        self.clipboard_append(self.email_entry.get())
        
    def _save(self):
        val = self.id_entry.get().strip()
        # Parse URL if needed
        match = re.search(r"spreadsheets/d/([a-zA-Z0-9-_]+)", val)
        if match:
            val = match.group(1)
            
        self.on_save_callback(val)
        self.destroy()

class LocalExcelApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        self.title(WINDOW_TITLE)
        self.geometry("1100x800")
        self.minsize(1000, 720)
        
        self.workbook_var = StringVar(value=str(DEFAULT_WORKBOOK_PATH))
        self.fiscal_year_var = StringVar()
        self.status_var = StringVar(value="Sẵn sàng thao tác.")
        self._busy = False

        self._load_config()
        self._build_layout()
        
    def _load_config(self):
        from gsheet_sync import DEFAULT_CONFIG_PATH
        try:
            with open(DEFAULT_CONFIG_PATH, encoding="utf-8") as f:
                config = json.load(f)
            self.fiscal_year_var.set(config.get("fiscal_year", "2026"))
        except Exception:
            self.fiscal_year_var.set("2026")

    def _build_layout(self) -> None:
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)
        
        # HERo
        hero = ctk.CTkFrame(self, fg_color="#17324d", corner_radius=0)
        hero.grid(row=0, column=0, sticky="ew")
        
        hero_inner = ctk.CTkFrame(hero, fg_color="transparent")
        hero_inner.pack(fill="x", padx=30, pady=20)
        
        title_row = ctk.CTkFrame(hero_inner, fg_color="transparent")
        title_row.pack(fill="x")
        
        ctk.CTkLabel(title_row, text="HR-NEXUS Local Excel Runner", font=ctk.CTkFont(size=24, weight="bold"), text_color="#fff7e6").pack(side="left")
        
        settings_btn = ctk.CTkButton(title_row, text="⚙️", width=40, height=40, fg_color="#2b4d6e", hover_color="#3a638b", font=ctk.CTkFont(size=20), command=self._open_settings)
        settings_btn.pack(side="right")
        
        ctk.CTkLabel(hero_inner, text="Hệ thống tổng hợp tự động Master Data L&D. Phiên bản Desktop CTK.", text_color="#d7e5f3").pack(anchor="w", pady=(5, 0))
        
        # MAIN CONTAINER
        container = ctk.CTkFrame(self, fg_color="#f4efe6", corner_radius=0)
        container.grid(row=1, column=0, sticky="nsew")
        container.columnconfigure(0, weight=3)
        container.columnconfigure(1, weight=2)
        container.rowconfigure(1, weight=1)

        # LEFT CARD (Controls & Log)
        controls = ctk.CTkFrame(container, fg_color="white", corner_radius=12)
        controls.grid(row=0, column=0, sticky="nsew", padx=(30, 15), pady=20)
        
        ctk.CTkLabel(controls, text="Cấu Hình Đầu Vào", font=ctk.CTkFont(size=14, weight="bold"), text_color="#17324d").pack(anchor="w", padx=20, pady=(20, 10))
        
        wb_row = ctk.CTkFrame(controls, fg_color="transparent")
        wb_row.pack(fill="x", padx=20)
        ctk.CTkEntry(wb_row, textvariable=self.workbook_var, height=36).pack(side="left", fill="x", expand=True)
        self.browse_button = ctk.CTkButton(wb_row, text="Chọn file", width=80, height=36, fg_color="#f0f0f0", text_color="#17324d", hover_color="#e0e0e0", command=self._browse_workbook)
        self.browse_button.pack(side="left", padx=(10, 0))
        
        fy_row = ctk.CTkFrame(controls, fg_color="transparent")
        fy_row.pack(fill="x", padx=20, pady=15)
        ctk.CTkLabel(fy_row, text="Năm báo cáo:", text_color="#3e4f5f").pack(side="left")
        self.fiscal_entry = ctk.CTkEntry(fy_row, textvariable=self.fiscal_year_var, width=100, height=36)
        self.fiscal_entry.pack(side="left", padx=(10, 0))
        
        btns = ctk.CTkFrame(controls, fg_color="transparent")
        btns.pack(fill="x", padx=20, pady=(10, 0))
        
        self.check_button = ctk.CTkButton(btns, text="🔍 Kiểm tra nguồn", height=40, font=ctk.CTkFont(weight="bold"), command=self._inspect_workbook)
        self.check_button.pack(side="left", padx=(0, 10))
        
        self.run_button = ctk.CTkButton(btns, text="⚙️ Đồng bộ local", height=40, font=ctk.CTkFont(weight="bold"), command=self._run_refresh)
        self.run_button.pack(side="left", padx=(0, 10))
        
        self.sync_button = ctk.CTkButton(btns, text="☁️ Kéo từ GG Sheets", height=40, fg_color="#10b981", hover_color="#059669", font=ctk.CTkFont(weight="bold"), command=self._run_gsheet_sync)
        self.sync_button.pack(side="left", padx=(0, 10))
        
        self.push_full_button = ctk.CTkButton(btns, text="☁️ Đẩy lên GG Sheets", height=40, fg_color="#f59e0b", hover_color="#d97706", font=ctk.CTkFont(weight="bold"), command=self._run_push_full)
        self.push_full_button.pack(side="left", padx=(0, 10))
        
        ctk.CTkLabel(controls, textvariable=self.status_var, text_color="#6a7d8d").pack(anchor="w", padx=20, pady=15)
        
        # LOG
        log_card = ctk.CTkFrame(container, fg_color="white", corner_radius=12)
        log_card.grid(row=1, column=0, sticky="nsew", padx=(30, 15), pady=(0, 20))
        
        ctk.CTkLabel(log_card, text="Nhật Ký Xử Lý", font=ctk.CTkFont(size=14, weight="bold"), text_color="#17324d").pack(anchor="w", padx=20, pady=(20, 5))
        
        self.log_text = ctk.CTkTextbox(log_card, font=ctk.CTkFont(family="Consolas", size=12), fg_color="#15283c", text_color="#edf3f8", wrap="word", corner_radius=8)
        self.log_text.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        self.log_text.configure(state="disabled")

        # RIGHT CARD (Guide)
        guide_card = ctk.CTkFrame(container, fg_color="white", corner_radius=12)
        guide_card.grid(row=0, column=1, rowspan=2, sticky="nsew", padx=(15, 30), pady=20)
        
        ctk.CTkLabel(guide_card, text="Admin cần nhập gì?", font=ctk.CTkFont(size=14, weight="bold"), text_color="#17324d").pack(anchor="w", padx=20, pady=(20, 10))
        
        guide_text = (
            "1. 'Danh mục nhân viên': cập nhật master nhân sự, phòng ban.\n\n"
            "2. 'Lớp đào tạo': cập nhật thông tin lớp, chi phí.\n\n"
            "3. 'Data raw học viên': dữ liệu học viên từng lớp.\n\n"
            "4. 'Dữ liệu đào tạo': tự động build, KHÔNG NHẬP TAY."
        )
        ctk.CTkLabel(guide_card, text=guide_text, text_color="#3e4f5f", justify="left", wraplength=280).pack(anchor="nw", padx=20)
        
        ctk.CTkLabel(guide_card, text="Trình tự thao tác", font=ctk.CTkFont(size=14, weight="bold"), text_color="#17324d").pack(anchor="w", padx=20, pady=(30, 10))
        
        flow_text = (
            "1. Đóng file Excel.\n"
            "2. Bấm 'Kiểm tra nguồn'.\n"
            "3. Bấm 'Đồng bộ local'.\n"
            "4. Đẩy/Kéo Google Sheets tuỳ nhu cầu."
        )
        ctk.CTkLabel(guide_card, text=flow_text, text_color="#3e4f5f", justify="left", wraplength=280).pack(anchor="nw", padx=20)
        
        # open workbook button
        self.open_button = ctk.CTkButton(guide_card, text="Mở File Excel Hiện Tại", fg_color="#f0f0f0", hover_color="#e0e0e0", text_color="#17324d", command=self._open_workbook)
        self.open_button.pack(anchor="w", padx=20, pady=20)

    def _open_settings(self):
        from gsheet_sync import DEFAULT_CONFIG_PATH
        try:
            with open(DEFAULT_CONFIG_PATH, encoding="utf-8") as f:
                config = json.load(f)
            current_id = config.get("spreadsheet_id", "")
        except Exception:
            current_id = ""
            
        def on_save(new_id):
            if new_id and new_id != current_id:
                try:
                    with open(DEFAULT_CONFIG_PATH, encoding="utf-8") as f:
                        config = json.load(f)
                except Exception:
                    config = {}
                
                config["spreadsheet_id"] = new_id
                
                with open(DEFAULT_CONFIG_PATH, "w", encoding="utf-8") as f:
                    json.dump(config, f, indent=2)
                
                self._append_log(f"✅ Đã cập nhật Spreadsheet ID: {new_id[:10]}...")
                messagebox.showinfo("Thành công", "Đã lưu cấu hình Google Sheets mới!")
                
        SettingsWindow(self, current_id, on_save)

    def _browse_workbook(self) -> None:
        path = filedialog.askopenfilename(
            title="Chọn workbook local",
            filetypes=[("Excel Workbook", "*.xlsx"), ("All files", "*.*")],
            initialdir=str(DEFAULT_WORKBOOK_PATH.parent),
        )
        if path:
            self.workbook_var.set(path)
            
    def _open_workbook(self) -> None:
        path = self.workbook_var.get()
        if os.path.exists(path):
            webbrowser.open(path)
        else:
            messagebox.showwarning("Khong tim thay file", f"Khong tim thay file: {path}")

    def _inspect_workbook(self) -> None:
        self._run_in_background("Đang kiểm tra...", self._inspect_task)

    def _run_refresh(self) -> None:
        self._run_in_background("Đang đồng bộ local...", self._refresh_task)

    def _run_gsheet_sync(self) -> None:
        self._run_in_background("Đang đồng bộ từ Google Sheets...", self._gsheet_sync_task)

    def _inspect_task(self) -> None:
        inspect_workbook(self.workbook_var.get(), fiscal_year=self._fiscal_year(), logger=self._append_log)

    def _refresh_task(self) -> None:
        result = run_local_refresh(
            self.workbook_var.get(),
            fiscal_year=self._fiscal_year(),
            backup=True,
            logger=self._append_log,
        )
        backup_text = str(result.backup_path) if result.backup_path else "không"
        self._append_log(f"Hoàn tất. Backup: {backup_text}")

    def _gsheet_sync_task(self) -> None:
        from gsheet_sync import run_sync
        result = run_sync(logger=self._append_log)
        if not result.success:
            raise RuntimeError(result.errors[0] if result.errors else "Đồng bộ thất bại")

    def _run_push_full(self) -> None:
        confirmed = messagebox.askyesno(
            WINDOW_TITLE,
            "CẢNH BÁO: Thao tác này sẽ XOÁ TOÀN BỘ dữ liệu trên Google Sheets "
            "và thay thế bằng file Excel hiện tại.\n\n"
            "Chắc chắn tiếp tục?",
            icon="warning"
        )
        if not confirmed:
            return
        self._run_in_background("Đang đẩy lên Google Sheets...", self._push_full_task)

    def _push_full_task(self) -> None:
        from gsheet_sync import run_push_full
        result = run_push_full(
            workbook_path=self.workbook_var.get(),
            logger=self._append_log,
        )
        if not result.success:
            raise RuntimeError(result.errors[0] if result.errors else "Đẩy dữ liệu thất bại")

    def _run_in_background(self, busy_text: str, task) -> None:
        if self._busy:
            return
        self._busy = True
        self._set_controls_state("disabled")
        self.status_var.set(busy_text)

        def worker() -> None:
            try:
                current_fy = self._fiscal_year()
                if current_fy:
                    self._save_fiscal_year_to_config(current_fy)
                    
                task()
            except Exception as exc:
                error_detail = traceback.format_exc()
                self._append_log(f"Lỗi:\n{error_detail}")
                self.after(0, lambda: messagebox.showerror("Lỗi", str(exc)))
            finally:
                self.after(0, self._on_task_done)

        threading.Thread(target=worker, daemon=True).start()

    def _save_fiscal_year_to_config(self, fy: str):
        from gsheet_sync import DEFAULT_CONFIG_PATH
        try:
            with open(DEFAULT_CONFIG_PATH, encoding="utf-8") as f:
                config = json.load(f)
            if config.get("fiscal_year") != fy:
                config["fiscal_year"] = fy
                with open(DEFAULT_CONFIG_PATH, "w", encoding="utf-8") as f:
                    json.dump(config, f, indent=2)
        except Exception:
            pass

    def _on_task_done(self) -> None:
        self._busy = False
        self._set_controls_state("normal")
        self.status_var.set("Hoàn tất.")

    def _set_controls_state(self, state: str) -> None:
        self.check_button.configure(state=state)
        self.run_button.configure(state=state)
        self.sync_button.configure(state=state)
        self.push_full_button.configure(state=state)
        self.browse_button.configure(state=state)
        self.fiscal_entry.configure(state=state)
        self.open_button.configure(state=state)

    def _append_log(self, text: str) -> None:
        def do_append() -> None:
            self.log_text.configure(state="normal")
            self.log_text.insert("end", text + "\n")
            self.log_text.see("end")
            self.log_text.configure(state="disabled")

        self.after(0, do_append)

    def _fiscal_year(self) -> str | None:
        value = self.fiscal_year_var.get().strip()
        return value if value else None

def main() -> None:
    app = LocalExcelApp()
    app.mainloop()

if __name__ == "__main__":
    main()
