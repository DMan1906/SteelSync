import os
import re
import shutil
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
# pyrefly: ignore [missing-import]
import customtkinter as ctk
import pandas as pd
from pathlib import Path
import sys
import json
import stat
import subprocess
import tempfile
import urllib.request
from collections import Counter

# openpyxl for advanced Excel formatting
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side, Alignment, PatternFill
from openpyxl.utils import get_column_letter

# ─── Permissions & File Access Management ────────────────────────────────────

def remove_readonly(path):
    """
    Remove read-only attribute from a file or folder so any app can modify or overwrite it.
    Works on Windows (clearing win32 FILE_ATTRIBUTE_READONLY) and POSIX.
    """
    path_str = str(path)
    try:
        if os.path.isdir(path_str):
            os.chmod(path_str, stat.S_IRWXU | stat.S_IRWXG | stat.S_IRWXO)
        else:
            os.chmod(path_str, stat.S_IREAD | stat.S_IWRITE | stat.S_IRGRP | stat.S_IWGRP | stat.S_IROTH | stat.S_IWOTH)
    except Exception:
        pass

    if sys.platform.startswith("win"):
        try:
            import ctypes
            attrs = ctypes.windll.kernel32.GetFileAttributesW(path_str)
            FILE_ATTRIBUTE_READONLY = 0x01
            FILE_ATTRIBUTE_NORMAL = 0x80
            if attrs != -1 and (attrs & FILE_ATTRIBUTE_READONLY):
                new_attrs = attrs & ~FILE_ATTRIBUTE_READONLY
                if new_attrs == 0:
                    new_attrs = FILE_ATTRIBUTE_NORMAL
                ctypes.windll.kernel32.SetFileAttributesW(path_str, new_attrs)
        except Exception:
            pass


def reset_to_default_permissions(path, is_dir=None, recursive=False):
    """
    Ensure the path has default Windows inherited permissions and is not read-only.
    Does NOT add any special or custom ACLs (keeps pure default Windows inheritance).
    """
    path_str = str(path)
    if not os.path.exists(path_str):
        return

    remove_readonly(path_str)

    if sys.platform.startswith("win"):
        try:
            icacls_bin = os.path.join(os.environ.get("SystemRoot", r"C:\Windows"), "System32", "icacls.exe")
            if not os.path.isfile(icacls_bin):
                icacls_bin = "icacls"

            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = 0  # SW_HIDE
            creationflags = getattr(subprocess, 'CREATE_NO_WINDOW', 0x08000000)

            args = [icacls_bin, path_str, "/reset"]
            if recursive:
                args.extend(["/T", "/C", "/Q"])
            else:
                args.extend(["/C", "/Q"])

            subprocess.run(
                args,
                capture_output=True,
                startupinfo=startupinfo,
                creationflags=creationflags,
                timeout=30 if recursive else 5
            )
        except Exception:
            pass

grant_full_permissions = reset_to_default_permissions


def prepare_destination(dst):
    """
    Ensure destination file or folder path is writeable before writing, copying, or moving to it.
    If it exists, removes read-only attribute.
    """
    try:
        path_str = str(dst)
        if os.path.exists(path_str):
            remove_readonly(path_str)
    except Exception:
        pass


def make_dir_with_permissions(dir_path):
    """
    Create directory and ensure default Windows inheritance and no read-only attribute.
    """
    path_str = str(dir_path)
    os.makedirs(path_str, exist_ok=True)
    remove_readonly(path_str)
    return path_str


def ensure_full_permissions_tree(root_dir):
    """
    Recursively remove read-only attribute and ensure default Windows inheritance
    without any special ACLs.
    """
    path_str = str(root_dir)
    if not os.path.exists(path_str):
        return

    try:
        remove_readonly(path_str)
        for root, dirs, files in os.walk(path_str):
            for d in dirs:
                remove_readonly(os.path.join(root, d))
            for f in files:
                remove_readonly(os.path.join(root, f))
    except Exception:
        pass

    reset_to_default_permissions(path_str, is_dir=True, recursive=True)


def safe_copy(src, dst):
    """
    Copy file ensuring destination is writeable, binary identical, and inherits default Windows permissions.
    """
    src_str = str(src)
    dst_str = str(dst)
    prepare_destination(dst_str)
    shutil.copyfile(src_str, dst_str)
    remove_readonly(dst_str)


def safe_move(src, dst):
    """
    Move file ensuring destination is writeable, falling back to copy+delete if locked or across drives.
    """
    src_str = str(src)
    dst_str = str(dst)
    prepare_destination(dst_str)
    remove_readonly(src_str)
    try:
        shutil.move(src_str, dst_str)
    except Exception:
        safe_copy(src_str, dst_str)
        try:
            remove_readonly(src_str)
            os.remove(src_str)
        except Exception:
            pass

    remove_readonly(dst_str)


def safe_save_workbook(wb, filepath, log_callback=None):
    """
    Save openpyxl Workbook ensuring destination is writeable, handling file-in-use locks gracefully.
    """
    path_str = str(filepath)
    prepare_destination(path_str)
    try:
        wb.save(path_str)
        remove_readonly(path_str)
        return path_str
    except PermissionError:
        import time
        base, ext = os.path.splitext(path_str)
        timestamp = time.strftime("_%Y%m%d_%H%M%S")
        alt_path = f"{base}{timestamp}{ext}"
        if log_callback:
            log_callback(f"⚠ Fișierul {os.path.basename(path_str)} este deschis în altă aplicație (Excel). Salvat ca: {os.path.basename(alt_path)}")
        prepare_destination(alt_path)
        wb.save(alt_path)
        remove_readonly(alt_path)
        return alt_path


def safe_df_to_excel(df, filepath, index=False, log_callback=None):
    """
    Save pandas DataFrame to Excel, handling file-in-use locks gracefully.
    """
    path_str = str(filepath)
    prepare_destination(path_str)
    try:
        df.to_excel(path_str, index=index)
        remove_readonly(path_str)
        return path_str
    except PermissionError:
        import time
        base, ext = os.path.splitext(path_str)
        timestamp = time.strftime("_%Y%m%d_%H%M%S")
        alt_path = f"{base}{timestamp}{ext}"
        if log_callback:
            log_callback(f"⚠ Fișierul {os.path.basename(path_str)} este deschis în altă aplicație (Excel). Salvat ca: {os.path.basename(alt_path)}")
        prepare_destination(alt_path)
        df.to_excel(alt_path, index=index)
        remove_readonly(alt_path)
        return alt_path


def safe_write_lines(dest_path, lines):
    """
    Write text lines to a file ensuring destination is writeable and inherits default Windows permissions.
    """
    path_str = str(dest_path)
    prepare_destination(path_str)
    with open(path_str, 'w', encoding='utf-8') as file:
        file.writelines(lines)
    remove_readonly(path_str)


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath(os.path.dirname(__file__)), relative_path)


def _get_settings_path():
    """Return path for the persistent settings JSON file."""
    appdata = os.environ.get('APPDATA')
    if not appdata:
        appdata = os.path.expanduser('~\\AppData\\Roaming')
    app_folder = os.path.join(appdata, "SteelSync")
    make_dir_with_permissions(app_folder)
    return os.path.join(app_folder, "steelsync_settings.json")

# ─── Theme ────────────────────────────────────────────────────────────────
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

# Color palette
_BG          = ("#f2f2f7", "#1a1a2e")   # main window background
_CARD_BG     = ("#ffffff", "#16213e")   # card / frame background
_CARD_BORDER = ("#d1d1d6", "#0f3460")   # subtle border tint
_ACCENT      = ("#007aff", "#00b4d8")   # teal accent (buttons, progress)
_ACCENT_HOV  = ("#005bb5", "#0096c7")   # hover state
_GREEN       = ("#34c759", "#06d6a0")   # start button / success
_GREEN_HOV   = ("#248a3d", "#05b98b")
_TEXT_PRI    = ("#1c1c1e", "#e0e0e0")   # primary text
_TEXT_SEC    = ("#8e8e93", "#94a3b8")   # secondary / muted text
_ENTRY_BG    = ("#e5e5ea", "#0d1b2a")   # entry field background
_LOG_BG      = ("#ffffff", "#0b0f19")   # log terminal background
_DANGER      = ("#ff3b30", "#ef476f")   # error color
_ORANGE      = ("#ff9500", "#f4a261")   # optional / info accent

SETTINGS_PATH = _get_settings_path()
APP_VERSION = "1.0.5"

GITHUB_REPO = "DMan1906/SteelSync"
GITHUB_API_LATEST = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"


def parse_version(v_str):
    """
    Parse 'v1.0.1t', 'v1.0.1', '1.0.2' into a tuple comparable with other versions.
    For each part, splits into (integer_part, string_suffix).
    E.g. '1.0.1t' -> ((1, ''), (0, ''), (1, 't'))
         '1.0.1'  -> ((1, ''), (0, ''), (1, ''))
    """
    clean = re.sub(r'^[vV]', '', str(v_str).strip())
    parts = []
    for p in clean.split('.'):
        m = re.match(r'^(\d+)(.*)$', p)
        if m:
            num = int(m.group(1))
            suffix = m.group(2).strip().lower()
            parts.append((num, suffix))
        else:
            parts.append((0, p.lower()))
    while len(parts) < 3:
        parts.append((0, ''))
    return tuple(parts)


def is_newer_version(latest_tag, current_ver=APP_VERSION):
    """Return True if latest_tag is strictly newer than current_ver."""
    try:
        clean_latest = re.sub(r'^[vV]', '', str(latest_tag).strip())
        clean_current = re.sub(r'^[vV]', '', str(current_ver).strip())
        if clean_latest.lower() == clean_current.lower():
            return False

        if parse_version(clean_latest) > parse_version(clean_current):
            return True

        if clean_latest.startswith(clean_current) and len(clean_latest) > len(clean_current):
            return True

        return False
    except Exception:
        return False


def fetch_latest_release_info(timeout=3.5):
    """
    Fetch release info from GitHub API.
    Returns dict with tag_name, body, html_url, asset_url, asset_size, asset_name or None.
    """
    req = urllib.request.Request(
        GITHUB_API_LATEST,
        headers={
            "User-Agent": "SteelSync-Toolkit-Updater",
            "Accept": "application/vnd.github.v3+json"
        }
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as response:
            if response.status == 200:
                data = json.loads(response.read().decode('utf-8'))
                tag = data.get("tag_name", "").strip()
                body = data.get("body", "").strip()
                html_url = data.get("html_url", "")
                
                assets = data.get("assets", [])
                asset_url = None
                asset_size = 0
                asset_name = "SteelSync.exe"
                for a in assets:
                    if a.get("name", "").lower().endswith(".exe"):
                        asset_url = a.get("browser_download_url")
                        asset_size = a.get("size", 0)
                        asset_name = a.get("name")
                        break
                
                if not asset_url and tag:
                    asset_url = f"https://github.com/{GITHUB_REPO}/releases/download/{tag}/SteelSync.exe"

                return {
                    "tag_name": tag,
                    "body": body,
                    "html_url": html_url,
                    "asset_url": asset_url,
                    "asset_size": asset_size,
                    "asset_name": asset_name
                }
    except Exception:
        return None
    return None


class UpdateDialog(ctk.CTkToplevel):
    """
    Modern modal dialog notifying the user about a new version with release notes,
    progress bar for downloading, and automatic restart via batch script.
    """
    def __init__(self, parent, release_info):
        super().__init__(parent)
        self.parent_app = parent
        self.release_info = release_info
        self.tag_name = release_info.get("tag_name", "")
        self.asset_url = release_info.get("asset_url", "")
        self.asset_name = release_info.get("asset_name", "SteelSync.exe")
        self.asset_size = release_info.get("asset_size", 0)
        self.body = release_info.get("body", "").strip() or "Îmbunătățiri și optimizări de performanță."

        self.title("Actualizare SteelSync Toolkit")
        self.geometry("560x520")
        self.minsize(500, 480)
        self.configure(fg_color=_BG)
        
        self.transient(parent)
        self.grab_set()

        self.update_idletasks()
        try:
            x = parent.winfo_x() + (parent.winfo_width() // 2) - 280
            y = parent.winfo_y() + (parent.winfo_height() // 2) - 260
            self.geometry(f"+{max(0, x)}+{max(0, y)}")
        except Exception:
            pass

        self._cancelled = False
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self._create_widgets()

    def on_close(self):
        self._cancelled = True
        self.destroy()

    def _create_widgets(self):
        container = ctk.CTkFrame(self, fg_color="transparent")
        container.pack(fill="both", expand=True, padx=20, pady=20)

        # Header card
        header_card = ctk.CTkFrame(container, fg_color=_CARD_BG, border_color=_CARD_BORDER, border_width=1, corner_radius=12)
        header_card.pack(fill="x", pady=(0, 14), padx=2)

        hdr_inner = ctk.CTkFrame(header_card, fg_color="transparent")
        hdr_inner.pack(fill="x", padx=16, pady=14)

        title_lbl = ctk.CTkLabel(
            hdr_inner,
            text=f"🚀 Versiune nouă disponibilă: {self.tag_name}",
            font=("Segoe UI", 16, "bold"),
            text_color=_TEXT_PRI,
            anchor="w"
        )
        title_lbl.pack(fill="x")

        sub_lbl = ctk.CTkLabel(
            hdr_inner,
            text=f"Versiunea curentă instalată: v{APP_VERSION}",
            font=("Segoe UI", 12),
            text_color=_TEXT_SEC,
            anchor="w"
        )
        sub_lbl.pack(fill="x", pady=(2, 0))

        # Changelog / Release notes
        notes_card = ctk.CTkFrame(container, fg_color=_CARD_BG, border_color=_CARD_BORDER, border_width=1, corner_radius=12)
        notes_card.pack(fill="both", expand=True, pady=(0, 14), padx=2)

        notes_hdr = ctk.CTkFrame(notes_card, fg_color="transparent")
        notes_hdr.pack(fill="x", padx=16, pady=(12, 6))

        notes_title = ctk.CTkLabel(
            notes_hdr,
            text="Ce aduce nou această versiune:",
            font=("Segoe UI", 13, "bold"),
            text_color=_ACCENT,
            anchor="w"
        )
        notes_title.pack(fill="x")

        self.notes_box = ctk.CTkTextbox(
            notes_card,
            font=("Segoe UI", 12),
            fg_color=_ENTRY_BG,
            text_color=_TEXT_PRI,
            corner_radius=8,
            wrap="word"
        )
        self.notes_box.pack(fill="both", expand=True, padx=16, pady=(0, 14))
        self.notes_box.insert("1.0", self.body)
        self.notes_box.configure(state="disabled")

        # Progress / Status Frame
        self.progress_frame = ctk.CTkFrame(container, fg_color="transparent")
        self.progress_frame.pack(fill="x", pady=(0, 12))

        self.status_lbl = ctk.CTkLabel(
            self.progress_frame,
            text="Apasă pe 'Actualizează acum' pentru descărcare și instalare automată.",
            font=("Segoe UI", 11),
            text_color=_TEXT_SEC,
            anchor="w"
        )
        self.status_lbl.pack(fill="x", padx=2, pady=(0, 4))

        self.progress_bar = ctk.CTkProgressBar(
            self.progress_frame,
            progress_color=_ACCENT,
            height=10,
            corner_radius=5
        )
        self.progress_bar.pack(fill="x", padx=2)
        self.progress_bar.set(0)

        # Buttons Frame
        self.btn_frame = ctk.CTkFrame(container, fg_color="transparent")
        self.btn_frame.pack(fill="x", pady=(0, 2))

        self.update_btn = ctk.CTkButton(
            self.btn_frame,
            text="🚀 Actualizează acum",
            font=("Segoe UI", 13, "bold"),
            fg_color=_GREEN,
            hover_color=_GREEN_HOV,
            height=38,
            corner_radius=8,
            command=self.start_download
        )
        self.update_btn.pack(side="left", fill="x", expand=True, padx=(0, 6))

        self.later_btn = ctk.CTkButton(
            self.btn_frame,
            text="Mai târziu",
            font=("Segoe UI", 12),
            fg_color=_CARD_BG,
            hover_color=_CARD_BORDER,
            text_color=_TEXT_PRI,
            height=38,
            corner_radius=8,
            command=self.destroy
        )
        self.later_btn.pack(side="left", padx=6)

        self.ignore_btn = ctk.CTkButton(
            self.btn_frame,
            text="Nu mai întreba",
            font=("Segoe UI", 11),
            fg_color="transparent",
            hover_color=_CARD_BORDER,
            text_color=_TEXT_SEC,
            height=38,
            corner_radius=8,
            command=self.ignore_this_version
        )
        self.ignore_btn.pack(side="left", padx=(6, 0))

    def start_download(self):
        self.update_btn.configure(state="disabled")
        self.later_btn.configure(state="disabled")
        self.ignore_btn.configure(state="disabled")
        self.status_lbl.configure(text="Se inițializează conexiunea...", text_color=_TEXT_PRI)
        threading.Thread(target=self._download_worker, daemon=True).start()

    def _download_worker(self):
        url = self.asset_url
        if not url:
            url = f"https://github.com/{GITHUB_REPO}/releases/download/{self.tag_name}/SteelSync.exe"

        temp_dir = tempfile.gettempdir()
        temp_exe = os.path.join(temp_dir, f"SteelSync_update_{self.tag_name}.exe")

        req = urllib.request.Request(
            url,
            headers={
                "User-Agent": "SteelSync-Toolkit-Updater"
            }
        )

        try:
            with urllib.request.urlopen(req, timeout=30) as resp, open(temp_exe, "wb") as out_f:
                content_len = resp.getheader("Content-Length")
                total_bytes = int(content_len) if content_len else self.asset_size
                downloaded = 0
                chunk_size = 65536  # 64 KB

                while True:
                    if self._cancelled:
                        return
                    chunk = resp.read(chunk_size)
                    if not chunk:
                        break
                    out_f.write(chunk)
                    downloaded += len(chunk)

                    if total_bytes > 0:
                        prog = min(1.0, downloaded / total_bytes)
                        mb_down = downloaded / (1024 * 1024)
                        mb_total = total_bytes / (1024 * 1024)
                        status_text = f"Se descarcă: {mb_down:.1f} MB din {mb_total:.1f} MB ({int(prog * 100)}%)"
                    else:
                        mb_down = downloaded / (1024 * 1024)
                        prog = 0.5
                        status_text = f"Se descarcă: {mb_down:.1f} MB..."

                    self.after(0, self._update_progress_ui, prog, status_text)

            if not self._cancelled:
                self.after(0, lambda: self._apply_update(temp_exe))

        except Exception as e:
            if not self._cancelled:
                self.after(0, lambda: self._handle_download_error(str(e)))

    def _update_progress_ui(self, prog, status_text):
        try:
            self.progress_bar.set(prog)
            self.status_lbl.configure(text=status_text)
        except Exception:
            pass

    def _handle_download_error(self, err_msg):
        self.status_lbl.configure(text=f"Eroare descărcare: {err_msg}", text_color=_DANGER)
        self.update_btn.configure(state="normal", text="Reîncearcă")
        self.later_btn.configure(state="normal")
        self.ignore_btn.configure(state="normal")

    def ignore_this_version(self):
        self.parent_app.ignored_update_version = self.tag_name
        self.parent_app.save_settings()
        self.destroy()

    def _apply_update(self, temp_exe):
        self.status_lbl.configure(text="Descărcare finalizată! Se aplică actualizarea...", text_color=_GREEN)
        self.progress_bar.set(1.0)

        is_frozen = getattr(sys, 'frozen', False)
        if not is_frozen:
            messagebox.showinfo(
                "Mod Dezvoltare (Python)",
                f"Actualizarea {self.tag_name} a fost descărcată cu succes în:\n{temp_exe}\n\n"
                "(În versiunea compilată SteelSync.exe, aplicația va fi înlocuită automat și va reporni)."
            )
            self.destroy()
            return

        current_exe = os.path.abspath(sys.executable)
        target_dir = os.path.dirname(current_exe)
        exe_name = os.path.basename(current_exe)
        bat_path = os.path.join(tempfile.gettempdir(), "steelsync_updater.bat")

        bat_content = f"""@echo off
chcp 65001 > nul
set "_MEIPASS2="
set "_MEIPASS="
set "TARGET_EXE={current_exe}"
set "TARGET_DIR={target_dir}"
set "NEW_EXE={temp_exe}"
set "EXE_NAME={exe_name}"
set /a RETRIES=0

:wait_loop
set /a RETRIES+=1
timeout /t 1 /nobreak > nul
tasklist /fi "IMAGENAME eq %EXE_NAME%" 2>nul | find /i "%EXE_NAME%" >nul
if not errorlevel 1 (
    if %RETRIES% lss 10 goto wait_loop
    taskkill /f /im "%EXE_NAME%" >nul 2>&1
    timeout /t 1 /nobreak > nul
)

timeout /t 2 /nobreak > nul

set /a COPY_RETRIES=0
:copy_loop
set /a COPY_RETRIES+=1
copy /y "%NEW_EXE%" "%TARGET_EXE%" > nul
if errorlevel 1 (
    if %COPY_RETRIES% lss 10 (
        timeout /t 1 /nobreak > nul
        goto copy_loop
    )
    move /y "%NEW_EXE%" "%TARGET_EXE%" > nul
)

if exist "%NEW_EXE%" del /f /q "%NEW_EXE%" > nul 2>&1

timeout /t 1 /nobreak > nul
set "_MEIPASS2="
set "_MEIPASS="
cd /d "%TARGET_DIR%"
start "" /D "%TARGET_DIR%" "%TARGET_EXE%"

(goto) 2>nul & del /f /q "%~f0"
exit
"""

        try:
            with open(bat_path, "w", encoding="utf-8") as f:
                f.write(bat_content)

            clean_env = os.environ.copy()
            for var in list(clean_env.keys()):
                if var.startswith('_MEI') or var in ('_MEIPASS2', '_MEIPASS', 'PYI_CHILD_SPLASH'):
                    clean_env.pop(var, None)

            if 'PATH' in clean_env:
                clean_paths = [p for p in clean_env['PATH'].split(os.pathsep) if '_MEI' not in p and '_mei' not in p]
                clean_env['PATH'] = os.pathsep.join(clean_paths)

            creationflags = getattr(subprocess, 'CREATE_NO_WINDOW', 0x08000000)
            subprocess.Popen(
                ["cmd.exe", "/c", bat_path],
                creationflags=creationflags,
                close_fds=True,
                env=clean_env
            )

            self.parent_app.destroy()
            os._exit(0)
        except Exception as e:
            messagebox.showerror("Eroare Actualizare", f"Nu s-a putut porni scriptul de actualizare:\n{e}")
            self.update_btn.configure(state="normal")
            self.later_btn.configure(state="normal")


class UnifiedApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("⚙ SteelSync Toolkit")
        self.geometry("960x860")
        self.minsize(850, 780)
        
        # Set application icon if it exists
        icon_path = resource_path("app_icon.ico")
        if os.path.exists(icon_path):
            self.iconbitmap(icon_path)
            try:
                self.iconbitmap(default=icon_path)
            except Exception:
                pass
            
        self.configure(fg_color=_BG)
        
        self.appearance_mode_var = ctk.StringVar(value="Dark")

        # ── Extractor Variables ──
        self.source_dir = ctk.StringVar()
        self.dest_dir = ctk.StringVar()
        self.duplicate_handling = ctk.StringVar(value="Keep Newest")
        self.extract_mode = ctk.StringVar(value="Doar Deficit")
        self.excel_file = ctk.StringVar()
        self.col_name = ctk.StringVar()
        
        # ── NC Sorter Variables ──
        self.nc_source_folder = ctk.StringVar()
        self.nc_action_var = ctk.StringVar(value="Copy")
        self.nc_opt_hot_cold = ctk.BooleanVar(value=False)
        self.nc_opt_separate_tubes = ctk.BooleanVar(value=False)
        self.nc_opt_verify_geometry = ctk.BooleanVar(value=True)
        self.nc_opt_no_header = ctk.BooleanVar(value=False)
        self.nc_opt_no_colors = ctk.BooleanVar(value=False)
        # NC Sorter – Excel / BOM for quantity lookup
        self.nc_excel_file = ctk.StringVar()
        self.nc_part_col = ctk.StringVar()
        self.nc_qty_col = ctk.StringVar()
        
        # ── Modify NC Variables ──
        self.mod_nc_dir = ctk.StringVar()
        self.mod_nc_find = ctk.StringVar()
        self.mod_nc_replace = ctk.StringVar()
        self.mod_nc_excel = ctk.StringVar()
        self.mod_nc_create_subfolder = ctk.BooleanVar(value=True)

        # ── Extension tracking ──
        self.available_extensions = ['.nc', '.pdf', '.dwg', '.dxf', '.step', '.stp']
        self.extension_vars = {}

        # ── Update tracking ──
        self.ignored_update_version = ""
        self._update_dialog = None

        self.create_widgets()
        self.load_settings()

        # Check for updates automatically in background (silent & non-blocking)
        self.after(1500, lambda: self.start_update_check(manual=False))

        # Save settings on close
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    # ══════════════════════════════════════════════════════════════════════
    #  SETTINGS PERSISTENCE
    # ══════════════════════════════════════════════════════════════════════
    def load_settings(self):
        """Load persisted user settings from JSON."""
        try:
            with open(SETTINGS_PATH, 'r', encoding='utf-8') as f:
                s = json.load(f)
            self.ignored_update_version = s.get("ignored_update_version", "")
            self.appearance_mode_var.set(s.get("appearance_mode", "Dark"))
            ctk.set_appearance_mode(self.appearance_mode_var.get())
            self.source_dir.set(s.get("source_dir", ""))
            self.dest_dir.set(s.get("dest_dir", ""))
            self.excel_file.set(s.get("excel_file", ""))
            self.col_name.set(s.get("col_name", ""))
            self.duplicate_handling.set(s.get("duplicate_handling", "Keep Newest"))
            self.extract_mode.set(s.get("extract_mode", "Doar Deficit"))
            self.nc_source_folder.set(s.get("nc_source_folder", ""))
            self.nc_action_var.set(s.get("nc_action_var", "Copy"))
            self.nc_opt_hot_cold.set(s.get("nc_opt_hot_cold", False))
            self.nc_opt_separate_tubes.set(s.get("nc_opt_separate_tubes", False))
            self.nc_opt_verify_geometry.set(s.get("nc_opt_verify_geometry", True))
            self.nc_opt_no_header.set(s.get("nc_opt_no_header", False))
            self.nc_opt_no_colors.set(s.get("nc_opt_no_colors", False))
            self.nc_excel_file.set(s.get("nc_excel_file", ""))
            self.nc_part_col.set(s.get("nc_part_col", ""))
            self.nc_qty_col.set(s.get("nc_qty_col", ""))
            
            # Modify NC settings
            self.mod_nc_dir.set(s.get("mod_nc_dir", ""))
            self.mod_nc_find.set(s.get("mod_nc_find", ""))
            self.mod_nc_replace.set(s.get("mod_nc_replace", ""))
            self.mod_nc_excel.set(s.get("mod_nc_excel", ""))
            self.mod_nc_create_subfolder.set(s.get("mod_nc_create_subfolder", True))

            # Restore extension selections
            saved_exts = s.get("extensions", {})
            for ext, val in saved_exts.items():
                if ext in self.extension_vars:
                    self.extension_vars[ext].set(val)
                elif ext not in self.available_extensions:
                    self.available_extensions.append(ext)
                    self.extension_vars[ext] = ctk.BooleanVar(value=val)
            self.draw_extensions()
        except (FileNotFoundError, json.JSONDecodeError):
            pass
        except Exception:
            pass

    def save_settings(self):
        """Persist current user settings to JSON."""
        ext_map = {ext: var.get() for ext, var in self.extension_vars.items()}
        settings = {
            "appearance_mode": self.appearance_mode_var.get(),
            "source_dir": self.source_dir.get(),
            "dest_dir": self.dest_dir.get(),
            "excel_file": self.excel_file.get(),
            "col_name": self.col_name.get(),
            "duplicate_handling": self.duplicate_handling.get(),
            "extract_mode": self.extract_mode.get(),
            "nc_source_folder": self.nc_source_folder.get(),
            "nc_action_var": self.nc_action_var.get(),
            "nc_opt_hot_cold": self.nc_opt_hot_cold.get(),
            "nc_opt_separate_tubes": self.nc_opt_separate_tubes.get(),
            "nc_opt_verify_geometry": self.nc_opt_verify_geometry.get(),
            "nc_opt_no_header": self.nc_opt_no_header.get(),
            "nc_opt_no_colors": self.nc_opt_no_colors.get(),
            "nc_excel_file": self.nc_excel_file.get(),
            "nc_part_col": self.nc_part_col.get(),
            "nc_qty_col": self.nc_qty_col.get(),
            "mod_nc_dir": self.mod_nc_dir.get(),
            "mod_nc_find": self.mod_nc_find.get(),
            "mod_nc_replace": self.mod_nc_replace.get(),
            "mod_nc_excel": self.mod_nc_excel.get(),
            "mod_nc_create_subfolder": self.mod_nc_create_subfolder.get(),
            "ignored_update_version": getattr(self, "ignored_update_version", ""),
            "extensions": ext_map,
        }
        try:
            prepare_destination(SETTINGS_PATH)
            with open(SETTINGS_PATH, 'w', encoding='utf-8') as f:
                json.dump(settings, f, indent=2, ensure_ascii=False)
            remove_readonly(SETTINGS_PATH)
            grant_full_permissions(SETTINGS_PATH, is_dir=False)
        except Exception:
            pass

    def toggle_appearance_mode(self):
        ctk.set_appearance_mode(self.appearance_mode_var.get())
        self.save_settings()

    def on_close(self):
        self.save_settings()
        self.destroy()

    # ══════════════════════════════════════════════════════════════════════
    #  AUTO-UPDATE METHODS
    # ══════════════════════════════════════════════════════════════════════
    def start_update_check(self, manual=False):
        """Check for updates in a non-blocking background thread."""
        threading.Thread(target=self._check_update_worker, args=(manual,), daemon=True).start()

    def _check_update_worker(self, manual=False):
        info = fetch_latest_release_info(timeout=4.0 if not manual else 7.0)
        if not info:
            if manual:
                self.after(0, lambda: messagebox.showinfo(
                    "Verificare Actualizări",
                    "Nu s-a putut contacta serverul GitHub.\nVerifică conexiunea la internet sau încearcă mai târziu."
                ))
            return

        tag = info.get("tag_name", "")
        if is_newer_version(tag, APP_VERSION):
            if not manual and tag == getattr(self, "ignored_update_version", ""):
                return
            self.after(0, lambda: self.show_update_dialog(info))
        else:
            if manual:
                self.after(0, lambda: messagebox.showinfo(
                    "Verificare Actualizări",
                    f"Aplicația este la zi!\nVersiunea curentă este v{APP_VERSION}."
                ))

    def show_update_dialog(self, release_info):
        if getattr(self, "_update_dialog", None) is not None and self._update_dialog.winfo_exists():
            self._update_dialog.lift()
            self._update_dialog.focus_force()
            return
        self._update_dialog = UpdateDialog(self, release_info)

    # ══════════════════════════════════════════════════════════════════════
    #  UI HELPERS
    # ══════════════════════════════════════════════════════════════════════
    def _make_card(self, parent, **kw):
        return ctk.CTkFrame(
            parent,
            fg_color=_CARD_BG,
            corner_radius=12,
            border_width=1,
            border_color=_CARD_BORDER,
            **kw
        )

    @staticmethod
    def _label(parent, text, **kw):
        defaults = dict(text_color=_TEXT_SEC, font=("Segoe UI", 13))
        defaults.update(kw)
        return ctk.CTkLabel(parent, text=text, **defaults)

    @staticmethod
    def _entry(parent, textvariable, **kw):
        defaults = dict(
            textvariable=textvariable,
            fg_color=_ENTRY_BG,
            border_color=_CARD_BORDER,
            text_color=_TEXT_PRI,
            font=("Segoe UI", 13),
            corner_radius=8,
        )
        defaults.update(kw)
        return ctk.CTkEntry(parent, **defaults)

    @staticmethod
    def _browse_btn(parent, command):
        return ctk.CTkButton(
            parent,
            text="📁 Browse",
            command=command,
            width=110,
            height=34,
            fg_color=_ACCENT,
            hover_color=_ACCENT_HOV,
            corner_radius=8,
            font=("Segoe UI", 12, "bold"),
        )

    # ══════════════════════════════════════════════════════════════════════
    #  MAIN LAYOUT
    # ══════════════════════════════════════════════════════════════════════
    def create_widgets(self):
        container = ctk.CTkFrame(self, fg_color="transparent")
        container.pack(fill="both", expand=True, padx=18, pady=14)
        
        top_bar = ctk.CTkFrame(container, fg_color="transparent")
        top_bar.pack(fill="x", pady=(0, 10))
        
        title_lbl = ctk.CTkLabel(
            top_bar,
            text="SteelSync Toolkit",
            font=("Segoe UI", 24, "bold"),
            text_color=_TEXT_PRI,
        )
        title_lbl.pack(side="left")
        
        theme_switch = ctk.CTkSwitch(
            top_bar, text="Dark Mode", variable=self.appearance_mode_var,
            onvalue="Dark", offvalue="Light", command=self.toggle_appearance_mode,
            progress_color=_ACCENT, text_color=_TEXT_PRI
        )
        theme_switch.pack(side="right", padx=(8, 0))

        self.version_lbl = ctk.CTkLabel(
            top_bar,
            text=f"v{APP_VERSION}",
            font=("Segoe UI", 11),
            text_color=_TEXT_SEC,
            cursor="hand2"
        )
        self.version_lbl.pack(side="right", padx=(0, 16))
        self.version_lbl.bind("<Button-1>", lambda e: self.start_update_check(manual=True))

        self.tabview = ctk.CTkTabview(container, fg_color=_CARD_BG, border_color=_CARD_BORDER, border_width=1)
        self.tabview.pack(fill="both", expand=True)
        
        self.tab_ext = self.tabview.add("📄 File Extractor")
        self.tab_nc = self.tabview.add("🔧 NC Sorter")
        self.tab_mod = self.tabview.add("📝 Modify NC")
        
        self.create_extractor_tab()
        self.create_nc_sorter_tab()
        self.create_mod_nc_tab()

    # ══════════════════════════════════════════════════════════════════════
    #  EXTRACTOR TAB
    # ══════════════════════════════════════════════════════════════════════
    def create_extractor_tab(self):
        container = ctk.CTkFrame(self.tab_ext, fg_color="transparent")
        container.pack(fill="both", expand=True, padx=10, pady=10)
        container.grid_columnconfigure(0, weight=1)
        
        subtitle_lbl = ctk.CTkLabel(
            container,
            text="Caută și copiază fișierele proiectului pe baza unei liste Excel",
            font=("Segoe UI", 12),
            text_color=_TEXT_SEC,
        )
        subtitle_lbl.grid(row=0, column=0, pady=(0, 10), sticky="w")

        # ──── Paths card ────
        paths_card = self._make_card(container)
        paths_card.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        paths_card.grid_columnconfigure(1, weight=1)

        section_lbl = ctk.CTkLabel(paths_card, text="📂 Directoare & Excel", font=("Segoe UI", 14, "bold"), text_color=_TEXT_PRI)
        section_lbl.grid(row=0, column=0, columnspan=3, padx=16, pady=(10, 6), sticky="w")

        self._label(paths_card, text="Sursă:").grid(row=1, column=0, padx=(16, 6), pady=6, sticky="w")
        self._entry(paths_card, self.source_dir).grid(row=1, column=1, padx=4, pady=6, sticky="ew")
        self._browse_btn(paths_card, self.browse_source).grid(row=1, column=2, padx=(4, 16), pady=6)

        self._label(paths_card, text="Destinație:").grid(row=2, column=0, padx=(16, 6), pady=6, sticky="w")
        self._entry(paths_card, self.dest_dir).grid(row=2, column=1, padx=4, pady=6, sticky="ew")
        self._browse_btn(paths_card, self.browse_dest).grid(row=2, column=2, padx=(4, 16), pady=6)

        self._label(paths_card, text="Fișier Excel:").grid(row=3, column=0, padx=(16, 6), pady=(6, 14), sticky="w")
        
        excel_inner = ctk.CTkFrame(paths_card, fg_color="transparent")
        excel_inner.grid(row=3, column=1, padx=4, pady=(6, 14), sticky="ew")
        excel_inner.grid_columnconfigure(0, weight=1)

        self._entry(excel_inner, self.excel_file).grid(row=0, column=0, sticky="ew")
        self._label(excel_inner, text="Col:", font=("Segoe UI", 11)).grid(row=0, column=1, padx=(8, 2))
        self._entry(excel_inner, self.col_name, width=80, placeholder_text="auto").grid(row=0, column=2, padx=(0, 0))

        self._browse_btn(paths_card, self.browse_excel).grid(row=3, column=2, padx=(4, 16), pady=(6, 14))

        # ──── Extensions card ────
        ext_card = self._make_card(container)
        ext_card.grid(row=2, column=0, sticky="ew", pady=(0, 10))

        section_lbl2 = ctk.CTkLabel(ext_card, text="🔖 Extensii Fișiere", font=("Segoe UI", 14, "bold"), text_color=_TEXT_PRI)
        section_lbl2.pack(anchor="w", padx=16, pady=(10, 4))

        self.ext_inner_frame = ctk.CTkFrame(ext_card, fg_color="transparent")
        self.ext_inner_frame.pack(anchor="w", padx=16, pady=(0, 4))

        self.draw_extensions()

        add_row = ctk.CTkFrame(ext_card, fg_color="transparent")
        add_row.pack(anchor="w", padx=16, pady=(0, 10))
        self.new_ext_entry = ctk.CTkEntry(
            add_row, width=90, placeholder_text=".ext",
            fg_color=_ENTRY_BG, border_color=_CARD_BORDER,
            text_color=_TEXT_PRI, corner_radius=8, font=("Segoe UI", 12)
        )
        self.new_ext_entry.pack(side="left", padx=(0, 6))
        ctk.CTkButton(
            add_row, text="＋ Add", width=70, height=30,
            fg_color=_ACCENT, hover_color=_ACCENT_HOV,
            corner_radius=8, font=("Segoe UI", 12, "bold"),
            command=self.add_extension,
        ).pack(side="left")

        # ──── Settings row ────
        settings_card = self._make_card(container)
        settings_card.grid(row=3, column=0, sticky="ew", pady=(0, 10))

        section_lbl3 = ctk.CTkLabel(settings_card, text="⚙ Setări", font=("Segoe UI", 14, "bold"), text_color=_TEXT_PRI)
        section_lbl3.pack(side="left", padx=16, pady=10)

        self._label(settings_card, text="Duplicate:").pack(side="left", padx=(20, 6), pady=10)
        ctk.CTkOptionMenu(
            settings_card,
            variable=self.duplicate_handling,
            values=["Keep Newest", "Keep Oldest"],
            fg_color=_ENTRY_BG,
            button_color=_ACCENT,
            button_hover_color=_ACCENT_HOV,
            dropdown_fg_color=_CARD_BG,
            dropdown_hover_color=_ACCENT,
            text_color=_TEXT_PRI,
            font=("Segoe UI", 13),
            corner_radius=8,
            width=130,
        ).pack(side="left", padx=(0, 10), pady=10)
        
        self._label(settings_card, text="Extracție:").pack(side="left", padx=(10, 6), pady=10)
        ctk.CTkOptionMenu(
            settings_card,
            variable=self.extract_mode,
            values=["Doar Deficit", "Extrage Tot"],
            fg_color=_ENTRY_BG,
            button_color=_ACCENT,
            button_hover_color=_ACCENT_HOV,
            dropdown_fg_color=_CARD_BG,
            dropdown_hover_color=_ACCENT,
            text_color=_TEXT_PRI,
            font=("Segoe UI", 13),
            corner_radius=8,
            width=130,
        ).pack(side="left", padx=(0, 16), pady=10)

        # ──── Start & Progress ────
        self.start_btn = ctk.CTkButton(
            container,
            text="▶ Start Extracție",
            command=self.start_extraction,
            fg_color=_GREEN,
            hover_color=_GREEN_HOV,
            text_color="#0b0f19",
            height=40,
            corner_radius=12,
            font=("Segoe UI", 16, "bold"),
        )
        self.start_btn.grid(row=4, column=0, pady=(4, 8), sticky="ew")

        self.progress_bar = ctk.CTkProgressBar(
            container, progress_color=_ACCENT, fg_color=_ENTRY_BG, corner_radius=6, height=8
        )
        self.progress_bar.grid(row=5, column=0, sticky="ew", pady=(0, 8))
        self.progress_bar.set(0)

        # ──── Log terminal ────
        log_card = self._make_card(container)
        log_card.grid(row=6, column=0, sticky="nsew", pady=(0, 0))
        container.grid_rowconfigure(6, weight=1)

        self.log_box = ctk.CTkTextbox(
            log_card, state="disabled", font=("Cascadia Code", 12),
            fg_color=_LOG_BG, text_color="#a3e635", corner_radius=8, border_width=0
        )
        self.log_box.pack(fill="both", expand=True, padx=10, pady=10)

    def draw_extensions(self):
        for widget in self.ext_inner_frame.winfo_children():
            widget.destroy()
        col = 0
        row = 0
        for ext in self.available_extensions:
            if ext not in self.extension_vars:
                self.extension_vars[ext] = ctk.BooleanVar(value=True if ext in ['.nc', '.pdf', '.dwg'] else False)
            cb = ctk.CTkCheckBox(
                self.ext_inner_frame, text=ext, variable=self.extension_vars[ext],
                fg_color=_ACCENT, hover_color=_ACCENT_HOV, border_color=_CARD_BORDER,
                text_color=_TEXT_PRI, font=("Segoe UI", 13), corner_radius=4,
                checkbox_width=20, checkbox_height=20,
            )
            cb.grid(row=row, column=col, padx=(0, 18), pady=5, sticky="w")
            col += 1
            if col > 4:
                col = 0
                row += 1

    def add_extension(self):
        new_ext = self.new_ext_entry.get().strip().lower()
        if new_ext:
            if not new_ext.startswith('.'):
                new_ext = '.' + new_ext
            if new_ext not in self.available_extensions:
                self.available_extensions.append(new_ext)
                self.extension_vars[new_ext] = ctk.BooleanVar(value=True)
                self.draw_extensions()
            self.new_ext_entry.delete(0, 'end')

    def browse_source(self):
        folder = filedialog.askdirectory()
        if folder:
            self.source_dir.set(folder)

    def browse_dest(self):
        folder = filedialog.askdirectory()
        if folder:
            self.dest_dir.set(folder)

    def browse_excel(self):
        file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file:
            self.excel_file.set(file)

    def log(self, message):
        self.after(0, self._insert_log, message)

    def _insert_log(self, message):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", message + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def set_progress(self, value):
        self.after(0, self.progress_bar.set, value)

    def extraction_finished(self):
        self.after(0, lambda: self.start_btn.configure(state="normal"))

    def get_part_regex(self, part_name):
        clean_chars = re.sub(r'[-_\s\.]', '', str(part_name))
        escaped_chars = [re.escape(c) for c in clean_chars]
        pattern = r'[\s\-_\.]*'.join(escaped_chars)
        full_pattern = r'(?i)^' + pattern + r'(?:[\s\-_\.]+|$)'
        return re.compile(full_pattern)

    def natural_keys(self, text):
        return [int(c) if c.isdigit() else c for c in re.split(r'(\d+)', text)]

    def start_extraction(self):
        if not self.source_dir.get() or not os.path.isdir(self.source_dir.get()):
            messagebox.showerror("Eroare", "Director sursă invalid")
            return
        if not self.dest_dir.get():
            messagebox.showerror("Eroare", "Director destinație invalid")
            return
        make_dir_with_permissions(self.dest_dir.get())
        if not self.excel_file.get() or not os.path.isfile(self.excel_file.get()):
            messagebox.showerror("Eroare", "Fișier Excel invalid")
            return
        
        selected_exts = []
        for ext, var in self.extension_vars.items():
            if var.get():
                selected_exts.append(ext)
                if ext == '.nc':
                    selected_exts.append('.nc1')
        if not selected_exts:
            messagebox.showerror("Eroare", "Selectați cel puțin o extensie")
            return

        self.save_settings()
        self.start_btn.configure(state="disabled")
        self.set_progress(0)
        self.log("=" * 50)
        self.log("Extracția a început...")
        
        threading.Thread(target=self.run_extraction_process, args=(selected_exts,), daemon=True).start()

    def _detect_column(self, df, keywords, fallback_idx=None):
        """Auto-detect a column by searching for keywords in the header."""
        for col in df.columns:
            col_lower = str(col).lower().strip()
            if any(kw in col_lower for kw in keywords):
                return col
        if fallback_idx is not None and fallback_idx < len(df.columns):
            return df.columns[fallback_idx]
        return None

    def _generate_bom_excel(self, dest_path, bom_data):
        """
        Generate a BOM Excel as a standard continuous list.
        bom_data: list of dicts with keys 'Denumire', 'Cantitate', 'Introduse', 'Ramase'
        """
        wb = Workbook()
        ws = wb.active
        ws.title = "BOM - Piese Ramase"

        headers = ["Nr.", "Denumire Piesă", "Cant.", "Introd.", "Rămase"]
        col_widths = [5.5, 24, 7, 7, 7]

        thin_side = Side(style='thin')
        thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
        header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        header_font_white = Font(bold=True, size=9, name="Segoe UI", color="FFFFFF")
        data_font = Font(size=9, name="Segoe UI")
        alt_fill = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
        warn_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        center_align = Alignment(horizontal='center', vertical='center')
        left_align = Alignment(horizontal='left', vertical='center')

        # Headers
        for h_idx, header in enumerate(headers):
            cell = ws.cell(row=1, column=h_idx + 1, value=header)
            cell.font = header_font_white
            cell.border = thin_border
            cell.fill = header_fill
            cell.alignment = center_align

        # Data
        for data_idx, item in enumerate(bom_data):
            row_num = 2 + data_idx
            use_alt = (data_idx % 2 == 1)

            # Nr.
            cell = ws.cell(row=row_num, column=1, value=data_idx + 1)
            cell.font = data_font; cell.border = thin_border; cell.alignment = center_align
            if use_alt: cell.fill = alt_fill

            # Denumire
            cell = ws.cell(row=row_num, column=2, value=item["Denumire"])
            cell.font = data_font; cell.border = thin_border; cell.alignment = left_align
            if use_alt: cell.fill = alt_fill

            # Cantitate totală
            cell = ws.cell(row=row_num, column=3, value=item["Cantitate"])
            cell.font = data_font; cell.border = thin_border; cell.alignment = center_align
            if use_alt: cell.fill = alt_fill

            # Introduse
            cell = ws.cell(row=row_num, column=4, value=item["Introduse"])
            cell.font = data_font; cell.border = thin_border; cell.alignment = center_align
            if use_alt: cell.fill = alt_fill

            # Rămase (highlight)
            cell = ws.cell(row=row_num, column=5, value=item["Ramase"])
            cell.font = Font(size=9, name="Segoe UI", bold=True)
            cell.border = thin_border; cell.alignment = center_align
            cell.fill = warn_fill

        for w_idx, width in enumerate(col_widths):
            ws.column_dimensions[get_column_letter(w_idx + 1)].width = width

        ws.freeze_panes = "A2"
        ws.print_title_rows = '1:1'

        bom_path = os.path.join(dest_path, "BOM_piese_ramase.xlsx")
        return safe_save_workbook(wb, bom_path, self.log)

    def run_extraction_process(self, selected_exts):
        try:
            self.log(f"Citire fișier Excel: {self.excel_file.get()}")
            df = pd.read_excel(self.excel_file.get())
            self.log(f"Coloane găsite: {list(df.columns)}")

            # ── Auto-detect the 3 columns ──
            name_keywords = ["denumire", "piesa", "parte", "mark", "name", "poz", "element"]
            qty_keywords = ["cantitat", "qty", "quantity", "buc", "pieces", "stk"]
            plan_keywords = ["plan", "croi", "introduse", "nested", "nesting"]

            if self.col_name.get().strip():
                col_name = self.col_name.get().strip()
                if col_name not in df.columns:
                    self.log(f"Eroare: Coloana '{col_name}' nu a fost găsită în Excel.")
                    self.extraction_finished()
                    return
            else:
                col_name = self._detect_column(df, name_keywords, fallback_idx=0)

            col_qty = self._detect_column(df, qty_keywords, fallback_idx=1)
            col_plan = self._detect_column(df, plan_keywords, fallback_idx=2 if len(df.columns) >= 3 else None)

            self.log(f"  Col. Denumire: '{col_name}'")
            self.log(f"  Col. Cantitate: '{col_qty}'")
            self.log(f"  Col. Plan Croi: '{col_plan}'")

            # ── Read part names ──
            if col_name and col_name in df.columns:
                all_part_names = df[col_name].dropna().astype(str).tolist()
            else:
                all_part_names = df.iloc[:, 0].dropna().astype(str).tolist()

            # ── Filter parts: Introduse < Cantitate ──
            bom_data = []
            filtered_parts = set()

            if col_qty and col_plan and col_qty in df.columns and col_plan in df.columns:
                self.log("\n── Verificare Plan de Croi vs Cantitate ──")
                skipped = 0
                for idx in df.index:
                    part = str(df[col_name][idx]).strip() if col_name in df.columns else str(df.iloc[idx, 0]).strip()
                    if not part or part == 'nan':
                        continue
                    try:
                        qty_total = int(float(df[col_qty][idx]))
                    except (ValueError, TypeError):
                        qty_total = 0
                    try:
                        qty_plan = int(float(df[col_plan][idx]))
                    except (ValueError, TypeError):
                        qty_plan = 0

                    if qty_plan < qty_total:
                        remaining = qty_total - qty_plan
                        filtered_parts.add(part)
                        bom_data.append({
                            "Denumire": part,
                            "Cantitate": qty_total,
                            "Introduse": qty_plan,
                            "Ramase": remaining,
                        })
                        self.log(f"  ⚠ {part}: Cantitate={qty_total}, Introduse={qty_plan}, Rămase={remaining}")
                    else:
                        skipped += 1

                self.log(f"\n  Total piese cu deficit: {len(filtered_parts)}")
                self.log(f"  Piese complete (skip): {skipped}")

                if self.extract_mode.get() == "Extrage Tot":
                    self.log("\n  ℹ Mod 'Extrage Tot' activat - se vor extrage toate piesele.")
                    part_names = all_part_names
                else:
                    if not filtered_parts:
                        self.log("\n✅ Toate piesele sunt complet introduse pe plan de croi! Nu este nevoie de extracție.")
                        self.set_progress(1.0)
                        self.after(0, lambda: messagebox.showinfo("Info", "Toate piesele sunt complet introduse pe plan de croi.\nNu mai rămân piese de extras."))
                        self.extraction_finished()
                        return
                    # Use only filtered parts for extraction
                    part_names = list(filtered_parts)
            else:
                self.log("  ⓘ Nu s-au detectat coloanele Cantitate/Plan Croi — se extrag toate piesele.")
                part_names = all_part_names

            part_regexes = {p: self.get_part_regex(p) for p in set(part_names)}
            self.log(f"\nPiese de extras: {len(part_regexes)}")
            self.log("Scanare director sursă...")
            
            files_db = {}
            source_path = Path(self.source_dir.get())
            
            all_files = []
            for root, _, files in os.walk(source_path):
                for f in files:
                    all_files.append(Path(root) / f)
                    
            total_files = len(all_files)
            self.log(f"Total fișiere găsite: {total_files}")
            
            for idx, filepath in enumerate(all_files):
                if idx % 500 == 0:
                    self.set_progress((idx / total_files) * 0.4) 
                    
                ext = filepath.suffix.lower()
                if ext not in selected_exts:
                    continue
                    
                stem = filepath.stem
                match = re.match(r'(?i)^(.*?)(?:[\s\-_]+rev[\s\-_]*([a-z0-9\.]+))?$', stem)
                if match:
                    file_base_name = match.group(1).strip()
                    revision = match.group(2).lower() if match.group(2) else ''
                else:
                    file_base_name = stem.strip()
                    revision = ''
                    
                mod_time = filepath.stat().st_mtime
                
                matched_part = None
                for part, regex in part_regexes.items():
                    if regex.match(file_base_name):
                        matched_part = part
                        break
                
                if matched_part:
                    if matched_part not in files_db:
                        files_db[matched_part] = {}
                    if file_base_name not in files_db[matched_part]:
                        files_db[matched_part][file_base_name] = {}
                    if ext not in files_db[matched_part][file_base_name]:
                        files_db[matched_part][file_base_name][ext] = []
                        
                    files_db[matched_part][file_base_name][ext].append((revision, filepath, mod_time))
            
            self.set_progress(0.4)
            self.log("Scanare completă. Filtrare fișiere conform regulilor...")
            
            keep_newest = (self.duplicate_handling.get() == "Keep Newest")
            files_to_copy = []
            
            found_parts_count = 0
            
            part_keys = list(part_regexes.keys())
            for part_idx, part in enumerate(part_keys):
                if part_idx % max(1, len(part_keys) // 10) == 0:
                    self.set_progress(0.4 + (part_idx / len(part_keys)) * 0.2)
                
                if part in files_db:
                    found_parts_count += 1
                    for base_name, ext_dict in files_db[part].items():
                        for ext, candidates in ext_dict.items():
                            def sort_key(item):
                                rev, fpath, mtime = item
                                has_rev = (rev != '')
                                rev_sort = self.natural_keys(rev)
                                time_sort = mtime if keep_newest else -mtime
                                return (has_rev, rev_sort, time_sort)
                                
                            best_candidate = sorted(candidates, key=sort_key)[-1]
                            files_to_copy.append((part, best_candidate[1]))
                        
            self.set_progress(0.6)
            self.log(f"Găsite {found_parts_count} din {len(part_regexes)} piese.")
            self.log(f"Se copiază {len(files_to_copy)} fișiere...")
            
            dest_path = Path(self.dest_dir.get())
            make_dir_with_permissions(str(dest_path))
            
            part_to_ext_found = {p: set() for p in part_names}
            
            for idx, (orig_part, filepath) in enumerate(files_to_copy):
                if idx % max(1, len(files_to_copy) // 100) == 0:
                    self.set_progress(0.6 + (idx / len(files_to_copy)) * 0.35)
                
                part_to_ext_found[orig_part].add(filepath.suffix.lower())
                dest_file = dest_path / filepath.name
                try:
                    safe_copy(filepath, dest_file)
                    self.log(f"Copiat: {filepath.name}")
                except Exception as e:
                    self.log(f"Eroare copiere {filepath.name}: {str(e)}")

            # ── Generare Raport Status Extracție ──
            self.log("\n── Generare Raport Status Extracție ──")
            try:
                status_path = os.path.join(dest_path, "Raport_Status_Extractie.xlsx")
                wb_status = Workbook()
                ws_status = wb_status.active
                ws_status.title = "Status"
                
                display_exts = [ext for ext, var in self.extension_vars.items() if var.get()]
                
                headers = ["Piesa"] + display_exts
                for col_idx, header in enumerate(headers, 1):
                    ws_status.cell(row=1, column=col_idx, value=header).font = Font(bold=True)
                
                green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
                red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                green_font = Font(color="006100")
                red_font = Font(color="9C0006")
                
                for row_idx, part in enumerate(sorted(part_names), 2):
                    ws_status.cell(row=row_idx, column=1, value=part)
                    for col_idx, ext in enumerate(display_exts, 2):
                        if ext == '.nc':
                            found = '.nc' in part_to_ext_found[part] or '.nc1' in part_to_ext_found[part]
                        else:
                            found = ext in part_to_ext_found[part]
                            
                        cell = ws_status.cell(row=row_idx, column=col_idx, value="OK" if found else "X")
                        cell.fill = green_fill if found else red_fill
                        cell.font = green_font if found else red_font
                        cell.alignment = Alignment(horizontal="center")
                        
                status_path = safe_save_workbook(wb_status, status_path, self.log)
                self.log(f"Raport status generat: {status_path}")
            except Exception as e:
                self.log(f"Eroare generare raport status: {e}")

            # ── Generate BOM Excel for remaining parts ──
            if bom_data:
                self.log("\n── Generare BOM piese rămase ──")
                try:
                    bom_path = self._generate_bom_excel(str(dest_path), bom_data)
                    self.log(f"📊 BOM generat: {bom_path} ({len(bom_data)} piese)")
                except Exception as e:
                    self.log(f"Eroare generare BOM: {e}")

            # ── Asigurare permisiuni citire/scriere complete pe întregul folder destinație ──
            self.log("\n🔒 Aplicare permisiuni complete citire/scriere (toate aplicațiile)...")
            ensure_full_permissions_tree(str(dest_path))
            self.log("✅ Permisiuni asigurate: Toate aplicațiile au acces complet la fișiere și foldere.")

            self.set_progress(1.0)
            self.log("\n✅ Extracție completă cu succes!")
            if bom_data:
                self.after(0, lambda: messagebox.showinfo("Succes", f"Extracție finalizată!\n\n{len(bom_data)} piese cu deficit identificate.\nBOM generat în directorul destinație."))
            else:
                self.after(0, lambda: messagebox.showinfo("Succes", "Extracția s-a finalizat cu succes!"))
            
        except Exception as e:
            self.log(f"A apărut o eroare: {str(e)}")
            self.after(0, lambda err=e: messagebox.showerror("Eroare", f"A apărut o eroare:\n{str(err)}"))
        finally:
            self.extraction_finished()

    # ══════════════════════════════════════════════════════════════════════
    #  NC SORTER TAB
    # ══════════════════════════════════════════════════════════════════════
    def create_nc_sorter_tab(self):
        container = ctk.CTkFrame(self.tab_nc, fg_color="transparent")
        container.pack(fill="both", expand=True, padx=10, pady=10)
        container.grid_columnconfigure(0, weight=1)
        
        subtitle_lbl = ctk.CTkLabel(
            container,
            text="Sortare fișiere NC/DSTV în foldere pe categorii cu rapoarte Excel",
            font=("Segoe UI", 12),
            text_color=_TEXT_SEC,
        )
        subtitle_lbl.grid(row=0, column=0, pady=(0, 8), sticky="w")

        # ──── Source Folder Card ────
        paths_card = self._make_card(container)
        paths_card.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        paths_card.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(paths_card, text="📂 Director Sursă", font=("Segoe UI", 14, "bold"), text_color=_TEXT_PRI).grid(
            row=0, column=0, columnspan=3, padx=16, pady=(10, 6), sticky="w")

        self._label(paths_card, text="Sursă:").grid(row=1, column=0, padx=(16, 6), pady=(6, 14), sticky="w")
        self._entry(paths_card, self.nc_source_folder).grid(row=1, column=1, padx=4, pady=(6, 14), sticky="ew")
        self._browse_btn(paths_card, self.nc_browse_source).grid(row=1, column=2, padx=(4, 16), pady=(6, 14))

        # ──── Excel / BOM Card (Optional) ────
        excel_card = self._make_card(container)
        excel_card.grid(row=2, column=0, sticky="ew", pady=(0, 8))
        excel_card.grid_columnconfigure(1, weight=1)

        excel_header_frame = ctk.CTkFrame(excel_card, fg_color="transparent")
        excel_header_frame.grid(row=0, column=0, columnspan=3, padx=16, pady=(10, 2), sticky="w")
        ctk.CTkLabel(excel_header_frame, text="📊 Excel / BOM", font=("Segoe UI", 14, "bold"), text_color=_TEXT_PRI).pack(side="left")
        ctk.CTkLabel(excel_header_frame, text="  (opțional — pentru raportul cu nr. bucăți)", font=("Segoe UI", 11), text_color=_ORANGE).pack(side="left")

        self._label(excel_card, text="Excel:").grid(row=1, column=0, padx=(16, 6), pady=5, sticky="w")
        self._entry(excel_card, self.nc_excel_file).grid(row=1, column=1, padx=4, pady=5, sticky="ew")
        self._browse_btn(excel_card, self.nc_browse_excel).grid(row=1, column=2, padx=(4, 16), pady=5)

        col_frame = ctk.CTkFrame(excel_card, fg_color="transparent")
        col_frame.grid(row=2, column=0, columnspan=3, padx=16, pady=(2, 12), sticky="w")
        
        self._label(col_frame, text="Col. Piese:", font=("Segoe UI", 11)).pack(side="left", padx=(0, 4))
        self._entry(col_frame, self.nc_part_col, width=100, placeholder_text="auto").pack(side="left", padx=(0, 16))
        self._label(col_frame, text="Col. Cantități:", font=("Segoe UI", 11)).pack(side="left", padx=(0, 4))
        self._entry(col_frame, self.nc_qty_col, width=100, placeholder_text="auto").pack(side="left")

        # ──── Settings Card ────
        settings_card = self._make_card(container)
        settings_card.grid(row=3, column=0, sticky="ew", pady=(0, 8))

        ctk.CTkLabel(settings_card, text="⚙ Setări Sortare", font=("Segoe UI", 14, "bold"), text_color=_TEXT_PRI).pack(
            anchor="w", padx=16, pady=(10, 6))

        # Action row
        action_frame = ctk.CTkFrame(settings_card, fg_color="transparent")
        action_frame.pack(anchor="w", padx=16, pady=(0, 6))
        self._label(action_frame, text="Acțiune:").pack(side="left", padx=(0, 10))
        ctk.CTkRadioButton(action_frame, text="Copiază fișiere", variable=self.nc_action_var, value="Copy",
                           fg_color=_ACCENT, hover_color=_ACCENT_HOV, text_color=_TEXT_PRI).pack(side="left", padx=(0, 16))
        ctk.CTkRadioButton(action_frame, text="Mută fișiere", variable=self.nc_action_var, value="Move",
                           fg_color=_ACCENT, hover_color=_ACCENT_HOV, text_color=_TEXT_PRI).pack(side="left")

        # Options – each on its own row for clarity
        opt_container = ctk.CTkFrame(settings_card, fg_color="transparent")
        opt_container.pack(anchor="w", padx=16, pady=(4, 12))

        ctk.CTkCheckBox(opt_container, text="Sortare profile Laminat la Cald / Rece", variable=self.nc_opt_hot_cold,
                        fg_color=_ACCENT, hover_color=_ACCENT_HOV, border_color=_CARD_BORDER,
                        text_color=_TEXT_PRI, font=("Segoe UI", 12)).pack(anchor="w", pady=3)
        ctk.CTkCheckBox(opt_container, text="Separare Profile de Țevi", variable=self.nc_opt_separate_tubes,
                        fg_color=_ACCENT, hover_color=_ACCENT_HOV, border_color=_CARD_BORDER,
                        text_color=_TEXT_PRI, font=("Segoe UI", 12)).pack(anchor="w", pady=3)
        ctk.CTkCheckBox(opt_container, text="Validare Geometrică (Erori 3D)", variable=self.nc_opt_verify_geometry,
                        fg_color=_ACCENT, hover_color=_ACCENT_HOV, border_color=_CARD_BORDER,
                        text_color=_TEXT_PRI, font=("Segoe UI", 12)).pack(anchor="w", pady=3)
        ctk.CTkCheckBox(opt_container, text="Fără antet (header) în rapoarte Excel", variable=self.nc_opt_no_header,
                        fg_color=_ACCENT, hover_color=_ACCENT_HOV, border_color=_CARD_BORDER,
                        text_color=_TEXT_PRI, font=("Segoe UI", 12)).pack(anchor="w", pady=3)
        ctk.CTkCheckBox(opt_container, text="Fără culori în rapoarte Excel", variable=self.nc_opt_no_colors,
                        fg_color=_ACCENT, hover_color=_ACCENT_HOV, border_color=_CARD_BORDER,
                        text_color=_TEXT_PRI, font=("Segoe UI", 12)).pack(anchor="w", pady=3)

        # ──── Start & Progress ────
        self.nc_start_btn = ctk.CTkButton(
            container,
            text="▶ Start Sortare",
            command=self.nc_start_sorting,
            fg_color=_GREEN,
            hover_color=_GREEN_HOV,
            text_color="#0b0f19",
            height=40,
            corner_radius=12,
            font=("Segoe UI", 16, "bold"),
        )
        self.nc_start_btn.grid(row=4, column=0, pady=(4, 8), sticky="ew")

        self.nc_progress_bar = ctk.CTkProgressBar(
            container, progress_color=_ACCENT, fg_color=_ENTRY_BG, corner_radius=6, height=8
        )
        self.nc_progress_bar.grid(row=5, column=0, sticky="ew", pady=(0, 8))
        self.nc_progress_bar.set(0)

        # ──── Log terminal ────
        log_card = self._make_card(container)
        log_card.grid(row=6, column=0, sticky="nsew", pady=(0, 0))
        container.grid_rowconfigure(6, weight=1)

        self.nc_log_box = ctk.CTkTextbox(
            log_card, state="disabled", font=("Cascadia Code", 12),
            fg_color=_LOG_BG, text_color="#a3e635", corner_radius=8, border_width=0
        )
        self.nc_log_box.pack(fill="both", expand=True, padx=10, pady=10)

    def nc_browse_source(self):
        folder = filedialog.askdirectory(title="Selectare Director Sursă")
        if folder:
            self.nc_source_folder.set(folder)

    def nc_browse_excel(self):
        file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file:
            self.nc_excel_file.set(file)

    def nc_log(self, message):
        self.after(0, self._nc_insert_log, message)

    def _nc_insert_log(self, message):
        self.nc_log_box.configure(state="normal")
        self.nc_log_box.insert("end", message + "\n")
        self.nc_log_box.see("end")
        self.nc_log_box.configure(state="disabled")

    def nc_start_sorting(self):
        src = self.nc_source_folder.get()
        if not src:
            messagebox.showerror("Eroare", "Selectați un director sursă.")
            return
        if not os.path.exists(src):
            messagebox.showerror("Eroare", "Directorul sursă nu există.")
            return

        self.save_settings()
        self.nc_start_btn.configure(state="disabled")
        self.nc_progress_bar.set(0)
        
        self.nc_log_box.configure(state="normal")
        self.nc_log_box.delete("1.0", "end")
        self.nc_log_box.configure(state="disabled")

        excel_file = self.nc_excel_file.get()
        part_col = self.nc_part_col.get().strip()
        qty_col = self.nc_qty_col.get().strip()

        # Daca nu este selectat niciun Excel, verifica daca exista BOM-ul generat anterior
        if not excel_file:
            auto_bom_path = os.path.join(src, "BOM_piese_ramase.xlsx")
            if os.path.exists(auto_bom_path):
                excel_file = auto_bom_path
                self.nc_log(f"Folosesc automat BOM-ul generat anterior: {auto_bom_path}")
                # Setam automat coloanele daca nu au fost completate
                if not part_col:
                    part_col = "Denumire Piesă"
                if not qty_col:
                    qty_col = "Rămase"

        threading.Thread(
            target=self.nc_process_files,
            args=(
                src,
                self.nc_action_var.get(),
                self.nc_opt_hot_cold.get(),
                self.nc_opt_separate_tubes.get(),
                self.nc_opt_verify_geometry.get(),
                self.nc_opt_no_header.get(),
                self.nc_opt_no_colors.get(),
                excel_file,
                part_col,
                qty_col,
            ),
            daemon=True
        ).start()

    # ══════════════════════════════════════════════════════════════════════
    #  MODIFY NC LOGIC & UI
    # ══════════════════════════════════════════════════════════════════════
    def create_mod_nc_tab(self):
        container = ctk.CTkFrame(self.tab_mod, fg_color="transparent")
        container.pack(fill="both", expand=True, padx=10, pady=10)
        container.grid_columnconfigure(0, weight=1)

        subtitle_lbl = ctk.CTkLabel(
            container,
            text="Modifică și ajustează fișierele NC (Text, Cantități, Rotație)",
            font=("Segoe UI", 12),
            text_color=_TEXT_SEC,
        )
        subtitle_lbl.grid(row=0, column=0, pady=(0, 10), sticky="w")

        # ──── Settings Card ────
        settings_card = self._make_card(container)
        settings_card.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        settings_card.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(settings_card, text="⚙ Setări Globale", font=("Segoe UI", 14, "bold"), text_color=_TEXT_PRI).grid(row=0, column=0, columnspan=3, padx=16, pady=(10, 6), sticky="w")
        
        self._label(settings_card, text="Folder NC:").grid(row=1, column=0, padx=(16, 6), pady=6, sticky="w")
        self._entry(settings_card, self.mod_nc_dir).grid(row=1, column=1, padx=4, pady=6, sticky="ew")
        self._browse_btn(settings_card, self.mod_nc_browse_dir).grid(row=1, column=2, padx=(4, 16), pady=6)

        cb_subfolder = ctk.CTkCheckBox(settings_card, text="Creează subfolder 'Modified' (copiază fișierele)", variable=self.mod_nc_create_subfolder)
        cb_subfolder.grid(row=2, column=1, padx=4, pady=(6, 14), sticky="w")

        # ──── Actions Card ────
        actions_card = self._make_card(container)
        actions_card.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        actions_card.grid_columnconfigure(1, weight=1)

        # A. Text Replace
        ctk.CTkLabel(actions_card, text="A. Înlocuire Text", font=("Segoe UI", 14, "bold"), text_color=_TEXT_PRI).grid(row=0, column=0, columnspan=3, padx=16, pady=(10, 6), sticky="w")
        
        self._label(actions_card, text="Caută:").grid(row=1, column=0, padx=(16, 6), pady=6, sticky="w")
        self._entry(actions_card, self.mod_nc_find).grid(row=1, column=1, padx=4, pady=6, sticky="ew")
        self._label(actions_card, text="Înlocuiește:").grid(row=2, column=0, padx=(16, 6), pady=6, sticky="w")
        self._entry(actions_card, self.mod_nc_replace).grid(row=2, column=1, padx=4, pady=6, sticky="ew")
        btn_replace = ctk.CTkButton(actions_card, text="Înlocuiește Text", command=self.mod_nc_replace_text)
        btn_replace.grid(row=1, column=2, rowspan=2, padx=(4, 16), pady=6)

        # B. Rectify Number
        ctk.CTkLabel(actions_card, text="B. Rectificare Cantități din Excel", font=("Segoe UI", 14, "bold"), text_color=_TEXT_PRI).grid(row=3, column=0, columnspan=3, padx=16, pady=(10, 6), sticky="w")
        self._label(actions_card, text="Fișier Excel:").grid(row=4, column=0, padx=(16, 6), pady=6, sticky="w")
        self._entry(actions_card, self.mod_nc_excel).grid(row=4, column=1, padx=4, pady=6, sticky="ew")
        self._browse_btn(actions_card, self.mod_nc_browse_excel).grid(row=4, column=2, padx=(4, 16), pady=6)
        btn_rectify = ctk.CTkButton(actions_card, text="Rectifică Număr", command=self.mod_nc_rectify_qty)
        btn_rectify.grid(row=5, column=1, padx=4, pady=(6, 14), sticky="w")

        # ──── Log box ────
        self.mod_nc_log_box = ctk.CTkTextbox(
            container,
            fg_color=_ENTRY_BG,
            text_color=_TEXT_PRI,
            border_color=_CARD_BORDER,
            border_width=1,
            corner_radius=8,
            font=("Consolas", 12),
        )
        self.mod_nc_log_box.grid(row=3, column=0, sticky="nsew", pady=(0, 0))
        container.grid_rowconfigure(3, weight=1)

    def mod_nc_browse_dir(self):
        d = filedialog.askdirectory(title="Selectează folderul cu fișiere NC")
        if d:
            self.mod_nc_dir.set(d)

    def mod_nc_browse_excel(self):
        f = filedialog.askopenfilename(title="Selectează fișierul Excel cu cantități", filetypes=[("Excel", "*.xlsx *.xls")])
        if f:
            self.mod_nc_excel.set(f)

    def mod_nc_log(self, msg):
        self.mod_nc_log_box.insert("end", msg + "\n")
        self.mod_nc_log_box.see("end")
        self.update()

    def _mod_nc_get_files(self):
        folder = self.mod_nc_dir.get()
        if not folder or not os.path.exists(folder):
            messagebox.showwarning("Atenție", "Te rog selectează un folder valid cu fișiere NC.")
            return None, None
        
        nc_files = [f for f in os.listdir(folder) if f.lower().endswith(('.nc', '.nc1'))]
        if not nc_files:
            self.mod_nc_log("⚠ Nu s-au găsit fișiere NC în folderul selectat.")
            return None, None
            
        out_folder = folder
        if self.mod_nc_create_subfolder.get():
            out_folder = os.path.join(folder, "Modified")
            make_dir_with_permissions(out_folder)
            self.mod_nc_log(f"Subfolder 'Modified' creat: {out_folder}")
        else:
            make_dir_with_permissions(out_folder)
            
        return nc_files, out_folder

    def mod_nc_replace_text(self):
        find_txt = self.mod_nc_find.get()
        rep_txt = self.mod_nc_replace.get()
        if not find_txt:
            messagebox.showwarning("Atenție", "Introdu textul de căutat!")
            return
            
        nc_files, out_folder = self._mod_nc_get_files()
        if not nc_files: return
        
        self.mod_nc_log(f"\n--- Începere Înlocuire Text '{find_txt}' -> '{rep_txt}' ---")
        modified_count = 0
        
        for f in nc_files:
            src_path = os.path.join(self.mod_nc_dir.get(), f)
            dest_path = os.path.join(out_folder, f)
            
            with open(src_path, 'r', encoding='utf-8', errors='ignore') as file:
                lines = file.readlines()
                
            changed = False
            for i in range(len(lines)):
                if lines[i].strip() == find_txt.strip():
                    lines[i] = lines[i].replace(find_txt.strip(), rep_txt.strip())
                    changed = True
                    
            if changed:
                modified_count += 1
                safe_write_lines(dest_path, lines)
                self.mod_nc_log(f"✅ Modificat: {f}")
            elif src_path != dest_path:
                safe_copy(src_path, dest_path)
                
        ensure_full_permissions_tree(out_folder)
        self.mod_nc_log("🔒 Permisiuni asigurate: Toate aplicațiile au acces complet (citire/scriere).")
        self.mod_nc_log(f"Proces finalizat. {modified_count} fișiere modificate.\n")

    def mod_nc_rectify_qty(self):
        excel = self.mod_nc_excel.get()
        if not excel or not os.path.exists(excel):
            messagebox.showwarning("Atenție", "Te rog selectează un fișier Excel valid.")
            return
            
        nc_files, out_folder = self._mod_nc_get_files()
        if not nc_files: return
        
        self.mod_nc_log(f"\n--- Începere Rectificare Cantități din Excel ---")
        self.mod_nc_log(f"Se citește excel-ul: {os.path.basename(excel)}")
        
        try:
            quantity_map = self._build_quantity_map(excel, None, None)
        except Exception as e:
            self.mod_nc_log(f"Eroare citire Excel: {e}")
            return
            
        modified_count = 0
        
        for f in nc_files:
            src_path = os.path.join(self.mod_nc_dir.get(), f)
            dest_path = os.path.join(out_folder, f)
            basename = os.path.splitext(f)[0]
            
            qty = self._lookup_quantity(basename, quantity_map)
            
            with open(src_path, 'r', encoding='utf-8', errors='ignore') as file:
                lines = file.readlines()
                
            if qty is not None and len(lines) > 7:
                lines[7] = f"  {qty}\n"
                safe_write_lines(dest_path, lines)
                self.mod_nc_log(f"✅ {f}: Cantitate actualizată -> {qty}")
                modified_count += 1
            else:
                if qty is None:
                    self.mod_nc_log(f"⚠ {f}: Nu a fost găsit în Excel.")
                if src_path != dest_path:
                    safe_copy(src_path, dest_path)
                    
        ensure_full_permissions_tree(out_folder)
        self.mod_nc_log("🔒 Permisiuni asigurate: Toate aplicațiile au acces complet (citire/scriere).")
        self.mod_nc_log(f"Proces finalizat. {modified_count} fișiere modificate.\n")

    # ══════════════════════════════════════════════════════════════════════
    #  NC SORTING LOGIC
    # ══════════════════════════════════════════════════════════════════════
    def determine_category(self, profile, material, opt_hot_cold=False, opt_separate_tubes=False):
        prof_upper = profile.strip().upper()
        mat_upper = material.strip().upper()

        inox_materials = ["INOX", "1.4301", "304", "316", "1.4404", "1.4307", "1.4571", "1.4401", "1.4016", "A2", "A4"]
        alu_materials = ["ALUMINIU", "AW", "ALU", "5754", "6082", "6060", "7075"]

        # Step 0: Check for Rails (SINE)
        if any(x in prof_upper for x in ["40X60", "60X40", "50X30", "30X50"]):
            return "SINE", "SINA"

        # Step 1: T-URI (Everything starting with T: T, TB, TABLA, T=...)
        if prof_upper.startswith("T"):
            return "T-URI", "T-URI"

        # Step 2: Check for Bars
        is_bar = False
        bar_prefixes = ["RO", "GST", "VKT", "SQ", "PT", "RD", "FI", "ARM"]
        for p in bar_prefixes:
            if prof_upper.startswith(p):
                is_bar = True
                break
        
        if not is_bar and (prof_upper.startswith("D") or prof_upper.startswith("M")):
            rest = prof_upper[1:].strip()
            if rest and (rest[0].isdigit() or rest[0] in ['_', '-']):
                is_bar = True
            elif not rest:
                is_bar = True
                
        if is_bar:
            if any(x in mat_upper for x in inox_materials):
                return os.path.join("BARE", "INOX"), "BARA INOX"
            elif any(x in mat_upper for x in alu_materials):
                return os.path.join("BARE", "ALUMINIU"), "BARA ALUMINIU"
            else:
                return "BARE", "BARA"
        
        # Step 3: Check for Box Girders
        if "HSQ" in prof_upper or "HSQEN" in prof_upper or "CHESON" in prof_upper:
            return "CHESON", "CHESON"
        
        # Step 4: Check for Plates
        is_plate = False
        plate_prefixes = ["PL", "BL", "FB", "FL", "SH", "SHEET"]
        for p in plate_prefixes:
            if prof_upper.startswith(p):
                is_plate = True
                break
                
        if not is_plate and prof_upper.startswith("B") and len(prof_upper) > 1 and prof_upper[1].isdigit():
            is_plate = True

        if not is_plate and len(prof_upper) > 0 and prof_upper[0].isdigit():
            is_plate = True

        if prof_upper.startswith("PL") and len(prof_upper) > 2 and (prof_upper[2].isdigit() or prof_upper[2] in ['-', '_']):
            is_plate = True

        if is_plate:
            if "Z35" in mat_upper:
                return os.path.join("PLACI", "Z35"), "PLACA"
            elif "S355" in mat_upper:
                return os.path.join("PLACI", "S355"), "PLACA"
            elif "S235" in mat_upper:
                return os.path.join("PLACI", "S235"), "PLACA"
            elif any(x in mat_upper for x in inox_materials):
                return os.path.join("PLACI", "INOX"), "PLACA INOX"
            elif any(x in mat_upper for x in alu_materials):
                return os.path.join("PLACI", "ALUMINIU"), "PLACA ALUMINIU"
            else:
                return os.path.join("PLACI", "Other"), "PLACA"

        # Step 5: Profiles (Catch-all)
        tube_hot_prefixes = ["HF", "VK"]
        tube_cold_prefixes = ["KKR", "RHS", "SHS", "CHS", "CFRHS", "KRR", "ROHR", "TUBE", "PIPE", "O"]

        is_tube = False
        is_tube_hot = False

        for p in tube_hot_prefixes:
            if prof_upper.startswith(p):
                is_tube = True
                is_tube_hot = True
                break
                
        if not is_tube:
            for p in tube_cold_prefixes:
                if prof_upper.startswith(p):
                    is_tube = True
                    break

        if is_tube:
            base_folder = "TEVI" if opt_separate_tubes else "PROFILE"
            if opt_hot_cold:
                if is_tube_hot:
                    return os.path.join(base_folder, "LAMINAT LA CALD"), "TEAVA CALD"
                else:
                    return os.path.join(base_folder, "LAMINAT LA RECE"), "TEAVA RECE"
            else:
                return base_folder, "TEAVA"

        # All other standard profiles (HEA, IPE, L, UPN, Z, C, etc.)
        if opt_hot_cold:
            return os.path.join("PROFILE", "LAMINAT LA RECE"), "PROFIL LAMINAT LA RECE"
        else:
            return "PROFILE", "PROFIL"

    def parse_nc_file(self, filepath):
        profile = ""
        material = ""
        
        try:
            with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
                lines = [line.strip() for line in f.readlines()[:15]]
                
                mat_idx = -1
                safe_grades = ["S355", "S235", "Z35", "S275", "S460", "STEEL", "B500", "K500", "INOX", "1.4301", "1.4404", "ALUMINIU", "ALU", "EN AW", "AW-", "S420", "S690", "CORTEN", "PC52", "OB37", "1.4307", "1.4571", "1.4401", "1.4016", "5754", "6082", "6060", "7075", "C45", "42CRMO4"]
                exact_grades = ["304", "316", "304L", "316L", "A2", "A4", "8.8", "10.9"]
                
                for i, line in enumerate(lines):
                    upper_line = line.upper()
                    if i < 4: continue
                    if "**" in upper_line or ".NC" in upper_line: continue
                    
                    if any(grade in upper_line for grade in safe_grades) or upper_line in exact_grades:
                        mat_idx = i
                        break
                
                profile = ""
                if mat_idx != -1:
                    material = lines[mat_idx]
                    
                    candidates = []
                    if mat_idx + 2 < len(lines):
                        candidates.append(lines[mat_idx + 2])
                    if mat_idx > 0:
                        candidates.append(lines[mat_idx - 1])
                        
                    exclude_marks = []
                    if len(lines) > 3: exclude_marks.append(lines[3].upper())
                    if len(lines) > 1: exclude_marks.append(lines[1].upper().replace("**", "").strip())
                    if len(lines) > 2: exclude_marks.append(lines[2].upper())
                    
                    exclude_marks = [m for m in exclude_marks if not m.isdigit()]
                    
                    prefixes = ["PL", "B", "BL", "RO", "D", "O", "HEA", "HEB", "IPE", "UPN", "RHS", "CHS", "SHS", "CFRHS", "HFRHS", "HSQEN", "HSQ", "CHESON", "L", "VKR", "GST", "ARM", "M", "UNP", "UPE", "INP", "FL", "HEM", "HL", "HD", "UB", "UC", "W", "UAP", "PFC", "C", "EA", "UA", "KKR", "KRR", "ROHR", "VKT", "SQ", "PT", "RD", "FI", "T", "TB", "Z", "FB", "FLT", "TABLA", "SHEET", "SH"]
                    prefixes.sort(key=len, reverse=True)
                    
                    for cand in candidates:
                        upper_cand = cand.upper()
                        if upper_cand in exclude_marks:
                            continue
                            
                        is_valid = False
                        if upper_cand and upper_cand[0].isdigit():
                            is_valid = True
                        else:
                            for p in prefixes:
                                if upper_cand.startswith(p):
                                    rest = upper_cand[len(p):].strip()
                                    if not rest or rest[0].isdigit() or rest[0] in ['_', '-', '*', 'M', 'X', '=', 'E', 'N']:
                                        is_valid = True
                                        break
                                        
                        if is_valid:
                            profile = cand
                            break
                            
                    if not profile:
                        profile = lines[mat_idx - 1] if mat_idx > 0 else ""
                
                if mat_idx == -1:
                    safe_fallback_prefixes = ["PL", "BL", "RO", "HEA", "HEB", "IPE", "UPN", "RHS", "CHS", "SHS", "CFRHS", "HFRHS", "HSQEN", "HSQ", "CHESON", "VKR", "GST", "HEM", "HL", "HD", "UB", "UC", "UAP", "PFC", "KKR", "KRR", "ROHR", "VKT", "SQ", "PT", "RD", "FI", "FB", "FLT", "TABLA", "SHEET"]
                    for i, line in enumerate(lines):
                        upper_line = line.upper()
                        if any(upper_line.startswith(p) for p in safe_fallback_prefixes):
                            profile = line
                            if i + 1 < len(lines):
                                material = lines[i + 1]
                            break
                    
                    if not profile and len(lines) >= 6:
                        profile = lines[3] if len(lines) > 3 else ""
                        material = lines[5] if len(lines) > 5 else ""
                        
        except Exception as e:
            self.nc_log(f"Eroare citire {os.path.basename(filepath)}: {e}")
            
        return profile, material

    def verify_geometry(self, filepath, subfolder):
        try:
            with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
                lines = [l.strip() for l in f.readlines()]
                
            is_plate = "PLACI" in subfolder.upper()
            
            if is_plate:
                in_op_block = False
                for line in lines:
                    if line in ["AK", "IK", "BO", "PU", "KO"]:
                        in_op_block = True
                        continue
                    if line == "EN":
                        in_op_block = False
                        continue
                        
                    if in_op_block and len(line) > 0:
                        face = line[0].lower()
                        if face in ['o', 'u']:
                            return False, "Placă cu operațiuni 3D (fețe o/u) - posibil profil exportat greșit"

            if not is_plate and len(lines) > 9:
                shape_code = lines[9].upper()
                if shape_code == "B":
                    return False, "Profil/Țeavă exportată ca Placă 2D desfășurată (Shape Code 'B')"

        except Exception:
            pass
            
        return True, ""

    # ══════════════════════════════════════════════════════════════════════
    #  QUANTITY MAP & LOOKUP
    # ══════════════════════════════════════════════════════════════════════
    @staticmethod
    def _normalize_name(name):
        """Remove common separators for fuzzy part-name matching."""
        return re.sub(r'[-_\s\.]', '', str(name)).upper()

    def _build_quantity_map(self, excel_path, part_col_name, qty_col_name):
        """
        Build a dict { normalized_part_name: quantity } from the BOM Excel.
        Returns an empty dict on failure.
        """
        try:
            # First, check if explicit columns are given
            if part_col_name or qty_col_name:
                df = pd.read_excel(excel_path)
                part_series = df[part_col_name] if (part_col_name and part_col_name in df.columns) else df.iloc[:, 0]
                qty_series = df[qty_col_name] if (qty_col_name and qty_col_name in df.columns) else df.iloc[:, 1]
                
                result = {}
                for idx in part_series.index:
                    pname = str(part_series[idx]).strip()
                    if not pname or pname.lower() == 'nan': continue
                    norm = self._normalize_name(pname)
                    try:
                        qty = int(float(qty_series[idx]))
                    except (ValueError, TypeError):
                        qty = 0
                    result[norm] = result.get(norm, 0) + qty
                return result

            # If NO explicit columns, use generic scan to support multi-section layouts
            df_raw = pd.read_excel(excel_path, header=None)
            result = {}
            qty_keywords = ['qty', 'quantity', 'buc', 'cantitate', 'cant', 'nr', 'count', 'stk', 'anz', 'pcs', 'pieces', 'bucati', 'ramase']
            part_keywords = ['denumire', 'pies', 'part', 'name', 'nume', 'reper']
            
            num_cols = df_raw.shape[1]
            used_cols = set()
            valid_pairs = []
            
            for c in range(num_cols - 1):
                if c in used_cols or (c+1) in used_cols:
                    continue
                
                valid_rows = 0
                is_qty_header = False
                
                # Check for headers in the first 5 rows
                for r in range(min(5, df_raw.shape[0])):
                    val_q_clean = ''.join(ch for ch in str(df_raw.iloc[r, c+1]).lower() if ch.isalpha())
                    if any(kw in val_q_clean for kw in qty_keywords):
                        is_qty_header = True
                        break
                
                # Check for data rows
                for r in range(min(20, df_raw.shape[0])):
                    p = str(df_raw.iloc[r, c]).strip()
                    q = df_raw.iloc[r, c+1]
                    if p and p.lower() != 'nan' and not any(kw in p.lower() for kw in part_keywords + qty_keywords):
                        try:
                            int(float(q))
                            valid_rows += 1
                        except (ValueError, TypeError):
                            pass
                
                # Valid if it's the first pair (0,1), has a known header, or has at least 1 row of valid data
                if c == 0 or is_qty_header or valid_rows >= 1:
                    valid_pairs.append((c, c+1))
                    used_cols.add(c)
                    used_cols.add(c+1)
            
            for p_idx, q_idx in valid_pairs:
                for r in range(df_raw.shape[0]):
                    pname = str(df_raw.iloc[r, p_idx]).strip()
                    qty_val = df_raw.iloc[r, q_idx]
                    
                    if not pname or pname.lower() == 'nan': continue
                    if any(kw in pname.lower() for kw in part_keywords + qty_keywords): continue
                    
                    try:
                        qty = int(float(qty_val))
                    except (ValueError, TypeError):
                        continue
                        
                    norm = self._normalize_name(pname)
                    result[norm] = result.get(norm, 0) + qty
                    
            return result
            
        except Exception as e:
            self.nc_log(f"  Nu s-a putut citi Excel-ul BOM: {e}")
            return {}

    def _lookup_quantity(self, nc_basename, quantity_map):
        """Look up the quantity for a part from the quantity map."""
        norm = self._normalize_name(nc_basename)
        if norm in quantity_map:
            return quantity_map[norm]
        return None

    # ══════════════════════════════════════════════════════════════════════
    #  PER-CATEGORY EXCEL GENERATION
    # ══════════════════════════════════════════════════════════════════════
    def _get_major_category(self, subfolder):
        """Extract the top-level category from a subfolder path."""
        parts = subfolder.replace("/", os.sep).replace("\\", os.sep).split(os.sep)
        return parts[0] if parts else subfolder

    def generate_category_excels(self, out_dir, report_data, quantity_map, no_header, no_colors):
        """Generate one formatted Excel per major category folder, split into sheets by material."""
        import re
        
        def sanitize_sheet_name(name):
            # Excel sheet names max length is 31, and cannot contain \ / * ? : [ ]
            if not name:
                return "Nedefinit"
            safe_name = re.sub(r'[\\/\*\?\:\[\]]', '_', str(name))
            return safe_name[:31]

        # Group by major category
        categories = {}
        for item in report_data:
            major = self._get_major_category(item["Cale Destinatie"])
            if major not in categories:
                categories[major] = {}
            
            # Further group by material
            mat = sanitize_sheet_name(item.get("Material Extras", "Nedefinit"))
            if mat not in categories[major]:
                categories[major][mat] = []
            categories[major][mat].append(item)

        for cat_name, materials_dict in categories.items():
            cat_dir = os.path.join(out_dir, cat_name)
            make_dir_with_permissions(cat_dir)

            sheets_data = {}
            total_parts = 0
            
            for mat_name, items in materials_dict.items():
                # Build unique-part list (deduplicate by filename stem)
                seen = {}
                for item in items:
                    nc_base = os.path.splitext(item["Nume Fisier"])[0]
                    if nc_base not in seen:
                        qty = self._lookup_quantity(nc_base, quantity_map)
                        row = {
                            "Denumire": nc_base,
                            "Bucati": qty if qty is not None else "-",
                        }
                        seen[nc_base] = row
                
                excel_data = sorted(seen.values(), key=lambda x: x["Denumire"])
                if excel_data:
                    sheets_data[mat_name] = excel_data
                    total_parts += len(excel_data)

            if not sheets_data:
                continue

            excel_path = os.path.join(cat_dir, f"raport_{cat_name}.xlsx")
            try:
                self.write_a4_excel(excel_path, sheets_data, no_header, no_colors)
                self.nc_log(f"  📊 Raport categorie generat: raport_{cat_name}.xlsx ({total_parts} piese, {len(sheets_data)} materiale)")
            except Exception as e:
                self.nc_log(f"  Eroare generare raport {cat_name}: {e}")

    def write_a4_excel(self, filepath, sheets_data, no_header, no_colors):
        """
        Write an A4-optimized Excel file with multi-section layout for each sheet.
        sheets_data: dict of { sheet_name: [ list of item dicts ] }
        """
        wb = Workbook()
        # Remove default sheet
        if "Sheet" in wb.sheetnames:
            wb.remove(wb["Sheet"])
            
        ROWS_PER_SECTION = 34
        headers = ["Denumire Piesă", "Cant."]
        col_widths = [24, 7]
        SECTIONS_PER_PAGE = 4
        COLS_PER_SECTION = len(headers)
        SPACER_WIDTH = 1.5

        # ── Styles ──
        thin_side = Side(style='thin')
        thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
        header_font_white = Font(bold=True, size=9, name="Segoe UI", color="FFFFFF") if not no_colors else Font(bold=True, size=9, name="Segoe UI")
        data_font = Font(size=9, name="Segoe UI")
        header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid") if not no_colors else PatternFill()
        alt_fill = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid") if not no_colors else PatternFill()
        center_align = Alignment(horizontal='center', vertical='center')
        left_align = Alignment(horizontal='left', vertical='center')

        for sheet_name, data in sheets_data.items():
            ws = wb.create_sheet(title=sheet_name)
            
            total_items = len(data)
            total_sections = max(1, (total_items + ROWS_PER_SECTION - 1) // ROWS_PER_SECTION)

            for sec_idx in range(total_sections):
                col_start = sec_idx * (COLS_PER_SECTION + 1) + 1  # +1 spacer column
                row_offset = 0

                # ── Header row ──
                if not no_header:
                    for h_idx, header in enumerate(headers):
                        cell = ws.cell(row=1, column=col_start + h_idx, value=header)
                        cell.font = header_font_white
                        cell.border = thin_border
                        if not no_colors:
                            cell.fill = header_fill
                        cell.alignment = center_align
                    row_offset = 1

                # ── Data rows ──
                sec_start = sec_idx * ROWS_PER_SECTION
                sec_end = min(sec_start + ROWS_PER_SECTION, total_items)

                for data_idx in range(sec_start, sec_end):
                    row_num = row_offset + 1 + (data_idx - sec_start)
                    item = data[data_idx]
                    use_alt = (data_idx - sec_start) % 2 == 1

                    # Denumire
                    cell = ws.cell(row=row_num, column=col_start, value=item["Denumire"])
                    cell.font = data_font
                    cell.border = thin_border
                    cell.alignment = left_align
                    if use_alt and not no_colors: cell.fill = alt_fill

                    # Bucăți
                    cell = ws.cell(row=row_num, column=col_start + 1, value=item["Bucati"])
                    cell.font = data_font
                    cell.border = thin_border
                    cell.alignment = center_align
                    if use_alt and not no_colors: cell.fill = alt_fill

                # ── Column dimensions ──
                for w_idx, width in enumerate(col_widths):
                    ws.column_dimensions[get_column_letter(col_start + w_idx)].width = width
                
                # Spacer column dimension
                if sec_idx < total_sections - 1:
                    ws.column_dimensions[get_column_letter(col_start + COLS_PER_SECTION)].width = SPACER_WIDTH

            # ── Page setup ──
            ws.page_setup.paperSize = ws.PAPERSIZE_A4
            ws.page_setup.orientation = 'landscape'
            ws.page_margins.left = 0.35
            ws.page_margins.right = 0.35
            ws.page_margins.top = 0.45
            ws.page_margins.bottom = 0.45
            
            ws.sheet_properties.pageSetUpPr.fitToPage = True
            num_pages_wide = max(1, (total_sections + SECTIONS_PER_PAGE - 1) // SECTIONS_PER_PAGE)
            ws.page_setup.fitToWidth = num_pages_wide
            ws.page_setup.fitToHeight = 1
            
            ws.freeze_panes = "A2"
            ws.print_title_rows = '1:1'

        # Fallback if no sheets were added
        if not wb.sheetnames:
            wb.create_sheet(title="No Data")

        safe_save_workbook(wb, filepath, log_callback=self.nc_log)

    def generate_project_summary_excel(self, out_dir, report_data, quantity_map, no_header, no_colors):
        groups = {}
        for item in report_data:
            sec = str(item.get("Sectiune", "Necunoscut")).strip()
            if not sec: sec = "Necunoscut"
            import re
            safe_sec = re.sub(r'[\\/\*\?\:\[\]]', '_', sec)[:31]
            if not safe_sec: safe_sec = "Necunoscut"
            
            if safe_sec not in groups:
                groups[safe_sec] = []
                
            nc_base = os.path.splitext(item["Nume Fisier"])[0]
            qty = self._lookup_quantity(nc_base, quantity_map)
            
            groups[safe_sec].append({
                "Denumire": nc_base,
                "Material": item["Material Extras"],
                "Bucati": qty if qty is not None else "-"
            })
            
        wb = Workbook()
        if "Sheet" in wb.sheetnames:
            wb.remove(wb["Sheet"])
            
        header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid") if not no_colors else PatternFill()
        header_font = Font(bold=True, color="FFFFFF") if not no_colors else Font(bold=True)
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        
        for sheet_name in sorted(groups.keys()):
            ws = wb.create_sheet(title=sheet_name)
            
            row_idx = 1
            if not no_header:
                headers = ["Denumire Piesă", "Material", "Bucăți"]
                for col_idx, h in enumerate(headers, 1):
                    cell = ws.cell(row=row_idx, column=col_idx, value=h)
                    if not no_colors:
                        cell.fill = header_fill
                    cell.font = header_font
                    cell.border = thin_border
                row_idx += 1
                
            data = sorted(groups[sheet_name], key=lambda x: x["Denumire"])
            for row_data in data:
                cell1 = ws.cell(row=row_idx, column=1, value=row_data["Denumire"])
                cell2 = ws.cell(row=row_idx, column=2, value=row_data["Material"])
                cell3 = ws.cell(row=row_idx, column=3, value=row_data["Bucati"])
                for c in [cell1, cell2, cell3]:
                    c.border = thin_border
                row_idx += 1
                
            ws.column_dimensions['A'].width = 25
            ws.column_dimensions['B'].width = 15
            ws.column_dimensions['C'].width = 10
            
        if not wb.sheetnames:
            wb.create_sheet("Empty")
            
        report_path = os.path.join(out_dir, "raport_grosimi_si_profile.xlsx")
        try:
            report_path = safe_save_workbook(wb, report_path, log_callback=self.nc_log)
            self.nc_log(f"Raport centralizator grosimi/profile generat: {report_path}")
        except Exception as e:
            self.nc_log(f"Eroare generare raport centralizator: {e}")

    # ══════════════════════════════════════════════════════════════════════
    #  NC PROCESS FILES (main worker)
    # ══════════════════════════════════════════════════════════════════════
    def nc_process_files(self, src, action, opt_hot_cold, opt_separate_tubes, opt_verify_geometry,
                         opt_no_header, opt_no_colors, excel_path, part_col, qty_col):
        out = os.path.join(src, "Sorted")
        make_dir_with_permissions(out)
        self.nc_log(f"Începere sortare din {src} spre {out}")
        self.nc_log(f"Acțiune: {action}\n")

        # ── Build quantity map if Excel provided ──
        quantity_map = {}
        if excel_path and os.path.isfile(excel_path):
            self.nc_log("Citire Excel BOM pentru cantități...")
            quantity_map = self._build_quantity_map(excel_path, part_col, qty_col)
            self.nc_log(f"  Hartă cantități: {len(quantity_map)} piese identificate.\n")

        # ── Collect NC files ──
        nc_files = []
        for root, _, files in os.walk(src):
            # Skip the output folder itself
            if os.path.normpath(root).startswith(os.path.normpath(out)):
                continue
            for file in files:
                if file.lower().endswith(('.nc', '.nc1')):
                    nc_files.append(os.path.join(root, file))

        total_files = len(nc_files)
        if total_files == 0:
            self.nc_log("Nu au fost găsite fișiere .nc în directorul sursă.")
            self.after(0, lambda: self.nc_start_btn.configure(state="normal"))
            return

        report_data = []

        for i, filepath in enumerate(nc_files):
            filename = os.path.basename(filepath)
            profile, material = self.parse_nc_file(filepath)
            
            self.nc_log(f"[{filename}] Profil: '{profile}', Material: '{material}'")

            subfolder, elem_type = self.determine_category(profile, material, opt_hot_cold, opt_separate_tubes)
            
            if opt_verify_geometry:
                is_valid, reason = self.verify_geometry(filepath, subfolder)
                if not is_valid:
                    self.nc_log(f"  [!] EROARE GEOMETRIE: {reason} -> Mutat la UNSORTED")
                    subfolder = "UNSORTED"
                    elem_type = "EROARE"

            dest_dir = os.path.join(out, subfolder)
            make_dir_with_permissions(dest_dir)
            dest_path = os.path.join(dest_dir, filename)
            
            try:
                if action == "Copy":
                    safe_copy(filepath, dest_path)
                    act_text = "Copiat"
                else:
                    safe_move(filepath, dest_path)
                    act_text = "Mutat"
                
                self.nc_log(f"{act_text} {filename} -> {subfolder}")
                
                # --- Find and Move Associated Files (PDFs, DWGs, etc) ---
                nc_base = os.path.splitext(filename)[0]
                nc_id = re.sub(r'[^a-zA-Z0-9]', '', nc_base).upper()
                if nc_id:
                    regex_str = "^" + "[^a-zA-Z0-9]*".join(list(nc_id)) + r"(?![a-zA-Z0-9])"
                    pattern = re.compile(regex_str, re.IGNORECASE)
                    
                    parent_dir = os.path.dirname(filepath)
                    for sib in os.listdir(parent_dir):
                        if sib == filename or sib.lower().endswith(('.nc', '.nc1')):
                            continue
                            
                        sib_base = os.path.splitext(sib)[0]
                        if pattern.match(sib_base):
                            sib_path = os.path.join(parent_dir, sib)
                            if not os.path.isfile(sib_path):
                                continue
                                
                            sib_dest_path = os.path.join(dest_dir, sib)
                            try:
                                if action == "Copy":
                                    safe_copy(sib_path, sib_dest_path)
                                else:
                                    safe_move(sib_path, sib_dest_path)
                                
                                self.nc_log(f"  + {act_text} fișier asociat: {sib}")
                            except Exception as e:
                                self.nc_log(f"  Eroare {action.lower()} fișier asociat {sib}: {e}")
                # --------------------------------------------------------
                
                report_data.append({
                    "Nume Fisier": filename,
                    "Tip Element": elem_type,
                    "Sectiune": profile,
                    "Material Extras": material,
                    "Cale Destinatie": subfolder
                })
            except Exception as e:
                self.nc_log(f"Eroare {action.lower()} {filename}: {e}")

            progress_val = (i + 1) / total_files
            self.after(0, self.nc_progress_bar.set, progress_val)

        # ── Generate master Excel Report ──
        self.nc_log("\nGenerare raport Excel principal...")
        report_path = os.path.join(out, "raport_piese_nc.xlsx")
        try:
            df = pd.DataFrame(report_data)
            report_path = safe_df_to_excel(df, report_path, index=False, log_callback=self.nc_log)
            self.nc_log(f"Raport generat: {report_path}")
        except Exception as e:
            self.nc_log(f"Eroare generare raport: {e}")

        # ── Generate per-category Excel reports ──
        self.nc_log("\nGenerare rapoarte Excel pe categorii...")
        try:
            self.generate_category_excels(out, report_data, quantity_map, opt_no_header, opt_no_colors)
        except Exception as e:
            self.nc_log(f"Eroare generare rapoarte categorii: {e}")

        # ── Generate Master Project Summary ──
        self.nc_log("\nGenerare raport centralizator proiect...")
        try:
            self.generate_project_summary_excel(out, report_data, quantity_map, opt_no_header, opt_no_colors)
        except Exception as e:
            self.nc_log(f"Eroare generare raport centralizator: {e}")

        # ── Asigurare permisiuni citire/scriere complete pe întregul folder Sorted ──
        self.nc_log("\n🔒 Aplicare permisiuni complete citire/scriere (toate aplicațiile)...")
        ensure_full_permissions_tree(out)
        self.nc_log("✅ Permisiuni asigurate: Toate aplicațiile au acces complet la fișiere și foldere.")

        self.nc_log("\n✅ Sortare completă!")
        self.after(0, lambda: self.nc_start_btn.configure(state="normal"))

if __name__ == "__main__":
    import sys
    # On Windows, tell the OS this is a distinct app so the taskbar icon updates correctly
    if sys.platform.startswith("win"):
        try:
            import ctypes
            myappid = f"steelsync.toolkit.app.{APP_VERSION}"
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        except Exception:
            pass

    app = UnifiedApp()
    app.mainloop()
