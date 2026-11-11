import os
import re
import shutil
import threading
import tkinter as tk
from pdf_to_dwg import PdfToCadConverter
import metalix_importer
import dwg_organizer
from tkinter import filedialog, messagebox, ttk
# pyrefly: ignore [missing-import]
import customtkinter as ctk
from pathlib import Path
import sys
import json
import stat
import subprocess
import tempfile
import urllib.request
from collections import Counter

# ─── High-performance lazy imports for instant startup ───────────────────────
_pd = None

def get_pd():
    global _pd
    if _pd is None:
        import pandas as __pd
        _pd = __pd
    return _pd

class _LazyPandasProxy:
    def __getattr__(self, name):
        return getattr(get_pd(), name)

pd = _LazyPandasProxy()

def _load_openpyxl():
    global Workbook, load_workbook, Font, Border, Side, Alignment, PatternFill, get_column_letter
    from openpyxl import Workbook as _WB, load_workbook as _LWB
    from openpyxl.styles import Font as _F, Border as _B, Side as _S, Alignment as _A, PatternFill as _PF
    from openpyxl.utils import get_column_letter as _GCL
    Workbook = _WB
    load_workbook = _LWB
    Font = _F
    Border = _B
    Side = _S
    Alignment = _A
    PatternFill = _PF
    get_column_letter = _GCL

def load_workbook(*args, **kwargs):
    _load_openpyxl()
    return globals()['load_workbook'](*args, **kwargs)

def Workbook(*args, **kwargs):
    _load_openpyxl()
    return globals()['Workbook'](*args, **kwargs)

def Font(*args, **kwargs):
    _load_openpyxl()
    return globals()['Font'](*args, **kwargs)

def Border(*args, **kwargs):
    _load_openpyxl()
    return globals()['Border'](*args, **kwargs)

def Side(*args, **kwargs):
    _load_openpyxl()
    return globals()['Side'](*args, **kwargs)

def Alignment(*args, **kwargs):
    _load_openpyxl()
    return globals()['Alignment'](*args, **kwargs)

def PatternFill(*args, **kwargs):
    _load_openpyxl()
    return globals()['PatternFill'](*args, **kwargs)

def get_column_letter(*args, **kwargs):
    _load_openpyxl()
    return globals()['get_column_letter'](*args, **kwargs)

def _preload_heavy_libraries():
    """Silently preloads pandas and openpyxl in background once UI is rendered."""
    try:
        get_pd()
        _load_openpyxl()
    except Exception:
        pass

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


def extract_nc_part_code(nc_path):
    """Extract piece mark from standard DSTV NC/NC1 file header."""
    try:
        with open(nc_path, 'r', encoding='utf-8', errors='ignore') as f:
            lines = [f.readline().strip() for _ in range(12)]
        for line in lines[2:8]:
            clean = re.sub(r'[^A-Za-z0-9]', '', line).upper()
            if len(clean) >= 2 and any(ch.isalpha() for ch in clean) and any(ch.isdigit() for ch in clean):
                return clean
    except Exception:
        pass
    return None

def extract_pdf_header_part_code(pdf_path, known_keys=None):
    """
    Safely extract part code or title block information from PDF using pypdf.
    Looks for headers like 'STYKKLISTE FOR <COD>', 'POS <COD>', etc.
    """
    try:
        from pypdf import PdfReader
        reader = PdfReader(str(pdf_path))
        if not reader.pages:
            return None
        for page in reader.pages[:2]:
            text = page.extract_text() or ""
            if not text:
                continue
            
            m = re.search(r'STYKKLISTE\s+FOR\s+([A-Za-z0-9\-_]+)', text, re.IGNORECASE)
            if m:
                clean = re.sub(r'[^A-Za-z0-9]', '', m.group(1)).upper()
                if clean:
                    return clean
            
            m = re.search(r'\bPOS\.?\s*[:\-]?\s*([A-Za-z0-9\-_]+)', text, re.IGNORECASE)
            if m:
                clean = re.sub(r'[^A-Za-z0-9]', '', m.group(1)).upper()
                if clean:
                    return clean

            m = re.search(r'(?:TEGNINGSNUMMER|TEGN\.?\s*NR\.?|DRAWING\s*NO\.?)\s*[:\-]?\s*([A-Za-z0-9\-_]+)', text, re.IGNORECASE)
            if m:
                clean = re.sub(r'[^A-Za-z0-9]', '', m.group(1)).upper()
                if clean:
                    return clean
                    
            if known_keys:
                clean_text = re.sub(r'[^A-Za-z0-9]', '', text).upper()
                for k in sorted(known_keys, key=len, reverse=True):
                    pattern = rf'(?<![A-Za-z0-9]){re.escape(k)}(?!\d)'
                    if re.search(pattern, clean_text):
                        return k
    except Exception:
        pass
    return None


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
_RED         = _DANGER
_RED_HOV     = ("#d70015", "#d90429")
_ORANGE      = ("#ff9500", "#f4a261")   # optional / info accent

SETTINGS_PATH = _get_settings_path()
APP_VERSION = "1.7.5"

GITHUB_REPO = "DMan1906/SteelSync"
GITHUB_API_LATEST = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"

# ─── Tab Configuration & Metadata ───────────────────────────────────────────
DEFAULT_TAB_ORDER = ["extractor", "nc_sorter", "mod_nc", "dwg_finder", "pdf_to_dwg", "imports_gen"]
TAB_METADATA = {
    "extractor": {"title": "📄 File Extractor", "icon": "📄"},
    "nc_sorter": {"title": "🔧 NC Sorter", "icon": "🔧"},
    "mod_nc": {"title": "📝 Modify NC", "icon": "📝"},
    "dwg_finder": {"title": "🔍 DWG Finder", "icon": "🔍"},
    "pdf_to_dwg": {"title": "📐 PDF to DWG", "icon": "📐"},
    "imports_gen": {"title": "📑 Imports Generator", "icon": "📑"},
}
TAB_NAME_MAP = {
    "extractor": "extractor",
    "file extractor": "extractor",
    "📄 file extractor": "extractor",
    "nc_sorter": "nc_sorter",
    "nc sorter": "nc_sorter",
    "🔧 nc sorter": "nc_sorter",
    "mod_nc": "mod_nc",
    "modify nc": "mod_nc",
    "📝 modify nc": "mod_nc",
    "dwg_finder": "dwg_finder",
    "dwg finder": "dwg_finder",
    "🔍 dwg finder": "dwg_finder",
    "cautare dwg": "dwg_finder",
    "căutare dwg": "dwg_finder",
    "pdf_to_dwg": "pdf_to_dwg",
    "pdf to dwg": "pdf_to_dwg",
    "📐 pdf to dwg": "pdf_to_dwg",
    "pdf2dwg": "pdf_to_dwg",
    "cad generator": "pdf_to_dwg",
    "imports_gen": "imports_gen",
    "imports generator": "imports_gen",
    "generator import": "imports_gen",
    "generator importuri": "imports_gen",
    "📑 imports generator": "imports_gen",
    "metalix": "imports_gen",
    "import metalix": "imports_gen",
}

class _TabViewCompat:
    """Backwards-compatibility shim for any legacy references to self.tabview."""
    def __init__(self, app):
        self._app = app
    def set(self, name):
        norm = str(name).lower().strip()
        tid = TAB_NAME_MAP.get(norm)
        if tid:
            self._app.select_tab(tid)
            return
        for tid, meta in TAB_METADATA.items():
            if norm in meta["title"].lower():
                self._app.select_tab(tid)
                return
    def get(self):
        return TAB_METADATA.get(self._app.active_tab_id, {}).get("title", "")



def _is_geometry_visible(geom_str):
    """
    Check if the window described by geom_str (e.g. '960x860+100+100' or '960x860+-1500+100')
    is visible on any connected monitor across the Windows virtual desktop.
    Prevents windows from reopening off-screen if a monitor was disconnected.
    """
    if not geom_str or not isinstance(geom_str, str):
        return False
    try:
        m = re.match(r'^(\d+)x(\d+)\+?(-?\d+)\+?(-?\d+)$', geom_str.strip())
        if not m:
            return False
        w, h, x, y = int(m.group(1)), int(m.group(2)), int(m.group(3)), int(m.group(4))
        if w < 100 or h < 100:
            return False

        if os.name == 'nt':
            import ctypes
            user32 = ctypes.windll.user32
            vx = user32.GetSystemMetrics(76)   # SM_XVIRTUALSCREEN
            vy = user32.GetSystemMetrics(77)   # SM_YVIRTUALSCREEN
            vw = user32.GetSystemMetrics(78)   # SM_CXVIRTUALSCREEN
            vh = user32.GetSystemMetrics(79)   # SM_CYVIRTUALSCREEN
            if vw > 0 and vh > 0:
                ix1 = max(x, vx)
                iy1 = max(y, vy)
                ix2 = min(x + w, vx + vw)
                iy2 = min(y + h, vy + vh)
                return (ix2 - ix1 >= 100) and (iy2 - iy1 >= 100)
    except Exception:
        pass
    return True


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


class ProjectSelectionDialog(ctk.CTkToplevel):
    """
    Modal dialog allowing the user to select one project folder when multiple match.
    """
    def __init__(self, parent, matches, search_term):
        super().__init__(parent)
        self.parent_app = parent
        self.matches = matches
        self.selected_path = None

        self.title("Selectare Folder Proiect")
        self.geometry("720x460")
        self.minsize(620, 360)
        self.configure(fg_color=_BG)
        
        self.transient(parent)
        self.grab_set()

        self.update_idletasks()
        try:
            x = parent.winfo_x() + (parent.winfo_width() // 2) - 360
            y = parent.winfo_y() + (parent.winfo_height() // 2) - 230
            self.geometry(f"+{max(0, x)}+{max(0, y)}")
        except Exception:
            pass

        self._create_widgets(search_term)

    def _create_widgets(self, search_term):
        container = ctk.CTkFrame(self, fg_color="transparent")
        container.pack(fill="both", expand=True, padx=18, pady=16)

        hdr_card = ctk.CTkFrame(container, fg_color=_CARD_BG, border_color=_CARD_BORDER, border_width=1, corner_radius=10)
        hdr_card.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(
            hdr_card,
            text=f"🔍 S-au găsit {len(self.matches)} foldere de proiect pentru: '{search_term}'",
            font=("Segoe UI", 14, "bold"),
            text_color=_TEXT_PRI,
            anchor="w"
        ).pack(fill="x", padx=14, pady=(10, 2))

        ctk.CTkLabel(
            hdr_card,
            text="Alege folderul dorit din lista de mai jos (dublu-click sau selectează și apasă butonul verde):",
            font=("Segoe UI", 11),
            text_color=_TEXT_SEC,
            anchor="w"
        ).pack(fill="x", padx=14, pady=(0, 10))

        tree_card = ctk.CTkFrame(container, fg_color=_CARD_BG, border_color=_CARD_BORDER, border_width=1, corner_radius=10)
        tree_card.pack(fill="both", expand=True, pady=(0, 12))

        tree_inner = tk.Frame(tree_card, bg="#0d1b2a" if ctk.get_appearance_mode() == "Dark" else "#ffffff")
        tree_inner.pack(fill="both", expand=True, padx=8, pady=8)
        tree_inner.grid_rowconfigure(0, weight=1)
        tree_inner.grid_columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(
            tree_inner,
            columns=("name", "path"),
            show="headings",
            style="DWG.Treeview",
            selectmode="browse"
        )
        self.tree.heading("name", text="Folder Proiect")
        self.tree.heading("path", text="Cale Completă")
        self.tree.column("name", width=220, anchor="w")
        self.tree.column("path", width=440, anchor="w")

        vsb = ttk.Scrollbar(tree_inner, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_inner, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        for p in self.matches:
            self.tree.insert("", "end", values=(os.path.basename(p), p))

        first_child = self.tree.get_children()
        if first_child:
            self.tree.selection_set(first_child[0])

        self.tree.bind("<Double-1>", self._on_double_click)

        btn_box = ctk.CTkFrame(container, fg_color="transparent")
        btn_box.pack(fill="x")

        ctk.CTkButton(
            btn_box,
            text="Renunță",
            width=100,
            height=32,
            fg_color=_ENTRY_BG,
            hover_color=_CARD_BORDER,
            text_color=_TEXT_PRI,
            command=self.destroy
        ).pack(side="right", padx=(8, 0))

        ctk.CTkButton(
            btn_box,
            text="✔ Selectează Proiect",
            width=160,
            height=32,
            fg_color=_GREEN,
            hover_color=_GREEN_HOV,
            text_color="#0b0f19",
            font=("Segoe UI", 12, "bold"),
            command=self._on_select
        ).pack(side="right")

    def _on_double_click(self, event):
        self._on_select()

    def _on_select(self):
        sel = self.tree.selection()
        if sel:
            vals = self.tree.item(sel[0], "values")
            if vals and len(vals) >= 2:
                self.selected_path = vals[1]
        self.destroy()


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
set "_PYI_APPLICATION_HOME_DIR="
set "_PYI_ARCHIVE_FILE="
set "_PYI_PARENT_PROCESS_LEVEL="
set "_PYI_SPLASH_IPC="
set "_MEIPASS2="
set "_MEIPASS="
set "TARGET_EXE={current_exe}"
set "TARGET_DIR={target_dir}"
set "NEW_EXE={temp_exe}"
set "EXE_NAME={exe_name}"
set /a RETRIES=0

:wait_loop
set /a RETRIES+=1
tasklist /fi "IMAGENAME eq %EXE_NAME%" 2>nul | find /i "%EXE_NAME%" >nul
if not errorlevel 1 (
    if %RETRIES% geq 25 (
        taskkill /f /im "%EXE_NAME%" >nul 2>&1
    )
    ping -n 1 -w 150 127.255.255.255 > nul 2>&1
    goto wait_loop
)

:: Instant Hot-Swap
copy /y "%NEW_EXE%" "%TARGET_EXE%" > nul 2>&1
if errorlevel 1 (
    ren "%TARGET_EXE%" "%EXE_NAME%.old_%RANDOM%" > nul 2>&1
    move /y "%NEW_EXE%" "%TARGET_EXE%" > nul 2>&1
    if errorlevel 1 copy /y "%NEW_EXE%" "%TARGET_EXE%" > nul 2>&1
)

:: Launch updated app immediately
set "_PYI_APPLICATION_HOME_DIR="
set "_PYI_ARCHIVE_FILE="
set "_PYI_PARENT_PROCESS_LEVEL="
set "_PYI_SPLASH_IPC="
set "_MEIPASS2="
set "_MEIPASS="
cd /d "%TARGET_DIR%"
explorer.exe "%TARGET_EXE%"

:: Background cleanup
if exist "%NEW_EXE%" del /f /q "%NEW_EXE%" > nul 2>&1
del /f /q "%TARGET_DIR%\\*.old_*" > nul 2>&1
(goto) 2>nul & del /f /q "%~f0"
exit
"""

        try:
            with open(bat_path, "w", encoding="utf-8") as f:
                f.write(bat_content)

            clean_env = os.environ.copy()
            for var in list(clean_env.keys()):
                uvar = var.upper()
                if (uvar.startswith('_PYI') or uvar.startswith('PYI') or
                    uvar.startswith('_MEI') or uvar.startswith('MEI') or
                    'PYTHON' in uvar):
                    clean_env.pop(var, None)

            if 'PATH' in clean_env:
                clean_paths = [p for p in clean_env['PATH'].split(os.pathsep)
                               if '_MEI' not in p.upper() and 'PYTHON' not in p.upper()]
                clean_env['PATH'] = os.pathsep.join(clean_paths)

            creationflags = getattr(subprocess, 'CREATE_NO_WINDOW', 0x08000000)
            subprocess.Popen(
                ["cmd.exe", "/c", bat_path],
                creationflags=creationflags,
                close_fds=True,
                env=clean_env
            )

            try:
                self.parent_app.save_settings()
            except Exception:
                pass

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
        self._last_normal_geometry = None
        self.geometry("960x860")
        self.minsize(900, 650)
        
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
        self.extractor_auto_search = ctk.BooleanVar(value=False)
        self.extractor_proj_search = ctk.StringVar()
        self.extract_assembly_var = ctk.BooleanVar(value=False)
        self.assembly_excel_file = ctk.StringVar()
        self.assembly_col_name = ctk.StringVar()
        
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

        # ── DWG Finder Variables ──
        self.dwg_parent_dir = ctk.StringVar()
        self.dwg_proj_num = ctk.StringVar()
        self.dwg_pdf_dir = ctk.StringVar()
        self.dwg_dest_dir = ctk.StringVar()
        self.dwg_source_type = ctk.StringVar(value="PDF (.pdf)")
        self.dwg_target_type = ctk.StringVar(value="DWG (.dwg)")

        # ── PDF to DWG Variables ──
        self.pdf_cad_mode = ctk.StringVar(value="Folder Întreg (Batch)")
        self.pdf_cad_single_file = ctk.StringVar()
        self.pdf_cad_batch_dir = ctk.StringVar()
        self.pdf_cad_dest_dir = ctk.StringVar()
        self.pdf_cad_format = ctk.StringVar(value="DWG (.dwg cu fallback .dxf)")
        self.pdf_cad_oda_path = ctk.StringVar(value=PdfToCadConverter._find_default_oda_exe() or "")
        self.pdf_cad_only_front = ctk.BooleanVar(value=True)
        self.pdf_cad_calibrate = ctk.BooleanVar(value=True)
        self.dwg_results = {"found": [], "missing": []}
        self._dwg_searching = False
        self._dwg_cancel_requested = False

        # ── DWG Tools Variables (Integrated inside DWG Finder) ──
        self.dwg_subview_mode = ctk.StringVar(value="🔍 Căutare DWG în Proiecte")
        # Subview 2: Organizare pe piese din BOM
        self.dwg_org_excel = ctk.StringVar()
        self.dwg_org_source = ctk.StringVar()
        self.dwg_org_dest = ctk.StringVar()
        self.dwg_org_action = ctk.StringVar(value="move")
        self.dwg_org_sheet = ctk.StringVar(value="-")
        self.dwg_org_asmb_col = ctk.StringVar(value="-")
        self.dwg_org_part_col = ctk.StringVar(value="-")
        self.dwg_org_variations = ctk.BooleanVar(value=True)
        self.dwg_org_clean_empty = ctk.BooleanVar(value=True)
        self.dwg_org_open_dest = ctk.BooleanVar(value=True)
        self.dwg_org_running = False
        self.dwg_org_cancel = False
        self.dwg_org_last_report = None
        self.dwg_org_excel_meta = None
        # Subview 3: Regrupare DWG pe grosimi
        self.dwg_regroup_source = ctk.StringVar()
        self.dwg_regroup_dest = ctk.StringVar()
        self.dwg_regroup_action = ctk.StringVar(value="copy")
        self.dwg_regroup_normalize = ctk.BooleanVar(value=True)
        self.dwg_regroup_clean_empty = ctk.BooleanVar(value=True)
        self.dwg_regroup_open_dest = ctk.BooleanVar(value=True)
        self.dwg_regroup_running = False
        self.dwg_regroup_cancel = False
        self.dwg_regroup_last_report = None

        # ── Imports Generator Variables (Dedicated New Tab) ──
        self.imports_selected_files = []
        self.imports_parsed_cache = {}
        self.imports_last_output_dir = None
        self.imports_last_generated_excel = None
        self.imports_running = False
        self.imports_cancel_requested = False
        self.imports_export_mode = ctk.StringVar(value="individual")
        self.imports_start_row_mode = ctk.StringVar(value="row2_empty1")
        self.imports_merge_dups = ctk.BooleanVar(value=False)
        self.imports_consolidated_name = ctk.StringVar(value="IMPORT.xlsx")
        self.imports_output_dir = ctk.StringVar(value=os.path.join(os.path.expanduser("~"), "Desktop"))
        self.imports_filter_text = ctk.StringVar(value="")
        self.imports_remove_ext = ctk.BooleanVar(value=True)
        self.imports_remove_index = ctk.BooleanVar(value=True)
        self.imports_keep_plus = ctk.BooleanVar(value=True)
        self.imports_remove_special = ctk.BooleanVar(value=True)
        self.imports_uppercase = ctk.BooleanVar(value=True)
        self.imports_open_folder = ctk.BooleanVar(value=True)

        # ── Extension tracking ──
        self.available_extensions = ['.nc', '.pdf', '.dwg', '.dxf', '.step', '.stp']
        self.extension_vars = {}

        # ── Update tracking ──
        self.ignored_update_version = ""
        self._update_dialog = None

        # ── Directory Browse Memories (Independent per Requirement #3) ──
        self.extractor_source_last_dir = ""
        self.extractor_dest_last_dir = ""
        self.global_last_dir = ""
        self._settings_loaded = False
        self._saved_settings_raw = {}

        # ── Tab Ordering & State (Requirement #2 with Drag & Drop) ──
        self.tab_order = list(DEFAULT_TAB_ORDER)
        self.active_tab_id = "extractor"
        self._drag_data = None
        self._load_saved_variables()

        self.create_widgets()
        self._apply_settings_ui()
        self.bind("<Configure>", self._on_window_configure)
        self.bind("<Control-Left>", lambda e: self.move_active_tab_left())
        self.bind("<Control-Right>", lambda e: self.move_active_tab_right())
        self.bind("<Control-Tab>", lambda e: self.cycle_tab_forward())
        self.bind("<Control-Shift-Tab>", lambda e: self.cycle_tab_backward())
        self.bind("<Control-ISO_Left_Tab>", lambda e: self.cycle_tab_backward())

        # Preload heavy modules (pandas, openpyxl) asynchronously in background
        threading.Thread(target=_preload_heavy_libraries, daemon=True).start()

        # Check for updates automatically in background (silent & non-blocking)
        self.after(1500, lambda: self.start_update_check(manual=False))

        # Save settings on close
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def _on_window_configure(self, event=None):
        """Smooth, debounced tracking of window resizing and position."""
        if event is not None and getattr(event, "widget", None) != self:
            return

        # Debounce geometry saving (prevents Win32 lag during mouse drag)
        if getattr(self, "_geom_save_timer", None) is not None:
            try:
                self.after_cancel(self._geom_save_timer)
            except Exception:
                pass
        self._geom_save_timer = self.after(200, self._record_normal_geometry)

        # Trigger responsive extension relayout if width changed (debounced to avoid Win32 resize stutter)
        curr_w = event.width if event else self.winfo_width()
        last_w = getattr(self, "_last_relayout_w", 0)
        if abs(curr_w - last_w) >= 20:
            self._last_relayout_w = curr_w
            if getattr(self, "_ext_relayout_timer", None) is not None:
                try:
                    self.after_cancel(self._ext_relayout_timer)
                except Exception:
                    pass
            self._ext_relayout_timer = self.after(120, lambda w=curr_w: self.reposition_extensions(w))

    def _record_normal_geometry(self):
        try:
            if self.state() == "normal":
                self._last_normal_geometry = self.geometry()
        except Exception:
            pass

    # ══════════════════════════════════════════════════════════════════════
    #  SETTINGS PERSISTENCE
    # ══════════════════════════════════════════════════════════════════════
    def _load_saved_variables(self):
        """Pre-load all variables and values from settings JSON before widgets are constructed."""
        try:
            if not os.path.exists(SETTINGS_PATH):
                return
            with open(SETTINGS_PATH, 'r', encoding='utf-8') as f:
                s = json.load(f)
            self._saved_settings_raw = s

            self.ignored_update_version = s.get("ignored_update_version", "")
            self.appearance_mode_var.set(s.get("appearance_mode", "Dark"))
            ctk.set_appearance_mode(self.appearance_mode_var.get())
            self.source_dir.set(s.get("source_dir", ""))
            self.dest_dir.set(s.get("dest_dir", ""))
            self.excel_file.set(s.get("excel_file", ""))
            self.col_name.set(s.get("col_name", ""))
            self.extractor_auto_search.set(s.get("extractor_auto_search", False))
            self.extractor_proj_search.set(s.get("extractor_proj_search", ""))
            self.extract_assembly_var.set(s.get("extract_assembly_var", False))
            self.assembly_excel_file.set(s.get("assembly_excel_file", ""))
            self.assembly_col_name.set(s.get("assembly_col_name", ""))
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

            # DWG Finder settings
            self.dwg_parent_dir.set(s.get("dwg_parent_dir", ""))
            self.dwg_proj_num.set(s.get("dwg_proj_num", ""))
            self.dwg_pdf_dir.set(s.get("dwg_pdf_dir", ""))
            self.dwg_dest_dir.set(s.get("dwg_dest_dir", ""))
            self.dwg_source_type.set(s.get("dwg_source_type", "PDF (.pdf)"))
            self.dwg_target_type.set(s.get("dwg_target_type", "DWG (.dwg)"))

            # PDF to DWG settings
            self.pdf_cad_mode.set(s.get("pdf_cad_mode", "Folder Întreg (Batch)"))
            self.pdf_cad_single_file.set(s.get("pdf_cad_single_file", ""))
            self.pdf_cad_batch_dir.set(s.get("pdf_cad_batch_dir", ""))
            self.pdf_cad_dest_dir.set(s.get("pdf_cad_dest_dir", ""))
            self.pdf_cad_format.set(s.get("pdf_cad_format", "DWG (.dwg cu fallback .dxf)"))
            saved_oda = s.get("pdf_cad_oda_path", "").strip()
            if not saved_oda or not os.path.isfile(saved_oda):
                saved_oda = PdfToCadConverter._find_default_oda_exe() or ""
            self.pdf_cad_oda_path.set(saved_oda)
            self.pdf_cad_only_front.set(s.get("pdf_cad_only_front", True))
            self.pdf_cad_calibrate.set(s.get("pdf_cad_calibrate", True))

            # DWG Tools settings
            self.dwg_subview_mode.set(s.get("dwg_subview_mode", "🔍 Căutare DWG în Proiecte"))
            self.dwg_org_excel.set(s.get("dwg_org_excel", ""))
            self.dwg_org_source.set(s.get("dwg_org_source", ""))
            self.dwg_org_dest.set(s.get("dwg_org_dest", ""))
            self.dwg_org_action.set(s.get("dwg_org_action", "move"))
            self.dwg_regroup_source.set(s.get("dwg_regroup_source", ""))
            self.dwg_regroup_dest.set(s.get("dwg_regroup_dest", ""))
            self.dwg_regroup_action.set(s.get("dwg_regroup_action", "copy"))

            # Imports Generator settings
            self.imports_export_mode.set(s.get("imports_export_mode", "individual"))
            self.imports_start_row_mode.set(s.get("imports_start_row_mode", "row2_empty1"))
            self.imports_merge_dups.set(s.get("imports_merge_dups", False))
            self.imports_consolidated_name.set(s.get("imports_consolidated_name", "IMPORT.xlsx"))
            default_desk = os.path.join(os.path.expanduser("~"), "Desktop")
            self.imports_output_dir.set(s.get("imports_output_dir", default_desk))
            self.imports_open_folder.set(s.get("imports_open_folder", True))

            # Restore browse directory memories (Requirement #3)
            self.extractor_source_last_dir = s.get("extractor_source_last_dir", "")
            if not self.extractor_source_last_dir:
                self.extractor_source_last_dir = self.source_dir.get()
            self.extractor_dest_last_dir = s.get("extractor_dest_last_dir", "")
            if not self.extractor_dest_last_dir:
                self.extractor_dest_last_dir = self.dest_dir.get()
            self.global_last_dir = s.get("global_last_dir", "")
            if not self.global_last_dir:
                self.global_last_dir = self.dest_dir.get() or self.nc_source_folder.get()

            # Pre-load Tab Ordering & Active Tab (Requirement #2)
            raw_order = s.get("tab_order", [])
            valid_order = []
            for item in raw_order:
                tid = TAB_NAME_MAP.get(str(item).lower().strip(), item)
                if tid in DEFAULT_TAB_ORDER and tid not in valid_order:
                    valid_order.append(tid)
            for tid in DEFAULT_TAB_ORDER:
                if tid not in valid_order:
                    valid_order.append(tid)
            self.tab_order = valid_order

            saved_active = s.get("active_tab_id", "")
            mapped_active = TAB_NAME_MAP.get(str(saved_active).lower().strip(), saved_active)
            if mapped_active in self.tab_order:
                self.active_tab_id = mapped_active
            else:
                self.active_tab_id = self.tab_order[0]
        except Exception:
            pass

    def _apply_settings_ui(self):
        """Apply loaded settings to UI elements after create_widgets() finishes."""
        s = getattr(self, "_saved_settings_raw", {})
        try:
            # 1. Apply tab order & active tab
            self.apply_tab_order()
            self.select_tab(self.active_tab_id)

            # 2. Restore extension selections
            saved_exts = s.get("extensions", {})
            for ext, val in saved_exts.items():
                if ext in self.extension_vars:
                    self.extension_vars[ext].set(val)
                elif ext not in self.available_extensions:
                    self.available_extensions.append(ext)
                    self.extension_vars[ext] = ctk.BooleanVar(value=val)
            self.draw_extensions()

            # 3. Synchronize toggle UI states
            if hasattr(self, "_on_toggle_auto_search"):
                self._on_toggle_auto_search()
            if hasattr(self, "_on_toggle_extract_assembly"):
                self._on_toggle_extract_assembly()

            # 4. Restore window geometry and position (multi-monitor safe)
            saved_geom = s.get("window_geometry", "")
            if saved_geom and _is_geometry_visible(saved_geom):
                try:
                    m = re.match(r'^(\d+)x(\d+)\+?(-?\d+)\+?(-?\d+)$', saved_geom.strip())
                    if m:
                        gw, gh, gx, gy = int(m.group(1)), int(m.group(2)), int(m.group(3)), int(m.group(4))
                        gw = max(gw, 900)
                        gh = max(gh, 650)
                        saved_geom = f"{gw}x{gh}+{gx}+{gy}"
                    self.geometry(saved_geom)
                    self._last_normal_geometry = saved_geom
                except Exception:
                    pass
            if s.get("window_maximized", False):
                try:
                    self.after(50, lambda: self.state("zoomed"))
                except Exception:
                    pass
        except Exception:
            pass
        finally:
            self._settings_loaded = True
            self.save_settings()

    def load_settings(self):
        """Load persisted user settings from JSON and apply to UI."""
        self._load_saved_variables()
        self._apply_settings_ui()

    def save_settings(self):
        """Persist current user settings to JSON."""
        if not getattr(self, "_settings_loaded", False):
            return
        is_zoomed = False
        try:
            is_zoomed = (self.state() == "zoomed")
        except Exception:
            pass

        current_geom = None
        if is_zoomed:
            current_geom = self._last_normal_geometry or "960x860"
        else:
            try:
                current_geom = self.geometry()
                self._last_normal_geometry = current_geom
            except Exception:
                current_geom = self._last_normal_geometry or "960x860"

        ext_map = {ext: var.get() for ext, var in self.extension_vars.items()}
        settings = {
            "window_geometry": current_geom,
            "window_maximized": is_zoomed,
            "appearance_mode": self.appearance_mode_var.get(),
            "source_dir": self.source_dir.get(),
            "dest_dir": self.dest_dir.get(),
            "excel_file": self.excel_file.get(),
            "col_name": self.col_name.get(),
            "extractor_auto_search": self.extractor_auto_search.get(),
            "extractor_proj_search": self.extractor_proj_search.get(),
            "extract_assembly_var": self.extract_assembly_var.get(),
            "assembly_excel_file": self.assembly_excel_file.get(),
            "assembly_col_name": self.assembly_col_name.get(),
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
            "dwg_parent_dir": self.dwg_parent_dir.get(),
            "dwg_proj_num": self.dwg_proj_num.get(),
            "dwg_pdf_dir": self.dwg_pdf_dir.get(),
            "dwg_dest_dir": self.dwg_dest_dir.get(),
            "dwg_source_type": self.dwg_source_type.get(),
            "dwg_target_type": self.dwg_target_type.get(),
            "pdf_cad_mode": self.pdf_cad_mode.get(),
            "pdf_cad_single_file": self.pdf_cad_single_file.get(),
            "pdf_cad_batch_dir": self.pdf_cad_batch_dir.get(),
            "pdf_cad_dest_dir": self.pdf_cad_dest_dir.get(),
            "pdf_cad_format": self.pdf_cad_format.get(),
            "pdf_cad_oda_path": self.pdf_cad_oda_path.get(),
            "pdf_cad_only_front": self.pdf_cad_only_front.get(),
            "pdf_cad_calibrate": self.pdf_cad_calibrate.get(),
            "dwg_subview_mode": self.dwg_subview_mode.get(),
            "dwg_org_excel": self.dwg_org_excel.get(),
            "dwg_org_source": self.dwg_org_source.get(),
            "dwg_org_dest": self.dwg_org_dest.get(),
            "dwg_org_action": self.dwg_org_action.get(),
            "dwg_regroup_source": self.dwg_regroup_source.get(),
            "dwg_regroup_dest": self.dwg_regroup_dest.get(),
            "dwg_regroup_action": self.dwg_regroup_action.get(),
            "imports_export_mode": self.imports_export_mode.get(),
            "imports_start_row_mode": self.imports_start_row_mode.get(),
            "imports_merge_dups": self.imports_merge_dups.get(),
            "imports_consolidated_name": self.imports_consolidated_name.get(),
            "imports_output_dir": self.imports_output_dir.get(),
            "imports_open_folder": self.imports_open_folder.get(),
            "ignored_update_version": getattr(self, "ignored_update_version", ""),
            "extensions": ext_map,
            "extractor_source_last_dir": getattr(self, "extractor_source_last_dir", ""),
            "extractor_dest_last_dir": getattr(self, "extractor_dest_last_dir", ""), 
            "global_last_dir": getattr(self, "global_last_dir", ""),
            "tab_order": getattr(self, "tab_order", list(DEFAULT_TAB_ORDER)),
            "active_tab_id": getattr(self, "active_tab_id", "extractor"),
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
        if hasattr(self, "_update_tab_buttons_style"):
            self._update_tab_buttons_style()
        if hasattr(self, "_setup_treeview_styles"):
            self._setup_treeview_styles()
        self.save_settings()

    def on_close(self):
        if hasattr(self, "_update_tab_buttons_style"):
            self._update_tab_buttons_style()
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
    #  DIRECTORY BROWSE HELPERS (INDEPENDENT & CONTEXT-AWARE MEMORIES)
    # ══════════════════════════════════════════════════════════════════════
    def _get_browse_initial_dir(self, target_type, current_val=""):
        """
        Determine the initial directory for file/folder dialogs according to User Rule:
        - 'extractor_source': independent memory (source_dir / extractor_source_last_dir)
        - 'extractor_dest': independent memory (dest_dir / extractor_dest_last_dir)
        - 'dest_based': based on File Extractor Destinatie (dest_dir / extractor_dest_last_dir)
        - 'dest_sorted_based': based on Destinatie + "\\Sorted" (or "\\SORTED", fallback dest_dir)
        """
        def _as_valid_dir(p):
            if not p:
                return None
            norm = os.path.normpath(str(p).strip())
            if os.path.isdir(norm):
                return norm
            parent = os.path.dirname(norm)
            if parent and os.path.isdir(parent):
                return parent
            return None

        # 1. Extractor Source (Independent)
        if target_type == "extractor_source":
            d = _as_valid_dir(current_val) or _as_valid_dir(getattr(self, "extractor_source_last_dir", "")) or _as_valid_dir(self.source_dir.get())
            return d or None

        # 2. Extractor Dest (Independent)
        if target_type == "extractor_dest":
            d = _as_valid_dir(current_val) or _as_valid_dir(getattr(self, "extractor_dest_last_dir", "")) or _as_valid_dir(self.dest_dir.get())
            return d or None

        # Base directory from File Extractor Destinatie
        dest_base = _as_valid_dir(self.dest_dir.get()) or _as_valid_dir(getattr(self, "extractor_dest_last_dir", ""))

        # 3. Dest-based (Excel Extractor, Assembly Excel, NC Sorter Source, NC Sorter Excel)
        if target_type == "dest_based":
            if dest_base:
                return dest_base
            return _as_valid_dir(current_val) or _as_valid_dir(getattr(self, "global_last_dir", "")) or None

        # 4. Dest-Sorted-based (Modify NC Folder, Modify NC Excel) -> Destinatie + "\Sorted"
        if target_type == "dest_sorted_based":
            if dest_base:
                for sub in ("Sorted", "SORTED", "sorted"):
                    candidate = os.path.join(dest_base, sub)
                    if os.path.isdir(candidate):
                        return candidate
                return dest_base
            return _as_valid_dir(current_val) or _as_valid_dir(getattr(self, "global_last_dir", "")) or None

        return _as_valid_dir(current_val) or None

    def _get_global_initial_dir(self, current_val=""):
        """Get the initial directory for global browse dialogs (fallback helper)."""
        return self._get_browse_initial_dir("dest_based", current_val)

    def _set_global_last_dir(self, chosen_path):
        """Update global last dir memory and persist."""
        if not chosen_path:
            return
        norm = os.path.normpath(chosen_path)
        if os.path.isdir(norm):
            self.global_last_dir = norm
        elif os.path.isfile(norm):
            self.global_last_dir = os.path.dirname(norm)
        else:
            _, ext = os.path.splitext(norm)
            if ext:
                self.global_last_dir = os.path.dirname(norm)
            else:
                self.global_last_dir = norm
        self.save_settings()

    # ══════════════════════════════════════════════════════════════════════
    #  TAB MANAGER & REORDERING METHODS (REQUIREMENT #2 WITH DRAG & DROP)
    # ══════════════════════════════════════════════════════════════════════
    def _load_saved_tab_order(self):
        """Pre-load tab order from settings before widgets are constructed."""
        try:
            if os.path.exists(SETTINGS_PATH):
                with open(SETTINGS_PATH, 'r', encoding='utf-8') as f:
                    s = json.load(f)
                raw_order = s.get("tab_order", [])
                valid_order = []
                for item in raw_order:
                    tid = TAB_NAME_MAP.get(str(item).lower().strip(), item)
                    if tid in DEFAULT_TAB_ORDER and tid not in valid_order:
                        valid_order.append(tid)
                for tid in DEFAULT_TAB_ORDER:
                    if tid not in valid_order:
                        valid_order.append(tid)
                self.tab_order = valid_order

                saved_active = s.get("active_tab_id", "")
                mapped_active = TAB_NAME_MAP.get(str(saved_active).lower().strip(), saved_active)
                if mapped_active in self.tab_order:
                    self.active_tab_id = mapped_active
                else:
                    self.active_tab_id = self.tab_order[0]
        except Exception:
            pass

    def _setup_tabview_drag_and_drop(self):
        """Enable intuitive in-place drag-and-drop reordering directly on the floating tabs."""
        if not hasattr(self, "tabview") or not hasattr(self.tabview, "_segmented_button"):
            return
        sb = self.tabview._segmented_button
        title_to_id = {meta["title"]: tid for tid, meta in TAB_METADATA.items()}

        def on_press(event, tab_title):
            self._drag_data = {
                "title": tab_title,
                "start_x": event.x_root,
                "swapped": False,
            }
            try:
                self.tabview.set(tab_title)
                tid = title_to_id.get(tab_title)
                if tid:
                    self.active_tab_id = tid
            except Exception:
                pass

        def on_motion(event):
            if not getattr(self, "_drag_data", None):
                return
            tab_title = self._drag_data.get("title")
            if not tab_title or tab_title not in sb._value_list:
                return

            idx = sb._value_list.index(tab_title)

            # Dragging right
            if idx < len(sb._value_list) - 1:
                next_title = sb._value_list[idx + 1]
                next_btn = sb._buttons_dict.get(next_title)
                if next_btn:
                    mid_x = next_btn.winfo_rootx() + (next_btn.winfo_width() // 2)
                    if event.x_root > mid_x:
                        sb._value_list[idx], sb._value_list[idx + 1] = sb._value_list[idx + 1], sb._value_list[idx]
                        self.tabview._name_list = list(sb._value_list)
                        sb._create_button_grid()
                        for k in range(len(sb._value_list)):
                            sb._configure_button_corners_for_index(k)
                        self.update_idletasks()
                        self._drag_data["swapped"] = True
                        return

            # Dragging left
            if idx > 0:
                prev_title = sb._value_list[idx - 1]
                prev_btn = sb._buttons_dict.get(prev_title)
                if prev_btn:
                    mid_x = prev_btn.winfo_rootx() + (prev_btn.winfo_width() // 2)
                    if event.x_root < mid_x:
                        sb._value_list[idx], sb._value_list[idx - 1] = sb._value_list[idx - 1], sb._value_list[idx]
                        self.tabview._name_list = list(sb._value_list)
                        sb._create_button_grid()
                        for k in range(len(sb._value_list)):
                            sb._configure_button_corners_for_index(k)
                        self.update_idletasks()
                        self._drag_data["swapped"] = True
                        return

        def on_release(event):
            if getattr(self, "_drag_data", None) and self._drag_data.get("swapped"):
                new_order = [title_to_id[t] for t in sb._value_list if t in title_to_id]
                if new_order:
                    self.tab_order = new_order
                    self.save_settings()
            self._drag_data = None

        def on_right_click(event, tab_title):
            tid = title_to_id.get(tab_title)
            if tid:
                self.show_tab_context_menu(event, tid)

        for tab_title, btn in sb._buttons_dict.items():
            widgets = [btn]
            if hasattr(btn, "_canvas") and btn._canvas:
                widgets.append(btn._canvas)
            if hasattr(btn, "_text_label") and btn._text_label:
                widgets.append(btn._text_label)

            for w in widgets:
                try:
                    w.bind("<Button-1>", lambda e, t=tab_title: on_press(e, t), add="+")
                    w.bind("<B1-Motion>", on_motion, add="+")
                    w.bind("<ButtonRelease-1>", on_release, add="+")
                    w.bind("<Button-3>", lambda e, t=tab_title: on_right_click(e, t), add="+")
                except Exception:
                    pass

    def apply_tab_order(self):
        """Reorder segmented button items according to self.tab_order."""
        if not hasattr(self, "tabview") or not hasattr(self.tabview, "_segmented_button"):
            return
        sb = self.tabview._segmented_button
        new_values = [TAB_METADATA[tid]["title"] for tid in self.tab_order if tid in TAB_METADATA]
        if set(new_values) == set(sb._value_list):
            sb._value_list = new_values
            self.tabview._name_list = list(sb._value_list)
            sb._create_button_grid()
            for k in range(len(sb._value_list)):
                sb._configure_button_corners_for_index(k)

    def select_tab(self, tab_id):
        """Switch to tab_id in tabview."""
        norm_id = TAB_NAME_MAP.get(str(tab_id).lower().strip(), tab_id)
        if norm_id not in TAB_METADATA:
            return
        self.active_tab_id = norm_id
        meta = TAB_METADATA[norm_id]
        try:
            self.tabview.set(meta["title"])
        except Exception:
            pass
        self.save_settings()

    def _on_tab_changed(self):
        """Called when user switches tabs via CTkTabview."""
        try:
            current_name = self.tabview.get()
            norm = str(current_name).lower().strip()
            tid = TAB_NAME_MAP.get(norm)
            if tid:
                self.active_tab_id = tid
                self.save_settings()
        except Exception:
            pass

    def move_active_tab_left(self):
        """Move the currently active tab one position to the left."""
        if self.active_tab_id not in self.tab_order:
            return
        idx = self.tab_order.index(self.active_tab_id)
        if idx > 0:
            self.tab_order[idx], self.tab_order[idx - 1] = self.tab_order[idx - 1], self.tab_order[idx]
            self.apply_tab_order()
            self.save_settings()

    def move_active_tab_right(self):
        """Move the currently active tab one position to the right."""
        if self.active_tab_id not in self.tab_order:
            return
        idx = self.tab_order.index(self.active_tab_id)
        if idx < len(self.tab_order) - 1:
            self.tab_order[idx], self.tab_order[idx + 1] = self.tab_order[idx + 1], self.tab_order[idx]
            self.apply_tab_order()
            self.save_settings()

    def reset_tab_order(self):
        """Reset tab order to default."""
        self.tab_order = list(DEFAULT_TAB_ORDER)
        self.apply_tab_order()
        self.save_settings()

    def show_tab_context_menu(self, event, tab_id):
        """Show context menu for reordering tabs on right click."""
        self.select_tab(tab_id)
        menu = tk.Menu(self, tearoff=0)
        idx = self.tab_order.index(tab_id) if tab_id in self.tab_order else 0
        if idx > 0:
            menu.add_command(label="⬅ Mută la stânga", command=self.move_active_tab_left)
        else:
            menu.add_command(label="⬅ Mută la stânga", state="disabled")

        if idx < len(self.tab_order) - 1:
            menu.add_command(label="➡ Mută la dreapta", command=self.move_active_tab_right)
        else:
            menu.add_command(label="➡ Mută la dreapta", state="disabled")

        menu.add_separator()
        menu.add_command(label="🔄 Resetează ordinea inițială", command=self.reset_tab_order)
        menu.post(event.x_root, event.y_root)

    def cycle_tab_forward(self):
        """Ctrl+Tab: switch to next tab."""
        idx = self.tab_order.index(self.active_tab_id)
        next_idx = (idx + 1) % len(self.tab_order)
        self.select_tab(self.tab_order[next_idx])

    def cycle_tab_backward(self):
        """Ctrl+Shift+Tab: switch to previous tab."""
        idx = self.tab_order.index(self.active_tab_id)
        prev_idx = (idx - 1) % len(self.tab_order)
        self.select_tab(self.tab_order[prev_idx])

    # ══════════════════════════════════════════════════════════════════════
    #  MAIN LAYOUT
    # ══════════════════════════════════════════════════════════════════════
    def create_widgets(self):
        container = ctk.CTkFrame(self, fg_color="transparent")
        container.pack(fill="both", expand=True, padx=14, pady=10)
        
        top_bar = ctk.CTkFrame(container, fg_color="transparent")
        top_bar.pack(fill="x", pady=(0, 6))
        
        title_lbl = ctk.CTkLabel(
            top_bar,
            text="SteelSync Toolkit",
            font=("Segoe UI", 20, "bold"),
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

        # ──── Native Floating CTkTabview (Requirement #2 with Drag & Drop) ────
        self.tabview = ctk.CTkTabview(
            container,
            fg_color=_CARD_BG,
            border_color=_CARD_BORDER,
            border_width=1,
            segmented_button_fg_color=_ENTRY_BG,
            segmented_button_unselected_color=_ENTRY_BG,
            segmented_button_unselected_hover_color=_CARD_BORDER,
            segmented_button_selected_color=_ACCENT,
            segmented_button_selected_hover_color=_ACCENT_HOV,
            text_color=_TEXT_PRI,
            command=self._on_tab_changed
        )
        self.tabview.pack(fill="both", expand=True)
        if hasattr(self.tabview, "_segmented_button"):
            try:
                self.tabview._segmented_button.configure(border_width=1)
            except Exception:
                pass

        # Tab Frames (instantiated in current tab_order)
        self.tab_frames = {}
        for tab_id in self.tab_order:
            meta = TAB_METADATA[tab_id]
            frame = self.tabview.add(meta["title"])
            self.tab_frames[tab_id] = frame

        self.tab_ext = self.tab_frames["extractor"]
        self.tab_nc = self.tab_frames["nc_sorter"]
        self.tab_mod = self.tab_frames["mod_nc"]
        self.tab_dwg = self.tab_frames["dwg_finder"]
        self.tab_pdf_cad = self.tab_frames["pdf_to_dwg"]
        self.tab_imports = self.tab_frames["imports_gen"]

        # Backwards-compatibility wrapper for any self.tabview.set calls
        _orig_tabview_set = self.tabview.set
        def _safe_tabview_set(name):
            norm = str(name).lower().strip()
            tid = TAB_NAME_MAP.get(norm)
            if tid and tid in TAB_METADATA:
                return _orig_tabview_set(TAB_METADATA[tid]["title"])
            return _orig_tabview_set(name)
        self.tabview.set = _safe_tabview_set

        self.create_extractor_tab()
        self.create_nc_sorter_tab()
        self.create_mod_nc_tab()
        self.create_dwg_finder_tab()
        self.create_pdf_to_dwg_tab()
        self.create_imports_generator_tab()

        self._setup_tabview_drag_and_drop()

        active_meta = TAB_METADATA.get(self.active_tab_id, TAB_METADATA["extractor"])
        try:
            self.tabview.set(active_meta["title"])
        except Exception:
            pass

    # ══════════════════════════════════════════════════════════════════════
    #  EXTRACTOR TAB
    # ══════════════════════════════════════════════════════════════════════
    def create_extractor_tab(self):
        container = ctk.CTkFrame(self.tab_ext, fg_color="transparent")
        container.pack(fill="both", expand=True, padx=12, pady=8)
        container.grid_columnconfigure(0, weight=1)
        
        subtitle_lbl = ctk.CTkLabel(
            container,
            text="Caută și copiază fișierele proiectului pe baza unei liste Excel",
            font=("Segoe UI", 12),
            text_color=_TEXT_SEC,
        )
        subtitle_lbl.grid(row=0, column=0, pady=(0, 6), sticky="w")

        # ──── Paths card ────
        paths_card = self._make_card(container)
        paths_card.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        paths_card.grid_columnconfigure(0, weight=0)
        paths_card.grid_columnconfigure(1, weight=1)
        paths_card.grid_columnconfigure(2, weight=0)

        section_lbl = ctk.CTkLabel(paths_card, text="📂 Directoare & Excel", font=("Segoe UI", 14, "bold"), text_color=_TEXT_PRI)
        section_lbl.grid(row=0, column=0, columnspan=3, padx=16, pady=(8, 4), sticky="w")

        # Auto project search checkbox
        self.cb_auto_proj = ctk.CTkCheckBox(
            paths_card,
            text="Caută automat proiectul în Folderul Mamă",
            variable=self.extractor_auto_search,
            command=self._on_toggle_auto_search,
            font=("Segoe UI", 12),
            fg_color=_ACCENT,
            hover_color=_ACCENT_HOV,
            border_color=_CARD_BORDER,
            text_color=_TEXT_PRI
        )
        self.cb_auto_proj.grid(row=1, column=1, columnspan=2, padx=4, pady=(2, 6), sticky="w")

        self.lbl_extractor_source = self._label(paths_card, text="Sursă:")
        self.lbl_extractor_source.grid(row=2, column=0, padx=(16, 6), pady=4, sticky="w")
        self._entry(paths_card, self.source_dir).grid(row=2, column=1, padx=4, pady=4, sticky="ew")
        self._browse_btn(paths_card, self.browse_source).grid(row=2, column=2, padx=(4, 16), pady=4)

        # Dynamic Project search frame (shown when auto search is enabled)
        self.proj_search_frame = ctk.CTkFrame(paths_card, fg_color="transparent")
        self.proj_search_frame.grid_columnconfigure(0, weight=0)
        self.proj_search_frame.grid_columnconfigure(1, weight=1)
        self.proj_search_frame.grid_columnconfigure(2, weight=0)

        self._label(self.proj_search_frame, text="Nr. / Denumire Proiect:").grid(row=0, column=0, padx=(0, 6), pady=2, sticky="w")
        self._entry(self.proj_search_frame, self.extractor_proj_search, placeholder_text="ex: 1242 sau Fiskebrygga").grid(row=0, column=1, padx=4, pady=2, sticky="ew")
        self.btn_verify_proj = ctk.CTkButton(
            self.proj_search_frame,
            text="🔍 Verifică",
            width=110,
            height=34,
            font=("Segoe UI", 12, "bold"),
            fg_color=_ENTRY_BG,
            hover_color=_ACCENT,
            text_color=_TEXT_PRI,
            corner_radius=8,
            command=self.find_project_folder_interactive
        )
        self.btn_verify_proj.grid(row=0, column=2, padx=(4, 0), pady=2)

        self._label(paths_card, text="Destinație:").grid(row=4, column=0, padx=(16, 6), pady=4, sticky="w")
        self._entry(paths_card, self.dest_dir).grid(row=4, column=1, padx=4, pady=4, sticky="ew")
        self._browse_btn(paths_card, self.browse_dest).grid(row=4, column=2, padx=(4, 16), pady=4)

        self._label(paths_card, text="Fișier Excel:").grid(row=5, column=0, padx=(16, 6), pady=(4, 10), sticky="w")
        
        excel_inner = ctk.CTkFrame(paths_card, fg_color="transparent")
        excel_inner.grid(row=5, column=1, padx=4, pady=(4, 10), sticky="ew")
        excel_inner.grid_columnconfigure(0, weight=1)
        excel_inner.grid_columnconfigure(1, weight=0)
        excel_inner.grid_columnconfigure(2, weight=0)

        self._entry(excel_inner, self.excel_file).grid(row=0, column=0, sticky="ew")
        self._label(excel_inner, text="Col:", font=("Segoe UI", 11)).grid(row=0, column=1, padx=(8, 2))
        self._entry(excel_inner, self.col_name, width=80, placeholder_text="auto").grid(row=0, column=2, padx=(0, 0))

        self._browse_btn(paths_card, self.browse_excel).grid(row=5, column=2, padx=(4, 16), pady=(4, 10))

        # Assembly drawings extraction checkbox (Row 6)
        self.cb_extract_assembly = ctk.CTkCheckBox(
            paths_card,
            text="Extrage și Desenele de Ansamblu din Excel Ansambluri",
            variable=self.extract_assembly_var,
            command=self._on_toggle_extract_assembly,
            font=("Segoe UI", 12),
            fg_color=_ACCENT,
            hover_color=_ACCENT_HOV,
            border_color=_CARD_BORDER,
            text_color=_TEXT_PRI
        )
        self.cb_extract_assembly.grid(row=6, column=1, columnspan=2, padx=4, pady=(2, 6), sticky="w")

        # Dynamic Assembly Excel frame (Row 7)
        self.assembly_frame = ctk.CTkFrame(paths_card, fg_color="transparent")
        self.assembly_frame.grid_columnconfigure(0, weight=0)
        self.assembly_frame.grid_columnconfigure(1, weight=1)
        self.assembly_frame.grid_columnconfigure(2, weight=0)

        self._label(self.assembly_frame, text="Excel Ansambluri:").grid(row=0, column=0, padx=(0, 6), pady=2, sticky="w")
        asmb_inner = ctk.CTkFrame(self.assembly_frame, fg_color="transparent")
        asmb_inner.grid(row=0, column=1, padx=4, pady=2, sticky="ew")
        asmb_inner.grid_columnconfigure(0, weight=1)
        asmb_inner.grid_columnconfigure(1, weight=0)
        asmb_inner.grid_columnconfigure(2, weight=0)

        self._entry(asmb_inner, self.assembly_excel_file).grid(row=0, column=0, sticky="ew")
        self._label(asmb_inner, text="Col:", font=("Segoe UI", 11)).grid(row=0, column=1, padx=(8, 2))
        self._entry(asmb_inner, self.assembly_col_name, width=80, placeholder_text="auto").grid(row=0, column=2, padx=(0, 0))

        self._browse_btn(self.assembly_frame, self.browse_assembly_excel).grid(row=0, column=2, padx=(4, 0), pady=2)

        self._on_toggle_auto_search()
        self._on_toggle_extract_assembly()

        # ──── Extensions card ────
        ext_card = self._make_card(container)
        ext_card.grid(row=2, column=0, sticky="ew", pady=(0, 8))
        self.ext_card = ext_card

        section_lbl2 = ctk.CTkLabel(ext_card, text="🔖 Extensii Fișiere", font=("Segoe UI", 14, "bold"), text_color=_TEXT_PRI)
        section_lbl2.pack(anchor="w", padx=16, pady=(8, 2))

        self.ext_inner_frame = ctk.CTkFrame(ext_card, fg_color="transparent")
        self.ext_inner_frame.pack(fill="x", padx=16, pady=(0, 8))

        # Row frames container for clean horizontal row flow (no grid columns bug)
        self.ext_row_frames = [ctk.CTkFrame(self.ext_inner_frame, fg_color="transparent") for _ in range(3)]

        # Add controls container for unified responsive flow
        self.ext_add_frame = ctk.CTkFrame(self.ext_inner_frame, fg_color="transparent")
        self.new_ext_entry = ctk.CTkEntry(
            self.ext_add_frame, width=80, height=28, placeholder_text=".ext",
            fg_color=_ENTRY_BG, border_color=_CARD_BORDER,
            text_color=_TEXT_PRI, corner_radius=6, font=("Segoe UI", 12)
        )
        self.new_ext_entry.pack(side="left", padx=(0, 6))
        self.new_ext_entry.bind("<Return>", lambda e: self.add_extension())

        self.add_ext_btn = ctk.CTkButton(
            self.ext_add_frame, text="＋ Add", width=65, height=28,
            fg_color=_ACCENT, hover_color=_ACCENT_HOV,
            corner_radius=6, font=("Segoe UI", 12, "bold"),
            command=self.add_extension,
        )
        self.add_ext_btn.pack(side="left")

        self.ext_checkboxes = {}
        self.draw_extensions()

        # ──── Settings row ────
        settings_card = self._make_card(container)
        settings_card.grid(row=3, column=0, sticky="ew", pady=(0, 8))

        section_lbl3 = ctk.CTkLabel(settings_card, text="⚙ Setări", font=("Segoe UI", 14, "bold"), text_color=_TEXT_PRI)
        section_lbl3.pack(side="left", padx=16, pady=8)

        self._label(settings_card, text="Duplicate:").pack(side="left", padx=(16, 6), pady=8)
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
        ).pack(side="left", padx=(0, 10), pady=8)
        
        self._label(settings_card, text="Extracție:").pack(side="left", padx=(10, 6), pady=8)
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
        ).pack(side="left", padx=(0, 16), pady=8)

        # ──── Start & Progress ────
        self.start_btn = ctk.CTkButton(
            container,
            text="▶ Start Extracție",
            command=self.start_extraction,
            fg_color=_GREEN,
            hover_color=_GREEN_HOV,
            text_color="#0b0f19",
            height=38,
            corner_radius=10,
            font=("Segoe UI", 15, "bold"),
        )
        self.start_btn.grid(row=4, column=0, pady=(2, 4), sticky="ew")

        self.progress_bar = ctk.CTkProgressBar(
            container, progress_color=_ACCENT, fg_color=_ENTRY_BG, corner_radius=6, height=6
        )
        self.progress_bar.grid(row=5, column=0, sticky="ew", pady=(0, 4))
        self.progress_bar.set(0)

        # ──── Log terminal ────
        log_card = self._make_card(container)
        log_card.grid(row=6, column=0, sticky="nsew", pady=(0, 0))
        container.grid_rowconfigure(6, weight=1, minsize=100)

        self.log_box = ctk.CTkTextbox(
            log_card, state="disabled", font=("Cascadia Code", 12),
            fg_color=_LOG_BG, text_color="#a3e635", corner_radius=8, border_width=0
        )
        self.log_box.pack(fill="both", expand=True, padx=8, pady=8)

    def draw_extensions(self):
        """Create or update checkboxes for available extensions without destroying widgets."""
        if not hasattr(self, "ext_checkboxes"):
            self.ext_checkboxes = {}

        for ext in self.available_extensions:
            if ext not in self.extension_vars:
                self.extension_vars[ext] = ctk.BooleanVar(value=True if ext in ['.nc', '.pdf', '.dwg'] else False)
            if ext not in self.ext_checkboxes:
                cb = ctk.CTkCheckBox(
                    self.ext_inner_frame, text=ext, variable=self.extension_vars[ext],
                    fg_color=_ACCENT, hover_color=_ACCENT_HOV, border_color=_CARD_BORDER,
                    text_color=_TEXT_PRI, font=("Segoe UI", 12), corner_radius=4,
                    width=0, checkbox_width=18, checkbox_height=18,
                )
                self.ext_checkboxes[ext] = cb

        for ext in list(self.ext_checkboxes.keys()):
            if ext not in self.available_extensions:
                self.ext_checkboxes[ext].destroy()
                del self.ext_checkboxes[ext]

        self.reposition_extensions()

    def reposition_extensions(self, target_width=None):
        """Dynamically flow checkboxes and Add input into clean row frames, wrapping neatly without grid gaps."""
        if not hasattr(self, "ext_inner_frame") or not hasattr(self, "ext_add_frame") or not hasattr(self, "ext_row_frames"):
            return

        # Determine usable width
        if target_width is not None and target_width > 100:
            usable_w = max(250, target_width - 80)
        elif hasattr(self, "ext_card") and self.ext_card.winfo_width() > 100:
            usable_w = max(250, self.ext_card.winfo_width() - 36)
        elif self.winfo_width() > 100:
            usable_w = max(250, self.winfo_width() - 80)
        else:
            usable_w = 850

        # Build items to pack: checkboxes then ext_add_frame
        items = []
        for ext in self.available_extensions:
            cb = self.ext_checkboxes.get(ext)
            if cb:
                w = max(cb.winfo_reqwidth(), len(ext) * 8 + 36) + 14
                items.append((cb, w))
        add_w = max(self.ext_add_frame.winfo_reqwidth(), 145) + 8
        items.append((self.ext_add_frame, add_w))

        # Compute rows based on usable_w
        rows = [[]]
        cur_w = 0
        for widget, w in items:
            if cur_w + w > usable_w and cur_w > 0:
                rows.append([(widget, w)])
                cur_w = w
            else:
                rows[-1].append((widget, w))
                cur_w += w

        # Optimization: Do nothing if the packing signature has not changed
        signature = tuple(tuple(id(w) for w, _ in r) for r in rows)
        if getattr(self, "_last_ext_flow_signature", None) == signature:
            return
        self._last_ext_flow_signature = signature

        # Ensure enough row frames exist
        while len(self.ext_row_frames) < len(rows):
            rf = ctk.CTkFrame(self.ext_inner_frame, fg_color="transparent")
            self.ext_row_frames.append(rf)

        # Repack row frames and items
        for r_idx, r_items in enumerate(rows):
            rf = self.ext_row_frames[r_idx]
            rf.pack(fill="x", pady=2, anchor="w")
            for widget, _ in r_items:
                widget.pack_forget()
                widget.pack(in_=rf, side="left", padx=(0, 14), pady=2)

        # Unpack any unused row frames
        for r_idx in range(len(rows), len(self.ext_row_frames)):
            self.ext_row_frames[r_idx].pack_forget()

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

    def _on_toggle_auto_search(self):
        """Toggle UI between manual project folder selection and Parent Folder + Project search."""
        if hasattr(self, "cb_auto_proj") and hasattr(self, "proj_search_frame"):
            if self.extractor_auto_search.get():
                self.lbl_extractor_source.configure(text="Folder Mamă Clienți:")
                self.proj_search_frame.grid(row=3, column=0, columnspan=3, padx=16, pady=4, sticky="ew")
            else:
                self.lbl_extractor_source.configure(text="Sursă:")
                self.proj_search_frame.grid_forget()
            self.save_settings()

    def _on_toggle_extract_assembly(self):
        """Toggle UI visibility of Assembly Excel selector frame."""
        if hasattr(self, "cb_extract_assembly") and hasattr(self, "assembly_frame"):
            if self.extract_assembly_var.get():
                self.assembly_frame.grid(row=7, column=0, columnspan=3, padx=16, pady=(0, 10), sticky="ew")
            else:
                self.assembly_frame.grid_forget()
            self.save_settings()

    def browse_assembly_excel(self):
        """Browse Assembly Excel (opens in File Extractor Destinatie folder)."""
        initial = self._get_browse_initial_dir("dest_based", self.assembly_excel_file.get())
        f = filedialog.askopenfilename(
            title="Selectează Fișierul Excel cu Ansambluri",
            initialdir=initial,
            filetypes=[("Excel Files", "*.xlsx *.xls *.xlsm"), ("All Files", "*.*")]
        )
        if f:
            norm = os.path.normpath(f)
            self.assembly_excel_file.set(norm)
            self.save_settings()

    def is_assembly_match(self, filename, asmb_code):
        """Check if a drawing filename corresponds to asmb_code with hyphen/revision tolerance."""
        stem = re.sub(r'\.[^.]+$', '', filename)
        clean_asmb = re.sub(r'[-_\s\.]', '', asmb_code).upper()
        if not clean_asmb:
            return False
        
        # 1. Exact match with stripped punctuation
        clean_stem = re.sub(r'[-_\s\.]', '', stem).upper()
        if clean_stem == clean_asmb:
            return True

        # 2. Token / chunk match: split by underscore, space, parentheses, brackets
        chunks = re.split(r'[\s_()\[\],;]+', stem)
        for c in chunks:
            clean_c = re.sub(r'[-_\s\.]', '', c).upper()
            if clean_c == clean_asmb:
                return True

        # 3. Regex with word-like boundaries allowing hyphens inside asmb
        asmb_pattern = r'[-_\s]*'.join([re.escape(ch) for ch in clean_asmb])
        regex = rf'(?:^|[\s_()\[\],;])({asmb_pattern})(?:$|[\s_()\[\],;]|(?![A-Za-z0-9]))'
        if re.search(regex, stem, re.IGNORECASE):
            return True
        return False

    def parse_assembly_excel(self, excel_path, custom_col=""):
        """Extract unique assembly codes from all sheets of an assembly Excel file."""
        asmb_codes = []
        seen = set()
        
        try:
            xl = pd.ExcelFile(excel_path)
            try:
                for sheet_name in xl.sheet_names:
                    try:
                        df = xl.parse(sheet_name, header=None)
                    except Exception:
                        continue
                    if df.empty or len(df) < 2:
                        continue

                    header_row_idx = None
                    asmb_col_idx = None

                    asmb_keywords = ["denumire ansamblu", "cod ansamblu", "ansamblu", "assembly", "marca", "main part", "mark"]
                    footer_keywords = ["total", "intocmit", "verificat", "aprobat", "greutate", "kg", "data", "semnatura", "buc", "suma"]

                    max_header_scan = min(15, len(df))
                    for r in range(max_header_scan):
                        row_vals = [str(x).strip().lower() for x in df.iloc[r].tolist() if pd.notna(x)]
                        if len(row_vals) < 2:
                            continue

                        if custom_col:
                            c_lower = custom_col.lower()
                            for c_idx in range(len(df.columns)):
                                val = str(df.iloc[r, c_idx]).strip().lower()
                                if pd.notna(df.iloc[r, c_idx]) and c_lower in val:
                                    header_row_idx = r
                                    asmb_col_idx = c_idx
                                    break
                            if asmb_col_idx is not None:
                                break

                        for c_idx in range(len(df.columns)):
                            val = str(df.iloc[r, c_idx]).strip().lower()
                            if pd.isna(df.iloc[r, c_idx]) or not val or val == 'nan':
                                continue
                            if 'lista' in val or 'proiect' in val:
                                continue
                            for kw in asmb_keywords:
                                if kw == val or (kw in val and len(val) < 30):
                                    header_row_idx = r
                                    asmb_col_idx = c_idx
                                    break
                            if asmb_col_idx is not None:
                                break
                        if asmb_col_idx is not None:
                            break

                    if asmb_col_idx is None:
                        if len(df.columns) >= 2:
                            asmb_col_idx = 1
                            header_row_idx = 0

                    start_row = (header_row_idx + 1) if header_row_idx is not None else 1
                    for r in range(start_row, len(df)):
                        val = str(df.iloc[r, asmb_col_idx]).strip()
                        val_lower = val.lower()
                        
                        if not val or val_lower in ("nan", "none", "-", "null", "undefined"):
                            continue
                        if any(kw in val_lower for kw in footer_keywords):
                            continue
                        if any(kw in val_lower for kw in asmb_keywords):
                            continue
                        val_clean = val.strip()
                        if len(val_clean) >= 2 and val_clean not in seen:
                            seen.add(val_clean)
                            asmb_codes.append(val_clean)
            finally:
                try:
                    xl.close()
                except Exception:
                    pass
        except Exception as e:
            self.log(f"Eroare citire Excel Ansambluri: {e}")
            
        return asmb_codes

    def extract_assembly_drawings(self, source_path, dest_path, assembly_codes, selected_exts):
        """Search and extract assembly drawings into a dedicated 'Desene Ansamblu' folder and generate a status report."""
        dest_asmb_dir = Path(dest_path) / "Desene Ansamblu"
        dest_asmb_dir.mkdir(parents=True, exist_ok=True)
        
        drawing_exts = {ext for ext in selected_exts if ext in ('.pdf', '.dwg', '.dxf', '.step', '.stp')}
        if not drawing_exts:
            drawing_exts = {'.pdf', '.dwg'}

        self.log(f"  Căutare desene de ansamblu ({', '.join(sorted(drawing_exts))}) în {source_path}...")
        
        all_drawings = []
        for root, _, files in os.walk(source_path, followlinks=True):
            for f in files:
                ext = os.path.splitext(f)[1].lower()
                if ext in drawing_exts:
                    all_drawings.append(Path(root) / f)

        self.log(f"  Total fișiere desene indexate în proiect: {len(all_drawings)}")

        asmb_results = []
        found_count = 0
        missing_count = 0

        for asmb_code in assembly_codes:
            matched_files = []
            for filepath in all_drawings:
                if self.is_assembly_match(filepath.name, asmb_code):
                    matched_files.append(filepath)

            if matched_files:
                if self.duplicate_handling.get() == "Keep Oldest":
                    matched_files.sort(key=lambda p: (p.stat().st_mtime, p.name))
                else:
                    matched_files.sort(key=lambda p: (p.stat().st_mtime, p.name), reverse=True)
                
                copied_for_asmb = []
                by_ext = {}
                for mf in matched_files:
                    e = mf.suffix.lower()
                    if e not in by_ext:
                        by_ext[e] = mf

                for e, target_file in by_ext.items():
                    dest_file = dest_asmb_dir / target_file.name
                    try:
                        shutil.copy2(target_file, dest_file)
                        copied_for_asmb.append(target_file.name)
                    except Exception as err:
                        self.log(f"    Eroare copiere {target_file.name}: {err}")

                found_count += 1
                asmb_results.append({
                    "code": asmb_code,
                    "status": "GĂSIT",
                    "files": ", ".join(copied_for_asmb),
                    "source": str(matched_files[0].parent)
                })
                self.log(f"  ✔ Ansamblu {asmb_code} -> Găsit: {', '.join(copied_for_asmb)}")
            else:
                missing_count += 1
                asmb_results.append({
                    "code": asmb_code,
                    "status": "LIPSĂ",
                    "files": "-",
                    "source": "-"
                })
                self.log(f"  ❌ Ansamblu {asmb_code} -> LIPSĂ")

        self.log(f"\n  Rezumat Ansambluri: {found_count} găsite, {missing_count} lipsă din {len(assembly_codes)} cerute.")

        # Raport_Status_Ansambluri.xlsx
        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "Status Ansambluri"
            
            hdr_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
            hdr_fill = PatternFill(fill_type="solid", fgColor="1F4068")
            green_fill = PatternFill(fill_type="solid", fgColor="D4EDDA")
            green_font = Font(name="Segoe UI", size=10, bold=True, color="155724")
            red_fill = PatternFill(fill_type="solid", fgColor="FFCCCC")
            red_font = Font(name="Segoe UI", size=10, bold=True, color="721C24")
            reg_font = Font(name="Segoe UI", size=10)
            thin_border = Border(
                left=Side(style='thin', color='CCCCCC'),
                right=Side(style='thin', color='CCCCCC'),
                top=Side(style='thin', color='CCCCCC'),
                bottom=Side(style='thin', color='CCCCCC')
            )

            headers = ["Nr. Crt.", "Cod Ansamblu", "Status", "Fișiere Găsite", "Director Sursă"]
            ws.append(headers)
            for c_idx in range(1, 6):
                cell = ws.cell(row=1, column=c_idx)
                cell.font = hdr_font
                cell.fill = hdr_fill
                cell.alignment = Alignment(horizontal="center", vertical="center")

            for idx, item in enumerate(asmb_results, start=1):
                row_idx = idx + 1
                is_found = (item["status"] == "GĂSIT")
                
                c_num = ws.cell(row=row_idx, column=1, value=idx)
                c_code = ws.cell(row=row_idx, column=2, value=item["code"])
                c_status = ws.cell(row=row_idx, column=3, value=item["status"])
                c_files = ws.cell(row=row_idx, column=4, value=item["files"])
                c_source = ws.cell(row=row_idx, column=5, value=item["source"])

                c_num.alignment = Alignment(horizontal="center")
                c_status.alignment = Alignment(horizontal="center")
                c_status.fill = green_fill if is_found else red_fill
                c_status.font = green_font if is_found else red_font

                for c in (c_num, c_code, c_files, c_source):
                    c.font = reg_font
                for col_idx in range(1, 6):
                    ws.cell(row=row_idx, column=col_idx).border = thin_border

            for col in ws.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

            report_path = os.path.join(str(dest_asmb_dir), "Raport_Status_Ansambluri.xlsx")
            safe_save_workbook(wb, report_path, self.log)
            self.log(f"  📄 Raport generat: {report_path}")
        except Exception as e:
            self.log(f"  Eroare generare raport status ansambluri: {e}")

    def find_matching_project_folders(self, parent_dir, search_term, progress_callback=None):
        """
        Scan parent_dir strictly following the hierarchy:
        FOLDER MAMA -> CLIENT -> PROIECT -> FISIERE PROIECT.
        Searches ONLY for the project directory name and NEVER descends inside the project folder itself.
        """
        if not parent_dir or not os.path.isdir(parent_dir):
            return []

        tokens = [t.strip() for t in re.split(r'[\s\-_\.,\(\)\[\]]+', search_term) if t.strip()]
        if not tokens:
            return []

        def match_name(name):
            for token in tokens:
                pattern = rf'(?:^|[\s\-_\.,\(\)\[\]]){re.escape(token)}(?:$|[\s\-_\.,\(\)\[\]])'
                if not re.search(pattern, name, re.IGNORECASE):
                    return False
            return True

        matches = []
        parent_base = os.path.basename(parent_dir.rstrip('/\\'))
        if parent_base and match_name(parent_base):
            return [parent_dir]

        # Level 1: Direct children of parent_dir (Clients or direct Projects)
        client_dirs = []
        try:
            with os.scandir(parent_dir) as it:
                for entry in it:
                    try:
                        if entry.is_dir(follow_symlinks=False):
                            if match_name(entry.name):
                                matches.append(entry.path)
                            else:
                                client_dirs.append((entry.name, entry.path))
                    except (PermissionError, OSError):
                        continue
        except (PermissionError, OSError) as e:
            if progress_callback:
                progress_callback(f"Eroare accesare folder mamă: {e}")
            return matches

        # Level 2: Direct children of client folders (Projects)
        total_clients = len(client_dirs)
        for idx, (client_name, client_path) in enumerate(client_dirs, 1):
            if progress_callback and (idx % 5 == 0 or idx == 1 or idx == total_clients):
                progress_callback(f"Verificare client ({idx}/{total_clients}): {client_name}...")

            try:
                with os.scandir(client_path) as it:
                    for proj_entry in it:
                        try:
                            # IMPORTANT: Check directory name ONLY! NEVER scan inside proj_entry!
                            if proj_entry.is_dir(follow_symlinks=False):
                                if match_name(proj_entry.name):
                                    if proj_entry.path not in matches:
                                        matches.append(proj_entry.path)
                        except (PermissionError, OSError):
                            continue
            except (PermissionError, OSError):
                continue

        return matches

    def resolve_project_folder(self):
        """Resolve project folder from parent dir and search term. Returns folder path or None."""
        parent_dir = self.source_dir.get().strip()
        search_term = self.extractor_proj_search.get().strip()

        if not parent_dir or not os.path.exists(parent_dir):
            messagebox.showwarning("Atenție", "Te rog selectează un Folder Mamă valid.")
            return None

        if not search_term:
            messagebox.showwarning("Atenție", "Te rog introdu numărul sau denumirea proiectului căutat (ex: 1147).")
            return None

        matches = self.find_matching_project_folders(parent_dir, search_term)

        if not matches:
            messagebox.showwarning(
                "Proiect Negăsit",
                f"Nu s-a găsit niciun folder de proiect care să conțină '{search_term}' ca și cuvânt întreg în:\n{parent_dir}"
            )
            return None
        elif len(matches) == 1:
            chosen = matches[0]
            self.log(f"✅ Proiect identificat automat: {os.path.basename(chosen)} ({chosen})")
            return chosen
        else:
            dlg = ProjectSelectionDialog(self, matches, search_term)
            self.wait_window(dlg)
            chosen = dlg.selected_path
            if chosen:
                self.log(f"✅ Proiect selectat manual: {os.path.basename(chosen)} ({chosen})")
                return chosen
            else:
                self.log("⚠ Selecția proiectului a fost anulată.")
                return None

    def find_project_folder_interactive(self):
        """Button click handler to test and verify project search asynchronously without freezing UI."""
        parent_dir = self.source_dir.get().strip()
        search_term = self.extractor_proj_search.get().strip()

        if not parent_dir or not os.path.exists(parent_dir):
            messagebox.showwarning("Atenție", "Te rog selectează un Folder Mamă valid.")
            return

        if not search_term:
            messagebox.showwarning("Atenție", "Te rog introdu numărul sau denumirea proiectului căutat (ex: 1147).")
            return

        if getattr(self, "_proj_search_in_progress", False):
            return

        self._proj_search_in_progress = True
        if hasattr(self, "btn_verify_proj"):
            self.btn_verify_proj.configure(text="⏳ Căutare...", state="disabled")
        self.log(f"\n🔍 Căutare asincronă proiect '{search_term}' în: {parent_dir}...")

        def _worker():
            try:
                def _prog(msg):
                    self.after(0, lambda m=msg: self.log(f"  {m}"))
                matches = self.find_matching_project_folders(parent_dir, search_term, progress_callback=_prog)
            except Exception as e:
                matches = []
                self.log(f"Eroare căutare proiect: {e}")
            self.after(0, lambda: self._on_verify_project_complete(matches, search_term, parent_dir))

        threading.Thread(target=_worker, daemon=True).start()

    def _on_verify_project_complete(self, matches, search_term, parent_dir):
        """Callback on UI thread when asynchronous project search completes."""
        self._proj_search_in_progress = False
        if hasattr(self, "btn_verify_proj"):
            self.btn_verify_proj.configure(text="🔍 Verifică", state="normal")

        if not matches:
            self.log(f"❌ Proiectul '{search_term}' nu a fost găsit în {parent_dir}.")
            messagebox.showwarning(
                "Proiect Negăsit",
                f"Nu s-a găsit niciun folder de proiect care să conțină '{search_term}' ca și cuvânt întreg în:\n\n{parent_dir}"
            )
        elif len(matches) == 1:
            chosen = matches[0]
            self.log(f"✅ Proiect identificat cu succes:\n   Folder: {os.path.basename(chosen)}\n   Cale: {chosen}")
            messagebox.showinfo(
                "Proiect Identificat",
                f"A fost identificat cu succes folderul de proiect:\n\n📁 {os.path.basename(chosen)}\n\nCale: {chosen}"
            )
        else:
            self.log(f"⚠ S-au găsit {len(matches)} foldere care corespund termenului '{search_term}'. Deschidere dialog selecție...")
            dlg = ProjectSelectionDialog(self, matches, search_term)
            self.wait_window(dlg)
            chosen = dlg.selected_path
            if chosen:
                self.log(f"✅ Proiect selectat: {os.path.basename(chosen)} ({chosen})")
            else:
                self.log("⚠ Selecția proiectului a fost anulată.")

    def browse_source(self):
        """Browse source folder exclusively for File Extractor (Independent memory)."""
        initial = self._get_browse_initial_dir("extractor_source", self.source_dir.get())
        title = "Selectează Folderul Mamă al Clienților" if self.extractor_auto_search.get() else "Selectare Director Sursă (File Extractor)"
        folder = filedialog.askdirectory(title=title, initialdir=initial)
        if folder:
            norm = os.path.normpath(folder)
            self.source_dir.set(norm)
            self.extractor_source_last_dir = norm
            self.save_settings()

    def browse_dest(self):
        """Browse destination folder (independent memory for File Extractor)."""
        initial = self._get_browse_initial_dir("extractor_dest", self.dest_dir.get())
        folder = filedialog.askdirectory(title="Selectare Director Destinație", initialdir=initial)
        if folder:
            norm = os.path.normpath(folder)
            self.dest_dir.set(norm)
            self.extractor_dest_last_dir = norm
            self.global_last_dir = norm
            self.save_settings()

    def browse_excel(self):
        """Browse Excel file (opens in File Extractor Destinație folder)."""
        initial = self._get_browse_initial_dir("dest_based", self.excel_file.get())
        file = filedialog.askopenfilename(
            title="Selectare Fișier Excel",
            initialdir=initial,
            filetypes=[("Excel files", "*.xlsx *.xls *.xlsm"), ("All Files", "*.*")]
        )
        if file:
            norm = os.path.normpath(file)
            self.excel_file.set(norm)
            self.save_settings()

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
        if self.extractor_auto_search.get():
            effective_source = self.resolve_project_folder()
            if not effective_source:
                return
        else:
            effective_source = self.source_dir.get().strip()
            if not effective_source or not os.path.isdir(effective_source):
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

        if hasattr(self, "_update_tab_buttons_style"):
            self._update_tab_buttons_style()
        self.save_settings()
        self.start_btn.configure(state="disabled")
        self.set_progress(0)
        self.log("=" * 50)
        self.log("Extracția a început...")
        self.log(f"📁 Director sursă proiect: {effective_source}")
        
        threading.Thread(target=self.run_extraction_process, args=(selected_exts, effective_source), daemon=True).start()

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

    def run_extraction_process(self, selected_exts, effective_source=None):
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
            source_dir_to_use = effective_source if effective_source else self.source_dir.get()
            source_path = Path(source_dir_to_use)
            
            all_files = []
            for root, _, files in os.walk(source_path, followlinks=True):
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
                
                # Identificare repere lipsă și evidențiere în rapoartele de piese
                missing_parts = set()
                for part in part_names:
                    is_missing = False
                    for ext in display_exts:
                        if ext == '.nc':
                            found = ('.nc' in part_to_ext_found[part] or '.nc1' in part_to_ext_found[part])
                        else:
                            found = (ext in part_to_ext_found[part])
                        if not found:
                            is_missing = True
                            break
                    if is_missing:
                        missing_parts.add(part)
                
                if missing_parts:
                    missing_norms = {self._normalize_name(p) for p in missing_parts}
                    self.log(f"🖍️ Evidențiere {len(missing_parts)} repere lipsă cu roșu în rapoartele de piese...")
                    self.highlight_missing_parts_in_reports(str(dest_path), missing_norms)
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

            # ── Extracție Desene de Ansamblu din Excel Ansambluri ──
            if self.extract_assembly_var.get() and self.assembly_excel_file.get().strip():
                self.log("\n── Extracție Desene de Ansamblu ──")
                asmb_excel = self.assembly_excel_file.get().strip()
                if os.path.exists(asmb_excel):
                    asmb_codes = self.parse_assembly_excel(asmb_excel, self.assembly_col_name.get().strip())
                    self.log(f"  Total ansambluri unice identificate în Excel: {len(asmb_codes)}")
                    if asmb_codes:
                        self.extract_assembly_drawings(
                            source_path=source_path,
                            dest_path=dest_path,
                            assembly_codes=asmb_codes,
                            selected_exts=selected_exts
                        )
                else:
                    self.log(f"  ⚠ Fișierul Excel de ansambluri nu există: {asmb_excel}")

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
        container.pack(fill="both", expand=True, padx=12, pady=8)
        container.grid_columnconfigure(0, weight=1)
        
        subtitle_lbl = ctk.CTkLabel(
            container,
            text="Sortare fișiere NC/DSTV în foldere pe categorii cu rapoarte Excel",
            font=("Segoe UI", 12),
            text_color=_TEXT_SEC,
        )
        subtitle_lbl.grid(row=0, column=0, pady=(0, 6), sticky="w")

        # ──── Source Folder Card ────
        paths_card = self._make_card(container)
        paths_card.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        paths_card.grid_columnconfigure(0, weight=0)
        paths_card.grid_columnconfigure(1, weight=1)
        paths_card.grid_columnconfigure(2, weight=0)

        ctk.CTkLabel(paths_card, text="📂 Director Sursă", font=("Segoe UI", 14, "bold"), text_color=_TEXT_PRI).grid(
            row=0, column=0, columnspan=3, padx=16, pady=(8, 4), sticky="w")

        self._label(paths_card, text="Sursă:").grid(row=1, column=0, padx=(16, 6), pady=(4, 10), sticky="w")
        self._entry(paths_card, self.nc_source_folder).grid(row=1, column=1, padx=4, pady=(4, 10), sticky="ew")
        self._browse_btn(paths_card, self.nc_browse_source).grid(row=1, column=2, padx=(4, 16), pady=(4, 10))

        # ──── Excel / BOM Card (Optional) ────
        excel_card = self._make_card(container)
        excel_card.grid(row=2, column=0, sticky="ew", pady=(0, 8))
        excel_card.grid_columnconfigure(0, weight=0)
        excel_card.grid_columnconfigure(1, weight=1)
        excel_card.grid_columnconfigure(2, weight=0)

        excel_header_frame = ctk.CTkFrame(excel_card, fg_color="transparent")
        excel_header_frame.grid(row=0, column=0, columnspan=3, padx=16, pady=(8, 2), sticky="w")
        ctk.CTkLabel(excel_header_frame, text="📊 Excel / BOM", font=("Segoe UI", 14, "bold"), text_color=_TEXT_PRI).pack(side="left")
        ctk.CTkLabel(excel_header_frame, text="  (opțional — pentru raportul cu nr. bucăți)", font=("Segoe UI", 11), text_color=_ORANGE).pack(side="left")

        self._label(excel_card, text="Excel:").grid(row=1, column=0, padx=(16, 6), pady=4, sticky="w")
        self._entry(excel_card, self.nc_excel_file).grid(row=1, column=1, padx=4, pady=4, sticky="ew")
        self._browse_btn(excel_card, self.nc_browse_excel).grid(row=1, column=2, padx=(4, 16), pady=4)

        col_frame = ctk.CTkFrame(excel_card, fg_color="transparent")
        col_frame.grid(row=2, column=0, columnspan=3, padx=16, pady=(2, 8), sticky="w")
        
        self._label(col_frame, text="Col. Piese:", font=("Segoe UI", 11)).pack(side="left", padx=(0, 4))
        self._entry(col_frame, self.nc_part_col, width=100, placeholder_text="auto").pack(side="left", padx=(0, 16))
        self._label(col_frame, text="Col. Cantități:", font=("Segoe UI", 11)).pack(side="left", padx=(0, 4))
        self._entry(col_frame, self.nc_qty_col, width=100, placeholder_text="auto").pack(side="left")

        # ──── Settings Card ────
        settings_card = self._make_card(container)
        settings_card.grid(row=3, column=0, sticky="ew", pady=(0, 8))

        ctk.CTkLabel(settings_card, text="⚙ Setări Sortare", font=("Segoe UI", 14, "bold"), text_color=_TEXT_PRI).pack(
            anchor="w", padx=16, pady=(8, 4))

        # Action row
        action_frame = ctk.CTkFrame(settings_card, fg_color="transparent")
        action_frame.pack(anchor="w", padx=16, pady=(0, 4))
        self._label(action_frame, text="Acțiune:").pack(side="left", padx=(0, 10))
        ctk.CTkRadioButton(action_frame, text="Copiază fișiere", variable=self.nc_action_var, value="Copy",
                           fg_color=_ACCENT, hover_color=_ACCENT_HOV, text_color=_TEXT_PRI).pack(side="left", padx=(0, 16))
        ctk.CTkRadioButton(action_frame, text="Mută fișiere", variable=self.nc_action_var, value="Move",
                           fg_color=_ACCENT, hover_color=_ACCENT_HOV, text_color=_TEXT_PRI).pack(side="left")

        # Options – 2 balanced columns
        opt_container = ctk.CTkFrame(settings_card, fg_color="transparent")
        opt_container.pack(fill="x", padx=16, pady=(2, 8))
        opt_container.grid_columnconfigure(0, weight=1)
        opt_container.grid_columnconfigure(1, weight=1)

        ctk.CTkCheckBox(opt_container, text="Sortare profile Laminat la Cald / Rece", variable=self.nc_opt_hot_cold,
                        fg_color=_ACCENT, hover_color=_ACCENT_HOV, border_color=_CARD_BORDER,
                        text_color=_TEXT_PRI, font=("Segoe UI", 12)).grid(row=0, column=0, sticky="w", pady=2)
        ctk.CTkCheckBox(opt_container, text="Separare Profile de Țevi", variable=self.nc_opt_separate_tubes,
                        fg_color=_ACCENT, hover_color=_ACCENT_HOV, border_color=_CARD_BORDER,
                        text_color=_TEXT_PRI, font=("Segoe UI", 12)).grid(row=1, column=0, sticky="w", pady=2)
        ctk.CTkCheckBox(opt_container, text="Validare Geometrică (Erori 3D)", variable=self.nc_opt_verify_geometry,
                        fg_color=_ACCENT, hover_color=_ACCENT_HOV, border_color=_CARD_BORDER,
                        text_color=_TEXT_PRI, font=("Segoe UI", 12)).grid(row=2, column=0, sticky="w", pady=2)
        ctk.CTkCheckBox(opt_container, text="Fără antet (header) în rapoarte Excel", variable=self.nc_opt_no_header,
                        fg_color=_ACCENT, hover_color=_ACCENT_HOV, border_color=_CARD_BORDER,
                        text_color=_TEXT_PRI, font=("Segoe UI", 12)).grid(row=0, column=1, sticky="w", pady=2, padx=(10, 0))
        ctk.CTkCheckBox(opt_container, text="Fără culori în rapoarte Excel", variable=self.nc_opt_no_colors,
                        fg_color=_ACCENT, hover_color=_ACCENT_HOV, border_color=_CARD_BORDER,
                        text_color=_TEXT_PRI, font=("Segoe UI", 12)).grid(row=1, column=1, sticky="w", pady=2, padx=(10, 0))

        # ──── Start & Progress ────
        self.nc_start_btn = ctk.CTkButton(
            container,
            text="▶ Start Sortare",
            command=self.nc_start_sorting,
            fg_color=_GREEN,
            hover_color=_GREEN_HOV,
            text_color="#0b0f19",
            height=38,
            corner_radius=10,
            font=("Segoe UI", 15, "bold"),
        )
        self.nc_start_btn.grid(row=4, column=0, pady=(2, 4), sticky="ew")

        self.nc_progress_bar = ctk.CTkProgressBar(
            container, progress_color=_ACCENT, fg_color=_ENTRY_BG, corner_radius=6, height=6
        )
        self.nc_progress_bar.grid(row=5, column=0, sticky="ew", pady=(0, 4))
        self.nc_progress_bar.set(0)

        # ──── Log terminal ────
        log_card = self._make_card(container)
        log_card.grid(row=6, column=0, sticky="nsew", pady=(0, 0))
        container.grid_rowconfigure(6, weight=1, minsize=100)

        self.nc_log_box = ctk.CTkTextbox(
            log_card, state="disabled", font=("Cascadia Code", 12),
            fg_color=_LOG_BG, text_color="#a3e635", corner_radius=8, border_width=0
        )
        self.nc_log_box.pack(fill="both", expand=True, padx=8, pady=8)

    def nc_browse_source(self):
        """Browse source folder for NC Sorter (opens in File Extractor Destinație folder)."""
        initial = self._get_browse_initial_dir("dest_based", self.nc_source_folder.get())
        folder = filedialog.askdirectory(title="Selectare Director Sursă", initialdir=initial)
        if folder:
            norm = os.path.normpath(folder)
            self.nc_source_folder.set(norm)
            self.save_settings()

    def nc_browse_excel(self):
        """Browse BOM Excel for NC Sorter (opens in File Extractor Destinație folder)."""
        initial = self._get_browse_initial_dir("dest_based", self.nc_excel_file.get())
        file = filedialog.askopenfilename(
            title="Selectare Fișier Excel BOM",
            initialdir=initial,
            filetypes=[("Excel files", "*.xlsx *.xls *.xlsm"), ("All Files", "*.*")]
        )
        if file:
            norm = os.path.normpath(file)
            self.nc_excel_file.set(norm)
            self.save_settings()

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

        if hasattr(self, "_update_tab_buttons_style"):
            self._update_tab_buttons_style()
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
        container.pack(fill="both", expand=True, padx=12, pady=8)
        container.grid_columnconfigure(0, weight=1)

        subtitle_lbl = ctk.CTkLabel(
            container,
            text="Modifică și ajustează fișierele NC (Text, Cantități, Rotație)",
            font=("Segoe UI", 12),
            text_color=_TEXT_SEC,
        )
        subtitle_lbl.grid(row=0, column=0, pady=(0, 6), sticky="w")

        # ──── Settings Card ────
        settings_card = self._make_card(container)
        settings_card.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        settings_card.grid_columnconfigure(0, weight=0)
        settings_card.grid_columnconfigure(1, weight=1)
        settings_card.grid_columnconfigure(2, weight=0)

        ctk.CTkLabel(settings_card, text="⚙ Setări Globale", font=("Segoe UI", 14, "bold"), text_color=_TEXT_PRI).grid(row=0, column=0, columnspan=3, padx=16, pady=(8, 4), sticky="w")
        
        self._label(settings_card, text="Folder NC:").grid(row=1, column=0, padx=(16, 6), pady=4, sticky="w")
        self._entry(settings_card, self.mod_nc_dir).grid(row=1, column=1, padx=4, pady=4, sticky="ew")
        self._browse_btn(settings_card, self.mod_nc_browse_dir).grid(row=1, column=2, padx=(4, 16), pady=4)

        cb_subfolder = ctk.CTkCheckBox(settings_card, text="Creează subfolder 'Modified' (copiază fișierele)", variable=self.mod_nc_create_subfolder)
        cb_subfolder.grid(row=2, column=1, padx=4, pady=(4, 10), sticky="w")

        # ──── Actions Card ────
        actions_card = self._make_card(container)
        actions_card.grid(row=2, column=0, sticky="ew", pady=(0, 8))
        actions_card.grid_columnconfigure(0, weight=0)
        actions_card.grid_columnconfigure(1, weight=1)
        actions_card.grid_columnconfigure(2, weight=0)

        # A. Text Replace
        ctk.CTkLabel(actions_card, text="A. Înlocuire Text", font=("Segoe UI", 14, "bold"), text_color=_TEXT_PRI).grid(row=0, column=0, columnspan=3, padx=16, pady=(8, 4), sticky="w")
        
        self._label(actions_card, text="Caută:").grid(row=1, column=0, padx=(16, 6), pady=4, sticky="w")
        self._entry(actions_card, self.mod_nc_find).grid(row=1, column=1, padx=4, pady=4, sticky="ew")
        self._label(actions_card, text="Înlocuiește:").grid(row=2, column=0, padx=(16, 6), pady=4, sticky="w")
        self._entry(actions_card, self.mod_nc_replace).grid(row=2, column=1, padx=4, pady=4, sticky="ew")
        btn_replace = ctk.CTkButton(actions_card, text="Înlocuiește Text", width=120, height=32, font=("Segoe UI", 12, "bold"), command=self.mod_nc_replace_text)
        btn_replace.grid(row=1, column=2, rowspan=2, padx=(4, 16), pady=4)

        # B. Rectify Number
        ctk.CTkLabel(actions_card, text="B. Rectificare Cantități din Excel", font=("Segoe UI", 14, "bold"), text_color=_TEXT_PRI).grid(row=3, column=0, columnspan=3, padx=16, pady=(8, 4), sticky="w")
        self._label(actions_card, text="Fișier Excel:").grid(row=4, column=0, padx=(16, 6), pady=4, sticky="w")
        self._entry(actions_card, self.mod_nc_excel).grid(row=4, column=1, padx=4, pady=4, sticky="ew")
        self._browse_btn(actions_card, self.mod_nc_browse_excel).grid(row=4, column=2, padx=(4, 16), pady=4)
        btn_rectify = ctk.CTkButton(actions_card, text="Rectifică Număr", width=120, height=32, font=("Segoe UI", 12, "bold"), command=self.mod_nc_rectify_qty)
        btn_rectify.grid(row=5, column=1, padx=4, pady=(4, 10), sticky="w")

        # ──── Log terminal ────
        log_card = self._make_card(container)
        log_card.grid(row=3, column=0, sticky="nsew", pady=(0, 0))
        container.grid_rowconfigure(3, weight=1, minsize=100)

        self.mod_nc_log_box = ctk.CTkTextbox(
            log_card,
            state="disabled",
            fg_color=_LOG_BG,
            text_color="#a3e635",
            corner_radius=8,
            border_width=0,
            font=("Cascadia Code", 12),
        )
        self.mod_nc_log_box.pack(fill="both", expand=True, padx=8, pady=8)

    def mod_nc_browse_dir(self):
        """Browse NC folder for Modify NC (opens in File Extractor Destinație + \\Sorted)."""
        initial = self._get_browse_initial_dir("dest_sorted_based", self.mod_nc_dir.get())
        d = filedialog.askdirectory(title="Selectează folderul cu fișiere NC", initialdir=initial)
        if d:
            norm = os.path.normpath(d)
            self.mod_nc_dir.set(norm)
            self.save_settings()

    def mod_nc_browse_excel(self):
        """Browse quantities Excel for Modify NC (opens in File Extractor Destinație + \\Sorted)."""
        initial = self._get_browse_initial_dir("dest_sorted_based", self.mod_nc_excel.get())
        f = filedialog.askopenfilename(
            title="Selectează fișierul Excel cu cantități",
            initialdir=initial,
            filetypes=[("Excel files", "*.xlsx *.xls *.xlsm"), ("All Files", "*.*")]
        )
        if f:
            norm = os.path.normpath(f)
            self.mod_nc_excel.set(norm)
            self.save_settings()

    def mod_nc_log(self, msg):
        self.after(0, self._mod_nc_insert_log, msg)

    def _mod_nc_insert_log(self, msg):
        self.mod_nc_log_box.configure(state="normal")
        self.mod_nc_log_box.insert("end", msg + "\n")
        self.mod_nc_log_box.see("end")
        self.mod_nc_log_box.configure(state="disabled")

    def _mod_nc_get_files(self):
        folder = self.mod_nc_dir.get()
        if not folder or not os.path.exists(folder):
            messagebox.showwarning("Atenție", "Te rog selectează un folder valid cu fișiere NC.")
            return None, None
        
        nc_files = []
        for root, dirs, files in os.walk(folder, followlinks=True):
            if os.path.basename(root).lower() == "modified":
                continue
            for f in files:
                if f.lower().endswith(('.nc', '.nc1')):
                    rel_p = os.path.relpath(os.path.join(root, f), folder)
                    nc_files.append(rel_p)

        if not nc_files:
            self.mod_nc_log("⚠ Nu s-au găsit fișiere NC în folderul selectat sau în subfolderele acestuia.")
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
            make_dir_with_permissions(os.path.dirname(dest_path))
            
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
            make_dir_with_permissions(os.path.dirname(dest_path))
            basename = os.path.splitext(os.path.basename(f))[0]
            
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
    #  DWG FINDER TAB & METHODS
    # ══════════════════════════════════════════════════════════════════════
    def _setup_treeview_styles(self):
        """Configure ttk Treeview styles to integrate smoothly with Dark / Light appearance."""
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        is_dark = (self.appearance_mode_var.get() == "Dark")
        bg_tree = "#0d1b2a" if is_dark else "#ffffff"
        fg_tree = "#e0e0e0" if is_dark else "#1c1c1e"
        hdr_bg = "#16213e" if is_dark else "#e5e5ea"
        hdr_fg = "#00b4d8" if is_dark else "#005bb5"
        sel_bg = "#007aff"

        style.configure(
            "DWG.Treeview",
            background=bg_tree,
            foreground=fg_tree,
            fieldbackground=bg_tree,
            font=("Segoe UI", 10),
            rowheight=26,
            borderwidth=0,
        )
        style.configure(
            "DWG.Treeview.Heading",
            background=hdr_bg,
            foreground=hdr_fg,
            font=("Segoe UI", 10, "bold"),
            relief="flat",
        )
        style.map(
            "DWG.Treeview",
            background=[("selected", sel_bg)],
            foreground=[("selected", "#ffffff")]
        )
        style.map(
            "DWG.Treeview.Heading",
            background=[("active", hdr_bg)],
            foreground=[("active", hdr_fg)]
        )

    def create_dwg_finder_tab(self):
        """Construct the layout for the DWG Finder tab with integrated DWG Tools."""
        self._setup_treeview_styles()

        # ── Top Subview Selector: DWG Finder vs Organizare Piese vs Regrupare DWG ──
        top_ctrl = ctk.CTkFrame(self.tab_dwg, fg_color="transparent")
        top_ctrl.pack(fill="x", padx=12, pady=(6, 2))

        self.dwg_subview_seg = ctk.CTkSegmentedButton(
            top_ctrl,
            values=[
                "🔍 Căutare DWG în Proiecte",
                "📦 Organizare Piese (BOM ➜ Piese)",
                "🔄 Regrupare DWG pe Grosimi"
            ],
            variable=self.dwg_subview_mode,
            command=self._on_dwg_subview_change,
            height=34,
            font=("Segoe UI", 12, "bold")
        )
        self.dwg_subview_seg.pack(side="left")

        # Container 1: Căutare DWG în Proiecte
        self.dwg_search_container = ctk.CTkFrame(self.tab_dwg, fg_color="transparent")
        container = self.dwg_search_container
        container.grid_columnconfigure(0, weight=1)
        container.grid_rowconfigure(5, weight=1)

        subtitle_lbl = ctk.CTkLabel(
            container,
            text="Caută și extrage desene DWG existente în arborele de proiecte pe baza pieselor PDF",
            font=("Segoe UI", 12),
            text_color=_TEXT_SEC,
        )
        subtitle_lbl.grid(row=0, column=0, pady=(0, 6), sticky="w")

        # ──── Card 1: Director Mamă & Proiect ────
        proj_card = self._make_card(container)
        proj_card.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        proj_card.grid_columnconfigure(0, weight=0)
        proj_card.grid_columnconfigure(1, weight=1)
        proj_card.grid_columnconfigure(2, weight=0)

        ctk.CTkLabel(
            proj_card, text="🏢 Director Mamă & Număr Proiect",
            font=("Segoe UI", 14, "bold"), text_color=_TEXT_PRI
        ).grid(row=0, column=0, columnspan=3, padx=16, pady=(8, 4), sticky="w")

        self._label(proj_card, text="Director Mamă:").grid(row=1, column=0, padx=(16, 6), pady=4, sticky="w")
        self._entry(proj_card, self.dwg_parent_dir).grid(row=1, column=1, padx=4, pady=4, sticky="ew")
        self._browse_btn(proj_card, self.dwg_browse_parent_dir).grid(row=1, column=2, padx=(4, 16), pady=4)

        self._label(proj_card, text="Număr Proiect:").grid(row=2, column=0, padx=(16, 6), pady=(4, 10), sticky="w")
        proj_inner = ctk.CTkFrame(proj_card, fg_color="transparent")
        proj_inner.grid(row=2, column=1, columnspan=2, padx=4, pady=(4, 10), sticky="w")
        self._entry(proj_inner, self.dwg_proj_num, placeholder_text="ex: 1242", width=140).pack(side="left", padx=(0, 10))
        self._label(proj_inner, text="(scanează automat folderele care conțin acest număr de proiect)", font=("Segoe UI", 11), text_color=_TEXT_SEC).pack(side="left")

        # ──── Card 2: Piese Sursă (PDF / NC / NC1) & Folder Salvare ────
        files_card = self._make_card(container)
        files_card.grid(row=2, column=0, sticky="ew", pady=(0, 8))
        files_card.grid_columnconfigure(0, weight=0)
        files_card.grid_columnconfigure(1, weight=1)
        files_card.grid_columnconfigure(2, weight=0)

        ctk.CTkLabel(
            files_card, text="📁 Piese Sursă (PDF / NC) & Salvare Desene",
            font=("Segoe UI", 14, "bold"), text_color=_TEXT_PRI
        ).grid(row=0, column=0, columnspan=3, padx=16, pady=(8, 4), sticky="w")

        # Row 1: Source & Target Selectors
        opt_frame = ctk.CTkFrame(files_card, fg_color="transparent")
        opt_frame.grid(row=1, column=0, columnspan=3, padx=16, pady=(2, 6), sticky="ew")

        self._label(opt_frame, text="Format Piese Sursă:").pack(side="left", padx=(0, 6))
        self.dwg_source_seg = ctk.CTkSegmentedButton(
            opt_frame,
            values=["PDF (.pdf)", "NC / NC1 (.nc, .nc1)", "Auto (Toate)"],
            variable=self.dwg_source_type,
            command=self._on_dwg_source_type_change,
            height=28,
            font=("Segoe UI", 11, "bold")
        )
        self.dwg_source_seg.pack(side="left", padx=(0, 18))

        self._label(opt_frame, text="Caută în Proiect:").pack(side="left", padx=(0, 6))
        self.dwg_target_seg = ctk.CTkSegmentedButton(
            opt_frame,
            values=["DWG (.dwg)", "PDF (.pdf)", "DWG & PDF"],
            variable=self.dwg_target_type,
            height=28,
            font=("Segoe UI", 11, "bold")
        )
        self.dwg_target_seg.pack(side="left")

        # Row 2: Folder Piese
        src_lbl_text = "Folder Piese (NC/NC1):" if "NC" in self.dwg_source_type.get() else ("Folder Piese (PDF / NC):" if "Auto" in self.dwg_source_type.get() else "Folder Piese (PDF):")
        self.dwg_piese_lbl = self._label(files_card, text=src_lbl_text)
        self.dwg_piese_lbl.grid(row=2, column=0, padx=(16, 6), pady=4, sticky="w")
        self._entry(files_card, self.dwg_pdf_dir).grid(row=2, column=1, padx=4, pady=4, sticky="ew")
        self._browse_btn(files_card, self.dwg_browse_pdf_dir).grid(row=2, column=2, padx=(4, 16), pady=4)

        # Row 3: Folder Destinație
        self._label(files_card, text="Folder Destinație:").grid(row=3, column=0, padx=(16, 6), pady=(4, 10), sticky="w")
        self._entry(files_card, self.dwg_dest_dir, placeholder_text="Opțional: implicit subdirector 'Desene_Gasite' în folderul sursă").grid(row=3, column=1, padx=4, pady=(4, 10), sticky="ew")
        self._browse_btn(files_card, self.dwg_browse_dest_dir).grid(row=3, column=2, padx=(4, 16), pady=(4, 10))

        # ──── Action Row ────
        action_box = ctk.CTkFrame(container, fg_color="transparent")
        action_box.grid(row=3, column=0, sticky="ew", pady=(2, 4))
        action_box.grid_columnconfigure(0, weight=1)
        action_box.grid_columnconfigure(1, weight=0)

        self.dwg_start_btn = ctk.CTkButton(
            action_box,
            text="▶ Caută și Copiază Desene Existente",
            command=self.dwg_start_search,
            fg_color=_GREEN,
            hover_color=_GREEN_HOV,
            text_color="#0b0f19",
            height=38,
            corner_radius=10,
            font=("Segoe UI", 15, "bold"),
        )
        self.dwg_start_btn.grid(row=0, column=0, sticky="ew", padx=(0, 8))

        self.dwg_cancel_btn = ctk.CTkButton(
            action_box,
            text="⏹ Oprește",
            command=self.dwg_cancel_search,
            fg_color=_ENTRY_BG,
            hover_color=_DANGER,
            text_color=_TEXT_PRI,
            width=100,
            height=38,
            corner_radius=10,
            font=("Segoe UI", 13, "bold"),
            state="disabled",
        )
        self.dwg_cancel_btn.grid(row=0, column=1, sticky="e")

        # Progress bar & Status label
        self.dwg_progress_bar = ctk.CTkProgressBar(
            container, progress_color=_ACCENT, fg_color=_ENTRY_BG, corner_radius=6, height=6
        )
        self.dwg_progress_bar.grid(row=4, column=0, sticky="ew", pady=(0, 4))
        self.dwg_progress_bar.set(0)

        # ──── Card 3: Results & Logs (Expands vertically) ────
        results_card = self._make_card(container)
        results_card.grid(row=5, column=0, sticky="nsew", pady=(0, 0))

        # Top Bar inside Results Card
        res_hdr = ctk.CTkFrame(results_card, fg_color="transparent")
        res_hdr.pack(fill="x", padx=10, pady=(8, 6))

        self.dwg_view_seg = ctk.CTkSegmentedButton(
            res_hdr,
            values=["🟢 Găsite (0)", "🔴 Lipsă (0)", "📜 Jurnal"],
            command=self.dwg_change_view,
            font=("Segoe UI", 12, "bold"),
        )
        self.dwg_view_seg.pack(side="left")
        self.dwg_view_seg.set("🟢 Găsite (0)")

        self.dwg_open_dest_btn = ctk.CTkButton(
            res_hdr,
            text="📂 Deschide Destinația",
            width=140,
            height=30,
            fg_color=_ENTRY_BG,
            hover_color=_ACCENT,
            text_color=_TEXT_PRI,
            font=("Segoe UI", 11, "bold"),
            command=self.dwg_open_destination,
            state="disabled",
        )
        self.dwg_open_dest_btn.pack(side="right", padx=(6, 0))

        self.dwg_export_excel_btn = ctk.CTkButton(
            res_hdr,
            text="📊 Export Excel",
            width=120,
            height=30,
            fg_color=_ENTRY_BG,
            hover_color=_GREEN,
            text_color=_TEXT_PRI,
            font=("Segoe UI", 11, "bold"),
            command=self.dwg_export_excel_report,
            state="disabled",
        )
        self.dwg_export_excel_btn.pack(side="right", padx=(6, 0))

        self.dwg_copy_missing_btn = ctk.CTkButton(
            res_hdr,
            text="📋 Copiază Piesele Lipsă",
            width=160,
            height=30,
            fg_color=_ENTRY_BG,
            hover_color=_ORANGE,
            text_color=_TEXT_PRI,
            font=("Segoe UI", 11, "bold"),
            command=self.dwg_copy_missing_pieces,
            state="disabled",
        )
        self.dwg_copy_missing_btn.pack(side="right")

        # Container for views
        self.dwg_views_container = ctk.CTkFrame(results_card, fg_color="transparent")
        self.dwg_views_container.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        self.dwg_views_container.grid_rowconfigure(0, weight=1)
        self.dwg_views_container.grid_columnconfigure(0, weight=1)

        # View 1: Found DWGs Table
        self.dwg_found_frame = ctk.CTkFrame(self.dwg_views_container, fg_color="transparent")
        self.dwg_found_frame.grid_rowconfigure(0, weight=1)
        self.dwg_found_frame.grid_columnconfigure(0, weight=1)

        found_cols = ("piece_code", "pdf_file", "dwg_file", "subfolder", "source_dir")
        self.dwg_found_tree = ttk.Treeview(
            self.dwg_found_frame,
            columns=found_cols,
            show="headings",
            style="DWG.Treeview",
            selectmode="browse",
        )
        self.dwg_found_tree.heading("piece_code", text="Cod Piesă")
        self.dwg_found_tree.heading("pdf_file", text="Fișier Piesă Sursă")
        self.dwg_found_tree.heading("dwg_file", text="Desen Găsit (Proiect)")
        self.dwg_found_tree.heading("subfolder", text="Subfolder / Grosime")
        self.dwg_found_tree.heading("source_dir", text="Folder Sursă în Proiect")

        self.dwg_found_tree.column("piece_code", width=130, anchor="w")
        self.dwg_found_tree.column("pdf_file", width=200, anchor="w")
        self.dwg_found_tree.column("dwg_file", width=180, anchor="w")
        self.dwg_found_tree.column("subfolder", width=130, anchor="center")
        self.dwg_found_tree.column("source_dir", width=300, anchor="w")

        found_vsb = ttk.Scrollbar(self.dwg_found_frame, orient="vertical", command=self.dwg_found_tree.yview)
        found_hsb = ttk.Scrollbar(self.dwg_found_frame, orient="horizontal", command=self.dwg_found_tree.xview)
        self.dwg_found_tree.configure(yscrollcommand=found_vsb.set, xscrollcommand=found_hsb.set)

        self.dwg_found_tree.grid(row=0, column=0, sticky="nsew")
        found_vsb.grid(row=0, column=1, sticky="ns")
        found_hsb.grid(row=1, column=0, sticky="ew")
        self.dwg_found_tree.bind("<Double-1>", self._on_found_tree_double_click)

        # View 2: Missing Parts Table
        self.dwg_missing_frame = ctk.CTkFrame(self.dwg_views_container, fg_color="transparent")
        self.dwg_missing_frame.grid_rowconfigure(0, weight=1)
        self.dwg_missing_frame.grid_columnconfigure(0, weight=1)

        missing_cols = ("piece_code", "pdf_file", "status")
        self.dwg_missing_tree = ttk.Treeview(
            self.dwg_missing_frame,
            columns=missing_cols,
            show="headings",
            style="DWG.Treeview",
            selectmode="browse",
        )
        self.dwg_missing_tree.heading("piece_code", text="Cod Piesă")
        self.dwg_missing_tree.heading("pdf_file", text="Fișier Piesă Sursă")
        self.dwg_missing_tree.heading("status", text="Stare")

        self.dwg_missing_tree.column("piece_code", width=180, anchor="w")
        self.dwg_missing_tree.column("pdf_file", width=300, anchor="w")
        self.dwg_missing_tree.column("status", width=350, anchor="w")

        missing_vsb = ttk.Scrollbar(self.dwg_missing_frame, orient="vertical", command=self.dwg_missing_tree.yview)
        missing_hsb = ttk.Scrollbar(self.dwg_missing_frame, orient="horizontal", command=self.dwg_missing_tree.xview)
        self.dwg_missing_tree.configure(yscrollcommand=missing_vsb.set, xscrollcommand=missing_hsb.set)

        self.dwg_missing_tree.grid(row=0, column=0, sticky="nsew")
        missing_vsb.grid(row=0, column=1, sticky="ns")
        missing_hsb.grid(row=1, column=0, sticky="ew")
        self.dwg_missing_tree.bind("<Double-1>", self._on_missing_tree_double_click)

        # Action bar under missing parts: generate DWG directly from PDF
        missing_action_bar = ctk.CTkFrame(self.dwg_missing_frame, fg_color="transparent")
        missing_action_bar.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(8, 2))

        self.btn_gen_missing_dwg = ctk.CTkButton(
            missing_action_bar,
            text="⚡ Generează DWG pentru Piesele Lipsă (din PDF)",
            font=("Segoe UI", 12, "bold"),
            fg_color=_ACCENT,
            hover_color=_ACCENT_HOV,
            height=32,
            corner_radius=8,
            command=self.generate_dwg_for_missing_parts
        )
        self.btn_gen_missing_dwg.pack(side="left", padx=2)

        # View 3: Execution Log Terminal
        self.dwg_log_frame = ctk.CTkFrame(self.dwg_views_container, fg_color="transparent")
        self.dwg_log_frame.grid_rowconfigure(0, weight=1)
        self.dwg_log_frame.grid_columnconfigure(0, weight=1)

        self.dwg_log_box = ctk.CTkTextbox(
            self.dwg_log_frame,
            state="disabled",
            font=("Cascadia Code", 12),
            fg_color=_LOG_BG,
            text_color="#a3e635",
            corner_radius=8,
            border_width=0,
        )
        self.dwg_log_box.grid(row=0, column=0, sticky="nsew")

        # Display initial view inside search container
        self.dwg_found_frame.grid(row=0, column=0, sticky="nsew")

        # Container 2: Organizare Piese (BOM -> Piese)
        self.dwg_org_container = ctk.CTkFrame(self.tab_dwg, fg_color="transparent")
        self._build_dwg_organize_view()

        # Container 3: Regrupare DWG pe Grosimi
        self.dwg_regroup_container = ctk.CTkFrame(self.tab_dwg, fg_color="transparent")
        self._build_dwg_regroup_view()

        # Set initial active subview
        self._on_dwg_subview_change(self.dwg_subview_mode.get())

    def _on_dwg_source_type_change(self, val=None):
        val = val or self.dwg_source_type.get()
        if hasattr(self, "dwg_piese_lbl"):
            if "NC" in val:
                self.dwg_piese_lbl.configure(text="Folder Piese (NC/NC1):")
            elif "Auto" in val:
                self.dwg_piese_lbl.configure(text="Folder Piese (PDF / NC):")
            else:
                self.dwg_piese_lbl.configure(text="Folder Piese (PDF):")
        self.save_settings()

    def dwg_browse_parent_dir(self):
        """Browse root parent directory for projects (uses global directory memory - Requirement #3)."""
        initial = self._get_global_initial_dir(self.dwg_parent_dir.get())
        d = filedialog.askdirectory(title="Selectează Directorul Mamă al Proiectelor", initialdir=initial)
        if d:
            norm = os.path.normpath(d)
            self.dwg_parent_dir.set(norm)
            self._set_global_last_dir(norm)

    def dwg_browse_pdf_dir(self):
        """Browse folder with part PDF files to be drawn (uses global directory memory - Requirement #3)."""
        initial = self._get_global_initial_dir(self.dwg_pdf_dir.get())
        d = filedialog.askdirectory(title="Selectează Folderul cu Piese de Desenat (PDF)", initialdir=initial)
        if d:
            norm = os.path.normpath(d)
            self.dwg_pdf_dir.set(norm)
            self._set_global_last_dir(norm)

    def dwg_browse_dest_dir(self):
        """Browse destination folder for found DWGs (uses global directory memory - Requirement #3)."""
        initial = self._get_global_initial_dir(self.dwg_dest_dir.get())
        d = filedialog.askdirectory(title="Selectează Folderul Destinație pentru DWG-uri Găsite", initialdir=initial)
        if d:
            norm = os.path.normpath(d)
            self.dwg_dest_dir.set(norm)
            self._set_global_last_dir(norm)

    def dwg_change_view(self, val=None):
        """Switch between Found, Missing, and Log views."""
        val = val or self.dwg_view_seg.get()
        self.dwg_found_frame.grid_forget()
        self.dwg_missing_frame.grid_forget()
        self.dwg_log_frame.grid_forget()

        if "Găsite" in val or "Gasite" in val:
            self.dwg_found_frame.grid(row=0, column=0, sticky="nsew")
        elif "Lipsă" in val or "Lipsa" in val:
            self.dwg_missing_frame.grid(row=0, column=0, sticky="nsew")
        else:
            self.dwg_log_frame.grid(row=0, column=0, sticky="nsew")

    def dwg_log(self, msg):
        """Thread-safe logging method for DWG Finder."""
        self.after(0, self._dwg_insert_log, msg)

    def _dwg_insert_log(self, msg):
        self.dwg_log_box.configure(state="normal")
        self.dwg_log_box.insert("end", str(msg) + "\n")
        self.dwg_log_box.see("end")
        self.dwg_log_box.configure(state="disabled")

    def _on_found_tree_double_click(self, event):
        """Open or reveal found DWG file on double click."""
        sel = self.dwg_found_tree.selection()
        if not sel:
            return
        vals = self.dwg_found_tree.item(sel[0], "values")
        if vals and len(vals) >= 3:
            dwg_file = vals[2]
            subfolder = vals[3] if len(vals) >= 4 and vals[3] != "-" else ""
            dest_dir = self.dwg_dest_dir.get().strip() or os.path.join(self.dwg_pdf_dir.get().strip(), "DWG_Gasite")
            target_path = os.path.join(dest_dir, subfolder, dwg_file) if subfolder else os.path.join(dest_dir, dwg_file)
            if not os.path.exists(target_path):
                target_path = os.path.join(dest_dir, dwg_file)
            if not os.path.exists(target_path):
                for item in self.dwg_results.get("found", []):
                    if item.get("dwg_file") == dwg_file:
                        target_path = item.get("dest_path") or item.get("source_path")
                        break
            if os.path.exists(target_path):
                try:
                    os.startfile(target_path)
                except Exception as e:
                    messagebox.showerror("Eroare", f"Nu s-a putut deschide fișierul:\n{e}")

    def _on_missing_tree_double_click(self, event):
        """Open missing part PDF file on double click."""
        sel = self.dwg_missing_tree.selection()
        if not sel:
            return
        vals = self.dwg_missing_tree.item(sel[0], "values")
        if vals and len(vals) >= 2:
            pdf_file = vals[1]
            pdf_dir = self.dwg_pdf_dir.get().strip()
            target_path = os.path.join(pdf_dir, pdf_file)
            if os.path.exists(target_path):
                try:
                    os.startfile(target_path)
                except Exception as e:
                    messagebox.showerror("Eroare", f"Nu s-a putut deschide fișierul:\n{e}")

    def dwg_start_search(self):
        """Validate inputs and launch background worker thread for DWG extraction."""
        if getattr(self, "_dwg_searching", False):
            return

        parent_dir = self.dwg_parent_dir.get().strip()
        proj_num = self.dwg_proj_num.get().strip()
        pdf_dir = self.dwg_pdf_dir.get().strip()
        dest_dir = self.dwg_dest_dir.get().strip()

        if not parent_dir or not os.path.exists(parent_dir):
            messagebox.showwarning("Atenție", "Te rog selectează un Director Mamă valid pentru proiecte.")
            return

        if not proj_num:
            messagebox.showwarning("Atenție", "Te rog introdu Numărul Proiectului (ex: 1242).")
            return

        if not pdf_dir or not os.path.exists(pdf_dir):
            messagebox.showwarning("Atenție", "Te rog selectează Folderul cu piese de desenat (PDF).")
            return

        if not dest_dir:
            dest_dir = os.path.join(pdf_dir, "Desene_Gasite")
            self.dwg_dest_dir.set(dest_dir)

        make_dir_with_permissions(dest_dir)

        # Clear UI state
        self._dwg_searching = True
        self._dwg_cancel_requested = False
        self.dwg_start_btn.configure(state="disabled")
        self.dwg_cancel_btn.configure(state="normal")
        self.dwg_open_dest_btn.configure(state="disabled")
        self.dwg_export_excel_btn.configure(state="disabled")
        self.dwg_copy_missing_btn.configure(state="disabled")
        self.dwg_progress_bar.set(0)

        # Clear previous items
        for item in self.dwg_found_tree.get_children():
            self.dwg_found_tree.delete(item)
        for item in self.dwg_missing_tree.get_children():
            self.dwg_missing_tree.delete(item)

        self.dwg_log_box.configure(state="normal")
        self.dwg_log_box.delete("1.0", "end")
        self.dwg_log_box.configure(state="disabled")

        # Switch view to Log during execution
        self.dwg_view_seg.set("📜 Jurnal")
        self.dwg_change_view("📜 Jurnal")

        self.save_settings()

        threading.Thread(
            target=self._dwg_search_worker,
            args=(parent_dir, proj_num, pdf_dir, dest_dir),
            daemon=True
        ).start()

    def dwg_cancel_search(self):
        """Request cooperative cancellation of DWG Finder worker thread."""
        if getattr(self, "_dwg_searching", False):
            self._dwg_cancel_requested = True
            self.dwg_cancel_btn.configure(state="disabled")
            self.dwg_log("\n⚠ Solicitare de oprire primită. Se finalizează operațiunea curentă...")

    def _dwg_search_worker(self, parent_dir, proj_num, pdf_dir, dest_dir):
        """Worker thread executing CAD matching & extraction with configurable source and target formats."""
        target_choice = getattr(self, "dwg_target_type", None)
        target_val = target_choice.get() if target_choice else "DWG (.dwg)"
        target_exts = set()
        if "DWG" in target_val:
            target_exts.add(".dwg")
        if "PDF" in target_val:
            target_exts.add(".pdf")
        if not target_exts:
            target_exts = {".dwg"}

        source_choice = getattr(self, "dwg_source_type", None)
        source_val = source_choice.get() if source_choice else "PDF (.pdf)"
        source_exts = set()
        if "PDF" in source_val:
            source_exts.add(".pdf")
        if "NC" in source_val:
            source_exts.add(".nc")
            source_exts.add(".nc1")
        if "Auto" in source_val or not source_exts:
            source_exts = {".pdf", ".nc", ".nc1"}

        self.dwg_log("=" * 60)
        self.dwg_log(f"🚀 Pornire Căutare Desene Existente (SteelSync v{APP_VERSION})")
        self.dwg_log(f"• Director Mamă: {parent_dir}")
        self.dwg_log(f"• Număr Proiect Căutat: '{proj_num}'")
        self.dwg_log(f"• Folder Piese Sursă: {pdf_dir} ({', '.join(sorted(source_exts))})")
        self.dwg_log(f"• Tip Desene Căutate în Proiect: {', '.join(sorted(target_exts))}")
        self.dwg_log(f"• Folder Salvare Desene: {dest_dir}")
        self.dwg_log("=" * 60 + "\n")

        make_dir_with_permissions(dest_dir)

        # ─── Pasul 1: Localizarea Folderelor de Proiect în Directorul Mamă ───
        self.dwg_log("🔍 Pasul 1: Scanare directoare proiect în Directorul Mamă (recursiv în toate subfolderele)...")
        proj_pattern = re.compile(rf'(?<!\d){re.escape(proj_num)}(?!\d)', re.IGNORECASE)
        matched_proj_dirs = []

        if proj_pattern.search(os.path.basename(parent_dir)):
            matched_proj_dirs.append(parent_dir)

        for root, dirs, files in os.walk(parent_dir, followlinks=True):
            if self._dwg_cancel_requested:
                break
            matched_here = []
            for d in dirs:
                if proj_pattern.search(d):
                    full_d = os.path.join(root, d)
                    if full_d not in matched_proj_dirs:
                        matched_proj_dirs.append(full_d)
                    matched_here.append(d)
            for d in matched_here:
                dirs.remove(d)

        if not matched_proj_dirs and os.path.isdir(parent_dir):
            self.dwg_log(f"ℹ Nu s-a găsit niciun subfolder cu numărul '{proj_num}'. Se utilizează directorul mamă ca proiect direct.")
            matched_proj_dirs.append(parent_dir)
        elif matched_proj_dirs:
            self.dwg_log(f"✅ S-au identificat {len(matched_proj_dirs)} directoare de proiect asociate:")
            for p in matched_proj_dirs:
                self.dwg_log(f"   ↳ {p}")

        # ─── Indexare Desene existente în directoarele găsite (până la capăt) ───
        self.dwg_log(f"\n📂 Indexare recursivă a desenelor ({', '.join(sorted(target_exts))}) din toate subfolderele proiectului...")
        dwg_records = []  # list of dicts: path, filename, stem, ext, clean_key, mtime, subfolder, proj_dir
        generic_folder_names = {'dwg', 'dwgs', 'cad', 'drawings', 'desene', 'model', 'tekla', 'advance steel', 'export', 'fisiere'}
        thickness_regex = re.compile(
            r'^(?:t|pl|p|thk|grosime)?[_\s-]*\d+(?:[\.,]\d+)?[_\s-]*(?:mm)?(?:[_\s-]+[a-z0-9]+)?$|^\d+(?:[\.,]\d+)?$',
            re.IGNORECASE
        )

        def detect_thickness_subfolder(dwg_full_path, project_root_path):
            try:
                rel = os.path.relpath(os.path.dirname(dwg_full_path), project_root_path)
            except Exception:
                rel = os.path.dirname(dwg_full_path)
            parts = [p.strip() for p in rel.split(os.sep) if p.strip() and p.strip() != '.']
            if not parts:
                return ""
            for p in reversed(parts):
                if p.lower() in generic_folder_names:
                    continue
                if thickness_regex.match(p):
                    return p
            if parts[-1].lower() not in generic_folder_names:
                return parts[-1]
            for p in reversed(parts):
                if p.lower() not in generic_folder_names:
                    return p
            return ""

        scanned_dirs_count = 0
        for p_dir in matched_proj_dirs:
            if self._dwg_cancel_requested:
                break
            for root, dirs, files in os.walk(p_dir, followlinks=True):
                if self._dwg_cancel_requested:
                    break
                scanned_dirs_count += 1
                for f in files:
                    f_ext = os.path.splitext(f)[1].lower()
                    if f_ext in target_exts:
                        full_p = os.path.join(root, f)
                        stem = Path(f).stem
                        clean_key = re.sub(r'[^A-Za-z0-9]', '', stem).upper()
                        if not clean_key:
                            continue
                        try:
                            mtime = os.path.getmtime(full_p)
                        except Exception:
                            mtime = 0

                        subfolder = detect_thickness_subfolder(full_p, p_dir)

                        dwg_records.append({
                            "path": full_p,
                            "filename": f,
                            "stem": stem,
                            "ext": f_ext,
                            "clean_key": clean_key,
                            "mtime": mtime,
                            "subfolder": subfolder,
                            "proj_dir": p_dir,
                        })

        self.dwg_log(f"✅ Indexare completă: {len(dwg_records)} desene găsite în {scanned_dirs_count} foldere scanate până la capăt.")

        # ─── Helper de potrivire variații ───
        def is_dwg_variation(dwg_k, base_k):
            """Checks if dwg_k is exact match or valid variation of base_k."""
            if dwg_k == base_k:
                return True, "Exact"
            if dwg_k.startswith(base_k):
                suffix = dwg_k[len(base_k):]
                if len(suffix) > 0 and not suffix[0].isdigit():
                    if re.match(r'^[A-Z0-9]{1,5}$', suffix):
                        return True, f"Variație ({suffix})"
            return False, ""

        # ─── Pasul 2 & 3: Scanare Piese Sursă și Potrivire cu Desenele ───
        self.dwg_log(f"\n📄 Pasul 2: Scanare piese ({', '.join(sorted(source_exts))}) din folderul de piese...")
        src_entries = []  # list of (filename, full_path, ext)
        try:
            for root, dirs, files in os.walk(pdf_dir, followlinks=True):
                if self._dwg_cancel_requested:
                    break
                for f in files:
                    f_ext = os.path.splitext(f)[1].lower()
                    if f_ext in source_exts:
                        src_entries.append((f, os.path.join(root, f), f_ext))
        except Exception as e:
            self.dwg_log(f"❌ Eroare la citirea folderului de piese: {e}")
            self.after(0, self._dwg_on_search_complete, [], [], dest_dir, False)
            return

        total_pieces = len(src_entries)
        self.dwg_log(f"S-au găsit {total_pieces} fișiere piese sursă de analizat.")

        found_list = []
        missing_list = []
        copied_dest_paths = set()

        for idx, (src_name, src_path, src_ext) in enumerate(src_entries):
            if self._dwg_cancel_requested:
                self.dwg_log("\n🛑 Procesul de căutare a fost întrerupt de utilizator.")
                break

            progress_val = (idx + 1) / max(1, total_pieces)
            self.after(0, self.dwg_progress_bar.set, progress_val)

            src_stem = Path(src_name).stem

            # ── Extragere coduri piesă din numele fișierului sursă ──
            chunks = re.split(r'[\s_()\[\],;]+', src_stem)
            clean_chunks = [re.sub(r'[^A-Za-z0-9]', '', c).upper() for c in chunks if re.sub(r'[^A-Za-z0-9]', '', c)]
            seqs = re.findall(r'[A-Za-z]+[-_]?\d+', src_stem)
            clean_seqs = [re.sub(r'[^A-Za-z0-9]', '', s).upper() for s in seqs if re.sub(r'[^A-Za-z0-9]', '', s)]
            clean_stem = re.sub(r'[^A-Za-z0-9]', '', src_stem).upper()

            candidate_keys = []
            for s in clean_seqs:
                if s not in candidate_keys: candidate_keys.append(s)
            for c in clean_chunks:
                if c not in candidate_keys: candidate_keys.append(c)
            if clean_stem not in candidate_keys:
                candidate_keys.append(clean_stem)

            # Dacă este fișier NC / NC1, extrage codul piesei din antetul DSTV ca fallback suplimentar
            if src_ext in ('.nc', '.nc1'):
                nc_hdr_mark = extract_nc_part_code(src_path)
                if nc_hdr_mark and nc_hdr_mark not in candidate_keys:
                    candidate_keys.append(nc_hdr_mark)

            matching_for_piece = []
            used_records_for_piece = set()

            # Verificare candidați
            for cand in candidate_keys:
                if len(cand) < 2:
                    continue
                for rec in dwg_records:
                    if rec["path"] in used_records_for_piece:
                        continue
                    is_match, reason = is_dwg_variation(rec["clean_key"], cand)
                    if is_match:
                        matching_for_piece.append((rec, f"Nume Fișier ({src_ext}): {reason}"))
                        used_records_for_piece.add(rec["path"])

            # Fallback Antet PDF dacă fișierul sursă este PDF și nu s-a potrivit nimic din nume
            if not matching_for_piece and src_ext == '.pdf':
                all_dwg_keys = sorted(list({r["clean_key"] for r in dwg_records}), key=len, reverse=True)
                header_code = extract_pdf_header_part_code(src_path, known_keys=all_dwg_keys)
                if header_code:
                    for rec in dwg_records:
                        if rec["path"] in used_records_for_piece:
                            continue
                        is_match, reason = is_dwg_variation(rec["clean_key"], header_code)
                        if is_match:
                            matching_for_piece.append((rec, f"Antet PDF: {reason}"))
                            used_records_for_piece.add(rec["path"])

            # Copiere în destinație (cu păstrarea folderului de grosime)
            if matching_for_piece:
                for rec, method_desc in matching_for_piece:
                    sub = rec["subfolder"]
                    if sub:
                        target_sub_dir = os.path.join(dest_dir, sub)
                        make_dir_with_permissions(target_sub_dir)
                        dst_file = os.path.join(target_sub_dir, rec["filename"])
                    else:
                        dst_file = os.path.join(dest_dir, rec["filename"])

                    try:
                        safe_copy(rec["path"], dst_file)
                        copied_dest_paths.add(dst_file)
                    except Exception as e:
                        self.dwg_log(f"  ❌ Eroare copiere {rec['filename']}: {e}")

                    found_list.append({
                        "piece_code": rec["clean_key"],
                        "pdf_file": src_name,
                        "dwg_file": rec["filename"],
                        "subfolder": sub or "-",
                        "source_dir": os.path.dirname(rec["path"]),
                        "source_path": rec["path"],
                        "dest_path": dst_file,
                        "method": method_desc,
                    })
                    sub_info = f" [Folder: {sub}]" if sub else ""
                    self.dwg_log(f"✔ [GĂSIT] {src_name}  ➔  {rec['filename']}{sub_info}  ({method_desc})")
            else:
                candidate = clean_seqs[0] if clean_seqs else (clean_chunks[0] if clean_chunks else clean_stem)
                missing_list.append({
                    "piece_code": candidate or src_stem,
                    "pdf_file": src_name,
                    "status": "Desen Lipsă în Proiect",
                })
                self.dwg_log(f"✖ [LIPSĂ] {src_name}  (Nu există desen în proiect)")

        # Permissions check on destination
        if found_list:
            ensure_full_permissions_tree(dest_dir)

        self.dwg_log("\n" + "=" * 60)
        self.dwg_log("🎉 Operațiune finalizată!")
        self.dwg_log(f"• Total piese analizate: {len(found_list) + len(missing_list)}")
        self.dwg_log(f"• 🟢 Desene Găsite și Copiate: {len(found_list)}")
        self.dwg_log(f"• 🔴 Piese Fără Desen: {len(missing_list)}")
        self.dwg_log(f"• Folder Salvare: {dest_dir}")
        self.dwg_log("=" * 60 + "\n")

        self.after(0, self._dwg_on_search_complete, found_list, missing_list, dest_dir, self._dwg_cancel_requested)

    def _dwg_on_search_complete(self, found_list, missing_list, dest_dir, was_cancelled):
        """Update GUI after DWG search finishes."""
        self._dwg_searching = False
        self._dwg_cancel_requested = False
        self.dwg_start_btn.configure(state="normal")
        self.dwg_cancel_btn.configure(state="disabled")

        self.dwg_results = {"found": found_list, "missing": missing_list}

        # Update button labels with counts
        seg_values = [f"🟢 Găsite ({len(found_list)})", f"🔴 Lipsă ({len(missing_list)})", "📜 Jurnal"]
        self.dwg_view_seg.configure(values=seg_values)

        # Populate Found Treeview
        for item in found_list:
            self.dwg_found_tree.insert(
                "", "end",
                values=(item["piece_code"], item["pdf_file"], item["dwg_file"], item.get("subfolder", "-"), item["source_dir"])
            )

        # Populate Missing Treeview
        for item in missing_list:
            self.dwg_missing_tree.insert(
                "", "end",
                values=(item["piece_code"], item["pdf_file"], item["status"])
            )

        # Enable action buttons if results exist
        if found_list or os.path.exists(dest_dir):
            self.dwg_open_dest_btn.configure(state="normal")
        if found_list or missing_list:
            self.dwg_export_excel_btn.configure(state="normal")
        if missing_list:
            self.dwg_copy_missing_btn.configure(state="normal")

        # Automatically switch to Found if found, else Missing
        if found_list:
            self.dwg_view_seg.set(seg_values[0])
            self.dwg_change_view(seg_values[0])
        elif missing_list:
            self.dwg_view_seg.set(seg_values[1])
            self.dwg_change_view(seg_values[1])

        if was_cancelled:
            messagebox.showinfo("Căutare Oprită", f"Operațiunea a fost întreruptă.\nS-au extras {len(found_list)} fișiere DWG.")
        else:
            messagebox.showinfo(
                "Căutare Finalizată",
                f"Proces finalizat cu succes!\n\n"
                f"🟢 Piese cu DWG Găsit: {len(found_list)}\n"
                f"🔴 Piese Lipsă (Necesită desenare): {len(missing_list)}\n\n"
                f"Fișierele DWG găsite au fost copiate în:\n{dest_dir}"
            )

    def dwg_open_destination(self):
        """Open the DWG destination folder in Windows Explorer."""
        dest_dir = self.dwg_dest_dir.get().strip() or os.path.join(self.dwg_pdf_dir.get().strip(), "DWG_Gasite")
        if dest_dir and os.path.exists(dest_dir):
            try:
                os.startfile(dest_dir)
            except Exception as e:
                messagebox.showerror("Eroare", f"Nu s-a putut deschide folderul:\n{e}")
        else:
            messagebox.showwarning("Atenție", "Folderul destinație nu există încă.")

    def dwg_copy_missing_pieces(self):
        """Copy the list of missing parts to the Windows clipboard."""
        missing = self.dwg_results.get("missing", [])
        if not missing:
            messagebox.showinfo("Informație", "Nu există piese lipsă de copiat.")
            return

        lines = [item["piece_code"] for item in missing if item.get("piece_code")]
        clipboard_txt = "\n".join(lines)

        try:
            self.clipboard_clear()
            self.clipboard_append(clipboard_txt)
            messagebox.showinfo(
                "Copiat în Clipboard",
                f"Au fost copiate {len(lines)} coduri de piese lipsă în clipboard.\nLe poți insera direct în Excel sau trimite spre desenare."
            )
        except Exception as e:
            messagebox.showerror("Eroare", f"Nu s-a putut copia în clipboard:\n{e}")

    def dwg_export_excel_report(self):
        """Generate and export a styled Excel report for found and missing DWGs."""
        found = self.dwg_results.get("found", [])
        missing = self.dwg_results.get("missing", [])

        if not found and not missing:
            messagebox.showinfo("Informație", "Nu există date pentru export. Rulați mai întâi căutarea.")
            return

        proj_num = self.dwg_proj_num.get().strip() or "Proiect"
        dest_dir = self.dwg_dest_dir.get().strip() or self.dwg_pdf_dir.get().strip()
        initial_dir = self._get_global_initial_dir(dest_dir)
        default_file = f"Raport_DWG_Proiect_{proj_num}.xlsx"

        save_path = filedialog.asksaveasfilename(
            title="Salvare Raport Excel Căutare DWG",
            initialdir=initial_dir,
            initialfile=default_file,
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not save_path:
            return

        try:
            wb = Workbook()

            # Styling
            thin_side = Side(style='thin', color='BFBFBF')
            thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
            header_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
            data_font = Font(name="Segoe UI", size=10)
            center_align = Alignment(horizontal='center', vertical='center')
            left_align = Alignment(horizontal='left', vertical='center')

            green_header = PatternFill(start_color="107C41", end_color="107C41", fill_type="solid")
            red_header = PatternFill(start_color="C53929", end_color="C53929", fill_type="solid")
            alt_fill = PatternFill(start_color="F2F5F8", end_color="F2F5F8", fill_type="solid")

            # ── Sheet 1: Piese cu DWG Gasit ──
            ws_found = wb.active
            ws_found.title = "Piese cu DWG Gasit"
            headers_found = ["Nr. Crt.", "Cod Piesă", "Fișier Piesă Sursă", "Desen Găsit (Proiect)", "Subfolder / Grosime", "Folder Sursă Desen", "Cale Destinație"]
            ws_found.append(headers_found)

            for col_idx in range(1, len(headers_found) + 1):
                cell = ws_found.cell(row=1, column=col_idx)
                cell.font = header_font
                cell.fill = green_header
                cell.alignment = center_align
                cell.border = thin_border

            for row_idx, item in enumerate(found, start=2):
                row_data = [
                    row_idx - 1,
                    item["piece_code"],
                    item["pdf_file"],
                    item["dwg_file"],
                    item.get("subfolder", "-"),
                    item["source_dir"],
                    item["dest_path"]
                ]
                ws_found.append(row_data)
                fill_to_apply = alt_fill if row_idx % 2 == 0 else PatternFill()
                for col_idx in range(1, len(row_data) + 1):
                    c = ws_found.cell(row=row_idx, column=col_idx)
                    c.font = data_font
                    c.border = thin_border
                    if fill_to_apply.fill_type:
                        c.fill = fill_to_apply
                    c.alignment = center_align if col_idx in [1, 2] else left_align

            for col in ws_found.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = get_column_letter(col[0].column)
                ws_found.column_dimensions[col_letter].width = max(max_len + 3, 12)

            # ── Sheet 2: Piese Fara DWG (Lipsa) ──
            ws_missing = wb.create_sheet(title="Piese Fara DWG (Lipsa)")
            headers_missing = ["Nr. Crt.", "Cod Piesă", "Fișier PDF (Desenat)", "Stare / Necesitate"]
            ws_missing.append(headers_missing)

            for col_idx in range(1, len(headers_missing) + 1):
                cell = ws_missing.cell(row=1, column=col_idx)
                cell.font = header_font
                cell.fill = red_header
                cell.alignment = center_align
                cell.border = thin_border

            for row_idx, item in enumerate(missing, start=2):
                row_data = [
                    row_idx - 1,
                    item["piece_code"],
                    item["pdf_file"],
                    item["status"]
                ]
                ws_missing.append(row_data)
                fill_to_apply = alt_fill if row_idx % 2 == 0 else PatternFill()
                for col_idx in range(1, len(row_data) + 1):
                    c = ws_missing.cell(row=row_idx, column=col_idx)
                    c.font = data_font
                    c.border = thin_border
                    if fill_to_apply.fill_type:
                        c.fill = fill_to_apply
                    c.alignment = center_align if col_idx in [1, 2] else left_align

            for col in ws_missing.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = get_column_letter(col[0].column)
                ws_missing.column_dimensions[col_letter].width = max(max_len + 3, 14)

            safe_save_workbook(wb, save_path)
            self._set_global_last_dir(save_path)
            messagebox.showinfo("Succes", f"Raportul Excel a fost salvat cu succes:\n{save_path}")
        except Exception as e:
            messagebox.showerror("Eroare", f"Nu s-a putut salva raportul Excel:\n{e}")

    # ══════════════════════════════════════════════════════════════════════
    #  PDF TO CAD (DWG / DXF) GENERATOR
    # ══════════════════════════════════════════════════════════════════════
    def generate_dwg_for_missing_parts(self):
        """Generate DWG/DXF files directly for missing parts identified in DWG Finder."""
        items = self.dwg_missing_tree.get_children()
        if not items:
            messagebox.showinfo("Info", "Nu există piese lipsă în listă pentru a genera DWG-uri.")
            return

        missing_pdfs = []
        pdf_dir = self.dwg_pdf_dir.get().strip()

        for item_id in items:
            vals = self.dwg_missing_tree.item(item_id, "values")
            if vals and len(vals) >= 2:
                pdf_name = vals[1]
                if pdf_dir and os.path.isdir(pdf_dir):
                    candidate = os.path.join(pdf_dir, pdf_name)
                    if os.path.isfile(candidate):
                        missing_pdfs.append(candidate)

        if not missing_pdfs:
            messagebox.showwarning("Atenție", "Nu s-au putut localiza fișierele PDF sursă pe disc în folderul specificat.")
            return

        dest_dir = self.dwg_dest_dir.get().strip()
        if not dest_dir:
            dest_dir = os.path.join(pdf_dir, "DWG_Gasite") if pdf_dir else os.getcwd()

        os.makedirs(dest_dir, exist_ok=True)
        self.dwg_change_view("log")
        self.dwg_log(f"\n🚀 Generare automată CAD pentru {len(missing_pdfs)} piese lipsă din PDF...")

        def _worker():
            converter = PdfToCadConverter(oda_path=self.pdf_cad_oda_path.get().strip())
            fmt = "dwg" if "dwg" in self.dwg_target_type.get().lower() else "dxf"
            succ, fail, res = converter.convert_batch(
                pdf_paths=missing_pdfs,
                output_dir=dest_dir,
                output_format=fmt,
                only_front_view=True,
                calibrate_1to1=True,
                progress_callback=None,
                log_callback=lambda m: self.after(0, lambda msg=m: self.dwg_log(msg))
            )
            def _done():
                self.dwg_log(f"\n✅ Generare CAD finalizată: {succ} create cu succes, {fail} eșuate.")
                messagebox.showinfo(
                    "Generare CAD Finalizată",
                    f"Au fost generate {succ} fișiere CAD pentru piesele lipsă în:\n{dest_dir}"
                )
            self.after(0, _done)

        threading.Thread(target=_worker, daemon=True).start()

    def create_pdf_to_dwg_tab(self):
        """Construct the UI for PDF to DWG/DXF 1:1 Generator tab."""
        parent = self.tab_pdf_cad
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(3, weight=1)

        # ── Card 1: Header / Description ──
        hdr_card = ctk.CTkFrame(parent, fg_color=_CARD_BG, border_color=_CARD_BORDER, border_width=1, corner_radius=10)
        hdr_card.grid(row=0, column=0, sticky="ew", padx=16, pady=(12, 8))
        hdr_card.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(hdr_card, text="📐", font=("Segoe UI Emoji", 26)).grid(row=0, column=0, rowspan=2, padx=(16, 12), pady=10)
        ctk.CTkLabel(hdr_card, text="PDF to DWG Generator (Geometrie 1:1 Fără Cote)", font=("Segoe UI", 15, "bold"), text_color=_TEXT_PRI).grid(row=0, column=1, sticky="w", pady=(8, 0))
        ctk.CTkLabel(
            hdr_card,
            text="Generează fișiere DWG/DXF conținând DOAR conturul piesei (fără cote, fără texte, fără cartuș), calibrate la scară reală 1:1 (mm).",
            font=("Segoe UI", 11),
            text_color=_TEXT_SEC,
            wraplength=700,
            justify="left"
        ).grid(row=1, column=1, sticky="w", pady=(0, 8))

        # ── Card 2: Configuration Inputs ──
        cfg_card = ctk.CTkFrame(parent, fg_color=_CARD_BG, border_color=_CARD_BORDER, border_width=1, corner_radius=10)
        cfg_card.grid(row=1, column=0, sticky="ew", padx=16, pady=4)
        cfg_card.grid_columnconfigure(1, weight=1)

        # Row 0: Mode selection (Batch vs Single)
        ctk.CTkLabel(cfg_card, text="Mod de Lucru:", font=("Segoe UI", 12, "bold")).grid(row=0, column=0, sticky="w", padx=14, pady=8)
        self.pdf_cad_mode_seg = ctk.CTkSegmentedButton(
            cfg_card,
            values=["Folder Întreg (Batch)", "Fișier Unic (Single)"],
            variable=self.pdf_cad_mode,
            command=self._on_pdf_cad_mode_change,
            selected_color=_ACCENT,
            selected_hover_color=_ACCENT_HOV
        )
        self.pdf_cad_mode_seg.grid(row=0, column=1, sticky="w", padx=8, pady=8)

        # Row 1: Source File / Folder
        self.pdf_cad_src_lbl = ctk.CTkLabel(cfg_card, text="Folder PDF-uri:", font=("Segoe UI", 12))
        self.pdf_cad_src_lbl.grid(row=1, column=0, sticky="w", padx=14, pady=6)

        self.pdf_cad_src_entry = ctk.CTkEntry(cfg_card, textvariable=self.pdf_cad_batch_dir, height=32, corner_radius=6)
        self.pdf_cad_src_entry.grid(row=1, column=1, sticky="ew", padx=8, pady=6)

        self.pdf_cad_src_btn = ctk.CTkButton(
            cfg_card,
            text="📁 Browse",
            width=90,
            height=32,
            corner_radius=6,
            fg_color=_ACCENT,
            hover_color=_ACCENT_HOV,
            command=self.browse_pdf_cad_source
        )
        self.pdf_cad_src_btn.grid(row=1, column=2, padx=(0, 14), pady=6)

        # Row 2: Destination Folder
        ctk.CTkLabel(cfg_card, text="Folder Destinație:", font=("Segoe UI", 12)).grid(row=2, column=0, sticky="w", padx=14, pady=6)
        self.pdf_cad_dest_entry = ctk.CTkEntry(cfg_card, textvariable=self.pdf_cad_dest_dir, placeholder_text="Implicit: subfolder CAD_1to1 în folderul sursă", height=32, corner_radius=6)
        self.pdf_cad_dest_entry.grid(row=2, column=1, sticky="ew", padx=8, pady=6)

        ctk.CTkButton(
            cfg_card,
            text="📁 Browse",
            width=90,
            height=32,
            corner_radius=6,
            fg_color=_ACCENT,
            hover_color=_ACCENT_HOV,
            command=self.browse_pdf_cad_dest
        ).grid(row=2, column=2, padx=(0, 14), pady=6)

        # Row 3: Format & Options
        ctk.CTkLabel(cfg_card, text="Format Export:", font=("Segoe UI", 12)).grid(row=3, column=0, sticky="w", padx=14, pady=6)

        fmt_frame = ctk.CTkFrame(cfg_card, fg_color="transparent")
        fmt_frame.grid(row=3, column=1, columnspan=2, sticky="ew", padx=8, pady=6)

        self.pdf_cad_fmt_opt = ctk.CTkOptionMenu(
            fmt_frame,
            values=["DWG (.dwg cu fallback .dxf)", "DXF (.dxf R2010)", "DWG Direct (necesită ODA)"],
            variable=self.pdf_cad_format,
            width=230,
            height=30,
            corner_radius=6
        )
        self.pdf_cad_fmt_opt.pack(side="left", padx=(0, 12))

        ctk.CTkCheckBox(fmt_frame, text="Doar Vederea Principală (Front)", variable=self.pdf_cad_only_front).pack(side="left", padx=8)
        ctk.CTkCheckBox(fmt_frame, text="Calibrare Automată 1:1 din Stykkliste", variable=self.pdf_cad_calibrate).pack(side="left", padx=8)

        # Row 4: ODA Converter Configuration (Collapsible or optional)
        oda_frame = ctk.CTkFrame(cfg_card, fg_color="#182335", corner_radius=6)
        oda_frame.grid(row=4, column=0, columnspan=3, sticky="ew", padx=14, pady=(6, 12))
        oda_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(oda_frame, text="ODA File Converter (opțional pentru DWG nativ):", font=("Segoe UI", 11)).grid(row=0, column=0, sticky="w", padx=10, pady=6)
        ctk.CTkEntry(oda_frame, textvariable=self.pdf_cad_oda_path, placeholder_text="Cale ODAFileConverter.exe (detectat automat dacă e instalat)", height=28, corner_radius=4).grid(row=0, column=1, sticky="ew", padx=8, pady=6)
        ctk.CTkButton(oda_frame, text="📁 Browse", width=70, height=28, corner_radius=4, command=self.browse_pdf_cad_oda).grid(row=0, column=2, padx=4, pady=6)
        ctk.CTkButton(oda_frame, text="🔍 Test ODA", width=75, height=28, corner_radius=4, command=self.test_oda_installation).grid(row=0, column=3, padx=(4, 10), pady=6)

        # ── Card 3: Action & Progress ──
        act_card = ctk.CTkFrame(parent, fg_color="transparent")
        act_card.grid(row=2, column=0, sticky="ew", padx=16, pady=4)
        act_card.grid_columnconfigure(0, weight=1)

        self.btn_start_pdf_cad = ctk.CTkButton(
            act_card,
            text="🚀 Generează Fișiere CAD (DWG / DXF)",
            font=("Segoe UI", 13, "bold"),
            fg_color=_ACCENT,
            hover_color=_ACCENT_HOV,
            height=38,
            corner_radius=8,
            command=self.start_pdf_to_cad_process
        )
        self.btn_start_pdf_cad.pack(fill="x", pady=(2, 6))

        self.pdf_cad_progress = ctk.CTkProgressBar(act_card, height=10, corner_radius=5)
        self.pdf_cad_progress.pack(fill="x", pady=2)
        self.pdf_cad_progress.set(0)

        # ── Card 4: Terminal Log ──
        log_frame = ctk.CTkFrame(parent, fg_color=_CARD_BG, border_color=_CARD_BORDER, border_width=1, corner_radius=10)
        log_frame.grid(row=3, column=0, sticky="nsew", padx=16, pady=(4, 12))
        log_frame.grid_rowconfigure(0, weight=1)
        log_frame.grid_columnconfigure(0, weight=1)

        self.pdf_cad_log_box = ctk.CTkTextbox(
            log_frame,
            font=("Cascadia Code", 11),
            fg_color=_LOG_BG,
            text_color="#a3e635",
            corner_radius=8
        )
        self.pdf_cad_log_box.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)

        # Update initial mode visibility
        self._on_pdf_cad_mode_change()

    def _on_pdf_cad_mode_change(self, val=None):
        mode = self.pdf_cad_mode.get()
        is_single = "single" in mode.lower()
        if is_single:
            self.pdf_cad_src_lbl.configure(text="Fișier PDF:")
            self.pdf_cad_src_entry.configure(textvariable=self.pdf_cad_single_file)
        else:
            self.pdf_cad_src_lbl.configure(text="Folder PDF-uri:")
            self.pdf_cad_src_entry.configure(textvariable=self.pdf_cad_batch_dir)

    def browse_pdf_cad_source(self):
        mode = self.pdf_cad_mode.get()
        is_single = "single" in mode.lower()
        if is_single:
            f = filedialog.askopenfilename(
                title="Selectează Fișier Desen PDF",
                filetypes=[("Fișiere PDF", "*.pdf"), ("Toate Fișierele", "*.*")],
                initialdir=getattr(self, "global_last_dir", None)
            )
            if f:
                self.pdf_cad_single_file.set(f)
                self._set_global_last_dir(f)
        else:
            d = filedialog.askdirectory(
                title="Selectează Folderul cu Desene PDF",
                initialdir=getattr(self, "global_last_dir", None)
            )
            if d:
                self.pdf_cad_batch_dir.set(d)
                self._set_global_last_dir(d)

    def browse_pdf_cad_dest(self):
        d = filedialog.askdirectory(
            title="Selectează Folderul Destinație pentru CAD",
            initialdir=getattr(self, "global_last_dir", None)
        )
        if d:
            self.pdf_cad_dest_dir.set(d)
            self._set_global_last_dir(d)

    def browse_pdf_cad_oda(self):
        f = filedialog.askopenfilename(
            title="Selectează Executabilul ODAFileConverter.exe",
            filetypes=[("Executabil ODA", "ODAFileConverter.exe"), ("Aplicații", "*.exe")]
        )
        if f:
            self.pdf_cad_oda_path.set(f)

    def test_oda_installation(self):
        oda_path = self.pdf_cad_oda_path.get().strip() or None
        converter = PdfToCadConverter(oda_path=oda_path)
        if converter.has_oda():
            messagebox.showinfo("Status ODA", "✅ ODA File Converter este configurat și funcțional! Puteți genera direct fișiere .DWG native.")
        else:
            messagebox.showwarning(
                "Status ODA",
                "⚠ ODA File Converter nu a fost detectat.\n\nAplicația va genera fișiere .DXF 1:1 native (compatibile cu orice versiune de AutoCAD/CAD).\n\nDacă doriți salvarea directă ca .DWG nativ, instalați 'ODA File Converter'."
            )

    def pdf_cad_log(self, msg):
        if hasattr(self, "pdf_cad_log_box"):
            self.pdf_cad_log_box.insert("end", msg + "\n")
            self.pdf_cad_log_box.see("end")

    def start_pdf_to_cad_process(self):
        """Handler for generating CAD files asynchronously."""
        if not PdfToCadConverter.is_available():
            messagebox.showerror("Eroare", "Librăriile 'pymupdf' și 'ezdxf' nu sunt disponibile.")
            return

        mode = self.pdf_cad_mode.get()
        is_single = "single" in mode.lower()

        if is_single:
            single_f = self.pdf_cad_single_file.get().strip()
            if not single_f or not os.path.isfile(single_f):
                messagebox.showerror("Eroare", "Te rog selectează un fișier PDF valid.")
                return
            target_files = [single_f]
            dest_dir = self.pdf_cad_dest_dir.get().strip() or os.path.join(os.path.dirname(single_f), "CAD_1to1")
        else:
            batch_d = self.pdf_cad_batch_dir.get().strip()
            if not batch_d or not os.path.isdir(batch_d):
                messagebox.showerror("Eroare", "Te rog selectează un folder valid cu fișiere PDF.")
                return
            target_files = [
                os.path.join(batch_d, f)
                for f in sorted(os.listdir(batch_d))
                if f.lower().endswith(".pdf") and os.path.isfile(os.path.join(batch_d, f))
            ]
            if not target_files:
                messagebox.showwarning("Atenție", "Nu s-au găsit fișiere .pdf în folderul selectat.")
                return
            dest_dir = self.pdf_cad_dest_dir.get().strip() or os.path.join(batch_d, "CAD_1to1")

        os.makedirs(dest_dir, exist_ok=True)
        raw_fmt = self.pdf_cad_format.get().lower()
        export_fmt = "dxf" if "dxf" in raw_fmt else "dwg"
        only_front = self.pdf_cad_only_front.get()
        calibrate = self.pdf_cad_calibrate.get()
        oda_path = self.pdf_cad_oda_path.get().strip()

        self.btn_start_pdf_cad.configure(state="disabled", text="⏳ Generare în curs...")
        self.pdf_cad_progress.set(0)
        self.pdf_cad_log_box.delete("1.0", "end")
        self.pdf_cad_log(f"🚀 Începere conversie CAD ({len(target_files)} fișiere)...")
        self.pdf_cad_log(f"📁 Destinație: {dest_dir}\n")

        def _worker():
            converter = PdfToCadConverter(oda_path=oda_path)
            succ, fail, res = converter.convert_batch(
                pdf_paths=target_files,
                output_dir=dest_dir,
                output_format=export_fmt,
                only_front_view=only_front,
                calibrate_1to1=calibrate,
                progress_callback=lambda p: self.after(0, lambda val=p: self.pdf_cad_progress.set(val)),
                log_callback=lambda m: self.after(0, lambda msg=m: self.pdf_cad_log(msg))
            )
            self.after(0, lambda: self._on_pdf_cad_process_complete(succ, fail, dest_dir))

        threading.Thread(target=_worker, daemon=True).start()

    def _on_pdf_cad_process_complete(self, succ, fail, dest_dir):
        self.btn_start_pdf_cad.configure(state="normal", text="🚀 Generează Fișiere CAD (DWG / DXF)")
        self.pdf_cad_progress.set(1.0)
        self.pdf_cad_log(f"\n{'='*50}")
        self.pdf_cad_log(f"✅ Conversie finalizată: {succ} fișiere generate cu succes, {fail} eșuate.")
        self.pdf_cad_log(f"📁 Director rezultate: {dest_dir}")
        messagebox.showinfo(
            "Conversie Finalizată",
            f"Procesul de generare CAD a fost finalizat!\n\nSucces: {succ} fișiere\nEșuate: {fail}\n\nLocație: {dest_dir}"
        )


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
        Build a dict { normalized_part_name: quantity } from the BOM Excel across ALL sheets.
        Handles both explicit columns, standard named tables, and multi-section A4 layouts
        with multiple side-by-side (part, qty) column pairs per sheet.
        Returns an empty dict on failure.
        """
        try:
            xl = pd.ExcelFile(excel_path)
            result = {}
            qty_keywords = ['qty', 'quantity', 'buc', 'cantitate', 'cant', 'count', 'stk', 'anz', 'pcs', 'pieces', 'bucati', 'ramase']
            part_keywords = ['denumire', 'pies', 'part', 'name', 'nume', 'reper', 'poz', 'element', 'mark']
            nr_keywords = ['nr', 'crt', 'no', 'index', 'item']

            for sheet_name in xl.sheet_names:
                try:
                    df = xl.parse(sheet_name)
                except Exception:
                    continue

                if df.empty:
                    continue

                cols = list(df.columns)
                num_cols = len(cols)
                part_cols = []
                qty_cols = []

                for i, col in enumerate(cols):
                    c_clean = str(col).lower().strip()
                    base_clean = re.sub(r'\.\d+$', '', c_clean).strip()
                    is_nr = any(kw == base_clean or c_clean.startswith(kw + ' ') or c_clean.startswith(kw + '.') for kw in nr_keywords)

                    is_part = False
                    if "unnamed" not in c_clean:
                        if part_col_name and part_col_name.lower() in c_clean:
                            is_part = True
                        elif not part_col_name and any(kw in base_clean for kw in part_keywords) and not is_nr:
                            is_part = True

                    is_qty = False
                    if qty_col_name and qty_col_name.lower() in c_clean:
                        is_qty = True
                    elif not qty_col_name and any(kw in base_clean for kw in qty_keywords):
                        is_qty = True

                    if is_part:
                        part_cols.append((i, col))
                    elif is_qty:
                        qty_cols.append((i, col))

                pairs = []
                for p_idx, p_col in part_cols:
                    next_p_idx = num_cols
                    for np_idx, _ in part_cols:
                        if np_idx > p_idx:
                            next_p_idx = np_idx
                            break

                    cand_q = [q for q in qty_cols if p_idx < q[0] < next_p_idx]
                    if cand_q:
                        pairs.append((p_col, cand_q[0][1]))

                if pairs:
                    for p_col, q_col in pairs:
                        for idx in df.index:
                            pname = str(df[p_col][idx]).strip()
                            if not pname or pname.lower() == 'nan':
                                continue
                            p_clean = pname.lower()
                            if any(kw in p_clean for kw in part_keywords + qty_keywords + nr_keywords):
                                continue
                            try:
                                qty = int(float(df[q_col][idx]))
                            except (ValueError, TypeError):
                                continue
                            norm = self._normalize_name(pname)
                            result[norm] = result.get(norm, 0) + qty
                    continue

                # Fallback: multi-section generic scan for unheadered or complex multi-column sheets
                try:
                    df_raw = xl.parse(sheet_name, header=None)
                    num_cols_raw = df_raw.shape[1]
                    if num_cols_raw < 2:
                        continue
                    used_cols = set()
                    valid_pairs = []

                    for c in range(num_cols_raw - 1):
                        if c in used_cols or (c + 1) in used_cols:
                            continue

                        # Filter out pure row-number columns (1, 2, 3...)
                        c_text_count = 0
                        c_num_count = 0
                        for r in range(min(15, df_raw.shape[0])):
                            v = str(df_raw.iloc[r, c]).strip()
                            if not v or v.lower() == 'nan':
                                continue
                            if any(ch.isalpha() for ch in v):
                                c_text_count += 1
                            elif v.isdigit():
                                c_num_count += 1
                        if c_num_count > 3 and c_text_count <= 1:
                            continue

                        valid_rows = 0
                        is_qty_header = False

                        for r in range(min(5, df_raw.shape[0])):
                            val_q_clean = ''.join(ch for ch in str(df_raw.iloc[r, c + 1]).lower() if ch.isalpha())
                            if any(kw in val_q_clean for kw in qty_keywords):
                                is_qty_header = True
                                break

                        for r in range(min(20, df_raw.shape[0])):
                            p = str(df_raw.iloc[r, c]).strip()
                            q = df_raw.iloc[r, c + 1]
                            if p and p.lower() != 'nan' and not any(kw in p.lower() for kw in part_keywords + qty_keywords + nr_keywords):
                                try:
                                    int(float(q))
                                    valid_rows += 1
                                except (ValueError, TypeError):
                                    pass

                        if is_qty_header or valid_rows >= 1:
                            valid_pairs.append((c, c + 1))
                            used_cols.add(c)
                            used_cols.add(c + 1)

                    for p_idx, q_idx in valid_pairs:
                        for r in range(df_raw.shape[0]):
                            pname = str(df_raw.iloc[r, p_idx]).strip()
                            qty_val = df_raw.iloc[r, q_idx]

                            if not pname or pname.lower() == 'nan':
                                continue
                            if any(kw in pname.lower() for kw in part_keywords + qty_keywords + nr_keywords):
                                continue

                            try:
                                qty = int(float(qty_val))
                            except (ValueError, TypeError):
                                continue

                            norm = self._normalize_name(pname)
                            result[norm] = result.get(norm, 0) + qty
                except Exception:
                    continue

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

    def highlight_missing_parts_in_reports(self, folder_path, missing_norms):
        """
        Highlight missing part rows in red (FF9999) across all parts reports in folder_path.
        Maintains 100% of rows, columns, and existing formulas/structure.
        """
        if not missing_norms or not os.path.exists(folder_path):
            return
        red_fill = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")
        for root, _, files in os.walk(folder_path):
            for f in files:
                f_lower = f.lower()
                if not f_lower.endswith(".xlsx"):
                    continue
                if any(k in f_lower for k in ["status", "dwg", "bom"]):
                    continue
                if not f_lower.startswith("raport"):
                    continue
                
                excel_path = os.path.join(root, f)
                try:
                    wb = load_workbook(excel_path)
                    modified = False
                    for ws in wb.worksheets:
                        for row in ws.iter_rows():
                            row_matches = False
                            for cell in row:
                                if cell.value is not None:
                                    val_str = str(cell.value).strip()
                                    if val_str and self._normalize_name(val_str) in missing_norms:
                                        row_matches = True
                                        break
                            if row_matches:
                                modified = True
                                for cell in row:
                                    if cell.value is not None:
                                        cell.fill = red_fill
                    if modified:
                        safe_save_workbook(wb, excel_path, log_callback=self.log)
                        self.log(f"  ✅ Repere lipsă evidențiate în: {f}")
                except Exception as e:
                    self.log(f"  Avertisment evidențiere {f}: {e}")

    def _get_missing_parts_from_status(self, folder_path):
        """Find Raport_Status_Extractie.xlsx and return a set of normalized missing part names."""
        missing = set()
        search_dirs = [folder_path, os.path.dirname(folder_path)]
        if hasattr(self, 'nc_source_dir') and self.nc_source_dir.get():
            search_dirs.append(self.nc_source_dir.get())
        if hasattr(self, 'dest_dir') and self.dest_dir.get():
            search_dirs.append(self.dest_dir.get())

        for s_dir in search_dirs:
            if not s_dir or not os.path.isdir(s_dir):
                continue
            status_file = os.path.join(s_dir, "Raport_Status_Extractie.xlsx")
            if os.path.isfile(status_file):
                try:
                    wb = load_workbook(status_file, data_only=True)
                    ws = wb.active
                    for row in ws.iter_rows(min_row=2, values_only=True):
                        if not row or not row[0]:
                            continue
                        part_name = str(row[0]).strip()
                        has_x = any(str(c).strip().upper() == 'X' for c in row[1:] if c is not None)
                        if has_x:
                            missing.add(self._normalize_name(part_name))
                    if missing:
                        break
                except Exception:
                    pass
        return missing

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

                missing_norms = self._get_missing_parts_from_status(os.path.dirname(filepath))
                missing_fill = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")

                for data_idx in range(sec_start, sec_end):
                    row_num = row_offset + 1 + (data_idx - sec_start)
                    item = data[data_idx]
                    use_alt = (data_idx - sec_start) % 2 == 1
                    is_missing = (not no_colors) and (self._normalize_name(item["Denumire"]) in missing_norms)

                    # Denumire
                    cell = ws.cell(row=row_num, column=col_start, value=item["Denumire"])
                    cell.font = data_font
                    cell.border = thin_border
                    cell.alignment = left_align
                    if is_missing:
                        cell.fill = missing_fill
                    elif use_alt and not no_colors:
                        cell.fill = alt_fill

                    # Bucăți
                    cell = ws.cell(row=row_num, column=col_start + 1, value=item["Bucati"])
                    cell.font = data_font
                    cell.border = thin_border
                    cell.alignment = center_align
                    if is_missing:
                        cell.fill = missing_fill
                    elif use_alt and not no_colors:
                        cell.fill = alt_fill

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
                
            missing_norms = self._get_missing_parts_from_status(out_dir)
            missing_fill = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")
            data = sorted(groups[sheet_name], key=lambda x: x["Denumire"])
            for row_data in data:
                cell1 = ws.cell(row=row_idx, column=1, value=row_data["Denumire"])
                cell2 = ws.cell(row=row_idx, column=2, value=row_data["Material"])
                cell3 = ws.cell(row=row_idx, column=3, value=row_data["Bucati"])
                for c in [cell1, cell2, cell3]:
                    c.border = thin_border
                if not no_colors and self._normalize_name(row_data["Denumire"]) in missing_norms:
                    for c in [cell1, cell2, cell3]:
                        c.fill = missing_fill
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


    # ══════════════════════════════════════════════════════════════════════
    #  DWG TOOLS INTEGRATION (INSIDE DWG FINDER TAB)
    # ══════════════════════════════════════════════════════════════════════
    def _on_dwg_subview_change(self, val=None):
        """Switch between DWG Search, Organizare Piese (BOM), and Regrupare DWG."""
        mode = self.dwg_subview_mode.get()
        if hasattr(self, "dwg_search_container"):
            self.dwg_search_container.pack_forget()
        if hasattr(self, "dwg_org_container"):
            self.dwg_org_container.pack_forget()
        if hasattr(self, "dwg_regroup_container"):
            self.dwg_regroup_container.pack_forget()

        if "căutare" in mode.lower():
            self.dwg_search_container.pack(fill="both", expand=True, padx=12, pady=6)
        elif "organizare" in mode.lower():
            self.dwg_org_container.pack(fill="both", expand=True, padx=12, pady=6)
        elif "regrupare" in mode.lower():
            self.dwg_regroup_container.pack(fill="both", expand=True, padx=12, pady=6)
        self.save_settings()

    def _build_dwg_organize_view(self):
        """Build UI for Modulul 1: Extragere pe Piese (BOM -> Piese)."""
        parent = self.dwg_org_container
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=1)

        scroll = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=4, pady=4)
        scroll.grid_columnconfigure(0, weight=1)

        # Header card
        hdr = self._make_card(scroll)
        hdr.pack(fill="x", pady=(0, 8))
        hdr.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(hdr, text="📦", font=("Segoe UI Emoji", 24)).grid(row=0, column=0, rowspan=2, padx=(16, 10), pady=10)
        ctk.CTkLabel(hdr, text="Organizare Automată pe Piese (BOM ➜ Piese)", font=("Segoe UI", 15, "bold"), text_color=_TEXT_PRI).grid(row=0, column=1, sticky="w", pady=(8, 0))
        ctk.CTkLabel(hdr, text="Citește tabelul BOM Excel și creează un folder pentru fiecare piesă: mută fișierul NC, desenele PDF de ansamblu și DWG-urile pe grosimi.", font=("Segoe UI", 12), text_color=_TEXT_SEC, wraplength=850, justify="left").grid(row=1, column=1, sticky="w", pady=(0, 8))

        # Card 1: Excel & Columns
        card_xl = self._make_card(scroll)
        card_xl.pack(fill="x", pady=(0, 8))
        card_xl.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(card_xl, text="1. Tabel Excel (Ansambluri & Piese)", font=("Segoe UI", 13, "bold"), text_color=_TEXT_PRI).grid(row=0, column=0, columnspan=3, sticky="w", padx=16, pady=(10, 4))

        self._label(card_xl, text="Fișier Excel BOM:").grid(row=1, column=0, sticky="w", padx=(16, 6), pady=4)
        self._entry(card_xl, self.dwg_org_excel, placeholder_text="Selectează fișierul Excel (.xlsx, .xls)...").grid(row=1, column=1, sticky="ew", padx=4, pady=4)
        self._browse_btn(card_xl, self._browse_dwg_org_excel).grid(row=1, column=2, padx=(4, 16), pady=4)

        # Sheet & Columns selection row
        cols_fr = ctk.CTkFrame(card_xl, fg_color="transparent")
        cols_fr.grid(row=2, column=0, columnspan=3, sticky="ew", padx=12, pady=(4, 6))
        cols_fr.grid_columnconfigure((1, 3, 5), weight=1)

        self._label(cols_fr, text="Foaie (Sheet):").grid(row=0, column=0, sticky="w", padx=4, pady=2)
        self.combo_dwg_org_sheet = ctk.CTkComboBox(cols_fr, values=["-"], command=self._on_dwg_org_sheet_change, variable=self.dwg_org_sheet)
        self.combo_dwg_org_sheet.grid(row=0, column=1, sticky="ew", padx=4, pady=2)

        self._label(cols_fr, text="Col. Ansamblu:").grid(row=0, column=2, sticky="w", padx=(12, 4), pady=2)
        self.combo_dwg_org_asmb_col = ctk.CTkComboBox(cols_fr, values=["-"], command=self._on_dwg_org_column_change, variable=self.dwg_org_asmb_col)
        self.combo_dwg_org_asmb_col.grid(row=0, column=3, sticky="ew", padx=4, pady=2)

        self._label(cols_fr, text="Col. Piesă:").grid(row=0, column=4, sticky="w", padx=(12, 4), pady=2)
        self.combo_dwg_org_part_col = ctk.CTkComboBox(cols_fr, values=["-"], command=self._on_dwg_org_column_change, variable=self.dwg_org_part_col)
        self.combo_dwg_org_part_col.grid(row=0, column=5, sticky="ew", padx=4, pady=2)

        self.lbl_dwg_org_excel_info = ctk.CTkLabel(card_xl, text="ℹ️ Selectează un fișier Excel pentru detectarea automată a coloanelor.", font=("Segoe UI", 11), text_color=_TEXT_SEC)
        self.lbl_dwg_org_excel_info.grid(row=3, column=0, columnspan=3, sticky="w", padx=16, pady=(0, 8))

        # Card 2: Foldere Sursă & Destinație
        card_dirs = self._make_card(scroll)
        card_dirs.pack(fill="x", pady=(0, 8))
        card_dirs.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(card_dirs, text="2. Foldere de Lucru (Sursă & Destinație)", font=("Segoe UI", 13, "bold"), text_color=_TEXT_PRI).grid(row=0, column=0, columnspan=3, sticky="w", padx=16, pady=(10, 4))

        self._label(card_dirs, text="Folder Sursă:").grid(row=1, column=0, sticky="w", padx=(16, 6), pady=4)
        self._entry(card_dirs, self.dwg_org_source, placeholder_text="Folder ce conține PDF-uri, NC-uri și subfoldere cu grosimi...").grid(row=1, column=1, sticky="ew", padx=4, pady=4)
        self._browse_btn(card_dirs, self._browse_dwg_org_source).grid(row=1, column=2, padx=(4, 16), pady=4)

        self.lbl_dwg_org_source_info = ctk.CTkLabel(card_dirs, text="", font=("Segoe UI", 11), text_color=_TEXT_SEC)
        self.lbl_dwg_org_source_info.grid(row=2, column=1, sticky="w", padx=4, pady=(0, 4))

        self._label(card_dirs, text="Folder Destinație:").grid(row=3, column=0, sticky="w", padx=(16, 6), pady=4)
        self._entry(card_dirs, self.dwg_org_dest, placeholder_text="Folderul unde se vor crea folderele fiecărei piese...").grid(row=3, column=1, sticky="ew", padx=4, pady=4)
        self._browse_btn(card_dirs, self._browse_dwg_org_dest).grid(row=3, column=2, padx=(4, 16), pady=4)

        # Card 3: Opțiuni & Butoane
        card_opts = self._make_card(scroll)
        card_opts.pack(fill="x", pady=(0, 8))

        ctk.CTkLabel(card_opts, text="3. Mod Operare și Opțiuni", font=("Segoe UI", 13, "bold"), text_color=_TEXT_PRI).pack(anchor="w", padx=16, pady=(10, 4))

        sub_opts = ctk.CTkFrame(card_opts, fg_color="transparent")
        sub_opts.pack(fill="x", padx=16, pady=(0, 8))

        ctk.CTkRadioButton(sub_opts, text="Mută fișierele (Move)", variable=self.dwg_org_action, value="move").pack(side="left", padx=(0, 20))
        ctk.CTkRadioButton(sub_opts, text="Copiază fișierele (Copy - păstrează originalele)", variable=self.dwg_org_action, value="copy").pack(side="left", padx=(0, 20))
        ctk.CTkCheckBox(sub_opts, text="Include variații DWG (ex. T, L, C)", variable=self.dwg_org_variations).pack(side="left", padx=(0, 20))
        ctk.CTkCheckBox(sub_opts, text="Curăță folderele goale la final (Mutare)", variable=self.dwg_org_clean_empty).pack(side="left")

        # Action Buttons
        btn_bar = ctk.CTkFrame(scroll, fg_color="transparent")
        btn_bar.pack(fill="x", pady=(2, 6))
        btn_bar.grid_columnconfigure((0, 1), weight=1)

        self.btn_dwg_org_run = ctk.CTkButton(
            btn_bar, text="🚀  ÎNCEPE ORGANIZAREA PE PIESE",
            font=("Segoe UI", 13, "bold"), height=40,
            fg_color=_GREEN, hover_color=_GREEN_HOV,
            command=self._start_dwg_org_pipeline
        )
        self.btn_dwg_org_run.grid(row=0, column=0, padx=(0, 6), sticky="ew")

        self.btn_dwg_org_cancel = ctk.CTkButton(
            btn_bar, text="🛑  Oprește / Anulează",
            font=("Segoe UI", 13), height=40,
            fg_color=_RED, hover_color=_RED_HOV,
            state="disabled",
            command=self._cancel_dwg_org_pipeline
        )
        self.btn_dwg_org_cancel.grid(row=0, column=1, padx=(6, 0), sticky="ew")

        self.dwg_org_progress = ctk.CTkProgressBar(scroll, height=10)
        self.dwg_org_progress.pack(fill="x", pady=(4, 2))
        self.dwg_org_progress.set(0)

        self.lbl_dwg_org_status = ctk.CTkLabel(scroll, text="Așteptare pornire...", font=("Segoe UI", 11), text_color=_TEXT_SEC)
        self.lbl_dwg_org_status.pack(anchor="w", padx=4, pady=(0, 6))

        # Log & post actions
        card_log = self._make_card(scroll)
        card_log.pack(fill="both", expand=True, pady=(0, 6))

        head_log = ctk.CTkFrame(card_log, fg_color="transparent")
        head_log.pack(fill="x", padx=12, pady=(8, 4))
        ctk.CTkLabel(head_log, text="Jurnal Activitate:", font=("Segoe UI", 12, "bold"), text_color=_TEXT_PRI).pack(side="left")
        ctk.CTkButton(head_log, text="🧹 Curăță", width=80, height=26, fg_color=_ACCENT, hover_color=_ACCENT_HOV, command=lambda: self._dwg_org_clear_log()).pack(side="right")

        self.dwg_org_log_box = ctk.CTkTextbox(card_log, height=180, font=("Cascadia Code", 11), fg_color=_LOG_BG, text_color="#a3e635")
        self.dwg_org_log_box.pack(fill="both", expand=True, padx=12, pady=(0, 8))

        post_bar = ctk.CTkFrame(scroll, fg_color="transparent")
        post_bar.pack(fill="x", pady=(2, 8))
        self.btn_dwg_org_open_dest = ctk.CTkButton(post_bar, text="📁 Deschide Folder Destinație", width=190, height=32, state="disabled", fg_color=_ACCENT, hover_color=_ACCENT_HOV, command=self._open_dwg_org_dest)
        self.btn_dwg_org_open_dest.pack(side="left", padx=(0, 8))
        self.btn_dwg_org_open_rep = ctk.CTkButton(post_bar, text="📊 Deschide Raport Excel", width=190, height=32, state="disabled", fg_color=_ACCENT, hover_color=_ACCENT_HOV, command=self._open_dwg_org_report)
        self.btn_dwg_org_open_rep.pack(side="left")

    def _browse_dwg_org_excel(self):
        initial = self._get_browse_initial_dir("dest_based", self.dwg_org_excel.get())
        f = filedialog.askopenfilename(title="Selectează Fișierul Excel BOM", initialdir=initial, filetypes=[("Excel", "*.xlsx *.xls *.xlsm"), ("All Files", "*.*")])
        if f:
            norm = os.path.normpath(f)
            self.dwg_org_excel.set(norm)
            self.save_settings()
            self._on_dwg_org_excel_selected(norm)

    def _on_dwg_org_excel_selected(self, filepath):
        try:
            sheets, active_sh, h_row, asmb_idx, part_idx, headers = dwg_organizer.detect_excel_structure(filepath)
            self.dwg_org_excel_meta = {
                "sheets": sheets, "active_sheet": active_sh, "header_row": h_row,
                "asmb_idx": asmb_idx, "part_idx": part_idx, "headers": headers
            }
            if sheets:
                self.combo_dwg_org_sheet.configure(values=sheets)
                self.dwg_org_sheet.set(active_sh)
            col_vals = [f"Col {idx}: {name}" for idx, name in headers.items()]
            if col_vals:
                self.combo_dwg_org_asmb_col.configure(values=col_vals)
                self.combo_dwg_org_part_col.configure(values=col_vals)
                if asmb_idx in headers:
                    self.dwg_org_asmb_col.set(f"Col {asmb_idx}: {headers[asmb_idx]}")
                if part_idx in headers:
                    self.dwg_org_part_col.set(f"Col {part_idx}: {headers[part_idx]}")
            self.lbl_dwg_org_excel_info.configure(
                text=f"✅ Detectat: Sheet '{active_sh}', Rând Antet: {h_row}, Asmb: Col {asmb_idx}, Piesă: Col {part_idx}",
                text_color=_GREEN[1]
            )
        except Exception as e:
            self.lbl_dwg_org_excel_info.configure(text=f"⚠ Eroare la citirea Excel-ului: {e}", text_color=_RED[1])

    def _on_dwg_org_sheet_change(self, val):
        f = self.dwg_org_excel.get().strip()
        if f and os.path.isfile(f):
            try:
                sheets, active_sh, h_row, asmb_idx, part_idx, headers = dwg_organizer.detect_excel_structure(f, sheet_name=val)
                self.dwg_org_excel_meta = {
                    "sheets": sheets, "active_sheet": active_sh, "header_row": h_row,
                    "asmb_idx": asmb_idx, "part_idx": part_idx, "headers": headers
                }
                col_vals = [f"Col {idx}: {name}" for idx, name in headers.items()]
                if col_vals:
                    self.combo_dwg_org_asmb_col.configure(values=col_vals)
                    self.combo_dwg_org_part_col.configure(values=col_vals)
                    if asmb_idx in headers:
                        self.dwg_org_asmb_col.set(f"Col {asmb_idx}: {headers[asmb_idx]}")
                    if part_idx in headers:
                        self.dwg_org_part_col.set(f"Col {part_idx}: {headers[part_idx]}")
            except Exception:
                pass

    def _on_dwg_org_column_change(self, val=None):
        pass

    def _browse_dwg_org_source(self):
        initial = self._get_browse_initial_dir("dest_based", self.dwg_org_source.get())
        d = filedialog.askdirectory(title="Selectează Folderul Sursă", initialdir=initial)
        if d:
            norm = os.path.normpath(d)
            self.dwg_org_source.set(norm)
            self.save_settings()
            try:
                nc_files, pdf_files, dwg_files, dwg_map = dwg_organizer.scan_source_files(norm)
                total_dwg = sum(len(v) for v in dwg_map.values())
                self.lbl_dwg_org_source_info.configure(
                    text=f"Găsit în sursă: {len(nc_files)} NC-uri, {len(pdf_files)} PDF-uri, {total_dwg} DWG-uri în {len(dwg_map)} grosimi.",
                    text_color=_GREEN[1]
                )
            except Exception:
                pass

    def _browse_dwg_org_dest(self):
        initial = self._get_browse_initial_dir("dest_based", self.dwg_org_dest.get())
        d = filedialog.askdirectory(title="Selectează Folderul Destinație Piese", initialdir=initial)
        if d:
            norm = os.path.normpath(d)
            self.dwg_org_dest.set(norm)
            self.save_settings()

    def _start_dwg_org_pipeline(self):
        if self.dwg_org_running:
            return
        excel_p = self.dwg_org_excel.get().strip()
        src_p = self.dwg_org_source.get().strip()
        dst_p = self.dwg_org_dest.get().strip()

        if not excel_p or not os.path.isfile(excel_p):
            messagebox.showwarning("Atenție", "Selectează un fișier Excel BOM valid.")
            return
        if not src_p or not os.path.isdir(src_p):
            messagebox.showwarning("Atenție", "Selectează un folder Sursă valid.")
            return
        if not dst_p:
            messagebox.showwarning("Atenție", "Selectează un folder Destinație.")
            return

        meta = getattr(self, "dwg_org_excel_meta", None)
        h_row = meta.get("header_row", 1) if meta else 1
        sh_name = self.dwg_org_sheet.get()
        if sh_name == "-":
            sh_name = meta.get("active_sheet") if meta else None

        # Extract col indices
        def parse_col_idx(val_str, fallback):
            m = re.match(r'Col\s+(\d+)', val_str)
            return int(m.group(1)) if m else fallback

        asmb_idx = parse_col_idx(self.dwg_org_asmb_col.get(), meta.get("asmb_idx", 1) if meta else 1)
        part_idx = parse_col_idx(self.dwg_org_part_col.get(), meta.get("part_idx", 2) if meta else 2)

        self.dwg_org_running = True
        self.dwg_org_cancel = False
        self.btn_dwg_org_run.configure(state="disabled", text="⏳ Se organizează...")
        self.btn_dwg_org_cancel.configure(state="normal")
        self.dwg_org_progress.set(0)
        self._dwg_org_clear_log()

        mode = self.dwg_org_action.get()
        allow_vars = self.dwg_org_variations.get()
        clean_empty = self.dwg_org_clean_empty.get()

        def worker():
            def prog_cb(curr, total, msg):
                self.after(0, lambda: self._update_dwg_org_progress(curr, total, msg))
            def log_cb(txt, tag="INFO"):
                self.after(0, lambda: self._dwg_org_log(txt, tag))
            def cancel_cb():
                return self.dwg_org_cancel

            res = dwg_organizer.run_pipeline(
                excel_path=excel_p,
                sheet_name=sh_name,
                header_row=h_row,
                asmb_col_idx=asmb_idx,
                part_col_idx=part_idx,
                source_dir=src_p,
                dest_dir=dst_p,
                mode=mode,
                allow_variations=allow_vars,
                clean_empty_folders=clean_empty,
                progress_callback=prog_cb,
                log_callback=log_cb,
                is_cancelled=cancel_cb
            )
            self.after(0, lambda: self._on_dwg_org_complete(res))

        threading.Thread(target=worker, daemon=True).start()

    def _cancel_dwg_org_pipeline(self):
        self.dwg_org_cancel = True
        self._dwg_org_log("🛑 Solicitare de oprire trimisă...", "WARNING")

    def _update_dwg_org_progress(self, curr, total, msg):
        frac = curr / max(total, 1)
        self.dwg_org_progress.set(frac)
        self.lbl_dwg_org_status.configure(text=msg)

    def _dwg_org_log(self, text, tag="INFO"):
        self.dwg_org_log_box.configure(state="normal")
        self.dwg_org_log_box.insert("end", f"{text}\n")
        self.dwg_org_log_box.see("end")
        self.dwg_org_log_box.configure(state="disabled")

    def _dwg_org_clear_log(self):
        self.dwg_org_log_box.configure(state="normal")
        self.dwg_org_log_box.delete("1.0", "end")
        self.dwg_org_log_box.configure(state="disabled")

    def _on_dwg_org_complete(self, res):
        self.dwg_org_running = False
        self.btn_dwg_org_run.configure(state="normal", text="🚀  ÎNCEPE ORGANIZAREA PE PIESE")
        self.btn_dwg_org_cancel.configure(state="disabled")
        self.dwg_org_progress.set(1.0)

        rep_path = res.get("report_path") if isinstance(res, dict) else None
        self.dwg_org_last_report = rep_path
        if rep_path and os.path.exists(rep_path):
            self.btn_dwg_org_open_rep.configure(state="normal")
        dst = self.dwg_org_dest.get().strip()
        if dst and os.path.exists(dst):
            self.btn_dwg_org_open_dest.configure(state="normal")

        if res.get("cancelled"):
            messagebox.showinfo("Oprit", "Procesul de organizare a fost oprit de utilizator.")
        elif res.get("error"):
            messagebox.showerror("Eroare", f"A apărut o eroare la organizare:\n{res['error']}")
        else:
            messagebox.showinfo(
                "Organizare Finalizată",
                f"Organizarea pe piese a fost finalizată cu succes!\n\n"
                f"• Piese procesate: {res.get('processed_parts', 0)}\n"
                f"• NC-uri transferate: {res.get('nc_transferred', 0)}\n"
                f"• PDF-uri transferate: {res.get('pdf_transferred', 0)}\n"
                f"• DWG-uri transferate: {res.get('dwg_transferred', 0)}\n\n"
                f"Folder Destinație:\n{dst}"
            )

    def _open_dwg_org_dest(self):
        d = self.dwg_org_dest.get().strip()
        if d and os.path.exists(d):
            os.startfile(d)

    def _open_dwg_org_report(self):
        if self.dwg_org_last_report and os.path.exists(self.dwg_org_last_report):
            os.startfile(self.dwg_org_last_report)

    # ───────────────────────────────────────────────────────────────────
    # Subview 3: Regrupare DWG pe Grosimi
    # ───────────────────────────────────────────────────────────────────
    def _build_dwg_regroup_view(self):
        """Build UI for Modulul 2: Regrupare DWG pe Grosimi (Piese -> Grosimi)."""
        parent = self.dwg_regroup_container
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=1)

        scroll = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=4, pady=4)
        scroll.grid_columnconfigure(0, weight=1)

        # Header card
        hdr = self._make_card(scroll)
        hdr.pack(fill="x", pady=(0, 8))
        hdr.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(hdr, text="🔄", font=("Segoe UI Emoji", 24)).grid(row=0, column=0, rowspan=2, padx=(16, 10), pady=10)
        ctk.CTkLabel(hdr, text="Regrupare DWG pe Grosimi de Tablă (Piese ➜ Grosimi)", font=("Segoe UI", 15, "bold"), text_color=_TEXT_PRI).grid(row=0, column=1, sticky="w", pady=(8, 0))
        ctk.CTkLabel(hdr, text="Scanează folderele de piese organizate, extrage exclusiv fișierele DWG și le regrupează în foldere unificate după grosime (T10, T12, T15 etc.).", font=("Segoe UI", 12), text_color=_TEXT_SEC, wraplength=850, justify="left").grid(row=1, column=1, sticky="w", pady=(0, 8))

        # Card 1: Directoare
        card_dirs = self._make_card(scroll)
        card_dirs.pack(fill="x", pady=(0, 8))
        card_dirs.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(card_dirs, text="1. Foldere de Lucru (Sursă Piese & Destinație Grosimi)", font=("Segoe UI", 13, "bold"), text_color=_TEXT_PRI).grid(row=0, column=0, columnspan=3, sticky="w", padx=16, pady=(10, 4))

        self._label(card_dirs, text="Folder Piese (Sursă):").grid(row=1, column=0, sticky="w", padx=(16, 6), pady=4)
        self._entry(card_dirs, self.dwg_regroup_source, placeholder_text="Selectează folderul ce conține folderele de piese...").grid(row=1, column=1, sticky="ew", padx=4, pady=4)
        self._browse_btn(card_dirs, self._browse_dwg_regroup_source).grid(row=1, column=2, padx=(4, 16), pady=4)

        self.lbl_dwg_rg_src_info = ctk.CTkLabel(card_dirs, text="", font=("Segoe UI", 11), text_color=_TEXT_SEC)
        self.lbl_dwg_rg_src_info.grid(row=2, column=1, sticky="w", padx=4, pady=(0, 4))

        self._label(card_dirs, text="Folder Destinație Grosimi:").grid(row=3, column=0, sticky="w", padx=(16, 6), pady=4)
        self._entry(card_dirs, self.dwg_regroup_dest, placeholder_text="Folderul unde se vor crea folderele cu grosimi (T10, T15...)...").grid(row=3, column=1, sticky="ew", padx=4, pady=4)
        self._browse_btn(card_dirs, self._browse_dwg_regroup_dest).grid(row=3, column=2, padx=(4, 16), pady=4)

        # Card 2: Opțiuni
        card_opts = self._make_card(scroll)
        card_opts.pack(fill="x", pady=(0, 8))

        ctk.CTkLabel(card_opts, text="2. Opțiuni Regrupare DWG", font=("Segoe UI", 13, "bold"), text_color=_TEXT_PRI).pack(anchor="w", padx=16, pady=(10, 4))
        sub_opts = ctk.CTkFrame(card_opts, fg_color="transparent")
        sub_opts.pack(fill="x", padx=16, pady=(0, 8))

        ctk.CTkRadioButton(sub_opts, text="Copiază fișierele DWG (Copy - păstrează originalele)", variable=self.dwg_regroup_action, value="copy").pack(side="left", padx=(0, 20))
        ctk.CTkRadioButton(sub_opts, text="Mută fișierele DWG (Move)", variable=self.dwg_regroup_action, value="move").pack(side="left", padx=(0, 20))
        ctk.CTkCheckBox(sub_opts, text="Unifică denumirile (ex: t10 -> T10)", variable=self.dwg_regroup_normalize).pack(side="left")

        # Action Buttons
        btn_bar = ctk.CTkFrame(scroll, fg_color="transparent")
        btn_bar.pack(fill="x", pady=(2, 6))
        btn_bar.grid_columnconfigure((0, 1), weight=1)

        self.btn_dwg_rg_run = ctk.CTkButton(
            btn_bar, text="🔄  ÎNCEPE REGRUPAREA DWG PE GROSIMI",
            font=("Segoe UI", 13, "bold"), height=40,
            fg_color=_GREEN, hover_color=_GREEN_HOV,
            command=self._start_dwg_regroup_pipeline
        )
        self.btn_dwg_rg_run.grid(row=0, column=0, padx=(0, 6), sticky="ew")

        self.btn_dwg_rg_cancel = ctk.CTkButton(
            btn_bar, text="🛑  Oprește / Anulează",
            font=("Segoe UI", 13), height=40,
            fg_color=_RED, hover_color=_RED_HOV,
            state="disabled",
            command=self._cancel_dwg_regroup_pipeline
        )
        self.btn_dwg_rg_cancel.grid(row=0, column=1, padx=(6, 0), sticky="ew")

        self.dwg_rg_progress = ctk.CTkProgressBar(scroll, height=10)
        self.dwg_rg_progress.pack(fill="x", pady=(4, 2))
        self.dwg_rg_progress.set(0)

        self.lbl_dwg_rg_status = ctk.CTkLabel(scroll, text="Așteptare pornire...", font=("Segoe UI", 11), text_color=_TEXT_SEC)
        self.lbl_dwg_rg_status.pack(anchor="w", padx=4, pady=(0, 6))

        # Log & post actions
        card_log = self._make_card(scroll)
        card_log.pack(fill="both", expand=True, pady=(0, 6))

        head_log = ctk.CTkFrame(card_log, fg_color="transparent")
        head_log.pack(fill="x", padx=12, pady=(8, 4))
        ctk.CTkLabel(head_log, text="Jurnal Regrupare:", font=("Segoe UI", 12, "bold"), text_color=_TEXT_PRI).pack(side="left")
        ctk.CTkButton(head_log, text="🧹 Curăță", width=80, height=26, fg_color=_ACCENT, hover_color=_ACCENT_HOV, command=lambda: self._dwg_rg_clear_log()).pack(side="right")

        self.dwg_rg_log_box = ctk.CTkTextbox(card_log, height=180, font=("Cascadia Code", 11), fg_color=_LOG_BG, text_color="#a3e635")
        self.dwg_rg_log_box.pack(fill="both", expand=True, padx=12, pady=(0, 8))

        post_bar = ctk.CTkFrame(scroll, fg_color="transparent")
        post_bar.pack(fill="x", pady=(2, 8))
        self.btn_dwg_rg_open_dest = ctk.CTkButton(post_bar, text="📁 Deschide Folder Grosimi", width=190, height=32, state="disabled", fg_color=_ACCENT, hover_color=_ACCENT_HOV, command=self._open_dwg_regroup_dest)
        self.btn_dwg_rg_open_dest.pack(side="left", padx=(0, 8))
        self.btn_dwg_rg_open_rep = ctk.CTkButton(post_bar, text="📊 Deschide Raport Excel", width=190, height=32, state="disabled", fg_color=_ACCENT, hover_color=_ACCENT_HOV, command=self._open_dwg_regroup_report)
        self.btn_dwg_rg_open_rep.pack(side="left")

    def _browse_dwg_regroup_source(self):
        initial = self._get_browse_initial_dir("dest_based", self.dwg_regroup_source.get())
        d = filedialog.askdirectory(title="Selectează Folderul cu Piese", initialdir=initial)
        if d:
            norm = os.path.normpath(d)
            self.dwg_regroup_source.set(norm)
            self.save_settings()
            try:
                total, thk_map, unclass = dwg_organizer.scan_dwgs_for_regrouping(norm)
                self.lbl_dwg_rg_src_info.configure(
                    text=f"Găsit: {total} fișiere DWG în {len(thk_map)} grosimi diferite ({len(unclass)} neclasificate).",
                    text_color=_GREEN[1]
                )
            except Exception:
                pass

    def _browse_dwg_regroup_dest(self):
        initial = self._get_browse_initial_dir("dest_based", self.dwg_regroup_dest.get())
        d = filedialog.askdirectory(title="Selectează Folderul Destinație Grosimi", initialdir=initial)
        if d:
            norm = os.path.normpath(d)
            self.dwg_regroup_dest.set(norm)
            self.save_settings()

    def _start_dwg_regroup_pipeline(self):
        if self.dwg_regroup_running:
            return
        src_p = self.dwg_regroup_source.get().strip()
        dst_p = self.dwg_regroup_dest.get().strip()

        if not src_p or not os.path.isdir(src_p):
            messagebox.showwarning("Atenție", "Selectează un folder Sursă valid ce conține folderele de piese.")
            return
        if not dst_p:
            messagebox.showwarning("Atenție", "Selectează un folder Destinație pentru grosimi.")
            return

        self.dwg_regroup_running = True
        self.dwg_regroup_cancel = False
        self.btn_dwg_rg_run.configure(state="disabled", text="⏳ Se regrupează...")
        self.btn_dwg_rg_cancel.configure(state="normal")
        self.dwg_rg_progress.set(0)
        self._dwg_rg_clear_log()

        mode = self.dwg_regroup_action.get()
        norm_case = self.dwg_regroup_normalize.get()

        def worker():
            def prog_cb(curr, total, msg):
                self.after(0, lambda: self._update_dwg_regroup_progress(curr, total, msg))
            def log_cb(txt, tag="INFO"):
                self.after(0, lambda: self._dwg_regroup_log(txt, tag))
            def cancel_cb():
                return self.dwg_regroup_cancel

            res = dwg_organizer.run_regroup_pipeline(
                source_dir=src_p,
                dest_dir=dst_p,
                mode=mode,
                normalize_case=norm_case,
                clean_empty_folders=True,
                progress_callback=prog_cb,
                log_callback=log_cb,
                is_cancelled=cancel_cb
            )
            self.after(0, lambda: self._on_dwg_regroup_complete(res))

        threading.Thread(target=worker, daemon=True).start()

    def _cancel_dwg_regroup_pipeline(self):
        self.dwg_regroup_cancel = True
        self._dwg_regroup_log("🛑 Solicitare de oprire trimisă...", "WARNING")

    def _update_dwg_regroup_progress(self, curr, total, msg):
        frac = curr / max(total, 1)
        self.dwg_rg_progress.set(frac)
        self.lbl_dwg_rg_status.configure(text=msg)

    def _dwg_regroup_log(self, text, tag="INFO"):
        self.dwg_rg_log_box.configure(state="normal")
        self.dwg_rg_log_box.insert("end", f"{text}\n")
        self.dwg_rg_log_box.see("end")
        self.dwg_rg_log_box.configure(state="disabled")

    def _dwg_rg_clear_log(self):
        self.dwg_rg_log_box.configure(state="normal")
        self.dwg_rg_log_box.delete("1.0", "end")
        self.dwg_rg_log_box.configure(state="disabled")

    def _on_dwg_regroup_complete(self, res):
        self.dwg_regroup_running = False
        self.btn_dwg_rg_run.configure(state="normal", text="🔄  ÎNCEPE REGRUPAREA DWG PE GROSIMI")
        self.btn_dwg_rg_cancel.configure(state="disabled")
        self.dwg_rg_progress.set(1.0)

        rep_path = res.get("report_path") if isinstance(res, dict) else None
        self.dwg_regroup_last_report = rep_path
        if rep_path and os.path.exists(rep_path):
            self.btn_dwg_rg_open_rep.configure(state="normal")
        dst = self.dwg_regroup_dest.get().strip()
        if dst and os.path.exists(dst):
            self.btn_dwg_rg_open_dest.configure(state="normal")

        if res.get("cancelled"):
            messagebox.showinfo("Oprit", "Procesul de regrupare a fost oprit de utilizator.")
        elif res.get("error"):
            messagebox.showerror("Eroare", f"A apărut o eroare la regrupare:\n{res['error']}")
        else:
            messagebox.showinfo(
                "Regrupare Finalizată",
                f"Regruparea DWG pe grosimi a fost finalizată cu succes!\n\n"
                f"• Fișiere DWG transferate: {res.get('transferred', 0)}\n"
                f"• Grosimi unificate: {len(res.get('thickness_counts', {}))}\n\n"
                f"Folder Destinație:\n{dst}"
            )

    def _open_dwg_regroup_dest(self):
        d = self.dwg_regroup_dest.get().strip()
        if d and os.path.exists(d):
            os.startfile(d)

    def _open_dwg_regroup_report(self):
        if self.dwg_regroup_last_report and os.path.exists(self.dwg_regroup_last_report):
            os.startfile(self.dwg_regroup_last_report)


    # ══════════════════════════════════════════════════════════════════════
    #  IMPORTS GENERATOR TAB (METALIX WORD .DOC / .DOCX -> .XLSX)
    # ══════════════════════════════════════════════════════════════════════
    def create_imports_generator_tab(self):
        """Construct the dedicated tab for Imports Generator (Word -> Excel)."""
        parent = self.tab_imports
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=1)

        scroll = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=8, pady=6)
        scroll.grid_columnconfigure(0, weight=1)

        # Header Card
        hdr = self._make_card(scroll)
        hdr.pack(fill="x", pady=(0, 8))
        hdr.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(hdr, text="📑", font=("Segoe UI Emoji", 26)).grid(row=0, column=0, rowspan=2, padx=(16, 12), pady=10)
        ctk.CTkLabel(hdr, text="Generator Import-uri Excel din Word (Metalix)", font=("Segoe UI", 15, "bold"), text_color=_TEXT_PRI).grid(row=0, column=1, sticky="w", pady=(8, 0))
        ctk.CTkLabel(hdr, text="Extrage automat piesele și cantitățile din documente Word (.doc / .docx) și generează fișiere Excel de import (.xlsx) cu denumiri curățate.", font=("Segoe UI", 12), text_color=_TEXT_SEC, wraplength=880, justify="left").grid(row=1, column=1, sticky="w", pady=(0, 8))

        # Card 1: Selecție Documente Word
        card_sel = self._make_card(scroll)
        card_sel.pack(fill="x", pady=(0, 8))

        ctk.CTkLabel(card_sel, text="1. Selectează Documente Word (.doc / .docx)", font=("Segoe UI", 13, "bold"), text_color=_TEXT_PRI).pack(anchor="w", padx=16, pady=(10, 6))

        btn_bar1 = ctk.CTkFrame(card_sel, fg_color="transparent")
        btn_bar1.pack(fill="x", padx=16, pady=(0, 8))

        ctk.CTkButton(btn_bar1, text="📄 Selectează Fișier(e) Word", font=("Segoe UI", 12, "bold"), height=34, fg_color=_ACCENT, hover_color=_ACCENT_HOV, command=self._on_imports_pick_word_files).pack(side="left", padx=(0, 8))
        ctk.CTkButton(btn_bar1, text="➕ Adaugă Fișiere", font=("Segoe UI", 12), height=34, width=120, fg_color="#3b82f6", hover_color="#2563eb", command=lambda: self._on_imports_pick_word_files(append=True)).pack(side="left", padx=(0, 8))
        ctk.CTkButton(btn_bar1, text="📂 Încarcă din Folder...", font=("Segoe UI", 12, "bold"), height=34, width=160, fg_color="#0d9488", hover_color="#0f766e", command=self._on_imports_pick_folder).pack(side="left", padx=(0, 8))

        self._label(btn_bar1, text="Filtru:").pack(side="left", padx=(6, 4))
        self._entry(btn_bar1, self.imports_filter_text, placeholder_text="ex: 1147-46", width=120, height=34).pack(side="left", padx=(0, 6))
        ctk.CTkButton(btn_bar1, text="🔍 Plasma (D:)", font=("Segoe UI", 12), height=34, width=110, fg_color="#4f46e5", hover_color="#4338ca", command=self._on_imports_quick_search_plasma).pack(side="left", padx=(0, 8))
        ctk.CTkButton(btn_bar1, text="🗑️ Golește", font=("Segoe UI", 12), height=34, width=80, fg_color=_RED, hover_color=_RED_HOV, command=self._on_imports_clear_files).pack(side="right")

        self.lbl_imports_summary = ctk.CTkLabel(card_sel, text="Niciun document Word selectat. Apasă pe «Selectează Fișier(e) Word» sau caută după cod de proiect.", font=("Segoe UI", 11, "bold"), text_color=_TEXT_SEC)
        self.lbl_imports_summary.pack(anchor="w", padx=16, pady=(0, 10))

        # Card 2: Previzualizare Live
        card_prev = self._make_card(scroll)
        card_prev.pack(fill="both", expand=True, pady=(0, 8))

        head_prev = ctk.CTkFrame(card_prev, fg_color="transparent")
        head_prev.pack(fill="x", padx=16, pady=(8, 4))
        ctk.CTkLabel(head_prev, text="2. Previzualizare Piese Extrase (Original Word ➜ Curățat Excel)", font=("Segoe UI", 13, "bold"), text_color=_TEXT_PRI).pack(side="left")
        ctk.CTkButton(head_prev, text="📋 Copiază Tabelul în Clipboard", font=("Segoe UI", 11, "bold"), height=28, fg_color="#d97706", hover_color="#b45309", command=self._on_imports_copy_to_clipboard).pack(side="right")

        tbl_container = ctk.CTkFrame(card_prev, fg_color="transparent")
        tbl_container.pack(fill="both", expand=True, padx=14, pady=(2, 10))
        tbl_container.grid_columnconfigure(0, weight=2)
        tbl_container.grid_columnconfigure(1, weight=5)
        tbl_container.grid_rowconfigure(0, weight=1)

        # Docs Tree
        box_docs = ctk.CTkFrame(tbl_container, fg_color=_ENTRY_BG, corner_radius=8)
        box_docs.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        box_docs.grid_rowconfigure(1, weight=1)
        box_docs.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(box_docs, text="📑 Documente Word încărcate:", font=("Segoe UI", 11, "bold"), text_color=_TEXT_PRI).grid(row=0, column=0, columnspan=2, sticky="w", padx=8, pady=(6, 2))

        self.tree_imports_docs = ttk.Treeview(box_docs, columns=("doc_name", "rows", "qty"), show="headings", height=8, selectmode="browse", style="DWG.Treeview")
        self.tree_imports_docs.heading("doc_name", text="Document Word")
        self.tree_imports_docs.heading("rows", text="Poz.")
        self.tree_imports_docs.heading("qty", text="Buc.")
        self.tree_imports_docs.column("doc_name", width=170, anchor="w")
        self.tree_imports_docs.column("rows", width=45, anchor="center")
        self.tree_imports_docs.column("qty", width=50, anchor="center")
        self.tree_imports_docs.grid(row=1, column=0, sticky="nsew", padx=(6, 0), pady=(0, 6))

        s_docs = ttk.Scrollbar(box_docs, orient="vertical", command=self.tree_imports_docs.yview)
        self.tree_imports_docs.configure(yscrollcommand=s_docs.set)
        s_docs.grid(row=1, column=1, sticky="ns", padx=(0, 4), pady=(0, 6))
        self.tree_imports_docs.bind("<<TreeviewSelect>>", self._on_imports_doc_tree_select)

        # Parts Tree
        box_parts = ctk.CTkFrame(tbl_container, fg_color=_ENTRY_BG, corner_radius=8)
        box_parts.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        box_parts.grid_rowconfigure(1, weight=1)
        box_parts.grid_columnconfigure(0, weight=1)

        self.lbl_imports_parts_title = ctk.CTkLabel(box_parts, text="🔩 Piese extrase (Col A = Denumire curățată | Col B = Cantitate):", font=("Segoe UI", 11, "bold"), text_color=_TEXT_PRI)
        self.lbl_imports_parts_title.grid(row=0, column=0, columnspan=2, sticky="w", padx=8, pady=(6, 2))

        self.tree_imports_parts = ttk.Treeview(box_parts, columns=("source", "no", "raw_name", "clean_name", "qty", "dims"), show="headings", height=8, style="DWG.Treeview")
        self.tree_imports_parts.heading("source", text="Fișier Word")
        self.tree_imports_parts.heading("no", text="No.")
        self.tree_imports_parts.heading("raw_name", text="Part name (Word)")
        self.tree_imports_parts.heading("clean_name", text="➜ EXCEL (Col A)")
        self.tree_imports_parts.heading("qty", text="Cant. (Col B)")
        self.tree_imports_parts.heading("dims", text="Dimensiuni (mm)")
        self.tree_imports_parts.column("source", width=140, anchor="w")
        self.tree_imports_parts.column("no", width=40, anchor="center")
        self.tree_imports_parts.column("raw_name", width=150, anchor="w")
        self.tree_imports_parts.column("clean_name", width=130, anchor="center")
        self.tree_imports_parts.column("qty", width=75, anchor="center")
        self.tree_imports_parts.column("dims", width=110, anchor="center")
        self.tree_imports_parts.grid(row=1, column=0, sticky="nsew", padx=(6, 0), pady=(0, 6))

        s_parts = ttk.Scrollbar(box_parts, orient="vertical", command=self.tree_imports_parts.yview)
        self.tree_imports_parts.configure(yscrollcommand=s_parts.set)
        s_parts.grid(row=1, column=1, sticky="ns", padx=(0, 4), pady=(0, 6))

        # Card 3: Opțiuni & Format Excel
        card_fmt = self._make_card(scroll)
        card_fmt.pack(fill="x", pady=(0, 8))
        card_fmt.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(card_fmt, text="3. Opțiuni Generare Excel & Folder Salvare", font=("Segoe UI", 13, "bold"), text_color=_TEXT_PRI).grid(row=0, column=0, columnspan=3, sticky="w", padx=16, pady=(10, 6))

        self._label(card_fmt, text="Mod generare Excel:").grid(row=1, column=0, sticky="w", padx=(16, 6), pady=4)
        m_frame = ctk.CTkFrame(card_fmt, fg_color="transparent")
        m_frame.grid(row=1, column=1, columnspan=2, sticky="w", padx=4, pady=4)
        ctk.CTkRadioButton(m_frame, text="Batch: Câte un Excel pentru fiecare Word (1:1)", variable=self.imports_export_mode, value="individual", command=self._on_imports_mode_change).pack(side="left", padx=(0, 14))
        ctk.CTkRadioButton(m_frame, text="Un singur Excel consolidat", variable=self.imports_export_mode, value="consolidated_single", command=self._on_imports_mode_change).pack(side="left", padx=(0, 14))
        ctk.CTkRadioButton(m_frame, text="Un Excel (Sheet per Word)", variable=self.imports_export_mode, value="consolidated_sheets", command=self._on_imports_mode_change).pack(side="left")

        self._label(card_fmt, text="Format tabel Excel:").grid(row=2, column=0, sticky="w", padx=(16, 6), pady=4)
        r_frame = ctk.CTkFrame(card_fmt, fg_color="transparent")
        r_frame.grid(row=2, column=1, columnspan=2, sticky="w", padx=4, pady=4)
        ctk.CTkRadioButton(r_frame, text="Rândul 2 (Rândul 1 gol - stil IMPORT.xlsx)", variable=self.imports_start_row_mode, value="row2_empty1").pack(side="left", padx=(0, 14))
        ctk.CTkRadioButton(r_frame, text="Rândul 1", variable=self.imports_start_row_mode, value="row1").pack(side="left", padx=(0, 14))
        ctk.CTkRadioButton(r_frame, text="Cu Antet (Denumire | Cantitate)", variable=self.imports_start_row_mode, value="header_row1").pack(side="left")

        extra_fmt = ctk.CTkFrame(card_fmt, fg_color="transparent")
        extra_fmt.grid(row=3, column=1, columnspan=2, sticky="w", padx=4, pady=4)
        ctk.CTkCheckBox(extra_fmt, text="Cumulează cantitățile la duplicate", variable=self.imports_merge_dups, command=self._imports_refresh_preview_tables).pack(side="left", padx=(0, 16))
        self._label(extra_fmt, text="Nume fișier consolidat:").pack(side="left", padx=(0, 6))
        self.entry_imports_cons_name = self._entry(extra_fmt, self.imports_consolidated_name, width=160, height=28)
        self.entry_imports_cons_name.pack(side="left")

        self._label(card_fmt, text="Folder Destinație:").grid(row=4, column=0, sticky="w", padx=(16, 6), pady=(4, 10))
        self._entry(card_fmt, self.imports_output_dir).grid(row=4, column=1, sticky="ew", padx=4, pady=(4, 10))
        self._browse_btn(card_fmt, self._on_imports_pick_output_dir).grid(row=4, column=2, padx=(4, 16), pady=(4, 10))

        # Card 4: Butoane Acțiune & Progres
        card_act = self._make_card(scroll)
        card_act.pack(fill="x", pady=(0, 8))

        act_bar = ctk.CTkFrame(card_act, fg_color="transparent")
        act_bar.pack(fill="x", padx=16, pady=(10, 6))

        self.btn_imports_generate = ctk.CTkButton(act_bar, text="🚀 GENEREAZĂ EXCEL (.XLSX)", font=("Segoe UI", 13, "bold"), height=42, fg_color=_GREEN, hover_color=_GREEN_HOV, command=self._on_imports_generate_excel)
        self.btn_imports_generate.pack(side="left", fill="x", expand=True, padx=(0, 8))

        self.btn_imports_quick_desk = ctk.CTkButton(act_bar, text="⚡ Salvează în Desktop\\\\IMPORT.xlsx", font=("Segoe UI", 12, "bold"), height=42, fg_color=_ACCENT, hover_color=_ACCENT_HOV, command=self._on_imports_quick_save_desktop)
        self.btn_imports_quick_desk.pack(side="left", padx=(0, 8))

        self.btn_imports_open_xl = ctk.CTkButton(act_bar, text="📊 Deschide Excel", font=("Segoe UI", 12), height=42, width=140, state="disabled", fg_color="#475569", hover_color="#334155", command=self._on_imports_open_last_excel)
        self.btn_imports_open_xl.pack(side="left", padx=(0, 8))

        self.btn_imports_open_fld = ctk.CTkButton(act_bar, text="📁 Deschide Folder", font=("Segoe UI", 12), height=42, width=140, fg_color="#475569", hover_color="#334155", command=self._on_imports_open_output_folder)
        self.btn_imports_open_fld.pack(side="left")

        self.imports_progress = ctk.CTkProgressBar(card_act, height=10)
        self.imports_progress.pack(fill="x", padx=16, pady=(2, 4))
        self.imports_progress.set(0)

        self.lbl_imports_status = ctk.CTkLabel(card_act, text="Pregătit. Încarcă documente Word pentru a genera fișierele Excel.", font=("Segoe UI", 11), text_color=_TEXT_SEC)
        self.lbl_imports_status.pack(anchor="w", padx=16, pady=(0, 10))

    # Handlers for Imports Generator
    def _on_imports_pick_word_files(self, append=False):
        default_dir = r"D:\Programe plasma" if os.path.exists(r"D:\Programe plasma") else self._get_browse_initial_dir("dest_based") or os.getcwd()
        files = filedialog.askopenfilenames(
            title="Selectează documente Word (.doc / .docx)",
            initialdir=default_dir,
            filetypes=[("Documente Word", "*.doc *.docx *.DOC *.DOCX"), ("Toate fișierele", "*.*")]
        )
        if files:
            self._imports_load_word_files(list(files), append=append)

    def _on_imports_pick_folder(self):
        default_dir = r"D:\Programe plasma" if os.path.exists(r"D:\Programe plasma") else self._get_browse_initial_dir("dest_based") or os.getcwd()
        folder = filedialog.askdirectory(title="Selectează folderul cu documente Word", initialdir=default_dir)
        if folder:
            self._imports_scan_folder(folder)

    def _on_imports_quick_search_plasma(self):
        plasma_dir = r"D:\Programe plasma"
        if not os.path.exists(plasma_dir):
            messagebox.showwarning("Folder inexistent", "Folderul D:\\Programe plasma nu a fost găsit pe acest calculator.\nFolosește butonul «Încarcă din Folder...».")
            return
        flt = self.imports_filter_text.get().strip()
        if not flt:
            messagebox.showinfo("Introdu un filtru", "Scrie în căsuța «Filtru» codul proiectului (ex: 1147-46) pentru a căuta rapid în D:\\Programe plasma.")
            return
        self._imports_scan_folder(plasma_dir)

    def _imports_scan_folder(self, folder):
        flt = self.imports_filter_text.get().strip().lower()
        matched = []
        try:
            for entry in os.listdir(folder):
                if entry.startswith("~$"):
                    continue
                if entry.lower().endswith((".doc", ".docx")):
                    if not flt or flt in entry.lower():
                        matched.append(os.path.join(folder, entry))
        except Exception as exc:
            messagebox.showerror("Eroare", f"Nu s-a putut citi folderul:\n{exc}")
            return
        if not matched:
            messagebox.showinfo("Niciun fișier găsit", f"Nu s-au găsit documente Word în:\n{folder}" + (f"\ncu filtrul «{flt}»." if flt else "."))
            return
        # Natural sort
        def natural_sort_key(path_str):
            name = os.path.basename(path_str)
            return [int(tok) if tok.isdigit() else tok.lower() for tok in re.split(r'(\d+)', name)]
        matched.sort(key=natural_sort_key)
        self._imports_load_word_files(matched, append=False)

    def _on_imports_clear_files(self):
        self.imports_selected_files.clear()
        self.imports_parsed_cache.clear()
        self._imports_refresh_preview_tables()
        self.lbl_imports_summary.configure(text="Lista a fost golită. Selectează unul sau mai multe documente Word (.doc / .docx).")
        self.lbl_imports_status.configure(text="Lista de fișiere a fost resetată.")

    def _imports_load_word_files(self, file_paths, append=False):
        if not append:
            self.imports_selected_files.clear()
            self.imports_parsed_cache.clear()

        def natural_sort_key(path_str):
            name = os.path.basename(path_str)
            return [int(tok) if tok.isdigit() else tok.lower() for tok in re.split(r'(\d+)', name)]

        existing = set(os.path.normcase(os.path.abspath(p)) for p in self.imports_selected_files)
        for p in file_paths:
            norm = os.path.normcase(os.path.abspath(p))
            if norm not in existing:
                self.imports_selected_files.append(p)
                existing.add(norm)

        self.imports_selected_files.sort(key=natural_sort_key)

        for p in self.imports_selected_files:
            if p not in self.imports_parsed_cache:
                try:
                    self.imports_parsed_cache[p] = metalix_importer.parse_word_document(p)
                except Exception as exc:
                    self.imports_parsed_cache[p] = {
                        "file_path": p, "file_name": os.path.basename(p),
                        "metadata": {}, "parts": [], "error": str(exc)
                    }

        self._imports_refresh_preview_tables()

    def _imports_refresh_preview_tables(self):
        for item in self.tree_imports_docs.get_children():
            self.tree_imports_docs.delete(item)
        for item in self.tree_imports_parts.get_children():
            self.tree_imports_parts.delete(item)

        if not self.imports_selected_files:
            return

        merge_dups = self.imports_merge_dups.get()
        total_rows = 0
        total_qty = 0

        if len(self.imports_selected_files) > 1:
            for p in self.imports_selected_files:
                d = self.imports_parsed_cache.get(p, {})
                pts = d.get("parts", [])
                if merge_dups:
                    pts = metalix_importer.aggregate_parts(pts)
                total_rows += len(pts)
                total_qty += sum(x["quantity"] for x in pts)

            self.tree_imports_docs.insert(
                "", "end", iid="__ALL__",
                values=(f"📚 TOATE ({len(self.imports_selected_files)} fișiere)", total_rows, total_qty)
            )

        total_rows_check = 0
        total_qty_check = 0
        for p in self.imports_selected_files:
            d = self.imports_parsed_cache.get(p, {})
            pts = d.get("parts", [])
            if merge_dups:
                pts = metalix_importer.aggregate_parts(pts)
            r_cnt = len(pts)
            q_cnt = sum(x["quantity"] for x in pts)
            total_rows_check += r_cnt
            total_qty_check += q_cnt
            self.tree_imports_docs.insert("", "end", iid=p, values=(os.path.basename(p), r_cnt, q_cnt))

        self.lbl_imports_summary.configure(
            text=f"✅ {len(self.imports_selected_files)} document(e) Word încărcate  |  Total poziții: {total_rows_check}  |  Total bucăți: {total_qty_check}",
            text_color=_GREEN[1]
        )

        first_iid = "__ALL__" if len(self.imports_selected_files) > 1 else self.imports_selected_files[0]
        self.tree_imports_docs.selection_set(first_iid)
        self._imports_populate_parts_tree(first_iid)

    def _on_imports_doc_tree_select(self, _event=None):
        sel = self.tree_imports_docs.selection()
        if sel:
            self._imports_populate_parts_tree(sel[0])

    def _imports_populate_parts_tree(self, selected_iid):
        for item in self.tree_imports_parts.get_children():
            self.tree_imports_parts.delete(item)

        merge_dups = self.imports_merge_dups.get()
        if selected_iid == "__ALL__":
            target_files = self.imports_selected_files
            self.lbl_imports_parts_title.configure(text=f"🔩 Piese extrase din TOATE cele {len(target_files)} documente Word:")
        else:
            target_files = [selected_iid]
            self.lbl_imports_parts_title.configure(text=f"🔩 Piese extrase din: {os.path.basename(selected_iid)}")

        for fpath in target_files:
            d = self.imports_parsed_cache.get(fpath, {})
            fname = os.path.basename(fpath)
            pts = d.get("parts", [])
            if merge_dups:
                pts = metalix_importer.aggregate_parts(pts)
            for p in pts:
                dims = f"{p.get('size_x', '')} x {p.get('size_y', '')}".strip(" x")
                self.tree_imports_parts.insert(
                    "", "end",
                    values=(fname, p.get("no", ""), p.get("raw_name", ""), p.get("clean_name", ""), p.get("quantity", ""), dims)
                )

    def _on_imports_mode_change(self):
        mode = self.imports_export_mode.get()
        if hasattr(self, "entry_imports_cons_name"):
            if mode == "individual":
                self.entry_imports_cons_name.configure(state="disabled")
            else:
                self.entry_imports_cons_name.configure(state="normal")
        self.save_settings()

    def _on_imports_pick_output_dir(self):
        initial = self.imports_output_dir.get().strip() or os.path.join(os.path.expanduser("~"), "Desktop")
        folder = filedialog.askdirectory(title="Selectează folderul destinație pentru Excel", initialdir=initial)
        if folder:
            norm = os.path.normpath(folder)
            self.imports_output_dir.set(norm)
            self.save_settings()

    def _on_imports_copy_to_clipboard(self):
        if not self.imports_selected_files:
            messagebox.showwarning("Atenție", "Selectează mai întâi cel puțin un document Word.")
            return

        sel = self.tree_imports_docs.selection()
        selected_iid = sel[0] if sel else ("__ALL__" if len(self.imports_selected_files) > 1 else self.imports_selected_files[0])
        target_files = self.imports_selected_files if selected_iid == "__ALL__" else [selected_iid]

        merge_dups = self.imports_merge_dups.get()
        all_parts = []
        for fpath in target_files:
            d = self.imports_parsed_cache.get(fpath, {})
            all_parts.extend(d.get("parts", []))

        if merge_dups:
            all_parts = metalix_importer.aggregate_parts(all_parts)

        if not all_parts:
            messagebox.showwarning("Atenție", "Nu există piese de copiat.")
            return

        lines = [f"{p['clean_name']}\t{p['quantity']}" for p in all_parts]
        tsv_text = "\n".join(lines)
        self.clipboard_clear()
        self.clipboard_append(tsv_text)
        self.update()

        self.lbl_imports_status.configure(text=f"📋 S-au copiat în Clipboard {len(all_parts)} rânduri (Col A + Col B). Poți da Paste (Ctrl+V) direct în Excel!")
        messagebox.showinfo("Copiat în Clipboard", f"Au fost copiate {len(all_parts)} poziții!\n\nAcum poți da direct Paste (Ctrl+V) în Excel.")

    def _get_imports_start_row_settings(self):
        mode = self.imports_start_row_mode.get()
        if mode == "row2_empty1":
            return 2, False
        elif mode == "row1":
            return 1, False
        else:  # header_row1
            return 2, True

    def _on_imports_quick_save_desktop(self):
        if not self.imports_selected_files:
            messagebox.showwarning("Atenție", "Selectează mai întâi unul sau mai multe documente Word.")
            return

        sel = self.tree_imports_docs.selection()
        selected_iid = sel[0] if sel else ("__ALL__" if len(self.imports_selected_files) > 1 else self.imports_selected_files[0])
        target_files = self.imports_selected_files if selected_iid == "__ALL__" else [selected_iid]

        merge_dups = self.imports_merge_dups.get()
        all_parts = []
        for fpath in target_files:
            d = self.imports_parsed_cache.get(fpath, {})
            all_parts.extend(d.get("parts", []))
        if merge_dups:
            all_parts = metalix_importer.aggregate_parts(all_parts)

        if not all_parts:
            messagebox.showwarning("Atenție", "Nu s-au găsit piese în documentele selectate.")
            return

        desk = os.path.join(os.path.expanduser("~"), "Desktop")
        import_path = os.path.join(desk, "IMPORT.xlsx")
        start_row, include_header = self._get_imports_start_row_settings()

        try:
            metalix_importer.write_import_excel(
                parts=all_parts, output_path=import_path,
                start_row=start_row, include_header=include_header
            )
            self.imports_last_generated_excel = import_path
            self.imports_last_output_dir = desk
            self.btn_imports_open_xl.configure(state="normal")
            self.lbl_imports_status.configure(text=f"✅ Salvat cu succes în {import_path} ({len(all_parts)} poziții, {sum(p['quantity'] for p in all_parts)} bucăți).")
            messagebox.showinfo("IMPORT.xlsx Salvat", f"Fișierul a fost salvat pe Desktop:\n{import_path}\n\n• Poziții: {len(all_parts)}\n• Total bucăți: {sum(p['quantity'] for p in all_parts)}")
        except PermissionError:
            messagebox.showerror("Fișier deschis", f"Nu se poate salva în:\n{import_path}\n\nÎnchide fișierul IMPORT.xlsx din Excel și încearcă din nou!")
        except Exception as exc:
            messagebox.showerror("Eroare la salvare", f"A apărut o eroare:\n{exc}")

    def _on_imports_generate_excel(self):
        if self.imports_running:
            return
        if not self.imports_selected_files:
            messagebox.showwarning("Atenție", "Selectează mai întâi cel puțin un document Word (.doc / .docx)!")
            return

        output_dir = self.imports_output_dir.get().strip() or os.path.join(os.path.expanduser("~"), "Desktop")
        mode = self.imports_export_mode.get()
        consolidated_name = self.imports_consolidated_name.get().strip() or "IMPORT.xlsx"
        start_row, include_header = self._get_imports_start_row_settings()
        merge_dups = self.imports_merge_dups.get()

        self.imports_running = True
        self.btn_imports_generate.configure(state="disabled", text="⏳ Se generează...")
        self.imports_progress.set(0)

        def worker():
            def on_progress(curr, total, msg):
                self.after(0, lambda: self._imports_update_progress(curr, total, msg))

            report = metalix_importer.process_batch_documents(
                word_files=list(self.imports_selected_files),
                output_dir=output_dir,
                mode=mode,
                consolidated_filename=consolidated_name,
                start_row=start_row,
                include_header=include_header,
                merge_duplicates=merge_dups,
                progress_callback=on_progress
            )
            self.after(0, lambda: self._on_imports_generation_complete(report, output_dir))

        threading.Thread(target=worker, daemon=True).start()

    def _imports_update_progress(self, curr, total, msg):
        frac = curr / max(total, 1)
        self.imports_progress.set(frac)
        self.lbl_imports_status.configure(text=msg)

    def _on_imports_generation_complete(self, report, output_dir):
        self.imports_running = False
        self.btn_imports_generate.configure(state="normal", text="🚀 GENEREAZĂ EXCEL (.XLSX)")
        self.imports_progress.set(1.0)

        gen_files = report.get("generated_files", [])
        errors = report.get("errors", [])
        self.imports_last_output_dir = output_dir

        if gen_files:
            self.imports_last_generated_excel = gen_files[-1]
            self.btn_imports_open_xl.configure(state="normal")

        summary_msg = f"✅ Generare finalizată! Documente procesate: {report.get('processed_docs', len(gen_files))}/{report.get('total_docs', len(self.imports_selected_files))} | Fișiere create: {len(gen_files)}"
        self.lbl_imports_status.configure(text=summary_msg)

        if errors:
            err_preview = "\n".join(errors[:6])
            messagebox.showwarning("Finalizat cu atenționări", f"{summary_msg}\n\nFolder:\n{output_dir}\n\nAtenționări:\n{err_preview}")
        else:
            messagebox.showinfo(
                "Generare Reușită",
                f"Toate fișierele Excel au fost generate cu succes!\n\n"
                f"• Fișiere Excel generate: {len(gen_files)}\n"
                f"• Salvare în:\n{output_dir}"
            )

    def _on_imports_open_last_excel(self):
        if self.imports_last_generated_excel and os.path.exists(self.imports_last_generated_excel):
            os.startfile(self.imports_last_generated_excel)
        else:
            messagebox.showwarning("Fișier negăsit", "Nu există încă un fișier Excel generat.")

    def _on_imports_open_output_folder(self):
        target = self.imports_last_output_dir or self.imports_output_dir.get().strip() or os.path.join(os.path.expanduser("~"), "Desktop")
        if os.path.exists(target):
            os.startfile(target)
        else:
            messagebox.showwarning("Folder inexistent", f"Folderul nu există încă:\n{target}")


if __name__ == "__main__":
    import sys
    # On Windows, tell the OS this is a distinct app so the taskbar icon updates correctly
    if sys.platform.startswith("win"):
        try:
            import ctypes
            myappid = f"steelsync.toolkit.app.{APP_VERSION}"
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

            # Prevent multiple instances from running simultaneously
            _single_instance_mutex = ctypes.windll.kernel32.CreateMutexW(None, False, "SteelSync_Toolkit_SingleInstance_Mutex")
            if ctypes.windll.kernel32.GetLastError() == 183:  # ERROR_ALREADY_EXISTS
                sys.exit(0)
        except Exception:
            pass

    app = UnifiedApp()
    app.mainloop()
