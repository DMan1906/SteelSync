"""
Aplicație Desktop Modernă: Mini Extragere & Organizare NC, PDF & DWG
Include 2 module principale:
1. Extragere pe Piese (BOM Excel -> Foldere pe Piese cu NC, PDF-uri de ansamblu și DWG-uri pe grosimi)
2. Regrupare DWG pe Grosimi (Piese -> Foldere Unificate cu Grosimi de Tablă)
"""

import os
import sys
import threading
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
from openpyxl.utils import get_column_letter

from core import (
    detect_excel_structure,
    parse_excel_parts,
    scan_source_files,
    run_pipeline,
    sanitize_folder_name,
    get_column_letter,
    scan_dwgs_for_regrouping,
    run_regroup_pipeline
)

# App Appearance configuration
ctk.set_appearance_mode("Dark")  # "Dark", "Light", or "System"
ctk.set_default_color_theme("blue")  # "blue", "green", "dark-blue"


class MiniExtragereNCApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Mini Extragere & Organizare NC, PDF & DWG")
        self.geometry("1020, 880")
        self.minsize(860, 700)

        # Try to set window icon
        icon_path = os.path.join(os.path.dirname(__file__), "app_icon.ico")
        if os.path.exists(icon_path):
            try:
                self.iconbitmap(icon_path)
            except Exception:
                pass

        # Pipeline 1 state (Extract)
        self._is_running = False
        self._cancel_requested = False
        self.last_report_path = None
        self.last_dest_dir = None
        self.excel_metadata = None

        # Pipeline 2 state (Regroup)
        self._regroup_running = False
        self._regroup_cancel_requested = False
        self.last_regroup_report = None
        self.last_regroup_dest = None

        self._build_ui()

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Tabview container
        self.tabview = ctk.CTkTabview(self, corner_radius=10)
        self.tabview.grid(row=0, column=0, sticky="nsew", padx=12, pady=10)

        self.tab_extract = self.tabview.add("📦 1. Extragere pe Piese (BOM -> Piese)")
        self.tab_regroup = self.tabview.add("🔄 2. Regrupare DWG pe Grosimi (Piese -> Grosimi)")

        self._build_extract_tab(self.tab_extract)
        self._build_regroup_tab(self.tab_regroup)

    # ═════════════════════════════════════════════════════════════════════════
    # 📦 TAB 1: EXTRAGERE PE PIESE (BOM -> PIESE)
    # ═════════════════════════════════════════════════════════════════════════
    def _build_extract_tab(self, parent):
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=1)

        main_scroll = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        main_scroll.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        main_scroll.grid_columnconfigure(0, weight=1)

        # ─── HEADER ───
        header_frame = ctk.CTkFrame(main_scroll, corner_radius=10, fg_color=("gray85", "#1e222d"))
        header_frame.pack(fill="x", pady=(0, 10), padx=2)

        title_lbl = ctk.CTkLabel(
            header_frame,
            text="📁 ORGANIZARE AUTOMATĂ PE PIESE (NC, PDF & DWG)",
            font=ctk.CTkFont(size=18, weight="bold")
        )
        title_lbl.pack(anchor="w", padx=14, pady=(10, 2))

        sub_lbl = ctk.CTkLabel(
            header_frame,
            text="Citește tabelul Excel BOM și creează câte un folder pentru fiecare piesă: mută fișierul NC, desenele de ansamblu (PDF) și piesele DWG pe subfoldere de grosimi de tablă.",
            font=ctk.CTkFont(size=12),
            text_color=("gray40", "gray70"),
            wraplength=920,
            justify="left"
        )
        sub_lbl.pack(anchor="w", padx=14, pady=(0, 10))

        # ─── CARD 1: FISIER EXCEL ───
        card_excel = ctk.CTkFrame(main_scroll, corner_radius=10)
        card_excel.pack(fill="x", pady=6, padx=2)
        card_excel.grid_columnconfigure(1, weight=1)

        lbl_c1 = ctk.CTkLabel(
            card_excel,
            text="1. Tabel Excel (Ansambluri & Piese)",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        lbl_c1.grid(row=0, column=0, columnspan=3, sticky="w", padx=14, pady=(10, 6))

        ctk.CTkLabel(card_excel, text="Fișier Excel:").grid(row=1, column=0, sticky="w", padx=(14, 8), pady=4)
        self.entry_excel = ctk.CTkEntry(card_excel, placeholder_text="Selectează fișierul Excel (.xlsx, .xls)...")
        self.entry_excel.grid(row=1, column=1, sticky="ew", padx=4, pady=4)
        btn_browse_excel = ctk.CTkButton(card_excel, text="Răsfoiește...", width=110, command=self._browse_excel)
        btn_browse_excel.grid(row=1, column=2, padx=(4, 14), pady=4)

        # Sheet & Columns selection row
        cols_frame = ctk.CTkFrame(card_excel, fg_color="transparent")
        cols_frame.grid(row=2, column=0, columnspan=3, sticky="ew", padx=10, pady=(4, 8))
        cols_frame.grid_columnconfigure((1, 3, 5), weight=1)

        ctk.CTkLabel(cols_frame, text="Foaie (Sheet):").grid(row=0, column=0, sticky="w", padx=4, pady=4)
        self.combo_sheet = ctk.CTkComboBox(cols_frame, values=["-"], command=self._on_sheet_change)
        self.combo_sheet.grid(row=0, column=1, sticky="ew", padx=4, pady=4)

        ctk.CTkLabel(cols_frame, text="Col. Ansamblu:").grid(row=0, column=2, sticky="w", padx=(12, 4), pady=4)
        self.combo_asmb_col = ctk.CTkComboBox(cols_frame, values=["-"], command=self._on_column_change)
        self.combo_asmb_col.grid(row=0, column=3, sticky="ew", padx=4, pady=4)

        ctk.CTkLabel(cols_frame, text="Col. Piesă:").grid(row=0, column=4, sticky="w", padx=(12, 4), pady=4)
        self.combo_part_col = ctk.CTkComboBox(cols_frame, values=["-"], command=self._on_column_change)
        self.combo_part_col.grid(row=0, column=5, sticky="ew", padx=4, pady=4)

        self.lbl_excel_info = ctk.CTkLabel(
            card_excel,
            text="ℹ️ Încărcați un fișier Excel pentru a începe.",
            font=ctk.CTkFont(size=11, slant="italic"),
            text_color=("gray40", "gray65")
        )
        self.lbl_excel_info.grid(row=3, column=0, columnspan=3, sticky="w", padx=14, pady=(0, 10))

        # ─── CARD 2: DIRECTOARE (SURSA & DESTINATIE) ───
        card_dirs = ctk.CTkFrame(main_scroll, corner_radius=10)
        card_dirs.pack(fill="x", pady=6, padx=2)
        card_dirs.grid_columnconfigure(1, weight=1)

        lbl_c2 = ctk.CTkLabel(
            card_dirs,
            text="2. Foldere de Lucru (Sursă & Destinație)",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        lbl_c2.grid(row=0, column=0, columnspan=3, sticky="w", padx=14, pady=(10, 6))

        ctk.CTkLabel(card_dirs, text="Folder Sursă:").grid(row=1, column=0, sticky="w", padx=(14, 8), pady=4)
        self.entry_source = ctk.CTkEntry(
            card_dirs,
            placeholder_text="Folderul ce conține PDF-uri, NC-uri și subfoldere cu grosimi..."
        )
        self.entry_source.grid(row=1, column=1, sticky="ew", padx=4, pady=4)
        btn_browse_src = ctk.CTkButton(card_dirs, text="Răsfoiește...", width=110, command=self._browse_source)
        btn_browse_src.grid(row=1, column=2, padx=(4, 14), pady=4)

        self.lbl_source_info = ctk.CTkLabel(
            card_dirs,
            text="",
            font=ctk.CTkFont(size=11),
            text_color=("gray40", "#64B5F6")
        )
        self.lbl_source_info.grid(row=2, column=1, sticky="w", padx=4, pady=(0, 4))

        ctk.CTkLabel(card_dirs, text="Folder Destinație:").grid(row=3, column=0, sticky="w", padx=(14, 8), pady=4)
        self.entry_dest = ctk.CTkEntry(
            card_dirs,
            placeholder_text="Folderul unde se vor crea folderele fiecărei piese..."
        )
        self.entry_dest.grid(row=3, column=1, sticky="ew", padx=4, pady=4)
        btn_browse_dst = ctk.CTkButton(card_dirs, text="Răsfoiește...", width=110, command=self._browse_dest)
        btn_browse_dst.grid(row=3, column=2, padx=(4, 14), pady=4)

        # ─── CARD 3: OPTIUNI SI SETARI ───
        card_opts = ctk.CTkFrame(main_scroll, corner_radius=10)
        card_opts.pack(fill="x", pady=6, padx=2)

        lbl_c3 = ctk.CTkLabel(
            card_opts,
            text="3. Mod Operare și Opțiuni",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        lbl_c3.pack(anchor="w", padx=14, pady=(10, 6))

        opts_subframe = ctk.CTkFrame(card_opts, fg_color="transparent")
        opts_subframe.pack(fill="x", padx=14, pady=(0, 10))

        self.var_mode = tk.StringVar(value="move")
        r_move = ctk.CTkRadioButton(
            opts_subframe,
            text="Mută fișierele (Move)",
            variable=self.var_mode,
            value="move"
        )
        r_move.grid(row=0, column=0, padx=(0, 20), pady=4, sticky="w")

        r_copy = ctk.CTkRadioButton(
            opts_subframe,
            text="Copiază fișierele (Copy - păstrează originalele)",
            variable=self.var_mode,
            value="copy"
        )
        r_copy.grid(row=0, column=1, padx=(0, 20), pady=4, sticky="w")

        self.var_variations = tk.BooleanVar(value=True)
        chk_var = ctk.CTkCheckBox(
            opts_subframe,
            text="Include variații DWG (ex. T, L, C, etc.)",
            variable=self.var_variations
        )
        chk_var.grid(row=1, column=0, padx=(0, 20), pady=6, sticky="w")

        self.var_clean_empty = tk.BooleanVar(value=True)
        chk_clean = ctk.CTkCheckBox(
            opts_subframe,
            text="Curăță folderele goale din sursă la final (în modul Mutare)",
            variable=self.var_clean_empty
        )
        chk_clean.grid(row=1, column=1, padx=(0, 20), pady=6, sticky="w")

        self.var_open_dest = tk.BooleanVar(value=True)
        chk_open = ctk.CTkCheckBox(
            opts_subframe,
            text="Deschide automat folderul destinație la finalizare",
            variable=self.var_open_dest
        )
        chk_open.grid(row=2, column=0, padx=(0, 20), pady=4, sticky="w")

        # ─── ACTION BUTTONS & PROGRESS ───
        action_frame = ctk.CTkFrame(main_scroll, fg_color="transparent")
        action_frame.pack(fill="x", pady=8, padx=2)
        action_frame.grid_columnconfigure((0, 1), weight=1)

        self.btn_run = ctk.CTkButton(
            action_frame,
            text="🚀  ÎNCEPE ORGANIZAREA / MUTAREA",
            font=ctk.CTkFont(size=15, weight="bold"),
            fg_color="#1f7a1f",
            hover_color="#186318",
            height=44,
            command=self._start_pipeline
        )
        self.btn_run.grid(row=0, column=0, padx=6, pady=4, sticky="ew")

        self.btn_cancel = ctk.CTkButton(
            action_frame,
            text="🛑  Oprește / Anulează",
            font=ctk.CTkFont(size=14),
            fg_color="#8B0000",
            hover_color="#5C0000",
            height=44,
            state="disabled",
            command=self._cancel_pipeline
        )
        self.btn_cancel.grid(row=0, column=1, padx=6, pady=4, sticky="ew")

        self.progress_bar = ctk.CTkProgressBar(main_scroll)
        self.progress_bar.pack(fill="x", padx=6, pady=(6, 2))
        self.progress_bar.set(0)

        self.lbl_progress_status = ctk.CTkLabel(
            main_scroll,
            text="Așteptare pornire...",
            font=ctk.CTkFont(size=12),
            text_color=("gray40", "gray70")
        )
        self.lbl_progress_status.pack(anchor="w", padx=8, pady=(0, 6))

        # ─── CARD 4: JURNAL ───
        card_log = ctk.CTkFrame(main_scroll, corner_radius=10)
        card_log.pack(fill="both", expand=True, pady=6, padx=2)

        log_head_frame = ctk.CTkFrame(card_log, fg_color="transparent")
        log_head_frame.pack(fill="x", padx=12, pady=(8, 4))

        ctk.CTkLabel(
            log_head_frame,
            text="Jurnal Activitate și Rezultate:",
            font=ctk.CTkFont(size=13, weight="bold")
        ).pack(side="left")

        btn_clear_log = ctk.CTkButton(
            log_head_frame,
            text="🧹 Șterge Jurnal",
            width=90,
            height=26,
            command=self._clear_log
        )
        btn_clear_log.pack(side="right")

        self.log_box = ctk.CTkTextbox(
            card_log,
            height=200,
            font=ctk.CTkFont(family="Consolas", size=11),
            activate_scrollbars=True
        )
        self.log_box.pack(fill="both", expand=True, padx=12, pady=(4, 8))

        try:
            self.log_box._textbox.tag_config("SUCCESS", foreground="#4CAF50")
            self.log_box._textbox.tag_config("WARNING", foreground="#FF9800")
            self.log_box._textbox.tag_config("ERROR", foreground="#F44336")
            self.log_box._textbox.tag_config("INFO", foreground="#E0E0E0")
        except Exception:
            pass

        post_frame = ctk.CTkFrame(main_scroll, fg_color="transparent")
        post_frame.pack(fill="x", pady=(4, 10), padx=2)

        self.btn_open_dest = ctk.CTkButton(
            post_frame,
            text="📁 Deschide Folder Destinație",
            width=200,
            state="disabled",
            command=self._open_dest_folder
        )
        self.btn_open_dest.pack(side="left", padx=6)

        self.btn_open_report = ctk.CTkButton(
            post_frame,
            text="📊 Deschide Raport Excel",
            width=200,
            state="disabled",
            command=self._open_excel_report
        )
        self.btn_open_report.pack(side="left", padx=6)

    # ═════════════════════════════════════════════════════════════════════════
    # 🔄 TAB 2: REGRUPARE INVERSĂ (PIESE -> FOLDERE GROSIMI DWG)
    # ═════════════════════════════════════════════════════════════════════════
    def _build_regroup_tab(self, parent):
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=1)

        regroup_scroll = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        regroup_scroll.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        regroup_scroll.grid_columnconfigure(0, weight=1)

        # ─── HEADER ───
        header_frame = ctk.CTkFrame(regroup_scroll, corner_radius=10, fg_color=("gray85", "#1e222d"))
        header_frame.pack(fill="x", pady=(0, 10), padx=2)

        title_lbl = ctk.CTkLabel(
            header_frame,
            text="🔄 REGRUPARE INVERSĂ: DOAR DWG-URILE PE GROSIMI DE TABLĂ",
            font=ctk.CTkFont(size=18, weight="bold")
        )
        title_lbl.pack(anchor="w", padx=14, pady=(10, 2))

        sub_lbl = ctk.CTkLabel(
            header_frame,
            text="Selectați folderul ce conține folderele de piese (ex. Piese_Organizate). Programul va colecta DOAR fișierele DWG din toate piesele și le va regrupa în foldere dedicate fiecărei grosimi (ex: T10, T12, T15 etc.), punând toate piesele de aceeași grosime în același folder.",
            font=ctk.CTkFont(size=12),
            text_color=("gray40", "gray70"),
            wraplength=920,
            justify="left"
        )
        sub_lbl.pack(anchor="w", padx=14, pady=(0, 10))

        # ─── CARD 1: SELECTIE FOLDERE ───
        card_dirs = ctk.CTkFrame(regroup_scroll, corner_radius=10)
        card_dirs.pack(fill="x", pady=6, padx=2)
        card_dirs.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            card_dirs,
            text="1. Foldere de Lucru (Sursă Piese & Destinație Grosimi)",
            font=ctk.CTkFont(size=14, weight="bold")
        ).grid(row=0, column=0, columnspan=3, sticky="w", padx=14, pady=(10, 6))

        # Regroup Source
        ctk.CTkLabel(card_dirs, text="Folder cu Piese (Sursă):").grid(row=1, column=0, sticky="w", padx=(14, 8), pady=4)
        self.entry_regroup_src = ctk.CTkEntry(
            card_dirs,
            placeholder_text="Selectați folderul ce conține folderele de piese (ex: Piese_Organizate)..."
        )
        self.entry_regroup_src.grid(row=1, column=1, sticky="ew", padx=4, pady=4)
        btn_browse_rg_src = ctk.CTkButton(card_dirs, text="Răsfoiește...", width=110, command=self._browse_regroup_source)
        btn_browse_rg_src.grid(row=1, column=2, padx=(4, 14), pady=4)

        self.lbl_regroup_src_info = ctk.CTkLabel(
            card_dirs,
            text="ℹ️ Selectați folderul cu piese pentru scanare automată.",
            font=ctk.CTkFont(size=11),
            text_color=("gray40", "#64B5F6"),
            wraplength=880,
            justify="left"
        )
        self.lbl_regroup_src_info.grid(row=2, column=1, sticky="w", padx=4, pady=(0, 4))

        # Regroup Dest
        ctk.CTkLabel(card_dirs, text="Folder Destinație (Grosimi):").grid(row=3, column=0, sticky="w", padx=(14, 8), pady=4)
        self.entry_regroup_dst = ctk.CTkEntry(
            card_dirs,
            placeholder_text="Folderul unde se vor crea folderele cu grosimi (T10, T12, etc.)..."
        )
        self.entry_regroup_dst.grid(row=3, column=1, sticky="ew", padx=4, pady=4)
        btn_browse_rg_dst = ctk.CTkButton(card_dirs, text="Răsfoiește...", width=110, command=self._browse_regroup_dest)
        btn_browse_rg_dst.grid(row=3, column=2, padx=(4, 14), pady=4)

        # ─── CARD 2: OPTIUNI REGRUPARE ───
        card_opts = ctk.CTkFrame(regroup_scroll, corner_radius=10)
        card_opts.pack(fill="x", pady=6, padx=2)

        ctk.CTkLabel(
            card_opts,
            text="2. Opțiuni Regrupare DWG",
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(anchor="w", padx=14, pady=(10, 6))

        opts_subframe = ctk.CTkFrame(card_opts, fg_color="transparent")
        opts_subframe.pack(fill="x", padx=14, pady=(0, 10))

        self.var_regroup_mode = tk.StringVar(value="copy")
        r_copy = ctk.CTkRadioButton(
            opts_subframe,
            text="Copiază fișierele DWG (Copy - păstrează originalele în folderele de piese)",
            variable=self.var_regroup_mode,
            value="copy"
        )
        r_copy.grid(row=0, column=0, padx=(0, 20), pady=4, sticky="w")

        r_move = ctk.CTkRadioButton(
            opts_subframe,
            text="Mută fișierele DWG (Move - le scoate din folderele de piese)",
            variable=self.var_regroup_mode,
            value="move"
        )
        r_move.grid(row=0, column=1, padx=(0, 20), pady=4, sticky="w")

        self.var_regroup_normalize = tk.BooleanVar(value=True)
        chk_norm = ctk.CTkCheckBox(
            opts_subframe,
            text="Uniformizează denumirile grosimilor (ex: t10 devine T10)",
            variable=self.var_regroup_normalize
        )
        chk_norm.grid(row=1, column=0, padx=(0, 20), pady=6, sticky="w")

        self.var_regroup_clean_empty = tk.BooleanVar(value=False)
        chk_clean = ctk.CTkCheckBox(
            opts_subframe,
            text="Curăță folderele goale din sursă la final (în modul Mutare)",
            variable=self.var_regroup_clean_empty
        )
        chk_clean.grid(row=1, column=1, padx=(0, 20), pady=6, sticky="w")

        self.var_regroup_open_dest = tk.BooleanVar(value=True)
        chk_open = ctk.CTkCheckBox(
            opts_subframe,
            text="Deschide automat folderul destinație la finalizare",
            variable=self.var_regroup_open_dest
        )
        chk_open.grid(row=2, column=0, padx=(0, 20), pady=4, sticky="w")

        # ─── ACTION BUTTONS & PROGRESS ───
        action_frame = ctk.CTkFrame(regroup_scroll, fg_color="transparent")
        action_frame.pack(fill="x", pady=8, padx=2)
        action_frame.grid_columnconfigure((0, 1), weight=1)

        self.btn_regroup_run = ctk.CTkButton(
            action_frame,
            text="🚀  REGRUPEAZĂ DWG-URILE PE GROSIMI",
            font=ctk.CTkFont(size=15, weight="bold"),
            fg_color="#0D47A1",
            hover_color="#0A3880",
            height=44,
            command=self._start_regroup_pipeline
        )
        self.btn_regroup_run.grid(row=0, column=0, padx=6, pady=4, sticky="ew")

        self.btn_regroup_cancel = ctk.CTkButton(
            action_frame,
            text="🛑  Oprește / Anulează",
            font=ctk.CTkFont(size=14),
            fg_color="#8B0000",
            hover_color="#5C0000",
            height=44,
            state="disabled",
            command=self._cancel_regroup_pipeline
        )
        self.btn_regroup_cancel.grid(row=0, column=1, padx=6, pady=4, sticky="ew")

        self.progress_bar_regroup = ctk.CTkProgressBar(regroup_scroll)
        self.progress_bar_regroup.pack(fill="x", padx=6, pady=(6, 2))
        self.progress_bar_regroup.set(0)

        self.lbl_regroup_status = ctk.CTkLabel(
            regroup_scroll,
            text="Așteptare pornire regrupare...",
            font=ctk.CTkFont(size=12),
            text_color=("gray40", "gray70")
        )
        self.lbl_regroup_status.pack(anchor="w", padx=8, pady=(0, 6))

        # ─── CARD 3: JURNAL REGRUPARE ───
        card_log = ctk.CTkFrame(regroup_scroll, corner_radius=10)
        card_log.pack(fill="both", expand=True, pady=6, padx=2)

        log_head_frame = ctk.CTkFrame(card_log, fg_color="transparent")
        log_head_frame.pack(fill="x", padx=12, pady=(8, 4))

        ctk.CTkLabel(
            log_head_frame,
            text="Jurnal Regrupare DWG:",
            font=ctk.CTkFont(size=13, weight="bold")
        ).pack(side="left")

        btn_clear_log = ctk.CTkButton(
            log_head_frame,
            text="🧹 Șterge Jurnal",
            width=90,
            height=26,
            command=self._clear_regroup_log
        )
        btn_clear_log.pack(side="right")

        self.log_box_regroup = ctk.CTkTextbox(
            card_log,
            height=200,
            font=ctk.CTkFont(family="Consolas", size=11),
            activate_scrollbars=True
        )
        self.log_box_regroup.pack(fill="both", expand=True, padx=12, pady=(4, 8))

        try:
            self.log_box_regroup._textbox.tag_config("SUCCESS", foreground="#4CAF50")
            self.log_box_regroup._textbox.tag_config("WARNING", foreground="#FF9800")
            self.log_box_regroup._textbox.tag_config("ERROR", foreground="#F44336")
            self.log_box_regroup._textbox.tag_config("INFO", foreground="#E0E0E0")
        except Exception:
            pass

        post_frame = ctk.CTkFrame(regroup_scroll, fg_color="transparent")
        post_frame.pack(fill="x", pady=(4, 10), padx=2)

        self.btn_open_regroup_dest = ctk.CTkButton(
            post_frame,
            text="📁 Deschide Folder Destinație",
            width=200,
            state="disabled",
            command=self._open_regroup_dest_folder
        )
        self.btn_open_regroup_dest.pack(side="left", padx=6)

        self.btn_open_regroup_report = ctk.CTkButton(
            post_frame,
            text="📊 Deschide Raport Excel",
            width=200,
            state="disabled",
            command=self._open_regroup_excel_report
        )
        self.btn_open_regroup_report.pack(side="left", padx=6)

    # ═════════════════════════════════════════════════════════════════════════
    # 📦 TAB 1 HANDLERS (EXTRACT)
    # ═════════════════════════════════════════════════════════════════════════
    def _browse_excel(self):
        f = filedialog.askopenfilename(
            title="Selectați fișierul Excel cu piese",
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")]
        )
        if f:
            self.entry_excel.delete(0, "end")
            self.entry_excel.insert(0, f)
            self._load_excel_structure(f)

    def _load_excel_structure(self, filepath):
        try:
            self.excel_metadata = detect_excel_structure(filepath)
            sheets = self.excel_metadata["sheets"]
            self.combo_sheet.configure(values=sheets)
            self.combo_sheet.set(self.excel_metadata["selected_sheet"])
            self._update_column_dropdowns()
        except Exception as e:
            messagebox.showerror("Eroare Excel", f"Nu s-a putut citi fișierul Excel:\n{e}")

    def _on_sheet_change(self, sheet_val):
        excel_path = self.entry_excel.get().strip()
        if excel_path and os.path.exists(excel_path):
            try:
                self.excel_metadata = detect_excel_structure(excel_path, sheet_name=sheet_val)
                self._update_column_dropdowns()
            except Exception as e:
                messagebox.showerror("Eroare Foaie", f"Eroare la schimbarea foii:\n{e}")

    def _on_column_change(self, val=None):
        self._update_preview()

    def _update_column_dropdowns(self):
        if not self.excel_metadata:
            return
        headers = self.excel_metadata.get("headers", [])
        if not headers:
            self.combo_asmb_col.configure(values=["-"])
            self.combo_part_col.configure(values=["-"])
            self.combo_asmb_col.set("-")
            self.combo_part_col.set("-")
            return

        col_labels = [f"Col. {get_column_letter(idx+1)}: {name}" for idx, name in headers]
        self.combo_asmb_col.configure(values=col_labels)
        self.combo_part_col.configure(values=col_labels)

        a_idx = self.excel_metadata["asmb_col_idx"]
        p_idx = self.excel_metadata["part_col_idx"]

        if a_idx is not None and a_idx < len(col_labels):
            self.combo_asmb_col.set(col_labels[a_idx])
        elif col_labels:
            self.combo_asmb_col.set(col_labels[0])

        if p_idx is not None and p_idx < len(col_labels):
            self.combo_part_col.set(col_labels[p_idx])
        elif len(col_labels) > 1:
            self.combo_part_col.set(col_labels[1])

        self._update_preview()

    def _update_preview(self):
        if not self.excel_metadata:
            return
        excel_path = self.entry_excel.get().strip()
        if not excel_path or not os.path.isfile(excel_path):
            return

        headers = self.excel_metadata.get("headers", [])
        a_choice = self.combo_asmb_col.get()
        p_choice = self.combo_part_col.get()

        a_idx = self.excel_metadata.get("asmb_col_idx", 0)
        p_idx = self.excel_metadata.get("part_col_idx", 1)

        for idx, name in headers:
            col_str = f"Col. {get_column_letter(idx+1)}: {name}"
            if col_str == a_choice:
                a_idx = idx
            if col_str == p_choice:
                p_idx = idx

        try:
            parts = parse_excel_parts(
                excel_path,
                self.combo_sheet.get(),
                self.excel_metadata.get("header_row", 1),
                a_idx,
                p_idx
            )
            asmb_count = sum(len(a) for a in parts.values())
            self.lbl_excel_info.configure(
                text=f"✔️ Detectate cu succes: {len(parts)} piese unice, {asmb_count} asocieri de ansambluri (Antet la rândul {self.excel_metadata.get('header_row', 1)}).",
                text_color="#81C784"
            )
        except Exception as e:
            self.lbl_excel_info.configure(
                text=f"ℹ️ Eroare parsare preview: {e}",
                text_color="#FFB74D"
            )

    def _browse_source(self):
        d = filedialog.askdirectory(title="Selectați folderul sursă")
        if d:
            self.entry_source.delete(0, "end")
            self.entry_source.insert(0, d)

            dest_sugg = os.path.join(d, "Piese_Organizate")
            if not self.entry_dest.get().strip():
                self.entry_dest.delete(0, "end")
                self.entry_dest.insert(0, dest_sugg)

            threading.Thread(target=self._quick_scan_source, args=(d,), daemon=True).start()

    def _quick_scan_source(self, src_dir):
        try:
            data = scan_source_files(src_dir)
            ncs = len(data["nc_files"])
            pdfs = len(data["pdf_files"])
            dwgs = len(data["dwg_files"])
            thks = data["thicknesses"]
            thk_str = f"în {len(thks)} grosimi ({', '.join(thks[:4])}{'...' if len(thks)>4 else ''})" if thks else "fără grosimi"
            info_text = f"✔️ Sursă: {ncs} fișiere NC, {pdfs} PDF-uri, {dwgs} DWG-uri {thk_str}"
            self.after(0, lambda: self.lbl_source_info.configure(text=info_text))
        except Exception:
            pass

    def _browse_dest(self):
        d = filedialog.askdirectory(title="Selectați folderul destinație")
        if d:
            self.entry_dest.delete(0, "end")
            self.entry_dest.insert(0, d)

    def _open_dest_folder(self):
        dest = self.last_dest_dir or self.entry_dest.get().strip()
        if dest and os.path.exists(dest):
            if sys.platform.startswith("win"):
                os.startfile(dest)
            else:
                subprocess.run(["xdg-open", dest])

    def _open_excel_report(self):
        if self.last_report_path and os.path.exists(self.last_report_path):
            if sys.platform.startswith("win"):
                os.startfile(self.last_report_path)
            else:
                subprocess.run(["xdg-open", self.last_report_path])
        else:
            messagebox.showinfo("Raport", "Raportul Excel nu a fost găsit sau procesul nu s-a finalizat încă.")

    def _clear_log(self):
        self.log_box.delete("1.0", "end")

    def _log(self, message, level="INFO"):
        def append():
            tag = level if level in ["SUCCESS", "WARNING", "ERROR"] else "INFO"
            try:
                self.log_box._textbox.insert("end", str(message) + "\n", tag)
            except Exception:
                self.log_box.insert("end", str(message) + "\n")
            self.log_box.see("end")
        self.after(0, append)

    def _set_progress(self, pct, text=""):
        def update():
            self.progress_bar.set(pct)
            if text:
                self.lbl_progress_status.configure(text=f"{int(pct * 100)}% - {text}")
        self.after(0, update)

    def _start_pipeline(self):
        if self._is_running:
            return

        excel_path = self.entry_excel.get().strip()
        source_dir = self.entry_source.get().strip()
        dest_dir = self.entry_dest.get().strip()

        if not excel_path or not os.path.isfile(excel_path):
            messagebox.showwarning("Atenție", "Vă rugăm să selectați un fișier Excel valid!")
            return
        if not source_dir or not os.path.isdir(source_dir):
            messagebox.showwarning("Atenție", "Vă rugăm să selectați un folder sursă valid!")
            return
        if not dest_dir:
            dest_dir = os.path.join(source_dir, "Piese_Organizate")
            self.entry_dest.insert(0, dest_dir)

        headers = self.excel_metadata.get("headers", []) if self.excel_metadata else []
        a_choice = self.combo_asmb_col.get()
        p_choice = self.combo_part_col.get()

        a_idx = 0
        p_idx = 1
        for idx, name in headers:
            col_str = f"Col. {get_column_letter(idx+1)}: {name}"
            if col_str == a_choice:
                a_idx = idx
            if col_str == p_choice:
                p_idx = idx

        sheet_name = self.combo_sheet.get()
        header_row = self.excel_metadata.get("header_row", 1) if self.excel_metadata else 1

        mode = self.var_mode.get()
        allow_variations = self.var_variations.get()
        clean_empty = self.var_clean_empty.get()

        self._is_running = True
        self._cancel_requested = False
        self.btn_run.configure(state="disabled")
        self.btn_cancel.configure(state="normal")
        self.btn_open_dest.configure(state="disabled")
        self.btn_open_report.configure(state="disabled")
        self._clear_log()

        threading.Thread(
            target=self._worker_thread,
            args=(excel_path, sheet_name, header_row, a_idx, p_idx, source_dir, dest_dir, mode, allow_variations, clean_empty),
            daemon=True
        ).start()

    def _cancel_pipeline(self):
        if self._is_running:
            self._cancel_requested = True
            self.btn_cancel.configure(state="disabled")
            self._log("⚠️ Se solicită oprirea procesului...", "WARNING")

    def _worker_thread(self, excel_path, sheet_name, header_row, a_idx, p_idx, source_dir, dest_dir, mode, allow_variations, clean_empty):
        summary_stats = None
        report_path = None
        error_msg = None

        try:
            summary_stats, results_list, report_path = run_pipeline(
                excel_path=excel_path,
                sheet_name=sheet_name,
                header_row=header_row,
                asmb_col_idx=a_idx,
                part_col_idx=p_idx,
                source_dir=source_dir,
                dest_dir=dest_dir,
                mode=mode,
                allow_variations=allow_variations,
                clean_empty_folders=clean_empty,
                progress_callback=self._set_progress,
                log_callback=self._log,
                is_cancelled=lambda: self._cancel_requested
            )
            self.last_dest_dir = dest_dir
            self.last_report_path = report_path
        except Exception as e:
            error_msg = str(e)
            self._log(f"\n❌ EROARE: {e}", "ERROR")

        self.after(0, lambda: self._on_pipeline_finished(summary_stats, report_path, error_msg, dest_dir))

    def _on_pipeline_finished(self, summary_stats, report_path, error_msg, dest_dir):
        self._is_running = False
        self.btn_run.configure(state="normal")
        self.btn_cancel.configure(state="disabled")

        if error_msg:
            messagebox.showerror("Eroare la procesare", f"A apărut o eroare:\n{error_msg}")
            return

        if self._cancel_requested:
            messagebox.showinfo("Oprit", "Procesul a fost oprit de utilizator.")
            return

        self.btn_open_dest.configure(state="normal")
        if report_path and os.path.exists(report_path):
            self.btn_open_report.configure(state="normal")

        if self.var_open_dest.get() and os.path.exists(dest_dir):
            try:
                if sys.platform.startswith("win"):
                    os.startfile(dest_dir)
            except Exception:
                pass

        total = summary_stats.get("total_parts", 0) if summary_stats else 0
        complete = summary_stats.get("complete_parts", 0) if summary_stats else 0
        partial = summary_stats.get("partial_parts", 0) if summary_stats else 0
        missing = summary_stats.get("missing_parts", 0) if summary_stats else 0

        messagebox.showinfo(
            "Finalizat cu Succes",
            f"Procesul de organizare s-a încheiat!\n\n"
            f"• Total piese: {total}\n"
            f"• Complete (NC + PDF + DWG): {complete}\n"
            f"• Parțiale: {partial}\n"
            f"• Fără fișiere găsite: {missing}\n\n"
            f"Folder destinație: {dest_dir}\n"
            f"Raport Excel: {os.path.basename(report_path) if report_path else '-'}"
        )

    # ═════════════════════════════════════════════════════════════════════════
    # 🔄 TAB 2 HANDLERS (REGROUP)
    # ═════════════════════════════════════════════════════════════════════════
    def _browse_regroup_source(self):
        d = filedialog.askdirectory(title="Selectați folderul cu piese (sursa pentru regrupare)")
        if d:
            self.entry_regroup_src.delete(0, "end")
            self.entry_regroup_src.insert(0, d)

            suggested_dst = d.rstrip("\\/") + "_DWG_Grosimi"
            if not self.entry_regroup_dst.get().strip():
                self.entry_regroup_dst.delete(0, "end")
                self.entry_regroup_dst.insert(0, suggested_dst)

            threading.Thread(target=self._quick_scan_regroup, args=(d,), daemon=True).start()

    def _quick_scan_regroup(self, src_dir):
        try:
            scan = scan_dwgs_for_regrouping(src_dir, normalize_case=self.var_regroup_normalize.get())
            total = scan["total_dwgs"]
            groups = scan["thickness_groups"]
            if total == 0:
                msg = "⚠️ Nu a fost găsit niciun fișier .dwg în folderul selectat."
                color = "#FFB74D"
            else:
                summary_parts = [f"{thk} ({len(items)})" for thk, items in sorted(groups.items())]
                msg = f"✔️ Detectate {total} fișiere DWG în {len(groups)} grosimi: {', '.join(summary_parts)}"
                color = "#81C784"
            self.after(0, lambda: self.lbl_regroup_src_info.configure(text=msg, text_color=color))
        except Exception as e:
            self.after(0, lambda: self.lbl_regroup_src_info.configure(text=f"Eroare scanare: {e}", text_color="#FFB74D"))

    def _browse_regroup_dest(self):
        d = filedialog.askdirectory(title="Selectați folderul destinație pentru grosimi")
        if d:
            self.entry_regroup_dst.delete(0, "end")
            self.entry_regroup_dst.insert(0, d)

    def _open_regroup_dest_folder(self):
        dest = self.last_regroup_dest or self.entry_regroup_dst.get().strip()
        if dest and os.path.exists(dest):
            if sys.platform.startswith("win"):
                os.startfile(dest)
            else:
                subprocess.run(["xdg-open", dest])

    def _open_regroup_excel_report(self):
        if self.last_regroup_report and os.path.exists(self.last_regroup_report):
            if sys.platform.startswith("win"):
                os.startfile(self.last_regroup_report)
            else:
                subprocess.run(["xdg-open", self.last_regroup_report])
        else:
            messagebox.showinfo("Raport", "Raportul Excel nu a fost găsit sau procesul nu s-a finalizat încă.")

    def _clear_regroup_log(self):
        self.log_box_regroup.delete("1.0", "end")

    def _log_regroup(self, message, level="INFO"):
        def append():
            tag = level if level in ["SUCCESS", "WARNING", "ERROR"] else "INFO"
            try:
                self.log_box_regroup._textbox.insert("end", str(message) + "\n", tag)
            except Exception:
                self.log_box_regroup.insert("end", str(message) + "\n")
            self.log_box_regroup.see("end")
        self.after(0, append)

    def _set_regroup_progress(self, pct, text=""):
        def update():
            self.progress_bar_regroup.set(pct)
            if text:
                self.lbl_regroup_status.configure(text=f"{int(pct * 100)}% - {text}")
        self.after(0, update)

    def _start_regroup_pipeline(self):
        if self._regroup_running:
            return

        source_dir = self.entry_regroup_src.get().strip()
        dest_dir = self.entry_regroup_dst.get().strip()

        if not source_dir or not os.path.isdir(source_dir):
            messagebox.showwarning("Atenție", "Vă rugăm să selectați un folder sursă valid!")
            return
        if not dest_dir:
            dest_dir = source_dir.rstrip("\\/") + "_DWG_Grosimi"
            self.entry_regroup_dst.insert(0, dest_dir)

        mode = self.var_regroup_mode.get()
        normalize_case = self.var_regroup_normalize.get()
        clean_empty = self.var_regroup_clean_empty.get()

        self._regroup_running = True
        self._regroup_cancel_requested = False
        self.btn_regroup_run.configure(state="disabled")
        self.btn_regroup_cancel.configure(state="normal")
        self.btn_open_regroup_dest.configure(state="disabled")
        self.btn_open_regroup_report.configure(state="disabled")
        self._clear_regroup_log()

        threading.Thread(
            target=self._worker_regroup_thread,
            args=(source_dir, dest_dir, mode, normalize_case, clean_empty),
            daemon=True
        ).start()

    def _cancel_regroup_pipeline(self):
        if self._regroup_running:
            self._regroup_cancel_requested = True
            self.btn_regroup_cancel.configure(state="disabled")
            self._log_regroup("⚠️ Se solicită oprirea regrupării...", "WARNING")

    def _worker_regroup_thread(self, source_dir, dest_dir, mode, normalize_case, clean_empty):
        summary_stats = None
        report_path = None
        error_msg = None

        try:
            summary_stats, report_path = run_regroup_pipeline(
                source_dir=source_dir,
                dest_dir=dest_dir,
                mode=mode,
                normalize_case=normalize_case,
                clean_empty_folders=clean_empty,
                progress_callback=self._set_regroup_progress,
                log_callback=self._log_regroup,
                is_cancelled=lambda: self._regroup_cancel_requested
            )
            self.last_regroup_dest = dest_dir
            self.last_regroup_report = report_path
        except Exception as e:
            error_msg = str(e)
            self._log_regroup(f"\n❌ EROARE: {e}", "ERROR")

        self.after(0, lambda: self._on_regroup_finished(summary_stats, report_path, error_msg, dest_dir))

    def _on_regroup_finished(self, summary_stats, report_path, error_msg, dest_dir):
        self._regroup_running = False
        self.btn_regroup_run.configure(state="normal")
        self.btn_regroup_cancel.configure(state="disabled")

        if error_msg:
            messagebox.showerror("Eroare la Regrupare", f"A apărut o eroare:\n{error_msg}")
            return

        if self._regroup_cancel_requested:
            messagebox.showinfo("Oprit", "Regruparea a fost oprită de utilizator.")
            return

        self.btn_open_regroup_dest.configure(state="normal")
        if report_path and os.path.exists(report_path):
            self.btn_open_regroup_report.configure(state="normal")

        if self.var_regroup_open_dest.get() and os.path.exists(dest_dir):
            try:
                if sys.platform.startswith("win"):
                    os.startfile(dest_dir)
            except Exception:
                pass

        total = summary_stats.get("total_dwgs", 0) if summary_stats else 0
        thk_count = summary_stats.get("total_thicknesses", 0) if summary_stats else 0
        breakdown = summary_stats.get("thickness_breakdown", {}) if summary_stats else {}

        breakdown_str = "\n".join([f"  • {k}: {v} DWG-uri" for k, v in sorted(breakdown.items())])

        messagebox.showinfo(
            "Regrupare Finalizată",
            f"Procesul de regrupare s-a încheiat cu succes!\n\n"
            f"• Total DWG-uri transferate: {total}\n"
            f"• Foldere de grosime create: {thk_count}\n\n"
            f"Distribuție pe grosimi:\n{breakdown_str}\n\n"
            f"Folder destinație: {dest_dir}\n"
            f"Raport Excel: {os.path.basename(report_path) if report_path else '-'}"
        )


def main():
    app = MiniExtragereNCApp()
    app.mainloop()


if __name__ == "__main__":
    main()
