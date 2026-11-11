"""
Aplicație Desktop Modernă: Generator Import-uri Excel din Word (Metalix .DOC / .DOCX)
Transformă automat unul sau mai multe documente Word (.doc / .docx) în fișiere Excel (.xlsx)
cu denumirea piesei curățată:
- Fără extensie (.dft / .DFT)
- Fără caractere speciale (-, _, spații etc.)
- Fără sufixe _001, _002 etc.
- Cu păstrarea excepțiilor +1.1, +1.2, +1.3 etc.
"""

import os
import re
import sys
import threading
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk

from core import (
    clean_part_name,
    parse_word_document,
    aggregate_parts,
    write_import_excel,
    process_batch_documents,
)

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")


def natural_sort_key(path_str: str):
    """Cheie de sortare naturală (ex: Steel-5-1, Steel-5-2, ..., Steel-5-10)."""
    name = os.path.basename(path_str)
    return [int(tok) if tok.isdigit() else tok.lower() for tok in re.split(r'(\d+)', name)]


class GeneratorImportExcelApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Generator Import-uri Excel din Word (.DOC / .DOCX -> .XLSX)")
        self.geometry("1120x860")
        self.minsize(940, 700)

        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "app_icon.ico")
        if os.path.exists(icon_path):
            try:
                self.iconbitmap(icon_path)
            except Exception:
                pass

        self.selected_files = []
        self.parsed_cache = {}  # file_path -> parsed dict
        self.last_output_dir = None
        self.last_generated_excel = None
        self._is_running = False

        # Locație implicită pentru dialogul de fișiere Word
        self.default_word_dir = r"D:\Programe plasma" if os.path.exists(r"D:\Programe plasma") else os.path.expanduser("~")
        self.default_desktop_dir = os.path.join(os.path.expanduser("~"), "Desktop")

        self._build_ui()

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        main_scroll = ctk.CTkScrollableFrame(self, fg_color="transparent")
        main_scroll.grid(row=0, column=0, sticky="nsew", padx=12, pady=10)
        main_scroll.grid_columnconfigure(0, weight=1)

        # ═════════════════════════════════════════════════════════════════════
        # HEADER
        # ═════════════════════════════════════════════════════════════════════
        header_frame = ctk.CTkFrame(main_scroll, corner_radius=10, fg_color=("gray85", "#1e222d"))
        header_frame.pack(fill="x", pady=(0, 10), padx=2)

        title_lbl = ctk.CTkLabel(
            header_frame,
            text="📄 ➜ 📊 GENERATOR IMPORT-URI EXCEL DIN WORD (METALIX)",
            font=ctk.CTkFont(size=19, weight="bold")
        )
        title_lbl.pack(anchor="w", padx=16, pady=(10, 2))

        sub_lbl = ctk.CTkLabel(
            header_frame,
            text="Extrage automat piesele și cantitățile din fișierele Word (.doc / .docx), elimină .dft, caracterele speciale și _001/_002, păstrând +1.1, +1.2, +1.3.",
            font=ctk.CTkFont(size=12),
            text_color=("gray30", "gray70")
        )
        sub_lbl.pack(anchor="w", padx=16, pady=(0, 10))

        # ═════════════════════════════════════════════════════════════════════
        # PASUL 1: SELECȚIE DOCUMENTE WORD (UNUL SAU BATCH)
        # ═════════════════════════════════════════════════════════════════════
        step1_frame = ctk.CTkFrame(main_scroll, corner_radius=10)
        step1_frame.pack(fill="x", pady=(0, 10), padx=2)
        step1_frame.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            step1_frame,
            text="1️⃣  SELECTEAZĂ DOCUMENTUL WORD SAU UN BATCH DE DOCUMENTE (.DOC / .DOCX)",
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(anchor="w", padx=14, pady=(10, 6))

        btn_row = ctk.CTkFrame(step1_frame, fg_color="transparent")
        btn_row.pack(fill="x", padx=14, pady=(0, 8))

        self.btn_pick_files = ctk.CTkButton(
            btn_row,
            text="📄 Selectează Fișier(e) Word (.doc / .docx)",
            font=ctk.CTkFont(size=13, weight="bold"),
            height=36,
            fg_color="#2563eb",
            hover_color="#1d4ed8",
            command=self._on_pick_word_files
        )
        self.btn_pick_files.pack(side="left", padx=(0, 8))

        self.btn_add_files = ctk.CTkButton(
            btn_row,
            text="➕ Adaugă Fișiere",
            font=ctk.CTkFont(size=12),
            height=36,
             width=130,
            fg_color="#3b82f6",
            hover_color="#2563eb",
            command=lambda: self._on_pick_word_files(append=True)
        )
        self.btn_add_files.pack(side="left", padx=(0, 8))

        self.btn_pick_folder = ctk.CTkButton(
            btn_row,
            text="📁 Încarcă din Folder...",
            font=ctk.CTkFont(size=12, weight="bold"),
            height=36,
            width=165,
            fg_color="#0d9488",
            hover_color="#0f766e",
            command=self._on_pick_folder
        )
        self.btn_pick_folder.pack(side="left", padx=(0, 8))

        # Filtru rapid după cod proiect (ex: 1147-46)
        ctk.CTkLabel(btn_row, text="Filtru nume (opțional):", font=ctk.CTkFont(size=12)).pack(side="left", padx=(6, 4))
        self.filter_var = ctk.StringVar(value="")
        self.entry_filter = ctk.CTkEntry(
            btn_row,
            textvariable=self.filter_var,
            placeholder_text="ex: 1147-46",
            width=125,
            height=34
        )
        self.entry_filter.pack(side="left", padx=(0, 6))

        self.btn_quick_search = ctk.CTkButton(
            btn_row,
            text="🔍 Caută în D:\\Programe plasma",
            font=ctk.CTkFont(size=12),
            height=34,
            fg_color="#4f46e5",
            hover_color="#4338ca",
            command=self._on_quick_search_plasma
        )
        self.btn_quick_search.pack(side="left", padx=(0, 8))

        self.btn_clear_files = ctk.CTkButton(
            btn_row,
            text="🗑️ Golește",
            font=ctk.CTkFont(size=12),
            height=34,
            width=90,
            fg_color="#dc2626",
            hover_color="#b91c1c",
            command=self._on_clear_files
        )
        self.btn_clear_files.pack(side="right")

        # Bara info selecție
        self.lbl_selection_summary = ctk.CTkLabel(
            step1_frame,
            text="Niciun document Word selectat. Apasă pe «Selectează Fișier(e) Word» sau caută după proiect (ex: 1147-46).",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=("#1e40af", "#60a5fa")
        )
        self.lbl_selection_summary.pack(anchor="w", padx=14, pady=(0, 10))

        # ═════════════════════════════════════════════════════════════════════
        # PASUL 2: PREVIZUALIZARE LIVE (DOCUMENTE & PIESE CURĂȚATE)
        # ═════════════════════════════════════════════════════════════════════
        step2_frame = ctk.CTkFrame(main_scroll, corner_radius=10)
        step2_frame.pack(fill="both", expand=True, pady=(0, 10), padx=2)
        step2_frame.grid_columnconfigure(0, weight=1)
        step2_frame.grid_columnconfigure(1, weight=2)

        top_prev = ctk.CTkFrame(step2_frame, fg_color="transparent")
        top_prev.pack(fill="x", padx=14, pady=(10, 4))

        ctk.CTkLabel(
            top_prev,
            text="2️⃣  PREVIZUALIZARE PIESE EXTRASE (ORIGINAL .DFT ➜ CURĂȚAT PENTRU EXCEL)",
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(side="left")

        self.btn_copy_clipboard = ctk.CTkButton(
            top_prev,
            text="📋 Copiază Tabelul în Clipboard (Paste în Excel)",
            font=ctk.CTkFont(size=12, weight="bold"),
            height=30,
            fg_color="#d97706",
            hover_color="#b45309",
            command=self._on_copy_to_clipboard
        )
        self.btn_copy_clipboard.pack(side="right")

        tables_container = ctk.CTkFrame(step2_frame, fg_color="transparent")
        tables_container.pack(fill="both", expand=True, padx=14, pady=(2, 10))
        tables_container.grid_columnconfigure(0, weight=2)
        tables_container.grid_columnconfigure(1, weight=5)
        tables_container.grid_rowconfigure(0, weight=1)

        # Stânga: Lista de fișiere Word selectate
        docs_box = ctk.CTkFrame(tables_container, corner_radius=8)
        docs_box.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        docs_box.grid_rowconfigure(1, weight=1)
        docs_box.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            docs_box,
            text="📑 Documente Word încărcate (click pt. filtrare):",
            font=ctk.CTkFont(size=12, weight="bold")
        ).grid(row=0, column=0, columnspan=2, sticky="w", padx=8, pady=(6, 4))

        self._configure_treeview_style()

        self.tree_docs = ttk.Treeview(
            docs_box,
            columns=("doc_name", "rows", "qty"),
            show="headings",
            height=10,
            selectmode="browse"
        )
        self.tree_docs.heading("doc_name", text="Document Word")
        self.tree_docs.heading("rows", text="Poz.")
        self.tree_docs.heading("qty", text="Buc.")
        self.tree_docs.column("doc_name", width=190, anchor="w")
        self.tree_docs.column("rows", width=50, anchor="center")
        self.tree_docs.column("qty", width=55, anchor="center")
        self.tree_docs.grid(row=1, column=0, sticky="nsew", padx=(8, 0), pady=(0, 8))

        scroll_docs = ttk.Scrollbar(docs_box, orient="vertical", command=self.tree_docs.yview)
        self.tree_docs.configure(yscrollcommand=scroll_docs.set)
        scroll_docs.grid(row=1, column=1, sticky="ns", padx=(0, 6), pady=(0, 8))

        self.tree_docs.bind("<<TreeviewSelect>>", self._on_doc_tree_select)

        # Dreapta: Tabelul de piese (No. | Denumire Word (.dft) | Denumire Excel | Cantitate | Greutate)
        parts_box = ctk.CTkFrame(tables_container, corner_radius=8)
        parts_box.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        parts_box.grid_rowconfigure(1, weight=1)
        parts_box.grid_columnconfigure(0, weight=1)

        self.lbl_parts_table_title = ctk.CTkLabel(
            parts_box,
            text="🔩 Piese extrase (Col A = Denumire curățată | Col B = Cantitate):",
            font=ctk.CTkFont(size=12, weight="bold")
        )
        self.lbl_parts_table_title.grid(row=0, column=0, columnspan=2, sticky="w", padx=8, pady=(6, 4))

        self.tree_parts = ttk.Treeview(
            parts_box,
            columns=("source", "no", "raw_name", "clean_name", "qty", "dims"),
            show="headings",
            height=10
        )
        self.tree_parts.heading("source", text="Fișier Word")
        self.tree_parts.heading("no", text="No.")
        self.tree_parts.heading("raw_name", text="Part name (în Word)")
        self.tree_parts.heading("clean_name", text="➜ EXCEL (Col A)")
        self.tree_parts.heading("qty", text="Cant. (Col B)")
        self.tree_parts.heading("dims", text="Dimensiuni (mm)")

        self.tree_parts.column("source", width=155, anchor="w")
        self.tree_parts.column("no", width=45, anchor="center")
        self.tree_parts.column("raw_name", width=165, anchor="w")
        self.tree_parts.column("clean_name", width=145, anchor="center")
        self.tree_parts.column("qty", width=85, anchor="center")
        self.tree_parts.column("dims", width=120, anchor="center")
        self.tree_parts.grid(row=1, column=0, sticky="nsew", padx=(8, 0), pady=(0, 8))

        scroll_parts = ttk.Scrollbar(parts_box, orient="vertical", command=self.tree_parts.yview)
        self.tree_parts.configure(yscrollcommand=scroll_parts.set)
        scroll_parts.grid(row=1, column=1, sticky="ns", padx=(0, 6), pady=(0, 8))

        # ═════════════════════════════════════════════════════════════════════
        # PASUL 3: SETĂRI EXPORT EXCEL & DESTINAȚIE
        # ═════════════════════════════════════════════════════════════════════
        step3_frame = ctk.CTkFrame(main_scroll, corner_radius=10)
        step3_frame.pack(fill="x", pady=(0, 10), padx=2)
        step3_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            step3_frame,
            text="3️⃣  OPȚIUNI GENERARE EXCEL (.XLSX) ȘI FOLDER DESTINAȚIE",
            font=ctk.CTkFont(size=14, weight="bold")
        ).grid(row=0, column=0, columnspan=3, sticky="w", padx=14, pady=(10, 8))

        # Mod de export
        ctk.CTkLabel(
            step3_frame,
            text="Mod generare Excel:",
            font=ctk.CTkFont(size=12, weight="bold")
        ).grid(row=1, column=0, sticky="w", padx=14, pady=4)

        self.export_mode_var = ctk.StringVar(value="individual")
        mode_frame = ctk.CTkFrame(step3_frame, fg_color="transparent")
        mode_frame.grid(row=1, column=1, columnspan=2, sticky="w", padx=6, pady=4)

        ctk.CTkRadioButton(
            mode_frame,
            text="Batch: Câte un Excel pentru fiecare Word (1:1)",
            variable=self.export_mode_var,
            value="individual",
            command=self._on_mode_change
        ).pack(side="left", padx=(0, 16))

        ctk.CTkRadioButton(
            mode_frame,
            text="Un singur Excel consolidat (toate piesele împreună)",
            variable=self.export_mode_var,
            value="consolidated_single",
            command=self._on_mode_change
        ).pack(side="left", padx=(0, 16))

        ctk.CTkRadioButton(
            mode_frame,
            text="Un singur Excel (câte un Sheet per Word)",
            variable=self.export_mode_var,
            value="consolidated_sheets",
            command=self._on_mode_change
        ).pack(side="left")

        # Format rând de start în Excel
        ctk.CTkLabel(
            step3_frame,
            text="Format tabel Excel:",
            font=ctk.CTkFont(size=12, weight="bold")
        ).grid(row=2, column=0, sticky="w", padx=14, pady=4)

        opts_frame = ctk.CTkFrame(step3_frame, fg_color="transparent")
        opts_frame.grid(row=2, column=1, columnspan=2, sticky="w", padx=6, pady=4)

        self.start_row_mode = ctk.StringVar(value="row2_empty1")
        ctk.CTkRadioButton(
            opts_frame,
            text="Începe de la Rândul 2 (Rândul 1 gol - identic cu IMPORT.xlsx)",
            variable=self.start_row_mode,
            value="row2_empty1"
        ).pack(side="left", padx=(0, 16))

        ctk.CTkRadioButton(
            opts_frame,
            text="Începe de la Rândul 1",
            variable=self.start_row_mode,
            value="row1"
        ).pack(side="left", padx=(0, 16))

        ctk.CTkRadioButton(
            opts_frame,
            text="Cu Antet pe Rândul 1 (Denumire piesa | Cantitate)",
            variable=self.start_row_mode,
            value="header_row1"
        ).pack(side="left")

        # Opțiune cumulare duplicate + nume fișier consolidat
        extra_frame = ctk.CTkFrame(step3_frame, fg_color="transparent")
        extra_frame.grid(row=3, column=1, columnspan=2, sticky="w", padx=6, pady=4)

        self.merge_dups_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            extra_frame,
            text="Cumulează cantitățile dacă aceeași piesă apare de mai multe ori",
            variable=self.merge_dups_var,
            command=self._refresh_preview_tables
        ).pack(side="left", padx=(0, 20))

        ctk.CTkLabel(extra_frame, text="Nume fișier (pt. mod consolidat):").pack(side="left", padx=(0, 6))
        self.consolidated_name_var = ctk.StringVar(value="IMPORT.xlsx")
        self.entry_consolidated_name = ctk.CTkEntry(
            extra_frame,
            textvariable=self.consolidated_name_var,
            width=180,
            height=30
        )
        self.entry_consolidated_name.pack(side="left")

        # Folder destinație
        ctk.CTkLabel(
            step3_frame,
            text="Folder Destinație Excel:",
            font=ctk.CTkFont(size=12, weight="bold")
        ).grid(row=4, column=0, sticky="w", padx=14, pady=(6, 12))

        self.output_dir_var = ctk.StringVar(value=self.default_desktop_dir)
        self.entry_output_dir = ctk.CTkEntry(
            step3_frame,
            textvariable=self.output_dir_var,
            height=34
        )
        self.entry_output_dir.grid(row=4, column=1, sticky="ew", padx=6, pady=(6, 12))

        btn_browse_out = ctk.CTkButton(
            step3_frame,
            text="📂 Alege Folder...",
            width=140,
            height=34,
            command=self._on_pick_output_dir
        )
        btn_browse_out.grid(row=4, column=2, sticky="e", padx=14, pady=(6, 12))

        # ═════════════════════════════════════════════════════════════════════
        # PASUL 4: BUTOANE ACȚIUNE & PROGRES
        # ═════════════════════════════════════════════════════════════════════
        action_frame = ctk.CTkFrame(main_scroll, corner_radius=10)
        action_frame.pack(fill="x", pady=(0, 6), padx=2)

        btn_bar = ctk.CTkFrame(action_frame, fg_color="transparent")
        btn_bar.pack(fill="x", padx=14, pady=(12, 8))

        self.btn_generate = ctk.CTkButton(
            btn_bar,
            text="🚀 GENEREAZĂ EXCEL(.XLSX)",
            font=ctk.CTkFont(size=15, weight="bold"),
            height=44,
            fg_color="#16a34a",
            hover_color="#15803d",
            command=self._on_generate_excel
        )
        self.btn_generate.pack(side="left", fill="x", expand=True, padx=(0, 8))

        self.btn_overwrite_import = ctk.CTkButton(
            btn_bar,
            text="⚡ Salvează direct în Desktop\\IMPORT.xlsx",
            font=ctk.CTkFont(size=13, weight="bold"),
            height=44,
            fg_color="#2563eb",
            hover_color="#1d4ed8",
            command=self._on_quick_save_desktop_import
        )
        self.btn_overwrite_import.pack(side="left", padx=(0, 8))

        self.btn_open_excel = ctk.CTkButton(
            btn_bar,
            text="📊 Deschide Excel-ul",
            font=ctk.CTkFont(size=12, weight="bold"),
            height=44,
            width=160,
            fg_color="#475569",
            hover_color="#334155",
            state="disabled",
            command=self._on_open_last_excel
        )
        self.btn_open_excel.pack(side="left", padx=(0, 8))

        self.btn_open_folder = ctk.CTkButton(
            btn_bar,
            text="📂 Deschide Folderul",
            font=ctk.CTkFont(size=12, weight="bold"),
            height=44,
            width=160,
            fg_color="#475569",
            hover_color="#334155",
            command=self._on_open_output_folder
        )
        self.btn_open_folder.pack(side="left")

        self.progress_bar = ctk.CTkProgressBar(action_frame, height=10)
        self.progress_bar.pack(fill="x", padx=14, pady=(0, 6))
        self.progress_bar.set(0)

        self.lbl_status = ctk.CTkLabel(
            action_frame,
            text="Pregătit. Selectează unul sau mai multe documente Word (.doc / .docx) pentru a începe.",
            font=ctk.CTkFont(size=12),
            text_color=("gray20", "gray80")
        )
        self.lbl_status.pack(anchor="w", padx=14, pady=(0, 10))

    def _configure_treeview_style(self):
        style = ttk.Style()
        style.theme_use("default")
        style.configure(
            "Treeview",
            background="#1e222d",
            foreground="#f1f5f9",
            rowheight=26,
            fieldbackground="#1e222d",
            borderwidth=0,
            font=("Segoe UI", 10)
        )
        style.configure(
            "Treeview.Heading",
            background="#0f172a",
            foreground="#93c5fd",
            font=("Segoe UI", 10, "bold"),
            relief="flat"
        )
        style.map(
            "Treeview",
            background=[("selected", "#2563eb")],
            foreground=[("selected", "#ffffff")]
        )

    # ═════════════════════════════════════════════════════════════════════════
    # HANDLERE SELECȚIE FIȘIERE WORD
    # ═════════════════════════════════════════════════════════════════════════
    def _on_pick_word_files(self, append: bool = False):
        files = filedialog.askopenfilenames(
            title="Selectează unul sau mai multe documente Word (.doc / .docx)",
            initialdir=self.default_word_dir,
            filetypes=[
                ("Documente Word", "*.doc *.docx *.DOC *.DOCX"),
                ("Toate fișierele", "*.*"),
            ]
        )
        if not files:
            return
        self.default_word_dir = os.path.dirname(files[0])
        self._load_word_files(list(files), append=append)

    def _on_pick_folder(self):
        folder = filedialog.askdirectory(
            title="Selectează folderul cu documente Word (.doc / .docx)",
            initialdir=self.default_word_dir
        )
        if not folder:
            return
        self.default_word_dir = folder
        self._scan_folder_for_docs(folder)

    def _on_quick_search_plasma(self):
        plasma_dir = r"D:\Programe plasma"
        if not os.path.exists(plasma_dir):
            messagebox.showwarning(
                "Folder inexistent",
                "Folderul D:\\Programe plasma nu a fost găsit pe acest calculator.\nFolosește butonul «Încarcă din Folder...»."
            )
            return
        flt = self.filter_var.get().strip()
        if not flt:
            messagebox.showinfo(
                "Introdu un filtru de proiect",
                "Scrie în căsuța «Filtru nume» codul proiectului (de exemplu: 1147-46) ca să încarci rapid doar fișierele acelui proiect din D:\\Programe plasma."
            )
            return
        self._scan_folder_for_docs(plasma_dir)

    def _scan_folder_for_docs(self, folder: str):
        flt = self.filter_var.get().strip().lower()
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
            messagebox.showinfo(
                "Niciun fișier găsit",
                f"Nu s-au găsit documente Word (.doc / .docx) în:\n{folder}"
                + (f"\ncu filtrul «{flt}»." if flt else ".")
            )
            return

        matched.sort(key=natural_sort_key)
        self._load_word_files(matched, append=False)

    def _on_clear_files(self):
        self.selected_files.clear()
        self.parsed_cache.clear()
        self._refresh_preview_tables()
        self.lbl_selection_summary.configure(
            text="Lista a fost golită. Selectează unul sau mai multe documente Word (.doc / .docx)."
        )
        self.lbl_status.configure(text="Lista de fișiere a fost resetată.")

    def _load_word_files(self, file_paths, append: bool = False):
        if not append:
            self.selected_files.clear()
            self.parsed_cache.clear()

        existing = set(os.path.normcase(os.path.abspath(p)) for p in self.selected_files)
        for p in file_paths:
            norm = os.path.normcase(os.path.abspath(p))
            if norm not in existing:
                self.selected_files.append(p)
                existing.add(norm)

        self.selected_files.sort(key=natural_sort_key)

        # Parsăm documentele pentru previzualizare instantanee
        for p in self.selected_files:
            if p not in self.parsed_cache:
                try:
                    self.parsed_cache[p] = parse_word_document(p)
                except Exception as exc:
                    self.parsed_cache[p] = {
                        "file_path": p,
                        "file_name": os.path.basename(p),
                        "metadata": {},
                        "parts": [],
                        "error": str(exc),
                    }

        self._refresh_preview_tables()

    def _refresh_preview_tables(self):
        for item in self.tree_docs.get_children():
            self.tree_docs.delete(item)
        for item in self.tree_parts.get_children():
            self.tree_parts.delete(item)

        if not self.selected_files:
            return

        merge_dups = self.merge_dups_var.get()
        total_rows = 0
        total_qty = 0

        # Adăugăm opțiunea "[TOATE DOCUMENTELE]" pe primul rând dacă sunt mai multe fișiere
        if len(self.selected_files) > 1:
            for p in self.selected_files:
                d = self.parsed_cache.get(p, {})
                pts = d.get("parts", [])
                if merge_dups:
                    pts = aggregate_parts(pts)
                total_rows += len(pts)
                total_qty += sum(x["quantity"] for x in pts)

            self.tree_docs.insert(
                "",
                "end",
                iid="__ALL__",
                values=(f"📚 TOATE ({len(self.selected_files)} fișiere)", total_rows, total_qty)
            )

        total_rows_check = 0
        total_qty_check = 0
        for p in self.selected_files:
            d = self.parsed_cache.get(p, {})
            pts = d.get("parts", [])
            if merge_dups:
                pts = aggregate_parts(pts)
            r_cnt = len(pts)
            q_cnt = sum(x["quantity"] for x in pts)
            total_rows_check += r_cnt
            total_qty_check += q_cnt
            self.tree_docs.insert(
                "",
                "end",
                iid=p,
                values=(os.path.basename(p), r_cnt, q_cnt)
            )

        self.lbl_selection_summary.configure(
            text=f"✅ {len(self.selected_files)} document(e) Word încărcate  |  Total poziții: {total_rows_check}  |  Total bucăți piese: {total_qty_check}"
        )

        # Selectăm implicit primul rând din tree_docs pentru a popula tabelul de piese
        first_iid = "__ALL__" if len(self.selected_files) > 1 else self.selected_files[0]
        self.tree_docs.selection_set(first_iid)
        self._populate_parts_tree(first_iid)

    def _on_doc_tree_select(self, _event=None):
        sel = self.tree_docs.selection()
        if not sel:
            return
        self._populate_parts_tree(sel[0])

    def _populate_parts_tree(self, selected_iid: str):
        for item in self.tree_parts.get_children():
            self.tree_parts.delete(item)

        merge_dups = self.merge_dups_var.get()

        if selected_iid == "__ALL__":
            target_files = self.selected_files
            self.lbl_parts_table_title.configure(
                text=f"🔩 Piese extrase din TOATE cele {len(target_files)} documente Word:"
            )
        else:
            target_files = [selected_iid]
            self.lbl_parts_table_title.configure(
                text=f"🔩 Piese extrase din: {os.path.basename(selected_iid)}"
            )

        for fpath in target_files:
            d = self.parsed_cache.get(fpath, {})
            fname = os.path.basename(fpath)
            pts = d.get("parts", [])
            if merge_dups:
                pts = aggregate_parts(pts)
            for p in pts:
                dims = f"{p.get('size_x', '')} x {p.get('size_y', '')}".strip(" x")
                self.tree_parts.insert(
                    "",
                    "end",
                    values=(
                        fname,
                        p.get("no", ""),
                        p.get("raw_name", ""),
                        p.get("clean_name", ""),
                        p.get("quantity", ""),
                        dims,
                    )
                )

    def _on_mode_change(self):
        mode = self.export_mode_var.get()
        if mode == "individual":
            self.entry_consolidated_name.configure(state="disabled")
        else:
            self.entry_consolidated_name.configure(state="normal")

    def _on_pick_output_dir(self):
        folder = filedialog.askdirectory(
            title="Selectează folderul unde vor fi salvate fișierele Excel",
            initialdir=self.output_dir_var.get() or self.default_desktop_dir
        )
        if folder:
            self.output_dir_var.set(folder)

    # ═════════════════════════════════════════════════════════════════════════
    # COPIERE ÎN CLIPBOARD & SALVARE RAPIDĂ IMPORT.XLSX
    # ═════════════════════════════════════════════════════════════════════════
    def _on_copy_to_clipboard(self):
        if not self.selected_files:
            messagebox.showwarning("Atenție", "Selectează mai întâi cel puțin un document Word.")
            return

        sel = self.tree_docs.selection()
        selected_iid = sel[0] if sel else ("__ALL__" if len(self.selected_files) > 1 else self.selected_files[0])
        target_files = self.selected_files if selected_iid == "__ALL__" else [selected_iid]

        merge_dups = self.merge_dups_var.get()
        all_parts = []
        for fpath in target_files:
            d = self.parsed_cache.get(fpath, {})
            all_parts.extend(d.get("parts", []))

        if merge_dups:
            all_parts = aggregate_parts(all_parts)

        if not all_parts:
            messagebox.showwarning("Atenție", "Nu există piese de copiat.")
            return

        lines = [f"{p['clean_name']}\t{p['quantity']}" for p in all_parts]
        tsv_text = "\n".join(lines)
        self.clipboard_clear()
        self.clipboard_append(tsv_text)
        self.update()

        self.lbl_status.configure(
            text=f"📋 S-au copiat în Clipboard {len(all_parts)} rânduri (Col A + Col B). Poți da direct Ctrl+V în Excel!"
        )
        messagebox.showinfo(
            "Copiat în Clipboard",
            f"Au fost copiate {len(all_parts)} poziții!\n\nAcum poți da direct Paste (Ctrl+V) în Excel."
        )

    def _get_start_row_settings(self):
        mode = self.start_row_mode.get()
        if mode == "row2_empty1":
            return 2, False
        elif mode == "row1":
            return 1, False
        else:  # header_row1
            return 2, True

    def _on_quick_save_desktop_import(self):
        if not self.selected_files:
            messagebox.showwarning("Atenție", "Selectează mai întâi unul sau mai multe documente Word.")
            return

        sel = self.tree_docs.selection()
        selected_iid = sel[0] if sel else ("__ALL__" if len(self.selected_files) > 1 else self.selected_files[0])
        target_files = self.selected_files if selected_iid == "__ALL__" else [selected_iid]

        merge_dups = self.merge_dups_var.get()
        all_parts = []
        for fpath in target_files:
            d = self.parsed_cache.get(fpath, {})
            all_parts.extend(d.get("parts", []))
        if merge_dups:
            all_parts = aggregate_parts(all_parts)

        if not all_parts:
            messagebox.showwarning("Atenție", "Nu s-au găsit piese în documentele selectate.")
            return

        import_path = os.path.join(self.default_desktop_dir, "IMPORT.xlsx")
        start_row, include_header = self._get_start_row_settings()

        try:
            write_import_excel(
                parts=all_parts,
                output_path=import_path,
                start_row=start_row,
                include_header=include_header,
            )
            self.last_generated_excel = import_path
            self.last_output_dir = self.default_desktop_dir
            self.btn_open_excel.configure(state="normal")
            self.lbl_status.configure(
                text=f"✅ Salvat cu succes în {import_path} ({len(all_parts)} poziții, {sum(p['quantity'] for p in all_parts)} bucăți)."
            )
            messagebox.showinfo(
                "IMPORT.xlsx Actualizat",
                f"Fișierul a fost salvat cu succes:\n{import_path}\n\n"
                f"• Poziții scrise: {len(all_parts)}\n"
                f"• Total bucăți: {sum(p['quantity'] for p in all_parts)}"
            )
        except PermissionError:
            messagebox.showerror(
                "Fișier deschis în Excel",
                f"Nu se poate salva în:\n{import_path}\n\nÎnchide mai întâi fișierul IMPORT.xlsx din Microsoft Excel și încearcă din nou!"
            )
        except Exception as exc:
            messagebox.showerror("Eroare la salvare", f"A apărut o eroare:\n{exc}")

    # ═════════════════════════════════════════════════════════════════════════
    # GENERARE BATCH / CONSOLIDAT
    # ═════════════════════════════════════════════════════════════════════════
    def _on_generate_excel(self):
        if self._is_running:
            return
        if not self.selected_files:
            messagebox.showwarning("Atenție", "Selectează mai întâi cel puțin un document Word (.doc / .docx)!")
            return

        output_dir = self.output_dir_var.get().strip() or self.default_desktop_dir
        mode = self.export_mode_var.get()
        consolidated_name = self.consolidated_name_var.get().strip() or "IMPORT.xlsx"
        start_row, include_header = self._get_start_row_settings()
        merge_dups = self.merge_dups_var.get()

        self._is_running = True
        self.btn_generate.configure(state="disabled", text="⏳ Se generează...")
        self.progress_bar.set(0)

        def worker():
            def on_progress(curr, total, msg):
                self.after(0, lambda: self._update_progress(curr, total, msg))

            report = process_batch_documents(
                word_files=list(self.selected_files),
                output_dir=output_dir,
                mode=mode,
                consolidated_filename=consolidated_name,
                start_row=start_row,
                include_header=include_header,
                merge_duplicates=merge_dups,
                progress_callback=on_progress,
            )
            self.after(0, lambda: self._on_generation_complete(report, output_dir))

        threading.Thread(target=worker, daemon=True).start()

    def _update_progress(self, curr: int, total: int, msg: str):
        frac = curr / max(total, 1)
        self.progress_bar.set(frac)
        self.lbl_status.configure(text=msg)

    def _on_generation_complete(self, report: dict, output_dir: str):
        self._is_running = False
        self.btn_generate.configure(state="normal", text="🚀 GENEREAZĂ EXCEL(.XLSX)")
        self.progress_bar.set(1.0)

        gen_files = report.get("generated_files", [])
        errors = report.get("errors", [])
        self.last_output_dir = output_dir

        if gen_files:
            self.last_generated_excel = gen_files[-1]
            self.btn_open_excel.configure(state="normal")

        summary_msg = (
            f"✅ Generare finalizată! Documente procesate: {report['processed_docs']}/{report['total_docs']} | "
            f"Fișiere Excel create: {len(gen_files)} | Total bucăți: {report['total_qty_sum']}"
        )
        self.lbl_status.configure(text=summary_msg)

        if errors:
            err_preview = "\n".join(errors[:8])
            if len(errors) > 8:
                err_preview += f"\n... și încă {len(errors) - 8} atenționări."
            messagebox.showwarning(
                "Finalizat cu atenționări",
                f"{summary_msg}\n\nFolder destinație:\n{output_dir}\n\nAtenționări:\n{err_preview}"
            )
        else:
            messagebox.showinfo(
                "Generare Reușită",
                f"Toate fișierele Excel au fost generate cu succes!\n\n"
                f"• Documente Word procesate: {report['processed_docs']}\n"
                f"• Fișiere Excel generate: {len(gen_files)}\n"
                f"• Total poziții piese: {report['total_parts_rows']}\n"
                f"• Total bucăți piese: {report['total_qty_sum']}\n\n"
                f"Salvate în:\n{output_dir}"
            )

    def _on_open_last_excel(self):
        if self.last_generated_excel and os.path.exists(self.last_generated_excel):
            os.startfile(self.last_generated_excel)
        else:
            messagebox.showwarning("Fișier negăsit", "Nu există încă un fișier Excel generat.")

    def _on_open_output_folder(self):
        target = self.last_output_dir or self.output_dir_var.get().strip() or self.default_desktop_dir
        if os.path.exists(target):
            os.startfile(target)
        else:
            messagebox.showwarning("Folder inexistent", f"Folderul nu există încă:\n{target}")


if __name__ == "__main__":
    app = GeneratorImportExcelApp()
    app.mainloop()

