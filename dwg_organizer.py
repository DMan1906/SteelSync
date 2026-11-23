"""
Core logic for Mini Extragere & Organizare NC, PDF & DWG.
Handles:
- Excel file parsing & auto-column detection
- Source folder scanning
- Part, Assembly, and DWG variation matching
- Plate thickness subfolder detection
- Safe Move / Copy execution
- Detailed Excel & Text report generation
"""

import os
import re
import shutil
import stat
import sys
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Regex to detect plate thickness folders (e.g., '10', 't10', 'PL12', '15mm', 'thk 8', 'grosime 20', '20_S355')
THICKNESS_REGEX = re.compile(
    r'^(?:t|pl|p|thk|grosime)?[_\s-]*\d+(?:[\.,]\d+)?[_\s-]*(?:mm)?(?:[_\s-]+[a-z0-9]+)?$|^\d+(?:[\.,]\d+)?$',
    re.IGNORECASE
)

GENERIC_DIR_NAMES = {
    'dwg', 'dwgs', 'cad', 'drawings', 'desene', 'model', 'table', 'placi', 'fisiere', 'export', 'nc', 'pdf'
}


def sanitize_folder_name(name):
    """Sanitize string for Windows folder name."""
    s = re.sub(r'[\\/:*?"<>|]', '_', str(name))
    s = s.strip('. ')
    return s or "Piesa_Necunoscuta"


def clean_alphanumeric(text):
    """Return only uppercase alphanumeric characters."""
    return re.sub(r'[^A-Za-z0-9]', '', str(text)).upper()


def is_nc_match(filename, part_code):
    """
    Check if NC filename matches the part code.
    Tolerance for special characters, extensions (.nc, .nc1), and clean alphanumeric.
    """
    ext = os.path.splitext(filename)[1].lower()
    if ext not in ['.nc', '.nc1']:
        return False
    stem = os.path.splitext(filename)[0].strip()
    p_str = str(part_code).strip()
    if not p_str or not stem:
        return False

    # 1. Exact match (case-insensitive)
    if stem.upper() == p_str.upper():
        return True

    # 2. Stripped alphanumeric match
    clean_p = clean_alphanumeric(p_str)
    clean_s = clean_alphanumeric(stem)
    if clean_p and clean_s == clean_p:
        return True

    # 3. Special characters translation (e.g. * -> x, _, -)
    alt_p = re.sub(r'[*]', 'x', p_str)
    if clean_alphanumeric(alt_p) == clean_s:
        return True

    return False


def is_assembly_match(filename, asmb_code):
    """
    Check if a PDF drawing filename corresponds to asmb_code.
    Tolerance for special characters, project prefixes, and revisions.
    """
    ext = os.path.splitext(filename)[1].lower()
    if ext != '.pdf':
        return False
    stem = os.path.splitext(filename)[0].strip()
    asmb_str = str(asmb_code).strip()
    if not asmb_str or not stem:
        return False

    # 1. Exact match
    if stem.upper() == asmb_str.upper():
        return True

    clean_asmb = clean_alphanumeric(asmb_str)
    clean_stem = clean_alphanumeric(stem)

    # 2. Cleaned alphanumeric match
    if clean_asmb and clean_stem == clean_asmb:
        return True

    # 3. Token chunk match (split by spaces, underscores, hyphens, brackets)
    chunks = [c for c in re.split(r'[\s_()\[\],;]+', stem) if c]
    for c in chunks:
        if clean_alphanumeric(c) == clean_asmb:
            return True

    # 4. Asmb substring match in clean stem (with boundary check)
    if clean_asmb and len(clean_asmb) >= 3 and clean_asmb in clean_stem:
        start = 0
        while True:
            idx = clean_stem.find(clean_asmb, start)
            if idx == -1:
                break
            after_idx = idx + len(clean_asmb)
            bad_left = False
            if idx > 0:
                if clean_asmb[0].isalpha() and clean_stem[idx - 1].isalpha():
                    bad_left = True
                elif clean_asmb[0].isdigit() and clean_stem[idx - 1].isdigit():
                    bad_left = True
            # Avoid false positives where code is a prefix of a larger number (e.g. B20 matching B200)
            bad_right = (
                after_idx < len(clean_stem)
                and clean_asmb[-1].isdigit()
                and clean_stem[after_idx].isdigit()
            )
            if not bad_left and not bad_right:
                return True
            start = idx + 1
        return False

    return False


def is_dwg_variation(filename, base_part, allow_variations=True):
    """
    Check if DWG filename matches base_part (exact or CAD variations T, L, C, etc.).
    Returns (True/False, Note).
    """
    ext = os.path.splitext(filename)[1].lower()
    if ext != '.dwg':
        return False, ""
    stem = os.path.splitext(filename)[0].strip()
    base_str = str(base_part).strip()

    # 1. Exact match
    if stem.upper() == base_str.upper():
        return True, "Exact"

    clean_dwg = clean_alphanumeric(stem)
    clean_base = clean_alphanumeric(base_str)

    if clean_dwg == clean_base:
        return True, "Exact (curățat)"

    if not allow_variations:
        return False, ""

    # 2. Prefix variation (e.g. base_part + T, L, C, T1, L1, C1, etc.)
    if clean_dwg.startswith(clean_base):
        suffix = clean_dwg[len(clean_base):]
        # Suffix must not start with a digit (so '101' will NOT match '1015' or '1010')
        if len(suffix) > 0 and not suffix[0].isdigit():
            if re.match(r'^[A-Z0-9]{1,6}$', suffix):
                return True, f"Variație ({suffix})"

    return False, ""


def get_dwg_thickness_folder(dwg_full_path, source_root):
    """
    Detect the plate thickness folder where a DWG resides.
    Traverses parent directories from deepest to root.
    """
    try:
        rel = os.path.relpath(os.path.dirname(dwg_full_path), source_root)
    except Exception:
        rel = os.path.dirname(dwg_full_path)

    parts = [p.strip() for p in rel.split(os.sep) if p.strip() and p.strip() != '.']
    if not parts:
        return ""

    # Search backwards for regex match
    for p in reversed(parts):
        if p.lower() in GENERIC_DIR_NAMES:
            continue
        if THICKNESS_REGEX.match(p):
            return p

    # Fallback: take closest parent folder that is not generic
    for p in reversed(parts):
        if p.lower() not in GENERIC_DIR_NAMES:
            return p

    return parts[-1]


def detect_excel_structure(filepath, sheet_name=None):
    """
    Inspect Excel file to find available sheets, header row,
    and auto-detect 'Ansamblu' and 'Piesă' column indices.
    """
    wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
    try:
        sheets = wb.sheetnames
        if not sheet_name or sheet_name not in sheets:
            sheet_name = sheets[0]
        ws = wb[sheet_name]

        asmb_keywords = [
            r'ansambl', r'asembl', r'assembly', r'main.*part', r'marca', r'mark', r'poz.*ans', r'denumire.*ans'
        ]
        part_keywords = [
            r'^denumire$', r'pies[aă]', r'piese', r'part', r'single.*part', r'poz.*pies[aă]', r'profil', r'pozitie', r'reper', r'denumire.*pie'
        ]

        header_row_idx = None
        asmb_col = None
        part_col = None
        headers = []

        # Scan first 30 rows
        for r_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
            if r_idx > 30:
                break
            row_strs = [str(c).strip() if c is not None else "" for c in row]
            a_matches = []
            p_matches = []
            for c_idx, val in enumerate(row_strs):
                if not val:
                    continue
                val_lower = val.lower()
                if any(re.search(k, val_lower) for k in asmb_keywords):
                    a_matches.append((c_idx, val))
                elif any(re.search(k, val_lower) for k in part_keywords):
                    p_matches.append((c_idx, val))

            if a_matches and p_matches:
                header_row_idx = r_idx
                asmb_col = a_matches[0][0]
                part_col = p_matches[0][0]
                headers = [
                    (i, val if val else f"Coloana {get_column_letter(i+1)}")
                    for i, val in enumerate(row_strs)
                ]
                break

        # If both not found on same row, fallback to first row with strings
        if header_row_idx is None:
            for r_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
                if r_idx > 30:
                    break
                row_strs = [str(c).strip() if c is not None else "" for c in row]
                non_empty = [c for c in row_strs if c]
                if len(non_empty) >= 2:
                    header_row_idx = r_idx
                    headers = [
                        (i, val if val else f"Coloana {get_column_letter(i+1)}")
                        for i, val in enumerate(row_strs)
                    ]
                    for c_idx, val in enumerate(row_strs):
                        val_lower = val.lower()
                        if asmb_col is None and any(re.search(k, val_lower) for k in asmb_keywords):
                            asmb_col = c_idx
                        if part_col is None and any(re.search(k, val_lower) for k in part_keywords):
                            part_col = c_idx
                    break

        # If still not detected, avoid picking index column ('Nr.', 'Nr. crt.', '#', 'No.') as assembly column
        if asmb_col is None and headers and part_col != 0:
            col0_name = headers[0][1].strip()
            if not re.match(r'^(?:nr\.?(?:\s*crt\.?)?|#|no\.?)$', col0_name, re.IGNORECASE):
                asmb_col = 0
        if part_col is None and len(headers) > 1:
            part_col = 1

        return {
            "sheets": sheets,
            "selected_sheet": sheet_name,
            "header_row": header_row_idx or 1,
            "asmb_col_idx": asmb_col,
            "part_col_idx": part_col,
            "headers": headers
        }
    finally:
        wb.close()


def parse_excel_parts(filepath, sheet_name, header_row, asmb_col_idx, part_col_idx):
    """
    Parse Excel table into parts dictionary with hierarchical assembly forward-fill:
    {
       part_name: [asmb_name_1, asmb_name_2, ...]
    }
    """
    wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
    try:
        ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.active

        parts_map = {}
        ignored_vals = {'none', 'nan', 'null', 'total', 'suma', 'sum', ''}
        current_asmb = ""

        start_row = header_row + 1
        for row in ws.iter_rows(min_row=start_row, values_only=True):
            if not row:
                continue
            asmb_val = ""
            part_val = ""

            if asmb_col_idx is not None and asmb_col_idx < len(row) and row[asmb_col_idx] is not None:
                asmb_val = str(row[asmb_col_idx]).strip()

            if part_col_idx is not None and part_col_idx < len(row) and row[part_col_idx] is not None:
                part_val = str(row[part_col_idx]).strip()

            if asmb_val and asmb_val.lower() not in ignored_vals:
                current_asmb = asmb_val

            if not part_val or part_val.lower() in ignored_vals:
                continue

            effective_asmb = asmb_val if (asmb_val and asmb_val.lower() not in ignored_vals) else current_asmb
            sub_parts = [p.strip() for p in part_val.split(",") if p.strip()]
            for single_part in sub_parts:
                if single_part.lower() in ignored_vals:
                    continue
                if single_part not in parts_map:
                    parts_map[single_part] = []
                if effective_asmb and effective_asmb.lower() not in ignored_vals and effective_asmb not in parts_map[single_part]:
                    parts_map[single_part].append(effective_asmb)

        return parts_map
    finally:
        wb.close()


def extract_filter_parts(filter_path):
    """
    Extract unique target part codes from either:
    - an Excel report/BOM file (all sheets, all part columns, e.g. raport_CHESON.xlsx), OR
    - a folder containing part subfolders (e.g. DSB-2094), NC files (.nc/.nc1), or PDF files.
    Returns {"parts": [...], "source_type": str, "count": int}.
    """
    if not filter_path or not os.path.exists(filter_path):
        raise ValueError(f"Calea de filtrare nu există: {filter_path}")

    seen_clean = set()
    result_parts = []
    ignored_vals = {'none', 'nan', 'null', 'total', 'suma', 'sum', '-', ''}
    header_skip_re = re.compile(
        r'^(?:nr\.?(?:\s*crt\.?)?|#|no\.?|cant\.?(?:itate[a]?)?|buc[aă]?[tț]?i?|material|greutate.*|poz.*|reper|pies[aă]|denumire.*)$',
        re.IGNORECASE
    )
    index_col_re = re.compile(r'^(?:nr\.?(?:\s*crt\.?)?|#|no\.?|cant\.?|buc)$', re.IGNORECASE)

    def add_part(raw_val):
        for token in str(raw_val or "").split(","):
            t = token.strip()
            if not t or t.lower() in ignored_vals or header_skip_re.match(t):
                continue
            cl = clean_alphanumeric(t)
            if not cl or len(cl) < 2 or cl in seen_clean:
                continue
            seen_clean.add(cl)
            result_parts.append(t)

    if os.path.isfile(filter_path):
        part_keywords = [
            r'^denumire$', r'pies[aă]', r'piese', r'part', r'single.*part',
            r'poz.*pies[aă]', r'profil', r'pozitie', r'reper', r'denumire.*pie'
        ]
        wb = openpyxl.load_workbook(filter_path, data_only=True, read_only=True)
        try:
            for sh_name in wb.sheetnames:
                ws = wb[sh_name]
                rows_sample = []
                for r_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
                    rows_sample.append((r_idx, row))
                    if r_idx >= 30:
                        break
                if not rows_sample:
                    continue

                header_row_idx = None
                part_cols = []
                for r_idx, row in rows_sample:
                    row_strs = [str(c).strip() if c is not None else "" for c in row]
                    matched_cols = [
                        c_idx for c_idx, val in enumerate(row_strs)
                        if val and any(re.search(k, val.lower()) for k in part_keywords)
                    ]
                    if matched_cols:
                        header_row_idx = r_idx
                        part_cols = matched_cols
                        break

                if not part_cols:
                    for r_idx, row in rows_sample:
                        row_strs = [str(c).strip() if c is not None else "" for c in row]
                        non_empty = [(i, v) for i, v in enumerate(row_strs) if v]
                        if non_empty:
                            # Skip leading index column if present
                            candidate = non_empty[0][0]
                            if len(non_empty) > 1 and (index_col_re.match(non_empty[0][1]) or non_empty[0][1].isdigit()):
                                candidate = non_empty[1][0]
                            header_row_idx = r_idx if any(not v.replace('.', '', 1).isdigit() for _, v in non_empty) else 0
                            part_cols = [candidate]
                            break

                start_row = (header_row_idx or 0) + 1
                for row in ws.iter_rows(min_row=start_row, values_only=True):
                    if not row:
                        continue
                    for c_idx in part_cols:
                        if c_idx < len(row) and row[c_idx] is not None:
                            val = row[c_idx]
                            if isinstance(val, (int, float)) and not isinstance(val, bool):
                                continue
                            add_part(val)
        finally:
            wb.close()

        return {
            "parts": result_parts,
            "source_type": f"Excel ({os.path.basename(filter_path)})",
            "count": len(result_parts),
        }

    # Directory mode: direct subfolders + .nc/.nc1 + .pdf (with recursive fallback)
    excluded_dirs = GENERIC_DIR_NAMES | {'modified', 'desene ansamblu', 'sorted', 'placi', 'profile', 'bare', 'cheson'}
    material_dirs = {'s235', 's355', 's275', 's235jr', 's355j2', 's355j0', 'z35', 'inox', 'aluminiu'}

    def clean_pdf_part_stem(stem_str):
        cleaned = re.sub(r'[\s_-]*(?:W|A|G|M)[\s_-]*drawing(?:\s*\(.*\))?$', '', stem_str, flags=re.IGNORECASE).strip()
        return cleaned

    def is_valid_part_dir(dir_name):
        dl = dir_name.strip().lower()
        if not dl or dl.startswith('.') or dl.startswith('_'):
            return False
        if dl in excluded_dirs or dl in material_dirs:
            return False
        if THICKNESS_REGEX.match(dir_name.strip()):
            return False
        return True

    entries = sorted(os.scandir(filter_path), key=lambda e: e.name.lower())
    for entry in entries:
        if entry.is_dir():
            if is_valid_part_dir(entry.name):
                add_part(entry.name)
        elif entry.is_file():
            ext = os.path.splitext(entry.name)[1].lower()
            stem = os.path.splitext(entry.name)[0].strip()
            if ext in ('.nc', '.nc1'):
                add_part(stem)
            elif ext == '.pdf':
                add_part(clean_pdf_part_stem(stem))

    if not result_parts:
        for root, dirs, files in os.walk(filter_path, followlinks=True):
            for d in sorted(dirs):
                if is_valid_part_dir(d):
                    add_part(d)
            for f in sorted(files):
                ext = os.path.splitext(f)[1].lower()
                stem = os.path.splitext(f)[0].strip()
                if ext in ('.nc', '.nc1'):
                    add_part(stem)
                elif ext == '.pdf':
                    add_part(clean_pdf_part_stem(stem))

    return {
        "parts": result_parts,
        "source_type": f"Folder ({os.path.basename(os.path.normpath(filter_path))})",
        "count": len(result_parts),
    }


def find_assembly_pdf_dir(excel_path=None, source_dir=None, explicit_dir=None):
    """Locate an extracted assembly drawings folder ('Desene Ansamblu' / 'Assembly Drawings') with PDFs."""
    if explicit_dir:
        exp_s = str(explicit_dir).strip()
        if exp_s and os.path.isdir(exp_s):
            return os.path.normpath(exp_s)

    target_names = ("desene ansamblu", "desene_ansamblu", "assembly drawings", "assembly_drawings", "ansambluri")
    base_parents = []
    if excel_path:
        ep = os.path.dirname(os.path.abspath(str(excel_path).strip()))
        if ep:
            base_parents.append(ep)
    if source_dir:
        sp = os.path.abspath(str(source_dir).strip())
        if sp:
            base_parents.append(sp)
            p1 = os.path.dirname(sp)
            p2 = os.path.dirname(p1) if p1 else ""
            if p1:
                base_parents.append(p1)
            if p2:
                base_parents.append(p2)

    seen_dirs = set()
    for bp in base_parents:
        if not bp or not os.path.isdir(bp):
            continue
        norm_bp = os.path.normcase(os.path.abspath(bp))
        if norm_bp in seen_dirs:
            continue
        seen_dirs.add(norm_bp)
        try:
            for entry in os.scandir(bp):
                if entry.is_dir() and entry.name.strip().lower() in target_names:
                    has_pdf = any(
                        f.lower().endswith(".pdf")
                        for _, _, files in os.walk(entry.path, followlinks=True)
                        for f in files
                    )
                    if has_pdf:
                        return os.path.normpath(entry.path)
        except OSError:
            pass
    return ""


def scan_pdf_directory(pdf_dir):
    """Walk a dedicated assembly drawings folder and return PDF records marked from_asmb_dir=True."""
    pdf_files = []
    if not pdf_dir or not os.path.isdir(pdf_dir):
        return pdf_files
    for root, _, files in os.walk(pdf_dir, followlinks=True):
        for f in sorted(files):
            if os.path.splitext(f)[1].lower() == ".pdf":
                full_path = os.path.join(root, f)
                pdf_files.append({
                    "filename": f,
                    "path": full_path,
                    "ext": ".pdf",
                    "stem": os.path.splitext(f)[0],
                    "from_asmb_dir": True,
                })
    return pdf_files


def scan_source_files(source_dir, skip_dirs=None):
    """
    Walk source directory recursively and index NC, PDF, and DWG files.
    """
    nc_files = []
    pdf_files = []
    dwg_files = []
    thicknesses_found = set()
    if not source_dir or not os.path.isdir(source_dir):
        return {
            "nc_files": nc_files,
            "pdf_files": pdf_files,
            "dwg_files": dwg_files,
            "thicknesses": [],
        }

    skip_norm = {os.path.normcase(os.path.abspath(d)) for d in (skip_dirs or ()) if d}

    for root, dirs, files in os.walk(source_dir, followlinks=True):
        if skip_norm:
            dirs[:] = [
                d for d in dirs
                if os.path.normcase(os.path.abspath(os.path.join(root, d))) not in skip_norm
            ]
        for f in files:
            full_path = os.path.join(root, f)
            ext = os.path.splitext(f)[1].lower()
            stem = os.path.splitext(f)[0]

            item = {
                "filename": f,
                "path": full_path,
                "ext": ext,
                "stem": stem
            }

            if ext in ['.nc', '.nc1']:
                nc_files.append(item)
            elif ext == '.pdf':
                pdf_files.append(item)
            elif ext == '.dwg':
                thk = get_dwg_thickness_folder(full_path, source_dir)
                item["thickness"] = thk
                if thk:
                    thicknesses_found.add(thk)
                dwg_files.append(item)

    return {
        "nc_files": nc_files,
        "pdf_files": pdf_files,
        "dwg_files": dwg_files,
        "thicknesses": sorted(list(thicknesses_found))
    }


def generate_excel_report(report_path, results_list, summary_stats):
    """Generate a formatted Excel report summarizing the run."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Raport Organizare"
    ws.views.sheetView[0].showGridLines = True

    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    title_font = Font(name="Segoe UI", size=14, bold=True, color="1F497D")
    normal_font = Font(name="Segoe UI", size=10)
    bold_font = Font(name="Segoe UI", size=10, bold=True)

    fill_ok = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    font_ok = Font(name="Segoe UI", size=10, color="006100", bold=True)

    fill_partial = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    font_partial = Font(name="Segoe UI", size=10, color="9C6500", bold=True)

    fill_missing = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    font_missing = Font(name="Segoe UI", size=10, color="9C0006", bold=True)

    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    # Title & summary block
    ws["A1"] = "RAPORT ORGANIZARE PIESE (NC, PDF & DWG)"
    ws["A1"].font = title_font

    ws["A3"] = f"Total piese analizate: {summary_stats.get('total_parts', 0)}"
    ws["A4"] = f"Piese complete (NC+PDF+DWG): {summary_stats.get('complete_parts', 0)}"
    ws["A5"] = f"Piese parțiale: {summary_stats.get('partial_parts', 0)}"
    ws["A6"] = f"Piese fără fișiere găsite: {summary_stats.get('missing_parts', 0)}"
    ws["A7"] = f"Mod operare: {summary_stats.get('mode', 'Mutare').upper()}"
    ws["A8"] = f"Total fișiere transferate: {summary_stats.get('total_files_transferred', 0)}"

    for r in range(3, 9):
        ws[f"A{r}"].font = bold_font

    # Table headers
    headers = [
        "Nr.",
        "Denumire Piesă (Excel)",
        "Folder Destinație",
        "Ansambluri Asociate",
        "Fișiere NC Găsite",
        "Desene PDF Ansamblu Găsite",
        "Piese DWG Găsite (Grosimi & Variații)",
        "Status Piesă"
    ]

    header_row = 10
    for col_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Data rows
    for idx, item in enumerate(results_list, start=1):
        curr_row = header_row + idx
        ws.cell(row=curr_row, column=1, value=idx).alignment = Alignment(horizontal="center")
        ws.cell(row=curr_row, column=2, value=item["part"])
        ws.cell(row=curr_row, column=3, value=item["folder"])
        ws.cell(row=curr_row, column=4, value=", ".join(item["assemblies"]) if item["assemblies"] else "-")
        
        # NC cell
        nc_text = ", ".join(item["nc_found"]) if item["nc_found"] else "LIPSĂ"
        ws.cell(row=curr_row, column=5, value=nc_text)

        # PDF cell
        pdf_text = ", ".join(item["pdf_found"]) if item["pdf_found"] else "LIPSĂ"
        ws.cell(row=curr_row, column=6, value=pdf_text)

        # DWG cell
        if item["dwg_found"]:
            dwg_details = [f"{d['name']} [{d.get('thickness', 'Fără')}]" for d in item["dwg_found"]]
            dwg_text = f"({len(item['dwg_found'])}) " + ", ".join(dwg_details)
        else:
            dwg_text = "LIPSĂ"
        ws.cell(row=curr_row, column=7, value=dwg_text)

        # Status cell
        status_cell = ws.cell(row=curr_row, column=8, value=item["status"])
        status_cell.alignment = Alignment(horizontal="center", vertical="center")
        if item["status"] == "COMPLET":
            status_cell.fill = fill_ok
            status_cell.font = font_ok
        elif item["status"] == "PARȚIAL":
            status_cell.fill = fill_partial
            status_cell.font = font_partial
        else:
            status_cell.fill = fill_missing
            status_cell.font = font_missing

        for col_idx in range(1, 9):
            ws.cell(row=curr_row, column=col_idx).border = thin_border
            if col_idx != 8:
                ws.cell(row=curr_row, column=col_idx).font = normal_font

    # Auto-adjust column widths
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.row < header_row:
                continue
            val_str = str(cell.value or "")
            if len(val_str) > max_len:
                max_len = len(val_str)
        ws.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 45)

    wb.save(report_path)


def run_pipeline(
    excel_path,
    sheet_name,
    header_row,
    asmb_col_idx,
    part_col_idx,
    source_dir,
    dest_dir,
    mode="move",
    allow_variations=True,
    clean_empty_folders=True,
    progress_callback=None,
    log_callback=None,
    is_cancelled=None,
    filter_path=None,
    include_nc=True,
    include_pdf=True,
    include_dwg=True,
    flat_pdf_only=False,
    pdf_source_dir=None,
    create_asmb_folder=False,
):
    """
    Full pipeline runner.
    Calls progress_callback(fraction, current_message) and log_callback(msg, level).
    Supports optional filtering by part/cheson list (Excel or folder), dedicated
    extracted assembly PDF folder (Desene Ansamblu), collective output subfolder,
    selective file-type extraction (NC / PDF / DWG), and flat PDF mode.
    Returns (summary_dict, results_list, report_path).
    """
    def log(msg, level="INFO"):
        if log_callback:
            log_callback(msg, level)

    def set_progress(pct, msg=""):
        if progress_callback:
            progress_callback(pct, msg)

    def _norm_abs(p):
        return os.path.normcase(os.path.abspath(p))

    def _is_same_file(src_p, dst_p):
        if not os.path.exists(dst_p):
            return False
        try:
            return os.path.samefile(src_p, dst_p)
        except Exception:
            return _norm_abs(src_p) == _norm_abs(dst_p)

    eff_nc = False if flat_pdf_only else bool(include_nc)
    eff_pdf = True if flat_pdf_only else bool(include_pdf)
    eff_dwg = False if flat_pdf_only else bool(include_dwg)

    resolved_pdf_dir = ""
    if eff_pdf:
        resolved_pdf_dir = find_assembly_pdf_dir(
            excel_path=excel_path,
            source_dir=source_dir,
            explicit_dir=pdf_source_dir,
        )
    is_separate_pdf_dir = bool(resolved_pdf_dir) and (
        not source_dir or _norm_abs(resolved_pdf_dir) != _norm_abs(source_dir)
    )

    log("=" * 60, "INFO")
    log("🚀 Inițializare proces organizare piese...", "INFO")
    log(f"📄 Fișier Excel BOM: {os.path.basename(excel_path)} (Foaia: {sheet_name})", "INFO")
    if filter_path:
        log(f"🎯 Filtru Piese/Chesoane: {filter_path}", "INFO")
    log(f"📁 Folder Sursă (NC/DWG): {source_dir or '(Neconfigurat - doar PDF Ansamblu)'}", "INFO")
    if resolved_pdf_dir:
        src_kind = "selectat" if (pdf_source_dir and str(pdf_source_dir).strip()) else "detectat automat"
        log(f"📋 Folder Desene Ansamblu (PDF, {src_kind}): {resolved_pdf_dir}", "INFO")
    log(f"📁 Folder Destinație: {dest_dir}", "INFO")
    log(f"⚙️ Mod Operare: {'MUTARE (Move)' if mode == 'move' else 'COPIERE (Copy)'}", "INFO")
    if flat_pdf_only:
        log(
            "📄 Mod Special: Doar PDF-uri de ansamblu direct în "
            + ("Destinație/Desene Ansamblu" if create_asmb_folder else "Destinație (plat)"),
            "INFO",
        )
    else:
        active_types = [t for t, en in (("NC", eff_nc), ("PDF Ansamblu", eff_pdf), ("DWG", eff_dwg)) if en]
        log(f"📦 Tipuri active: {', '.join(active_types) if active_types else 'Niciunul'}", "INFO")
        if eff_pdf and create_asmb_folder:
            log("📂 Se va crea și folderul centralizat «Desene Ansamblu» în Destinație.", "INFO")
        if eff_dwg:
            log(f"🔍 Suport Variații DWG (T, L, C): {'DA' if allow_variations else 'NU'}", "INFO")
    log("=" * 60, "INFO")

    # 1. Parse Excel BOM & optional part filter
    set_progress(0.05, "Parsare tabel Excel BOM...")
    bom_parts_map = parse_excel_parts(excel_path, sheet_name, header_row, asmb_col_idx, part_col_idx)
    if not bom_parts_map:
        raise ValueError("Nu a fost găsită nicio piesă validă în fișierul Excel BOM!")

    if filter_path:
        set_progress(0.10, "Extragere piese din filtru...")
        filter_info = extract_filter_parts(filter_path)
        filter_parts = filter_info["parts"]
        if not filter_parts:
            raise ValueError(f"Nu a fost găsită nicio piesă validă în filtrul selectat: {filter_path}")

        clean_bom_map = {}
        for bp, asmbs in bom_parts_map.items():
            cl = clean_alphanumeric(bp)
            if not cl:
                continue
            clean_bom_map.setdefault(cl, [])
            for a in asmbs:
                if a not in clean_bom_map[cl]:
                    clean_bom_map[cl].append(a)

        parts_map = {}
        matched_in_bom = 0
        for fp in filter_parts:
            asmbs = list(clean_bom_map.get(clean_alphanumeric(fp), []))
            parts_map[fp] = asmbs
            if asmbs:
                matched_in_bom += 1

        total_parts = len(parts_map)
        all_asmb_count = sum(len(asmbs) for asmbs in parts_map.values())
        log(
            f"✔️ Filtru aplicat ({filter_info['source_type']}): {total_parts} piese, "
            f"{matched_in_bom} găsite în BOM ({all_asmb_count} asocieri de ansambluri).",
            "SUCCESS",
        )
    else:
        parts_map = bom_parts_map
        total_parts = len(parts_map)
        all_asmb_count = sum(len(asmbs) for asmbs in parts_map.values())
        log(f"✔️ Excel analizat cu succes: {total_parts} piese unice, {all_asmb_count} asocieri de ansambluri.", "SUCCESS")

    if is_cancelled and is_cancelled():
        return None, None, None

    # 2. Scan source directory + dedicated assembly PDF directory
    set_progress(0.15, "Scanare surse (NC, PDF Ansamblu, DWG)...")
    asmb_out_dir = os.path.join(dest_dir, "Desene Ansamblu")
    skip_dirs = [asmb_out_dir] if (is_separate_pdf_dir and create_asmb_folder) else None
    src_data = scan_source_files(source_dir, skip_dirs=skip_dirs) if source_dir else {
        "nc_files": [], "pdf_files": [], "dwg_files": [], "thicknesses": []
    }
    nc_files = src_data["nc_files"]
    dwg_files = src_data["dwg_files"]
    thk_list = src_data["thicknesses"]

    pdf_files = []
    seen_pdf_paths = set()
    if eff_pdf and resolved_pdf_dir:
        extra_pdfs = scan_pdf_directory(resolved_pdf_dir)
        for epf in extra_pdfs:
            np_path = _norm_abs(epf["path"])
            if not is_separate_pdf_dir:
                epf = {**epf, "from_asmb_dir": False}
            if np_path not in seen_pdf_paths:
                seen_pdf_paths.add(np_path)
                pdf_files.append(epf)
    for spf in src_data["pdf_files"]:
        np_path = _norm_abs(spf["path"])
        if np_path not in seen_pdf_paths:
            seen_pdf_paths.add(np_path)
            pdf_files.append(spf)

    log("✔️ Surse scanate:", "SUCCESS")
    log(f"   • {len(nc_files)} fișiere NC (.nc, .nc1)", "INFO")
    log(
        f"   • {len(pdf_files)} desene PDF (ansambluri)"
        + (" [copiere protejată din Folder Desene Ansamblu]" if is_separate_pdf_dir else ""),
        "INFO",
    )
    log(f"   • {len(dwg_files)} piese DWG în {len(thk_list)} subfoldere de grosime ({', '.join(thk_list) if thk_list else 'niciunul'})", "INFO")

    if is_cancelled and is_cancelled():
        return None, None, None

    # 3. Process parts
    os.makedirs(dest_dir, exist_ok=True)
    if eff_pdf and create_asmb_folder:
        os.makedirs(asmb_out_dir, exist_ok=True)

    existing_dest_dirs = {}
    if not flat_pdf_only:
        try:
            for entry in os.scandir(dest_dir):
                if entry.is_dir():
                    if entry.name.strip().lower() == "desene ansamblu":
                        continue
                    cl_dir = clean_alphanumeric(entry.name)
                    if cl_dir and cl_dir not in existing_dest_dirs:
                        existing_dest_dirs[cl_dir] = entry.name
        except Exception:
            pass

    results_list = []
    files_to_delete_from_source = set()
    copied_dest_paths = set()
    copied_flat_pdfs = set()
    copied_asmb_folder_pdfs = set()
    total_files_transferred = 0

    complete_parts = 0
    partial_parts = 0
    missing_parts = 0

    start_prog = 0.2
    end_prog = 0.85

    for idx, (part, asmbs) in enumerate(parts_map.items()):
        if is_cancelled and is_cancelled():
            log("🛑 Procesul a fost oprit de utilizator.", "WARNING")
            break

        curr_prog = start_prog + (end_prog - start_prog) * (idx / max(1, total_parts))
        set_progress(curr_prog, f"Procesare piesa ({idx+1}/{total_parts}): {part}")

        if flat_pdf_only:
            folder_name = "(Desene Ansamblu)" if create_asmb_folder else "(Direct în Destinație)"
            part_dest_dir = asmb_out_dir if create_asmb_folder else dest_dir
        else:
            cl_part = clean_alphanumeric(part)
            folder_name = existing_dest_dirs.get(cl_part) or sanitize_folder_name(part)
            if cl_part:
                existing_dest_dirs[cl_part] = folder_name
            part_dest_dir = os.path.join(dest_dir, folder_name)
            os.makedirs(part_dest_dir, exist_ok=True)

        part_nc_found = []
        part_pdf_found = []
        part_dwg_found = []

        # A. NC Match
        if eff_nc:
            for ncf in nc_files:
                if is_nc_match(ncf["filename"], part):
                    part_nc_found.append(ncf["filename"])
                    dst_file = os.path.join(part_dest_dir, ncf["filename"])
                    copied_dest_paths.add(_norm_abs(dst_file))
                    if _is_same_file(ncf["path"], dst_file):
                        continue
                    try:
                        shutil.copy2(ncf["path"], dst_file)
                        files_to_delete_from_source.add(ncf["path"])
                        total_files_transferred += 1
                    except Exception as e:
                        log(f"❌ Eroare la copiere NC '{ncf['filename']}': {e}", "ERROR")

        # B. PDF Match for each associated assembly
        if eff_pdf:
            matched_pdf_names_for_part = set()
            for asmb in asmbs:
                for pf in pdf_files:
                    if is_assembly_match(pf["filename"], asmb):
                        fn_lower = pf["filename"].lower()
                        if fn_lower in matched_pdf_names_for_part:
                            continue
                        matched_pdf_names_for_part.add(fn_lower)
                        part_pdf_found.append(pf["filename"])

                        dst_file = os.path.join(part_dest_dir, pf["filename"])
                        norm_dst = _norm_abs(dst_file)
                        copied_dest_paths.add(norm_dst)

                        skip_part_copy = flat_pdf_only and (norm_dst in copied_flat_pdfs)
                        if not skip_part_copy:
                            if _is_same_file(pf["path"], dst_file):
                                if flat_pdf_only:
                                    copied_flat_pdfs.add(norm_dst)
                            else:
                                try:
                                    shutil.copy2(pf["path"], dst_file)
                                    if not pf.get("from_asmb_dir"):
                                        files_to_delete_from_source.add(pf["path"])
                                    total_files_transferred += 1
                                    if flat_pdf_only:
                                        copied_flat_pdfs.add(norm_dst)
                                except Exception as e:
                                    log(f"❌ Eroare la copiere PDF '{pf['filename']}': {e}", "ERROR")

                        if create_asmb_folder:
                            asmb_dst = os.path.join(asmb_out_dir, pf["filename"])
                            norm_asmb_dst = _norm_abs(asmb_dst)
                            copied_dest_paths.add(norm_asmb_dst)
                            if norm_asmb_dst not in copied_asmb_folder_pdfs:
                                copied_asmb_folder_pdfs.add(norm_asmb_dst)
                                if not _is_same_file(pf["path"], asmb_dst):
                                    try:
                                        shutil.copy2(pf["path"], asmb_dst)
                                        if not pf.get("from_asmb_dir"):
                                            files_to_delete_from_source.add(pf["path"])
                                        total_files_transferred += 1
                                    except Exception as e:
                                        log(f"❌ Eroare la copiere PDF în Desene Ansamblu '{pf['filename']}': {e}", "ERROR")

        # C. DWG Match in thickness subfolders
        if eff_dwg:
            for df in dwg_files:
                is_match, note = is_dwg_variation(df["filename"], part, allow_variations=allow_variations)
                if is_match:
                    thk = df.get("thickness", "")
                    if thk:
                        thk_dir = os.path.join(part_dest_dir, sanitize_folder_name(thk))
                    else:
                        thk_dir = part_dest_dir
                    os.makedirs(thk_dir, exist_ok=True)

                    dst_file = os.path.join(thk_dir, df["filename"])
                    copied_dest_paths.add(_norm_abs(dst_file))
                    part_dwg_found.append({
                        "name": df["filename"],
                        "thickness": thk or "Rădăcină",
                        "note": note
                    })
                    if _is_same_file(df["path"], dst_file):
                        continue
                    try:
                        shutil.copy2(df["path"], dst_file)
                        files_to_delete_from_source.add(df["path"])
                        total_files_transferred += 1
                    except Exception as e:
                        log(f"❌ Eroare la copiere DWG '{df['filename']}': {e}", "ERROR")

        # Evaluate status relative to active extraction modes
        if flat_pdf_only:
            if len(part_pdf_found) > 0:
                status = "COMPLET"
                complete_parts += 1
                log_lvl = "SUCCESS"
            else:
                status = "LIPSĂ"
                missing_parts += 1
                log_lvl = "ERROR"
        else:
            checks = []
            any_found = False
            if eff_nc:
                checks.append(len(part_nc_found) > 0)
                any_found = any_found or (len(part_nc_found) > 0)
            if eff_pdf:
                checks.append(len(part_pdf_found) > 0 or (len(asmbs) == 0 and not filter_path))
                any_found = any_found or (len(part_pdf_found) > 0)
            if eff_dwg:
                checks.append(len(part_dwg_found) > 0)
                any_found = any_found or (len(part_dwg_found) > 0)

            if checks and all(checks) and any_found:
                status = "COMPLET"
                complete_parts += 1
                log_lvl = "SUCCESS"
            elif any_found:
                status = "PARȚIAL"
                partial_parts += 1
                log_lvl = "WARNING"
            else:
                status = "LIPSĂ"
                missing_parts += 1
                log_lvl = "ERROR"

        results_list.append({
            "part": part,
            "folder": folder_name,
            "assemblies": asmbs,
            "nc_found": part_nc_found,
            "pdf_found": part_pdf_found,
            "dwg_found": part_dwg_found,
            "status": status
        })

        dwg_desc = f"{len(part_dwg_found)} DWG" if part_dwg_found else "0 DWG"
        log(f"[{status}] Piesă '{part}' (Asmb: {', '.join(asmbs) if asmbs else '-'}) -> NC: {len(part_nc_found)}, PDF: {len(part_pdf_found)}, {dwg_desc}", log_lvl)

    # 4. Clean up source files if mode == 'move' and not cancelled
    safe_delete_paths = [
        p for p in files_to_delete_from_source
        if _norm_abs(p) not in copied_dest_paths
    ]
    if mode == "move" and safe_delete_paths and not (is_cancelled and is_cancelled()):
        set_progress(0.88, "Ștergere fișiere mutate din sursă...")
        log(f"\n🗑️ Ștergere {len(safe_delete_paths)} fișiere originale din sursă...", "INFO")
        del_count = 0
        for src_path in safe_delete_paths:
            try:
                if os.path.exists(src_path):
                    # Clear read-only if present
                    try:
                        os.chmod(src_path, stat.S_IWRITE)
                    except Exception:
                        pass
                    os.remove(src_path)
                    del_count += 1
            except Exception as e:
                log(f"⚠️ Nu s-a putut șterge fișierul sursă '{os.path.basename(src_path)}': {e}", "WARNING")
        log(f"✔️ S-au șters {del_count} fișiere din sursă.", "SUCCESS")

        # Clean empty directories in source
        if clean_empty_folders and source_dir and os.path.isdir(source_dir):
            set_progress(0.92, "Curățare foldere goale...")
            removed_dirs = 0
            for root, dirs, files in os.walk(source_dir, topdown=False):
                if _norm_abs(root) == _norm_abs(source_dir):
                    continue
                if eff_pdf and create_asmb_folder and _norm_abs(root) == _norm_abs(asmb_out_dir):
                    continue
                try:
                    if not os.listdir(root):
                        os.rmdir(root)
                        removed_dirs += 1
                except Exception:
                    pass
            if removed_dirs > 0:
                log(f"✔️ S-au curățat {removed_dirs} foldere goale din sursă.", "INFO")

    # 5. Generate Excel Report
    set_progress(0.95, "Generare raport Excel...")
    report_path = os.path.join(dest_dir, "Raport_Organizare_Piese.xlsx")
    summary_stats = {
        "total_parts": total_parts,
        "complete_parts": complete_parts,
        "partial_parts": partial_parts,
        "missing_parts": missing_parts,
        "mode": "Mutare" if mode == "move" else "Copiere",
        "total_files_transferred": total_files_transferred,
        "asmb_folder_pdfs": len(copied_asmb_folder_pdfs),
        "asmb_folder_path": asmb_out_dir if (create_asmb_folder and copied_asmb_folder_pdfs) else None,
    }

    try:
        generate_excel_report(report_path, results_list, summary_stats)
        log(f"✔️ Raport salvat: {os.path.basename(report_path)}", "SUCCESS")
    except Exception as e:
        log(f"⚠️ Nu s-a putut genera raportul Excel: {e}", "WARNING")

    # Save log to text file
    set_progress(1.0, "Finalizat!")
    log("\n" + "=" * 60, "INFO")
    log("🎉 PROCES FINALIZAT CU SUCCES!", "SUCCESS")
    log(f"📊 Total piese: {total_parts} | Complete: {complete_parts} | Parțiale: {partial_parts} | Lipsă: {missing_parts}", "INFO")
    log(f"📂 Folder destinație: {dest_dir}", "INFO")
    log("=" * 60, "INFO")

    return summary_stats, results_list, report_path


# ─────────────────────────────────────────────────────────────────────────────
# 🔄 FUNCȚIA INVERSĂ: REGRUPARE DWG-URI PE GROSIMI DE TABLĂ
# ─────────────────────────────────────────────────────────────────────────────

def detect_dwg_thickness(dwg_full_path, root_dir, normalize_case=True):
    """
    Detects the thickness folder for a DWG file located in a piece subfolder hierarchy.
    E.g. Root/DSB2004/T10/DSB2004C.dwg -> 'T10'
    """
    try:
        rel = os.path.relpath(os.path.dirname(dwg_full_path), root_dir)
    except Exception:
        rel = os.path.dirname(dwg_full_path)

    parts = [p.strip() for p in rel.split(os.sep) if p.strip() and p.strip() != '.']
    if not parts:
        return '_Fara_Grosime'

    # Check closest folder first (bottom-up)
    for p in reversed(parts):
        if p.lower() in GENERIC_DIR_NAMES:
            continue
        if THICKNESS_REGEX.match(p):
            return p.upper() if normalize_case else p

    # Fallback to closest non-generic folder
    for p in reversed(parts):
        if p.lower() not in GENERIC_DIR_NAMES:
            return p.upper() if normalize_case else p

    return '_Fara_Grosime'


def scan_dwgs_for_regrouping(source_dir, normalize_case=True):
    """
    Scan source directory recursively for DWG files and group them by plate thickness.
    """
    groups = {}
    dwg_records = []

    for root, dirs, files in os.walk(source_dir, followlinks=True):
        for f in files:
            if f.lower().endswith('.dwg'):
                full_p = os.path.join(root, f)
                thk = detect_dwg_thickness(full_p, source_dir, normalize_case=normalize_case)
                rel = os.path.relpath(full_p, source_dir)
                p_parts = rel.split(os.sep)
                parent_part = p_parts[0] if len(p_parts) > 1 else "Rădăcină"

                rec = {
                    "filename": f,
                    "full_path": full_p,
                    "thickness": thk,
                    "parent_part": parent_part,
                    "size_bytes": os.path.getsize(full_p) if os.path.exists(full_p) else 0
                }
                dwg_records.append(rec)
                groups.setdefault(thk, []).append(rec)

    return {
        "total_dwgs": len(dwg_records),
        "thickness_groups": groups,
        "thicknesses": sorted(list(groups.keys())),
        "records": dwg_records
    }


def generate_regroup_excel_report(report_path, transferred_records, summary_stats):
    """Generate Excel report for regrouped DWGs."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Regrupare Grosimi DWG"
    ws.views.sheetView[0].showGridLines = True

    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    title_font = Font(name="Segoe UI", size=14, bold=True, color="1F497D")
    normal_font = Font(name="Segoe UI", size=10)
    bold_font = Font(name="Segoe UI", size=10, bold=True)

    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    # Title & summary block
    ws["A1"] = "RAPORT REGRUPARE DWG-URI PE GROSIMI DE TABLĂ"
    ws["A1"].font = title_font

    ws["A3"] = f"Total fișiere DWG procesate: {summary_stats.get('total_dwgs', 0)}"
    ws["A4"] = f"Număr grosimi identificate: {summary_stats.get('total_thicknesses', 0)}"
    ws["A5"] = f"Mod operare: {summary_stats.get('mode', 'Copiere').upper()}"
    ws["A6"] = f"Folder Destinație: {summary_stats.get('dest_dir', '')}"

    for r in range(3, 7):
        ws[f"A{r}"].font = bold_font

    headers = [
        "Nr.",
        "Grosime Tablă (Folder)",
        "Denumire Fișier DWG",
        "Piesă / Folder Sursă",
        "Destinație Relativă",
        "Mărime (KB)",
        "Status"
    ]

    header_row = 8
    for col_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    fill_ok = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    font_ok = Font(name="Segoe UI", size=10, color="006100", bold=True)

    for idx, rec in enumerate(transferred_records, start=1):
        curr_row = header_row + idx
        ws.cell(row=curr_row, column=1, value=idx).alignment = Alignment(horizontal="center")
        ws.cell(row=curr_row, column=2, value=rec.get("thickness", ""))
        ws.cell(row=curr_row, column=3, value=rec.get("filename", ""))
        ws.cell(row=curr_row, column=4, value=rec.get("parent_part", ""))
        ws.cell(row=curr_row, column=5, value=rec.get("dest_rel", ""))
        
        size_kb = round(rec.get("size_bytes", 0) / 1024, 1)
        ws.cell(row=curr_row, column=6, value=size_kb).alignment = Alignment(horizontal="right")

        status_cell = ws.cell(row=curr_row, column=7, value="REGRUPAT")
        status_cell.fill = fill_ok
        status_cell.font = font_ok
        status_cell.alignment = Alignment(horizontal="center")

        for col_idx in range(1, 8):
            ws.cell(row=curr_row, column=col_idx).border = thin_border
            if col_idx not in (1, 6, 7):
                ws.cell(row=curr_row, column=col_idx).font = normal_font

    # Column widths
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.row < header_row:
                continue
            val_str = str(cell.value or "")
            if len(val_str) > max_len:
                max_len = len(val_str)
        ws.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 50)

    wb.save(report_path)


def run_regroup_pipeline(
    source_dir,
    dest_dir,
    mode="copy",
    normalize_case=True,
    clean_empty_folders=True,
    progress_callback=None,
    log_callback=None,
    is_cancelled=None
):
    """
    Executes the regrouping of all DWG files from piece folders into unified thickness folders.
    """
    def log(msg, level="INFO"):
        if log_callback:
            log_callback(msg, level)

    def set_progress(pct, msg=""):
        if progress_callback:
            progress_callback(pct, msg)

    log("=" * 60, "INFO")
    log("🔄 Inițializare Regrupare DWG-uri pe Grosimi de Tablă...", "INFO")
    log(f"📁 Folder Sursă (cu foldere piese): {source_dir}", "INFO")
    log(f"📁 Folder Destinație (pe grosimi): {dest_dir}", "INFO")
    log(f"⚙️ Mod Operare: {'MUTARE (Move)' if mode == 'move' else 'COPIERE (Copy)'}", "INFO")
    log("=" * 60, "INFO")

    # 1. Scan source
    set_progress(0.1, "Scanare fișiere DWG din structura sursă...")
    scan_data = scan_dwgs_for_regrouping(source_dir, normalize_case=normalize_case)
    dwg_records = scan_data["records"]
    total_dwgs = scan_data["total_dwgs"]
    groups = scan_data["thickness_groups"]

    if total_dwgs == 0:
        raise ValueError("Nu a fost găsit niciun fișier .dwg în folderul sursă selectat!")

    log(f"✔️ S-au detectat {total_dwgs} fișiere DWG în {len(groups)} grosimi:", "SUCCESS")
    for thk, items in sorted(groups.items()):
        log(f"   • {thk}: {len(items)} piese DWG", "INFO")

    if is_cancelled and is_cancelled():
        return None, None

    os.makedirs(dest_dir, exist_ok=True)
    transferred_records = []
    files_to_delete = []

    start_prog = 0.2
    end_prog = 0.85

    # 2. Transfer files into thickness folders
    for idx, rec in enumerate(dwg_records):
        if is_cancelled and is_cancelled():
            log("🛑 Procesul a fost oprit de utilizator.", "WARNING")
            break

        curr_prog = start_prog + (end_prog - start_prog) * (idx / max(1, total_dwgs))
        set_progress(curr_prog, f"Regrupare ({idx+1}/{total_dwgs}): {rec['filename']} -> {rec['thickness']}")

        thk = rec["thickness"]
        target_thk_dir = os.path.join(dest_dir, sanitize_folder_name(thk))
        os.makedirs(target_thk_dir, exist_ok=True)

        # Handle filename collisions safely
        dest_filename = rec["filename"]
        dest_path = os.path.join(target_thk_dir, dest_filename)

        if os.path.exists(dest_path):
            try:
                # Check if it's identical
                if os.path.samefile(rec["full_path"], dest_path):
                    pass
                else:
                    # Rename with parent piece identifier
                    stem, ext = os.path.splitext(dest_filename)
                    disambig = sanitize_folder_name(rec["parent_part"])
                    dest_filename = f"{stem} ({disambig}){ext}"
                    dest_path = os.path.join(target_thk_dir, dest_filename)
                    # If still exists, add counter
                    counter = 1
                    while os.path.exists(dest_path):
                        dest_filename = f"{stem} ({disambig}_{counter}){ext}"
                        dest_path = os.path.join(target_thk_dir, dest_filename)
                        counter += 1
            except Exception:
                pass

        try:
            shutil.copy2(rec["full_path"], dest_path)
            files_to_delete.append(rec["full_path"])
            dest_rel = os.path.join(thk, dest_filename)
            rec["dest_path"] = dest_path
            rec["dest_rel"] = dest_rel
            transferred_records.append(rec)
            log(f"  [{thk}] Copiat '{rec['filename']}' din piesa '{rec['parent_part']}'", "SUCCESS")
        except Exception as e:
            log(f"❌ Eroare la copiere '{rec['filename']}': {e}", "ERROR")

    # 3. Clean up source files if mode == 'move' and not cancelled
    if mode == "move" and files_to_delete and not (is_cancelled and is_cancelled()):
        set_progress(0.88, "Ștergere fișiere mutate din sursă...")
        log(f"\n🗑️ Ștergere {len(files_to_delete)} fișiere DWG din sursă...", "INFO")
        del_count = 0
        for src_path in files_to_delete:
            try:
                if os.path.exists(src_path):
                    try:
                        os.chmod(src_path, stat.S_IWRITE)
                    except Exception:
                        pass
                    os.remove(src_path)
                    del_count += 1
            except Exception as e:
                log(f"⚠️ Nu s-a putut șterge '{os.path.basename(src_path)}': {e}", "WARNING")
        log(f"✔️ S-au șters {del_count} fișiere DWG din sursă.", "SUCCESS")

        # Clean empty folders in source
        if clean_empty_folders:
            set_progress(0.92, "Curățare foldere goale...")
            removed_dirs = 0
            for root, dirs, files in os.walk(source_dir, topdown=False):
                if root == source_dir:
                    continue
                try:
                    if not os.listdir(root):
                        os.rmdir(root)
                        removed_dirs += 1
                except Exception:
                    pass
            if removed_dirs > 0:
                log(f"✔️ S-au curățat {removed_dirs} foldere goale din sursă.", "INFO")

    # 4. Generate Excel Report
    set_progress(0.95, "Generare raport Excel...")
    report_path = os.path.join(dest_dir, "Raport_Regrupare_DWG.xlsx")
    summary_stats = {
        "total_dwgs": len(transferred_records),
        "total_thicknesses": len(groups),
        "thickness_breakdown": {thk: len(items) for thk, items in groups.items()},
        "mode": "Mutare" if mode == "move" else "Copiere",
        "dest_dir": dest_dir
    }

    try:
        generate_regroup_excel_report(report_path, transferred_records, summary_stats)
        log(f"✔️ Raport Excel salvat: {os.path.basename(report_path)}", "SUCCESS")
    except Exception as e:
        log(f"⚠️ Nu s-a putut genera raportul Excel: {e}", "WARNING")

    set_progress(1.0, "Finalizat!")
    log("\n" + "=" * 60, "INFO")
    log("🎉 REGRUPARE FINALIZATĂ CU SUCCES!", "SUCCESS")
    log(f"📊 Total DWG-uri regrupate: {len(transferred_records)} în {len(groups)} foldere de grosime.", "INFO")
    log(f"📂 Folder destinație: {dest_dir}", "INFO")
    log("=" * 60, "INFO")

    return summary_stats, report_path

