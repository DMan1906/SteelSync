"""
Core logic pentru Generare Import-uri Excel din documente Word (Metalix .DOC / .DOCX).
Funcționalități:
- Citire ultra-rapidă a fișierelor .DOC (format binar OLE2 Word 97-2003 generat de Metalix, RTF) și .DOCX
- Extragere tabel 'Parts' (No., Part name, Quantity, Weight, Size X, Size Y) + metadate antet (SubNest, Material, Part qty.)
- Curățare denumire piesă:
  * Eliminare extensie (.dft, .DFT, .dxf, .dwg, .nc, .nc1)
  * Eliminare sufixe de tip _001, _002, _003 etc. (inclusiv când apar înainte de +1.1, ex: DSB2419L_001+1.1.DFT -> DSB2419L+1.1)
  * Eliminare caractere speciale (-, _, spații etc., ex: DSB-2259.dft -> DSB2259)
  * Păstrare excepții +1.1, +1.2, +1.3 etc.
- Generare fișiere Excel (.xlsx) individuale (batch 1:1) sau consolidate, formatate identic cu IMPORT.xlsx.
"""

import os
import re
from pathlib import Path
from typing import List, Dict, Any, Optional, Tuple, Callable
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill


# Extensii CAD/CAM care trebuie eliminate din denumirea piesei
CAD_EXTENSIONS_REGEX = re.compile(
    r'\.(?:dft|dxf|dwg|nc1?|prt|par|stp|step|igs|iges)$',
    re.IGNORECASE
)

# Sufixe de multiplicare/indexare Metalix (_001, _002, _01 etc.)
# Se elimină:
# 1) Orice _0... (ex: _001, _002, _01, _02) oriunde în nume
# 2) Orice _<cifre> care apare DUPĂ un cod de piesă ce conține deja cifre (ex: DSB2273L1_001, DSB2040_1)
#    Notă: În 'DVIDSP_184.dft' sau 'IGDP_12.dft', înainte de '_' sunt doar litere, deci '184'/'12' este chiar numărul piesei
#    și se elimină doar caracterul '_' (rezultând DVIDSP184 / IGDP12).
INDEX_SUFFIX_REGEX = re.compile(
    r'(?:(?<=\d)_[0-9]+|(?<=\d[A-Za-z])_[0-9]+|(?<=\d[A-Za-z][0-9])_[0-9]+|(?<=\d[A-Za-z]{2})_[0-9]+|_0+[0-9]+)'
)

# Excepție: sufixe de tip +1.1, +1.2, +1.3, +2.1 etc.
PLUS_EXCEPTION_REGEX = re.compile(r'(\+\d+(?:\.\d+)+|\+\d+\.\d+)$')


def clean_part_name(
    raw_name: str,
    remove_extensions: bool = True,
    remove_index_suffix: bool = True,
    keep_plus_decimal: bool = True,
    uppercase: bool = True
) -> str:
    """
    Curăță denumirea piesei conform regulilor:
    - Elimină extensia (.dft / .DFT etc.)
    - Elimină sufixele de tip _001, _002 etc. (chiar și înainte de +1.1)
    - Păstrează +1.1, +1.2, +1.3 etc.
    - Elimină toate celelalte caractere speciale (-, _, spații etc.)

    Exemple:
      'DSB-2259.dft'          -> 'DSB2259'
      'DSP-2037.dft'          -> 'DSP2037'
      'DSB2273L1_001.dft'     -> 'DSB2273L1'
      'DSB2046T_002.dft'      -> 'DSB2046T'
      'DSB2052L+1.1.DFT'      -> 'DSB2052L+1.1'
      'DSB2419L_001+1.1.DFT'  -> 'DSB2419L+1.1'
      'DSB2128L1_001+1.1.DFT' -> 'DSB2128L1+1.1'
    """
    if not raw_name:
        return ""

    # Elimină caractere de control Word (\x07, \r, \n, \x01 etc.)
    s = re.sub(r'[\x00-\x1f]', '', str(raw_name)).strip()
    if not s:
        return ""

    # 1. Elimină extensia (.dft, .dxf, etc.)
    if remove_extensions:
        s = CAD_EXTENSIONS_REGEX.sub('', s).strip()

    # 2. Elimină _001, _002 etc. (înainte de extragerea +1.1 pentru a acoperi 'DSB2419L_001+1.1')
    if remove_index_suffix:
        s = INDEX_SUFFIX_REGEX.sub('', s)

    # 3. Extrage și protejează sufixul +1.1, +1.2, +1.3 etc.
    plus_suffix = ""
    if keep_plus_decimal:
        m_plus = PLUS_EXCEPTION_REGEX.search(s)
        if m_plus:
            plus_suffix = m_plus.group(1)
            s = s[:m_plus.start()]

    # Asigură că nu a mai rămas un _001 imediat înainte de +1.1
    if remove_index_suffix:
        s = INDEX_SUFFIX_REGEX.sub('', s)

    # Dacă există text adițional după spațiu în denumiri vechi (ex: 'DSB1003 T10 2BUC'),
    # păstrăm tot sau curățăm caracterele speciale.
    # 4. Elimină toate caracterele speciale din corpul denumirii (lăsând doar litere și cifre)
    cleaned_body = re.sub(r'[^A-Za-z0-9]', '', s)
    if uppercase:
        cleaned_body = cleaned_body.upper()

    return cleaned_body + plus_suffix


def _clean_cell_text(cell_raw: str) -> str:
    """Curăță textul brut dintr-o celulă de tabel Word."""
    if not cell_raw:
        return ""
    # Elimină caractere de control Word (\x01 = imagine inline, \x07 = delimitator celulă, \r, \n, \x0c)
    s = re.sub(r'[\x00-\x1f]', ' ', str(cell_raw))
    return re.sub(r'\s+', ' ', s).strip()


def _parse_binary_doc_text(text: str, file_path: str) -> Dict[str, Any]:
    """
    Parsează conținutul text extras dintr-un document Word binar (.DOC) Metalix.
    În formatul binar Word, celulele de tabel sunt delimitate prin '\\x07', iar sfârșitul de rând prin '\\x07\\x07'.
    """
    metadata = {
        "subnest": "",
        "material": "",
        "sheet_size": "",
        "sheets_qty": 1,
        "expected_part_qty": None,
        "machine": "",
        "date": "",
    }

    # Extragere metadate din antetul Metalix
    m_sub = re.search(r'SubNest:\s*([^\r\n\x07]+)', text, re.IGNORECASE)
    if m_sub:
        metadata["subnest"] = m_sub.group(1).strip()

    m_mat = re.search(r'Material:\s*([^\r\n\x07]+)', text, re.IGNORECASE)
    if m_mat:
        metadata["material"] = m_mat.group(1).strip()

    m_sh_qty = re.search(r'Sheets\s*qty\.?:\s*(\d+)', text, re.IGNORECASE)
    if m_sh_qty:
        metadata["sheets_qty"] = int(m_sh_qty.group(1))

    m_part_qty = re.search(r'Part\s*qty\.?:\s*(\d+)', text, re.IGNORECASE)
    if m_part_qty:
        metadata["expected_part_qty"] = int(m_part_qty.group(1))

    m_mach = re.search(r'Machine:\s*([^\r\n\x07]+)', text, re.IGNORECASE)
    if m_mach:
        metadata["machine"] = m_mach.group(1).strip()

    # Căutare tabel 'Parts' (conține 'Part name' și 'Quantity')
    parts_list: List[Dict[str, Any]] = []

    # Pot exista mai multe apariții ale antetului dacă tabelul se întinde pe mai multe pagini
    search_pos = 0
    while True:
        idx = text.find("Part name", search_pos)
        if idx == -1:
            # Încercare case-insensitive
            m_pn = re.search(r'Part\s+name', text[search_pos:], re.IGNORECASE)
            if not m_pn:
                break
            idx = search_pos + m_pn.start()

        sub = text[idx:]
        raw_rows = sub.split('\x07\x07')
        found_in_block = 0

        for r in raw_rows[1:]:
            cells = [_clean_cell_text(c) for c in r.split('\x07')]
            # Eliminăm celulele goale de la capăt dacă există
            while cells and cells[-1] == "":
                cells.pop()

            if not cells:
                if found_in_block > 0:
                    break
                continue

            # Dacă întâlnim un antet repetat pe pagina următoare (No. | Part name | Quantity)
            if len(cells) >= 2 and ('no' in cells[0].lower() or 'part name' in cells[1].lower()):
                continue

            # Rând valid de piesă: Col 0 = No. (număr), Col 1 = Part name, Col 2 = Quantity (număr)
            first_col = cells[0].strip()
            if len(cells) >= 3 and re.match(r'^\d+$', first_col):
                part_raw = cells[1].strip()
                qty_raw = cells[2].strip()
                m_q = re.search(r'\d+', qty_raw)
                if part_raw and m_q:
                    qty = int(m_q.group(0))
                    weight = cells[4].strip() if len(cells) > 4 else ""
                    size_x = cells[5].strip() if len(cells) > 5 else ""
                    size_y = cells[6].strip() if len(cells) > 6 else ""

                    parts_list.append({
                        "no": int(first_col),
                        "raw_name": part_raw,
                        "clean_name": clean_part_name(part_raw),
                        "quantity": qty,
                        "weight": weight,
                        "size_x": size_x,
                        "size_y": size_y,
                    })
                    found_in_block += 1
                    continue

            # Dacă am extras deja rânduri din acest bloc și am dat de finalul tabelului
            if found_in_block > 0:
                break

        search_pos = idx + 9
        # Dacă am găsit deja toate piesele așteptate, ne oprim
        if metadata["expected_part_qty"] and sum(p["quantity"] for p in parts_list) >= metadata["expected_part_qty"]:
            break

    return {
        "file_path": file_path,
        "file_name": os.path.basename(file_path),
        "metadata": metadata,
        "parts": parts_list,
    }


def _parse_docx_file(file_path: str) -> Dict[str, Any]:
    """Parsează un document .docx folosind python-docx."""
    import docx

    doc = docx.Document(file_path)
    metadata = {
        "subnest": "",
        "material": "",
        "sheet_size": "",
        "sheets_qty": 1,
        "expected_part_qty": None,
        "machine": "",
        "date": "",
    }

    # Extragere text din paragrafe pentru metadate
    full_text_lines = [p.text for p in doc.paragraphs if p.text.strip()]
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if cell.text.strip():
                    full_text_lines.append(cell.text.strip())
    joined_text = "\n".join(full_text_lines)

    m_sub = re.search(r'SubNest:\s*([^\r\n]+)', joined_text, re.IGNORECASE)
    if m_sub:
        metadata["subnest"] = m_sub.group(1).strip()

    m_part_qty = re.search(r'Part\s*qty\.?:\s*(\d+)', joined_text, re.IGNORECASE)
    if m_part_qty:
        metadata["expected_part_qty"] = int(m_part_qty.group(1))

    parts_list: List[Dict[str, Any]] = []

    for table in doc.tables:
        if not table.rows:
            continue
        header_cells = [c.text.strip().lower() for c in table.rows[0].cells]
        # Identificăm coloanele 'part name' și 'quantity'
        part_col_idx = None
        qty_col_idx = None
        no_col_idx = 0

        for idx, h_text in enumerate(header_cells):
            if "part name" in h_text or "denumire" in h_text or "piesa" in h_text or "part" == h_text:
                part_col_idx = idx
            elif "quantity" in h_text or "qty" in h_text or "cant" in h_text:
                qty_col_idx = idx
            elif h_text in ("no.", "no", "nr.", "nr"):
                no_col_idx = idx

        if part_col_idx is None:
            continue
        if qty_col_idx is None:
            qty_col_idx = part_col_idx + 1

        for row_idx, row in enumerate(table.rows[1:], start=1):
            cells = [c.text.strip() for c in row.cells]
            if len(cells) <= max(part_col_idx, qty_col_idx):
                continue
            part_raw = cells[part_col_idx].strip()
            qty_raw = cells[qty_col_idx].strip()
            if not part_raw or part_raw.lower() == "part name":
                continue
            m_q = re.search(r'\d+', qty_raw)
            if not m_q:
                continue
            no_val = row_idx
            if no_col_idx < len(cells) and cells[no_col_idx].strip().isdigit():
                no_val = int(cells[no_col_idx].strip())

            parts_list.append({
                "no": no_val,
                "raw_name": part_raw,
                "clean_name": clean_part_name(part_raw),
                "quantity": int(m_q.group(0)),
                "weight": cells[4].strip() if len(cells) > 4 else "",
                "size_x": cells[5].strip() if len(cells) > 5 else "",
                "size_y": cells[6].strip() if len(cells) > 6 else "",
            })

    return {
        "file_path": file_path,
        "file_name": os.path.basename(file_path),
        "metadata": metadata,
        "parts": parts_list,
    }


def parse_word_document(file_path: str) -> Dict[str, Any]:
    """
    Parsează un fișier Word (.doc sau .docx) și returnează dicționarul cu metadate și lista de piese.
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Fișierul nu există: {file_path}")

    with open(file_path, "rb") as f:
        raw_bytes = f.read()

    # 1. Dacă este arhivă ZIP (format .docx modern), folosim _parse_docx_file
    if raw_bytes[:4] == b'PK\x03\x04':
        return _parse_docx_file(file_path)

    # 2. Format binar OLE2 (.DOC Word 97-2003, generat de Metalix)
    # Încercăm mai întâi decodarea CP1252 (standardul Metalix pentru tabele ASCII/ANSI)
    text_cp1252 = raw_bytes.decode("cp1252", errors="ignore")
    if "Part name" in text_cp1252 or "part name" in text_cp1252.lower():
        res = _parse_binary_doc_text(text_cp1252, file_path)
        if res["parts"]:
            return res

    # 3. Fallback UTF-16LE (dacă documentul .DOC a fost salvat cu Unicode în Word)
    text_utf16 = raw_bytes.decode("utf-16le", errors="ignore")
    if "Part name" in text_utf16 or "part name" in text_utf16.lower():
        res = _parse_binary_doc_text(text_utf16, file_path)
        if res["parts"]:
            return res

    # Returnează rezultatul chiar dacă nu are piese (pentru raportare)
    return _parse_binary_doc_text(text_cp1252, file_path)


def aggregate_parts(parts: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Cumulează cantitățile pentru piesele care au aceeași denumire curățată (clean_name),
    păstrând ordinea primei apariții.
    """
    grouped: Dict[str, Dict[str, Any]] = {}
    for p in parts:
        key = p["clean_name"]
        if not key:
            continue
        if key not in grouped:
            grouped[key] = dict(p)
        else:
            grouped[key]["quantity"] += p["quantity"]
            if p["raw_name"] not in grouped[key]["raw_name"]:
                grouped[key]["raw_name"] += f", {p['raw_name']}"
    return list(grouped.values())


def write_import_excel(
    parts: List[Dict[str, Any]],
    output_path: str,
    start_row: int = 2,
    include_header: bool = False,
    style_borders: bool = True,
    empty_bordered_rows: int = 15,
    sheet_name: str = "Sheet1"
) -> str:
    """
    Generează un fișier Excel (.xlsx) identic ca structură și stil cu IMPORT.xlsx:
    - Coloana A: Denumirea piesei curățată (fără .dft, fără caractere speciale, fără _001, cu +1.1 păstrat)
    - Coloana B: Cantitatea (număr întreg)
    - Implicit începe de la rândul 2 (rândul 1 liber), font Verdana 10, centrat, chenar negru medium.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name[:31]

    font_data = Font(name="Verdana", size=10, bold=False)
    font_header = Font(name="Verdana", size=10, bold=True)
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    medium_side = Side(style="medium", color="000000")
    border_box = Border(left=medium_side, right=medium_side, top=medium_side, bottom=medium_side)

    # Setăm înălțimea rândului 1 dacă start_row == 2
    if start_row >= 2 and not include_header:
        ws.row_dimensions[1].height = 15.75

    current_row = start_row
    if include_header:
        header_row = 1 if start_row <= 2 else start_row - 1
        c_h1 = ws.cell(row=header_row, column=1, value="Denumire piesa")
        c_h2 = ws.cell(row=header_row, column=2, value="Cantitate")
        c_h1.font = font_header
        c_h2.font = font_header
        c_h1.alignment = align_center
        c_h2.alignment = align_center
        if style_borders:
            c_h1.border = border_box
            c_h2.border = border_box
        ws.row_dimensions[header_row].height = 22.0
        if current_row <= header_row:
            current_row = header_row + 1

    # Scriem piesele în coloanele A și B
    for p in parts:
        c_name = ws.cell(row=current_row, column=1, value=p["clean_name"])
        c_qty = ws.cell(row=current_row, column=2, value=int(p["quantity"]))

        c_name.font = font_data
        c_qty.font = font_data
        c_name.alignment = align_center
        c_qty.alignment = align_center

        if style_borders:
            c_name.border = border_box
            c_qty.border = border_box

        ws.row_dimensions[current_row].height = 24.0
        current_row += 1

    # Adăugăm câteva rânduri goale cu chenar sub tabel (exact ca în imaginea Excel atașată)
    if style_borders and empty_bordered_rows > 0:
        for _ in range(empty_bordered_rows):
            c_e1 = ws.cell(row=current_row, column=1, value=None)
            c_e2 = ws.cell(row=current_row, column=2, value=None)
            c_e1.font = font_data
            c_e2.font = font_data
            c_e1.alignment = align_center
            c_e2.alignment = align_center
            c_e1.border = border_box
            c_e2.border = border_box
            ws.row_dimensions[current_row].height = 15.75
            current_row += 1

    # Lățimi coloane (ajustate ca în imagine, dar suficient de încăpătoare)
    ws.column_dimensions["A"].width = 16.5
    ws.column_dimensions["B"].width = 10.5

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    wb.save(output_path)
    return output_path


def process_batch_documents(
    word_files: List[str],
    output_dir: str,
    mode: str = "individual",  # "individual" (1 Excel per Word) | "consolidated_single" | "consolidated_sheets"
    consolidated_filename: str = "IMPORT_CONSOLIDAT.xlsx",
    start_row: int = 2,
    include_header: bool = False,
    merge_duplicates: bool = False,
    progress_callback: Optional[Callable[[int, int, str], None]] = None
) -> Dict[str, Any]:
    """
    Procesează o listă de documente Word (.doc / .docx) și generează fișierul / fișierele Excel.
    Returnează un raport detaliat.
    """
    os.makedirs(output_dir, exist_ok=True)

    parsed_docs = []
    generated_files = []
    errors = []
    total_parts_rows = 0
    total_qty_sum = 0

    total = len(word_files)
    for idx, fpath in enumerate(word_files, start=1):
        fname = os.path.basename(fpath)
        if progress_callback:
            progress_callback(idx, total, f"Citire {fname} ({idx}/{total})...")
        try:
            doc_data = parse_word_document(fpath)
            parts = doc_data["parts"]
            if merge_duplicates:
                parts = aggregate_parts(parts)
                doc_data["parts"] = parts

            if not parts:
                errors.append(f"{fname}: Nu s-a găsit tabelul 'Parts' sau nu conține piese.")
                continue

            parsed_docs.append(doc_data)
            total_parts_rows += len(parts)
            total_qty_sum += sum(p["quantity"] for p in parts)

            if mode == "individual":
                stem = os.path.splitext(fname)[0]
                out_xlsx = os.path.join(output_dir, f"{stem}.xlsx")
                write_import_excel(
                    parts=parts,
                    output_path=out_xlsx,
                    start_row=start_row,
                    include_header=include_header,
                )
                generated_files.append(out_xlsx)
        except Exception as exc:
            errors.append(f"{fname}: Eroare la procesare - {exc}")

    # Moduri consolidate
    if parsed_docs and mode == "consolidated_single":
        all_parts = []
        for d in parsed_docs:
            all_parts.extend(d["parts"])
        if merge_duplicates:
            all_parts = aggregate_parts(all_parts)

        if not consolidated_filename.lower().endswith(".xlsx"):
            consolidated_filename += ".xlsx"
        out_xlsx = os.path.join(output_dir, consolidated_filename)
        write_import_excel(
            parts=all_parts,
            output_path=out_xlsx,
            start_row=start_row,
            include_header=include_header,
        )
        generated_files.append(out_xlsx)

    elif parsed_docs and mode == "consolidated_sheets":
        if not consolidated_filename.lower().endswith(".xlsx"):
            consolidated_filename += ".xlsx"
        out_xlsx = os.path.join(output_dir, consolidated_filename)
        wb = openpyxl.Workbook()
        # Eliminăm sheet-ul implicit la final după ce adăugăm sheet-urile documentelor
        default_ws = wb.active

        font_data = Font(name="Verdana", size=10)
        font_header = Font(name="Verdana", size=10, bold=True)
        align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
        medium_side = Side(style="medium", color="000000")
        border_box = Border(left=medium_side, right=medium_side, top=medium_side, bottom=medium_side)

        used_titles = set()
        for d in parsed_docs:
            stem = os.path.splitext(d["file_name"])[0]
            # Curățăm caracterele interzise în nume de sheet Excel
            clean_title = re.sub(r'[\\/*?:\[\]]', '_', stem)[:31] or "Sheet"
            base_title = clean_title
            counter = 2
            while clean_title.upper() in used_titles:
                suffix = f"_{counter}"
                clean_title = base_title[:31 - len(suffix)] + suffix
                counter += 1
            used_titles.add(clean_title.upper())

            ws = wb.create_sheet(title=clean_title)
            if start_row >= 2 and not include_header:
                ws.row_dimensions[1].height = 15.75

            curr = start_row
            if include_header:
                h_row = 1 if start_row <= 2 else start_row - 1
                ch1 = ws.cell(row=h_row, column=1, value="Denumire piesa")
                ch2 = ws.cell(row=h_row, column=2, value="Cantitate")
                ch1.font = font_header
                ch2.font = font_header
                ch1.alignment = align_center
                ch2.alignment = align_center
                ch1.border = border_box
                ch2.border = border_box
                ws.row_dimensions[h_row].height = 22.0
                if curr <= h_row:
                    curr = h_row + 1

            for p in d["parts"]:
                c1 = ws.cell(row=curr, column=1, value=p["clean_name"])
                c2 = ws.cell(row=curr, column=2, value=int(p["quantity"]))
                c1.font = font_data
                c2.font = font_data
                c1.alignment = align_center
                c2.alignment = align_center
                c1.border = border_box
                c2.border = border_box
                ws.row_dimensions[curr].height = 24.0
                curr += 1

            for _ in range(10):
                ce1 = ws.cell(row=curr, column=1, value=None)
                ce2 = ws.cell(row=curr, column=2, value=None)
                ce1.border = border_box
                ce2.border = border_box
                ws.row_dimensions[curr].height = 15.75
                curr += 1

            ws.column_dimensions["A"].width = 16.5
            ws.column_dimensions["B"].width = 10.5

        if len(wb.sheetnames) > 1:
            wb.remove(default_ws)
        wb.save(out_xlsx)
        generated_files.append(out_xlsx)

    return {
        "processed_docs": len(parsed_docs),
        "total_docs": total,
        "total_parts_rows": total_parts_rows,
        "total_qty_sum": total_qty_sum,
        "generated_files": generated_files,
        "errors": errors,
        "docs_data": parsed_docs,
    }

