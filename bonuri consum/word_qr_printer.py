from __future__ import annotations

import io
import os
import re
import shutil
import tempfile
import time
from pathlib import Path
from typing import Optional

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt
from PIL import Image

_OLE_DOC_MAGIC = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"
_ZIP_DOCX_MAGIC = b"PK\x03\x04"


def _is_ole_doc(path: Path) -> bool:
    """Verifică semnătura binară OLE2 (.doc vechi Word 97-2003)."""
    try:
        with path.open("rb") as f:
            return f.read(8) == _OLE_DOC_MAGIC
    except OSError:
        return False


def _is_zip_docx(path: Path) -> bool:
    """Verifică semnătura ZIP (.docx OpenXML)."""
    try:
        with path.open("rb") as f:
            return f.read(4) == _ZIP_DOCX_MAGIC
    except OSError:
        return False


def _convert_docx_to_doc_com(docx_path: Path, dst_doc_path: Path) -> None:
    """Convertește un fișier .docx în format binar Word 97-2003 (.doc) prin Word COM."""
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise RuntimeError(
            f"Salvarea în format original .DOC necesită pywin32 și Microsoft Word: {exc}"
        ) from exc

    word = None
    doc = None
    pythoncom.CoInitialize()
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(
            str(docx_path.resolve()),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
        )
        # 0 = wdFormatDocument (.doc binar Word 97-2003)
        doc.SaveAs2(str(dst_doc_path.resolve()), FileFormat=0)
        doc.Close(SaveChanges=False)
        doc = None
    finally:
        if doc is not None:
            try:
                doc.Close(SaveChanges=False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit(SaveChanges=False)
            except Exception:
                pass
            del word
        pythoncom.CoUninitialize()

    if not dst_doc_path.exists() or dst_doc_path.stat().st_size <= 0 or not _is_ole_doc(dst_doc_path):
        raise RuntimeError(
            f"Conversia Word COM nu a produs un fișier .doc binar valid: {dst_doc_path.name}"
        )


def _atomic_replace_file(src_file: Path, dst_file: Path) -> None:
    """Înlocuiește atomic fișierul destinație printr-un fișier temporar din același folder."""
    dst_file.parent.mkdir(parents=True, exist_ok=True)
    same_dir_tmp = dst_file.parent / f"_qr_tmp_{os.getpid()}_{dst_file.name}"
    try:
        shutil.copyfile(src_file, same_dir_tmp)
        try:
            os.replace(same_dir_tmp, dst_file)
        except PermissionError as exc:
            raise PermissionError(
                f"Nu s-a putut suprascrie '{dst_file.name}' deoarece fișierul este deschis în Word sau protejat la scriere. Închideți fișierul și reîncercați."
            ) from exc
    finally:
        try:
            if same_dir_tmp.exists():
                same_dir_tmp.unlink()
        except OSError:
            pass


def _normalize_cell_text(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip().lower()


def find_qr_cell(doc: Document):
    """
    Identifică celula destinație pentru codul QR din tabelul 'Laser parameters'.

    Caută antetul 'QR code' (preferând tabelul care conține și coloane precum
    'Gas', 'Lens', 'Laser table', 'Speed' sau 'Used power') și returnează
    celula imediat dedesubtul antetului.
    """
    candidates = []
    param_hints = {"gas", "lens", "laser table", "beam diameter", "speed", "used power"}

    for table in doc.tables:
        for r_idx, row in enumerate(table.rows):
            row_texts = [_normalize_cell_text(c.text) for c in row.cells]
            has_param_hint = any(
                any(hint in txt for hint in param_hints) for txt in row_texts
            )
            for c_idx, txt in enumerate(row_texts):
                if txt in {"qr code", "qrcode"} or "qr code" in txt:
                    score = (2 if has_param_hint else 0) + (1 if r_idx == 0 else 0)
                    candidates.append((score, table, r_idx, c_idx))

    if not candidates:
        raise ValueError(
            "Nu s-a găsit coloana cu antetul 'QR code' în tabelele documentului Word."
        )

    candidates.sort(key=lambda item: item[0], reverse=True)
    _, target_table, header_row_idx, col_idx = candidates[0]

    target_row_idx = header_row_idx + 1
    if target_row_idx >= len(target_table.rows):
        target_table.add_row()

    return target_table.rows[target_row_idx].cells[col_idx]


def insert_qr_into_docx(
    source_docx_path: str | Path,
    qr_image_path: str | Path,
    output_dir: str | Path = "output_print",
    output_filename: Optional[str] = None,
    qr_size_cm: float = 2.2,
    target_overwrite_path: Optional[str | Path] = None,
) -> Path:
    """
    Inserează imaginea PNG a codului QR în celula 'QR code' a documentului Word.

    Dacă target_overwrite_path este specificat, suprascrie atomic fișierul țintă
    (inclusiv reconversia în format binar .DOC prin Word COM dacă ținta este .doc).
    Altfel salvează o copie modificată în output_dir.
    """
    src_docx = Path(source_docx_path)
    qr_img = Path(qr_image_path)
    out_dir = Path(output_dir)

    if not src_docx.exists():
        raise FileNotFoundError(f"Documentul Word sursă nu există: {src_docx}")
    if not qr_img.exists():
        raise FileNotFoundError(f"Imaginea QR nu există: {qr_img}")

    with Image.open(qr_img) as img:
        w_px, h_px = img.size
    if w_px <= 0 or h_px <= 0:
        raise ValueError(f"Imagine QR invalidă (dimensiuni {w_px}x{h_px}): {qr_img}")

    if w_px >= h_px:
        width_cm = qr_size_cm
        height_cm = qr_size_cm * (h_px / w_px)
    else:
        height_cm = qr_size_cm
        width_cm = qr_size_cm * (w_px / h_px)

    doc = Document(io.BytesIO(src_docx.read_bytes()))
    cell = find_qr_cell(doc)

    # Curăță orice paragraf/text existent în celulă păstrând un singur paragraf centrat
    for p in cell.paragraphs[1:]:
        p._element.getparent().remove(p._element)
    para = cell.paragraphs[0]
    para.text = ""
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para.paragraph_format.space_before = Pt(2)
    para.paragraph_format.space_after = Pt(2)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    run = para.add_run()
    run.add_picture(str(qr_img), width=Cm(width_cm), height=Cm(height_cm))

    if target_overwrite_path is not None:
        target = Path(target_overwrite_path)
        wants_legacy_doc = (target.suffix.lower() == ".doc") and (
            not target.exists() or _is_ole_doc(target) or not _is_zip_docx(target)
        )
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp_path = Path(tmpdir)
            stage_docx = tmp_path / "qr_stage.docx"
            doc.save(str(stage_docx))
            if wants_legacy_doc:
                stage_doc = tmp_path / "qr_stage.doc"
                _convert_docx_to_doc_com(stage_docx, stage_doc)
                _atomic_replace_file(stage_doc, target)
            else:
                _atomic_replace_file(stage_docx, target)
        return target

    out_dir.mkdir(parents=True, exist_ok=True)
    out_name = output_filename or f"{src_docx.stem}_QR.docx"
    out_path = out_dir / out_name
    doc.save(str(out_path))
    return out_path


DEFAULT_PRINTER_OPTION = "Imprimanta implicită Windows"


def list_windows_printers() -> tuple[list[str], str]:
    """Returnează lista imprimantelor instalate în Windows și numele imprimantei implicite."""
    printers: list[str] = []
    default_name = ""
    try:
        import win32print

        try:
            default_name = str(win32print.GetDefaultPrinter() or "").strip()
        except Exception:
            default_name = ""

        raw = win32print.EnumPrinters(
            win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
        )
        seen: set[str] = set()
        for entry in raw:
            if len(entry) >= 3 and entry[2]:
                name = str(entry[2]).strip()
                if name and name not in seen:
                    seen.add(name)
                    printers.append(name)
    except Exception:
        return [], ""

    if default_name:
        if default_name in printers:
            printers.remove(default_name)
        printers.insert(0, default_name)

    return printers, default_name


def print_docx_windows(
    filepath: str | Path,
    delay_seconds: float = 3.0,
    printer_name: Optional[str] = None,
) -> None:
    """Trimite documentul .docx la imprimanta aleasă (sau implicită) din Windows."""
    import win32api

    target = Path(filepath).resolve()
    if not target.exists():
        raise FileNotFoundError(f"Fișierul pentru printare nu există: {target}")

    chosen = (printer_name or "").strip()
    if chosen and chosen != DEFAULT_PRINTER_OPTION:
        safe_printer = chosen.replace('"', "")
        rc = win32api.ShellExecute(0, "printto", str(target), f'"{safe_printer}"', ".", 0)
        if rc <= 32:
            raise RuntimeError(
                f"ShellExecute('printto', '{safe_printer}') a eșuat pentru '{target.name}' cu codul de eroare {rc}."
            )
    else:
        rc = win32api.ShellExecute(0, "print", str(target), None, ".", 0)
        if rc <= 32:
            raise RuntimeError(
                f"ShellExecute('print') a eșuat pentru '{target.name}' cu codul de eroare {rc}."
            )

    if delay_seconds > 0:
        time.sleep(delay_seconds)


def insert_qr_and_optionally_print(
    source_docx_path: str | Path,
    qr_image_path: str | Path,
    output_dir: str | Path = "output_print",
    output_filename: Optional[str] = None,
    qr_size_cm: float = 2.2,
    send_to_printer: bool = False,
    print_delay: float = 3.0,
    printer_name: Optional[str] = None,
    overwrite_original: bool = False,
    original_word_path: Optional[str | Path] = None,
) -> tuple[Path, bool]:
    """Inserează QR în documentul Word și, dacă este activat, îl trimite la imprimantă."""
    target_overwrite: Optional[Path] = None
    if overwrite_original:
        target_overwrite = (
            Path(original_word_path)
            if (original_word_path and str(original_word_path).strip())
            else Path(source_docx_path)
        )

    out_path = insert_qr_into_docx(
        source_docx_path=source_docx_path,
        qr_image_path=qr_image_path,
        output_dir=output_dir,
        output_filename=output_filename,
        qr_size_cm=qr_size_cm,
        target_overwrite_path=target_overwrite,
    )
    printed = False
    if send_to_printer:
        print_docx_windows(
            out_path,
            delay_seconds=print_delay,
            printer_name=printer_name,
        )
        printed = True
    return out_path, printed
