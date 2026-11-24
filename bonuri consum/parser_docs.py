from __future__ import annotations

import logging
import re
import shutil
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable, Optional

from docx import Document

logger = logging.getLogger(__name__)

OLE_DOC_MAGIC = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"
ZIP_DOCX_MAGIC = b"PK\x03\x04"

MATERIAL_TOKENS = {
    "steel",
    "inox",
    "stainless",
    "alu",
    "aluminium",
    "aluminum",
    "s235",
    "s355",
    "s275",
    "dc01",
    "hardox",
    "galv",
    "cupru",
    "alama",
    "brass",
}

SUBNEST_RE = re.compile(r"SubNest\s*:\s*([^\s,;]+)", re.IGNORECASE)
PART_WEIGHT_RE = re.compile(r"Part\s*weight\s*:\s*([0-9]+(?:[.,][0-9]+)?)", re.IGNORECASE)
MACHINE_RE = re.compile(r"Machine\s*:\s*(.+)", re.IGNORECASE)
LASER_RE = re.compile(r"Laser\s*:\s*(.+)", re.IGNORECASE)
TRAILING_NUM_RE = re.compile(r"(\d+)\s*$")
THICKNESS_TOKEN_RE = re.compile(r"^\d+(?:[.,]\d+)?(?:mm)?$", re.IGNORECASE)
WORD_EXCEL_EXT_RE = re.compile(r"\.(?:docx?|xlsx)$", re.IGNORECASE)
NATURAL_SPLIT_RE = re.compile(r"(\d+)")


@dataclass
class CuttingPlanRecord:
    """Reprezintă datele extrase dintr-un plan de croi Word și fișierul Excel asociat."""

    plan_name: str
    subnest_name: Optional[str]
    name_validated: bool
    part_weight_raw: Optional[float]
    part_weight_int: Optional[int]
    cutting_type: str
    machine_info: str
    project_query: str
    source_word_path: str
    working_docx_path: str
    excel_path: Optional[str]
    excel_found: bool
    parts_count: int = 0
    warnings: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)
    is_ready: bool = False


def is_ole_doc(path: Path) -> bool:
    """Verifică semnătura binară OLE2 (.doc vechi Word 97-2003)."""
    try:
        with path.open("rb") as f:
            return f.read(8) == OLE_DOC_MAGIC
    except OSError:
        return False


def is_zip_docx(path: Path) -> bool:
    """Verifică semnătura ZIP (.docx OpenXML)."""
    try:
        with path.open("rb") as f:
            return f.read(4) == ZIP_DOCX_MAGIC
    except OSError:
        return False


def _cached_docx_path(src_path: Path, cache_dir: Path) -> Path:
    stat = src_path.stat()
    safe_stem = re.sub(r"[^A-Za-z0-9._-]+", "_", src_path.stem)
    return cache_dir / f"{safe_stem}__{stat.st_size}_{stat.st_mtime_ns}.docx"


def convert_doc_batch_to_docx(
    doc_paths: Iterable[Path],
    cache_dir: Path,
) -> tuple[dict[Path, Path], dict[Path, str]]:
    """
    Pregătește fișierele .doc/.docx pentru python-docx.

    - Fișierele .docx native sunt folosite direct.
    - Fișierele OLE2 .DOC sunt convertite o singură dată într-o instanță Word COM
      și salvate în cache_dir (cheie: stem + dimensiune + mtime_ns).
    """
    cache_dir.mkdir(parents=True, exist_ok=True)
    resolved_map: dict[Path, Path] = {}
    errors_map: dict[Path, str] = {}
    to_convert: list[tuple[Path, Path]] = []

    for raw_path in doc_paths:
        src = Path(raw_path)
        if not src.exists():
            errors_map[src] = f"Fișierul nu există: {src}"
            continue

        if is_zip_docx(src):
            if src.suffix.lower() == ".docx":
                resolved_map[src] = src
            else:
                cached = _cached_docx_path(src, cache_dir)
                if not cached.exists():
                    shutil.copy2(src, cached)
                resolved_map[src] = cached
            continue

        if is_ole_doc(src) or src.suffix.lower() == ".doc":
            cached = _cached_docx_path(src, cache_dir)
            if cached.exists() and cached.stat().st_size > 0 and is_zip_docx(cached):
                resolved_map[src] = cached
            else:
                to_convert.append((src, cached))
            continue

        errors_map[src] = f"Format Word nesuportat: {src.name}"

    if not to_convert:
        return resolved_map, errors_map

    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        for src, _ in to_convert:
            errors_map[src] = f"Conversia .DOC necesită pywin32 și Microsoft Word: {exc}"
        return resolved_map, errors_map

    word = None
    pythoncom.CoInitialize()
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0  # wdAlertsNone

        for src, dst in to_convert:
            doc = None
            tmp_dst = dst.with_suffix(".tmp.docx")
            try:
                if tmp_dst.exists():
                    tmp_dst.unlink()
                doc = word.Documents.Open(
                    str(src.resolve()),
                    ConfirmConversions=False,
                    ReadOnly=True,
                    AddToRecentFiles=False,
                )
                # 16 = wdFormatXMLDocument (.docx)
                doc.SaveAs2(str(tmp_dst.resolve()), FileFormat=16)
                doc.Close(SaveChanges=False)
                doc = None

                if is_zip_docx(tmp_dst):
                    if dst.exists():
                        dst.unlink()
                    tmp_dst.replace(dst)
                    resolved_map[src] = dst
                else:
                    errors_map[src] = f"Conversia Word COM nu a produs un .docx valid pentru {src.name}."
            except Exception as exc:
                errors_map[src] = f"Eroare conversie .DOC -> .docx ({src.name}): {exc}"
            finally:
                if doc is not None:
                    try:
                        doc.Close(SaveChanges=False)
                    except Exception:
                        pass
                if tmp_dst.exists():
                    try:
                        tmp_dst.unlink()
                    except OSError:
                        pass
    except Exception as exc:
        for src, _ in to_convert:
            if src not in resolved_map and src not in errors_map:
                errors_map[src] = f"Nu s-a putut inițializa Microsoft Word COM: {exc}"
    finally:
        if word is not None:
            try:
                word.Quit(SaveChanges=False)
            except Exception:
                pass
            del word
        pythoncom.CoUninitialize()

    return resolved_map, errors_map


def derive_project_query(plan_name: str) -> tuple[str, Optional[str]]:
    """
    Deduce prefixul comenzii din numele planului.

    Exemplu:
      1242-49-1-Steel-6-4 -> '1242 - 49.1'
      1009-01-Steel-10-23 -> '1009 - 01' (fallback 2 segmente când al 3-lea e material)
    """
    parts = [p.strip() for p in plan_name.split("-") if p.strip()]
    if len(parts) >= 3 and parts[2].lower() not in MATERIAL_TOKENS:
        return f"{parts[0]} - {parts[1]}.{parts[2]}", None
    if len(parts) >= 2:
        return (
            f"{parts[0]} - {parts[1]}",
            f"Cod comandă dedus pe 2 segmente ('{parts[0]} - {parts[1]}') deoarece segmentul 3 ('{parts[2] if len(parts) > 2 else ''}') este material/lipsă.",
        )
    return plan_name, f"Nu s-au găsit cel puțin 2 segmente separate de '-' în '{plan_name}'."


def _natural_sort_tokens(text: str) -> tuple[tuple[int, int | str], ...]:
    """Descompune un șir în token-uri tipate (număr/text) pentru sortare naturală sigură."""
    tokens: list[tuple[int, int | str]] = []
    for part in NATURAL_SPLIT_RE.split(text):
        if not part:
            continue
        if part.isdigit():
            tokens.append((0, int(part)))
        else:
            tokens.append((1, part.lower()))
    return tuple(tokens)


def plan_sort_key(
    name_or_path: str | Path,
) -> tuple[
    tuple[tuple[int, int | str], ...],
    int,
    int,
    tuple[tuple[int, int | str], ...],
]:
    """
    Cheie de sortare pentru planuri Metalix: grupează după prefixul comenzii și
    ordonează crescător după ultimul număr din numele planului (ex: -8-1 înainte de -10-2).
    """
    if isinstance(name_or_path, Path):
        stem = name_or_path.stem.strip()
    else:
        stem = WORD_EXCEL_EXT_RE.sub("", str(name_or_path).strip())

    trailing_match = TRAILING_NUM_RE.search(stem)
    has_no_num = 0 if trailing_match else 1
    last_num = int(trailing_match.group(1)) if trailing_match else 0

    parts = [p.strip() for p in stem.split("-") if p.strip()]
    group_prefix: Optional[str] = None
    for idx in range(1, len(parts)):
        if parts[idx].lower() in MATERIAL_TOKENS:
            group_prefix = "-".join(parts[:idx])
            break

    if (
        group_prefix is None
        and len(parts) >= 4
        and parts[-1].isdigit()
        and THICKNESS_TOKEN_RE.match(parts[-2])
    ):
        group_prefix = "-".join(parts[:-3])

    if group_prefix is None:
        group_prefix = stem[: trailing_match.start()].rstrip("-_ ") if trailing_match else stem

    return (
        _natural_sort_tokens(group_prefix),
        has_no_num,
        last_num,
        _natural_sort_tokens(stem),
    )


def _collect_document_lines(doc: Document) -> tuple[list[str], list[str]]:
    """Colectează liniile de text din paragrafe și din toate tabelele documentului."""
    lines: list[str] = []
    joined_rows: list[str] = []

    for para in doc.paragraphs:
        txt = re.sub(r"\s+", " ", para.text).strip()
        if txt:
            lines.append(txt)

    for table in doc.tables:
        for row in table.rows:
            seen_tcs = set()
            row_cells: list[str] = []
            for cell in row.cells:
                if cell._tc in seen_tcs:
                    continue
                seen_tcs.add(cell._tc)
                cell_txt = re.sub(r"\s+", " ", cell.text).strip()
                if cell_txt:
                    row_cells.append(cell_txt)
                    for subline in cell.text.splitlines():
                        clean_sub = re.sub(r"\s+", " ", subline).strip()
                        if clean_sub:
                            lines.append(clean_sub)
            if row_cells:
                joined_rows.append(" ".join(row_cells))

    return lines, joined_rows


def read_parts_from_excel(excel_path: str | Path) -> list[tuple[str, int]]:
    """
    Citește perechile (Denumire piesă, Cantitate) dintr-un fișier .xlsx Metalix.

    Ignoră rândurile goale, capul de tabel (ex: 'Denumire piesa' / 'Cantitatea')
    și eventualele rânduri de total.
    """
    import openpyxl

    path = Path(excel_path)
    if not path.exists():
        raise FileNotFoundError(f"Fișierul Excel nu există: {path}")

    wb = openpyxl.load_workbook(str(path), data_only=True, read_only=True)
    parts: list[tuple[str, int]] = []
    try:
        ws = wb.active
        header_tokens = {
            "denumire piesa",
            "denumire piesă",
            "piesa",
            "part name",
            "part",
        }
        qty_header_tokens = {"cantitatea", "cantitate", "cant.", "cant", "qty", "quantity"}

        for row in ws.iter_rows(values_only=True):
            if not row or len(row) < 2:
                continue
            part_raw, qty_raw = row[0], row[1]
            if part_raw is None or qty_raw is None:
                continue

            part_name = str(part_raw).strip()
            if not part_name:
                continue
            part_low = part_name.lower()
            if part_low in header_tokens or part_low.startswith("total"):
                continue

            if isinstance(qty_raw, str):
                qty_s = qty_raw.strip().replace(",", ".")
                if qty_s.lower() in qty_header_tokens or not qty_s:
                    continue
                try:
                    qty_val = int(round(float(qty_s)))
                except ValueError:
                    continue
            elif isinstance(qty_raw, (int, float)):
                qty_val = int(round(float(qty_raw)))
            else:
                continue

            if qty_val > 0:
                parts.append((part_name, qty_val))
    finally:
        wb.close()

    return parts


def parse_single_docx(
    working_docx_path: Path,
    source_word_path: Path,
    excel_map: dict[str, Path],
) -> CuttingPlanRecord:
    """Extrage metadatele unui plan de croi dintr-un document .docx și verifică perechea .xlsx."""
    plan_name = source_word_path.stem.strip()
    warnings: list[str] = []
    errors: list[str] = []

    project_query, query_warning = derive_project_query(plan_name)
    if query_warning:
        warnings.append(query_warning)

    excel_path_obj = excel_map.get(plan_name.lower())
    excel_found = bool(excel_path_obj and excel_path_obj.exists())
    parts_count = 0
    if not excel_found:
        warnings.append("Lipsește fișierul Excel (.xlsx) cu același nume de bază.")
    elif excel_path_obj is not None:
        try:
            parts_count = len(read_parts_from_excel(excel_path_obj))
            if parts_count == 0:
                warnings.append("Fișierul Excel nu conține rânduri cu piese valide.")
        except Exception as exc:
            warnings.append(f"Nu s-au putut citi piesele din Excel: {exc}")

    subnest_name: Optional[str] = None
    part_weight_raw: Optional[float] = None
    part_weight_int: Optional[int] = None
    machine_val: Optional[str] = None
    laser_val: Optional[str] = None

    try:
        doc = Document(str(working_docx_path))
        lines, joined_rows = _collect_document_lines(doc)
        search_pool = joined_rows + lines + ["\n".join(lines)]

        for candidate in search_pool:
            m = SUBNEST_RE.search(candidate)
            if m:
                subnest_name = m.group(1).strip().rstrip(",;")
                break

        for candidate in search_pool:
            m = PART_WEIGHT_RE.search(candidate)
            if m:
                num_str = m.group(1).replace(",", ".")
                part_weight_raw = float(num_str)
                part_weight_int = int(round(part_weight_raw))
                break

        for candidate in lines + joined_rows:
            if machine_val is None:
                m_mach = MACHINE_RE.search(candidate)
                if m_mach:
                    val = m_mach.group(1).strip()
                    if val and not val.lower().startswith("laser:"):
                        machine_val = val
            if laser_val is None:
                m_las = LASER_RE.search(candidate)
                if m_las:
                    val = m_las.group(1).strip()
                    # Evită potrivirea cu antetul tabelului "Laser table" / "Laser parameters"
                    if val and not any(
                        val.lower().startswith(skip) for skip in ("table", "parameters")
                    ):
                        laser_val = val
    except Exception as exc:
        errors.append(f"Eroare la citirea documentului Word: {exc}")

    if subnest_name is None:
        name_validated = False
        warnings.append("Nu s-a găsit câmpul 'SubNest:' în textul documentului.")
    else:
        name_validated = subnest_name.lower() == plan_name.lower()
        if not name_validated:
            warnings.append(
                f"SubNest din document ('{subnest_name}') diferă de numele fișierului ('{plan_name}')."
            )

    if part_weight_int is None and not errors:
        errors.append("Nu s-a găsit valoarea 'Part weight:' în document.")

    machine_parts = []
    if machine_val:
        machine_parts.append(f"Machine: {machine_val}")
    if laser_val:
        machine_parts.append(f"Laser: {laser_val}")
    machine_info = " | ".join(machine_parts) if machine_parts else "Nedetectat"

    if not machine_parts and not errors:
        warnings.append("Nu s-au detectat câmpurile 'Machine:' / 'Laser:'; debitarea a fost setată implicit pe 'Plasma'.")

    # În intranetul NORSTAL, atât debitările de Laser cât și cele de Plasma
    # se încadrează sub opțiunea 'Plasma' din dropdown-ul 'Debitare'.
    cutting_type = "Plasma"
    is_ready = bool(
        excel_found
        and part_weight_int is not None
        and project_query
        and len(errors) == 0
    )

    return CuttingPlanRecord(
        plan_name=plan_name,
        subnest_name=subnest_name,
        name_validated=name_validated,
        part_weight_raw=part_weight_raw,
        part_weight_int=part_weight_int,
        cutting_type=cutting_type,
        machine_info=machine_info,
        project_query=project_query,
        source_word_path=str(source_word_path),
        working_docx_path=str(working_docx_path),
        excel_path=str(excel_path_obj) if excel_path_obj else None,
        excel_found=excel_found,
        parts_count=parts_count,
        warnings=warnings,
        errors=errors,
        is_ready=is_ready,
    )


def parse_files(
    word_paths: Iterable[Path],
    excel_paths: Iterable[Path],
    cache_dir: Path | str = "work_tmp/converted_docx",
) -> list[CuttingPlanRecord]:
    """Convertește lotul de fișiere Word (.doc/.docx), le parsează și le împerechează cu .xlsx."""
    word_list = sorted({Path(p) for p in word_paths}, key=plan_sort_key)
    excel_map: dict[str, Path] = {}
    for ep in excel_paths:
        p = Path(ep)
        if p.exists() and p.suffix.lower() == ".xlsx" and not p.name.startswith("~$"):
            excel_map[p.stem.lower()] = p

    converted_map, conv_errors = convert_doc_batch_to_docx(word_list, Path(cache_dir))

    records: list[CuttingPlanRecord] = []
    for src_word in word_list:
        if src_word in conv_errors:
            plan_name = src_word.stem.strip()
            project_query, _ = derive_project_query(plan_name)
            excel_p = excel_map.get(plan_name.lower())
            records.append(
                CuttingPlanRecord(
                    plan_name=plan_name,
                    subnest_name=None,
                    name_validated=False,
                    part_weight_raw=None,
                    part_weight_int=None,
                    cutting_type="Laser",
                    machine_info="Eroare conversie",
                    project_query=project_query,
                    source_word_path=str(src_word),
                    working_docx_path="",
                    excel_path=str(excel_p) if excel_p else None,
                    excel_found=bool(excel_p and excel_p.exists()),
                    warnings=[],
                    errors=[conv_errors[src_word]],
                    is_ready=False,
                )
            )
            continue

        working_docx = converted_map[src_word]
        records.append(parse_single_docx(working_docx, src_word, excel_map))

    return records


def discover_folder_files(
    word_dir: Path | str,
    excel_dir: Optional[Path | str] = None,
    name_filter: str = "",
) -> tuple[list[Path], list[Path]]:
    """Scanează un folder local (non-recursiv) pentru fișiere .doc/.docx și .xlsx."""
    w_dir = Path(word_dir)
    e_dir = Path(excel_dir) if excel_dir and str(excel_dir).strip() else w_dir
    filt = name_filter.strip().lower()

    word_files: list[Path] = []
    if w_dir.is_dir():
        for entry in w_dir.iterdir():
            if not entry.is_file() or entry.name.startswith("~$"):
                continue
            if entry.suffix.lower() in {".doc", ".docx"}:
                if not filt or filt in entry.stem.lower():
                    word_files.append(entry)

    excel_files: list[Path] = []
    if e_dir.is_dir():
        for entry in e_dir.iterdir():
            if not entry.is_file() or entry.name.startswith("~$"):
                continue
            if entry.suffix.lower() == ".xlsx":
                excel_files.append(entry)

    word_files.sort(key=plan_sort_key)
    excel_files.sort(key=plan_sort_key)
    return word_files, excel_files
