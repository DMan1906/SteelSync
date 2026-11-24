from __future__ import annotations

import json
import os
from dataclasses import asdict, fields, is_dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Optional, Type, TypeVar

from dotenv import load_dotenv

from browser_bot import Stage1PlanResult, Stage2PlanResult
from parser_docs import CuttingPlanRecord, plan_sort_key

load_dotenv(override=False)

APP_DIR = Path(__file__).resolve().parent

INPUT_MODE_LOCAL = "📁 Folder local"
INPUT_MODE_UPLOAD = "📤 Încărcare fișiere (Drag & Drop)"
INPUT_MODE_OPTIONS = [INPUT_MODE_LOCAL, INPUT_MODE_UPLOAD]

VALID_STAGES = {"idle", "parsed", "stage1_done", "stage2_done"}

T = TypeVar("T")


def get_settings_path() -> Path:
    """Returnează calea fișierului JSON pentru setările persistente."""
    override = os.environ.get("NORSTAL_SETTINGS_PATH", "").strip()
    if override:
        return Path(override)
    return APP_DIR / "user_settings.json"


def get_session_path() -> Path:
    """Returnează calea fișierului JSON pentru ultima sesiune de date."""
    override = os.environ.get("NORSTAL_SESSION_PATH", "").strip()
    if override:
        return Path(override)
    return APP_DIR / "work_tmp" / "last_session.json"


def _atomic_write_json(path: Path, payload: dict[str, Any]) -> None:
    """Scrie atomic un dicționar JSON UTF-8 folosind fișier temporar + os.replace."""
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = path.with_suffix(path.suffix + ".tmp")
    content = json.dumps(payload, ensure_ascii=False, indent=2)
    tmp_path.write_text(content, encoding="utf-8")
    os.replace(tmp_path, path)


def _coerce_bool(val: Any, default: bool) -> bool:
    if isinstance(val, bool):
        return val
    if isinstance(val, (int, float)):
        return bool(val)
    if isinstance(val, str):
        norm = val.strip().lower()
        if norm in {"true", "1", "yes", "da"}:
            return True
        if norm in {"false", "0", "no", "nu"}:
            return False
    return default


def _coerce_int(val: Any, default: int = 0, min_val: Optional[int] = None) -> int:
    try:
        parsed = int(float(val))
    except (TypeError, ValueError):
        parsed = default
    if min_val is not None:
        parsed = max(min_val, parsed)
    return parsed


def _coerce_float(val: Any, default: float, min_val: float, max_val: float) -> float:
    try:
        parsed = float(val)
    except (TypeError, ValueError):
        parsed = default
    return max(min_val, min(max_val, parsed))


def get_default_settings() -> dict[str, Any]:
    """Returnează setările implicite (citind și variabilele din .env)."""
    return {
        "intranet_url": os.getenv("NORSTAL_URL", "http://10.0.1.132/NORSTAL/"),
        "intranet_user": os.getenv("NORSTAL_USER", ""),
        "intranet_pass": os.getenv("NORSTAL_PASSWORD", ""),
        "visible_mode": True,
        "use_full_http_mode": False,
        "use_http_mode": False,
        "use_import_mode": False,
        "send_to_printer": False,
        "printer_name": "",
        "overwrite_original_word": True,
        "print_delay": 3.0,
        "output_dir_str": "output_print",
        "input_mode": INPUT_MODE_LOCAL,
        "word_dir_str": os.getenv("DEFAULT_WORD_DIR", r"d:\Programe plasma"),
        "excel_dir_str": os.getenv("DEFAULT_EXCEL_DIR", ""),
        "name_filter": "",
    }


def sanitize_settings(raw: Any) -> dict[str, Any]:
    """Validează și normalizează un dicționar de setări față de valorile implicite."""
    defaults = get_default_settings()
    src = raw if isinstance(raw, dict) else {}

    result: dict[str, Any] = {
        "intranet_url": str(src["intranet_url"]) if src.get("intranet_url") is not None else defaults["intranet_url"],
        "intranet_user": str(src["intranet_user"]) if src.get("intranet_user") is not None else defaults["intranet_user"],
        "intranet_pass": str(src["intranet_pass"]) if src.get("intranet_pass") is not None else defaults["intranet_pass"],
        "visible_mode": _coerce_bool(src.get("visible_mode"), defaults["visible_mode"]),
        "use_full_http_mode": _coerce_bool(src.get("use_full_http_mode"), defaults["use_full_http_mode"]),
        "use_http_mode": _coerce_bool(src.get("use_http_mode"), defaults["use_http_mode"]),
        "use_import_mode": _coerce_bool(src.get("use_import_mode"), defaults["use_import_mode"]),
        "send_to_printer": _coerce_bool(src.get("send_to_printer"), defaults["send_to_printer"]),
        "printer_name": str(src["printer_name"]).strip() if src.get("printer_name") is not None else defaults["printer_name"],
        "overwrite_original_word": _coerce_bool(
            src.get("overwrite_original_word"), defaults["overwrite_original_word"]
        ),
        "print_delay": _coerce_float(src.get("print_delay"), defaults["print_delay"], 0.0, 15.0),
        "output_dir_str": str(src["output_dir_str"]) if src.get("output_dir_str") is not None else defaults["output_dir_str"],
        "input_mode": (
            str(src["input_mode"])
            if src.get("input_mode") in INPUT_MODE_OPTIONS
            else defaults["input_mode"]
        ),
        "word_dir_str": str(src["word_dir_str"]) if src.get("word_dir_str") is not None else defaults["word_dir_str"],
        "excel_dir_str": str(src["excel_dir_str"]) if src.get("excel_dir_str") is not None else defaults["excel_dir_str"],
        "name_filter": str(src["name_filter"]) if src.get("name_filter") is not None else defaults["name_filter"],
    }

    if result["use_full_http_mode"]:
        result["use_http_mode"] = False
        result["use_import_mode"] = False
    elif result["use_http_mode"] and result["use_import_mode"]:
        result["use_import_mode"] = False

    return result


def load_user_settings(path: Optional[Path] = None) -> dict[str, Any]:
    """Încarcă setările utilizatorului din JSON, cu fallback sigur la valorile implicite."""
    target = path or get_settings_path()
    if not target.exists():
        return sanitize_settings({})
    try:
        raw = json.loads(target.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError, ValueError):
        return sanitize_settings({})
    return sanitize_settings(raw)


def save_user_settings(settings: dict[str, Any], path: Optional[Path] = None) -> dict[str, Any]:
    """Validează și salvează atomic setările utilizatorului în JSON."""
    clean = sanitize_settings(settings)
    target = path or get_settings_path()
    _atomic_write_json(target, clean)
    return clean


def clear_user_settings(path: Optional[Path] = None) -> None:
    """Șterge fișierul cu setările salvate, dacă există."""
    target = path or get_settings_path()
    target.unlink(missing_ok=True)


def get_empty_session() -> dict[str, Any]:
    """Returnează structura goală pentru datele de sesiune."""
    return {
        "stage": "idle",
        "plans": [],
        "active_plans": [],
        "plan_overrides": {},
        "stage1_results": [],
        "stage2_results": [],
        "discovered_word": [],
        "discovered_excel": [],
        "selected_word_names": [],
        "selected_s2_names": None,
        "bot_logs": [],
        "updated_at": "",
    }


def _serialize_dc_list(items: list[Any]) -> list[dict[str, Any]]:
    serialized: list[dict[str, Any]] = []
    for item in items:
        if is_dataclass(item):
            serialized.append(asdict(item))
        elif isinstance(item, dict):
            serialized.append(dict(item))
    return serialized


def _deserialize_dataclasses(raw_items: Any, cls: Type[T]) -> list[T]:
    """Deserializare defensivă pentru liste de dataclass-uri (ignoră cheile necunoscute)."""
    if not isinstance(raw_items, list):
        return []
    valid_names = {f.name for f in fields(cls)}
    items: list[T] = []
    for entry in raw_items:
        if is_dataclass(entry) and isinstance(entry, cls):
            items.append(entry)
            continue
        if not isinstance(entry, dict):
            continue
        filtered = {k: v for k, v in entry.items() if k in valid_names}
        try:
            items.append(cls(**filtered))
        except TypeError:
            continue
    return items


def _deserialize_plans(raw_items: Any) -> list[CuttingPlanRecord]:
    plans = _deserialize_dataclasses(raw_items, CuttingPlanRecord)
    missing_word_msg = "Fișierul Word convertit nu mai există pe disc — re-analizați fișierele."
    missing_excel_msg = "Fișierul .xlsx asociat nu mai există pe disc."

    for rec in plans:
        rec.warnings = [str(w) for w in (rec.warnings or [])]
        rec.errors = [str(e) for e in (rec.errors or [])]
        rec.parts_count = _coerce_int(rec.parts_count, default=0, min_val=0)
        if rec.part_weight_int is not None:
            rec.part_weight_int = _coerce_int(rec.part_weight_int, default=0, min_val=0)

        if not rec.working_docx_path or not Path(rec.working_docx_path).exists():
            if missing_word_msg not in rec.warnings:
                rec.warnings.append(missing_word_msg)
            rec.is_ready = False

        if rec.excel_found and (not rec.excel_path or not Path(rec.excel_path).exists()):
            rec.excel_found = False
            rec.is_ready = False
            if missing_excel_msg not in rec.warnings:
                rec.warnings.append(missing_excel_msg)

    return plans


def _normalize_plan_overrides(raw_overrides: Any) -> dict[str, dict[str, Any]]:
    if not isinstance(raw_overrides, dict):
        return {}
    normalized: dict[str, dict[str, Any]] = {}
    for raw_name, raw_val in raw_overrides.items():
        if not isinstance(raw_val, dict):
            continue
        plan_name = str(raw_name)
        cutting = str(raw_val.get("cutting_type") or "Plasma").strip()
        if cutting not in {"Plasma", "Laser"}:
            cutting = "Plasma"
        normalized[plan_name] = {
            "include": _coerce_bool(raw_val.get("include"), True),
            "cutting_type": cutting,
            "project_query": str(raw_val.get("project_query") or "").strip(),
            "part_weight_int": _coerce_int(raw_val.get("part_weight_int"), default=0, min_val=0),
        }
    return normalized


def build_session_payload(
    *,
    stage: str,
    plans: list[CuttingPlanRecord],
    active_plans: list[CuttingPlanRecord],
    plan_overrides: dict[str, dict[str, Any]],
    stage1_results: list[Stage1PlanResult],
    stage2_results: list[Stage2PlanResult],
    discovered_word: list[Any],
    discovered_excel: list[Any],
    selected_word_names: list[str],
    selected_s2_names: Optional[list[str]],
    bot_logs: list[str],
    updated_at: str = "",
) -> dict[str, Any]:
    """Construiește payload-ul JSON serializabil pentru starea curentă a sesiunii."""
    norm_stage = stage if stage in VALID_STAGES else "idle"
    if not plans:
        norm_stage = "idle"

    return {
        "stage": norm_stage,
        "plans": _serialize_dc_list(plans),
        "active_plans": _serialize_dc_list(active_plans),
        "plan_overrides": _normalize_plan_overrides(plan_overrides),
        "stage1_results": _serialize_dc_list(stage1_results),
        "stage2_results": _serialize_dc_list(stage2_results),
        "discovered_word": [str(p) for p in discovered_word],
        "discovered_excel": [str(p) for p in discovered_excel],
        "selected_word_names": [str(name) for name in selected_word_names],
        "selected_s2_names": [str(name) for name in selected_s2_names] if selected_s2_names is not None else None,
        "bot_logs": [str(line) for line in (bot_logs or [])][-300:],
        "updated_at": str(updated_at or ""),
    }


def session_fingerprint(payload: dict[str, Any]) -> str:
    """Generează o amprentă JSON canonică ignorând updated_at pentru detecția schimbărilor."""
    comparable = {k: v for k, v in payload.items() if k != "updated_at"}
    return json.dumps(comparable, ensure_ascii=False, sort_keys=True)


def load_session_data(path: Optional[Path] = None) -> dict[str, Any]:
    """Încarcă ultima sesiune salvată, verificând existența fișierelor pe disc."""
    empty = get_empty_session()
    target = path or get_session_path()
    if not target.exists():
        return empty
    try:
        raw = json.loads(target.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError, ValueError):
        return empty
    if not isinstance(raw, dict):
        return empty

    plans = sorted(
        _deserialize_plans(raw.get("plans")),
        key=lambda p: plan_sort_key(p.plan_name),
    )
    active_plans = sorted(
        _deserialize_plans(raw.get("active_plans")),
        key=lambda p: plan_sort_key(p.plan_name),
    )
    stage1_results = _deserialize_dataclasses(raw.get("stage1_results"), Stage1PlanResult)
    stage2_results = _deserialize_dataclasses(raw.get("stage2_results"), Stage2PlanResult)

    raw_stage = raw.get("stage", "idle")
    stage = raw_stage if raw_stage in VALID_STAGES else "idle"
    if not plans:
        stage = "idle"

    disc_w = sorted(
        [
            Path(p)
            for p in (raw.get("discovered_word") or [])
            if isinstance(p, str) and p.strip() and Path(p).exists()
        ],
        key=plan_sort_key,
    )
    disc_e = sorted(
        [
            Path(p)
            for p in (raw.get("discovered_excel") or [])
            if isinstance(p, str) and p.strip() and Path(p).exists()
        ],
        key=plan_sort_key,
    )

    valid_w_names = {p.name for p in disc_w}
    raw_sel_w = raw.get("selected_word_names")
    if isinstance(raw_sel_w, list):
        selected_word_names = sorted(
            [str(n) for n in raw_sel_w if str(n) in valid_w_names],
            key=plan_sort_key,
        )
    else:
        selected_word_names = []

    known_plan_names = {p.plan_name for p in plans} | {p.plan_name for p in active_plans}
    raw_sel_s2 = raw.get("selected_s2_names")
    if isinstance(raw_sel_s2, list):
        selected_s2_names: Optional[list[str]] = sorted(
            [str(n) for n in raw_sel_s2 if str(n) in known_plan_names],
            key=plan_sort_key,
        )
    else:
        selected_s2_names = None

    raw_logs = raw.get("bot_logs")
    bot_logs = [str(line) for line in raw_logs][-300:] if isinstance(raw_logs, list) else []

    return {
        "stage": stage,
        "plans": plans,
        "active_plans": active_plans,
        "plan_overrides": _normalize_plan_overrides(raw.get("plan_overrides")),
        "stage1_results": stage1_results,
        "stage2_results": stage2_results,
        "discovered_word": disc_w,
        "discovered_excel": disc_e,
        "selected_word_names": selected_word_names,
        "selected_s2_names": selected_s2_names,
        "bot_logs": bot_logs,
        "updated_at": str(raw.get("updated_at") or ""),
    }


def save_session_data(payload: dict[str, Any], path: Optional[Path] = None) -> None:
    """Salvează atomic starea sesiunii cu marcaj temporal updated_at."""
    target = path or get_session_path()
    to_write = dict(payload)
    to_write["updated_at"] = datetime.now().isoformat(timespec="seconds")
    _atomic_write_json(target, to_write)


def clear_session_data(path: Optional[Path] = None) -> None:
    """Șterge fișierul cu ultima sesiune salvată, dacă există."""
    target = path or get_session_path()
    target.unlink(missing_ok=True)
