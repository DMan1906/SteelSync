from __future__ import annotations

import html
import importlib
import inspect
import json
import os
import time
from dataclasses import replace
from datetime import datetime
from pathlib import Path
from typing import Any, Optional

import pandas as pd
import streamlit as st
from dotenv import load_dotenv

import browser_bot
import parser_docs
import settings_store
import word_qr_printer


def _reload_local_module_if_needed(mod: Any, force: bool = False) -> tuple[Any, bool]:
    """Reîncarcă un modul local dacă lipsește un simbol așteptat sau i s-a modificat mtime-ul."""
    mod_file = getattr(mod, "__file__", None)
    try:
        current_mtime: Optional[float] = (
            os.path.getmtime(mod_file) if mod_file and os.path.exists(mod_file) else None
        )
    except OSError:
        current_mtime = None

    loaded_mtime = getattr(mod, "_st_loaded_mtime", None)
    needs_reload = force or (
        loaded_mtime is not None
        and current_mtime is not None
        and current_mtime != loaded_mtime
    )
    if needs_reload:
        mod = importlib.reload(mod)
        try:
            current_mtime = (
                os.path.getmtime(mod_file) if mod_file and os.path.exists(mod_file) else None
            )
        except OSError:
            current_mtime = None
        if current_mtime is not None:
            setattr(mod, "_st_loaded_mtime", current_mtime)
        return mod, True

    if loaded_mtime is None and current_mtime is not None:
        setattr(mod, "_st_loaded_mtime", current_mtime)
    return mod, False


# Reîncărcare în ordine topologică a dependențelor pentru hot-reload Streamlit
_parser_stale = not hasattr(parser_docs, "plan_sort_key")
parser_docs, _parser_reloaded = _reload_local_module_if_needed(
    parser_docs, force=_parser_stale
)

_word_stale = (
    not hasattr(word_qr_printer, "DEFAULT_PRINTER_OPTION")
    or not hasattr(word_qr_printer, "list_windows_printers")
    or not hasattr(word_qr_printer, "insert_qr_and_optionally_print")
    or "printer_name"
    not in inspect.signature(word_qr_printer.insert_qr_and_optionally_print).parameters
    or "overwrite_original"
    not in inspect.signature(word_qr_printer.insert_qr_and_optionally_print).parameters
)
word_qr_printer, _word_reloaded = _reload_local_module_if_needed(
    word_qr_printer, force=_word_stale
)

_bot_cls = getattr(browser_bot, "NorstalBrowserBot", None)
_bot_stale = (
    _parser_reloaded
    or _word_reloaded
    or _bot_cls is None
    or not hasattr(browser_bot, "_is_uploaded_temp_word_path")
    or not hasattr(_bot_cls, "start_stage1_async")
    or not hasattr(_bot_cls, "start_stage2_async")
    or "clean_red_rows" not in inspect.signature(_bot_cls.start_stage1_async).parameters
    or "printer_name" not in inspect.signature(_bot_cls.start_stage2_async).parameters
    or "overwrite_original" not in inspect.signature(_bot_cls.start_stage2_async).parameters
    or "clean_red_rows" not in inspect.signature(_bot_cls.start_stage2_async).parameters
)
browser_bot, _bot_reloaded = _reload_local_module_if_needed(browser_bot, force=_bot_stale)

_settings_stale = (
    _parser_reloaded
    or _bot_reloaded
    or not hasattr(settings_store, "get_default_settings")
    or "printer_name" not in settings_store.get_default_settings()
    or "overwrite_original_word" not in settings_store.get_default_settings()
    or "clean_red_stage1" not in settings_store.get_default_settings()
    or "clean_red_stage2" not in settings_store.get_default_settings()
)
settings_store, _ = _reload_local_module_if_needed(settings_store, force=_settings_stale)

NorstalBrowserBot = browser_bot.NorstalBrowserBot
Stage1PlanResult = browser_bot.Stage1PlanResult
Stage2PlanResult = browser_bot.Stage2PlanResult
_is_uploaded_temp_word_path = browser_bot._is_uploaded_temp_word_path

CuttingPlanRecord = parser_docs.CuttingPlanRecord
discover_folder_files = parser_docs.discover_folder_files
parse_files = parser_docs.parse_files
plan_sort_key = parser_docs.plan_sort_key

INPUT_MODE_LOCAL = settings_store.INPUT_MODE_LOCAL
INPUT_MODE_OPTIONS = settings_store.INPUT_MODE_OPTIONS
build_session_payload = settings_store.build_session_payload
clear_session_data = settings_store.clear_session_data
clear_user_settings = settings_store.clear_user_settings
load_session_data = settings_store.load_session_data
load_user_settings = settings_store.load_user_settings
save_session_data = settings_store.save_session_data
save_user_settings = settings_store.save_user_settings
session_fingerprint = settings_store.session_fingerprint

DEFAULT_PRINTER_OPTION = word_qr_printer.DEFAULT_PRINTER_OPTION
list_windows_printers = word_qr_printer.list_windows_printers

_CHECKBOX_SUPPORTS_PERSIST_STATE = "persist_state" in inspect.signature(st.checkbox).parameters
_CB_PERSIST_KW = {"persist_state": "session"} if _CHECKBOX_SUPPORTS_PERSIST_STATE else {}

load_dotenv()

st.set_page_config(
    page_title="NORSTAL — Automatizare Planuri Croi",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ==============================================================================
# Stilizare Dark Mode & Componente Vizuale
# ==============================================================================
st.markdown(
    """
<style>
    .stApp {
        background-color: #0E1117;
        color: #E6EDF3;
    }
    .step-bar {
        display: flex;
        gap: 12px;
        margin: 0.5rem 0 1.25rem 0;
        flex-wrap: wrap;
    }
    .step-pill {
        flex: 1;
        min-width: 220px;
        padding: 12px 16px;
        border-radius: 10px;
        border: 1px solid #30363D;
        background: #161B22;
        font-size: 0.92rem;
        font-weight: 600;
        color: #8B949E;
    }
    .step-pill.active {
        border-color: #00B4D8;
        background: rgba(0, 180, 216, 0.12);
        color: #E6EDF3;
        box-shadow: 0 0 12px rgba(0, 180, 216, 0.18);
    }
    .step-pill.done {
        border-color: #2EA043;
        background: rgba(46, 160, 67, 0.14);
        color: #7EE787;
    }
    .log-console {
        background: #0B0F14;
        border: 1px solid #30363D;
        border-radius: 8px;
        padding: 12px 14px;
        font-family: "Cascadia Code", "Consolas", monospace;
        font-size: 0.82rem;
        line-height: 1.45;
        max-height: 360px;
        overflow-y: auto;
        white-space: pre-wrap;
        word-break: break-word;
    }
    .log-err { color: #FF7B72; }
    .log-warn { color: #D29922; }
    .log-info { color: #A5D6FF; }
    .badge-print-on {
        display: inline-block;
        padding: 4px 10px;
        border-radius: 999px;
        background: rgba(46, 160, 67, 0.18);
        border: 1px solid #2EA043;
        color: #7EE787;
        font-size: 0.82rem;
        font-weight: 600;
    }
    .badge-print-off {
        display: inline-block;
        padding: 4px 10px;
        border-radius: 999px;
        background: rgba(210, 153, 34, 0.18);
        border: 1px solid #D29922;
        color: #F2CC60;
        font-size: 0.82rem;
        font-weight: 600;
    }
</style>
""",
    unsafe_allow_html=True,
)


def _current_session_payload() -> dict[str, Any]:
    """Construiește snapshot-ul serializabil din starea curentă Streamlit."""
    return build_session_payload(
        stage=st.session_state.get("stage", "idle"),
        plans=st.session_state.get("plans", []),
        active_plans=st.session_state.get("active_plans", []),
        plan_overrides=st.session_state.get("plan_overrides", {}),
        stage1_results=st.session_state.get("stage1_results", []),
        stage2_results=st.session_state.get("stage2_results", []),
        discovered_word=st.session_state.get("discovered_word", []),
        discovered_excel=st.session_state.get("discovered_excel", []),
        selected_word_names=st.session_state.get("saved_selected_word_names", []),
        selected_s2_names=st.session_state.get("saved_s2_names"),
        bot_logs=st.session_state.get("bot_logs", []),
    )


def _persist_settings_if_changed(current_settings: dict[str, Any]) -> None:
    """Salvează setările pe disc doar dacă s-au modificat față de ultimul snapshot."""
    candidate_json = json.dumps(current_settings, ensure_ascii=False, sort_keys=True)
    if candidate_json == st.session_state.get("_last_saved_settings_json"):
        return
    clean = save_user_settings(current_settings)
    st.session_state["saved_settings"] = clean
    st.session_state["_last_saved_settings_json"] = json.dumps(
        clean, ensure_ascii=False, sort_keys=True
    )


def _persist_session_if_changed(force: bool = False) -> None:
    """Salvează datele sesiunii pe disc doar dacă s-au modificat (sau la force=True)."""
    payload = _current_session_payload()
    sig = session_fingerprint(payload)
    if force or sig != st.session_state.get("_last_saved_session_sig"):
        save_session_data(payload)
        st.session_state["_last_saved_session_sig"] = sig


def _init_state() -> None:
    """Inițializează starea aplicației o singură dată per sesiune din fișierele salvate."""
    if st.session_state.get("_state_initialized"):
        return

    saved_settings = load_user_settings()
    saved_session = load_session_data()

    st.session_state["saved_settings"] = dict(saved_settings)
    st.session_state["_last_saved_settings_json"] = json.dumps(
        saved_settings, ensure_ascii=False, sort_keys=True
    )

    st.session_state["stage"] = saved_session["stage"]
    st.session_state["plans"] = saved_session["plans"]
    st.session_state["active_plans"] = saved_session["active_plans"]
    st.session_state["base_plan_overrides"] = dict(saved_session["plan_overrides"])
    st.session_state["plan_overrides"] = dict(saved_session["plan_overrides"])
    st.session_state["plans_version"] = 0
    st.session_state["stage1_results"] = saved_session["stage1_results"]
    st.session_state["stage2_results"] = saved_session["stage2_results"]
    st.session_state["discovered_word"] = saved_session["discovered_word"]
    st.session_state["discovered_excel"] = saved_session["discovered_excel"]
    st.session_state["saved_selected_word_names"] = list(saved_session["selected_word_names"])
    st.session_state["saved_s2_names"] = (
        list(saved_session["selected_s2_names"])
        if saved_session["selected_s2_names"] is not None
        else None
    )
    st.session_state["bot_logs"] = list(saved_session["bot_logs"])
    st.session_state["bot"] = None

    initial_payload = _current_session_payload()
    st.session_state["_last_saved_session_sig"] = session_fingerprint(initial_payload)

    st.session_state["cfg_intranet_url"] = saved_settings["intranet_url"]
    st.session_state["cfg_intranet_user"] = saved_settings["intranet_user"]
    st.session_state["cfg_intranet_pass"] = saved_settings["intranet_pass"]
    st.session_state["cfg_visible_mode"] = bool(saved_settings["visible_mode"])
    st.session_state["use_full_http_mode"] = bool(saved_settings["use_full_http_mode"])
    st.session_state["use_http_mode"] = bool(saved_settings["use_http_mode"])
    st.session_state["use_import_mode"] = bool(saved_settings["use_import_mode"])
    st.session_state["cfg_clean_red_stage1"] = bool(saved_settings["clean_red_stage1"])
    st.session_state["cfg_clean_red_stage2"] = bool(saved_settings["clean_red_stage2"])
    st.session_state["cfg_send_to_printer"] = bool(saved_settings["send_to_printer"])
    st.session_state["cfg_printer_name"] = saved_settings["printer_name"] or DEFAULT_PRINTER_OPTION
    st.session_state["cfg_overwrite_original_word"] = bool(
        saved_settings["overwrite_original_word"]
    )
    st.session_state["cfg_print_delay"] = float(saved_settings["print_delay"])
    st.session_state["cfg_output_dir_str"] = saved_settings["output_dir_str"]
    st.session_state["cfg_input_mode"] = saved_settings["input_mode"]
    st.session_state["cfg_word_dir_str"] = saved_settings["word_dir_str"]
    st.session_state["cfg_excel_dir_str"] = saved_settings["excel_dir_str"]
    st.session_state["cfg_name_filter"] = saved_settings["name_filter"]

    st.session_state["_state_initialized"] = True


def _render_logs_box(placeholder, logs: list[str]) -> None:
    if not logs:
        placeholder.markdown(
            '<div class="log-console"><span class="log-info">Nicio înregistrare încă.</span></div>',
            unsafe_allow_html=True,
        )
        return

    lines_html = []
    for raw in logs[-300:]:
        esc = html.escape(raw)
        if "[ERROR]" in raw:
            cls = "log-err"
        elif "[WARN]" in raw:
            cls = "log-warn"
        else:
            cls = "log-info"
        lines_html.append(f'<div class="{cls}">{esc}</div>')

    placeholder.markdown(
        f'<div class="log-console">{"".join(lines_html)}</div>',
        unsafe_allow_html=True,
    )


def _render_stepper(stage: str) -> None:
    s1_cls = "done" if stage in {"parsed", "stage1_done", "stage2_done"} else "active"
    s2_cls = (
        "done"
        if stage in {"stage1_done", "stage2_done"}
        else ("active" if stage == "parsed" else "")
    )
    s3_cls = (
        "done"
        if stage == "stage2_done"
        else ("active" if stage == "stage1_done" else "")
    )
    st.markdown(
        f"""
<div class="step-bar">
  <div class="step-pill {s1_cls}">1️⃣ Încărcare &amp; Parsare Fișiere</div>
  <div class="step-pill {s2_cls}">2️⃣ Etapa 1: Încărcare Intranet (ROȘU opțional)</div>
  <div class="step-pill {s3_cls}">3️⃣ Verificare Umană → Etapa 2: Curățare/Validare, QR &amp; Print</div>
</div>
""",
        unsafe_allow_html=True,
    )


def _close_bot_and_reset(reset_plans: bool = False) -> None:
    bot: Optional[NorstalBrowserBot] = st.session_state.get("bot")
    if bot is not None:
        try:
            bot.close()
        except Exception:
            pass
    st.session_state["bot"] = None
    st.session_state["stage1_results"] = []
    st.session_state["stage2_results"] = []
    st.session_state["active_plans"] = []
    st.session_state["saved_s2_names"] = None
    st.session_state.pop("cfg_selected_s2_names", None)
    st.session_state["bot_logs"] = []

    if reset_plans:
        st.session_state["plans"] = []
        st.session_state["stage"] = "idle"
        st.session_state["base_plan_overrides"] = {}
        st.session_state["plan_overrides"] = {}
        st.session_state["discovered_word"] = []
        st.session_state["discovered_excel"] = []
        st.session_state["saved_selected_word_names"] = []
        st.session_state.pop("cfg_selected_word_names", None)
        st.session_state["plans_version"] = int(st.session_state.get("plans_version", 0)) + 1
        clear_session_data()
        st.session_state["_last_saved_session_sig"] = session_fingerprint(
            _current_session_payload()
        )
    else:
        st.session_state["stage"] = "parsed" if st.session_state.get("plans") else "idle"
        _persist_session_if_changed(force=True)


def _reset_all_saved_state() -> None:
    """Șterge toate setările și datele salvate și reinițializează din valorile implicite/.env."""
    bot: Optional[NorstalBrowserBot] = st.session_state.get("bot")
    if bot is not None:
        try:
            bot.close()
        except Exception:
            pass
    clear_user_settings()
    clear_session_data()

    static_keys = {
        "_state_initialized",
        "_last_saved_settings_json",
        "_last_saved_session_sig",
        "saved_settings",
        "stage",
        "plans",
        "active_plans",
        "base_plan_overrides",
        "plan_overrides",
        "plans_version",
        "stage1_results",
        "stage2_results",
        "discovered_word",
        "discovered_excel",
        "saved_selected_word_names",
        "saved_s2_names",
        "bot_logs",
        "bot",
        "use_full_http_mode",
        "use_http_mode",
        "use_import_mode",
    }
    for key in list(st.session_state.keys()):
        if (
            key in static_keys
            or key.startswith("cfg_")
            or key.startswith("plans_preview_editor")
        ):
            st.session_state.pop(key, None)


_init_state()

# ==============================================================================
# Sidebar Configurare
# ==============================================================================
with st.sidebar:
    st.header("⚙️ Configurare Intranet")
    intranet_url = st.text_input(
        "URL Intranet NORSTAL",
        key="cfg_intranet_url",
    )
    intranet_user = st.text_input(
        "Utilizator (User)",
        key="cfg_intranet_user",
        placeholder="ex: DanielMan",
    )
    intranet_pass = st.text_input(
        "Parolă",
        key="cfg_intranet_pass",
        type="password",
    )

    st.divider()
    st.subheader("🖥️ Browser & Printare")
    visible_mode = st.toggle(
        "Mod Vizibil (Browser deschis)",
        key="cfg_visible_mode",
        help="Când este activ, browserul Chromium rămâne vizibil pe ecran pentru debug și verificare vizuală.",
    )

    def _enforce_exclusive_mode(active_key: str) -> None:
        if not st.session_state.get(active_key):
            return
        for mode_key in ("use_full_http_mode", "use_http_mode", "use_import_mode"):
            if mode_key != active_key:
                st.session_state[mode_key] = False

    st.markdown("**Mod încărcare planuri & piese**")
    use_full_http_mode = st.checkbox(
        "🚀 Full HTTP — tot planul rapid (P1 + P2 + P3)",
        key="use_full_http_mode",
        on_change=_enforce_exclusive_mode,
        args=("use_full_http_mode",),
        help="Calibrează automat pe Planul #1 prin UI, apoi introduce Planurile #2..#N (Panou 1, Panou 2, toate piesele Panou 3 și curățarea ROȘU) direct prin HTTP Request (~5–10x mai rapid).",
    )
    use_http_mode = st.checkbox(
        "⚡ Încărcare piese prin HTTP Request (Panou 3)",
        key="use_http_mode",
        on_change=_enforce_exclusive_mode,
        args=("use_http_mode",),
        help="Introduce prima piesă prin '+' pentru a prelua request-ul sesiunii, apoi trimite restul pieselor direct prin HTTP POST.",
    )
    use_import_mode = st.checkbox(
        "📥 Încărcare piese prin Import Excel (.xlsx)",
        key="use_import_mode",
        on_change=_enforce_exclusive_mode,
        args=("use_import_mode",),
        help="Folosește modalul verde 'Import' din Panoul 3 pentru încărcarea fișierului .xlsx.",
    )
    if use_full_http_mode:
        panel3_mode = "full_http"
        st.caption("Mod activ: **Full HTTP (P1 + P2 + P3 rapid)**")
    elif use_http_mode:
        panel3_mode = "http_request"
        st.caption("Mod activ: **HTTP Request rapid (Panou 3)**")
    elif use_import_mode:
        panel3_mode = "import_excel"
        st.caption("Mod activ: **Import Excel (.xlsx)**")
    else:
        panel3_mode = "manual_plus"
        st.caption("Mod activ: **Adăugare piesă cu piesă (+)**")

    send_to_printer = st.checkbox(
        "Trimite la imprimantă",
        key="cfg_send_to_printer",
        help="Debifați pentru a testa mai întâi doar generarea documentelor Word cu cod QR.",
    )
    if send_to_printer:
        st.markdown('<span class="badge-print-on">🖨️ Printare automată ACTIVĂ</span>', unsafe_allow_html=True)
    else:
        st.markdown('<span class="badge-print-off">🧪 Mod TEST: doar generare Word + QR</span>', unsafe_allow_html=True)

    installed_printers, default_win_printer = list_windows_printers()
    printer_options = [DEFAULT_PRINTER_OPTION] + installed_printers
    if st.session_state.get("cfg_printer_name") not in printer_options:
        st.session_state["cfg_printer_name"] = DEFAULT_PRINTER_OPTION

    selected_printer_opt = st.selectbox(
        "Imprimantă",
        options=printer_options,
        key="cfg_printer_name",
        format_func=lambda opt: (
            f"Imprimanta implicită ({default_win_printer})"
            if (opt == DEFAULT_PRINTER_OPTION and default_win_printer)
            else opt
        ),
        help="Alegeți imprimanta Windows pentru printarea automată în Etapa 2.",
    )
    effective_printer_name = (
        "" if selected_printer_opt == DEFAULT_PRINTER_OPTION else str(selected_printer_opt)
    )
    printer_display = (
        effective_printer_name or default_win_printer or DEFAULT_PRINTER_OPTION
    )

    print_delay = st.number_input(
        "Pauză între printări (secunde)",
        min_value=0.0,
        max_value=15.0,
        step=0.5,
        key="cfg_print_delay",
    )
    if "cfg_overwrite_original_word" not in st.session_state:
        st.session_state["cfg_overwrite_original_word"] = bool(
            st.session_state.get("saved_settings", {}).get("overwrite_original_word", True)
        )
    overwrite_original_word = st.checkbox(
        "📝 Suprascrie documentele Word originale cu QR",
        key="cfg_overwrite_original_word",
        help="Salvează codul QR direct în locul fișierelor originale (.doc / .docx) din folderul local. Debifați pentru a salva copii noi în folderul de output.",
    )
    if overwrite_original_word:
        st.caption("Fișierele originale (.doc / .docx) vor fi actualizate direct pe disc cu codul QR.")

    output_dir_str = st.text_input(
        "Folder output (copii / Drag & Drop)",
        key="cfg_output_dir_str",
        help="Folosit când opțiunea de suprascriere a originalelor este debifată sau la încărcare prin Drag & Drop.",
    )

    st.divider()
    st.subheader("🌐 Stare Sesiune Browser")
    current_bot: Optional[NorstalBrowserBot] = st.session_state.get("bot")
    if current_bot is not None and current_bot.is_alive():
        st.success("🟢 Sesiune Playwright activă")
    else:
        st.info("⚪ Browser nepornit")

    if st.button("🛑 Închide Browser & Resetează Etapele", width="stretch"):
        _close_bot_and_reset(reset_plans=False)
        st.rerun()

    st.divider()
    st.caption("💾 Setările și datele introduse se salvează automat local.")
    if st.button("🗑️ Șterge lotul salvat", width="stretch"):
        _close_bot_and_reset(reset_plans=True)
        st.rerun()
    if st.button("♻️ Resetează toate setările", width="stretch"):
        _reset_all_saved_state()
        st.rerun()

# ==============================================================================
# Header & Stepper
# ==============================================================================
st.title("⚙️ Automatizare Planuri de Croi — NORSTAL")
st.caption(
    "Încărcare automată din Metalix (.doc/.docx + .xlsx) → Intranet jqGrid → Verificare umană → Validare, QR & Printare."
)
_render_stepper(st.session_state["stage"])

def _apply_parsed_plans(parsed: list[CuttingPlanRecord]) -> None:
    """Actualizează starea lotului după parsare și persistă sesiunea."""
    st.session_state["plans"] = parsed
    st.session_state["stage"] = "parsed" if parsed else "idle"
    st.session_state["active_plans"] = []
    st.session_state["stage1_results"] = []
    st.session_state["stage2_results"] = []
    st.session_state["base_plan_overrides"] = {}
    st.session_state["plan_overrides"] = {}
    st.session_state["saved_s2_names"] = None
    st.session_state.pop("cfg_selected_s2_names", None)
    st.session_state["plans_version"] = int(st.session_state.get("plans_version", 0)) + 1
    _persist_session_if_changed(force=True)


# ==============================================================================
# Secțiunea 1: Selecție / Încărcare Fișiere & Parsare
# ==============================================================================
st.subheader("1. Selecție fișiere și previzualizare date")

input_mode = st.radio(
    "Mod alimentare fișiere",
    options=INPUT_MODE_OPTIONS,
    horizontal=True,
    key="cfg_input_mode",
)

word_paths_to_parse: list[Path] = []
excel_paths_to_parse: list[Path] = []

if input_mode == INPUT_MODE_LOCAL:
    for fld, wkey in (
        ("word_dir_str", "cfg_word_dir_str"),
        ("excel_dir_str", "cfg_excel_dir_str"),
        ("name_filter", "cfg_name_filter"),
    ):
        if wkey not in st.session_state:
            st.session_state[wkey] = st.session_state["saved_settings"].get(fld, "")

    col_d1, col_d2, col_d3 = st.columns([2, 2, 1.4])
    with col_d1:
        word_dir_str = st.text_input(
            "Folder fișiere Word (.doc / .docx)",
            key="cfg_word_dir_str",
        )
    with col_d2:
        excel_dir_str = st.text_input(
            "Folder fișiere Excel (.xlsx) — opțional",
            key="cfg_excel_dir_str",
            placeholder="Lăsați gol dacă sunt în același folder cu Word",
        )
    with col_d3:
        name_filter = st.text_input(
            "Filtru nume (opțional)",
            key="cfg_name_filter",
            placeholder="ex: 1242-49-2",
        )
else:
    word_dir_str = str(st.session_state["saved_settings"].get("word_dir_str", ""))
    excel_dir_str = str(st.session_state["saved_settings"].get("excel_dir_str", ""))
    name_filter = str(st.session_state["saved_settings"].get("name_filter", ""))

def _build_current_settings_dict() -> dict[str, Any]:
    """Construiește dicționarul complet de setări din valorile curente ale interfeței."""
    saved = st.session_state.get("saved_settings", {})
    return {
        "intranet_url": intranet_url,
        "intranet_user": intranet_user,
        "intranet_pass": intranet_pass,
        "visible_mode": bool(visible_mode),
        "use_full_http_mode": bool(use_full_http_mode),
        "use_http_mode": bool(use_http_mode),
        "use_import_mode": bool(use_import_mode),
        "clean_red_stage1": bool(
            st.session_state.get("cfg_clean_red_stage1", saved.get("clean_red_stage1", True))
        ),
        "clean_red_stage2": bool(
            st.session_state.get("cfg_clean_red_stage2", saved.get("clean_red_stage2", False))
        ),
        "send_to_printer": bool(send_to_printer),
        "printer_name": effective_printer_name,
        "overwrite_original_word": bool(overwrite_original_word),
        "print_delay": float(print_delay),
        "output_dir_str": output_dir_str,
        "input_mode": input_mode,
        "word_dir_str": word_dir_str,
        "excel_dir_str": excel_dir_str,
        "name_filter": name_filter,
    }


_persist_settings_if_changed(_build_current_settings_dict())

if input_mode == INPUT_MODE_LOCAL:
    if st.button("🔎 Scanează folderul"):
        w_found, e_found = discover_folder_files(word_dir_str, excel_dir_str, name_filter=name_filter)
        st.session_state["discovered_word"] = w_found
        st.session_state["discovered_excel"] = e_found
        default_sel = w_found if len(w_found) <= 25 else []
        default_names = [p.name for p in default_sel]
        st.session_state["saved_selected_word_names"] = list(default_names)
        st.session_state["cfg_selected_word_names"] = list(default_names)
        _persist_session_if_changed(force=True)

    disc_w: list[Path] = sorted(st.session_state.get("discovered_word", []), key=plan_sort_key)
    disc_e: list[Path] = st.session_state.get("discovered_excel", [])

    if disc_w or disc_e:
        st.caption(f"Găsite în folder: **{len(disc_w)}** fișiere Word și **{len(disc_e)}** fișiere `.xlsx`.")
        valid_word_names = [p.name for p in disc_w]
        valid_word_set = set(valid_word_names)
        if "cfg_selected_word_names" not in st.session_state:
            st.session_state["cfg_selected_word_names"] = [
                n for n in st.session_state.get("saved_selected_word_names", []) if n in valid_word_set
            ]
        else:
            st.session_state["cfg_selected_word_names"] = [
                n for n in st.session_state["cfg_selected_word_names"] if n in valid_word_set
            ]

        selected_names = st.multiselect(
            "Selectați planurile Word de analizat",
            options=valid_word_names,
            key="cfg_selected_word_names",
            help="Dacă folderul are multe fișiere, folosiți filtrul de nume de mai sus sau alegeți exact planurile din lot.",
        )
        st.session_state["saved_selected_word_names"] = list(selected_names)
        sel_set = set(selected_names)
        word_paths_to_parse = [p for p in disc_w if p.name in sel_set]
        excel_paths_to_parse = disc_e
else:
    uploaded_files = st.file_uploader(
        "Încărcați fișierele Word (.doc / .docx) și Excel (.xlsx)",
        type=["doc", "docx", "xlsx"],
        accept_multiple_files=True,
    )
    if uploaded_files:
        upload_dir = Path("work_tmp/uploads") / datetime.now().strftime("%Y%m%d_%H%M%S")
        if st.button("🔍 Analizează & Împerechează fișierele încărcate", type="primary"):
            upload_dir.mkdir(parents=True, exist_ok=True)
            for uf in uploaded_files:
                dest = upload_dir / Path(uf.name).name
                dest.write_bytes(uf.getbuffer())
                if dest.suffix.lower() in {".doc", ".docx"}:
                    word_paths_to_parse.append(dest)
                elif dest.suffix.lower() == ".xlsx":
                    excel_paths_to_parse.append(dest)

            with st.spinner("Conversie .DOC (dacă este cazul), parsare și împerechere..."):
                parsed = parse_files(word_paths_to_parse, excel_paths_to_parse)
            _apply_parsed_plans(parsed)
            st.rerun()

if input_mode == INPUT_MODE_LOCAL and word_paths_to_parse:
    if st.button("🔍 Analizează & Împerechează fișierele selectate", type="primary"):
        with st.spinner("Conversie .DOC (Word COM) și extragere date din tabele..."):
            parsed = parse_files(word_paths_to_parse, excel_paths_to_parse)
        _apply_parsed_plans(parsed)
        st.rerun()

# ==============================================================================
# Tabel Previzualizare Planuri Parsate
# ==============================================================================
raw_plans: list[CuttingPlanRecord] = st.session_state.get("plans", [])
plans = sorted(raw_plans, key=lambda p: plan_sort_key(p.plan_name))
if [p.plan_name for p in plans] != [p.plan_name for p in raw_plans]:
    st.session_state["plans"] = plans
    st.session_state["plans_version"] = int(st.session_state.get("plans_version", 0)) + 1
edited_records: list[CuttingPlanRecord] = []

if plans:
    ready_count = sum(1 for p in plans if p.is_ready)
    missing_xlsx = sum(1 for p in plans if not p.excel_found)
    laser_machine_count = sum(1 for p in plans if "plasma" not in p.machine_info.lower())
    plasma_machine_count = sum(1 for p in plans if "plasma" in p.machine_info.lower())

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total planuri Word", len(plans))
    m2.metric("Perechi complete (Word + Excel)", ready_count)
    m3.metric("Lipsă .xlsx", missing_xlsx)
    m4.metric("Mașină Word (Laser / Plasma) → Plasma", f"{laser_machine_count} / {plasma_machine_count}")

    if missing_xlsx > 0:
        missing_names = [p.plan_name for p in plans if not p.excel_found]
        st.warning(
            f"⚠️ Lipsesc fișierele `.xlsx` pentru {missing_xlsx} plan(uri): {', '.join(missing_names)}"
        )

    base_overrides: dict[str, dict[str, Any]] = st.session_state.get("base_plan_overrides", {})
    rows_df = []
    for p in plans:
        ov = base_overrides.get(p.plan_name, {})
        notes = p.errors + p.warnings
        include_val = bool(ov["include"]) if "include" in ov else bool(p.is_ready)
        debitare_val = str(ov.get("cutting_type") or p.cutting_type)
        if debitare_val not in {"Plasma", "Laser"}:
            debitare_val = p.cutting_type if p.cutting_type in {"Plasma", "Laser"} else "Plasma"
        comanda_val = str(ov["project_query"]) if "project_query" in ov else str(p.project_query)
        weight_val = (
            int(ov["part_weight_int"])
            if "part_weight_int" in ov
            else (int(p.part_weight_int) if p.part_weight_int is not None else 0)
        )
        rows_df.append(
            {
                "Include": include_val,
                "Plan croi": p.plan_name,
                "SubNest validat": "✅ Da" if p.name_validated else "⚠️ Verifică",
                "Mașină (Word)": p.machine_info,
                "Debitare": debitare_val,
                "Comanda (prefix)": comanda_val,
                "Greutate (kg)": weight_val,
                "Excel (.xlsx)": Path(p.excel_path).name if p.excel_found and p.excel_path else "❌ LIPSEȘTE",
                "Piese (.xlsx)": p.parts_count,
                "Observații": "Gata" if p.is_ready and not notes else " | ".join(notes),
            }
        )

    df_preview = pd.DataFrame(rows_df)
    editor_version = int(st.session_state.get("plans_version", 0))
    edited_df = st.data_editor(
        df_preview,
        width="stretch",
        hide_index=True,
        key=f"plans_preview_editor_{editor_version}",
        column_config={
            "Include": st.column_config.CheckboxColumn("Include", default=True),
            "Debitare": st.column_config.SelectboxColumn("Debitare", options=["Plasma", "Laser"], required=True),
            "Comanda (prefix)": st.column_config.TextColumn("Comanda (prefix)"),
            "Greutate (kg)": st.column_config.NumberColumn("Greutate (kg)", min_value=0, step=1),
        },
        disabled=["Plan croi", "SubNest validat", "Mașină (Word)", "Excel (.xlsx)", "Piese (.xlsx)", "Observații"],
    )

    current_overrides: dict[str, dict[str, Any]] = {}
    for idx, rec in enumerate(plans):
        row = edited_df.iloc[idx]
        raw_inc = row["Include"]
        inc_val = False if pd.isna(raw_inc) else bool(raw_inc)

        raw_deb = row["Debitare"]
        deb_val = rec.cutting_type if pd.isna(raw_deb) or not str(raw_deb).strip() else str(raw_deb).strip()
        if deb_val not in {"Plasma", "Laser"}:
            deb_val = "Plasma"

        raw_cmd = row["Comanda (prefix)"]
        cmd_val = "" if pd.isna(raw_cmd) else str(raw_cmd).strip()

        raw_wt = row["Greutate (kg)"]
        if pd.isna(raw_wt):
            wt_val = 0
        else:
            try:
                wt_val = max(0, int(float(raw_wt)))
            except (TypeError, ValueError):
                wt_val = 0

        current_overrides[rec.plan_name] = {
            "include": inc_val,
            "cutting_type": deb_val,
            "project_query": cmd_val,
            "part_weight_int": wt_val,
        }

        if inc_val and rec.excel_found and wt_val > 0:
            updated = replace(
                rec,
                cutting_type=deb_val,
                project_query=cmd_val,
                part_weight_int=wt_val,
            )
            edited_records.append(updated)
        elif inc_val and not rec.excel_found:
            st.error(f"Planul `{rec.plan_name}` nu poate fi inclus fără fișierul `.xlsx` asociat.")

    edited_records.sort(key=lambda rec: plan_sort_key(rec.plan_name))
    st.session_state["plan_overrides"] = current_overrides

# ==============================================================================
# Secțiunea 2: Etapa 1 — Încărcare în Intranet & Curățare
# ==============================================================================
if plans:
    st.divider()
    st.subheader("2. Etapa 1 — Încărcare în Intranet & Curățare rânduri")

    c1, c2 = st.columns([2, 3])
    with c1:
        st.write(f"Planuri pregătite pentru rulare: **{len(edited_records)}**")
    with c2:
        if not visible_mode and (not intranet_user or not intranet_pass):
            st.warning("În modul Headless trebuie completate User și Parolă în bara laterală.")

    def _get_or_create_bot(force_new: bool = False) -> NorstalBrowserBot:
        existing: Optional[NorstalBrowserBot] = st.session_state.get("bot")
        if (
            not force_new
            and existing is not None
            and isinstance(existing, NorstalBrowserBot)
            and existing.is_alive()
        ):
            return existing
        if existing is not None:
            try:
                existing.close()
            except Exception:
                pass
        new_bot = NorstalBrowserBot(
            base_url=intranet_url,
            username=intranet_user,
            password=intranet_pass,
            headless=not visible_mode,
        )
        st.session_state["bot"] = new_bot
        return new_bot

    def _run_stage1_with_ui(
        bot: NorstalBrowserBot,
        batch_plans: list[CuttingPlanRecord],
        progress_slot,
        log_slot,
        clean_red_rows: bool,
    ) -> Optional[list[Stage1PlanResult]]:
        fut = bot.start_stage1_async(
            batch_plans,
            panel3_mode=panel3_mode,
            clean_red_rows=clean_red_rows,
        )
        pbar = progress_slot.progress(0.0, text="Inițializare browser Playwright...")

        while not fut.done():
            prog = bot.get_progress()
            tot = max(1, int(prog.get("total") or len(batch_plans)))
            cur = int(prog.get("current") or 0)
            pbar.progress(min(1.0, cur / tot), text=prog.get("message") or "Procesare Etapa 1...")
            _render_logs_box(log_slot, bot.get_logs())
            time.sleep(0.25)

        try:
            res: list[Stage1PlanResult] = fut.result()
            st.session_state["bot_logs"] = bot.get_logs()[-300:]
            pbar.progress(1.0, text="Etapa 1 finalizată!")
            return res
        except Exception as exc:
            logs_now = bot.get_logs()
            st.session_state["bot_logs"] = logs_now[-300:]
            _render_logs_box(log_slot, logs_now)
            st.error(f"Etapa 1 a întâmpinat o eroare critică: {exc}")
            return None

    if "cfg_clean_red_stage1" not in st.session_state:
        st.session_state["cfg_clean_red_stage1"] = bool(
            st.session_state.get("saved_settings", {}).get("clean_red_stage1", True)
        )
    clean_red_stage1 = st.checkbox(
        "🧹 Curăță automat piesele ROȘII în Etapa 1",
        key="cfg_clean_red_stage1",
        help="Debifați pentru a NU șterge piesele ROȘII în Etapa 1 — acestea rămân în Panoul 3 pentru verificare manuală sau curățare separată în Etapa 2.",
        **_CB_PERSIST_KW,
    )
    if not clean_red_stage1:
        st.info(
            "ℹ️ Curățarea automată a pieselor ROȘII este **dezactivată** în Etapa 1 — toate piesele rămân în Panoul 3 și le puteți curăța separat în **Etapa 2**."
        )

    s1_btn_label = (
        "🚀 Pornește Etapa 1 (Încărcare în Intranet & Curățare ROȘU)"
        if clean_red_stage1
        else "🚀 Pornește Etapa 1 (Încărcare fără ștergere ROȘU)"
    )
    start_s1 = st.button(
        s1_btn_label,
        type="primary",
        disabled=(len(edited_records) == 0),
    )

    progress_slot_s1 = st.empty()
    log_slot_live = st.empty()

    if start_s1:
        _close_bot_and_reset(reset_plans=False)
        bot = _get_or_create_bot(force_new=True)
        st.session_state["active_plans"] = edited_records
        s1_res = _run_stage1_with_ui(
            bot,
            edited_records,
            progress_slot_s1,
            log_slot_live,
            clean_red_rows=bool(clean_red_stage1),
        )
        if s1_res is not None:
            st.session_state["stage1_results"] = s1_res
            st.session_state["stage"] = "stage1_done"
            st.session_state["saved_s2_names"] = None
            st.session_state.pop("cfg_selected_s2_names", None)
            _persist_session_if_changed(force=True)
            st.rerun()
        else:
            _persist_session_if_changed(force=True)

# Rezultate Etapa 1 + Banner verificare umană
s1_results: list[Stage1PlanResult] = st.session_state.get("stage1_results", [])
if st.session_state["stage"] in {"stage1_done", "stage2_done"}:
    failed_s1 = [r for r in s1_results if not r.success]
    if failed_s1:
        st.warning(
            f"⚠️ Etapa 1 finalizată cu **{len(failed_s1)}** plan(uri) cu eroare/goale: "
            + ", ".join(r.plan_name for r in failed_s1)
            + ". Le puteți reîncerca mai jos sau repara manual în intranet înainte de Etapa 2."
        )
    else:
        st.success(
            "✅ **Etapa 1 finalizată. Verificați datele în intranet, apoi apăsați butonul de Validare.**"
        )

    if s1_results:
        s1_df = pd.DataFrame(
            [
                {
                    "Plan croi": r.plan_name,
                    "Status": "✅ OK" if r.success else "❌ Eroare",
                    "Comanda potrivită": r.matched_comanda or "-",
                    "Piese adăugate": r.parts_added,
                    "Respinse": r.parts_rejected,
                    "Rânduri ROȘII șterse": r.red_deleted,
                    "Rânduri ROȘII păstrate": getattr(r, "red_kept", 0),
                    "Rânduri ROZ păstrate": r.pink_kept,
                    "Eroare": r.error or "-",
                }
                for r in s1_results
            ]
        )
        st.dataframe(s1_df, width="stretch", hide_index=True)

    if failed_s1 and plans:
        failed_names_set = {r.plan_name for r in failed_s1}
        retry_pool = edited_records or st.session_state.get("active_plans", []) or plans
        retry_plans = sorted(
            [p for p in retry_pool if p.plan_name in failed_names_set],
            key=lambda rec: plan_sort_key(rec.plan_name),
        )
        retry_progress = st.empty()
        retry_logs = st.empty()
        if retry_plans and st.button(f"🔁 Reîncearcă Etapa 1 doar pentru planurile cu eroare ({len(retry_plans)})"):
            bot = _get_or_create_bot(force_new=False)
            retry_res = _run_stage1_with_ui(
                bot,
                retry_plans,
                retry_progress,
                retry_logs,
                clean_red_rows=bool(clean_red_stage1),
            )
            if retry_res is not None:
                retry_map = {r.plan_name: r for r in retry_res}
                merged = [retry_map.get(r.plan_name, r) for r in s1_results]
                existing_names = {r.plan_name for r in merged}
                for r in retry_res:
                    if r.plan_name not in existing_names:
                        merged.append(r)
                merged.sort(key=lambda r: plan_sort_key(r.plan_name))
                st.session_state["stage1_results"] = merged
                st.session_state["saved_s2_names"] = None
                st.session_state.pop("cfg_selected_s2_names", None)
                _persist_session_if_changed(force=True)
                st.rerun()
            else:
                _persist_session_if_changed(force=True)

# ==============================================================================
# Secțiunea 3: Etapa 2 — Validare, Copiere QR în Word & Printare
# ==============================================================================
if plans:
    st.divider()
    st.subheader("3. Etapa 2 — Validare, Copiere QR în Word & Printare")

    active_plans: list[CuttingPlanRecord] = st.session_state.get("active_plans", [])
    plan_lookup: dict[str, CuttingPlanRecord] = {p.plan_name: p for p in plans}
    for p in active_plans:
        plan_lookup[p.plan_name] = p
    for p in edited_records:
        plan_lookup[p.plan_name] = p

    if s1_results:
        default_s2_names = [r.plan_name for r in s1_results if r.success and r.plan_name in plan_lookup]
    else:
        default_s2_names = [p.plan_name for p in edited_records if p.plan_name in plan_lookup]
    default_s2_names.sort(key=plan_sort_key)

    available_s2_options = sorted(plan_lookup.keys(), key=plan_sort_key)
    available_s2_set = set(available_s2_options)
    saved_s2 = st.session_state.get("saved_s2_names")
    if "cfg_selected_s2_names" not in st.session_state:
        if saved_s2 is not None:
            st.session_state["cfg_selected_s2_names"] = sorted(
                [n for n in saved_s2 if n in available_s2_set],
                key=plan_sort_key,
            )
        else:
            st.session_state["cfg_selected_s2_names"] = list(default_s2_names)
    else:
        st.session_state["cfg_selected_s2_names"] = sorted(
            [n for n in st.session_state["cfg_selected_s2_names"] if n in available_s2_set],
            key=plan_sort_key,
        )

    selected_s2_names = st.multiselect(
        "Planuri de validat, generat QR și printat în Etapa 2",
        options=available_s2_options,
        key="cfg_selected_s2_names",
        help="Implicit sunt selectate planurile finalizate cu succes în Etapa 1. Puteți adăuga și planuri reparate manual sau elimina planuri goale.",
    )
    st.session_state["saved_s2_names"] = list(selected_s2_names)
    stage2_candidates = sorted(
        [plan_lookup[name] for name in selected_s2_names],
        key=lambda rec: plan_sort_key(rec.plan_name),
    )

    if "cfg_clean_red_stage2" not in st.session_state:
        st.session_state["cfg_clean_red_stage2"] = bool(
            st.session_state.get("saved_settings", {}).get("clean_red_stage2", False)
        )
    clean_red_stage2 = st.checkbox(
        "🧹 Curăță piesele ROȘII în Etapa 2 (înainte de validare)",
        key="cfg_clean_red_stage2",
        help="Opțiune separată pentru Pasul 2: șterge automat piesele ROȘII din Panoul 3 pentru fiecare plan selectat, imediat înainte de bifarea 'Valid' și generarea codului QR.",
        **_CB_PERSIST_KW,
    )

    bot_alive = current_bot is not None and current_bot.is_alive()
    if not bot_alive:
        st.info("ℹ️ Browserul nu este deschis momentan — la pornirea Etapei 2 se va deschide automat și va continua cu validarea.")

    if overwrite_original_word:
        save_dest_desc = "va **suprascrie documentele Word originale** (`.doc` / `.docx`) cu codul QR"
    else:
        save_dest_desc = f"va salva copii noi Word + QR în `{output_dir_str}`"

    clean_prefix = (
        "Va **curăța mai întâi piesele ROȘII** din Panoul 3, apoi va"
        if clean_red_stage2
        else "Va"
    )
    mode_info = (
        f"{clean_prefix} valida planurile în intranet, {save_dest_desc} și le va trimite la imprimanta **{printer_display}**."
        if send_to_printer
        else f"{clean_prefix} valida planurile, va extrage QR și {save_dest_desc} (fără printare)."
    )
    st.info(f"ℹ️ Mod curent: {mode_info}")

    if overwrite_original_word and (
        input_mode != INPUT_MODE_LOCAL
        or any(_is_uploaded_temp_word_path(p.source_word_path) for p in stage2_candidates)
    ):
        st.warning(
            "⚠️ Pentru fișierele încărcate prin **Drag & Drop**, browserul nu oferă calea originală de pe disc. "
            f"Acestea vor fi salvate în folderul `{output_dir_str}`. Pentru a suprascrie direct fișierele originale, folosiți modul **📁 Folder local**."
        )

    s2_btn_label = (
        "✅ Curăță ROȘU, Validează, Inserează QR & Printează"
        if clean_red_stage2
        else "✅ Validează, Inserează QR & Printează"
    )
    start_s2 = st.button(
        s2_btn_label,
        type="primary",
        disabled=(len(stage2_candidates) == 0),
    )

    progress_slot_s2 = st.empty()
    log_slot_s2 = st.empty()

    if start_s2:
        bot2 = _get_or_create_bot(force_new=False)
        fut2 = bot2.start_stage2_async(
            plans=stage2_candidates,
            output_dir=Path(output_dir_str),
            qr_temp_dir=Path("temp_qr"),
            send_to_printer=send_to_printer,
            print_delay=float(print_delay),
            printer_name=effective_printer_name or None,
            overwrite_original=bool(overwrite_original_word),
            clean_red_rows=bool(clean_red_stage2),
        )
        pbar2 = progress_slot_s2.progress(0.0, text="Inițializare Etapa 2...")

        while not fut2.done():
            prog = bot2.get_progress()
            tot = max(1, int(prog.get("total") or len(stage2_candidates)))
            cur = int(prog.get("current") or 0)
            pbar2.progress(min(1.0, cur / tot), text=prog.get("message") or "Procesare Etapa 2...")
            _render_logs_box(log_slot_s2, bot2.get_logs())
            time.sleep(0.25)

        try:
            s2_res: list[Stage2PlanResult] = fut2.result()
            st.session_state["bot_logs"] = bot2.get_logs()[-300:]
            st.session_state["stage2_results"] = s2_res
            st.session_state["stage"] = "stage2_done"
            _persist_session_if_changed(force=True)
            pbar2.progress(1.0, text="Etapa 2 finalizată!")
            st.rerun()
        except Exception as exc:
            logs_now = bot2.get_logs()
            st.session_state["bot_logs"] = logs_now[-300:]
            _render_logs_box(log_slot_s2, logs_now)
            _persist_session_if_changed(force=True)
            st.error(f"Etapa 2 a întâmpinat o eroare critică: {exc}")

s2_results: list[Stage2PlanResult] = st.session_state.get("stage2_results", [])
if s2_results:
    st.subheader("📄 Rezultate Etapa 2 & Previzualizare QR")
    s2_df = pd.DataFrame(
        [
            {
                "Plan croi": r.plan_name,
                "Status": "✅ OK" if r.success else "❌ Eroare",
                "Rânduri ROȘII șterse": getattr(r, "red_deleted", 0),
                "Fișier Word + QR": r.output_docx_path or "-",
                "Printat": "🖨️ Da" if r.printed else "Nu (test)",
                "Eroare": r.error or "-",
            }
            for r in s2_results
        ]
    )
    st.dataframe(s2_df, width="stretch", hide_index=True)

    qr_items = [r for r in s2_results if r.qr_path and Path(r.qr_path).exists()]
    if qr_items:
        cols = st.columns(min(4, len(qr_items)))
        for idx, item in enumerate(qr_items):
            with cols[idx % len(cols)]:
                st.image(item.qr_path, caption=item.plan_name, width=160)

# ==============================================================================
# Consolă Persistentă de Log-uri
# ==============================================================================
bot_for_logs: Optional[NorstalBrowserBot] = st.session_state.get("bot")
if bot_for_logs is not None:
    logs = bot_for_logs.get_logs()[-300:]
    if logs:
        st.session_state["bot_logs"] = logs
else:
    logs = st.session_state.get("bot_logs", [])

if logs:
    with st.expander("📋 Jurnal Execuție în Timp Real (Log-uri)", expanded=True):
        log_box = st.empty()
        _render_logs_box(log_box, logs)

_persist_settings_if_changed(_build_current_settings_dict())
_persist_session_if_changed()
