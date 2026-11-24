from __future__ import annotations

import base64
import concurrent.futures
import io
import json
import queue
import re
import threading
import time
import unicodedata
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Optional
from urllib.parse import parse_qsl, urlencode, urljoin, urlsplit, urlunsplit

from PIL import Image
from playwright.sync_api import Browser, BrowserContext, Frame, Locator, Page, Playwright, sync_playwright

from parser_docs import CuttingPlanRecord, read_parts_from_excel
from word_qr_printer import insert_qr_and_optionally_print

# ==============================================================================
# SELECTORI CONFIGURABILI PENTRU INTRANETUL NORSTAL (PHP + Bootstrap + jqGrid)
# Dacă structura DOM a paginii pro_introducere_planuri_croi.php diferă,
# ajustațiile pot fi făcute centralizat aici.
# ==============================================================================
SELECTORS: dict[str, Any] = {
    "login": {
        "login_link": "#log_in_a",
        "logout_indicator": "#log_out",
        "login_form": "#log_form",
        "username_inputs": [
            "#log_form input[type='text']",
            "#log_form input[name*='user' i]",
            "#log_form input[name*='nume' i]",
            "#log_form input[name*='mail' i]",
            "input[name='username']",
            "input[name='user']",
            "input[id*='user' i]",
        ],
        "password_inputs": [
            "#log_form input[type='password']",
            "input[type='password']",
        ],
        "submit_buttons": [
            "#log_form button[type='submit']",
            "#log_form input[type='submit']",
            "#log_form button",
            "#log_form input[type='button']",
            "button:has-text('Login')",
            "button:has-text('Autentificare')",
            "input[type='submit']",
        ],
    },
    "navbar": {
        "productie_menu": "#myNavbar a.dropdown-toggle:has-text('Productie')",
        "introducere_date_menu": "#myNavbar a:has-text('Introducere date')",
        "planuri_croi_link": "#myNavbar a[href*='pro_introducere_planuri_croi.php']",
    },
    "iframe": "iframe#iframe_a, iframe[name='iframe_a']",
    "maintenance_overlay": "#stop_working",
    "panels": {
        "panel1_titles": ["Productie-Plan de croi", "Productie - Plan de croi", "Plan de croi"],
        "panel2_titles": ["Comenzi / planul de croi", "Comenzi /planul de croi", "Comenzi"],
        "panel3_titles": ["Piese din comanda", "Piese"],
    },
    "jqgrid": {
        "grid_box": ".ui-jqgrid",
        "title_text": ".ui-jqgrid-titlebar .ui-jqgrid-title, .ui-jqgrid-caption",
        "add_button": "td[id^='add_'], [id^='add_'], .ui-pg-button:has(.ui-icon-plus), .ui-pg-button:has(.fa-plus), .ui-pg-button:has(.glyphicon-plus), .ui-pg-button[title*='Adaug' i], .ui-pg-button[title*='Add' i]",
        "edit_pager_button": "td[id^='edit_'], [id^='edit_'], .ui-pg-button:has(.ui-icon-pencil), .ui-pg-button:has(.fa-pencil), .ui-pg-button:has(.glyphicon-pencil), .ui-pg-button[title*='Modific' i]",
        "rows": "table.ui-jqgrid-btable tr.jqgrow:not(.jqgfirstrow)",
        "inline_edit": ".ui-inline-edit, div[id^='jEditButton_'], span.ui-icon-pencil, i.fa-pencil, i.fa-edit, span.glyphicon-pencil, [title*='Modific' i]",
        "inline_delete": ".ui-inline-del, div[id^='jDeleteButton_'], span.ui-icon-trash, i.fa-trash, i.fa-trash-o, span.glyphicon-trash, [title*='Șterg' i], [title*='Sterg' i]",
        "loading": ".loading, .jqgrid-overlay",
    },
    "modals": {
        "visible_dialog": ".ui-dialog:visible, .ui-jqdialog:visible, .modal.in, .modal.show, div[role='dialog']:visible",
        "form_error": "#FormError, .FormError",
        "save_buttons": [
            "a#sData",
            "button:has-text('Salvează')",
            "a:has-text('Salvează')",
            "button:has-text('Salveaza')",
            "a:has-text('Salveaza')",
            "input[type='button'][value*='Salveaz' i]",
            "input[type='submit'][value*='Salveaz' i]",
        ],
        "delete_confirm_buttons": [
            "a#dData",
            "button:has-text('Șterge')",
            "a:has-text('Șterge')",
            "button:has-text('Sterge')",
            "a:has-text('Sterge')",
            "input[value*='Sterg' i]",
            "input[value*='Șterg' i]",
        ],
        "import_buttons": [
            "button:has-text('Import')",
            "a:has-text('Import')",
            "input[type='button'][value*='Import' i]",
        ],
        "upload_buttons": [
            "button:has-text('Upload')",
            "a:has-text('Upload')",
            "input[type='submit'][value*='Upload' i]",
            "input[type='button'][value*='Upload' i]",
            "button:has-text('Încarcă')",
            "button:has-text('Incarca')",
        ],
        "close_buttons": [
            "button:has-text('Close')",
            "a:has-text('Close')",
            "button:has-text('Închide')",
            "button:has-text('Inchide')",
            "input[value*='Close' i]",
            ".ui-dialog-buttonpane button",
            "button.close",
            ".ui-dialog-titlebar-close",
        ],
        "cancel_buttons": [
            "a#cData",
            "button:has-text('Anulare')",
            "a:has-text('Anulare')",
            "button:has-text('Cancel')",
            "a:has-text('Cancel')",
            ".ui-jqdialog-titlebar-close",
            ".ui-dialog-titlebar-close",
        ],
    },
    "select2": {
        "dropdown_search_inputs": [
            ".select2-search--dropdown .select2-search__field",
            ".select2-container--open .select2-search__field",
            "#select2-drop input.select2-input",
            ".select2-search__field",
        ],
        "result_options": [
            ".select2-results__option:not(.loading-results):not(.select2-results__message)",
            "#select2-drop .select2-results li.select2-result-selectable",
            ".select2-results li:not(.select2-no-results)",
        ],
    },
}


_UPLOADS_DIR = (Path(__file__).resolve().parent / "work_tmp" / "uploads").resolve()


def _is_uploaded_temp_word_path(word_path: str | Path) -> bool:
    """Returnează True dacă calea Word provine din upload temporar Drag & Drop."""
    if not word_path or not str(word_path).strip():
        return False
    try:
        resolved = Path(word_path).resolve()
        resolved.relative_to(_UPLOADS_DIR)
        return True
    except (OSError, ValueError):
        return False


@dataclass
class Stage1PlanResult:
    plan_name: str
    success: bool
    matched_comanda: str = ""
    parts_added: int = 0
    parts_rejected: int = 0
    red_deleted: int = 0
    pink_kept: int = 0
    error: str = ""


@dataclass
class Stage2PlanResult:
    plan_name: str
    success: bool
    qr_path: str = ""
    output_docx_path: str = ""
    printed: bool = False
    error: str = ""


def _strip_accents(text: str) -> str:
    nfkd = unicodedata.normalize("NFKD", text or "")
    return "".join(ch for ch in nfkd if not unicodedata.combining(ch))


def _norm(text: str) -> str:
    return re.sub(r"\s+", " ", _strip_accents(text)).strip().lower()


def parse_rgb(color_str: str) -> Optional[tuple[int, int, int, float]]:
    """Parsează un string CSS rgb()/rgba() și returnează (r, g, b, alpha) sau None."""
    if not color_str:
        return None
    s = color_str.strip().lower()
    if s in {"transparent", "none", "inherit", "initial", ""}:
        return None
    m = re.match(
        r"rgba?\(\s*([0-9.]+)\s*[,\s]\s*([0-9.]+)\s*[,\s]\s*([0-9.]+)(?:\s*[,/]\s*([0-9.]+%?))?\s*\)",
        s,
    )
    if not m:
        return None
    r = int(round(float(m.group(1))))
    g = int(round(float(m.group(2))))
    b = int(round(float(m.group(3))))
    alpha = 1.0
    if m.group(4) is not None:
        raw_a = m.group(4).strip()
        alpha = float(raw_a[:-1]) / 100.0 if raw_a.endswith("%") else float(raw_a)
    if alpha <= 0.01:
        return None
    return r, g, b, alpha


def _is_red_rgb(r: int, g: int, b: int) -> bool:
    if r > 180 and g < 80 and b < 80:
        return True
    if r >= 150 and g < 90 and b < 90 and (r - max(g, b)) >= 100:
        return True
    return False


def _is_pink_rgb(r: int, g: int, b: int) -> bool:
    if _is_red_rgb(r, g, b):
        return False
    # Exclude alb/gri deschis și galbenul de selecție jqGrid (#fbec88: g=236, b=136)
    if r > 180 and g >= 95 and b >= 95 and (r - g) >= 15 and b >= (g - 30):
        return True
    return False


def classify_row_color(
    tr_bg: str,
    cell_bgs: list[str],
    first_cell_text: str,
    row_classes: str = "",
    cell_classes: Optional[list[str]] = None,
) -> tuple[str, str]:
    """
    Clasifică un rând din tabelul 3 în 'red', 'pink' sau 'normal'.

    Prioritate:
      1. Culoarea celulelor (prima celulă are prioritate față de tr, util dacă rândul
         este selectat în jqGrid și primește highlight pe tr)
      2. Culoarea de fundal a tr-ului
      3. Clase CSS sugestive (red/rosu/danger vs pink/roz)
      4. Textul primei celule == '0' -> 'red'
    """
    ordered_colors: list[tuple[str, str]] = []
    for idx, bg in enumerate(cell_bgs):
        ordered_colors.append((f"cell[{idx}]", bg))
    ordered_colors.append(("tr", tr_bg))

    for source, bg in ordered_colors:
        parsed = parse_rgb(bg)
        if not parsed:
            continue
        r, g, b, _ = parsed
        if _is_red_rgb(r, g, b):
            return "red", f"RGB roșu pe {source}: ({r}, {g}, {b})"
        if _is_pink_rgb(r, g, b):
            return "pink", f"RGB roz pe {source}: ({r}, {g}, {b})"

    all_cls = " ".join([row_classes or ""] + (cell_classes or [])).lower()
    if any(tok in all_cls for tok in ("pink", "roz", "warning-pink")):
        return "pink", f"clasă CSS roz ({all_cls.strip()})"
    if any(tok in all_cls for tok in ("row-red", "bg-danger", "text-danger", "rosu")):
        return "red", f"clasă CSS roșie ({all_cls.strip()})"

    if (first_cell_text or "").strip() == "0":
        return "red", "textul primei celule este '0'"

    return "normal", "fără fundal roșu/roz"


_STATIC_EXT_RE = re.compile(r"\.(?:js|css|png|jpe?g|gif|svg|woff2?|ttf|eot|ico|map)(?:\?|$)", re.IGNORECASE)


@dataclass
class _CapturedRequest:
    method: str
    url: str
    content_type: str
    post_data: str
    response_status: int
    response_text: str


@dataclass
class _FullHttpBlueprint:
    cutting_type: str
    p1_post_url: str
    p1_post_fmt: str
    p1_post_payload: Any
    p1_plan_key: str
    p1_get_url: str
    p1_id_key: str
    p1_match_key: str
    p1_post_id_hint: Optional[tuple[str, Any]]
    p2_get_url: str
    p2_get_plan_params: list[str]
    p2_post_url: str
    p2_post_url_plan_params: list[str]
    p2_post_fmt: str
    p2_post_payload: Any
    p2_post_body_plan_keys: list[str]
    p2_comanda_key: str
    p2_weight_key: str
    p2_id_key: str
    p2_post_id_hint: Optional[tuple[str, Any]]
    p3_get_url: str
    p3_get_rowid_params: list[str]
    p3_get_comanda_params: list[str]
    p3_get_plan_params: list[str]
    p3_post_url: str
    p3_post_url_rowid_params: list[str]
    p3_post_url_comanda_params: list[str]
    p3_post_url_plan_params: list[str]
    p3_post_fmt: str
    p3_post_payload: Any
    p3_post_body_rowid_keys: list[str]
    p3_post_body_comanda_keys: list[str]
    p3_post_body_plan_keys: list[str]
    p3_part_key: str
    p3_qty_key: str
    sample_plan_id: str
    sample_rowid2: str
    sample_comanda_id: str
    comanda_options: list[tuple[str, str]] = field(default_factory=list)
    comanda_cache: dict[str, tuple[str, str]] = field(default_factory=dict)
    p1_get_method: str = "GET"
    p1_get_post_data: str = ""
    p1_get_content_type: str = ""
    p2_get_method: str = "GET"
    p2_get_post_data: str = ""
    p2_get_content_type: str = ""
    p2_get_body_plan_keys: list[str] = field(default_factory=list)
    p3_get_method: str = "GET"
    p3_get_post_data: str = ""
    p3_get_content_type: str = ""
    p3_get_body_rowid_keys: list[str] = field(default_factory=list)
    p3_get_body_comanda_keys: list[str] = field(default_factory=list)
    p3_get_body_plan_keys: list[str] = field(default_factory=list)


class _HttpTrafficRecorder:
    """Înregistrează pasiv request-urile XHR/fetch/document (GET/POST) pentru calibrare Full HTTP."""

    def __init__(self, page: Page) -> None:
        self._page = page
        self.events: list[_CapturedRequest] = []
        self._active = True
        self._page.on("response", self._on_response)

    def _on_response(self, resp: Any) -> None:
        if not self._active:
            return
        try:
            req = resp.request
            method = (req.method or "").upper()
            if method not in {"GET", "POST"}:
                return
            url = req.url or ""
            if not url or "check_if_stop" in url or _STATIC_EXT_RE.search(url):
                return
            rtype = (req.resource_type or "").lower()
            if rtype and rtype not in {"xhr", "fetch", "document", "other"}:
                return
            try:
                ct = req.headers.get("content-type", "")
            except Exception:
                ct = ""
            try:
                post_data = req.post_data or ""
            except Exception:
                post_data = ""
            try:
                body_txt = resp.text() or ""
            except Exception:
                body_txt = ""
            self.events.append(
                _CapturedRequest(
                    method=method,
                    url=url,
                    content_type=ct,
                    post_data=post_data,
                    response_status=int(resp.status),
                    response_text=body_txt,
                )
            )
        except Exception:
            pass

    def mark(self) -> int:
        return len(self.events)

    def since(self, idx: int) -> list[_CapturedRequest]:
        return list(self.events[idx:])

    def stop(self) -> None:
        self._active = False
        try:
            self._page.remove_listener("response", self._on_response)
        except Exception:
            pass


class NorstalBrowserBot:
    """
    Automatizare Playwright pentru intranetul NORSTAL pe un singur thread dedicat.

    Păstrează aceeași instanță de browser, context și pagină între Etapa 1 și Etapa 2,
    permițând apeluri sigure din thread-urile efemere ale Streamlit.
    """

    def __init__(
        self,
        base_url: str = "http://10.0.1.132/NORSTAL/",
        username: str = "",
        password: str = "",
        headless: bool = False,
        timeout_ms: int = 20000,
    ) -> None:
        self.base_url = base_url.rstrip("/") + "/"
        self.username = username
        self.password = password
        self.headless = headless
        self.timeout_ms = timeout_ms

        self._lock = threading.Lock()
        self._logs: list[str] = []
        self._progress: dict[str, Any] = {
            "active": False,
            "stage": "",
            "current": 0,
            "total": 0,
            "plan": "",
            "message": "",
        }
        self._browser_open = False
        self._closed = False
        self._last_p3_capture: Optional[dict[str, Any]] = None

        self._playwright: Optional[Playwright] = None
        self._browser: Optional[Browser] = None
        self._context: Optional[BrowserContext] = None
        self._page: Optional[Page] = None

        self._cmd_queue: queue.Queue[
            tuple[Optional[Callable[..., Any]], tuple[Any, ...], dict[str, Any], Optional[concurrent.futures.Future]]
        ] = queue.Queue()
        self._worker = threading.Thread(
            target=self._worker_loop,
            name="NorstalPlaywrightWorker",
            daemon=True,
        )
        self._worker.start()

    # ------------------------------------------------------------------
    # Logging & Progress (Thread-safe)
    # ------------------------------------------------------------------
    def _log(self, msg: str, level: str = "INFO") -> None:
        ts = datetime.now().strftime("%H:%M:%S")
        line = f"[{ts}] [{level}] {msg}"
        with self._lock:
            self._logs.append(line)
            try:
                log_path = Path("work_tmp/bot.log")
                log_path.parent.mkdir(parents=True, exist_ok=True)
                with log_path.open("a", encoding="utf-8") as fh:
                    fh.write(line + "\n")
            except OSError:
                pass

    def get_logs(self) -> list[str]:
        with self._lock:
            return list(self._logs)

    def clear_logs(self) -> None:
        with self._lock:
            self._logs.clear()

    def _set_progress(
        self,
        *,
        active: Optional[bool] = None,
        stage: Optional[str] = None,
        current: Optional[int] = None,
        total: Optional[int] = None,
        plan: Optional[str] = None,
        message: Optional[str] = None,
    ) -> None:
        with self._lock:
            if active is not None:
                self._progress["active"] = active
            if stage is not None:
                self._progress["stage"] = stage
            if current is not None:
                self._progress["current"] = current
            if total is not None:
                self._progress["total"] = total
            if plan is not None:
                self._progress["plan"] = plan
            if message is not None:
                self._progress["message"] = message

    def get_progress(self) -> dict[str, Any]:
        with self._lock:
            return dict(self._progress)

    def is_alive(self) -> bool:
        with self._lock:
            return not self._closed and self._worker.is_alive() and self._browser_open

    # ------------------------------------------------------------------
    # Worker Thread Lifecycle
    # ------------------------------------------------------------------
    def _worker_loop(self) -> None:
        while True:
            fn, args, kwargs, fut = self._cmd_queue.get()
            if fn is None:
                try:
                    self._teardown_browser()
                finally:
                    if fut is not None and not fut.done():
                        fut.set_result(True)
                    self._cmd_queue.task_done()
                break

            try:
                result = fn(*args, **kwargs)
                if fut is not None and not fut.done():
                    fut.set_result(result)
            except Exception as exc:
                self._log(f"Eroare execuție worker: {exc}", level="ERROR")
                if fut is not None and not fut.done():
                    fut.set_exception(exc)
            finally:
                self._cmd_queue.task_done()

    def _submit(self, fn: Callable[..., Any], *args: Any, **kwargs: Any) -> concurrent.futures.Future:
        if self._closed:
            raise RuntimeError("NorstalBrowserBot a fost închis.")
        fut: concurrent.futures.Future = concurrent.futures.Future()
        self._cmd_queue.put((fn, args, kwargs, fut))
        return fut

    def close(self, timeout: float = 10.0) -> None:
        if self._closed:
            return
        self._closed = True
        fut: concurrent.futures.Future = concurrent.futures.Future()
        self._cmd_queue.put((None, (), {}, fut))
        try:
            fut.result(timeout=timeout)
        except Exception:
            pass
        self._worker.join(timeout=timeout)

    def _teardown_browser(self) -> None:
        for obj, method in (
            (self._page, "close"),
            (self._context, "close"),
            (self._browser, "close"),
            (self._playwright, "stop"),
        ):
            if obj is not None:
                try:
                    getattr(obj, method)()
                except Exception:
                    pass
        self._page = None
        self._context = None
        self._browser = None
        self._playwright = None
        with self._lock:
            self._browser_open = False

    def _ensure_browser(self) -> Page:
        if self._page is not None and not self._page.is_closed():
            return self._page

        self._teardown_browser()
        self._log(f"Pornire Chromium Playwright (headless={self.headless})...")
        self._playwright = sync_playwright().start()
        launch_args = ["--start-maximized"] if not self.headless else []
        self._browser = self._playwright.chromium.launch(
            headless=self.headless,
            args=launch_args,
        )
        if self.headless:
            self._context = self._browser.new_context(viewport={"width": 1600, "height": 950})
        else:
            self._context = self._browser.new_context(no_viewport=True)

        self._context.set_default_timeout(self.timeout_ms)
        self._page = self._context.new_page()
        self._page.on("dialog", self._handle_dialog)
        with self._lock:
            self._browser_open = True
        return self._page

    def _handle_dialog(self, dialog: Any) -> None:
        try:
            self._log(f"Dialog browser ({dialog.type}): {dialog.message}", level="WARN")
            dialog.accept()
        except Exception:
            pass

    # ------------------------------------------------------------------
    # Public Async / Sync Entrypoints
    # ------------------------------------------------------------------
    def start_stage1_async(
        self,
        plans: list[CuttingPlanRecord],
        panel3_mode: str = "manual_plus",
    ) -> concurrent.futures.Future:
        return self._submit(self._run_stage1, plans, panel3_mode)

    def start_stage2_async(
        self,
        plans: list[CuttingPlanRecord],
        output_dir: Path | str = "output_print",
        qr_temp_dir: Path | str = "temp_qr",
        send_to_printer: bool = False,
        print_delay: float = 3.0,
        printer_name: Optional[str] = None,
        overwrite_original: bool = True,
    ) -> concurrent.futures.Future:
        return self._submit(
            self._run_stage2,
            plans,
            Path(output_dir),
            Path(qr_temp_dir),
            send_to_printer,
            print_delay,
            printer_name,
            overwrite_original,
        )

    def run_custom_task(self, fn: Callable[[Page, "NorstalBrowserBot"], Any]) -> concurrent.futures.Future:
        """Permite rularea unei operații pe thread-ul dedicat (util și pentru teste)."""
        def _runner() -> Any:
            page = self._ensure_browser()
            return fn(page, self)

        return self._submit(_runner)

    # ------------------------------------------------------------------
    # Frame, Maintenance & AJAX Synchronization
    # ------------------------------------------------------------------
    def _wait_no_maintenance(self, page: Page, max_wait_s: float = 120.0) -> None:
        overlay = page.locator(SELECTORS["maintenance_overlay"]).first
        try:
            if overlay.count() > 0 and overlay.is_visible():
                self._log("Overlay de mentenanță (#stop_working) activ; se așteaptă...", level="WARN")
                overlay.wait_for(state="hidden", timeout=int(max_wait_s * 1000))
                self._log("Overlay de mentenanță dispărut; execuția continuă.")
        except Exception:
            pass

    def _get_work_frame(self, page: Page) -> Page | Frame:
        iframe_loc = page.locator(SELECTORS["iframe"]).first
        if iframe_loc.count() > 0:
            elem = iframe_loc.element_handle()
            if elem is not None:
                frame = elem.content_frame()
                if frame is not None:
                    return frame
        return page

    def _wait_frame_ajax(
        self,
        page: Page,
        frame: Page | Frame,
        timeout_ms: Optional[int] = None,
        wait_network_idle: bool = True,
    ) -> None:
        t_ms = timeout_ms or self.timeout_ms
        self._wait_no_maintenance(page)
        try:
            frame.wait_for_function(
                "() => !window.jQuery || typeof window.jQuery.active !== 'number' || window.jQuery.active === 0",
                timeout=t_ms,
            )
        except Exception:
            pass

        # Așteaptă ascunderea overlay-urilor 'Loading...' din jqGrid
        try:
            frame.wait_for_function(
                """() => {
                    const els = document.querySelectorAll('.loading, .jqgrid-overlay');
                    for (const el of els) {
                        const style = window.getComputedStyle(el);
                        if (style.display !== 'none' && style.visibility !== 'hidden' && el.offsetParent !== null) {
                            return false;
                        }
                    }
                    return true;
                }""",
                timeout=t_ms,
            )
        except Exception:
            pass

        if wait_network_idle:
            try:
                frame.wait_for_load_state("networkidle", timeout=1500)
            except Exception:
                pass
        time.sleep(0.15)

    def _save_error_screenshot(self, page: Page, plan_name: str, step_tag: str) -> None:
        try:
            err_dir = Path("work_tmp/errors")
            err_dir.mkdir(parents=True, exist_ok=True)
            safe_name = re.sub(r"[^A-Za-z0-9._-]+", "_", f"{plan_name}_{step_tag}")
            out_png = err_dir / f"{safe_name}_{datetime.now().strftime('%H%M%S')}.png"
            page.screenshot(path=str(out_png), full_page=True)
            self._log(f"Screenshot eroare salvat: {out_png}", level="WARN")
        except Exception:
            pass

    # ------------------------------------------------------------------
    # Login & Navigation
    # ------------------------------------------------------------------
    def _login_and_navigate(self, page: Page) -> Page | Frame:
        # Fast-path: dacă pagina de lucru este deja deschisă și autentificată, nu reîncărca index.php
        try:
            existing_frame = self._get_work_frame(page)
            logout_loc = page.locator(SELECTORS["login"]["logout_indicator"]).first
            if (
                page.url
                and self.base_url.rstrip("/").lower() in page.url.lower()
                and logout_loc.count() > 0
                and logout_loc.is_visible()
                and existing_frame.locator(SELECTORS["jqgrid"]["grid_box"]).count() >= 3
            ):
                self._wait_frame_ajax(page, existing_frame)
                return existing_frame
        except Exception:
            pass

        index_url = urljoin(self.base_url, "index.php")
        self._log(f"Deschidere intranet: {index_url}")
        page.goto(index_url, wait_until="domcontentloaded")
        self._wait_no_maintenance(page)

        logout_loc = page.locator(SELECTORS["login"]["logout_indicator"]).first
        already_logged = False
        try:
            already_logged = logout_loc.count() > 0 and logout_loc.is_visible()
        except Exception:
            already_logged = False

        if not already_logged:
            self._log("Sesiune neautentificată; se inițiază login-ul...")
            login_link = page.locator(SELECTORS["login"]["login_link"]).first
            if login_link.count() > 0 and login_link.is_visible():
                login_link.click()
                page.wait_for_timeout(300)

            pwd_input = self._first_visible(page, SELECTORS["login"]["password_inputs"])
            if pwd_input is not None:
                if not self.username or not self.password:
                    if self.headless:
                        raise RuntimeError(
                            "Formularul de login este activ, dar lipsesc credențialele (User/Parolă) în modul Headless."
                        )
                    self._log(
                        "Credențialele nu sunt completate; autentificați-vă manual în browser (timeout 120s)...",
                        level="WARN",
                    )
                    logout_loc.wait_for(state="visible", timeout=120000)
                else:
                    user_input = self._first_visible(page, SELECTORS["login"]["username_inputs"])
                    if user_input is None:
                        raise RuntimeError("Nu s-a găsit câmpul de utilizator în formularul de login.")
                    user_input.fill(self.username)
                    pwd_input.fill(self.password)
                    submit_btn = self._first_visible(page, SELECTORS["login"]["submit_buttons"])
                    if submit_btn is not None:
                        submit_btn.click()
                    else:
                        pwd_input.press("Enter")

                    try:
                        logout_loc.wait_for(state="visible", timeout=self.timeout_ms)
                    except Exception as exc:
                        raise RuntimeError(
                            "Autentificarea în intranet nu a confirmat sesiunea (#log_out a rămas ascuns). Verificați user/parola."
                        ) from exc
                    self._log("Autentificare reușită.")
        else:
            self._log("Sesiune deja autentificată în intranet.")

        # Navigare prin meniu: Productie -> Introducere date -> Introducere planuri de croi
        self._log("Navigare către Productie -> Introducere date -> Introducere planuri de croi...")
        nav_ok = False
        try:
            prod_menu = page.locator(SELECTORS["navbar"]["productie_menu"]).first
            prod_menu.hover(timeout=3000)
            page.wait_for_timeout(150)

            intro_menu = page.locator(SELECTORS["navbar"]["introducere_date_menu"]).first
            if not intro_menu.is_visible():
                prod_menu.click(timeout=2000)
                page.wait_for_timeout(150)
            intro_menu.hover(timeout=3000)
            page.wait_for_timeout(150)

            target_link = page.locator(SELECTORS["navbar"]["planuri_croi_link"]).first
            if not target_link.is_visible():
                intro_menu.click(timeout=2000)
                page.wait_for_timeout(150)
            if target_link.is_visible():
                target_link.click(timeout=3000)
                nav_ok = True
        except Exception:
            nav_ok = False

        if not nav_ok:
            # Fallback direct în iframe_a sau link direct
            self._log("Se folosește fallback-ul de încărcare directă în iframe_a...", level="WARN")
            iframe_loc = page.locator(SELECTORS["iframe"]).first
            if iframe_loc.count() > 0:
                page.evaluate(
                    "() => { const f = document.getElementById('iframe_a'); if (f) f.src = 'pro_introducere_planuri_croi.php'; }"
                )
            else:
                page.goto(urljoin(self.base_url, "pro_introducere_planuri_croi.php"), wait_until="domcontentloaded")

        # Așteaptă încărcarea tabelelor jqGrid în iframe_a
        deadline = time.time() + (self.timeout_ms / 1000.0)
        frame = self._get_work_frame(page)
        while time.time() < deadline:
            frame = self._get_work_frame(page)
            try:
                if frame.locator(SELECTORS["jqgrid"]["grid_box"]).count() >= 1:
                    break
            except Exception:
                pass
            time.sleep(0.25)

        self._wait_frame_ajax(page, frame)
        grids_count = frame.locator(SELECTORS["jqgrid"]["grid_box"]).count()
        if grids_count == 0:
            raise RuntimeError("Pagina 'Introducere planuri de croi' nu a încărcat niciun tabel jqGrid.")
        self._log(f"Pagină încărcată ({grids_count} panouri jqGrid detectate).")
        return frame

    # ------------------------------------------------------------------
    # Generic Locator & Grid Helpers
    # ------------------------------------------------------------------
    def _first_visible(self, scope: Page | Frame | Locator, selectors: list[str]) -> Optional[Locator]:
        for sel in selectors:
            try:
                loc = scope.locator(sel)
                cnt = loc.count()
                for idx in range(cnt):
                    cand = loc.nth(idx)
                    if cand.is_visible():
                        return cand
            except Exception:
                continue
        return None

    def _find_panel_grid(self, frame: Page | Frame, panel_key: str, fallback_index: int) -> Locator:
        """
        Găsește containerul `.ui-jqgrid` pentru unul dintre cele 3 panouri
        folosind titlul panoului (și fallback pe ordinea de la stânga la dreapta).
        """
        title_candidates = [_norm(t) for t in SELECTORS["panels"][panel_key]]
        idx = frame.evaluate(
            """({candidates, fallbackIndex}) => {
                const strip = (s) => (s || '')
                    .normalize('NFKD')
                    .replace(/[\\u0300-\\u036f]/g, '')
                    .replace(/\\s+/g, ' ')
                    .trim()
                    .toLowerCase();

                const grids = Array.from(document.querySelectorAll('.ui-jqgrid'));
                if (!grids.length) return -1;

                // Sortare vizuală stânga->dreapta pentru fallback determinist
                const entries = grids.map((g, domIdx) => {
                    const rect = g.getBoundingClientRect();
                    const titleEl = g.querySelector('.ui-jqgrid-titlebar .ui-jqgrid-title, .ui-jqgrid-caption');
                    const panelParent = g.closest('.panel, .box, .card, fieldset') || g.parentElement;
                    const headerEl = panelParent ? panelParent.querySelector('.panel-heading, .box-header, .card-header, legend, h3, h4') : null;
                    const texts = [
                        titleEl ? titleEl.textContent : '',
                        headerEl ? headerEl.textContent : ''
                    ].map(strip);
                    return { domIdx, x: rect.left, y: rect.top, texts };
                });

                for (const c of candidates) {
                    for (const e of entries) {
                        if (e.texts.some(t => t && (t === c || t.includes(c)))) {
                            return e.domIdx;
                        }
                    }
                }

                const sorted = [...entries].sort((a, b) => (a.x - b.x) || (a.y - b.y));
                if (fallbackIndex < sorted.length) {
                    return sorted[fallbackIndex].domIdx;
                }
                return -1;
            }""",
            {"candidates": title_candidates, "fallbackIndex": fallback_index},
        )
        if idx < 0:
            raise RuntimeError(f"Nu s-a putut localiza panoul jqGrid pentru '{panel_key}'.")
        return frame.locator(SELECTORS["jqgrid"]["grid_box"]).nth(idx)

    def _get_visible_modal(self, page: Page, frame: Page | Frame) -> Optional[Locator]:
        for scope in (frame, page):
            loc = scope.locator(SELECTORS["modals"]["visible_dialog"])
            cnt = loc.count()
            # Ia ultimul dialog vizibil (de obicei cel aflat deasupra în z-index)
            for idx in range(cnt - 1, -1, -1):
                cand = loc.nth(idx)
                try:
                    if cand.is_visible():
                        return cand
                except Exception:
                    continue
        return None

    def _wait_visible_modal(self, page: Page, frame: Page | Frame, timeout_ms: Optional[int] = None) -> Locator:
        deadline = time.time() + ((timeout_ms or self.timeout_ms) / 1000.0)
        while time.time() < deadline:
            modal = self._get_visible_modal(page, frame)
            if modal is not None:
                return modal
            time.sleep(0.1)
        raise TimeoutError("Nu a apărut niciun modal vizibil în timpul așteptat.")

    def _wait_modal_closed(self, page: Page, frame: Page | Frame, modal: Locator, timeout_ms: Optional[int] = None) -> None:
        t_ms = timeout_ms or self.timeout_ms
        try:
            modal.wait_for(state="hidden", timeout=t_ms)
        except Exception:
            # Verifică dacă a rămas o eroare de validare în modal
            err_loc = modal.locator(SELECTORS["modals"]["form_error"]).first
            if err_loc.count() > 0 and err_loc.is_visible():
                err_txt = (err_loc.inner_text() or "").strip()
                if err_txt:
                    raise RuntimeError(f"Eroare formular intranet: {err_txt}")
            raise
        self._wait_overlays_hidden(page, frame, timeout_ms=min(5000, t_ms))
        self._wait_frame_ajax(page, frame)

    def _wait_overlays_hidden(self, page: Page, frame: Page | Frame, timeout_ms: int = 5000) -> None:
        for scope in (frame, page):
            try:
                ov = scope.locator(".ui-widget-overlay:visible").first
                if ov.count() > 0 and ov.is_visible():
                    ov.wait_for(state="hidden", timeout=timeout_ms)
            except Exception:
                pass

    def _dismiss_any_open_modals(self, page: Page, frame: Page | Frame) -> None:
        """Închide defensiv orice dialog/overlay rămas deschis între planuri."""
        for _ in range(5):
            modal = self._get_visible_modal(page, frame)
            if modal is None:
                break
            closed_attempt = False
            close_btn = self._first_visible(
                modal,
                [
                    ".ui-dialog-titlebar-close",
                    ".ui-jqdialog-titlebar-close",
                    "button.close",
                    *SELECTORS["modals"]["close_buttons"],
                ],
            )
            if close_btn is not None:
                try:
                    close_btn.click(timeout=2000)
                    closed_attempt = True
                except Exception:
                    pass
            if not closed_attempt:
                try:
                    page.keyboard.press("Escape")
                except Exception:
                    pass
            try:
                modal.wait_for(state="hidden", timeout=2500)
            except Exception:
                break
        self._wait_overlays_hidden(page, frame, timeout_ms=3000)

    def _click_save_in_modal(self, page: Page, frame: Page | Frame, modal: Locator) -> None:
        btn = self._first_visible(modal, SELECTORS["modals"]["save_buttons"])
        if btn is None:
            raise RuntimeError("Nu s-a găsit butonul 'Salvează' în modal.")
        btn.click()
        # Așteaptă puțin pentru eventuală eroare de formular înainte de închidere
        time.sleep(0.2)
        err_loc = modal.locator(SELECTORS["modals"]["form_error"]).first
        if err_loc.count() > 0 and err_loc.is_visible():
            err_txt = (err_loc.inner_text() or "").strip()
            if err_txt:
                raise RuntimeError(f"Eroare la salvare în intranet: {err_txt}")
        self._wait_modal_closed(page, frame, modal)

    def _fill_modal_field(self, modal: Locator, label_variants: list[str], value: str) -> None:
        norm_variants = [_norm(v) for v in label_variants]
        filled = modal.evaluate(
            """(el, {variants, val}) => {
                const strip = (s) => (s || '')
                    .normalize('NFKD')
                    .replace(/[\\u0300-\\u036f]/g, '')
                    .replace(/\\s+/g, ' ')
                    .trim()
                    .toLowerCase();

                const setVal = (input) => {
                    input.focus();
                    input.value = val;
                    input.dispatchEvent(new Event('input', { bubbles: true }));
                    input.dispatchEvent(new Event('change', { bubbles: true }));
                    input.dispatchEvent(new Event('blur', { bubbles: true }));
                };

                // 1. Căutare în rândurile jqGrid FormEdit (tr.FormData, apoi .form-group, apoi tr fără tr imbricate)
                const formRows = Array.from(el.querySelectorAll('tr.FormData'));
                const extraRows = Array.from(el.querySelectorAll('.form-group')).concat(
                    Array.from(el.querySelectorAll('tr')).filter(r => !r.classList.contains('FormData') && !r.querySelector('tr'))
                );
                const rows = formRows.concat(extraRows);
                for (const r of rows) {
                    const caption = r.querySelector('.CaptionTD, label, td:first-child');
                    const txt = strip(caption ? caption.textContent : '');
                    if (txt && variants.some(v => txt === v || txt.includes(v))) {
                        const input = r.querySelector('input:not([type="hidden"]):not([type="checkbox"]):not([type="button"]):not([type="submit"]), textarea');
                        if (input && !input.disabled && !input.readOnly) {
                            setVal(input);
                            return true;
                        }
                    }
                }

                // 2. Fallback după atributul name/id al input-ului
                const inputs = Array.from(el.querySelectorAll('input:not([type="hidden"]), textarea'));
                for (const inp of inputs) {
                    const key = strip((inp.name || '') + ' ' + (inp.id || '') + ' ' + (inp.placeholder || ''));
                    if (variants.some(v => key.includes(v.replace(/\\s+/g, '')) || key.includes(v))) {
                        setVal(inp);
                        return true;
                    }
                }
                return false;
            }""",
            {"variants": norm_variants, "val": value},
        )
        if not filled:
            raise RuntimeError(f"Nu s-a găsit câmpul {label_variants} în modal.")

    def _inspect_modal_field(self, modal: Locator, label_variants: list[str]) -> Optional[dict[str, str]]:
        """Citește valoarea și atributele name/id ale unui câmp din modalul curent."""
        norm_variants = [_norm(v) for v in label_variants]
        return modal.evaluate(
            """(el, variants) => {
                const strip = (s) => (s || '')
                    .normalize('NFKD')
                    .replace(/[\\u0300-\\u036f]/g, '')
                    .replace(/\\s+/g, ' ')
                    .trim()
                    .toLowerCase();

                const formRows = Array.from(el.querySelectorAll('tr.FormData'));
                const extraRows = Array.from(el.querySelectorAll('.form-group')).concat(
                    Array.from(el.querySelectorAll('tr')).filter(r => !r.classList.contains('FormData') && !r.querySelector('tr'))
                );
                const rows = formRows.concat(extraRows);
                for (const r of rows) {
                    const caption = r.querySelector('.CaptionTD, label, td:first-child');
                    const txt = strip(caption ? caption.textContent : '');
                    if (txt && variants.some(v => txt === v || txt.includes(v))) {
                        const input = r.querySelector('input:not([type="hidden"]):not([type="checkbox"]):not([type="button"]):not([type="submit"]), textarea');
                        if (input && !input.disabled && !input.readOnly) {
                            return { value: input.value || '', name: input.name || '', id: input.id || '' };
                        }
                    }
                }
                const inputs = Array.from(el.querySelectorAll('input:not([type="hidden"]), textarea'));
                for (const inp of inputs) {
                    const key = strip((inp.name || '') + ' ' + (inp.id || '') + ' ' + (inp.placeholder || ''));
                    if (variants.some(v => key.includes(v.replace(/\\s+/g, '')) || key.includes(v))) {
                        return { value: inp.value || '', name: inp.name || '', id: inp.id || '' };
                    }
                }
                return null;
            }""",
            norm_variants,
        )

    def _fill_and_verify_modal_field(
        self,
        modal: Locator,
        label_variants: list[str],
        value: str,
        max_attempts: int = 3,
    ) -> dict[str, str]:
        """
        Completează un câmp din modal și verifică efectiv că valoarea a rămas în DOM
        (previne pierderea valorii dacă jqGrid resetează formularul imediat după deschidere).
        """
        expected = value.strip()
        last_info: Optional[dict[str, str]] = None
        for _ in range(max_attempts):
            self._fill_modal_field(modal, label_variants, value)
            time.sleep(0.08)
            last_info = self._inspect_modal_field(modal, label_variants)
            if last_info is not None and (last_info.get("value") or "").strip() == expected:
                return last_info
            time.sleep(0.10)
        raise RuntimeError(
            f"Câmpul {label_variants} nu a reținut valoarea '{value}' în modal (valoare curentă: {last_info})."
        )

    def _select_modal_option_if_present(self, modal: Locator, label_variants: list[str], target_text: str) -> None:
        norm_variants = [_norm(v) for v in label_variants]
        modal.evaluate(
            """(el, {variants, target}) => {
                const strip = (s) => (s || '')
                    .normalize('NFKD')
                    .replace(/[\\u0300-\\u036f]/g, '')
                    .replace(/\\s+/g, ' ')
                    .trim()
                    .toLowerCase();
                const tNorm = strip(target);

                const rows = Array.from(el.querySelectorAll('tr.FormData, tr, .form-group'));
                for (const r of rows) {
                    const caption = r.querySelector('.CaptionTD, label, td:first-child');
                    const txt = strip(caption ? caption.textContent : '');
                    if (txt && variants.some(v => txt === v || txt.includes(v))) {
                        const sel = r.querySelector('select');
                        if (sel) {
                            const opt = Array.from(sel.options).find(o => strip(o.textContent).includes(tNorm) || strip(o.value).includes(tNorm));
                            if (opt && sel.value !== opt.value) {
                                sel.value = opt.value;
                                sel.dispatchEvent(new Event('change', { bubbles: true }));
                            }
                            return true;
                        }
                    }
                }
                return false;
            }""",
            {"variants": norm_variants, "target": target_text},
        )

    def _select_debitare_dropdown(self, page: Page, frame: Page | Frame, cutting_type: str) -> None:
        """Selectează tipul de debitare din dropdown-ul 'Debitare' din antetul Panoului 1."""
        target = _norm(cutting_type)
        result = frame.evaluate(
            """(targetNorm) => {
                const strip = (s) => (s || '')
                    .normalize('NFKD')
                    .replace(/[\\u0300-\\u036f]/g, '')
                    .replace(/\\s+/g, ' ')
                    .trim()
                    .toLowerCase();

                const selects = Array.from(document.querySelectorAll('select'));
                let bestSel = null;
                let bestScore = -1;
                for (const s of selects) {
                    const optTexts = Array.from(s.options).map(o => strip(o.textContent));
                    const nearLabel = s.closest('div, td, label, .form-group, tr')?.textContent || '';
                    const idName = (s.id || '') + ' ' + (s.name || '');
                    let score = 0;
                    if (strip(idName).includes('debitare')) score += 10;
                    if (strip(nearLabel).includes('debitare')) score += 6;
                    if (optTexts.some(t => t.includes('plasma'))) score += 5;
                    if (s.offsetParent !== null) score += 2;
                    if (score >= 5 && score > bestScore) {
                        bestScore = score;
                        bestSel = s;
                    }
                }
                if (!bestSel) return { found: false, changed: false };

                const opts = Array.from(bestSel.options);
                const availableOptions = opts.map(o => o.textContent.trim()).filter(Boolean);
                let match = opts.find(o => {
                    const t = strip(o.textContent);
                    const v = strip(o.value);
                    return t === targetNorm || t.includes(targetNorm) || v === targetNorm;
                });

                let fallbackUsed = false;
                if (!match && targetNorm === 'laser') {
                    match = opts.find(o => strip(o.textContent).includes('plasma') || strip(o.value).includes('plasma'));
                    if (match) fallbackUsed = true;
                }
                if (!match) {
                    return { found: true, matched: false, availableOptions };
                }

                const changed = bestSel.value !== match.value;
                bestSel.value = match.value;
                bestSel.dispatchEvent(new Event('change', { bubbles: true }));
                if (window.jQuery) {
                    window.jQuery(bestSel).val(match.value).trigger('change');
                }
                return {
                    found: true,
                    matched: true,
                    changed,
                    fallbackUsed,
                    text: match.textContent.trim(),
                    availableOptions,
                };
            }""",
            target,
        )
        if result.get("found") and result.get("matched"):
            if result.get("fallbackUsed"):
                self._log(
                    f"Opțiunea '{cutting_type}' nu există în dropdown-ul Debitare; s-a selectat automat '{result.get('text')}' (schimbat={result.get('changed')})."
                )
            else:
                self._log(f"Debitare selectată: '{result.get('text')}' (schimbat={result.get('changed')})")
            self._wait_frame_ajax(page, frame)
        elif not result.get("found"):
            self._log(
                "Nu s-a detectat dropdown-ul extern 'Debitare'; se va verifica la nivelul modalului.",
                level="WARN",
            )
        else:
            avail = result.get("availableOptions") or []
            raise RuntimeError(
                f"Nu s-a găsit opțiunea '{cutting_type}' în dropdown-ul Debitare (disponibile: {avail})."
            )

    def _clear_grid_search_filters(self, page: Page, frame: Page | Frame, grid: Locator) -> None:
        """Golește filtrele din bara de căutare jqGrid (`Caută...`) dacă au rămas completate."""
        cleared = False
        try:
            search_inputs = grid.locator("tr.ui-search-toolbar input:visible")
            cnt = search_inputs.count()
            last_cleared_inp = None
            for i in range(cnt):
                inp = search_inputs.nth(i)
                try:
                    if (inp.input_value() or "").strip():
                        inp.fill("")
                        last_cleared_inp = inp
                        cleared = True
                except Exception:
                    continue
            if last_cleared_inp is not None:
                try:
                    last_cleared_inp.press("Enter")
                except Exception:
                    pass
        except Exception:
            cleared = False

        if cleared:
            self._wait_frame_ajax(page, frame)

    def _find_and_select_grid_row(
        self,
        page: Page,
        frame: Page | Frame,
        grid: Locator,
        match_text: str,
        col_hints: list[str],
        click_to_select: bool = True,
    ) -> Locator:
        """
        Găsește un rând din jqGrid care conține `match_text`. Dacă nu e vizibil în pagina curentă,
        încearcă filtrarea folosind bara de căutare jqGrid pe coloana potrivită.
        """
        target_norm = _norm(match_text)

        def _locate_in_dom() -> Optional[Locator]:
            rows = grid.locator(SELECTORS["jqgrid"]["rows"])
            cnt = rows.count()
            # Căutare exactă întâi, apoi substring
            for exact in (True, False):
                for idx in range(cnt):
                    r = rows.nth(idx)
                    try:
                        txt = _norm(r.inner_text())
                    except Exception:
                        continue
                    if exact:
                        cell_texts = [_norm(c) for c in r.locator("td").all_inner_texts()]
                        if target_norm in cell_texts:
                            return r
                    else:
                        if target_norm in txt:
                            return r
            return None

        row = _locate_in_dom()
        filter_used = False
        if row is None:
            # Încearcă filtrul din toolbar-ul jqGrid (ex: coloana 'Plan croi')
            filter_used = bool(
                grid.evaluate(
                    """(gridEl, {colHints, query}) => {
                        const strip = (s) => (s || '')
                            .normalize('NFKD')
                            .replace(/[\\u0300-\\u036f]/g, '')
                            .replace(/\\s+/g, ' ')
                            .trim()
                            .toLowerCase();

                        const headers = Array.from(gridEl.querySelectorAll('.ui-jqgrid-htable tr.ui-jqgrid-labels th'));
                        const searchCells = Array.from(gridEl.querySelectorAll('.ui-jqgrid-htable tr.ui-search-toolbar th'));
                        if (!headers.length || !searchCells.length) return false;

                        for (let i = 0; i < headers.length && i < searchCells.length; i++) {
                            const hTxt = strip(headers[i].textContent);
                            if (colHints.some(h => hTxt.includes(strip(h)))) {
                                const inp = searchCells[i].querySelector('input');
                                if (inp) {
                                    inp.focus();
                                    inp.value = query;
                                    inp.dispatchEvent(new Event('input', { bubbles: true }));
                                    inp.dispatchEvent(new Event('change', { bubbles: true }));
                                    return true;
                                }
                            }
                        }
                        return false;
                    }""",
                    {"colHints": col_hints, "query": match_text},
                )
            )
            if filter_used:
                try:
                    search_inputs = grid.locator("tr.ui-search-toolbar input:visible")
                    for i in range(search_inputs.count()):
                        val = search_inputs.nth(i).input_value()
                        if val == match_text:
                            search_inputs.nth(i).press("Enter")
                            break
                except Exception:
                    pass
                self._wait_frame_ajax(page, frame)
                row = _locate_in_dom()

        if row is None:
            if filter_used:
                self._clear_grid_search_filters(page, frame, grid)
            raise RuntimeError(f"Nu s-a găsit rândul pentru '{match_text}' în tabelul jqGrid.")

        if click_to_select:
            self._click_grid_row_to_select(page, frame, row)
        return row

    def _click_grid_row_to_select(self, page: Page, frame: Page | Frame, row: Locator) -> None:
        """Dă click pe o celulă de date (nu pe coloana Actions) pentru a declanșa onSelectRow."""
        cells = row.locator("td:visible")
        clicked = False
        for idx in range(cells.count()):
            cell = cells.nth(idx)
            try:
                aria = (cell.get_attribute("aria-describedby") or "").lower()
                if any(skip in aria for skip in ("act", "action", "cb", "rn")):
                    continue
                cell.click()
                clicked = True
                break
            except Exception:
                continue
        if not clicked:
            row.click()
        self._wait_frame_ajax(page, frame)

    def _read_modal_select_info(
        self, modal: Locator, label_hints: list[str]
    ) -> tuple[str, str, list[tuple[str, str]]]:
        """Citește (name/id câmp, valoare curentă, opțiuni [(val, text)]) pentru un dropdown din modal."""
        norm_hints = [_norm(h) for h in label_hints]
        try:
            res = modal.evaluate(
                """(el, hints) => {
                    const strip = (s) => (s || '')
                        .normalize('NFKD')
                        .replace(/[\\u0300-\\u036f]/g, '')
                        .replace(/\\s+/g, ' ')
                        .trim()
                        .toLowerCase();
                    const rows = Array.from(el.querySelectorAll('tr.FormData, .form-group, tr'));
                    for (const r of rows) {
                        const txt = strip(r.textContent);
                        if (!hints.some(h => txt.includes(h))) continue;
                        const sel = r.querySelector('select');
                        if (sel) {
                            const fieldName = sel.name || sel.id || '';
                            const opts = Array.from(sel.options)
                                .map(o => [String(o.value || '').trim(), String(o.textContent || '').trim()])
                                .filter(([v, t]) => v || t);
                            let val = String(sel.value || '').trim();
                            if (!val && sel.selectedIndex >= 0 && sel.options[sel.selectedIndex]) {
                                val = String(sel.options[sel.selectedIndex].value || '').trim();
                            }
                            return { fieldName, val, opts };
                        }
                        const inp = r.querySelector('input[type="hidden"], input.select2-offscreen, input');
                        if (inp) {
                            const fieldName = inp.name || inp.id || '';
                            const val = String(inp.value || '').trim();
                            const s2 = r.querySelector('.select2-selection__rendered, .select2-chosen');
                            const s2Txt = s2 ? String(s2.textContent || '').trim() : '';
                            return { fieldName, val, opts: (val || s2Txt) ? [[val, s2Txt]] : [] };
                        }
                    }
                    return { fieldName: '', val: '', opts: [] };
                }""",
                norm_hints,
            )
            field_name = str(res.get("fieldName") or "")
            val = str(res.get("val") or "")
            opts = [(str(o[0]), str(o[1])) for o in (res.get("opts") or []) if len(o) >= 2]
            return field_name, val, opts
        except Exception:
            return "", "", []

    def _select2_choose(
        self,
        page: Page,
        frame: Page | Frame,
        modal: Locator,
        label_hints: list[str],
        query: str,
    ) -> str:
        """
        Deschide dropdown-ul Select2 pentru 'Comanda', tastează `query` (ex: '1242 - 49.1'),
        așteaptă strict o opțiune filtrată care conține `query` și verifică selecția în DOM.
        """
        norm_hints = [_norm(h) for h in label_hints]
        q_norm = _norm(query)

        # Așteaptă stabilizarea inițializării modalului / Select2 (evită resetul afterShowForm)
        self._wait_frame_ajax(page, frame, wait_network_idle=False)
        time.sleep(0.25)

        field_container = modal
        rows = modal.locator("tr.FormData, .form-group, tr")
        for idx in range(rows.count()):
            cand = rows.nth(idx)
            try:
                txt = _norm(cand.inner_text())
                if any(h in txt for h in norm_hints):
                    field_container = cand
                    break
            except Exception:
                continue

        def _read_selected_text() -> str:
            try:
                res = field_container.evaluate(
                    """(el) => {
                        const sel = el.querySelector('select');
                        const selTxt = (sel && sel.selectedIndex >= 0 && sel.options[sel.selectedIndex])
                            ? (sel.options[sel.selectedIndex].textContent || '').trim()
                            : '';
                        const s2 = el.querySelector('.select2-selection__rendered, .select2-chosen');
                        const s2Txt = s2 ? (s2.textContent || '').trim() : '';
                        return { selTxt, s2Txt };
                    }"""
                )
                sel_txt = res.get("selTxt") or ""
                s2_txt = res.get("s2Txt") or ""
                if s2_txt and q_norm in _norm(s2_txt):
                    return s2_txt
                if sel_txt and q_norm in _norm(sel_txt):
                    return sel_txt
                return s2_txt or sel_txt
            except Exception:
                return ""

        for attempt in range(3):
            s2_trigger = field_container.locator(
                ".select2-selection, .select2-choice, .select2-container a, .select2-container"
            ).first

            chosen_text = ""
            if s2_trigger.count() > 0 and s2_trigger.is_visible():
                s2_trigger.click()
                time.sleep(0.2)

                search_input: Optional[Locator] = None
                active_scope: Page | Frame = frame
                for scope in (frame, page):
                    inp = self._first_visible(scope, SELECTORS["select2"]["dropdown_search_inputs"])
                    if inp is not None:
                        search_input = inp
                        active_scope = scope
                        break

                if search_input is not None:
                    search_input.fill("")
                    search_input.type(query, delay=35)
                    time.sleep(0.30)
                    self._wait_frame_ajax(page, frame, wait_network_idle=False)
                else:
                    active_scope = frame

                deadline = time.time() + 10.0
                matched_opt: Optional[Locator] = None
                while time.time() < deadline:
                    for sel in SELECTORS["select2"]["result_options"]:
                        opts = active_scope.locator(sel)
                        cnt = opts.count()
                        if cnt == 0:
                            continue
                        for i in range(cnt):
                            opt = opts.nth(i)
                            try:
                                if not opt.is_visible():
                                    continue
                                txt = (opt.inner_text() or "").strip()
                                if not txt or _norm(txt) in {"searching...", "cautare...", "loading..."}:
                                    continue
                                if q_norm in _norm(txt):
                                    matched_opt = opt
                                    chosen_text = txt
                                    break
                            except Exception:
                                continue
                        if matched_opt is not None:
                            break
                    if matched_opt is not None:
                        break
                    time.sleep(0.15)

                if matched_opt is None:
                    try:
                        page.keyboard.press("Escape")
                    except Exception:
                        pass
                else:
                    matched_opt.click()
                    self._wait_frame_ajax(page, frame, wait_network_idle=False)
            else:
                native_select = field_container.locator("select").first
                if native_select.count() > 0:
                    picked = native_select.evaluate(
                        """(sel, qNorm) => {
                            const strip = (s) => (s || '')
                                .normalize('NFKD')
                                .replace(/[\\u0300-\\u036f]/g, '')
                                .replace(/\\s+/g, ' ')
                                .trim()
                                .toLowerCase();
                            const opt = Array.from(sel.options).find(o => strip(o.textContent).includes(qNorm));
                            if (!opt) return null;
                            sel.value = opt.value;
                            sel.dispatchEvent(new Event('change', { bubbles: true }));
                            if (window.jQuery) {
                                window.jQuery(sel).val(opt.value).trigger('change');
                            }
                            return opt.textContent.trim();
                        }""",
                        q_norm,
                    )
                    if picked:
                        chosen_text = str(picked)
                        self._wait_frame_ajax(page, frame, wait_network_idle=False)
                else:
                    raise RuntimeError(
                        f"Nu s-a găsit dropdown-ul Select2/'<select>' pentru 'Comanda' (căutare '{query}')."
                    )

            time.sleep(0.15)
            verified_txt = _read_selected_text()
            if verified_txt and q_norm in _norm(verified_txt):
                return verified_txt
            if chosen_text and q_norm in _norm(chosen_text) and not verified_txt:
                return chosen_text

            if attempt < 2:
                self._log(
                    f"Selecția Comanda ('{verified_txt or 'gol'}') nu corespunde cu '{query}'; reîncercare ({attempt + 2}/3)...",
                    level="WARN",
                )
                time.sleep(0.25)

        raise RuntimeError(
            f"Nu s-a putut selecta și confirma comanda '{query}' în dropdown-ul Select2."
        )

    # ------------------------------------------------------------------
    # Stage 1: Upload & Clean Red/Pink Rows
    # ------------------------------------------------------------------
    def _find_panel3_import_button(self, frame: Page | Frame, panel3_grid: Locator) -> Locator:
        # Caută mai întâi în interiorul sau în părintele imediat al Panoului 3
        btn_idx = frame.evaluate(
            """(gridEl) => {
                const candidates = [];
                let cur = gridEl;
                for (let depth = 0; depth < 4 && cur; depth++) {
                    const btns = Array.from(cur.querySelectorAll('button, a.btn, input[type="button"], a'));
                    for (const b of btns) {
                        const txt = (b.textContent || b.value || '').trim().toLowerCase();
                        if (txt.includes('import') && b.offsetParent !== null) {
                            candidates.push(b);
                        }
                    }
                    if (candidates.length) break;
                    cur = cur.parentElement;
                }
                const target = candidates[0];
                if (!target) return -1;
                const allBtns = Array.from(document.querySelectorAll('button, a, input[type="button"]'));
                return allBtns.indexOf(target);
            }""",
            panel3_grid.element_handle(),
        )
        if btn_idx >= 0:
            return frame.locator("button, a, input[type='button']").nth(btn_idx)

        fallback = self._first_visible(frame, SELECTORS["modals"]["import_buttons"])
        if fallback is not None:
            return fallback
        raise RuntimeError("Nu s-a găsit butonul 'Import' pentru Panoul 3.")

    def _find_import_scope_and_file_input(
        self,
        page: Page,
        frame: Page | Frame,
        import_modal: Locator,
        timeout_ms: int = 15000,
    ) -> tuple[Page | Frame | Locator, Locator]:
        """
        Caută `<input type='file'>` atât direct în dialogul de Import, cât și în orice
        `<iframe>` intern al popup-ului jQuery UI (așteptând încărcarea asincronă).
        """
        deadline = time.time() + (timeout_ms / 1000.0)
        while time.time() < deadline:
            # 1. Direct în containerul modalului
            try:
                direct = import_modal.locator("input[type='file']").first
                if direct.count() > 0:
                    return import_modal, direct
            except Exception:
                pass

            # 2. În iframe-urile din interiorul modalului (.ui-dialog iframe)
            try:
                iframes = import_modal.locator("iframe")
                for i in range(iframes.count()):
                    elem = iframes.nth(i).element_handle()
                    if elem is None:
                        continue
                    cf = elem.content_frame()
                    if cf is not None:
                        cand = cf.locator("input[type='file']").first
                        if cand.count() > 0:
                            return cf, cand
            except Exception:
                pass

            # 3. În frame-urile copil ale frame-ului de lucru (iframe_a)
            if isinstance(frame, Frame):
                for cf in frame.child_frames:
                    try:
                        cand = cf.locator("input[type='file']").first
                        if cand.count() > 0:
                            return cf, cand
                    except Exception:
                        continue

            # 4. În toate frame-urile paginii (ex. dialog montat la nivelul părinte)
            for pf in page.frames:
                if pf == page.main_frame or pf == frame:
                    continue
                try:
                    cand = pf.locator("input[type='file']").first
                    if cand.count() > 0:
                        return pf, cand
                except Exception:
                    continue

            time.sleep(0.2)

        raise RuntimeError(
            "Nu s-a găsit <input type='file'> nici în modalul de Import, nici în iframe-ul său intern."
        )

    def _upload_excel_in_import_modal(
        self,
        page: Page,
        frame: Page | Frame,
        import_modal: Locator,
        excel_path: str | Path,
        plan_name: str,
    ) -> None:
        """Gestionează selecția fișierului .xlsx, apăsarea Upload și închiderea popup-ului Import."""
        upload_scope, file_input = self._find_import_scope_and_file_input(page, frame, import_modal)
        file_input.set_input_files(str(Path(excel_path).resolve()))
        self._log(f"[{plan_name}] Încărcare Excel: {Path(excel_path).name}")

        upload_selectors = SELECTORS["modals"]["upload_buttons"] + [
            "input[type='submit']",
            "button[type='submit']",
        ]
        upload_btn = self._first_visible(upload_scope, upload_selectors)
        if upload_btn is None and upload_scope is not import_modal:
            upload_btn = self._first_visible(import_modal, upload_selectors)
        if upload_btn is None:
            raise RuntimeError("Nu s-a găsit butonul 'Upload' în modalul de Import.")

        clicked = False
        try:
            with page.expect_response(
                lambda r: r.request.method == "POST" and "check_if_stop" not in r.url,
                timeout=max(self.timeout_ms, 30000),
            ):
                upload_btn.click()
                clicked = True
        except Exception as exc:
            if not clicked:
                raise
            self._log(f"[{plan_name}] Avertisment așteptare POST Upload: {exc}", level="WARN")

        if isinstance(upload_scope, Frame):
            try:
                upload_scope.wait_for_load_state("domcontentloaded", timeout=10000)
            except Exception:
                pass
        self._wait_frame_ajax(page, frame, timeout_ms=max(self.timeout_ms, 30000))

        # Loghează un scurt rezumat dacă popup-ul afișează confirmare/mesaj după upload
        try:
            body_text = (
                upload_scope.locator("body").inner_text()
                if isinstance(upload_scope, Frame)
                else import_modal.inner_text()
            )
            compact = re.sub(r"\s+", " ", body_text or "").strip()
            if compact and "alegeti fisierul pt import" not in compact.lower():
                self._log(f"[{plan_name}] Răspuns popup Import: {compact[:160]}")
        except Exception:
            pass

        # Închide modalul dacă nu s-a auto-închis după upload
        still_open = False
        try:
            still_open = import_modal.is_visible()
        except Exception:
            still_open = False

        if still_open:
            close_btn = None
            if upload_scope is not import_modal:
                close_btn = self._first_visible(upload_scope, SELECTORS["modals"]["close_buttons"][:-1])
            if close_btn is None:
                close_btn = self._first_visible(import_modal, SELECTORS["modals"]["close_buttons"])

            if close_btn is not None:
                close_btn.click()
                self._wait_modal_closed(page, frame, import_modal)
            else:
                self._dismiss_any_open_modals(page, frame)
        else:
            self._wait_overlays_hidden(page, frame, timeout_ms=5000)
            self._wait_frame_ajax(page, frame)

    def _close_add_part_modal(self, page: Page, frame: Page | Frame) -> None:
        """Închide modalul de adăugare piese dacă a rămas deschis (closeAfterAdd: false)."""
        open_modal = self._get_visible_modal(page, frame)
        if open_modal is None:
            return
        cancel_btn = self._first_visible(open_modal, SELECTORS["modals"]["cancel_buttons"])
        if cancel_btn is not None:
            try:
                cancel_btn.click()
                self._wait_modal_closed(page, frame, open_modal)
                return
            except Exception:
                pass
        self._dismiss_any_open_modals(page, frame)

    def _clear_modal_form_error(self, modal: Locator) -> None:
        try:
            modal.evaluate(
                """(el) => {
                    const err = el.querySelector('#FormError, .FormError');
                    if (err) {
                        err.style.display = 'none';
                        const td = err.querySelector('td');
                        if (td) td.textContent = '';
                    }
                }"""
            )
        except Exception:
            pass

    def _read_modal_form_error(self, modal: Locator) -> str:
        try:
            if not modal.is_visible():
                return ""
            err_loc = modal.locator(SELECTORS["modals"]["form_error"]).first
            if err_loc.count() > 0 and err_loc.is_visible():
                return (err_loc.inner_text() or "").strip()
        except Exception:
            pass
        return ""

    def _add_parts_via_plus_list(
        self,
        page: Page,
        frame: Page | Frame,
        panel3_grid: Locator,
        parts: list[tuple[str, int]],
        plan_name: str,
        start_index: int = 1,
        total_parts: Optional[int] = None,
    ) -> tuple[int, int]:
        """Introduce o listă de piese prin modalul '+' al Panoului 3, cu verificare & retry."""
        tot = total_parts or len(parts)
        parts_added = 0
        parts_rejected = 0

        for offset, (part_name, qty) in enumerate(parts):
            idx = start_index + offset
            for attempt in range(3):
                modal = self._get_visible_modal(page, frame)
                if modal is None:
                    add3 = self._first_visible(panel3_grid, [SELECTORS["jqgrid"]["add_button"]])
                    if add3 is None:
                        raise RuntimeError("Nu s-a găsit butonul '+' în Panoul 3.")
                    add3.click()
                    modal = self._wait_visible_modal(page, frame)
                    self._wait_frame_ajax(page, frame)
                    # Așteaptă finalizarea inițializării jqGrid (afterShowForm / reset)
                    time.sleep(0.30)

                self._fill_and_verify_modal_field(
                    modal,
                    ["Denumire piesa", "Denumire piesă", "Denumire", "Piesa"],
                    part_name,
                )
                self._fill_and_verify_modal_field(
                    modal,
                    ["Cant.", "Cant", "Cantitate", "Cantitatea"],
                    str(qty),
                )
                self._clear_modal_form_error(modal)

                save_btn = self._first_visible(modal, SELECTORS["modals"]["save_buttons"])
                if save_btn is None:
                    raise RuntimeError("Nu s-a găsit butonul 'Salvează' în modalul de adăugare piesă.")
                save_btn.click()

                self._wait_frame_ajax(page, frame, wait_network_idle=False)
                time.sleep(0.12)

                err_txt = self._read_modal_form_error(modal)
                if err_txt:
                    err_norm = _norm(err_txt)
                    if ("obligatoriu" in err_norm or "required" in err_norm) and attempt < 2:
                        self._log(
                            f"  [{plan_name}] [{idx}/{tot}] Câmp resetat la '{part_name}'; reîncercare ({attempt + 2}/3)...",
                            level="WARN",
                        )
                        time.sleep(0.25)
                        continue
                    parts_rejected += 1
                    self._log(
                        f"  [{plan_name}] [{idx}/{tot}] Piesa '{part_name}' (cant. {qty}) respinsă: {err_txt}",
                        level="WARN",
                    )
                    break

                parts_added += 1
                self._log(f"  [{plan_name}] [{idx}/{tot}] Adăugat: {part_name} x {qty}")
                break

        self._close_add_part_modal(page, frame)
        self._wait_frame_ajax(page, frame)
        return parts_added, parts_rejected

    def _add_parts_via_plus_in_panel3(
        self,
        page: Page,
        frame: Page | Frame,
        panel3_grid: Locator,
        excel_path: str | Path,
        plan_name: str,
    ) -> tuple[int, int]:
        """
        Citește piesele (Denumire piesă, Cantitate) din .xlsx și le introduce pe rând
        prin butonul '+' ('Adăugare înregistrare') din Panoul 3.
        """
        parts = read_parts_from_excel(excel_path)
        if not parts:
            raise RuntimeError(f"Fișierul Excel '{Path(excel_path).name}' nu conține piese valide.")

        self._log(f"[{plan_name}] Adăugare {len(parts)} piese prin '+' în Panoul 3...")
        return self._add_parts_via_plus_list(
            page=page,
            frame=frame,
            panel3_grid=panel3_grid,
            parts=parts,
            plan_name=plan_name,
        )

    _JQGRID_PAGING_KEYS = {
        "_search",
        "nd",
        "rows",
        "page",
        "sidx",
        "sord",
        "filters",
        "searchfield",
        "searchoper",
        "searchstring",
    }

    @staticmethod
    def _parse_post_payload(post_data: str, content_type: str) -> Optional[tuple[str, Any]]:
        """Parsează corpul unui POST ca ('json', dict) sau ('form', list[(k, v)])."""
        if not post_data:
            return None
        is_json = "application/json" in (content_type or "").lower() or post_data.lstrip().startswith("{")
        if is_json:
            try:
                obj = json.loads(post_data)
            except Exception:
                return None
            if isinstance(obj, dict):
                return "json", obj
            return None
        pairs = parse_qsl(post_data, keep_blank_values=True)
        if pairs:
            return "form", pairs
        return None

    @staticmethod
    def _encode_post_payload(
        fmt: str,
        base_payload: Any,
        updates: dict[str, Any],
        frame_url: str,
    ) -> tuple[Any, dict[str, str]]:
        """Construiește body-ul și headerele XHR pentru un replay HTTP POST."""
        headers = {
            "X-Requested-With": "XMLHttpRequest",
            "Referer": frame_url,
        }
        if fmt == "json":
            payload = dict(base_payload) if isinstance(base_payload, dict) else {}
            payload.update(updates)
            headers["Content-Type"] = "application/json; charset=UTF-8"
            return json.dumps(payload), headers

        base_pairs = list(base_payload) if isinstance(base_payload, list) else []
        updated_pairs: list[tuple[str, str]] = []
        seen: set[str] = set()
        for k, v in base_pairs:
            if k in updates:
                updated_pairs.append((k, str(updates[k])))
                seen.add(k)
            else:
                updated_pairs.append((k, str(v)))
        for k, v in updates.items():
            if k not in seen:
                updated_pairs.append((k, str(v)))
        headers["Content-Type"] = "application/x-www-form-urlencoded; charset=UTF-8"
        return urlencode(updated_pairs), headers

    @staticmethod
    def _split_query_raw(query: str) -> list[tuple[str, str, bool]]:
        """Păstrează distincția între flag-uri fără '=' (?list2) și parametri ?k=v."""
        items: list[tuple[str, str, bool]] = []
        if not query:
            return items
        for seg in query.split("&"):
            if seg == "":
                continue
            has_eq = "=" in seg
            parsed = parse_qsl(seg, keep_blank_values=True)
            if parsed:
                k, v = parsed[0]
                items.append((k, v, has_eq))
        return items

    @classmethod
    def _replace_url_params(
        cls,
        url: str,
        param_updates: dict[str, str],
        refresh_nd: bool = False,
        strip_jqgrid_paging: bool = False,
        add_missing: bool = False,
    ) -> str:
        """Actualizează selectiv parametrii query dintr-un URL păstrând flag-urile fără '='."""
        sp = urlsplit(url)
        raw_items = cls._split_query_raw(sp.query)
        out_parts: list[str] = []
        seen: set[str] = set()
        nd_val = str(int(time.time() * 1000))

        for k, v, has_eq in raw_items:
            k_low = k.lower()
            if strip_jqgrid_paging and k_low in cls._JQGRID_PAGING_KEYS:
                continue
            if k in param_updates:
                out_parts.append(urlencode([(k, str(param_updates[k]))]))
                seen.add(k)
            elif refresh_nd and k_low == "nd":
                out_parts.append(urlencode([(k, nd_val)]))
                seen.add(k)
            elif has_eq:
                out_parts.append(urlencode([(k, v)]))
                seen.add(k)
            else:
                out_parts.append(k)
                seen.add(k)

        if add_missing:
            for k, v in param_updates.items():
                if k not in seen:
                    out_parts.append(urlencode([(k, str(v))]))

        return urlunsplit((sp.scheme, sp.netloc, sp.path, "&".join(out_parts), sp.fragment))

    @classmethod
    def _classify_id_params(
        cls,
        pairs: list[tuple[str, str]],
        targets: dict[str, str],
        role_hints: dict[str, tuple[str, ...]],
        extra_skip_keys: Optional[set[str]] = None,
    ) -> dict[str, list[str]]:
        """
        Identifică numele parametrilor care conțin ID-urile țintă (`plan`, `rowid`, `comanda`),
        rezolvând sigur eventualele coliziuni numerice (ex. două ID-uri egale cu '1').
        """
        skip = set(cls._JQGRID_PAGING_KEYS) | {"oper", "grid_id", "_empty"}
        if extra_skip_keys:
            skip |= {s.lower() for s in extra_skip_keys}

        classified: dict[str, list[str]] = {role: [] for role in targets}
        for k, v in pairs:
            k_low = k.lower()
            if k_low in skip:
                continue
            val_s = str(v).strip()
            if not val_s:
                continue
            matched_roles = [r for r, tval in targets.items() if tval and val_s == tval]
            if not matched_roles:
                continue
            if len(matched_roles) == 1:
                chosen = matched_roles[0]
            else:
                best_role = matched_roles[0]
                best_score = -1
                for r in matched_roles:
                    for hint in role_hints.get(r, ()):
                        if k_low == hint:
                            score = 100 + len(hint)
                        elif hint in k_low:
                            score = 10 + len(hint)
                        else:
                            score = 0
                        if score > best_score:
                            best_score = score
                            best_role = r
                chosen = best_role
            if k not in classified[chosen]:
                classified[chosen].append(k)
        return classified

    @classmethod
    def _payload_to_pairs(cls, fmt: str, base_payload: Any) -> list[tuple[str, str]]:
        if fmt == "json" and isinstance(base_payload, dict):
            return [(str(k), str(v)) for k, v in base_payload.items()]
        if isinstance(base_payload, list):
            return [(str(k), str(v)) for k, v in base_payload]
        return []

    @staticmethod
    def _detect_post_id_hint(response_text: str, known_id: str) -> Optional[tuple[str, Any]]:
        """Detectează cum întoarce serverul noul ID după un POST de inserare."""
        target = (known_id or "").strip()
        body = (response_text or "").strip()
        if not target or not body:
            return None
        if body == target:
            return ("plain", None)
        try:
            data = json.loads(body)
        except Exception:
            return None
        if isinstance(data, list):
            for idx, item in enumerate(data):
                if str(item).strip() == target:
                    return ("list", idx)
        elif isinstance(data, dict):
            for k, v in data.items():
                if str(v).strip() == target:
                    return ("dict", k)
        return None

    @staticmethod
    def _extract_id_from_post_response(response_text: str, hint: Optional[tuple[str, Any]]) -> str:
        body = (response_text or "").strip()
        if not body or hint is None:
            return ""
        kind, key = hint
        if kind == "plain":
            return body if re.match(r"^[A-Za-z0-9_-]+$", body) else ""
        try:
            data = json.loads(body)
        except Exception:
            return ""
        if kind == "list" and isinstance(data, list) and isinstance(key, int) and 0 <= key < len(data):
            val = str(data[key]).strip()
            return val if val and val.lower() not in {"true", "false", "none", "null"} else ""
        if kind == "dict" and isinstance(data, dict) and key in data:
            val = str(data[key]).strip()
            return val if val and val.lower() not in {"true", "false", "none", "null"} else ""
        return ""

    @staticmethod
    def _parse_jqgrid_json_rows(body_text: str) -> Optional[list[Any]]:
        body = (body_text or "").strip()
        if not body or not body.startswith(("{", "[")):
            return None
        try:
            data = json.loads(body)
        except Exception:
            return None
        if isinstance(data, dict) and isinstance(data.get("rows"), list):
            return data["rows"]
        if isinstance(data, list):
            return data
        return None

    @staticmethod
    def _jqgrid_row_values(row: Any, id_key: str = "", match_key: str = "") -> tuple[str, list[str], str]:
        """Returnează (row_id, toate textele celulelor, textul coloanei dedicate)."""
        if not isinstance(row, dict):
            return "", [], ""
        row_id = str(row.get(id_key) or row.get("id") or "").strip()
        cell = row.get("cell")
        if isinstance(cell, list):
            vals = [str(x or "").strip() for x in cell]
            return row_id, vals, ""
        vals = [str(v or "").strip() for k, v in row.items()]
        match_val = str(row.get(match_key) or "").strip() if match_key else ""
        return row_id, vals, match_val

    @classmethod
    def _detect_jqgrid_keys(
        cls, rows: list[Any], known_id: str, match_value: str
    ) -> tuple[str, str]:
        """Identifică (id_key, match_key) pentru rânduri JSON jqGrid în format obiect sau cell[]."""
        target_id = (known_id or "").strip()
        target_norm = _norm(match_value)
        for row in rows:
            if not isinstance(row, dict):
                continue
            if isinstance(row.get("cell"), list):
                rid = str(row.get("id") or "").strip()
                if not target_id or rid == target_id or any(target_norm and target_norm in _norm(str(c)) for c in row["cell"]):
                    return "", ""
                continue
            rid_key = ""
            m_key = ""
            for k, v in row.items():
                vs = str(v or "").strip()
                if target_id and vs == target_id and not rid_key:
                    rid_key = k
                if target_norm and _norm(vs) == target_norm:
                    m_key = k
            if not m_key and target_norm:
                for k, v in row.items():
                    if target_norm in _norm(str(v or "")):
                        m_key = k
                        break
            if rid_key or m_key:
                return rid_key or "id", m_key
        return "id", ""

    @classmethod
    def _find_row_id_in_jqgrid_json(
        cls,
        rows: list[Any],
        match_text: str,
        id_key: str = "",
        match_key: str = "",
        single_row_fallback: bool = False,
    ) -> str:
        """Găsește row_id-ul corespunzător din răspunsul JSON jqGrid (exact înainte de substring)."""
        target_norm = _norm(match_text)

        def _sort_key(rid: str) -> tuple[int, int | str]:
            try:
                return (1, int(rid))
            except ValueError:
                return (0, rid)

        exact_hits: list[str] = []
        sub_hits: list[str] = []
        for row in rows:
            rid, vals, dedicated_val = cls._jqgrid_row_values(row, id_key=id_key, match_key=match_key)
            if not rid:
                continue
            check_vals = [dedicated_val] if dedicated_val else vals
            norm_vals = [_norm(x) for x in check_vals if x]
            if target_norm and any(x == target_norm for x in norm_vals):
                exact_hits.append(rid)
            elif target_norm and any(target_norm in x for x in norm_vals):
                sub_hits.append(rid)

        if exact_hits:
            return max(exact_hits, key=_sort_key)
        if sub_hits:
            return max(sub_hits, key=_sort_key)
        if single_row_fallback and len(rows) == 1:
            rid, _, _ = cls._jqgrid_row_values(rows[0], id_key=id_key, match_key=match_key)
            return rid
        return ""

    def _build_replay_payload_template(
        self,
        post_data: str,
        content_type: str,
        first_part: str,
        first_qty: int,
        dom_part_name: str = "",
        dom_qty_name: str = "",
    ) -> Optional[tuple[str, Any, str, str]]:
        """Identifică formatul și cheile exacte pentru piesă/cantitate din POST-ul capturat."""
        parsed = self._parse_post_payload(post_data, content_type)
        if parsed is None:
            return None
        fmt, payload_data = parsed
        qty_str = str(first_qty)

        if fmt == "json":
            obj = payload_data
            part_key = dom_part_name if dom_part_name in obj else ""
            qty_key = dom_qty_name if dom_qty_name in obj else ""
            if not part_key:
                for k, v in obj.items():
                    if str(v).strip() == first_part:
                        part_key = k
                        break
            if not qty_key:
                for k, v in obj.items():
                    if k != part_key and str(v).strip() == qty_str and k.lower() not in {"id", "oper", "id_plan", "id_comanda"}:
                        qty_key = k
                        break
            if part_key and qty_key:
                return "json", obj, part_key, qty_key
            return None

        pairs = payload_data
        keys = [k for k, _ in pairs]
        part_key = dom_part_name if dom_part_name in keys else ""
        qty_key = dom_qty_name if dom_qty_name in keys else ""

        if not part_key:
            part_cands = [k for k, v in pairs if v.strip() == first_part]
            if part_cands:
                preferred = [k for k in part_cands if any(tok in k.lower() for tok in ("piesa", "part", "denum"))]
                part_key = preferred[0] if preferred else part_cands[0]

        if not qty_key:
            skip_keys = {"id", "oper", "id_plan", "id_comanda", "_empty"}
            qty_cands = [
                k for k, v in pairs
                if k != part_key and k.lower() not in skip_keys and v.strip() == qty_str
            ]
            if qty_cands:
                preferred = [k for k in qty_cands if any(tok in k.lower() for tok in ("cant", "qty"))]
                qty_key = preferred[0] if preferred else qty_cands[0]

        if part_key and qty_key:
            return "form", pairs, part_key, qty_key
        return None

    @staticmethod
    def _is_http_part_response_ok(status: int, body_text: str) -> tuple[bool, str]:
        if status >= 400:
            return False, f"HTTP {status}"
        body = (body_text or "").strip()
        if not body:
            return True, ""
        try:
            data = json.loads(body)
            if isinstance(data, list) and len(data) >= 1 and data[0] is False:
                return False, str(data[1]) if len(data) > 1 else "Server JSON [false]"
            if isinstance(data, dict):
                if (
                    data.get("success") is False
                    or data.get("ok") is False
                    or str(data.get("status", "")).lower() == "error"
                    or bool(data.get("error"))
                ):
                    msg = data.get("message") or data.get("error") or data.get("msg") or body[:120]
                    return False, str(msg)
        except Exception:
            pass
        if body.lower().startswith(("eroare", "error:")):
            return False, body[:120]
        return True, ""

    def _reload_jqgrid(
        self,
        page: Page,
        frame: Page | Frame,
        grid: Locator,
        new_base_url: Optional[str] = None,
        edit_url: Optional[str] = None,
        post_data_updates: Optional[dict[str, str]] = None,
        wait_network_idle: bool = True,
    ) -> None:
        """Actualizează opțional url/editurl/postData și declanșează reloadGrid pe un tabel jqGrid."""
        reloaded = False
        try:
            reloaded = bool(
                grid.evaluate(
                    """(gridEl, cfg) => {
                        const tbl = gridEl.querySelector('table.ui-jqgrid-btable');
                        if (!tbl || !window.jQuery) return false;
                        const $t = window.jQuery(tbl);
                        if ($t.jqGrid && (cfg.url || cfg.editurl || cfg.postDataUpdates)) {
                            const params = {};
                            if (cfg.url) params.url = cfg.url;
                            if (cfg.editurl) params.editurl = cfg.editurl;
                            const curPost = $t.jqGrid('getGridParam', 'postData');
                            if (curPost && typeof curPost === 'object') {
                                const nextPost = Object.assign({}, curPost);
                                let changed = false;
                                if (cfg.url) {
                                    nextPost._search = false;
                                    changed = true;
                                }
                                if (cfg.postDataUpdates) {
                                    for (const [k, v] of Object.entries(cfg.postDataUpdates)) {
                                        if (k in nextPost) {
                                            nextPost[k] = v;
                                            changed = true;
                                        }
                                    }
                                }
                                if (changed) params.postData = nextPost;
                            }
                            $t.jqGrid('setGridParam', params);
                        }
                        if ($t.trigger) {
                            $t.trigger('reloadGrid');
                            return true;
                        }
                        return false;
                    }""",
                    {
                        "url": new_base_url or "",
                        "editurl": edit_url or "",
                        "postDataUpdates": post_data_updates or {},
                    },
                )
            )
        except Exception:
            reloaded = False

        if not reloaded:
            refresh_btn = self._first_visible(
                grid,
                ["td[id^='refresh_'], [id^='refresh_'], .ui-pg-button:has(.ui-icon-refresh)"],
            )
            if refresh_btn is not None:
                refresh_btn.click()

        self._wait_frame_ajax(page, frame, wait_network_idle=wait_network_idle)

    def _reload_panel3_grid(self, page: Page, frame: Page | Frame, panel3_grid: Locator) -> None:
        self._reload_jqgrid(page, frame, panel3_grid, wait_network_idle=True)

    def _add_parts_via_http_in_panel3(
        self,
        page: Page,
        frame: Page | Frame,
        panel3_grid: Locator,
        excel_path: str | Path,
        plan_name: str,
    ) -> tuple[int, int]:
        """
        Mod rapid HTTP Request:
          1. Introduce Piesa #1 prin '+' pentru a captura exact URL-ul și payload-ul POST din sesiune.
          2. Trimite piesele #2..#N direct prin HTTP POST folosind cookie-urile sesiunii Playwright.
          3. Reîncarcă Panoul 3 o singură dată la final.
        """
        self._last_p3_capture = None
        parts = read_parts_from_excel(excel_path)
        if not parts:
            raise RuntimeError(f"Fișierul Excel '{Path(excel_path).name}' nu conține piese valide.")

        total = len(parts)
        first_part, first_qty = parts[0]
        self._log(f"[{plan_name}] Adăugare {total} piese prin HTTP Request rapid (Capture #1 + Replay)...")

        parts_added = 0
        parts_rejected = 0
        captured_req = None
        dom_part_name = ""
        dom_qty_name = ""

        # --- Piesa #1 (Capture-1 prin modal '+') ---
        for attempt in range(3):
            modal = self._get_visible_modal(page, frame)
            if modal is None:
                add3 = self._first_visible(panel3_grid, [SELECTORS["jqgrid"]["add_button"]])
                if add3 is None:
                    raise RuntimeError("Nu s-a găsit butonul '+' în Panoul 3.")
                add3.click()
                modal = self._wait_visible_modal(page, frame)
                self._wait_frame_ajax(page, frame)
                time.sleep(0.30)

            p_info = self._fill_and_verify_modal_field(
                modal,
                ["Denumire piesa", "Denumire piesă", "Denumire", "Piesa"],
                first_part,
            )
            q_info = self._fill_and_verify_modal_field(
                modal,
                ["Cant.", "Cant", "Cantitate", "Cantitatea"],
                str(first_qty),
            )
            dom_part_name = p_info.get("name") or p_info.get("id") or dom_part_name
            dom_qty_name = q_info.get("name") or q_info.get("id") or dom_qty_name
            self._clear_modal_form_error(modal)

            save_btn = self._first_visible(modal, SELECTORS["modals"]["save_buttons"])
            if save_btn is None:
                raise RuntimeError("Nu s-a găsit butonul 'Salvează' în modalul de adăugare piesă.")

            clicked = False
            try:
                with page.expect_request(
                    lambda r: r.method == "POST" and "check_if_stop" not in r.url,
                    timeout=15000,
                ) as req_info:
                    save_btn.click()
                    clicked = True
                captured_req = req_info.value
            except Exception as exc:
                if not clicked:
                    raise
                self._log(f"[{plan_name}] Avertisment captură POST Piesa #1: {exc}", level="WARN")

            self._wait_frame_ajax(page, frame)
            time.sleep(0.15)

            err_txt = self._read_modal_form_error(modal)
            if err_txt:
                err_norm = _norm(err_txt)
                if ("obligatoriu" in err_norm or "required" in err_norm) and attempt < 2:
                    self._log(
                        f"  [{plan_name}] [1/{total}] Câmp resetat la '{first_part}'; reîncercare ({attempt + 2}/3)...",
                        level="WARN",
                    )
                    time.sleep(0.25)
                    continue
                parts_rejected += 1
                self._log(
                    f"  [{plan_name}] [1/{total}] Piesa '{first_part}' (cant. {first_qty}) respinsă: {err_txt}",
                    level="WARN",
                )
                break

            parts_added += 1
            self._log(f"  [{plan_name}] [1/{total}] Adăugat (capture): {first_part} x {first_qty}")
            break

        self._close_add_part_modal(page, frame)

        # Extragem șablonul POST din Piesa #1 chiar dacă total == 1 (necesar și pentru Full HTTP)
        template = None
        req_url = ""
        if captured_req is not None:
            req_url = captured_req.url
            req_ct = captured_req.headers.get("content-type", "")
            template = self._build_replay_payload_template(
                post_data=captured_req.post_data or "",
                content_type=req_ct,
                first_part=first_part,
                first_qty=first_qty,
                dom_part_name=dom_part_name,
                dom_qty_name=dom_qty_name,
            )

        if template is not None and req_url:
            self._last_p3_capture = {
                "url": req_url,
                "fmt": template[0],
                "base_payload": template[1],
                "part_key": template[2],
                "qty_key": template[3],
            }

        if total == 1:
            self._reload_panel3_grid(page, frame, panel3_grid)
            return parts_added, parts_rejected

        if template is None or self._context is None or not req_url:
            self._log(
                f"[{plan_name}] Nu s-a putut extrage șablonul POST; se continuă cu modalul '+' pentru restul de {total - 1} piese.",
                level="WARN",
            )
            add_rest, rej_rest = self._add_parts_via_plus_list(
                page=page,
                frame=frame,
                panel3_grid=panel3_grid,
                parts=parts[1:],
                plan_name=plan_name,
                start_index=2,
                total_parts=total,
            )
            return parts_added + add_rest, parts_rejected + rej_rest

        fmt, base_payload, part_key, qty_key = template
        frame_url = frame.url if hasattr(frame, "url") and frame.url else self.base_url
        self._log(f"[{plan_name}] Replay HTTP POST -> {req_url} (câmpuri: {part_key}, {qty_key})")

        for idx, (part_name, qty) in enumerate(parts[1:], start=2):
            body_arg, headers = self._encode_post_payload(
                fmt=fmt,
                base_payload=base_payload,
                updates={
                    part_key: part_name,
                    qty_key: qty if fmt == "json" else str(qty),
                },
                frame_url=frame_url,
            )
            resp = self._context.request.post(req_url, data=body_arg, headers=headers)
            ok, err_detail = self._is_http_part_response_ok(resp.status, resp.text())
            if ok:
                parts_added += 1
                self._log(f"  [{plan_name}] [{idx}/{total}] HTTP OK: {part_name} x {qty}")
            else:
                parts_rejected += 1
                self._log(
                    f"  [{plan_name}] [{idx}/{total}] Piesa '{part_name}' (cant. {qty}) respinsă HTTP: {err_detail}",
                    level="WARN",
                )

        self._reload_panel3_grid(page, frame, panel3_grid)
        return parts_added, parts_rejected

    def _inspect_panel3_rows(self, panel3_grid: Locator) -> list[dict[str, Any]]:
        """Citește rândurile și culorile CSS calculate din Panoul 3."""
        return panel3_grid.evaluate(
            """(gridEl) => {
                const rows = Array.from(gridEl.querySelectorAll('table.ui-jqgrid-btable tr.jqgrow:not(.jqgfirstrow)'));
                return rows.map((tr) => {
                    const trBg = window.getComputedStyle(tr).backgroundColor || '';
                    const tds = Array.from(tr.querySelectorAll('td'));
                    const cellBgs = [];
                    const cellClasses = [];
                    let firstDataText = '';
                    for (const td of tds) {
                        const aria = (td.getAttribute('aria-describedby') || '').toLowerCase();
                        const cls = td.className || '';
                        const style = window.getComputedStyle(td);
                        cellBgs.push(style.backgroundColor || '');
                        cellClasses.push(cls);
                        if (!firstDataText && style.display !== 'none' && !aria.includes('act') && !aria.includes('cb') && !aria.includes('rn')) {
                            firstDataText = (td.textContent || '').trim();
                        }
                    }
                    if (!firstDataText && tds.length > 0) {
                        firstDataText = (tds[0].textContent || '').trim();
                    }
                    return {
                        id: tr.id || '',
                        trBg,
                        rowClasses: tr.className || '',
                        cellBgs,
                        cellClasses,
                        firstCellText: firstDataText,
                        preview: tds.slice(0, 4).map(t => (t.textContent || '').trim()).join(' | ')
                    };
                });
            }"""
        )

    def _clean_panel3_rows(self, page: Page, frame: Page | Frame, panel3_grid: Locator, plan_name: str) -> tuple[int, int]:
        """Scanează rândurile din Panoul 3: șterge ROȘU, păstrează ROZ."""
        red_deleted = 0
        pink_kept = 0
        kept_ids: set[str] = set()
        failed_delete_ids: set[str] = set()

        for _ in range(500):  # Guard anti-buclă infinită
            rows_info = self._inspect_panel3_rows(panel3_grid)

            next_red_row: Optional[dict[str, Any]] = None
            for idx, info in enumerate(rows_info):
                row_key = info["id"] or f"idx_{idx}_{info['preview']}"
                if row_key in kept_ids or row_key in failed_delete_ids:
                    continue

                category, reason = classify_row_color(
                    tr_bg=info["trBg"],
                    cell_bgs=info["cellBgs"],
                    first_cell_text=info["firstCellText"],
                    row_classes=info["rowClasses"],
                    cell_classes=info["cellClasses"],
                )
                if category == "red":
                    next_red_row = {"row_key": row_key, "id": info["id"], "idx": idx, "reason": reason, "preview": info["preview"]}
                    break
                elif category == "pink":
                    kept_ids.add(row_key)
                    pink_kept += 1
                    self._log(f"  [{plan_name}] Rând ROZ păstrat ({reason}): {info['preview']}")
                else:
                    kept_ids.add(row_key)

            if next_red_row is None:
                break

            row_id = next_red_row["id"]
            self._log(f"  [{plan_name}] Ștergere rând ROȘU ({next_red_row['reason']}): {next_red_row['preview']}")
            if row_id:
                safe_id = row_id.replace('"', '\\"')
                row_loc = panel3_grid.locator(f'tr[id="{safe_id}"]')
            else:
                row_loc = panel3_grid.locator(SELECTORS["jqgrid"]["rows"]).nth(next_red_row["idx"])

            del_btn = self._first_visible(row_loc, [SELECTORS["jqgrid"]["inline_delete"]])
            if del_btn is not None:
                del_btn.click()
            else:
                # Fallback: selectează rândul și apasă coșul de gunoi din pager
                row_loc.click()
                pager_del = panel3_grid.locator("td[id^='del_'], .ui-pg-button:has(.ui-icon-trash)").first
                if pager_del.count() == 0 or not pager_del.is_visible():
                    failed_delete_ids.add(next_red_row["row_key"])
                    self._log(f"  [{plan_name}] Nu s-a găsit butonul de ștergere pentru rândul {row_id}.", level="WARN")
                    continue
                pager_del.click()

            del_modal = self._wait_visible_modal(page, frame)
            confirm_btn = self._first_visible(del_modal, SELECTORS["modals"]["delete_confirm_buttons"])
            if confirm_btn is None:
                raise RuntimeError("Nu s-a găsit butonul de confirmare 'Șterge' în modalul de ștergere.")
            confirm_btn.click()
            self._wait_modal_closed(page, frame, del_modal)
            red_deleted += 1

        return red_deleted, pink_kept

    def _resolve_comanda_from_blueprint(
        self, bp: _FullHttpBlueprint, project_query: str
    ) -> Optional[tuple[str, str]]:
        """Rezolvă (comanda_id, matched_comanda) din cache sau din lista opțiunilor Select2 capturată."""
        q_norm = _norm(project_query)
        if not q_norm:
            return None
        if q_norm in bp.comanda_cache:
            return bp.comanda_cache[q_norm]
        for val, txt in bp.comanda_options:
            if val and q_norm in _norm(txt):
                bp.comanda_cache[q_norm] = (val, txt)
                return val, txt
        return None

    @staticmethod
    def _try_float(val: str) -> Optional[float]:
        s = (val or "").strip().replace(",", ".")
        if not s or not re.match(r"^-?\d+(?:\.\d+)?$", s):
            return None
        try:
            return float(s)
        except ValueError:
            return None

    @classmethod
    def _value_matches(cls, actual: str, expected: str) -> bool:
        a = (actual or "").strip()
        e = (expected or "").strip()
        if not a or not e:
            return False
        if a == e or _norm(a) == _norm(e):
            return True
        fa = cls._try_float(a)
        fe = cls._try_float(e)
        if fa is not None and fe is not None and abs(fa - fe) <= 1e-6:
            return True
        return False

    @classmethod
    def _find_key_in_pairs_or_dom(
        cls,
        pairs: list[tuple[str, str]],
        dom_name: str,
        value: str,
        hints: tuple[str, ...],
        exclude_keys: Optional[set[str]] = None,
        alt_values: Optional[list[str]] = None,
    ) -> str:
        """Găsește numele exact al unui câmp din POST după atributul DOM, valoare (inclusiv numeric) sau hint."""
        excl = {e for e in (exclude_keys or set()) if e}
        keys = [k for k, _ in pairs if k not in excl]
        if dom_name and dom_name in keys:
            return dom_name

        candidates = [c for c in ([value] + list(alt_values or [])) if (c or "").strip()]
        for cand in candidates:
            exact = [k for k, v in pairs if k not in excl and cls._value_matches(str(v), cand)]
            if exact:
                for k in exact:
                    if any(h in k.lower() for h in hints):
                        return k
                return exact[0]

        for k in keys:
            if any(h in k.lower() for h in hints):
                return k
        return ""

    def _select_best_add_post(
        self,
        events: list[_CapturedRequest],
        expected_values: list[str],
        dom_names: list[str],
    ) -> Optional[_CapturedRequest]:
        """
        Alege POST-ul real de adăugare (`oper=add` / câmpuri modal) dintr-un segment de trafic,
        ignorând eventualele POST-uri de reloadGrid (`_search`/`rows`/`page`) sau hook-uri secundare.
        """
        exp_clean = [v.strip() for v in expected_values if (v or "").strip()]
        dom_clean = [d.strip() for d in dom_names if (d or "").strip()]

        best_event: Optional[_CapturedRequest] = None
        best_rank: Optional[tuple[int, int]] = None

        for idx, ev in enumerate(events):
            if ev.method != "POST":
                continue
            parsed = self._parse_post_payload(ev.post_data, ev.content_type)
            if parsed is None:
                continue
            fmt, payload = parsed
            pairs = self._payload_to_pairs(fmt, payload)
            keys_set = {k for k, _ in pairs}
            keys_low = {k.lower() for k in keys_set}
            kv_low = {k.lower(): str(v).strip() for k, v in pairs}

            oper_val = kv_low.get("oper", "").lower()
            if oper_val in {"del", "search"}:
                continue
            is_grid_fetch = (
                oper_val != "add"
                and {"_search", "rows", "page"}.issubset(keys_low)
            )
            if is_grid_fetch:
                continue

            score = 0
            if oper_val == "add":
                score += 100
            for d in dom_clean:
                if d in keys_set:
                    score += 35
            for ev_val in exp_clean:
                ev_norm = _norm(ev_val)
                if any(self._value_matches(str(v), ev_val) for _, v in pairs):
                    score += 25
                elif ev_norm and any(ev_norm in _norm(str(v)) for _, v in pairs if str(v).strip()):
                    score += 10

            if score == 0 and oper_val != "add" and len(events) > 1:
                # Păstrează totuși un fallback neutru dacă niciun alt POST nu are scor mai mare
                score = 1

            rank = (score, -idx)
            if best_rank is None or rank > best_rank:
                best_rank = rank
                best_event = ev

        return best_event

    def _extract_grid_fetch_events(
        self, events: list[_CapturedRequest]
    ) -> list[tuple[_CapturedRequest, list[Any]]]:
        """Extrage răspunsurile JSON jqGrid atât din request-uri GET, cât și din mtype='POST'."""
        fetches: list[tuple[_CapturedRequest, list[Any]]] = []
        for ev in events:
            if ev.method == "POST":
                parsed = self._parse_post_payload(ev.post_data, ev.content_type)
                if parsed is not None:
                    pairs = self._payload_to_pairs(*parsed)
                    kv_low = {k.lower(): str(v).strip().lower() for k, v in pairs}
                    if kv_low.get("oper") in {"add", "edit", "del"}:
                        continue
            elif ev.method != "GET":
                continue

            rows = self._parse_jqgrid_json_rows(ev.response_text)
            if rows is not None:
                fetches.append((ev, rows))
        return fetches

    def _build_blueprint_from_calibration(
        self,
        *,
        cutting_type: str,
        plan: CuttingPlanRecord,
        sample_plan_id: str,
        sample_rowid2: str,
        sample_comanda_id: str,
        matched_comanda: str,
        comanda_options: list[tuple[str, str]],
        dom_p1_name: str,
        dom_comanda_name: str,
        dom_weight_name: str,
        p1_events: list[_CapturedRequest],
        p2_priming_events: list[_CapturedRequest],
        p2_after_events: list[_CapturedRequest],
        p3_initial_events: list[_CapturedRequest],
        p3_after_events: list[_CapturedRequest],
        p3_capture: Optional[dict[str, Any]],
    ) -> Optional[_FullHttpBlueprint]:
        """Construiește șablonul HTTP complet pentru Panou 1 + Panou 2 + Panou 3 din traficul Planului #1."""
        def _fail(reason: str) -> None:
            self._log(f"[{plan.plan_name}] Calibrare Full HTTP oprită: {reason}", level="WARN")

        def _summarize_posts(ev_list: list[_CapturedRequest]) -> list[tuple[str, list[str]]]:
            out: list[tuple[str, list[str]]] = []
            for ev in ev_list:
                if ev.method != "POST":
                    continue
                parsed = self._parse_post_payload(ev.post_data, ev.content_type)
                keys = [k for k, _ in self._payload_to_pairs(*parsed)] if parsed is not None else []
                out.append((ev.url, keys))
            return out

        weight_str = str(plan.part_weight_int) if plan.part_weight_int is not None else ""
        p2_post_pre = self._select_best_add_post(
            p2_after_events,
            expected_values=[sample_comanda_id, matched_comanda, weight_str],
            dom_names=[dom_comanda_name, dom_weight_name],
        )

        if not sample_comanda_id and p3_capture:
            for k, v in self._split_query_raw_pairs(str(p3_capture.get("url") or "")):
                if v.strip() and any(tok in k.lower() for tok in ("comanda", "proiect")):
                    sample_comanda_id = v.strip()
                    break

        if not sample_comanda_id and p2_post_pre is not None:
            p2_parsed_fb = self._parse_post_payload(p2_post_pre.post_data, p2_post_pre.content_type)
            if p2_parsed_fb is not None:
                for k, v in self._payload_to_pairs(*p2_parsed_fb):
                    if not str(v).strip() or k.lower() in {"id", "oper"}:
                        continue
                    if (dom_comanda_name and k == dom_comanda_name) or any(
                        tok in k.lower() for tok in ("comanda", "proiect")
                    ):
                        sample_comanda_id = str(v).strip()
                        break

        if not sample_plan_id or not sample_rowid2 or not sample_comanda_id or not p3_capture:
            _fail(
                f"date lipsă (plan_id='{sample_plan_id}', rowid2='{sample_rowid2}', "
                f"comanda_id='{sample_comanda_id}', p3_capture={bool(p3_capture)})"
            )
            return None

        # --- 1. Panou 1 POST & GET/POST-fetch ---
        p1_post = self._select_best_add_post(
            p1_events,
            expected_values=[plan.plan_name],
            dom_names=[dom_p1_name],
        )
        if p1_post is None:
            _fail(f"nu s-a capturat POST valid pentru Panoul 1 (posts={_summarize_posts(p1_events)})")
            return None
        p1_parsed = self._parse_post_payload(p1_post.post_data, p1_post.content_type)
        if p1_parsed is None:
            _fail(f"payload POST Panoul 1 necunoscut ({p1_post.url})")
            return None
        p1_fmt, p1_payload = p1_parsed
        p1_pairs = self._payload_to_pairs(p1_fmt, p1_payload)
        p1_plan_key = self._find_key_in_pairs_or_dom(
            p1_pairs,
            dom_p1_name,
            plan.plan_name,
            ("plan", "croi", "nume", "denum"),
            exclude_keys={"id", "oper"},
        )
        if not p1_plan_key:
            _fail(
                f"nu s-a identificat câmpul 'Plan croi' în POST Panoul 1 "
                f"(url='{p1_post.url}', keys={[k for k, _ in p1_pairs]})"
            )
            return None
        p1_post_id_hint = self._detect_post_id_hint(p1_post.response_text, sample_plan_id)

        p1_json_gets = self._extract_grid_fetch_events(p1_events)
        p1_get: Optional[_CapturedRequest] = None
        p1_rows: list[Any] = []
        for e, rows in reversed(p1_json_gets):
            rid = self._find_row_id_in_jqgrid_json(rows, plan.plan_name)
            if rid == sample_plan_id or any(
                sample_plan_id == self._jqgrid_row_values(r)[0] for r in rows
            ):
                p1_get = e
                p1_rows = rows
                break
        if p1_get is None and p1_json_gets:
            p1_get, p1_rows = p1_json_gets[-1]
        if p1_get is None and p1_post_id_hint is None:
            _fail("nu s-a capturat nici fetch JSON, nici ID în POST pentru Panoul 1")
            return None

        p1_id_key, p1_match_key = (
            self._detect_jqgrid_keys(p1_rows, sample_plan_id, plan.plan_name)
            if p1_rows
            else ("id", "")
        )

        # --- 2. Panou 2 POST & GET/POST-fetch ---
        p2_post = self._select_best_add_post(
            p2_after_events,
            expected_values=[sample_comanda_id, matched_comanda, weight_str],
            dom_names=[dom_comanda_name, dom_weight_name],
        )
        if p2_post is None:
            _fail(f"nu s-a capturat POST valid pentru Panoul 2 (posts={_summarize_posts(p2_after_events)})")
            return None
        p2_parsed = self._parse_post_payload(p2_post.post_data, p2_post.content_type)
        if p2_parsed is None:
            _fail(f"payload POST Panoul 2 necunoscut ({p2_post.url})")
            return None
        p2_fmt, p2_payload = p2_parsed
        p2_pairs = self._payload_to_pairs(p2_fmt, p2_payload)

        p2_comanda_key = self._find_key_in_pairs_or_dom(
            p2_pairs,
            dom_comanda_name,
            sample_comanda_id,
            ("comanda", "proiect", "order"),
            exclude_keys={"id", "oper"},
            alt_values=[matched_comanda] if matched_comanda else None,
        )
        p2_weight_key = self._find_key_in_pairs_or_dom(
            p2_pairs,
            dom_weight_name,
            weight_str,
            ("greut", "weight"),
            exclude_keys={"id", "oper", p2_comanda_key},
        )
        if not p2_comanda_key or not p2_weight_key:
            _fail(
                f"nu s-au identificat câmpurile Panou 2 (comanda='{p2_comanda_key}', greutate='{p2_weight_key}', "
                f"dom_cmd='{dom_comanda_name}', dom_wt='{dom_weight_name}', "
                f"url='{p2_post.url}', keys={[k for k, _ in p2_pairs]}, "
                f"posts={_summarize_posts(p2_after_events)})"
            )
            return None

        plan_url_hints = {"plan": ("plan", "rowid", "parent", "master", "id")}
        p2_post_url_plan_params = self._classify_id_params(
            self._split_query_raw_pairs(p2_post.url),
            targets={"plan": sample_plan_id},
            role_hints=plan_url_hints,
        )["plan"]
        p2_post_body_plan_keys = self._classify_id_params(
            p2_pairs,
            targets={"plan": sample_plan_id},
            role_hints=plan_url_hints,
            extra_skip_keys={"id", "oper", p2_comanda_key, p2_weight_key},
        )["plan"]
        p2_post_id_hint = self._detect_post_id_hint(p2_post.response_text, sample_rowid2)

        p2_after_gets = self._extract_grid_fetch_events(p2_after_events)
        p2_priming_gets = self._extract_grid_fetch_events(p2_priming_events)

        def _fetch_has_plan_ref(ev: _CapturedRequest) -> bool:
            if self._classify_id_params(self._split_query_raw_pairs(ev.url), {"plan": sample_plan_id}, plan_url_hints)["plan"]:
                return True
            if ev.method == "POST":
                ep = self._parse_post_payload(ev.post_data, ev.content_type)
                if ep is not None:
                    return bool(self._classify_id_params(self._payload_to_pairs(*ep), {"plan": sample_plan_id}, plan_url_hints)["plan"])
            return False

        p2_after_pick: Optional[tuple[_CapturedRequest, list[Any]]] = next(
            ((e, rows) for e, rows in reversed(p2_after_gets) if any(sample_rowid2 == self._jqgrid_row_values(r)[0] for r in rows)),
            p2_after_gets[-1] if p2_after_gets else None,
        )
        p2_priming_pick: Optional[tuple[_CapturedRequest, list[Any]]] = next(
            ((e, rows) for e, rows in p2_priming_gets if _fetch_has_plan_ref(e)),
            p2_priming_gets[0] if p2_priming_gets else None,
        )

        p2_get_req: Optional[_CapturedRequest] = None
        p2_id_key = "id"
        if p2_after_pick is not None:
            p2_get_req = p2_after_pick[0]
            p2_id_key, _ = self._detect_jqgrid_keys(p2_after_pick[1], sample_rowid2, matched_comanda)
            if p2_priming_pick is not None and not _fetch_has_plan_ref(p2_get_req):
                p2_get_req = p2_priming_pick[0]
        elif p2_priming_pick is not None:
            p2_get_req = p2_priming_pick[0]

        if p2_get_req is None and p2_post_id_hint is None:
            _fail("nu s-a capturat nici fetch JSON, nici ID în POST pentru Panoul 2")
            return None

        p2_get_url = p2_get_req.url if p2_get_req is not None else ""
        p2_get_plan_params = (
            self._classify_id_params(self._split_query_raw_pairs(p2_get_url), {"plan": sample_plan_id}, plan_url_hints)["plan"]
            if p2_get_url
            else []
        )
        p2_get_body_plan_keys: list[str] = []
        if p2_get_req is not None and p2_get_req.method == "POST":
            p2_get_parsed = self._parse_post_payload(p2_get_req.post_data, p2_get_req.content_type)
            if p2_get_parsed is not None:
                p2_get_body_plan_keys = self._classify_id_params(
                    self._payload_to_pairs(*p2_get_parsed), {"plan": sample_plan_id}, plan_url_hints
                )["plan"]

        # --- 3. Panou 3 POST & GET/POST-fetch ---
        p3_targets = {"rowid": sample_rowid2, "comanda": sample_comanda_id, "plan": sample_plan_id}
        p3_role_hints = {
            "rowid": ("rowid", "id_leg", "parent", "master", "id"),
            "comanda": ("comanda", "proiect", "order"),
            "plan": ("id_plan", "plan"),
        }

        p3_after_gets = [e for e, _ in self._extract_grid_fetch_events(p3_after_events)]
        p3_init_gets = [e for e, _ in self._extract_grid_fetch_events(p3_initial_events)]
        p3_get = next(
            (e for e in reversed(p3_after_gets) if sample_rowid2 in e.url or "list3" in e.url.lower()),
            None,
        ) or (p3_after_gets[-1] if p3_after_gets else None)
        if p3_get is None:
            p3_get = next(
                (e for e in p3_init_gets if sample_rowid2 in e.url or "list3" in e.url.lower()),
                None,
            ) or (p3_init_gets[-1] if p3_init_gets else None)
        if p3_get is None:
            _fail("nu s-a capturat fetch JSON pentru Panoul 3")
            return None

        p3_get_roles = self._classify_id_params(self._split_query_raw_pairs(p3_get.url), p3_targets, p3_role_hints)
        p3_get_body_roles: dict[str, list[str]] = {"rowid": [], "comanda": [], "plan": []}
        if p3_get.method == "POST":
            p3_get_parsed = self._parse_post_payload(p3_get.post_data, p3_get.content_type)
            if p3_get_parsed is not None:
                p3_get_body_roles = self._classify_id_params(
                    self._payload_to_pairs(*p3_get_parsed), p3_targets, p3_role_hints
                )

        p3_post_url = str(p3_capture["url"])
        p3_post_url_roles = self._classify_id_params(self._split_query_raw_pairs(p3_post_url), p3_targets, p3_role_hints)
        p3_post_pairs = self._payload_to_pairs(str(p3_capture["fmt"]), p3_capture["base_payload"])
        p3_post_body_roles = self._classify_id_params(
            p3_post_pairs,
            p3_targets,
            p3_role_hints,
            extra_skip_keys={"id", "oper", str(p3_capture["part_key"]), str(p3_capture["qty_key"])},
        )

        # Verifică să existe cel puțin o legătură dinamică pentru Panou 2 -> Panou 3
        has_p2_link = bool(
            p2_get_plan_params
            or p2_get_body_plan_keys
            or p2_post_url_plan_params
            or p2_post_body_plan_keys
        )
        has_p3_link = bool(
            p3_get_roles["rowid"]
            or p3_get_body_roles["rowid"]
            or p3_post_url_roles["rowid"]
            or p3_post_body_roles["rowid"]
        )
        if not has_p2_link or not has_p3_link:
            _fail(f"legătură dinamică incompletă (has_p2_link={has_p2_link}, has_p3_link={has_p3_link})")
            return None

        comanda_cache: dict[str, tuple[str, str]] = {}
        if _norm(plan.project_query):
            comanda_cache[_norm(plan.project_query)] = (sample_comanda_id, matched_comanda)

        return _FullHttpBlueprint(
            cutting_type=cutting_type,
            p1_post_url=p1_post.url,
            p1_post_fmt=p1_fmt,
            p1_post_payload=p1_payload,
            p1_plan_key=p1_plan_key,
            p1_get_url=p1_get.url if p1_get is not None else "",
            p1_id_key=p1_id_key,
            p1_match_key=p1_match_key,
            p1_post_id_hint=p1_post_id_hint,
            p2_get_url=p2_get_url,
            p2_get_plan_params=p2_get_plan_params,
            p2_post_url=p2_post.url,
            p2_post_url_plan_params=p2_post_url_plan_params,
            p2_post_fmt=p2_fmt,
            p2_post_payload=p2_payload,
            p2_post_body_plan_keys=p2_post_body_plan_keys,
            p2_comanda_key=p2_comanda_key,
            p2_weight_key=p2_weight_key,
            p2_id_key=p2_id_key,
            p2_post_id_hint=p2_post_id_hint,
            p3_get_url=p3_get.url,
            p3_get_rowid_params=p3_get_roles["rowid"],
            p3_get_comanda_params=p3_get_roles["comanda"],
            p3_get_plan_params=p3_get_roles["plan"],
            p3_post_url=p3_post_url,
            p3_post_url_rowid_params=p3_post_url_roles["rowid"],
            p3_post_url_comanda_params=p3_post_url_roles["comanda"],
            p3_post_url_plan_params=p3_post_url_roles["plan"],
            p3_post_fmt=str(p3_capture["fmt"]),
            p3_post_payload=p3_capture["base_payload"],
            p3_post_body_rowid_keys=p3_post_body_roles["rowid"],
            p3_post_body_comanda_keys=p3_post_body_roles["comanda"],
            p3_post_body_plan_keys=p3_post_body_roles["plan"],
            p3_part_key=str(p3_capture["part_key"]),
            p3_qty_key=str(p3_capture["qty_key"]),
            sample_plan_id=sample_plan_id,
            sample_rowid2=sample_rowid2,
            sample_comanda_id=sample_comanda_id,
            comanda_options=comanda_options,
            comanda_cache=comanda_cache,
            p1_get_method=p1_get.method if p1_get is not None else "GET",
            p1_get_post_data=p1_get.post_data if p1_get is not None else "",
            p1_get_content_type=p1_get.content_type if p1_get is not None else "",
            p2_get_method=p2_get_req.method if p2_get_req is not None else "GET",
            p2_get_post_data=p2_get_req.post_data if p2_get_req is not None else "",
            p2_get_content_type=p2_get_req.content_type if p2_get_req is not None else "",
            p2_get_body_plan_keys=p2_get_body_plan_keys,
            p3_get_method=p3_get.method,
            p3_get_post_data=p3_get.post_data,
            p3_get_content_type=p3_get.content_type,
            p3_get_body_rowid_keys=p3_get_body_roles["rowid"],
            p3_get_body_comanda_keys=p3_get_body_roles["comanda"],
            p3_get_body_plan_keys=p3_get_body_roles["plan"],
        )

    @classmethod
    def _split_query_raw_pairs(cls, url: str) -> list[tuple[str, str]]:
        return [(k, v) for k, v, _ in cls._split_query_raw(urlsplit(url).query)]

    def _finalize_stage1_result(
        self,
        page: Page,
        panel3: Locator,
        plan: CuttingPlanRecord,
        matched_comanda: str,
        parts_added: int,
        parts_rejected: int,
        red_deleted: int,
        pink_kept: int,
        mode_tag: str = "Etapa 1",
    ) -> Stage1PlanResult:
        """Verifică dacă au rămas piese valide în Panoul 3 și construiește rezultatul planului."""
        final_rows = 0
        try:
            final_rows = panel3.locator(SELECTORS["jqgrid"]["rows"]).count()
        except Exception:
            final_rows = 0

        if final_rows == 0:
            empty_msg = "Plan gol: 0 piese rămase în Panoul 3 după încărcare/curățare."
            self._log(f"[{plan.plan_name}] {empty_msg}", level="ERROR")
            self._save_error_screenshot(page, plan.plan_name, "stage1_empty")
            return Stage1PlanResult(
                plan_name=plan.plan_name,
                success=False,
                matched_comanda=matched_comanda,
                parts_added=parts_added,
                parts_rejected=parts_rejected,
                red_deleted=red_deleted,
                pink_kept=pink_kept,
                error=empty_msg,
            )

        self._log(
            f"[{plan.plan_name}] Finalizat {mode_tag} — Piese: {parts_added} (respinse: {parts_rejected}), "
            f"Rânduri finale Panou 3: {final_rows}, ROȘII șterse: {red_deleted}, ROZ păstrate: {pink_kept}."
        )
        return Stage1PlanResult(
            plan_name=plan.plan_name,
            success=True,
            matched_comanda=matched_comanda,
            parts_added=parts_added,
            parts_rejected=parts_rejected,
            red_deleted=red_deleted,
            pink_kept=pink_kept,
        )

    def _run_single_plan_ui(
        self,
        page: Page,
        plan: CuttingPlanRecord,
        effective_panel3_mode: str,
        calibrate: bool = False,
        cache_blueprint: Optional[_FullHttpBlueprint] = None,
    ) -> tuple[Stage1PlanResult, Optional[_FullHttpBlueprint]]:
        """Execută un plan prin UI (opțional capturând și șablonul pentru Full HTTP)."""
        recorder = _HttpTrafficRecorder(page) if calibrate else None
        blueprint: Optional[_FullHttpBlueprint] = None
        try:
            frame = self._get_work_frame(page)
            self._dismiss_any_open_modals(page, frame)

            # --- PANOU 1: Productie-Plan de croi ---
            self._select_debitare_dropdown(page, frame, plan.cutting_type)
            panel1 = self._find_panel_grid(frame, "panel1_titles", fallback_index=0)

            add1 = self._first_visible(panel1, [SELECTORS["jqgrid"]["add_button"]])
            if add1 is None:
                raise RuntimeError("Nu s-a găsit butonul '+' în Panoul 1.")
            add1.click()

            modal1 = self._wait_visible_modal(page, frame)
            time.sleep(0.20)
            self._select_modal_option_if_present(modal1, ["Debitare"], plan.cutting_type)
            p1_info = self._fill_and_verify_modal_field(
                modal1, ["Plan croi", "Plan de croi", "Plan"], plan.plan_name
            )
            dom_p1_name = p1_info.get("name") or p1_info.get("id") or ""

            m_p1_save = recorder.mark() if recorder is not None else 0
            self._click_save_in_modal(page, frame, modal1)
            self._log(f"[{plan.plan_name}] Plan adăugat în Panoul 1.")

            row1 = self._find_and_select_grid_row(
                page,
                frame,
                panel1,
                plan.plan_name,
                col_hints=["Plan croi", "Plan"],
                click_to_select=False,
            )
            sample_plan_id = (row1.get_attribute("id") or "").strip() if recorder is not None else ""
            p1_events = recorder.since(m_p1_save) if recorder is not None else []

            m_row1_click = recorder.mark() if recorder is not None else 0
            self._click_grid_row_to_select(page, frame, row1)

            # --- PANOU 2: Comenzi / planul de croi ---
            panel2 = self._find_panel_grid(frame, "panel2_titles", fallback_index=1)
            self._clear_grid_search_filters(page, frame, panel2)
            p2_priming_events = recorder.since(m_row1_click) if recorder is not None else []

            add2 = self._first_visible(panel2, [SELECTORS["jqgrid"]["add_button"]])
            if add2 is None:
                raise RuntimeError("Nu s-a găsit butonul '+' în Panoul 2.")
            add2.click()

            modal2 = self._wait_visible_modal(page, frame)
            matched_comanda = self._select2_choose(
                page,
                frame,
                modal2,
                label_hints=["Comanda", "Proiect"],
                query=plan.project_query,
            )
            self._log(f"[{plan.plan_name}] Comandă selectată: '{matched_comanda}'")
            dom_comanda_name, sample_comanda_id, comanda_options = self._read_modal_select_info(
                modal2, ["Comanda", "Proiect"]
            )
            if cache_blueprint is not None and sample_comanda_id and _norm(plan.project_query):
                cache_blueprint.comanda_cache[_norm(plan.project_query)] = (
                    sample_comanda_id,
                    matched_comanda,
                )

            w_info = self._fill_and_verify_modal_field(
                modal2, ["Greutatea", "Greutate"], str(plan.part_weight_int)
            )
            dom_weight_name = w_info.get("name") or w_info.get("id") or ""

            m_p2_save = recorder.mark() if recorder is not None else 0
            self._click_save_in_modal(page, frame, modal2)

            row2 = self._find_and_select_grid_row(
                page,
                frame,
                panel2,
                matched_comanda or plan.project_query,
                col_hints=["Comanda", "Proiect"],
                click_to_select=False,
            )
            sample_rowid2 = (row2.get_attribute("id") or "").strip() if recorder is not None else ""
            p2_after_events = recorder.since(m_p2_save) if recorder is not None else []

            m_row2_click = recorder.mark() if recorder is not None else 0
            self._click_grid_row_to_select(page, frame, row2)

            # --- PANOU 3: Piese din comanda ---
            panel3 = self._find_panel_grid(frame, "panel3_titles", fallback_index=2)
            self._wait_frame_ajax(page, frame)
            time.sleep(0.20)
            p3_initial_events = recorder.since(m_row2_click) if recorder is not None else []

            m_p3_ops = recorder.mark() if recorder is not None else 0
            parts_added = 0
            parts_rejected = 0

            if effective_panel3_mode == "http_request":
                parts_added, parts_rejected = self._add_parts_via_http_in_panel3(
                    page=page,
                    frame=frame,
                    panel3_grid=panel3,
                    excel_path=plan.excel_path or "",
                    plan_name=plan.plan_name,
                )
            elif effective_panel3_mode == "import_excel":
                import_btn = self._find_panel3_import_button(frame, panel3)
                import_btn.click()

                import_modal = self._wait_visible_modal(page, frame)
                self._upload_excel_in_import_modal(
                    page=page,
                    frame=frame,
                    import_modal=import_modal,
                    excel_path=plan.excel_path or "",
                    plan_name=plan.plan_name,
                )
            else:
                parts_added, parts_rejected = self._add_parts_via_plus_in_panel3(
                    page=page,
                    frame=frame,
                    panel3_grid=panel3,
                    excel_path=plan.excel_path or "",
                    plan_name=plan.plan_name,
                )

            p3_after_events = recorder.since(m_p3_ops) if recorder is not None else []
            if recorder is not None:
                recorder.stop()
                blueprint = self._build_blueprint_from_calibration(
                    cutting_type=plan.cutting_type,
                    plan=plan,
                    sample_plan_id=sample_plan_id,
                    sample_rowid2=sample_rowid2,
                    sample_comanda_id=sample_comanda_id,
                    matched_comanda=matched_comanda,
                    comanda_options=comanda_options,
                    dom_p1_name=dom_p1_name,
                    dom_comanda_name=dom_comanda_name,
                    dom_weight_name=dom_weight_name,
                    p1_events=p1_events,
                    p2_priming_events=p2_priming_events,
                    p2_after_events=p2_after_events,
                    p3_initial_events=p3_initial_events,
                    p3_after_events=p3_after_events,
                    p3_capture=self._last_p3_capture,
                )
                if blueprint is not None:
                    self._log(
                        f"[{plan.plan_name}] Calibrare Full HTTP reușită "
                        f"(Debitare='{plan.cutting_type}', Comenzi cunoscute={len(blueprint.comanda_options)}, "
                        f"P2 planParams={blueprint.p2_get_plan_params or blueprint.p2_post_url_plan_params or blueprint.p2_post_body_plan_keys}, "
                        f"P3 rowidParams={blueprint.p3_post_url_rowid_params or blueprint.p3_get_rowid_params})."
                    )
                else:
                    self._log(
                        f"[{plan.plan_name}] Calibrarea Full HTTP incompletă; se va folosi modul HTTP standard (Panou 3).",
                        level="WARN",
                    )

            red_deleted, pink_kept = self._clean_panel3_rows(page, frame, panel3, plan.plan_name)
            res = self._finalize_stage1_result(
                page=page,
                panel3=panel3,
                plan=plan,
                matched_comanda=matched_comanda,
                parts_added=parts_added,
                parts_rejected=parts_rejected,
                red_deleted=red_deleted,
                pink_kept=pink_kept,
                mode_tag="Etapa 1",
            )
            return res, blueprint
        finally:
            if recorder is not None:
                recorder.stop()

    def _send_grid_fetch(
        self,
        *,
        url: str,
        method: str,
        post_data: str,
        content_type: str,
        url_updates: dict[str, str],
        body_updates: dict[str, Any],
        paging_updates: Optional[dict[str, str]],
        frame_url: str,
        refresh_nd: bool = True,
        add_missing: bool = False,
    ) -> Any:
        """Execută un fetch jqGrid fie ca GET, fie ca POST (mtype='POST')."""
        assert self._context is not None
        paging = paging_updates or {}
        if method.upper() != "POST":
            merged_url = dict(url_updates)
            merged_url.update(paging)
            target_url = self._replace_url_params(
                url,
                merged_url,
                refresh_nd=refresh_nd,
                add_missing=add_missing,
            )
            return self._context.request.get(
                target_url,
                headers={"X-Requested-With": "XMLHttpRequest", "Referer": frame_url},
            )

        target_url = (
            self._replace_url_params(url, url_updates, refresh_nd=refresh_nd)
            if (url_updates or refresh_nd)
            else url
        )
        parsed = self._parse_post_payload(post_data, content_type)
        fmt, base_payload = parsed if parsed is not None else ("form", [])
        merged_body: dict[str, Any] = dict(body_updates)
        merged_body.update(paging)
        if refresh_nd:
            base_keys = {k.lower() for k, _ in self._payload_to_pairs(fmt, base_payload)}
            if "nd" in base_keys and "nd" not in merged_body:
                merged_body["nd"] = str(int(time.time() * 1000))
        body_arg, headers = self._encode_post_payload(
            fmt=fmt,
            base_payload=base_payload,
            updates=merged_body,
            frame_url=frame_url,
        )
        return self._context.request.post(target_url, data=body_arg, headers=headers)

    def _lookup_p1_id_via_get(
        self, bp: _FullHttpBlueprint, plan_name: str, frame_url: str
    ) -> str:
        """Caută row_id-ul planului nou creat prin fetch JSON (GET sau POST) pe Panoul 1."""
        if self._context is None or not bp.p1_get_url:
            return ""

        def _fetch_p1(paging: dict[str, str]) -> Any:
            return self._send_grid_fetch(
                url=bp.p1_get_url,
                method=bp.p1_get_method,
                post_data=bp.p1_get_post_data,
                content_type=bp.p1_get_content_type,
                url_updates={},
                body_updates={},
                paging_updates=paging,
                frame_url=frame_url,
                refresh_nd=True,
                add_missing=True,
            )

        # Încercarea 1: pagina 1 cu rows=100 (acoperă rapid sortările descendente sau loturile mici)
        resp1 = _fetch_p1({"rows": "100", "page": "1", "_search": "false"})
        rows1 = self._parse_jqgrid_json_rows(resp1.text())
        if rows1:
            hit = self._find_row_id_in_jqgrid_json(
                rows1, plan_name, id_key=bp.p1_id_key, match_key=bp.p1_match_key
            )
            if hit:
                return hit

        # Încercarea 2: filtrare toolbar jqGrid după coloana planului
        if bp.p1_match_key:
            filters_json = json.dumps(
                {
                    "groupOp": "AND",
                    "rules": [{"field": bp.p1_match_key, "op": "cn", "data": plan_name}],
                }
            )
            resp2 = _fetch_p1(
                {
                    "_search": "true",
                    "page": "1",
                    "rows": "50",
                    bp.p1_match_key: plan_name,
                    "filters": filters_json,
                }
            )
            rows2 = self._parse_jqgrid_json_rows(resp2.text())
            if rows2:
                hit = self._find_row_id_in_jqgrid_json(
                    rows2, plan_name, id_key=bp.p1_id_key, match_key=bp.p1_match_key
                )
                if hit:
                    return hit

        # Încercarea 3: sortare descendentă pe pagina 1
        resp3 = _fetch_p1({"rows": "100", "page": "1", "_search": "false", "sord": "desc"})
        rows3 = self._parse_jqgrid_json_rows(resp3.text())
        if rows3:
            hit = self._find_row_id_in_jqgrid_json(
                rows3, plan_name, id_key=bp.p1_id_key, match_key=bp.p1_match_key
            )
            if hit:
                return hit

        # Încercarea 4: ultima pagină dacă serverul sortează crescător și ignoră sord=desc
        try:
            meta1 = json.loads(resp1.text() or "{}")
            tot_pages = int(meta1.get("total") or 1) if isinstance(meta1, dict) else 1
        except Exception:
            tot_pages = 1
        if tot_pages > 1:
            resp4 = _fetch_p1({"rows": "100", "page": str(tot_pages), "_search": "false"})
            rows4 = self._parse_jqgrid_json_rows(resp4.text())
            if rows4:
                return self._find_row_id_in_jqgrid_json(
                    rows4, plan_name, id_key=bp.p1_id_key, match_key=bp.p1_match_key
                )
        return ""

    def _run_plan_full_http(
        self,
        page: Page,
        frame: Page | Frame,
        bp: _FullHttpBlueprint,
        plan: CuttingPlanRecord,
    ) -> Optional[Stage1PlanResult]:
        """
        Execută complet un plan (#2..#N) prin HTTP Request (Panou 1 + Panou 2 + toate piesele Panou 3)
        plus verificarea vizuală ROȘU/ROZ în DOM și ștergerea rapidă a rândurilor ROȘII.
        Returnează None dacă comanda nu poate fi rezolvată din cache/opțiuni (semnal de fallback UI).
        """
        if self._context is None:
            return None

        # --- Pas A: Rezolvare comandă ---
        comanda_hit = self._resolve_comanda_from_blueprint(bp, plan.project_query)
        if comanda_hit is None:
            self._log(
                f"[{plan.plan_name}] Comanda '{plan.project_query}' nu există în cache-ul Select2; se rulează prin UI pentru preluare.",
                level="WARN",
            )
            return None
        comanda_id, matched_comanda = comanda_hit

        parts = read_parts_from_excel(plan.excel_path or "")
        if not parts:
            raise RuntimeError(f"Fișierul Excel '{Path(plan.excel_path or '').name}' nu conține piese valide.")

        frame_url = frame.url if hasattr(frame, "url") and frame.url else self.base_url

        # --- Pas B: Panou 1 (creare plan via HTTP POST) ---
        p1_body, p1_headers = self._encode_post_payload(
            fmt=bp.p1_post_fmt,
            base_payload=bp.p1_post_payload,
            updates={bp.p1_plan_key: plan.plan_name},
            frame_url=frame_url,
        )
        resp1 = self._context.request.post(bp.p1_post_url, data=p1_body, headers=p1_headers)
        ok1, err1 = self._is_http_part_response_ok(resp1.status, resp1.text())
        if not ok1:
            raise RuntimeError(f"Panou 1 HTTP refuzat pentru '{plan.plan_name}': {err1}")

        new_plan_id = self._extract_id_from_post_response(resp1.text(), bp.p1_post_id_hint)
        if not new_plan_id:
            new_plan_id = self._lookup_p1_id_via_get(bp, plan.plan_name, frame_url)
        if not new_plan_id:
            raise RuntimeError(f"Nu s-a putut determina ID-ul noului plan '{plan.plan_name}' din Panoul 1.")

        # --- Pas C: Panou 2 (priming fetch + creare legătură comandă via HTTP POST) ---
        p2_get_url_upd = {p: new_plan_id for p in bp.p2_get_plan_params}
        p2_get_body_upd = {k: new_plan_id for k in bp.p2_get_body_plan_keys}
        if bp.p2_get_url:
            self._send_grid_fetch(
                url=bp.p2_get_url,
                method=bp.p2_get_method,
                post_data=bp.p2_get_post_data,
                content_type=bp.p2_get_content_type,
                url_updates=p2_get_url_upd,
                body_updates=p2_get_body_upd,
                paging_updates=None,
                frame_url=frame_url,
                refresh_nd=True,
            )

        p2_post_url_plan = self._replace_url_params(
            bp.p2_post_url,
            {p: new_plan_id for p in bp.p2_post_url_plan_params},
        )
        p2_updates: dict[str, Any] = {
            bp.p2_comanda_key: comanda_id,
            bp.p2_weight_key: plan.part_weight_int if bp.p2_post_fmt == "json" else str(plan.part_weight_int),
        }
        for k in bp.p2_post_body_plan_keys:
            p2_updates[k] = new_plan_id

        p2_body, p2_headers = self._encode_post_payload(
            fmt=bp.p2_post_fmt,
            base_payload=bp.p2_post_payload,
            updates=p2_updates,
            frame_url=frame_url,
        )
        resp2 = self._context.request.post(p2_post_url_plan, data=p2_body, headers=p2_headers)
        ok2, err2 = self._is_http_part_response_ok(resp2.status, resp2.text())
        if not ok2:
            raise RuntimeError(f"Panou 2 HTTP refuzat pentru '{plan.plan_name}': {err2}")

        new_rowid2 = self._extract_id_from_post_response(resp2.text(), bp.p2_post_id_hint)
        if not new_rowid2 and bp.p2_get_url:
            resp2_get = self._send_grid_fetch(
                url=bp.p2_get_url,
                method=bp.p2_get_method,
                post_data=bp.p2_get_post_data,
                content_type=bp.p2_get_content_type,
                url_updates=p2_get_url_upd,
                body_updates=p2_get_body_upd,
                paging_updates=None,
                frame_url=frame_url,
                refresh_nd=True,
            )
            rows2 = self._parse_jqgrid_json_rows(resp2_get.text())
            if rows2:
                new_rowid2 = self._find_row_id_in_jqgrid_json(
                    rows2,
                    matched_comanda or plan.project_query,
                    id_key=bp.p2_id_key,
                    single_row_fallback=True,
                )
        if not new_rowid2:
            raise RuntimeError(
                f"Nu s-a putut determina ID-ul legăturii Panou 2 pentru '{plan.plan_name}' (plan_id={new_plan_id})."
            )

        # --- Pas D: Panou 3 (priming fetch + toate piesele 1..N via HTTP POST) ---
        p3_get_updates: dict[str, str] = {}
        for p in bp.p3_get_rowid_params:
            p3_get_updates[p] = new_rowid2
        for p in bp.p3_get_comanda_params:
            p3_get_updates[p] = comanda_id
        for p in bp.p3_get_plan_params:
            p3_get_updates[p] = new_plan_id

        p3_get_body_updates: dict[str, str] = {}
        for k in bp.p3_get_body_rowid_keys:
            p3_get_body_updates[k] = new_rowid2
        for k in bp.p3_get_body_comanda_keys:
            p3_get_body_updates[k] = comanda_id
        for k in bp.p3_get_body_plan_keys:
            p3_get_body_updates[k] = new_plan_id

        p3_get_url_plan = self._replace_url_params(bp.p3_get_url, p3_get_updates, refresh_nd=True)
        self._send_grid_fetch(
            url=bp.p3_get_url,
            method=bp.p3_get_method,
            post_data=bp.p3_get_post_data,
            content_type=bp.p3_get_content_type,
            url_updates=p3_get_updates,
            body_updates=p3_get_body_updates,
            paging_updates=None,
            frame_url=frame_url,
            refresh_nd=True,
        )

        p3_post_url_updates: dict[str, str] = {}
        for p in bp.p3_post_url_rowid_params:
            p3_post_url_updates[p] = new_rowid2
        for p in bp.p3_post_url_comanda_params:
            p3_post_url_updates[p] = comanda_id
        for p in bp.p3_post_url_plan_params:
            p3_post_url_updates[p] = new_plan_id
        p3_post_url_plan = self._replace_url_params(bp.p3_post_url, p3_post_url_updates)

        p3_post_fixed_body: dict[str, Any] = {}
        for k in bp.p3_post_body_rowid_keys:
            p3_post_fixed_body[k] = new_rowid2
        for k in bp.p3_post_body_comanda_keys:
            p3_post_fixed_body[k] = comanda_id
        for k in bp.p3_post_body_plan_keys:
            p3_post_fixed_body[k] = new_plan_id

        total_parts = len(parts)
        self._log(
            f"[{plan.plan_name}] Full HTTP: P1(id={new_plan_id}) -> P2(rowid={new_rowid2}, "
            f"Comanda='{matched_comanda}') -> P3({total_parts} piese via HTTP)..."
        )

        parts_added = 0
        parts_rejected = 0
        for idx, (part_name, qty) in enumerate(parts, start=1):
            p3_updates = dict(p3_post_fixed_body)
            p3_updates[bp.p3_part_key] = part_name
            p3_updates[bp.p3_qty_key] = qty if bp.p3_post_fmt == "json" else str(qty)
            body_arg, headers = self._encode_post_payload(
                fmt=bp.p3_post_fmt,
                base_payload=bp.p3_post_payload,
                updates=p3_updates,
                frame_url=frame_url,
            )
            resp = self._context.request.post(p3_post_url_plan, data=body_arg, headers=headers)
            ok, err_detail = self._is_http_part_response_ok(resp.status, resp.text())
            if ok:
                parts_added += 1
                self._log(f"  [{plan.plan_name}] [{idx}/{total_parts}] HTTP OK: {part_name} x {qty}")
            else:
                parts_rejected += 1
                self._log(
                    f"  [{plan.plan_name}] [{idx}/{total_parts}] Piesa '{part_name}' (cant. {qty}) respinsă HTTP: {err_detail}",
                    level="WARN",
                )

        # --- Pas E: Reîncărcare Panou 3 în DOM + clasificare ROȘU/ROZ + ștergere HTTP rapidă ---
        panel3 = self._find_panel_grid(frame, "panel3_titles", fallback_index=2)
        p3_grid_base_url = self._replace_url_params(p3_get_url_plan, {}, strip_jqgrid_paging=True)
        p3_postdata_updates = {**p3_get_updates, **p3_get_body_updates}
        self._reload_jqgrid(
            page,
            frame,
            panel3,
            new_base_url=p3_grid_base_url,
            edit_url=p3_post_url_plan,
            post_data_updates=p3_postdata_updates,
            wait_network_idle=False,
        )

        rows_info = self._inspect_panel3_rows(panel3)
        red_rows: list[dict[str, Any]] = []
        pink_kept = 0
        missing_row_id = False

        for info in rows_info:
            category, reason = classify_row_color(
                tr_bg=info["trBg"],
                cell_bgs=info["cellBgs"],
                first_cell_text=info["firstCellText"],
                row_classes=info["rowClasses"],
                cell_classes=info["cellClasses"],
            )
            if category == "red":
                red_rows.append({"id": info["id"], "reason": reason, "preview": info["preview"]})
                if not info["id"]:
                    missing_row_id = True
            elif category == "pink":
                pink_kept += 1
                self._log(f"  [{plan.plan_name}] Rând ROZ păstrat ({reason}): {info['preview']}")

        red_deleted = 0
        if red_rows:
            if missing_row_id:
                red_deleted, extra_pink = self._clean_panel3_rows(page, frame, panel3, plan.plan_name)
                if pink_kept == 0 and extra_pink > 0:
                    pink_kept = extra_pink
            else:
                targeted_ids: set[str] = set()
                for rr in red_rows:
                    rid = str(rr["id"])
                    self._log(f"  [{plan.plan_name}] Ștergere HTTP rând ROȘU ({rr['reason']}): {rr['preview']}")
                    del_body, del_headers = self._encode_post_payload(
                        fmt="form",
                        base_payload=[],
                        updates={"oper": "del", "id": rid},
                        frame_url=frame_url,
                    )
                    del_resp = self._context.request.post(p3_post_url_plan, data=del_body, headers=del_headers)
                    ok_del, err_del = self._is_http_part_response_ok(del_resp.status, del_resp.text())
                    if ok_del:
                        targeted_ids.add(rid)
                    else:
                        self._log(
                            f"  [{plan.plan_name}] Ștergere HTTP eșuată pentru id={rid}: {err_del}",
                            level="WARN",
                        )

                if targeted_ids:
                    self._reload_jqgrid(
                        page,
                        frame,
                        panel3,
                        new_base_url=p3_grid_base_url,
                        edit_url=p3_post_url_plan,
                        post_data_updates=p3_get_updates,
                        wait_network_idle=False,
                    )
                rows_after = self._inspect_panel3_rows(panel3)
                remaining_ids = {str(info["id"]) for info in rows_after if info["id"]}
                red_deleted = len(targeted_ids - remaining_ids)

                still_has_red = any(
                    classify_row_color(
                        tr_bg=info["trBg"],
                        cell_bgs=info["cellBgs"],
                        first_cell_text=info["firstCellText"],
                        row_classes=info["rowClasses"],
                        cell_classes=info["cellClasses"],
                    )[0]
                    == "red"
                    for info in rows_after
                )
                if still_has_red:
                    extra_del, _ = self._clean_panel3_rows(page, frame, panel3, plan.plan_name)
                    red_deleted += extra_del

        return self._finalize_stage1_result(
            page=page,
            panel3=panel3,
            plan=plan,
            matched_comanda=matched_comanda,
            parts_added=parts_added,
            parts_rejected=parts_rejected,
            red_deleted=red_deleted,
            pink_kept=pink_kept,
            mode_tag="Full HTTP",
        )

    def _run_stage1(
        self,
        plans: list[CuttingPlanRecord],
        panel3_mode: str = "manual_plus",
    ) -> list[Stage1PlanResult]:
        results: list[Stage1PlanResult] = []
        total = len(plans)
        self._set_progress(active=True, stage="Etapa 1", current=0, total=total, plan="", message="Conectare la intranet...")

        page = self._ensure_browser()
        frame = self._login_and_navigate(page)
        blueprints: dict[str, _FullHttpBlueprint] = {}
        used_full_http = False

        for idx, plan in enumerate(plans, start=1):
            if page.is_closed():
                self._log("Browserul a fost închis manual; execuția lotului a fost oprită.", level="WARN")
                break

            self._set_progress(
                active=True,
                stage="Etapa 1",
                current=idx - 1,
                total=total,
                plan=plan.plan_name,
                message=f"Procesare {plan.plan_name} ({idx}/{total})...",
            )
            self._log(
                f"=== ETAPA 1 [{idx}/{total}]: {plan.plan_name} "
                f"(Debitare={plan.cutting_type}, Comanda='{plan.project_query}', "
                f"Greutate={plan.part_weight_int}, ModPiese={panel3_mode}) ==="
            )

            if not plan.excel_path or not Path(plan.excel_path).exists():
                msg = "Lipsește fișierul .xlsx asociat."
                self._log(f"[{plan.plan_name}] Sărit: {msg}", level="ERROR")
                results.append(Stage1PlanResult(plan_name=plan.plan_name, success=False, error=msg))
                continue

            if plan.part_weight_int is None:
                msg = "Lipsește greutatea extrasă din Word."
                self._log(f"[{plan.plan_name}] Sărit: {msg}", level="ERROR")
                results.append(Stage1PlanResult(plan_name=plan.plan_name, success=False, error=msg))
                continue

            try:
                frame = self._get_work_frame(page)
                ct_key = _norm(plan.cutting_type)
                res: Optional[Stage1PlanResult] = None

                if panel3_mode == "full_http" and ct_key in blueprints:
                    res = self._run_plan_full_http(page, frame, blueprints[ct_key], plan)
                    if res is not None:
                        used_full_http = True
                    else:
                        res, _ = self._run_single_plan_ui(
                            page=page,
                            plan=plan,
                            effective_panel3_mode="http_request",
                            calibrate=False,
                            cache_blueprint=blueprints[ct_key],
                        )
                elif panel3_mode == "full_http":
                    res, new_bp = self._run_single_plan_ui(
                        page=page,
                        plan=plan,
                        effective_panel3_mode="http_request",
                        calibrate=True,
                    )
                    if new_bp is not None:
                        blueprints[ct_key] = new_bp
                else:
                    res, _ = self._run_single_plan_ui(
                        page=page,
                        plan=plan,
                        effective_panel3_mode=panel3_mode,
                        calibrate=False,
                    )

                results.append(res)
            except Exception as exc:
                err_msg = str(exc)
                if page.is_closed() or "target page, context or browser has been closed" in err_msg.lower():
                    self._log("Browserul a fost închis manual; execuția lotului a fost oprită.", level="WARN")
                    results.append(
                        Stage1PlanResult(
                            plan_name=plan.plan_name,
                            success=False,
                            error="Browser închis manual.",
                        )
                    )
                    break

                self._log(f"[{plan.plan_name}] EROARE Etapa 1: {exc}", level="ERROR")
                self._save_error_screenshot(page, plan.plan_name, "stage1")
                self._dismiss_any_open_modals(page, self._get_work_frame(page))
                results.append(
                    Stage1PlanResult(
                        plan_name=plan.plan_name,
                        success=False,
                        error=err_msg,
                    )
                )

            self._set_progress(current=idx)

        if used_full_http and not page.is_closed():
            try:
                frame = self._get_work_frame(page)
                panel1 = self._find_panel_grid(frame, "panel1_titles", fallback_index=0)
                self._clear_grid_search_filters(page, frame, panel1)
                self._reload_jqgrid(page, frame, panel1, wait_network_idle=False)
                self._log("Panoul 1 reîncărcat pentru verificarea vizuală a tuturor planurilor adăugate.")
            except Exception as exc:
                self._log(f"Avertisment reîncărcare finală Panou 1: {exc}", level="WARN")

        self._set_progress(active=False, message="Etapa 1 finalizată. Sesiunea browser rămâne deschisă pentru verificare.")
        self._log("=== ETAPA 1 FINALIZATĂ. Sesiunea browser a fost păstrată deschisă. ===")
        return results

    # ------------------------------------------------------------------
    # Stage 2: Validate, Capture QR, Insert into Word & Print
    # ------------------------------------------------------------------
    def _extract_qr_png(
        self,
        page: Page,
        frame: Page | Frame,
        panel1: Locator,
        plan_name: str,
        qr_temp_dir: Path,
    ) -> Path:
        """Localizează QRcode în rândul planului din Panoul 1 și îl salvează ca PNG."""
        qr_temp_dir.mkdir(parents=True, exist_ok=True)
        out_png = qr_temp_dir / f"{re.sub(r'[^A-Za-z0-9._-]+', '_', plan_name)}_qr.png"

        # Găsește indexul coloanei QRcode din antetul Panoului 1
        qr_col_idx = panel1.evaluate(
            """(gridEl) => {
                const headers = Array.from(gridEl.querySelectorAll('.ui-jqgrid-htable tr.ui-jqgrid-labels th'));
                for (let i = 0; i < headers.length; i++) {
                    const txt = (headers[i].textContent || '').trim().toLowerCase();
                    if (txt.includes('qr')) return i;
                }
                return -1;
            }"""
        )

        deadline = time.time() + 15.0
        qr_el: Optional[Locator] = None
        tag_name = ""

        while time.time() < deadline:
            row = self._find_and_select_grid_row(
                page,
                frame,
                panel1,
                plan_name,
                col_hints=["Plan croi", "Plan"],
                click_to_select=False,
            )
            search_scope = row.locator("td").nth(qr_col_idx) if qr_col_idx >= 0 else row
            cand = search_scope.locator("img, canvas").first
            if cand.count() > 0 and cand.is_visible():
                tag_name = (cand.evaluate("e => e.tagName.toLowerCase()") or "").strip()
                if tag_name == "img":
                    loaded = cand.evaluate("e => Boolean(e.complete && e.naturalWidth > 0)")
                    if loaded:
                        qr_el = cand
                        break
                else:
                    qr_el = cand
                    break
            time.sleep(0.35)

        if qr_el is None:
            raise RuntimeError(f"Codul QR nu a apărut în coloana 'QRcode' pentru '{plan_name}'.")

        saved = False
        if tag_name == "img":
            src = (qr_el.get_attribute("src") or "").strip()
            if src.startswith("data:image/"):
                try:
                    _, b64_part = src.split(",", 1)
                    raw_bytes = base64.b64decode(b64_part)
                    with Image.open(io.BytesIO(raw_bytes)) as im:
                        im.save(out_png, format="PNG")
                    saved = True
                except Exception:
                    saved = False
            elif src and self._context is not None:
                try:
                    frame_url = frame.url if hasattr(frame, "url") and frame.url else self.base_url
                    abs_url = urljoin(frame_url, src)
                    resp = self._context.request.get(abs_url)
                    if resp.ok:
                        raw_bytes = resp.body()
                        with Image.open(io.BytesIO(raw_bytes)) as im:
                            im.save(out_png, format="PNG")
                        saved = True
                except Exception:
                    saved = False

        if not saved and tag_name == "canvas":
            try:
                data_url = qr_el.evaluate("c => c.toDataURL('image/png')")
                if isinstance(data_url, str) and "," in data_url:
                    _, b64_part = data_url.split(",", 1)
                    raw_bytes = base64.b64decode(b64_part)
                    with Image.open(io.BytesIO(raw_bytes)) as im:
                        im.save(out_png, format="PNG")
                    saved = True
            except Exception:
                saved = False

        if not saved:
            qr_el.scroll_into_view_if_needed()
            qr_el.screenshot(path=str(out_png))
            with Image.open(out_png) as im:
                im.verify()

        return out_png

    def _run_stage2(
        self,
        plans: list[CuttingPlanRecord],
        output_dir: Path,
        qr_temp_dir: Path,
        send_to_printer: bool,
        print_delay: float,
        printer_name: Optional[str] = None,
        overwrite_original: bool = True,
    ) -> list[Stage2PlanResult]:
        if self._page is None or self._page.is_closed():
            self._log("Browserul nu este deschis; se pornește și se navighează automat pentru Etapa 2...")
            page = self._ensure_browser()
            frame = self._login_and_navigate(page)
        else:
            page = self._page
            frame = self._get_work_frame(page)
            try:
                if frame.locator(SELECTORS["jqgrid"]["grid_box"]).count() == 0:
                    frame = self._login_and_navigate(page)
            except Exception:
                frame = self._login_and_navigate(page)

        results: list[Stage2PlanResult] = []
        total = len(plans)
        self._set_progress(active=True, stage="Etapa 2", current=0, total=total, plan="", message="Pornire validare și QR...")
        last_panel1: Optional[Locator] = None

        for idx, plan in enumerate(plans, start=1):
            if page.is_closed():
                self._log("Browserul a fost închis manual; execuția lotului a fost oprită.", level="WARN")
                break

            self._set_progress(
                active=True,
                stage="Etapa 2",
                current=idx - 1,
                total=total,
                plan=plan.plan_name,
                message=f"Validare & QR {plan.plan_name} ({idx}/{total})...",
            )
            self._log(f"=== ETAPA 2 [{idx}/{total}]: {plan.plan_name} ===")

            if not plan.working_docx_path or not Path(plan.working_docx_path).exists():
                msg = "Lipsește documentul Word pregătit (.docx)."
                self._log(f"[{plan.plan_name}] Sărit: {msg}", level="ERROR")
                results.append(Stage2PlanResult(plan_name=plan.plan_name, success=False, error=msg))
                continue

            try:
                frame = self._get_work_frame(page)
                self._dismiss_any_open_modals(page, frame)
                self._select_debitare_dropdown(page, frame, plan.cutting_type)
                panel1 = self._find_panel_grid(frame, "panel1_titles", fallback_index=0)
                last_panel1 = panel1

                row = self._find_and_select_grid_row(
                    page,
                    frame,
                    panel1,
                    plan.plan_name,
                    col_hints=["Plan croi", "Plan"],
                    click_to_select=False,
                )

                edit_btn = self._first_visible(row, [SELECTORS["jqgrid"]["inline_edit"]])
                if edit_btn is not None:
                    edit_btn.click()
                else:
                    row.click()
                    self._wait_frame_ajax(page, frame)
                    pager_edit = self._first_visible(panel1, [SELECTORS["jqgrid"]["edit_pager_button"]])
                    if pager_edit is None:
                        raise RuntimeError("Nu s-a găsit butonul de editare (creion) pentru rândul din Panoul 1.")
                    pager_edit.click()

                edit_modal = self._wait_visible_modal(page, frame)
                valid_res = edit_modal.evaluate(
                    """(el) => {
                        const strip = (s) => (s || '')
                            .normalize('NFKD')
                            .replace(/[\\u0300-\\u036f]/g, '')
                            .replace(/\\s+/g, ' ')
                            .trim()
                            .toLowerCase();

                        const rows = Array.from(el.querySelectorAll('tr.FormData, tr, .form-group'));
                        for (const r of rows) {
                            const txt = strip(r.textContent);
                            if (txt.includes('valid')) {
                                const cb = r.querySelector('input[type="checkbox"]');
                                if (cb) {
                                    const wasAlreadyChecked = Boolean(cb.checked);
                                    if (!wasAlreadyChecked) cb.click();
                                    return { found: true, wasAlreadyChecked };
                                }
                            }
                        }
                        const byAttr = el.querySelector('input[type="checkbox"][name*="valid" i], input[type="checkbox"][id*="valid" i]');
                        if (byAttr) {
                            const wasAlreadyChecked = Boolean(byAttr.checked);
                            if (!wasAlreadyChecked) byAttr.click();
                            return { found: true, wasAlreadyChecked };
                        }
                        return { found: false, wasAlreadyChecked: false };
                    }"""
                )
                if not valid_res.get("found"):
                    raise RuntimeError("Nu s-a găsit checkbox-ul 'Valid' în modalul 'Modificare înregistrare'.")

                if valid_res.get("wasAlreadyChecked"):
                    self._close_add_part_modal(page, frame)
                    self._log(f"[{plan.plan_name}] Planul avea deja bifa 'Valid' activă.")
                else:
                    self._click_save_in_modal(page, frame, edit_modal)
                    self._log(f"[{plan.plan_name}] Bifat 'Valid' și salvat.")

                qr_png_path = self._extract_qr_png(page, frame, panel1, plan.plan_name, qr_temp_dir)
                self._log(f"[{plan.plan_name}] QR salvat: {qr_png_path.name}")

                orig_path_str = (plan.source_word_path or "").strip()
                is_upload_copy = bool(orig_path_str) and _is_uploaded_temp_word_path(orig_path_str)
                do_overwrite = bool(overwrite_original and orig_path_str and not is_upload_copy)
                if overwrite_original and is_upload_copy:
                    self._log(
                        f"[{plan.plan_name}] Fișier încărcat prin Drag & Drop (fără cale originală pe disc); se va salva în '{output_dir}'.",
                        level="WARN",
                    )

                out_docx, printed = insert_qr_and_optionally_print(
                    source_docx_path=plan.working_docx_path,
                    qr_image_path=qr_png_path,
                    output_dir=output_dir,
                    output_filename=f"{plan.plan_name}_QR.docx",
                    qr_size_cm=2.2,
                    send_to_printer=send_to_printer,
                    print_delay=print_delay,
                    printer_name=printer_name,
                    overwrite_original=do_overwrite,
                    original_word_path=(orig_path_str if do_overwrite else None),
                )
                printer_label = (printer_name or "").strip() or "imprimanta implicită"
                word_action_msg = (
                    f"Original suprascris cu QR: {out_docx}"
                    if do_overwrite
                    else f"Word generat: {out_docx.name}"
                )
                self._log(
                    f"[{plan.plan_name}] {word_action_msg}"
                    + (f" | Trimis la imprimantă ({printer_label})." if printed else " | Printare dezactivată (mod test).")
                )

                results.append(
                    Stage2PlanResult(
                        plan_name=plan.plan_name,
                        success=True,
                        qr_path=str(qr_png_path),
                        output_docx_path=str(out_docx),
                        printed=printed,
                    )
                )
            except Exception as exc:
                err_msg = str(exc)
                if page.is_closed() or "target page, context or browser has been closed" in err_msg.lower():
                    self._log("Browserul a fost închis manual; execuția lotului a fost oprită.", level="WARN")
                    results.append(
                        Stage2PlanResult(
                            plan_name=plan.plan_name,
                            success=False,
                            error="Browser închis manual.",
                        )
                    )
                    break

                self._log(f"[{plan.plan_name}] EROARE Etapa 2: {exc}", level="ERROR")
                self._save_error_screenshot(page, plan.plan_name, "stage2")
                self._dismiss_any_open_modals(page, self._get_work_frame(page))
                results.append(
                    Stage2PlanResult(
                        plan_name=plan.plan_name,
                        success=False,
                        error=err_msg,
                    )
                )

            self._set_progress(current=idx)

        if last_panel1 is not None and not page.is_closed():
            try:
                self._clear_grid_search_filters(page, self._get_work_frame(page), last_panel1)
            except Exception:
                pass

        self._set_progress(active=False, message="Etapa 2 finalizată.")
        self._log("=== ETAPA 2 FINALIZATĂ. ===")
        return results
