"""Launcher și manager de ciclu de viață pentru aplicația web Streamlit din `bonuri consum`."""

from __future__ import annotations

import os
import shutil
import socket
import subprocess
import sys
import tempfile
import threading
import time
import urllib.error
import urllib.request
import webbrowser
from pathlib import Path
from typing import IO, Optional


_BUNDLE_SOURCE_FILES: tuple[str, ...] = (
    "app.py",
    "browser_bot.py",
    "parser_docs.py",
    "settings_store.py",
    "word_qr_printer.py",
    "requirements.txt",
)


def _creationflags_no_window() -> int:
    if sys.platform.startswith("win"):
        return getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)
    return 0


def build_clean_subprocess_env(python_exe: Optional[Path] = None) -> dict[str, str]:
    """Curăță variabilele PyInstaller (_PYI*/_MEI*) și Python înainte de a lansa procesul copil."""
    env = os.environ.copy()
    for var in list(env.keys()):
        uvar = var.upper()
        if (
            uvar.startswith("_PYI")
            or uvar.startswith("PYI")
            or uvar.startswith("_MEI")
            or uvar.startswith("MEI")
            or uvar in {"PYTHONHOME", "PYTHONPATH", "TCL_LIBRARY", "TK_LIBRARY"}
        ):
            env.pop(var, None)

    if "PATH" in env:
        clean_parts = [
            p for p in env["PATH"].split(os.pathsep)
            if p and "_MEI" not in p.upper()
        ]
        if python_exe is not None:
            scripts_dir = str(Path(python_exe).resolve().parent)
            clean_parts = [scripts_dir] + [p for p in clean_parts if p.lower() != scripts_dir.lower()]
        env["PATH"] = os.pathsep.join(clean_parts)
    elif python_exe is not None:
        env["PATH"] = str(Path(python_exe).resolve().parent)

    env["STREAMLIT_BROWSER_GATHER_USAGE_STATS"] = "false"
    env["PYTHONUTF8"] = "1"
    return env


def _sync_bundled_app(bundle_dir: Path) -> Optional[Path]:
    """Sincronizează sursele curate din PyInstaller _MEIPASS în %APPDATA%\\SteelSync\\bonuri_consum."""
    if not (bundle_dir / "app.py").is_file():
        return None
    appdata = os.environ.get("APPDATA") or str(Path.home())
    target_dir = Path(appdata) / "SteelSync" / "bonuri_consum"
    target_dir.mkdir(parents=True, exist_ok=True)

    for rel in _BUNDLE_SOURCE_FILES:
        src = bundle_dir / rel
        if src.is_file():
            shutil.copy2(src, target_dir / rel)

    src_cfg = bundle_dir / ".streamlit" / "config.toml"
    if src_cfg.is_file():
        dst_st_dir = target_dir / ".streamlit"
        dst_st_dir.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src_cfg, dst_st_dir / "config.toml")

    return target_dir if (target_dir / "app.py").is_file() else None


def resolve_app_dir(
    saved_dir: str = "",
    base_dir_override: Optional[Path] = None,
) -> Optional[Path]:
    """Găsește folderul aplicației `bonuri consum` care conține `app.py`."""
    candidates: list[Path] = []
    if saved_dir and str(saved_dir).strip():
        candidates.append(Path(str(saved_dir).strip()))

    if base_dir_override is not None:
        base_dir = Path(base_dir_override).resolve()
    elif getattr(sys, "frozen", False):
        base_dir = Path(sys.executable).resolve().parent
    else:
        base_dir = Path(__file__).resolve().parent

    for parent in (base_dir, base_dir.parent):
        candidates.append(parent / "bonuri consum")
        candidates.append(parent / "Bonuri consum")
        candidates.append(parent / "AUTOMATIZARE INTRODUCERE PLANURI CROI")

    docs_dir = Path.home() / "Documents"
    candidates.extend([
        docs_dir / "daniel" / "aplicatie sortare piese" / "bonuri consum",
        docs_dir / "daniel" / "AUTOMATIZARE INTRODUCERE PLANURI CROI",
    ])

    seen: set[str] = set()
    for cand in candidates:
        try:
            key = os.path.normcase(str(cand.resolve()))
        except OSError:
            key = os.path.normcase(str(cand))
        if key in seen:
            continue
        seen.add(key)
        if cand.is_dir() and (cand / "app.py").is_file():
            return cand.resolve()

    # Scanare limitată (până la adâncime 3) sub Documents pentru un folder "bonuri consum" cu app.py
    if docs_dir.is_dir():
        queue: list[tuple[Path, int]] = [(docs_dir, 0)]
        while queue:
            cur, depth = queue.pop(0)
            try:
                children = sorted(cur.iterdir(), key=lambda p: p.name.lower())
            except OSError:
                continue
            for child in children:
                if not child.is_dir():
                    continue
                low = child.name.lower()
                if low in {".git", ".venv", "venv", "node_modules", "__pycache__", ".next", "dist", "build"}:
                    continue
                if low == "bonuri consum" and (child / "app.py").is_file():
                    return child.resolve()
                if depth + 1 < 3:
                    queue.append((child, depth + 1))

    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        synced = _sync_bundled_app(Path(meipass) / "bonuri_consum")
        if synced is not None:
            return synced

    return None


def _python_has_streamlit(py_path: Path) -> bool:
    """Verifică rapid dacă interpretorul Python indicat are modulul `streamlit` instalat."""
    if not py_path.is_file():
        return False
    if (py_path.parent / "streamlit.exe").is_file():
        return True
    for sp in (
        py_path.parent.parent / "Lib" / "site-packages" / "streamlit",
        py_path.parent / "Lib" / "site-packages" / "streamlit",
    ):
        if sp.is_dir():
            return True
    try:
        proc = subprocess.run(
            [str(py_path), "-c", "import streamlit"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            env=build_clean_subprocess_env(py_path),
            timeout=5,
            creationflags=_creationflags_no_window(),
        )
        return proc.returncode == 0
    except Exception:
        return False


def resolve_python_for_streamlit(
    app_dir: Path,
    saved_python: str = "",
) -> Optional[Path]:
    """Identifică un interpretor Python valid (preferând .venv) care are Streamlit instalat."""
    candidates: list[Path] = []
    if saved_python and str(saved_python).strip():
        candidates.append(Path(str(saved_python).strip()))

    candidates.extend([
        app_dir / ".venv" / "Scripts" / "python.exe",
        app_dir / "venv" / "Scripts" / "python.exe",
    ])

    # Scanează mediile virtuale din proiectele învecinate / părinte (ex. AUTOMATIZARE ...\.venv)
    for root in (app_dir.parent, app_dir.parent.parent):
        if root.is_dir():
            candidates.extend([
                root / ".venv" / "Scripts" / "python.exe",
                root / "venv" / "Scripts" / "python.exe",
            ])
            try:
                children = sorted(root.iterdir(), key=lambda p: p.name.lower())
            except OSError:
                children = []
            for child in children:
                if child.is_dir():
                    candidates.extend([
                        child / ".venv" / "Scripts" / "python.exe",
                        child / "venv" / "Scripts" / "python.exe",
                    ])

    st_cmd = shutil.which("streamlit")
    if st_cmd:
        candidates.append(Path(st_cmd).resolve().parent / "python.exe")

    if not getattr(sys, "frozen", False):
        candidates.append(Path(sys.executable).resolve())

    py_cmd = shutil.which("python")
    if py_cmd:
        candidates.append(Path(py_cmd).resolve())

    seen: set[str] = set()
    for py in candidates:
        try:
            key = os.path.normcase(str(py.resolve()))
        except OSError:
            key = os.path.normcase(str(py))
        if key in seen:
            continue
        seen.add(key)
        if _python_has_streamlit(py):
            return py.resolve()

    return None


def find_free_port(start_port: int = 8501, end_port: int = 8520) -> int:
    """Returnează primul port TCP liber pe 127.0.0.1 în intervalul dat."""
    for port in range(start_port, end_port + 1):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            try:
                sock.bind(("127.0.0.1", port))
                return port
            except OSError:
                continue
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind(("127.0.0.1", 0))
        return int(sock.getsockname()[1])


class BonConsumLauncher:
    """Gestionează pornirea, verificarea stării, deschiderea în browser și oprirea aplicației Streamlit."""

    def __init__(self) -> None:
        self._lock = threading.Lock()
        self.process: Optional[subprocess.Popen[bytes]] = None
        self.port: Optional[int] = None
        self.app_dir: Optional[Path] = None
        self.python_exe: Optional[Path] = None
        self.log_path: Path = Path(tempfile.gettempdir()) / "steelsync_bon_consum.log"
        self._log_file: Optional[IO[bytes]] = None

    @staticmethod
    def _check_health(port: int, timeout: float = 0.6) -> bool:
        url = f"http://127.0.0.1:{port}/_stcore/health"
        try:
            req = urllib.request.Request(url, method="GET")
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                return resp.status == 200
        except (urllib.error.URLError, TimeoutError, OSError):
            return False

    def is_running(self) -> bool:
        with self._lock:
            if self.process is None or self.port is None:
                return False
            if self.process.poll() is not None:
                return False
            port = self.port
        return self._check_health(port, timeout=0.6)

    def _read_log_tail(self, max_lines: int = 25) -> str:
        try:
            if self.log_path.is_file():
                lines = self.log_path.read_text(encoding="utf-8", errors="replace").splitlines()
                return "\n".join(lines[-max_lines:])
        except OSError:
            pass
        return ""

    def _close_log_handle(self) -> None:
        if self._log_file is not None:
            try:
                self._log_file.flush()
            except Exception:
                pass
            try:
                self._log_file.close()
            except Exception:
                pass
            self._log_file = None

    def start_or_open(
        self,
        saved_dir: str = "",
        saved_python: str = "",
        preferred_port: int = 8501,
        open_browser: bool = True,
        startup_timeout: float = 25.0,
    ) -> dict[str, object]:
        """Pornește serverul Streamlit dacă nu rulează deja și deschide URL-ul în browser."""
        with self._lock:
            existing_proc = self.process
            existing_port = self.port
            existing_dir = self.app_dir
            existing_py = self.python_exe

        if (
            existing_proc is not None
            and existing_proc.poll() is None
            and existing_port is not None
            and self._check_health(existing_port, timeout=0.6)
        ):
            url = f"http://127.0.0.1:{existing_port}"
            if open_browser:
                webbrowser.open(url, new=2)
            return {
                "status": "opened",
                "url": url,
                "port": existing_port,
                "app_dir": str(existing_dir) if existing_dir else "",
                "python": str(existing_py) if existing_py else "",
                "pid": existing_proc.pid,
            }

        self.stop()

        app_dir = resolve_app_dir(saved_dir=saved_dir)
        if app_dir is None:
            return {
                "status": "missing_dir",
                "message": "Folderul 'bonuri consum' (cu app.py) nu a fost găsit automat.",
            }

        python_exe = resolve_python_for_streamlit(app_dir, saved_python=saved_python)
        if python_exe is None:
            return {
                "status": "missing_python",
                "app_dir": str(app_dir),
                "message": "Nu a fost găsit un interpretor Python (.venv) cu Streamlit instalat.",
            }

        port = find_free_port(preferred_port, preferred_port + 19)
        url = f"http://127.0.0.1:{port}"
        cmd = [
            str(python_exe),
            "-m",
            "streamlit",
            "run",
            str(app_dir / "app.py"),
            "--server.address",
            "127.0.0.1",
            "--server.port",
            str(port),
            "--server.headless",
            "true",
            "--browser.gatherUsageStats",
            "false",
        ]
        env = build_clean_subprocess_env(python_exe)

        try:
            self._close_log_handle()
            self._log_file = open(self.log_path, "wb")
            proc = subprocess.Popen(
                cmd,
                cwd=str(app_dir),
                env=env,
                stdout=self._log_file,
                stderr=subprocess.STDOUT,
                stdin=subprocess.DEVNULL,
                creationflags=_creationflags_no_window(),
            )
        except Exception as exc:
            self._close_log_handle()
            return {
                "status": "error",
                "app_dir": str(app_dir),
                "python": str(python_exe),
                "message": f"Eroare la lansarea procesului Streamlit: {exc}",
            }

        with self._lock:
            self.process = proc
            self.port = port
            self.app_dir = app_dir
            self.python_exe = python_exe

        deadline = time.monotonic() + max(1.0, float(startup_timeout))
        while time.monotonic() < deadline:
            if proc.poll() is not None:
                tail = self._read_log_tail()
                self.stop()
                return {
                    "status": "error",
                    "app_dir": str(app_dir),
                    "python": str(python_exe),
                    "message": f"Serverul Streamlit s-a închis prematur (cod {proc.returncode}).",
                    "log_tail": tail,
                }
            if self._check_health(port, timeout=0.5):
                if open_browser:
                    webbrowser.open(url, new=2)
                return {
                    "status": "started",
                    "url": url,
                    "port": port,
                    "app_dir": str(app_dir),
                    "python": str(python_exe),
                    "pid": proc.pid,
                }
            time.sleep(0.25)

        tail = self._read_log_tail()
        self.stop()
        return {
            "status": "error",
            "app_dir": str(app_dir),
            "python": str(python_exe),
            "message": f"Serverul Streamlit nu a răspuns în {startup_timeout:.0f}s pe {url}.",
            "log_tail": tail,
        }

    def stop(self) -> None:
        """Oprește procesul copil Streamlit (și arborii săi pe Windows)."""
        with self._lock:
            proc = self.process
            self.process = None
            self.port = None

        if proc is not None:
            try:
                if proc.poll() is None:
                    proc.terminate()
                    try:
                        proc.wait(timeout=3.0)
                    except subprocess.TimeoutExpired:
                        if sys.platform.startswith("win"):
                            subprocess.run(
                                ["taskkill", "/F", "/T", "/PID", str(proc.pid)],
                                stdout=subprocess.DEVNULL,
                                stderr=subprocess.DEVNULL,
                                creationflags=_creationflags_no_window(),
                                timeout=5,
                            )
                        else:
                            proc.kill()
                        proc.wait(timeout=2.0)
            except Exception:
                pass

        self._close_log_handle()
