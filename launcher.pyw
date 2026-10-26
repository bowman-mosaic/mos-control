"""Windows desktop launcher for Mosaic Control Center.

Starts the Flask control server without a console window, opens the UI in a
dedicated Edge/Chrome app window, and gracefully stops the server when that
window closes.
"""

from __future__ import annotations

import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile
import time
import urllib.error
import urllib.request


ROOT = Path(__file__).resolve().parent
APP_URL = "http://localhost:8081"
SHUTDOWN_URL = f"{APP_URL}/api/shutdown"
STARTUP_TIMEOUT_SECONDS = 30
SHUTDOWN_TIMEOUT_SECONDS = 30


def _show_error(message: str) -> None:
    """Display launcher errors even though this script has no console."""
    try:
        import tkinter
        from tkinter import messagebox

        root = tkinter.Tk()
        root.withdraw()
        messagebox.showerror("Mosaic Control Center", message, parent=root)
        root.destroy()
    except Exception:
        pass


def _find_browser() -> Path | None:
    """Find Edge or Chrome, preferring Edge on Windows."""
    candidates = [
        shutil.which("msedge"),
        os.environ.get("PROGRAMFILES(X86)", "")
        + r"\Microsoft\Edge\Application\msedge.exe",
        os.environ.get("PROGRAMFILES", "")
        + r"\Microsoft\Edge\Application\msedge.exe",
        os.environ.get("LOCALAPPDATA", "")
        + r"\Microsoft\Edge\Application\msedge.exe",
        shutil.which("chrome"),
        os.environ.get("PROGRAMFILES", "")
        + r"\Google\Chrome\Application\chrome.exe",
        os.environ.get("PROGRAMFILES(X86)", "")
        + r"\Google\Chrome\Application\chrome.exe",
        os.environ.get("LOCALAPPDATA", "")
        + r"\Google\Chrome\Application\chrome.exe",
    ]
    for candidate in candidates:
        if candidate and Path(candidate).is_file():
            return Path(candidate)
    return None


def _find_server_python() -> list[str]:
    """Find a Python installation containing the server dependencies."""
    candidates: list[list[str]] = [
        [str(ROOT / ".venv" / "Scripts" / "python.exe")],
        [str(ROOT / "venv" / "Scripts" / "python.exe")],
        [sys.executable],
    ]

    py_launcher = shutil.which("py")
    if py_launcher:
        try:
            result = subprocess.run(
                [py_launcher, "-0p"],
                capture_output=True,
                text=True,
                timeout=10,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
            versions = re.findall(
                r"-V:([0-9]+(?:\.[0-9]+)?)", result.stdout + result.stderr
            )
            candidates.extend([py_launcher, f"-{version}"] for version in versions)
        except (OSError, subprocess.SubprocessError):
            pass

    checked: set[tuple[str, ...]] = set()
    for command in candidates:
        key = tuple(command)
        if key in checked:
            continue
        checked.add(key)
        if len(command) == 1 and not Path(command[0]).is_file():
            continue
        try:
            result = subprocess.run(
                command + ["-c", "import flask, flask_sock"],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                timeout=15,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
            if result.returncode == 0:
                return command
        except (OSError, subprocess.SubprocessError):
            continue

    raise RuntimeError(
        "No Python installation with the server dependencies was found.\n\n"
        "Install them with:\n"
        "py -3.11 -m pip install -r requirements.txt"
    )


def _start_server(log_file, python_command: list[str]) -> subprocess.Popen:
    env = os.environ.copy()
    env["PYTHONUNBUFFERED"] = "1"
    creation_flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    return subprocess.Popen(
        python_command
        + [str(ROOT / "control_server.py"), "--port", "8081"],
        cwd=ROOT,
        env=env,
        stdout=log_file,
        stderr=subprocess.STDOUT,
        creationflags=creation_flags,
    )


def _wait_for_server(server: subprocess.Popen) -> None:
    deadline = time.monotonic() + STARTUP_TIMEOUT_SECONDS
    while time.monotonic() < deadline:
        if server.poll() is not None:
            raise RuntimeError(
                f"The control server stopped during startup (exit code "
                f"{server.returncode})."
            )
        try:
            with urllib.request.urlopen(APP_URL, timeout=0.5) as response:
                if response.status < 500:
                    return
        except (OSError, urllib.error.URLError):
            pass
        time.sleep(0.2)
    raise RuntimeError(
        f"The control server did not become ready at {APP_URL} within "
        f"{STARTUP_TIMEOUT_SECONDS} seconds."
    )


def _start_browser(browser: Path, profile_dir: Path) -> subprocess.Popen:
    return subprocess.Popen(
        [
            str(browser),
            f"--app={APP_URL}",
            f"--user-data-dir={profile_dir}",
            "--no-first-run",
            "--no-default-browser-check",
            "--disable-background-mode",
            "--disable-features=msEdgeFirstRunExperience",
            "--window-size=1500,950",
        ],
        cwd=ROOT,
    )


def _stop_server(server: subprocess.Popen | None) -> None:
    if server is None or server.poll() is not None:
        return

    # This endpoint runs the same hardware cleanup used by Ctrl+C.
    try:
        request = urllib.request.Request(SHUTDOWN_URL, method="POST")
        with urllib.request.urlopen(request, timeout=3):
            pass
    except (OSError, urllib.error.URLError):
        # If HTTP is already unavailable, fall through to process cleanup.
        pass

    try:
        server.wait(timeout=SHUTDOWN_TIMEOUT_SECONDS)
        return
    except subprocess.TimeoutExpired:
        server.terminate()

    try:
        server.wait(timeout=5)
    except subprocess.TimeoutExpired:
        server.kill()
        server.wait(timeout=5)


def _read_log(log_path: Path, log_file) -> str:
    try:
        log_file.flush()
        text = log_path.read_text(encoding="utf-8", errors="replace").strip()
        if text:
            return f"\n\nServer output:\n{text[-5000:]}"
    except OSError:
        pass
    return ""


def main() -> None:
    server = None
    browser_profile = None
    failed = False
    log_path = Path(tempfile.gettempdir()) / f"mosaic-control-{os.getpid()}.log"

    with log_path.open("w", encoding="utf-8") as log_file:
        try:
            browser = _find_browser()
            if browser is None:
                raise RuntimeError(
                    "Microsoft Edge or Google Chrome could not be found."
                )

            server_python = _find_server_python()
            server = _start_server(log_file, server_python)
            _wait_for_server(server)

            browser_profile = Path(
                tempfile.mkdtemp(prefix="mosaic-control-browser-")
            )
            browser_process = _start_browser(browser, browser_profile)
            browser_process.wait()
        except Exception as exc:
            failed = True
            _show_error(str(exc) + _read_log(log_path, log_file))
        finally:
            _stop_server(server)

    if browser_profile is not None:
        shutil.rmtree(browser_profile, ignore_errors=True)
    if not failed:
        try:
            log_path.unlink()
        except OSError:
            pass


if __name__ == "__main__":
    main()
