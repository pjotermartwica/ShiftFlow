"""
ShiftFlow Bootstrap Loader
--------------------------
Punkt wejścia skompilowanego .exe.
Przy każdym uruchomieniu:
  1. Sprawdza GitHub Raw version.txt (timeout 5 s).
  2. Jeśli jest nowsza wersja: pobiera main.py (~100 KB) do
     %LOCALAPPDATA%/ShiftFlow/ bez przebudowy .exe.
  3. Uruchamia lokalny main.py przez exec() — korzysta z bundlowanego
     Pythona i PySide6, więc wszystkie zależności są wypełnione.
  4. Awaryjnie: jeśli brak sieci lub niespójny plik, uruchamia
     main.py bundlowany wewnątrz .exe.

Przebudowy .exe wymagają TYLKO zmiany zależności (nowy PySide6, nowe
paczki). Każdą zmianę kodu wystarczy wgrać na GitHub.
"""

import sys
import os
import re as _re

# ---------------------------------------------------------------------------
# Konfiguracja
# ---------------------------------------------------------------------------
_GITHUB_RAW = "https://raw.githubusercontent.com/pjotermartwica/ShiftFlow/main"
_APP_DIR = os.path.join(
    os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "ShiftFlow"
)
_LOCAL_SCRIPT  = os.path.join(_APP_DIR, "main.py")
_LOCAL_VER     = os.path.join(_APP_DIR, "version.txt")
_LOCAL_CFG     = os.path.join(_APP_DIR, "config.json")
_LOG_FILE      = os.path.join(_APP_DIR, "shiftflow_log.txt")


# ---------------------------------------------------------------------------
# Logging do pliku (działa PRZED inicjalizacją Qt)
# ---------------------------------------------------------------------------

def _log(msg: str) -> None:
    """Zapisz linię do shiftflow_log.txt. Nigdy nie rzuca wyjątków."""
    try:
        os.makedirs(_APP_DIR, exist_ok=True)
        import datetime
        ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(_LOG_FILE, "a", encoding="utf-8") as f:
            f.write(f"[{ts}] {msg}\n")
    except Exception:
        pass


def _fatal(title: str, msg: str) -> None:
    """Pokaż błąd krytyczny (tkinter, bo PySide6 może nie działać) i zakończ."""
    _log(f"FATAL: {title}: {msg}")
    try:
        import tkinter as tk
        from tkinter import messagebox
        r = tk.Tk()
        r.withdraw()
        messagebox.showerror(title, f"{msg}\n\nSzczegóły w:\n{_LOG_FILE}")
        r.destroy()
    except Exception:
        pass
    sys.exit(1)


# ---------------------------------------------------------------------------
# Narzędzia
# ---------------------------------------------------------------------------

def _ver_parts(v: str) -> list[int]:
    return [int(x) for x in _re.findall(r"\d+", v)]


def _read_local_version() -> str:
    """Odczytaj wersję z version.txt, a jeśli brak — ze wbudowanego main.py."""
    if os.path.isfile(_LOCAL_VER):
        try:
            return open(_LOCAL_VER, encoding="utf-8").read().strip()
        except Exception:
            pass
    if os.path.isfile(_LOCAL_SCRIPT):
        try:
            for line in open(_LOCAL_SCRIPT, encoding="utf-8"):
                if line.startswith("__version__"):
                    m = _re.search(r'"([^"]+)"', line)
                    if m:
                        return m.group(1)
        except Exception:
            pass
    # Bundlowany fallback: wyciągnij z MEIPASS\main.py
    if getattr(sys, "frozen", False):
        bundled = os.path.join(sys._MEIPASS, "main.py")  # type: ignore[attr-defined]
        try:
            for line in open(bundled, encoding="utf-8"):
                if line.startswith("__version__"):
                    m = _re.search(r'"([^"]+)"', line)
                    if m:
                        return m.group(1)
        except Exception:
            pass
    return "0.0.0"


def _fetch_remote_version() -> str | None:
    import urllib.request
    try:
        req = urllib.request.Request(
            f"{_GITHUB_RAW}/version.txt",
            headers={"User-Agent": "ShiftFlow-Bootstrap"},
        )
        with urllib.request.urlopen(req, timeout=5) as r:
            return r.read().decode("utf-8-sig").strip()
    except Exception:
        return None


def _download_script(remote_ver: str) -> bool:
    """Pobierz main.py z GitHub Raw i zapisz w APP_DIR."""
    import urllib.request
    try:
        req = urllib.request.Request(
            f"{_GITHUB_RAW}/main.py",
            headers={"User-Agent": "ShiftFlow-Bootstrap"},
        )
        with urllib.request.urlopen(req, timeout=30) as r:
            data = r.read()
        os.makedirs(_APP_DIR, exist_ok=True)
        tmp = _LOCAL_SCRIPT + ".tmp"
        with open(tmp, "wb") as f:
            f.write(data)
        os.replace(tmp, _LOCAL_SCRIPT)
        with open(_LOCAL_VER, "w", encoding="utf-8") as f:
            f.write(remote_ver)
        return True
    except Exception:
        return False


def _ensure_default_files():
    """
    Przy pierwszym uruchomieniu skopiuj config.json i harmonogram.json
    z bundla do APP_DIR, jeśli jeszcze tam nie ma.
    """
    if not getattr(sys, "frozen", False):
        return
    meipass = sys._MEIPASS  # type: ignore[attr-defined]
    for name in ("config.json", "harmonogram.json"):
        dst = os.path.join(_APP_DIR, name)
        src = os.path.join(meipass, name)
        if not os.path.isfile(dst) and os.path.isfile(src):
            try:
                import shutil
                shutil.copy2(src, dst)
            except Exception:
                pass


def _get_script_path() -> str:
    """Zwróć ścieżkę do main.py do wykonania (zewnętrzny lub bundlowany)."""
    if os.path.isfile(_LOCAL_SCRIPT):
        return _LOCAL_SCRIPT
    if getattr(sys, "frozen", False):
        return os.path.join(sys._MEIPASS, "main.py")  # type: ignore[attr-defined]
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "main.py")


# ---------------------------------------------------------------------------
# Punkt wejścia
# ---------------------------------------------------------------------------

def main() -> None:
    os.makedirs(_APP_DIR, exist_ok=True)
    _log("=" * 60)
    _log(f"Bootstrap start  |  Python {sys.version}  |  frozen={getattr(sys,'frozen',False)}")
    _log(f"APP_DIR: {_APP_DIR}")

    _ensure_default_files()

    # Ustaw CWD na APP_DIR: harmonogram.json, config.json, brain_memory.json
    # będą zapisywane obok siebie, a nie obok .exe (która może być gdzie indziej).
    os.chdir(_APP_DIR)
    _log(f"CWD ustawione na: {_APP_DIR}")

    # --- Cicha aktualizacja skryptu ---
    try:
        remote_ver = _fetch_remote_version()
        if remote_ver:
            local_ver = _read_local_version()
            _log(f"Wersja: lokalna={local_ver}  zdalna={remote_ver}")
            if _ver_parts(remote_ver) > _ver_parts(local_ver):
                ok = _download_script(remote_ver)
                _log(f"Pobieranie skryptu: {'OK' if ok else 'BLAD'}")
        else:
            _log("Brak sieci lub GitHub niedostepny — pomijam aktualizacje")
    except Exception as e:
        _log(f"Blad aktualizacji (niekrytyczny): {e}")

    # --- Uruchom main.py w tym samym procesie ---
    script = _get_script_path()
    _log(f"Uruchamiam skrypt: {script}  |  istnieje={os.path.isfile(script)}")

    if not os.path.isfile(script):
        _fatal(
            "ShiftFlow — brak skryptu",
            f"Nie znaleziono main.py:\n{script}\n\nZainstaluj ponownie aplikację.",
        )

    globs: dict = {
        "__file__": script,
        "__name__": "__main__",
        "__spec__": None,
        "__builtins__": __builtins__,
    }
    try:
        with open(script, "r", encoding="utf-8") as f:
            src = f.read()
        _log("exec() main.py — start")
        exec(compile(src, script, "exec"), globs)  # noqa: S102
        _log("exec() main.py — zakonczone normalnie")
    except SystemExit:
        # Normalne wyjście przez sys.exit() — nie loguj jako błąd
        raise
    except Exception as exc:
        import traceback
        tb = traceback.format_exc()
        _log(f"WYJATEK w main.py:\n{tb}")
        _fatal(
            "ShiftFlow — błąd krytyczny",
            f"{exc}\n\nSzczegóły w logu:\n{_LOG_FILE}",
        )


if __name__ == "__main__":
    main()
