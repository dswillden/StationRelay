"""
StationRelay
============
Submit lot numbers into a shared Excel print queue.
Operators mark lots as done — first come, first served.

Excel column layout (starting at configured column, default A):
    Col A  -- Lot Number
    Col B  -- Submitted By  (Windows display name)
    Col C  -- Submitted At  (YYYY-MM-DD HH:MM:SS)
    Col D  -- Printed At    (YYYY-MM-DD HH:MM:SS, written when marked done)

Requirements:
    pip install pywin32

Run:
    python app.py

Build single EXE:
    pip install pyinstaller
    pyinstaller --onefile --windowed --name StationRelay app.py
"""

import json
import os
import threading
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, ttk

# ---------------------------------------------------------------------------
# Config
#
# Two-layer config system:
#   1. BAKED_DEFAULTS  — compiled into the EXE; provides the shared Excel path
#                        and sensible defaults for every new user.
#   2. User override   — %APPDATA%\StationRelay\config.json, created on first
#                        save.  Merged on top of BAKED_DEFAULTS so users can
#                        override any key without losing new defaults added in
#                        future EXE releases.
# ---------------------------------------------------------------------------

# These values are baked in at build time.  Edit before running PyInstaller.
#
# excel_path accepts two formats:
#   1. Absolute path  — used as-is (e.g. user manually picked a file in Settings)
#   2. OneDrive-relative path — a path that does NOT start with a drive letter or
#      UNC prefix.  At runtime _resolve_excel_path() prepends the machine's
#      OneDrive root (%OneDriveCommercial% → %OneDrive% → fallback).
#      Example:  "AMER-LS-eBR Implementation-Bottling Team - Documents/BTL_QuickShare.xlsx"
#
BAKED_DEFAULTS: dict = {
    "excel_path":    r"AMER-LS-eBR Implementation-Bottling Team - Documents/BTL_QuickShare.xlsx",
    "sheet_name":    "Sheet1",
    "column":        "A",
    "theme":         "light",
    "always_on_top": False,
}

# Fallback if AppData is unavailable for some reason
_FALLBACK_CONFIG_FILE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "config.json"
)


def _user_config_path() -> str:
    """Return %APPDATA%\\StationRelay\\config.json, creating the folder if needed."""
    appdata = os.environ.get("APPDATA") or os.path.expanduser("~")
    folder = os.path.join(appdata, "StationRelay")
    try:
        os.makedirs(folder, exist_ok=True)
        return os.path.join(folder, "config.json")
    except Exception:
        return _FALLBACK_CONFIG_FILE


def _onedrive_root() -> str:
    """Return this machine's OneDrive root from env vars, or empty string."""
    return (
        os.environ.get("OneDriveCommercial")
        or os.environ.get("OneDrive")
        or ""
    )


def _resolve_excel_path(raw: str) -> str:
    """
    Normalise *raw* to an absolute path that is valid on this machine.

    Three cases are handled:

    1. OneDrive-relative path (no drive letter prefix, e.g. stored in
       BAKED_DEFAULTS):
           "AMER-LS-eBR .../BTL_QuickShare.xlsx"
       → prepend this machine's OneDrive root.

    2. Absolute path whose OneDrive root belongs to a *different* user
       (e.g. a config.json written on dev machine with username 2020303
       but loaded on a machine with a different username):
           "C:/Users/2020303/OneDrive - Revvity/AMER-LS-.../BTL_QuickShare.xlsx"
       → strip everything up to and including the OneDrive root folder,
         keep the relative tail, re-prepend this machine's root.

    3. Absolute path that is already valid on this machine (user manually
       chose a file via Settings browse dialog):
       → return unchanged.

    Priority of OneDrive roots: %OneDriveCommercial% > %OneDrive%.
    If no env var is set, the path is returned as-is (best effort).
    """
    import re

    raw = raw.replace("\\", "/").strip()
    my_root = _onedrive_root()

    # Helper: normalise a root for comparison
    def _norm(p: str) -> str:
        return os.path.normcase(os.path.normpath(p))

    # Case 1 — not absolute at all → treat as OneDrive-relative
    if not re.match(r'^[A-Za-z]:[/\\]', raw) and not raw.startswith("//") and not raw.startswith("\\\\"):
        if not my_root:
            import sys
            print(
                "[StationRelay] WARNING: OneDriveCommercial / OneDrive env vars not set. "
                f"Cannot resolve relative excel_path: {raw!r}",
                file=sys.stderr,
            )
            return raw
        return os.path.normpath(os.path.join(my_root, raw))

    # Case 2 — absolute path, but OneDrive root doesn't match this machine.
    # Detect by looking for a "OneDrive" folder component in the path.
    if my_root:
        norm_raw = _norm(raw)
        norm_my_root = _norm(my_root)

        # Already points inside this machine's OneDrive → fine as-is (Case 3)
        if norm_raw.startswith(norm_my_root + os.sep) or norm_raw.startswith(norm_my_root.lower() + os.sep):
            return os.path.normpath(raw)

        # Look for any "OneDrive*" folder in the path and re-root from there
        # e.g. C:/Users/2020303/OneDrive - Revvity/Folder/File.xlsx
        #   → tail = Folder/File.xlsx
        #   → result = <my_root>/Folder/File.xlsx
        onedrive_re = re.compile(r'^(.*?[/\\])(OneDrive[^/\\]*)[/\\](.+)$', re.IGNORECASE)
        m = onedrive_re.match(raw)
        if m:
            tail = m.group(3)  # everything after the OneDrive root folder
            return os.path.normpath(os.path.join(my_root, tail))

    # Case 3 — absolute, not a OneDrive path (or no env var) → return as-is
    return os.path.normpath(raw)


def load_config() -> dict:
    """
    Return merged config: BAKED_DEFAULTS <- user overrides.
    The user file only needs to store keys the user has changed.
    excel_path is resolved to an absolute path at load time.
    """
    cfg = dict(BAKED_DEFAULTS)
    user_file = _user_config_path()
    if os.path.exists(user_file):
        try:
            with open(user_file, "r", encoding="utf-8") as f:
                user_data = json.load(f)
            cfg.update(user_data)
        except Exception:
            pass
    # Resolve relative OneDrive paths to absolute on this machine
    cfg["excel_path"] = _resolve_excel_path(cfg.get("excel_path", ""))
    return cfg


def save_config(cfg: dict) -> None:
    """Write only the user-override file (AppData), never the baked defaults."""
    user_file = _user_config_path()
    with open(user_file, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2)


# ---------------------------------------------------------------------------
# Theme tokens
# ---------------------------------------------------------------------------

THEMES: dict[str, dict] = {
    "light": {
        "bg":           "#f4f4f6",
        "surface":      "#ffffff",
        "surface2":     "#ebebef",
        "border":       "#d0d0d8",
        "accent":       "#5b4ef8",
        "accent_hover": "#4a3de0",
        "accent_fg":    "#ffffff",
        "text":         "#111118",
        "muted":        "#7070a0",
        "success":      "#1a7a40",
        "error":        "#c0302a",
        "warning":      "#b06010",
        "done_bg":      "#f0f0f4",
        "done_fg":      "#aaaacc",
        "row_alt":      "#f9f9fc",
        "tag_pending":  "#e8e8f8",
        "tag_done":     "#e8f4ee",
    },
    "dark": {
        "bg":           "#1a1a28",
        "surface":      "#24243a",
        "surface2":     "#2e2e46",
        "border":       "#3a3a58",
        "accent":       "#7c6af7",
        "accent_hover": "#6a59e0",
        "accent_fg":    "#ffffff",
        "text":         "#e2e2f2",
        "muted":        "#8080aa",
        "success":      "#4caf7d",
        "error":        "#f06060",
        "warning":      "#f0a060",
        "done_bg":      "#202030",
        "done_fg":      "#505070",
        "row_alt":      "#222238",
        "tag_pending":  "#33335a",
        "tag_done":     "#1e3a28",
    },
}

# Active theme dict — mutated in-place on toggle so all refs stay valid
T: dict = dict(THEMES["light"])

F = {
    "title":   ("Segoe UI", 15, "bold"),
    "heading": ("Segoe UI", 11, "bold"),
    "normal":  ("Segoe UI", 10),
    "small":   ("Segoe UI", 9),
    "mono":    ("Consolas", 12, "bold"),
    "mono_sm": ("Consolas", 9),
    "btn":     ("Segoe UI", 10, "bold"),
    "btn_sm":  ("Segoe UI", 9),
}


def apply_theme(name: str) -> None:
    """Mutate T in-place to the chosen theme palette."""
    T.update(THEMES[name])


# ---------------------------------------------------------------------------
# Button factory
# ---------------------------------------------------------------------------

def styled_button(parent, text, command, style="accent", font_key="btn", **kwargs):
    colours = {
        "accent":  lambda: (T["accent"],   T["accent_fg"], T["accent_hover"]),
        "muted":   lambda: (T["surface2"], T["muted"],     T["border"]),
        "success": lambda: ("#1e5a34",     T["success"],   "#25723f"),
        "copy":    lambda: (T["surface2"], T["accent"],    T["border"]),
        "pin":     lambda: (T["tag_pending"], T["accent"], T["border"]),
    }
    bg, fg, hover = colours.get(style, colours["muted"])()
    return tk.Button(
        parent, text=text, command=command,
        font=F[font_key], relief=tk.FLAT, cursor="hand2",
        bg=bg, fg=fg, activebackground=hover, activeforeground=fg,
        bd=0,
        padx=kwargs.pop("padx", 14),
        pady=kwargs.pop("pady", 6),
        **kwargs,
    )


# ---------------------------------------------------------------------------
# Excel COM helpers
# ---------------------------------------------------------------------------

def _format_exc(exc: Exception) -> str:
    """
    Return a human-readable error string from *exc*.

    COM errors (pywintypes.com_error) carry rich info in .args but often
    have str(exc) == 'None'.  This extracts the real description.
    """
    # pywintypes.com_error → args is (hresult, source, description, ...)
    if hasattr(exc, "args") and isinstance(exc.args, tuple) and len(exc.args) >= 3:
        desc = exc.args[2]
        src  = exc.args[1] if len(exc.args) > 1 else ""
        if desc:
            return f"{src}: {desc}" if src else str(desc)

    s = str(exc)
    if s and s != "None":
        return s

    return f"{type(exc).__name__}: (no details — see Windows Event Viewer)"


def _kill_excel_for_file(file_path: str) -> None:
    """
    If any visible or background Excel.exe process has *file_path* open,
    close that workbook (or kill the process as a last resort).

    This prevents COM DispatchEx from hanging because another Excel instance
    already holds a write-lock on the file.  Safe to call even if no Excel
    is running.
    """
    import subprocess
    target = os.path.normcase(os.path.abspath(file_path))
    target_name = os.path.basename(target).lower()

    # Strategy 1: Try COM — connect to each running Excel and close the wb
    try:
        import pythoncom  # type: ignore
        import win32com.client  # type: ignore

        ctx = pythoncom.CreateBindCtx(0)  # type: ignore[call-arg]
        rot = pythoncom.GetRunningObjectTable()

        for moniker in rot:  # type: ignore[union-attr]
            try:
                display = moniker.GetDisplayName(ctx, None)
            except Exception:
                continue
            if target_name in display.lower():
                try:
                    obj = rot.GetObject(moniker)
                    wb = win32com.client.Dispatch(
                        obj.QueryInterface(pythoncom.IID_IDispatch)
                    )
                    wb.Close(SaveChanges=False)
                except Exception:
                    pass
    except Exception:
        pass

    # Strategy 2: Brute-force — kill all EXCEL.EXE.
    # Only do this if the file is STILL locked (test by trying to open it).
    try:
        with open(target, "r+b"):
            return  # file is not locked, we're good
    except PermissionError:
        pass
    except Exception:
        return  # file missing or other issue — nothing to kill

    # File is locked — kill Excel processes
    try:
        subprocess.run(
            ["taskkill", "/F", "/IM", "EXCEL.EXE"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            timeout=5,
        )
        import time
        time.sleep(0.5)  # give OS a moment to release handles
    except Exception:
        pass


def get_display_name() -> str:
    try:
        import win32api  # type: ignore
        name = win32api.GetUserNameEx(3)
        if name and name.strip():
            return name.strip()
    except Exception:
        pass
    return os.environ.get("USERNAME") or os.getlogin()


def _open_excel_hidden(excel_path: str):
    """Open *excel_path* in a fresh hidden Excel instance.

    Kills any existing Excel process holding the file first to avoid
    write-lock conflicts.
    """
    _kill_excel_for_file(excel_path)
    import win32com.client  # type: ignore
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = excel.Workbooks.Open(excel_path)
    return excel, wb


def _get_sheet(wb, sheet_name: str):
    for sh in wb.Sheets:
        if sh.Name.strip().lower() == sheet_name.strip().lower():
            return sh
    raise ValueError(
        f"Sheet '{sheet_name}' not found.\n"
        f"Available: {[s.Name for s in wb.Sheets]}"
    )


def _col_letter_to_index(col: str) -> int:
    col = col.upper().strip()
    result = 0
    for ch in col:
        result = result * 26 + (ord(ch) - ord("A") + 1)
    return result


def _next_empty_row(sheet, col_index: int) -> int:
    used_rows = sheet.UsedRange.Rows.Count
    last_row = 0
    for row in range(1, used_rows + 2):
        cell = sheet.Cells(row, col_index)
        if cell.Value is not None and str(cell.Value).strip() != "":
            last_row = row
    return last_row + 1


def _nudge_onedrive(file_path: str) -> None:
    try:
        onedrive_root = (os.environ.get("OneDriveCommercial")
                         or os.environ.get("OneDrive", ""))
        norm_path = os.path.normcase(os.path.abspath(file_path))
        norm_root = (os.path.normcase(os.path.abspath(onedrive_root))
                     if onedrive_root else "")
        if norm_root and not norm_path.startswith(norm_root):
            return
        now = datetime.now().timestamp()
        os.utime(file_path, (now, now))
    except Exception:
        pass


def _com_session(func):
    """
    Decorator that wraps a function in a COM-initialized thread context.

    Also installs a RetryMessageFilter for the duration of the call so that
    RPC_E_CALL_REJECTED (0x80010001 — "Call was rejected by callee") errors
    from a busy Excel instance are automatically retried by the COM runtime
    rather than immediately raised.
    """
    def wrapper(*args, **kwargs):
        import pythoncom  # type: ignore
        pythoncom.CoInitialize()
        _filter = _RetryMessageFilter()
        _filter.register()
        try:
            return func(*args, **kwargs)
        finally:
            _filter.unregister()
            pythoncom.CoUninitialize()
    return wrapper


class _RetryMessageFilter:
    """
    COM IMessageFilter implementation that tells the RPC runtime to
    retry calls automatically when Excel (or any COM server) is busy.

    Without this, a busy Excel throws RPC_E_CALL_REJECTED immediately.
    With this, the runtime retries for up to RETRY_MS milliseconds.
    """

    RETRY_MS = 10_000  # retry for up to 10 seconds

    def register(self) -> None:
        try:
            import pythoncom      # type: ignore
            import win32com.server.util  # type: ignore

            _retry_ms = self.RETRY_MS

            class _Filter:
                def HandleInComingCall(self, dwCallType, hTaskCaller,
                                       dwTickCount, lpInterfaceInfo):
                    return 0  # SERVERCALL_ISHANDLED

                def RetryRejectedCall(self, hTaskCallee, dwTickCount,
                                      dwRejectType):
                    if dwTickCount < _retry_ms:
                        return 100  # retry after 100 ms
                    return -1  # cancel

                def MessagePending(self, hTaskCallee, dwTickCount,
                                   dwPendingType):
                    return 2  # PENDINGMSG_WAITNOPROCESS

            com_filter = win32com.server.util.wrap(  # type: ignore[attr-defined]
                _Filter(),
                pythoncom.IID_IMessageFilter,        # type: ignore[attr-defined]
            )
            self._prev = pythoncom.CoRegisterMessageFilter(com_filter)  # type: ignore[attr-defined]
        except Exception:
            self._prev = None

    def unregister(self) -> None:
        try:
            if self._prev is not None:
                import pythoncom  # type: ignore
                pythoncom.CoRegisterMessageFilter(self._prev)  # type: ignore[attr-defined]
        except Exception:
            pass


def _with_retry(max_attempts: int = 3, delay: float = 1.5):
    """
    Decorator that retries a function on exception up to *max_attempts* times,
    waiting *delay* seconds between attempts.  Designed for COM write functions
    where OneDrive file-lock collisions cause transient errors.
    Re-raises the last exception if all attempts fail.
    """
    import functools
    import time

    def decorator(func):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            last_exc: Exception | None = None
            for attempt in range(1, max_attempts + 1):
                try:
                    return func(*args, **kwargs)
                except Exception as exc:
                    last_exc = exc
                    if attempt < max_attempts:
                        time.sleep(delay)
            raise last_exc  # type: ignore[misc]
        return wrapper
    return decorator


@_with_retry(max_attempts=3, delay=2.5)
@_com_session
def append_to_excel(excel_path: str, sheet_name: str, column: str, lot: str) -> None:
    display_name = get_display_name()
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    excel = wb = None
    try:
        excel, wb = _open_excel_hidden(excel_path)
        sheet = _get_sheet(wb, sheet_name)
        col = _col_letter_to_index(column)
        row = _next_empty_row(sheet, col)
        sheet.Cells(row, col).Value     = lot
        sheet.Cells(row, col + 1).Value = display_name
        sheet.Cells(row, col + 2).Value = timestamp
        wb.Save()
    finally:
        if wb:
            try: wb.Close(SaveChanges=False)
            except Exception: pass
        if excel:
            try: excel.Quit()
            except Exception: pass
    _nudge_onedrive(excel_path)


@_com_session
def read_from_excel(excel_path: str, sheet_name: str, column: str) -> list:
    """
    Read all rows from the Excel queue sheet.
    Raises on failure — callers must handle the exception and surface it to the user.
    """
    excel = wb = None
    rows = []
    try:
        excel, wb = _open_excel_hidden(excel_path)
        sheet = _get_sheet(wb, sheet_name)
        col = _col_letter_to_index(column)
        used = sheet.UsedRange.Rows.Count
        for r in range(1, used + 1):
            val = sheet.Cells(r, col).Value
            if val is None or str(val).strip() == "":
                continue
            rows.append({
                "row":          r,
                "lot":          str(val).strip(),
                "name":         str(sheet.Cells(r, col + 1).Value or ""),
                "submitted_at": str(sheet.Cells(r, col + 2).Value or ""),
                "printed_at":   str(sheet.Cells(r, col + 3).Value or ""),
            })
    finally:
        if wb:
            try: wb.Close(SaveChanges=False)
            except Exception: pass
        if excel:
            try: excel.Quit()
            except Exception: pass
    return rows


@_with_retry(max_attempts=3, delay=2.5)
@_com_session
def mark_done_in_excel(excel_path: str, sheet_name: str,
                        column: str, row_number: int) -> None:
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    excel = wb = None
    try:
        excel, wb = _open_excel_hidden(excel_path)
        sheet = _get_sheet(wb, sheet_name)
        col = _col_letter_to_index(column)
        sheet.Cells(row_number, col + 3).Value = timestamp
        wb.Save()
    finally:
        if wb:
            try: wb.Close(SaveChanges=False)
            except Exception: pass
        if excel:
            try: excel.Quit()
            except Exception: pass
    _nudge_onedrive(excel_path)


@_with_retry(max_attempts=3, delay=2.5)
@_com_session
def unmark_done_in_excel(excel_path: str, sheet_name: str,
                          column: str, row_number: int) -> None:
    excel = wb = None
    try:
        excel, wb = _open_excel_hidden(excel_path)
        sheet = _get_sheet(wb, sheet_name)
        col = _col_letter_to_index(column)
        sheet.Cells(row_number, col + 3).Value = None
        wb.Save()
    finally:
        if wb:
            try: wb.Close(SaveChanges=False)
            except Exception: pass
        if excel:
            try: excel.Quit()
            except Exception: pass
    _nudge_onedrive(excel_path)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _parse_dt(s: str) -> datetime:
    """Parse a submitted_at string to datetime; return epoch on failure."""
    try:
        return datetime.strptime(s.strip(), "%Y-%m-%d %H:%M:%S")
    except Exception:
        return datetime(1970, 1, 1)


def _sort_rows(rows: list) -> list:
    """Sort rows oldest-first by submitted_at (exact second precision)."""
    return sorted(rows, key=lambda r: _parse_dt(r["submitted_at"]))


# ---------------------------------------------------------------------------
# Settings dialog
# ---------------------------------------------------------------------------

class SettingsDialog(tk.Toplevel):
    def __init__(self, parent: "StationRelayApp"):
        super().__init__(parent)
        self.parent_app = parent
        self.title("Settings")
        self.resizable(False, False)
        self.configure(bg=T["bg"])
        self.grab_set()
        self.update_idletasks()
        px = parent.winfo_x() + parent.winfo_width()  // 2
        py = parent.winfo_y() + parent.winfo_height() // 2
        self.geometry(f"440x280+{px - 220}+{py - 140}")
        self._build()

    def _lbl(self, parent, text):
        return tk.Label(parent, text=text, bg=T["bg"], fg=T["muted"], font=F["small"])

    def _entry(self, parent, textvariable, width=36):
        return tk.Entry(
            parent, textvariable=textvariable, width=width,
            font=F["normal"], relief=tk.FLAT,
            bg=T["surface2"], fg=T["text"],
            insertbackground=T["text"],
            highlightthickness=1,
            highlightbackground=T["border"],
            highlightcolor=T["accent"],
        )

    def _build(self):
        tk.Label(self, text="Settings", bg=T["bg"], fg=T["text"],
                 font=F["heading"]).pack(anchor="w", padx=24, pady=(20, 4))
        tk.Frame(self, height=1, bg=T["border"]).pack(fill=tk.X, padx=24)

        self._lbl(self, "Excel file path").pack(anchor="w", padx=24, pady=(10, 2))
        path_row = tk.Frame(self, bg=T["bg"])
        path_row.pack(fill=tk.X, padx=24)

        self._path_var = tk.StringVar(
            value=self.parent_app.config_data.get("excel_path", ""))
        pe = self._entry(path_row, self._path_var, width=32)
        pe.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=5)
        styled_button(path_row, "…", self._browse,
                      style="muted", padx=10, pady=5).pack(side=tk.LEFT, padx=(6, 0))

        sc = tk.Frame(self, bg=T["bg"])
        sc.pack(fill=tk.X, padx=24, pady=(12, 0))
        self._lbl(sc, "Sheet name").grid(row=0, column=0, sticky="w")
        self._lbl(sc, "Start column").grid(row=0, column=2, sticky="w", padx=(20, 0))

        self._sheet_var = tk.StringVar(
            value=self.parent_app.config_data.get("sheet_name", "Sheet1"))
        self._col_var = tk.StringVar(
            value=self.parent_app.config_data.get("column", "A"))

        self._entry(sc, self._sheet_var, width=18).grid(row=1, column=0, ipady=5, sticky="w")
        self._entry(sc, self._col_var,   width=6 ).grid(row=1, column=2, ipady=5,
                                                          sticky="w", padx=(20, 0))

        btn_row = tk.Frame(self, bg=T["bg"])
        btn_row.pack(fill=tk.X, padx=24, pady=(20, 16))
        styled_button(btn_row, "Cancel", self.destroy, style="muted").pack(side=tk.RIGHT, padx=(8, 0))
        styled_button(btn_row, "Save",   self._save,  style="accent").pack(side=tk.RIGHT)

    def _browse(self):
        initial = self._path_var.get()
        initial_dir = os.path.dirname(initial) if initial else os.path.expanduser("~")
        path = filedialog.askopenfilename(
            title="Select Excel file",
            initialdir=initial_dir,
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")],
        )
        if path:
            self._path_var.set(path)

    def _save(self):
        self.parent_app.config_data["excel_path"] = self._path_var.get().strip()
        self.parent_app.config_data["sheet_name"] = self._sheet_var.get().strip() or "Sheet1"
        self.parent_app.config_data["column"]     = self._col_var.get().strip().upper() or "A"
        save_config(self.parent_app.config_data)
        self.destroy()


# ---------------------------------------------------------------------------
# Queue row widget
# ---------------------------------------------------------------------------

class QueueRow(tk.Frame):
    def __init__(self, parent, row_data: dict, on_toggle, is_done: bool = False,
                 index: int = 0):
        bg = T["done_bg"] if is_done else (T["surface"] if index % 2 == 0 else T["row_alt"])
        super().__init__(parent, bg=bg)
        self._data = row_data
        self._build(bg, is_done, on_toggle)

    def _build(self, bg: str, is_done: bool, on_toggle):
        fg      = T["done_fg"] if is_done else T["text"]
        fg_mono = T["done_fg"] if is_done else T["accent"]

        # Left pad
        tk.Frame(self, width=12, bg=bg).pack(side=tk.LEFT)

        # Queue position badge
        badge_bg = T["tag_done"] if is_done else T["tag_pending"]
        badge_fg = T["done_fg"]  if is_done else T["muted"]
        badge = tk.Frame(self, bg=bg, width=28)
        badge.pack(side=tk.LEFT, pady=10)
        badge.pack_propagate(False)
        tk.Label(badge, text=str(self._data.get("queue_pos", "")),
                 bg=badge_bg, fg=badge_fg, font=F["mono_sm"],
                 padx=4, pady=2).pack(expand=True)

        tk.Frame(self, width=10, bg=bg).pack(side=tk.LEFT)

        # Lot number
        lot_font = F["mono"] if not is_done else ("Consolas", 12)
        tk.Label(self, text=self._data["lot"],
                 bg=bg, fg=fg_mono, font=lot_font,
                 width=14, anchor="w").pack(side=tk.LEFT)

        # Copy button
        def _copy():
            self.clipboard_clear()
            self.clipboard_append(self._data["lot"])
            self.update()
            cb.config(text="Copied!")
            self.after(1400, lambda: cb.config(text="Copy"))

        cb = styled_button(self, "Copy", _copy, style="copy",
                            font_key="btn_sm", padx=7, pady=2)
        cb.pack(side=tk.LEFT, padx=(0, 14))

        # Meta
        meta = tk.Frame(self, bg=bg)
        meta.pack(side=tk.LEFT, fill=tk.X, expand=True, pady=8)
        tk.Label(meta, text=self._data["name"],
                 bg=bg, fg=fg, font=F["small"], anchor="w").pack(anchor="w")
        time_text = self._data["submitted_at"]
        if is_done and self._data["printed_at"]:
            time_text += f"   →   Done {self._data['printed_at']}"
        tk.Label(meta, text=time_text,
                 bg=bg, fg=T["muted"], font=F["small"], anchor="w").pack(anchor="w")

        # Action button
        tk.Frame(self, width=8, bg=bg).pack(side=tk.RIGHT)
        if is_done:
            styled_button(self, "Undo", lambda: on_toggle(self._data, False),
                           style="muted", font_key="btn_sm", padx=10, pady=4
                           ).pack(side=tk.RIGHT, pady=8)
        else:
            styled_button(self, "Mark Done", lambda: on_toggle(self._data, True),
                           style="success", font_key="btn_sm", padx=10, pady=4
                           ).pack(side=tk.RIGHT, pady=8)

        # Divider
        tk.Frame(self, height=1, bg=T["border"]).place(
            relx=0, rely=1.0, relwidth=1.0, anchor="sw")


# ---------------------------------------------------------------------------
# Main application
# ---------------------------------------------------------------------------

class StationRelayApp(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("StationRelay")
        self.resizable(True, True)
        self.minsize(580, 480)

        self.config_data = load_config()

        # Apply saved theme
        self._theme_name = self.config_data.get("theme", "light")
        apply_theme(self._theme_name)
        self.configure(bg=T["bg"])

        # Cached row data for optimistic UI
        self._rows: list = []

        self._build_ui()

        # Restore always-on-top
        if self.config_data.get("always_on_top", False):
            self.attributes("-topmost", True)

    # ------------------------------------------------------------------
    # Build UI
    # ------------------------------------------------------------------

    def _build_ui(self):
        self._build_titlebar()
        self._build_tabs()

    def _build_titlebar(self):
        self._bar = tk.Frame(self, bg=T["bg"])
        self._bar.pack(fill=tk.X, padx=0, pady=0)

        # App title
        self._title_lbl = tk.Label(
            self._bar, text="StationRelay",
            bg=T["bg"], fg=T["accent"], font=F["title"],
            padx=18, pady=12,
        )
        self._title_lbl.pack(side=tk.LEFT)

        self._sub_lbl = tk.Label(
            self._bar, text="Print Queue",
            bg=T["bg"], fg=T["muted"], font=F["small"],
        )
        self._sub_lbl.pack(side=tk.LEFT)

        # Right-side controls
        right = tk.Frame(self._bar, bg=T["bg"])
        right.pack(side=tk.RIGHT, padx=10, pady=8)

        # Always-on-top pin toggle
        self._pin_var = tk.BooleanVar(
            value=self.config_data.get("always_on_top", False))
        self._pin_btn = tk.Button(
            right,
            text="📌 Pinned" if self._pin_var.get() else "📌 Pin",
            command=self._toggle_pin,
            font=F["btn_sm"], relief=tk.FLAT, cursor="hand2",
            bg=T["accent"] if self._pin_var.get() else T["surface2"],
            fg=T["accent_fg"] if self._pin_var.get() else T["muted"],
            activebackground=T["accent_hover"],
            activeforeground=T["accent_fg"],
            bd=0, padx=10, pady=5,
        )
        self._pin_btn.pack(side=tk.LEFT, padx=(0, 6))

        # Theme toggle
        self._theme_btn = tk.Button(
            right,
            text="☀ Light" if self._theme_name == "dark" else "🌙 Dark",
            command=self._toggle_theme,
            font=F["btn_sm"], relief=tk.FLAT, cursor="hand2",
            bg=T["surface2"], fg=T["muted"],
            activebackground=T["border"],
            activeforeground=T["text"],
            bd=0, padx=10, pady=5,
        )
        self._theme_btn.pack(side=tk.LEFT, padx=(0, 6))

        # Settings
        self._settings_btn = tk.Button(
            right,
            text="⚙ Settings",
            command=self._open_settings,
            font=F["btn_sm"], relief=tk.FLAT, cursor="hand2",
            bg=T["surface2"], fg=T["muted"],
            activebackground=T["border"],
            activeforeground=T["text"],
            bd=0, padx=10, pady=5,
        )
        self._settings_btn.pack(side=tk.LEFT)

        # Divider line
        self._divider = tk.Frame(self, height=1, bg=T["border"])
        self._divider.pack(fill=tk.X)

    def _build_tabs(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        self._apply_tab_style(style)

        self._nb = ttk.Notebook(self, style="SR.TNotebook")
        self._nb.pack(fill=tk.BOTH, expand=True)

        self._build_submit_tab()
        self._build_queue_tab()
        self._build_completed_tab()

        self._nb.bind("<<NotebookTabChanged>>", self._on_tab_changed)

    def _apply_tab_style(self, style: ttk.Style):
        style.configure("SR.TNotebook",
                         background=T["bg"], borderwidth=0, tabmargins=0)
        style.configure("SR.TNotebook.Tab",
                         background=T["surface2"], foreground=T["muted"],
                         font=F["normal"], padding=(18, 7), borderwidth=0)
        style.map("SR.TNotebook.Tab",
                  background=[("selected", T["bg"])],
                  foreground=[("selected", T["text"])])

    # ---- Submit tab ----------------------------------------------------

    def _build_submit_tab(self):
        self._submit_outer = tk.Frame(self._nb, bg=T["bg"])
        self._nb.add(self._submit_outer, text="  Submit  ")

        card = tk.Frame(self._submit_outer, bg=T["surface"], padx=36, pady=32)
        card.place(relx=0.5, rely=0.38, anchor="center")

        self._submit_card = card

        tk.Label(card, text="Submit a Lot Number",
                 bg=T["surface"], fg=T["text"], font=F["title"]).pack(anchor="w")
        tk.Label(card, text="Enter the lot number and press Submit or hit Enter.",
                 bg=T["surface"], fg=T["muted"], font=F["small"]).pack(
                 anchor="w", pady=(4, 22))

        tk.Label(card, text="LOT NUMBER",
                 bg=T["surface"], fg=T["muted"], font=F["small"]).pack(anchor="w")

        row = tk.Frame(card, bg=T["surface"])
        row.pack(fill=tk.X, pady=(4, 0))

        self._lot_var = tk.StringVar()
        self._lot_entry = tk.Entry(
            row,
            textvariable=self._lot_var,
            font=("Consolas", 16),
            relief=tk.FLAT,
            bg=T["surface2"],
            fg=T["text"],
            insertbackground=T["accent"],
            highlightthickness=2,
            highlightbackground=T["border"],
            highlightcolor=T["accent"],
            width=18,
        )
        self._lot_entry.pack(side=tk.LEFT, ipady=10, padx=(0, 10))
        self._lot_entry.bind("<Return>", lambda _e: self._on_send())
        self._lot_entry.focus_set()

        self._send_btn = styled_button(row, "Submit", self._on_send,
                                        style="accent", padx=20, pady=10)
        self._send_btn.pack(side=tk.LEFT)

        self._status_var = tk.StringVar(value="")
        self._status_lbl = tk.Label(
            card, textvariable=self._status_var,
            bg=T["surface"], fg=T["muted"], font=F["small"],
        )
        self._status_lbl.pack(anchor="w", pady=(12, 0))

    # ---- Queue tab -----------------------------------------------------

    def _build_queue_tab(self):
        self._queue_outer = tk.Frame(self._nb, bg=T["bg"])
        self._nb.add(self._queue_outer, text="  Print Queue  ")
        self._queue_status_var = tk.StringVar(value="")
        self._queue_inner = self._build_list_tab(
            self._queue_outer, self._queue_status_var, self._load_queue)

    # ---- Completed tab -------------------------------------------------

    def _build_completed_tab(self):
        self._comp_outer = tk.Frame(self._nb, bg=T["bg"])
        self._nb.add(self._comp_outer, text="  Completed  ")
        self._comp_status_var = tk.StringVar(value="")
        self._comp_inner = self._build_list_tab(
            self._comp_outer, self._comp_status_var, self._load_queue)

    # ---- Shared list-tab builder ---------------------------------------

    def _build_list_tab(self, outer: tk.Frame,
                         status_var: tk.StringVar,
                         refresh_cmd) -> tk.Frame:
        toolbar = tk.Frame(outer, bg=T["bg"], pady=6)
        toolbar.pack(fill=tk.X, padx=16)
        tk.Label(toolbar, textvariable=status_var,
                 bg=T["bg"], fg=T["muted"], font=F["small"]).pack(side=tk.LEFT)
        styled_button(toolbar, "Refresh", refresh_cmd,
                      style="muted", font_key="btn_sm",
                      padx=12, pady=3).pack(side=tk.RIGHT)

        hdr = tk.Frame(outer, bg=T["surface2"], padx=14, pady=5)
        hdr.pack(fill=tk.X, padx=16)
        for txt, w in [("#", 4), ("Lot Number", 14), ("", 8),
                        ("Submitted By", 18), ("Submitted At", 20), ("", 0)]:
            tk.Label(hdr, text=txt, bg=T["surface2"], fg=T["muted"],
                     font=F["small"], width=w, anchor="w").pack(side=tk.LEFT)

        scroll_frame, inner = self._make_scroll_area(outer)
        scroll_frame.pack(fill=tk.BOTH, expand=True, padx=16, pady=(0, 8))
        return inner

    # ---- Scroll area ---------------------------------------------------

    def _make_scroll_area(self, parent) -> tuple:
        container = tk.Frame(parent, bg=T["bg"])
        canvas = tk.Canvas(container, bg=T["bg"], highlightthickness=0)
        vsb = tk.Scrollbar(container, orient="vertical", command=canvas.yview,
                            bg=T["surface2"], troughcolor=T["bg"],
                            activebackground=T["border"])
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        inner = tk.Frame(canvas, bg=T["bg"])
        win = canvas.create_window((0, 0), window=inner, anchor="nw")

        inner.bind("<Configure>",
                   lambda e: (canvas.configure(scrollregion=canvas.bbox("all")),
                               canvas.itemconfig(win, width=canvas.winfo_width())))
        canvas.bind("<Configure>",
                    lambda e: canvas.itemconfig(win, width=e.width))

        def _scroll(e):
            canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")

        canvas.bind("<Enter>",  lambda _: canvas.bind_all("<MouseWheel>", _scroll))
        canvas.bind("<Leave>",  lambda _: canvas.unbind_all("<MouseWheel>"))

        return container, inner

    # ------------------------------------------------------------------
    # Theme toggle
    # ------------------------------------------------------------------

    def _toggle_theme(self):
        self._theme_name = "light" if self._theme_name == "dark" else "dark"
        apply_theme(self._theme_name)
        self.config_data["theme"] = self._theme_name
        save_config(self.config_data)
        # Full rebuild — simplest reliable way to recolour everything
        self._full_rebuild()

    def _full_rebuild(self):
        self.configure(bg=T["bg"])
        for w in self.winfo_children():
            w.destroy()
        self._build_ui()
        self._load_queue()

    # ------------------------------------------------------------------
    # Always-on-top toggle
    # ------------------------------------------------------------------

    def _toggle_pin(self):
        new_val = not self._pin_var.get()
        self._pin_var.set(new_val)
        self.attributes("-topmost", new_val)
        self.config_data["always_on_top"] = new_val
        save_config(self.config_data)
        if new_val:
            self._pin_btn.config(text="📌 Pinned",
                                  bg=T["accent"], fg=T["accent_fg"])
        else:
            self._pin_btn.config(text="📌 Pin",
                                  bg=T["surface2"], fg=T["muted"])

    # ------------------------------------------------------------------
    # Settings
    # ------------------------------------------------------------------

    def _open_settings(self):
        SettingsDialog(self)

    # ------------------------------------------------------------------
    # Submit
    # ------------------------------------------------------------------

    def _on_send(self):
        lot = self._lot_var.get().strip()
        if not lot:
            self._set_status("Please enter a lot number.", T["error"])
            return

        excel_path = self.config_data.get("excel_path", "").strip()
        if not excel_path:
            messagebox.showerror("No Excel file",
                                 "Open ⚙ Settings and select an Excel file.")
            return
        if not os.path.isfile(excel_path):
            messagebox.showerror("File not found",
                                 f"Cannot find:\n{excel_path}\n\nUpdate in ⚙ Settings.")
            return

        self._send_btn.config(state=tk.DISABLED, text="Submitting…")
        self._set_status("Saving…", T["muted"])

        threading.Thread(
            target=self._send_worker,
            args=(excel_path,
                  self.config_data.get("sheet_name", "Sheet1"),
                  self.config_data.get("column", "A"),
                  lot),
            daemon=True,
        ).start()

    def _send_worker(self, excel_path, sheet_name, column, lot):
        try:
            append_to_excel(excel_path, sheet_name, column, lot)
            self.after(0, lambda: self._send_success(lot))
        except Exception as exc:
            msg = _format_exc(exc)
            self.after(0, lambda: self._send_error(msg))

    def _send_success(self, lot: str):
        self._lot_var.set("")
        self._send_btn.config(state=tk.NORMAL, text="Submit")
        self._set_status(f"Lot {lot} submitted.", T["success"])
        self._lot_entry.focus_set()
        self._load_queue()

    def _send_error(self, message: str):
        self._send_btn.config(state=tk.NORMAL, text="Submit")
        self._set_status("Submit failed.", T["error"])
        messagebox.showerror("Submit failed", message)

    def _set_status(self, msg: str, colour: str):
        self._status_var.set(msg)
        self._status_lbl.config(fg=colour)

    # ------------------------------------------------------------------
    # Queue loading
    # ------------------------------------------------------------------

    def _on_tab_changed(self, _event=None):
        tab = self._nb.tab(self._nb.select(), "text").strip()
        if tab in ("Print Queue", "Completed"):
            self._load_queue()

    def _load_queue(self):
        excel_path = self.config_data.get("excel_path", "").strip()
        if not excel_path or not os.path.isfile(excel_path):
            self._queue_status_var.set("No Excel file — open ⚙ Settings.")
            self._comp_status_var.set("No Excel file — open ⚙ Settings.")
            return
        self._queue_status_var.set("Loading…")
        self._comp_status_var.set("Loading…")
        threading.Thread(
            target=self._queue_worker,
            args=(excel_path,
                  self.config_data.get("sheet_name", "Sheet1"),
                  self.config_data.get("column", "A")),
            daemon=True,
        ).start()

    def _queue_worker(self, excel_path, sheet_name, column):
        try:
            rows = read_from_excel(excel_path, sheet_name, column)
            self.after(0, lambda: self._store_and_render(rows))
        except Exception as exc:
            msg = _format_exc(exc)
            self.after(0, lambda: self._queue_load_error(msg))

    def _store_and_render(self, rows: list):
        self._rows = rows
        self._render_rows(rows)

    def _queue_load_error(self, message: str):
        ts = datetime.now().strftime("%H:%M:%S")
        err = f"Load failed at {ts}"
        self._queue_status_var.set(err)
        self._comp_status_var.set(err)
        messagebox.showerror(
            "Queue load failed",
            f"{message}\n\n"
            "Possible causes:\n"
            "  • Excel file is open on this PC — close it and retry\n"
            "  • OneDrive is syncing — wait a moment and retry\n"
            "  • File path is wrong — check ⚙ Settings",
        )

    def _render_rows(self, rows: list):
        # Sort all rows by submitted_at, then split
        sorted_rows = _sort_rows(rows)
        pending   = [r for r in sorted_rows if not r["printed_at"]]
        completed = [r for r in sorted_rows if r["printed_at"]]

        for i, r in enumerate(pending, 1):
            r["queue_pos"] = i
        for i, r in enumerate(completed, 1):
            r["queue_pos"] = i

        self._render_list(self._queue_inner, pending,   is_done=False)
        self._render_list(self._comp_inner,  completed, is_done=True)

        p, d, total = len(pending), len(completed), len(rows)
        ts = datetime.now().strftime("%H:%M:%S")
        self._queue_status_var.set(
            f"{p} pending  ·  {d} done  ·  {total} total  —  {ts}")
        self._comp_status_var.set(
            f"{d} completed  ·  {total} total  —  {ts}")

    def _render_list(self, container: tk.Frame, rows: list, is_done: bool):
        for w in container.winfo_children():
            w.destroy()
        if not rows:
            msg = ("No completed lots yet."
                   if is_done else
                   "Queue is empty — submit a lot on the Submit tab.")
            tk.Label(container, text=msg,
                     bg=T["bg"], fg=T["muted"], font=F["small"]).pack(pady=30)
            return
        for i, rd in enumerate(rows):
            QueueRow(container, rd,
                     on_toggle=self._toggle_done,
                     is_done=is_done, index=i).pack(fill=tk.X)

    # ------------------------------------------------------------------
    # Toggle done — optimistic UI first, Excel sync in background
    # ------------------------------------------------------------------

    def _toggle_done(self, row_data: dict, mark_as_done: bool):
        # --- Optimistic update: mutate cached rows and re-render now ---
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for r in self._rows:
            if r["row"] == row_data["row"]:
                r["printed_at"] = now_str if mark_as_done else ""
                break
        self._render_rows(self._rows)

        # --- Background Excel write ------------------------------------
        excel_path = self.config_data.get("excel_path", "").strip()
        if not excel_path:
            return

        sheet_name = self.config_data.get("sheet_name", "Sheet1")
        column     = self.config_data.get("column", "A")
        row_num    = row_data["row"]
        status_var = self._queue_status_var

        def _worker():
            try:
                if mark_as_done:
                    mark_done_in_excel(excel_path, sheet_name, column, row_num)
                else:
                    unmark_done_in_excel(excel_path, sheet_name, column, row_num)
                # Silent background re-sync to confirm Excel agrees
                self.after(0, self._load_queue)
            except Exception as exc:
                msg = _format_exc(exc)
                self.after(0, lambda: (
                    messagebox.showerror("Save error", msg),
                    # Rollback optimistic change
                    self._load_queue()
                ))

        status_var.set("Syncing to Excel…")
        threading.Thread(target=_worker, daemon=True).start()


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    app = StationRelayApp()
    app.mainloop()
