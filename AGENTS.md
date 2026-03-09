# StationRelay — Agent Guide

## Project Overview

StationRelay is a single-file Python desktop application (`app.py`).  
The user pastes text into a Tkinter UI and clicks **Send**. The app opens a
configured Excel file invisibly via Windows COM (`pywin32`), appends a new row
with the pasted text, the user's full display name, and a timestamp, saves, and
closes Excel silently. It then nudges OneDrive to upload immediately.

**Platform:** Windows only (COM + `pywin32` + Tkinter).  
**Python:** 3.10+ recommended. Uses f-strings, `|` union types are not used yet.

---

## File Structure

```
app.py            # Entire application — UI + COM logic + config helpers
requirements.txt  # Runtime deps (pywin32 only; tkinter ships with Python)
config.json       # Auto-generated at runtime — user settings (gitignored)
AGENTS.md         # This file
```

No build step, no bundler, no virtual environment required beyond `pip install pywin32`.

---

## Running the App

```powershell
# Dev (shows console, useful for print-debugging)
python app.py

# Windowless (production-style)
pythonw app.py
```

---

## Install / Setup

```powershell
pip install pywin32
```

That is the only third-party dependency. `tkinter`, `json`, `os`, `threading`,
`datetime`, and `sys` are all stdlib.

---

## Build / Package as EXE

```powershell
pip install pyinstaller
pyinstaller --onefile --windowed app.py
# Output: dist/app.exe
```

---

## Linting and Formatting

No linter or formatter is configured yet. Follow these conventions:

```powershell
# Recommended — install manually if needed
pip install black flake8

# Format
black app.py

# Lint
flake8 app.py --max-line-length=100
```

Target: **zero flake8 warnings**, black-formatted output.

---

## Testing

There are no automated tests yet. When adding tests:

```powershell
# Run all tests
python -m pytest tests/

# Run a single test function
python -m pytest tests/test_excel.py::test_col_letter_to_index -v

# Run without pywin32/COM (mock it)
python -m pytest tests/ -k "not com" -v
```

Place tests in a `tests/` directory. Mock all COM calls with `unittest.mock`
— never open real Excel in tests.

---

## Code Style

### Imports
- Stdlib imports first, then third-party, then local — each group separated by a blank line.
- Platform-specific imports (`win32com`, `pythoncom`, `win32api`) must be inside
  the function body, not at module level. This allows the module to import on
  non-Windows machines for testing.
- Always add `# type: ignore` to `win32*` / `pythoncom` imports — stubs don't exist.

```python
# Good
def append_to_excel(...) -> None:
    import pythoncom          # type: ignore
    import win32com.client    # type: ignore
    ...
```

### Type Hints
- All public functions must have full type annotations on parameters and return type.
- Private helpers (prefixed `_`) should also be annotated.
- Use `dict` / `list` (not `Dict` / `List`) — Python 3.9+ generics.
- `Optional[X]` → use `X | None` only if targeting 3.10+; otherwise keep `Optional`.

```python
def load_config() -> dict:          # good
def save_config(cfg: dict) -> None: # good
```

### Naming
| Kind | Convention | Example |
|---|---|---|
| Functions / methods | `snake_case` | `append_to_excel` |
| Private helpers | `_snake_case` | `_col_letter_to_index` |
| Classes | `PascalCase` | `StationRelayApp` |
| Constants | `UPPER_SNAKE` | `CONFIG_FILE`, `DEFAULT_CONFIG` |
| Tkinter instance vars | `_name_var` / `_name_btn` | `_send_btn`, `_path_var` |

### String formatting
- Use f-strings for all string interpolation.
- Multiline strings: use implicit concatenation or `\n` joins — avoid `%` formatting.

### Line length
- Soft limit: **100 characters**. Hard limit: **120**.
- Break long `tk.Widget(...)` calls onto multiple lines with trailing comma.

### Section headers
Separate logical sections with a consistent comment banner:

```python
# ---------------------------------------------------------------------------
# Section name
# ---------------------------------------------------------------------------
```

---

## Architecture Notes

### Single-file design
Everything lives in `app.py`. If the file grows beyond ~600 lines, extract into:
- `excel_writer.py` — COM automation (`append_to_excel`, helpers)
- `config.py` — `load_config`, `save_config`, `DEFAULT_CONFIG`
- `app.py` — UI only

### Threading model
- The main thread owns all Tkinter widgets. **Never touch widgets from a background thread.**
- COM work runs on a `daemon=True` `threading.Thread`.
- Use `self.after(0, callback)` to marshal results back to the UI thread.

### COM lifecycle (critical)
Always follow this pattern on the background thread:

```python
pythoncom.CoInitialize()
try:
    excel = win32com.client.DispatchEx("Excel.Application")  # DispatchEx, NOT Dispatch
    excel.Visible = False
    excel.DisplayAlerts = False
    # ... work ...
    wb.Save()
finally:
    try: wb.Close(SaveChanges=False)
    except Exception: pass
    try: excel.Quit()
    except Exception: pass
    wb = None
    excel = None          # release COM refs before CoUninitialize
    pythoncom.CoUninitialize()
```

- Use `DispatchEx` (not `Dispatch`) — always spawns a fresh Excel process,
  preventing zombie-instance hangs on repeated sends.
- Call `CoInitialize` / `CoUninitialize` around every background-thread COM session.
- Always `Quit()` Excel — a leaked process will block the next send.

### Config persistence
Settings are stored in `config.json` next to `app.py`. The config dict is
always merged with `DEFAULT_CONFIG` on load so new keys are never missing.
Never write credentials or sensitive paths to `config.json`.

### OneDrive sync nudge
After saving, `_nudge_onedrive()` calls `os.utime()` on the file to bump
`LastWriteTime`. OneDrive's file-system watcher reacts within ~1-2 seconds.
This function must never raise — wrap the entire body in `try/except Exception: pass`.

---

## Common Pitfalls

1. **`Dispatch` vs `DispatchEx`** — `Dispatch` reuses a running Excel instance;
   if a previous send left a zombie, the second send hangs. Always use `DispatchEx`.

2. **COM on wrong thread** — calling `win32com` without `CoInitialize` on a
   non-main thread raises `CoInitialize has not been called`. Always initialize
   COM at the top of the worker function.

3. **Touching widgets from background thread** — causes silent corruption or
   crashes. Use `self.after(0, fn)` to schedule UI updates on the main thread.

4. **`wb.Close(SaveChanges=False)` after `wb.Save()`** — always pass
   `SaveChanges=False`; otherwise Excel prompts for confirmation on close.

5. **Column index is 1-based in COM** — `sheet.Cells(row, 1)` is column A.
   `_col_letter_to_index("A")` returns `1`, not `0`.
