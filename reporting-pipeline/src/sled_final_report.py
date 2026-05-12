# ---------------- Standard Library ----------------
import gc
import logging
import os
import sys
import re
import time
import shutil
import importlib
import subprocess
import traceback
from pathlib import Path
from datetime import datetime, timedelta
from datetime import datetime, date

# ---------------- Windows / COM Libraries ----------------
import pythoncom
import pywintypes
from requests import session
import win32com.client
import win32com.client as win32
from win32com.client import GetObject, Dispatch
import win32clipboard

# ---------------- Data / Excel Libraries ----------------
import pandas as pd
import pyodbc
import xlwings as xw
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Border, Side, Font

# ---------------- Package Checker ----------------
required_packages = [
    "pandas",
    "pyodbc",
    "xlwings",
    "openpyxl",
    "pywin32",
]

def install_if_missing(package_name: str):
    """Install a package if it is not already available."""
    try:
        importlib.import_module(package_name)
    except ImportError:
        print(f"Package '{package_name}' not found. Installing...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", package_name])

# # Check and install missing packages
# for pkg in required_packages:
#     install_if_missing(pkg)

# print("All required packages are installed.")

# subprocess.Popen(r'/path/to/file Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe') 
# time.sleep(3)

# sapguiauto = win32com.client.GetObject("SAPGUI")
# application = sapguiauto.GetScriptingEngine
# connection = application.OpenConnection("Login", True)
# session = connection.Children(0)


# session.findById("wnd[0]/path/to/file").text = "100"
# session.findById("wnd[0]/path/to/file").text = "DUMMY_USER"
# sap_password = os.getenv("SAP_PASSWORD")

# if not sap_password:
#     raise ValueError("SAP_PASSWORD environment variable is not set.")

# session.findById("wnd[0]/path/to/file").text = sap_password
# session.findById("wnd[0]/path/to/file").text = "EN"
# session.findById("wnd[0]").sendVKey(0)

# ---------------- Date Calculations ----------------
# Calculate yesterday's date in DD.MM.YYYY format
yesterday = (datetime.now() - timedelta(days=3)).strftime("%d.%m.%Y")

# Previous week: last 7 days
start_date = (datetime.now() - timedelta(days=7)).strftime('%Y-%m-%d')
end_date = datetime.now().strftime('%Y-%m-%d')

# ---------------- CONFIG ----------------
# --- Base directories ---
BASE_DIR = Path(r"~\Demo Company\Master Data Governance - Documents\Projects\shelf life").expanduser()
EXTEND_DIR = BASE_DIR / "Extended"
CREATED_DIR = BASE_DIR / "Created"
CHANGES_DIR = BASE_DIR / "Changes"

SOURCE_ROOT = Path(r"~\Demo Company\Master Data Governance - Documents\Projects\shelf life\Created").expanduser()
DEST_FOLDER = Path(r"~\Demo Company\Master Data Governance - Documents\Projects\shelf life\Load to Sharepoint").expanduser()

NEW_NAME = "load file.xlsx"

# --- Filenames & paths ---
NEW_LISTINGS_PATH = CREATED_DIR / f"new_listings_{yesterday}.xlsx"
NEW_LISTINGS_FILENAME = NEW_LISTINGS_PATH.name

HIERARCHY_PATH = CREATED_DIR / f"Hierarchy_{yesterday}.xlsx"
HIERARCHY_FILENAME = f"Hierarchy_{yesterday}.xlsx"

INFORECORD = CREATED_DIR / f"Inforecord_{yesterday}.xlsx"
PURCH_ORG_PATH = CREATED_DIR / f"Purch_Org_{yesterday}.xlsx"
PURCH_ORG_DESC_PATH = CREATED_DIR / f"Purch_Org_Desc_{yesterday}.xlsx"
INFORECORD_FILENAME = INFORECORD.name
PURCH_ORG_FILENAME = PURCH_ORG_PATH.name
PURCH_ORG_DESC_FILENAME = PURCH_ORG_DESC_PATH.name

CHANGE_LOG_PATH = CHANGES_DIR / f"Change_Log_{start_date}_{end_date}.xlsx"
CHANGE_LOG_FILENAME = CHANGE_LOG_PATH.name
HIERARCHY_PATH_CHNGE = CHANGES_DIR / HIERARCHY_FILENAME
INFORECORD_1 = CHANGES_DIR / f"Inforecord_{yesterday}.xlsx"
PURCH_ORG_PATH_1 = CHANGES_DIR / f"Purch_Org_{yesterday}.xlsx"
PURCH_ORG_PATH_2 = CHANGES_DIR / f"Purch_Org_Desc_{yesterday}.xlsx"
INFORECORD_FILENAME_1 = INFORECORD.name
PURCH_ORG_FILENAME_1 = PURCH_ORG_PATH_1.name
PURCH_ORG_DESC_FILENAME_1 = PURCH_ORG_PATH_2.name

EXTEND_PATH = EXTEND_DIR / f'Extended_To_DC_Shelf_Life_{yesterday}.xlsx'
EXTENDED_DATA_FILENAME = EXTEND_PATH.name

SOURCE_FILE_EXTEND = f"Shelf_Life_Data_{yesterday}.xlsx"
SOURCE_FILE_EXTEND_PATH = EXTEND_DIR / f"Shelf_Life_Data_{yesterday}.xlsx"
HIERARCHY_FILE_EXTEND = EXTEND_DIR / HIERARCHY_FILENAME

REPORT_FILE = CREATED_DIR / f"new_listings_{yesterday}_report.xlsx"
EXTEND_REPORT_FILE = EXTEND_DIR / f"Extended_To_DC_Shelf_Life_Report_{yesterday}.xlsx"

# --- Excel settings ---
SHEET_NAME = "Sheet1"   # sheet to use in both workbooks (falls back to first sheet)
MAP_RETURN_COLS = {"BMC Description": "I","Divisional Buyer": "L",}
NOT_FOUND_TEXT = "Not found"
XL_UP = -4162
HIER_TRANSACTION = None

NEW_HEADERS_EXTEND = ["Divisional Buyer", "Date Created", "Min. Rem. Shelf Life", "Total shelf life", "X-site artl status"]

# --- Timeout & sleep ---
WAIT_FOR_EXCEL_OPEN_SEC = 15
SHORT_SLEEP = 0.5
LONG_SLEEP = 2.0

# --- Access DB settings ---
ACCESS_DB = BASE_DIR / "Shelf_Life.accdb"
ACCESS_TABLES = {
    "created": "Shelf_Life_Created",
    "extended": "Shelf_Life_Extended",
    "change_log": "Shelf_Life_Change_Log"
}

RUN_CONTEXT = {
    "run_date": None,
    "days_back_created": None,
    "days_back_extended": None,
}

RUN_DATE = None
DAYS_BACK_CREATED = None
DAYS_BACK_EXTENDED = None

# ---------------- HELPERS ----------------
# ---------------- File Helpers ----------------
def wait_for_file(path: Path, timeout: int = 30):
    """Wait until a file exists and is ready."""
    start = time.time()
    while not path.exists():
        if time.time() - start > timeout:
            raise FileNotFoundError(f"{path} not found after {timeout}s")
        time.sleep(1)
    time.sleep(1)  # extra buffer
    print(f"File ready: {path}")

def safe_move(src: Path, dst: Path, retries=10, delay=2):
    """Move a file with retries if it's in use."""
    for i in range(retries):
        try:
            shutil.move(str(src), str(dst))
            print(f"Moved {src.name} -> {dst}")
            return
        except PermissionError:
            print(f"{src.name} still in use... retry {i+1}/path/to/file")
            time.sleep(delay)
    raise PermissionError(f"Could not move {src.name} after {retries} retries")

def move_files_to_folder(files: list[Path], folder: Path):
    """Move a list of files to a folder, creating folder if needed."""
    folder.mkdir(exist_ok=True)
    for f in files:
        if f.exists():
            safe_move(f, folder / f.name)
        else:
            print(f"File not found, skipping: {f.name}")
    print(f"All files moved to {folder}")

def extract_date_from_filename(path: Path) -> str:
    """Extract date from filename; returns DD.MM.YYYY."""
    name = path.name
    m = re.search(r"(\d{2}[.\-]\d{2}[.\-]\d{4})", name)
    if m:
        return m.group(1).replace("-", ".")
    m2 = re.search(r"(\d{4}[.\-]\d{2}[.\-]\d{2})", name)
    if m2:
        parts = re.split(r"[.\-]", m2.group(1))
        return f"{parts[2]}.{parts[1]}.{parts[0]}"
    return datetime.now().strftime("%d.%m.%Y")

def get_sap_session():
    """
    Attach to an available SAP GUI scripting session.
    Tries multiple COM access patterns and safely enumerates sessions.
    """
    errors = []

    def _first_session(app):
        # Try to find any available session in any connection
        try:
            for c in range(app.Children.Count):
                conn = app.Children(c)
                if conn is None:
                    continue
                for s in range(conn.Children.Count):
                    ses = conn.Children(s)
                    if ses is not None:
                        return ses
        except Exception as e:
            raise e
        return None

    # Pattern 1: GetObject("SAPGUI")
    try:
        sapguiauto = win32com.client.GetObject("SAPGUI")
        app = sapguiauto.GetScriptingEngine
        ses = _first_session(app)
        if ses:
            print("Attached to SAP GUI session (GetObject 'SAPGUI').")
            return ses
        raise RuntimeError("SAP scripting engine found but no active sessions.")
    except Exception as e:
        errors.append(("GetObject('SAPGUI')", e))

    # Pattern 2: Dispatch("Sapgui.ScriptingCtrl.1")
    try:
        sapguiauto = win32com.client.Dispatch("Sapgui.ScriptingCtrl.1")
        app = sapguiauto.GetScriptingEngine
        ses = _first_session(app)
        if ses:
            print("Attached to SAP GUI session (Dispatch 'Sapgui.ScriptingCtrl.1').")
            return ses
        raise RuntimeError("SAP scripting control found but no active sessions.")
    except Exception as e:
        errors.append(("Dispatch('Sapgui.ScriptingCtrl.1')", e))

    # Pattern 3: Dispatch("SAPGUI") (some environments register this differently)
    try:
        sapguiauto = win32com.client.Dispatch("SAPGUI")
        app = sapguiauto.GetScriptingEngine
        ses = _first_session(app)
        if ses:
            print("Attached to SAP GUI session (Dispatch 'SAPGUI').")
            return ses
        raise RuntimeError("SAP GUI object found but no active sessions.")
    except Exception as e:
        errors.append(("Dispatch('SAPGUI')", e))

    # If everything fails, print a useful summary
    print("Unable to attach to SAP GUI session.")
    for name, err in errors:
        print(f"--- {name} failed ---")
        print(err)

    print("\nChecklist:")
    print("- SAP Logon is open")
    print("- You are logged into a system (a session exists)")
    print("- SAP GUI Scripting is enabled (SAP Logon Options) and server allows it")
    print("- Python and SAP are running under the same Windows user/path/to/file")
    raise RuntimeError("Could not attach to SAP GUI session via COM. See errors above.")

def get_worksheet_safe(workbook, sheet_name=None, index=1, retries=20, delay=2):
    """
    Safely get an Excel worksheet even if Excel is busy.
    Falls back from sheet_name to index.
    Prints available sheets if access keeps failing.
    """
    last_error = None

    for i in range(retries):
        try:
            # Try by name first if provided
            if sheet_name:
                return workbook.Worksheets(sheet_name)

            # Fallback to index
            return workbook.Worksheets(index)

        except Exception as e:
            last_error = e
            print(f"Excel busy or worksheet not ready, retry {i+1}/path/to/file")
            time.sleep(delay)

    # Final diagnostics
    try:
        sheet_names = [ws.Name for ws in workbook.Worksheets]
        print("Available sheets in workbook:", sheet_names)
    except Exception as e:
        print("Could not list worksheets — Excel likely blocked (dialog / Protected View).")
        print("Details:", e)

    raise RuntimeError(
        "Could not access Excel worksheet after retries. "
        "Excel may be blocked by Protected View, file lock, or a hidden dialog."
    )

# ---------------- Clipboard Helpers ----------------
def set_clipboard_text(text: str):
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardText(text)
    win32clipboard.CloseClipboard()

def get_clipboard_text() -> str:
    win32clipboard.OpenClipboard()
    try:
        data = win32clipboard.GetClipboardData(win32clipboard.CF_TEXT)
        return data.decode("utf-8")
    except TypeError:
        return ""
    finally:
        win32clipboard.CloseClipboard()

def copy_column_to_clipboard(file_path: Path, column_index: int):
    """Copy a column from Excel to clipboard (0-based column_index)."""
    df = pd.read_excel(file_path)
    col = df.iloc[:, column_index].dropna().astype(str)
    set_clipboard_text("\r\n".join(col))
    print(f"Column {column_index} copied to clipboard from {file_path.name}")

# ---------------- Excel Helpers ----------------
def kill_excel():
    """Kill all Excel processes."""
    subprocess.run(["taskkill", "/path/to/file", "/path/to/file", "EXCEL.EXE", "/path/to/file"],
                   check=False, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    print("All Excel processes terminated.")

def wait_for_excel_workbook(name: str, timeout: int = 20):
    """Wait until an Excel workbook with `name` is open; returns (excel_app, workbook)."""
    end = time.time() + timeout
    while time.time() < end:
        try:
            excel_app = win32com.client.GetObject(None, "Excel.Application")
            for wb in excel_app.Workbooks:
                if os.path.basename(wb.FullName).lower() == name.lower():
                    return excel_app, wb
        except Exception:
            pass
        time.sleep(1)
    return None, None

def copy_column_from_workbook(wb, column_letter='D', start_row=2):
    """Copy a column from an open workbook; returns list of strings (bulk-safe)."""
    from win32com.client import constants as XL

    ws = wb.Worksheets(1)
    col_idx = col_letter_to_index(column_letter)

    try:
        last_row = ws.Cells(ws.Rows.Count, col_idx).End(XL.xlUp).Row
    except Exception:
        used = ws.UsedRange
        last_row = used.Row + used.Rows.Count - 1

    if last_row < start_row:
        return []

    rng = ws.Range(ws.Cells(start_row, col_idx), ws.Cells(last_row, col_idx))
    vals = rng.Value  # bulk read

    out = []
    if vals is None:
        return out

    # Multiple rows -> tuple of tuples: ((v1,), (v2,), ...)
    if isinstance(vals, tuple) and len(vals) > 0 and isinstance(vals[0], tuple):
        for row in vals:
            v = row[0]
            out.append("" if v is None else str(v))
        return out

    # Single row / single cell -> scalar or (v,)
    if isinstance(vals, tuple):
        v = vals[0]
        return ["" if v is None else str(v)]

    return ["" if vals is None else str(vals)]

def open_workbook_with_retry(excel, path: Path, retries=8, delay=0.6, readonly_fallback=False):
    """Open workbook with retries and optional readonly fallback."""
    path_str = str(path)
    for attempt in range(1, retries + 1):
        try:
            wb = excel.Workbooks.Open(path_str)
            print(f"Opened {path.name} (attempt {attempt})")
            return wb
        except pywintypes.com_error as e:
            if readonly_fallback:
                try:
                    wb = excel.Workbooks.Open(path_str, ReadOnly=True)
                    print(f"Opened {path.name} as ReadOnly fallback")
                    return wb
                except Exception:
                    pass
            time.sleep(delay)
    raise RuntimeError(f"Failed to open {path}")

def build_hierarchy_lookup_fast(ws_h, map_return_cols):
    """
    FAST: read key col H + only return cols in one range, then build dict in Python.
    Expects headers in row 1 and data from row 2.
    """
    key_col_idx = col_letter_to_index("H")
    ret_idxs = {h: col_letter_to_index(col) for h, col in map_return_cols.items()}

    max_col = max([key_col_idx] + list(ret_idxs.values()))

    last_row = ws_h.Cells(ws_h.Rows.Count, key_col_idx).End(XL_UP).Row
    if last_row < 2:
        return {}

    # Bulk read H..max_col for all rows
    data = ws_h.Range(ws_h.Cells(2, 1), ws_h.Cells(last_row, max_col)).Value2
    if not data:
        return {}

    # Normalize single-row -> tuple of tuples
    if isinstance(data, tuple) and data and not isinstance(data[0], tuple):
        data = (data,)

    lookup = {}
    for row in data:
        k = row[key_col_idx - 1]  # because row is 0-based list/path/to/file
        if k is None:
            continue
        ks = str(k).strip()
        if ks.endswith(".0"):
            ks = ks[:-2]
        if not ks:
            continue

        row_dict = {}
        for header, idx in ret_idxs.items():
            v = row[idx - 1]
            row_dict[header] = v
        lookup[ks] = row_dict

    return lookup

def apply_basic_formatting_simple(ws, excel_app=None, header_row=1):
    """Fast + simple formatting for small sheets (<200 rows)."""
    # xlwings -> COM
    if hasattr(ws, "api"):
        ws = ws.api

    if excel_app is None:
        try:
            excel_app = ws.Application
        except Exception:
            excel_app = None

    used = ws.UsedRange
    last_row = used.Row + used.Rows.Count - 1
    last_col = used.Column + used.Columns.Count - 1
    if last_row < header_row or last_col < 1:
        return

    rng_header = ws.Range(ws.Cells(header_row, 1), ws.Cells(header_row, last_col))
    rng_all = ws.Range(ws.Cells(header_row, 1), ws.Cells(last_row, last_col))

    # Base
    try:
        rng_all.Font.Name = "Calibri"
        rng_all.Font.Size = 11
        rng_all.WrapText = False
    except Exception:
        pass

    # Header (bulk)
    try:
        rng_header.Font.Bold = True
        rng_header.HorizontalAlignment = -4108  # xlCenter
        rng_header.VerticalAlignment = -4108    # xlCenter
        rng_header.Interior.Color = 14277081  # light gray
        rng_header.AutoFilter()
    except Exception:
        pass

    # Freeze panes
    try:
        ws.Activate()
        ws.Cells(header_row + 1, 1).Select()
        if excel_app and excel_app.ActiveWindow:
            excel_app.ActiveWindow.FreezePanes = False
            excel_app.ActiveWindow.FreezePanes = True
    except Exception:
        pass

    # ---- Find "Min. Rem. Shelf Life" column ----
    target_col = None
    try:
        header_vals = rng_header.Value2
        if isinstance(header_vals, tuple):
            header_vals = header_vals[0]  # 1-row range -> tuple-of-tuples

        for i, v in enumerate(header_vals, start=1):
            if v and str(v).strip().lower() == "min. rem. shelf life":
                target_col = i
                break
    except Exception:
        target_col = None

    # ---- Conditional format 999 -> yellow ----
    if target_col:
        col_rng = ws.Range(ws.Cells(header_row + 1, target_col), ws.Cells(last_row, target_col))

        try:
            col_rng.FormatConditions.Delete()
        except Exception:
            pass

        fc = col_rng.FormatConditions.Add(
            Type=1,          # xlCellValue
            Operator=3,      # xlEqual
            Formula1="=999"  # IMPORTANT: include '='
        )
        fc.Interior.Color = 65535  # Yellow
        fc.StopIfTrue = False

    # AutoFit is fine for <200 rows
    try:
        ws.UsedRange.Columns.AutoFit()
    except Exception:
        pass

def open_book_safe(app, path, retries=8, delay=0.6):
    last = None
    for i in range(retries):
        try:
            # update_links=False prevents link prompts; read_only=False normal
            wb = app.books.open(str(path), update_links=False, read_only=False)
            return wb
        except Exception as e:
            last = e
            time.sleep(delay)
    raise last

def get_sheet_safe(wb, index=0, retries=10, delay=0.5):
    last = None
    for i in range(retries):
        try:
            ws = wb.sheets[index]
            # "touch" the sheet to ensure COM object is alive
            _ = ws.range("A1").options(empty="").value
            return ws
        except Exception as e:
            last = e
            time.sleep(delay)
    raise last

def com_retry(fn, retries=8, delay=0.6):
    last = None
    for i in range(retries):
        try:
            return fn()
        except pywintypes.com_error as e:
            last = e
            time.sleep(delay)
    raise last

# ---------------- Utility Helpers ----------------
def col_letter_to_index(col_letter: str) -> int:
    """Convert Excel column letter to 1-based index."""
    col_letter = col_letter.upper()
    col_index = 0
    for i, ch in enumerate(reversed(col_letter)):
        col_index += (ord(ch) - ord("A") + 1) * (26 ** i)
    return col_index

def wait_for_file_stable(path: Path, timeout: int = 180, stable_seconds: int = 3):
    """
    Wait until a file exists AND its size stays unchanged for stable_seconds.
    Helps a lot with OneDrive/path/to/file sync + SAP exports.
    """
    start = time.time()
    last_size = -1
    stable_start = None

    while time.time() - start < timeout:
        if path.exists():
            try:
                size = path.stat().st_size
            except OSError:
                time.sleep(1)
                continue

            if size > 0 and size == last_size:
                if stable_start is None:
                    stable_start = time.time()
                if time.time() - stable_start >= stable_seconds:
                    print(f"File stable: {path}")
                    return
            else:
                stable_start = None
                last_size = size

        time.sleep(1)

    raise TimeoutError(f"File not stable after {timeout}s: {path}")

#------------------------------------------------------------------------ Created Report Extract ------------------------------------------------------------------
def sap_export_created():
    try:
        session = get_sap_session()
    except Exception as e:
        print("ERROR: could not connect to SAP scripting:", e)
        traceback.print_exc()
        return

    try:
        print("Running SE16N -> MARA -> export ...")
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/path/to/file").text = "/path/to/file"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/path/to/file").text = "MARA"
        session.findById("wnd[0]/path/to/file").setFocus()
        session.findById("wnd[0]/path/to/file").caretPosition = 4
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/path/to/file").text = "SHELF_LIFE"
        session.findById("wnd[0]/path/to/file").text = ""  # as in your snippet
        # set your date field - from your snippet:
        session.findById("wnd[0]/path/to/file").text = yesterday
        session.findById("wnd[0]/path/to/file").text = "HAWA"
        session.findById("wnd[0]/path/to/file").setFocus()
        session.findById("wnd[0]/path/to/file").caretPosition = 10
        session.findById("wnd[0]/path/to/file").setFocus()
        session.findById("wnd[0]/path/to/file").press()
        session.findById("wnd[1]/path/to/file").text = "G*"
        # session.findById("wnd[1]/path/to/file").text = "F*"
        session.findById("wnd[1]/path/to/file").text = "P*"
        session.findById("wnd[1]/path/to/file").setFocus()
        session.findById("wnd[1]/path/to/file").caretPosition = 2
        session.findById("wnd[1]/path/to/file").press()
        # execute
        session.findById("wnd[0]/path/to/file").press()

        # select all rows, select MATNR column, context menu export (&XXL) as in your snippet
        time.sleep(LONG_SLEEP)
        session.findById("wnd[0]/path/to/file").currentCellRow = -1
        session.findById("wnd[0]/path/to/file").selectColumn("MATNR")
        session.findById("wnd[0]/path/to/file").contextMenu()
        session.findById("wnd[0]/path/to/file").selectContextMenuItem("&XXL")
        # export dialog - set path and filename then press Save (btn[11] per your snippet)
        time.sleep(SHORT_SLEEP)
        session.findById("wnd[1]/path/to/file").press()  # press Continue/path/to/file on the Excel export dialog - per your recorded script
        time.sleep(SHORT_SLEEP)
        session.findById("wnd[1]/path/to/file").text = str(CREATED_DIR)
        session.findById("wnd[1]/path/to/file").text = NEW_LISTINGS_FILENAME
        session.findById("wnd[1]/path/to/file").caretPosition = 23
        session.findById("wnd[1]/path/to/file").press()  # Save

        # Wait a bit for SAP to finish and for Excel to auto-open the exported file
        print("Waiting for Excel to open the exported file...")
        excel_app, wb = wait_for_excel_workbook(NEW_LISTINGS_FILENAME, timeout=WAIT_FOR_EXCEL_OPEN_SEC)
        if excel_app is None or wb is None:
            print(f"WARNING: Excel file {NEW_LISTINGS_FILENAME} didn't open within {WAIT_FOR_EXCEL_OPEN_SEC}s. Trying to locate file on disk.")
            # fallback: try opening file directly if path exists
            exported_path = os.path.join(CREATED_DIR, NEW_LISTINGS_FILENAME)
            if os.path.exists(exported_path):
                # Launch Excel and open it
                print("Opening workbook directly from disk:", exported_path)
                excel_app = Dispatch("Excel.Application")
                excel_app.Visible = True
                wb = excel_app.Workbooks.Open(exported_path)
            else:
                raise RuntimeError(f"Exported file not found at {exported_path}")

        # Copy column D to clipboard
        values = copy_column_from_workbook(wb, column_letter='D', start_row=2)
        set_clipboard_text("\r\n".join(values))
        print(f"Column D copied to clipboard from {NEW_LISTINGS_FILENAME} ({len(values)} values)")

        # prepare clipboard text as newline-separated (kept for manual use if needed)
        clipboard_text = "\r\n".join(values)
        set_clipboard_text(clipboard_text)
        print("Values copied to clipboard.")

        print("SAP Export for created report completed.")
        
# -------------------------- Open hierarchy transaction, run and export (NO paste popup) --------------------------
        print("Starting hierarchy transaction steps...")
        session.findById("wnd[0]/path/to/file").text = "/path/to/file"
        session.findById("wnd[0]").sendVKey(0)

        time.sleep(1)
        # Run/path/to/file and drill-down per your script
        session.findById("wnd[0]/path/to/file").press()
        time.sleep(1)
        # The script you provided then selects a row and double-clicks to drill into details:
        session.findById("wnd[1]/path/to/file").currentCellRow = 6
        session.findById("wnd[1]/path/to/file").selectedRows = "6"
        session.findById("wnd[1]/path/to/file").doubleClickCurrentCell()
        time.sleep(0.5)
        # press the MATKL button (from your snippet)
        session.findById("wnd[0]/path/to/file").press()

        # Your original snippet then configured some selection values
        session.findById("wnd[1]/path/to/file").press()
        time.sleep(0.2)
        session.findById("wnd[1]/path/to/file").press()
        time.sleep(0.2)
        session.findById("wnd[1]/path/to/file"
                         "tblSAPLALDBSINGLE/path/to/file").text = "F08010201"
        session.findById("wnd[1]/path/to/file"
                         "tblSAPLALDBSINGLE/path/to/file").text = "F77070101"
        session.findById("wnd[1]/path/to/file"
                         "tblSAPLALDBSINGLE/path/to/file").text = "F04010101"
        session.findById("wnd[1]/path/to/file"
                         "tblSAPLALDBSINGLE/path/to/file").text = "F77070101"
        session.findById("wnd[1]/path/to/file"
                         "tblSAPLALDBSINGLE/path/to/file").text = "F04010101"
        session.findById("wnd[1]/path/to/file"
                         "tblSAPLALDBSINGLE/path/to/file").text = "C27640201"
        session.findById("wnd[1]/path/to/file"
                         "tblSAPLALDBSINGLE/path/to/file").text = "F77070101"
        session.findById("wnd[1]/path/to/file"
                         "tblSAPLALDBSINGLE/path/to/file").text = "F77080110"
        time.sleep(0.2)
        session.findById("wnd[1]/path/to/file").press()
        session.findById("wnd[1]/path/to/file").press()

        # Back to main screen, run the report
        time.sleep(0.5)
        session.findById("wnd[0]/path/to/file").press()  # run

        time.sleep(1)
        # Select the LVL1_DESCR column (per your snippet) and export directly, no paste popup
        grid = session.findById("wnd[0]/path/to/file")
        grid.currentCellRow = -1
        grid.selectColumn("LVL1_DESCR")
        grid.contextMenu()
        grid.selectContextMenuItem("&XXL")
        time.sleep(0.5)

        # Confirm export and set file path/path/to/file – this is now the ONLY dialog we touch
        session.findById("wnd[1]/path/to/file").press()
        session.findById("wnd[1]/path/to/file").text = str(CREATED_DIR)
        session.findById("wnd[1]/path/to/file").text = HIERARCHY_FILENAME
        session.findById("wnd[1]/path/to/file").caretPosition = len(HIERARCHY_FILENAME)
        session.findById("wnd[1]/path/to/file").press()

        # Wait for Excel to auto-open the exported hierarchy workbook
        print("Waiting for hierarchy Excel to open...")
        excel_app2, wb2 = wait_for_excel_workbook(HIERARCHY_FILENAME, timeout=WAIT_FOR_EXCEL_OPEN_SEC)
        if excel_app2 is None or wb2 is None:
            print(f"WARNING: Hierarchy Excel {HIERARCHY_FILENAME} did not open automatically. Checking on disk.")
            exported_path2 = os.path.join(CREATED_DIR, HIERARCHY_FILENAME)
            if os.path.exists(exported_path2):
                excel_app2 = Dispatch("Excel.Application")
                try:
                    excel_app2.Visible = True
                except AttributeError:
                    print("Could not set Excel visibility, continuing anyway.")
                wb2 = excel_app2.Workbooks.Open(exported_path2)
            else:
                print("Could not find exported hierarchy file on disk. Continuing to final steps.")

        # Copy column A to clipboard
        values = copy_column_from_workbook(wb, column_letter='A', start_row=2)
        set_clipboard_text("\r\n".join(values))
        print(f"Column A copied to clipboard from {NEW_LISTINGS_FILENAME} ({len(values)} values)")

        # prepare clipboard text as newline-separated (kept for manual use if needed)
        clipboard_text = "\r\n".join(values)
        set_clipboard_text(clipboard_text)
        print("Values copied to clipboard.")


        print("Starting Purch Org steps...")
        session.findById("wnd[0]/path/to/file").text = "/path/to/file"
        session.findById("wnd[0]").sendVKey (0)
        session.findById("wnd[0]/path/to/file").text = "EINA"
        session.findById("wnd[0]/path/to/file").setFocus()
        session.findById("wnd[0]/path/to/file").caretPosition = 4
        session.findById("wnd[0]").sendVKey (0)
        session.findById("wnd[0]/path/to/file").setFocus()
        session.findById("wnd[0]/path/to/file").press()
        session.findById("wnd[1]/path/to/file").press()
        session.findById("wnd[1]/path/to/file").press()
        session.findById("wnd[0]/path/to/file").text = ""
        session.findById("wnd[0]/path/to/file").setFocus()
        session.findById("wnd[0]/path/to/file").caretPosition = 0
        session.findById("wnd[0]/path/to/file").press()
        session.findById("wnd[0]/path/to/file").setCurrentCell(-1,"MATKL")
        session.findById("wnd[0]/path/to/file").selectColumn ("MATKL")
        session.findById("wnd[0]/path/to/file").contextMenu()
        session.findById("wnd[0]/path/to/file").selectContextMenuItem ("&XXL")
        session.findById("wnd[1]/path/to/file").press()
        session.findById("wnd[1]/path/to/file").text = str(CREATED_DIR)
        session.findById("wnd[1]/path/to/file").text = INFORECORD_FILENAME
        session.findById("wnd[1]/path/to/file").press()

        # Wait for Excel to auto-open the exported hierarchy workbook
        print("Waiting for Excel to open...")
        excel_app2, wb2 = wait_for_excel_workbook(INFORECORD_FILENAME, timeout=WAIT_FOR_EXCEL_OPEN_SEC)
        if excel_app2 is None or wb2 is None:
            print(f"WARNING: Excel {INFORECORD_FILENAME} did not open automatically. Checking on disk.")
            exported_path2 = os.path.join(CREATED_DIR, INFORECORD_FILENAME)
            if os.path.exists(exported_path2):
                excel_app2 = Dispatch("Excel.Application")
                try:
                    excel_app2.Visible = True
                except AttributeError:
                    print("Could not set Excel visibility, continuing anyway.")
                wb2 = excel_app2.Workbooks.Open(exported_path2)
            else:
                print("Could not find exported hierarchy file on disk. Continuing to final steps.")

        # Copy column D to clipboard
        values = copy_column_from_workbook(wb2, column_letter='A', start_row=2)
        set_clipboard_text("\r\n".join(values))
        print(f"Column A copied to clipboard from {INFORECORD_FILENAME} ({len(values)} values)")

        # prepare clipboard text as newline-separated (kept for manual use if needed)
        clipboard_text = "\r\n".join(values)
        set_clipboard_text(clipboard_text)
        print("Values copied to clipboard.")

        session.findById("wnd[0]/path/to/file").text = "/path/to/file"
        session.findById("wnd[0]").sendVKey (0)
        session.findById("wnd[0]/path/to/file").text = "EINE"
        session.findById("wnd[0]/path/to/file").setFocus()
        session.findById("wnd[0]/path/to/file").caretPosition = 4
        session.findById("wnd[0]").sendVKey (0)
        session.findById("wnd[0]/path/to/file").setFocus()
        session.findById("wnd[0]/path/to/file").press()
        session.findById("wnd[1]/path/to/file").press()
        session.findById("wnd[1]/path/to/file").press()
        session.findById("wnd[0]/path/to/file").text = ""
        session.findById("wnd[0]/path/to/file").setFocus()
        session.findById("wnd[0]/path/to/file").caretPosition = 0
        session.findById("wnd[0]/path/to/file").press()
        session.findById("wnd[0]/path/to/file").currentCellRow = -1
        session.findById("wnd[0]/path/to/file").selectColumn ("INFNR")
        session.findById("wnd[0]/path/to/file").contextMenu()
        session.findById("wnd[0]/path/to/file").selectContextMenuItem ("&XXL")
        session.findById("wnd[1]/path/to/file").press()
        session.findById("wnd[1]/path/to/file").text = str(CREATED_DIR)
        session.findById("wnd[1]/path/to/file").text = PURCH_ORG_FILENAME
        session.findById("wnd[1]/path/to/file").press()

        # Wait for Excel to auto-open the exported hierarchy workbook
        print("Waiting for Excel to open...")
        excel_app2, wb2 = wait_for_excel_workbook(PURCH_ORG_FILENAME, timeout=WAIT_FOR_EXCEL_OPEN_SEC)
        if excel_app2 is None or wb2 is None:
            print(f"WARNING: Excel {PURCH_ORG_FILENAME} did not open automatically. Checking on disk.")
            exported_path2 = os.path.join(CREATED_DIR, PURCH_ORG_FILENAME)
            if os.path.exists(exported_path2):
                excel_app2 = Dispatch("Excel.Application")
                try:
                    excel_app2.Visible = True
                except AttributeError:
                    print("Could not set Excel visibility, continuing anyway.")
                wb2 = excel_app2.Workbooks.Open(exported_path2)
            else:
                print("Could not find exported hierarchy file on disk. Continuing to final steps.")

        # Copy column D to clipboard
        values = copy_column_from_workbook(wb2, column_letter='B', start_row=2)
        set_clipboard_text("\r\n".join(values))
        print(f"Column B copied to clipboard from {PURCH_ORG_FILENAME} ({len(values)} values)")

        # prepare clipboard text as newline-separated (kept for manual use if needed)
        clipboard_text = "\r\n".join(values)
        set_clipboard_text(clipboard_text)
        print("Values copied to clipboard.")

        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/path/to/file").text = "/path/to/file"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/path/to/file").text = "T024E"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/path/to/file").text = "PURCH_ORG_DE"
        session.findById("wnd[0]/path/to/file").setFocus()
        session.findById("wnd[0]/path/to/file").press()
        time.sleep(0.2)
        session.findById("wnd[1]/path/to/file").press()
        time.sleep(0.2)
        session.findById("wnd[1]/path/to/file").press()
        session.findById("wnd[0]/path/to/file").text = ""
        session.findById("wnd[0]/path/to/file").setFocus()
        session.findById("wnd[0]/path/to/file").caretPosition = 0
        session.findById("wnd[0]/path/to/file").press()
        session.findById("wnd[0]/path/to/file").setCurrentCell(-1, "EKOTX")
        session.findById("wnd[0]/path/to/file").selectColumn("EKOTX")
        session.findById("wnd[0]/path/to/file").contextMenu()
        session.findById("wnd[0]/path/to/file").selectContextMenuItem("&XXL")
        session.findById("wnd[1]/path/to/file").press()
        session.findById("wnd[1]/path/to/file").text = str(CREATED_DIR)
        session.findById("wnd[1]/path/to/file").text = PURCH_ORG_DESC_FILENAME
        session.findById("wnd[1]/path/to/file").caretPosition = 11
        session.findById("wnd[1]/path/to/file").press()

    except Exception as e:
        print("ERROR during automation:")
        traceback.print_exc()
        # try to close excel to avoid leaving things open
        try:
            subprocess.run(
                ["taskkill", "/path/to/file", "/path/to/file", "EXCEL.EXE"],
                check=False,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
        except Exception:
            pass

def create_reference_file():
    wait_for_file_stable(INFORECORD, stable_seconds=3)
    wait_for_file_stable(PURCH_ORG_PATH, stable_seconds=3)
    wait_for_file_stable(PURCH_ORG_DESC_PATH, stable_seconds=3)

    output_path = CREATED_DIR / "Reference.xlsx"
    if output_path.exists():
        try:
            output_path.unlink()
        except Exception:
            pass

    # --- Read source files ---
    info_df = pd.read_excel(INFORECORD, sheet_name=0)
    po_df = pd.read_excel(PURCH_ORG_PATH, sheet_name=0)
    po_desc_df = pd.read_excel(PURCH_ORG_DESC_PATH, sheet_name=0)

    # --- Validate required columns ---
    if "Purchasing Info Rec." not in info_df.columns or "Article" not in info_df.columns:
        raise KeyError(
            "Inforecord file must contain 'Purchasing Info Rec.' and 'Article' columns"
        )

    # --- Prepare Purch Org lookup ---
    po_key_col = po_df.columns[0]
    po_val_col = po_df.columns[1]
    po_df[po_key_col] = pd.to_numeric(po_df[po_key_col], errors="coerce")

    po_map = dict(
        zip(
            po_df[po_key_col].dropna().astype(int),
            po_df[po_val_col]
        )
    )

    # --- Prepare Purch Org Description lookup ---
    desc_key_col = po_desc_df.columns[0]
    desc_val_col = po_desc_df.columns[1]

    desc_map = dict(
        zip(
            po_desc_df[desc_key_col].astype(str).str.strip(),
            po_desc_df[desc_val_col]
        )
    )

    # --- Build Reference dataframe ---
    ref = pd.DataFrame({
        "Purchasing Info Rec.": info_df["Purchasing Info Rec."],
        "Article": info_df["Article"],
    })

    ref_key = pd.to_numeric(ref["Purchasing Info Rec."], errors="coerce")
    ref["Purch Org"] = ref_key.map(
        lambda x: po_map.get(int(x)) if pd.notna(x) else None
    )

    ref["Purch Org"] = ref["Purch Org"].fillna("0000")
    ref["Purch Org Description"] = ref["Purch Org"].astype(str).str.strip().map(desc_map).fillna(NOT_FOUND_TEXT)

    # --- Write Reference.xlsx ---
    ref.to_excel(output_path, index=False)

    # --- Delete source files ONLY after success ---
    for f in [INFORECORD, PURCH_ORG_PATH, PURCH_ORG_DESC_PATH]:
        try:
            if f.exists():
                f.unlink()
                print(f"Deleted source file: {f.name}")
        except Exception as e:
            print(f"Could not delete {f.name}: {e}")

    print("Reference.xlsx created and source files deleted.")
    return output_path

def excel_formatting_created():
    """
    Created report formatting (NO Excel COM).
    - Inserts Purch Org + Description into C:D (lookup from Reference.xlsx by Article)
      * Purch Org default = "0000" when missing/path/to/file
    - Inserts Hierarchy fields into H:I (lookup from Hierarchy by key)
    - Applies light formatting
    - Saves report and moves files to dated folder
    """

    # --- Make sure Excel isn't holding file locks
    kill_excel()
    time.sleep(1)

    date_str = extract_date_from_filename(NEW_LISTINGS_PATH)
    output_path = CREATED_DIR / f"new_listings_{date_str}_report.xlsx"
    reference_path = CREATED_DIR / "Reference.xlsx"

    # --- stability checks (OneDrive + SAP)
    for p in (NEW_LISTINGS_PATH, HIERARCHY_PATH, reference_path):
        if not p.exists():
            raise FileNotFoundError(f"Missing file: {p}")
        wait_for_file_stable(p, timeout=180, stable_seconds=3)

    print("Formatting output will be:", output_path)

    # -----------------------
    # Helpers
    # -----------------------
    def norm_key(v):
        if v is None:
            return ""
        s = str(v).strip()
        return s[:-2] if s.endswith(".0") else s

    def find_col_by_header(ws, name, header_row=1, max_cols=250):
        target = name.strip().lower()
        for c in range(1, max_cols + 1):
            v = ws.cell(row=header_row, column=c).value
            if v is not None and str(v).strip().lower() == target:
                return c
        return None

    def last_row_by_col(ws, col_idx, header_row=1):
        r = ws.max_row
        while r > header_row:
            v = ws.cell(row=r, column=col_idx).value
            if v is not None and str(v).strip() != "":
                return r
            r -= 1
        return header_row

    # -----------------------
    # Load workbooks (openpyxl)
    # -----------------------
    wb_new = load_workbook(NEW_LISTINGS_PATH)
    ws_new = wb_new.worksheets[0]

    wb_ref = load_workbook(reference_path, data_only=True)
    ws_ref = wb_ref.worksheets[0]

    wb_h = load_workbook(HIERARCHY_PATH, data_only=True)
    ws_h = wb_h.worksheets[0]

    # -----------------------
    # 1) Insert Purch Org cols in NEW (C:D)
    # -----------------------
    ws_new.insert_cols(3, amount=2)
    ws_new.cell(row=1, column=3).value = "Purch Org"
    ws_new.cell(row=1, column=4).value = "Purch Org Description"

    article_col_new = find_col_by_header(ws_new, "Article")
    if not article_col_new:
        raise RuntimeError("Could not find 'Article' column in New Listings.")

    last_new = last_row_by_col(ws_new, article_col_new, header_row=1)

    # -----------------------
    # 2) Build Reference lookup: Article -> (Purch Org, Purch Org Desc)
    # -----------------------
    ref_article_col = find_col_by_header(ws_ref, "Article")
    ref_org_col     = find_col_by_header(ws_ref, "Purch Org")
    ref_desc_col    = find_col_by_header(ws_ref, "Purch Org Description")

    if not (ref_article_col and ref_org_col and ref_desc_col):
        raise RuntimeError("Reference.xlsx missing headers: Article / Purch Org / Purch Org Description")

    last_ref = last_row_by_col(ws_ref, ref_article_col, header_row=1)

    ref_lookup = {}
    for r in range(2, last_ref + 1):
        k = norm_key(ws_ref.cell(row=r, column=ref_article_col).value)
        if not k:
            continue
        org = ws_ref.cell(row=r, column=ref_org_col).value
        desc = ws_ref.cell(row=r, column=ref_desc_col).value
        ref_lookup[k] = (org, desc)

    # Fill C:D
    for r in range(2, last_new + 1):
        k = norm_key(ws_new.cell(row=r, column=article_col_new).value)
        org, desc = ref_lookup.get(k, (None, None))

        # your rule: if reference has no info for Purch Org -> "0000"
        if org is None or str(org).strip() == "":
            org = "0000"

        if desc is None or str(desc).strip() == "":
            desc = NOT_FOUND_TEXT

        ws_new.cell(row=r, column=3).value = org
        ws_new.cell(row=r, column=4).value = desc

    # -----------------------
    # 3) Hierarchy lookup (key col H; return cols per MAP_RETURN_COLS)
    # -----------------------
    h_key_idx = col_letter_to_index("H")
    ret_idxs = {name: col_letter_to_index(col) for name, col in MAP_RETURN_COLS.items()}

    last_h = last_row_by_col(ws_h, h_key_idx, header_row=1)

    hier_lookup = {}
    for r in range(2, last_h + 1):
        k = norm_key(ws_h.cell(row=r, column=h_key_idx).value)
        if not k:
            continue
        row_dict = {}
        for header, cidx in ret_idxs.items():
            row_dict[header] = ws_h.cell(row=r, column=cidx).value
        hier_lookup[k] = row_dict

    # Insert 2 cols at H -> new columns become H:I
    insert_at = col_letter_to_index("H")
    ws_new.insert_cols(insert_at, amount=2)

    out_headers = list(MAP_RETURN_COLS.keys())[:2]
    ws_new.cell(row=1, column=insert_at).value = out_headers[0]
    ws_new.cell(row=1, column=insert_at + 1).value = out_headers[1]

    # Keys in NEW column F (your current logic)
    key_col_new = col_letter_to_index("F")
    last_new2 = last_row_by_col(ws_new, key_col_new, header_row=1)

    for r in range(2, last_new2 + 1):
        k = norm_key(ws_new.cell(row=r, column=key_col_new).value)
        row = hier_lookup.get(k, {})

        v1 = row.get(out_headers[0], "0000")
        v2 = row.get(out_headers[1], NOT_FOUND_TEXT)

        if v1 is None or str(v1).strip() == "":
            v1 = "0000"
        if v2 is None or str(v2).strip() == "":
            v2 = NOT_FOUND_TEXT

        ws_new.cell(row=r, column=insert_at).value = v1
        ws_new.cell(row=r, column=insert_at + 1).value = v2

    # -----------------------
    # 4) Basic formatting (fast)
    # -----------------------
    header_fill = PatternFill("solid", fgColor="BFBFBF")
    thin = Side(border_style="thin", color=14277081)
    header_border = Border(top=thin, left=thin, right=thin, bottom=thin)
    header_font = Font(bold=True)
    header_align = Alignment(horizontal="center", vertical="center")

    max_row = ws_new.max_row
    max_col = ws_new.max_column

    # ---- Format header row ----
    for c in range(1, max_col + 1):
        cell = ws_new.cell(row=1, column=c)
        cell.font = header_font
        cell.alignment = header_align
        cell.fill = header_fill
        cell.border = header_border

    # ---- Freeze header row + apply autofilter ----
    ws_new.freeze_panes = "A2"
    ws_new.auto_filter.ref = f"A1:{get_column_letter(max_col)}{max_row}"

    # ---- Autofit-style column width ----
    for c in range(1, max_col + 1):
        col_letter = get_column_letter(c)
        max_len = 0
        for r in range(1, min(max_row, 220) + 1):
            v = ws_new.cell(row=r, column=c).value
            if v is None:
                continue
            max_len = max(max_len, len(str(v)))
        ws_new.column_dimensions[col_letter].width = min(max(10, max_len + 2), 45)

    # ------------------------------------------------------
    # Apply Custom Formatting: highlight "999" values yellow
    # ------------------------------------------------------
    def apply_custom_formatting(ws):
        # ---- Locate "Min. Rem. Shelf Life" column ----
        target_col = None
        max_col = ws.max_column

        for col in range(1, max_col + 1):
            value = ws.cell(row=1, column=col).value
            if value and str(value).strip().lower() in [
                "min. rem. shelf life",
                "min rem shelf life",
            ]:
                target_col = col
                break

        if not target_col:
            print("Column not found: Min. Rem. Shelf Life")
            return

        # ---- Highlight 999 values ----
        yellow = PatternFill("solid", fgColor="FFFF00")
        max_row = ws.max_row

        for r in range(2, max_row + 1):
            v = ws.cell(row=r, column=target_col).value
            if v is None:
                continue

            try:
                v_clean = str(int(float(v)))
            except:
                v_clean = str(v).strip()

            if v_clean == "999":
                ws.cell(row=r, column=target_col).fill = yellow

    # Run custom formatting
    apply_custom_formatting(ws_new)

    # -----------------------
    # 5) Save + move
    # -----------------------
    if output_path.exists():
        output_path.unlink()

    wb_new.save(output_path)

    dest_folder = CREATED_DIR / date_str
    dest_folder.mkdir(parents=True, exist_ok=True)

    move_files_to_folder([HIERARCHY_PATH, NEW_LISTINGS_PATH, output_path], dest_folder)

    final_path = dest_folder / output_path.name
    print("Created report:", final_path)
    return final_path

def prepare_dataframe_for_access_created(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = df.columns.str.strip()

    rename_map = {
        "Min. Rem. Shelf Life": "Min Rem Shelf Life",
    }
    df = df.rename(columns=rename_map)

    required_cols = [
        "Article",
        "Article description",
        "Purch Org",
        "Purch Org Description",
        "Material Type",
        "Merchandise Category",
        "Created On",
        "BMC Description",
        "Divisional Buyer",
        "Min Rem Shelf Life",
        "Total shelf life",
        "X-site artl status",
    ]

    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise KeyError(f"Missing required columns: {missing}")

    df = df[required_cols]

    df["Created On"] = pd.to_datetime(df["Created On"], errors="coerce").dt.date
    df["Min Rem Shelf Life"] = pd.to_numeric(df["Min Rem Shelf Life"], errors="coerce")
    df["X-site artl status"] = df["X-site artl status"].astype(str)

    df = df.where(pd.notna(df), None)
    return df

def append_to_access_created(report_path, db_path, table_name):
    print("🚀 append_to_access_created() started")

    report_path = Path(report_path)

    if not report_path.exists():
        raise FileNotFoundError(f"Report file not found: {report_path}")

    print("Reading Excel report...")
    df = pd.read_excel(report_path, sheet_name=0)
    df = prepare_dataframe_for_access_created(df)

    print(f"Columns ready for Access: {list(df.columns)}")
    print(f"Rows in report: {len(df)}")

    conn_str = r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=" + str(db_path)
    conn = pyodbc.connect(conn_str)
    cursor = conn.cursor()

    # ---- Check for existing records (optional but good) ----
    existing_records = set()
    try:
        cursor.execute(
            f"SELECT [Article], [Article description], [Created On] FROM [{table_name}]"
        )
        for row in cursor.fetchall():
            if all(x is not None for x in row):
                existing_records.add((
                    str(row[0]).strip(),
                    str(row[1]).strip().lower(),
                    row[2]
                ))
        print(f"Existing records found: {len(existing_records)}")
    except Exception as e:
        print(f"Could not check existing records: {e}")

    # ---- Remove duplicates ----
    df["record_key"] = list(zip(
        df["Article"],
        df["Article description"].str.lower(),
        df["Created On"]
    ))

    original_count = len(df)
    df = df[~df["record_key"].isin(existing_records)]
    df = df.drop(columns=["record_key"])

    print(f"Filtered out {original_count - len(df)} duplicates")

    if df.empty:
        print("No new records to insert.")
        cursor.close()
        conn.close()
        return

    # ---- Build insert SQL ----
    columns = df.columns.tolist()
    col_sql = ", ".join(f"[{c}]" for c in columns)
    placeholders = ", ".join("?" for _ in columns)

    insert_sql = f"INSERT INTO [{table_name}] ({col_sql}) VALUES ({placeholders})"

    data = [[None if pd.isna(x) else x for x in row] for row in df.values.tolist()]

    print("Sample SQL:", insert_sql)
    print("Sample row:", data[0])

    inserted_rows = 0
    for i, row in enumerate(data, 1):
        try:
            cursor.execute(insert_sql, row)
            inserted_rows += 1
        except Exception as e:
            print(f"Row {i} failed:", e)

    conn.commit()
    print(f"Append completed. {inserted_rows}/path/to/file rows added.")

    cursor.close()
    conn.close()

def convert_sheet_to_table(filepath, sheet_index=0, table_name="LoadTable"):
    """Create an Excel table from the used range on the target sheet."""

    wb = load_workbook(filepath)
    ws = wb.worksheets[sheet_index]

    # 1) Remove standalone AutoFilter to prevent conflicts with a table UI
    try:
        if ws.auto_filter and ws.auto_filter.ref:
            ws.auto_filter.ref = None
    except Exception:
        pass

    # 2) Find last used column from contiguous non-empty headers in row 1
    last_col = 0
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=1, column=c).value
        if v is not None and str(v).strip() != "":
            last_col = c
        else:
            if last_col > 0:  # stop at first gap after seeing headers
                break

    if last_col == 0:
        print("No headers found in row 1. Cannot create a table.")
        wb.save(filepath)
        try:
            wb.close()
        except Exception:
            pass
        return

    # 3) Find last used row by scanning 1..last_col for the first empty row gap
    last_row = 1
    for r in range(2, ws.max_row + 1):
        row_has_data = any(
            (ws.cell(row=r, column=c).value is not None 
             and str(ws.cell(row=r, column=c).value).strip() != "")
            for c in range(1, last_col + 1)
        )
        if row_has_data:
            last_row = r
        else:
            break

    if last_row < 2:
        print("Not enough data to create table (need header + at least 1 row).")
        wb.save(filepath)
        try:
            wb.close()
        except Exception:
            pass
        return

    # 4) Header must be unmerged and non-empty
    for merged in list(ws.merged_cells.ranges):
        if merged.min_row == 1 or merged.max_row == 1:
            ws.unmerge_cells(str(merged))
    for c in range(1, last_col + 1):
        cell = ws.cell(row=1, column=c)
        if cell.value is None or str(cell.value).strip() == "":
            cell.value = f"Column{c}"

    # 5) Build table ref: A1 : <last_col><last_row>
    ref = f"A1:{get_column_letter(last_col)}{last_row}"

    # 6) Ensure unique table name
    existing_names = {t.displayName for t in ws._tables}
    name = table_name
    i = 2
    while name in existing_names:
        name = f"{table_name}{i}"
        i += 1

    # 7) Create table
    tbl = Table(displayName=name, ref=ref)
    style = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    tbl.tableStyleInfo = style
    ws.add_table(tbl)

    wb.save(filepath)
    try:
        wb.close()
    except Exception:
        pass
    print(f"Created table '{name}' with range {ref}")

def copy_yesterday_file():
    date_str = yesterday
    filename = f"new_listings_{date_str}_report.xlsx"

    source_folder = SOURCE_ROOT / date_str
    source_file = source_folder / filename

    print(f"Looking for: {source_file}")

    if not source_file.exists():
        print("File not found.")

    DEST_FOLDER.mkdir(parents=True, exist_ok=True)

    destination = DEST_FOLDER / NEW_NAME
    shutil.copy2(source_file, destination)

    print(f"Copied successfully t/path/to/file")

    # 🔥 NEW STEP
    convert_sheet_to_table(destination)

#---------------------------------------------------------------------------- Extended Data Report Script -------------------------------------------------------------------
def sap_export_extended():
    try:
        session = get_sap_session()
    except Exception as e:
        print("ERROR: could not connect to SAP scripting:", e)
        traceback.print_exc()
        return

    print("Exporting Extended Data ...")
    # Maximize the SAP window
    session.findById("wnd[0]").maximize()

    # Navigate to the transaction
    session.findById("wnd[0]/path/to/file").text = "/path/to/file"
    session.findById("wnd[0]").sendVKey(0)

    # Set wizard action
    session.findById("wnd[1]/path/to/file").key = "Extend"

    # Set date range
    session.findById("wnd[1]/path/to/file").text = yesterday
    session.findById("wnd[1]/path/to/file").text = yesterday
    # Set Department Variable
    session.findById("wnd[1]/path/to/file").press()
    session.findById("wnd[2]/path/to/file").text = "G*"
    session.findById("wnd[2]/path/to/file").text = "P*"
    session.findById("wnd[2]/path/to/file").setFocus()
    session.findById("wnd[2]/path/to/file").caretPosition = 2
    session.findById("wnd[2]").sendVKey (8)

    # Tick the checkbox
    session.findById("wnd[1]/path/to/file").selected = True

    # Focus and set caret position
    session.findById("wnd[1]/path/to/file").setFocus()
    session.findById("wnd[1]/path/to/file").caretPosition = 10

    # Execute
    session.findById("wnd[1]").sendVKey(8)

    session.findById("wnd[0]/path/to/file").pressToolbarContextButton ("&MB_VARIANT")
    session.findById("wnd[0]/path/to/file").selectContextMenuItem ("&LOAD")
    session.findById("wnd[1]/path/to/file").currentCellRow = 3
    session.findById("wnd[1]/path/to/file").firstVisibleRow = 1
    session.findById("wnd[1]/path/to/file").selectedRows = "3"
    session.findById("wnd[1]/path/to/file").clickCurrentCell()
    session.findById("wnd[0]/path/to/file").currentCellRow = -1
    session.findById("wnd[0]/path/to/file").selectColumn ("ARTNR")
    session.findById("wnd[0]/path/to/file").contextMenu()
    session.findById("wnd[0]/path/to/file").selectContextMenuItem ("&XXL")
    session.findById("wnd[1]/path/to/file").press()
    time.sleep(2)  # wait for the file dialog to appear
    session.findById("wnd[1]/path/to/file").text = str(EXTEND_DIR)
    session.findById("wnd[1]/path/to/file").text = f'Extended_To_DC_Shelf_Life_{yesterday}.xlsx'
    session.findById("wnd[1]/path/to/file").caretPosition = 30
    session.findById("wnd[1]/path/to/file").press()

    # Wait a bit for SAP to finish and for Excel to auto-open the exported file
    print("Waiting for Excel to open the exported file...")
    excel_app, wb = wait_for_excel_workbook(EXTENDED_DATA_FILENAME, timeout=WAIT_FOR_EXCEL_OPEN_SEC)
    if excel_app is None or wb is None:
        print(f"WARNING: Excel file {EXTENDED_DATA_FILENAME} didn't open within {WAIT_FOR_EXCEL_OPEN_SEC}s. Trying to locate file on disk.")
        # fallback: try opening file directly if path exists
        exported_path = os.path.join(EXTEND_DIR, EXTENDED_DATA_FILENAME)
        if os.path.exists(exported_path):
            # Launch Excel and open it
            print("Opening workbook directly from disk:", exported_path)
            excel_app = Dispatch("Excel.Application")
            excel_app.Visible = True
            wb = excel_app.Workbooks.Open(exported_path)
        else:
            raise RuntimeError(f"Exported file not found at {exported_path}")

    # Copy column D to clipboard
    values = copy_column_from_workbook(wb, column_letter='A', start_row=2)
    set_clipboard_text("\r\n".join(values))
    print(f"Column A copied to clipboard from {EXTENDED_DATA_FILENAME} ({len(values)} values)")

    # prepare clipboard text as newline-separated (kept for manual use if needed)
    clipboard_text = "\r\n".join(values)
    set_clipboard_text(clipboard_text)
    print("Values copied to clipboard.")

    print("Extracting Shelf Life Data...")
    session.findById("wnd[0]/path/to/file").text = "/path/to/file"
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/path/to/file").text = "MARA"
    session.findById("wnd[0]/path/to/file").caretPosition = 4
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/path/to/file").text = "SHELF_LIFE"
    session.findById("wnd[0]/path/to/file").setFocus()
    session.findById("wnd[0]/path/to/file").press()
    session.findById("wnd[1]/path/to/file").press()
    session.findById("wnd[1]/path/to/file").press()
    session.findById("wnd[0]/path/to/file").text = ""
    session.findById("wnd[0]/path/to/file").setFocus()
    session.findById("wnd[0]/path/to/file").caretPosition = 0
    session.findById("wnd[0]/path/to/file").press()
    session.findById("wnd[0]/path/to/file").currentCellRow = -1
    session.findById("wnd[0]/path/to/file").selectColumn ("MATNR")
    session.findById("wnd[0]/path/to/file").contextMenu()
    session.findById("wnd[0]/path/to/file").selectContextMenuItem ("&XXL")
    session.findById("wnd[1]/path/to/file").press()
    session.findById("wnd[1]/path/to/file").text = str(EXTEND_DIR)
    session.findById("wnd[1]/path/to/file").text = f'Shelf_Life_Data_{yesterday}.xlsx'
    session.findById("wnd[1]/path/to/file").caretPosition = 20
    session.findById("wnd[1]/path/to/file").press()

    # Wait a bit for SAP to finish and for Excel to auto-open the exported file
    print("Waiting for Excel to open the exported file...")
    excel_app, wb = wait_for_excel_workbook(SOURCE_FILE_EXTEND, timeout=WAIT_FOR_EXCEL_OPEN_SEC)
    if excel_app is None or wb is None:
        print(f"WARNING: Excel file {SOURCE_FILE_EXTEND} didn't open within {WAIT_FOR_EXCEL_OPEN_SEC}s. Trying to locate file on disk.")
        # fallback: try opening file directly if path exists
        exported_path = os.path.join(EXTEND_DIR, SOURCE_FILE_EXTEND)
        if os.path.exists(exported_path):
            # Launch Excel and open it
            print("Opening workbook directly from disk:", exported_path)
            excel_app = Dispatch("Excel.Application")
            excel_app.Visible = True
            wb = excel_app.Workbooks.Open(exported_path)
        else:
            raise RuntimeError(f"Exported file not found at {exported_path}")

    # Copy column D to clipboard
    values = copy_column_from_workbook(wb, column_letter='D', start_row=2)
    set_clipboard_text("\r\n".join(values))
    print(f"Column D copied to clipboard from {SOURCE_FILE_EXTEND} ({len(values)} values)")

    # prepare clipboard text as newline-separated (kept for manual use if needed)
    clipboard_text = "\r\n".join(values)
    set_clipboard_text(clipboard_text)
    print("Values copied to clipboard.")

    print("Extracting Hierarchy Data...")
    session.findById("wnd[0]/path/to/file").text = "/path/to/file"
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(1)
    # Run/path/to/file and drill-down per your script
    session.findById("wnd[0]/path/to/file").press()
    time.sleep(1)
    # The script you provided then selects a row and double-clicks to drill into details:
    session.findById("wnd[1]/path/to/file").currentCellRow = 6
    session.findById("wnd[1]/path/to/file").selectedRows = "6"
    session.findById("wnd[1]/path/to/file").doubleClickCurrentCell()
    time.sleep(0.5)
    # press the MATKL button (from your snippet)
    session.findById("wnd[0]/path/to/file").press()

    # Your original snippet then configured some selection values
    session.findById("wnd[1]/path/to/file").press()
    time.sleep(0.2)
    session.findById("wnd[1]/path/to/file").press()
    time.sleep(0.2)
    session.findById("wnd[1]/path/to/file"
                        "tblSAPLALDBSINGLE/path/to/file").text = "F08010201"
    session.findById("wnd[1]/path/to/file"
                        "tblSAPLALDBSINGLE/path/to/file").text = "F77070101"
    session.findById("wnd[1]/path/to/file"
                        "tblSAPLALDBSINGLE/path/to/file").text = "F04010101"
    session.findById("wnd[1]/path/to/file"
                        "tblSAPLALDBSINGLE/path/to/file").text = "F77070101"
    session.findById("wnd[1]/path/to/file"
                        "tblSAPLALDBSINGLE/path/to/file").text = "F04010101"
    session.findById("wnd[1]/path/to/file"
                        "tblSAPLALDBSINGLE/path/to/file").text = "C27640201"
    session.findById("wnd[1]/path/to/file"
                        "tblSAPLALDBSINGLE/path/to/file").text = "F77070101"
    session.findById("wnd[1]/path/to/file"
                        "tblSAPLALDBSINGLE/path/to/file").text = "F77080110"
    time.sleep(0.2)
    session.findById("wnd[1]/path/to/file").press()
    session.findById("wnd[1]/path/to/file").press()

    # Back to main screen, run the report
    time.sleep(0.5)
    session.findById("wnd[0]/path/to/file").press()  # run

    time.sleep(1)
    # Select the LVL1_DESCR column (per your snippet) and export directly, no paste popup
    grid = session.findById("wnd[0]/path/to/file")
    grid.currentCellRow = -1
    grid.selectColumn("LVL1_DESCR")
    grid.contextMenu()
    grid.selectContextMenuItem("&XXL")
    time.sleep(0.5)
    # Confirm export and set file path/path/to/file – this is now the ONLY dialog we touch
    session.findById("wnd[1]/path/to/file").press()
    session.findById("wnd[1]/path/to/file").text = str(EXTEND_DIR)
    session.findById("wnd[1]/path/to/file").text = F'Hierarchy_{yesterday}.xlsx'
    session.findById("wnd[1]/path/to/file").caretPosition = len(F'Hierarchy_{yesterday}.xlsx')
    session.findById("wnd[1]/path/to/file").press()
    time.sleep(7)  # wait for file to save
    
    print("SAP Export for extended report completed.")
    
def excel_formatting_extended():
    kill_excel()
    time.sleep(2)

    # Resolve paths
    date_str = extract_date_from_filename(EXTEND_PATH)
    target_path = Path(EXTEND_PATH).resolve()
    hierarchy_path = Path(HIERARCHY_FILE_EXTEND).resolve()
    source_path = Path(SOURCE_FILE_EXTEND_PATH).resolve()

    # Ensure files exist + stable (OneDrive)
    for p in [target_path, hierarchy_path, source_path]:
        if not p.exists():
            raise FileNotFoundError(f"Missing file: {p}")
        wait_for_file_stable(p, timeout=180, stable_seconds=3)

    # --- Read hierarchy + source WITHOUT Excel COM ---
    # Hierarchy: we need H..L, and we use column H as key, column L as value
    try:
        hier_df = pd.read_excel(
            hierarchy_path,
            sheet_name=0,
            usecols="H:L",
            engine="openpyxl"
        )
    except Exception as e:
        raise RuntimeError(
            f"Hierarchy file couldn't be read as a real xlsx. "
            f"Open it manually once in Excel and Save As .xlsx, then rerun.\nFile: {hierarchy_path}"
        ) from e

    # Build lookup: key = col H, value = col L (5th col in H..L)
    hier_df = hier_df.dropna(subset=[hier_df.columns[0]])
    hier_keys = hier_df.iloc[:, 0].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    hier_vals = hier_df.iloc[:, 4].fillna("").astype(str)
    lookup_dict = dict(zip(hier_keys, hier_vals))

    # Source: A..H, key=A, created_on=E, min=F, total=G, status=H
    try:
        src_df = pd.read_excel(
            source_path,
            sheet_name=0,
            usecols="A:H",
            engine="openpyxl"
        )
    except Exception as e:
        raise RuntimeError(
            f"Source file couldn't be read as a real xlsx. "
            f"Open it manually once in Excel and Save As .xlsx, then rerun.\nFile: {source_path}"
        ) from e

    src_df = src_df.dropna(subset=[src_df.columns[0]])
    src_key = src_df.iloc[:, 0].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)

    lookup_created_on = dict(zip(src_key, src_df.iloc[:, 4]))
    lookup_min_rem = dict(zip(src_key, pd.to_numeric(src_df.iloc[:, 5], errors="coerce").fillna(0).astype(int)))
    lookup_total_shelf_life = dict(zip(src_key, pd.to_numeric(src_df.iloc[:, 6], errors="coerce").fillna(0).astype(int)))

    # status -> 2-digit string
    stat_series = src_df.iloc[:, 7]
    stat_clean = (
        pd.to_numeric(stat_series, errors="coerce")
        .fillna(stat_series.astype(str).str.strip())
    )
    def fmt_status(v):
        try:
            return str(int(float(v))).zfill(2)
        except Exception:
            s = str(v).strip()
            return s.zfill(2) if s else ""
    lookup_xsite_status = dict(zip(src_key, [fmt_status(v) for v in stat_clean]))

    # --- Now open ONLY the target workbook with xlwings and write results ---
    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False

    wb_target = None
    try:
        wb_target = app.books.open(str(target_path), update_links=False, read_only=False)
        ws_target = wb_target.sheets[0]

        # Rename header F1
        if ws_target.range("F1").value == "Date Created":
            ws_target.range("F1").value = "Date Extended"

        # Write headers H..L
        headers = NEW_HEADERS_EXTEND
        start_col = 8  # H
        ws_target.range((1, start_col), (1, start_col + len(headers) - 1)).value = headers

        # Fill H from hierarchy lookup using target col D keys
        last_row_d = ws_target.range("D" + str(ws_target.cells.last_cell.row)).end("up").row
        d_vals = ws_target.range(f"D2:D{last_row_d}").value
        if d_vals is None:
            d_vals = []
        elif isinstance(d_vals, (str, int, float, datetime)):
            d_vals = [d_vals]

        out_h = []
        for v in d_vals:
            if v is None:
                out_h.append([""])
                continue
            k = str(int(v)) if isinstance(v, float) and v.is_integer() else str(v)
            k = k.strip()
            if k.endswith(".0"):
                k = k[:-2]
            out_h.append([lookup_dict.get(k, "")])

        if out_h:
            ws_target.range(f"H2:H{1 + len(out_h)}").value = out_h

        # Fill I..L from source lookups using target col A keys
        last_row_a = ws_target.range("A" + str(ws_target.cells.last_cell.row)).end("up").row
        a_vals = ws_target.range(f"A2:A{last_row_a}").value
        if a_vals is None:
            a_vals = []
        elif isinstance(a_vals, (str, int, float, datetime)):
            a_vals = [a_vals]

        out_i, out_j, out_k, out_l = [], [], [], []
        for v in a_vals:
            if v is None:
                kk = ""
            else:
                kk = str(int(v)) if isinstance(v, float) and hasattr(v, "is_integer") and v.is_integer() else str(v)
                kk = kk.strip()
                if kk.endswith(".0"):
                    kk = kk[:-2]

            out_i.append([lookup_created_on.get(kk, "")])
            out_j.append([lookup_min_rem.get(kk, 0)])
            out_k.append([lookup_total_shelf_life.get(kk, 0)])
            out_l.append([lookup_xsite_status.get(kk, "")])

        if out_i:
            ws_target.range(f"I2:I{len(out_i) + 1}").value = out_i
            ws_target.range(f"J2:J{len(out_j) + 1}").value = out_j
            ws_target.range(f"K2:K{len(out_k) + 1}").value = out_k
            ws_target.range(f"L2:L{len(out_l) + 1}").value = out_l

        ws_target.range("I:I").number_format = "dd/path/to/file"
        ws_target.range("J:J").number_format = "0"
        ws_target.range("K:K").number_format = "0"
        ws_target.range("L:L").number_format = "00"

        # IMPORTANT: pass app, not undefined "excel"
        apply_basic_formatting_simple(ws_target, app, header_row=1)

        wb_target.save()

    finally:
        try:
            if wb_target is not None:
                wb_target.close()
        except Exception:
            pass
        try:
            app.display_alerts = True
        except Exception:
            pass
        try:
            app.quit()
        except Exception:
            pass

    # Move files into dated folder
    time.sleep(2)
    dest_folder = EXTEND_DIR / date_str
    dest_folder.mkdir(parents=True, exist_ok=True)

    move_files_to_folder([target_path, source_path, hierarchy_path], dest_folder)

    final_moved_report_path = dest_folder / target_path.name
    wait_for_file(final_moved_report_path, timeout=60)
    return final_moved_report_path

def prepare_dataframe_for_access_extended(df: pd.DataFrame) -> pd.DataFrame:
    # Remove extra spaces
    df.columns = df.columns.str.strip()

    # Rename columns uniquely if duplicates exist
    # Here we assume the first "Date Created" is the original creation date of the record,
    # the second is the date the shelf life was extended.
    cols = list(df.columns)
    date_created_count = cols.count("Date Created")
    if date_created_count > 1:
        new_cols = []
        dc_seen = 0
        for c in cols:
            if c == "Date Created":
                if dc_seen == 0:
                    new_cols.append("Date Extended")  # first occurrence
                else:
                    new_cols.append("Date Created")   # second occurrence
                dc_seen += 1
            else:
                new_cols.append(c)
        df.columns = new_cols

    # Rename other columns to match Access schema
    rename_map = {
        "Min. Rem. Shelf Life": "Min Rem Shelf Life",
        "Mdse Catgry Desc.": "Mdse Catgry Desc",
        # "Divisional Buyer" already matches
    }
    df = df.rename(columns=rename_map)

    # Define required column order
    required_cols = [
        "Article",
        "Article Description",
        "Material Type",
        "Merchandise Category",
        "Mdse Catgry Desc",
        "Date Extended",
        "Site",
        "Divisional Buyer",
        "Date Created",
        "Min Rem Shelf Life",
        "Total shelf life",
        "X-site artl status"
    ]

    # Ensure all required columns exist
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise KeyError(f"Missing required columns: {missing}")

    # Reorder columns
    df = df[required_cols]

    # Convert types to match Access
    df["Date Created"] = pd.to_datetime(df["Date Created"], errors="coerce")  # Date/path/to/file
    df["Date Extended"] = pd.to_datetime(df["Date Extended"], errors="coerce")  # Date/path/to/file
    df["Min Rem Shelf Life"] = pd.to_numeric(df["Min Rem Shelf Life"], errors="coerce")  # Number
    df["X-site artl status"] = df["X-site artl status"].astype(str)  # Short Text

    # Replace NaN/path/to/file with None for Access
    df = df.where(pd.notna(df), None)

    return df

def append_to_access_extended(report_path: Path, db_path: Path, table_name: str):
    if not report_path.exists():
        raise FileNotFoundError(f"Report not found: {report_path}")
    if not db_path.exists():
        raise FileNotFoundError(f"Access DB not found: {db_path}")

    print("Reading Excel report...")
    df = pd.read_excel(report_path, sheet_name=0)
    df = prepare_dataframe_for_access_extended(df)
    print(f"Columns ready for Access: {list(df.columns)}")
    print(f"Rows to append: {len(df)}")

    conn_str = r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=" + str(db_path)
    conn = pyodbc.connect(conn_str)
    cursor = conn.cursor()

    columns = df.columns.tolist()
    col_sql = ", ".join(f"[{c}]" for c in columns)
    placeholders = ", ".join("?" for _ in columns)
    insert_sql = f"INSERT INTO [{table_name}] ({col_sql}) VALUES ({placeholders})"

    data = [[None if pd.isna(x) else x for x in row] for row in df.values.tolist()]

    # Debug: print first row
    if data:
        print("Sample SQL insert:", insert_sql)
        print("Sample row data:", data[0])

    # Execute insert
    inserted_rows = 0
    for idx, row in enumerate(data, 1):
        try:
            cursor.execute(insert_sql, row)
            inserted_rows += 1
        except Exception as e:
            print(f"Row {idx} failed to insert:", e)

    conn.commit()
    print(f"Append completed. {inserted_rows}/path/to/file rows added.")

    cursor.close()
    conn.close()


# -------------------------- CHANGE LOG --------------------------
def _last_7_days_range():
    """Rolling last 7 days: start = now-7d (00:00), end = today (23:59:59.1)."""
    now = datetime.now()
    start = (now - timedelta(days=7)).replace(hour=0, minute=0, second=0, microsecond=0)
    end = now.replace(hour=23, minute=59, second=59, microsecond=1)
    return start, end

def _fetch_articles_for_range(db_path: Path, table_name: str, date_column: str, article_column: str = "Article") -> list[str]:
    """Query DISTINCT articles for the last-7-days window (parameterized) with cleaned, no-decimal article values."""
    start_dt, end_dt = _last_7_days_range()

    conn_str = r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=" + str(db_path)
    conn = pyodbc.connect(conn_str)
    cursor = conn.cursor()

    tbl = f"[{table_name}]"
    dtc = f"[{date_column}]"
    arc = f"[{article_column}]"

    query = f"""
        SELECT DISTINCT {arc}
        FROM {tbl}
        WHERE {dtc} BETWEEN ? AND ?
          AND {arc} IS NOT NULL
    """
    cursor.execute(query, (start_dt, end_dt))

    def clean_article(val):
        """Normalize to a clean article string:
           - Remove .0 / .000 decimals
           - Preserve leading zeros
           - Keep alphanumeric codes unchanged
        """
        if val is None:
            return None

        # Case 1: numeric types (float / int from Access)
        if isinstance(val, int):
            return str(val)

        if isinstance(val, float):
            if val.is_integer():
                return str(int(val))  # 1.0 → '1'
            # Rare case: true fractional values (shouldn't occur)
            return format(val, 'f').rstrip('0').rstrip('.')

        # Case 2: string types
        s = str(val).strip()
        if not s:
            return None

        # Remove trailing .0/path/to/file if all zeros
        if '.' in s:
            left, right = s.split('.', 1)
            if left.isdigit() and set(right) <= {'0'}:
                return left  # "1.0" → "1"

        # Pure digits (with or without leading zeros)
        return s  # keep leading zeros untouched

    # Build result list
    results = []
    for row in cursor.fetchall():
        cleaned = clean_article(row[0])
        if cleaned:
            results.append(cleaned)

    cursor.close()
    conn.close()
    return results

def copy_articles_from_previous_week(dedupe: bool = False):
    """
    Combine Articles from last 7 days across the two tables and copy to clipboard ONCE.
      - ACCESS_TABLES["created"]  → "Created On"
      - ACCESS_TABLES["extended"] → "Date Created"
    If dedupe=False (default): keep ALL rows (Created + Extended) -> totals add (e.g., 235 + 43 = 278).
    If dedupe=True: unique only.
    """
    created  = _fetch_articles_for_range(ACCESS_DB, ACCESS_TABLES["created"],  "Created On",  "Article")
    extended = _fetch_articles_for_range(ACCESS_DB, ACCESS_TABLES["extended"], "Date Created", "Article")

    # ---- CHANGE HERE: no de-duplication by default ----
    combined = created + extended

    if dedupe:
        seen = set()
        combined = [x for x in combined if not (x in seen or seen.add(x))]

    print("Starting to copy articles from the last 7 days (Created + Extended)...")
    print(f"• {ACCESS_TABLES['created']}:  {len(created)} article(s)")
    print(f"• {ACCESS_TABLES['extended']}: {len(extended)} article(s)")
    print(f"• Combined total (dedupe={'ON' if dedupe else 'OFF'}): {len(combined)}")

    if not combined:
        print("No Articles found in the last 7 days.")
        return combined

    # Single clipboard write (CRLF for SAP multiple selection)
    text = "\r\n".join(combined)
    if 'set_clipboard_text' in globals() and callable(globals()['set_clipboard_text']):
        set_clipboard_text(text)
        method = "set_clipboard_text"
    else:
        try:
            import pyperclip
            pyperclip.copy(text)
            method = "pyperclip"
        except Exception:
            import win32clipboard, win32con
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, text)
            win32clipboard.CloseClipboard()
            method = "win32clipboard"

    print(f"Copied {len(combined)} Articles to clipboard ({method}).")
    preview = ", ".join(combined[:10]) + (" ..." if len(combined) > 10 else "")
    print("Preview:", preview)
    return combined

def sap_export_change():
    try:
        session = get_sap_session()
    except Exception as e:
        print("ERROR: could not connect to SAP scripting:", e)
        traceback.print_exc()
        return
    
    print("Running SE16N -> MARA -> export ...")
    print("Running SE16N -> MARA -> export ...")
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/path/to/file").text = "/path/to/file"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/path/to/file").text = "MARA"
    session.findById("wnd[0]/path/to/file").setFocus()
    session.findById("wnd[0]/path/to/file").caretPosition = 4
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/path/to/file").text = "SHELF_LIFE"
    session.findById("wnd[0]/path/to/file").text = ""
    session.findById("wnd[0]/path/to/file").press()
    session.findById("wnd[0]/path/to/file").setFocus()
    session.findById("wnd[0]/path/to/file").press()
    session.findById("wnd[1]/path/to/file").press()
    session.findById("wnd[1]/path/to/file").press()
    session.findById("wnd[0]/path/to/file").text = "HAWA"
    session.findById("wnd[0]/path/to/file").setFocus()
    session.findById("wnd[0]/path/to/file").caretPosition = 10
    session.findById("wnd[0]/path/to/file").setFocus()
    session.findById("wnd[0]/path/to/file").press()
    session.findById("wnd[1]/path/to/file").text = "G*"
    # session.findById("wnd[1]/path/to/file").text = "F*"
    session.findById("wnd[1]/path/to/file").text = "P*"
    session.findById("wnd[1]/path/to/file").setFocus()
    session.findById("wnd[1]/path/to/file").caretPosition = 2
    session.findById("wnd[1]/path/to/file").press()
    # execute
    session.findById("wnd[0]/path/to/file").press()

    # select all rows, select MATNR column, context menu export (&XXL) as in your snippet
    time.sleep(LONG_SLEEP)
    session.findById("wnd[0]/path/to/file").currentCellRow = -1
    session.findById("wnd[0]/path/to/file").selectColumn("MATNR")
    session.findById("wnd[0]/path/to/file").contextMenu()
    session.findById("wnd[0]/path/to/file").selectContextMenuItem("&XXL")
    # export dialog - set path and filename then press Save (btn[11] per your snippet)
    time.sleep(SHORT_SLEEP)
    session.findById("wnd[1]/path/to/file").press()  # press Continue/path/to/file on the Excel export dialog - per your recorded script
    time.sleep(SHORT_SLEEP)
    session.findById("wnd[1]/path/to/file").text = str(CHANGES_DIR)
    session.findById("wnd[1]/path/to/file").text = CHANGE_LOG_FILENAME
    session.findById("wnd[1]/path/to/file").caretPosition = 23
    session.findById("wnd[1]/path/to/file").press()  # Save

    # Wait a bit for SAP to finish and for Excel to auto-open the exported file
    print("Waiting for Excel to open the exported file...")
    excel_app, wb = wait_for_excel_workbook(CHANGE_LOG_FILENAME, timeout=WAIT_FOR_EXCEL_OPEN_SEC)
    if excel_app is None or wb is None:
        print(f"WARNING: Excel file {CHANGE_LOG_FILENAME} didn't open within {WAIT_FOR_EXCEL_OPEN_SEC}s. Trying to locate file on disk.")
        # fallback: try opening file directly if path exists
        exported_path = os.path.join(CHANGES_DIR, CHANGE_LOG_FILENAME)
        if os.path.exists(exported_path):
            # Launch Excel and open it
            print("Opening workbook directly from disk:", exported_path)
            excel_app = Dispatch("Excel.Application")
            excel_app.Visible = True
            wb = excel_app.Workbooks.Open(exported_path)
        else:
            raise RuntimeError(f"Exported file not found at {exported_path}")

    # Copy column D to clipboard
    values = copy_column_from_workbook(wb, column_letter='D', start_row=2)
    set_clipboard_text("\r\n".join(values))
    print(f"Column D copied to clipboard from {CHANGE_LOG_FILENAME} ({len(values)} values)")

    # prepare clipboard text as newline-separated (kept for manual use if needed)
    clipboard_text = "\r\n".join(values)
    set_clipboard_text(clipboard_text)
    print("Values copied to clipboard.")

# -------------------------- Open hierarchy transaction, run and export (NO paste popup) --------------------------
    print("Starting hierarchy transaction steps...")
    session.findById("wnd[0]/path/to/file").text = "/path/to/file"
    session.findById("wnd[0]").sendVKey(0)

    time.sleep(1)
    # Run/path/to/file and drill-down per your script
    session.findById("wnd[0]/path/to/file").press()
    time.sleep(1)
    # The script you provided then selects a row and double-clicks to drill into details:
    session.findById("wnd[1]/path/to/file").currentCellRow = 6
    session.findById("wnd[1]/path/to/file").selectedRows = "6"
    session.findById("wnd[1]/path/to/file").doubleClickCurrentCell()
    time.sleep(0.5)
    # press the MATKL button (from your snippet)
    session.findById("wnd[0]/path/to/file").press()

    # Your original snippet then configured some selection values
    session.findById("wnd[1]/path/to/file").press()
    time.sleep(0.2)
    session.findById("wnd[1]/path/to/file").press()
    time.sleep(0.2)
    session.findById("wnd[1]/path/to/file"
                        "tblSAPLALDBSINGLE/path/to/file").text = "F08010201"
    session.findById("wnd[1]/path/to/file"
                        "tblSAPLALDBSINGLE/path/to/file").text = "F77070101"
    session.findById("wnd[1]/path/to/file"
                        "tblSAPLALDBSINGLE/path/to/file").text = "F04010101"
    session.findById("wnd[1]/path/to/file"
                        "tblSAPLALDBSINGLE/path/to/file").text = "F77070101"
    session.findById("wnd[1]/path/to/file"
                        "tblSAPLALDBSINGLE/path/to/file").text = "F04010101"
    session.findById("wnd[1]/path/to/file"
                        "tblSAPLALDBSINGLE/path/to/file").text = "C27640201"
    session.findById("wnd[1]/path/to/file"
                        "tblSAPLALDBSINGLE/path/to/file").text = "F77070101"
    session.findById("wnd[1]/path/to/file"
                        "tblSAPLALDBSINGLE/path/to/file").text = "F77080110"
    time.sleep(0.2)
    session.findById("wnd[1]/path/to/file").press()
    session.findById("wnd[1]/path/to/file").press()

    # Back to main screen, run the report
    time.sleep(0.5)
    session.findById("wnd[0]/path/to/file").press()  # run

    time.sleep(1)
    # Select the LVL1_DESCR column (per your snippet) and export directly, no paste popup
    grid = session.findById("wnd[0]/path/to/file")
    grid.currentCellRow = -1
    grid.selectColumn("LVL1_DESCR")
    grid.contextMenu()
    grid.selectContextMenuItem("&XXL")
    time.sleep(0.5)

    # Confirm export and set file path/path/to/file – this is now the ONLY dialog we touch
    session.findById("wnd[1]/path/to/file").press()
    session.findById("wnd[1]/path/to/file").text = str(CHANGES_DIR)
    session.findById("wnd[1]/path/to/file").text = HIERARCHY_FILENAME
    session.findById("wnd[1]/path/to/file").caretPosition = len(HIERARCHY_FILENAME)
    session.findById("wnd[1]/path/to/file").press()

    time.sleep(1)

    # Copy column A to clipboard
    values = copy_column_from_workbook(wb, column_letter='A', start_row=2)
    set_clipboard_text("\r\n".join(values))
    print(f"Column A copied to clipboard from {CHANGE_LOG_FILENAME} ({len(values)} values)")

    # prepare clipboard text as newline-separated (kept for manual use if needed)
    clipboard_text = "\r\n".join(values)
    set_clipboard_text(clipboard_text)
    print("Values copied to clipboard.")

    print("Starting Purch Org steps...")
    session.findById("wnd[0]/path/to/file").text = "/path/to/file"
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/path/to/file").text = "EINA"
    session.findById("wnd[0]/path/to/file").setFocus()
    session.findById("wnd[0]/path/to/file").caretPosition = 4
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/path/to/file").setFocus()
    session.findById("wnd[0]/path/to/file").press()
    session.findById("wnd[1]/path/to/file").press()
    session.findById("wnd[1]/path/to/file").press()
    session.findById("wnd[0]/path/to/file").text = ""
    session.findById("wnd[0]/path/to/file").setFocus()
    session.findById("wnd[0]/path/to/file").caretPosition = 0
    session.findById("wnd[0]/path/to/file").press()
    session.findById("wnd[0]/path/to/file").setCurrentCell(-1,"MATKL")
    session.findById("wnd[0]/path/to/file").selectColumn ("MATKL")
    session.findById("wnd[0]/path/to/file").contextMenu()
    session.findById("wnd[0]/path/to/file").selectContextMenuItem ("&XXL")
    session.findById("wnd[1]/path/to/file").press()
    session.findById("wnd[1]/path/to/file").text = str(CHANGES_DIR)
    session.findById("wnd[1]/path/to/file").text = INFORECORD_FILENAME_1
    session.findById("wnd[1]/path/to/file").press()

    # Wait for Excel
    print("Waiting for Excel to open...")
    excel_app2, wb2 = wait_for_excel_workbook(INFORECORD_FILENAME_1, timeout=WAIT_FOR_EXCEL_OPEN_SEC)
    if excel_app2 is None or wb2 is None:
        print(f"WARNING: Excel {INFORECORD_FILENAME_1} did not open automatically. Checking on disk.")
        exported_path2 = os.path.join(CHANGES_DIR, INFORECORD_FILENAME_1)
        if os.path.exists(exported_path2):
            excel_app2 = Dispatch("Excel.Application")
            try:
                excel_app2.Visible = True
            except AttributeError:
                print("Could not set Excel visibility, continuing anyway.")
            wb2 = excel_app2.Workbooks.Open(exported_path2)
        else:
            print("Could not find exported hierarchy file on disk. Continuing to final steps.")

    # Copy column D to clipboard
    values = copy_column_from_workbook(wb2, column_letter='A', start_row=2)
    set_clipboard_text("\r\n".join(values))
    print(f"Column A copied to clipboard from {INFORECORD_FILENAME_1} ({len(values)} values)")

    # prepare clipboard text as newline-separated (kept for manual use if needed)
    clipboard_text = "\r\n".join(values)
    set_clipboard_text(clipboard_text)
    print("Values copied to clipboard.")

    session.findById("wnd[0]/path/to/file").text = "/path/to/file"
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/path/to/file").text = "EINE"
    session.findById("wnd[0]/path/to/file").setFocus()
    session.findById("wnd[0]/path/to/file").caretPosition = 4
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/path/to/file").setFocus()
    session.findById("wnd[0]/path/to/file").press()
    session.findById("wnd[1]/path/to/file").press()
    session.findById("wnd[1]/path/to/file").press()
    session.findById("wnd[0]/path/to/file").text = ""
    session.findById("wnd[0]/path/to/file").setFocus()
    session.findById("wnd[0]/path/to/file").caretPosition = 0
    session.findById("wnd[0]/path/to/file").press()
    session.findById("wnd[0]/path/to/file").currentCellRow = -1
    session.findById("wnd[0]/path/to/file").selectColumn ("INFNR")
    session.findById("wnd[0]/path/to/file").contextMenu()
    session.findById("wnd[0]/path/to/file").selectContextMenuItem ("&XXL")
    session.findById("wnd[1]/path/to/file").press()
    session.findById("wnd[1]/path/to/file").text = str(CHANGES_DIR)
    session.findById("wnd[1]/path/to/file").text = PURCH_ORG_FILENAME_1
    session.findById("wnd[1]/path/to/file").press()

    # Wait for Excel to auto-open the exported hierarchy workbook
    print("Waiting for Excel to open...")
    excel_app2, wb2 = wait_for_excel_workbook(PURCH_ORG_FILENAME_1, timeout=WAIT_FOR_EXCEL_OPEN_SEC)
    if excel_app2 is None or wb2 is None:
        print(f"WARNING: Excel {PURCH_ORG_FILENAME_1} did not open automatically. Checking on disk.")
        exported_path2 = os.path.join(CHANGES_DIR, PURCH_ORG_FILENAME_1)
        if os.path.exists(exported_path2):
            excel_app2 = Dispatch("Excel.Application")
            try:
                excel_app2.Visible = True
            except AttributeError:
                print("Could not set Excel visibility, continuing anyway.")
            wb2 = excel_app2.Workbooks.Open(exported_path2)
        else:
            print("Could not find exported hierarchy file on disk. Continuing to final steps.")

    # Copy column D to clipboard
    values = copy_column_from_workbook(wb2, column_letter='B', start_row=2)
    set_clipboard_text("\r\n".join(values))
    print(f"Column B copied to clipboard from {PURCH_ORG_FILENAME_1} ({len(values)} values)")

    # prepare clipboard text as newline-separated (kept for manual use if needed)
    clipboard_text = "\r\n".join(values)
    set_clipboard_text(clipboard_text)
    print("Values copied to clipboard.")

    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/path/to/file").text = "/path/to/file"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/path/to/file").text = "T024E"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/path/to/file").text = "PURCH_ORG_DE"
    session.findById("wnd[0]/path/to/file").setFocus()
    session.findById("wnd[0]/path/to/file").press()
    time.sleep(0.2)
    session.findById("wnd[1]/path/to/file").press()
    time.sleep(0.2)
    session.findById("wnd[1]/path/to/file").press()
    session.findById("wnd[0]/path/to/file").text = ""
    session.findById("wnd[0]/path/to/file").setFocus()
    session.findById("wnd[0]/path/to/file").caretPosition = 0
    session.findById("wnd[0]/path/to/file").press()
    session.findById("wnd[0]/path/to/file").setCurrentCell(-1, "EKOTX")
    session.findById("wnd[0]/path/to/file").selectColumn("EKOTX")
    session.findById("wnd[0]/path/to/file").contextMenu()
    session.findById("wnd[0]/path/to/file").selectContextMenuItem("&XXL")
    session.findById("wnd[1]/path/to/file").press()
    session.findById("wnd[1]/path/to/file").text = str(CHANGES_DIR)
    session.findById("wnd[1]/path/to/file").text = PURCH_ORG_DESC_FILENAME_1
    session.findById("wnd[1]/path/to/file").caretPosition = 11
    session.findById("wnd[1]/path/to/file").press()

def create_reference_file_log():
    wait_for_file_stable(INFORECORD_1, stable_seconds=3)
    wait_for_file_stable(PURCH_ORG_PATH_1, stable_seconds=3)
    wait_for_file_stable(PURCH_ORG_PATH_2, stable_seconds=3)

    output_path = CHANGES_DIR / "Reference.xlsx"
    if output_path.exists():
        try:
            output_path.unlink()
        except Exception:
            pass

    # --- Read source files ---
    info_df = pd.read_excel(INFORECORD_1, sheet_name=0)
    po_df = pd.read_excel(PURCH_ORG_PATH_1, sheet_name=0)
    po_desc_df = pd.read_excel(PURCH_ORG_PATH_2, sheet_name=0)

    # --- Validate required columns ---
    if "Purchasing Info Rec." not in info_df.columns or "Article" not in info_df.columns:
        raise KeyError(
            "Inforecord file must contain 'Purchasing Info Rec.' and 'Article' columns"
        )

    # --- Prepare Purch Org lookup ---
    po_key_col = po_df.columns[0]
    po_val_col = po_df.columns[1]
    po_df[po_key_col] = pd.to_numeric(po_df[po_key_col], errors="coerce")

    po_map = dict(
        zip(
            po_df[po_key_col].dropna().astype(int),
            po_df[po_val_col]
        )
    )

    # --- Prepare Purch Org Description lookup ---
    desc_key_col = po_desc_df.columns[0]
    desc_val_col = po_desc_df.columns[1]

    desc_map = dict(
        zip(
            po_desc_df[desc_key_col].astype(str).str.strip(),
            po_desc_df[desc_val_col]
        )
    )

    # --- Build Reference dataframe ---
    ref = pd.DataFrame({
        "Purchasing Info Rec.": info_df["Purchasing Info Rec."],
        "Article": info_df["Article"],
    })

    ref_key = pd.to_numeric(ref["Purchasing Info Rec."], errors="coerce")
    ref["Purch Org"] = ref_key.map(
        lambda x: po_map.get(int(x)) if pd.notna(x) else None
    )

    ref["Purch Org"] = ref["Purch Org"].fillna("0000")
    ref["Purch Org Description"] = ref["Purch Org"].astype(str).str.strip().map(desc_map).fillna(NOT_FOUND_TEXT)

    # --- Write Reference.xlsx ---
    ref.to_excel(output_path, index=False)

    # --- Delete source files ONLY after success ---
    for f in [INFORECORD_1, PURCH_ORG_PATH_1, PURCH_ORG_PATH_2]:
        try:
            if f.exists():
                f.unlink()
                print(f"Deleted source file: {f.name}")
        except Exception as e:
            print(f"Could not delete {f.name}: {e}")

    print("Reference.xlsx created and source files deleted.")
    return output_path

def build_hierarchy_lookup_openpyxl(ws_h, map_return_cols):
    """
    Openpyxl-only hierarchy lookup for Change Log.
    Key column: H
    """
    key_col = col_letter_to_index("H")
    ret_cols = {
        header: col_letter_to_index(col)
        for header, col in map_return_cols.items()
    }

    lookup = {}

    for r in range(2, ws_h.max_row + 1):
        k = ws_h.cell(r, key_col).value
        if k is None:
            continue

        ks = str(k).strip()
        if ks.endswith(".0"):
            ks = ks[:-2]
        if not ks:
            continue

        row = {}
        for header, cidx in ret_cols.items():
            row[header] = ws_h.cell(r, cidx).value

        lookup[ks] = row

    return lookup

def apply_basic_formatting_openpyxl(ws):
    """
    Openpyxl-only formatting:
    - Header formatting
    - Autofilter
    - Freeze header
    - Column auto-width
    """

    header_fill = PatternFill("solid", fgColor="BFBFBF")
    header_font = Font(bold=True)
    header_align = Alignment(horizontal="center", vertical="center")
    thin = Side(border_style="thin", color="1")
    header_border = Border(top=thin, left=thin, right=thin, bottom=thin)

    max_row = ws.max_row
    max_col = ws.max_column

    # ---- Header row ----
    for c in range(1, max_col + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_align
        cell.border = header_border

    # ---- Freeze + filter ----
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(max_col)}{max_row}"

    # ---- Autofit-ish column width ----
    for c in range(1, max_col + 1):
        max_len = 0
        for r in range(1, min(max_row, 250) + 1):
            v = ws.cell(row=r, column=c).value
            if v:
                max_len = max(max_len, len(str(v)))
        ws.column_dimensions[get_column_letter(c)].width = min(max(10, max_len + 2), 45)

def excel_formatting_changes():
    """
    Change Log formatting using:
    - Reference.xlsx (Purch Org + Description)
    - Hierarchy.xlsx (MAP_RETURN_COLS)
    """

    kill_excel()
    time.sleep(2)

    # ---------------- Paths ----------------
    change_path   = CHANGE_LOG_PATH
    hierarchy_path = CHANGES_DIR / HIERARCHY_FILENAME
    reference_path = CHANGES_DIR / "Reference.xlsx"

    required_files = [change_path, hierarchy_path, reference_path]
    for p in required_files:
        if not p.exists():
            raise FileNotFoundError(f"Missing file: {p}")
        wait_for_file_stable(p, timeout=180, stable_seconds=3)

    print("All source files stable")

    # ---------------- Load workbooks ----------------
    wb_log = load_workbook(change_path)
    ws_log = wb_log.worksheets[0]

    wb_ref = load_workbook(reference_path, data_only=True)
    ws_ref = wb_ref.worksheets[0]

    wb_h = load_workbook(hierarchy_path, data_only=True)
    ws_h = wb_h.worksheets[0]

    # ---------------- Helpers ----------------
    def norm(v):
        if v is None:
            return ""
        s = str(v).strip()
        return s[:-2] if s.endswith(".0") else s

    def find_col(ws, name):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(1, c).value
            if v and str(v).strip().lower() == name.lower():
                return c
        return None

   
    # Build Reference lookup (Article → Purch Org / Desc)
   
    print("Building Reference lookup")

    ref_article_col = find_col(ws_ref, "Article")
    ref_org_col     = find_col(ws_ref, "Purch Org")
    ref_desc_col    = find_col(ws_ref, "Purch Org Description")

    if not all([ref_article_col, ref_org_col, ref_desc_col]):
        raise RuntimeError("Reference.xlsx missing required columns")

    ref_lookup = {}
    for r in range(2, ws_ref.max_row + 1):
        key = norm(ws_ref.cell(r, ref_article_col).value)
        if not key:
            continue

        org  = ws_ref.cell(r, ref_org_col).value or "0000"
        desc = ws_ref.cell(r, ref_desc_col).value or NOT_FOUND_TEXT
        ref_lookup[key] = (org, desc)

   
    #Insert Purch Org columns (C:D)
   
    print("Inserting Purch Org columns")

    ws_log.insert_cols(3, amount=2)
    ws_log.cell(1, 3).value = "Purch Org"
    ws_log.cell(1, 4).value = "Purch Org Description"

    article_col = find_col(ws_log, "Article")
    if not article_col:
        raise RuntimeError("Change Log missing Article column")

    for r in range(2, ws_log.max_row + 1):
        key = norm(ws_log.cell(r, article_col).value)
        org, desc = ref_lookup.get(key, ("0000", NOT_FOUND_TEXT))
        ws_log.cell(r, 3).value = org
        ws_log.cell(r, 4).value = desc

   
    #Hierarchy lookup (reuse fast logic)
   
    print("Applying hierarchy mapping")

    hier_lookup = build_hierarchy_lookup_openpyxl(ws_h, MAP_RETURN_COLS)

    # Insert hierarchy columns after D
    insert_at = col_letter_to_index("E")
    headers = list(MAP_RETURN_COLS.keys())[:2]

    ws_log.insert_cols(insert_at, amount=len(headers))
    for i, h in enumerate(headers):
        ws_log.cell(1, insert_at + i).value = h

    key_col = col_letter_to_index("H")
    for r in range(2, ws_log.max_row + 1):
        key = norm(ws_log.cell(r, key_col).value)
        row = hier_lookup.get(key, {})

        for i, h in enumerate(headers):
            ws_log.cell(
                r,
                insert_at + i
            ).value = row.get(h, NOT_FOUND_TEXT) or NOT_FOUND_TEXT

   
    #Formatting
   
    print("Applying formatting")
    apply_basic_formatting_openpyxl(ws_log)

    # ---------------- Save ----------------
    wb_log.save(change_path)
    wb_log.close()
    wb_ref.close()
    wb_h.close()

    # ---------------- Move files ----------------
    date_str = extract_date_from_filename(change_path)
    dest = CHANGES_DIR / date_str
    dest.mkdir(parents=True, exist_ok=True)

    move_files_to_folder(
        [change_path, hierarchy_path, reference_path],
        dest
    )

    final_path = dest / change_path.name
    wait_for_file(final_path, timeout=60)

    print("Change Log formatting complete:", final_path)
    return final_path

def prepare_dataframe_for_access_change_log(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = df.columns.str.strip()

    # Normalize column names you expect
    rename_map = {
        "Min. Rem. Shelf Life": "Min Rem Shelf Life",
    }
    df = df.rename(columns=rename_map)

    required_cols = [
        "Article",
        "Article description",
        "Material Type",
        "Merchandise Category",
        "BMC Description",
        "Divisional Buyer",
        "Created On",
        "Min Rem Shelf Life",
        "Total shelf life",
        "X-site artl status",
    ]

    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise KeyError(f"Missing required columns for Access (change log): {missing}")

    df = df[required_cols].copy()

    # ---- Key fields ----
    # Treat Article as TEXT and trim
    df["Article"] = df["Article"].astype(str).str.strip()

    # Parse Created On; keep full precision, but also compute second-level for Access granularity
    df["Created On"] = pd.to_datetime(df["Created On"], errors="coerce")
    df["__CreatedOn_sec__"] = df["Created On"].dt.floor("s")

    # ---- Types ----
    df["Min Rem Shelf Life"] = pd.to_numeric(df["Min Rem Shelf Life"], errors="coerce")
    df["Total shelf life"]   = pd.to_numeric(df["Total shelf life"], errors="coerce")
    df["X-site artl status"] = df["X-site artl status"].astype("string")

    # ---- Drop rows missing key parts ----
    df = df.dropna(subset=["Article", "Created On", "__CreatedOn_sec__"])

    # ---- In-file de-duplication at Access's time grain ----
    # This prevents multiple rows that would collide in Access (same Article, same second)
    df = df.sort_values(["Article", "Created On"]).drop_duplicates(
        subset=["Article", "__CreatedOn_sec__"], keep="first"
    )

    # Use the second-level timestamp for the value we send to Access
    df["Created On"] = df["__CreatedOn_sec__"]
    df = df.drop(columns=["__CreatedOn_sec__"])

    # ---- Convert NaN/path/to/file -> None for Access NULLs ----
    df = df.where(pd.notna(df), None)

    return df

def append_to_access_change_log(report_path: Path, db_path: Path, table_name: str):
    report_path = Path(report_path)

    if not report_path.exists():
        raise FileNotFoundError(f"Report not found: {report_path}")
    if not db_path.exists():
        raise FileNotFoundError(f"Access DB not found: {db_path}")

    print("Reading Excel report...")
    df = pd.read_excel(report_path, sheet_name=0, engine="openpyxl")
    df = prepare_dataframe_for_access_change_log(df)

    print(f"Columns ready for Access: {list(df.columns)}")
    print(f"Rows to append (after in-file de-dupe): {len(df)}")

    conn_str = r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=" + str(db_path)
    conn = pyodbc.connect(conn_str)
    cursor = conn.cursor()

    columns = df.columns.tolist()
    col_sql = ", ".join(f"[{c}]" for c in columns)
    placeholders = ", ".join("?" for _ in columns)
    insert_sql = f"INSERT INTO [{table_name}] ({col_sql}) VALUES ({placeholders})"
    select_sql = f"SELECT 1 FROM [{table_name}] WHERE [Article] = ? AND [Created On] = ?"

    data_rows = df.values.tolist()

    # Precompute column indices for readability
    i_article = columns.index("Article")
    i_created = columns.index("Created On")

    inserted_rows = 0
    skipped_existing = 0
    errors = 0

    for idx, row in enumerate(data_rows, 1):
        try:
            # Normalize key fields for the lookup
            art = row[i_article]
            crd = row[i_created]

            if art is not None:
                art = str(art).strip()

            # 1) Idempotency check: skip if exists
            cursor.execute(select_sql, (art, crd))
            if cursor.fetchone():
                skipped_existing += 1
                continue

            # 2) Insert row
            cursor.execute(insert_sql, row)
            inserted_rows += 1

        except Exception as e:
            errors += 1
            print(f"Row {idx} failed: {row}")
            print("   Error:", e)

    conn.commit()
    cursor.close()
    conn.close()

    print(f"Append completed. Inserted: {inserted_rows}, Skipped existing: {skipped_existing}, Errors: {errors}")

# -------------------------- MAIN SCRIPT --------------------------
if __name__ == "__main__":
    # -------------------------- SAP exports ----------------------------
    sap_export_created()
    kill_excel()
    sap_export_extended()

    # -------------------------- Reference file creation ----------------------------
    kill_excel()
    create_reference_file()
    print("Reference file created.")
    time.sleep(3)

    # -------------------------- Excel Formatting & Access Import ----------------------------
    final_moved_report_path = excel_formatting_extended()
    
    if final_moved_report_path and Path(final_moved_report_path).exists():
        print("Excel Formatting completed.")

        append_to_access_extended(
            report_path=final_moved_report_path,
            db_path=ACCESS_DB,
            table_name=ACCESS_TABLES["extended"]
        )
        print("Data appended to Access database.")
    else:
        print(f"Report not found, skipping Access append: {final_moved_report_path}")

    final_report_path = excel_formatting_created()

    if final_report_path and Path(final_report_path).exists():
        print("Excel Formatting completed.")

        append_to_access_created(
            report_path=final_report_path,
            db_path=ACCESS_DB,
            table_name=ACCESS_TABLES["created"]
        )
        print("Data appended to Access database.")
    else:
        print(f"Report not found, skipping Access append: {final_report_path}")

    copy_yesterday_file()
    
    # -------------------------- Change Log exports ----------------------------
    copy_articles_from_previous_week()
    sap_export_change()
    print("SAP Export completed.")

    kill_excel()
    create_reference_file_log()
    print("Reference file created.")
    time.sleep(3)

    final_change_log_path = excel_formatting_changes()
    print("Excel Formatting completed.")

    if final_change_log_path and Path(final_change_log_path).exists():
        append_to_access_change_log(
            report_path=final_change_log_path,
            db_path=ACCESS_DB,
            table_name=ACCESS_TABLES["change_log"]
        )
        print("Change log appended to Access database.")
    else:
        print(f"Report not found, skipping Access append: {final_change_log_path}")
