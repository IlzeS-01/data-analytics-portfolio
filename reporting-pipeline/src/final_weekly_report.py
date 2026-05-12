# -*- coding: utf-8 -*-
"""
Weekly Report V5 - A Python script to automate the extraction, formatting, and reporting of SAP data for Demo Company's Finance and Administration team.

v1.4.2
------
- Fix: exceptions excel formatting has been corrected
- Includes: automatic installation of required packages, multi-threading for simultaneous SAP data extraction, and comprehensive data formatting.
- Added: added print statements that indicates if packages are installed or not
"""
# AUTO-INSTALL REQUIRED PACKAGES
import sys
import subprocess
import importlib
from typing import List

REQUIRED = [
    "pandas",
    "numpy",
    "openpyxl",
    "pyodbc",
    "pywin32",  # for win32com.client
]

def ensure(pkgs: List[str]):
    py = sys.executable
    for p in pkgs:
        try:
            # special handling for pywin32
            if p == "pywin32":
                importlib.import_module("win32com.client")
            else:
                importlib.import_module(p)
            print(f"[INSTALLED] {p}")
        except Exception:
            print(f"[INSTALLING] {p} ...")
            subprocess.check_call([py, "-m", "pip", "install", p])

ensure(REQUIRED)

# IMPORTS
import win32com.client
import time
import threading
import pythoncom
from datetime import datetime, timedelta
import os
import pandas as pd
import queue
import win32com.client as win32
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook
import shutil
import numpy as np
import pyodbc
import warnings
import re

warnings.simplefilter(action='ignore', category=FutureWarning)

# MAIN SETUP
# ================= PATHS =================
base_path = os.path.expanduser(r"~\Demo Company\Finance and Administration Team Site - Documents\Datos\VM & EDI Report")

# Exports
exceptions_export_path = os.path.join(base_path, "Exceptions Archive", "Version4 (SAP)")
greenlight_export_path = os.path.join(base_path, "Greenlight Archive", "Week Breakdown")
log_file_location = os.path.join(base_path, "Log File")
os.makedirs(exceptions_export_path, exist_ok=True)
os.makedirs(greenlight_export_path, exist_ok=True)

# Specific files
exceptions_lookup_file = os.path.join(base_path, "Exceptions Archive", "Exceptions.xlsx")
country_code_file = "Country Code.xlsx"
country_code_file_path = os.path.join(base_path, "Greenlight Archive", country_code_file)
exceptions_path = os.path.join(base_path, "Exceptions Archive")
access_db_path = os.path.join(base_path,"Reporting", "Report.accdb")

# ================= DATE SETUP =================
today = datetime.today()
this_week_monday = today - timedelta(days=today.weekday())
# how many weeks back do you want?

weeks_back = 1   # 1 = last week, 2 = week before that, etc.

report_monday = this_week_monday - timedelta(weeks=weeks_back)
report_sunday = report_monday + timedelta(days=6)


# Full week range (for convenience)
days = [(report_monday + timedelta(days=i)).strftime("%d.%m.%Y") for i in range(7)]
monday_str = report_monday.strftime("%d.%m.%Y")
sunday_str = report_sunday.strftime("%d.%m.%Y")
print(f"Report Week: {monday_str} to {sunday_str}")

# ================= FINANCIAL YEAR & WEEKS =================
financial_year = 2027
start_fw1 = datetime.strptime("02.03.2026", "%d.%m.%Y")

financial_weeks = {
    f"W{i+1:02}": (
        start_fw1 + timedelta(weeks=i),
        start_fw1 + timedelta(weeks=i, days=6),
    )
    for i in range(53)
}

def get_financial_week(date_to_check):
    if pd.isna(date_to_check):
        return ""
    if not isinstance(date_to_check, datetime):
        try:
            date_to_check = pd.to_datetime(date_to_check)
        except Exception:
            return ""
    for fw, (start, end) in financial_weeks.items():
        if start <= date_to_check <= end:
            return fw
    return "Unknown"

# Week label for reporting
week_label = get_financial_week(report_monday)

# SCRIPT 1: SAP DATA EXTRACT
# ======================= HELPERS =======================
def wait_for_id(session, sap_id, timeout=20, interval=0.5):
    t0 = time.time()
    while True:
        try:
            session.findById(sap_id)
            return True
        except Exception:
            if time.time() - t0 > timeout:
                return False
            time.sleep(interval)

def get_session(session_index, retries=20, delay=2):
    sapguiauto = win32com.client.GetObject("SAPGUI")
    application = sapguiauto.GetScriptingEngine
    connection = application.Children(0)
    for _ in range(retries):
        if connection.Children.Count > session_index:
            return connection.Children(session_index)
        time.sleep(delay)
    raise RuntimeError(f"SAP session {session_index} not found after waiting")

# ======================= THREAD WORKERS =======================
log_records = []
log_lock = threading.Lock()

def run_sqvi():
    pythoncom.CoInitialize()
    try:
        session = get_session(0)
        start_time = datetime.now()

        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/path/to/file").text = "/path/to/file"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/path/to/file").text = "EXCEPTIONS_RPT"
        session.findById("wnd[0]").sendVKey(8)

        # Open query selection
        if not wait_for_id(session, "wnd[0]/path/to/file", timeout=10):
            raise RuntimeError("SQVI: Query list button not available")
        session.findById("wnd[0]/path/to/file").press()

        if not wait_for_id(session, "wnd[1]/path/to/file", timeout=15):
            raise RuntimeError("SQVI: Query chooser grid not available")
        shell = session.findById("wnd[1]/path/to/file")
        shell.selectedRows = "0"
        shell.doubleClickCurrentCell()

        # Set date range
        if not wait_for_id(session, "wnd[0]/path/to/file", timeout=10):
            raise RuntimeError("SQVI: Date field LOW not available")
        session.findById("wnd[0]/path/to/file").text = monday_str
        session.findById("wnd[0]/path/to/file").text = sunday_str
        session.findById("wnd[0]").sendVKey(8)  # Execute

        # Export to Excel
        alv = session.findById("wnd[0]/path/to/file")
        alv.pressToolbarContextButton("&MB_EXPORT")
        alv.selectContextMenuItem("&XXL")
        session.findById("wnd[1]/path/to/file").press()
        time.sleep(5)  # Wait for dialog to open
        session.findById("wnd[1]/path/to/file").text = exceptions_export_path
        session.findById("wnd[1]/path/to/file").text = f"exceptions_{week_label}.xlsx"
        session.findById("wnd[1]/path/to/file").press()  # Save
        print(f"SQVI export completed")

        status = "Success"
    except Exception as e:
        print(f"SQVI error: {e!r}")
        status = f"Error: {e}"
    finally:
        end_time = datetime.now()
        with log_lock:
            log_records.append({
                "Session": 1, "Report": "SQVI", "Start": start_time, "End": end_time,
                "Duration_sec": (end_time - start_time).total_seconds(), "Status": status
            })
        pythoncom.CoUninitialize()

def run_greenlight():
    pythoncom.CoInitialize()
    try:
        session = get_session(1)
        start_time = datetime.now()

        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/path/to/file").text = "/path/to/file"
        session.findById("wnd[0]").sendVKey(0)

        # Open variant list
        if not wait_for_id(session, "wnd[0]/path/to/file", timeout=10):
            raise RuntimeError("Greenlight: Variant button not available")
        session.findById("wnd[0]/path/to/file").press()

        if not wait_for_id(session, "wnd[1]/path/to/file", timeout=15):
            raise RuntimeError("Greenlight: Variant chooser grid not available")
        shell = session.findById("wnd[1]/path/to/file")
        shell.currentCellRow = 1
        shell.selectedRows = "1"
        shell.doubleClickCurrentCell()

        # Dates + execute
        if not wait_for_id(session, "wnd[0]/path/to/file", timeout=10):
            raise RuntimeError("Greenlight: LOW date field not available")
        session.findById("wnd[0]/path/to/file").text = monday_str
        session.findById("wnd[0]/path/to/file").text = sunday_str
        session.findById("wnd[0]/path/to/file").press()

        # --- Wait for ALV grid to be ready ---
        def wait_for_grid(session, timeout=15):
            start = time.time()
            while time.time() - start < timeout:
                try:
                    grid = session.findById("wnd[0]/path/to/file")
                    if grid.RowCount > 0:
                        return grid
                except:
                    pass
                time.sleep(0.5)
            raise RuntimeError("Greenlight: ALV grid not available")

        shell_grid = wait_for_grid(session)

        # --- Export to Excel ---
        shell_grid.setCurrentCell(-1, "IVTYP")
        shell_grid.selectColumn("IVTYP")
        shell_grid.selectedRows = ""
        shell_grid.contextMenu()
        time.sleep(0.5)
        shell_grid.selectContextMenuItem("&XXL")
        session.findById("wnd[1]/path/to/file").press()
        time.sleep(2)
        session.findById("wnd[1]/path/to/file").text = greenlight_export_path
        session.findById("wnd[1]/path/to/file").text = f"{week_label} Greenlight.xlsx"
        session.findById("wnd[1]").sendVKey(11)
        print(f"Greenlight export completed")

        status = "Success"
    except Exception as e:
        print(f"Greenlight error: {e!r}")
        status = f"Error: {e}"
    finally:
        end_time = datetime.now()
        with log_lock:
            log_records.append({
                "Session": 2, "Report": "Greenlight", "Start": start_time, "End": end_time,
                "Duration_sec": (end_time - start_time).total_seconds(), "Status": status
            })
        pythoncom.CoUninitialize()

def run_threads():
    t1 = threading.Thread(target=run_sqvi, name="SQVI-Thread")
    t2 = threading.Thread(target=run_greenlight, name="Greenlight-Thread")
    t1.start()
    t2.start()
    t1.join()
    t2.join()

# ======================= CHECK AND CREATE SQVI QUERY =======================
query_name = "EXCEPTIONS_RPT"
try:
    session = get_session(0)
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/path/to/file").text = "SQVI"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/path/to/file").text = query_name
    session.findById("wnd[0]/path/to/file").press()

    time.sleep(2)
    try:
        session.findById("wnd[1]")  # New window → query does NOT exist
        query_exists = False
    except:
        query_exists = True  # No new window → query exists

except Exception as e:
    print(f"Error checking query: {e!r}")
    query_exists = False

if query_exists:
    print(f"SQVI query '{query_name}' exists. Running threads...")
    time.sleep(2)
    # ======================= PREP SESSIONS (main thread only) =======================
    sapguiauto_main = win32com.client.GetObject("SAPGUI")
    application_main = sapguiauto_main.GetScriptingEngine
    connection_main = application_main.Children(0)

    # Ensure 2 sessions exist
    while connection_main.Children.Count < 2:
        connection_main.Children(0).createSession()
        time.sleep(3)
    run_threads()

else:
    print(f"SQVI query '{query_name}' does not exist. Creating it now...")
    time.sleep(2)  # Wait for the session to stabilize

    def sqvi_setup():
        pythoncom.CoInitialize()
        try:
            # ======================= SQVI SETUP =======================
            session.findById("wnd[0]").maximize()
            session.findById("wnd[0]/path/to/file").text = "/path/to/file"
            session.findById("wnd[0]").sendVKey(0)

            # Open your query
            session.findById("wnd[0]/path/to/file").text = "EXCEPTIONS_RPT"
            session.findById("wnd[0]/path/to/file").press()
            session.findById("wnd[1]/path/to/file").text = "EXCEPTIONS_RPT"

            # Table join selection
            session.findById("wnd[1]/path/to/file").setFocus()
            session.findById("wnd[1]/path/to/file").key = "Table join"
            session.findById("wnd[1]/path/to/file").press()

            # Add tables
            tables = [
                "/path/to/file",
                "/path/to/file",
                "/path/to/file",
                "/path/to/file",
                "/path/to/file",
                "LFA1"
            ]

            for table in tables:
                session.findById("wnd[0]/path/to/file").press()
                session.findById("wnd[1]/path/to/file").text = table
                session.findById("wnd[1]/path/to/file").caretPosition = len(table)
                session.findById("wnd[1]").sendVKey(0)

            # Join tables
            session.findById("wnd[0]/path/to/file").press()
            session.findById("wnd[1]/path/to/file").key = "/path/to/file"
            session.findById("wnd[1]/path/to/file").key = "LFA1"
            session.findById("wnd[1]/path/to/file").setFocus()
            session.findById("wnd[1]/path/to/file").press()
            session.findById("wnd[0]/path/to/file").press()

            # Maximize and prepare ALV grid
            session.findById("wnd[0]").maximize()
            session.findById("wnd[0]").resizeWorkingPane(233, 39, False)
            session.findById("wnd[0]/path/to/file").dockerPixelSize = 898

            # ======================= COLUMN SELECTION =======================
            grid_id = "wnd[0]/path/to/file"

            # Wait for the grid to load
            start_time = time.time()
            timeout = 3
            grid = None
            while time.time() - start_time < timeout:
                try:
                    grid = session.findById(grid_id)
                    if hasattr(grid, "RowCount") and grid.RowCount > 0:
                        break
                except:
                    pass
                time.sleep(0.5)

            if grid is None:
                raise Exception(f"Grid {grid_id} not ready after {timeout}s")

            session.findById("wnd[0]").maximize()

            # VBA-like Python-safe sequence
            session.findById(grid_id).expandNode("          1")
            session.findById(grid_id).selectItem("          2", "COL1")
            session.findById(grid_id).ensureVisibleHorizontalItem("          2", "COL1")
            session.findById(grid_id).topNode = "          0"
            session.findById(grid_id).changeCheckbox("          2", "COL1", True)
            session.findById(grid_id).selectItem("          2", "COL2")
            session.findById(grid_id).ensureVisibleHorizontalItem("          2", "COL2")
            session.findById(grid_id).changeCheckbox("          2", "COL2", True)

            session.findById(grid_id).selectItem("          5", "COL1")
            session.findById(grid_id).ensureVisibleHorizontalItem("          5", "COL1")
            session.findById(grid_id).changeCheckbox("          5", "COL1", True)
            session.findById(grid_id).selectItem("          5", "COL2")
            session.findById(grid_id).ensureVisibleHorizontalItem("          5", "COL2")
            session.findById(grid_id).changeCheckbox("          5", "COL2", True)

            session.findById(grid_id).selectItem("          6", "COL1")
            session.findById(grid_id).ensureVisibleHorizontalItem("          6", "COL1")
            session.findById(grid_id).changeCheckbox("          6", "COL1", True)
            session.findById(grid_id).selectItem("          6", "COL2")
            session.findById(grid_id).ensureVisibleHorizontalItem("          6", "COL2")
            session.findById(grid_id).changeCheckbox("          6", "COL2", True)

            session.findById(grid_id).selectItem("         10", "COL1")
            session.findById(grid_id).ensureVisibleHorizontalItem("         10", "COL1")
            session.findById(grid_id).changeCheckbox("         10", "COL1", True)
            session.findById(grid_id).selectItem("         10", "COL2")
            session.findById(grid_id).ensureVisibleHorizontalItem("         10", "COL2")
            session.findById(grid_id).changeCheckbox("         10", "COL2", True)

            session.findById(grid_id).collapseNode("          1")
            session.findById(grid_id).expandNode("         17")
            session.findById(grid_id).selectItem("         21", "COL1")
            session.findById(grid_id).ensureVisibleHorizontalItem("         21", "COL1")
            session.findById(grid_id).topNode = "         17"
            session.findById(grid_id).changeCheckbox("         21", "COL1", True)
            session.findById(grid_id).selectItem("         21", "COL2")
            session.findById(grid_id).ensureVisibleHorizontalItem("         21", "COL2")
            session.findById(grid_id).changeCheckbox("         21", "COL2", True)

            session.findById(grid_id).selectItem("         22", "COL1")
            session.findById(grid_id).ensureVisibleHorizontalItem("         22", "COL1")
            session.findById(grid_id).changeCheckbox("         22", "COL1", True)
            session.findById(grid_id).selectItem("         22", "COL2")
            session.findById(grid_id).ensureVisibleHorizontalItem("         22", "COL2")
            session.findById(grid_id).changeCheckbox("         22", "COL2", True)

            session.findById(grid_id).selectItem("         30", "COL1")
            session.findById(grid_id).ensureVisibleHorizontalItem("         30", "COL1")
            session.findById(grid_id).changeCheckbox("         30", "COL1", True)
            session.findById(grid_id).selectItem("         30", "COL2")
            session.findById(grid_id).ensureVisibleHorizontalItem("         30", "COL2")
            session.findById(grid_id).changeCheckbox("         30", "COL2", True)

            session.findById(grid_id).selectItem("         33", "COL1")
            session.findById(grid_id).ensureVisibleHorizontalItem("         33", "COL1")
            session.findById(grid_id).changeCheckbox("         33", "COL1", True)
            session.findById(grid_id).selectItem("         33", "COL2")
            session.findById(grid_id).ensureVisibleHorizontalItem("         33", "COL2")
            session.findById(grid_id).changeCheckbox("         33", "COL2", True)

            session.findById(grid_id).collapseNode("         17")
            session.findById(grid_id).expandNode("         46")
            session.findById(grid_id).selectItem("         55", "COL1")
            session.findById(grid_id).ensureVisibleHorizontalItem("         55", "COL1")
            session.findById(grid_id).topNode = "          0"
            session.findById(grid_id).changeCheckbox("         55", "COL1", True)
            session.findById(grid_id).selectItem("         55", "COL2")
            session.findById(grid_id).ensureVisibleHorizontalItem("         55", "COL2")
            session.findById(grid_id).changeCheckbox("         55", "COL2", True)

            session.findById(grid_id).collapseNode("         46")
            session.findById(grid_id).expandNode("         65")
            session.findById(grid_id).selectItem("         68", "COL1")
            session.findById(grid_id).ensureVisibleHorizontalItem("         68", "COL1")
            session.findById(grid_id).topNode = "          0"
            session.findById(grid_id).changeCheckbox("         68", "COL1", True)
            session.findById(grid_id).selectItem("         68", "COL2")
            session.findById(grid_id).ensureVisibleHorizontalItem("         68", "COL2")
            session.findById(grid_id).changeCheckbox("         68", "COL2", True)

            session.findById(grid_id).collapseNode("         65")
            session.findById(grid_id).expandNode("         69")
            session.findById(grid_id).selectItem("         73", "COL1")
            session.findById(grid_id).ensureVisibleHorizontalItem("         73", "COL1")
            session.findById(grid_id).topNode = "          0"
            session.findById(grid_id).changeCheckbox("         73", "COL1", True)
            session.findById(grid_id).selectItem("         73", "COL2")
            session.findById(grid_id).ensureVisibleHorizontalItem("         73", "COL2")
            session.findById(grid_id).changeCheckbox("         73", "COL2", True)

            session.findById(grid_id).collapseNode("         69")
            session.findById(grid_id).expandNode("         74")
            session.findById(grid_id).selectItem("         77", "COL1")
            session.findById(grid_id).ensureVisibleHorizontalItem("         77", "COL1")
            session.findById(grid_id).topNode = "         74"
            session.findById(grid_id).changeCheckbox("         77", "COL1", True)
            session.findById(grid_id).selectItem("         77", "COL2")
            session.findById(grid_id).ensureVisibleHorizontalItem("         77", "COL2")
            session.findById(grid_id).changeCheckbox("         77", "COL2", True)

            session.findById(grid_id).collapseNode("         74")
            session.findById(grid_id).topNode = "          0"

            # ======================= MOVE DISPLAY FIELDS =======================
            display_field_rows = [7, 8, 5, 10, 11, 8, 9]
            display_table_id = "wnd[0]/path/to/file"
            btn_mvup_id = "wnd[0]/path/to/file"

            for row in display_field_rows:
                session.findById(display_table_id).getAbsoluteRow(row).selected = True
                session.findById(btn_mvup_id).press()

            time.sleep(0.5)

            # ======================= SELECTION FIELD REORDER =======================
            session.findById("wnd[0]").maximize()
            session.findById("wnd[0]/path/to/file").select()

            session.findById("wnd[0]/path/to/file").getAbsoluteRow(7).selected = True
            session.findById("wnd[0]/path/to/file").setFocus()
            session.findById("wnd[0]/path/to/file").caretPosition = 0

            # Press Move Up multiple times
            for _ in range(7):
                session.findById("wnd[0]/path/to/file").press()

            session.findById("wnd[0]/path/to/file").getAbsoluteRow(0).selected = False
            session.findById("wnd[0]/path/to/file").getAbsoluteRow(7).selected = True

            # Move up again
            for _ in range(6):
                session.findById("wnd[0]/path/to/file").press()

            session.findById("wnd[0]/path/to/file").getAbsoluteRow(1).selected = False
            session.findById("wnd[0]/path/to/file").getAbsoluteRow(11).selected = True
            session.findById("wnd[0]/path/to/file").setFocus()
            session.findById("wnd[0]/path/to/file").caretPosition = 0

            # Move up multiple times
            for _ in range(9):
                session.findById("wnd[0]/path/to/file").press()

            # Deselect and select rows
            session.findById("wnd[0]/path/to/file").getAbsoluteRow(2).selected = False
            for row in [9, 10, 11]:
                session.findById("wnd[0]/path/to/file").getAbsoluteRow(row).selected = True

            # More move up presses
            for _ in range(6):
                session.findById("wnd[0]/path/to/file").press()

            # Deselect other rows and focus on row 8
            for row in [3, 4, 5]:
                session.findById("wnd[0]/path/to/file").getAbsoluteRow(row).selected = False

            session.findById("wnd[0]/path/to/file").getAbsoluteRow(8).selected = True
            session.findById("wnd[0]/path/to/file").setFocus()
            session.findById("wnd[0]/path/to/file").caretPosition = 0

            # Final move up presses
            for _ in range(2):
                session.findById("wnd[0]/path/to/file").press()

            # ======================= FINAL VARIANT SETUP =======================
            session.findById("wnd[0]").maximize()
            session.findById("wnd[0]/path/to/file").getAbsoluteRow(8).selected = "false"
            session.findById("wnd[0]/path/to/file").press()
            session.findById("wnd[0]/path/to/file").press()
            session.findById("wnd[0]/path/to/file").press()
            session.findById("wnd[1]/path/to/file").press()
            session.findById("wnd[1]/path/to/file").select()
            session.findById("wnd[1]/path/to/file").text = "DUMMY_ID"
            session.findById("wnd[1]/path/to/file").text = "DUMMY_ID"
            session.findById("wnd[1]/path/to/file").text = "DUMMY_ID"
            session.findById("wnd[1]/path/to/file").text = "DUMMY_ID"
            session.findById("wnd[1]/path/to/file").setFocus()
            session.findById("wnd[1]/path/to/file").caretPosition = 10
            session.findById("wnd[1]/path/to/file").press()
            session.findById("wnd[0]/path/to/file").press()
            session.findById("wnd[1]/path/to/file").select()
            session.findById("wnd[1]/path/to/file").text = "A"
            session.findById("wnd[1]/path/to/file").text = "I"
            session.findById("wnd[1]/path/to/file").text = "N"
            session.findById("wnd[1]/path/to/file").setFocus()
            session.findById("wnd[1]/path/to/file").caretPosition = 1
            session.findById("wnd[1]/path/to/file").press()
            session.findById("wnd[0]/path/to/file").press()
            session.findById("wnd[0]/path/to/file").text = "DEFAULT_VIEW"
            session.findById("wnd[0]/path/to/file").text = "Default_View"
            session.findById("wnd[0]/path/to/file").setFocus()
            session.findById("wnd[0]/path/to/file").caretPosition = 12
            session.findById("wnd[0]/path/to/file").press()

            time.sleep(2)  # Allow time for the query to save
            print("SQVI setup completed successfully. Running threads...")
            # ======================= PREP SESSIONS (main thread only) =======================
            sapguiauto_main = win32com.client.GetObject("SAPGUI")
            application_main = sapguiauto_main.GetScriptingEngine
            connection_main = application_main.Children(0)

            # Ensure 2 sessions exist
            while connection_main.Children.Count < 2:
                connection_main.Children(0).createSession()
                time.sleep(3)

            run_threads()
        except Exception as e:
            print(f"Error during SQVI setup: {e!r}")
        finally:
            pythoncom.CoUninitialize()

    # Run setup
    sqvi_setup()

# ======================= SAVE LOG =======================
log_df = pd.DataFrame(log_records)
log_file = os.path.join(log_file_location, f"session_run_log_{week_label}.xlsx")
log_df.to_excel(log_file, index=False)
print(f"📊 Session log saved: {log_file}")
print("🎉 SQVI and Greenlight ran simultaneously successfully!")
time.sleep(3)  # Ensure threads complete before proceeding

 # SCRIPT 2: DATA FORMATTING=
# ----------------- CLEANING HELPERS ------------------
def clean_vendor_number(x):
    try:
        return str(int(float(x))) if pd.notna(x) else ""
    except:
        return ""

def clean_company_code(x):
    try:
        return str(int(float(x))) if pd.notna(x) else ""
    except:
        return ""

# ----------------- CLOSE EXCEL FILES ------------------
def close_all_excel_files():
    try:
        time.sleep(3)
        excel = win32.GetActiveObject("Excel.Application")
        for workbook in list(excel.Workbooks):
            workbook.Close(SaveChanges=True)
        excel.Quit()
        print("All Excel files closed.")
    except Exception as e:
        print(f"No running instance of Excel found or an error occurred.\nError: {e}")

close_all_excel_files()

time.sleep(5) 
#--------------------------------------------------------------- GREENLIGHT DATA FORMATTING ------------------------------------------------------------------------------
filename = f"{week_label} Greenlight.xlsx"
file_path = os.path.join(greenlight_export_path, filename)
rename_map = {
    "IV category": "IV_Category", "Supplier": "Supplier", "Company Code": "Company_Code",
    "GLN": "GLN", "VAT Registration No.": "VAT_Registration_No", "Name 1": "Name",
    "Inv. created Immediately": "Invoices_Created",
    "Inv. posted immediately": "Invoices_Posted_Without_Error",
    "Inv. percent Immediately": "Success_Rate", "Country": "Country_Code",
    "Group key": "Group_Key", "Account Group": "Account_Group", "Gross invoice amount": "Gross_Invoice_Amount",
    "Currency": "Currency"
}

desired_order = [
    "Financial_Year", "Financial_Week", "IV_Category", "IV_Description", "Supplier",
    "Company_Code", "Country", "GLN", "VAT_Registration_No", "Name",
    "Invoices_Created", "Invoices_Posted_Without_Error", "Invoices_Including_Exceptions",
    "Success_Rate", "Country_Code", "Group_Key", "Account_Group", "Gross_Invoice_Amount", "Currency"
]

if os.path.exists(file_path):
    df = pd.read_excel(file_path, engine='openpyxl')
    df.columns = df.columns.str.strip()
    df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns}, inplace=True)
    df["GLN"] = pd.to_numeric(df["GLN"], errors='coerce').fillna(0).astype(int).astype(str)

    for col in desired_order:
        if col not in df.columns:
            df[col] = ""

    df["Financial_Year"] = financial_year
    df["Financial_Week"] = df["Date"].apply(get_financial_week) if "Date" in df.columns else week_label
    df["IV_Description"] = df["IV_Category"].apply(lambda x: "Non-EDI" if x == 3 else "EDI")

    if "IV_Desc" in df.columns:
        df.drop(columns=["IV_Desc"], inplace=True)

    df["Invoices_Created"] = pd.to_numeric(df["Invoices_Created"], errors='coerce').fillna(0)
    df["Invoices_Posted_Without_Error"] = pd.to_numeric(df["Invoices_Posted_Without_Error"], errors='coerce').fillna(0)
    df["Invoices_Including_Exceptions"] = df["Invoices_Created"] - df["Invoices_Posted_Without_Error"]

    df["Success_Rate"] = np.where(
        df["Invoices_Created"] != 0,
        df["Invoices_Posted_Without_Error"] / df["Invoices_Created"],
        0
    )
    df["Success_Rate"] = df["Success_Rate"].apply(lambda x: f"{x:.0%}")

    df = df.iloc[:-1]

    extra_cols = [col for col in df.columns if col not in desired_order]
    df = df[desired_order + extra_cols]

    # Merge Country
    country_df = pd.read_excel(country_code_file_path, engine='openpyxl')
    country_df.columns = country_df.columns.str.strip()
    if 'Company Code' in country_df.columns:
        country_df.rename(columns={"Company Code": "Company_Code"}, inplace=True)
    if 'Country' in df.columns:
        df.drop(columns=['Country'], inplace=True)

    df = df.merge(country_df[['Company_Code', 'Country']], on='Company_Code', how='left')
    if 'Country_x' in df.columns and 'Country_y' in df.columns:
        df.drop(columns=['Country_x', 'Country_y'], inplace=True)

    # Reinsert Country column
    cols = list(df.columns)
    if 'Country' in cols and 'Company_Code' in cols:
        cols.remove('Country')
        cols.insert(cols.index('Company_Code') + 1, 'Country')
        df = df[cols]

    df = df.drop(columns=[col for col in df.columns if col == 'Country_Code' or col.startswith('Country_Code.')], errors='ignore')
    df.to_excel(file_path, index=False, sheet_name="Sheet1")

    try:
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        wb = excel.Workbooks.Open(file_path)
        ws = wb.Sheets(1)
        ws.Columns("H").NumberFormat = "@"
        wb.Save()
        wb.Close()
        excel.Quit()
    except Exception as e:
        print(f"Excel formatting failed: {e}")

    print(f"Greenlight data formatted and saved to {file_path}")

#--------------------------------------------------------------- EXCEPTIONS DATA FORMATTING ------------------------------------------------------------------------------
# ----------------- COLUMN RENAMING & ORDER ------------------
rename_map = {
    "CoCd": "Company Code",
    "IDoc number": "IDoc number",
    "Exception code": "Exception Code",
    "IDOC status": "IDOC Status",
    "ET created date": "ET created date",
    "Partner number": "Partner number",
    "Exception type": "Message Type",
    "Pur. Doc.": "Purchase Order",
    "Name 1": "Vendor Name",
    "Reference": "Reference",
    "Supplier": "Vendor Number",
    "Exception code text": "Exception Code Text"
}

target_columns = [
    "Financial Year", "Financial Week", "Company Code", "Vendor Number", "Vendor Name", "ET created date",
    "Rank", "Message Type", "Exception Code", "Exception Code Text", "Exception Message", "Purchase Order",
    "Source File"
]

# ----------------- FIND LATEST WEEK FILE ------------------
def find_latest_week_file(folder):
    files = [f for f in os.listdir(folder) if re.match(r'exceptions_W\d{2}\.xlsx', f)]
    if not files:
        raise FileNotFoundError("No exceptions_Wxx.xlsx files found in the folder.")
    latest_week = max(files, key=lambda x: int(re.search(r'W(\d{2})', x).group(1)))
    return os.path.join(folder, latest_week)

exceptions_file_path = find_latest_week_file(exceptions_export_path)
print(f"Processing latest file: {os.path.basename(exceptions_file_path)}")

# ----------------- LOAD DATA ------------------
df = pd.read_excel(exceptions_file_path)
lookup_df = pd.read_excel(exceptions_lookup_file, sheet_name="Sheet1")

# Strip spaces from column names
df.columns = df.columns.str.strip()
lookup_df.columns = lookup_df.columns.str.strip()

print("Columns in exceptions file:", df.columns.tolist())
print("Columns in lookup file:", lookup_df.columns.tolist())

# ----------------- TEMPORARY RENAME FOR MERGE ------------------
df.rename(columns={"Exception code": "Exception Code", "Exception code text": "Exception Code Text"}, inplace=True)

# ----------------- MERGE WITH LOOKUP ------------------
df = df.merge(
    lookup_df,
    how="left",
    on="Exception Code",
    suffixes=("", "_lookup")  # avoids _x/path/to/file
)

# ----------------- CLEAN UP MERGED COLUMNS ------------------
# Combine Exception Code Text from original and lookup
if "Exception Code Text_lookup" in df.columns:
    df["Exception Code Text"] = df["Exception Code Text"].combine_first(df.pop("Exception Code Text_lookup"))

# Remove any other remaining _lookup duplicates automatically
df = df.loc[:, ~df.columns.str.endswith("_lookup")]

# ----------------- RENAME COLUMNS ------------------
df.rename(columns=rename_map, inplace=True)


# ----------------- CLEAN DATA ------------------
df["Vendor Number"] = df["Vendor Number"].apply(clean_vendor_number)
df["Company Code"] = df["Company Code"].apply(clean_company_code)

# Map Status values
if "Message Type" in df.columns:
    df["Message Type"] = df["Message Type"].replace({"E": "Failed", "W": "Warning"})

# ----------------- FINANCIAL WEEK ------------------
df["Financial Week"] = df["ET created date"].apply(get_financial_week)
df["Financial Year"] = financial_year

# ----------------- DELETE BLANK ROWS ------------------
df = df.dropna(subset=["Company Code", "Vendor Number", "Vendor Name"], how='all')

# ----------------- ADD SOURCE FILE ------------------
df["Source File"] = os.path.basename(exceptions_file_path)

# ----------------- CREATE WEEK FOLDER ------------------
week_label = df["Financial Week"].max() or "Unknown"
week_folder = os.path.join(exceptions_export_path, week_label)
os.makedirs(week_folder, exist_ok=True)
print(f"Created folder for week: {week_label}")

# ----------------- MOVE ORIGINAL FILE ------------------
original_dest = os.path.join(week_folder, f"Original_{os.path.basename(exceptions_file_path)}")
shutil.move(exceptions_file_path, original_dest)
print(f"Original file moved: {os.path.basename(original_dest)}")

# ----------------- SAVE FORMATTED FILE ------------------
formatted_file = os.path.join(week_folder, f"Exceptions_Formatted_{week_label}.xlsx")
df = df[[col for col in target_columns if col in df.columns]]
df.to_excel(formatted_file, index=False)
print(f"Formatted file saved: {os.path.basename(formatted_file)}")

# ----------------- GENERATE SUMMARY ------------------
# Group by all required columns except Exception Code Text
summary = df.groupby(
    ["Company Code", "Vendor Number", "Vendor Name", "Message Type", "Exception Code", "Financial Year", "Financial Week"]
).agg(
    Exception_Code_Count=pd.NamedAgg(column="Exception Code", aggfunc="count"),
    Exception_Code_Text=pd.NamedAgg(column="Exception Code Text", aggfunc=lambda x: x.mode()[0] if not x.mode().empty else ""),
    Rank=pd.NamedAgg(column="Rank", aggfunc=lambda x: x.mode()[0] if not x.mode().empty else ""),
    Unique_Purchase_Orders=pd.NamedAgg(column="Purchase Order", aggfunc=pd.Series.nunique),
    Vendor_Number_Unique=pd.NamedAgg(column="Vendor Number", aggfunc=pd.Series.nunique)
).reset_index()

# Reorder columns
final_cols = [
    "Financial Year", "Financial Week", "Company Code", "Vendor Number", "Vendor Name",
    "Message Type", "Exception Code", "Exception_Code_Text", "Exception_Code_Count",
    "Rank", "Unique_Purchase_Orders", "Vendor_Number_Unique"
]
summary = summary[final_cols]
summary.rename(columns={"Exception_Code_Text": "Exception Code Text", "Exception_Code_Count": "Exception Code Count"}, inplace=True)

# ----------------- SAVE SUMMARY FILE ------------------
summary_file = os.path.join(week_folder, f"Exceptions_Summary_{week_label}.xlsx")
summary.to_excel(summary_file, index=False)
print(f"Summary file saved: {os.path.basename(summary_file)}")

# SCRIPT 3: ACCESS UPLOAD===
# ================== DATABASE CONNECTION ==================
conn_str = f'DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={access_db_path};'
conn = pyodbc.connect(conn_str)
cursor = conn.cursor()

# ================== HELPERS ==================
def safe_int_str(val):
    try:
        return str(int(float(val))) if pd.notna(val) and str(val).strip() else None
    except:
        return None


def format_percent(val):
    if pd.isna(val):
        return ''
    try:
        val = float(str(val).replace('%', '').strip())
        if val <= 1.0:
            val *= 100
        return f"{round(val)}%"
    except:
        return str(val)


# ================== IMPORT EXCEPTIONS ==================
def import_exceptions(batch_size=100):
    week_folders = [
        f for f in os.listdir(exceptions_export_path)
        if os.path.isdir(os.path.join(exceptions_export_path, f)) and re.match(r"W\d{2}", f)
    ]
    newest_week_folder = sorted(week_folders, key=lambda x: int(x[1:]))[-1]

    exceptions_file_name = f"Exceptions_Summary_{newest_week_folder}.xlsx"
    exceptions_file_path = os.path.join(exceptions_export_path, newest_week_folder, exceptions_file_name)

    print(f"Importing Exceptions from {newest_week_folder} folder")

    summary_df = pd.read_excel(exceptions_file_path)
    summary_df.columns = summary_df.columns.str.strip()
    summary_df = summary_df[summary_df["Financial Week"] == newest_week_folder]

    rows_inserted = 0
    skipped_rows = []

    insert_sql = '''
        INSERT INTO Exceptions
        ([Financial Year], [Financial Week], [Company Code], [Vendor Number], [Vendor Name],
         [Message Type], [Exception Code], [Exception Code Text], [Exception Code Count], [Rank],
         [Unique_Purchase_Orders])
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    '''

    # Prepare insert values
    insert_values = []
    for _, row in summary_df.iterrows():
        values = (
            int(row['Financial Year']) if pd.notna(row['Financial Year']) else None,
            str(row['Financial Week']),
            str(int(float(row['Company Code']))) if pd.notna(row['Company Code']) and str(row['Company Code']).strip() != "" else None,
            int(float(row['Vendor Number'])) if pd.notna(row['Vendor Number']) else None,
            str(row['Vendor Name']) if pd.notna(row['Vendor Name']) else None,
            str(row['Message Type']) if pd.notna(row['Message Type']) else None,
            str(row['Exception Code']) if pd.notna(row['Exception Code']) else None,
            str(row['Exception Code Text']) if pd.notna(row['Exception Code Text']) else None,
            int(row['Exception Code Count']) if pd.notna(row['Exception Code Count']) else 0,
            int(row['Rank']) if pd.notna(row['Rank']) else None,
            int(row['Unique_Purchase_Orders']) if pd.notna(row['Unique_Purchase_Orders']) else None
        )
        insert_values.append(values)

    # Insert in batches
    for i in range(0, len(insert_values), batch_size):
        batch = insert_values[i:i+batch_size]
        for values in batch:
            try:
                cursor.execute(insert_sql, values)
                rows_inserted += 1
            except Exception as e:
                skipped_rows.append((values, str(e)))
        conn.commit()  # commit after each batch

    print(f"Exceptions Import completed: {rows_inserted} rows inserted, {len(skipped_rows)} skipped.")
    if skipped_rows:
        print("Skipped rows and errors:")
        for val, err in skipped_rows:
            print(val, "->", err)


# ================== IMPORT GREENLIGHT ==================
def import_greenlight(batch_size=500):
    greenlight_files = [
        f for f in os.listdir(greenlight_export_path)
        if re.match(r"W\d{2} Greenlight\.xlsx", f)
    ]
    newest_greenlight_file = sorted(
        greenlight_files,
        key=lambda x: int(re.findall(r"W(\d{2})", x)[0])
    )[-1]
    greenlight_file_path = os.path.join(greenlight_export_path, newest_greenlight_file)

    print(f"Importing Greenlight Data from {newest_greenlight_file}")

    df = pd.read_excel(greenlight_file_path)
    df.columns = [
        "Financial_Year", "Financial_Week", "IV_Category", "IV_Description",
        "Supplier", "Company_Code", "Country", "GLN", "VAT_Registration_No", "Name",
        "Invoices_Created", "Invoices_Posted_Without_Error",
        "Invoices_Including_Exceptions", "Success_Rate", "Group_Key",
        "Account_Group", "Gross_Invoice_Amount", "Currency"
    ]
    df.columns = df.columns.str.strip()

    # Convert columns used for duplicate checking to string
    for col in ["Financial_Year", "Financial_Week", "IV_Category", "Supplier", "Company_Code"]:
        df[col] = df[col].fillna(0).astype(str).str.strip()

    # Drop duplicates based on key columns
    df = df.drop_duplicates(subset=["Financial_Year", "Financial_Week", "IV_Category", "Supplier", "Company_Code"])
    print(f"{len(df)} unique rows to insert after duplicate filtering.")

    insert_sql = '''
    INSERT INTO Greenlight
    ([Financial_Year], [Financial_Week], [IV_Category], [IV_Description],
    [Supplier], [Company_Code], [Country], [GLN], [VAT_Registration_No], [Name],
    [Invoices_Created], [Invoices_Posted_Without_Error], [Invoices_Including_Exceptions], 
    [Success_Rate], [Group_Key], [Account_Group], [Gross_Invoice_Amount], [Currency])
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    '''

    insert_values = []
    for _, row in df.iterrows():
        vals = (
            int(row['Financial_Year']) if row['Financial_Year'].isdigit() else 0,
            row['Financial_Week'],
            int(row['IV_Category']) if row['IV_Category'].isdigit() else 0,
            str(row['IV_Description']),
            row['Supplier'],
            row['Company_Code'],
            str(row['Country']),
            str(row['GLN']),
            str(row['VAT_Registration_No']),
            str(row['Name']),
            int(row['Invoices_Created']) if pd.notna(row['Invoices_Created']) else 0,
            int(row['Invoices_Posted_Without_Error']) if pd.notna(row['Invoices_Posted_Without_Error']) else 0,
            int(row['Invoices_Including_Exceptions']) if pd.notna(row['Invoices_Including_Exceptions']) else 0,
            format_percent(row['Success_Rate']),
            str(row['Group_Key']),
            str(row['Account_Group']),
            float(row['Gross_Invoice_Amount']) if pd.notna(row['Gross_Invoice_Amount']) else 0.0,
            str(row['Currency'])
        )
        insert_values.append(vals)

    inserted = 0
    skipped_rows = []

    for i in range(0, len(insert_values), batch_size):
        batch = insert_values[i:i+batch_size]
        for values in batch:
            try:
                cursor.execute(insert_sql, values)
                inserted += 1
            except Exception as e:
                skipped_rows.append((values, str(e)))
        conn.commit()  # commit each batch

    print(f"Greenlight Import completed: {inserted} rows inserted, {len(skipped_rows)} skipped.")
    if skipped_rows:
        print("Skipped rows and errors:")
        for val, err in skipped_rows:
            print(val, "->", err)

# # ================== 12-WEEK CLEANUP ==================
# def cleanup_old_weeks(table_name, week_column="Financial_Week", keep_weeks=12, batch_size=1):
#     table_name_br = f"[{table_name}]"
#     week_column_br = f"[{week_column}]"

#     # Get all distinct weeks
#     cursor.execute(f"SELECT DISTINCT {week_column_br} FROM {table_name_br}")
#     all_weeks = [row[0] for row in cursor.fetchall()]

#     # Filter out None or empty values
#     all_weeks = [w.strip() for w in all_weeks if w and str(w).strip()]

#     if not all_weeks:
#         print(f"No weeks found in {table_name} to clean up.")
#         return

#     # Sort weeks numerically
#     sorted_weeks = sorted(all_weeks, key=lambda x: int(re.findall(r"W(\d+)", x)[0]))

#     # Determine weeks to delete
#     weeks_to_delete = sorted_weeks[:-keep_weeks]  # keep only the most recent `keep_weeks` weeks
#     if not weeks_to_delete:
#         print(f"No old weeks to delete from {table_name}")
#         return

#     total_deleted = 0
#     for i in range(0, len(weeks_to_delete), batch_size):
#         batch = weeks_to_delete[i:i+batch_size]
#         placeholders = ",".join("?" * len(batch))
#         delete_sql = f"DELETE FROM {table_name_br} WHERE {week_column_br} IN ({placeholders})"
#         cursor.execute(delete_sql, batch)
#         total_deleted += len(batch)

#     print(f"Deleted {total_deleted} old weeks from {table_name}: {weeks_to_delete}")

# ================== RUN IMPORT + CLEANUP ==================
import_greenlight()
# cleanup_old_weeks("Greenlight", week_column="Financial_Week")  # Cleanup old weeks in Greenlight

import_exceptions()
# cleanup_old_weeks("Exceptions", week_column="Financial Week")  # Cleanup old weeks in Exceptions

conn.commit()
cursor.close()
conn.close()

print("All data imported successfully and database connection closed.")