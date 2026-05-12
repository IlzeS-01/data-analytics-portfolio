# -*- coding: utf-8 -*-
"""
B-BBEE Response Tracker – Updates Excel Status Column based on email responses
and appends responses to tracker Excel with Test Mode support
Includes Excel formatting for Vendors sheet:
- Headers bold, centered, white text on dark blue
- Freeze top row
- Auto-adjust column widths
- 'Sent' column highlighted if X
- Newly updated 'Status' cells highlighted light blue
- Thin borders
"""

import os
import json
import pandas as pd
from dotenv import load_dotenv
load_dotenv()
import win32com.client as win32
from filelock import FileLock
from datetime import datetime, timedelta
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import re
import traceback
import argparse

# --- Configuration ---
BASE_DIR = os.path.expanduser(r"~\Demo Company\Finance and Administration Team Site - Documents\Datos\B-BBEE")
REAL_DIR = os.path.expanduser(r"~\Demo Company\Finance and Administration Team Site - Documents\B-BBEE")
os.makedirs(BASE_DIR, exist_ok=True)

INPUT_FILE = os.path.join(REAL_DIR, "FY26 B-BBEE Procurement File (CDA).xlsx")
if not os.path.exists(INPUT_FILE):
    raise FileNotFoundError(f"Source file not found: {INPUT_FILE}")
print(f"Using source file: {INPUT_FILE}")

TRACKING_DIR = os.path.join(BASE_DIR, "B-BBEE Mail Tracker")
os.makedirs(TRACKING_DIR, exist_ok=True)
TRACKING_FILE = os.path.join(TRACKING_DIR, "B-BBEE_Email_Tracker.json")
LOCK_FILE = TRACKING_FILE + ".lock"

# --- Test Mode ---
TEST_MODE = False
TEST_EMAIL = "demo.user@company.com"
TEST_CC = ["demo.user@company.com"]
if TEST_MODE:
    TRACKING_FILE = os.path.join(TRACKING_DIR, "B-BBEE_Email_Tracker_TEST.json")
    LOCK_FILE = TRACKING_FILE + ".lock"
    print(f"TEST MODE ON – using test tracker: {TRACKING_FILE}")

# --- Tracker functions ---
def load_tracker():
    with FileLock(LOCK_FILE):
        if os.path.exists(TRACKING_FILE):
            with open(TRACKING_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        else:
            return []

def save_tracker(tracker):
    with FileLock(LOCK_FILE):
        with open(TRACKING_FILE, "w", encoding="utf-8") as f:
            json.dump(tracker, f, indent=4, ensure_ascii=False)

# --- Load Excel sheet (robust version) ---
def load_excel():
    """
    Loads the Excel file (either .xls or .xlsx) and returns a cleaned DataFrame.
    Ensures 'Vendor account', 'Company', 'Status', and 'Sent' columns exist.
    """
    if not os.path.exists(INPUT_FILE):
        raise FileNotFoundError(f"Source file not found: {INPUT_FILE}")

    ext = os.path.splitext(INPUT_FILE)[1].lower()
    if ext == '.xlsx':
        engine = 'openpyxl'
    elif ext == '.xls':
        engine = 'xlrd'
    else:
        raise ValueError(f"Unsupported Excel file type: {ext}")

    try:
        xls = pd.ExcelFile(INPUT_FILE, engine=engine)
    except Exception as e:
        raise ValueError(f"Failed to read Excel file. It may be corrupted: {e}")

    sheet_to_use = "Vendors" if "Vendors" in xls.sheet_names else xls.sheet_names[0]
    df = pd.read_excel(xls, sheet_name=sheet_to_use, engine=engine)

    # Clean up columns
    df['Vendor account'] = pd.to_numeric(df.get('Vendor account', pd.Series()), errors='coerce').astype('Int64')
    df['Company'] = df.get('Company', pd.Series()).astype(str).str.strip()
    df['Status'] = df.get('Status', pd.Series()).astype(str).str.strip()
    if 'Sent' not in df.columns:
        df['Sent'] = ''
    if 'Responded' not in df.columns:
        df['Responded'] = ''

    print(f"Loaded sheet '{sheet_to_use}' from {INPUT_FILE} ({len(df)} rows).")
    return df

# --- Format Vendors sheet ---
def format_vendors_sheet(file_path, updated_status_rows=None):
    wb = load_workbook(file_path)
    if 'Vendors' not in wb.sheetnames:
        print("'Vendors' sheet not found. Skipping formatting.")
        return

    sheet = wb['Vendors']

    # Header formatting
    header_fill = PatternFill(start_color="DUMMY_ID", end_color="DUMMY_ID", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    header_alignment = Alignment(horizontal='center', vertical='center')
    for cell in sheet[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment

    # Freeze top row
    sheet.freeze_panes = sheet['A2']

    # Auto column width
    for col in sheet.columns:
        max_length = max(len(str(cell.value)) if cell.value else 0 for cell in col)
        sheet.column_dimensions[col[0].column_letter].width = max_length + 2

    # Identify columns
    sent_col = None
    status_col = None
    responded_col = None              # ← NEW

    for idx, cell in enumerate(sheet[1], start=1):
        if cell.value == 'Sent':
            sent_col = idx
        if cell.value == 'Status':
            status_col = idx
        if cell.value == 'Responded':  # ← NEW
            responded_col = idx

    # Highlight Sent cells
    if sent_col:
        for row in sheet.iter_rows(min_row=2, min_col=sent_col, max_col=sent_col):
            for cell in row:
                if str(cell.value).strip().upper() == 'X':
                    cell.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
                    cell.font = Font(bold=True, color="DUMMY_ID")

    # Highlight updated Status cells
    if status_col and updated_status_rows:
        light_blue_fill = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
        bold_font = Font(bold=True)
        for r_idx in updated_status_rows:
            cell = sheet.cell(row=r_idx + 2, column=status_col)
            cell.fill = light_blue_fill
            cell.font = bold_font
    # Highlight Responded cells (same light-blue style) when they contain "X"
    if responded_col:
        for row in sheet.iter_rows(min_row=2, min_col=responded_col, max_col=responded_col):
            for cell in row:
                if str(cell.value).strip().upper() == 'X':
                    cell.fill = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
                    cell.font = Font(bold=True)

    # Thin borders
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                         top=Side(style='thin'), bottom=Side(style='thin'))
    for row in sheet.iter_rows():
        for cell in row:
            cell.border = thin_border

    wb.save(file_path)
    print(f"Vendors sheet formatted in {file_path}")

# --- Save Excel sheet ---
def save_excel(df, updated_status_indices=None):
    """
    Saves the DataFrame to the 'Vendors' sheet in INPUT_FILE.
    Handles pd.NA values safely for openpyxl and applies formatting.
    """
    # Replace all pd.NA / NaN with None
    df_clean = df.copy()
    for col in df_clean.columns:
        df_clean[col] = df_clean[col].apply(lambda x: None if pd.isna(x) else x)

    # Load workbook and remove old Vendors sheet if exists
    wb = load_workbook(INPUT_FILE)
    if 'Vendors' in wb.sheetnames:
        wb.remove(wb['Vendors'])
    new_sheet = wb.create_sheet('Vendors', 0)

    # Write headers
    for col_idx, col_name in enumerate(df_clean.columns, start=1):
        new_sheet.cell(row=1, column=col_idx, value=col_name)

    # Write data
    for row_idx, row_data in enumerate(df_clean.itertuples(index=False), start=2):
        for col_idx, value in enumerate(row_data, start=1):
            new_sheet.cell(row=row_idx, column=col_idx, value=value)

    wb.save(INPUT_FILE)
    print(f"Excel sheet updated: {INPUT_FILE}")

    # Format Vendors sheet
    if updated_status_indices:
        format_vendors_sheet(INPUT_FILE, updated_status_rows=updated_status_indices)

# --- Append responses to tracker Excel ---
def append_to_tracker_excel(new_entries):
    tracker_excel_file = os.path.join(TRACKING_DIR, "B-BBEE_Email_Tracker.xlsx")
    sheet_name = "Response"

    if os.path.exists(tracker_excel_file):
        wb = load_workbook(tracker_excel_file)
    else:
        wb = Workbook()

    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
    else:
        ws = wb.create_sheet(sheet_name)
        headers = list(new_entries[0].keys())
        ws.append(headers)

    for entry in new_entries:
        ws.append(list(entry.values()))

    wb.save(tracker_excel_file)
    print(f"Appended {len(new_entries)} responses to tracker Excel: {tracker_excel_file}")

def _get_range_datetimes(range_key):
    """
    Returns a tuple (start_dt, end_dt) in local time for the provided range_key.
    If range_key == 'all', returns (None, None) meaning no filtering.
    Supported keys: 'today', 'yesterday', 'last7', 'all'
    """
    now = datetime.now()
    start = None
    end = None
    if range_key == 'today':
        start = datetime(now.year, now.month, now.day, 0, 0, 0)
        end = datetime(now.year, now.month, now.day, 23, 59, 59)
    elif range_key == 'yesterday':
        y = now - timedelta(days=1)
        start = datetime(y.year, y.month, y.day, 0, 0, 0)
        end = datetime(y.year, y.month, y.day, 23, 59, 59)
    elif range_key == 'last7':
        # Last 7 days inclusive of today (i.e., today and previous 6 days)
        from_day = now - timedelta(days=6)
        start = datetime(from_day.year, from_day.month, from_day.day, 0, 0, 0)
        end = datetime(now.year, now.month, now.day, 23, 59, 59)
    elif range_key == 'all':
        start, end = None, None
    else:
        # default to all
        start, end = None, None
    return start, end

def track_responses_and_update_excel(range_key='all'):
    tracker = load_tracker()
    df = load_excel()
    updated_status_indices = []
    new_entries = []
    skipped_subjects = []

    # Determine date range filter
    start_dt, end_dt = _get_range_datetimes(range_key)
    if start_dt and end_dt:
        print(f"Filtering messages: {range_key} -> {start_dt.isoformat()} to {end_dt.isoformat()}")
    else:
        print("Processing all messages (no date filter).")

    # helper to interpret true/path/to/file safely (kept inside function for minimal changes)
    def is_true(x):
        if isinstance(x, bool):
            return x
        if x is None:
            return False
        return str(x).strip().lower() in ("true", "yes", "y", "1")

    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)
    target_folder = inbox.Folders["B-BEE Flow - Response FY26"]

    messages = target_folder.Items
    messages.Sort("[ReceivedTime]", True)
    print(f"Found {messages.Count} messages in folder.")

    processed_count = 0
    appended_count = 0
    skipped_count = 0

    for i in range(1, messages.Count + 1):
        try:
            msg = messages.Item(i)
            # ReceivedTime is a datetime object provided by pywin32
            try:
                received_dt = msg.ReceivedTime  # datetime object
            except Exception:
                # If ReceivedTime is not available, skip
                continue

            # Apply date filter if specified
            if start_dt and end_dt:
                # ReceivedTime should be comparable with naive local datetimes; if it's timezone-aware you may need adjustments
                if received_dt < start_dt or received_dt > end_dt:
                    continue

            subject_original = (msg.Subject or "").strip()

            if "b-bbee response:" not in subject_original.lower():
                continue

            clean_subj = subject_original.replace("B-BBEE Response:", "").strip()
            vendor_number_match = re.search(r"\d+", clean_subj)
            if not vendor_number_match:
                skipped_subjects.append(subject_original)
                skipped_count += 1
                continue
            vendor_number = int(vendor_number_match.group(0))

            parts = re.split(r"\s*[-–—]\s*", clean_subj)
            response_text = parts[-1].strip() if len(parts) > 1 else "Unknown"

            sender = msg.SenderEmailAddress or ""
            received_time = received_dt.strftime("%Y-%m-%d %H:%M:%S")

            # ------------------ PATCH START ------------------
            # Find the original "Email Sent" entry for this vendor — but DO NOT modify it.
            sent_entry = None
            for entry in tracker:
                if str(entry.get("Vendor Number")) == str(vendor_number) and is_true(entry.get("Email Sent")):
                    sent_entry = entry
                    break

            # Create a separate response-only entry and append it (links back to sent entry if present)
            response_entry = {
                "Date/path/to/file": received_time,
                "Sender": sender,
                "Vendor Number": vendor_number,
                "Vendor Name": sent_entry.get("Vendor Name", "") if sent_entry else "",
                "Response": response_text,
                "Response Date": received_time,
                "Response Sender": sender,
                "RepliedTo": subject_original,
                "IsResponse": True,
                "Linked Sent Date/path/to/file": sent_entry.get("Date/path/to/file") if sent_entry else None,
                # Keep Email Sent context if present; otherwise False
                "Email Sent": sent_entry.get("Email Sent") if sent_entry else False
            }

            tracker.append(response_entry)
            new_entries.append({
                "Date/path/to/file": received_time,
                "Sender": sender,
                "Vendor Number": vendor_number,
                "Vendor Name": response_entry.get("Vendor Name", ""),
                "Response": response_text,
                "RepliedTo": subject_original,
                "Email Sent": response_entry.get("Email Sent"),
                "IsResponse": True,
                "Linked Sent Date/path/to/file": response_entry.get("Linked Sent Date/path/to/file")
            })

            # If there was no sent entry to link to, count it as appended (no prior send found)
            if sent_entry is None:
                appended_count += 1
            # ------------------- PATCH END -------------------

            # Update the Vendors sheet Status column if vendor exists in Excel
            mask = df['Vendor account'].fillna(0).astype(int) == vendor_number
            if mask.any():
                idx = df[mask].index[0]
                df.at[idx, 'Status'] = response_text
                df.at[idx, 'Responded'] = 'X'
                updated_status_indices.append(idx)
                processed_count += 1
                print(f"Updated Status for Vendor {vendor_number}: {response_text}")
            else:
                print(f"Vendor {vendor_number} not found in Excel (subject: {subject_original})")

            try:
                msg.UnRead = False
            except Exception:
                pass

        except Exception as e:
            skipped_subjects.append(subject_original if 'subject_original' in locals() else str(i))
            skipped_count += 1
            print(f"Error processing message at index {i}: {e}")
            traceback.print_exc()

    save_tracker(tracker)
    save_excel(df, updated_status_indices=updated_status_indices)
    if new_entries:
        append_to_tracker_excel(new_entries)

    if skipped_subjects:
        skipped_log = os.path.join(TRACKING_DIR, "skipped_subjects.log")
        with open(skipped_log, "a", encoding="utf-8") as f:
            f.write(f"\n\n--- Run at {datetime.now().isoformat()} ---\n")
            for s in skipped_subjects:
                f.write(s + "\n")
        print(f"{len(skipped_subjects)} subjects skipped — see {skipped_log}")

    print(f"Summary: Excel statuses updated: {processed_count}, tracker appended: {appended_count}, skipped: {skipped_count}.")

# --- Run ---
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="B-BBEE Response Tracker - process Outlook responses.")
    parser.add_argument("-r", "--range", choices=["today", "yesterday", "last7", "all"], default="last7",
                        help="Date range to process (default: last7). Options: today, yesterday, last7, all")
    args = parser.parse_args()

    # call main with selected date range
    track_responses_and_update_excel(range_key=args.range)
