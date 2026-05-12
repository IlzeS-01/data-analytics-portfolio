# -*- coding: utf-8 -*- 
"""
CDA B_BBEE Emails V12 – Batch-Driven Logic with Max Emails & Skipped Statuses

Features:
- Sends emails in batches per CDA (default 10 per run)
- Skips vendors with statuses: Vendor Closed, Non Compliant, Cert Received, Cert is valid for FY26
- Marks vendors as sent in tracker (JSON only)
- Maximum emails per CDA per run enforced (respects test mode)
- Logs whether an email is NEW or REMINDER
- Preserves all original sheets
- Formats Vendors sheet in Excel
"""

import os, sys, subprocess, importlib, time, warnings, urllib.parse, json, shutil
from datetime import datetime
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from filelock import FileLock
import win32com.client as win32
warnings.simplefilter("ignore", category=FutureWarning)
warnings.simplefilter("ignore", category=UserWarning)

def safe_install(mod, pkg=None):
    pkg = pkg or mod
    try:
        importlib.import_module(mod)
        print(f"{pkg} already installed")
    except ImportError:
        print(f"Installing {pkg}...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg])

def check_and_install_packages():
    required = {
        "pandas": "pandas",
        "openpyxl": "openpyxl",
        "filelock": "filelock",
        "win32com.client": "pywin32",
    }
    for mod, pkg in required.items():
        safe_install(mod, pkg)

check_and_install_packages()

# --- Configuration ---
BASE_DIR = os.path.expanduser(r"~\Demo Company\Finance and Administration Team Site - Documents\Datos\B-BBEE")
REAL_DIR = os.path.expanduser(r"~\Demo Company\Finance and Administration Team Site - Documents\B-BBEE")
os.makedirs(BASE_DIR, exist_ok=True)

INPUT_FILE = os.path.join(REAL_DIR, "FY26 B-BBEE Procurement File (CDA).xlsx")
if not os.path.exists(INPUT_FILE):
    raise FileNotFoundError(f"Source file not found: {INPUT_FILE}")
print(f"Using source file: {INPUT_FILE}")

EMAIL_DELAY = 2
MAX_EMAILS_PER_CDA_PER_DAY = 10
MAX_TEST_EMAILS = 1

TRACKING_DIR = os.path.join(BASE_DIR, "B-BBEE Mail Tracker")
os.makedirs(TRACKING_DIR, exist_ok=True)
TRACKING_FILE = os.path.join(TRACKING_DIR, "B-BBEE_Email_Tracker.json")
LOCK_FILE = TRACKING_FILE + ".lock"

STATUS_TEMPLATES = {
    "To be contacted": "Please contact this vendor ASAP to obtain the necessary B-BBEE documentation.",
    "Vendor Closed": "No further action is required. This vendor is marked as closed.",
    "Emailed Vendor": "Vendor has been emailed. Follow up if no response.",
    "Called Vendor": "Vendor has been called. Follow up if no feedback.",
    "Cert is valid for FY26": "No action required. Certificate is valid for FY26.",
    "In progress": "Documentation process is in progress. Monitor and follow up within 1 month.",
    "Waiting on feedback from the vendor": "Waiting on vendor feedback.",
    "Non Compliant": "Please ensure you received the relevant documentation to confirm that the vendor is non-compliant.",
    "Cert Received": "B-BBEE certificate has been received. Please rename the certificate.",
    "Contact buyer for support": "The buyer has been contacted. Follow up if no feedback."
}

# --- Test Mode ---
TEST_MODE = True
TEST_EMAIL = "demo.user@company.com"
TEST_CC = ["demodirector.user@company.com"]

if TEST_MODE:
    TRACKING_FILE = os.path.join(TRACKING_DIR, "B-BBEE_Email_Tracker_TEST.json")
    LOCK_FILE = TRACKING_FILE + ".lock"
    print(f"TEST MODE ON – using test tracker: {TRACKING_FILE}")

# --- Button Config ---
BUTTON_COLORS = ["#DUMMY_ID", "#DUMMY_ID"]
BUTTON_OPTIONS = [
    "Non Compliant", "Vendor Closed", "In progress", "Waiting on feedback from the vendor",
    "Cert Received", "Emailed Vendor", "Called Vendor", "Contact buyer for support",
    "Vendor Applying for Certificate", "Not My Vendor"
]

def generate_buttons(vendor_number, vendor_name):
    buttons_html = '<table border="0" cellspacing="0" cellpadding="0" style="margin:10px 0;">'
    custom_message = f"Hi There,\n\nResponse for vendor {vendor_number} - {vendor_name} :\n"
    for i, label in enumerate(BUTTON_OPTIONS):
        if i % 4 == 0:
            if i > 0: buttons_html += '<tr><td height="10"></path/to/file'
            buttons_html += '<tr>'
        color = BUTTON_COLORS[i % len(BUTTON_COLORS)]
        subject = urllib.parse.quote(f"B-BBEE Response: {vendor_number} – {vendor_name} – {label}")
        body_text = urllib.parse.quote(f"{custom_message}\n{label}")
        buttons_html += f'''
        <td align="center" bgcolor="{color}" style="border-radius:4px; padding:2px 16px;">
            <a href="mailto:{TEST_EMAIL if TEST_MODE else "demo.user@company.com"}?subject={subject}&body={body_text}"
               style="font-family:Segoe UI, Arial, sans-serif; font-size:10pt; color:#ffffff; text-decoration:none; display:inline-block; padding:2px 0;">
               {label}
            </path/to/file
        </path/to/file
        <td width="5">&nbsp;</path/to/file
        '''
        if (i + 1) % 4 == 0: buttons_html += '</path/to/file'
    if len(BUTTON_OPTIONS) % 4 != 0: buttons_html += '</path/to/file'
    buttons_html += '</path/to/file'
    return buttons_html

# --- DataFrame Helpers ---
def copy_vendors_sheet():
    xls = pd.ExcelFile(INPUT_FILE)
    sheet_to_use = "Vendors" if "Vendors" in xls.sheet_names else xls.sheet_names[0]
    df = pd.read_excel(xls, sheet_name=sheet_to_use)
    for col in ['Status', 'CDA', 'Company']:
        if col in df.columns: df[col] = df[col].astype(str).str.strip()
    print(f"Copied '{sheet_to_use}' sheet from {INPUT_FILE}")
    return df

def validate_columns(df):
    required_cols = ['Vendor account', 'Company', 'VAT Number', 'CDA', 'Status', 'CDA Email', 'BMC']
    missing = [c for c in required_cols if c not in df.columns]
    if missing: raise ValueError(f"Missing columns in Excel: {missing}")
    df['Vendor account'] = pd.to_numeric(df['Vendor account'], errors='coerce').astype('Int64')
    if 'Sent' not in df.columns: df['Sent'] = ''
    return df

# --- Tracker ---
def load_tracker():
    try:
        with open(TRACKING_FILE, "r", encoding="utf-8") as f: return json.load(f)
    except (json.JSONDecodeError, FileNotFoundError): return []

def save_tracker(tracker):
    with FileLock(LOCK_FILE):
        with open(TRACKING_FILE, "w", encoding="utf-8") as f: json.dump(tracker, f, indent=4, ensure_ascii=False)

# --- Email ---
def create_new_email_body(row):
    action_required = STATUS_TEMPLATES.get(row['Status'], '')
    vendor_number = row['Vendor account']
    vendor_name = row['Company']
    buttons_html = generate_buttons(vendor_number, vendor_name)
    return f"""
    <html><body style="font-family:'Segoe UI', Arial, sans-serif; font-size:10pt; color:#DUMMY_ID;">
        <p>Dear {row['CDA']},</path/to/file
        <p><strong style="color:#c82613;">Action Required:</path/to/file {action_required}</path/to/file
        <table border='1' cellpadding='5' cellspacing='0' 
            style="border-collapse:collapse; table-layout:fixed; width:473px; font-family:'Segoe UI', Arial, sans-serif; font-size:10pt; color:#DUMMY_ID;">
            <tr><th style="background-color:#DUMMY_ID; text-align:center; color:white;">Field</path/to/file
                <th style="background-color:#DUMMY_ID; text-align:center; color:white;">Vendor Details</path/to/file
            <tr><td>Vendor Number</path/to/file
            <tr><td>Vendor Name</path/to/file
            <tr><td>Tax Number</path/to/file'VAT Number']</path/to/file
            <tr><td>Status</path/to/file'Status']</path/to/file
            <tr><td>BMC</path/to/file'BMC']</path/to/file
        </path/to/file
        {buttons_html}
        <p>Regards,<br>Demo Company Team</path/to/file
    </path/to/file
    """

def send_email(row):
    recipient = TEST_EMAIL if TEST_MODE else row['CDA Email']
    cc_list = TEST_CC if TEST_MODE else []
    subject = f"B-BBEE Follow-Up: {row['Vendor account']} – {row['Company']}"
    body = create_new_email_body(row)
    try:
        outlook = win32.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)
        mail.To = recipient
        mail.CC = ';'.join(cc_list) if cc_list else ''
        mail.Subject = subject
        mail.HTMLBody = body
        mail.Send()
        print(f"Email sent to {recipient}")
    except Exception as e:
        print(f"Failed to send email to {recipient}: {e}")

# --- Log Tracker + Mark Excel ---
def log_email_json(row, df):
    data = {
        "Date/path/to/file": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Recipient": TEST_EMAIL if TEST_MODE else row['CDA Email'],
        "Vendor Number": row['Vendor account'],
        "Vendor Name": row['Company'],
        "Status": row['Status'],
        "BMC": row['BMC'],
        "Email Sent": True
    }
    tracker = load_tracker()
    tracker.append(data)
    save_tracker(tracker)
    df.loc[df['Vendor account'] == row['Vendor account'], 'Sent'] = 'X'
    print(f"Logged & marked Sent for vendor {row['Vendor account']}")

# --- Excel Formatting ---
def format_vendors_sheet(file_path):
    wb = load_workbook(file_path)
    sheet = wb['Vendors']

    header_fill = PatternFill(start_color="DUMMY_ID", end_color="DUMMY_ID", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    header_alignment = Alignment(horizontal='center', vertical='center')
    for cell in sheet[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment

    sheet.freeze_panes = sheet['A2']

    for col in sheet.columns:
        max_length = max(len(str(cell.value)) if cell.value else 0 for cell in col)
        sheet.column_dimensions[col[0].column_letter].width = max_length + 2
    sent_col = None
    for idx, cell in enumerate(sheet[1], start=1):
        if cell.value == 'Sent':
            sent_col = idx
            break
    if sent_col:
        for row in sheet.iter_rows(min_row=2, min_col=sent_col, max_col=sent_col):
            for cell in row:
                if str(cell.value).strip().upper() == 'X':
                    cell.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
                    cell.font = Font(bold=True, color="DUMMY_ID")
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                         top=Side(style='thin'), bottom=Side(style='thin'))
    for row in sheet.iter_rows():
        for cell in row:
            cell.border = thin_border

    wb.save(file_path)
    print(f"Vendors sheet formatted in {file_path}")

# --- Main Process ---
def process_bbbee_emails():
    df = copy_vendors_sheet()
    df = validate_columns(df)
    tracker = load_tracker()
    sent_vendors = {(t.get("Vendor Number"), t.get("Recipient")) for t in tracker}
    SKIP_STATUSES = {"Vendor Closed", "Non Compliant", "Cert Received", "Cert is valid for FY26"}

    cda_groups = df.groupby('CDA Email')
    for cda_email, group in cda_groups:
        emails_sent = 0
        max_emails = MAX_TEST_EMAILS if TEST_MODE else MAX_EMAILS_PER_CDA_PER_DAY
        eligible_rows = []
        for _, row in group.iterrows():
            vid = row['Vendor account']
            rec = row['CDA Email']
            if str(row.get('Sent', '')).strip().upper() == 'X': continue
            if row['Status'] in SKIP_STATUSES: continue
            if (vid, rec) in sent_vendors: continue
            eligible_rows.append(row)

        unsent_vendors = pd.DataFrame(eligible_rows)
        if unsent_vendors.empty: continue

        for _, row in unsent_vendors.iterrows():
            if emails_sent >= max_emails:
                break

            if str(row['CDA Email']).strip().lower() == "no cda assigned":
                print(f"Skipped vendor {row['Vendor account']} - No CDA Assigned")
                continue

            send_email(row)
            log_email_json(row, df)
            emails_sent += 1
            time.sleep(EMAIL_DELAY)

    # Backup
    backup_dir = os.path.join(os.path.dirname(INPUT_FILE), "Backups")
    os.makedirs(backup_dir, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_file = os.path.join(backup_dir, f"FY26_BBBEE_Backup_{timestamp}.xlsx")
    shutil.copy2(INPUT_FILE, backup_file)
    print(f"Backup created: {backup_file}")

    # Save updated Sent column and preserve all sheets
    wb = load_workbook(INPUT_FILE)
    if 'Vendors' in wb.sheetnames: wb.remove(wb['Vendors'])
    new_sheet = wb.create_sheet('Vendors', 0)
    for col_idx, col_name in enumerate(df.columns, start=1):
        new_sheet.cell(row=1, column=col_idx, value=col_name)
    for row_idx, row_data in enumerate(df.itertuples(index=False), start=2):
        for col_idx, value in enumerate(row_data, start=1):
            new_sheet.cell(row=row_idx, column=col_idx, value=value)
    wb.save(INPUT_FILE)
    print(f"'Sent' column updated in original file: {INPUT_FILE}")

    format_vendors_sheet(INPUT_FILE)

if __name__ == "__main__":
    process_bbbee_emails()
