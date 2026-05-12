import win32com.client
import time

EXPORT_PATH = r"/path/to/file Company\Finance and Administration Team Site - Documents\EDI\EDI Document Set Up Report\Raw Data"
BATCH_SIZE = 3000

SapGuiAuto = win32com.client.GetObject("SAPGUI")
application = SapGuiAuto.GetScriptingEngine
connection = application.Children(0)
session = connection.Children(0)

def export_table(table_name):
    print(f"Exporting {table_name}...")

    session.findById("wnd[0]/path/to/file").text = "/path/to/file"
    session.findById("wnd[0]").sendVKey(0)

    session.findById("wnd[0]/path/to/file").text = table_name
    session.findById("wnd[0]").sendVKey(0)

    session.findById("wnd[0]/path/to/file").text = "ACTIVE_VNDRS"
    session.findById("wnd[0]/path/to/file").text = ""

    session.findById("wnd[0]/path/to/file").text = "DUMMY_ID"
    session.findById("wnd[0]/path/to/file").text = "DUMMY_ID"

    session.findById("wnd[0]/path/to/file").press()

    export_current_result(table_name)

def export_current_result(filename):
    session.findById("wnd[0]/path/to/file").pressToolbarContextButton("&MB_EXPORT")
    session.findById("wnd[0]/path/to/file").selectContextMenuItem("&XXL")

    session.findById("wnd[1]/path/to/file").press()
    session.findById("wnd[1]/path/to/file").text = EXPORT_PATH
    session.findById("wnd[1]/path/to/file").text = f"{filename}.xlsx"
    session.findById("wnd[1]/path/to/file").press()

    print(f"{filename} exported.")
    time.sleep(3)

def get_clean_vendor_list():
    time.sleep(5)

    excel = win32com.client.GetObject(None, "Excel.Application")

    for wb in excel.Workbooks:
        if "LFB1" in wb.Name:
            ws = wb.Worksheets(1)

            last_row = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
            values = ws.Range(f"A2:A{last_row}").Value

            # Flatten + clean
            vendors = list(set([
                str(v[0]).strip() for v in values if v[0]
            ]))

            print(f"Vendors loaded: {len(vendors)}")
            return vendors

    print("LFB1 not found.")
    return []

def copy_to_clipboard(values):
    text = "\n".join(values)

    import win32clipboard
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardText(text)
    win32clipboard.CloseClipboard()

def run_table(table_name, btn_row=None, btn_col=None,
              extra_filter=None, filter_col=None,
              variant=None, vendors=None, vendor_field="LIFNR"):

    print(f"Running {table_name}...")

    session.findById("wnd[0]/path/to/file").text = "/path/to/file"
    session.findById("wnd[0]").sendVKey(0)

    session.findById("wnd[0]/path/to/file").text = table_name
    session.findById("wnd[0]").sendVKey(0)

    time.sleep(1)

    # Vendor-based tables
    if vendors:
        session.findById(
            f"wnd[0]/path/to/file"
        ).press()

        time.sleep(1)
        session.findById("wnd[1]/path/to/file").press()
        session.findById("wnd[1]/path/to/file").press()

    # Variant
    if variant:
        session.findById("wnd[0]/path/to/file").text = variant

    # Filters
    if extra_filter and filter_col is not None:
        session.findById(
            f"wnd[0]/path/to/file"
        ).text = extra_filter

    session.findById("wnd[0]/path/to/file").text = ""
    session.findById("wnd[0]/path/to/file").press()

    export_current_result(table_name)

# MAIN FLOW
export_table("LFA1")
export_table("LFB1")

vendors = get_clean_vendor_list()

for i in range(0, len(vendors), BATCH_SIZE):
    batch = vendors[i:i + BATCH_SIZE]

    print(f"Processing batch {i // BATCH_SIZE + 1}")

    copy_to_clipboard(batch)

    run_table("LFM1", 4, 1, vendors=batch, vendor_field="LIFNR")
    run_table("WYT3", 4, 1, extra_filter="OA", filter_col=5, vendors=batch, vendor_field="LIFNR")
    run_table("B048", 4, 4, variant="CREDIT_ADVIC", vendors=batch, vendor_field="LIFNR")
    run_table("B904", 4, 3, extra_filter="ZETI", filter_col=2, vendors=batch, vendor_field="LIFNR")

run_table("B507", extra_filter="ZNEU", filter_col=2)
print("ALL PROCESSING COMPLETE")