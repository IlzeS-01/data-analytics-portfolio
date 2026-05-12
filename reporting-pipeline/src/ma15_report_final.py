import os
import win32com.client
import time
import clipboard

# SAP CONNECTION
sapguiauto = win32com.client.GetObject("SAPGUI")
application = sapguiauto.GetScriptingEngine
connection = application.Children(0)
session = connection.Children(0)

# DASHBOARD
session.findById("wnd[0]").maximize()
session.findById("wnd[0]/path/to/file").expandNode("F00003")
session.findById("wnd[0]/path/to/file").selectedNode = "F00082"
session.findById("wnd[0]/path/to/file").topNode = "Favo"
session.findById("wnd[0]/path/to/file").doubleClickNode("F00082")
session.findById("wnd[1]/path/to/file").press()
session.findById("wnd[2]/path/to/file").text = "RPT_DSHBRD"
session.findById("wnd[2]/path/to/file").text = "DUMMY_USER"
session.findById("wnd[2]/path/to/file").setFocus()
session.findById("wnd[2]/path/to/file").caretPosition = 9
session.findById("wnd[2]/path/to/file").press()
session.findById("wnd[1]/path/to/file").press()
session.findById("wnd[0]/path/to/file").pressToolbarButton("&MB_VARIANT")
session.findById("wnd[1]/path/to/file").setCurrentCell(5,"TEXT")
session.findById("wnd[1]/path/to/file").selectedRows = "5"
session.findById("wnd[1]/path/to/file").clickCurrentCell()
session.findById("wnd[0]/path/to/file").setCurrentCell(-1,"")
session.findById("wnd[0]/path/to/file").selectAll()
session.findById("wnd[0]/path/to/file").contextMenu()
session.findById("wnd[0]/path/to/file").selectContextMenuItem("&XXL")
session.findById("wnd[1]/path/to/file").press()
session.findById("wnd[1]/path/to/file").text = "DSHBRD_DATA.XLSX"
session.findById("wnd[1]/path/to/file").caretPosition = 16
session.findById("wnd[1]/path/to/file").press()

# Path to export folder
export_folder = r"/path/to/file"

# Specific file name
file_name = "DSHBRD_DATA.XLSX"
file_path = os.path.join(export_folder, file_name)

# Check if file exists
if not os.path.exists(file_path):
    raise FileNotFoundError(f"File {file_name} not found in {export_folder}")

# Open Excel file
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False  # Keep Excel in the background
wb = excel.Workbooks.Open(file_path)
ws = wb.Sheets("Sheet1")  # Adjust sheet name if needed

# Read data range (column B, adjust as needed)
data_range = ws.Range("B2:B10000").Value  # Returns a tuple of tuples


wb.Close(SaveChanges=False)
excel.Quit()
sap_data = [str(row[0]).replace('#', '').strip() for row in data_range if row and row[0] is not None]
data_string = "\r\n".join(sap_data)
clipboard.copy(data_string)
print(data_string)

# NEW SESSION
session.createSession()
time.sleep(2) 

new_session = None
for i in range(connection.Children.Count):
    new_session = connection.Children(i)

session = new_session  # Switch to the new session
print("Switched to the new SAP session!")

# VENDOR DETAIL
session.findById("wnd[0]").maximize()
session.findById("wnd[0]/path/to/file").expandNode("F00003")
session.findById("wnd[0]/path/to/file").selectedNode = "F00018"
session.findById("wnd[0]/path/to/file").topNode = "Favo"
session.findById("wnd[0]/path/to/file").doubleClickNode("F00018")
session.findById("wnd[0]/path/to/file").text = "MA15"
session.findById("wnd[0]/path/to/file").setFocus()
session.findById("wnd[0]/path/to/file").caretPosition = 4
session.findById("wnd[0]/path/to/file").press()
session.findById("wnd[1]/path/to/file").press()
session.findById("wnd[1]/path/to/file").press()
session.findById("wnd[1]/path/to/file").press()
session.findById("wnd[0]/path/to/file").key = "ALL"
session.findById("wnd[0]/path/to/file").setFocus()
session.findById("wnd[0]/path/to/file").press()
session.findById("wnd[0]/path/to/file").pressToolbarContextButton("&MB_EXPORT")
session.findById("wnd[0]/path/to/file").selectContextMenuItem("&XXL")
session.findById("wnd[1]/path/to/file").press()
session.findById("wnd[1]/path/to/file").text = "MA15_LISTED_DATA.XLSX"
session.findById("wnd[1]/path/to/file").caretPosition = 21
session.findById("wnd[1]/path/to/file").press()

# Excel Data Extraction From MA15 Listed
# Path to export folder
export_folder = r"/path/to/file"

# Specific file name
file_name = "MA15_LISTED_DATA.XLSX"
file_path = os.path.join(export_folder, file_name)

# Check if file exists
if not os.path.exists(file_path):
    raise FileNotFoundError(f"File {file_name} not found in {export_folder}")

# Open Excel file
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False  # Keep Excel in the background
wb = excel.Workbooks.Open(file_path)
ws = wb.Sheets("Sheet1")  # Adjust sheet name if needed

# Read data range (column C, adjust as needed)
data_range = ws.Range("C2:C10000").Value  # Returns a tuple of tuples


wb.Close(SaveChanges=False)
excel.Quit()
sap_data = [str(row[0]).replace('#', '').strip() for row in data_range if row and row[0] is not None]
data_string = "\r\n".join(sap_data)
clipboard.copy(data_string)
print(data_string)

# NEW SESSION
session.createSession()
time.sleep(2) 

new_session = None
for i in range(connection.Children.Count):
    new_session = connection.Children(i)

session = new_session  # Switch to the new session
print("Switched to the new SAP session!")

# ME2M
session.findById("wnd[0]").maximize()
session.findById("wnd[0]/path/to/file").text = "ME2M"
session.findById("wnd[0]").sendVKey(0)
session.findById("wnd[0]/path/to/file").press()
session.findById("wnd[1]/path/to/file").press()
session.findById("wnd[1]/path/to/file").currentCellColumn = "TEXT"
session.findById("wnd[1]/path/to/file").selectedRows = "0"
session.findById("wnd[1]/path/to/file").doubleClickCurrentCell()
session.findById("wnd[0]/path/to/file").press()
session.findById("wnd[1]/path/to/file").press()
session.findById("wnd[1]/path/to/file").press()
session.findById("wnd[1]/path/to/file").press()
session.findById("wnd[0]/path/to/file").text = "20.12.2024"
session.findById("wnd[0]/path/to/file").setFocus()
session.findById("wnd[0]/path/to/file").caretPosition = 10
session.findById("wnd[0]/path/to/file").press()
session.findById("wnd[0]/path/to/file").press()
session.findById("wnd[1]/path/to/file").press()
session.findById("wnd[1]/path/to/file").text = "PO_DATA.XLSX"
session.findById("wnd[1]/path/to/file").caretPosition = 12
session.findById("wnd[1]/path/to/file").press()
session.findById("wnd[0]/path/to/file").press()
session.findById("wnd[0]/path/to/file").setCurrentCell(-1,"SUPERFIELD")
session.findById("wnd[0]/path/to/file").selectColumn("SUPERFIELD")
session.findById("wnd[0]/path/to/file").contextMenu()
session.findById("wnd[0]/path/to/file").selectContextMenuItem("&COL0")
session.findById("wnd[1]/path/to/file").currentCellRow = 1
session.findById("wnd[1]/path/to/file").selectedRows = "1"
session.findById("wnd[1]/path/to/file").press()
session.findById("wnd[1]/path/to/file").setCurrentCell(30,"DO_SUM")
session.findById("wnd[1]/path/to/file").firstVisibleRow = 18
session.findById("wnd[1]/path/to/file").selectedRows = "30"
session.findById("wnd[1]/path/to/file").clickCurrentCell()
session.findById("wnd[1]/path/to/file").press()
session.findById("wnd[0]/path/to/file").press()
session.findById("wnd[1]/path/to/file").press()
session.findById("wnd[1]/path/to/file").text = "DELIVERY_DATA.XLSX"
session.findById("wnd[1]/path/to/file").caretPosition = 18
session.findById("wnd[1]/path/to/file").press()

# Excel Data Extraction From MA15 Listed
# Path to export folder
export_folder = r"/path/to/file"

# Specific file name
file_name = "MA15_LISTED_DATA.XLSX"
file_path = os.path.join(export_folder, file_name)

# Check if file exists
if not os.path.exists(file_path):
    raise FileNotFoundError(f"File {file_name} not found in {export_folder}")

# Open Excel file
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False  # Keep Excel in the background
wb = excel.Workbooks.Open(file_path)
ws = wb.Sheets("Sheet1")  # Adjust sheet name if needed

# Read data range (column C, adjust as needed)
data_range = ws.Range("C2:C10000").Value  # Returns a tuple of tuples


wb.Close(SaveChanges=False)
excel.Quit()
sap_data = [str(row[0]).replace('#', '').strip() for row in data_range if row and row[0] is not None]
data_string = "\r\n".join(sap_data)
clipboard.copy(data_string)
print(data_string)

# NEW SESSION
session.createSession()
time.sleep(2) 

new_session = None
for i in range(connection.Children.Count):
    new_session = connection.Children(i)

session = new_session  # Switch to the new session
print("Switched to the new SAP session!")

# VENDOR DETAIL REPORT G* STORES-------------------
session.findById("wnd[0]").maximize()
session.findById("wnd[0]/path/to/file").expandNode("F00003")
session.findById("wnd[0]/path/to/file").selectedNode = "F00018"
session.findById("wnd[0]/path/to/file").topNode = "Favo"
session.findById("wnd[0]/path/to/file").doubleClickNode("F00018")
session.findById("wnd[0]/path/to/file").press()
session.findById("wnd[1]/path/to/file").press()
session.findById("wnd[0]/path/to/file").press()
session.findById("wnd[1]/path/to/file").press()
session.findById("wnd[1]/path/to/file").press()
session.findById("wnd[1]/path/to/file").press()
session.findById("wnd[0]/path/to/file").press()
session.findById("wnd[0]/path/to/file").pressToolbarContextButton("&MB_EXPORT")
session.findById("wnd[0]/path/to/file").selectContextMenuItem("&XXL")
session.findById("wnd[1]/path/to/file").press()
session.findById("wnd[1]/path/to/file").text = "G_STORE_LISTED.XLSX"
session.findById("wnd[1]/path/to/file").caretPosition = 19
session.findById("wnd[1]/path/to/file").press()

# Excel Data Extraction From G* STORES Listed
# Path to export folder
export_folder = r"/path/to/file"

# Specific file name
file_name = "G_STORE_LISTED.XLSX"
file_path = os.path.join(export_folder, file_name)

# Check if file exists
if not os.path.exists(file_path):
    raise FileNotFoundError(f"File {file_name} not found in {export_folder}")

# Open Excel file
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False  # Keep Excel in the background
wb = excel.Workbooks.Open(file_path)
ws = wb.Sheets("Sheet1")  # Adjust sheet name if needed

# Read data range (column C, adjust as needed)
data_range = ws.Range("C2:C10000").Value  # Returns a tuple of tuples


wb.Close(SaveChanges=False)
excel.Quit()
sap_data = [str(row[0]).replace('#', '').strip() for row in data_range if row and row[0] is not None]
data_string = "\r\n".join(sap_data)
clipboard.copy(data_string)
print(data_string)

# NEW SESSION
session.createSession()
time.sleep(2) 

new_session = None
for i in range(connection.Children.Count):
    new_session = connection.Children(i)

session = new_session  # Switch to the new session
print("Switched to the new SAP session!")

# SE16N G* STORES PO'S
session.findById("wnd[0]").maximize()
session.findById("wnd[0]/path/to/file").text = "SE16N"
session.findById("wnd[0]").sendVKey(0) 
session.findById("wnd[0]/path/to/file").text = "EKPO"
session.findById("wnd[0]/path/to/file").setFocus()
session.findById("wnd[0]/path/to/file").caretPosition = 4
session.findById("wnd[0]").sendVKey(0)
session.findById("wnd[0]/path/to/file").setFocus()
session.findById("wnd[0]/path/to/file").press()
session.findById("wnd[1]/path/to/file").press()
session.findById("wnd[1]/path/to/file").press()
session.findById("wnd[1]/path/to/file").press()
session.findById("wnd[0]/path/to/file").setFocus()
session.findById("wnd[0]/path/to/file").press()
session.findById("wnd[1]/path/to/file").currentCellRow = 3
session.findById("wnd[1]/path/to/file").selectedRows = "3"
session.findById("wnd[1]/path/to/file").doubleClickCurrentCell()
session.findById("wnd[0]/path/to/file").text = ""
session.findById("wnd[0]/path/to/file").text = "G*"
session.findById("wnd[0]/path/to/file").setFocus()
session.findById("wnd[0]/path/to/file").caretPosition = 0
session.findById("wnd[0]/path/to/file").press()
session.findById("wnd[0]/path/to/file").pressToolbarButton("&MB_VARIANT")
session.findById("wnd[1]/path/to/file").setCurrentCell(4,"TEXT")
session.findById("wnd[1]/path/to/file").selectedRows = "4"
session.findById("wnd[1]/path/to/file").clickCurrentCell()
session.findById("wnd[0]/path/to/file").pressToolbarContextButton("&MB_EXPORT")
session.findById("wnd[0]/path/to/file").selectContextMenuItem("&XXL")
session.findById("wnd[1]/path/to/file").press()
session.findById("wnd[1]/path/to/file").text = "STORE_PO.XLSX"
session.findById("wnd[1]/path/to/file").caretPosition = 13
session.findById("wnd[1]/path/to/file").press()

# Excel Data Extraction From G* STORE PO
# Path to export folder
export_folder = r"/path/to/file"

# Specific file name
file_name = "STORE_PO.XLSX"
file_path = os.path.join(export_folder, file_name)

# Check if file exists
if not os.path.exists(file_path):
    raise FileNotFoundError(f"File {file_name} not found in {export_folder}")

# Open Excel file
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False  # Keep Excel in the background
wb = excel.Workbooks.Open(file_path)
ws = wb.Sheets("Sheet1")  # Adjust sheet name if needed

# Read data range (column C, adjust as needed)
data_range = ws.Range("A2:A10000").Value  # Returns a tuple of tuples


wb.Close(SaveChanges=False)
excel.Quit()
sap_data = [str(row[0]).replace('#', '').strip() for row in data_range if row and row[0] is not None]
data_string = "\r\n".join(sap_data)
clipboard.copy(data_string)
print(data_string)

# NEW SESSION
session.createSession()
time.sleep(2) 

new_session = None
for i in range(connection.Children.Count):
    new_session = connection.Children(i)

session = new_session  # Switch to the new session
print("Switched to the new SAP session!")

# SE16N G* STORES PO DELIVERY
session.findById("wnd[0]").maximize()
session.findById("wnd[0]/path/to/file").text = "SE16N"
session.findById("wnd[0]").sendVKey(0)
session.findById("wnd[0]/path/to/file").text = "/path/to/file"
session.findById("wnd[0]/path/to/file").caretPosition = 15
session.findById("wnd[0]").sendVKey(0)
session.findById("wnd[0]/path/to/file").setFocus()
session.findById("wnd[0]/path/to/file").press()
session.findById("wnd[1]/path/to/file").press()
session.findById("wnd[1]/path/to/file").press()

# Excel Data Extraction From G* STORE PO'S
# Path to export folder
export_folder = r"/path/to/file"

# Specific file name
file_name = "STORE_PO.XLSX"
file_path = os.path.join(export_folder, file_name)

# Check if file exists
if not os.path.exists(file_path):
    raise FileNotFoundError(f"File {file_name} not found in {export_folder}")

# Open Excel file
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False  # Keep Excel in the background
wb = excel.Workbooks.Open(file_path)
ws = wb.Sheets("Sheet1")  # Adjust sheet name if needed

# Read data range (column C, adjust as needed)
data_range = ws.Range("C2:C500000").Value  # Returns a tuple of tuples


wb.Close(SaveChanges=False)
excel.Quit()
sap_data = [str(row[0]).replace('#', '').strip() for row in data_range if row and row[0] is not None]
data_string = "\r\n".join(sap_data)
clipboard.copy(data_string)
print(data_string)

session.findById("wnd[0]/path/to/file").setFocus()
session.findById("wnd[0]/path/to/file").press()
session.findById("wnd[1]/path/to/file").press()
session.findById("wnd[1]/path/to/file").press()
session.findById("wnd[1]/path/to/file").press()
session.findById("wnd[0]/path/to/file").setFocus()
session.findById("wnd[0]/path/to/file").press()
session.findById("wnd[1]/path/to/file").currentCellRow = 3
session.findById("wnd[1]/path/to/file").selectedRows = "3"
session.findById("wnd[1]/path/to/file").doubleClickCurrentCell()
session.findById("wnd[0]/path/to/file").setFocus()
session.findById("wnd[0]/path/to/file").press()
session.findById("wnd[1]/path/to/file").currentCellRow = 3
session.findById("wnd[1]/path/to/file").selectedRows = "3"
session.findById("wnd[1]/path/to/file").doubleClickCurrentCell()
session.findById("wnd[0]/path/to/file").text = "X"
session.findById("wnd[0]/path/to/file").text = "MA15"
session.findById("wnd[0]/path/to/file").setFocus
session.findById("wnd[0]/path/to/file").caretPosition = 4
session.findById("wnd[0]/path/to/file").verticalScrollbar.position = 1
session.findById("wnd[0]/path/to/file").verticalScrollbar.position = 2
session.findById("wnd[0]/path/to/file").verticalScrollbar.position = 3
session.findById("wnd[0]/path/to/file").verticalScrollbar.position = 4
session.findById("wnd[0]/path/to/file").verticalScrollbar.position = 5
session.findById("wnd[0]/path/to/file").verticalScrollbar.position = 6
session.findById("wnd[0]/path/to/file").verticalScrollbar.position = 7
session.findById("wnd[0]/path/to/file").verticalScrollbar.position = 8
session.findById("wnd[0]/path/to/file").setFocus()
session.findById("wnd[0]/path/to/file").press()
session.findById("wnd[1]/path/to/file").text = "NB"
session.findById("wnd[1]/path/to/file").text = "AO"
session.findById("wnd[1]/path/to/file").text = "MDAO"
session.findById("wnd[1]/path/to/file").text = "ZTCN"
session.findById("wnd[1]/path/to/file").text = "ZA"
session.findById("wnd[1]/path/to/file").setFocus()
session.findById("wnd[1]/path/to/file").caretPosition = 2
session.findById("wnd[1]/path/to/file").press()
session.findById("wnd[0]/path/to/file").setFocus()
session.findById("wnd[0]/path/to/file").press()
session.findById("wnd[1]/path/to/file").currentCellRow = 8
session.findById("wnd[1]/path/to/file").selectedRows = "8"
session.findById("wnd[1]/path/to/file").press()
session.findById("wnd[0]/path/to/file").text = "20.12.2024"
session.findById("wnd[0]/path/to/file").setFocus()
session.findById("wnd[0]/path/to/file").caretPosition = 10
session.findById("wnd[0]/path/to/file").press()
session.findById("wnd[0]/path/to/file").pressToolbarButton("&MB_VARIANT")
session.findById("wnd[1]/path/to/file").setCurrentCell(13,"TEXT")
session.findById("wnd[1]/path/to/file").firstVisibleRow = 4
session.findById("wnd[1]/path/to/file").selectedRows = "13"
session.findById("wnd[1]/path/to/file").clickCurrentCell()
session.findById("wnd[0]/path/to/file").pressToolbarContextButton("&MB_EXPORT")
session.findById("wnd[0]/path/to/file").selectContextMenuItem("&XXL")
session.findById("wnd[1]/path/to/file").press()
session.findById("wnd[1]/path/to/file").text = "STORE_PO_DELV.XLSX"
session.findById("wnd[1]/path/to/file").caretPosition = 18
session.findById("wnd[1]/path/to/file").press()

# Excel Data Extraction
export_folder = r"/path/to/file"
file_name = "STORE_PO_DELV.XLSX"
file_path = os.path.join(export_folder, file_name)

if not os.path.exists(file_path):
    raise FileNotFoundError(f"File {file_name} not found in {export_folder}")

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
wb = excel.Workbooks.Open(file_path)
ws = wb.Sheets("Sheet1")
data_range = ws.Range("D2:D10000").Value
wb.Close(SaveChanges=False)
excel.Quit()

sap_data = [str(row[0]).replace('#', '').strip() for row in data_range if row and row[0] is not None]
data_string = "\r\n".join(sap_data)
clipboard.copy(data_string)
print(data_string)

# -Close Last SAP Session ----------------
sapguiauto = win32com.client.GetObject("SAPGUI")
application = sapguiauto.GetScriptingEngine
connection = application.Children(application.Children.Count - 1)
session = connection.Children(connection.Children.Count - 1)
session.EndTransaction()
session.findById("wnd[0]").Close()

print("Closed the last opened SAP session.")
time.sleep(1)

# -SAP CONNECTION
sapguiauto = win32com.client.GetObject("SAPGUI")
application = sapguiauto.GetScriptingEngine
connection = application.Children(0)
session = connection.Children(0)

# NEW SESSION
session.createSession()
time.sleep(2) 

new_session = None
for i in range(connection.Children.Count):
    new_session = connection.Children(i)

session = new_session  # Switch to the new session
print("Switched to the new SAP session!")

# -Run SAP Script ----------------
session.findById("wnd[0]").maximize()
session.findById("wnd[0]/path/to/file").text = "MB51"
session.findById("wnd[0]").sendVKey(0)
session.findById("wnd[0]/path/to/file").press()
session.findById("wnd[1]/path/to/file").press()
session.findById("wnd[1]/path/to/file").press()
session.findById("wnd[1]/path/to/file").press()

# -Extract & Copy Data ----------------
time.sleep(5)
if not os.path.exists(file_path):
    raise FileNotFoundError(f"File {file_name} not found in {export_folder}")

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
wb = excel.Workbooks.Open(file_path)
ws = wb.Sheets("Sheet1")
data_range = ws.Range("E2:E10000").Value
wb.Close(SaveChanges=False)
excel.Quit()

sap_data = [str(row[0]).replace('#', '').strip() for row in data_range if row and row[0] is not None]
data_string = "\r\n".join(sap_data)
clipboard.copy(data_string)
print("Data copied to clipboard.")

# -SAP Filtering ----------------
session.findById("wnd[0]/path/to/file").press()
session.findById("wnd[1]/path/to/file").press()
session.findById("wnd[1]/path/to/file").press()
session.findById("wnd[1]/path/to/file").press()
session.findById("wnd[0]/path/to/file").press()
session.findById("wnd[1]/path/to/file").press()
session.findById("wnd[1]/path/to/file").text = "251"
session.findById("wnd[1]/path/to/file").caretPosition = 3
session.findById("wnd[1]/path/to/file").press()
session.findById("wnd[0]/path/to/file").press()

# -Export Data ----------------
session.findById("wnd[0]/path/to/file").setCurrentCell(-1, "MAKTX")
session.findById("wnd[0]/path/to/file").selectColumn("MAKTX")
session.findById("wnd[0]/path/to/file").contextMenu()
session.findById("wnd[0]/path/to/file").selectContextMenuItem("&XXL")
session.findById("wnd[1]/path/to/file").press()
session.findById("wnd[1]/path/to/file").text = "STORE_POSTING_DATE.XLSX"
session.findById("wnd[1]/path/to/file").caretPosition = 23
session.findById("wnd[1]/path/to/file").press()