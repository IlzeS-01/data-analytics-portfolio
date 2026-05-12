from datetime import datetime, timedelta
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
try:
    import win32com.client as win32
except ImportError:
    win32 = None



# CONFIG 
EXCEL_PATH = Path(r"/path/to/file - Demo Company\Desktop\Weekly Summary.xlsx")
SHEET_NAME = "DailyLogs"

BOSS_NAME = "Boss"
EMAIL_TO = "demo.user@company.com"
SIGN_OFF_NAME = "Ilze"

FIELDS = [
    ("edi_integration", "EDI / Integration"),
    ("needs_blockers", "Needs / Blockers"),
    ("portal_tickets_4me", "Portal / Tickets / 4me"),
    ("other", "Other"),
]

COLUMNS = ["date"] + [field for field, _ in FIELDS]



# EXCEL STORAGE
class ExcelStorage:
    def __init__(self, excel_path: Path):
        self.excel_path = excel_path
        self.excel_path.parent.mkdir(parents=True, exist_ok=True)
        self._ensure_workbook_exists()

    def _ensure_workbook_exists(self):
        if not self.excel_path.exists():
            df = pd.DataFrame(columns=COLUMNS)
            with pd.ExcelWriter(self.excel_path, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name=SHEET_NAME, index=False)

    def read_all(self) -> pd.DataFrame:
        try:
            df = pd.read_excel(
                self.excel_path, sheet_name=SHEET_NAME, engine="openpyxl"
            )
        except ValueError:
            df = pd.DataFrame(columns=COLUMNS)

        for col in COLUMNS:
            if col not in df.columns:
                df[col] = ""

        df["date"] = pd.to_datetime(df["date"], errors="coerce")
        return df[COLUMNS].copy()

    def save_daily_entry(self, entry: dict):
        df = self.read_all()
        entry_date = pd.to_datetime(entry["date"])

        df = df[df["date"] != entry_date]
        new_row = pd.DataFrame([entry])
        new_row["date"] = pd.to_datetime(new_row["date"])

        combined = (
            pd.concat([df, new_row], ignore_index=True)
            .sort_values("date")
            .reset_index(drop=True)
        )

        combined["date"] = combined["date"].dt.strftime("%Y-%m-%d")
        with pd.ExcelWriter(self.excel_path, engine="openpyxl") as writer:
            combined.to_excel(writer, sheet_name=SHEET_NAME, index=False)

    def get_entry_by_date(self, date_str: str) -> dict:
        df = self.read_all()
        target = pd.to_datetime(date_str)
        row = df[df["date"] == target]
        if row.empty:
            return {}
        record = row.iloc[0].to_dict()
        return {
            field: "" if pd.isna(record.get(field)) else str(record.get(field))
            for field, _ in FIELDS
        }

    def get_week_entries(self, start: datetime, end: datetime) -> pd.DataFrame:
        df = self.read_all()
        mask = (df["date"] >= start.date()) & (df["date"] <= end.date())
        return df.loc[mask].sort_values("date")



# WEEKLY SUMMARY BUILDER
class WeeklySummaryBuilder:
    @staticmethod
    def current_week_range():
        today = datetime.now()
        monday = today - timedelta(days=today.weekday())
        friday = monday + timedelta(days=4)
        return monday, friday

    @staticmethod
    def subject(start, end):
        return f"Weekly Summary - {start:%d/%m/%Y} - {end:%d/%m/%Y}"

    @staticmethod
    def section(df, field, label):
        bullets = []

        for _, row in df.iterrows():
            raw_text = str(row.get(field, "")).strip()
            if not raw_text:
                continue

            date_str = pd.to_datetime(row["date"]).strftime("%d/%m/%Y")

            # Split multiple lines into separate bullets
            for line in raw_text.splitlines():
                line = line.strip()
                if line:
                    bullets.append(f"- {date_str}: {line}")

        if not bullets:
            bullets.append("- No updates recorded")

        return f"{label}:\n" + "\n".join(bullets)

    @staticmethod
    def email_body(df, boss_name, signoff):
        sections = [
            WeeklySummaryBuilder.section(df, field, label)
            for field, label in FIELDS
        ]
        return (
            f"Hi {boss_name},\n\n"
            "Please find below a brief summary for the week:"
            + "\n\n".join(sections)
            + f"\n\nRegards,\n{signoff}"
        )



# OUTLOOK
class OutlookDraftService:
    @staticmethod
    def create(to_email, subject, body):
        if win32 is None:
            raise RuntimeError("pywin32 not installed")
        outlook = win32.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To = to_email
        mail.Subject = subject
        mail.Body = body
        mail.Display()



# UI
class WeeklySummaryApp:
    def __init__(self, root):
        self.storage = ExcelStorage(EXCEL_PATH)

        self.root = root
        self.root.title("Weekly Summary")
        self.root.geometry("820x720")

        self.date_var = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        self.text_widgets = {}

        self._build_ui()
        self.load_entry()

    def _build_ui(self):
        header = ttk.Frame(self.root, padding=10)
        header.pack(fill="x")

        ttk.Entry(header, textvariable=self.date_var, width=14).pack(side="left")
        ttk.Button(header, text="Load", command=self.load_entry).pack(side="left", padx=5)
        ttk.Button(header, text="Save", command=self.save).pack(side="left", padx=5)
        ttk.Button(header, text="Preview", command=self.preview).pack(side="left", padx=5)
        ttk.Button(header, text="Outlook Draft", command=self.create_draft).pack(
            side="left", padx=5
        )

        body = ttk.Frame(self.root, padding=10)
        body.pack(fill="both", expand=True)

        for field, label in FIELDS:
            frame = ttk.LabelFrame(body, text=label)
            frame.pack(fill="x", pady=5)
            text = tk.Text(frame, height=5, wrap="word")
            text.pack(fill="x")
            self.text_widgets[field] = text

    def load_entry(self):
        entry = self.storage.get_entry_by_date(self.date_var.get())
        for field in self.text_widgets:
            self.text_widgets[field].delete("1.0", tk.END)
            self.text_widgets[field].insert("1.0", entry.get(field, ""))

    def save(self):
        entry = {"date": self.date_var.get()}
        for field, widget in self.text_widgets.items():
            entry[field] = widget.get("1.0", tk.END).strip()
        self.storage.save_daily_entry(entry)
        messagebox.showinfo("Saved", "Daily entry saved")

    def preview(self):
        start, end = WeeklySummaryBuilder.current_week_range()
        df = self.storage.get_week_entries(start, end)
        subject = WeeklySummaryBuilder.subject(start, end)
        body = WeeklySummaryBuilder.email_body(df, BOSS_NAME, SIGN_OFF_NAME)

        win = tk.Toplevel(self.root)
        win.title("Weekly Summary Preview")
        text = tk.Text(win, wrap="word")
        text.pack(fill="both", expand=True)
        text.insert("1.0", f"Subject: {subject}\n\n{body}")
        text.config(state="disabled")

    def create_draft(self):
        start, end = WeeklySummaryBuilder.current_week_range()
        df = self.storage.get_week_entries(start, end)
        subject = WeeklySummaryBuilder.subject(start, end)
        body = WeeklySummaryBuilder.email_body(df, BOSS_NAME, SIGN_OFF_NAME)
        OutlookDraftService.create(EMAIL_TO, subject, body)
        messagebox.showinfo("Success", "Outlook draft created")



# ENTRY POINT
if __name__ == "__main__":
    root = tk.Tk()
    WeeklySummaryApp(root)
    root.mainloop()