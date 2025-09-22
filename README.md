
import io
import json
import datetime as dt
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

APP_DIR = Path.home() / ".lieu_days_manager"
APP_DIR.mkdir(exist_ok=True)
CFG_FILE = APP_DIR / "config.json"

# ========= Utilities to remember last path =========
def load_config():
    if CFG_FILE.exists():
        try:
            return json.loads(CFG_FILE.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}

def save_config(cfg: dict):
    try:
        CFG_FILE.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass

cfg = load_config()
last_path = cfg.get("last_excel_path", "")

# Hide Streamlit's error helper links and long details
st.set_option("client.showErrorDetails", False)
st.markdown("""
<style>
div[data-testid="stException"] a {display:none !important;}
div[data-testid="stException"] details {display:none !important;}
</style>
""", unsafe_allow_html=True)

st.set_page_config(page_title="Lieu Days Manager", page_icon="🗓️", layout="wide")

st.title("🗓️ Lieu Days Manager")
st.caption("Select employee → Add/Minus/Lost → pick start/end date or single date → write to Excel → see live balances")

# --- Helpers ---
def load_or_create_workbook(xlsx_path: Path):
    if xlsx_path.exists():
        wb = load_workbook(xlsx_path)
    else:
        wb = Workbook()
    return wb

def ensure_staff_sheet(wb):
    if "Staff" not in wb.sheetnames:
        ws = wb.create_sheet("Staff")
        ws["A1"] = "Staff"
    return wb["Staff"]

def get_staff_list(wb):
    staff = []
    if "Staff" in wb.sheetnames:
        ws = wb["Staff"]
        for r in range(2, ws.max_row + 1):
            val = ws.cell(row=r, column=1).value
            if isinstance(val, str) and val.strip():
                staff.append(val.strip())
    if not staff:
        ws0 = wb[wb.sheetnames[0]]
        for r in range(2, ws0.max_row + 1):
            val = ws0.cell(row=r, column=1).value
            if isinstance(val, str) and val.strip():
                staff.append(val.strip())
    return sorted(set(staff))

def ensure_transactions_sheet(wb):
    # Ensure Transactions sheet with Start/End Date columns (upgrade legacy if needed)
    if "Transactions" not in wb.sheetnames:
        ws = wb.create_sheet("Transactions")
        headers = ["Staff", "Start Date", "End Date", "Type", "Days", "Note", "Entered By", "Timestamp"]
        for i, h in enumerate(headers, start=1):
            ws.cell(row=1, column=i, value=h)
        return ws

    ws = wb["Transactions"]
    hdrs = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column+1)]
    if "End Date" not in hdrs:
        if ws.cell(row=1, column=2).value and str(ws.cell(row=1, column=2).value).strip().lower() == "date":
            ws.cell(row=1, column=2, value="Start Date")
        ws.insert_cols(3)
        ws.cell(row=1, column=3, value="End Date")
        for r in range(2, ws.max_row+1):
            start_val = ws.cell(row=r, column=2).value
            ws.cell(row=r, column=3, value=start_val)
        names = ["Staff", "Start Date", "End Date", "Type", "Days", "Note", "Entered By", "Timestamp"]
        for idx, nm in enumerate(names, start=1):
            if ws.cell(row=1, column=idx).value is None:
                ws.cell(row=1, column=idx, value=nm)
    return ws

def transactions_to_df(ws):
    headers = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column+1)]
    header_map = {str(h).strip(): i+1 for i, h in enumerate(headers) if h is not None}
    def col(name, fallback):
        return header_map.get(name, fallback)

    rows = []
    for r in range(2, ws.max_row+1):
        if all(ws.cell(row=r, column=c).value in (None, "") for c in range(1, max(8, ws.max_column)+1)):
            continue
        staff = ws.cell(row=r, column=col("Staff", 1)).value
        start = ws.cell(row=r, column=col("Start Date", 2)).value
        end   = ws.cell(row=r, column=col("End Date", 3)).value
        typ   = ws.cell(row=r, column=col("Type", 4)).value
        days  = ws.cell(row=r, column=col("Days", 5)).value
        note  = ws.cell(row=r, column=col("Note", 6)).value
        entby = ws.cell(row=r, column=col("Entered By", 7)).value
        ts    = ws.cell(row=r, column=col("Timestamp", 8)).value
        rows.append([staff, start, end, typ, days, note, entby, ts])
    df = pd.DataFrame(rows, columns=["Staff","Start Date","End Date","Type","Days","Note","Entered By","Timestamp"])

    if not df.empty:
        df["Type"] = df["Type"].astype(str).str.strip()
        df["Days"] = pd.to_numeric(df["Days"], errors="coerce").fillna(0.0)
        def coerce_date(x):
            if pd.isna(x):
                return pd.NaT
            if isinstance(x, dt.date) and not isinstance(x, dt.datetime):
                return dt.datetime.combine(x, dt.time())
            try:
                return pd.to_datetime(x)
            except:
                return pd.NaT
        df["Start Date"] = df["Start Date"].apply(coerce_date)
        df["End Date"]   = df["End Date"].apply(coerce_date)
    return df

def write_tx_row(ws, staff, start_when, end_when, entry_type, days, note, user):
    next_row = ws.max_row + 1
    ws.cell(row=next_row, column=1, value=staff)
    ws.cell(row=next_row, column=2, value=start_when)
    ws.cell(row=next_row, column=3, value=end_when)
    ws.cell(row=next_row, column=4, value=entry_type)
    ws.cell(row=next_row, column=5, value=float(days))
    ws.cell(row=next_row, column=6, value=note or "")
    ws.cell(row=next_row, column=7, value=user or "")
    ws.cell(row=next_row, column=8, value=dt.datetime.now())

def compute_balance(df, name):
    if df.empty:
        return 0.0, 0.0, 0.0, 0.0
    d = df[df["Staff"] == name].copy()
    earned = d.loc[d["Type"].str.lower() == "earned", "Days"].sum()
    taken  = d.loc[d["Type"].str.lower() == "taken",  "Days"].sum()
    lost   = d.loc[d["Type"].str.lower() == "lost",   "Days"].sum()
    bal = earned - taken - lost
    return float(earned), float(taken), float(lost), float(bal)

def ensure_summary_sheet(wb):
    first = wb.sheetnames[0]
    ws = wb[first]
    if first != "Summary":
        ws.title = "Summary"
    return wb["Summary"]

def style_header(ws, start_col, end_col):
    fill = PatternFill("solid", fgColor="D9E1F2")
    bold = Font(bold=True)
    center = Alignment(horizontal="center", vertical="center")
    thin = Side(style="thin", color="999999")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    rng = ws[f"{get_column_letter(start_col)}3:{get_column_letter(end_col)}3"]
    for row in rng:
        for c in row:
            c.fill = fill
            c.font = bold
            c.alignment = center
            c.border = border

def clear_table(ws, start_row, end_row, start_col, end_col):
    for r in range(start_row, end_row+1):
        for c in range(start_col, end_col+1):
            ws.cell(row=r, column=c, value=None)

def set_cell_value_safe(ws, row:int, col:int, value):
    for mr in ws.merged_cells.ranges:
        if (mr.min_row <= row <= mr.max_row) and (mr.min_col <= col <= mr.max_col):
            ws.cell(row=mr.min_row, column=mr.min_col, value=value)
            return
    ws.cell(row=row, column=col, value=value)

def rebuild_summary(ws, staff_list, df_tx, year:int):
    ws.sheet_view.showGridLines = True
    ws.freeze_panes = "B4"

    months = ["January","February","March","April","May","June","July","August","September","October","November","December"]
    headers = ["Staff list"] + months + ["Current balance"]

    set_cell_value_safe(ws, 1, 1, f"Lieu Days Summary — Year: {year}")
    last_col = 1 + len(headers) - 1
    set_cell_value_safe(ws, 1, last_col, f"Last update: {dt.datetime.now().strftime('%Y-%m-%d %H:%M')}")

    for i, h in enumerate(headers, start=1):
        ws.cell(row=3, column=i, value=h)
    style_header(ws, 1, len(headers))

    total_rows = len(staff_list)
    clear_table(ws, 4, 4+total_rows+2, 1, len(headers))
    for idx, name in enumerate(staff_list, start=4):
        ws.cell(row=idx, column=1, value=name)

    if df_tx is None or df_tx.empty:
        return

    df = df_tx.copy()
    df["Month"] = df["Start Date"].dt.month
    df["Year"]  = df["Start Date"].dt.year

    for i, name in enumerate(staff_list, start=4):
        d_name = df[df["Staff"] == name]
        if d_name.empty:
            continue
        earned = d_name.loc[d_name["Type"].str.lower() == "earned", "Days"].sum()
        taken  = d_name.loc[d_name["Type"].str.lower() == "taken",  "Days"].sum()
        lost   = d_name.loc[d_name["Type"].str.lower() == "lost",   "Days"].sum()
        bal = earned - taken - lost
        ws.cell(row=i, column=1+12+1, value=round(float(bal),2))

        dny = d_name[d_name["Year"] == year]
        for m in range(1,13):
            dm = dny[dny["Month"] == m]
            if dm.empty:
                val = 0.0
            else:
                e = dm.loc[dm["Type"].str.lower() == "earned", "Days"].sum()
                t = dm.loc[dm["Type"].str.lower() == "taken",  "Days"].sum()
                l = dm.loc[dm["Type"].str.lower() == "lost",   "Days"].sum()
                val = e - t - l
            ws.cell(row=i, column=1+m, value=round(float(val),2))

# --- Sidebar: file + options ---
st.sidebar.header("Excel file")
mode = st.sidebar.radio("How would you like to open the Excel file?", ["Use path on my computer", "Upload .xlsx"], index=0)

xlsx_bytes = None
xlsx_path = None
remember = st.sidebar.checkbox("Remember this path", value=True)
clear_saved = st.sidebar.button("Clear saved path")

if clear_saved:
    cfg["last_excel_path"] = ""
    save_config(cfg)
    st.sidebar.success("Saved path cleared.")

if mode == "Use path on my computer":
    default_value = last_path if isinstance(last_path, str) else ""
    xlsx_path_str = st.sidebar.text_input("Path to your Excel (.xlsx)", value=default_value, placeholder=r"C:\Users\you\Documents\lieu_days.xlsx")
    if xlsx_path_str:
        xlsx_path = Path(xlsx_path_str)
        if xlsx_path.exists():
            st.sidebar.success("File found ✅")
            if remember:
                cfg["last_excel_path"] = str(xlsx_path)
                save_config(cfg)
        else:
            st.sidebar.info("Enter a valid path to enable editing.")
else:
    file = st.sidebar.file_uploader("Upload your .xlsx", type=["xlsx"])
    if file is not None:
        xlsx_bytes = file.read()
        # Do not overwrite last path on upload

# Year selector (2025–2050)
year = st.sidebar.selectbox("Summary year (for Jan–Dec table)", options=list(range(2025, 2051)), index=0)

# Load workbook
if 'xlsx_path' in locals() and xlsx_path is not None and xlsx_path.exists():
    wb = load_or_create_workbook(xlsx_path)
elif xlsx_bytes:
    tmp = io.BytesIO(xlsx_bytes)
    wb = load_workbook(tmp)
else:
    st.info("Load or upload your Excel file to continue.")
    st.stop()

# Ensure sheets
ws_staff = ensure_staff_sheet(wb)
ws_tx = ensure_transactions_sheet(wb)
ws_summary = ensure_summary_sheet(wb)

# Staff list
staff_list = get_staff_list(wb)

# --- Main form ---
col1, col2, col3 = st.columns([2,1,1])
with col1:
    staff = st.selectbox("Select employee", options=staff_list, index=0 if staff_list else None, placeholder="Choose staff")
with col2:
    action = st.selectbox("Action", options=["Add", "Minus", "Lost"])
with col3:
    use_range = st.checkbox("Use end date (range)", value=False)

date_row1, date_row2 = st.columns([2,2])
with date_row1:
    start_date = st.date_input("Start date", value=dt.date.today())
with date_row2:
    end_date = st.date_input("End date", value=dt.date.today()) if use_range else None

if use_range:
    if end_date is None or end_date < start_date:
        st.warning("End date must be the same as or after the start date.")
        computed_days = 0.0
    else:
        computed_days = float((end_date - start_date).days + 1)
    st.number_input("Days (computed; inclusive)", value=computed_days, step=0.5, disabled=True)
    days_value = computed_days
else:
    days_value = st.number_input("Days", min_value=0.0, step=0.5, value=1.0, format="%.2f")

note = st.text_input("Note (optional)", value="")

# Metrics
df_tx = transactions_to_df(ws_tx)
earned, taken, lost, bal = compute_balance(df_tx, staff) if staff else (0.0,0.0,0.0,0.0)

st.markdown("### Current balance for selected employee")
m1, m2, m3, m4 = st.columns(4)
m1.metric("Earned", f"{earned:.2f}")
m2.metric("Taken", f"{taken:.2f}")
m3.metric("Lost", f"{lost:.2f}")
m4.metric("Balance", f"{bal:.2f}")

user_name = st.text_input("Entered By (optional; auto-fill your name)")
go = st.button("➕ Add entry to Excel", type="primary")

if go:
    if not staff:
        st.error("Please select an employee."); st.stop()
    if days_value <= 0:
        st.error("Days must be greater than 0."); st.stop()

    entry_type = "Earned" if action == "Add" else ("Taken" if action == "Minus" else "Lost")
    start_dt = dt.datetime.combine(start_date, dt.time())
    end_dt = dt.datetime.combine(end_date, dt.time()) if end_date else start_dt

    write_tx_row(ws_tx, staff, start_dt, end_dt, entry_type, days_value, note, user_name)

    if 'xlsx_path' in locals() and xlsx_path is not None and xlsx_path.exists():
        wb.save(xlsx_path)
        st.success(f"Entry added → saved to {xlsx_path}")
        if remember:
            cfg["last_excel_path"] = str(xlsx_path)
            save_config(cfg)
    else:
        out = io.BytesIO()
        wb.save(out)
        st.download_button("Download updated .xlsx", data=out.getvalue(), file_name="lieu_days_updated.xlsx")
        st.success("Entry added. Click to download the updated file.")

    df_tx = transactions_to_df(ws_tx)
    earned, taken, lost, bal = compute_balance(df_tx, staff)
    st.toast(f"Balance for {staff}: {bal:.2f} (Earned {earned:.2f} / Taken {taken:.2f} / Lost {lost:.2f})")

# --- Summary writer ---
st.markdown("---")
st.subheader("📊 Excel Summary Table (Jan–Dec)")
st.caption("Monthly cells = net change (Earned − Taken − Lost) by the entry's Start date month. Lifetime Current balance in the last column.")

c1, c2 = st.columns([1,2])
with c1:
    if st.button("🔁 Rebuild Summary (write to Excel)"):
        rebuild_summary(ws_summary, staff_list, df_tx, year)
        if 'xlsx_path' in locals() and xlsx_path is not None and xlsx_path.exists():
            wb.save(xlsx_path)
            st.success(f"Summary updated for {year} → saved to {xlsx_path}")
            if remember:
                cfg["last_excel_path"] = str(xlsx_path)
                save_config(cfg)
        else:
            out = io.BytesIO()
            wb.save(out)
            st.download_button("Download updated .xlsx", data=out.getvalue(), file_name="lieu_days_updated.xlsx")
            st.success("Summary updated. Click to download the updated file.")
with c2:
    st.info("The app remembers the last valid Excel path (stored in ~/.lieu_days_manager/config.json). Use the checkbox to disable, or 'Clear saved path' to reset.")

# Recent transactions table
st.markdown("### Recent transactions for selected employee")
if df_tx.empty:
    st.info("No transactions yet.")
else:
    dsel = df_tx[df_tx["Staff"] == staff].copy() if staff else df_tx.copy()
    dsel = dsel.sort_values("Start Date", ascending=False).head(50)
    dsel["Start Date"] = dsel["Start Date"].dt.strftime("%Y-%m-%d").fillna("")
    dsel["End Date"]   = dsel["End Date"].dt.strftime("%Y-%m-%d").fillna("")
    st.dataframe(dsel, use_container_width=True)

st.divider()
st.caption("Tip: manage staff in the 'Staff' sheet (A2:A). Balance = Earned − Taken − Lost. Summary year selector is in the sidebar (2025–2050).")
