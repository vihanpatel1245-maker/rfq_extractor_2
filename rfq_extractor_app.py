import streamlit as st
import pdfplumber
import re
import io
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime, timedelta
import os
import json

# ── Authentication ─────────────────────────────────────────────────────────────
USERS = {"anany": "dada.niruma"}

def login_screen():
    st.title("🔒 Login")
    with st.form("login_form"):
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Login", type="primary", use_container_width=True)
    if submitted:
        if username in USERS and USERS[username] == password:
            st.session_state.authenticated = True
            st.session_state.username = username
            st.rerun()
        else:
            st.error("Invalid username or password")

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    login_screen()
    st.stop()

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(page_title="RFQ PDF → Excel", page_icon="📋", layout="centered")
st.title("PDF → Excel")

with st.sidebar:
    st.write(f"Logged in as **{st.session_state.get('username', '')}**")
    if st.button("Logout", use_container_width=True):
        st.session_state.authenticated = False
        st.session_state.pop("username", None)
        st.rerun()

st.markdown("""
<style>
div.stButton > button[kind="primary"],
div.stDownloadButton > button[kind="primary"] {
    background-color: #0E577C !important;
    border-color: #0E577C !important;
    color: white !important;
}
div.stButton > button[kind="primary"]:hover,
div.stDownloadButton > button[kind="primary"]:hover {
    background-color: #0a3f5a !important;
    border-color: #0a3f5a !important;
}
</style>
""", unsafe_allow_html=True)

# ── Columns ───────────────────────────────────────────────────────────────────
COLUMNS = [
    "Enquiry", "Cust", "RFQ D", "Due D", "Due Time",
    "Delivery Time", "Status", "SN", "CPN", "Description",
    "Enq. Part No*", "Enq. Mfg", "Qty", "Unit",
    "ITEM/VALUE.(EVALUATION)", "TP", "Remark", "ACT.Due DT"
]

EXCEL_DATE_BASE = datetime(1899, 12, 30)

def to_serial(d: datetime) -> int:
    return (d - EXCEL_DATE_BASE).days

# ── ITEM/VALUE.(EVALUATION) extractor ────────────────────────────────────────
def get_item_value(text: str, cust_code: str) -> str:
    code = cust_code.upper()
    if "N6" in code:
        m = re.search(
            r"(Year of manufacturing[^.]*?within\s+\d+\s+years?[^.]*\.)",
            text, re.IGNORECASE
        )
        if m:
            return m.group(1).strip()
    if "CH1" in code:
        m = re.search(
            r"(MANUFACTURING DATE CODE[^.\n]*?WITHIN\s+\d+\s+YEARS?[^.\n]*(?:IS MANDATORY|ORDER)[^.\n]*)",
            text, re.IGNORECASE
        )
        if m:
            return m.group(1).strip()
    return "N/A"

# ── Core extraction ───────────────────────────────────────────────────────────
def extract_from_pdf(pdf_bytes: bytes) -> list[dict]:
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        text = "\n".join(p.extract_text() or "" for p in pdf.pages)

    # 1. Enquiry
    enq_m = re.search(r"RFx number\s+(\d+)", text)
    enquiry = enq_m.group(1) if enq_m else "NA"

    # 2. Cust
    desc_m = re.search(r"Description:\s*(\S+)", text)
    cust = "BEL"
    if desc_m:
        parts = desc_m.group(1).split("/")
        cust = "BEL-" + parts[1] if len(parts) > 1 else "BEL-" + parts[0]

    # 2b. ITEM/VALUE
    item_value = get_item_value(text, cust)

    # 3. RFQ D
    rfq_d = datetime.today()

    # 4 & 5 & 18. Submission period
    sub_line = re.search(r"Submission period:\s*(.+)", text, re.IGNORECASE)
    if sub_line:
        all_dates = re.findall(r"(\d{2}\.\d{2}\.\d{4})\s+(\d{2}:\d{2})", sub_line.group(1))
        if all_dates:
            last_date, last_time = all_dates[-1]
            sub_date = datetime.strptime(last_date, "%d.%m.%Y")
            due_d    = sub_date - timedelta(days=3)
            t        = datetime.strptime(last_time, "%H:%M")
            due_time = t.strftime("%I:%M %p").lstrip("0")
            act_due  = sub_date
        else:
            due_d = due_time = act_due = None
    else:
        due_d = due_time = act_due = None

    # 6. Delivery Time
    lt_m = re.search(
        r"(?:OUR\s+REQUIRED\s+)?DELIVERY\s+SCHDULE?\s*:?\s*([\d\-–]+\s*(?:WEEKS?|DAYS?))",
        text, re.IGNORECASE
    )
    delivery_time = "NA"
    if lt_m:
        lt_raw  = lt_m.group(1).strip()
        range_m = re.match(r"(\d+)[\-–](\d+)\s*DAYS?", lt_raw, re.IGNORECASE)
        days_m  = re.match(r"(\d+)\s*DAYS?",            lt_raw, re.IGNORECASE)
        weeks_m = re.match(r"([\d\-–]+)\s*WEEKS?",      lt_raw, re.IGNORECASE)
        if range_m:
            avg = (int(range_m.group(1)) + int(range_m.group(2))) / 2
            delivery_time = f"{round(avg / 7)} WEEKS"
        elif days_m:
            delivery_time = f"{round(int(days_m.group(1)) / 7)} WEEKS"
        elif weeks_m:
            delivery_time = weeks_m.group(1) + " WEEKS"

    # 7+. Bid Details
    NOISE = re.compile(
        r"^(?:Item\s+Material|Qty/Unit|Quantity\s*$|Page\s+\d|Date\s*:|"
        r"Bid Invitation|Product no\.|info@|nklamin@|orantselectro@|"
        r"sales\.|support\.|\d{7,10}$)",
        re.IGNORECASE
    )
    in_bid    = False
    bid_lines = []
    for line in text.splitlines():
        s = line.strip()
        if not s:
            continue
        if re.search(r"^Bid Details$", s):
            in_bid = True
            continue
        if in_bid:
            if NOISE.match(s):
                continue
            bid_lines.append(s)

    rows = []
    i = 0
    while i < len(bid_lines):
        item_m = re.match(
            r"^(\d+)\s+(\d{8,})\s+(.+?)\s+([\d,\.]+)\s+([A-Z]+)$",
            bid_lines[i]
        )
        if item_m:
            sn      = int(item_m.group(1)) // 10
            cpn     = item_m.group(2)
            desc    = item_m.group(3).strip()
            qty_raw = item_m.group(4).replace(",", "")
            try:
                qty = float(qty_raw)
                qty = int(qty) if qty == int(qty) else qty
            except:
                qty = qty_raw
            unit = item_m.group(5)

            mfg_list, mpn_list = [], []
            while i + 1 < len(bid_lines):
                nxt = bid_lines[i + 1]
                if re.match(r"^\d+\s+\d{8,}", nxt):
                    break
                if "-" not in nxt and "/" not in nxt:
                    break
                nxt_n   = re.sub(r",\s*-", "-", nxt)
                split_m = re.match(r"^(.*?[A-Z\)\/\.])\s*-\s*(.+)$", nxt_n)
                if split_m:
                    mfg_list.append(split_m.group(1).strip().rstrip(",/").strip())
                    mpn_list.append(split_m.group(2).strip())
                else:
                    break
                i += 1

            mfg = " // ".join(mfg_list) if mfg_list else "NA"
            mpn = " // ".join(mpn_list) if mpn_list else "NA"

            rows.append({
                "Enquiry":                 enquiry,
                "Cust":                    cust,
                "RFQ D":                   rfq_d,
                "Due D":                   due_d,
                "Due Time":                due_time or "NA",
                "Delivery Time":           delivery_time,
                "Status":                  "Working",
                "SN":                      sn,
                "CPN":                     cpn,
                "Description":             desc,
                "Enq. Part No*":           mpn,
                "Enq. Mfg":                mfg,
                "Qty":                     qty,
                "Unit":                    unit,
                "ITEM/VALUE.(EVALUATION)": item_value,
                "TP":                      "",
                "Remark":                  "N/A",
                "ACT.Due DT":              act_due,
            })
        i += 1

    if not rows:
        rows.append({col: "NA" for col in COLUMNS} | {
            "Enquiry": enquiry, "Cust": cust,
            "RFQ D": rfq_d, "Due D": due_d,
            "Due Time": due_time or "NA",
            "Delivery Time": delivery_time,
            "Status": "Working",
            "ITEM/VALUE.(EVALUATION)": item_value,
            "Remark": "N/A", "ACT.Due DT": act_due, "TP": "",
        })

    return rows


# ── Excel builder ─────────────────────────────────────────────────────────────
def build_excel(all_rows: list[dict]) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "RFQ Log"

    hdr_fill  = PatternFill("solid", start_color="1F4E79", end_color="1F4E79")
    hdr_font  = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    hdr_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin      = Side(style="thin", color="BBBBBB")
    bdr       = Border(left=thin, right=thin, top=thin, bottom=thin)
    fill_w    = PatternFill("solid", start_color="FFFFFF", end_color="FFFFFF")
    fill_b    = PatternFill("solid", start_color="EBF3FB", end_color="EBF3FB")
    data_font = Font(name="Arial", size=10)
    d_align   = Alignment(horizontal="left", vertical="center")

    col_widths = [14, 10, 12, 12, 10, 14, 10, 5, 16, 40, 16, 28, 6, 6, 22, 8, 10, 12]
    date_cols  = {"RFQ D", "Due D", "ACT.Due DT"}

    for ci, (col, w) in enumerate(zip(COLUMNS, col_widths), 1):
        cell = ws.cell(row=1, column=ci, value=col)
        cell.font = hdr_font; cell.fill = hdr_fill
        cell.alignment = hdr_align; cell.border = bdr
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.row_dimensions[1].height = 28

    for ri, row in enumerate(all_rows, 2):
        fill = fill_w if ri % 2 == 0 else fill_b
        for ci, col in enumerate(COLUMNS, 1):
            val  = row.get(col, "")
            cell = ws.cell(row=ri, column=ci)
            if col in date_cols:
                if isinstance(val, datetime):
                    cell.value = to_serial(val)
                    cell.number_format = "DD-MMM-YY"
                else:
                    cell.value = str(val) if val else ""
            elif col == "SN":
                try:    cell.value = int(val)
                except: cell.value = val
            elif col == "Qty":
                try:
                    f = float(val)
                    cell.value = int(f) if f == int(f) else f
                except: cell.value = val
            else:
                cell.value = str(val) if val not in (None, "") else ""
            cell.font = data_font; cell.fill = fill
            cell.alignment = d_align; cell.border = bdr
        ws.row_dimensions[ri].height = 18

    ws.freeze_panes = "A2"
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ── Session state ─────────────────────────────────────────────────────────────
if "result_excel" not in st.session_state:
    st.session_state.result_excel = None
if "result_rows" not in st.session_state:
    st.session_state.result_rows  = []
if "result_count" not in st.session_state:
    st.session_state.result_count = 0

# ── UI ────────────────────────────────────────────────────────────────────────
uploaded_files = st.file_uploader(
    "Upload PDF",
    type="pdf",
    accept_multiple_files=True,
)

if uploaded_files:
    st.write(f"**{len(uploaded_files)} file(s) selected**")

    if st.button("Generate Excel", type="primary", use_container_width=True):
        st.session_state.result_excel = None
        st.session_state.result_rows  = []
        st.session_state.result_count = 0

        all_rows = []
        prog = st.progress(0)

        for i, uf in enumerate(uploaded_files):
            with st.spinner(f"Reading {uf.name}…"):
                try:
                    rows = extract_from_pdf(uf.read())
                    all_rows.extend(rows)
                except Exception as e:
                    st.error(f"❌  {uf.name}: {e}")
            prog.progress((i + 1) / len(uploaded_files))

        if all_rows:
            st.session_state.result_excel = build_excel(all_rows)
            st.session_state.result_rows  = all_rows
            st.session_state.result_count = len(uploaded_files)

# ── Download + preview ────────────────────────────────────────────────────────
if st.session_state.result_excel:
    st.markdown("---")
    st.download_button(
        label="Download RFQ_Log.xlsx",
        data=st.session_state.result_excel,
        file_name="RFQ_Log.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
        key="dl_combined",
    )

    st.subheader("Preview")
    preview_cols = ["Enquiry", "Cust", "Due D", "Due Time", "Delivery Time",
                    "SN", "CPN", "Description", "Enq. Part No*", "Enq. Mfg", "Qty", "Unit"]
    df = pd.DataFrame(st.session_state.result_rows)
    for dc in ["RFQ D", "Due D", "ACT.Due DT"]:
        if dc in df.columns:
            df[dc] = df[dc].apply(
                lambda x: x.strftime("%d-%b-%y") if isinstance(x, datetime) else str(x)
            )
    st.dataframe(
        df[[c for c in preview_cols if c in df.columns]],
        use_container_width=True
    )
    st.caption(
        f"Total: {len(st.session_state.result_rows)} row(s) "
        f"from {st.session_state.result_count} PDF(s)"
    )
