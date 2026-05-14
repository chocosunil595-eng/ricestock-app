# =============================================================================
# RiceStock - Quintal System
# Government FPS Rice Inventory Management System
# Developed for: Dealer - Balaram Shial | Code: 0201P100
# Author: Sunil Kumar
# =============================================================================

import streamlit as st
import pandas as pd
import sqlite3
import hashlib
import pytz
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime, date, timedelta
from io import BytesIO
import calendar
import openpyxl
from openpyxl.styles import (
    Font, Alignment, Border, Side, PatternFill, numbers
)
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="RiceStock - Quintal System",
    page_icon="🌾",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────────────────────────────────────
IST = pytz.timezone("Asia/Kolkata")
DB_PATH = "rice_inventory.db"
CARD_TYPES = ["PHH", "AAY", "SFSS"]
DEALER_NAME = "Dealer - Balaram Shial"
DEALER_CODE = "Code - 0201P100"
APP_NAME = "RiceStock"

# ─────────────────────────────────────────────────────────────────────────────
# GLOBAL CSS – Clean Green Agricultural Government Theme
# ─────────────────────────────────────────────────────────────────────────────
def inject_css():
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans:wght@400;500;600;700&family=Noto+Serif:wght@600;700&display=swap');

    html, body, [class*="css"] {
        font-family: 'Noto Sans', sans-serif;
    }

    /* ── Main background ── */
    .stApp {
        background: linear-gradient(160deg, #f0f7f0 0%, #e8f5e8 50%, #f5f9f0 100%);
    }

    /* ── Sidebar ── */
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #1a4731 0%, #2d6a4f 60%, #1b4332 100%);
        border-right: 3px solid #40916c;
    }
    section[data-testid="stSidebar"] * {
        color: #d8f3dc !important;
    }
    section[data-testid="stSidebar"] .stRadio label {
        color: #d8f3dc !important;
        font-size: 0.95rem;
    }
    section[data-testid="stSidebar"] hr {
        border-color: #40916c !important;
    }

    /* ── Cards / Metric boxes ── */
    .metric-card {
        background: white;
        border: 1px solid #b7e4c7;
        border-left: 5px solid #2d6a4f;
        border-radius: 10px;
        padding: 18px 22px;
        margin: 8px 0;
        box-shadow: 0 2px 8px rgba(45,106,79,0.08);
    }
    .metric-card h3 { margin: 0; font-size: 0.85rem; color: #555; text-transform: uppercase; letter-spacing: 0.05em; }
    .metric-card h2 { margin: 4px 0 0 0; font-size: 1.6rem; color: #1a4731; font-family: 'Noto Serif', serif; }
    .metric-card p  { margin: 2px 0 0 0; font-size: 0.8rem; color: #888; }

    /* ── Section headers ── */
    .section-header {
        background: linear-gradient(90deg, #2d6a4f, #40916c);
        color: white !important;
        padding: 10px 18px;
        border-radius: 8px;
        margin: 14px 0 10px 0;
        font-weight: 600;
        font-size: 1.05rem;
        letter-spacing: 0.03em;
    }

    /* ── Login card ── */
    .login-card {
        background: white;
        border-radius: 16px;
        padding: 40px;
        box-shadow: 0 8px 32px rgba(45,106,79,0.15);
        border-top: 6px solid #2d6a4f;
        max-width: 420px;
        margin: auto;
    }
    .login-title {
        font-family: 'Noto Serif', serif;
        font-size: 1.6rem;
        color: #1a4731;
        text-align: center;
        margin-bottom: 4px;
    }
    .login-sub {
        text-align: center;
        color: #40916c;
        font-size: 0.9rem;
        margin-bottom: 24px;
    }

    /* ── Greeting banner ── */
    .greeting-banner {
        background: linear-gradient(90deg, #1a4731, #2d6a4f, #40916c);
        color: white;
        padding: 12px 22px;
        border-radius: 10px;
        font-size: 1.1rem;
        font-weight: 600;
        margin-bottom: 18px;
        display: flex;
        justify-content: space-between;
        align-items: center;
    }

    /* ── Quintal badge ── */
    .q-badge {
        background: #d8f3dc;
        color: #1a4731;
        border: 1px solid #95d5b2;
        border-radius: 6px;
        padding: 2px 8px;
        font-weight: 600;
        font-size: 0.9rem;
        display: inline-block;
    }

    /* ── Info / warning boxes ── */
    .info-box {
        background: #e8f5e8;
        border: 1px solid #95d5b2;
        border-left: 4px solid #2d6a4f;
        border-radius: 8px;
        padding: 12px 16px;
        font-size: 0.9rem;
        color: #1a4731;
        margin: 8px 0;
    }
    .warn-box {
        background: #fff8e1;
        border: 1px solid #ffe082;
        border-left: 4px solid #f9a825;
        border-radius: 8px;
        padding: 12px 16px;
        font-size: 0.9rem;
        color: #5d4037;
        margin: 8px 0;
    }

    /* ── Footer ── */
    .footer {
        text-align: center;
        padding: 14px 0 4px 0;
        font-size: 0.8rem;
        color: #888;
        border-top: 1px solid #b7e4c7;
        margin-top: 40px;
    }
    .footer strong { color: #2d6a4f; }

    /* ── Streamlit overrides ── */
    .stButton > button {
        background: linear-gradient(90deg, #2d6a4f, #40916c);
        color: white;
        border: none;
        border-radius: 8px;
        font-weight: 600;
        padding: 8px 22px;
        transition: all 0.2s;
    }
    .stButton > button:hover {
        background: linear-gradient(90deg, #1a4731, #2d6a4f);
        box-shadow: 0 4px 12px rgba(45,106,79,0.3);
    }
    div[data-testid="stDataFrame"] { border-radius: 10px; overflow: hidden; }
    .stSelectbox > div > div { border-color: #95d5b2 !important; border-radius: 8px !important; }
    .stTextInput > div > div { border-radius: 8px !important; }
    .stDateInput > div > div { border-radius: 8px !important; }
    .stNumberInput > div > div { border-radius: 8px !important; }
    .stTabs [data-baseweb="tab"] { font-weight: 600; }
    .stTabs [data-baseweb="tab-list"] { border-bottom: 2px solid #b7e4c7; }
    </style>
    """, unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# DATABASE SETUP
# ─────────────────────────────────────────────────────────────────────────────
def get_connection():
    """Return a SQLite connection to the inventory DB."""
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    """Create tables if they don't exist."""
    conn = get_connection()
    cur = conn.cursor()

    # Inventory table (all quantities stored as float Kg)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS inventory (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            entry_date  TEXT NOT NULL,
            card_type   TEXT NOT NULL,
            opening_kg  REAL DEFAULT 0,
            received_kg REAL DEFAULT 0,
            total_kg    REAL DEFAULT 0,
            closing_kg  REAL DEFAULT 0,
            selling_kg  REAL DEFAULT 0,
            remarks     TEXT DEFAULT '',
            pickup_day  INTEGER DEFAULT NULL,
            created_at  TEXT,
            UNIQUE(entry_date, card_type)
        )
    """)
    # Migrate existing DB safely (adds column if not already present)
    try:
        cur.execute("ALTER TABLE inventory ADD COLUMN pickup_day INTEGER DEFAULT NULL")
        conn.commit()
    except Exception:
        pass  # Column already exists — safe to ignore

    # Settings / credentials table
    cur.execute("""
        CREATE TABLE IF NOT EXISTS settings (
            key   TEXT PRIMARY KEY,
            value TEXT NOT NULL
        )
    """)

    # Seed default credentials (stored as SHA-256 hashes)
    default_creds = [
        ("username", "Sunil"),
        ("password", hashlib.sha256("sunil123".encode()).hexdigest()),
    ]
    for k, v in default_creds:
        cur.execute(
            "INSERT OR IGNORE INTO settings (key, value) VALUES (?, ?)",
            (k, v)
        )

    conn.commit()
    conn.close()


# ─────────────────────────────────────────────────────────────────────────────
# UTILITY HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def now_ist() -> datetime:
    """Current datetime in IST."""
    return datetime.now(IST)


def today_ist() -> date:
    """Today's date in IST."""
    return now_ist().date()


def get_greeting(username: str) -> str:
    """Return time-appropriate greeting for the user."""
    hour = now_ist().hour
    if 0 <= hour < 12:
        period = "Good Morning"
    elif 12 <= hour < 17:
        period = "Good Afternoon"
    else:
        period = "Good Evening"
    return f"{period} Mr. {username} 🙏"


def format_qty(kg: float) -> str:
    """
    Smart display: if kg >= 100, show as 'X Q.YY Kg', else 'YY Kg'.
    Examples: 125.5 → '1 Q.25.5 Kg', 75.0 → '75 Kg', 0 → '0 Kg'
    """
    if kg is None:
        kg = 0.0
    kg = round(float(kg), 2)
    if kg >= 100:
        quintals = int(kg // 100)
        remainder = kg % 100
        if remainder == int(remainder):
            remainder = int(remainder)
        return f"{quintals} Q.{remainder} Kg"
    else:
        display = int(kg) if kg == int(kg) else kg
        return f"{display} Kg"


def hash_password(pw: str) -> str:
    return hashlib.sha256(pw.encode()).hexdigest()


def verify_login(username: str, password: str) -> bool:
    """Check credentials against DB."""
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("SELECT value FROM settings WHERE key='username'")
    db_user = cur.fetchone()
    cur.execute("SELECT value FROM settings WHERE key='password'")
    db_pass = cur.fetchone()
    conn.close()
    if db_user and db_pass:
        return (username.strip() == db_user["value"] and
                hash_password(password) == db_pass["value"])
    return False


def get_setting(key: str) -> str:
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("SELECT value FROM settings WHERE key=?", (key,))
    row = cur.fetchone()
    conn.close()
    return row["value"] if row else ""


def update_setting(key: str, value: str):
    conn = get_connection()
    conn.execute(
        "INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)",
        (key, value)
    )
    conn.commit()
    conn.close()


# ─────────────────────────────────────────────────────────────────────────────
# INVENTORY DATA ACCESS
# ─────────────────────────────────────────────────────────────────────────────
def get_previous_closing(entry_date: date, card_type: str) -> float:
    """
    Returns the closing_kg of the most recent entry BEFORE entry_date
    for the given card_type. Returns 0.0 if no prior entry exists.
    """
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("""
        SELECT closing_kg FROM inventory
        WHERE card_type = ? AND entry_date < ?
        ORDER BY entry_date DESC
        LIMIT 1
    """, (card_type, str(entry_date)))
    row = cur.fetchone()
    conn.close()
    return float(row["closing_kg"]) if row else 0.0


def upsert_entry(entry_date: date, card_type: str,
                 opening_kg: float, received_kg: float,
                 closing_kg: float, remarks: str = "",
                 pickup_day: int = None) -> dict:
    """
    Insert or update a daily entry, applying smart auto-adjustment.
    Returns a dict with final values and an optional adjustment message.
    """
    adj_message = None

    total_kg = opening_kg + received_kg

    # Smart auto-adjustment: if closing > total, reduce received
    if closing_kg > total_kg:
        excess = closing_kg - total_kg
        received_kg = max(0.0, received_kg - excess)
        total_kg = opening_kg + received_kg
        adj_message = (
            "⚠️ Closing was higher than Total. "
            "Received has been automatically adjusted."
        )

    selling_kg = max(0.0, total_kg - closing_kg)

    conn = get_connection()
    conn.execute("""
        INSERT INTO inventory
            (entry_date, card_type, opening_kg, received_kg,
             total_kg, closing_kg, selling_kg, remarks, pickup_day, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(entry_date, card_type) DO UPDATE SET
            opening_kg  = excluded.opening_kg,
            received_kg = excluded.received_kg,
            total_kg    = excluded.total_kg,
            closing_kg  = excluded.closing_kg,
            selling_kg  = excluded.selling_kg,
            remarks     = excluded.remarks,
            pickup_day  = excluded.pickup_day,
            created_at  = excluded.created_at
    """, (
        str(entry_date), card_type,
        round(opening_kg, 2), round(received_kg, 2),
        round(total_kg, 2), round(closing_kg, 2),
        round(selling_kg, 2), remarks, pickup_day,
        now_ist().strftime("%Y-%m-%d %H:%M:%S")
    ))
    conn.commit()
    conn.close()

    return {
        "opening_kg": opening_kg,
        "received_kg": received_kg,
        "total_kg": total_kg,
        "closing_kg": closing_kg,
        "selling_kg": selling_kg,
        "adj_message": adj_message,
    }


def get_entry(entry_date: date, card_type: str):
    """Fetch a single entry or None."""
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("""
        SELECT * FROM inventory
        WHERE entry_date = ? AND card_type = ?
    """, (str(entry_date), card_type))
    row = cur.fetchone()
    conn.close()
    return dict(row) if row else None


def get_entries_for_month(year: int, month: int,
                          card_type: str = None) -> pd.DataFrame:
    """Return all entries for a month, optionally filtered by card_type."""
    conn = get_connection()
    start = f"{year:04d}-{month:02d}-01"
    end_day = calendar.monthrange(year, month)[1]
    end = f"{year:04d}-{month:02d}-{end_day:02d}"
    if card_type:
        df = pd.read_sql_query("""
            SELECT * FROM inventory
            WHERE entry_date BETWEEN ? AND ? AND card_type = ?
            ORDER BY entry_date
        """, conn, params=(start, end, card_type))
    else:
        df = pd.read_sql_query("""
            SELECT * FROM inventory
            WHERE entry_date BETWEEN ? AND ?
            ORDER BY entry_date, card_type
        """, conn, params=(start, end))
    conn.close()
    return df


def get_entries_for_year(year: int) -> pd.DataFrame:
    conn = get_connection()
    df = pd.read_sql_query("""
        SELECT * FROM inventory
        WHERE entry_date LIKE ?
        ORDER BY entry_date, card_type
    """, conn, params=(f"{year:04d}%",))
    conn.close()
    return df


def get_latest_stock() -> dict:
    """
    Returns current (latest) closing_kg for each card type.
    """
    conn = get_connection()
    cur = conn.cursor()
    result = {}
    for ct in CARD_TYPES:
        cur.execute("""
            SELECT closing_kg FROM inventory
            WHERE card_type = ?
            ORDER BY entry_date DESC LIMIT 1
        """, (ct,))
        row = cur.fetchone()
        result[ct] = float(row["closing_kg"]) if row else 0.0
    conn.close()
    return result


def get_30day_trend() -> pd.DataFrame:
    """Last 30 days of data for trend chart."""
    end = today_ist()
    start = end - timedelta(days=29)
    conn = get_connection()
    df = pd.read_sql_query("""
        SELECT entry_date, card_type,
               closing_kg, selling_kg, received_kg
        FROM inventory
        WHERE entry_date BETWEEN ? AND ?
        ORDER BY entry_date
    """, conn, params=(str(start), str(end)))
    conn.close()
    return df


# ─────────────────────────────────────────────────────────────────────────────
# EXCEL EXPORT
# ─────────────────────────────────────────────────────────────────────────────
def build_excel_report(df: pd.DataFrame,
                        report_title: str,
                        period_label: str,
                        card_type: str) -> BytesIO:
    """
    Create a government-style FPS Stock Register Excel workbook.
    Columns: Date | Opening (Q) | Receipt (Q) | Total (Q) |
             No. of Cards | Qty Issued (Q) | Closing (Q) | Remarks
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "FPS Register"

    # ── Styles ──
    green_fill   = PatternFill("solid", fgColor="1A4731")
    header_fill  = PatternFill("solid", fgColor="2D6A4F")
    alt_fill     = PatternFill("solid", fgColor="E8F5E8")
    total_fill   = PatternFill("solid", fgColor="B7E4C7")
    white_fill   = PatternFill("solid", fgColor="FFFFFF")

    thin  = Side(style="thin",   color="2D6A4F")
    thick = Side(style="medium", color="1A4731")
    thin_border  = Border(left=thin,  right=thin,  top=thin,  bottom=thin)
    thick_border = Border(left=thick, right=thick, top=thick, bottom=thick)

    bold_white  = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
    bold_dark   = Font(name="Calibri", bold=True, color="1A4731", size=11)
    normal_font = Font(name="Calibri", size=10)
    total_font  = Font(name="Calibri", bold=True, color="1A4731", size=10)

    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left   = Alignment(horizontal="left",   vertical="center", wrap_text=True)

    # ── Row 1: Main title (merged) ──
    ws.merge_cells("A1:H1")
    c = ws["A1"]
    c.value = "FPS Stock Register / F.P.S. Code No."
    c.font = Font(name="Calibri", bold=True, color="FFFFFF", size=14)
    c.fill = green_fill
    c.alignment = center
    c.border = thick_border
    ws.row_dimensions[1].height = 28

    # ── Row 2: Dealer info ──
    ws.merge_cells("A2:D2")
    c = ws["A2"]
    c.value = f"{DEALER_NAME}    |    {DEALER_CODE}"
    c.font = Font(name="Calibri", bold=True, color="1A4731", size=10)
    c.fill = PatternFill("solid", fgColor="D8F3DC")
    c.alignment = left
    c.border = thin_border

    ws.merge_cells("E2:H2")
    c = ws["E2"]
    c.value = f"Card Type: {card_type}    |    {period_label}"
    c.font = Font(name="Calibri", bold=True, color="1A4731", size=10)
    c.fill = PatternFill("solid", fgColor="D8F3DC")
    c.alignment = center
    c.border = thin_border
    ws.row_dimensions[2].height = 20

    # ── Row 3: Report label ──
    ws.merge_cells("A3:H3")
    c = ws["A3"]
    c.value = report_title
    c.font = Font(name="Calibri", bold=True, color="FFFFFF", size=12)
    c.fill = header_fill
    c.alignment = center
    c.border = thick_border
    ws.row_dimensions[3].height = 22

    # ── Row 4: Column headers ──
    headers = [
        "Date of\nTransaction",
        "Opening\nBalance (Q)",
        "Receipt\n(Q)",
        "Total\n(Q)",
        "Number\nof Cards",
        "Quantity\nIssued (Q)",
        "Closing\nBalance (Q)",
        "Remarks",
    ]
    for col_idx, hdr in enumerate(headers, start=1):
        c = ws.cell(row=4, column=col_idx, value=hdr)
        c.font = bold_white
        c.fill = header_fill
        c.alignment = center
        c.border = thin_border
    ws.row_dimensions[4].height = 36

    # ── Column widths ──
    col_widths = [16, 16, 14, 14, 14, 16, 16, 22]
    for i, w in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # ── Data rows ──
    row_num = 5
    totals = {
        "opening": 0.0, "received": 0.0, "total": 0.0,
        "selling": 0.0, "closing": 0.0
    }

    for i, (_, row) in enumerate(df.iterrows()):
        fill = alt_fill if i % 2 == 0 else white_fill
        entry_date_str = str(row.get("entry_date", ""))
        try:
            dt = datetime.strptime(entry_date_str, "%Y-%m-%d")
            date_display = dt.strftime("%d-%b-%Y")
        except Exception:
            date_display = entry_date_str

        row_data = [
            date_display,
            format_qty(row.get("opening_kg", 0)),
            format_qty(row.get("received_kg", 0)),
            format_qty(row.get("total_kg", 0)),
            "",   # Number of cards – left blank (not tracked)
            format_qty(row.get("selling_kg", 0)),
            format_qty(row.get("closing_kg", 0)),
            row.get("remarks", ""),
        ]
        for col_idx, val in enumerate(row_data, start=1):
            c = ws.cell(row=row_num, column=col_idx, value=val)
            c.font = normal_font
            c.fill = fill
            c.alignment = center if col_idx != 8 else left
            c.border = thin_border
        ws.row_dimensions[row_num].height = 18

        totals["opening"]   += float(row.get("opening_kg", 0))
        totals["received"]  += float(row.get("received_kg", 0))
        totals["total"]     += float(row.get("total_kg", 0))
        totals["selling"]   += float(row.get("selling_kg", 0))
        totals["closing"]   += float(row.get("closing_kg", 0))
        row_num += 1

    # ── Totals row ──
    total_row = [
        "TOTAL",
        format_qty(totals["opening"]),
        format_qty(totals["received"]),
        format_qty(totals["total"]),
        "",
        format_qty(totals["selling"]),
        format_qty(totals["closing"]),
        "",
    ]
    for col_idx, val in enumerate(total_row, start=1):
        c = ws.cell(row=row_num, column=col_idx, value=val)
        c.font = total_font
        c.fill = total_fill
        c.alignment = center
        c.border = thin_border
    ws.row_dimensions[row_num].height = 20

    # ── Footer rows ──
    row_num += 2
    ws.merge_cells(f"A{row_num}:H{row_num}")
    c = ws.cell(row=row_num, column=1,
                value="Rice Inventory  |  Created by - Sunil Kumar  |  RiceStock - Quintal System")
    c.font = Font(name="Calibri", italic=True, color="888888", size=9)
    c.alignment = center

    # ── Freeze panes ──
    ws.freeze_panes = "A5"

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ─────────────────────────────────────────────────────────────────────────────
# SHARED UI COMPONENTS
# ─────────────────────────────────────────────────────────────────────────────
def render_greeting():
    username = get_setting("username")
    greeting = get_greeting(username)
    ist_now  = now_ist().strftime("%d %b %Y  |  %I:%M %p IST")
    st.markdown(f"""
        <div class="greeting-banner">
            <span>🌾 {greeting}</span>
            <span style="font-size:0.85rem; opacity:0.85;">📅 {ist_now}</span>
        </div>
    """, unsafe_allow_html=True)


def render_footer():
    st.markdown("""
        <div class="footer">
            <strong>Rice Inventory</strong><br>
            Created by - Sunil Kumar
        </div>
    """, unsafe_allow_html=True)


def section_header(text: str):
    st.markdown(f'<div class="section-header">▎ {text}</div>',
                unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# PAGE: LOGIN
# ─────────────────────────────────────────────────────────────────────────────
def page_login():
    st.markdown("""
        <div style="text-align:center; padding: 30px 0 10px 0;">
            <div style="font-size:3rem;">🌾</div>
            <h1 style="font-family:'Noto Serif',serif; color:#1a4731;
                       font-size:2rem; margin:0;">RiceStock</h1>
            <p style="color:#40916c; font-size:1rem; margin:0;">
                Quintal Inventory Management System
            </p>
        </div>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 1.4, 1])
    with col2:
        st.markdown(f"""
            <div class="login-card">
                <div class="login-title">{DEALER_NAME}</div>
                <div class="login-sub">{DEALER_CODE}</div>
            </div>
        """, unsafe_allow_html=True)

        st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)

        with st.form("login_form"):
            username = st.text_input("👤 Username", placeholder="Enter username")
            password = st.text_input("🔑 Password", type="password",
                                     placeholder="Enter password")
            submitted = st.form_submit_button("🔓 Login", use_container_width=True)

        if submitted:
            if verify_login(username, password):
                st.session_state["logged_in"] = True
                st.session_state["username"]   = username
                st.success("✅ Login successful! Redirecting…")
                st.rerun()
            else:
                st.error("❌ Invalid username or password. Please try again.")

    render_footer()


# ─────────────────────────────────────────────────────────────────────────────
# PAGE: DASHBOARD
# ─────────────────────────────────────────────────────────────────────────────
def page_dashboard():
    render_greeting()
    section_header("📊 Current Stock Overview")

    stock = get_latest_stock()
    card_colors = {"PHH": "#2d6a4f", "AAY": "#40916c", "SFSS": "#52b788"}
    max_stock = 1000  # progress bar denominator (Kg)

    cols = st.columns(3)
    card_labels = {
        "PHH":  ("Priority Household", "🏠"),
        "AAY":  ("Antyodaya Anna Yojana", "👨‍👩‍👧"),
        "SFSS": ("State Food Security Scheme", "🏛️"),
    }
    for i, ct in enumerate(CARD_TYPES):
        with cols[i]:
            kg  = stock.get(ct, 0.0)
            lbl, icon = card_labels[ct]
            pct = min(int((kg / max_stock) * 100), 100)
            st.markdown(f"""
                <div class="metric-card">
                    <h3>{icon} {ct} – {lbl}</h3>
                    <h2>{format_qty(kg)}</h2>
                    <p>Current Closing Balance</p>
                </div>
            """, unsafe_allow_html=True)
            st.progress(pct / 100, text=f"{pct}% of {format_qty(max_stock)}")

    # ── Summary totals ──
    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
    section_header("📋 All-Time Summary")

    conn = get_connection()
    summary = pd.read_sql_query("""
        SELECT
            SUM(received_kg) AS total_received,
            SUM(selling_kg)  AS total_sold,
            SUM(closing_kg)  AS total_closing
        FROM inventory
    """, conn)
    conn.close()

    s_cols = st.columns(3)
    summary_items = [
        ("📥 Total Received", summary["total_received"][0], "All-time inward stock"),
        ("📤 Total Sold",     summary["total_sold"][0],     "All-time issued stock"),
        ("📦 Current Stock",  sum(stock.values()),          "Combined closing balance"),
    ]
    for i, (lbl, val, sub) in enumerate(summary_items):
        val = val or 0.0
        with s_cols[i]:
            st.markdown(f"""
                <div class="metric-card">
                    <h3>{lbl}</h3>
                    <h2>{format_qty(val)}</h2>
                    <p>{sub}</p>
                </div>
            """, unsafe_allow_html=True)

    # ── 30-day trend chart ──
    section_header("📈 30-Day Stock Trend")
    trend_df = get_30day_trend()
    if trend_df.empty:
        st.markdown("""
            <div class="info-box">
                📭 No data available for the last 30 days.
                Start adding daily entries to see the trend chart here.
            </div>
        """, unsafe_allow_html=True)
    else:
        fig = go.Figure()
        for ct, color in card_colors.items():
            ct_df = trend_df[trend_df["card_type"] == ct]
            if not ct_df.empty:
                fig.add_trace(go.Scatter(
                    x=ct_df["entry_date"],
                    y=ct_df["closing_kg"],
                    name=f"{ct} – Closing",
                    mode="lines+markers",
                    line=dict(color=color, width=2.5),
                    marker=dict(size=6),
                    hovertemplate="%{x}<br>Closing: %{y:.1f} Kg<extra>" + ct + "</extra>",
                ))
        fig.update_layout(
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(240,247,240,0.6)",
            font=dict(family="Noto Sans", size=12, color="#1a4731"),
            legend=dict(orientation="h", y=-0.15),
            xaxis=dict(title="Date", gridcolor="#b7e4c7"),
            yaxis=dict(title="Stock (Kg)", gridcolor="#b7e4c7"),
            margin=dict(l=20, r=20, t=20, b=20),
            height=320,
        )
        st.plotly_chart(fig, use_container_width=True)

    render_footer()


# ─────────────────────────────────────────────────────────────────────────────
# PAGE: DAILY ENTRY
# ─────────────────────────────────────────────────────────────────────────────
def page_daily_entry():
    render_greeting()

    tab1, tab2 = st.tabs(["📝 Daily Entry", "📦 Bulk 2-Month Entry"])

    # ════════════════════════════════════════════════════════════
    # TAB 1 – Single Daily Entry
    # ════════════════════════════════════════════════════════════
    with tab1:
        section_header("📝 Add / Edit Daily Entry")

        col_a, col_b = st.columns([1, 1])
        with col_a:
            entry_date = st.date_input(
                "📅 Date",
                value=today_ist(),
                help="Select the entry date (IST)"
            )
        with col_b:
            card_type = st.selectbox(
                "🃏 Card Type",
                CARD_TYPES,
                help="Select ration card type"
            )

        # Auto-fill opening balance
        opening_kg = get_previous_closing(entry_date, card_type)
        existing   = get_entry(entry_date, card_type)

        st.markdown(f"""
            <div class="info-box">
                🔓 <strong>Opening Balance (Auto-filled):</strong>
                &nbsp;<span class="q-badge">{format_qty(opening_kg)}</span>
                &nbsp;— from previous {card_type} closing balance
                {f'&nbsp;|&nbsp; <strong>Entry exists for this date</strong>' if existing else ''}
            </div>
        """, unsafe_allow_html=True)

        col_c, col_d = st.columns(2)
        with col_c:
            default_recv = float(existing["received_kg"]) if existing else 0.0
            received_kg = st.number_input(
                "📥 Received (Kg)",
                min_value=0.0, step=0.5,
                value=default_recv,
                format="%.2f",
                help="Quantity received in Kg"
            )
        with col_d:
            total_kg = opening_kg + received_kg
            st.markdown(f"""
                <div class="metric-card" style="margin-top:28px;">
                    <h3>📊 Total (Auto-calculated)</h3>
                    <h2>{format_qty(total_kg)}</h2>
                </div>
            """, unsafe_allow_html=True)

        col_e, col_f = st.columns(2)
        with col_e:
            default_clos = float(existing["closing_kg"]) if existing else 0.0
            closing_kg = st.number_input(
                "📦 Closing Balance (Kg)",
                min_value=0.0, step=0.5,
                value=default_clos,
                format="%.2f",
                help="Enter actual closing balance in Kg"
            )
        with col_f:
            selling_preview = max(0.0, total_kg - closing_kg)
            st.markdown(f"""
                <div class="metric-card" style="margin-top:28px;">
                    <h3>📤 Selling (Auto-calculated)</h3>
                    <h2>{format_qty(selling_preview)}</h2>
                    <p>= Total − Closing</p>
                </div>
            """, unsafe_allow_html=True)

        remarks = st.text_input(
            "📝 Remarks (optional)",
            value=existing["remarks"] if existing else "",
            placeholder="E.g. Monthly distribution, Festival allocation…"
        )

        # Warn if closing > total
        if closing_kg > total_kg:
            st.markdown(f"""
                <div class="warn-box">
                    ⚠️ <strong>Closing ({format_qty(closing_kg)}) &gt;
                    Total ({format_qty(total_kg)}).</strong><br>
                    Received will be auto-adjusted on save.
                </div>
            """, unsafe_allow_html=True)

        if st.button("💾 Save Entry", use_container_width=True):
            result = upsert_entry(
                entry_date, card_type,
                opening_kg, received_kg, closing_kg, remarks
            )
            if result["adj_message"]:
                st.warning(result["adj_message"])
            st.success(
                f"✅ Entry saved!  "
                f"Opening: {format_qty(result['opening_kg'])} | "
                f"Received: {format_qty(result['received_kg'])} | "
                f"Total: {format_qty(result['total_kg'])} | "
                f"Closing: {format_qty(result['closing_kg'])} | "
                f"Selling: {format_qty(result['selling_kg'])}"
            )
            st.rerun()

    # ====================== BULK 2-MONTH ENTRY ======================
st.subheader("🧾 Bulk 2-Month Entry (Government 2-Month Ration)")

with st.form("bulk_form"):
    col1, col2 = st.columns(2)
    
    with col1:
        card_type = st.selectbox("Card Type", ["PHH", "AAY", "SFSS"], key="bulk_card")
        month1 = st.selectbox("Month 1", ["January","February","March","April","May","June","July","August","September","October","November","December"], key="m1")
        year1 = st.number_input("Year 1", value=2025, step=1, key="y1")
        
    with col2:
        month2 = st.selectbox("Month 2", ["January","February","March","April","May","June","July","August","September","October","November","December"], key="m2")
        year2 = st.number_input("Year 2", value=2025, step=1, key="y2")
        
    # ── Pickup Day inputs ──
        st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)
        section_header("📅 Distribution / Pickup Day Reference")
        st.markdown("""
            <div class="info-box">
                📌 Enter the day of the month on which ration was distributed
                (e.g. <strong>15</strong> means the 15th of that month).
                This is stored as a sell reference for records.
            </div>
        """, unsafe_allow_html=True)

        pd1_col, pd2_col = st.columns(2)
        with pd1_col:
            pickup_day1 = st.number_input(
                f"🗓️ Pickup Day – {m1} {y1}",
                min_value=1, max_value=31,
                value=1, step=1, key="pd1",
                help=f"Day of distribution in {m1} {y1}"
            )
            # Show the resolved full date as a reference label
            try:
                ref_date1 = date(int(y1), m1_num, int(pickup_day1))
                st.caption(f"📌 Distribution date: **{ref_date1.strftime('%d %b %Y')}**")
            except ValueError:
                st.caption("⚠️ Invalid day for this month")

        with pd2_col:
            pickup_day2 = st.number_input(
                f"🗓️ Pickup Day – {m2} {y2}",
                min_value=1, max_value=31,
                value=1, step=1, key="pd2",
                help=f"Day of distribution in {m2} {y2}"
            )
            try:
                ref_date2 = date(int(y2), m2_num, int(pickup_day2))
                st.caption(f"📌 Distribution date: **{ref_date2.strftime('%d %b %Y')}**")
            except ValueError:
                st.caption("⚠️ Invalid day for this month")

        if st.button("💾 Save Both Months", use_container_width=True):
            res1 = upsert_entry(
                date1, b_card, opening1, recv1, clos1,
                remarks=f"Bulk entry – {m1} {y1} | Pickup: {int(pickup_day1)}",
                pickup_day=int(pickup_day1)
            )
            res2 = upsert_entry(
                date2, b_card,
                res1["closing_kg"], recv2, clos2,
                remarks=f"Bulk entry – {m2} {y2} | Pickup: {int(pickup_day2)}",
                pickup_day=int(pickup_day2)
            )
# ─────────────────────────────────────────────────────────────────────────────
# PAGE: REPORTS
# ─────────────────────────────────────────────────────────────────────────────
def page_reports():
    render_greeting()
    section_header("📋 Reports & Export")

    months_list = list(calendar.month_name)[1:]
    current_year = today_ist().year

    export_tab, view_tab = st.tabs(["📤 Export Excel", "🔍 View Monthly Data"])

    # ════════════════════════════════════════════════════════════
    # EXPORT TAB
    # ════════════════════════════════════════════════════════════
    with export_tab:
        export_type = st.radio(
            "Export Type",
            ["Single Month", "Multiple Months Combined",
             "Full Year", "Combined 2 Months"],
            horizontal=True,
        )
        exp_card = st.selectbox("Card Type", CARD_TYPES, key="exp_card")

        # ── Single Month ──
        if export_type == "Single Month":
            ec1, ec2 = st.columns(2)
            with ec1:
                em = st.selectbox("Month", months_list,
                                   index=today_ist().month - 1, key="e_sm")
            with ec2:
                ey = st.number_input("Year", min_value=2020,
                                      max_value=2035, value=current_year, key="e_sy")
            em_num = months_list.index(em) + 1
            if st.button("⬇️ Download Excel", key="dl_single"):
                df = get_entries_for_month(int(ey), em_num, exp_card)
                if df.empty:
                    st.warning("No data found for this selection.")
                else:
                    period = f"{em} {ey}"
                    buf = build_excel_report(
                        df,
                        report_title=f"Month / Year - {period}",
                        period_label=period,
                        card_type=exp_card
                    )
                    st.download_button(
                        "📥 Click to Save Excel File",
                        data=buf,
                        file_name=f"FPS_Register_{exp_card}_{em}_{ey}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

        # ── Multiple Months Combined ──
        elif export_type == "Multiple Months Combined":
            sel_months = st.multiselect(
                "Select Months (hold Ctrl/Cmd to multi-select)",
                [f"{m} {y}" for y in range(current_year - 1, current_year + 2)
                 for m in months_list],
                default=[],
            )
            if st.button("⬇️ Prepare Combined Export", key="dl_multi"):
                if not sel_months:
                    st.warning("Please select at least one month.")
                else:
                    frames = []
                    for my in sel_months:
                        parts = my.rsplit(" ", 1)
                        m_num = months_list.index(parts[0]) + 1
                        y_num = int(parts[1])
                        frames.append(get_entries_for_month(y_num, m_num, exp_card))
                    combined = pd.concat(frames, ignore_index=True)
                    if combined.empty:
                        st.warning("No data found for selected months.")
                    else:
                        label = f"{sel_months[0]} to {sel_months[-1]}"
                        buf = build_excel_report(
                            combined,
                            report_title=f"Combined Report - {label}",
                            period_label=label,
                            card_type=exp_card,
                        )
                        st.download_button(
                            "📥 Click to Save Excel File",
                            data=buf,
                            file_name=f"FPS_Register_{exp_card}_Combined.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )

        # ── Full Year ──
        elif export_type == "Full Year":
            fy = st.number_input("Year", min_value=2020,
                                  max_value=2035, value=current_year, key="e_fy")
            if st.button("⬇️ Download Full Year Excel", key="dl_year"):
                df = get_entries_for_year(int(fy))
                df = df[df["card_type"] == exp_card]
                if df.empty:
                    st.warning("No data found for this year.")
                else:
                    buf = build_excel_report(
                        df,
                        report_title=f"Annual Report - {fy}",
                        period_label=f"Year {fy}",
                        card_type=exp_card,
                    )
                    st.download_button(
                        "📥 Click to Save Excel File",
                        data=buf,
                        file_name=f"FPS_Register_{exp_card}_{fy}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

        # ── Combined 2 Months ──
        else:  # Combined 2 Months
            r1, r2 = st.columns(2)
            with r1:
                cm1 = st.selectbox("Month 1", months_list,
                                    index=today_ist().month - 1, key="cm1")
                cy1 = st.number_input("Year 1", min_value=2020,
                                       max_value=2035, value=current_year, key="cy1")
            with r2:
                cm2 = st.selectbox("Month 2", months_list,
                                    index=min(today_ist().month, 11), key="cm2")
                cy2 = st.number_input("Year 2", min_value=2020,
                                       max_value=2035, value=current_year, key="cy2")
            if st.button("⬇️ Download Combined 2-Month Excel", key="dl_2m"):
                cm1_num = months_list.index(cm1) + 1
                cm2_num = months_list.index(cm2) + 1
                df1 = get_entries_for_month(int(cy1), cm1_num, exp_card)
                df2 = get_entries_for_month(int(cy2), cm2_num, exp_card)
                combined = pd.concat([df1, df2], ignore_index=True)
                if combined.empty:
                    st.warning("No data found for selected months.")
                else:
                    label = f"{cm1} {cy1} to {cm2} {cy2}"
                    buf = build_excel_report(
                        combined,
                        report_title=f"Combined Report - {label}",
                        period_label=label,
                        card_type=exp_card,
                    )
                    st.download_button(
                        "📥 Click to Save Excel File",
                        data=buf,
                        file_name=f"FPS_Register_{exp_card}_{cm1}{cy1}_{cm2}{cy2}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

    # ════════════════════════════════════════════════════════════
    # VIEW TAB
    # ════════════════════════════════════════════════════════════
    with view_tab:
        section_header("🔍 View Monthly Ledger")
        vc1, vc2, vc3 = st.columns(3)
        with vc1:
            vm = st.selectbox("Month", months_list,
                               index=today_ist().month - 1, key="v_m")
        with vc2:
            vy = st.number_input("Year", min_value=2020,
                                  max_value=2035, value=current_year, key="v_y")
        with vc3:
            v_card = st.selectbox("Card Type", ["All"] + CARD_TYPES, key="v_card")

        vm_num = months_list.index(vm) + 1
        df = get_entries_for_month(
            int(vy), vm_num,
            None if v_card == "All" else v_card
        )

        if df.empty:
            st.markdown("""
                <div class="info-box">
                    📭 No entries found for this selection.
                    Add entries via the Daily Entry page.
                </div>
            """, unsafe_allow_html=True)
        else:
           display_df["pickup_day"] = display_df["pickup_day"].apply(
                lambda x: f"Day {int(x)}" if pd.notna(x) and x else "—"
            )
            display_df = display_df.rename(columns={
                "entry_date":  "Date",
                "card_type":   "Card",
                "opening_kg":  "Opening",
                "received_kg": "Received",
                "total_kg":    "Total",
                "closing_kg":  "Closing",
                "selling_kg":  "Sold/Issued",
                "pickup_day":  "Pickup Day",
                "remarks":     "Remarks",
            })[["Date", "Card", "Opening", "Received",
                "Total", "Closing", "Sold/Issued", "Pickup Day", "Remarks"]]
            st.dataframe(display_df, use_container_width=True, hide_index=True)

            # Totals summary
            st.markdown("---")
            totals_row = {
                "Total Received": format_qty(df["received_kg"].sum()),
                "Total Sold":     format_qty(df["selling_kg"].sum()),
                "Total Closing":  format_qty(df["closing_kg"].sum()),
            }
            t_cols = st.columns(3)
            for i, (k, v) in enumerate(totals_row.items()):
                with t_cols[i]:
                    st.markdown(f"""
                        <div class="metric-card">
                            <h3>{k}</h3><h2>{v}</h2>
                        </div>
                    """, unsafe_allow_html=True)

    render_footer()


# ─────────────────────────────────────────────────────────────────────────────
# PAGE: SETTINGS
# ─────────────────────────────────────────────────────────────────────────────
def page_settings():
    render_greeting()
    section_header("⚙️ Account Settings")

    st.markdown("""
        <div class="info-box">
            🔐 Update your login credentials below.
            Changes take effect immediately.
        </div>
    """, unsafe_allow_html=True)

    with st.form("settings_form"):
        new_username = st.text_input(
            "👤 New Username",
            value=get_setting("username"),
            placeholder="Enter new username"
        )
        new_password = st.text_input(
            "🔑 New Password",
            type="password",
            placeholder="Enter new password (leave blank to keep current)"
        )
        confirm_pw = st.text_input(
            "🔑 Confirm Password",
            type="password",
            placeholder="Re-enter new password"
        )
        save = st.form_submit_button("💾 Save Changes", use_container_width=True)

    if save:
        if not new_username.strip():
            st.error("❌ Username cannot be empty.")
        elif new_password and new_password != confirm_pw:
            st.error("❌ Passwords do not match.")
        else:
            update_setting("username", new_username.strip())
            if new_password:
                update_setting("password", hash_password(new_password))
            st.success("✅ Settings saved successfully!")
            st.session_state["username"] = new_username.strip()
            st.rerun()

    # ── App info card ──
    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)
    section_header("ℹ️ System Information")
    st.markdown(f"""
        <div class="metric-card">
            <h3>Application Details</h3>
            <p><strong>App Name:</strong> {APP_NAME} - Quintal System</p>
            <p><strong>Dealer:</strong> {DEALER_NAME}</p>
            <p><strong>Code:</strong> {DEALER_CODE}</p>
            <p><strong>Database:</strong> SQLite (rice_inventory.db)</p>
            <p><strong>Timezone:</strong> Indian Standard Time (IST / UTC+5:30)</p>
            <p><strong>Version:</strong> 1.0.0</p>
        </div>
    """, unsafe_allow_html=True)

    render_footer()


# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR NAVIGATION
# ─────────────────────────────────────────────────────────────────────────────
def render_sidebar():
    with st.sidebar:
        st.markdown("""
            <div style="text-align:center; padding: 16px 0 8px 0;">
                <div style="font-size:2.2rem;">🌾</div>
                <div style="font-size:1.2rem; font-weight:700;
                            color:#d8f3dc; letter-spacing:0.05em;">RiceStock</div>
                <div style="font-size:0.75rem; color:#95d5b2;">
                    Quintal System
                </div>
            </div>
        """, unsafe_allow_html=True)
        st.divider()

        st.markdown(f"""
            <div style="font-size:0.8rem; color:#95d5b2;
                        padding: 0 4px 8px 4px; line-height:1.5;">
                👤 {get_setting('username')}<br>
                🏪 Balaram Shial<br>
                🔖 Code: 0201P100
            </div>
        """, unsafe_allow_html=True)
        st.divider()

        nav = st.radio(
            "Navigation",
            ["🏠 Dashboard", "📝 Daily Entry", "📋 Reports", "⚙️ Settings"],
            label_visibility="collapsed",
        )
        st.divider()

        if st.button("🔓 Logout", use_container_width=True):
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.rerun()

        st.markdown("""
            <div style="position:absolute; bottom:20px; left:0; right:0;
                        text-align:center; font-size:0.7rem;
                        color:#52b788; padding: 0 10px;">
                Rice Inventory<br>
                <span style="color:#40916c;">Created by - Sunil Kumar</span>
            </div>
        """, unsafe_allow_html=True)

    return nav


# ─────────────────────────────────────────────────────────────────────────────
# MAIN ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────
def main():
    init_db()
    inject_css()

    if "logged_in" not in st.session_state:
        st.session_state["logged_in"] = False

    if not st.session_state["logged_in"]:
        page_login()
        return

    nav = render_sidebar()

    if nav == "🏠 Dashboard":
        page_dashboard()
    elif nav == "📝 Daily Entry":
        page_daily_entry()
    elif nav == "📋 Reports":
        page_reports()
    elif nav == "⚙️ Settings":
        page_settings()


if __name__ == "__main__":
    main()
