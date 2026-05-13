"""
RiceStock - Quintal System
A professional Rice Inventory Management Web App built with Streamlit.
Single-file, production-ready implementation.
"""

# ─────────────────────────────────────────────────────────────────────────────
# IMPORTS
# ─────────────────────────────────────────────────────────────────────────────
import streamlit as st
import pandas as pd
import sqlite3
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from datetime import datetime, date, timedelta
import calendar
import io
import os
import shutil
import hashlib
import json
from pathlib import Path

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG  (must be first Streamlit call)
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
DB_PATH = "rice_inventory.db"
BACKUP_DIR = "backups"
CARD_TYPES = ["PHH", "AAY", "SFSS"]
TRANSACTION_TYPES = ["Opening", "Received", "Selling", "Closing"]
MONTHS = [calendar.month_name[i] for i in range(1, 13)]
DEFAULT_USER = "admin"
DEFAULT_PASS = hashlib.sha256("rice123".encode()).hexdigest()

# Low-stock threshold per card type (Kg)
LOW_STOCK_THRESHOLD = {"PHH": 500, "AAY": 300, "SFSS": 200}

# ─────────────────────────────────────────────────────────────────────────────
# THEME & CSS
# ─────────────────────────────────────────────────────────────────────────────
def load_css(dark_mode: bool = False):
    bg        = "#0f1b0f" if dark_mode else "#f0f7f0"
    surface   = "#1a2e1a" if dark_mode else "#ffffff"
    surface2  = "#243324" if dark_mode else "#e8f5e8"
    text      = "#e8f5e8" if dark_mode else "#1a2e1a"
    text2     = "#a8c8a8" if dark_mode else "#4a6b4a"
    border    = "#2d4d2d" if dark_mode else "#c8e6c8"
    accent    = "#4caf50"
    accent_h  = "#66bb6a"
    danger    = "#ef5350"
    warn      = "#ff9800"
    info      = "#29b6f6"

    st.markdown(f"""
    <style>
    /* ── Google Font ── */
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;500&display=swap');

    /* ── Root Variables ── */
    :root {{
        --bg:        {bg};
        --surface:   {surface};
        --surface2:  {surface2};
        --text:      {text};
        --text2:     {text2};
        --border:    {border};
        --accent:    {accent};
        --accent-h:  {accent_h};
        --danger:    {danger};
        --warn:      {warn};
        --info:      {info};
        --radius:    12px;
        --shadow:    0 4px 24px rgba(0,0,0,.12);
    }}

    /* ── Global ── */
    html, body, [class*="css"] {{
        font-family: 'Plus Jakarta Sans', sans-serif;
        background-color: var(--bg) !important;
        color: var(--text) !important;
    }}

    /* ── Main area ── */
    .main .block-container {{
        padding: 1.5rem 2rem 3rem;
        max-width: 1300px;
    }}

    /* ── Sidebar ── */
    section[data-testid="stSidebar"] {{
        background: var(--surface) !important;
        border-right: 1px solid var(--border) !important;
    }}
    section[data-testid="stSidebar"] * {{
        color: var(--text) !important;
    }}

    /* ── Cards ── */
    .rs-card {{
        background: var(--surface);
        border: 1px solid var(--border);
        border-radius: var(--radius);
        padding: 1.25rem 1.5rem;
        box-shadow: var(--shadow);
        margin-bottom: 1rem;
    }}
    .rs-card-accent {{
        border-left: 4px solid var(--accent);
    }}
    .rs-card-danger {{
        border-left: 4px solid var(--danger);
    }}
    .rs-card-warn {{
        border-left: 4px solid var(--warn);
    }}
    .rs-card-info {{
        border-left: 4px solid var(--info);
    }}

    /* ── Metric Cards ── */
    .rs-metric {{
        background: var(--surface);
        border: 1px solid var(--border);
        border-radius: var(--radius);
        padding: 1.1rem 1.25rem;
        text-align: center;
        box-shadow: var(--shadow);
    }}
    .rs-metric-label {{
        font-size: .75rem;
        font-weight: 600;
        letter-spacing: .08em;
        text-transform: uppercase;
        color: var(--text2);
        margin-bottom: .3rem;
    }}
    .rs-metric-value {{
        font-size: 1.6rem;
        font-weight: 800;
        color: var(--accent);
        line-height: 1.1;
        font-family: 'JetBrains Mono', monospace;
    }}
    .rs-metric-sub {{
        font-size: .72rem;
        color: var(--text2);
        margin-top: .25rem;
    }}

    /* ── Page Title ── */
    .rs-page-title {{
        font-size: 1.75rem;
        font-weight: 800;
        color: var(--text);
        display: flex;
        align-items: center;
        gap: .5rem;
        margin-bottom: 1.5rem;
        border-bottom: 2px solid var(--border);
        padding-bottom: .75rem;
    }}

    /* ── Section Header ── */
    .rs-section {{
        font-size: 1rem;
        font-weight: 700;
        color: var(--text2);
        text-transform: uppercase;
        letter-spacing: .07em;
        margin: 1.25rem 0 .75rem;
    }}

    /* ── Alert Boxes ── */
    .rs-alert-danger {{
        background: rgba(239,83,80,.12);
        border: 1px solid var(--danger);
        border-radius: var(--radius);
        padding: .85rem 1.1rem;
        color: var(--danger);
        font-weight: 600;
        font-size: .9rem;
    }}
    .rs-alert-warn {{
        background: rgba(255,152,0,.12);
        border: 1px solid var(--warn);
        border-radius: var(--radius);
        padding: .85rem 1.1rem;
        color: var(--warn);
        font-weight: 600;
        font-size: .9rem;
    }}
    .rs-alert-success {{
        background: rgba(76,175,80,.12);
        border: 1px solid var(--accent);
        border-radius: var(--radius);
        padding: .85rem 1.1rem;
        color: var(--accent);
        font-weight: 600;
        font-size: .9rem;
    }}

    /* ── Badge ── */
    .rs-badge {{
        display: inline-block;
        padding: .2rem .6rem;
        border-radius: 100px;
        font-size: .72rem;
        font-weight: 700;
        letter-spacing: .05em;
    }}
    .rs-badge-green  {{ background: rgba(76,175,80,.2);  color: #66bb6a; }}
    .rs-badge-blue   {{ background: rgba(41,182,246,.2); color: #29b6f6; }}
    .rs-badge-orange {{ background: rgba(255,152,0,.2);  color: #ffa726; }}
    .rs-badge-red    {{ background: rgba(239,83,80,.2);  color: #ef5350; }}

    /* ── Streamlit overrides ── */
    div[data-testid="metric-container"] {{
        background: var(--surface);
        border: 1px solid var(--border);
        border-radius: var(--radius);
        padding: 1rem 1.25rem;
        box-shadow: var(--shadow);
    }}
    div[data-testid="metric-container"] label {{
        color: var(--text2) !important;
        font-size: .78rem !important;
        font-weight: 600 !important;
        text-transform: uppercase;
        letter-spacing: .06em;
    }}
    div[data-testid="metric-container"] div[data-testid="metric-value"] {{
        color: var(--accent) !important;
        font-family: 'JetBrains Mono', monospace !important;
        font-weight: 800 !important;
    }}
    .stButton > button {{
        background: var(--accent) !important;
        color: #fff !important;
        border: none !important;
        border-radius: 8px !important;
        font-weight: 700 !important;
        font-size: .9rem !important;
        padding: .55rem 1.4rem !important;
        transition: background .2s ease;
        font-family: 'Plus Jakarta Sans', sans-serif !important;
    }}
    .stButton > button:hover {{
        background: var(--accent-h) !important;
    }}
    /* secondary button */
    .stButton > button[kind="secondary"] {{
        background: var(--surface2) !important;
        color: var(--text) !important;
        border: 1px solid var(--border) !important;
    }}
    .stSelectbox > div > div,
    .stTextInput > div > div > input,
    .stNumberInput > div > div > input,
    .stTextArea > div > div > textarea {{
        background: var(--surface2) !important;
        color: var(--text) !important;
        border: 1px solid var(--border) !important;
        border-radius: 8px !important;
        font-family: 'Plus Jakarta Sans', sans-serif !important;
    }}
    .stDataFrame {{
        border: 1px solid var(--border) !important;
        border-radius: var(--radius) !important;
        overflow: hidden;
    }}
    div[data-testid="stExpander"] {{
        border: 1px solid var(--border) !important;
        border-radius: var(--radius) !important;
        background: var(--surface) !important;
    }}
    .stProgress > div > div > div {{
        background: var(--accent) !important;
    }}
    .stTabs [data-baseweb="tab-list"] {{
        gap: .5rem;
    }}
    .stTabs [data-baseweb="tab"] {{
        background: var(--surface2) !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
        font-family: 'Plus Jakarta Sans', sans-serif !important;
    }}
    .stTabs [aria-selected="true"] {{
        background: var(--accent) !important;
        color: #fff !important;
    }}

    /* ── Logo area ── */
    .rs-logo {{
        display: flex;
        align-items: center;
        gap: .75rem;
        padding: 1rem 0 1.5rem;
        border-bottom: 1px solid var(--border);
        margin-bottom: 1rem;
    }}
    .rs-logo-icon {{
        font-size: 2rem;
        line-height: 1;
    }}
    .rs-logo-text {{
        font-size: 1.2rem;
        font-weight: 800;
        color: var(--accent);
        letter-spacing: -.01em;
    }}
    .rs-logo-sub {{
        font-size: .68rem;
        color: var(--text2);
        font-weight: 500;
        letter-spacing: .05em;
        text-transform: uppercase;
    }}

    /* ── Login card ── */
    .rs-login-wrap {{
        max-width: 420px;
        margin: 4rem auto;
        background: var(--surface);
        border: 1px solid var(--border);
        border-radius: 20px;
        padding: 2.5rem;
        box-shadow: 0 8px 48px rgba(0,0,0,.2);
    }}
    .rs-login-title {{
        font-size: 1.5rem;
        font-weight: 800;
        color: var(--text);
        margin-bottom: .25rem;
    }}
    .rs-login-sub {{
        font-size: .85rem;
        color: var(--text2);
        margin-bottom: 1.75rem;
    }}
    </style>
    """, unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# DATABASE HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def get_conn():
    """Return a SQLite connection (check_same_thread=False for Streamlit)."""
    return sqlite3.connect(DB_PATH, check_same_thread=False)


def create_tables():
    """Create all necessary tables if they don't exist."""
    conn = get_conn()
    c = conn.cursor()

    # Main inventory table
    c.execute("""
        CREATE TABLE IF NOT EXISTS inventory (
            id               INTEGER PRIMARY KEY AUTOINCREMENT,
            entry_date       TEXT    NOT NULL,
            day              INTEGER NOT NULL,
            month            INTEGER NOT NULL,
            year             INTEGER NOT NULL,
            card_type        TEXT    NOT NULL,
            transaction_type TEXT    NOT NULL,
            quantity_kg      REAL    NOT NULL,
            remarks          TEXT,
            created_at       TEXT    DEFAULT (datetime('now'))
        )
    """)

    # Users table
    c.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            username     TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            role         TEXT DEFAULT 'admin'
        )
    """)

    # Settings table
    c.execute("""
        CREATE TABLE IF NOT EXISTS settings (
            key   TEXT PRIMARY KEY,
            value TEXT
        )
    """)

    # Insert default admin if not present
    c.execute("SELECT COUNT(*) FROM users WHERE username=?", (DEFAULT_USER,))
    if c.fetchone()[0] == 0:
        c.execute("INSERT INTO users (username, password_hash) VALUES (?,?)",
                  (DEFAULT_USER, DEFAULT_PASS))

    conn.commit()
    conn.close()


def add_entry(entry_date, card_type, transaction_type, quantity_kg, remarks=""):
    """Insert a new inventory entry."""
    d = datetime.strptime(entry_date, "%Y-%m-%d")
    conn = get_conn()
    conn.execute("""
        INSERT INTO inventory (entry_date, day, month, year, card_type,
                               transaction_type, quantity_kg, remarks)
        VALUES (?,?,?,?,?,?,?,?)
    """, (entry_date, d.day, d.month, d.year,
          card_type, transaction_type, quantity_kg, remarks))
    conn.commit()
    conn.close()


def delete_entry(entry_id: int):
    """Delete an entry by ID."""
    conn = get_conn()
    conn.execute("DELETE FROM inventory WHERE id=?", (entry_id,))
    conn.commit()
    conn.close()


def get_all_entries() -> pd.DataFrame:
    """Return all inventory entries as a DataFrame."""
    conn = get_conn()
    df = pd.read_sql("SELECT * FROM inventory ORDER BY entry_date DESC, id DESC", conn)
    conn.close()
    return df


def get_entries_by_date(entry_date: str) -> pd.DataFrame:
    conn = get_conn()
    df = pd.read_sql(
        "SELECT * FROM inventory WHERE entry_date=? ORDER BY id",
        conn, params=(entry_date,)
    )
    conn.close()
    return df


def get_entries_by_month(month: int, year: int) -> pd.DataFrame:
    conn = get_conn()
    df = pd.read_sql(
        "SELECT * FROM inventory WHERE month=? AND year=? ORDER BY entry_date, id",
        conn, params=(month, year)
    )
    conn.close()
    return df


def get_entries_by_year(year: int) -> pd.DataFrame:
    conn = get_conn()
    df = pd.read_sql(
        "SELECT * FROM inventory WHERE year=? ORDER BY entry_date, id",
        conn, params=(year,)
    )
    conn.close()
    return df


def get_current_stock(card_type: str) -> float:
    """
    Calculate current stock for a card type.
    Formula: Opening + Received - Selling (most recent closing if present).
    """
    conn = get_conn()
    c = conn.cursor()

    # Sum by transaction type
    c.execute("""
        SELECT transaction_type, SUM(quantity_kg) FROM inventory
        WHERE card_type=?
        GROUP BY transaction_type
    """, (card_type,))
    rows = dict(c.fetchall())
    conn.close()

    opening  = rows.get("Opening",  0) or 0
    received = rows.get("Received", 0) or 0
    selling  = rows.get("Selling",  0) or 0
    closing  = rows.get("Closing",  0) or 0

    # If closing entries exist, use the last closing
    if closing > 0:
        return closing
    return opening + received - selling


def get_stock_trend(days: int = 30) -> pd.DataFrame:
    """Return daily closing stock per card type for the last N days."""
    conn = get_conn()
    cutoff = (date.today() - timedelta(days=days)).isoformat()
    df = pd.read_sql("""
        SELECT entry_date, card_type, transaction_type, quantity_kg
        FROM inventory
        WHERE entry_date >= ?
        ORDER BY entry_date
    """, conn, params=(cutoff,))
    conn.close()
    return df


def get_monthly_summary(month: int, year: int) -> pd.DataFrame:
    """Pivot table: day × card_type × transaction for a month."""
    df = get_entries_by_month(month, year)
    if df.empty:
        return df
    pivot = df.pivot_table(
        index=["entry_date", "day", "card_type"],
        columns="transaction_type",
        values="quantity_kg",
        aggfunc="sum"
    ).reset_index()
    pivot.columns.name = None
    # Fill missing columns
    for col in TRANSACTION_TYPES:
        if col not in pivot.columns:
            pivot[col] = 0
    pivot = pivot.fillna(0)
    return pivot


def verify_password(username: str, password: str) -> bool:
    conn = get_conn()
    c = conn.cursor()
    pw_hash = hashlib.sha256(password.encode()).hexdigest()
    c.execute("SELECT COUNT(*) FROM users WHERE username=? AND password_hash=?",
              (username, pw_hash))
    result = c.fetchone()[0] > 0
    conn.close()
    return result


def change_password(username: str, new_password: str):
    pw_hash = hashlib.sha256(new_password.encode()).hexdigest()
    conn = get_conn()
    conn.execute("UPDATE users SET password_hash=? WHERE username=?",
                 (pw_hash, username))
    conn.commit()
    conn.close()


def get_setting(key: str, default=None):
    conn = get_conn()
    c = conn.cursor()
    c.execute("SELECT value FROM settings WHERE key=?", (key,))
    row = c.fetchone()
    conn.close()
    return row[0] if row else default


def set_setting(key: str, value: str):
    conn = get_conn()
    conn.execute("INSERT OR REPLACE INTO settings (key,value) VALUES (?,?)",
                 (key, str(value)))
    conn.commit()
    conn.close()


def backup_database() -> bytes:
    """Return the raw bytes of the DB file for download."""
    with open(DB_PATH, "rb") as f:
        return f.read()


def restore_database(data: bytes):
    """Write uploaded bytes as the new DB file."""
    with open(DB_PATH, "wb") as f:
        f.write(data)


# ─────────────────────────────────────────────────────────────────────────────
# UNIT HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def kg_to_display(kg: float) -> str:
    """Smart unit display: ≥100 Kg → 'X Q.YY Kg', else 'XX.XX Kg'."""
    if kg is None:
        return "0 Kg"
    kg = round(kg, 2)
    if kg >= 100:
        quintals = int(kg // 100)
        remainder = round(kg % 100, 2)
        return f"{quintals} Q.{remainder:05.2f} Kg"
    return f"{kg:.2f} Kg"


def input_to_kg(value: float, unit: str) -> float:
    """Convert user input to Kg. unit is 'Kg' or 'Quintal'."""
    if unit == "Quintal":
        return round(value * 100, 4)
    return round(value, 4)


# ─────────────────────────────────────────────────────────────────────────────
# EXPORT HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def df_to_excel_bytes(df: pd.DataFrame, sheet_name="Data") -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
    return buf.getvalue()


def df_to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8")


# ─────────────────────────────────────────────────────────────────────────────
# CHART HELPERS
# ─────────────────────────────────────────────────────────────────────────────
PLOTLY_THEME_LIGHT = dict(
    paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
    font_color="#1a2e1a", gridcolor="#c8e6c8"
)
PLOTLY_THEME_DARK = dict(
    paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
    font_color="#e8f5e8", gridcolor="#2d4d2d"
)

def get_plotly_cfg(dark: bool) -> dict:
    return PLOTLY_THEME_DARK if dark else PLOTLY_THEME_LIGHT


# ─────────────────────────────────────────────────────────────────────────────
# SESSION STATE INIT
# ─────────────────────────────────────────────────────────────────────────────
def init_session():
    defaults = {
        "logged_in": False,
        "username": "",
        "dark_mode": False,
        "page": "🏠 Home",
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

    # Load theme preference from DB
    if "theme_loaded" not in st.session_state:
        val = get_setting("dark_mode", "false")
        st.session_state.dark_mode = val.lower() == "true"
        st.session_state.theme_loaded = True


# ─────────────────────────────────────────────────────────────────────────────
# LOGIN PAGE
# ─────────────────────────────────────────────────────────────────────────────
def page_login():
    st.markdown("""
    <div style="text-align:center; padding: 2rem 0 1rem;">
        <span style="font-size:4rem;">🌾</span>
        <h1 style="font-size:2.2rem; font-weight:900; margin:.3rem 0 .1rem; color:var(--accent);">
            RiceStock
        </h1>
        <p style="color:var(--text2); font-size:.95rem; letter-spacing:.05em;">
            QUINTAL SYSTEM &nbsp;•&nbsp; INVENTORY MANAGEMENT
        </p>
    </div>
    """, unsafe_allow_html=True)

    col_l, col_c, col_r = st.columns([1, 1.2, 1])
    with col_c:
        with st.container():
            st.markdown('<div class="rs-card">', unsafe_allow_html=True)
            st.markdown("#### 🔐 Sign In to your account")
            username = st.text_input("Username", placeholder="admin")
            password = st.text_input("Password", type="password", placeholder="••••••••")

            if st.button("Sign In →", use_container_width=True):
                if verify_password(username, password):
                    st.session_state.logged_in = True
                    st.session_state.username = username
                    st.rerun()
                else:
                    st.markdown(
                        '<div class="rs-alert-danger">❌ Invalid username or password.</div>',
                        unsafe_allow_html=True
                    )

            st.markdown("""
            <div style="margin-top:1.2rem; padding-top:1rem; border-top:1px solid var(--border);
                        font-size:.78rem; color:var(--text2); text-align:center;">
                Default credentials: <code>admin</code> / <code>rice123</code>
            </div>
            """, unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────
PAGES = [
    "🏠 Home",
    "📝 Daily Entry",
    "📊 Dashboard",
    "📋 Monthly Reports",
    "📅 Yearly Summary",
    "⬆️ Import / Export",
    "⚙️ Settings",
]

def render_sidebar():
    with st.sidebar:
        st.markdown("""
        <div class="rs-logo">
            <span class="rs-logo-icon">🌾</span>
            <div>
                <div class="rs-logo-text">RiceStock</div>
                <div class="rs-logo-sub">Quintal System</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        # Navigation
        st.markdown('<div class="rs-section">Navigation</div>', unsafe_allow_html=True)
        for p in PAGES:
            active = st.session_state.page == p
            if st.button(
                p,
                key=f"nav_{p}",
                use_container_width=True,
                type="primary" if active else "secondary",
            ):
                st.session_state.page = p
                st.rerun()

        # Quick stock summary
        st.markdown("---")
        st.markdown('<div class="rs-section">Quick Stock</div>', unsafe_allow_html=True)
        for ct in CARD_TYPES:
            stock = get_current_stock(ct)
            badge_cls = "rs-badge-green" if stock >= LOW_STOCK_THRESHOLD[ct] else "rs-badge-red"
            st.markdown(
                f'<span class="rs-badge {badge_cls}">{ct}</span>&nbsp; '
                f'<b style="font-family:JetBrains Mono,monospace;">{kg_to_display(stock)}</b><br>',
                unsafe_allow_html=True
            )

        st.markdown("---")
        st.markdown(
            f'<div style="font-size:.78rem; color:var(--text2);">👤 {st.session_state.username}</div>',
            unsafe_allow_html=True
        )
        if st.button("🚪 Sign Out", use_container_width=True, type="secondary"):
            st.session_state.logged_in = False
            st.session_state.username = ""
            st.rerun()


# ─────────────────────────────────────────────────────────────────────────────
# PAGE: HOME
# ─────────────────────────────────────────────────────────────────────────────
def page_home():
    st.markdown('<div class="rs-page-title">🏠 Home</div>', unsafe_allow_html=True)
    today = date.today()
    dark = st.session_state.dark_mode

    # Welcome banner
    st.markdown(f"""
    <div class="rs-card rs-card-accent">
        <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:1rem;">
            <div>
                <div style="font-size:1.2rem; font-weight:800;">👋 Welcome back, {st.session_state.username}!</div>
                <div style="color:var(--text2); font-size:.88rem; margin-top:.25rem;">
                    📅 {today.strftime("%A, %d %B %Y")} &nbsp;•&nbsp; RiceStock Quintal System
                </div>
            </div>
            <div style="text-align:right;">
                <div style="font-size:2.5rem; font-weight:900; color:var(--accent); font-family:'JetBrains Mono',monospace;">
                    {today.strftime("%H:%M")}
                </div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # Current stock metrics
    st.markdown('<div class="rs-section">📦 Current Stock Overview</div>', unsafe_allow_html=True)
    cols = st.columns(3)
    card_colors = {"PHH": "#4caf50", "AAY": "#29b6f6", "SFSS": "#ff9800"}
    for i, ct in enumerate(CARD_TYPES):
        stock = get_current_stock(ct)
        threshold = LOW_STOCK_THRESHOLD[ct]
        pct = min(stock / (threshold * 3) * 100, 100)
        status = "✅ Good" if stock >= threshold else "⚠️ Low"
        with cols[i]:
            st.markdown(f"""
            <div class="rs-card" style="border-left:4px solid {card_colors[ct]};">
                <div style="display:flex; justify-content:space-between; align-items:flex-start;">
                    <div>
                        <div class="rs-metric-label">{ct} Card</div>
                        <div style="font-size:1.4rem; font-weight:800; color:{card_colors[ct]};
                                    font-family:'JetBrains Mono',monospace;">
                            {kg_to_display(stock)}
                        </div>
                    </div>
                    <div style="font-size:1.8rem;">{"🌾" if i==0 else "🌿" if i==1 else "🍃"}</div>
                </div>
                <div style="margin-top:.75rem; background:var(--surface2); border-radius:100px; height:6px;">
                    <div style="width:{pct}%; background:{card_colors[ct]}; border-radius:100px; height:6px;"></div>
                </div>
                <div style="font-size:.72rem; color:var(--text2); margin-top:.4rem;">
                    {status} &nbsp;•&nbsp; Threshold: {threshold} Kg
                </div>
            </div>
            """, unsafe_allow_html=True)

    # Today's entries
    st.markdown('<div class="rs-section">📋 Today\'s Entries</div>', unsafe_allow_html=True)
    today_df = get_entries_by_date(today.isoformat())
    if today_df.empty:
        st.markdown(
            '<div class="rs-alert-warn">📭 No entries recorded for today. Go to Daily Entry to add.</div>',
            unsafe_allow_html=True
        )
    else:
        disp_df = today_df[["card_type", "transaction_type", "quantity_kg", "remarks"]].copy()
        disp_df["quantity_kg"] = disp_df["quantity_kg"].apply(kg_to_display)
        disp_df.columns = ["Card Type", "Transaction", "Quantity", "Remarks"]
        st.dataframe(disp_df, use_container_width=True, hide_index=True)

    # Low-stock alerts
    low_stocks = [(ct, get_current_stock(ct)) for ct in CARD_TYPES
                  if get_current_stock(ct) < LOW_STOCK_THRESHOLD[ct]]
    if low_stocks:
        st.markdown('<div class="rs-section">⚠️ Alerts</div>', unsafe_allow_html=True)
        for ct, stk in low_stocks:
            st.markdown(
                f'<div class="rs-alert-danger">🚨 <b>{ct}</b> stock is critically low: '
                f'{kg_to_display(stk)} (Threshold: {LOW_STOCK_THRESHOLD[ct]} Kg)</div>',
                unsafe_allow_html=True
            )

    # Mini trend chart
    st.markdown('<div class="rs-section">📈 Stock Trend (Last 30 Days)</div>', unsafe_allow_html=True)
    trend_df = get_stock_trend(30)
    if not trend_df.empty:
        _render_trend_chart(trend_df, dark)
    else:
        st.info("Not enough data to show trend chart yet.")


def _render_trend_chart(trend_df: pd.DataFrame, dark: bool):
    cfg = get_plotly_cfg(dark)
    # Aggregate daily by card type
    agg = (
        trend_df[trend_df["transaction_type"].isin(["Received", "Selling", "Opening"])]
        .groupby(["entry_date", "card_type"])["quantity_kg"]
        .sum()
        .reset_index()
    )
    if agg.empty:
        st.info("Insufficient data for trend chart.")
        return

    colors = {"PHH": "#4caf50", "AAY": "#29b6f6", "SFSS": "#ff9800"}
    fig = go.Figure()
    for ct in CARD_TYPES:
        sub = agg[agg["card_type"] == ct]
        if not sub.empty:
            fig.add_trace(go.Scatter(
                x=sub["entry_date"], y=sub["quantity_kg"],
                name=ct, mode="lines+markers",
                line=dict(color=colors[ct], width=2.5),
                marker=dict(size=6),
                fill="tozeroy", fillcolor=colors[ct].replace(")", ",.08)").replace("rgb", "rgba")
                    if "rgb" in colors[ct] else colors[ct] + "14",
            ))
    fig.update_layout(
        height=320, margin=dict(l=0, r=0, t=10, b=0),
        paper_bgcolor=cfg["paper_bgcolor"],
        plot_bgcolor=cfg["plot_bgcolor"],
        font_color=cfg["font_color"],
        legend=dict(orientation="h", yanchor="bottom", y=1, xanchor="right", x=1),
        xaxis=dict(showgrid=True, gridcolor=cfg["gridcolor"]),
        yaxis=dict(showgrid=True, gridcolor=cfg["gridcolor"]),
    )
    st.plotly_chart(fig, use_container_width=True)


# ─────────────────────────────────────────────────────────────────────────────
# PAGE: DAILY ENTRY
# ─────────────────────────────────────────────────────────────────────────────
def page_daily_entry():
    st.markdown('<div class="rs-page-title">📝 Daily Entry</div>', unsafe_allow_html=True)

    tab_add, tab_view, tab_del = st.tabs(["➕ Add Entry", "📄 View by Date", "🗑️ Delete Entry"])

    # ── Add Entry ──
    with tab_add:
        with st.form("entry_form", clear_on_submit=True):
            st.markdown('<div class="rs-section">Entry Details</div>', unsafe_allow_html=True)
            c1, c2 = st.columns(2)
            with c1:
                entry_date = st.date_input("📅 Date", value=date.today())
                card_type  = st.selectbox("🏷️ Card Type", CARD_TYPES)
            with c2:
                txn_type = st.selectbox("🔄 Transaction Type", TRANSACTION_TYPES)
                unit     = st.selectbox("⚖️ Unit", ["Kg", "Quintal"])

            qty_label = "Quantity (Kg)" if unit == "Kg" else "Quantity (Quintals)"
            qty_val   = st.number_input(qty_label, min_value=0.0, step=0.5, format="%.2f")
            remarks   = st.text_area("📝 Remarks (optional)", height=80)

            submitted = st.form_submit_button("✅ Add Entry", use_container_width=True)

        if submitted:
            qty_kg = input_to_kg(qty_val, unit)
            if qty_kg <= 0:
                st.markdown(
                    '<div class="rs-alert-danger">❌ Quantity must be greater than 0.</div>',
                    unsafe_allow_html=True
                )
            else:
                # Validate: closing stock cannot be negative
                if txn_type == "Selling":
                    current = get_current_stock(card_type)
                    if qty_kg > current:
                        st.markdown(
                            f'<div class="rs-alert-danger">❌ Selling quantity ({kg_to_display(qty_kg)}) '
                            f'exceeds current stock ({kg_to_display(current)}) for {card_type}!</div>',
                            unsafe_allow_html=True
                        )
                    else:
                        add_entry(entry_date.isoformat(), card_type, txn_type, qty_kg, remarks)
                        st.markdown(
                            f'<div class="rs-alert-success">✅ Entry added: <b>{txn_type}</b> '
                            f'<b>{kg_to_display(qty_kg)}</b> for <b>{card_type}</b> on '
                            f'<b>{entry_date}</b>.</div>',
                            unsafe_allow_html=True
                        )
                else:
                    add_entry(entry_date.isoformat(), card_type, txn_type, qty_kg, remarks)
                    st.markdown(
                        f'<div class="rs-alert-success">✅ Entry added: <b>{txn_type}</b> '
                        f'<b>{kg_to_display(qty_kg)}</b> for <b>{card_type}</b> on '
                        f'<b>{entry_date}</b>.</div>',
                        unsafe_allow_html=True
                    )

        # Show today's entries
        st.markdown('<div class="rs-section">📋 Today\'s Entries</div>', unsafe_allow_html=True)
        today_df = get_entries_by_date(date.today().isoformat())
        if today_df.empty:
            st.info("No entries today yet.")
        else:
            disp = today_df[["id", "card_type", "transaction_type", "quantity_kg", "remarks"]].copy()
            disp["quantity_kg"] = disp["quantity_kg"].apply(kg_to_display)
            disp.columns = ["ID", "Card", "Transaction", "Quantity", "Remarks"]
            st.dataframe(disp, use_container_width=True, hide_index=True)

    # ── View by Date ──
    with tab_view:
        sel_date = st.date_input("Select Date", value=date.today(), key="view_date")
        df = get_entries_by_date(sel_date.isoformat())
        if df.empty:
            st.info(f"No entries for {sel_date}.")
        else:
            disp = df[["id","card_type","transaction_type","quantity_kg","remarks"]].copy()
            disp["quantity_kg"] = disp["quantity_kg"].apply(kg_to_display)
            disp.columns = ["ID","Card","Transaction","Quantity","Remarks"]
            st.dataframe(disp, use_container_width=True, hide_index=True)

            # Per card summary
            st.markdown('<div class="rs-section">Summary</div>', unsafe_allow_html=True)
            summary = df.groupby(["card_type", "transaction_type"])["quantity_kg"].sum().reset_index()
            summary["quantity_kg"] = summary["quantity_kg"].apply(kg_to_display)
            summary.columns = ["Card","Transaction","Total Quantity"]
            st.dataframe(summary, use_container_width=True, hide_index=True)

    # ── Delete Entry ──
    with tab_del:
        st.markdown(
            '<div class="rs-alert-warn">⚠️ Deletion is permanent. Use with caution.</div>',
            unsafe_allow_html=True
        )
        del_date = st.date_input("Select Date to Load Entries", value=date.today(), key="del_date")
        df_del = get_entries_by_date(del_date.isoformat())
        if df_del.empty:
            st.info("No entries for selected date.")
        else:
            disp = df_del[["id","card_type","transaction_type","quantity_kg","remarks"]].copy()
            disp["quantity_kg"] = disp["quantity_kg"].apply(kg_to_display)
            disp.columns = ["ID","Card","Transaction","Quantity","Remarks"]
            st.dataframe(disp, use_container_width=True, hide_index=True)

            entry_id = st.number_input("Enter Entry ID to delete", min_value=1, step=1)
            if st.button("🗑️ Delete Entry", type="primary"):
                ids_available = df_del["id"].tolist()
                if entry_id in ids_available:
                    delete_entry(int(entry_id))
                    st.success(f"Entry {entry_id} deleted.")
                    st.rerun()
                else:
                    st.error("ID not found in the selected date's entries.")


# ─────────────────────────────────────────────────────────────────────────────
# PAGE: DASHBOARD
# ─────────────────────────────────────────────────────────────────────────────
def page_dashboard():
    st.markdown('<div class="rs-page-title">📊 Dashboard</div>', unsafe_allow_html=True)
    dark = st.session_state.dark_mode
    now  = datetime.now()
    month, year = now.month, now.year

    # ── Stock metrics ──
    st.markdown('<div class="rs-section">📦 Current Stock</div>', unsafe_allow_html=True)
    cols = st.columns(3)
    card_colors = {"PHH": "#4caf50", "AAY": "#29b6f6", "SFSS": "#ff9800"}
    stocks = {ct: get_current_stock(ct) for ct in CARD_TYPES}
    total_stock = sum(stocks.values())

    for i, ct in enumerate(CARD_TYPES):
        stock = stocks[ct]
        thresh = LOW_STOCK_THRESHOLD[ct]
        pct = min(stock / max(thresh * 3, 1) * 100, 100)
        with cols[i]:
            st.metric(
                label=f"{'🌾' if i==0 else '🌿' if i==1 else '🍃'} {ct} Stock",
                value=kg_to_display(stock),
                delta=f"Threshold: {thresh} Kg"
            )
            st.progress(int(pct))

    # ── This month summary ──
    st.markdown('<div class="rs-section">📅 This Month\'s Summary</div>', unsafe_allow_html=True)
    month_df = get_entries_by_month(month, year)

    def _month_total(txn): 
        return month_df[month_df["transaction_type"]==txn]["quantity_kg"].sum() if not month_df.empty else 0

    m_received = _month_total("Received")
    m_selling  = _month_total("Selling")
    m_opening  = _month_total("Opening")
    m_balance  = m_opening + m_received - m_selling

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("📥 Opening",  kg_to_display(m_opening))
    c2.metric("📦 Received", kg_to_display(m_received))
    c3.metric("📤 Sold",     kg_to_display(m_selling))
    c4.metric("💰 Balance",  kg_to_display(m_balance))

    # ── Overall summary ──
    all_df = get_all_entries()
    if not all_df.empty:
        st.markdown('<div class="rs-section">🌐 Overall Summary</div>', unsafe_allow_html=True)
        def _overall(txn):
            return all_df[all_df["transaction_type"]==txn]["quantity_kg"].sum()

        o1, o2, o3 = st.columns(3)
        o1.metric("📦 Total Received (All Time)", kg_to_display(_overall("Received")))
        o2.metric("📤 Total Sold (All Time)",     kg_to_display(_overall("Selling")))
        o3.metric("📊 Total Entries",             str(len(all_df)))

    # ── Low stock alerts ──
    low = [(ct, stocks[ct]) for ct in CARD_TYPES if stocks[ct] < LOW_STOCK_THRESHOLD[ct]]
    if low:
        st.markdown('<div class="rs-section">⚠️ Low Stock Alerts</div>', unsafe_allow_html=True)
        for ct, stk in low:
            st.markdown(
                f'<div class="rs-alert-danger">🚨 <b>{ct}</b>: {kg_to_display(stk)} '
                f'— Below threshold ({LOW_STOCK_THRESHOLD[ct]} Kg)</div>',
                unsafe_allow_html=True
            )

    # ── Trend chart ──
    st.markdown('<div class="rs-section">📈 30-Day Stock Movement</div>', unsafe_allow_html=True)
    trend_df = get_stock_trend(30)
    if not trend_df.empty:
        _render_trend_chart(trend_df, dark)
    else:
        st.info("No trend data available yet.")

    # ── Donut chart ──
    if total_stock > 0:
        st.markdown('<div class="rs-section">🥧 Stock Distribution</div>', unsafe_allow_html=True)
        cfg = get_plotly_cfg(dark)
        fig_pie = go.Figure(go.Pie(
            labels=list(stocks.keys()),
            values=list(stocks.values()),
            hole=.55,
            marker_colors=["#4caf50", "#29b6f6", "#ff9800"],
            textinfo="label+percent",
        ))
        fig_pie.add_annotation(
            text=f"<b>{kg_to_display(total_stock)}</b><br>Total",
            x=.5, y=.5, font_size=13,
            showarrow=False, font_color=cfg["font_color"]
        )
        fig_pie.update_layout(
            height=320, margin=dict(l=0,r=0,t=10,b=0),
            paper_bgcolor=cfg["paper_bgcolor"],
            font_color=cfg["font_color"],
            showlegend=True,
        )
        st.plotly_chart(fig_pie, use_container_width=True)

    # ── Transaction breakdown bar chart ──
    if not all_df.empty:
        st.markdown('<div class="rs-section">📊 Transaction Breakdown by Card</div>', unsafe_allow_html=True)
        grp = all_df.groupby(["card_type","transaction_type"])["quantity_kg"].sum().reset_index()
        cfg = get_plotly_cfg(dark)
        fig_bar = px.bar(
            grp, x="card_type", y="quantity_kg", color="transaction_type",
            barmode="group",
            color_discrete_map={
                "Opening": "#78909c", "Received": "#4caf50",
                "Selling": "#ef5350", "Closing": "#29b6f6"
            },
            labels={"quantity_kg": "Quantity (Kg)", "card_type": "Card Type", "transaction_type": "Transaction"},
        )
        fig_bar.update_layout(
            height=340, margin=dict(l=0,r=0,t=10,b=0),
            paper_bgcolor=cfg["paper_bgcolor"],
            plot_bgcolor=cfg["plot_bgcolor"],
            font_color=cfg["font_color"],
        )
        st.plotly_chart(fig_bar, use_container_width=True)


# ─────────────────────────────────────────────────────────────────────────────
# PAGE: MONTHLY REPORTS
# ─────────────────────────────────────────────────────────────────────────────
def page_monthly_reports():
    st.markdown('<div class="rs-page-title">📋 Monthly Reports</div>', unsafe_allow_html=True)
    dark = st.session_state.dark_mode
    now  = datetime.now()

    c1, c2 = st.columns(2)
    with c1:
        sel_month = st.selectbox("Month", MONTHS, index=now.month - 1)
    with c2:
        sel_year = st.selectbox("Year", list(range(now.year, now.year - 6, -1)), index=0)

    month_num = MONTHS.index(sel_month) + 1
    df = get_entries_by_month(month_num, sel_year)

    if df.empty:
        st.markdown(
            f'<div class="rs-alert-warn">📭 No data for {sel_month} {sel_year}.</div>',
            unsafe_allow_html=True
        )
        return

    # Summary cards
    st.markdown(f'<div class="rs-section">Summary — {sel_month} {sel_year}</div>', unsafe_allow_html=True)
    def _tot(txn): return df[df["transaction_type"]==txn]["quantity_kg"].sum()
    c1,c2,c3,c4 = st.columns(4)
    c1.metric("📥 Opening",  kg_to_display(_tot("Opening")))
    c2.metric("📦 Received", kg_to_display(_tot("Received")))
    c3.metric("📤 Sold",     kg_to_display(_tot("Selling")))
    c4.metric("🔒 Closing",  kg_to_display(_tot("Closing") or (_tot("Opening")+_tot("Received")-_tot("Selling"))))

    # Per card type tabs
    tab_all, *card_tabs = st.tabs(["All Cards"] + CARD_TYPES)

    def _show_card_table(data: pd.DataFrame, label: str):
        if data.empty:
            st.info(f"No data for {label}.")
            return
        pivot = data.pivot_table(
            index=["entry_date", "day"],
            columns="transaction_type",
            values="quantity_kg",
            aggfunc="sum"
        ).reset_index()
        pivot.columns.name = None
        for col in TRANSACTION_TYPES:
            if col not in pivot.columns:
                pivot[col] = 0
        pivot = pivot.fillna(0)
        display_pivot = pivot.copy()
        for col in TRANSACTION_TYPES:
            if col in display_pivot.columns:
                display_pivot[col] = display_pivot[col].apply(kg_to_display)
        display_pivot.columns = [
            c.replace("entry_date","Date").replace("day","Day") for c in display_pivot.columns
        ]
        st.dataframe(display_pivot, use_container_width=True, hide_index=True)
        return pivot

    with tab_all:
        _show_card_table(df, "All Cards")
        # Bar chart
        grp = df.groupby(["entry_date","transaction_type"])["quantity_kg"].sum().reset_index()
        cfg = get_plotly_cfg(dark)
        fig = px.bar(
            grp, x="entry_date", y="quantity_kg", color="transaction_type",
            barmode="stack",
            color_discrete_map={
                "Opening":"#78909c","Received":"#4caf50","Selling":"#ef5350","Closing":"#29b6f6"
            },
            labels={"quantity_kg":"Kg","entry_date":"Date","transaction_type":"Transaction"},
        )
        fig.update_layout(height=320, margin=dict(l=0,r=0,t=10,b=0),
                          paper_bgcolor=cfg["paper_bgcolor"],
                          plot_bgcolor=cfg["plot_bgcolor"],
                          font_color=cfg["font_color"])
        st.plotly_chart(fig, use_container_width=True)

    for ct, ctab in zip(CARD_TYPES, card_tabs):
        with ctab:
            _show_card_table(df[df["card_type"]==ct], ct)

    # Downloads
    st.markdown('<div class="rs-section">⬇️ Download</div>', unsafe_allow_html=True)
    dl1, dl2 = st.columns(2)
    with dl1:
        excel = df_to_excel_bytes(df, sheet_name=f"{sel_month}_{sel_year}")
        st.download_button(
            "📥 Download Excel",
            data=excel,
            file_name=f"ricestock_{sel_month}_{sel_year}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with dl2:
        csv = df_to_csv_bytes(df)
        st.download_button(
            "📥 Download CSV",
            data=csv,
            file_name=f"ricestock_{sel_month}_{sel_year}.csv",
            mime="text/csv",
            use_container_width=True,
        )


# ─────────────────────────────────────────────────────────────────────────────
# PAGE: YEARLY SUMMARY
# ─────────────────────────────────────────────────────────────────────────────
def page_yearly_summary():
    st.markdown('<div class="rs-page-title">📅 Yearly Summary</div>', unsafe_allow_html=True)
    dark = st.session_state.dark_mode
    now  = datetime.now()

    sel_year = st.selectbox("Year", list(range(now.year, now.year - 6, -1)), index=0)
    df = get_entries_by_year(sel_year)

    if df.empty:
        st.markdown(
            f'<div class="rs-alert-warn">📭 No data for {sel_year}.</div>',
            unsafe_allow_html=True
        )
        return

    # Overall year metrics
    def _tot(txn): return df[df["transaction_type"]==txn]["quantity_kg"].sum()
    st.markdown(f'<div class="rs-section">Year {sel_year} Overview</div>', unsafe_allow_html=True)
    c1,c2,c3,c4 = st.columns(4)
    c1.metric("📥 Total Opening",  kg_to_display(_tot("Opening")))
    c2.metric("📦 Total Received", kg_to_display(_tot("Received")))
    c3.metric("📤 Total Sold",     kg_to_display(_tot("Selling")))
    c4.metric("📊 Total Entries",  str(len(df)))

    # Monthly pivot table
    st.markdown('<div class="rs-section">📋 Month-wise Breakdown</div>', unsafe_allow_html=True)
    monthly = df.groupby(["month","card_type","transaction_type"])["quantity_kg"].sum().reset_index()

    # Build display table
    rows = []
    for m in range(1, 13):
        m_data = monthly[monthly["month"] == m]
        if m_data.empty:
            continue
        row = {"Month": calendar.month_name[m]}
        for ct in CARD_TYPES:
            ct_data = m_data[m_data["card_type"] == ct]
            for txn in ["Received","Selling"]:
                val = ct_data[ct_data["transaction_type"]==txn]["quantity_kg"].sum()
                row[f"{ct} {txn}"] = kg_to_display(val) if val else "-"
        rows.append(row)

    if rows:
        summary_df = pd.DataFrame(rows)
        st.dataframe(summary_df, use_container_width=True, hide_index=True)

    # Monthly bar chart
    st.markdown('<div class="rs-section">📊 Monthly Flow</div>', unsafe_allow_html=True)
    month_grp = df.groupby(["month","transaction_type"])["quantity_kg"].sum().reset_index()
    month_grp["month_name"] = month_grp["month"].apply(lambda x: calendar.month_abbr[x])
    cfg = get_plotly_cfg(dark)
    fig = px.bar(
        month_grp, x="month_name", y="quantity_kg", color="transaction_type",
        barmode="group",
        color_discrete_map={
            "Opening":"#78909c","Received":"#4caf50","Selling":"#ef5350","Closing":"#29b6f6"
        },
        labels={"quantity_kg":"Kg","month_name":"Month","transaction_type":"Transaction"},
    )
    fig.update_layout(
        height=360, margin=dict(l=0,r=0,t=10,b=0),
        paper_bgcolor=cfg["paper_bgcolor"],
        plot_bgcolor=cfg["plot_bgcolor"],
        font_color=cfg["font_color"],
    )
    st.plotly_chart(fig, use_container_width=True)

    # Per card type heatmap
    st.markdown('<div class="rs-section">🗓️ Card-wise Yearly Heatmap</div>', unsafe_allow_html=True)
    for ct in CARD_TYPES:
        with st.expander(f"🌾 {ct} — Monthly Received & Sold"):
            ct_data = df[df["card_type"] == ct]
            if ct_data.empty:
                st.write("No data.")
                continue
            ct_grp = ct_data.groupby(["month","transaction_type"])["quantity_kg"].sum().reset_index()
            ct_grp["month_name"] = ct_grp["month"].apply(lambda x: calendar.month_name[x])
            ct_fig = px.line(
                ct_grp, x="month_name", y="quantity_kg", color="transaction_type",
                markers=True,
                color_discrete_map={
                    "Opening":"#78909c","Received":"#4caf50","Selling":"#ef5350","Closing":"#29b6f6"
                },
                labels={"quantity_kg":"Kg","month_name":"Month"},
            )
            ct_fig.update_layout(
                height=260, margin=dict(l=0,r=0,t=10,b=0),
                paper_bgcolor=cfg["paper_bgcolor"],
                plot_bgcolor=cfg["plot_bgcolor"],
                font_color=cfg["font_color"],
            )
            st.plotly_chart(ct_fig, use_container_width=True)

    # Downloads
    st.markdown('<div class="rs-section">⬇️ Download</div>', unsafe_allow_html=True)
    d1, d2 = st.columns(2)
    with d1:
        st.download_button(
            "📥 Download Year Excel",
            data=df_to_excel_bytes(df, f"Year_{sel_year}"),
            file_name=f"ricestock_{sel_year}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with d2:
        st.download_button(
            "📥 Download Year CSV",
            data=df_to_csv_bytes(df),
            file_name=f"ricestock_{sel_year}.csv",
            mime="text/csv",
            use_container_width=True,
        )


# ─────────────────────────────────────────────────────────────────────────────
# PAGE: IMPORT / EXPORT
# ─────────────────────────────────────────────────────────────────────────────
def page_import_export():
    st.markdown('<div class="rs-page-title">⬆️ Import / Export</div>', unsafe_allow_html=True)
    tab_exp, tab_imp, tab_bak = st.tabs(["📤 Export Data", "📥 Import Data", "💾 Backup & Restore"])

    # ── Export ──
    with tab_exp:
        st.markdown('<div class="rs-section">Export Options</div>', unsafe_allow_html=True)
        exp_scope = st.radio("Scope", ["Full Data", "Month-wise", "Year-wise"], horizontal=True)

        if exp_scope == "Full Data":
            df = get_all_entries()
            if df.empty:
                st.info("No data to export.")
            else:
                e1, e2 = st.columns(2)
                with e1:
                    st.download_button(
                        "📥 Full Excel",
                        data=df_to_excel_bytes(df, "FullData"),
                        file_name="ricestock_full.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
                with e2:
                    st.download_button(
                        "📥 Full CSV",
                        data=df_to_csv_bytes(df),
                        file_name="ricestock_full.csv",
                        mime="text/csv",
                        use_container_width=True,
                    )
                st.markdown(f"**{len(df)} rows** ready to export.")

        elif exp_scope == "Month-wise":
            now = datetime.now()
            c1, c2 = st.columns(2)
            with c1:
                m = st.selectbox("Month", MONTHS, index=now.month-1, key="exp_m")
            with c2:
                y = st.selectbox("Year", list(range(now.year, now.year-6, -1)), key="exp_y")
            df = get_entries_by_month(MONTHS.index(m)+1, y)
            if df.empty:
                st.info(f"No data for {m} {y}.")
            else:
                e1, e2 = st.columns(2)
                with e1:
                    st.download_button("📥 Excel", df_to_excel_bytes(df, f"{m}_{y}"),
                                       f"ricestock_{m}_{y}.xlsx",
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                       use_container_width=True)
                with e2:
                    st.download_button("📥 CSV", df_to_csv_bytes(df),
                                       f"ricestock_{m}_{y}.csv", mime="text/csv",
                                       use_container_width=True)

        else:  # Year-wise
            y = st.selectbox("Year", list(range(datetime.now().year, datetime.now().year-6, -1)),
                             key="exp_yr")
            df = get_entries_by_year(y)
            if df.empty:
                st.info(f"No data for {y}.")
            else:
                e1, e2 = st.columns(2)
                with e1:
                    st.download_button("📥 Excel", df_to_excel_bytes(df, f"Year_{y}"),
                                       f"ricestock_{y}.xlsx",
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                       use_container_width=True)
                with e2:
                    st.download_button("📥 CSV", df_to_csv_bytes(df),
                                       f"ricestock_{y}.csv", mime="text/csv",
                                       use_container_width=True)

    # ── Import ──
    with tab_imp:
        st.markdown('<div class="rs-section">Import from Excel / CSV</div>', unsafe_allow_html=True)
        st.markdown("""
        <div class="rs-card rs-card-info">
            <b>📋 Required columns:</b><br>
            <code>entry_date, card_type, transaction_type, quantity_kg</code><br>
            <span style="font-size:.8rem;">Optional: <code>remarks</code></span><br><br>
            <b>Card Types:</b> PHH, AAY, SFSS &nbsp;&nbsp; <b>Transactions:</b> Opening, Received, Selling, Closing
        </div>
        """, unsafe_allow_html=True)

        uploaded = st.file_uploader("Upload Excel or CSV", type=["xlsx","csv","xls"])
        if uploaded:
            try:
                if uploaded.name.endswith(".csv"):
                    imp_df = pd.read_csv(uploaded)
                else:
                    imp_df = pd.read_excel(uploaded)

                st.markdown("**Preview (first 10 rows):**")
                st.dataframe(imp_df.head(10), use_container_width=True, hide_index=True)

                # Validate
                required = {"entry_date","card_type","transaction_type","quantity_kg"}
                missing = required - set(imp_df.columns.str.lower())
                if missing:
                    st.error(f"Missing columns: {missing}")
                else:
                    imp_df.columns = imp_df.columns.str.lower()
                    bad_cards = imp_df[~imp_df["card_type"].isin(CARD_TYPES)]
                    bad_txns  = imp_df[~imp_df["transaction_type"].isin(TRANSACTION_TYPES)]

                    if not bad_cards.empty:
                        st.warning(f"⚠️ {len(bad_cards)} rows have invalid card_type. They will be skipped.")
                    if not bad_txns.empty:
                        st.warning(f"⚠️ {len(bad_txns)} rows have invalid transaction_type. They will be skipped.")

                    valid = imp_df[
                        imp_df["card_type"].isin(CARD_TYPES) &
                        imp_df["transaction_type"].isin(TRANSACTION_TYPES)
                    ]

                    st.markdown(f"**{len(valid)} valid rows** ready to import.")
                    if st.button("⬆️ Import Valid Rows", use_container_width=True):
                        count = 0
                        for _, row in valid.iterrows():
                            try:
                                add_entry(
                                    str(row["entry_date"])[:10],
                                    str(row["card_type"]),
                                    str(row["transaction_type"]),
                                    float(row["quantity_kg"]),
                                    str(row.get("remarks","")) if pd.notna(row.get("remarks","")) else ""
                                )
                                count += 1
                            except Exception:
                                pass
                        st.markdown(
                            f'<div class="rs-alert-success">✅ Successfully imported {count} rows.</div>',
                            unsafe_allow_html=True
                        )
                        st.rerun()

            except Exception as e:
                st.error(f"Error reading file: {e}")

    # ── Backup & Restore ──
    with tab_bak:
        st.markdown('<div class="rs-section">Backup Database</div>', unsafe_allow_html=True)
        st.markdown(
            '<div class="rs-card">Download the full SQLite database file for safe-keeping.</div>',
            unsafe_allow_html=True
        )
        db_bytes = backup_database()
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        st.download_button(
            "💾 Download Database Backup",
            data=db_bytes,
            file_name=f"ricestock_backup_{ts}.db",
            mime="application/octet-stream",
            use_container_width=True,
        )

        st.markdown("---")
        st.markdown('<div class="rs-section">Restore Database</div>', unsafe_allow_html=True)
        st.markdown(
            '<div class="rs-alert-danger">⚠️ Restoring will OVERWRITE all current data. This cannot be undone.</div>',
            unsafe_allow_html=True
        )
        restore_file = st.file_uploader("Upload .db backup file", type=["db"])
        confirm = st.checkbox("I understand this will overwrite all current data")
        if st.button("🔄 Restore Database", type="primary") and restore_file and confirm:
            restore_database(restore_file.read())
            st.markdown(
                '<div class="rs-alert-success">✅ Database restored successfully. Please refresh the page.</div>',
                unsafe_allow_html=True
            )


# ─────────────────────────────────────────────────────────────────────────────
# PAGE: SETTINGS
# ─────────────────────────────────────────────────────────────────────────────
def page_settings():
    st.markdown('<div class="rs-page-title">⚙️ Settings</div>', unsafe_allow_html=True)

    tab_theme, tab_thresh, tab_pwd, tab_about = st.tabs(
        ["🎨 Theme", "🚨 Stock Thresholds", "🔑 Change Password", "ℹ️ About"]
    )

    # ── Theme ──
    with tab_theme:
        st.markdown('<div class="rs-section">Display Theme</div>', unsafe_allow_html=True)
        current_dark = st.session_state.dark_mode
        new_dark = st.toggle("🌙 Dark Mode", value=current_dark)
        if new_dark != current_dark:
            st.session_state.dark_mode = new_dark
            set_setting("dark_mode", str(new_dark).lower())
            st.rerun()
        st.markdown(
            f'<div class="rs-card">Current theme: <b>{"🌙 Dark" if st.session_state.dark_mode else "☀️ Light"}</b></div>',
            unsafe_allow_html=True
        )

    # ── Thresholds ──
    with tab_thresh:
        st.markdown('<div class="rs-section">Low Stock Alert Thresholds (Kg)</div>', unsafe_allow_html=True)
        st.markdown(
            '<div class="rs-card rs-card-warn">Alerts are triggered when stock falls below these values.</div>',
            unsafe_allow_html=True
        )
        for ct in CARD_TYPES:
            val = st.number_input(
                f"{ct} Threshold (Kg)",
                min_value=0, value=LOW_STOCK_THRESHOLD[ct], step=10,
                key=f"thresh_{ct}"
            )
            LOW_STOCK_THRESHOLD[ct] = val  # runtime update (not persisted to DB in this version)
        st.info("ℹ️ Threshold changes apply for this session only.")

    # ── Password ──
    with tab_pwd:
        st.markdown('<div class="rs-section">Change Password</div>', unsafe_allow_html=True)
        with st.form("pwd_form", clear_on_submit=True):
            old_pwd = st.text_input("Current Password", type="password")
            new_pwd = st.text_input("New Password", type="password")
            new_pwd2 = st.text_input("Confirm New Password", type="password")
            if st.form_submit_button("🔑 Update Password", use_container_width=True):
                if not verify_password(st.session_state.username, old_pwd):
                    st.error("Current password is incorrect.")
                elif new_pwd != new_pwd2:
                    st.error("New passwords do not match.")
                elif len(new_pwd) < 6:
                    st.error("Password must be at least 6 characters.")
                else:
                    change_password(st.session_state.username, new_pwd)
                    st.success("✅ Password updated successfully!")

    # ── About ──
    with tab_about:
        st.markdown("""
        <div class="rs-card rs-card-accent">
            <div style="font-size:2rem; margin-bottom:.5rem;">🌾</div>
            <div style="font-size:1.3rem; font-weight:800;">RiceStock — Quintal System</div>
            <div style="color:var(--text2); margin:.25rem 0 1rem; font-size:.88rem;">
                Professional Rice Inventory Management
            </div>
            <div style="display:grid; grid-template-columns:1fr 1fr; gap:.5rem; font-size:.85rem;">
                <div>📦 <b>Version:</b> 1.0.0</div>
                <div>🗄️ <b>Database:</b> SQLite</div>
                <div>⚙️ <b>Framework:</b> Streamlit</div>
                <div>📊 <b>Charts:</b> Plotly</div>
                <div>🐍 <b>Language:</b> Python 3.x</div>
                <div>📅 <b>Year:</b> 2025</div>
            </div>
        </div>
        <div class="rs-card" style="margin-top:.75rem;">
            <div style="font-weight:700; margin-bottom:.5rem;">📋 Features</div>
            <ul style="margin:0; padding-left:1.2rem; color:var(--text2); font-size:.88rem;">
                <li>Quintal / Kg smart unit display</li>
                <li>PHH · AAY · SFSS card types</li>
                <li>Full SQLite persistence</li>
                <li>Monthly &amp; Yearly reports</li>
                <li>Excel &amp; CSV import/export</li>
                <li>Database backup &amp; restore</li>
                <li>Dark / Light theme</li>
                <li>Low stock alerts</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)

        # DB stats
        all_df = get_all_entries()
        if not all_df.empty:
            st.markdown('<div class="rs-section">Database Statistics</div>', unsafe_allow_html=True)
            s1,s2,s3 = st.columns(3)
            s1.metric("Total Entries", len(all_df))
            s2.metric("Card Types", all_df["card_type"].nunique())
            s3.metric("Date Range",
                      f"{all_df['entry_date'].min()[:10]} → {all_df['entry_date'].max()[:10]}")

        db_size = os.path.getsize(DB_PATH) if os.path.exists(DB_PATH) else 0
        st.caption(f"📁 Database size: {db_size/1024:.1f} KB  •  Path: `{DB_PATH}`")


# ─────────────────────────────────────────────────────────────────────────────
# MAIN APP ROUTER
# ─────────────────────────────────────────────────────────────────────────────
def main():
    # Init DB and session
    create_tables()
    init_session()

    # Load CSS
    load_css(st.session_state.get("dark_mode", False))

    # Auth gate
    if not st.session_state.logged_in:
        page_login()
        return

    # Sidebar navigation
    render_sidebar()

    # Route to page
    page = st.session_state.page
    if page == "🏠 Home":
        page_home()
    elif page == "📝 Daily Entry":
        page_daily_entry()
    elif page == "📊 Dashboard":
        page_dashboard()
    elif page == "📋 Monthly Reports":
        page_monthly_reports()
    elif page == "📅 Yearly Summary":
        page_yearly_summary()
    elif page == "⬆️ Import / Export":
        page_import_export()
    elif page == "⚙️ Settings":
        page_settings()


if __name__ == "__main__":
    main()
