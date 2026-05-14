import streamlit as st
import pandas as pd
import sqlite3
from datetime import datetime
from zoneinfo import ZoneInfo
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side, Alignment, PatternFill
from io import BytesIO
import plotly.express as px

st.set_page_config(
    page_title="RiceStock - Official FPS Inventory System",
    page_icon="🌾",
    layout="wide"
)

DB_NAME = "rice_inventory.db"

st.markdown("""
<style>
.main {
    background-color: #f6fff6;
}

.stButton>button {
    background-color: #2e7d32;
    color: white;
    border-radius: 10px;
    border: none;
    padding: 0.6rem 1rem;
    font-weight: 600;
}

.stButton>button:hover {
    background-color: #1b5e20;
    color: white;
}

div[data-testid="stSidebar"] {
    background-color: #e8f5e9;
}

.block-container {
    padding-top: 1rem;
}

.footer {
    text-align:center;
    margin-top:40px;
    padding:10px;
    color:gray;
    font-size:14px;
}
</style>
""", unsafe_allow_html=True)

def connect_db():
    return sqlite3.connect(DB_NAME, check_same_thread=False)

conn = connect_db()
cursor = conn.cursor()

cursor.execute("""
CREATE TABLE IF NOT EXISTS inventory (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    entry_date TEXT,
    entry_time TEXT,
    opening_balance REAL,
    received REAL,
    total REAL,
    selling REAL,
    closing_balance REAL,
    remarks TEXT
)
""")
conn.commit()

def get_ist_now():
    return datetime.now(ZoneInfo("Asia/Kolkata"))

def get_greeting():
    hour = get_ist_now().hour

    if hour < 12:
        return "Good Morning Mr. Sunil"
    elif hour < 17:
        return "Good Afternoon Mr. Sunil"
    return "Good Evening Mr. Sunil"

def kg_to_display(kg):
    if kg >= 100:
        quintal = int(kg // 100)
        remaining = kg % 100
        return f"{quintal} Q.{remaining:.2f} Kg"
    return f"{kg:.2f} Kg"

def get_previous_closing():
    cursor.execute("""
    SELECT closing_balance FROM inventory
    ORDER BY id DESC LIMIT 1
    """)
    data = cursor.fetchone()

    if data:
        return float(data[0])
    return 0.0

def save_entry(data):
    cursor.execute("""
    INSERT INTO inventory (
        entry_date,
        entry_time,
        opening_balance,
        received,
        total,
        selling,
        closing_balance,
        remarks
    )
    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, data)
    conn.commit()

def load_data():
    return pd.read_sql_query(
        "SELECT * FROM inventory ORDER BY id DESC",
        conn
    )

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

def login_page():
    col1, col2, col3 = st.columns([1,2,1])

    with col2:
        st.markdown("<h1 style='text-align:center;color:#2e7d32;'>🌾 RiceStock</h1>", unsafe_allow_html=True)
        st.markdown("<h3 style='text-align:center;'>Dealer - Balaram Shial</h3>", unsafe_allow_html=True)
        st.markdown("<h5 style='text-align:center;color:gray;'>Code - 0201P100</h5>", unsafe_allow_html=True)

        username = st.text_input("Username")
        password = st.text_input("Password", type="password")

        if st.button("Login", use_container_width=True):
            if username == "Sunil" and password == "sunil123":
                st.session_state.logged_in = True
                st.success("Login Successful")
                st.rerun()
            else:
                st.error("Invalid Username or Password")

if not st.session_state.logged_in:
    login_page()
    st.stop()

st.sidebar.title("🌾 RiceStock")

menu = st.sidebar.radio(
    "Navigation",
    [
        "🏠 Home",
        "📝 Daily Entry",
        "📊 Dashboard",
        "📋 Monthly Reports",
        "📅 Yearly Summary",
        "⬆️ Import / Export",
        "⚙️ Settings"
    ]
)

if menu == "🏠 Home":
    st.title("RiceStock - Official FPS Inventory System")

    st.success(get_greeting())

    now = get_ist_now()

    col1, col2, col3 = st.columns(3)

    with col1:
        st.metric("Current Date", now.strftime("%d-%m-%Y"))

    with col2:
        st.metric("Current Time", now.strftime("%I:%M:%S %p"))

    with col3:
        st.metric("Current Stock", kg_to_display(get_previous_closing()))

    df = load_data()

    if not df.empty:
        chart = px.line(
            df.sort_values("id"),
            x="entry_date",
            y="closing_balance",
            title="Stock Movement Trend"
        )
        st.plotly_chart(chart, use_container_width=True)

elif menu == "📝 Daily Entry":
    st.title("📝 Daily Inventory Entry")

    opening = get_previous_closing()

    st.info(f"Opening Balance: {kg_to_display(opening)}")

    received = st.number_input(
        "Received Quantity (Kg)",
        min_value=0.0,
        step=1.0
    )

    closing = st.number_input(
        "Closing Balance (Kg)",
        min_value=0.0,
        step=1.0
    )

    remarks = st.text_input("Remarks")

    total = opening + received
    selling = total - closing

    if closing > total:
        received = closing - opening
        total = opening + received
        selling = 0

        st.warning("Received quantity auto-adjusted.")

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.metric("Opening", kg_to_display(opening))

    with col2:
        st.metric("Total", kg_to_display(total))

    with col3:
        st.metric("Selling", kg_to_display(selling))

    with col4:
        st.metric("Closing", kg_to_display(closing))

    if st.button("Save Entry"):
        now = get_ist_now()

        data = (
            now.strftime("%d-%m-%Y"),
            now.strftime("%I:%M:%S %p"),
            opening,
            received,
            total,
            selling,
            closing,
            remarks
        )

        save_entry(data)

        st.success("Inventory Entry Saved Successfully")

elif menu == "📊 Dashboard":
    st.title("📊 Inventory Dashboard")

    df = load_data()

    if not df.empty:
        total_received = df["received"].sum()
        total_sold = df["selling"].sum()
        current_stock = df.iloc[0]["closing_balance"]

        c1, c2, c3 = st.columns(3)

        with c1:
            st.metric("Total Received", kg_to_display(total_received))

        with c2:
            st.metric("Total Sold", kg_to_display(total_sold))

        with c3:
            st.metric("Current Stock", kg_to_display(current_stock))

        fig = px.bar(
            df.sort_values("id"),
            x="entry_date",
            y=["received", "selling", "closing_balance"],
            barmode="group"
        )

        st.plotly_chart(fig, use_container_width=True)

        st.dataframe(df, use_container_width=True)

elif menu == "📋 Monthly Reports":
    st.title("📋 Monthly Reports")

    df = load_data()

    if not df.empty:
        st.dataframe(df, use_container_width=True)

elif menu == "📅 Yearly Summary":
    st.title("📅 Yearly Summary")

    df = load_data()

    if not df.empty:
        summary = df.groupby(df["entry_date"]).agg({
            "received": "sum",
            "selling": "sum",
            "closing_balance": "last"
        }).reset_index()

        st.dataframe(summary, use_container_width=True)

elif menu == "⬆️ Import / Export":
    st.title("⬆️ Export FPS Stock Register")

    df = load_data()

    if not df.empty:

        def create_excel(dataframe):
            wb = Workbook()
            ws = wb.active
            ws.title = "FPS Register"

            title_font = Font(bold=True, size=14)

            thin = Side(style='thin')

            border = Border(
                left=thin,
                right=thin,
                top=thin,
                bottom=thin
            )

            ws.merge_cells('A1:G1')
            ws['A1'] = "FPS Stock Register"
            ws['A1'].font = title_font
            ws['A1'].alignment = Alignment(horizontal='center')

            headers = [
                "Date of Transaction",
                "Opening Balance (Quintal)",
                "Receipt (Quintal)",
                "Total (Quintal)",
                "Number of Cards",
                "Quantity Issued",
                "Closing Balance (Quintal)"
            ]

            ws.append(headers)

            for _, row in dataframe.iterrows():
                ws.append([
                    row["entry_date"],
                    row["opening_balance"] / 100,
                    row["received"] / 100,
                    row["total"] / 100,
                    "",
                    row["selling"] / 100,
                    row["closing_balance"] / 100
                ])

            for row in ws.iter_rows():
                for cell in row:
                    cell.border = border
                    cell.alignment = Alignment(horizontal='center')

            output = BytesIO()
            wb.save(output)
            output.seek(0)

            return output

        excel_file = create_excel(df)

        st.download_button(
            label="📥 Download Excel Report",
            data=excel_file,
            file_name="FPS_Stock_Register.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

elif menu == "⚙️ Settings":
    st.title("⚙️ Settings")

    st.write("""
    - Project Name: RiceStock
    - Developer: Sunil Kumar
    - Database: SQLite
    - Timezone: IST
    """)

st.markdown("""
<div class='footer'>
    <strong>Rice Inventory</strong><br>
    Created by - Sunil Kumar
</div>
""", unsafe_allow_html=True)
