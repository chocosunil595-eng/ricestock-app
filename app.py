import streamlit as st
import sqlite3
import pandas as pd
import datetime
import pytz
import plotly.express as px
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
import io

# ==========================================
# 1. CONFIGURATION & SETUP
# ==========================================
st.set_page_config(page_title="RiceStock - FPS Inventory", page_icon="🌾", layout="wide")

IST = pytz.timezone('Asia/Kolkata')

# Default Settings
DEFAULT_DEALER = "Balaram Shial"
DEFAULT_CODE = "0201P100"
DEFAULT_USER = "Sunil"
DEFAULT_PASS = "sunil123"

# Theme settings (Green/Agriculture)
st.markdown("""
    <style>
    .stButton>button {background-color: #2e7d32; color: white;}
    .stProgress .st-bo {background-color: #4caf50;}
    .footer {position: fixed; left: 0; bottom: 0; width: 100%; background-color: #f1f1f1; color: #333; text-align: center; padding: 10px; font-weight: bold; z-index: 100;}
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 2. DATABASE INITIALIZATION
# ==========================================
def init_db():
    conn = sqlite3.connect("rice_inventory.db")
    c = conn.cursor()
    # Inventory Table (Stores everything in KG)
    c.execute('''CREATE TABLE IF NOT EXISTS inventory (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    date TEXT,
                    card_type TEXT,
                    opening_kg REAL,
                    received_kg REAL,
                    total_kg REAL,
                    sold_kg REAL,
                    closing_kg REAL
                )''')
    # Settings/Credentials Table
    c.execute('''CREATE TABLE IF NOT EXISTS settings (
                    username TEXT, password TEXT, dealer_name TEXT, code TEXT
                )''')
    c.execute("SELECT * FROM settings")
    if not c.fetchone():
        c.execute("INSERT INTO settings VALUES (?, ?, ?, ?)", (DEFAULT_USER, DEFAULT_PASS, DEFAULT_DEALER, DEFAULT_CODE))
    conn.commit()
    conn.close()

init_db()

# ==========================================
# 3. HELPER FUNCTIONS
# ==========================================
def get_current_ist_time():
    return datetime.datetime.now(IST)

def get_greeting():
    hour = get_current_ist_time().hour
    if hour < 12: return "Good Morning"
    elif hour < 17: return "Good Afternoon"
    else: return "Good Evening"

def format_kg(kg_val):
    """Smart display: >= 100 Kg -> X Q.YY Kg, else XX Kg"""
    if kg_val is None: return "0 Kg"
    kg_val = round(float(kg_val), 2)
    if kg_val >= 100:
        q = int(kg_val // 100)
        rem = kg_val % 100
        return f"{q} Q.{int(rem):02d} Kg"
    else:
        return f"{int(kg_val)} Kg"

def get_db_connection():
    return sqlite3.connect("rice_inventory.db")

def get_settings():
    conn = get_db_connection()
    df = pd.read_sql_query("SELECT * FROM settings LIMIT 1", conn)
    conn.close()
    return df.iloc[0]

def get_previous_closing(card_type, target_date):
    conn = get_db_connection()
    c = conn.cursor()
    c.execute('''SELECT closing_kg FROM inventory 
                 WHERE card_type=? AND date < ? 
                 ORDER BY date DESC LIMIT 1''', (card_type, target_date.strftime("%Y-%m-%d")))
    row = c.fetchone()
    conn.close()
    return row[0] if row else 0.0

# ==========================================
# 4. AUTHENTICATION (LOGIN)
# ==========================================
if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False

if not st.session_state.logged_in:
    st.markdown("<h1 style='text-align: center; color: #2e7d32;'>🌾 RiceStock - FPS Inventory System</h1>", unsafe_allow_html=True)
    
    settings = get_settings()
    
    col1, col2, col3 = st.columns([1,2,1])
    with col2:
        st.markdown(f"<h3 style='text-align: center;'>Dealer - {settings['dealer_name']}</h3>", unsafe_allow_html=True)
        st.markdown(f"<h5 style='text-align: center;'>Code - {settings['code']}</h5>", unsafe_allow_html=True)
        st.divider()
        
        user_input = st.text_input("Username")
        pass_input = st.text_input("Password", type="password")
        if st.button("Login", use_container_width=True):
            if user_input == settings['username'] and pass_input == settings['password']:
                st.session_state.logged_in = True
                st.session_state.username = user_input
                st.rerun()
            else:
                st.error("Invalid Credentials!")
    st.stop()

# ==========================================
# 5. NAVIGATION & LAYOUT
# ==========================================
settings = get_settings()
st.sidebar.title("🌾 RiceStock")
st.sidebar.markdown(f"**Dealer:** {settings['dealer_name']}<br>**Code:** {settings['code']}", unsafe_allow_html=True)
st.sidebar.divider()

menu = ["🏠 Home", "📝 Daily Entry", "📊 Dashboard", "📋 Monthly Reports", "📅 Yearly Summary", "⚙️ Settings"]
choice = st.sidebar.radio("Navigation", menu)

st.sidebar.divider()
if st.sidebar.button("Logout"):
    st.session_state.logged_in = False
    st.rerun()

# ==========================================
# 6. PAGES LOGIC
# ==========================================

# --- HOME ---
if choice == "🏠 Home":
    st.title(f"{get_greeting()} Mr. {st.session_state.username}!")
    st.markdown(f"**Current Date & Time (IST):** {get_current_ist_time().strftime('%A, %d %b %Y - %I:%M %p')}")
    st.info("Welcome to RiceStock. Navigate using the sidebar to manage your daily FPS inventory.")

# --- DAILY ENTRY ---
elif choice == "📝 Daily Entry":
    st.title("📝 Daily Entry")
    
    col1, col2 = st.columns(2)
    with col1:
        entry_date = st.date_input("Transaction Date", value=get_current_ist_time().date())
    with col2:
        card_type = st.selectbox("Card Type", ["PHH", "AAY", "SFSS"])

    st.subheader(f"Inventory Details: {card_type}")
    
    # Auto-fetch opening balance
    opening_kg = get_previous_closing(card_type, entry_date)
    st.text_input("Opening Balance", value=format_kg(opening_kg), disabled=True)
    
    with st.form("entry_form"):
        received_kg = st.number_input("Received Quantity (in Kg)", min_value=0.0, step=1.0)
        closing_kg = st.number_input("Closing Balance (in Kg)", min_value=0.0, step=1.0)
        
        submitted = st.form_submit_button("Save Entry")
        
        if submitted:
            total_kg = opening_kg + received_kg
            
            # Logic: If user enters Closing > Total, adjust Received
            if closing_kg > total_kg:
                st.warning("Closing Balance exceeds Total. Automatically adjusting 'Received' quantity.")
                received_kg = closing_kg - opening_kg
                total_kg = opening_kg + received_kg
                sold_kg = 0.0
            else:
                sold_kg = total_kg - closing_kg
                
            conn = get_db_connection()
            c = conn.cursor()
            # Upsert logic based on date and card_type
            c.execute("SELECT id FROM inventory WHERE date=? AND card_type=?", (entry_date.strftime("%Y-%m-%d"), card_type))
            existing = c.fetchone()
            
            if existing:
                c.execute('''UPDATE inventory SET opening_kg=?, received_kg=?, total_kg=?, sold_kg=?, closing_kg=? 
                             WHERE id=?''', (opening_kg, received_kg, total_kg, sold_kg, closing_kg, existing[0]))
            else:
                c.execute('''INSERT INTO inventory (date, card_type, opening_kg, received_kg, total_kg, sold_kg, closing_kg) 
                             VALUES (?, ?, ?, ?, ?, ?, ?)''', 
                          (entry_date.strftime("%Y-%m-%d"), card_type, opening_kg, received_kg, total_kg, sold_kg, closing_kg))
            conn.commit()
            conn.close()
            st.success(f"Entry saved successfully! Sold: {format_kg(sold_kg)}")

# --- DASHBOARD ---
elif choice == "📊 Dashboard":
    st.title("📊 Inventory Dashboard")
    conn = get_db_connection()
    df = pd.read_sql_query("SELECT * FROM inventory ORDER BY date ASC", conn)
    conn.close()
    
    if not df.empty:
        # Latest Stock Cards
        st.subheader("Current Stock Status")
        latest_stocks = df.sort_values('date').groupby('card_type').tail(1)
        
        cols = st.columns(3)
        for i, card in enumerate(["PHH", "AAY", "SFSS"]):
            with cols[i]:
                card_data = latest_stocks[latest_stocks['card_type'] == card]
                current_stock = card_data['closing_kg'].values[0] if not card_data.empty else 0
                st.metric(label=f"{card} Current Stock", value=format_kg(current_stock))
                
        st.divider()
        
        # Charts
        st.subheader("30-Day Stock Movement (Sales)")
        recent_df = df.tail(90) # approx 30 days * 3 card types
        fig = px.bar(recent_df, x="date", y="sold_kg", color="card_type", barmode="group",
                     title="Daily Rice Distribution (Kg)", labels={'sold_kg':'Sold (Kg)', 'date': 'Date'})
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("No data available yet. Please make some daily entries.")

# --- REPORTS & EXPORT ---
elif choice in ["📋 Monthly Reports", "📅 Yearly Summary"]:
    st.title(choice)
    
    conn = get_db_connection()
    df = pd.read_sql_query("SELECT * FROM inventory", conn)
    conn.close()
    
    if df.empty:
        st.warning("No records found in database.")
    else:
        df['date'] = pd.to_datetime(df['date'])
        
        if choice == "📋 Monthly Reports":
            months = df['date'].dt.to_period('M').unique()
            selected_month = st.selectbox("Select Month", sorted(months, reverse=True))
            filtered_df = df[df['date'].dt.to_period('M') == selected_month]
            title_str = f"Month / Year - {selected_month}"
        else:
            years = df['date'].dt.year.unique()
            selected_year = st.selectbox("Select Year", sorted(years, reverse=True))
            filtered_df = df[df['date'].dt.year == selected_year]
            title_str = f"Year - {selected_year}"
            
        st.dataframe(filtered_df.sort_values(by=['date', 'card_type']))
        
        # EXCEL EXPORT (Government Style)
        if st.button("Export to Official Register (Excel)"):
            wb = Workbook()
            ws = wb.active
            ws.title = "Stock Register"
            
            # Styles
            title_font = Font(bold=True, size=14)
            header_font = Font(bold=True)
            center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
            thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            
            # Headers Setup
            ws.merge_cells('A1:H1')
            ws['A1'] = f"FPS Stock Register / F.P.S. Code No. {settings['code']}"
            ws['A1'].font = title_font
            ws['A1'].alignment = center_align
            
            ws.merge_cells('A2:H2')
            ws['A2'] = title_str
            ws['A2'].alignment = Alignment(horizontal="left", vertical="center")
            
            headers = ["Date of Transaction", "Opening Balance (Quintal)", "Receipt (Quintal)", 
                       "Total (Quintal)", "Number of Cards", "Quantity Issued/Sold", 
                       "Closing Balance (Quintal)", "Remarks"]
            
            ws.append(headers)
            for col in range(1, 9):
                cell = ws.cell(row=3, column=col)
                cell.font = header_font
                cell.alignment = center_align
                cell.border = thin_border
                ws.column_dimensions[cell.column_letter].width = 18
                
            # Data Rows (Converted to Quintal format logically for register)
            for _, row in filtered_df.sort_values(by=['date', 'card_type']).iterrows():
                data_row = [
                    row['date'].strftime('%d-%m-%Y') + f" ({row['card_type']})",
                    row['opening_kg'] / 100,
                    row['received_kg'] / 100,
                    row['total_kg'] / 100,
                    "", # Number of cards (manual or add later)
                    row['sold_kg'] / 100,
                    row['closing_kg'] / 100,
                    ""  # Remarks
                ]
                ws.append(data_row)
                
            for row in ws.iter_rows(min_row=4, max_row=ws.max_row, min_col=1, max_col=8):
                for cell in row:
                    cell.border = thin_border
                    cell.alignment = Alignment(horizontal="center")
            
            # Save to BytesIO
            output = io.BytesIO()
            wb.save(output)
            output.seek(0)
            
            st.download_button(label="📥 Download Excel Register",
                               data=output,
                               file_name=f"FPS_Stock_Register_{title_str.replace(' / ', '_')}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# --- SETTINGS ---
elif choice == "⚙️ Settings":
    st.title("⚙️ Settings")
    st.subheader("Update Credentials & Details")
    
    with st.form("settings_form"):
        new_dealer = st.text_input("Dealer Name", value=settings['dealer_name'])
        new_code = st.text_input("FPS Code", value=settings['code'])
        new_user = st.text_input("Username", value=settings['username'])
        new_pass = st.text_input("Password", type="password", value=settings['password'])
        
        if st.form_submit_button("Update Settings"):
            conn = get_db_connection()
            c = conn.cursor()
            c.execute("UPDATE settings SET dealer_name=?, code=?, username=?, password=?", 
                      (new_dealer, new_code, new_user, new_pass))
            conn.commit()
            conn.close()
            st.success("Settings updated successfully! Please log in again to see changes.")
            st.session_state.logged_in = False

# ==========================================
# 7. FOOTER
# ==========================================
st.markdown("""
<div class="footer">
    🌾 Rice Inventory | Created by - Sunil Kumar
</div>
""", unsafe_allow_html=True)
