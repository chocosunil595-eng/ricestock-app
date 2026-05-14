import streamlit as st
import sqlite3
import pandas as pd
import datetime
import pytz
import plotly.express as px
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side

# ==========================================
# 1. PAGE CONFIGURATION & CSS THEME
# ==========================================
st.set_page_config(page_title="RiceStock - FPS Inventory System", layout="wide")

# Custom Green Agricultural Theme & Footer
st.markdown("""
    <style>
    .stApp { background-color: #f7fdf7; }
    h1, h2, h3 { color: #2e7d32; }
    .stButton>button { background-color: #4caf50; color: white; border-radius: 5px; }
    .stButton>button:hover { background-color: #388e3c; }
    .footer { position: fixed; bottom: 0; left: 0; width: 100%; background-color: #2e7d32; color: white; text-align: center; padding: 10px; z-index: 1000; }
    </style>
    <div class="footer">
        <strong>Rice Inventory</strong> | Created by - Sunil Kumar
    </div>
""", unsafe_allow_html=True)

# ==========================================
# 2. DATABASE INITIALIZATION
# ==========================================
DB_FILE = 'rice_inventory.db'

def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    # Inventory Table
    c.execute('''CREATE TABLE IF NOT EXISTS inventory (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    entry_date TEXT, entry_month TEXT, card_type TEXT,
                    opening REAL, received REAL, total REAL, 
                    closing REAL, selling REAL, remarks TEXT
                )''')
    # Settings Table
    c.execute('''CREATE TABLE IF NOT EXISTS settings (
                    id INTEGER PRIMARY KEY CHECK (id = 1),
                    username TEXT, password TEXT, timezone TEXT
                )''')
    # Insert default settings if empty
    c.execute("SELECT * FROM settings WHERE id = 1")
    if not c.fetchone():
        c.execute("INSERT INTO settings (id, username, password, timezone) VALUES (1, 'Sunil', 'sunil123', 'Asia/Kolkata')")
    conn.commit()
    conn.close()

init_db()

# ==========================================
# 3. GLOBAL HELPERS & LOGIC
# ==========================================
def get_settings():
    conn = sqlite3.connect(DB_FILE)
    df = pd.read_sql_query("SELECT * FROM settings WHERE id=1", conn)
    conn.close()
    return df.iloc[0]

def get_ist_time():
    tz_string = get_settings()['timezone']
    return datetime.datetime.now(pytz.timezone(tz_string))

def format_qty(kg):
    if kg >= 100:
        q = int(kg // 100)
        rem = kg % 100
        return f"{q} Q.{rem:02g} Kg" if rem > 0 else f"{q} Q.00 Kg"
    return f"{kg:g} Kg"

def kg_to_q(kg):
    return round(kg / 100, 2)

# ==========================================
# 4. AUTHENTICATION (LOGIN)
# ==========================================
if 'logged_in' not in st.session_state:
    st.session_state['logged_in'] = False

settings = get_settings()

if not st.session_state['logged_in']:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("<h2 style='text-align: center;'>Dealer - Balaram Shial</h2>", unsafe_allow_html=True)
        st.markdown("<h4 style='text-align: center; color: gray;'>Code - 0201P100</h4>", unsafe_allow_html=True)
        st.write("---")
        
        user_input = st.text_input("Username")
        pass_input = st.text_input("Password", type="password")
        
        if st.button("Login", use_container_width=True):
            if user_input == settings['username'] and pass_input == settings['password']:
                st.session_state['logged_in'] = True
                st.rerun()
            else:
                st.error("Invalid Credentials. Please try again.")
    st.stop()

# ==========================================
# 5. DYNAMIC GREETING & SIDEBAR NAVIGATION
# ==========================================
current_time = get_ist_time()
hour = current_time.hour
if hour < 12: greeting = "Good Morning"
elif hour < 17: greeting = "Good Afternoon"
else: greeting = "Good Evening"

st.sidebar.markdown(f"### {greeting},<br>Mr. {settings['username']}", unsafe_allow_html=True)
st.sidebar.write("---")

menu = st.sidebar.radio("Navigation", ["🏠 Home", "📝 Daily Entry", "📊 Dashboard", "📋 Monthly Reports", "📅 Yearly Summary", "⬆️ Import / Export", "⚙️ Settings"])

if st.sidebar.button("Logout"):
    st.session_state['logged_in'] = False
    st.rerun()

# ==========================================
# 6. APP MODULES
# ==========================================
def get_previous_closing(card_type):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT closing FROM inventory WHERE card_type=? ORDER BY entry_date DESC LIMIT 1", (card_type,))
    result = c.fetchone()
    conn.close()
    return result[0] if result else 0.0

if menu == "🏠 Home" or menu == "📊 Dashboard":
    st.title("📊 Inventory Dashboard")
    st.write(f"**Welcome back, {settings['username']}!** Here is your current stock overview.")
    
    conn = sqlite3.connect(DB_FILE)
    df = pd.read_sql_query("SELECT * FROM inventory", conn)
    conn.close()

    col1, col2, col3 = st.columns(3)
    card_types = ['PHH', 'AAY', 'SFSS']
    
    for i, c_type in enumerate(card_types):
        latest = df[df['card_type'] == c_type].sort_values(by='entry_date', ascending=False).head(1)
        current_stock = latest['closing'].values[0] if not latest.empty else 0
        with [col1, col2, col3][i]:
            st.metric(label=f"Current {c_type} Stock", value=format_qty(current_stock))
            st.progress(min(int(current_stock / 10000 * 100), 100)) # Progress visual up to 10k kg

    st.write("---")
    st.subheader("📈 30-Day Stock Movement")
    if not df.empty:
        df['entry_date'] = pd.to_datetime(df['entry_date'])
        recent_df = df[df['entry_date'] >= (current_time.replace(tzinfo=None) - datetime.timedelta(days=30))]
        if not recent_df.empty:
            fig = px.line(recent_df, x='entry_date', y='closing', color='card_type', title="Closing Stock Over Time (Kg)")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No transaction data available for the last 30 days.")
    else:
        st.info("No data available yet. Please add entries in the 'Daily Entry' tab.")

elif menu == "📝 Daily Entry":
    st.title("📝 Daily Stock Entry")
    
    col1, col2 = st.columns(2)
    with col1:
        entry_date = st.date_input("Date", value=current_time.date())
        card_type = st.selectbox("Card Type", ["PHH", "AAY", "SFSS"])
    with col2:
        months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
        entry_month = st.selectbox("Allocation Month", months, index=current_time.month - 1)
        opening_bal = get_previous_closing(card_type)
        st.text_input("Opening Balance (Auto-fetched)", value=format_qty(opening_bal), disabled=True)

    col3, col4 = st.columns(2)
    with col3:
        received = st.number_input("Received (in Kg)", min_value=0.0, step=1.0)
    with col4:
        closing_bal = st.number_input("Closing Balance (in Kg)", min_value=0.0, step=1.0)

    remarks = st.text_input("Remarks (Optional)")

    if st.button("Save Entry", use_container_width=True):
        total = opening_bal + received
        
        # Special Auto-Adjustment Logic
        if closing_bal > total:
            excess = closing_bal - total
            received = received + excess
            total = opening_bal + received
            st.warning(f"⚠️ Closing was higher than Total. Received has been automatically adjusted to {received} Kg.")

        selling = total - closing_bal
        
        conn = sqlite3.connect(DB_FILE)
        c = conn.cursor()
        c.execute("""INSERT INTO inventory (entry_date, entry_month, card_type, opening, received, total, closing, selling, remarks) 
                     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""", 
                  (str(entry_date), entry_month, card_type, opening_bal, received, total, closing_bal, selling, remarks))
        conn.commit()
        conn.close()
        st.success("✅ Daily entry saved successfully!")

elif menu in ["📋 Monthly Reports", "📅 Yearly Summary", "⬆️ Import / Export"]:
    st.title("📋 Government Official Exports")
    st.info("Generate print-ready Excel sheets formatted exactly to Government FPS Stock Register standards (in Quintals).")
    
    conn = sqlite3.connect(DB_FILE)
    df = pd.read_sql_query("SELECT * FROM inventory", conn)
    conn.close()

    if df.empty:
        st.warning("No data available to export.")
    else:
        months = df['entry_month'].unique().tolist()
        selected_months = st.multiselect("Select Month(s) for Export", months, default=months[-1] if months else None)
        
        if st.button("Generate FPS Excel Report") and selected_months:
            filtered_df = df[df['entry_month'].isin(selected_months)].copy()
            filtered_df = filtered_df.sort_values(by='entry_date')

            # Create Excel File in Memory
            wb = Workbook()
            ws = wb.active
            
            # Formatting
            bold_font = Font(bold=True)
            center_aligned_text = Alignment(horizontal="center", vertical="center")
            thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

            # Headers
            title_month = f"Month / Year - {selected_months[0]}" if len(selected_months) == 1 else f"Combined Report - {', '.join(selected_months)}"
            ws.merge_cells('A1:H1')
            ws['A1'] = "FPS Stock Register / F.P.S. Code No. 0201P100"
            ws['A1'].font = bold_font
            ws['A1'].alignment = center_aligned_text
            
            ws.merge_cells('A2:H2')
            ws['A2'] = title_month
            ws['A2'].font = bold_font
            ws['A2'].alignment = center_aligned_text

            columns = ["Date of Transaction", "Opening Balance (Q)", "Receipt (Q)", "Total (Q)", "Number of Cards", "Quantity Issued (Q)", "Closing Balance (Q)", "Remarks"]
            ws.append(columns)
            
            for cell in ws[3]:
                cell.font = bold_font
                cell.alignment = center_aligned_text
                cell.border = thin_border

            # Data Rows (Converted to Quintal)
            for _, row in filtered_df.iterrows():
                data_row = [
                    row['entry_date'],
                    kg_to_q(row['opening']),
                    kg_to_q(row['received']),
                    kg_to_q(row['total']),
                    "", # Blank optional column
                    kg_to_q(row['selling']),
                    kg_to_q(row['closing']),
                    row['remarks'] if row['remarks'] else ""
                ]
                ws.append(data_row)
                for cell in ws[ws.max_row]:
                    cell.border = thin_border
                    cell.alignment = Alignment(horizontal="center")

            # Save to BytesIO
            excel_data = io.BytesIO()
            wb.save(excel_data)
            excel_data.seek(0)

            st.download_button(
                label="📥 Download Official Excel Register",
                data=excel_data,
                file_name=f"FPS_Register_{'_'.join(selected_months)}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

elif menu == "⚙️ Settings":
    st.title("⚙️ System Settings")
    
    with st.form("settings_form"):
        new_username = st.text_input("Username", value=settings['username'])
        new_password = st.text_input("New Password", type="password", value=settings['password'])
        timezones = pytz.all_timezones
        tz_index = timezones.index(settings['timezone']) if settings['timezone'] in timezones else timezones.index('Asia/Kolkata')
        new_tz = st.selectbox("System Timezone", timezones, index=tz_index)
        
        if st.form_submit_button("Save Settings"):
            conn = sqlite3.connect(DB_FILE)
            c = conn.cursor()
            c.execute("UPDATE settings SET username=?, password=?, timezone=? WHERE id=1", (new_username, new_password, new_tz))
            conn.commit()
            conn.close()
            st.success("Settings updated successfully! Changes will reflect on next interaction.")