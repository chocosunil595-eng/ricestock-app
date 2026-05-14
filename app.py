import streamlit as st
import pandas as pd
import sqlite3
from datetime import datetime, date, timedelta
from zoneinfo import ZoneInfo
import io
import plotly.express as px
import plotly.graph_objects as go
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows

# --- PAGE CONFIGURATION ---
st.set_page_config(
    page_title="RiceStock - FPS Inventory System",
    page_icon="🌾",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- THEME & STYLES ---
# Custom CSS for a professional, clean, green-themed look
custom_css = """
<style>
    :root {
        --primary-color: #2e7d32;
        --background-color: #f1f8e9;
        --text-color: #1b5e20;
    }
    
    .stApp {
        background-color: #ffffff;
    }
    
    .stSidebar {
        background-color: #f1f8e9 !important;
    }
    
    .stSidebar [data-testid="stSidebarNav"] span {
        color: #1b5e20;
        font-weight: 500;
    }
    
    h1, h2, h3 {
        color: #2e7d32 !important;
    }
    
    .metric-card {
        background-color: #e8f5e9;
        border-left: 5px solid #2e7d32;
        padding: 15px;
        border-radius: 5px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    
    .footer {
        position: fixed;
        left: 0;
        bottom: 0;
        width: 100%;
        background-color: #f1f8e9;
        color: #1b5e20;
        text-align: center;
        padding: 10px;
        font-size: 0.9em;
        border-top: 1px solid #c8e6c9;
        z-index: 999;
    }
    
    .login-box {
        background-color: #ffffff;
        padding: 30px;
        border-radius: 10px;
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
        border: 1px solid #c8e6c9;
    }
</style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

# --- CONSTANTS & CONFIG ---
IST = ZoneInfo('Asia/Kolkata')
DB_NAME = "rice_inventory.db"

# --- HELPER FUNCTIONS ---

def format_quantity(kg_value):
    """
    Formats the weight from Kg to Quintal and Kg format.
    ≥ 100 Kg → X Q.YY Kg
    < 100 Kg → XX Kg
    """
    if kg_value is None:
        return "0 Kg"
    
    try:
        kg = float(kg_value)
        if kg >= 100:
            quintals = int(kg // 100)
            remaining_kg = kg % 100
            # Format nicely, dropping .0 if it's an exact integer
            if remaining_kg == 0:
                return f"{quintals} Q"
            elif remaining_kg.is_integer():
                return f"{quintals} Q.{int(remaining_kg)} Kg"
            else:
                return f"{quintals} Q.{remaining_kg:.2f} Kg"
        else:
            if kg.is_integer():
                return f"{int(kg)} Kg"
            return f"{kg:.2f} Kg"
    except (ValueError, TypeError):
        return "0 Kg"

def format_quantity_quintal_only(kg_value):
    """Formats strictly in Quintals (decimal format) for official reports."""
    if kg_value is None:
        return "0.00"
    try:
        return f"{float(kg_value) / 100:.2f}"
    except (ValueError, TypeError):
        return "0.00"

def get_greeting():
    """Returns a greeting based on current IST time."""
    current_time = datetime.now(IST)
    hour = current_time.hour
    
    if 5 <= hour < 12:
        return "Good Morning"
    elif 12 <= hour < 17:
        return "Good Afternoon"
    else:
        return "Good Evening"

# --- DATABASE FUNCTIONS ---
def init_db():
    """Initializes the SQLite database and creates necessary tables."""
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS daily_stock (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            entry_date DATE UNIQUE,
            opening_balance REAL,
            received REAL,
            total REAL,
            selling REAL,
            closing_balance REAL,
            created_at TIMESTAMP,
            updated_at TIMESTAMP
        )
    ''')
    conn.commit()
    conn.close()

def get_latest_closing_balance():
    """Fetches the closing balance of the most recent entry."""
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute('''
        SELECT closing_balance 
        FROM daily_stock 
        ORDER BY entry_date DESC 
        LIMIT 1
    ''')
    result = c.fetchone()
    conn.close()
    return result[0] if result else 0.0

def get_entry_by_date(target_date):
    """Fetches an entry for a specific date."""
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute('SELECT * FROM daily_stock WHERE entry_date = ?', (target_date,))
    result = c.fetchone()
    conn.close()
    
    if result:
        # Map to dict for easier access
        columns = ['id', 'entry_date', 'opening_balance', 'received', 'total', 'selling', 'closing_balance', 'created_at', 'updated_at']
        return dict(zip(columns, result))
    return None

def get_all_data(start_date=None, end_date=None):
    """Fetches all data, optionally filtered by date range."""
    conn = sqlite3.connect(DB_NAME)
    query = "SELECT * FROM daily_stock"
    params = []
    
    if start_date and end_date:
        query += " WHERE entry_date BETWEEN ? AND ?"
        params.extend([start_date, end_date])
        
    query += " ORDER BY entry_date DESC"
    
    df = pd.read_sql_query(query, conn, params=params)
    conn.close()
    return df

def save_entry(entry_date, opening, received, total, selling, closing):
    """Saves or updates a daily entry."""
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    
    current_time = datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S")
    
    # Check if entry exists
    c.execute('SELECT id FROM daily_stock WHERE entry_date = ?', (entry_date,))
    exists = c.fetchone()
    
    if exists:
        c.execute('''
            UPDATE daily_stock 
            SET opening_balance=?, received=?, total=?, selling=?, closing_balance=?, updated_at=?
            WHERE entry_date=?
        ''', (opening, received, total, selling, closing, current_time, entry_date))
        st.toast("Entry updated successfully!", icon="✅")
    else:
        c.execute('''
            INSERT INTO daily_stock (entry_date, opening_balance, received, total, selling, closing_balance, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        ''', (entry_date, opening, received, total, selling, closing, current_time, current_time))
        st.toast("New entry saved successfully!", icon="✅")
        
    conn.commit()
    conn.close()

# --- EXPORT FUNCTIONALITY ---
def generate_official_excel(df, period_str):
    """Generates a print-ready Excel file formatted like an official FPS Stock Register."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Stock Register"
    
    # Define styles
    header_font = Font(name='Arial', bold=True, size=12)
    title_font = Font(name='Arial', bold=True, size=14)
    regular_font = Font(name='Arial', size=11)
    
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                         top=Side(style='thin'), bottom=Side(style='thin'))
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    
    # Header Section
    ws.merge_cells('A1:H1')
    ws['A1'] = "F.P.S. STOCK REGISTER"
    ws['A1'].font = title_font
    ws['A1'].alignment = center_align
    
    ws.merge_cells('A2:D2')
    ws['A2'] = f"Dealer Name: Balaram Shial"
    ws['A2'].font = header_font
    
    ws.merge_cells('E2:H2')
    ws['E2'] = f"F.P.S. Code No.: 0201P100"
    ws['E2'].font = header_font
    ws['E2'].alignment = Alignment(horizontal='right')
    
    ws.merge_cells('A3:H3')
    ws['A3'] = f"Period: {period_str}"
    ws['A3'].font = header_font
    ws['A3'].alignment = center_align
    
    # Empty row for spacing
    ws.append([])
    
    # Column Headers (Row 5)
    headers = [
        "Date of Transaction",
        "Opening Balance\n(Quintal)",
        "Receipt\n(Quintal)",
        "Total\n(Quintal)",
        "Number of\nCards",
        "Quantity Issued\n(Quintal)",
        "Closing Balance\n(Quintal)",
        "Remarks"
    ]
    
    ws.append(headers)
    
    # Style Headers
    for col_num, cell in enumerate(ws[5], 1):
        cell.font = header_font
        cell.border = thin_border
        cell.alignment = center_align
        cell.fill = PatternFill(start_color="E8F5E9", end_color="E8F5E9", fill_type="solid")
    
    # Set Column Widths
    col_widths = {'A': 15, 'B': 15, 'C': 15, 'D': 15, 'E': 12, 'F': 15, 'G': 15, 'H': 20}
    for col, width in col_widths.items():
        ws.column_dimensions[col].width = width
    
    ws.row_dimensions[5].height = 45
    
    # Sort dataframe chronological for export
    df_sorted = df.sort_values(by='entry_date', ascending=True)
    
    # Add Data
    for index, row in df_sorted.iterrows():
        date_str = pd.to_datetime(row['entry_date']).strftime("%d/%m/%Y")
        
        # Convert Kg to Quintals for official report
        op_bal = float(row['opening_balance']) / 100
        rect = float(row['received']) / 100
        tot = float(row['total']) / 100
        issued = float(row['selling']) / 100
        cl_bal = float(row['closing_balance']) / 100
        
        data_row = [
            date_str,
            f"{op_bal:.2f}",
            f"{rect:.2f}",
            f"{tot:.2f}",
            "", # Number of Cards (Blank/Optional)
            f"{issued:.2f}",
            f"{cl_bal:.2f}",
            ""  # Remarks
        ]
        ws.append(data_row)
        
        # Style Data Rows
        current_row = ws.max_row
        for cell in ws[current_row]:
            cell.font = regular_font
            cell.border = thin_border
            cell.alignment = center_align
            
    # Save to memory
    excel_file = io.BytesIO()
    wb.save(excel_file)
    excel_file.seek(0)
    
    return excel_file

# --- UI COMPONENTS ---

def login_page():
    """Renders the login interface."""
    st.markdown("<h1 style='text-align: center; color: #2e7d32; font-size: 3em;'>🌾 RiceStock</h1>", unsafe_allow_html=True)
    st.markdown("<h3 style='text-align: center; color: #555;'>Official FPS Inventory System</h3>", unsafe_allow_html=True)
    st.write("---")
    
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        st.markdown("<div class='login-box'>", unsafe_allow_html=True)
        st.markdown("<h2 style='text-align: center; margin-bottom: 20px;'>Dealer Login</h2>", unsafe_allow_html=True)
        st.markdown("**Dealer:** Balaram Shial  <br>  **Code:** 0201P100", unsafe_allow_html=True)
        
        with st.form("login_form"):
            username = st.text_input("Username")
            password = st.text_input("Password", type="password")
            submit = st.form_submit_button("Login", use_container_width=True)
            
            if submit:
                if username == "Sunil" and password == "sunil123":
                    st.session_state['logged_in'] = True
                    st.session_state['username'] = username
                    st.rerun()
                else:
                    st.error("Invalid Username or Password")
        st.markdown("</div>", unsafe_allow_html=True)

def home_page():
    """Renders the Home/Dashboard view."""
    greeting = get_greeting()
    st.title(f"{greeting}, Mr. Sunil! 👋")
    st.markdown("##### Welcome to your Rice Inventory Management System")
    st.write("---")
    
    # Summary Metrics
    today_str = datetime.now(IST).strftime("%Y-%m-%d")
    today_data = get_entry_by_date(today_str)
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.markdown("<div class='metric-card'>", unsafe_allow_html=True)
        st.markdown("**Today's Opening**")
        val = today_data['opening_balance'] if today_data else get_latest_closing_balance()
        st.markdown(f"<h3 style='color:#1b5e20;'>{format_quantity(val)}</h3>", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)
        
    with col2:
        st.markdown("<div class='metric-card'>", unsafe_allow_html=True)
        st.markdown("**Received Today**")
        val = today_data['received'] if today_data else 0.0
        st.markdown(f"<h3 style='color:#1976d2;'>{format_quantity(val)}</h3>", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)
        
    with col3:
        st.markdown("<div class='metric-card'>", unsafe_allow_html=True)
        st.markdown("**Sold Today**")
        val = today_data['selling'] if today_data else 0.0
        st.markdown(f"<h3 style='color:#d32f2f;'>{format_quantity(val)}</h3>", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)
        
    with col4:
        st.markdown("<div class='metric-card'>", unsafe_allow_html=True)
        st.markdown("**Current Stock**")
        val = today_data['closing_balance'] if today_data else get_latest_closing_balance()
        st.markdown(f"<h3 style='color:#fbc02d;'>{format_quantity(val)}</h3>", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

    st.write("<br>", unsafe_allow_html=True)
    
    # Recent Transactions Chart
    st.subheader("Recent Activity (Last 7 Entries)")
    df = get_all_data()
    if not df.empty:
        recent_df = df.head(7).sort_values('entry_date')
        
        fig = go.Figure()
        fig.add_trace(go.Bar(x=recent_df['entry_date'], y=recent_df['received'], name='Received (Kg)', marker_color='#1976d2'))
        fig.add_trace(go.Bar(x=recent_df['entry_date'], y=recent_df['selling'], name='Sold (Kg)', marker_color='#d32f2f'))
        fig.add_trace(go.Scatter(x=recent_df['entry_date'], y=recent_df['closing_balance'], name='Closing Stock (Kg)', mode='lines+markers', line=dict(color='#2e7d32', width=3)))
        
        fig.update_layout(
            barmode='group',
            plot_bgcolor='white',
            paper_bgcolor='white',
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            margin=dict(l=20, r=20, t=40, b=20)
        )
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("No data available yet. Start by adding a Daily Entry.")

def daily_entry_page():
    """Renders the Daily Entry form with smart calculations."""
    st.title("📝 Daily Entry")
    st.markdown("Enter stock details below. All internal values are stored in **Kg**.")
    
    # Date Selection
    entry_date = st.date_input("Select Date", datetime.now(IST).date(), max_value=datetime.now(IST).date())
    date_str = entry_date.strftime("%Y-%m-%d")
    
    # Fetch existing data for selected date
    existing_data = get_entry_by_date(date_str)
    
    # Determine Opening Balance
    if existing_data:
        default_opening = existing_data['opening_balance']
        is_edit = True
        st.info(f"Editing existing record for {date_str}")
    else:
        # If no entry for this date, fetch latest closing balance BEFORE this date
        conn = sqlite3.connect(DB_NAME)
        df_prev = pd.read_sql_query("SELECT closing_balance FROM daily_stock WHERE entry_date < ? ORDER BY entry_date DESC LIMIT 1", conn, params=(date_str,))
        conn.close()
        default_opening = df_prev.iloc[0]['closing_balance'] if not df_prev.empty else 0.0
        is_edit = False
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("### Input Values (in Kg)")
        # Inputs
        opening_input = st.number_input("Opening Balance (Kg)", min_value=0.0, value=float(default_opening), step=1.0, disabled=True, help="Auto-fetched from previous closing balance.")
        
        default_rec = float(existing_data['received']) if existing_data else 0.0
        received_input = st.number_input("Received Today (Kg)", min_value=0.0, value=default_rec, step=1.0)
        
        default_closing = float(existing_data['closing_balance']) if existing_data else float(default_opening)
        closing_input = st.number_input("Closing Balance (Kg)", min_value=0.0, value=default_closing, step=1.0)

    # --- SMART LOGIC CALCULATIONS ---
    # Total = Opening + Received
    calculated_total = opening_input + received_input
    
    # Selling = Total - Closing
    calculated_selling = calculated_total - closing_input
    
    # Edge case correction: If Closing > Total, automatically adjust Received and Selling
    auto_adjusted = False
    if closing_input > calculated_total:
        auto_adjusted = True
        # Adjust Received so that Total == Closing (meaning selling = 0)
        # Closing = Opening + Adjusted_Received
        adjusted_received = closing_input - opening_input
        calculated_total = opening_input + adjusted_received
        calculated_selling = 0.0
        
        st.warning(f"⚠️ Closing Balance ({format_quantity(closing_input)}) is greater than Total Stock. System has automatically adjusted 'Received' to {format_quantity(adjusted_received)} and 'Selling' to 0 Kg to balance the books.")
        # Force the variable to update for display and saving
        received_input = adjusted_received

    with col2:
        st.markdown("### Calculated Preview")
        st.markdown("<div style='background-color:#f8f9fa; padding:20px; border-radius:10px; border:1px solid #dee2e6;'>", unsafe_allow_html=True)
        
        st.markdown(f"**Opening Balance:** {format_quantity(opening_input)}")
        st.markdown(f"**+ Received:** {format_quantity(received_input)}")
        st.markdown(f"**= Total Stock:** {format_quantity(calculated_total)}")
        st.markdown(f"**- Sold (Issued):** <span style='color:#d32f2f; font-weight:bold;'>{format_quantity(calculated_selling)}</span>", unsafe_allow_html=True)
        st.markdown(f"**= Closing Balance:** <span style='color:#2e7d32; font-weight:bold;'>{format_quantity(closing_input)}</span>", unsafe_allow_html=True)
        
        st.markdown("</div>", unsafe_allow_html=True)
        
        if calculated_selling < 0 and not auto_adjusted:
            st.error("Error: Negative selling amount detected. Please check your inputs.")
            can_save = False
        else:
            can_save = True

    st.write("---")
    if st.button("💾 Save Entry", disabled=not can_save, type="primary"):
        with st.spinner("Saving data..."):
            save_entry(
                entry_date=date_str,
                opening=opening_input,
                received=received_input,
                total=calculated_total,
                selling=calculated_selling,
                closing=closing_input
            )
            # Fetch latest to refresh cache implicitly
            get_latest_closing_balance()
            st.success("Entry saved successfully!")
            # Use query params or rerun to refresh view
            st.rerun()

def view_reports_page():
    """Renders the data table and export functionality."""
    st.title("📊 Reports & Export")
    
    tab1, tab2 = st.tabs(["Monthly Report", "Yearly Report"])
    
    df = get_all_data()
    
    if df.empty:
        st.warning("No data available in the database.")
        return
        
    df['entry_date'] = pd.to_datetime(df['entry_date'])
    
    with tab1:
        st.subheader("Monthly Register")
        
        col1, col2 = st.columns(2)
        with col1:
            years = df['entry_date'].dt.year.unique()
            selected_year = st.selectbox("Select Year", years, key="month_year")
        with col2:
            months = range(1, 13)
            month_names = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
            selected_month_idx = st.selectbox("Select Month", range(len(month_names)), format_func=lambda x: month_names[x])
            selected_month = selected_month_idx + 1
            
        # Filter Data
        mask = (df['entry_date'].dt.year == selected_year) & (df['entry_date'].dt.month == selected_month)
        monthly_df = df[mask].copy()
        
        if not monthly_df.empty:
            st.write(f"Showing records for **{month_names[selected_month_idx]} {selected_year}**")
            
            # Display dataframe with formatted quantities
            display_df = monthly_df.copy()
            display_df['entry_date'] = display_df['entry_date'].dt.strftime('%d-%m-%Y')
            cols_to_format = ['opening_balance', 'received', 'total', 'selling', 'closing_balance']
            for col in cols_to_format:
                display_df[col] = display_df[col].apply(format_quantity)
                
            st.dataframe(display_df[['entry_date', 'opening_balance', 'received', 'total', 'selling', 'closing_balance']], use_container_width=True, hide_index=True)
            
            # Export Button
            period_str = f"{month_names[selected_month_idx]} / {selected_year}"
            excel_data = generate_official_excel(monthly_df, period_str)
            
            st.download_button(
                label="📥 Download Official Register (Excel)",
                data=excel_data,
                file_name=f"FPS_Register_{month_names[selected_month_idx]}_{selected_year}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
        else:
            st.info("No records found for the selected month.")

    with tab2:
        st.subheader("Yearly Summary")
        selected_year_only = st.selectbox("Select Year", years, key="year_only")
        
        yearly_df = df[df['entry_date'].dt.year == selected_year_only].copy()
        
        if not yearly_df.empty:
            # Aggregate monthly totals for the year
            yearly_df['Month'] = yearly_df['entry_date'].dt.strftime('%B')
            monthly_summary = yearly_df.groupby('Month').agg({
                'received': 'sum',
                'selling': 'sum'
            }).reset_index()
            
            # Sort months correctly
            month_dict = {m: i for i, m in enumerate(month_names)}
            monthly_summary['month_num'] = monthly_summary['Month'].map(month_dict)
            monthly_summary = monthly_summary.sort_values('month_num').drop('month_num', axis=1)
            
            # Format for display
            display_summary = monthly_summary.copy()
            display_summary['received'] = display_summary['received'].apply(format_quantity)
            display_summary['selling'] = display_summary['selling'].apply(format_quantity)
            
            st.dataframe(display_summary, use_container_width=True, hide_index=True)
            
            # Export
            period_str = f"Year {selected_year_only}"
            excel_data = generate_official_excel(yearly_df, period_str)
            
            st.download_button(
                label=f"📥 Download Yearly Data ({selected_year_only})",
                data=excel_data,
                file_name=f"FPS_Register_Year_{selected_year_only}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
        else:
             st.info("No records found for the selected year.")

# --- MAIN APP ROUTING ---
def main():
    """Main application entry point and routing."""
    init_db()
    
    if 'logged_in' not in st.session_state:
        st.session_state['logged_in'] = False
        
    if not st.session_state['logged_in']:
        login_page()
    else:
        # Sidebar Navigation
        with st.sidebar:
            st.markdown("<h2 style='text-align: center;'>🌾 RiceStock</h2>", unsafe_allow_html=True)
            st.markdown("<p style='text-align: center; color: #555; font-size: 0.9em;'>FPS Code: 0201P100</p>", unsafe_allow_html=True)
            st.write("---")
            
            page = st.radio("Navigation", 
                ["🏠 Home", "📝 Daily Entry", "📊 Reports & Export"],
                label_visibility="collapsed"
            )
            
            st.write("---")
            if st.button("🚪 Logout", use_container_width=True):
                st.session_state['logged_in'] = False
                st.rerun()
                
        # Main Content Area
        if page == "🏠 Home":
            home_page()
        elif page == "📝 Daily Entry":
            daily_entry_page()
        elif page == "📊 Reports & Export":
            view_reports_page()
            
        # Footer
        st.markdown(
            """
            <div class="footer">
                <b>Rice Inventory System</b> | Created by - Sunil Kumar
            </div>
            """,
            unsafe_allow_html=True
        )

if __name__ == "__main__":
    main()
