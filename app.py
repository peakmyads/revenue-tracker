"""
Revenue Tracker Software
Author: Sumit
Description:
Revenue, Cost, Billing & Collection Tracker
"""

import streamlit as st
import pandas as pd
import os
from datetime import datetime

FILE_PATH = r"D:\Sumit\PY\Tracker Software\master_config.xlsx"

from st_aggrid import AgGrid, GridOptionsBuilder, JsCode
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Table
from reportlab.lib import utils
from st_aggrid import JsCode

# -------------------------------
# GOOGLE SHEETS CONNECTION (GLOBAL)
# -------------------------------

import gspread
from google.oauth2.service_account import Credentials
import streamlit as st

SPREADSHEET_ID = "18FfRWCMShQSlDaC7UG0N3H00VuN-UbU3rfMTSegcluU"

@st.cache_resource
def get_gsheet_connection():
    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]

    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=scope
    )

    client = gspread.authorize(creds)
    return client

# Create connection ONCE
client = get_gsheet_connection()
spreadsheet = client.open_by_key(SPREADSHEET_ID)

@st.cache_data(ttl=60)
def calculate_kpis(df_master):

    df = df_master.copy()

    df["DSP $ (BC)"] = pd.to_numeric(df["DSP $ (BC)"], errors="coerce").fillna(0)
    df["SSP $ (BC)"] = pd.to_numeric(df["SSP $ (BC)"], errors="coerce").fillna(0)
    df["C DSP $"] = pd.to_numeric(df.get("C DSP $", 0), errors="coerce").fillna(0)
    df["C SSP $"] = pd.to_numeric(df.get("C SSP $", 0), errors="coerce").fillna(0)

    df["Net $ (BC)"] = df["DSP $ (BC)"] - df["SSP $ (BC)"]
    df["C Net $"] = df["C DSP $"] - df["C SSP $"]

    total_dsp = df["DSP $ (BC)"].sum()
    total_ssp = df["SSP $ (BC)"].sum()
    total_net = df["Net $ (BC)"].sum()

    total_c_dsp = df["C DSP $"].sum()
    total_c_ssp = df["C SSP $"].sum()
    total_c_net = df["C Net $"].sum()

    ivt = total_net - total_c_net
    ivt_percent = (ivt / total_dsp * 100) if total_dsp != 0 else 0
    c_profit_percent = (total_c_net / total_c_dsp * 100) if total_c_dsp != 0 else 0

    return (
        df,
        total_dsp,
        total_ssp,
        total_net,
        total_c_dsp,
        total_c_ssp,
        total_c_net,
        ivt,
        ivt_percent,
        c_profit_percent
    )

@st.cache_data(ttl=60)
def load_master_data_from_gsheet():
    worksheet = spreadsheet.worksheet("Master Data")
    data = worksheet.get_all_records()
    return pd.DataFrame(data)
    
@st.cache_data(ttl=60)
def load_partner_list_from_gsheet():
    worksheet = spreadsheet.worksheet("Partner List")
    data = worksheet.get_all_records()
    return pd.DataFrame(data)

def generate_dashboard_pdf(file_path, metrics_dict):
    doc = SimpleDocTemplate(file_path)
    elements = []

    styles = getSampleStyleSheet()
    elements.append(Paragraph("Revenue Dashboard Report", styles["Heading1"]))
    elements.append(Spacer(1, 0.5 * inch))

    data = [["Metric", "Value"]]

    for key, value in metrics_dict.items():
        data.append([key, value])

    table = Table(data)
    elements.append(table)

    doc.build(elements)


import numpy as np
import re
from datetime import datetime

def prepare_dataframe_for_gsheet(df: pd.DataFrame):
    clean_df = df.copy()
    date_columns = []

    for col in clean_df.columns:

        # 🔹 Try detect date columns by name
        if "date" in col.lower():
            parsed_dates = pd.to_datetime(clean_df[col], errors="coerce", dayfirst=True)

            if parsed_dates.notna().mean() > 0.5:
                clean_df[col] = parsed_dates
                date_columns.append(col)
                continue

        # 🔹 Try numeric detection
        numeric_series = pd.to_numeric(clean_df[col], errors="coerce")
        numeric_ratio = numeric_series.notna().mean()

        if numeric_ratio > 0.7:
            clean_df[col] = numeric_series
            clean_df[col].replace([np.inf, -np.inf], 0, inplace=True)
            clean_df[col].fillna(0, inplace=True)
            clean_df[col] = clean_df[col].round(2)
        else:
            clean_df[col] = clean_df[col].astype(str)
            clean_df[col].replace("nan", "", inplace=True)

    return clean_df, date_columns

# -------------------------------
# Utility Functions
# -------------------------------

def load_sheet(sheet_name: str) -> pd.DataFrame:
    if not os.path.exists(FILE_PATH):
        return pd.DataFrame()

    try:
        return pd.read_excel(FILE_PATH, sheet_name=sheet_name)
    except:
        return pd.DataFrame()


def save_sheet(df: pd.DataFrame, sheet_name: str):
    with pd.ExcelWriter(FILE_PATH, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)


def format_usd(value):
    try:
        return f"${value:,.2f}"
    except:
        return "$0.00"


# -------------------------------
# Streamlit App Config
# -------------------------------

st.set_page_config(
    page_title="PEAKADS LLP - Revenue Tracker",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
/* Remove top white padding */
.block-container {
    padding-top: 1rem !important;
    padding-bottom: 0rem !important;
    padding-left: 2rem !important;
    padding-right: 2rem !important;
}

/* Remove extra header spacing */
header {visibility: hidden;}

/* Remove footer */
footer {visibility: hidden;}

/* Remove default Streamlit top space */
[data-testid="stAppViewContainer"] {
    margin-top: -60px;
}

/* Make app use full height */
html, body, [class*="css"] {
    height: 100%;
}
</style>
""", unsafe_allow_html=True)

from datetime import date

def generate_financial_years():
    today = date.today()

    current_year = today.year
    current_month = today.month

    # FY starts April
    if current_month >= 4:
        fy_start = current_year
    else:
        fy_start = current_year - 1

    years = []

    # Always show 3 FY
    for i in range(3):
        start = fy_start - 2 + i
        years.append(f"{start}-{str(start+1)[-2:]}")

    # If today >= 1 April next FY start, add one more
    if today >= date(fy_start + 1, 4, 1):
        extra = fy_start + 1
        years.append(f"{extra}-{str(extra+1)[-2:]}")

    return sorted(years)


def get_fy_date_range(fy_string):
    start_year = int(fy_string.split("-")[0])
    start_date = pd.to_datetime(f"{start_year}-04-01")
    end_date = pd.to_datetime(f"{start_year+1}-03-31")
    return start_date, end_date


def get_quarter_range(fy_string, quarter):
    start_year = int(fy_string.split("-")[0])

    mapping = {
        "Q1": (4, 6),
        "Q2": (7, 9),
        "Q3": (10, 12),
        "Q4": (1, 3)
    }

    start_month, end_month = mapping[quarter]

    if quarter == "Q4":
        start = pd.Timestamp(start_year + 1, start_month, 1)
        end = pd.Timestamp(start_year + 1, end_month, 1) + pd.offsets.MonthEnd(0)
    else:
        start = pd.Timestamp(start_year, start_month, 1)
        end = pd.Timestamp(start_year, end_month, 1) + pd.offsets.MonthEnd(0)

    return start, end

st.markdown("""
<style>

/* KPI Main Container */
.kpi-container {
    background: linear-gradient(135deg, #f4f7fb, #e8eef7);
    padding: 25px;
    border-radius: 12px;
    margin-bottom: 25px;
    box-shadow: 0px 4px 10px rgba(0,0,0,0.08);
}

/* Individual KPI Card */
.kpi-card {
    background-color: white;
    padding: 18px;
    border-radius: 10px;
    box-shadow: 0px 2px 6px rgba(0,0,0,0.05);
    border-left: 6px solid #ccc;
    transition: transform 0.2s ease-in-out;
}

.kpi-card:hover {
    transform: translateY(-3px);
}

/* Positive KPI */
.kpi-positive {
    border-left: 6px solid #2ecc71 !important;
}

/* Negative KPI */
.kpi-negative {
    border-left: 6px solid #e74c3c !important;
}

/* KPI Label */
.kpi-label {
    font-size: 14px;
    font-weight: 600;
    color: #666;
}

/* KPI Value */
.kpi-value {
    font-size: 26px;
    font-weight: 800;
}

</style>
""", unsafe_allow_html=True)

st.markdown("""
<style>

/* KPI Background Container */
.kpi-container {
    background: linear-gradient(135deg, #87CEFA, #e8eef7);
    padding: 25px;
    border-radius: 12px;
    margin-bottom: 20px;
    box-shadow: 0px 4px 10px rgba(0,0,0,0.08);
}

/* Make metrics inside look clean */
div[data-testid="stMetric"] {
    background-color: #00FFFF;
    padding: 15px;
    border-radius: 10px;
    box-shadow: 0px 2px 6px rgba(0,0,0,0.05);
}

</style>
""", unsafe_allow_html=True)

st.markdown("""
<style>

/* ---------- TAB CONTAINER ---------- */
div[data-testid="stTabs"] {
    overflow-x: auto;
    white-space: nowrap;
}

/* ---------- FORCE BOLD TAB TEXT ---------- */
div[data-testid="stTabs"] button {
    font-weight: 800 !important;
    font-size: 17px !important;
    letter-spacing: 0.3px;
}

/* Ensure inner text is bold */
div[data-testid="stTabs"] button p {
    font-weight: 800 !important;
}

/* ---------- TAB BUTTON ---------- */
div[data-testid="stTabs"] button {
    font-weight: 700 !important;
    font-size: 16px !important;
    padding: 10px 20px !important;
    border-radius: 8px 8px 0px 0px !important;
    margin-right: 4px;
    background: linear-gradient(#26F7FD, #f0f0f0, #d6d6d6);
    border: 1px solid #c0c0c0 !important;
    box-shadow: 3px 3px 6px #b0b0b0, 
                -2px -2px 5px #ffffff;
    transition: all 0.25s ease-in-out;
    position: relative;
}

/* ---------- HOVER EFFECT ---------- */
div[data-testid="stTabs"] button:hover {
    transform: translateY(-2px);
    background: linear-gradient(145deg, #e6e6e6, #cccccc);
}

/* ---------- ACTIVE TAB ---------- */
div[data-testid="stTabs"] button[aria-selected="true"] {
    background: linear-gradient(145deg, #003366, #0059b3);
    color: white !important;
    box-shadow: inset 2px 2px 6px #002244,
                inset -2px -2px 6px #0066cc;
}

/* ---------- SUBTLE ANIMATED BOTTOM BORDER ---------- */
div[data-testid="stTabs"] button::after {
    content: "";
    position: absolute;
    left: 0;
    bottom: -3px;
    width: 0%;
    height: 3px;
    background-color: #ff4b4b;
    transition: width 0.3s ease-in-out;
}

div[data-testid="stTabs"] button[aria-selected="true"]::after {
    width: 100%;
}

/* ---------- RESPONSIVE MOBILE ---------- */
@media (max-width: 768px) {

    div[data-testid="stTabs"] button {
        font-size: 14px !important;
        padding: 8px 14px !important;
        margin-right: 2px;
    }

    div[data-testid="stTabs"] {
        overflow-x: auto;
        scrollbar-width: thin;
    }
}

</style>
""", unsafe_allow_html=True)

st.markdown("""
<style>

/* Target Streamlit Tabs */
div[data-testid="stTabs"] button {
    font-weight: 700 !important;
    font-size: 16px !important;
    padding: 10px 18px !important;
    border-radius: 8px 8px 0px 0px !important;
    background: linear-gradient(145deg, #f0f0f0, #87CEFA);
    border: 1px solid #c0c0c0 !important;
    box-shadow: 3px 3px 6px #b0b0b0, 
                -2px -2px 5px #ffffff;
    transition: all 0.2s ease-in-out;
}

/* Hover effect */
div[data-testid="stTabs"] button:hover {
    background: linear-gradient(145deg, #e6e6e6, #90EE90);
    transform: translateY(-2px);
}

/* Active Tab */
div[data-testid="stTabs"] button[aria-selected="true"] {
    background: linear-gradient(145deg, #003366, #0059b3);
    color: white !important;
    box-shadow: inset 2px 2px 6px #002244,
                inset -2px -2px 6px #0066cc;
}

</style>
""", unsafe_allow_html=True)

# -------------------------------
# DASHBOARD-STYLE COMPANY HEADER
# -------------------------------

logo_path = "peakads_logo.png"

st.markdown("""
<style>

/* Remove top spacing */
.block-container {
    padding-top: 0.2rem !important;
}

/* Company header row */
.company-header {
    display: flex;
    align-items: center;
    gap: 12px;
    font-size: 60px;
    font-weight: 700;
    color: #2D5DA1;
    margin-bottom: 5px;
}

/* Subtitle inline */
.company-sub {
    font-size: 30px;
    font-weight: 600;
    color: #FF5E0E;
    margin-left: 5px;
}

</style>
""", unsafe_allow_html=True)

col1, col2 = st.columns([0.6, 9])

with col1:
    st.image(logo_path, width=150)   # Same visual weight as dashboard icon

with col2:
    st.markdown("""
        <div class="company-header">
            PEAKADS LLP
            <span class="company-sub">Revenue & Billing Tracker</span>
        </div>
    """, unsafe_allow_html=True)

tabs = st.tabs([
    "📊 Dashboard",
    "📈 Summary",
    "📁 Master Data",
    " DSP (Customers)",
    " SSP (Vendors)",
    "📝 Partner Onboarding Form",
    "🤝 List of Partners"
])

# ====================================================
# 1️⃣ PARTNER ONBOARDING FORM
# ====================================================

with tabs[5]:

    st.header("📝Partner Onboarding Form")

    st.markdown("""
        <style>
        label {
            font-weight: 700 !important;
        }

        div.stButton > button {
            background-color: #003366;
            color: white;
            font-weight: 600;
            border-radius: 6px;
            padding: 8px 25px;
            border: none;
            transition: 0.2s ease-in-out;
        }

        div.stButton > button:hover {
            background-color: #0059b3;
            transform: scale(1.05);
        }
        </style>
    """, unsafe_allow_html=True)

    # ==============================
    # A4 LANDSCAPE GRID (3 COLUMNS)
    # ==============================

    col1, col2, col3 = st.columns(3)

    # ---------- COLUMN 1 ----------
    with col1:
        agreement_date = st.date_input(
            "Agreement Start Date",
            format="DD/MM/YYYY"
        )
        legal_name = st.text_input("Legal Entity Name", max_chars=50)
        short_name = st.text_input("Short Name using in Bidscube", max_chars=20)
        country_list = [
            "India (IN)", "United States (US)", "United Kingdom (UK)",
            "Singapore (SG)", "UAE (AE)", "Germany (DE)",
            "Australia (AU)", "Canada (CA)"
        ]
        country = st.selectbox("Country", country_list)

    # ---------- COLUMN 2 ----------
    with col2:
        entity_type = "Indian" if country == "India (IN)" else "Foreign"

        gstin = st.text_input(
            "GSTIN",
            max_chars=15,
            disabled=(entity_type != "Indian")
        )

        payment_terms = st.selectbox(
            "Payment Terms",
            ["Net 30", "Net 45", "Net 60", "Net 90"]
        )

        contact_person = st.text_input("Contact Person", max_chars=20)
        designation = st.text_input("Designation", max_chars=20)

    # ---------- COLUMN 3 ----------
    with col3:
        contact_no = st.text_input("Contact No.", max_chars=15)
        email1 = st.text_input("Email 1", max_chars=30)
        email2 = st.text_input("Email 2", max_chars=30)
        email3 = st.text_input("Email 3", max_chars=30)

    # ---------- FULL WIDTH ROW ----------
    address = st.text_area("Registered Address", height=80)

    col4, col5 = st.columns(2)

    with col4:
        finance_contact = st.text_input("Finance Contact", max_chars=20)

    with col5:
        finance_email = st.text_input("Finance Email", max_chars=30)

    # ---------- CENTER BUTTON ----------
    col_left, col_center, col_right = st.columns([2, 1, 2])
    with col_center:
        save_clicked = st.button("Save Partner")

    if save_clicked:

        with st.spinner("Saving Partner..."):

            partner_data = {
                "Agreement Start Date": agreement_date,
                "Legal Entity Name": legal_name,
                "Short Name using in Bidscube": short_name,
                "Registered Address": address,
                "Country": country,
                "Foreign / Indian Entity": entity_type,
                "GSTIN": gstin,
                "Payment Terms": payment_terms,
                "Contact Person": contact_person,
                "Designation": designation,
                "Contact No.": contact_no,
                "Email 1": email1,
                "Email 2": email2,
                "Email 3": email3,
                "Finance Contact": finance_contact,
                "Finance Email": finance_email
            }

            df = load_sheet("Partner List")
            new_row = pd.DataFrame([partner_data])

            if df.empty:
                df = new_row
            else:
                for col in new_row.columns:
                    if col not in df.columns:
                        df[col] = ""
                df = pd.concat([df, new_row], ignore_index=True)

            save_sheet(df, "Partner List")

        st.success("Successfully Saved in Google Sheet")


# ====================================================
# 2️⃣ MASTER DATA TAB
# ====================================================

with tabs[2]:

    st.header("📁Master Data")
    
    fy_list = generate_financial_years()

    col1, col2, col3 = st.columns(3)

    with col1:
        selected_fy = st.selectbox(
            "Financial Year",
            options=["All"] + fy_list,
            index=0,
            key="master_fy"
        )

    with col2:
        month_options = ["All"]

        if selected_fy != "All":
            fy_start = int(selected_fy.split("-")[0])
            months = pd.date_range(
                start=f"{fy_start}-04-01",
                end=f"{fy_start+1}-03-31",
                freq="MS"
            )
            month_options += months.strftime("%b-%Y").tolist()

        selected_month = st.selectbox(
            "Month",
            options=month_options,
            index=0,
            disabled=False,
            key="master_month"
        )

    with col3:
        selected_quarter = st.selectbox(
            "Quarter",
            options=["All", "Q1", "Q2", "Q3", "Q4"],
            index=0,
            key="master_quarter"
        )

    # Disable month if quarter selected
    if selected_quarter != "All":
        selected_month = "All"

    # 🔹 Load Master Data from Google
    if "master_df" not in st.session_state:
        st.session_state.master_df = load_master_data_from_gsheet()

    df_master = st.session_state.master_df
    
    df_filtered = df_master.copy()

    if selected_fy != "All":
        fy_start, fy_end = get_fy_date_range(selected_fy)

        df_filtered["Month"] = pd.to_datetime(df_filtered["Month"], errors="coerce")

        df_filtered = df_filtered[
            (df_filtered["Month"] >= fy_start) &
            (df_filtered["Month"] <= fy_end)
        ]

    if selected_quarter != "All" and selected_fy != "All":
        q_start, q_end = get_quarter_range(selected_fy, selected_quarter)

        df_filtered = df_filtered[
            (df_filtered["Month"] >= q_start) &
            (df_filtered["Month"] <= q_end)
        ]

    elif selected_month != "All":
        selected_month_dt = pd.to_datetime(selected_month, format="%b-%Y", errors="coerce")

        df_filtered = df_filtered[
            df_filtered["Month"] == selected_month_dt
        ]
    
    df_partner = load_partner_list_from_gsheet()

    if df_master.empty:
        st.warning("No Master Data Found")
        st.stop()

    # Convert Month to real datetime internally
    # Ensure Month is datetime BEFORE sorting
    df_filtered["Month"] = pd.to_datetime(
        df_filtered["Month"],
        errors="coerce"
    )

    df_master = df_filtered.sort_values("Month")

    # Convert to display format AFTER everything
    df_master["Month"] = df_master["Month"].dt.strftime("%b-%Y")
    
    if not df_master.empty:

        df_master["Net $ (BC)"] = df_master["DSP $ (BC)"] - df_master["SSP $ (BC)"]

        df_master["Month"] = pd.to_datetime(df_master["Month"])
        df_master["Month"] = df_master["Month"].dt.strftime("%b-%Y")

        for index, row in df_master.iterrows():

            if df_partner.empty:
                continue

            if "Short Name using in Bidscube" not in df_partner.columns:
                continue

            partner_match = df_partner[
                df_partner["Short Name using in Bidscube"] == row["Partner Name"]
            ]

            if partner_match.empty:
                continue

            country = partner_match.iloc[0].get("Country", "")
            gstin = partner_match.iloc[0].get("GSTIN", "")
            net_term = partner_match.iloc[0].get("Payment Terms", "")

            if country == "India (IN)":
                df_master.loc[index, "I/F"] = "Indian"
                df_master.loc[index, "USD/INR"] = "INR"
            else:
                df_master.loc[index, "I/F"] = "Foreign"
                df_master.loc[index, "USD/INR"] = "USD"

            df_master.loc[index, "GSTIN"] = gstin
            df_master.loc[index, "NET Term"] = net_term

        # Ensure required columns exist
        for col in ["C DSP $", "C SSP $", "C Net $", "Category (DSP/SSP)"]:
            if col not in df_master.columns:
                df_master[col] = 0.0

        
               
        with col2:
            search_text = st.text_input(
                "🔍 Search",
                placeholder="Global Search...",
                key="master_search"
            )
            
        with col3:
            st.markdown("<br>", unsafe_allow_html=True)  # aligns button vertically
            refresh_clicked = st.button("🔄 Refresh Master Data", key="master_refresh_button")
            
        if refresh_clicked:
            st.session_state.master_df = load_master_data_from_gsheet()
            st.session_state.last_saved_df = st.session_state.master_df.copy()
            st.success("Data refreshed from Google Sheet")
            st.rerun()
            
                        
        # ===== AGGRID MASTER DATA TABLE =====

        from st_aggrid import AgGrid, GridOptionsBuilder, JsCode

        # Ensure numeric columns
        numeric_cols = [
            "DSP $ (BC)",
            "SSP $ (BC)",
            "Net $ (BC)",
            "C DSP $",
            "C SSP $",
            "C Net $",
        ]

        df_master["C DSP $"] = pd.to_numeric(df_master["C DSP $"], errors="coerce").fillna(0)
        df_master["C SSP $"] = pd.to_numeric(df_master["C SSP $"], errors="coerce").fillna(0)

        month_comparator = JsCode("""
        function(date1, date2) {
            function parseMonth(str) {
                if (!str) return new Date(0);
                const [mon, year] = str.split("-");
                return new Date(mon + " 1, " + year);
            }
            const d1 = parseMonth(date1);
            const d2 = parseMonth(date2);
            return d1 - d2;
        }
        """)
        
        # -------- GRID BUILDER --------
        from st_aggrid import GridOptionsBuilder, JsCode
        
        # Currency formatter
        currency_formatter = JsCode("""
        function(params) {
            if (params.value == null || params.value === '') return '';
            return '$' + parseFloat(params.value).toLocaleString(undefined, {minimumFractionDigits: 2});
        }
        """)
        
        bg_style_js = JsCode("""
        function(params) {

            if (params.node.rowPinned) {
                return {};   // skip footer
            }

            let col = params.colDef.field;

            if (col === "C DSP $" || col === "C SSP $") {
                return { backgroundColor: "#FFF4E5" };   // very light orange
            }

            if (col === "C Net $") {
                return { backgroundColor: "#E8F5E9" };   // light green
            }

            if (col === "Net $ (BC)") {
                return { backgroundColor: "#F3E5F5" };   // light purple
                fontWeight: "bold"
            }

            return {};
        }
        """)
        
        gb = GridOptionsBuilder.from_dataframe(df_master)
        
        gb.configure_column(
            "Month",
            comparator=month_comparator
        )
        
        negative_style = JsCode("""
        function(params) {
            if (params.value < 0) {
                return {
                    color: 'red',
                    fontWeight: 'bold'
                };
            }
        }
        """)
        
        # Editable Table
        numeric_cols = [
            "DSP $ (BC)",
            "SSP $ (BC)",
            "Net $ (BC)",
            "C DSP $",
            "C SSP $",
            "C Net $",
        ]
        
        editable_cols = ["C DSP $", "C SSP $"]
        
        # Ensure numeric conversion before grid
        for col in numeric_cols:
            if col in df_master.columns:
                df_master[col] = pd.to_numeric(df_master[col], errors="coerce").fillna(0)

        # Apply uniform currency formatting to ALL numeric columns
        for col in numeric_cols:
            if col in df_master.columns:
                gb.configure_column(
                    col,
                    type=["numericColumn"],
                    editable=(col in editable_cols),
                    valueFormatter=currency_formatter,
                    cellStyle=bg_style_js
                )

                                
        from st_aggrid import JsCode

        # C Net = C DSP - C SSP (Excel style)
        net_value_getter = JsCode("""
        function(params) {
            let dsp = parseFloat(params.data["C DSP $"]) || 0;
            let ssp = parseFloat(params.data["C SSP $"]) || 0;
            return dsp - ssp;
        }
        """)

        # Category depends on C Net
        category_value_getter = JsCode("""
        function(params) {
        
            // 🚫 Do NOT apply to footer row
            if (params.node.rowPinned) {
                return "";
            }
            
            let dsp = parseFloat(params.data["C DSP $"]) || 0;
            let ssp = parseFloat(params.data["C SSP $"]) || 0;
            let net = dsp - ssp;
            return net >= 0 ? "DSP" : "SSP";
        }
        """)

        gb.configure_column(
            "C Net $",
            editable=False,
            type=["numericColumn"],
            valueGetter=net_value_getter,
            valueFormatter=currency_formatter,
            cellStyle=JsCode("""
                function(params) {

                    if (params.node.rowPinned) return {};

                    let style = {
                        backgroundColor: "#E8F5E9",   // light green
                        fontWeight: "bold"
                    };

                    if (params.value < 0) {
                        style.color = "red";
                    }

                    return style;
                }
            """)
        )
        
        gb.configure_column(
            "Net $ (BC)",
            type=["numericColumn"],
            valueFormatter=currency_formatter,
            cellStyle=JsCode("""
                function(params) {

                    if (params.node.rowPinned) return {};

                    let style = {
                        backgroundColor: "#F3E5F5",   // light purple
                        fontWeight: "bold"           // <-- BOLD ALWAYS
                    };

                    if (params.value < 0) {
                        style.color = "red";
                    }

                    return style;
                }
            """)
        )

        gb.configure_column(
            "Category (DSP/SSP)",
            editable=False,
            valueGetter=category_value_getter
        )
                
        # Freeze first 2 columns
        first_two_cols = df_master.columns[:2]

        for col in first_two_cols:
            gb.configure_column(col, pinned="left")
            
        from st_aggrid import JsCode

        pinned_style = JsCode("""
        function(params) {
            if (params.column.pinned) {
                return {
                    fontWeight: 'bold',
                    borderRight: '2px solid #003366'
                };
            }
        }
        """)

        for col in first_two_cols:
            gb.configure_column(
                col,
                pinned="left",
                cellStyle=pinned_style
            )
        
        gb.configure_default_column(
            resizable=True,
            sortable=True,
            filter=False
        )

        # Red negative styling
        negative_style = JsCode("""
        function(params) {
            if (params.value < 0) {
                return {color: 'red', fontWeight: 'bold'};
            }
        }
        """)

        # -------- REAL-TIME GRAND TOTAL (JS BASED) --------

        grand_total_js = JsCode("""
        function(api) {
            let totals = {};
            let numericCols = ["DSP $ (BC)", "SSP $ (BC)", "Net $ (BC)", "C DSP $", "C SSP $", "C Net $"];

            numericCols.forEach(col => totals[col] = 0);

            api.forEachNodeAfterFilter(function(node) {
                numericCols.forEach(function(col) {
                    let val = parseFloat(node.data[col]);
                    if (!isNaN(val)) {
                        totals[col] += val;
                    }
                });
            });

            let totalRow = { Month: "Grand Total" };
            numericCols.forEach(col => totalRow[col] = totals[col]);

            api.setPinnedBottomRowData([totalRow]);
        }
        """)
        
        gridOptions = gb.build()
        gridOptions["onFirstDataRendered"] = grand_total_js
        gridOptions["onFilterChanged"] = grand_total_js
        gridOptions["onModelUpdated"] = grand_total_js
        gridOptions["onCellValueChanged"] = grand_total_js
        
        if search_text:
            gridOptions["quickFilterText"] = search_text

        # Footer styling
        gridOptions["getRowStyle"] = JsCode("""
        function(params) {
            if (params.node.rowPinned) {
                return {
                    backgroundColor: '#003366',
                    color: 'white',
                    fontWeight: 'bold',
                    fontSize: '14px'
                };
            }
        }
        """)
        
        custom_css = {
            ".ag-header": {
                "background-color": "#003366 !important",
                "color": "white !important",
                "font-weight": "bold !important",
                "font-size": "14px !important"
            },
            ".ag-header-cell-label": {
                "color": "white !important",
                "font-weight": "bold !important"
            }
        }
        
        st.markdown("""
        <script>
        document.addEventListener("DOMContentLoaded", function() {
            const counters = document.querySelectorAll(".kpi-value");
            counters.forEach(counter => {
                const updateCount = () => {
                    const target = +counter.getAttribute("data-value");
                    const duration = 800;
                    const stepTime = 20;
                    const steps = duration / stepTime;
                    let count = 0;
                    const increment = target / steps;

                    const timer = setInterval(() => {
                        count += increment;
                        if (Math.abs(count) >= Math.abs(target)) {
                            counter.innerText = counter.innerText.includes('%')
                                ? target.toFixed(2) + '%'
                                : '$' + target.toLocaleString(undefined, {minimumFractionDigits: 2});
                            clearInterval(timer);
                        } else {
                            counter.innerText = counter.innerText.includes('%')
                                ? count.toFixed(2) + '%'
                                : '$' + count.toLocaleString(undefined, {minimumFractionDigits: 2});
                        }
                    }, stepTime);
                };
                updateCount();
            });
        });
        </script>
        """, unsafe_allow_html=True)

        # -------- GRAND TOTAL (Search + Month Reactive) --------

        filtered_df = df_master.copy()

        if search_text:
            search_lower = search_text.lower()

            mask = filtered_df.apply(
                lambda row: row.astype(str).str.lower().str.contains(search_lower).any(),
                axis=1
            )
            filtered_df = filtered_df[mask]

        grand_total_values = filtered_df.sum(numeric_only=True)

        grand_total_row = {col: "" for col in df_master.columns}
        grand_total_row["Month"] = "Grand Total"

        for col in numeric_cols:
            if col in grand_total_row:
                grand_total_row[col] = grand_total_values.get(col, 0)

        gridOptions["pinnedBottomRowData"] = [grand_total_row]
        
        from st_aggrid import GridUpdateMode
        
        grid_df = df_master.copy().reset_index(drop=True)

        grid_response = AgGrid(
            grid_df,
            gridOptions=gridOptions,
            allow_unsafe_jscode=True,
            update_mode=GridUpdateMode.MODEL_CHANGED,
            data_return_mode="AS_INPUT",
            fit_columns_on_grid_load=True,
            height=550,
            custom_css=custom_css
        )
        
                
        if grid_response["selected_rows"] is not None:
            pass  # ignore selection

        # =========================
        # AUTO SAVE ON EDIT (BATCH SAFE VERSION)
        # =========================

        updated_df = pd.DataFrame(grid_response["data"])
        
        if "last_saved_df" not in st.session_state:
            st.session_state.last_saved_df = grid_df.copy()

        data_changed = not updated_df.equals(st.session_state.last_saved_df)

        updated_df["C DSP $"] = pd.to_numeric(updated_df["C DSP $"], errors="coerce").fillna(0)
        updated_df["C SSP $"] = pd.to_numeric(updated_df["C SSP $"], errors="coerce").fillna(0)
        
        if data_changed:

            original_df = st.session_state.master_df.reset_index(drop=True)
            worksheet = spreadsheet.worksheet("Master Data")

            batch_requests = []

            worksheet = spreadsheet.worksheet("Master Data")
            sheet_data = worksheet.get_all_records()
            sheet_df = pd.DataFrame(sheet_data)

            batch_requests = []

            for _, row in updated_df.iterrows():

                match = sheet_df[
                    (sheet_df["Month"] == row["Month"]) &
                    (sheet_df["Partner Name"] == row["Partner Name"])
                ]

                if match.empty:
                    continue

                sheet_index = match.index[0] + 2  # +2 because sheet starts at row 2

                batch_requests.append({
                    "range": f"F{sheet_index}:H{sheet_index}",
                    "values": [[
                        float(row["C DSP $"]),
                        float(row["C SSP $"]),
                        float(row["C DSP $"]) - float(row["C SSP $"])
                    ]]
                })

            if batch_requests:
                worksheet.batch_update(batch_requests, value_input_option="USER_ENTERED")

            st.session_state.last_saved_df = updated_df.copy()
            st.toast("Auto-saved ✅")
                        
        df_master = df_master.reset_index(drop=True)
        updated_df = updated_df.reset_index(drop=True)

               
        worksheet = spreadsheet.worksheet("Master Data")

                
        # RED negative styling
        def highlight_negative(val):
            if isinstance(val, (int, float)) and val < 0:
                return "color: red; font-weight: bold;"
            return ""

               
    else:
        st.warning("No Master Data Found in Excel")

# ====================================================
# Other Tabs (Blank for Now)
# ====================================================

def render_kpi(label, value, is_currency=True):

    numeric_value = float(value)

    css_class = "kpi-positive" if numeric_value >= 0 else "kpi-negative"

    display_value = (
        f"${numeric_value:,.2f}" if is_currency
        else f"{numeric_value:.2f}%"
    )

    st.markdown(f"""
    <div class="kpi-card {css_class}">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value" data-value="{numeric_value}">
            {display_value}
        </div>
    </div>
    """, unsafe_allow_html=True)

with tabs[0]:

    st.header("📊 Dashboard")
    
    fy_list = generate_financial_years()

    col1, col2, col3 = st.columns(3)

    with col1:
        selected_fy = st.selectbox(
            "Financial Year",
            options=["All"] + fy_list,
            index=0,
            key="dashboard_fy"
        )

    with col2:
        month_options = ["All"]

        if selected_fy != "All":
            fy_start = int(selected_fy.split("-")[0])
            months = pd.date_range(
                start=f"{fy_start}-04-01",
                end=f"{fy_start+1}-03-31",
                freq="MS"
            )
            month_options += months.strftime("%b-%Y").tolist()

        selected_month = st.selectbox(
            "Month",
            options=month_options,
            index=0,
            disabled=False,
            key="dashboard_month"
        )

    with col3:
        selected_quarter = st.selectbox(
            "Quarter",
            options=["All", "Q1", "Q2", "Q3", "Q4"],
            index=0,
            key="dashboard_quarter"
        )

    # Disable month if quarter selected
    if selected_quarter != "All":
        selected_month = "All"

    df_master = load_master_data_from_gsheet()    
    
    df_filtered = df_master.copy()

    if selected_fy != "All":
        fy_start, fy_end = get_fy_date_range(selected_fy)

        df_filtered["Month"] = pd.to_datetime(df_filtered["Month"], errors="coerce")

        df_filtered = df_filtered[
            (df_filtered["Month"] >= fy_start) &
            (df_filtered["Month"] <= fy_end)
        ]

    if selected_quarter != "All" and selected_fy != "All":
        q_start, q_end = get_quarter_range(selected_fy, selected_quarter)

        df_filtered = df_filtered[
            (df_filtered["Month"] >= q_start) &
            (df_filtered["Month"] <= q_end)
        ]

    elif selected_month != "All":
        selected_month_dt = pd.to_datetime(selected_month, format="%b-%Y", errors="coerce")

        df_filtered = df_filtered[
            df_filtered["Month"] == selected_month_dt
        ]
    
    if df_master.empty:
        st.warning("No Master Data Available")
        st.stop()

    # 🔹 Convert Month properly
    df_filtered["Month"] = pd.to_datetime(
        df_filtered["Month"],
        errors="coerce"
    )

    df_master = df_filtered.sort_values("Month")
    df_master["Month"] = df_master["Month"].dt.strftime("%b-%Y")

    # 🔹 Month Filter
    months = df_master["Month"].dropna().unique().tolist()
    months_sorted = sorted(
        months,
        key=lambda x: pd.to_datetime(x, format="%b-%Y")
    )

    if df_master.empty:
        st.warning("No data available for selected filter.")
        st.stop()

    # 🔹 Calculate KPIs (ONLY ON FILTERED DATA)
    (
        df_master,
        total_dsp,
        total_ssp,
        total_net,
        total_c_dsp,
        total_c_ssp,
        total_c_net,
        ivt,
        ivt_percent,
        c_profit_percent
    ) = calculate_kpis(df_master)

    
    # 🔹 Subtabs
    subtabs = st.tabs([
        "📊 Key Financial Metrics",
        "📈 Monthly Revenue Trend",
        "🏆 Top 10 Partners by Net",
        "👥 Partner Onboarded"
    ])

    with subtabs[0]:

        # ---- KPI SECTION HERE ----
        st.markdown("### 📊Key Financial Metrics")

        c1, c2, c3 = st.columns(3)
        c4, c5, c6 = st.columns(3)
        c7, c8, c9 = st.columns(3)

        with c1:
            render_kpi("Total BC DSP $", total_dsp)

        with c2:
            render_kpi("Total BC SSP $", total_ssp)

        with c3:
            render_kpi("Total BC Net $", total_net)

        with c4:
            render_kpi("Total C DSP $", total_c_dsp)

        with c5:
            render_kpi("Total C SSP $", total_c_ssp)

        with c6:
            render_kpi("Total C Net $", total_c_net)

        with c7:
            render_kpi("IVT $", ivt)

        with c8:
            render_kpi("IVT %", ivt_percent, is_currency=False)

        with c9:
            render_kpi("C Profit %", c_profit_percent, is_currency=False)

        st.markdown('</div>', unsafe_allow_html=True)
        
        # ----------------------------
        # KPI Calculations
        # ----------------------------
        total_dsp = df_master["DSP $ (BC)"].sum()
        total_ssp = df_master["SSP $ (BC)"].sum()
        total_net = df_master["Net $ (BC)"].sum()

        total_c_dsp = df_master["C DSP $"].sum()
        total_c_ssp = df_master["C SSP $"].sum()

        ivt = total_net - total_c_net
        ivt_percent = (ivt / total_dsp * 100) if total_dsp != 0 else 0
        c_profit_percent = (total_c_net / total_c_dsp * 100) if total_c_dsp != 0 else 0

        k1, k2, k3 = st.columns(3)
        k4, k5, k6 = st.columns(3)
        k7, k8, k9 = st.columns(3)

    with subtabs[1]:    
        
        import altair as alt

        st.markdown("### 📈Monthly Revenue Trend")

        view_type = st.radio(
            "View By",
            ["Monthly", "Quarterly (FY)"],
            horizontal=True,
            key="revenue_view_toggle"
        )

        data = df_master.copy()
        data["Month"] = pd.to_datetime(data["Month"], errors="coerce")

        # ---------------- MONTHLY VIEW ----------------
        if view_type == "Monthly":

            data["MonthPeriod"] = data["Month"].dt.to_period("M")

            monthly = (
                data.groupby("MonthPeriod", as_index=False)
                .agg({"Net $ (BC)": "sum"})
            )

            monthly["Date"] = monthly["MonthPeriod"].dt.to_timestamp()
            monthly = monthly.sort_values("Date")
            monthly["Label"] = monthly["Date"].dt.strftime("%b-%Y")

            chart = (
                alt.Chart(monthly)
                .mark_line(point=True)
                .encode(
                    x=alt.X(
                        "Label:N",
                        sort=list(monthly["Label"]),
                        title="Month"
                    ),
                    y=alt.Y(
                        "Net $ (BC):Q",
                        title="Net Revenue ($)",
                        axis=alt.Axis(format="$,.0f")
                    ),
                    tooltip=[
                        alt.Tooltip("Label:N", title="Month"),
                        alt.Tooltip("Net $ (BC):Q", format=",.2f")
                    ]
                )
                .properties(height=400)
            )

            st.altair_chart(chart, use_container_width=True)


        # ---------------- QUARTERLY FY VIEW ----------------
        else:

            # FY calculation (April to March)
            def get_fy_year(date):
                return date.year if date.month >= 4 else date.year - 1

            def get_fy_quarter(date):
                if 4 <= date.month <= 6:
                    return "Q1"
                elif 7 <= date.month <= 9:
                    return "Q2"
                elif 10 <= date.month <= 12:
                    return "Q3"
                else:
                    return "Q4"

            data["FY"] = data["Month"].apply(get_fy_year)
            data["Quarter"] = data["Month"].apply(get_fy_quarter)

            data["FY_Label"] = (
                "FY "
                + data["FY"].astype(str)
                + "-"
                + (data["FY"] + 1).astype(str).str[-2:]
            )

            data["Quarter_Label"] = data["FY_Label"] + " " + data["Quarter"]

            quarterly = (
                data.groupby(["FY", "Quarter"], as_index=False)
                .agg({"Net $ (BC)": "sum"})
            )

            # Create proper sorting order
            quarter_order = {"Q1": 1, "Q2": 2, "Q3": 3, "Q4": 4}
            quarterly["QuarterOrder"] = quarterly["Quarter"].map(quarter_order)

            quarterly = quarterly.sort_values(["FY", "QuarterOrder"])

            quarterly["Label"] = (
                "FY "
                + quarterly["FY"].astype(str)
                + "-"
                + (quarterly["FY"] + 1).astype(str).str[-2:]
                + " "
                + quarterly["Quarter"]
            )

            chart = (
                alt.Chart(quarterly)
                .mark_bar()
                .encode(
                    x=alt.X(
                        "Label:N",
                        sort=list(quarterly["Label"]),
                        title="Quarter"
                    ),
                    y=alt.Y(
                        "Net $ (BC):Q",
                        title="Net Revenue ($)",
                        axis=alt.Axis(format="$,.0f")
                    ),
                    tooltip=[
                        alt.Tooltip("Label:N"),
                        alt.Tooltip("Net $ (BC):Q", format=",.2f")
                    ]
                )
                .properties(height=400)
            )

            st.altair_chart(chart, use_container_width=True)

    with subtabs[2]:    
        import altair as alt

        # ---- TOP 10 PARTNERS BY NET ----
        top10 = (
            df_master.groupby("Partner Name", as_index=False)
            .agg({"Net $ (BC)": "sum"})
        )

        # Sort highest to lowest
        top10 = top10.sort_values("Net $ (BC)", ascending=False)

        # Take top 10
        top10 = top10.head(10)

        # ---- ALTAIR BAR CHART ----
        chart = (
            alt.Chart(top10)
            .mark_bar()
            .encode(
                x=alt.X(
                    "Net $ (BC):Q",
                    title="Net Revenue ($)",
                    axis=alt.Axis(format="$,.0f")
                ),
                y=alt.Y(
                    "Partner Name:N",
                    sort='-x',   # 🔥 highest to lowest
                    title="Partner"
                ),
                tooltip=[
                    "Partner Name",
                    alt.Tooltip("Net $ (BC):Q", format=",.2f")
                ]
            )
            .properties(height=400)
        )

        st.markdown("### 🏆Top 10 Partners by Net")
        st.altair_chart(chart, use_container_width=True)
        
                        
    with subtabs[3]:

        st.markdown("### 👥 Partner Onboarded Overview")

        df_partner = load_partner_list_from_gsheet()

        if df_partner.empty:
            st.warning("No Partner Data Found")
            st.stop()

        df_partner = df_partner.dropna(how="all")

        if "Country" not in df_partner.columns:
            st.warning("Country column missing")
            st.stop()

        # -----------------------------
        # TOTAL COUNT (FULL WIDTH)
        # -----------------------------
        
        total_partners = len(df_partner)

        st.metric(
            "Total Partners Onboarded",
            total_partners
        )

        st.divider()
        
        import altair as alt

        # Convert Agreement Start Date
        df_partner["Agreement Start Date"] = pd.to_datetime(
            df_partner["Agreement Start Date"],
            errors="coerce"
        )

        df_partner = df_partner.dropna(subset=["Agreement Start Date"])

        # Create Month column
        df_partner["Month"] = df_partner["Agreement Start Date"].dt.to_period("M")

        monthly_counts = (
            df_partner.groupby("Month")
            .size()
            .reset_index(name="Partner Count")
        )

        monthly_counts["Month"] = monthly_counts["Month"].dt.to_timestamp()
        monthly_counts = monthly_counts.sort_values("Month")

        monthly_counts["Label"] = monthly_counts["Month"].dt.strftime("%b-%Y")

        st.markdown("### Total Partners Onboarded (Month-wise)")

        chart = (
            alt.Chart(monthly_counts)
            .mark_bar()
            .encode(
                x=alt.X(
                    "Month:T",   # 🔥 Use real datetime
                    title="Month",
                    axis=alt.Axis(format="%b-%Y")
                ),
                y=alt.Y(
                    "Partner Count:Q",
                    title="Partners Onboarded"
                ),
                tooltip=[
                    alt.Tooltip("Month:T", format="%b-%Y"),
                    "Partner Count"
                ]
            )
            .properties(height=250)
        )

        st.altair_chart(chart, use_container_width=True)

        st.markdown("</div>", unsafe_allow_html=True)

        # -----------------------------
        # COUNTRY-WISE COUNT
        # -----------------------------
        country_counts = (
            df_partner["Country"]
            .fillna("Unknown")
            .value_counts()
            .reset_index()
        )

        country_counts.columns = ["Country Name", "Country Count"]

        # ---- SIDE BY SIDE TABLE + CHART ----
        col1, col2 = st.columns([1, 2])

        with col1:
            st.markdown("#### Country-wise Count")
            st.dataframe(
                country_counts.reset_index(drop=True),
                use_container_width=True,
                height=350,
                hide_index=True
            )
        
with tabs[1]:

    st.header("📈 Summary")

    # ==============================
    # 4 EQUAL PARTS (2x2 GRID)
    # ==============================

    row1_col1, row1_col2 = st.columns(2)
    row2_col1, row2_col2 = st.columns(2)

    # ======================================================
    # 🟦 PART 1
    # ======================================================
    with row1_col1:

        st.subheader("Partner Summary - Monthwise")

        df_master = load_master_data_from_gsheet()

        if df_master.empty:
            st.warning("No Master Data Found")
            st.stop()

        # ---- Ensure correct types ----
        df_master["Month"] = pd.to_datetime(df_master["Month"], errors="coerce")
        df_master["C DSP $"] = pd.to_numeric(df_master["C DSP $"], errors="coerce").fillna(0)
        df_master["C SSP $"] = pd.to_numeric(df_master["C SSP $"], errors="coerce").fillna(0)

        # ---- Partner Dropdown ----
        partner_list = sorted(df_master["Partner Name"].dropna().unique().tolist())

        # Add placeholder as first option
        partner_options = ["Select Partner"] + partner_list

        # 🔥 Calculate dynamic width
        max_length = max(len(str(name)) for name in partner_list)
        dynamic_width = max(300, min(1000, max_length * 11))  # 11px per character

        # ---- Bold Label + Dynamic Width Styling ----
        st.markdown(f"""
        <style>
        /* Bold Label */
        label[data-testid="stWidgetLabel"] p {{
            font-weight: 800 !important;
        }}

        /* Resize ONLY this selectbox */
        div[data-testid="stSelectbox"] > div {{
            width: {dynamic_width}px !important;
        }}
        </style>
        """, unsafe_allow_html=True)

        selected_partner = st.selectbox(
            "Select Partner",
            partner_options,
            index=0,
            key="summary_partner_part1"
        )

        if selected_partner == "Select Partner":
            st.info("Please select a partner to view summary.")

        else:

            # ---- Filter Partner ----
            df_partner = df_master[
                df_master["Partner Name"] == selected_partner
            ].copy()

            # -------------------------------------------------------
            # REMOVE MONTHS MARKED GREEN OR LIGHT YELLOW IN DSP/SSP
            # -------------------------------------------------------

            def get_excluded_months(sheet_name, name_column, amount_column):
                try:
                    worksheet = spreadsheet.worksheet(sheet_name)
                    sheet_data = worksheet.get_all_records()
                    df_sheet = pd.DataFrame(sheet_data)

                    if df_sheet.empty:
                        return set()

                    df_sheet["Month"] = pd.to_datetime(
                        df_sheet["Month"],
                        errors="coerce"
                    )

                    df_sheet["Receivable $"] = pd.to_numeric(
                        df_sheet.get("Receivable $", df_sheet.get("Payable $", 0)),
                        errors="coerce"
                    ).fillna(0)

                    df_sheet[amount_column] = pd.to_numeric(
                        df_sheet.get(amount_column, 0),
                        errors="coerce"
                    ).fillna(0)

                    df_sheet = df_sheet[
                        df_sheet[name_column] == selected_partner
                    ]

                    green = df_sheet[
                        df_sheet[amount_column] == df_sheet["Receivable $"]
                    ]

                    yellow = df_sheet[
                        (df_sheet[amount_column] != 0) &
                        (df_sheet[amount_column] != df_sheet["Receivable $"])
                    ]

                    excluded = pd.concat([green, yellow])

                    return set(
                        excluded["Month"].dt.strftime("%b-%Y")
                    )

                except:
                    return set()
                    
            # -----------------------------
            # APPLY EXCLUSION LOGIC
            # -----------------------------

            excluded_dsp = get_excluded_months(
                "DSP (Customers)",
                "DSP Name",
                "Received Amount $"
            )

            excluded_ssp = get_excluded_months(
                "SSP (Vendors)",
                "SSP Name",
                "Paid Amount $"
            )

            excluded_months = excluded_dsp.union(excluded_ssp)

            # Convert Month to string for comparison
            df_partner["MonthStr"] = df_partner["Month"].dt.strftime("%b-%Y")

            # Remove excluded months
            df_partner = df_partner[
                ~df_partner["MonthStr"].isin(excluded_months)
            ]

            df_partner.drop(columns=["MonthStr"], inplace=True)        

            if df_partner.empty:
                st.warning("No Data for Selected Partner")
            else:
                # 🔥 BUILD SUMMARY ONLY HERE

                required_cols = ["C DSP $", "C SSP $"]
                for col in required_cols:
                    if col not in df_partner.columns:
                        df_partner[col] = 0.0

                df_partner["C DSP $"] = pd.to_numeric(df_partner["C DSP $"], errors="coerce").fillna(0)
                df_partner["C SSP $"] = pd.to_numeric(df_partner["C SSP $"], errors="coerce").fillna(0)

                df_summary = (
                    df_partner
                    .groupby("Month", as_index=False)
                    .agg({
                        "C DSP $": "sum",
                        "C SSP $": "sum"
                    })
                )

                df_summary["Offset $ USD"] = df_summary["C DSP $"] - df_summary["C SSP $"]
                df_summary["Month"] = df_summary["Month"].dt.strftime("%b-%Y")

                df_summary.rename(columns={
                    "C DSP $": "As DSP",
                    "C SSP $": "As SSP"
                }, inplace=True)

                total_row = {
                    "Month": "Total",
                    "As DSP": df_summary["As DSP"].sum(),
                    "As SSP": df_summary["As SSP"].sum(),
                    "Offset $ USD": df_summary["Offset $ USD"].sum()
                }

                df_summary = pd.concat(
                    [df_summary, pd.DataFrame([total_row])],
                    ignore_index=True
                )

            # ======================================================
            # AGGRID (MASTER STYLE)
            # ======================================================

            from st_aggrid import GridOptionsBuilder, AgGrid, JsCode
            from st_aggrid import GridUpdateMode

            currency_formatter = JsCode("""
            function(params) {
                if (params.value == null || params.value === '') return '';
                return '$' + parseFloat(params.value).toLocaleString(undefined, {minimumFractionDigits: 2});
            }
            """)

            gb = GridOptionsBuilder.from_dataframe(df_summary)

            numeric_cols = ["As DSP", "As SSP", "Offset $ USD"]

            from st_aggrid import JsCode

            currency_formatter = JsCode("""
            function(params) {
                if (params.value == null || params.value === '') return '';
                return '$' + parseFloat(params.value).toLocaleString(undefined, {minimumFractionDigits: 2});
            }
            """)

            offset_style = JsCode("""
            function(params) {

                if (params.data.Month === "Total") {
                    return {
                        backgroundColor: '#003366',
                        color: 'white',
                        fontWeight: 'bold'
                    };
                }

                let style = {
                    backgroundColor: "#E8F5E9",  // light green like C Net $
                    fontWeight: "bold"
                };

                if (params.value < 0) {
                    style.color = "red";
                }

                return style;
            }
            """)

            gb = GridOptionsBuilder.from_dataframe(df_summary)

            # Make ALL numeric columns bold
            for col in ["As DSP", "As SSP", "Offset $ USD"]:
                gb.configure_column(
                    col,
                    type=["numericColumn"],
                    valueFormatter=currency_formatter,
                    editable=False,
                    cellStyle=offset_style if col == "Offset $ USD" else JsCode("""
                        function(params) {

                            if (params.data.Month === "Total") {
                                return {
                                    backgroundColor: '#003366',
                                    color: 'white',
                                    fontWeight: 'bold'
                                };
                            }

                            return { fontWeight: 'bold' };
                        }
                    """)
                )

            # Footer style
            gridOptions = gb.build()

            gridOptions["getRowStyle"] = JsCode("""
            function(params) {
                if (params.data.Month === "Total") {
                    return {
                        backgroundColor: '#003366',
                        color: 'white',
                        fontWeight: 'bold'
                    };
                }
            }
            """)

            custom_css = {
                ".ag-header": {
                    "background-color": "#003366 !important",
                    "color": "white !important",
                    "font-weight": "bold !important"
                }
            }

            AgGrid(
                df_summary,
                gridOptions=gridOptions,
                allow_unsafe_jscode=True,
                fit_columns_on_grid_load=True,
                height=300,
                custom_css=custom_css,
                update_mode=GridUpdateMode.NO_UPDATE
            )

    # ======================================================
    # 🟨 PART 2
    # ======================================================
    with row1_col2:
        st.subheader("Part 2")
        st.info("Under Development")

    # ======================================================
    # 🟩 PART 3
    # ======================================================
    with row2_col1:
        st.subheader("Part 3")
        st.info("Under Development")

    # ======================================================
    # 🟥 PART 4
    # ======================================================
    with row2_col2:
        st.subheader("Part 4")
        st.info("Under Development")

# ====================================================
# 4️⃣ DSP (CUSTOMERS) TAB  (100% SSP CLONE)
# ====================================================

with tabs[3]:

    st.header("📤 DSP (Customers)")
    
    fy_list = generate_financial_years()

    col1, col2, col3 = st.columns(3)

    with col1:
        selected_fy = st.selectbox(
            "Financial Year",
            options=["All"] + fy_list,
            index=0,
            key="dsp_fy"
        )

    with col2:
        month_options = ["All"]

        if selected_fy != "All":
            fy_start = int(selected_fy.split("-")[0])
            months = pd.date_range(
                start=f"{fy_start}-04-01",
                end=f"{fy_start+1}-03-31",
                freq="MS"
            )
            month_options += months.strftime("%b-%Y").tolist()

        selected_month = st.selectbox(
            "Month",
            options=month_options,
            index=0,
            disabled=False,
            key="dsp_month"
        )

    with col3:
        selected_quarter = st.selectbox(
            "Quarter",
            options=["All", "Q1", "Q2", "Q3", "Q4"],
            index=0,
            key="dsp_quarter"
        )

    # Disable month if quarter selected
    if selected_quarter != "All":
        selected_month = "All"

    # ----------------------------------------
    # LOAD MASTER DATA
    # ----------------------------------------

    df_master = load_master_data_from_gsheet()
    df_partner = load_partner_list_from_gsheet()
    
    # 🔹 FILTER DSP CATEGORY ONLY
    df_master["C Net $"] = pd.to_numeric(
        df_master["C Net $"],
        errors="coerce"
    ).fillna(0)
    
    df_dsp = df_master[
        df_master["C Net $"] > 0
    ].copy()
    
    # ----------------------------------------
    # REBUILD USD/INR FROM PARTNER LIST
    # ----------------------------------------

    df_partner = load_partner_list_from_gsheet()

    if not df_partner.empty:

        for index, row in df_master.iterrows():

            partner_match = df_partner[
                df_partner["Short Name using in Bidscube"] == row["Partner Name"]
            ]

            if partner_match.empty:
                continue

            country = partner_match.iloc[0].get("Country", "")
            net_term = partner_match.iloc[0].get("Payment Terms", "")

            if country == "India (IN)":
                df_master.loc[index, "USD/INR"] = "INR"
            else:
                df_master.loc[index, "USD/INR"] = "USD"
                
            # NET TERM (IMPORTANT FIX)
            df_master.loc[index, "NET Term"] = net_term

    if df_master.empty:
        st.warning("No Master Data Found")
        st.stop()

    df_filtered = df_dsp.copy()

    df_filtered["Month"] = pd.to_datetime(
        df_filtered["Month"],
        errors="coerce"
    )
    
    if selected_fy != "All":
        fy_start, fy_end = get_fy_date_range(selected_fy)

        df_filtered["Month"] = pd.to_datetime(
            df_filtered["Month"],
            errors="coerce"
        )

        df_filtered = df_filtered[
            (df_filtered["Month"] >= fy_start) &
            (df_filtered["Month"] <= fy_end)
        ]

    if selected_quarter != "All" and selected_fy != "All":
        q_start, q_end = get_quarter_range(selected_fy, selected_quarter)

        df_filtered = df_filtered[
            (df_filtered["Month"] >= q_start) &
            (df_filtered["Month"] <= q_end)
        ]

    elif selected_month != "All":
        selected_month_dt = pd.to_datetime(
            selected_month,
            format="%b-%Y",
            errors="coerce"
        )

        df_filtered = df_filtered[
            df_filtered["Month"] == selected_month_dt
        ]

    if df_dsp.empty:
        st.warning("No DSP Customers Found")
        st.stop()

    # ----------------------------------------
    # BUILD DSP TABLE DATA
    # ----------------------------------------

    def calculate_due_date(month_str, net_term):
        try:
            month_dt = pd.to_datetime(month_str, format="%b-%Y")
            last_date = (month_dt + pd.offsets.MonthEnd(0)).date()

            days = 0
            if isinstance(net_term, str):
                match = re.search(r"\d+", net_term)
                if match:
                    days = int(match.group())

            due_date = last_date + pd.Timedelta(days=days + 1)
            return due_date.strftime("%d/%m/%Y")
        except:
            return ""

    sheet_name = "DSP (Customers)"

    worksheet = spreadsheet.worksheet(sheet_name)
    sheet_data = worksheet.get_all_records()
    df_sheet = pd.DataFrame(sheet_data)

    if not df_sheet.empty:

        df_sheet["Month"] = pd.to_datetime(
            df_sheet["Month"],
            errors="coerce"
        )

        # Apply SAME filters to sheet data
        df_dsp_final = df_sheet.copy()

        if selected_fy != "All":
            fy_start, fy_end = get_fy_date_range(selected_fy)
            df_dsp_final = df_dsp_final[
                (df_dsp_final["Month"] >= fy_start) &
                (df_dsp_final["Month"] <= fy_end)
            ]

        if selected_quarter != "All" and selected_fy != "All":
            q_start, q_end = get_quarter_range(selected_fy, selected_quarter)
            df_dsp_final = df_dsp_final[
                (df_dsp_final["Month"] >= q_start) &
                (df_dsp_final["Month"] <= q_end)
            ]

        elif selected_month != "All":
            selected_month_dt = pd.to_datetime(
                selected_month,
                format="%b-%Y",
                errors="coerce"
            )
            df_dsp_final = df_dsp_final[
                df_dsp_final["Month"] == selected_month_dt
            ]

        df_dsp_final["Month"] = df_dsp_final["Month"].dt.strftime("%b-%Y")

    else:
        dsp_rows = []

        for _, row in df_filtered.iterrows():

            receivable = abs(float(row["C Net $"]))

            dsp_rows.append({
                "Month": row["Month"].strftime("%b-%Y"),
				"DSP Name": row["Partner Name"],
				"Receivable $": receivable,
				"USD/INR": row.get("USD/INR", ""),
				"Due Date": calculate_due_date(row["Month"], net_term),
				"Received Date": "",
				"Received Amount $": 0.0,
				"Received In": "",
				"Shortage": receivable,
				"Reason": ""
            })

        df_dsp_final = pd.DataFrame(dsp_rows)

    # ----------------------------------------
    # SEARCH + REFRESH
    # ----------------------------------------

    col1, col2 = st.columns([3, 1])

    with col1:
        search_text = st.text_input("🔍 Search DSP", placeholder="Global search...")

    with col2:
        st.markdown("<br>", unsafe_allow_html=True)
        refresh = st.button("🔄 Refresh", key="dsp_refresh_button")

    if refresh:
        st.rerun()
    
       
    # ----------------------------------------
    # AGGRID (MASTER STYLE)
    # ----------------------------------------

    from st_aggrid import GridOptionsBuilder, AgGrid, JsCode
    from st_aggrid import GridUpdateMode

    currency_formatter = JsCode("""
    function(params) {
        if (params.value == null || params.value === '') return '';
        return '$' + parseFloat(params.value).toLocaleString(undefined, {minimumFractionDigits: 2});
    }
    """)

    shortage_getter = JsCode("""
    function(params) {
        let receivable = parseFloat(params.data["Receivable $"]) || 0;
        let received = parseFloat(params.data["Received Amount $"]) || 0;
        return receivable - received;
    }
    """)

    gb = GridOptionsBuilder.from_dataframe(df_dsp_final)
    
    
    from st_aggrid import JsCode   # make sure this import exists

    date_validation_js = JsCode("""
    function(params) {

        if (!params.newValue || params.newValue.trim() === "") {
            params.data["Received Date"] = "";
            return true;
        }

        let regex = /^\\d{2}\\/\\d{2}\\/\\d{4}$/;

        if (!regex.test(params.newValue)) {
            alert("Enter date in DD/MM/YYYY format");
            return false;
        }

        params.data["Received Date"] = params.newValue;
        return true;
    }
    """)

    date_editor_js = JsCode("""
    class DateEditor {

        init(params) {
            this.params = params;

            this.eInput = document.createElement('input');
            this.eInput.type = 'text';
            this.eInput.value = params.value || '';
            this.eInput.placeholder = 'DD/MM/YYYY';
            this.eInput.style.width = '100%';
            this.eInput.style.height = '100%';
            this.eInput.style.border = 'none';
            this.eInput.style.outline = 'none';
            this.eInput.style.fontSize = '14px';

            this.eInput.addEventListener('input', function(e) {
                let v = e.target.value.replace(/\\D/g,'');

                if (v.length >= 2) v = v.slice(0,2) + '/' + v.slice(2);
                if (v.length >= 5) v = v.slice(0,5) + '/' + v.slice(5);
                if (v.length > 10) v = v.slice(0,10);

                e.target.value = v;
            });
        }

        getGui() {
            return this.eInput;
        }

        afterGuiAttached() {
            this.eInput.focus();
            this.eInput.select();
        }

        getValue() {
            return this.eInput.value;
        }

        isCancelAfterEnd() {

            if (!this.eInput.value) return false;

            let regex = /^\\d{2}\\/\\d{2}\\/\\d{4}$/;
            return !regex.test(this.eInput.value);
        }
    }
    """)
    
    gb.configure_column(
        "Received Date",
        editable=True,
        singleClickEdit=True,
        cellEditor=date_editor_js,
        headerTooltip="Enter date in DD/MM/YYYY"
    )

    gb.configure_default_column(resizable=True, sortable=True)

    bold_currency_style = JsCode("""
    function(params) {

        // ❌ Do NOT apply to footer row
        if (params.node.rowPinned) {
            return {};
        }

        return {
            fontWeight: 'bold',
            color: params.value < 0 ? 'red' : 'black'
        };
    }
    """)

        
    gb.configure_column("Receivable $",
        type=["numericColumn"],
        editable=False,
        valueFormatter=currency_formatter,
        cellStyle=bold_currency_style
    )

    gb.configure_column("Received Amount $",
        type=["numericColumn"],
        editable=True,
        valueFormatter=currency_formatter,
        cellStyle=bold_currency_style
    )

    gb.configure_column("Shortage",
                        type=["numericColumn"],
                        editable=False,
                        valueGetter=shortage_getter,
                        valueFormatter=currency_formatter)

    received_in_editor_js = JsCode("""
    class ReceivedInEditor {

        init(params) {
            this.params = params;

            this.eSelect = document.createElement('select');
            this.eSelect.style.width = '100%';
            this.eSelect.style.height = '100%';
            this.eSelect.style.border = 'none';
            this.eSelect.style.outline = 'none';
            this.eSelect.style.fontSize = '14px';

            let placeholder = document.createElement('option');
            placeholder.value = '';
            placeholder.text = 'Select';
            placeholder.disabled = true;
            placeholder.selected = !params.value;
            this.eSelect.appendChild(placeholder);

            let values = ["Bank Remittance", "PayPal", "Payoneer", "Other"];

            values.forEach(function(val) {
                let option = document.createElement('option');
                option.value = val;
                option.text = val;
                if (params.value === val) {
                    option.selected = true;
                }
                this.eSelect.appendChild(option);
            }.bind(this));
        }

        getGui() {
            return this.eSelect;
        }

        afterGuiAttached() {
            this.eSelect.focus();
        }

        getValue() {
            return this.eSelect.value;
        }
    }
    """)
    
    gb.configure_column(
        "Received In",
        editable=True,
        singleClickEdit=True,
        cellEditor=received_in_editor_js,
        cellStyle=JsCode("""
            function(params) {

                if (params.node.rowPinned) return {};

                if (!params.value) {
                    return {
                        color: '#bfbfbf',
                        fontStyle: 'italic'
                    };
                }

                return {
                    color: 'black',
                    fontStyle: 'normal'
                };
            }
        """)
    )

    gb.configure_column("Reason", editable=True)

    # Freeze first 2 columns
    gb.configure_column("Month", pinned="left")
    gb.configure_column("DSP Name", pinned="left")

    # Footer Total
    total_js = JsCode("""
    function(api) {
        let totalReceivable = 0;
        let totalReceived = 0;

        api.forEachNodeAfterFilter(function(node) {
            totalReceivable += parseFloat(node.data["Receivable $"]) || 0;
            totalReceived += parseFloat(node.data["Received Amount $"]) || 0;
        });

        api.setPinnedBottomRowData([{
            "Month": "Grand Total",
            "Receivable $": Number(totalReceivable),
            "Received Amount $": Number(totalReceived),
            "Shortage": Number(totalReceivable - totalReceived)
        }]);
    }
    """)

    gridOptions = gb.build()
    gridOptions["onFirstDataRendered"] = total_js
    gridOptions["onCellValueChanged"] = total_js
    gridOptions["onFilterChanged"] = total_js
    
       
    # -------- GRAND TOTAL --------
    total_receivable = float(df_dsp_final["Receivable $"].sum())
    total_received = float(df_dsp_final["Received Amount $"].sum())

    grand_total_row = {
        "Month": "Grand Total",
        "Receivable $": float(total_receivable),
        "Received Amount $": float(total_received),
        "Shortage": float(total_receivable - total_received)
    }

    gridOptions["pinnedBottomRowData"] = [grand_total_row]

    gridOptions["getRowStyle"] = JsCode("""
    function(params) {

        if (params.node.rowPinned) {
            return {
                backgroundColor: '#003366',
                color: 'white',
                fontWeight: 'bold',
                fontSize: '14px'
            };
        }

        let receivable = parseFloat(params.data["Receivable $"]) || 0;
        let received = parseFloat(params.data["Received Amount $"]) || 0;
        let dueDateStr = params.data["Due Date"];

        if (!dueDateStr) return {};

        let parts = dueDateStr.split("/");
        if (parts.length !== 3) return {};

        let dueDate = new Date(parts[2], parts[1] - 1, parts[0]);

        let today = new Date();
        today.setHours(0,0,0,0);

        if (today > dueDate && received === 0) {
            return { backgroundColor: "#fdecea" };
        }

        if (received === receivable && receivable !== 0) {
            return { backgroundColor: "#90ee90" };
        }

        if (received !== 0 && received !== receivable) {
            return { backgroundColor: "#fff8e1" };
        }

        return {};
    }
    """)

    if search_text:
        gridOptions["quickFilterText"] = search_text

    custom_css = {
        ".ag-header": {
            "background-color": "#003366 !important",
            "color": "white !important",
            "font-weight": "bold !important"
        }
    }

    # Ensure JSON serializable types
    df_grid = df_dsp_final.copy()

    for col in df_grid.columns:
        if pd.api.types.is_numeric_dtype(df_grid[col]):
            df_grid[col] = df_grid[col].apply(
                lambda x: float(x) if pd.notnull(x) else None
            )

    df_grid = df_grid.where(pd.notnull(df_grid), None)

    grid_response = AgGrid(
        df_grid,
        gridOptions=gridOptions,
        allow_unsafe_jscode=True,
        update_mode=GridUpdateMode.MODEL_CHANGED,
        data_return_mode="AS_INPUT",
        reload_data=True,
        height=550,
        custom_css=custom_css
    )

    # ----------------------------------------
    # MANUAL SAVE BUTTON (FINAL STABLE)
    # ----------------------------------------

    col_left, col_center, col_right = st.columns([3, 1, 3])

    with col_center:
        save_clicked = st.button("💾 Save Changes", key="dsp_manual_save")

    if save_clicked:

        with st.spinner("Saving to Google Sheet..."):

            # 🔥 Force grid commit before reading data
            st.session_state["_force_commit"] = True

            updated_df = pd.DataFrame(grid_response["data"]).copy()
            
            # ---- SAFE NUMERIC CONVERSION ----
            updated_df["Receivable $"] = pd.to_numeric(
                updated_df["Receivable $"], errors="coerce"
            ).fillna(0)

            updated_df["Received Amount $"] = pd.to_numeric(
                updated_df["Received Amount $"], errors="coerce"
            ).fillna(0)

            # ---- SAFE DATE HANDLING (GOOGLE SAFE FORMAT) ----
            if "Received Date" in updated_df.columns:

                def normalize_date(x):

                    if pd.isna(x) or str(x).strip() == "":
                        return ""

                    # Case 1: dict from AgGrid
                    if isinstance(x, dict):
                        try:
                            year = x.get("year")
                            month = x.get("month")
                            day = x.get("date") or x.get("day")
                            return f"{int(year):04d}-{int(month):02d}-{int(day):02d}"
                        except:
                            return ""

                    # Case 2: string/ISO
                    try:
                        dt = pd.to_datetime(x, errors="coerce")
                        if pd.isna(dt):
                            return ""
                        return dt.strftime("%Y-%m-%d")   # 🔥 ISO FORMAT
                    except:
                        return ""

                updated_df["Received Date"] = updated_df["Received Date"].apply(normalize_date)

            # ---- SAFE STRING CLEAN ----
            for col in ["Received In", "Reason"]:
                updated_df[col] = (
                    updated_df[col]
                    .fillna("")
                    .astype(str)
                    .replace("nan", "")
                )

            # ---- RECALCULATE SHORTAGE ----
            updated_df["Shortage"] = (
                updated_df["Receivable $"] - updated_df["Received Amount $"]
            )

            # ---- SAVE FULL SHEET ----
            worksheet.clear()

            worksheet.update(
                [updated_df.columns.tolist()] +
                updated_df.values.tolist(),
                value_input_option="USER_ENTERED"
            )
            
            load_master_data_from_gsheet.clear()
            load_partner_list_from_gsheet.clear()
            st.rerun()

        st.success("DSP (Vendors) saved successfully ✅")
        st.write(updated_df[["DSP Name", "Received Date"]].head(10))

# ====================================================
# 5️⃣ SSP (VENDORS) TAB
# ====================================================

with tabs[4]:

    st.header("📤 SSP (Vendors)")
    
    fy_list = generate_financial_years()

    col1, col2, col3 = st.columns(3)

    with col1:
        selected_fy = st.selectbox(
            "Financial Year",
            options=["All"] + fy_list,
            index=0,
            key="ssp_fy"
        )

    with col2:
        month_options = ["All"]

        if selected_fy != "All":
            fy_start = int(selected_fy.split("-")[0])
            months = pd.date_range(
                start=f"{fy_start}-04-01",
                end=f"{fy_start+1}-03-31",
                freq="MS"
            )
            month_options += months.strftime("%b-%Y").tolist()

        selected_month = st.selectbox(
            "Month",
            options=month_options,
            index=0,
            disabled=False,
            key="ssp_month"
        )

    with col3:
        selected_quarter = st.selectbox(
            "Quarter",
            options=["All", "Q1", "Q2", "Q3", "Q4"],
            index=0,
            key="ssp_quarter"
        )

    # Disable month if quarter selected
    if selected_quarter != "All":
        selected_month = "All"

    # ----------------------------------------
    # LOAD MASTER DATA
    # ----------------------------------------

    df_master = load_master_data_from_gsheet()
    
    # ----------------------------------------
    # REBUILD USD/INR FROM PARTNER LIST
    # ----------------------------------------

    df_partner = load_partner_list_from_gsheet()

    if not df_partner.empty:

        for index, row in df_master.iterrows():

            partner_match = df_partner[
                df_partner["Short Name using in Bidscube"] == row["Partner Name"]
            ]

            if partner_match.empty:
                continue

            country = partner_match.iloc[0].get("Country", "")
            net_term = partner_match.iloc[0].get("Payment Terms", "")

            if country == "India (IN)":
                df_master.loc[index, "USD/INR"] = "INR"
            else:
                df_master.loc[index, "USD/INR"] = "USD"
                
            # NET TERM (IMPORTANT FIX)
            df_master.loc[index, "NET Term"] = net_term

    if df_master.empty:
        st.warning("No Master Data Found")
        st.stop()

    # Ensure numeric
    df_master["C Net $"] = pd.to_numeric(df_master["C Net $"], errors="coerce").fillna(0)

    # 🔹 FILTER SSP CATEGORY ONLY
    df_master["C Net $"] = pd.to_numeric(df_master["C Net $"], errors="coerce").fillna(0)
    df_ssp = df_master[df_master["C Net $"] < 0].copy()
    
    df_filtered = df_ssp.copy()

    df_filtered["Month"] = pd.to_datetime(
        df_filtered["Month"],
        errors="coerce"
    )
    
    if selected_fy != "All":
        fy_start, fy_end = get_fy_date_range(selected_fy)

        df_filtered["Month"] = pd.to_datetime(
            df_filtered["Month"],
            errors="coerce"
        )

        df_filtered = df_filtered[
            (df_filtered["Month"] >= fy_start) &
            (df_filtered["Month"] <= fy_end)
        ]

    if selected_quarter != "All" and selected_fy != "All":
        q_start, q_end = get_quarter_range(selected_fy, selected_quarter)

        df_filtered = df_filtered[
            (df_filtered["Month"] >= q_start) &
            (df_filtered["Month"] <= q_end)
        ]

    elif selected_month != "All":
        selected_month_dt = pd.to_datetime(
            selected_month,
            format="%b-%Y",
            errors="coerce"
        )

        df_filtered = df_filtered[
            df_filtered["Month"] == selected_month_dt
        ]

    if df_ssp.empty:
        st.warning("No SSP Vendors Found")
        st.stop()

    # ----------------------------------------
    # BUILD SSP TABLE DATA
    # ----------------------------------------

    def calculate_due_date(month_str, net_term):
        try:
            month_dt = pd.to_datetime(month_str, format="%b-%Y")
            last_date = (month_dt + pd.offsets.MonthEnd(0)).date()

            days = 0
            if isinstance(net_term, str):
                match = re.search(r"\d+", net_term)
                if match:
                    days = int(match.group())

            due_date = last_date + pd.Timedelta(days=days + 1)
            return due_date.strftime("%d/%m/%Y")
        except:
            return ""

    sheet_name = "SSP (Vendors)"

    worksheet = spreadsheet.worksheet(sheet_name)
    sheet_data = worksheet.get_all_records()
    df_sheet = pd.DataFrame(sheet_data)

    if not df_sheet.empty:

        df_sheet["Month"] = pd.to_datetime(
            df_sheet["Month"],
            errors="coerce"
        )

        # Apply SAME filters to sheet data
        df_ssp_final = df_sheet.copy()

        if selected_fy != "All":
            fy_start, fy_end = get_fy_date_range(selected_fy)
            df_ssp_final = df_ssp_final[
                (df_ssp_final["Month"] >= fy_start) &
                (df_ssp_final["Month"] <= fy_end)
            ]

        if selected_quarter != "All" and selected_fy != "All":
            q_start, q_end = get_quarter_range(selected_fy, selected_quarter)
            df_ssp_final = df_ssp_final[
                (df_ssp_final["Month"] >= q_start) &
                (df_ssp_final["Month"] <= q_end)
            ]

        elif selected_month != "All":
            selected_month_dt = pd.to_datetime(
                selected_month,
                format="%b-%Y",
                errors="coerce"
            )
            df_ssp_final = df_ssp_final[
                df_ssp_final["Month"] == selected_month_dt
            ]

        df_ssp_final["Month"] = df_ssp_final["Month"].dt.strftime("%b-%Y")

    else:
        ssp_rows = []

        for _, row in df_filtered.iterrows():

            payable = abs(float(row["C Net $"]))

            ssp_rows.append({
                "Month": row["Month"],
                "SSP Name": row["Partner Name"],
                "Payable $": payable,
                "USD/INR": row.get("USD/INR", ""),
                "Due Date": calculate_due_date(row["Month"], row.get("NET Term", "")),
                "Payment Date": "",
                "Paid Amount $": 0.0,
                "Paid From": "",
                "Shortage": payable,
                "Reason": ""
            })

        df_ssp_final = pd.DataFrame(ssp_rows)

    # ----------------------------------------
    # SEARCH + REFRESH
    # ----------------------------------------

    col1, col2 = st.columns([3, 1])

    with col1:
        search_text = st.text_input("🔍 Search SSP", placeholder="Global search...")

    with col2:
        st.markdown("<br>", unsafe_allow_html=True)
        refresh = st.button("🔄 Refresh", key="ssp_refresh_button")

    if refresh:
        st.rerun()
    
       
    # ----------------------------------------
    # AGGRID (MASTER STYLE)
    # ----------------------------------------

    from st_aggrid import GridOptionsBuilder, AgGrid, JsCode
    from st_aggrid import GridUpdateMode

    currency_formatter = JsCode("""
    function(params) {
        if (params.value == null || params.value === '') return '';
        return '$' + parseFloat(params.value).toLocaleString(undefined, {minimumFractionDigits: 2});
    }
    """)

    shortage_getter = JsCode("""
    function(params) {
        let payable = parseFloat(params.data["Payable $"]) || 0;
        let paid = parseFloat(params.data["Paid Amount $"]) || 0;
        return payable - paid;
    }
    """)

    gb = GridOptionsBuilder.from_dataframe(df_ssp_final)
    
    
    from st_aggrid import JsCode   # make sure this import exists

    date_validation_js = JsCode("""
    function(params) {

        if (!params.newValue || params.newValue.trim() === "") {
            params.data["Payment Date"] = "";
            return true;
        }

        let regex = /^\\d{2}\\/\\d{2}\\/\\d{4}$/;

        if (!regex.test(params.newValue)) {
            alert("Enter date in DD/MM/YYYY format");
            return false;
        }

        params.data["Payment Date"] = params.newValue;
        return true;
    }
    """)

    date_editor_js = JsCode("""
    class DateEditor {

        init(params) {
            this.params = params;

            this.eInput = document.createElement('input');
            this.eInput.type = 'text';
            this.eInput.value = params.value || '';
            this.eInput.placeholder = 'DD/MM/YYYY';
            this.eInput.style.width = '100%';
            this.eInput.style.height = '100%';
            this.eInput.style.border = 'none';
            this.eInput.style.outline = 'none';
            this.eInput.style.fontSize = '14px';

            this.eInput.addEventListener('input', function(e) {
                let v = e.target.value.replace(/\\D/g,'');

                if (v.length >= 2) v = v.slice(0,2) + '/' + v.slice(2);
                if (v.length >= 5) v = v.slice(0,5) + '/' + v.slice(5);
                if (v.length > 10) v = v.slice(0,10);

                e.target.value = v;
            });
        }

        getGui() {
            return this.eInput;
        }

        afterGuiAttached() {
            this.eInput.focus();
            this.eInput.select();
        }

        getValue() {
            return this.eInput.value;
        }

        isCancelAfterEnd() {

            if (!this.eInput.value) return false;

            let regex = /^\\d{2}\\/\\d{2}\\/\\d{4}$/;
            return !regex.test(this.eInput.value);
        }
    }
    """)
    
    gb.configure_column(
        "Payment Date",
        editable=True,
        singleClickEdit=True,
        cellEditor=date_editor_js,
        headerTooltip="Enter date in DD/MM/YYYY"
    )

    gb.configure_default_column(resizable=True, sortable=True)

    bold_currency_style = JsCode("""
    function(params) {

        // ❌ Do NOT apply to footer row
        if (params.node.rowPinned) {
            return {};
        }

        return {
            fontWeight: 'bold',
            color: params.value < 0 ? 'red' : 'black'
        };
    }
    """)

    
    
    gb.configure_column("Payable $",
        type=["numericColumn"],
        editable=False,
        valueFormatter=currency_formatter,
        cellStyle=bold_currency_style
    )

    gb.configure_column("Paid Amount $",
        type=["numericColumn"],
        editable=True,
        valueFormatter=currency_formatter,
        cellStyle=bold_currency_style
    )

    gb.configure_column("Shortage",
                        type=["numericColumn"],
                        editable=False,
                        valueGetter=shortage_getter,
                        valueFormatter=currency_formatter)

    paid_from_editor_js = JsCode("""
    class PaidFromEditor {

        init(params) {
            this.params = params;

            this.eSelect = document.createElement('select');
            this.eSelect.style.width = '100%';
            this.eSelect.style.height = '100%';
            this.eSelect.style.border = 'none';
            this.eSelect.style.outline = 'none';
            this.eSelect.style.fontSize = '14px';

            let placeholder = document.createElement('option');
            placeholder.value = '';
            placeholder.text = 'Select';
            placeholder.disabled = true;
            placeholder.selected = !params.value;
            this.eSelect.appendChild(placeholder);

            let values = ["Bank Remittance", "PayPal", "Payoneer", "Other"];

            values.forEach(function(val) {
                let option = document.createElement('option');
                option.value = val;
                option.text = val;
                if (params.value === val) {
                    option.selected = true;
                }
                this.eSelect.appendChild(option);
            }.bind(this));
        }

        getGui() {
            return this.eSelect;
        }

        afterGuiAttached() {
            this.eSelect.focus();
        }

        getValue() {
            return this.eSelect.value;
        }
    }
    """)
    
    gb.configure_column(
        "Paid From",
        editable=True,
        singleClickEdit=True,
        cellEditor=paid_from_editor_js,
        cellStyle=JsCode("""
            function(params) {

                if (params.node.rowPinned) return {};

                if (!params.value) {
                    return {
                        color: '#bfbfbf',
                        fontStyle: 'italic'
                    };
                }

                return {
                    color: 'black',
                    fontStyle: 'normal'
                };
            }
        """)
    )

    gb.configure_column("Reason", editable=True)

    # Freeze first 2 columns
    gb.configure_column("Month", pinned="left")
    gb.configure_column("SSP Name", pinned="left")

    # Footer Total
    total_js = JsCode("""
    function(api) {
        let totalPayable = 0;
        let totalPaid = 0;

        api.forEachNodeAfterFilter(function(node) {
            totalPayable += parseFloat(node.data["Payable $"]) || 0;
            totalPaid += parseFloat(node.data["Paid Amount $"]) || 0;
        });

        api.setPinnedBottomRowData([{
            "Month": "Grand Total",
            "Payable $": Number(totalPayable),
            "Paid Amount $": Number(totalPaid),
            "Shortage": Number(totalPayable - totalPaid)
        }]);
    }
    """)

    gridOptions = gb.build()
    gridOptions["onFirstDataRendered"] = total_js
    gridOptions["onCellValueChanged"] = total_js
    gridOptions["onFilterChanged"] = total_js
    
       
    # -------- GRAND TOTAL --------
    total_payable = float(df_ssp_final["Payable $"].sum())
    total_paid = float(df_ssp_final["Paid Amount $"].sum())

    grand_total_row = {
        "Month": "Grand Total",
        "Payable $": float(total_payable),
        "Paid Amount $": float(total_paid),
        "Shortage": float(total_payable - total_paid)
    }

    gridOptions["pinnedBottomRowData"] = [grand_total_row]

    gridOptions["getRowStyle"] = JsCode("""
    function(params) {

        if (params.node.rowPinned) {
            return {
                backgroundColor: '#003366',
                color: 'white',
                fontWeight: 'bold',
                fontSize: '14px'
            };
        }

        let payable = parseFloat(params.data["Payable $"]) || 0;
        let paid = parseFloat(params.data["Paid Amount $"]) || 0;
        let dueDateStr = params.data["Due Date"];

        if (!dueDateStr) return {};

        let parts = dueDateStr.split("/");
        if (parts.length !== 3) return {};

        let dueDate = new Date(parts[2], parts[1] - 1, parts[0]);

        let today = new Date();
        today.setHours(0,0,0,0);

        if (today > dueDate && paid === 0) {
            return { backgroundColor: "#fdecea" };
        }

        if (paid === payable && payable !== 0) {
            return { backgroundColor: "#90ee90" };
        }

        if (paid !== 0 && paid !== payable) {
            return { backgroundColor: "#fff8e1" };
        }

        return {};
    }
    """)

    if search_text:
        gridOptions["quickFilterText"] = search_text

    custom_css = {
        ".ag-header": {
            "background-color": "#003366 !important",
            "color": "white !important",
            "font-weight": "bold !important"
        }
    }

    # Ensure JSON serializable types
    df_grid = df_ssp_final.copy()

    for col in df_grid.columns:
        if pd.api.types.is_numeric_dtype(df_grid[col]):
            df_grid[col] = df_grid[col].apply(
                lambda x: float(x) if pd.notnull(x) else None
            )

    df_grid = df_grid.where(pd.notnull(df_grid), None)

    grid_response = AgGrid(
        df_grid,
        gridOptions=gridOptions,
        allow_unsafe_jscode=True,
        update_mode=GridUpdateMode.MODEL_CHANGED,
        data_return_mode="AS_INPUT",
        reload_data=True,
        height=550,
        custom_css=custom_css
    )

    # ----------------------------------------
    # MANUAL SAVE BUTTON (FINAL STABLE)
    # ----------------------------------------

    col_left, col_center, col_right = st.columns([3, 1, 3])

    with col_center:
        save_clicked = st.button("💾 Save Changes", key="ssp_manual_save")

    if save_clicked:

        with st.spinner("Saving to Google Sheet..."):

            # 🔥 Force grid commit before reading data
            st.session_state["_force_commit"] = True

            updated_df = pd.DataFrame(grid_response["data"]).copy()
            
            # ---- SAFE NUMERIC CONVERSION ----
            updated_df["Payable $"] = pd.to_numeric(
                updated_df["Payable $"], errors="coerce"
            ).fillna(0)

            updated_df["Paid Amount $"] = pd.to_numeric(
                updated_df["Paid Amount $"], errors="coerce"
            ).fillna(0)

            # ---- SAFE DATE HANDLING (GOOGLE SAFE FORMAT) ----
            if "Payment Date" in updated_df.columns:

                def normalize_date(x):

                    if pd.isna(x) or str(x).strip() == "":
                        return ""

                    # Case 1: dict from AgGrid
                    if isinstance(x, dict):
                        try:
                            year = x.get("year")
                            month = x.get("month")
                            day = x.get("date") or x.get("day")
                            return f"{int(year):04d}-{int(month):02d}-{int(day):02d}"
                        except:
                            return ""

                    # Case 2: string/ISO
                    try:
                        dt = pd.to_datetime(x, errors="coerce")
                        if pd.isna(dt):
                            return ""
                        return dt.strftime("%Y-%m-%d")   # 🔥 ISO FORMAT
                    except:
                        return ""

                updated_df["Payment Date"] = updated_df["Payment Date"].apply(normalize_date)

            # ---- SAFE STRING CLEAN ----
            for col in ["Paid From", "Reason"]:
                updated_df[col] = (
                    updated_df[col]
                    .fillna("")
                    .astype(str)
                    .replace("nan", "")
                )

            # ---- RECALCULATE SHORTAGE ----
            updated_df["Shortage"] = (
                updated_df["Payable $"] - updated_df["Paid Amount $"]
            )

            # ---- SAVE FULL SHEET ----
            worksheet.clear()

            worksheet.update(
                [updated_df.columns.tolist()] +
                updated_df.values.tolist(),
                value_input_option="USER_ENTERED"
            )
            
            load_master_data_from_gsheet.clear()
            load_partner_list_from_gsheet.clear()
            st.rerun()

        st.success("SSP (Vendors) saved successfully ✅")
        st.write(updated_df[["SSP Name", "Payment Date"]].head(10))

# ====================================================
# 7️⃣ LIST OF PARTNERS TAB
# ====================================================

with tabs[6]:

    st.header("🤝List of Partners")
    
    # -----------------------------
    # FILTER REQUIRED COLUMNS
    # -----------------------------
    required_columns = [
        "Agreement Start Date",
        "Short Name using in Bidscube",
        "Country",
        "Foreign / Indian Entity",
        "GSTIN",
        "Payment Terms",
        "Contact Person",
        "Email 1",
        "Finance Contact",
        "Finance Email"
    ]

    for col in required_columns:
        if col not in df_partner.columns:
            df_partner[col] = ""

    df_partner = df_partner[required_columns]

    # -----------------------------
    # SEARCH BAR
    # -----------------------------
    col1, col2 = st.columns([3, 1])

    with col1:
        search_text = st.text_input(
            "🔍 Search Partner",
            placeholder="Global search..."
        )

    with col2:
        st.markdown("<br>", unsafe_allow_html=True)
        refresh_clicked = st.button("🔄 Refresh", key="partner_refresh_button")

    if refresh_clicked:
        st.rerun()

    # -----------------------------
    # AGGRID TABLE (Master Style)
    # -----------------------------
    from st_aggrid import AgGrid, GridOptionsBuilder, JsCode

    gb = GridOptionsBuilder.from_dataframe(df_partner)

    gb.configure_default_column(
        resizable=True,
        sortable=True,
        filter=True
    )

    gb.configure_grid_options(
        domLayout="normal",
        suppressHorizontalScroll=False
    )
    
    gridOptions = gb.build()

    if search_text:
        gridOptions["quickFilterText"] = search_text

    custom_css = {
        ".ag-root-wrapper": {
            "overflow": "auto"
        },
        ".ag-body-horizontal-scroll": {
            "height": "8px"
        },
        ".ag-header": {
            "background-color": "#003366 !important",
            "color": "white !important",
            "font-weight": "bold !important",
            "font-size": "14px !important"
        },
        ".ag-header-cell-label": {
            "color": "white !important",
            "font-weight": "bold !important"
        }
    }

    AgGrid(
        df_partner,
        gridOptions=gridOptions,
        allow_unsafe_jscode=True,
        height=500,
        custom_css=custom_css
    )