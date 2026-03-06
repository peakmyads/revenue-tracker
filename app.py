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

import base64

def get_image_base64(path):
    with open(path, "rb") as img_file:
        return base64.b64encode(img_file.read()).decode()

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

@st.cache_resource
def get_gsheet_objects():
    client = get_gsheet_connection()
    spreadsheet = client.open_by_key(SPREADSHEET_ID)
    worksheets = {ws.title: ws for ws in spreadsheet.worksheets()}
    return spreadsheet, worksheets

spreadsheet, worksheets = get_gsheet_objects()

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

@st.cache_data(ttl=300)
def load_master_data_from_gsheet():
    worksheet = worksheets["Master Data"]
    data = worksheet.get_all_records()
    return pd.DataFrame(data)
    
@st.cache_data(ttl=60)
def load_partner_list_from_gsheet():
    worksheet = spreadsheet.worksheet("Partner List")
    data = worksheet.get_all_records()
    return pd.DataFrame(data)

# ADD THIS HERE
@st.cache_data(ttl=120)
def load_cost_centre():
    worksheet = spreadsheet.worksheet("Cost Centre")
    data = worksheet.get_all_records()
    return pd.DataFrame(data)

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

def clean_for_gsheet(df: pd.DataFrame) -> pd.DataFrame:
    import numpy as np
    df = df.copy()

    # Replace infinities
    df = df.replace([np.inf, -np.inf], 0)

    # Handle numeric columns
    numeric_cols = df.select_dtypes(include=["number"]).columns
    df[numeric_cols] = df[numeric_cols].fillna(0)

    # Handle non-numeric columns
    non_numeric_cols = df.select_dtypes(exclude=["number"]).columns
    df[non_numeric_cols] = df[non_numeric_cols].fillna("")

    return df

def load_dsp_sheet():
    worksheet = spreadsheet.worksheet("DSP (Customers)")
    data = worksheet.get_all_records()
    df = pd.DataFrame(data)

    if df.empty:
        return df

    # FIX: Ensure Month is datetime
    df["Month"] = pd.to_datetime(df["Month"], errors="coerce")

    df["Due Date"] = pd.to_datetime(df["Due Date"], errors="coerce", dayfirst=True)
    df["Received Date"] = pd.to_datetime(df["Received Date"], errors="coerce", dayfirst=True)

    df["Receivable $"] = pd.to_numeric(df["Receivable $"], errors="coerce").fillna(0)
    df["Received Amount $"] = pd.to_numeric(df["Received Amount $"], errors="coerce").fillna(0)

    df["Outstanding $"] = df["Receivable $"] - df["Received Amount $"]

    return df


def load_ssp_sheet():
    worksheet = spreadsheet.worksheet("SSP (Vendors)")
    data = worksheet.get_all_records()
    df = pd.DataFrame(data)

    if df.empty:
        return df
        
    # FIX: Ensure Month is datetime
    df["Month"] = pd.to_datetime(df["Month"], errors="coerce")

    df["Due Date"] = pd.to_datetime(df["Due Date"], errors="coerce", dayfirst=True)
    df["Payment Date"] = pd.to_datetime(df["Payment Date"], errors="coerce", dayfirst=True)

    df["Payable $"] = pd.to_numeric(df["Payable $"], errors="coerce").fillna(0)
    df["Paid Amount $"] = pd.to_numeric(df["Paid Amount $"], errors="coerce").fillna(0)

    df["Outstanding $"] = df["Payable $"] - df["Paid Amount $"]

    return df

# ==========================================
# CENTRAL DATA STORE (LOAD ONCE ONLY)
# ==========================================

def initialize_session_data():
    if "data_initialized" not in st.session_state:

        st.session_state.master_df = load_master_data_from_gsheet()
        st.session_state.partner_df = load_partner_list_from_gsheet()
        st.session_state.dsp_df = load_dsp_sheet()
        st.session_state.ssp_df = load_ssp_sheet()

        st.session_state.data_initialized = True

initialize_session_data()    
    
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

/* REMOVE ALL STREAMLIT DEFAULT SPACING */
html, body {
    margin: 0 !important;
    padding: 0 !important;
}

/* Remove top & bottom padding completely */
.block-container {
    padding-top: 0 !important;
    padding-bottom: 0 !important;

    /* Controlled left & right spacing */
    padding-left: 15px !important;
    padding-right: 15px !important;

    margin: 0 !important;
}

/* Remove Streamlit header + footer */
header { display: none !important; }
footer { display: none !important; }

/* Remove app container spacing */
[data-testid="stAppViewContainer"] {
    margin: 0 !important;
    padding: 0 !important;
}

/* Remove vertical block spacing */
[data-testid="stVerticalBlock"] {
    gap: 0rem !important;
}

</style>
""", unsafe_allow_html=True)

from datetime import date

def generate_financial_years():
    from datetime import date

    today = date.today()
    current_year = today.year
    current_month = today.month

    # Determine current FY start year
    if current_month >= 4:
        current_fy_start = current_year
    else:
        current_fy_start = current_year - 1

    fy_list = []

    # Always include current FY
    fy_list.append(
        f"{current_fy_start}-{str(current_fy_start + 1)[-2:]}"
    )

    # If new FY has started (April onwards),
    # add next FY only after it officially begins
    if today >= date(current_fy_start + 1, 4, 1):
        next_fy = current_fy_start + 1
        fy_list.append(
            f"{next_fy}-{str(next_fy + 1)[-2:]}"
        )

    return fy_list


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
    color: white !important;
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
    background: linear-gradient(#188BC2, #2D5DA1);
    border: 1px solid #c0c0c0 !important;
    box-shadow: 3px 3px 6px #b0b0b0, 
                -2px -2px 5px #ffffff;
    transition: all 0.2s ease-in-out;
}

/* Hover effect */
div[data-testid="stTabs"] button:hover {
    background: linear-gradient(#9932CC, #9932CC, #9932CC);
    transform: translateY(-2px);
}

/* Active Tab */
div[data-testid="stTabs"] button[aria-selected="true"] {
    background: linear-gradient(#FF5E0E, #FF8F00, #FF8F00);
    color: white !important;
    box-shadow: inset 2px 2px 6px #FF5E0E,
                inset -2px -2px 6px #FF8F00;
}

</style>
""", unsafe_allow_html=True)

st.markdown("""
<style>

/* ===== HEADER BANNER ===== */
.header-banner {
    background: linear-gradient(#475F94, #0076CE, #475F94);

    /* Controlled internal spacing */
    padding: 18px 40px;   /* 18px top/bottom, 40px left/right */

    border-radius: 0px;   /* remove rounded corners for full-width look */
    margin: 0px;          /* no outer margin */

    box-shadow: 0 4px 10px rgba(0,0,0,0.08);
}

/* Layout */
.header-container {
    display: flex;
    align-items: center;
    margin-bottom: 13px;
    gap: 20px;
}

/* Title */
.header-title {
    font-size: 48px;
    font-weight: 800;
    color: #FFFFFF;
}

/* Subtitle */
.header-sub {
    font-size: 30px;
    font-weight: 800;
    color: #FDFF00;
    margin-left: 10px;
}

</style>
""", unsafe_allow_html=True)

# -------------------------------
# DASHBOARD-STYLE COMPANY HEADER
# -------------------------------

logo_path = "peakads_logo.png"

st.markdown("""
<style>

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
    font-size: 40px;
    font-weight: 600;
    color: #FF5E0E;
    margin-left: 5px;
}

</style>
""", unsafe_allow_html=True)

st.markdown(f"""
<div class="header-banner">
    <div class="header-container">
        <img src="data:image/png;base64,{get_image_base64('peakads_logo.png')}" width="130">
        <div>
            <span class="header-title">PEAKADS LLP</span>
            <span class="header-sub">Management Tracker</span>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

tabs = st.tabs([
    "📊 Dashboard",
    "📈 Summary",
    "📁 Master Data",
    " DSP (Customers)",
    " SSP (Vendors)",
    "📝 Partner Onboarding Form",
    "🤝 List of Partners",
    "💰 Costs Centre"
])

# ====================================================
# 1️⃣ PARTNER ONBOARDING FORM
# ====================================================

with tabs[5]:

    st.header("📝Partner Onboarding Form")
    
    st.divider()

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
            opacity: 0.9;
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
        
    st.divider()

    # ---------- FULL WIDTH ROW ----------
    address = st.text_area("Registered Address", height=80)

    col4, col5 = st.columns(2)
    
    st.divider()

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

    col1, col2, col3, col4, col5 = st.columns([1.2, 1.2, 1, 1.8, 1])

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
            key="master_month"
        )

    with col3:
        selected_quarter = st.selectbox(
            "Quarter",
            options=["All", "Q1", "Q2", "Q3", "Q4"],
            index=0,
            key="master_quarter"
        )

    with col4:
        search_text = st.text_input(
            "🔍 Search",
            placeholder="Global Search...",
            key="master_search"
        )

    with col5:
        st.markdown("<br>", unsafe_allow_html=True)
        refresh_clicked = st.button(
            "🔄 Refresh",
            key="master_refresh_button"
        )
    
    if refresh_clicked:
        load_master_data_from_gsheet.clear()
        st.session_state.master_df = load_master_data_from_gsheet()
        st.rerun()

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
    
    df_partner = st.session_state.partner_df.copy()

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

        st.divider()
        
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
        
        grid_df = (
            df_master
            .reset_index(drop=True)
            .loc[:, ~df_master.columns.duplicated()]
            .copy()
        )

        grid_response = AgGrid(
            grid_df,
            gridOptions=gridOptions,
            allow_unsafe_jscode=True,
            update_mode=GridUpdateMode.VALUE_CHANGED,
            data_return_mode="AS_INPUT",
            fit_columns_on_grid_load=True,
            height=550,
            custom_css=custom_css
        )
        
                
        if grid_response["selected_rows"] is not None:
            pass  # ignore selection

        # =========================
        # AUTO SAVE ON EDIT (NO READ VERSION)
        # =========================

        if grid_response and grid_response.get("event") == "cellValueChanged":

            updated_df = pd.DataFrame(grid_response["data"]).reset_index(drop=True)

            editable_cols = ["C DSP $", "C SSP $"]

            for col in editable_cols:
                updated_df[col] = pd.to_numeric(updated_df[col], errors="coerce").fillna(0)

            previous_df = st.session_state.master_df.reset_index(drop=True)

            if not updated_df[editable_cols].equals(previous_df[editable_cols]):

                worksheet = worksheets["Master Data"]

                batch_requests = []

                for idx in range(len(updated_df)):

                    old_row = previous_df.iloc[idx]
                    new_row = updated_df.iloc[idx]

                    if (
                        float(old_row["C DSP $"]) != float(new_row["C DSP $"]) or
                        float(old_row["C SSP $"]) != float(new_row["C SSP $"])
                    ):

                        sheet_row_number = idx + 2

                        batch_requests.append({
                            "range": f"F{sheet_row_number}:H{sheet_row_number}",
                            "values": [[
                                float(new_row["C DSP $"]),
                                float(new_row["C SSP $"]),
                                float(new_row["C DSP $"]) - float(new_row["C SSP $"])
                            ]]
                        })

                if batch_requests:
                    worksheet.batch_update(batch_requests, value_input_option="USER_ENTERED")
                    st.session_state.master_df = updated_df.copy()
                    st.toast("Auto-saved ✅")
                                       
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

# ==========================================
# MODERN KPI COMPONENT
# ==========================================

import altair as alt
import pandas as pd

def render_premium_kpi(title, value, trend_data=None, is_currency=True):

    numeric_value = float(value)

    display_value = (
        f"${numeric_value:,.2f}"
        if is_currency
        else f"{numeric_value:.2f}%"
    )

    # ----------------------------------
    # 🎨 KPI COLOR LOGIC (Nature Based)
    # ----------------------------------

    title_lower = title.lower()

    # Revenue / Receivable → Green
    if "revenue" in title_lower or "receivable" in title_lower:
        bg = "linear-gradient(135deg, #0f5132, #198754)"
        text_color = "#ffffff"

    # Cost / Payable → Red
    elif "cost" in title_lower or "payable" in title_lower:
        bg = "linear-gradient(135deg, #842029, #dc3545)"
        text_color = "#ffffff"

    # Profit → Blue
    elif "profit" in title_lower:
        bg = "linear-gradient(135deg, #084298, #0d6efd)"
        text_color = "#ffffff"

    # IVT → Orange
    elif "ivt" in title_lower:
        bg = "linear-gradient(135deg, #646D7E, #3F829D)"
        text_color = "#ffffff"

    # % Metrics → Teal
    elif "%" in title_lower:
        bg = "linear-gradient(135deg, #004d40, #20c997)"
        text_color = "#ffffff"

    # Default → Grey
    else:
        bg = "linear-gradient(135deg, #343a40, #6c757d)"
        text_color = "#ffffff"

    st.markdown(f"""
    <div style="
        background: {bg};
        padding:20px;
        border-radius:16px;
        box-shadow:0 6px 18px rgba(0,0,0,0.2);
        margin-bottom:10px;
    ">
        <div style="font-size:18px;font-weight:900;color:#FFEF00;">
            {title}
        </div>
        <div style="font-size:30px;font-weight:900;color:#FFEF00;margin-top:6px;">
            {display_value}
        </div>
    </div>
    """, unsafe_allow_html=True)


# ==========================================
# CASH CONTROL ENGINE
# ==========================================

from datetime import datetime

def calculate_outstanding_metrics(dsp_df, ssp_df):
    today = pd.Timestamp.today()

    # DSP
    total_receivable = dsp_df["Receivable $"].sum()
    total_received = dsp_df["Received Amount $"].sum()
    total_outstanding_dsp = dsp_df["Outstanding $"].sum()
    overdue_dsp = dsp_df[
        (dsp_df["Outstanding $"] > 0) &
        (dsp_df["Due Date"] < today)
    ]["Outstanding $"].sum()

    # SSP
    total_payable = ssp_df["Payable $"].sum()
    total_paid = ssp_df["Paid Amount $"].sum()
    total_outstanding_ssp = ssp_df["Outstanding $"].sum()
    overdue_ssp = ssp_df[
        (ssp_df["Outstanding $"] > 0) &
        (ssp_df["Due Date"] < today)
    ]["Outstanding $"].sum()

    return {
        "total_receivable": total_receivable,
        "total_received": total_received,
        "total_outstanding_dsp": total_outstanding_dsp,
        "overdue_dsp": overdue_dsp,
        "total_payable": total_payable,
        "total_paid": total_paid,
        "total_outstanding_ssp": total_outstanding_ssp,
        "overdue_ssp": overdue_ssp
    }


def calculate_collection_efficiency(dsp_df, ssp_df):

    total_receivable = dsp_df["Receivable $"].sum()
    total_received = dsp_df["Received Amount $"].sum()

    total_payable = ssp_df["Payable $"].sum()
    total_paid = ssp_df["Paid Amount $"].sum()

    collection_pct = (
        (total_received / total_receivable) * 100
        if total_receivable != 0 else 0
    )

    payment_pct = (
        (total_paid / total_payable) * 100
        if total_payable != 0 else 0
    )

    return {
        "collection_pct": collection_pct,
        "payment_pct": payment_pct
    }

with tabs[0]:
    
    st.markdown("""
    <style>

    /* ===== MAKE SUBTABS STICKY ===== */
    div[data-testid="stTabs"] > div:first-child {
        position: sticky;
        top: 0;
        z-index: 999;
        background: white;
        padding-top: 5px;
    }

    /* Prevent content overlapping */
    div[data-testid="stTabs"] {
        background: white;
    }

    </style>
    """, unsafe_allow_html=True)

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
        
    st.divider()

    # Disable month if quarter selected
    if selected_quarter != "All":
        selected_month = "All"

    df_master = st.session_state.master_df.copy()    
    
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
        "🏆 Top 10 Partners",
        "👥 Partner Onboarded"
    ])

    with subtabs[0]:

        # ==========================================
        # LOAD CASH DATA FOR DASHBOARD
        # ==========================================

        dsp_df = st.session_state.dsp_df.copy()
        ssp_df = st.session_state.ssp_df.copy()

        # Default fallback to prevent NameError
        outstanding_metrics = {
            "total_outstanding_dsp": 0,
            "total_outstanding_ssp": 0,
            "net_working_capital": 0
        }

        efficiency_metrics = {
            "collection_pct": 0,
            "payment_pct": 0
        }

        if not dsp_df.empty and not ssp_df.empty:
            outstanding_metrics = calculate_outstanding_metrics(dsp_df, ssp_df)
            efficiency_metrics = calculate_collection_efficiency(dsp_df, ssp_df)
        
        st.markdown("## 📊 Financial Overview")
        st.divider()

        # ===== ROW 1 =====
        r1c1, r1c2, r1c3, r1c4, r1c5, r1c6 = st.columns(6)

        with r1c1:
            render_premium_kpi("Revenue", total_dsp, df_master["C DSP $"].tolist())

        with r1c2:
            render_premium_kpi("Cost", total_ssp, df_master["C SSP $"].tolist())

        with r1c3:
            render_premium_kpi("Gross Profit", total_c_net, df_master["C Net $"].tolist())
            
        with r1c4:
            render_premium_kpi("Gross Profit %", c_profit_percent, is_currency=False)
        
        with r1c5:
            render_premium_kpi("IVT $", ivt)

        with r1c6:
            render_premium_kpi("IVT %", ivt_percent, is_currency=False)

        st.markdown("<br>", unsafe_allow_html=True)
        st.divider()

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
                
        # ==========================================
        # CASH CONTROL KPIs
        # ==========================================

        dsp_df = st.session_state.dsp_df.copy()
        ssp_df = st.session_state.ssp_df.copy()

        if not dsp_df.empty and not ssp_df.empty:

            outstanding_metrics = calculate_outstanding_metrics(dsp_df, ssp_df)
            efficiency_metrics = calculate_collection_efficiency(dsp_df, ssp_df)

            st.markdown("### 💰 Cash Control Overview")
            st.divider()

            c1, c2, c3, c4, c5, c6 = st.columns(6)

            with c1:
                render_premium_kpi("Outstanding Receivable", outstanding_metrics["total_outstanding_dsp"])

            with c2:
                render_premium_kpi("Overdue Receivable", outstanding_metrics["overdue_dsp"])
                
            with c3:
                render_premium_kpi("Collection Efficiency %", efficiency_metrics["collection_pct"], is_currency=False)
            
            with c4:
                render_premium_kpi("Outstanding Payable", outstanding_metrics["total_outstanding_ssp"])

            with c5:
                render_premium_kpi("Overdue Payable", outstanding_metrics["overdue_ssp"])

            with c6:
                render_premium_kpi("Payment Efficiency %", efficiency_metrics["payment_pct"], is_currency=False)
                
        st.divider()
                
        # ===== ROW 2 =====
        r2c0, r2c1, r2c2, r2c3, r2c4 = st.columns(5)

        with r2c0:
            render_premium_kpi("Gross Profit (INR)", outstanding_metrics["total_outstanding_dsp"])
        
        with r2c1:
            render_premium_kpi("Direct Cost (INR)", outstanding_metrics["total_outstanding_dsp"])

        with r2c2:
            render_premium_kpi("Indirect Cost (INR)", outstanding_metrics["total_outstanding_ssp"])

        with r2c3:
            render_premium_kpi("Net Profit (INR)", efficiency_metrics["collection_pct"], is_currency=False)

        with r2c4:
            render_premium_kpi("Net Profit %", efficiency_metrics["payment_pct"], is_currency=False)

        st.markdown('</div>', unsafe_allow_html=True)
        
        st.divider()

    with subtabs[1]:    
        
        import altair as alt

        st.markdown("### 📈Monthly Revenue Trend")
        st.divider()

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

        # ---- TOP 10 PARTNERS ----
        top10 = (
            df_master.groupby("Partner Name", as_index=False)
            .agg({"C Net $": "sum"})
        )

        # Sort highest to lowest
        top10 = top10.sort_values("C Net $", ascending=False)

        # Take top 10
        top10 = top10.head(10)

        # ---- ALTAIR BAR CHART ----
        chart = (
            alt.Chart(top10)
            .mark_bar()
            .encode(
                x=alt.X(
                    "C Net $:Q",
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
                    alt.Tooltip("C Net $:Q", format=",.2f")
                ]
            )
            .properties(height=400)
        )

        st.markdown("### 🏆Top 10 Partners")
        st.divider()
        st.altair_chart(chart, use_container_width=True)
        
                        
    with subtabs[3]:

        st.markdown("### 👥 Partner Onboarded Overview")

        df_partner = st.session_state.partner_df.copy()

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
        
        # Apply Dashboard Filters

        if selected_fy != "All":
            fy_start, fy_end = get_fy_date_range(selected_fy)

            df_partner = df_partner[
                (df_partner["Agreement Start Date"] >= fy_start) &
                (df_partner["Agreement Start Date"] <= fy_end)
            ]

        if selected_quarter != "All" and selected_fy != "All":
            q_start, q_end = get_quarter_range(selected_fy, selected_quarter)

            df_partner = df_partner[
                (df_partner["Agreement Start Date"] >= q_start) &
                (df_partner["Agreement Start Date"] <= q_end)
            ]

        elif selected_month != "All":
            selected_month_dt = pd.to_datetime(selected_month, format="%b-%Y")

            df_partner = df_partner[
                df_partner["Agreement Start Date"].dt.to_period("M")
                == selected_month_dt.to_period("M")
            ]

        df_partner = df_partner.dropna(subset=["Agreement Start Date"])

        # Create Month column
        # Convert Start Date
        df_partner["Agreement Start Date"] = pd.to_datetime(
            df_partner["Agreement Start Date"],
            errors="coerce"
        )

        df_partner = df_partner.dropna(subset=["Agreement Start Date"])

        # Create Month column aligned to first day of month
        df_partner["Month"] = df_partner["Agreement Start Date"].dt.to_period("M").dt.to_timestamp()

        # Monthly counts
        monthly_counts = (
            df_partner.groupby("Month", as_index=False)
            .size()
            .rename(columns={"size": "Partner Count"})
        )

        monthly_counts = monthly_counts.sort_values("Month")
        monthly_counts = monthly_counts.sort_values("Month")

        monthly_counts["Label"] = monthly_counts["Month"].dt.strftime("%b-%Y")

        st.markdown("### Total Partners Onboarded (Month-wise)")

        chart = (
            alt.Chart(monthly_counts)
            .mark_bar()
            .encode(
                x=alt.X(
                    "Label:N",
                    sort=list(monthly_counts["Label"]),
                    title="Month"
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

        df_master = st.session_state.master_df.copy()

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
        
        st.divider()

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

            # ----------------------------------------
            # EXCLUSION LOGIC (NO GOOGLE CALLS)
            # ----------------------------------------

            dsp_df = st.session_state.dsp_df.copy()
            ssp_df = st.session_state.ssp_df.copy()

            excluded_dsp = set()
            excluded_ssp = set()

            # DSP exclusion
            if not dsp_df.empty:
                dsp_filtered = dsp_df[
                    dsp_df["DSP Name"] == selected_partner
                ].copy()

                dsp_filtered["MonthStr"] = dsp_filtered["Month"].dt.strftime("%b-%Y")

                green_dsp = dsp_filtered[
                    dsp_filtered["Received Amount $"] == dsp_filtered["Receivable $"]
                ]

                yellow_dsp = dsp_filtered[
                    (dsp_filtered["Received Amount $"] != 0) &
                    (dsp_filtered["Received Amount $"] != dsp_filtered["Receivable $"])
                ]

                excluded_dsp = set(
                    pd.concat([green_dsp, yellow_dsp])["MonthStr"]
                )

            # SSP exclusion
            if not ssp_df.empty:
                ssp_filtered = ssp_df[
                    ssp_df["SSP Name"] == selected_partner
                ].copy()

                ssp_filtered["MonthStr"] = ssp_filtered["Month"].dt.strftime("%b-%Y")

                green_ssp = ssp_filtered[
                    ssp_filtered["Paid Amount $"] == ssp_filtered["Payable $"]
                ]

                yellow_ssp = ssp_filtered[
                    (ssp_filtered["Paid Amount $"] != 0) &
                    (ssp_filtered["Paid Amount $"] != ssp_filtered["Payable $"])
                ]

                excluded_ssp = set(
                    pd.concat([green_ssp, yellow_ssp])["MonthStr"]
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

            month_comparator = JsCode("""
            function(date1, date2) {

                function parseMonth(str){
                    if(!str) return new Date(0);

                    const [mon, year] = str.split("-");
                    return new Date(mon + " 1, " + year);
                }

                const d1 = parseMonth(date1);
                const d2 = parseMonth(date2);

                return d1 - d2;
            }
            """)
            
            gb = GridOptionsBuilder.from_dataframe(df_summary)
            
            gb.configure_column(
                "Month",
                comparator=month_comparator
            )

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

    col1, col2, col3, col4, col5 = st.columns([1.2, 1.2, 1, 2, 1])

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
            key="dsp_month"
        )

    with col3:
        selected_quarter = st.selectbox(
            "Quarter",
            options=["All", "Q1", "Q2", "Q3", "Q4"],
            index=0,
            key="dsp_quarter"
        )

    with col4:
        search_text = st.text_input(
            "🔍 Search DSP",
            placeholder="Global search...",
            key="dsp_search"
        )

    with col5:
        st.markdown("<br>", unsafe_allow_html=True)
        refresh = st.button(
            "🔄 Refresh",
            key="dsp_refresh_button"
        )
        
    if refresh:
        st.rerun()   

    # Disable month if quarter selected
    if selected_quarter != "All":
        selected_month = "All"

    # ----------------------------------------
    # LOAD MASTER DATA
    # ----------------------------------------

    df_master = st.session_state.master_df.copy()
    df_partner = st.session_state.partner_df.copy()
    
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

    df_partner = st.session_state.partner_df.copy()

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
        
    st.divider()

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
    
    gb.configure_column(
        "Month",
        comparator=month_comparator
    )
    for col in df_dsp_final.columns:
        gb.configure_column(col, flex=1)
    
    
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

    from st_aggrid import JsCode

    date_cell_editor = JsCode("""
    class DatePickerEditor {
        init(params) {
            this.eInput = document.createElement('input');
            this.eInput.type = 'date';
            this.eInput.style.width = '100%';
            this.eInput.style.height = '100%';
            this.eInput.style.border = 'none';
            this.eInput.style.outline = 'none';

            if (params.value) {
                let parts = params.value.split('/');
                if (parts.length === 3) {
                    this.eInput.value = parts[2] + '-' + parts[1] + '-' + parts[0];
                }
            }
        }

        getGui() {
            return this.eInput;
        }

        afterGuiAttached() {
            this.eInput.focus();
        }

        getValue() {
            if (!this.eInput.value) return '';

            let date = new Date(this.eInput.value);
            let day = String(date.getDate()).padStart(2, '0');
            let month = String(date.getMonth() + 1).padStart(2, '0');
            let year = date.getFullYear();

            return day + '/' + month + '/' + year;
        }
    }
    """)
    
    gb.configure_column(
        "Received Date",
        editable=True,
        singleClickEdit=True,
        cellEditor=date_cell_editor
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
    gridOptions["domLayout"] = "normal"
    gridOptions["suppressHorizontalScroll"] = True
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
        fit_columns_on_grid_load=False,
        reload_data=True,
        height=550,
        width="100%",
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

        st.success("DSP (Customers) saved successfully ✅")
        st.write(updated_df[["DSP Name", "Received Date"]].head(10))

# ====================================================
# 5️⃣ SSP (VENDORS) TAB
# ====================================================

with tabs[4]:

    st.header("📤 SSP (Vendors)")
    
    fy_list = generate_financial_years()

    col1, col2, col3, col4, col5 = st.columns([1.2, 1.2, 1, 2, 1])

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
            key="ssp_month"
        )

    with col3:
        selected_quarter = st.selectbox(
            "Quarter",
            options=["All", "Q1", "Q2", "Q3", "Q4"],
            index=0,
            key="ssp_quarter"
        )

    with col4:
        search_text = st.text_input(
            "🔍 Search SSP",
            placeholder="Global search...",
            key="ssp_search"
        )

    with col5:
        st.markdown("<br>", unsafe_allow_html=True)
        refresh = st.button(
            "🔄 Refresh",
            key="ssp_refresh_button"
        )
        
    if refresh:
        st.rerun()

    # Disable month if quarter selected
    if selected_quarter != "All":
        selected_month = "All"

    # ----------------------------------------
    # LOAD MASTER DATA
    # ----------------------------------------

    df_master = st.session_state.master_df.copy()
    
    # ----------------------------------------
    # REBUILD USD/INR FROM PARTNER LIST
    # ----------------------------------------

    df_partner = st.session_state.partner_df.copy()

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
        
    st.divider()

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
    gb.configure_column(
        "Month",
        comparator=month_comparator
    )

    for col in df_ssp_final.columns:
        gb.configure_column(col, flex=1)
        
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

    date_cell_editor = JsCode("""
    class DatePickerEditor {
        init(params) {
            this.eInput = document.createElement('input');
            this.eInput.type = 'date';
            this.eInput.style.width = '100%';
            this.eInput.style.height = '100%';
            this.eInput.style.border = 'none';
            this.eInput.style.outline = 'none';

            if (params.value) {
                let parts = params.value.split('/');
                if (parts.length === 3) {
                    this.eInput.value = parts[2] + '-' + parts[1] + '-' + parts[0];
                }
            }
        }

        getGui() {
            return this.eInput;
        }

        afterGuiAttached() {
            this.eInput.focus();
        }

        getValue() {
            if (!this.eInput.value) return '';

            let date = new Date(this.eInput.value);
            let day = String(date.getDate()).padStart(2, '0');
            let month = String(date.getMonth() + 1).padStart(2, '0');
            let year = date.getFullYear();

            return day + '/' + month + '/' + year;
        }
    }
    """)
    
    gb.configure_column(
        "Payment Date",
        editable=True,
        singleClickEdit=True,
        cellEditor=date_cell_editor
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
    gridOptions["domLayout"] = "normal"
    gridOptions["suppressHorizontalScroll"] = True
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
        fit_columns_on_grid_load=False,
        reload_data=True,
        height=550,
        width="100%",
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
    
    df_partner["Agreement Start Date"] = pd.to_datetime(
        df_partner["Agreement Start Date"],
        errors="coerce",
        dayfirst=True
    )
    
    df_partner["Agreement Start Date"] = df_partner["Agreement Start Date"].dt.strftime("%d-%b-%Y")
    
    # -----------------------------
    # RENAME FOR UI DISPLAY ONLY
    # -----------------------------
    df_display = df_partner.rename(columns={
        "Short Name using in Bidscube": "Partner",
        "Agreement Start Date": "Start Date",
        "Foreign / Indian Entity": "Foreign / Indian"
    })
    
    gb = GridOptionsBuilder.from_dataframe(df_display)

    date_comparator = JsCode("""
    function(date1, date2) {
        if (!date1) return -1;
        if (!date2) return 1;

        const d1 = new Date(date1);
        const d2 = new Date(date2);

        return d1 - d2;
    }
    """)

    gb.configure_column(
        "Start Date",
        comparator=date_comparator
    )
    
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
        
    st.divider()

    # -----------------------------
    # AGGRID TABLE (Master Style)
    # -----------------------------
    from st_aggrid import AgGrid, GridOptionsBuilder, JsCode

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
        df_display,
        gridOptions=gridOptions,
        allow_unsafe_jscode=True,
        height=500,
        custom_css=custom_css
    )
    
# ====================================================
# 💰 DIRECT & INDIRECT COST TAB - FINAL STABLE
# ====================================================

with tabs[7]:

    st.header("💰 Costs Centre")

    col1, col2, col3 = st.columns([1.2, 1.2, 1])

    with col1:
        selected_fy = st.selectbox(
            "Financial Year",
            options=["All"] + fy_list,
            index=0,
            key="cost_fy_filter"
        )
        
    with col2:

        # -------------------------
        # ADD COST BUTTON
        # -------------------------

        if "open_cost_popup" not in st.session_state:
            st.session_state.open_cost_popup = False
        
        if st.button("➕ Add Cost"):
            st.session_state.open_cost_popup = True

        if st.session_state.open_cost_popup:
            
            if "reset_cost_form" not in st.session_state:
                st.session_state.reset_cost_form = False
            
            @st.dialog("Add New Cost")
            
            def add_cost_popup():
                
                if "reset_cost_form" not in st.session_state:
                    st.session_state.reset_cost_form = False

                if st.session_state.reset_cost_form:

                    st.session_state.cost_category = "Select"
                    st.session_state.cost_name = "Select"
                    st.session_state.sub_cost = "Select"

                    st.session_state.new_cost_name = ""
                    st.session_state.new_sub_cost = ""

                    st.session_state.cost_month = "Select"
                    st.session_state.cost_currency = "Select"

                    st.session_state.amount_usd = 0.0
                    st.session_state.fx_rate = 0.0
                    st.session_state.amount_inr = 0.0

                    st.session_state.reset_cost_form = False

                st.markdown("### Add Cost Details")
                
                cost_df = load_cost_centre()

                # -------------------------
                # CATEGORY
                # -------------------------

                category = st.selectbox(
                    "Category",
                    ["Select", "Direct", "Indirect"],
                    key="cost_category"
                )

                # -------------------------
                # COST NAME
                # -------------------------

                worksheet = spreadsheet.worksheet("Cost Centre")

                existing_data = worksheet.get_all_records()

                cost_names = sorted(
                    list(set([r["Cost Name"] for r in existing_data if r["Cost Name"]]))
                )

                if not cost_df.empty and category != "Select":

                    cost_names = sorted(
                        cost_df[cost_df["Category"] == category]["Cost Name"]
                        .dropna()
                        .unique()
                        .tolist()
                    )

                else:
                    cost_names = []

                cost_name = st.selectbox(
                    "Cost Name",
                    ["Select"] + cost_names,
                    key="cost_name"
                )

                new_cost_name = st.text_input(
                    "Add New Cost Name (optional)",
                    key="new_cost_name"
                )

                if new_cost_name:
                    cost_name = new_cost_name

                # -------------------------
                # SUB COST
                # -------------------------

                sub_cost_list = sorted(
                    list(set([r["Sub Cost"] for r in existing_data if r["Sub Cost"]]))
                )

                if not cost_df.empty and cost_name != "Select":

                    sub_cost_list = sorted(
                        cost_df[
                            (cost_df["Category"] == category) &
                            (cost_df["Cost Name"] == cost_name)
                        ]["Sub Cost"]
                        .dropna()
                        .unique()
                        .tolist()
                    )

                else:
                    sub_cost_list = []

                sub_cost = st.selectbox(
                    "Sub Cost",
                    ["Select"] + sub_cost_list,
                    key="sub_cost"
                )

                new_sub_cost = st.text_input(
                    "Add New Sub Cost (optional)",
                    key="new_sub_cost"
                )

                if new_sub_cost:
                    sub_cost = new_sub_cost

                # -------------------------
                # FINANCIAL YEAR
                # -------------------------

                from datetime import date

                today = date.today()

                if today.month >= 4:
                    current_fy_start = today.year
                else:
                    current_fy_start = today.year - 1

                fy_list = [
                    f"{current_fy_start}-{str(current_fy_start+1)[-2:]}"
                ]

                if today >= date(current_fy_start + 1, 4, 1):
                    next_fy = current_fy_start + 1
                    fy_list.append(
                        f"{next_fy}-{str(next_fy+1)[-2:]}"
                    )

                financial_year = st.selectbox(
                    "Financial Year",
                    fy_list,
                    index=0,
                    key="cost_fy"
                )

                # -------------------------
                # MONTH LIST
                # -------------------------

                import pandas as pd

                fy_start = int(financial_year.split("-")[0])

                months = pd.date_range(
                    start=f"{fy_start}-04-01",
                    end=pd.Timestamp.today(),
                    freq="MS"
                )

                month_list = months.strftime("%b-%Y").tolist()

                month = st.selectbox(
                    "Month",
                    ["Select"] + month_list,
                    key="cost_month"
                )

                # -------------------------
                # CURRENCY
                # -------------------------

                currency = st.selectbox(
                    "USD / INR",
                    ["Select", "USD", "INR"]
                )

                amount_usd = 0
                fx_rate = 0
                amount_inr = 0

                if currency == "USD":

                    amount_usd = st.number_input(
                        "Amount $",
                        min_value=0.0,
                        step=0.01,
                        key="amount_usd"
                    )

                    fx_rate = st.number_input(
                        "FX Rate",
                        min_value=0.0,
                        step=0.01,
                        key="fx_rate"
                    )

                elif currency == "INR":

                    amount_inr = st.number_input(
                        "Amount ₹",
                        min_value=0.0,
                        step=0.01,
                        key="amount_inr"
                    )

                st.divider()

                # -------------------------
                # SAVE BUTTON
                # -------------------------

                if st.button("Save Cost"):

                    worksheet = spreadsheet.worksheet("Cost Centre")
                    
                    existing = pd.DataFrame(
                        worksheet.get_all_records()
                    )

                    duplicate = existing[
                        (existing["Category"] == category) &
                        (existing["Cost Name"] == cost_name) &
                        (existing["Sub Cost"] == sub_cost) &
                        (existing["Financial Year"] == financial_year) &
                        (existing["Month"] == month)
                    ]

                    if not duplicate.empty:
                        st.error("This cost already exists for the selected month.")
                        st.stop()

                    headers = [
                        "Category",
                        "Cost Name",
                        "Sub Cost",
                        "Financial Year",
                        "Month",
                        "Currency",
                        "Amount USD",
                        "FX Rate",
                        "Amount INR"
                    ]

                    row = [
                        category,
                        cost_name,
                        sub_cost,
                        financial_year,
                        month,
                        currency,
                        amount_usd,
                        fx_rate,
                        amount_inr
                    ]

                    first_row = worksheet.row_values(1)

                    if first_row != headers:
                        worksheet.update("A1:I1", [headers])

                    worksheet.append_row(row, value_input_option="USER_ENTERED")
                    
                    load_cost_centre.clear()

                    st.success("Cost Saved Successfully")

                    st.session_state.reset_cost_form = True
                    
                    # FORCE TABLE REFRESH
                    st.session_state.cost_table_refresh += 1

                    st.rerun()
                    
            add_cost_popup()
            
    # --------------------------------
    # COST TABLE REFRESH STATE
    # --------------------------------

    if "cost_table_refresh" not in st.session_state:
        st.session_state.cost_table_refresh = 0
        
    # ================================
    # COST CENTRE TABLE
    # ================================

    st.divider()

    st.subheader("Cost Centre Summary")

    fy_list = generate_financial_years()

    df_cost = load_cost_centre()

    if df_cost.empty:
        st.info("No Cost Data Found")
        st.stop()

    # --------------------------------
    # FILTER BY FY
    # --------------------------------

    if selected_fy != "All":
        df_cost = df_cost[df_cost["Financial Year"] == selected_fy]

    if df_cost.empty:
        st.info("No data for selected FY")
        st.stop()

    # --------------------------------
    # HANDLE FINANCIAL YEAR
    # --------------------------------

    if selected_fy == "All":
        st.info("Please select a Financial Year to view Cost Centre")
        st.stop()

    start_year = int(selected_fy.split("-")[0])

    months = pd.date_range(
        start=f"{start_year}-04-01",
        end=f"{start_year+1}-03-31",
        freq="MS"
    )

    month_cols = [m.strftime("%b-%Y") for m in months]

    # --------------------------------
    # CREATE PARTICULARS
    # --------------------------------

    df_cost["Particulars"] = df_cost["Cost Name"] + " - " + df_cost["Sub Cost"]

    # =====================================================
    # DIRECT COST
    # =====================================================

    direct_df = df_cost[df_cost["Category"] == "Direct"]

    direct_usd = direct_df[direct_df["Currency"] == "USD"]

    usd_pivot = (
        direct_usd
        .pivot_table(
            index="Particulars",
            columns="Month",
            values="Amount USD",
            aggfunc="sum",
            fill_value=0
        )
    )

    usd_pivot = usd_pivot.reindex(columns=month_cols, fill_value=0)
    usd_pivot.insert(0, "Currency", "USD")
    usd_pivot.reset_index(inplace=True)

    total_usd = usd_pivot[month_cols].sum()

    fx_rate = (
        direct_usd
        .pivot_table(
            index="Month",
            values="FX Rate",
            aggfunc="mean"
        )
    )

    fx_rate = fx_rate.reindex(month_cols, fill_value=0)["FX Rate"]

    direct_inr_from_usd = total_usd * fx_rate

    direct_inr = direct_df[direct_df["Currency"] == "INR"]

    inr_pivot = (
        direct_inr
        .pivot_table(
            index="Particulars",
            columns="Month",
            values="Amount INR",
            aggfunc="sum",
            fill_value=0
        )
    )

    inr_pivot = inr_pivot.reindex(columns=month_cols, fill_value=0)
    inr_pivot.insert(0, "Currency", "INR")
    inr_pivot.reset_index(inplace=True)

    direct_inr_total = inr_pivot[month_cols].sum()

    total_direct_inr = direct_inr_total + direct_inr_from_usd

    # =====================================================
    # INDIRECT COST
    # =====================================================

    indirect_df = df_cost[df_cost["Category"] == "Indirect"]

    indirect_pivot = (
        indirect_df
        .pivot_table(
            index="Particulars",
            columns="Month",
            values="Amount INR",
            aggfunc="sum",
            fill_value=0
        )
    )

    indirect_pivot = indirect_pivot.reindex(columns=month_cols, fill_value=0)
    indirect_pivot.insert(0, "Currency", "INR")
    indirect_pivot.reset_index(inplace=True)

    indirect_pivot.rename(columns={"Sub Cost": "Particulars"}, inplace=True)

    total_indirect = indirect_pivot[month_cols].sum()

    # =====================================================
    # BUILD TABLE ROWS
    # =====================================================

    rows = []

    # DIRECT GROUP
    rows.append({"Particulars": "Direct Cost", "Currency": ""})

    rows += usd_pivot.to_dict("records")

    rows.append({
        "Particulars": "Total USD",
        "Currency": "USD",
        **total_usd.to_dict()
    })

    rows.append({
        "Particulars": "FX Rate",
        "Currency": "",
        **fx_rate.to_dict()
    })

    rows.append({
        "Particulars": "Direct Cost INR",
        "Currency": "INR",
        **direct_inr_from_usd.to_dict()
    })

    rows += inr_pivot.to_dict("records")

    rows.append({
        "Particulars": "Total Direct Cost INR",
        "Currency": "INR",
        **total_direct_inr.to_dict()
    })

    # INDIRECT GROUP
    rows.append({"Particulars": "Indirect Cost", "Currency": ""})

    rows += indirect_pivot.to_dict("records")

    rows.append({
        "Particulars": "Total Indirect Cost INR",
        "Currency": "INR",
        **total_indirect.to_dict()
    })

    df_table = pd.DataFrame(rows)

    df_table["Annual/FY Total"] = df_table[month_cols].sum(axis=1)

    df_table = df_table[["Particulars", "Currency"] + month_cols + ["Annual/FY Total"]]

    # =====================================================
    # ADD GROUP COLUMN
    # =====================================================

    df_table["Group"] = ""

    df_table.loc[df_table["Particulars"].str.contains("Direct Cost"), "Group"] = "Direct Cost"
    df_table.loc[df_table["Particulars"].str.contains("Indirect Cost"), "Group"] = "Indirect Cost"

    df_table["Group"] = df_table["Group"].replace("", method="ffill")

    # =====================================================
    # AGGRID
    # =====================================================

    from st_aggrid import AgGrid, GridOptionsBuilder, JsCode

    gb = GridOptionsBuilder.from_dataframe(df_table)

    # Hide grouping column
    gb.configure_column("Group", rowGroup=True, hide=True)

    currency_formatter = JsCode("""
    function(params){

        if(params.value == null || params.value === '') return '';

        let currency = params.data.Currency;

        if(currency === "USD"){
            return '$' + Number(params.value).toLocaleString(undefined,{minimumFractionDigits:2});
        }

        if(currency === "INR"){
            return '₹' + Number(params.value).toLocaleString(undefined,{minimumFractionDigits:2});
        }

        return params.value;
    }
    """)

    for col in month_cols + ["Annual/FY Total"]:
        gb.configure_column(
            col,
            type=["numericColumn"],
            valueFormatter=currency_formatter
        )

    gb.configure_default_column(resizable=True)

    gridOptions = gb.build()

    gridOptions["groupDefaultExpanded"] = 0   # collapsed by default

    # =====================================
    # FINANCIAL FOOTER STYLE
    # =====================================

    gridOptions["getRowStyle"] = JsCode("""
    function(params){

        if(!params.data) return;

        if(
            params.data.Particulars === "Total USD" ||
            params.data.Particulars === "Direct Cost INR" ||
            params.data.Particulars === "Total Direct Cost INR" ||
            params.data.Particulars === "Total Indirect Cost INR"
        ){
            return {
                backgroundColor:'#003366',
                color:'white',
                fontWeight:'bold'
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
        df_table,
        gridOptions=gridOptions,
        allow_unsafe_jscode=True,
        height=600,
        fit_columns_on_grid_load=True,
        custom_css=custom_css,
        key=f"cost_centre_grid_{st.session_state.cost_table_refresh}"
    )