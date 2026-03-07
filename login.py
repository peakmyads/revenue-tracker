import streamlit as st
from datetime import datetime
import pandas as pd

# ===============================
# DEFAULT USERS
# ===============================

DEFAULT_USERS = {
    "Admin": {"password": "Admin@1808", "role": "Admin"},
    "Sales": {"password": "Sales@123", "role": "Sales"},
    "Finance": {"password": "Fin@123", "role": "Finance"},
    "User": {"password": "User@123", "role": "User"}
}

# ===============================
# ROLE TAB ACCESS
# ===============================

ROLE_ACCESS = {
    "Admin": [
        "Dashboard",
        "Summary",
        "Master Data",
        "DSP (Customers)",
        "SSP (Vendors)",
        "Partner Onboarding Form",
        "List of Partners",
        "Costs Centre",
        "Admin Control"
    ],

    "Sales": [
        "Dashboard",
        "Summary",
        "Partner Onboarding Form",
        "List of Partners"
    ],

    "Finance": [
        "Dashboard",
        "Summary",
        "DSP (Customers)",
        "SSP (Vendors)",
        "Partner Onboarding Form",
        "List of Partners",
        "Costs Centre"
    ],

    "User": [
        "Dashboard",
        "Summary",
        "List of Partners"
    ]
}

# ===============================
# LOGIN LOG FUNCTION
# ===============================

def log_login(spreadsheet, username):

    try:
        worksheet = spreadsheet.worksheet("Login Logs")
    except:
        worksheet = spreadsheet.add_worksheet("Login Logs", rows=1000, cols=10)
        worksheet.append_row(["Timestamp", "Username"])

    worksheet.append_row([
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        username
    ])

# ===============================
# LOGIN SCREEN
# ===============================

def login_screen(spreadsheet):

    st.markdown("""
    <style>

    .login-box {
        width:380px;
        margin:auto;
        margin-top:80px;
        padding:40px;
        border-radius:12px;
        background:white;
        box-shadow:0 10px 25px rgba(0,0,0,0.2);
    }

    .login-title{
        font-size:28px;
        font-weight:700;
        margin-bottom:20px;
        color:#003366;
        text-align:center;
    }

    /* BLUE LOGIN BUTTON */
    div.stButton > button{
        background:#0076CE;
        color:white;
        font-weight:700;
        border-radius:8px;
        padding:10px;
    }

    div.stButton > button:hover{
        background:#0059b3;
    }

    </style>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([2,1,2])

    with col2:
        
        st.divider()
        
        st.markdown('<div class="login-title">🔐 Login</div>', unsafe_allow_html=True)

        username = st.selectbox(
            "User",
            ["Admin","Sales","Finance","User"]
        )
        st.divider()
        password = st.text_input(
            "Password",
            type="password"
        )

        st.divider()
        
        login_btn = st.button("Login", use_container_width=True)

        if login_btn:

            if username in DEFAULT_USERS and password == DEFAULT_USERS[username]["password"]:

                st.session_state.logged_in = True
                st.session_state.user = username
                st.session_state.role = DEFAULT_USERS[username]["role"]

                log_login(spreadsheet, username)

                st.rerun()

            else:
                st.error("Invalid credentials")

        st.markdown('</div>', unsafe_allow_html=True)

# ===============================
# PASSWORD CHANGE (ADMIN ONLY)
# ===============================

def admin_change_password():

    st.subheader("🔑 Change User Password")

    if st.session_state.role != "Admin":
        st.warning("Only Admin can change passwords")
        return

    user = st.selectbox(
        "Select User",
        ["Admin", "Sales", "Finance", "User"]
    )

    new_pass = st.text_input("New Password", type="password")

    st.divider()
    
    if st.button("Update Password"):

        DEFAULT_USERS[user]["password"] = new_pass

        st.success("Password Updated")

# ===============================
# ROLE ACCESS FUNCTION
# ===============================

def get_allowed_tabs():

    role = st.session_state.role

    return ROLE_ACCESS.get(role, [])