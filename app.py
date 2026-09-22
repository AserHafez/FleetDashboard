import os
import glob
import re
import logging
import pandas as pd
import streamlit as st
from datetime import datetime, timedelta
import xlwings as xw
import math
import io
import win32com.client
import pythoncom
import tempfile
from supabase import create_client, Client
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter
import json

# ==========================================
# LOGGING CONFIGURATION
# ==========================================
# Server-side log file (separate from what users see in the UI). This makes it
# possible to diagnose failures after the fact instead of relying on console
# `print()` output, which is easy to lose on a shared/scheduled server.
LOG_DIR = os.path.join(os.getcwd(), "logs")
os.makedirs(LOG_DIR, exist_ok=True)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
    handlers=[
        logging.FileHandler(os.path.join(LOG_DIR, "fleet_planning.log"), encoding="utf-8"),
        logging.StreamHandler(),
    ],
)
logger = logging.getLogger("fleet_planning")

# ==========================================
# PAGE CONFIGURATION (MUST BE FIRST)
# ==========================================
st.set_page_config(
    page_title="Fleet Maintenance Planning",
    page_icon="✈️",
    layout="wide"
)

# Modern Corporate Aviation Design System — matches the reference look:
# gradient header banner, white metric/info cards, status badges, and a
# clean tab bar. Keeps the same brand colors used elsewhere in this file.
st.markdown("""
    <style>
    :root {
        --brand-navy: #0F172A;
        --brand-navy-dark: #0B3D6B;
        --brand-blue: #004B87;
        --brand-blue-hover: #003366;
        --brand-sky: #0284C7;
        --brand-bg: #F8FAFC;
        --brand-card: #FFFFFF;
        --brand-border: #E2E8F0;
        --brand-text: #0F172A;
        --brand-muted: #64748B;
        --brand-success: #15803D;
        --brand-warning: #B45309;
        --brand-danger: #B91C1C;
    }

    .stApp {
        background-color: var(--brand-bg);
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
    }

    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 3rem;
    }

    /* ---------- Sidebar ---------- */
    [data-testid="stSidebar"] {
        background-color: #F1F5F9;
        border-right: 1px solid var(--brand-border);
    }
    [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 {
        color: var(--brand-navy) !important;
        font-weight: 700 !important;
    }

    /* ---------- Header banner ---------- */
    .header-card {
        background: linear-gradient(135deg, #0F172A 0%, #004B87 100%);
        padding: 1.5rem 2rem;
        border-radius: 12px;
        color: white;
        margin-bottom: 1.5rem;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }
    .header-card h1 {
        color: #FFFFFF !important;
        margin: 0;
        font-size: 1.8rem;
        font-weight: 700;
        letter-spacing: -0.02em;
    }
    .header-card p {
        color: #94A3B8;
        margin: 0.25rem 0 0 0;
        font-size: 0.95rem;
    }

    /* ---------- Buttons ---------- */
    .stButton > button {
        border-radius: 8px;
        font-weight: 600;
        padding: 0.5rem 1.25rem;
        transition: all 0.2s ease;
        border: none;
    }
    .stButton > button[kind="primary"] {
        background-color: var(--brand-blue);
        color: #FFFFFF;
    }
    .stButton > button[kind="primary"]:hover {
        background-color: var(--brand-blue-hover);
        box-shadow: 0 2px 4px rgba(0,0,0,0.15);
    }
    .stButton > button[kind="secondary"] {
        background-color: #FFFFFF;
        color: var(--brand-blue);
        border: 1px solid var(--brand-border);
    }
    .stButton > button[kind="secondary"]:hover {
        background-color: #F1F5F9;
    }
    .stDownloadButton > button {
        border-radius: 8px;
        font-weight: 600;
    }

    /* ---------- Typography ---------- */
    h1, h2, h3, h4, label {
        color: var(--brand-text) !important;
    }
    h1 { font-weight: 800 !important; letter-spacing: -0.02em; }
    h2, h3 { font-weight: 700 !important; }

    /* ---------- Metric / info cards ---------- */
    .metric-card {
        background-color: var(--brand-card);
        border: 1px solid var(--brand-border);
        border-radius: 10px;
        padding: 1.25rem;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        transition: transform 0.15s ease;
    }
    .metric-card:hover {
        transform: translateY(-2px);
    }
    .metric-title {
        color: var(--brand-muted);
        font-size: 0.85rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.05em;
    }
    .metric-value {
        color: var(--brand-blue);
        font-size: 1.8rem;
        font-weight: 700;
        margin: 0.25rem 0;
    }
    .metric-sub {
        color: var(--brand-text);
        font-size: 0.85rem;
        font-weight: 500;
    }

    /* Native st.metric widgets keep the same card treatment */
    [data-testid="stMetric"] {
        background-color: var(--brand-card);
        border: 1px solid var(--brand-border);
        border-radius: 10px;
        padding: 0.9rem 1rem 0.6rem 1rem;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    [data-testid="stMetricValue"] {
        color: var(--brand-blue) !important;
        font-weight: 700 !important;
    }
    [data-testid="stMetricLabel"] {
        color: var(--brand-muted) !important;
        font-weight: 600 !important;
        text-transform: uppercase;
        font-size: 0.72rem !important;
        letter-spacing: 0.04em;
    }

    /* ---------- Status badges ---------- */
    .badge {
        display: inline-block;
        padding: 0.25em 0.6em;
        font-size: 75%;
        font-weight: 700;
        line-height: 1;
        text-align: center;
        white-space: nowrap;
        vertical-align: baseline;
        border-radius: 0.375rem;
    }
    .badge-critical { background-color: #FEE2E2; color: #991B1B; border: 1px solid #FCA5A5; }
    .badge-high { background-color: #FFEDD5; color: #9A3412; border: 1px solid #FDBA74; }
    .badge-medium { background-color: #FEF3C7; color: #92400E; border: 1px solid #FCD34D; }
    .badge-low { background-color: #DCFCE7; color: #166534; border: 1px solid #86EFAC; }

    /* ---------- Tabs (main navigation) ---------- */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        border-bottom: 2px solid var(--brand-border);
        background-color: transparent;
    }
    .stTabs [data-baseweb="tab"] {
        height: 48px;
        padding: 0 16px;
        background-color: transparent;
        border-radius: 8px 8px 0 0;
        font-weight: 600;
        color: var(--brand-muted);
    }
    .stTabs [aria-selected="true"] {
        background-color: var(--brand-card) !important;
        color: var(--brand-blue) !important;
        border-bottom: 3px solid var(--brand-blue) !important;
    }

    /* ---------- Expanders (card look) ---------- */
    [data-testid="stExpander"] {
        background-color: var(--brand-card);
        border: 1px solid var(--brand-border);
        border-radius: 10px;
        box-shadow: 0 1px 2px rgba(15, 23, 42, 0.04);
    }
    [data-testid="stExpander"] summary {
        font-weight: 600;
        color: var(--brand-blue);
    }

    /* ---------- Inputs ---------- */
    .stTextInput input, .stNumberInput input, .stTextArea textarea,
    .stSelectbox > div > div, .stMultiSelect > div > div {
        border-radius: 8px !important;
    }

    /* ---------- Dataframes / tables ---------- */
    [data-testid="stDataFrame"] {
        border: 1px solid var(--brand-border);
        border-radius: 8px;
        overflow: hidden;
    }

    /* ---------- Alerts ---------- */
    div[data-testid="stAlert"] {
        border-radius: 8px;
    }

    /* ---------- Divider ---------- */
    hr {
        border-color: var(--brand-border) !important;
    }

    /* ---------- Bordered containers (login card, etc.) ---------- */
    [data-testid="stVerticalBlockBorderWrapper"] {
        background-color: var(--brand-card);
        border: 1px solid var(--brand-border) !important;
        border-radius: 12px !important;
        box-shadow: 0 1px 3px rgba(15, 23, 42, 0.06);
    }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# STORAGE CONFIGURATION (NETWORK / LOCAL)
# ==========================================
NETWORK_BASE_DIR = r"U:\Planning\CEO PLANNING"
NETWORK_SAVED_DIR = os.path.join(NETWORK_BASE_DIR, "saved_reports")

# Use U: drive if reachable, otherwise fall back to local working directory
if os.path.exists(NETWORK_BASE_DIR):
    SAVED_DIR = NETWORK_SAVED_DIR
else:
    SAVED_DIR = os.path.join(os.getcwd(), "saved_reports")

os.makedirs(SAVED_DIR, exist_ok=True)

# DEFAULT SYSTEM PATHS
DEFAULT_FOLDER_PATH = r"U:\Planning\CEO PLANNING\LDND Calculations\AirCrafts\\"
DEFAULT_TCI_PATH = r"Z:\AIRWORTHINESS DOSSIER (ENG HESHAM)\TCI\A320-214\\"
DEFAULT_AMP_PATH = r"U:\Planning\CEO PLANNING\AMP Aircairo Nov25.xls"

# ==========================================
# FILE SYSTEM SAFETY HELPERS
# ==========================================
def sanitize_filename(raw_name: str, fallback: str = "report") -> str:
    """
    Strips path separators, drive letters, and '..' segments from a
    user-supplied file name so it can never be used to write or read
    outside SAVED_DIR (path traversal protection). Only the base file
    name is kept - any directory component the user typed is discarded.
    """
    if not raw_name:
        return fallback
    # Keep only the final path segment; ignore any folders the user typed.
    candidate = os.path.basename(str(raw_name).strip())
    # Drop characters that are unsafe/meaningless in a file name.
    candidate = re.sub(r'[\\/:*?"<>|\x00-\x1f]', "", candidate).strip(" .")
    if not candidate:
        return fallback
    return candidate


def write_json_atomic(path: str, data: dict) -> None:
    """
    Writes JSON to disk atomically (write to a temp file, then rename over
    the target). This avoids leaving a half-written/corrupted metadata file
    if the process is interrupted or two users save at nearly the same time.
    """
    directory = os.path.dirname(path) or "."
    fd, tmp_path = tempfile.mkstemp(dir=directory, prefix=".tmp_", suffix=".json")
    try:
        with os.fdopen(fd, "w") as f:
            json.dump(data, f, indent=2)
        os.replace(tmp_path, path)
    except Exception:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        raise


@st.cache_data(show_spinner=False)
def read_excel_cached(file_path: str, mtime: float) -> pd.DataFrame:
    """
    Cached wrapper around pd.read_excel. `mtime` (the file's last-modified
    time) is part of the cache key, so a stale cache entry is automatically
    invalidated the moment the underlying file changes on disk.
    """
    return pd.read_excel(file_path)


def read_excel_fresh(file_path: str) -> pd.DataFrame:
    """Reads an Excel file, transparently caching on (path, mtime)."""
    mtime = os.path.getmtime(file_path)
    return read_excel_cached(file_path, mtime)


# ==========================================
# SUPABASE CONFIGURATION
# ==========================================
SUPABASE_URL = st.secrets.get("SUPABASE_URL", "")
SUPABASE_KEY = st.secrets.get("SUPABASE_KEY", "")
SUPABASE_CONFIGURED = bool(SUPABASE_URL) and bool(SUPABASE_KEY)

if not SUPABASE_CONFIGURED:
    # Fail loudly rather than silently degrading everyone to "viewer" with
    # no explanation, which is what happened before when Supabase calls
    # quietly raised and were swallowed by a bare except/pass.
    st.error(
        "⚠️ Supabase is not configured. Add SUPABASE_URL and SUPABASE_KEY "
        "to `.streamlit/secrets.toml` before running this app."
    )
    st.stop()

@st.cache_resource
def init_supabase() -> Client:
    return create_client(SUPABASE_URL, SUPABASE_KEY)

supabase = init_supabase()

# ==========================================
# SESSION STATE INITIALIZATION
# ==========================================
if "popover_key_rename" not in st.session_state:
    st.session_state.popover_key_rename = 0
if "popover_key_delete" not in st.session_state:
    st.session_state.popover_key_delete = 0
if "popover_key_sel" not in st.session_state:
    st.session_state.popover_key_sel = 0
if "popover_key_bow" not in st.session_state:
    st.session_state.popover_key_bow = 0
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if "user_email" not in st.session_state:
    st.session_state.user_email = None
if "user_id" not in st.session_state:
    st.session_state.user_id = None
if "role" not in st.session_state:
    st.session_state.role = "viewer"  # Default to viewer for new/unassigned users
if "auth_mode" not in st.session_state:
    st.session_state.auth_mode = "login"
if "all_results" not in st.session_state:
    st.session_state.all_results = None
if "aircraft_stats" not in st.session_state:
    st.session_state.aircraft_stats = None
if "preview_df" not in st.session_state:
    st.session_state.preview_df = None
if "export_mode" not in st.session_state:
    st.session_state.export_mode = None
if "select_all_state" not in st.session_state:
    st.session_state.select_all_state = False
if "search_filter" not in st.session_state:
    st.session_state.search_filter = ""
if "priority_filter" not in st.session_state:
    st.session_state.priority_filter = ["CRITICAL", "HIGH", "MEDIUM", "LOW"]

# Falls back to the original hardcoded address if not set in secrets, so
# existing deployments keep working; but new deployments should set this in
# .streamlit/secrets.toml instead of hardcoding an address in source code.
ADMIN_EMAIL = st.secrets.get("ADMIN_EMAIL", "aser.hafez@aircairo.com")

# Optional: restrict self-signup to a company email domain, e.g. "aircairo.com".
# Leave unset in secrets to keep the original behavior (any email can sign up).
ALLOWED_SIGNUP_DOMAIN = st.secrets.get("ALLOWED_SIGNUP_DOMAIN", "")

# ==========================================
# AUTHENTICATION & USER MANAGEMENT HELPERS
# ==========================================
def get_user_role(user_id, email):
    if email and email.lower() == ADMIN_EMAIL.lower():
        return "admin"
    try:
        res = supabase.table("profiles").select("role").eq("id", user_id).execute()
        if res.data and len(res.data) > 0:
            return res.data[0].get("role", "viewer")
        else:
            default_role = "viewer"
            supabase.table("profiles").upsert({
                "id": user_id,
                "email": email,
                "role": default_role
            }).execute()
            return default_role
    except Exception as e:
        logger.error("Could not fetch/assign role for user %s (%s): %s", email, user_id, e)
    return "viewer"

def fetch_all_user_profiles():
    try:
        res = supabase.table("profiles").select("*").execute()
        return res.data if res.data else []
    except Exception as e:
        st.error(f"Error fetching user list: {e}")
        return []

def update_user_role_in_db(user_id, new_role):
    try:
        res = supabase.table("profiles").update({"role": new_role}).eq("id", user_id).execute()
        if not res.data:
            return False, "Database update failed. Verify Supabase RLS policies for the 'profiles' table."
        if user_id == st.session_state.get("user_id"):
            st.session_state.role = new_role
        return True, f"Role successfully updated to '{new_role.upper()}'"
    except Exception as e:
        return False, f"Failed to update user role: {str(e)}"

def handle_login(email, password):
    if not email or not email.strip() or not password:
        st.error("Please enter both an email address and a password.")
        return
    try:
        res = supabase.auth.sign_in_with_password({"email": email.strip(), "password": password})
        st.session_state.authenticated = True
        st.session_state.user_email = res.user.email
        st.session_state.user_id = res.user.id
        st.session_state.role = get_user_role(res.user.id, res.user.email)
        st.success("Successfully logged in!")
        st.rerun()
    except Exception as e:
        logger.warning("Login failed for %s: %s", email, e)
        st.error(f"Login failed: {str(e)}")

def handle_signup(email, password, confirm_password):
    email = (email or "").strip()
    if not email or "@" not in email:
        st.error("Please enter a valid email address.")
        return
    if ALLOWED_SIGNUP_DOMAIN and not email.lower().endswith(f"@{ALLOWED_SIGNUP_DOMAIN.lower()}"):
        st.error(f"Sign-up is restricted to @{ALLOWED_SIGNUP_DOMAIN} email addresses.")
        return
    if password != confirm_password:
        st.error("Passwords do not match!")
        return
    if len(password) < 6:
        st.error("Password must be at least 6 characters long.")
        return

    try:
        res = supabase.auth.sign_up({"email": email, "password": password})
        st.success("Account created successfully! Check your email to verify your account before logging in.")
        st.session_state.auth_mode = "login"
    except Exception as e:
        logger.warning("Signup failed for %s: %s", email, e)
        st.error(f"Signup failed: {str(e)}")

def handle_password_reset(email):
    try:
        supabase.auth.reset_password_for_email(email)
        st.success(f"Password reset link sent to {email}. Check your inbox.")
        st.session_state.auth_mode = "login"
    except Exception as e:
        st.error(f"Reset request failed: {str(e)}")

# ==========================================
# AUTHENTICATION INTERFACE
# ==========================================
if not st.session_state.authenticated:
    st.markdown("""
        <div style="text-align: center; padding: 2rem 0;">
            <h1 style="color: #0F172A; font-weight: 800; font-size: 2.2rem;">✈️ Fleet Maintenance Planning Portal</h1>
            <p style="color: #64748B;">Aircraft maintenance & production planning</p>
        </div>
    """, unsafe_allow_html=True)

    col_a1, col_a2, col_a3 = st.columns([1, 1.2, 1])
    with col_a2:
        with st.container(border=True):
            if st.session_state.auth_mode == "login":
                st.subheader("🔑 Sign In")
                with st.form("login_form"):
                    email_in = st.text_input("Email Address", placeholder="you@company.com")
                    pass_in = st.text_input("Password", type="password", placeholder="••••••••")
                    submit_login = st.form_submit_button("Sign In", type="primary", use_container_width=True)
                    if submit_login:
                        handle_login(email_in, pass_in)

                col1, col2 = st.columns(2)
                with col1:
                    if st.button("Create Account", use_container_width=True):
                        st.session_state.auth_mode = "signup"
                        st.rerun()
                with col2:
                    if st.button("Forgot Password?", use_container_width=True):
                        st.session_state.auth_mode = "reset"
                        st.rerun()

            elif st.session_state.auth_mode == "signup":
                st.subheader("📝 Account Registration")
                st.caption("First-time setup for admins and new users.")
                with st.form("signup_form"):
                    email_in = st.text_input("Corporate Email Address", placeholder="you@company.com")
                    pass_in = st.text_input("Password", type="password", placeholder="At least 6 characters")
                    pass_confirm = st.text_input("Confirm Password", type="password")
                    submit_signup = st.form_submit_button("Register & Verify", type="primary", use_container_width=True)

                    if submit_signup:
                        handle_signup(email_in, pass_in, pass_confirm)

                if st.button("Back to Login", use_container_width=True):
                    st.session_state.auth_mode = "login"
                    st.rerun()

            elif st.session_state.auth_mode == "reset":
                st.subheader("🔒 Reset Password")
                with st.form("reset_form"):
                    email_in = st.text_input("Registered Email Address", placeholder="you@company.com")
                    submit_reset = st.form_submit_button("Send Reset Link", type="primary", use_container_width=True)

                    if submit_reset:
                        handle_password_reset(email_in)

                if st.button("Back to Login", use_container_width=True):
                    st.session_state.auth_mode = "login"
                    st.rerun()

    st.stop()

# ==========================================
# HEADER BAR (AUTHENTICATED)
# ==========================================
role_colors = {"admin": "#FFEDD5", "planner": "#DCFCE7", "production": "#DCFCE7", "viewer": "#E2E8F0"}
role_badge_bg = role_colors.get(st.session_state.role, "#E2E8F0")
st.markdown(f"""
    <div class="header-card">
        <div style="display: flex; justify-content: space-between; align-items: center;">
            <div>
                <h1>✈️ Fleet Maintenance Planning</h1>
                <p>Developed by Eng. Aser Hafez</p>
            </div>
            <div style="text-align: right;">
                <span class="badge badge-low" style="font-size: 0.85rem; padding: 0.4rem 0.8rem;">
                    👤 {st.session_state.user_email} ({st.session_state.role.upper()})
                </span>
            </div>
        </div>
    </div>
""", unsafe_allow_html=True)

# Sidebar user controls
with st.sidebar:
    if os.path.exists("logo.png"):
        st.image("logo.png", use_container_width=True)
    st.markdown("### ⚙️ System Controls")
    if st.button("🚪 Logout", type="secondary", use_container_width=True):
        supabase.auth.sign_out()
        st.session_state.authenticated = False
        st.session_state.user_email = None
        st.session_state.user_id = None
        st.session_state.role = "viewer"
        st.rerun()
    st.divider()

# --- SIDEBAR CONFIGURATION BY ROLE ---
st.sidebar.markdown(
    "<h3 style='margin-bottom:0;'>✈️ Fleet Ops</h3>"
    "<p style='color:#64748B; margin-top:0; font-size:0.85rem;'>Navigation & Criteria</p>",
    unsafe_allow_html=True
)
st.sidebar.divider()

# 1. ADMIN ROLE CONFIGURATION
if st.session_state.role == "admin":
    st.sidebar.subheader("🔒 Admin Folder Settings")
    folder_path = st.sidebar.text_input("Fleet Folder Path", value=DEFAULT_FOLDER_PATH)
    tci_path_input = st.sidebar.text_input("TCI Base Folder Path (Optional)", value=DEFAULT_TCI_PATH)
    AMP_FILE_PATH = st.sidebar.text_input("AMP File Path", value=DEFAULT_AMP_PATH)
    if st.sidebar.button("🔄 Refresh Aircraft Folder List"):
        scan_available_registrations.clear()
        st.rerun()

# 2. PLANNER ROLE CONFIGURATION
elif st.session_state.role == "planner":
    folder_path = DEFAULT_FOLDER_PATH
    tci_path_input = DEFAULT_TCI_PATH
    AMP_FILE_PATH = DEFAULT_AMP_PATH

# 3. PRODUCTION ROLE CONFIGURATION
elif st.session_state.role == "production":
    folder_path = DEFAULT_FOLDER_PATH
    tci_path_input = DEFAULT_TCI_PATH
    AMP_FILE_PATH = DEFAULT_AMP_PATH
    st.sidebar.info("📌 **Production Mode**\nOpen the **Production Workspace** tab to view and work on reports sent to you by the Planning team.")

# 4. UNASSIGNED / VIEWER ROLE CONFIGURATION
else:
    folder_path = DEFAULT_FOLDER_PATH
    tci_path_input = DEFAULT_TCI_PATH
    AMP_FILE_PATH = DEFAULT_AMP_PATH
    st.sidebar.warning("⏳ **Pending Role Assignment**\nYour account is assigned View-Only access until an admin assigns you a role.")

@st.cache_data(ttl=300, show_spinner=False)
def scan_available_registrations(base_folder: str) -> list:
    """
    Lists aircraft registration sub-folders under base_folder. Cached for 5
    minutes so a full os.walk of the (often network-mounted) fleet folder
    doesn't re-run on every single widget interaction/rerun in the app -
    previously this ran unconditionally on every Streamlit script rerun.
    """
    regs = []
    if os.path.exists(base_folder):
        for root, dirs, files in os.walk(base_folder):
            for d in dirs:
                if not d.startswith("~$") and "MASTER" not in d.upper():
                    regs.append(d.upper())
    return regs

# Show search options only for Admin & Planner
if st.session_state.role in ["admin", "planner"]:
    st.sidebar.subheader("🔍 Search Criteria")
    available_regs = ["ALL"] + scan_available_registrations(folder_path)

    target_reg = st.sidebar.selectbox("Aircraft Reg (or ALL)", options=available_regs, index=0)
    fh_limit = st.sidebar.number_input("FH Limit", value=600, step=50)
    fc_limit = st.sidebar.number_input("FC Limit", value=300, step=50)
    days_limit = st.sidebar.number_input("Days Limit", value=45, step=5)

    st.sidebar.markdown("---")
    st.sidebar.subheader("📊 Report Filter Criteria")

    section_options = [
        "ALL",
        "Checks Follow-Up",
        "Out Of Phase",
        "Base Maint OOP",
        "AD",
        "TCI"
    ]

    selected_sections = st.sidebar.multiselect(
        "Select Report Type(s):",
        options=section_options,
        default=["ALL"],
        help="Select one, multiple, or ALL reports to filter maintenance tasks."
    )

    st.sidebar.markdown("---")
    run_button = st.sidebar.button("▶ Run Fleet Review", type="primary", use_container_width=True)
else:
    run_button = False

# --- HELPER FUNCTIONS ---

def log_download_activity(filename, user_email):
    """Logs download event with file name, user, and timestamp."""
    DOWNLOAD_LOG_FILE = os.path.join(SAVED_DIR, "download_history.json")
    logs = []
    if os.path.exists(DOWNLOAD_LOG_FILE):
        try:
            with open(DOWNLOAD_LOG_FILE, "r") as f:
                logs = json.load(f)
        except Exception as e:
            logger.warning("Could not read download log, starting fresh: %s", e)
            logs = []

    log_entry = {
        "filename": filename,
        "downloaded_by": user_email if user_email else "Anonymous User",
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }

    logs.insert(0, log_entry)

    try:
        write_json_atomic(DOWNLOAD_LOG_FILE, logs)
    except Exception as e:
        logger.error("Error saving download log: %s", e)

    try:
        supabase.table("download_logs").insert(log_entry).execute()
    except Exception as e:
        logger.warning("Could not mirror download log to Supabase: %s", e)

def get_download_logs():
    """Retrieves all download activity logs."""
    DOWNLOAD_LOG_FILE = os.path.join(SAVED_DIR, "download_history.json")
    if os.path.exists(DOWNLOAD_LOG_FILE):
        try:
            with open(DOWNLOAD_LOG_FILE, "r") as f:
                data = json.load(f)
                return pd.DataFrame(data)
        except Exception as e:
            logger.warning("Could not read download log file: %s", e)
    return pd.DataFrame(columns=["filename", "downloaded_by", "timestamp"])

def map_skill_code_to_trade(skill_code):
    if not skill_code:
        return "B1 (Mechanical)"
        
    code = str(skill_code).strip().upper()
    
    if code in ["AV", "EL", "RA"]:
        return "B2 (Avionics)"
    elif code in ["AF", "CA", "EN", "UT"]:
        return "B1 (Mechanical)"
    elif "NDT" in code:
        return "NDT / Borescope"
    else:
        return "B1 (Mechanical)"

def clean_and_sum_cell_mh(val):
    if pd.isna(val) or not val:
        return 0.0
    
    lines = re.split(r'[\r\n]+', str(val))
    total_mh = 0.0
    
    for line in lines:
        cleaned_line = line.strip()
        if cleaned_line:
            match = re.search(r'[-+]?\d*\.?\d+', cleaned_line)
            if match:
                total_mh += float(match.group())
                
    return total_mh

def clean_man_hours(val):
    if pd.isna(val) or val is None:
        return ""
    s = str(val).strip().replace("\n", " ")
    matches = re.findall(r"\d+(?:\.\d+)?", s)
    if matches:
        return matches[0]
    return ""

def parse_mh_to_float(val):
    if pd.isna(val) or val is None:
        return 0.0
    try:
        clean_v = clean_man_hours(val)
        return float(clean_v) if clean_v else 0.0
    except (ValueError, TypeError):
        return 0.0

def _safe_quit_xw_app(app):
    """
    Quits a xlwings/Excel COM App instance without letting a COM error on
    shutdown crash the request or leave the failure completely silent. A raw
    `app.quit()` with no guard can raise if Excel is already unresponsive,
    which previously could abort cleanup and leave an orphaned EXCEL.EXE
    process running on the server.
    """
    try:
        app.quit()
    except Exception as e:
        logger.warning("Error while closing Excel COM instance: %s", e)


def open_wb_with_xlwings(app, file_path):
    try:
        wb = app.books.open(file_path, update_links=True, read_only=True)
        app.calculate()
        return wb
    except Exception as e:
        logger.error("Error opening %s via xlwings: %s", file_path, e)
        return None

EMAIL_PATTERN = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

def _is_valid_email_list(value: str) -> bool:
    """Accepts a single address or a ';'/','-separated list, as Outlook does."""
    if not value or not value.strip():
        return False
    parts = re.split(r"[;,]", value)
    return all(EMAIL_PATTERN.match(p.strip()) for p in parts if p.strip())


def send_bow_via_outlook(to_email, cc_email, subject, body_text, excel_bytes, filename, extra_attachments=None):
    if not _is_valid_email_list(to_email):
        return False, "Please enter a valid recipient email address."
    if cc_email and cc_email.strip() and not _is_valid_email_list(cc_email):
        return False, "Please enter a valid CC email address."

    pythoncom.CoInitialize()
    temp_dir = tempfile.mkdtemp()
    temp_file_path = os.path.join(temp_dir, filename)
    temp_files_to_clean = [temp_file_path]

    with open(temp_file_path, "wb") as f:
        f.write(excel_bytes.getvalue())

    additional_file_paths = []
    if extra_attachments:
        for uploaded_file in extra_attachments:
            uploaded_path = os.path.join(temp_dir, uploaded_file.name)
            with open(uploaded_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            additional_file_paths.append(uploaded_path)
            temp_files_to_clean.append(uploaded_path)

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.Display()
        mail.To = to_email
        if cc_email.strip():
            mail.CC = cc_email.strip()
            
        mail.Subject = subject
        html_body_text = body_text.replace("\n", "<br>")
        mail.HTMLBody = f"<div>{html_body_text}</div><br><br>" + mail.HTMLBody
        mail.Attachments.Add(temp_file_path)
        
        for extra_path in additional_file_paths:
            mail.Attachments.Add(extra_path)
        
        mail.Send()
        return True, "Email sent successfully via Outlook with attachments!"
    except Exception as e:
        logger.error("Failed to send BOW email via Outlook: %s", e)
        return False, f"Failed to send email via Outlook: {str(e)}"
    finally:
        pythoncom.CoUninitialize()
        for t_file in temp_files_to_clean:
            if os.path.exists(t_file):
                try:
                    os.remove(t_file)
                except OSError as e:
                    logger.warning("Could not remove temp file %s: %s", t_file, e)
        if os.path.exists(temp_dir):
            try:
                os.rmdir(temp_dir)
            except OSError as e:
                logger.warning("Could not remove temp dir %s: %s", temp_dir, e)

def get_acheck_clean_name(task_name):
    if not task_name:
        return ""
    clean = str(task_name).upper().strip()
    clean = clean.replace("\xa0", "").replace("CHECK", "").replace("PACKAGE", "")
    clean = clean.replace("-", "").replace("_", "").replace(" ", "").strip()
    
    mapping = {
        "1A": "1A", "A1": "1A", "A01": "1A", "01A": "1A", "A001": "1A", "A": "1A", "A(1)": "1A",
        "2A": "2A", "A2": "2A", "A02": "2A", "02A": "2A", "A002": "2A",
        "3A": "3A", "A3": "3A", "A03": "3A", "03A": "3A", "A003": "3A",
        "4A": "4A", "A4": "4A", "A04": "4A", "04A": "4A", "A004": "4A",
        "5A": "5A", "A5": "5A", "A05": "5A", "05A": "5A", "A005": "5A",
        "6A": "6A", "A6": "6A", "A06": "6A", "06A": "6A", "A006": "6A",
        "7A": "7A", "A7": "7A", "A07": "7A", "07A": "7A", "A007": "7A",
        "8A": "8A", "A8": "8A", "A08": "8A", "08A": "8A", "A008": "8A",
    }
    return mapping.get(clean, str(task_name).strip())

def get_acheck_subtasks(task_name):
    clean = get_acheck_clean_name(task_name)
    subtask_map = {
        "1A": ["1A"],
        "2A": ["1A", "2A"],
        "3A": ["1A", "3A"],
        "4A": ["1A", "2A", "4A"],
        "5A": ["1A", "5A"],
        "6A": ["1A", "2A", "3A", "6A"],
        "7A": ["1A", "7A"],
        "8A": ["1A", "2A", "4A", "8A"]
    }
    return subtask_map.get(clean, None)

def get_triggers(rem_fh, rem_fc, rem_days, fh_lim, fc_lim, days_lim):
    trigs = []
    if rem_fh is not None and isinstance(rem_fh, (int, float)) and rem_fh <= fh_lim:
        trigs.append("FH")
    if rem_fc is not None and isinstance(rem_fc, (int, float)) and rem_fc <= fc_lim:
        trigs.append("FC")
    if rem_days is not None and isinstance(rem_days, (int, float)) and rem_days <= days_lim:
        trigs.append("Days")
    return ", ".join(trigs)

def get_priority(rem_fh, rem_fc, rem_days):
    fh = rem_fh if isinstance(rem_fh, (int, float)) else None
    fc = rem_fc if isinstance(rem_fc, (int, float)) else None
    days = rem_days if isinstance(rem_days, (int, float)) else None

    if (fh is not None and fh <= 0) or (fc is not None and fc <= 0) or (days is not None and days <= 0):
        return "CRITICAL"
    if (fh is not None and fh < 100) or (fc is not None and fc < 50) or (days is not None and days < 10):
        return "HIGH"
    if ((fh is not None and 100 <= fh < 500) or 
        (fc is not None and 50 <= fc < 210) or 
        (days is not None and 10 <= days < 45)):
        return "MEDIUM"
    return "LOW"

def clean_num(val):
    if pd.isna(val) or val is None:
        return None
    try:
        v = float(str(val).replace(",", ".").replace(" ", ""))
        return int(round(v))
    except (ValueError, TypeError):
        return None

def format_int_cell(val):
    if pd.isna(val) or val is None:
        return ""
    try:
        return f"{int(round(float(val)))}"
    except (ValueError, TypeError):
        return str(val)

def get_subtask_values(df_checks, sub_task_name):
    target_clean = get_acheck_clean_name(sub_task_name)
    for idx, row in df_checks.iterrows():
        c_task = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
        if c_task and get_acheck_clean_name(c_task) == target_clean:
            if "YE" in c_task.upper() or c_task.upper() == "WEEKLY":
                return None, None, clean_num(row.iloc[12])
            else:
                return clean_num(row.iloc[12]), clean_num(row.iloc[13]), clean_num(row.iloc[14])
    return None, None, None

def get_sheet_dataframe(wb, sheet_name):
    try:
        sheet = wb.sheets[sheet_name]
        data = sheet.used_range.value
        if data:
            return pd.DataFrame(data)
    except Exception as e:
        # Missing sheets are expected (not every workbook has every tab),
        # so this stays at debug level rather than error/print noise.
        logger.debug("Sheet '%s' not read from %s: %s", sheet_name, getattr(wb, "fullname", "?"), e)
    return pd.DataFrame()

def calculate_est_remaining_days(row):
    fh = row.get("Remaining FH")
    fc = row.get("Remaining FC")
    days = row.get("Remaining Days")

    estimates = []

    def safe_float(val):
        if pd.isna(val) or val is None or str(val).strip() == "":
            return None
        try:
            v = float(val)
            return None if math.isnan(v) else v
        except (ValueError, TypeError):
            return None

    num_fh = safe_float(fh)
    num_fc = safe_float(fc)
    num_days = safe_float(days)

    if num_fh is not None:
        estimates.append(round(num_fh / 11))
    if num_fc is not None:
        estimates.append(round(num_fc / 7))
    if num_days is not None:
        estimates.append(round(num_days))

    return min(estimates) if estimates else ""

def process_single_aircraft(excel_path, xw_app, fh_lim, fc_lim, days_lim, tci_base_folder=""):
    results = []
    wb = open_wb_with_xlwings(xw_app, excel_path)
    if not wb:
        return "", 0, 0, pd.DataFrame()

    try:
        df_checks_full = get_sheet_dataframe(wb, "Checks Follow-Up")
        if not df_checks_full.empty and len(df_checks_full) >= 4:
            ac_reg = str(df_checks_full.iloc[1, 3]).strip() if pd.notna(df_checks_full.iloc[1, 3]) else os.path.basename(excel_path)
            cur_afh = clean_num(df_checks_full.iloc[3, 6]) or 0
            cur_afc = clean_num(df_checks_full.iloc[3, 8]) or 0
        else:
            ac_reg = os.path.basename(excel_path)
            cur_afh = 0
            cur_afc = 0

        written_dict = set()

        if not df_checks_full.empty and len(df_checks_full) > 8:
            df_checks = df_checks_full.iloc[8:].reset_index(drop=True)
            for idx, row in df_checks.iterrows():
                task_name = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
                
                if not task_name or any(x in task_name.upper() for x in ["OUT OF PHASE", "BASE MAINTENANCE", "ENGINE WASH"]):
                    continue

                if "YE" in task_name.upper() or task_name.upper() == "WEEKLY":
                    rem_fh, rem_fc = None, None
                    rem_days = clean_num(row.iloc[12])
                else:
                    rem_fh = clean_num(row.iloc[12])
                    rem_fc = clean_num(row.iloc[13])
                    rem_days = clean_num(row.iloc[14])

                if rem_fh is None and rem_fc is None and rem_days is None:
                    continue

                trigger = get_triggers(rem_fh, rem_fc, rem_days, fh_lim, fc_lim, days_lim)
                if trigger:
                    sub_tasks = get_acheck_subtasks(task_name)
                    clean_name = get_acheck_clean_name(task_name)

                    if sub_tasks:
                        for sub in sub_tasks:
                            key = f"{ac_reg}|Checks Follow-Up|{sub}"
                            if key not in written_dict:
                                written_dict.add(key)
                                if sub == clean_name:
                                    sfh, sfc, sdays, strig = rem_fh, rem_fc, rem_days, trigger
                                else:
                                    sfh, sfc, sdays = get_subtask_values(df_checks, sub)
                                    strig = clean_name
                                
                                sprio = get_priority(sfh, sfc, sdays)
                                results.append({
                                    "Reg": ac_reg, "Section": "Checks Follow-Up", "Task": sub, "PN": "", "SN": "",
                                    "Description": "", "Interval FH": "", "Interval FC": "", "Interval Days": "",
                                    "Remaining FH": sfh, "Remaining FC": sfc, "Remaining Days": sdays,
                                    "Trigger": strig, "Priority": sprio,
                                    "File Path": excel_path
                                })
                    else:
                        key = f"{ac_reg}|Checks Follow-Up|{task_name}"
                        if key not in written_dict:
                            written_dict.add(key)
                            prio = get_priority(rem_fh, rem_fc, rem_days)
                            results.append({
                                "Reg": ac_reg, "Section": "Checks Follow-Up", "Task": task_name, "PN": "", "SN": "",
                                "Description": "", "Interval FH": "", "Interval FC": "", "Interval Days": "",
                                "Remaining FH": rem_fh, "Remaining FC": rem_fc, "Remaining Days": rem_days,
                                "Trigger": trigger, "Priority": prio,
                                "File Path": excel_path
                            })

        def parse_oop_sheet(sheet_name, section_title):
            df_oop_full = get_sheet_dataframe(wb, sheet_name)
            if not df_oop_full.empty and len(df_oop_full) > 6:
                df_oop = df_oop_full.iloc[6:].reset_index(drop=True)
                for idx, row in df_oop.iterrows():
                    task_name = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else ""
                    if task_name:
                        desc = str(row.iloc[5]).replace("\n", " ")[:80] if pd.notna(row.iloc[5]) else ""
                        int_fh = format_int_cell(row.iloc[14])
                        int_fc = format_int_cell(row.iloc[15])
                        
                        raw_int_days = row.iloc[16]
                        if pd.notna(raw_int_days):
                            if isinstance(raw_int_days, (datetime, pd.Timestamp)):
                                int_days = raw_int_days.strftime("%Y-%m-%d")
                            else:
                                int_days = format_int_cell(raw_int_days)
                        else:
                            int_days = ""

                        rem_fh = clean_num(row.iloc[17])
                        rem_fc = clean_num(row.iloc[18])
                        rem_days = clean_num(row.iloc[19])

                        trigger = get_triggers(rem_fh, rem_fc, rem_days, fh_lim, fc_lim, days_lim)
                        if trigger:
                            results.append({
                                "Reg": ac_reg, "Section": section_title, "Task": task_name, "PN": "", "SN": "",
                                "Description": desc, "Interval FH": int_fh, "Interval FC": int_fc, "Interval Days": int_days,
                                "Remaining FH": rem_fh, "Remaining FC": rem_fc, "Remaining Days": rem_days,
                                "Trigger": trigger, "Priority": get_priority(rem_fh, rem_fc, rem_days),
                                "File Path": excel_path
                            })

        parse_oop_sheet("Out Of Phase", "Out Of Phase")
        parse_oop_sheet("Base Maintenance Out Of Phase", "Base Maint OOP")

        df_ad_full = get_sheet_dataframe(wb, "AD")
        if not df_ad_full.empty and len(df_ad_full) > 6:
            df_ad = df_ad_full.iloc[6:].reset_index(drop=True)
            for idx, row in df_ad.iterrows():
                task_name = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
                if task_name:
                    desc = str(row.iloc[3]).replace("\n", " ")[:80] if pd.notna(row.iloc[3]) else ""
                    ad_unit = str(row.iloc[12]).strip().upper() if pd.notna(row.iloc[12]) else ""
                    rem_val = clean_num(row.iloc[11])

                    rem_fh = rem_val if ad_unit == "H" else None
                    rem_fc = rem_val if ad_unit == "C" else None
                    rem_days = rem_val if ad_unit == "D" else None

                    trigger = get_triggers(rem_fh, rem_fc, rem_days, fh_lim, fc_lim, days_lim)
                    if trigger:
                        results.append({
                            "Reg": ac_reg, "Section": "AD", "Task": task_name, "PN": "", "SN": "",
                            "Description": desc, "Interval FH": "", "Interval FC": "", "Interval Days": "",
                            "Remaining FH": rem_fh, "Remaining FC": rem_fc, "Remaining Days": rem_days,
                            "Trigger": trigger, "Priority": get_priority(rem_fh, rem_fc, rem_days),
                            "File Path": excel_path
                        })

        if tci_base_folder and os.path.exists(tci_base_folder):
            try:
                for sub in os.listdir(tci_base_folder):
                    if ac_reg in sub and os.path.isdir(os.path.join(tci_base_folder, sub)):
                        tci_dir = os.path.join(tci_base_folder, sub)
                        tci_files = glob.glob(os.path.join(tci_dir, "*TCI UPDATING*.xls*"))
                        if tci_files:
                            wb_tci = open_wb_with_xlwings(xw_app, tci_files[0])
                            if wb_tci:
                                df_tci = get_sheet_dataframe(wb_tci, wb_tci.sheets[0].name)
                                wb_tci.close()

                                if not df_tci.empty:
                                    curr_task, curr_desc, curr_pn, curr_sn = "", "", "", ""
                                    for idx in range(1, len(df_tci)):
                                        row = df_tci.iloc[idx]
                                        val_a    = str(row.iloc[0]).strip() if len(row) > 0 and pd.notna(row.iloc[0]) else ""
                                        val_desc = str(row.iloc[2]).strip() if len(row) > 2 and pd.notna(row.iloc[2]) else ""
                                        val_pn   = clean_code_str(row.iloc[3]) if len(row) > 3 and pd.notna(row.iloc[3]) else ""
                                        val_sn   = clean_code_str(row.iloc[4]) if len(row) > 4 and pd.notna(row.iloc[4]) else ""

                                        if val_a:    curr_task = val_a
                                        if val_desc: curr_desc = val_desc
                                        if val_pn:   curr_pn   = val_pn
                                        if val_sn:   curr_sn   = val_sn

                                        if not curr_task:
                                            continue

                                        raw_days = row.iloc[16] if len(row) > 16 and pd.notna(row.iloc[16]) else ""
                                        raw_fh   = row.iloc[17] if len(row) > 17 and pd.notna(row.iloc[17]) else ""
                                        raw_fc   = row.iloc[18] if len(row) > 18 and pd.notna(row.iloc[18]) else ""

                                        if str(raw_days).startswith("#"): raw_days = ""
                                        if str(raw_fh).startswith("#"): raw_fh = ""
                                        if str(raw_fc).startswith("#"): raw_fc = ""

                                        rem_fh = None
                                        try:
                                            clean_fh = float(str(raw_fh).replace(" ", "").replace(",", ".")) if str(raw_fh).strip() else 0.0
                                        except ValueError:
                                            clean_fh = 0.0

                                        if clean_fh > 0:
                                            rem_fh = round(clean_fh, 2)

                                        rem_fc = None
                                        try:
                                            clean_fc = float(str(raw_fc).replace(" ", "").replace(",", ".")) if str(raw_fc).strip() else 0.0
                                        except ValueError:
                                            clean_fc = 0.0

                                        if clean_fc > 0:
                                            rem_fc = round(clean_fc, 2)

                                        rem_days = None
                                        if isinstance(raw_days, (datetime, pd.Timestamp)):
                                            today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
                                            rem_days = (raw_days.replace(tzinfo=None) - today).days
                                        elif isinstance(raw_days, str) and any(c in raw_days for c in ["-", "/", "."]) and len(raw_days.strip()) >= 8:
                                            d_val = pd.to_datetime(raw_days, errors='coerce')
                                            if pd.notna(d_val):
                                                today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
                                                rem_days = (d_val - today).days
                                            else:
                                                rem_days = clean_num(raw_days)
                                        elif str(raw_days).strip() != "":
                                            rem_days = clean_num(raw_days)

                                        if rem_fh is None and rem_fc is None and rem_days is None:
                                            continue

                                        trigger = get_triggers(rem_fh, rem_fc, rem_days, fh_lim, fc_lim, days_lim)
                                        if trigger:
                                            results.append({
                                                "Reg": ac_reg,
                                                "Section": "TCI",
                                                "Task": curr_task,
                                                "Description": curr_desc,
                                                "PN": curr_pn,
                                                "SN": curr_sn,
                                                "Interval FH": "",
                                                "Interval FC": "",
                                                "Interval Days": "",
                                                "Remaining FH": rem_fh if rem_fh is not None else "",
                                                "Remaining FC": rem_fc if rem_fc is not None else "",
                                                "Remaining Days": rem_days if rem_days is not None else "",
                                                "Trigger": trigger,
                                                "Priority": get_priority(rem_fh, rem_fc, rem_days),
                                                "File Path": excel_path
                                            })
            except Exception as e:
                logger.error("Error processing TCI for %s: %s", ac_reg, e)

    finally:
        wb.close()

    return ac_reg, cur_afh, cur_afc, pd.DataFrame(results)

def fetch_ldnd_export_data(selected_df):
    export_rows = []
    grouped_tasks = selected_df.groupby("File Path")

    xw_app = xw.App(visible=False, add_book=False)
    xw_app.display_alerts = False
    xw_app.screen_updating = False

    try:
        amp_lookup = {}
        if os.path.exists(AMP_FILE_PATH):
            try:
                wb_amp = open_wb_with_xlwings(xw_app, AMP_FILE_PATH)
                if wb_amp:
                    for sheet_obj in wb_amp.sheets:
                        df_amp = get_sheet_dataframe(wb_amp, sheet_obj.name)
                        if not df_amp.empty and df_amp.shape[1] > 20:
                            for _, r_amp in df_amp.iterrows():
                                t_num = str(r_amp.iloc[2]).strip() if pd.notna(r_amp.iloc[2]) else ""
                                skill_code = str(r_amp.iloc[8]).strip() if pd.notna(r_amp.iloc[8]) else ""
                                mh_val = r_amp.iloc[20] if pd.notna(r_amp.iloc[20]) else ""
                                
                                if t_num and "TASK" not in t_num.upper() and str(mh_val).strip().upper() != "TASKM.H.":
                                    key = t_num.upper()
                                    cell_total = clean_and_sum_cell_mh(mh_val)
                                    
                                    if key not in amp_lookup:
                                        amp_lookup[key] = {"mh": 0.0, "skill_code": skill_code}
                                    
                                    amp_lookup[key]["mh"] += cell_total
                                    if not amp_lookup[key]["skill_code"] and skill_code:
                                        amp_lookup[key]["skill_code"] = skill_code
                        
                    wb_amp.close()
            except Exception as e:
                logger.error("Error reading AMP file (%s): %s", AMP_FILE_PATH, e)

        for file_path, group in grouped_tasks:
            wb = open_wb_with_xlwings(xw_app, file_path)
            if not wb:
                continue

            try:
                ldnd_lookup = {}
                df_ldnd_full = get_sheet_dataframe(wb, "LDND")
                if not df_ldnd_full.empty and len(df_ldnd_full) > 6:
                    df_ldnd = df_ldnd_full.iloc[6:].reset_index(drop=True)
                    for _, r_ldnd in df_ldnd.iterrows():
                        task_num = str(r_ldnd.iloc[2]).strip() if pd.notna(r_ldnd.iloc[2]) else ""
                        if task_num:
                            desc_val = str(r_ldnd.iloc[5]).replace("\n", " ").strip() if pd.notna(r_ldnd.iloc[5]) else ""
                            code_val = str(r_ldnd.iloc[6]).strip() if pd.notna(r_ldnd.iloc[6]) else ""
                            ldnd_lookup[task_num] = {
                                "code": code_val,
                                "desc": desc_val
                            }

                for _, row in group.iterrows():
                    task_number = str(row.get("Task", "")).strip()
                    ui_desc = str(row.get("Description", "")).strip()
                    ui_section = str(row.get("Section", "")).strip()
                    ui_reg = str(row.get("Reg", "")).strip()
                    ui_pn = clean_code_str(row.get("PN", ""))
                    ui_sn = clean_code_str(row.get("SN", ""))

                    sheet_match = None
                    for sheet_obj in wb.sheets:
                        s_name = sheet_obj.name.strip()
                        if s_name.upper() == task_number.upper():
                            sheet_match = s_name
                            break

                    if sheet_match:
                        task_info = ldnd_lookup.get(task_number, {"code": "", "desc": ""})
                        parent_code = task_info["code"]
                        parent_desc = task_info["desc"] if task_info["desc"] else ui_desc

                        amp_info = amp_lookup.get(task_number.upper(), {"mh": 0.0, "skill_code": ""})
                        parent_mh = f"{round(amp_info['mh'], 4):g}" if amp_info['mh'] > 0 else ""
                        parent_skill = amp_info["skill_code"]
                        parent_trade = map_skill_code_to_trade(parent_skill)

                        export_rows.append({
                            "Reg": ui_reg,
                            "Section": ui_section,
                            "Task Code": parent_code,
                            "TASK Number": task_number,
                            "Description": parent_desc,
                            "Part Number (P/N)": ui_pn,
                            "Serial Number (S/N)": ui_sn,
                            "Skill Code": parent_skill,
                            "Trade Skill": parent_trade,
                            "Man Hours": parent_mh,
                            "Is Parent": True
                        })

                        df_check_tab = get_sheet_dataframe(wb, sheet_match)
                        if not df_check_tab.empty and len(df_check_tab) >= 7:
                            df_check_rows = df_check_tab.iloc[6:].reset_index(drop=True)
                            for _, r_tab in df_check_rows.iterrows():
                                sub_task_num = str(r_tab.iloc[2]).strip() if len(r_tab) > 2 and pd.notna(r_tab.iloc[2]) else ""
                                if sub_task_num and sub_task_num.upper() != "NONE":
                                    sub_desc = str(r_tab.iloc[5]).replace("\n", " ").strip() if len(r_tab) > 5 and pd.notna(r_tab.iloc[5]) else ""
                                    sub_code = str(r_tab.iloc[6]).strip() if len(r_tab) > 6 and pd.notna(r_tab.iloc[6]) else ""

                                    amp_info = amp_lookup.get(sub_task_num.upper(), {"mh": 0.0, "skill_code": ""})
                                    sub_mh = f"{round(amp_info['mh'], 4):g}" if amp_info['mh'] > 0 else ""
                                    sub_skill = amp_info["skill_code"]
                                    sub_trade = map_skill_code_to_trade(sub_skill)

                                    export_rows.append({
                                        "Reg": ui_reg,
                                        "Section": ui_section,
                                        "Task Code": sub_code,
                                        "TASK Number": sub_task_num,
                                        "Description": sub_desc if sub_desc else ui_desc,
                                        "Part Number (P/N)": "",
                                        "Serial Number (S/N)": "",
                                        "Skill Code": sub_skill,
                                        "Trade Skill": sub_trade,
                                        "Man Hours": sub_mh,
                                        "Is Parent": False
                                    })
                    else:
                        task_info = ldnd_lookup.get(task_number, {"code": "", "desc": ""})
                        final_code = task_info["code"]
                        final_desc = task_info["desc"] if task_info["desc"] else ui_desc

                        amp_info = amp_lookup.get(task_number.upper(), {"mh": 0.0, "skill_code": ""})
                        final_mh = f"{round(amp_info['mh'], 4):g}" if amp_info['mh'] > 0 else ""
                        final_skill = amp_info["skill_code"]
                        final_trade = map_skill_code_to_trade(final_skill)

                        export_rows.append({
                            "Reg": ui_reg,
                            "Section": ui_section,
                            "Task Code": final_code,
                            "TASK Number": task_number,
                            "Description": final_desc,
                            "Part Number (P/N)": ui_pn,
                            "Serial Number (S/N)": ui_sn,
                            "Skill Code": final_skill,
                            "Trade Skill": final_trade,
                            "Man Hours": final_mh,
                            "Is Parent": False
                        })

            finally:
                wb.close()

    finally:
        _safe_quit_xw_app(xw_app)

    return pd.DataFrame(export_rows)

def build_bow_dataframe(export_df):
    julian_5_digit = datetime.now().strftime("%y%j")
    bow_rows = []

    filtered_df = export_df
    if "Is Parent" in export_df.columns:
        filtered_df = export_df[export_df["Is Parent"] != True]

    for _, row in filtered_df.iterrows():
        full_reg = str(row.get("Reg", "")).strip()
        reg_no_su = full_reg.replace("SU-", "").replace("su-", "").strip()
        task_number = str(row.get("TASK Number", "")).strip()

        formatted_code = f"A320-MSC/{reg_no_su}/{julian_5_digit}/1"

        bow_rows.append({
            "A/C": reg_no_su,
            "Event": task_number,
            "W/O-Flags": formatted_code
        })

    return pd.DataFrame(bow_rows)

def generate_single_sheet_excel(df, sheet_name):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        worksheet = writer.sheets[sheet_name]
        
        if not df.empty:
            max_row = len(df) + 1
            max_col = len(df.columns)
            max_col_letter = get_column_letter(max_col)
            table_range = f"A1:{max_col_letter}{max_row}"
            
            clean_table_name = "".join([c for c in sheet_name if c.isalnum()]) + "Table"
            
            tab = Table(displayName=clean_table_name, ref=table_range)
            
            style = TableStyleInfo(
                name="TableStyleMedium9",
                showFirstColumn=False,
                showLastColumn=False,
                showRowStripes=True,
                showColumnStripes=False
            )
            tab.tableStyleInfo = style
            worksheet.add_table(tab)
            
            for col in worksheet.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = get_column_letter(col[0].column)
                worksheet.column_dimensions[col_letter].width = max(max_len + 4, 12)
                
    buffer.seek(0)
    return buffer

def clean_code_str(val):
    if pd.isna(val) or val is None:
        return ""
    s = str(val).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s

# ==========================================
# REPORT MANAGEMENT HELPER FUNCTIONS
# ==========================================
def get_current_user_email():
    return st.session_state.get("user_email", "System")

def get_saved_reports_list():
    """Scans SAVED_DIR for Excel reports and calculates task compliance, B1/B2, and Total Hours metadata."""
    if not os.path.exists(SAVED_DIR):
        return pd.DataFrame()

    reports_data = []
    current_user = get_current_user_email()

    try:
        for fname in os.listdir(SAVED_DIR):
            if fname.endswith((".xlsx", ".xls")) and not fname.startswith("~$"):
                fpath = os.path.join(SAVED_DIR, fname)
                json_path = fpath + ".json"
                
                meta = {}
                if os.path.exists(json_path):
                    try:
                        with open(json_path, "r") as mf:
                            meta = json.load(mf)
                    except Exception as e:
                        logger.warning("Could not read metadata sidecar %s: %s", json_path, e)

                file_stat = os.stat(fpath)
                mod_time = datetime.fromtimestamp(file_stat.st_mtime).strftime("%Y-%m-%d %H:%M:%S")

                tasks_meta = meta.get("tasks_compliance", {})
                
                total_tasks = 0
                completed_tasks = 0
                total_mh = 0.0
                completed_mh = 0.0
                b1_total_mh = 0.0
                b1_completed_mh = 0.0
                b2_total_mh = 0.0
                b2_completed_mh = 0.0

                last_completed_date = ""
                last_completed_by = ""
                latest_timestamp = ""

                # Try reading actual Excel content to extract tasks and man hours
                try:
                    excel_df = read_excel_fresh(fpath)
                    total_tasks = len(excel_df)

                    has_mh = "Man Hours" in excel_df.columns
                    
                    for idx, r in excel_df.iterrows():
                        task_id = str(r.get("TASK Number", r.get("Task", f"Task_{idx+1}"))).strip()
                        task_key = f"{idx}_{task_id}"
                        t_info = tasks_meta.get(task_key, {})
                        is_done = bool(t_info.get("done", False))

                        if is_done:
                            completed_tasks += 1
                            t_updated = t_info.get("updated_at", "")
                            if t_updated >= latest_timestamp:
                                latest_timestamp = t_updated
                                last_completed_date = t_info.get("done_date", "")
                                last_completed_by = t_info.get("checked_by", "")

                        mh_val = parse_mh_to_float(r.get("Man Hours")) if has_mh else 0.0

                        if "Trade Skill" in excel_df.columns:
                            trade_str = str(r.get("Trade Skill", ""))
                        elif "Skill Code" in excel_df.columns:
                            trade_str = map_skill_code_to_trade(r.get("Skill Code", ""))
                        else:
                            trade_str = "B1 (Mechanical)"

                        total_mh += mh_val
                        if is_done:
                            completed_mh += mh_val

                        if "B2" in trade_str.upper():
                            b2_total_mh += mh_val
                            if is_done:
                                b2_completed_mh += mh_val
                        else:
                            b1_total_mh += mh_val
                            if is_done:
                                b1_completed_mh += mh_val

                except Exception as e:
                    # Excel file couldn't be read (locked, corrupted, wrong format, etc.) -
                    # fall back to whatever we already recorded in the JSON sidecar.
                    logger.warning("Could not read report Excel '%s' for stats, using sidecar fallback: %s", fpath, e)
                    total_tasks = meta.get("total_tasks_count", len(tasks_meta))
                    for t_key, t_info in tasks_meta.items():
                        if t_info.get("done", False):
                            completed_tasks += 1
                            t_updated = t_info.get("updated_at", "")
                            if t_updated >= latest_timestamp:
                                latest_timestamp = t_updated
                                last_completed_date = t_info.get("done_date", "")
                                last_completed_by = t_info.get("checked_by", "")

                rem_tasks = total_tasks - completed_tasks
                rem_mh = max(0.0, total_mh - completed_mh)
                b1_rem_mh = max(0.0, b1_total_mh - b1_completed_mh)
                b2_rem_mh = max(0.0, b2_total_mh - b2_completed_mh)

                is_fully_done = (total_tasks > 0) and (completed_tasks == total_tasks)

                # Format human-readable summary columns
                tasks_disp = f"{completed_tasks} / {total_tasks} ({rem_tasks} rem)"
                
                if total_mh > 0:
                    total_mh_disp = f"{rem_mh:.1f}h rem ({total_mh:.1f}h total)"
                    b1_mh_disp = f"{b1_rem_mh:.1f}h rem ({b1_total_mh:.1f}h total)"
                    b2_mh_disp = f"{b2_rem_mh:.1f}h rem ({b2_total_mh:.1f}h total)"
                else:
                    total_mh_disp = "N/A"
                    b1_mh_disp = "N/A"
                    b2_mh_disp = "N/A"

                reports_data.append({
                    "filename": meta.get("filename", fname),
                    "file_path": fpath,
                    "report_type": meta.get("report_type", "Excel Work Package"),
                    "created_by": meta.get("created_by", current_user),
                    "created_at": meta.get("created_at", mod_time),
                    "total_tasks": total_tasks,
                    "completed_tasks": completed_tasks,
                    "rem_tasks": rem_tasks,
                    "total_tasks_display": tasks_disp,
                    "total_mh": total_mh,
                    "rem_mh": rem_mh,
                    "total_mh_display": total_mh_disp,
                    "b1_total_mh": b1_total_mh,
                    "b1_rem_mh": b1_rem_mh,
                    "b1_mh_display": b1_mh_disp,
                    "b2_total_mh": b2_total_mh,
                    "b2_rem_mh": b2_rem_mh,
                    "b2_mh_display": b2_mh_disp,
                    "is_fully_done": is_fully_done,
                    "last_completed_date": last_completed_date,
                    "last_completed_by": last_completed_by,
                    "status_label": "✅ COMPLIANT" if is_fully_done else f"⏳ IN PROGRESS ({completed_tasks}/{total_tasks})",
                    "published_to_production": meta.get("published_to_production", False),
                    "published_at": meta.get("published_at", ""),
                    "published_by": meta.get("published_by", "")
                })
    except Exception as e:
        logger.error("Error scanning saved reports directory: %s", e)

    df = pd.DataFrame(reports_data)
    if not df.empty and "created_at" in df.columns:
        df = df.sort_values(by="created_at", ascending=False)

    return df


def save_task_compliance(file_path, edited_df, current_user_email):
    """Saves row-by-row task compliance edits and man hours to the JSON sidecar metadata file."""
    meta_path = file_path + ".json"
    meta_data = {}

    if os.path.exists(meta_path):
        try:
            with open(meta_path, "r") as mf:
                meta_data = json.load(mf)
        except Exception as e:
            # If the sidecar is unreadable we fall back to an empty dict, which
            # means prior compliance history could be overwritten below - log
            # this loudly since it's a potential silent data-loss scenario.
            logger.error("Compliance sidecar %s is unreadable, starting fresh: %s", meta_path, e)

    existing_tasks = meta_data.get("tasks_compliance", {})
    new_tasks_meta = {}

    total_tasks = len(edited_df)
    completed_count = 0
    latest_update_time = ""
    last_user = ""
    last_date = ""

    total_mh = 0.0
    completed_mh = 0.0
    b1_total_mh = 0.0
    b1_completed_mh = 0.0
    b2_total_mh = 0.0
    b2_completed_mh = 0.0

    has_mh = "Man Hours" in edited_df.columns

    for idx, row in edited_df.iterrows():
        task_id = str(row.get("TASK Number", row.get("Task", f"Task_{idx+1}"))).strip()
        task_key = f"{idx}_{task_id}"

        is_done = bool(row.get("Done", False))
        raw_date = row.get("Completion Date", None)

        if raw_date and pd.notna(raw_date) and str(raw_date).strip() != "":
            done_date_str = str(raw_date).split(" ")[0]
        else:
            done_date_str = datetime.now().strftime("%Y-%m-%d") if is_done else ""

        prev_meta = existing_tasks.get(task_key, {})
        was_done = prev_meta.get("done", False)

        if is_done and not was_done:
            checked_by = current_user_email
            updated_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        elif is_done:
            checked_by = prev_meta.get("checked_by", current_user_email)
            updated_at = prev_meta.get("updated_at", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        else:
            checked_by = ""
            updated_at = ""

        new_tasks_meta[task_key] = {
            "done": is_done,
            "done_date": done_date_str,
            "checked_by": checked_by,
            "updated_at": updated_at
        }

        mh_val = parse_mh_to_float(row.get("Man Hours")) if has_mh else 0.0
        if "Trade Skill" in edited_df.columns:
            trade_str = str(row.get("Trade Skill", ""))
        elif "Skill Code" in edited_df.columns:
            trade_str = map_skill_code_to_trade(row.get("Skill Code", ""))
        else:
            trade_str = "B1 (Mechanical)"

        total_mh += mh_val
        if is_done:
            completed_count += 1
            completed_mh += mh_val
            if updated_at >= latest_update_time:
                latest_update_time = updated_at
                last_user = checked_by
                last_date = done_date_str

        if "B2" in trade_str.upper():
            b2_total_mh += mh_val
            if is_done:
                b2_completed_mh += mh_val
        else:
            b1_total_mh += mh_val
            if is_done:
                b1_completed_mh += mh_val

    is_all_done = (total_tasks > 0) and (completed_count == total_tasks)

    meta_data["tasks_compliance"] = new_tasks_meta
    meta_data["total_tasks_count"] = total_tasks
    meta_data["completed_tasks_count"] = completed_count
    meta_data["total_mh"] = total_mh
    meta_data["completed_mh"] = completed_mh
    meta_data["remaining_mh"] = max(0.0, total_mh - completed_mh)
    meta_data["b1_total_mh"] = b1_total_mh
    meta_data["b1_completed_mh"] = b1_completed_mh
    meta_data["b1_remaining_mh"] = max(0.0, b1_total_mh - b1_completed_mh)
    meta_data["b2_total_mh"] = b2_total_mh
    meta_data["b2_completed_mh"] = b2_completed_mh
    meta_data["b2_remaining_mh"] = max(0.0, b2_total_mh - b2_completed_mh)
    meta_data["is_fully_compliant"] = is_all_done
    meta_data["final_completion_date"] = last_date if is_all_done else ""
    meta_data["final_completed_by"] = last_user if is_all_done else ""
    meta_data["last_updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    try:
        write_json_atomic(meta_path, meta_data)
        return True, "Task compliance updates saved successfully!"
    except Exception as e:
        logger.error("Failed to save compliance updates for %s: %s", file_path, e)
        return False, f"Failed to save compliance updates: {str(e)}"


def update_report_compliance(file_path, is_compliant, compliance_date, user_email):
    """Updates or creates compliance metadata in the report's JSON sidecar file."""
    meta_path = file_path + ".json"
    meta_data = {}

    if os.path.exists(meta_path):
        try:
            with open(meta_path, "r") as mf:
                meta_data = json.load(mf)
        except Exception as e:
            logger.error("Compliance sidecar %s is unreadable, starting fresh: %s", meta_path, e)

    meta_data["is_compliant"] = is_compliant
    meta_data["compliance_date"] = str(compliance_date) if compliance_date else ""
    meta_data["compliant_by"] = user_email
    meta_data["compliance_updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    try:
        write_json_atomic(meta_path, meta_data)
        return True, "Compliance status saved successfully!"
    except Exception as e:
        logger.error("Failed to save compliance status for %s: %s", file_path, e)
        return False, f"Failed to save compliance status: {str(e)}"

def delete_saved_report(filename, file_path):
    """
    Deletes a report and its metadata sidecar. `file_path` must already be a
    path resolved from the saved-reports listing (never built directly from
    free-text user input), and is re-validated here to stay inside SAVED_DIR.
    """
    try:
        real_saved_dir = os.path.realpath(SAVED_DIR)
        real_target = os.path.realpath(file_path)
        if os.path.commonpath([real_saved_dir, real_target]) != real_saved_dir:
            logger.error("Blocked delete attempt outside SAVED_DIR: %s", file_path)
            return False, "Invalid file location."

        meta_path = file_path + ".json"
        if os.path.exists(file_path):
            os.remove(file_path)
        if os.path.exists(meta_path):
            os.remove(meta_path)
        return True, f"Deleted '{filename}' successfully!"
    except Exception as e:
        logger.error("Failed to delete file %s: %s", file_path, e)
        return False, f"Failed to delete file: {str(e)}"

def save_report(file_bytes, filename, report_type, user_email=None):
    """
    Saves an in-memory report to SAVED_DIR. The user-supplied `filename` is
    sanitized first (see sanitize_filename) so it can never contain path
    separators or '..' segments that would let it write outside SAVED_DIR.
    """
    clean_filename = sanitize_filename(filename, fallback=f"Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
    if not clean_filename.lower().endswith(".xlsx"):
        clean_filename += ".xlsx"

    base_name = clean_filename[:-5]
    file_path = os.path.join(SAVED_DIR, clean_filename)
    counter = 1
    while os.path.exists(file_path):
        clean_filename = f"{base_name} ({counter}).xlsx"
        file_path = os.path.join(SAVED_DIR, clean_filename)
        counter += 1

    meta_path = file_path + ".json"

    actual_user = user_email if (user_email and user_email != "System") else get_current_user_email()
    created_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    try:
        with open(file_path, "wb") as f:
            f.write(file_bytes.getvalue())

        meta_data = {
            "filename": clean_filename,
            "report_type": report_type,
            "created_by": actual_user,
            "created_at": created_time,
            "file_path": file_path,
            "published_to_production": False,
            "published_at": "",
            "published_by": ""
        }
        write_json_atomic(meta_path, meta_data)

        return True, f"Report successfully saved as '{clean_filename}'"
    except Exception as e:
        logger.error("Failed to save report '%s': %s", filename, e)
        return False, f"Failed to save report: {str(e)}"

def rename_saved_report(old_filename, new_filename):
    """
    Renames a saved report. `new_filename` is user-typed free text, so it is
    sanitized first (see sanitize_filename) to prevent path traversal - e.g.
    typing "../../secrets" can no longer make the app write outside SAVED_DIR.
    """
    clean_new_name = sanitize_filename(new_filename, fallback=old_filename)
    if not clean_new_name.lower().endswith(".xlsx"):
        clean_new_name += ".xlsx"

    if clean_new_name == old_filename:
        return True, "No change made to file name."

    base_name = clean_new_name[:-5]
    new_path = os.path.join(SAVED_DIR, clean_new_name)
    counter = 1
    while os.path.exists(new_path) and clean_new_name != old_filename:
        clean_new_name = f"{base_name} ({counter}).xlsx"
        new_path = os.path.join(SAVED_DIR, clean_new_name)
        counter += 1

    old_path = os.path.join(SAVED_DIR, old_filename)
    old_meta = old_path + ".json"
    new_meta = new_path + ".json"

    try:
        if os.path.exists(old_path):
            os.rename(old_path, new_path)
        if os.path.exists(old_meta):
            os.rename(old_meta, new_meta)
            with open(new_meta, "r") as mf:
                data = json.load(mf)
            data["filename"] = clean_new_name
            data["file_path"] = new_path
            write_json_atomic(new_meta, data)
        return True, f"Renamed to '{clean_new_name}' successfully!"
    except Exception as e:
        logger.error("Failed to rename '%s' to '%s': %s", old_filename, new_filename, e)
        return False, f"Failed to rename file: {str(e)}"

# ==========================================
# REPORT PUBLISH / RECALL HELPERS
# ==========================================
def publish_report_to_production(file_path: str, user_email: str):
    """
    Marks a saved report as published to Production by setting the
    `published_to_production` flag in its JSON sidecar to True.
    Until this is called, the report is invisible to production users.
    """
    meta_path = file_path + ".json"
    meta_data = {}
    if os.path.exists(meta_path):
        try:
            with open(meta_path, "r") as mf:
                meta_data = json.load(mf)
        except Exception as e:
            logger.error("Could not read sidecar %s before publish: %s", meta_path, e)
    meta_data["published_to_production"] = True
    meta_data["published_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    meta_data["published_by"] = user_email
    try:
        write_json_atomic(meta_path, meta_data)
        return True, "Report sent to Production successfully!"
    except Exception as e:
        logger.error("Failed to publish report %s: %s", file_path, e)
        return False, f"Failed to send report to Production: {str(e)}"


def recall_report_from_production(file_path: str, user_email: str):
    """
    Recalls a report from the Production Workspace by setting
    `published_to_production` back to False. The report returns to draft state
    and is no longer visible to production users in Tab 2.
    """
    meta_path = file_path + ".json"
    meta_data = {}
    if os.path.exists(meta_path):
        try:
            with open(meta_path, "r") as mf:
                meta_data = json.load(mf)
        except Exception as e:
            logger.error("Could not read sidecar %s before recall: %s", meta_path, e)
    meta_data["published_to_production"] = False
    meta_data["published_at"] = ""
    meta_data["published_by"] = ""
    try:
        write_json_atomic(meta_path, meta_data)
        return True, "Report recalled from Production successfully!"
    except Exception as e:
        logger.error("Failed to recall report %s: %s", file_path, e)
        return False, f"Failed to recall report from Production: {str(e)}"


# ==========================================
# HEADER: CURRENT DATE, JULIAN, AND LIVE TIME
# ==========================================
import streamlit.components.v1 as components

now_dt = datetime.now()
date_str = now_dt.strftime("%d-%b-%Y").upper()
julian_str = now_dt.strftime("%y%j")

html_code = f"""
<!DOCTYPE html>
<html>
<head>
    <style>
        body {{
            margin: 0; padding: 0; font-family: 'Inter', sans-serif; background-color: transparent; overflow: hidden;
        }}
        .banner {{
            display: flex; justify-content: flex-end; align-items: center; gap: 24px;
            padding: 10px 18px; margin-bottom: 2px;
            background: linear-gradient(to right, #ffffff, #f8fafc);
            border-radius: 8px; border-left: 4px solid #0284c7;
            box-shadow: 0 1px 3px rgba(0,0,0,0.05);
            border-top: 1px solid #f1f5f9; border-right: 1px solid #f1f5f9; border-bottom: 1px solid #f1f5f9;
        }}
        .item {{ display: flex; align-items: center; gap: 8px; }}
        .icon {{ font-size: 1.1rem; }}
        .label {{ font-size: 0.75rem; color: #64748b; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; }}
        .value {{ font-family: monospace; font-size: 1rem; font-weight: 600; color: #0f172a; }}
        .divider {{ width: 1px; height: 24px; background-color: #cbd5e1; }}
    </style>
</head>
<body>
    <div class="banner">
        <div class="item">
            <span class="icon">📅</span>
            <span class="label">Date</span>
            <span class="value">{date_str}</span>
        </div>
        <div class="divider"></div>
        <div class="item">
            <span class="icon">✈️</span>
            <span class="label">Julian</span>
            <span class="value">{julian_str}</span>
        </div>
        <div class="divider"></div>
        <div class="item">
            <span class="icon">🕒</span>
            <span class="label">Time</span>
            <span class="value" id="time-val">--:--:--</span>
        </div>
    </div>
    <script>
        function updateClock() {{
            const now = new Date();
            const hours = String(now.getHours()).padStart(2, '0');
            const minutes = String(now.getMinutes()).padStart(2, '0');
            const seconds = String(now.getSeconds()).padStart(2, '0');
            document.getElementById('time-val').textContent = hours + ':' + minutes + ':' + seconds;
        }}
        setInterval(updateClock, 1000);
        updateClock();
    </script>
</body>
</html>
"""
components.html(html_code, height=65)

# ==========================================
# MAIN CONTENT TABS
# ==========================================
tab1, tab2, tab3 = st.tabs(["🛫 Fleet Review", "⚙️ Production Workspace", "🛠️ Admin Tools"])

with tab3:
    if st.session_state.role == "admin":
        with st.expander("👥 **User Management & Role Assignment**", expanded=False):
            st.markdown("### Registered Users & Access Permissions")
            users_list = fetch_all_user_profiles()
        
            if users_list:
                u_df = pd.DataFrame(users_list)
            
                for col in ["email", "role", "id"]:
                    if col not in u_df.columns:
                        u_df[col] = ""

                st.dataframe(
                    u_df[["email", "role"]].rename(columns={
                        "email": "User Email Address",
                        "role": "Assigned System Role"
                    }),
                    use_container_width=True,
                    hide_index=True
                )

                st.markdown("---")
                st.markdown("#### Modify User Role")
                col_u1, col_u2, col_u3 = st.columns([2.5, 2, 1.5])

                user_options = {u.get("email", u["id"]): u for u in users_list if u.get("email")}

                with col_u1:
                    selected_user_email = st.selectbox(
                        "Select Target User:",
                        options=list(user_options.keys()),
                        key="admin_user_select"
                    )

                current_u_role = user_options[selected_user_email].get("role", "viewer").lower()

                with col_u2:
                    role_index = 0
                    if current_u_role == "viewer":
                        role_index = 0
                    elif current_u_role == "planner":
                        role_index = 1
                    elif current_u_role == "production":
                        role_index = 2
                    elif current_u_role == "admin":
                        role_index = 3

                    assigned_role = st.selectbox(
                        "Assign New Role:",
                        options=["viewer", "planner", "production", "admin"],
                        index=role_index,
                        key="admin_role_select"
                    )

                with col_u3:
                    st.write("")
                    st.write("")
                    if st.button("Update Role", type="primary", use_container_width=True):
                        target_u = user_options[selected_user_email]
                        ok, msg = update_user_role_in_db(target_u["id"], assigned_role)
                        if ok:
                            st.toast(msg, icon="✅")
                            st.rerun()
                        else:
                            st.error(msg)
    else:
        st.info("🔒 User management is available to Admin accounts only.")

with tab1:
    if st.session_state.role in ["admin", "planner"]:
        if run_button:
            st.session_state.preview_df = None
            st.session_state.export_mode = None

            if not os.path.exists(folder_path):
                st.error(f"The root folder path specified does not exist: {folder_path}")
            else:
                all_results = []
                aircraft_stats = []

                xw_app = xw.App(visible=False, add_book=False)
                xw_app.display_alerts = False
                xw_app.screen_updating = False

                try:
                    with st.spinner("Processing fleet files..."):
                        for root, dirs, files in os.walk(folder_path):
                            folder_name = os.path.basename(root).upper()
                        
                            if target_reg == "ALL" or target_reg in folder_name:
                                for file in files:
                                    file_upper = file.upper()
                                    if (file_upper.endswith((".XLSX", ".XLSM", ".XLS")) 
                                        and not file.startswith("~$") 
                                        and "MASTER" not in file_upper):
                                    
                                        file_p = os.path.join(root, file)
                                        try:
                                            reg, afh, afc, df_res = process_single_aircraft(
                                                file_p, xw_app, fh_limit, fc_limit, days_limit, tci_path_input
                                            )
                                        except Exception as e:
                                            # One bad/corrupted workbook shouldn't abort the whole
                                            # fleet scan - log it and keep going with the rest.
                                            logger.error("Failed to process %s: %s", file_p, e)
                                            st.warning(f"⚠️ Skipped '{file}' - could not be processed. See logs for details.")
                                            continue

                                        if reg and "MASTER" not in reg.upper():
                                            aircraft_stats.append({"Aircraft Reg": reg, "Updated AFH": afh, "Updated AFC": afc})
                                            if not df_res.empty:
                                                all_results.append(df_res)
                                            break
                finally:
                    _safe_quit_xw_app(xw_app)

                st.session_state.all_results = all_results
                st.session_state.aircraft_stats = aircraft_stats

        if st.session_state.aircraft_stats:
            st.subheader("✈️ Active Fleet Status Overview")
            cols = st.columns(min(len(st.session_state.aircraft_stats), 4))
            for i, stat in enumerate(st.session_state.aircraft_stats):
                with cols[i % 4]:
                    st.markdown(f"""
                        <div class="metric-card">
                            <div class="metric-title">AIRCRAFT REGISTRATION</div>
                            <div class="metric-value">{stat['Aircraft Reg']}</div>
                            <div class="metric-sub">⏱️ {stat['Updated AFH']:,} FH &nbsp;|&nbsp; 🔄 {stat['Updated AFC']:,} FC</div>
                        </div>
                    """, unsafe_allow_html=True)
            st.write("")

        if st.session_state.all_results:
            final_df = pd.concat(st.session_state.all_results, ignore_index=True)

            interval_cols = ["Interval FH", "Interval FC", "Interval Days"]
            final_df = final_df.drop(columns=[col for col in interval_cols if col in final_df.columns])
        
            final_df["EST Remaining Days"] = final_df.apply(calculate_est_remaining_days, axis=1)

            def get_est_due_date_obj(row):
                est_days = row.get("EST Remaining Days")
                if pd.isna(est_days) or est_days is None or str(est_days).strip() == "":
                    return pd.NaT
                try:
                    days_to_add = int(round(float(est_days)))
                    return datetime.now().replace(hour=0, minute=0, second=0, microsecond=0) + timedelta(days=days_to_add)
                except (ValueError, TypeError, OverflowError):
                    return pd.NaT

            final_df["EST Due Date"] = final_df.apply(get_est_due_date_obj, axis=1)

            num_cols = ["Remaining FH", "Remaining FC", "Remaining Days", "EST Remaining Days"]
            for col in num_cols:
                if col in final_df.columns:
                    final_df[col] = pd.to_numeric(final_df[col], errors='coerce')

            if "PN" in final_df.columns:
                final_df["PN"] = final_df["PN"].apply(clean_code_str)
            if "SN" in final_df.columns:
                final_df["SN"] = final_df["SN"].apply(clean_code_str)

            if selected_sections and "ALL" not in selected_sections:
                final_df = final_df[final_df["Section"].isin(selected_sections)].reset_index(drop=True)

            # Search & Priority Filter Controls
            st.markdown("### 🔍 Maintenance Review Items")
            c_f1, c_f2 = st.columns([2, 2])
            with c_f1:
                search_q = st.text_input(
                    "Filter by Task Number or Description:",
                    value=st.session_state.search_filter,
                    placeholder="e.g. NDT-42 or 'landing gear'"
                )
                st.session_state.search_filter = search_q
            with c_f2:
                prio_q = st.multiselect(
                    "Priority Filter:",
                    ["CRITICAL", "HIGH", "MEDIUM", "LOW"],
                    default=st.session_state.priority_filter
                )
                st.session_state.priority_filter = prio_q

            if search_q:
                final_df = final_df[
                    final_df["Task"].astype(str).str.contains(search_q, case=False, na=False) |
                    final_df["Description"].astype(str).str.contains(search_q, case=False, na=False)
                ]
            if prio_q and "Priority" in final_df.columns:
                final_df = final_df[final_df["Priority"].isin(prio_q)]

            final_df.index = range(1, len(final_df) + 1)

            col_hdr, col_chk = st.columns([4, 1])
            with col_hdr:
                st.caption(f"Showing **{len(final_df)}** matching maintenance items")
            with col_chk:
                select_all = st.checkbox("Select All Tasks", key="chk_select_all", value=st.session_state.select_all_state)
                st.session_state.select_all_state = select_all

            final_df.insert(0, "Select", select_all)

            column_order = [
                "Select", "Reg", "Section", "Task", "PN", "SN", "Description",
                "Remaining FH", "Remaining FC", "Remaining Days",
                "EST Remaining Days", "EST Due Date", "Trigger", "Priority", "File Path"
            ]
            final_cols = [c for c in column_order if c in final_df.columns]
            final_df = final_df[final_cols]
        
            edited_df = st.data_editor(
                final_df,
                use_container_width=True,
                height=450,
                key="main_task_editor",
                column_config={
                    "Select": st.column_config.CheckboxColumn(
                        "Select",
                        help="Select tasks to export or review",
                        default=False
                    ),
                    "Remaining FH": st.column_config.NumberColumn(
                        "Remaining FH",
                        format="%d"
                    ),
                    "Remaining FC": st.column_config.NumberColumn(
                        "Remaining FC",
                        format="%d"
                    ),
                    "Remaining Days": st.column_config.NumberColumn(
                        "Remaining Days",
                        format="%d"
                    ),
                    "EST Remaining Days": st.column_config.NumberColumn(
                        "EST Remaining Days",
                        format="%d"
                    ),
                    "EST Due Date": st.column_config.DateColumn(
                        "EST Due Date",
                        format="DD-MMM-YY"
                    ),
                    "PN": None,
                    "SN": None,
                    "File Path": None
                },
                disabled=[col for col in final_df.columns if col != "Select"]
            )

            selected_rows = edited_df[edited_df["Select"] == True]

            st.divider()
            st.subheader("🛠️ Actions & Reports")

            col_btn1, col_btn2 = st.columns(2)

            with col_btn1:
                btn_export_selected = st.button(
                    "📥 Preview & Export Selected Tasks",
                    type="primary",
                    use_container_width=True,
                    disabled=selected_rows.empty
                )

            with col_btn2:
                btn_export_bow = st.button(
                    "📄 Preview & Export BOW",
                    type="secondary",
                    use_container_width=True,
                    disabled=selected_rows.empty
                )

            if btn_export_selected:
                with st.spinner("Opening source file(s) and generating Selected Tasks report..."):
                    st.session_state.preview_df = fetch_ldnd_export_data(selected_rows)
                    st.session_state.export_mode = "SELECTED"

            if btn_export_bow:
                with st.spinner("Opening check sheet(s) and generating expanded BOW report..."):
                    st.session_state.preview_df = fetch_ldnd_export_data(selected_rows)
                    st.session_state.export_mode = "BOW"

        if st.session_state.preview_df is not None and st.session_state.export_mode:
            date_str = datetime.now().strftime("%d-%m-%y")

            # --- SELECTED TASKS SECTION ---
            if st.session_state.export_mode == "SELECTED":
                st.subheader("📋 Selected Tasks Preview & Trade Skill Breakdown")
            
                selected_display_df = st.session_state.preview_df.drop(columns=["Is Parent"], errors="ignore").copy()

                if "Skill Code" in selected_display_df.columns:
                    selected_display_df["Trade Skill"] = selected_display_df["Skill Code"].apply(map_skill_code_to_trade)

                edited_preview_df = st.data_editor(
                    selected_display_df,
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "Trade Skill": st.column_config.SelectboxColumn(
                            "Trade Skill",
                            help="Assigned license skill for task execution",
                            options=["B1 (Mechanical)", "B2 (Avionics)"],
                            required=True
                        )
                    },
                    disabled=[c for c in selected_display_df.columns if c != "Trade Skill"]
                )

                edited_preview_df["Parsed MH"] = edited_preview_df["Man Hours"].apply(parse_mh_to_float)

                b1_df = edited_preview_df[edited_preview_df["Trade Skill"] == "B1 (Mechanical)"]
                b2_df = edited_preview_df[edited_preview_df["Trade Skill"] == "B2 (Avionics)"]

                b1_hours = b1_df["Parsed MH"].sum()
                b2_hours = b2_df["Parsed MH"].sum()
                total_mh_decimal = b1_hours + b2_hours

                def format_hm(val_float):
                    hrs = int(val_float)
                    mins = int(round((val_float - hrs) * 60))
                    return f"{hrs}h {mins}m"

                st.markdown("### 🛠️ Resource & Skill Allocation Summary")
                col_m1, col_m2, col_m3, col_m4 = st.columns(4)

                with col_m1:
                    st.metric(
                        label="⏱️ Total Work Package MH",
                        value=format_hm(total_mh_decimal),
                        delta=f"{total_mh_decimal:.2f} Decimal Hrs"
                    )
                with col_m2:
                    st.metric(
                        label="🔧 B1 Mechanical Hours",
                        value=format_hm(b1_hours),
                        delta=f"{(b1_hours / total_mh_decimal * 100 if total_mh_decimal > 0 else 0):.1f}% of Total"
                    )
                with col_m3:
                    st.metric(
                        label="⚡ B2 Avionics Hours",
                        value=format_hm(b2_hours),
                        delta=f"{(b2_hours / total_mh_decimal * 100 if total_mh_decimal > 0 else 0):.1f}% of Total"
                    )
                with col_m4:
                    st.metric(
                        label="📊 Total Tasks Count",
                        value=f"{len(edited_preview_df)} Tasks",
                        delta=f"{len(b1_df)} B1 / {len(b2_df)} B2"
                    )

                chart_data = pd.DataFrame({
                    "Trade Skill": ["B1 (Mechanical)", "B2 (Avionics)"],
                    "Man Hours": [b1_hours, b2_hours]
                }).set_index("Trade Skill")
            
                st.bar_chart(chart_data)

                st.divider()

                export_final_df = edited_preview_df.drop(columns=["Parsed MH"], errors="ignore")
                excel_data = generate_single_sheet_excel(export_final_df, "Selected Tasks")
                default_filename = f"Selected_Tasks_{date_str}.xlsx"

                col_dl, col_sv = st.columns(2)
                with col_dl:
                    st.download_button(
                        label=f"⬇️ Download Excel ({len(export_final_df)} items)",
                        data=excel_data,
                        file_name=default_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        use_container_width=True
                    )
                with col_sv, st.popover("💾 Save As to Portal History", use_container_width=True, key=f"popover_sel_{st.session_state.popover_key_sel}"):
                    st.markdown("### Save Report")
                    custom_name = st.text_input("Saved File Name:", value=default_filename, key="save_as_sel_name")
                    if st.button("Confirm Save", type="primary", use_container_width=True, key="btn_confirm_save_sel"):
                        ok, msg = save_report(excel_data, custom_name, "Selected Tasks", st.session_state.get("user_email"))
                        if ok:
                            st.session_state.popover_key_sel += 1
                            st.toast(f"✅ {msg}", icon="💾")
                            st.rerun()
                        else:
                            st.error(msg)

            # --- BOW SECTION ---
            elif st.session_state.export_mode == "BOW":
                st.subheader("📄 BOW Sheet Preview")
                bow_df = build_bow_dataframe(st.session_state.preview_df)
                st.dataframe(bow_df, use_container_width=True, hide_index=True)

                excel_data = generate_single_sheet_excel(bow_df, "BOW")

                unique_ac = bow_df["A/C"].unique()
                ac_tag = unique_ac[0] if len(unique_ac) == 1 else "FLEET"
                bow_filename = f"{ac_tag} BOW {date_str}.xlsx"

                col_dl, col_sv = st.columns(2)
                with col_dl:
                    st.download_button(
                        label=f"⬇️ Download BOW Excel ({len(bow_df)} items)",
                        data=excel_data,
                        file_name=bow_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        use_container_width=True
                    )
                with col_sv, st.popover("💾 Save As to Portal History", use_container_width=True, key=f"popover_bow_{st.session_state.popover_key_bow}"):
                    st.markdown("### Save BOW Report")
                    custom_name = st.text_input("Saved File Name:", value=bow_filename, key="save_as_bow_name")
                    if st.button("Confirm Save", type="primary", use_container_width=True, key="btn_confirm_save_bow"):
                        ok, msg = save_report(excel_data, custom_name, "BOW", st.session_state.get("user_email"))
                        if ok:
                            st.session_state.popover_key_bow += 1
                            st.toast(f"✅ {msg}", icon="💾")
                            st.rerun()
                        else:
                            st.error(msg)

                st.divider()
                st.subheader("📧 Send BOW File via Outlook")

                col_to, col_cc = st.columns(2)
                with col_to:
                    email_to = st.text_input("To:", value="aser.hafez@aircairo.com")
                with col_cc:
                    email_cc = st.text_input("CC (Optional):", value="")

                default_subject = f"BOW Report for {ac_tag}"
                default_body = (
                    "Dear Eng. Aser,\n\n"
                    f"Please find the attached BOW Report for {ac_tag} file for your reference and review.\n\n"
                    "Kindly review the document and let me know if you require any additional information or clarification.\n\n"
                    "Thank you, and I look forward to your feedback."
                )

                email_subject = st.text_input("Subject:", value=default_subject)
                email_body = st.text_area("Body:", value=default_body, height=150)

                extra_files = st.file_uploader(
                    "📎 Attach Additional File(s) (Optional):",
                    accept_multiple_files=True,
                    help="Select additional files from your computer to attach along with the BOW report."
                )

                if st.button("📤 Send Email via Outlook", type="primary", use_container_width=True):
                    with st.spinner("Connecting to Outlook and sending email..."):
                        success, msg = send_bow_via_outlook(
                            to_email=email_to,
                            cc_email=email_cc,
                            subject=email_subject,
                            body_text=email_body,
                            excel_bytes=excel_data,
                            filename=bow_filename,
                            extra_attachments=extra_files
                        )
                        if success:
                            st.success(msg)
                        else:
                            st.error(msg)

        elif not st.session_state.all_results and run_button:
            st.info("No maintenance checks due within the specified thresholds.")

        # ==========================================
        # MY SAVED REPORTS — Publish-Gate Panel
        # Always visible to Admin & Planner in Tab 1
        # ==========================================
        st.divider()
        st.subheader("📂 My Saved Reports")
        st.caption(
            "Reports saved here are in **Draft** and invisible to Production until you explicitly "
            "send them. Click **🚀 Send to Production** to make a report available in the Production Workspace tab."
        )

        planner_reports_df = get_saved_reports_list()
        current_user_tab1 = st.session_state.get("user_email", "")
        is_admin_tab1 = st.session_state.role == "admin"

        # All planners and admins see all reports in this view
        # No filtering by creator needed.

        if planner_reports_df.empty:
            st.info("📋 No saved reports yet. Generate and save a report above to get started.")
        else:
            for _, rep_row in planner_reports_df.iterrows():
                rep_name     = rep_row["filename"]
                rep_path     = rep_row["file_path"]
                rep_created  = rep_row["created_at"]
                rep_by       = rep_row["created_by"]
                rep_tasks    = rep_row.get("total_tasks_display", "—")
                is_published = bool(rep_row.get("published_to_production", False))
                pub_at       = rep_row.get("published_at", "")
                pub_by       = rep_row.get("published_by", "")

                with st.container(border=True):
                    col_info, col_badge, col_pub, col_dl, col_ren, col_del = st.columns([2.5, 1.5, 1.8, 1.2, 1, 1])

                    with col_info:
                        st.markdown(f"**📄 {rep_name}**")
                        st.caption(f"Created by **{rep_by}** on {rep_created} · Tasks: {rep_tasks}")
                        if is_published and pub_at:
                            st.caption(f"✅ Sent to Production by **{pub_by}** on {pub_at}")

                    with col_badge:
                        if is_published:
                            st.markdown(
                                "<div style='padding:6px 10px;background:#DCFCE7;border-radius:8px;"
                                "border:1px solid #86EFAC;text-align:center;'>"
                                "<span style='color:#166534;font-weight:700;font-size:0.8rem;'>"
                                "✅ SENT TO PRODUCTION</span></div>",
                                unsafe_allow_html=True
                            )
                        else:
                            st.markdown(
                                "<div style='padding:6px 10px;background:#FEF3C7;border-radius:8px;"
                                "border:1px solid #FCD34D;text-align:center;'>"
                                "<span style='color:#92400E;font-weight:700;font-size:0.8rem;'>"
                                "📝 DRAFT</span></div>",
                                unsafe_allow_html=True
                            )

                    is_creator_tab1 = current_user_tab1.lower() == rep_by.lower()
                    can_manage_pub = is_admin_tab1 or is_creator_tab1
                    can_edit = is_admin_tab1 or is_creator_tab1

                    with col_pub:
                        if not is_published:
                            if can_manage_pub:
                                if st.button(
                                    "🚀 Send to Production",
                                    key=f"publish_{rep_name}",
                                    type="primary",
                                    use_container_width=True
                                ):
                                    ok, msg = publish_report_to_production(
                                        rep_path, current_user_tab1
                                    )
                                    if ok:
                                        st.toast(f"✅ {msg}", icon="🚀")
                                        st.rerun()
                                    else:
                                        st.error(msg)
                            else:
                                st.button(
                                    "🚀 Send to Production",
                                    key=f"publish_{rep_name}",
                                    disabled=True,
                                    use_container_width=True
                                )
                        else:
                            if can_manage_pub:
                                if st.button(
                                    "↩️ Recall from Production",
                                    key=f"recall_{rep_name}",
                                    type="secondary",
                                    use_container_width=True
                                ):
                                    ok, msg = recall_report_from_production(
                                        rep_path, current_user_tab1
                                    )
                                    if ok:
                                        st.toast(f"↩️ {msg}", icon="↩️")
                                        st.rerun()
                                    else:
                                        st.error(msg)
                            else:
                                st.button(
                                    "↩️ Recall from Production",
                                    key=f"recall_{rep_name}",
                                    disabled=True,
                                    use_container_width=True
                                )

                    with col_dl:
                        if os.path.exists(rep_path):
                            with open(rep_path, "rb") as file_bytes:
                                dl_clicked = st.download_button(
                                    label="⬇️ Download",
                                    data=file_bytes,
                                    file_name=rep_name,
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True,
                                    key=f"t1_dl_{rep_name}"
                                )
                                if dl_clicked:
                                    log_download_activity(rep_name, current_user_tab1)
                                    st.toast(f"📥 Download logged for {rep_name}", icon="ℹ️")
                        else:
                            st.button("⬇️ Missing", disabled=True, use_container_width=True, key=f"t1_dl_miss_{rep_name}")

                    with col_ren:
                        if can_edit:
                            with st.popover("✏️ Rename", use_container_width=True, key=f"t1_rename_{rep_name}_{st.session_state.popover_key_rename}"):
                                st.markdown("### Rename Report")
                                new_name_in = st.text_input("New Name:", value=rep_name, key=f"t1_rename_in_{rep_name}")
                                if st.button("Confirm Rename", type="primary", use_container_width=True, key=f"t1_btn_ren_{rep_name}"):
                                    if new_name_in.strip() and new_name_in != rep_name:
                                        ok, msg = rename_saved_report(rep_name, new_name_in)
                                        if ok:
                                            st.session_state.popover_key_rename += 1
                                            st.toast(f"✅ {msg}", icon="✏️")
                                            st.rerun()
                                        else:
                                            st.error(msg)
                        else:
                            st.button("✏️", disabled=True, use_container_width=True, key=f"t1_no_ren_{rep_name}")

                    with col_del:
                        if can_edit:
                            with st.popover("🗑️ Delete", use_container_width=True, key=f"t1_del_{rep_name}_{st.session_state.popover_key_delete}"):
                                st.markdown("⚠️ **Confirm Deletion**")
                                st.write(f"Are you sure you want to delete `{rep_name}`?")
                                if st.button("Yes, Delete", type="primary", use_container_width=True, key=f"t1_btn_del_{rep_name}"):
                                    ok, msg = delete_saved_report(rep_name, rep_path)
                                    if ok:
                                        st.session_state.popover_key_delete += 1
                                        st.toast(f"🗑️ {msg}", icon="🗑️")
                                        st.rerun()
                                    else:
                                        st.error(msg)
                        else:
                            st.button("🗑️", disabled=True, use_container_width=True, key=f"t1_no_del_{rep_name}")
    else:
        st.info("🔒 Fleet Review requires Admin or Planner access. Contact an admin to have your role assigned.")

with tab2:
    st.divider()
    expander_open = True if st.session_state.role == "production" else False

    with st.expander("⚙️ **Production Workspace — Reports Sent for Execution**", expanded=expander_open):
        all_saved_df = get_saved_reports_list()

        # Production only sees reports explicitly published by the Planner/Admin.
        # Admin and Planner retain full visibility of all published reports here too.
        if not all_saved_df.empty and "published_to_production" in all_saved_df.columns:
            saved_df = all_saved_df[
                all_saved_df["published_to_production"] == True
            ].reset_index(drop=True)
        else:
            saved_df = all_saved_df

        if saved_df.empty:
            st.info(
                "⏳ No reports have been sent to Production yet. "
                "Ask your Planner to open the **Fleet Review** tab and click "
                "**🚀 Send to Production** on a saved report."
            )
        else:
            # Display Overview Table of Saved Reports including Total Hours, B1, B2, and Tasks
            st.dataframe(
                saved_df[[
                    "filename", "report_type", "total_tasks_display", "total_mh_display", 
                    "b1_mh_display", "b2_mh_display", "status_label", "created_by", 
                    "created_at", "last_completed_date", "last_completed_by"
                ]].rename(columns={
                    "filename": "Report Name",
                    "report_type": "Report Type",
                    "total_tasks_display": "Tasks (Done/Total)",
                    "total_mh_display": "Total Hours (Rem/Total)",
                    "b1_mh_display": "B1 Hours (Rem/Total)",
                    "b2_mh_display": "B2 Hours (Rem/Total)",
                    "created_by": "Created By (Planner)",
                    "created_at": "Saved At",
                    "status_label": "Compliance Status",
                    "last_completed_date": "Final Completion Date",
                    "last_completed_by": "Completed By (Last Task)"
                }),
                use_container_width=True,
                hide_index=True
            )

            st.markdown("---")
        
            selected_file_name = st.selectbox(
                "Select a saved file to track task compliance & preview:",
                options=saved_df["filename"].tolist(),
                key="saved_reports_select_box"
            )

            if selected_file_name:
                row_data = saved_df[saved_df["filename"] == selected_file_name].iloc[0]
                creator = str(row_data.get("created_by", ""))
                fpath = str(row_data.get("file_path", ""))
            
                is_admin = st.session_state.role == "admin"
                is_production = st.session_state.role == "production"
                is_planner = st.session_state.role == "planner"
                can_edit_compliance = is_production or is_admin

                st.markdown(f"### 📋 Task Compliance Checklist: <span style='color: #004B87;'>{selected_file_name}</span>", unsafe_allow_html=True)

                if is_planner:
                    st.caption("🔒 **Planner Mode:** Viewing task compliance status in read-only mode.")
                elif is_production:
                    st.caption("⚙️ **Production Mode:** Check completed tasks, adjust completion dates, and click Save.")
                elif is_admin:
                    st.caption("🔑 **Admin Mode:** Full edit and save access.")
                else:
                    st.caption("👁️ **View-Only Mode:** Your account is awaiting role assignment by an admin.")

                if os.path.exists(fpath):
                    try:
                        excel_df = read_excel_fresh(fpath)

                        # Read stored task compliance from JSON sidecar
                        json_path = fpath + ".json"
                        tasks_meta = {}
                        if os.path.exists(json_path):
                            try:
                                with open(json_path, "r") as mf:
                                    meta_json = json.load(mf)
                                    tasks_meta = meta_json.get("tasks_compliance", {})
                            except Exception as e:
                                logger.error("Compliance sidecar %s is unreadable, showing no saved progress: %s", json_path, e)
                                st.warning("⚠️ Could not read saved compliance history for this report - it may show as unchecked below.")

                        done_list = []
                        date_list = []
                        user_list = []
                        saved_done_list = []
                        saved_date_list = []

                        for idx, r in excel_df.iterrows():
                            task_id = str(r.get("TASK Number", r.get("Task", f"Task_{idx+1}"))).strip()
                            task_key = f"{idx}_{task_id}"
                            t_info = tasks_meta.get(task_key, {})

                            is_done = bool(t_info.get("done", False))
                            done_list.append(is_done)
                            saved_done_list.append(is_done)
                        
                            raw_d = t_info.get("done_date", "")
                            if raw_d:
                                try:
                                    parsed_d = datetime.strptime(raw_d, "%Y-%m-%d").date()
                                except ValueError:
                                    parsed_d = datetime.now().date()
                            else:
                                parsed_d = datetime.now().date()
                        
                            date_list.append(parsed_d)
                            saved_date_list.append(parsed_d)

                            user_list.append(t_info.get("checked_by", ""))

                        # State key tracking
                        state_key = f"done_list_{selected_file_name}"
                        date_state_key = f"date_list_{selected_file_name}"
                        editor_key = f"editor_comp_{selected_file_name}"
                        select_all_key = f"select_all_toggle_{selected_file_name}"
                        bulk_date_key = f"bulk_date_input_{selected_file_name}"

                        if state_key not in st.session_state or len(st.session_state[state_key]) != len(excel_df):
                            st.session_state[state_key] = list(done_list)

                        if date_state_key not in st.session_state or len(st.session_state[date_state_key]) != len(excel_df):
                            st.session_state[date_state_key] = list(date_list)

                        if not is_admin:
                            for i, was_saved in enumerate(saved_done_list):
                                if was_saved:
                                    st.session_state[state_key][i] = True
                                    st.session_state[date_state_key][i] = saved_date_list[i]

                        # Callback: Bulk date handler
                        def on_apply_bulk_date():
                            chosen_date = st.session_state.get(bulk_date_key, datetime.now().date())
                            updated_dates = list(st.session_state[date_state_key])
                            current_done = st.session_state[state_key]
                        
                            for i, was_saved in enumerate(saved_done_list):
                                if current_done[i] and not was_saved:
                                    updated_dates[i] = chosen_date
                                
                            st.session_state[date_state_key] = updated_dates
                            if editor_key in st.session_state:
                                del st.session_state[editor_key]

                        # Callback: Select/Deselect All handler
                        def on_select_all_toggle():
                            val = st.session_state[select_all_key]
                            chosen_date = st.session_state.get(bulk_date_key, datetime.now().date())
                            updated_done = list(st.session_state[state_key])
                            updated_dates = list(st.session_state[date_state_key])

                            for i, was_saved in enumerate(saved_done_list):
                                if is_admin:
                                    if not was_saved:
                                        updated_done[i] = val
                                        if val:
                                            updated_dates[i] = chosen_date
                                    else:
                                        updated_done[i] = val
                                else:
                                    if not was_saved:
                                        updated_done[i] = val
                                        if val:
                                            updated_dates[i] = chosen_date
                                    else:
                                        updated_done[i] = True
                                        updated_dates[i] = saved_date_list[i]

                            st.session_state[state_key] = updated_done
                            st.session_state[date_state_key] = updated_dates
                            if editor_key in st.session_state:
                                del st.session_state[editor_key]

                        # Controls bar
                        col_sel_all, col_bulk_date, col_spacer = st.columns([2.5, 2.5, 1.0])
                    
                        with col_sel_all:
                            initial_all_done = all(st.session_state[state_key]) if st.session_state[state_key] else False
                            st.checkbox(
                                "☑️ **Select / Deselect All Tasks**",
                                value=initial_all_done,
                                disabled=not can_edit_compliance,
                                key=select_all_key,
                                on_change=on_select_all_toggle
                            )

                        with col_bulk_date:
                            default_date = st.session_state[date_state_key][0] if st.session_state[date_state_key] else datetime.now().date()
                            st.date_input(
                                "📅 **Bulk Set Completion Date**",
                                value=default_date,
                                disabled=not can_edit_compliance,
                                format="DD-MM-YYYY",
                                key=bulk_date_key,
                                on_change=on_apply_bulk_date
                            )

                        excel_df.insert(0, "Done", st.session_state[state_key])
                        excel_df["Completion Date"] = st.session_state[date_state_key]
                        excel_df["Completed By"] = user_list

                        if not is_admin and any(saved_done_list):
                            st.caption("🔒 **Lock Rule:** Tasks completed in previous saves are locked and cannot be edited or unchecked.")

                        # Calculate live progress and hours dynamically from current checklist state
                        total_c = len(excel_df)
                        done_c = int(sum(st.session_state[state_key]))
                        pct = (done_c / total_c) * 100 if total_c > 0 else 0

                        has_mh_col = "Man Hours" in excel_df.columns
                        total_mh_active = 0.0
                        completed_mh_active = 0.0
                        b1_total_active = 0.0
                        b1_completed_active = 0.0
                        b2_total_active = 0.0
                        b2_completed_active = 0.0

                        for idx, r in excel_df.iterrows():
                            task_done = bool(st.session_state[state_key][idx])
                            mh_val = parse_mh_to_float(r.get("Man Hours")) if has_mh_col else 0.0

                            if "Trade Skill" in excel_df.columns:
                                trade_str = str(r.get("Trade Skill", ""))
                            elif "Skill Code" in excel_df.columns:
                                trade_str = map_skill_code_to_trade(r.get("Skill Code", ""))
                            else:
                                trade_str = "B1 (Mechanical)"

                            total_mh_active += mh_val
                            if task_done:
                                completed_mh_active += mh_val

                            if "B2" in trade_str.upper():
                                b2_total_active += mh_val
                                if task_done:
                                    b2_completed_active += mh_val
                            else:
                                b1_total_active += mh_val
                                if task_done:
                                    b1_completed_active += mh_val

                        rem_mh_active = max(0.0, total_mh_active - completed_mh_active)
                        b1_rem_active = max(0.0, b1_total_active - b1_completed_active)
                        b2_rem_active = max(0.0, b2_total_active - b2_completed_active)

                        # Row 1: Task Counts
                        col_m1, col_m2, col_m3, col_m4 = st.columns(4)
                        with col_m1:
                            st.metric("📊 Total Tasks", f"{total_c}")
                        with col_m2:
                            st.metric("✅ Completed Tasks", f"{done_c}")
                        with col_m3:
                            st.metric("⏳ Remaining Tasks", f"{total_c - done_c}")
                        with col_m4:
                            st.metric("📈 Task Progress", f"{pct:.1f}%")

                        # Row 2: Man Hours Metrics (if available)
                        if total_mh_active > 0:
                            col_h1, col_h2, col_h3, col_h4 = st.columns(4)
                            with col_h1:
                                st.metric("⏱️ Total Package MH", f"{rem_mh_active:.1f}h Rem", delta=f"{total_mh_active:.1f}h Total")
                            with col_h2:
                                st.metric("🔧 B1 Mechanical MH", f"{b1_rem_active:.1f}h Rem", delta=f"{b1_total_active:.1f}h Total")
                            with col_h3:
                                st.metric("⚡ B2 Avionics MH", f"{b2_rem_active:.1f}h Rem", delta=f"{b2_total_active:.1f}h Total")
                            with col_h4:
                                st.metric("📊 MH Completed", f"{((completed_mh_active / total_mh_active) * 100):.1f}%")

                        st.progress(pct / 100.0)

                        # Task Data Editor
                        edited_task_df = st.data_editor(
                            excel_df,
                            use_container_width=True,
                            hide_index=True,
                            key=editor_key,
                            column_config={
                                "Done": st.column_config.CheckboxColumn(
                                    "Done",
                                    help="Check when task execution is completed",
                                    default=False
                                ),
                                "Completion Date": st.column_config.DateColumn(
                                    "Completion Date",
                                    format="DD-MM-YYYY"
                                ),
                                "Completed By": st.column_config.TextColumn(
                                    "Completed By (Auto)",
                                    help="User email who completed the task",
                                    disabled=True
                                )
                            },
                            disabled=[c for c in excel_df.columns if c not in ["Done", "Completion Date"]] if can_edit_compliance else True
                        )

                        # Auto-assign completion date on new checks
                        active_bulk_date = st.session_state.get(bulk_date_key, datetime.now().date())
                        old_done = st.session_state[state_key]
                        new_done = edited_task_df["Done"].tolist()

                        for i, was_saved in enumerate(saved_done_list):
                            if not is_admin and was_saved:
                                edited_task_df.at[i, "Done"] = True
                                edited_task_df.at[i, "Completion Date"] = saved_date_list[i]
                            else:
                                if not was_saved and not old_done[i] and new_done[i]:
                                    edited_task_df.at[i, "Completion Date"] = active_bulk_date

                        # Sync back manual edits
                        st.session_state[state_key] = edited_task_df["Done"].tolist()
                        st.session_state[date_state_key] = edited_task_df["Completion Date"].tolist()

                        # Save Progress Button
                        if can_edit_compliance:
                            if st.button("💾 Save Task Compliance Progress", type="primary", key=f"btn_save_tasks_{selected_file_name}", use_container_width=True):
                                active_user = st.session_state.get("user_email", "Production User")
                                ok, msg = save_task_compliance(fpath, edited_task_df, active_user)
                                if ok:
                                    if state_key in st.session_state:
                                        del st.session_state[state_key]
                                    if date_state_key in st.session_state:
                                        del st.session_state[date_state_key]
                                    st.toast(f"✅ {msg}", icon="✅")
                                    st.rerun()
                                else:
                                    st.error(msg)

                        st.markdown("---")

                        # Actions Bar: Download, Rename, Delete
                        col_act1, col_act2, col_act3 = st.columns(3)

                        # 1. DOWNLOAD
                        with col_act1:
                            with open(fpath, "rb") as file_bytes:
                                download_clicked = st.download_button(
                                    label=f"⬇️ Download Excel ({total_c} tasks)",
                                    data=file_bytes,
                                    file_name=selected_file_name,
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    type="primary",
                                    use_container_width=True,
                                    key=f"dl_btn_{selected_file_name}"
                                )
                                if download_clicked:
                                    current_user = st.session_state.get("user_email", "System User")
                                    log_download_activity(selected_file_name, current_user)
                                    st.toast(f"📥 Download logged for {selected_file_name}", icon="ℹ️")

                        # 2. RENAME
                        with col_act2:
                            is_creator = st.session_state.user_email and (creator.lower() == st.session_state.user_email.lower())
                            can_rename = is_admin or (is_planner and is_creator)
                            if can_rename:
                                with st.popover("✏️ Rename Report", use_container_width=True, key=f"popover_rename_{st.session_state.popover_key_rename}"):
                                    st.markdown("### Rename Report")
                                    new_name_in = st.text_input("New Name:", value=selected_file_name, key=f"rename_input_{selected_file_name}")
                                    if st.button("Confirm Rename", type="primary", use_container_width=True, key=f"btn_rename_{selected_file_name}"):
                                        if new_name_in.strip() and new_name_in != selected_file_name:
                                            ok, msg = rename_saved_report(selected_file_name, new_name_in)
                                            if ok:
                                                st.session_state.popover_key_rename += 1
                                                st.toast(f"✅ {msg}", icon="✏️")
                                                st.rerun()
                                            else:
                                                st.error(msg)
                            else:
                                st.button("✏️ Rename (Owner/Admin Only)", disabled=True, use_container_width=True)

                        # 3. DELETE
                        with col_act3:
                            is_creator = st.session_state.user_email and (creator.lower() == st.session_state.user_email.lower())
                            can_delete = is_admin or (is_planner and is_creator)
                        
                            if can_delete:
                                with st.popover("🗑️ Delete Report", use_container_width=True, key=f"popover_del_{st.session_state.popover_key_delete}"):
                                    st.markdown("⚠️ **Confirm Deletion**")
                                    st.write(f"Are you sure you want to delete `{selected_file_name}`?")
                                    if st.button("Yes, Delete File", type="primary", use_container_width=True, key=f"btn_del_{selected_file_name}"):
                                        ok, msg = delete_saved_report(selected_file_name, fpath)
                                        if ok:
                                            st.session_state.popover_key_delete += 1
                                            st.toast(f"🗑️ {msg}", icon="🗑️")
                                            st.rerun()
                                        else:
                                            st.error(msg)
                            else:
                                st.button("🗑️ Delete (Owner/Admin Only)", disabled=True, use_container_width=True)

                    except Exception as e:
                        st.error(f"Could not load tasks for report: {e}")
                else:
                    st.error("The selected file path could not be found on the server.")

with tab3:
    if st.session_state.role in ["admin", "planner"]:
        with st.expander("🔔 **Notifications: Report Downloads Log**", expanded=False):
            download_df = get_download_logs()
        
            if download_df.empty:
                st.info("No report downloads recorded yet.")
            else:
                st.caption("Real-time log of users downloading saved reports:")
            
                latest = download_df.iloc[0]
                col_n1, col_n2, col_n3 = st.columns(3)
            
                with col_n1:
                    st.markdown(
                        f"""
                        <div style="background-color: #f8fafc; padding: 10px 14px; border-radius: 8px; border-left: 4px solid #0284c7; border-top: 1px solid #e2e8f0; border-right: 1px solid #e2e8f0; border-bottom: 1px solid #e2e8f0;">
                            <span style="font-size: 0.75rem; color: #64748b; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px;">📄 Last Downloaded File</span><br>
                            <span style="font-size: 0.9rem; font-weight: 600; color: #0f172a; word-break: break-all;">{latest['filename']}</span>
                        </div>
                        """,
                        unsafe_allow_html=True
                    )
                with col_n2:
                    st.markdown(
                        f"""
                        <div style="background-color: #f8fafc; padding: 10px 14px; border-radius: 8px; border-left: 4px solid #0284c7; border-top: 1px solid #e2e8f0; border-right: 1px solid #e2e8f0; border-bottom: 1px solid #e2e8f0;">
                            <span style="font-size: 0.75rem; color: #64748b; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px;">👤 Downloaded By</span><br>
                            <span style="font-size: 0.9rem; font-weight: 600; color: #0f172a; word-break: break-all;">{latest['downloaded_by']}</span>
                        </div>
                        """,
                        unsafe_allow_html=True
                    )
                with col_n3:
                    st.markdown(
                        f"""
                        <div style="background-color: #f8fafc; padding: 10px 14px; border-radius: 8px; border-left: 4px solid #0284c7; border-top: 1px solid #e2e8f0; border-right: 1px solid #e2e8f0; border-bottom: 1px solid #e2e8f0;">
                            <span style="font-size: 0.75rem; color: #64748b; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px;">🕒 Date & Time</span><br>
                            <span style="font-size: 0.9rem; font-weight: 600; color: #0f172a;">{latest['timestamp']}</span>
                        </div>
                        """,
                        unsafe_allow_html=True
                    )

                st.markdown("---")
            
                st.dataframe(
                    download_df.rename(columns={
                        "filename": "Report Name",
                        "downloaded_by": "User Email",
                        "timestamp": "Date & Time (UTC)"
                    }),
                    use_container_width=True,
                    hide_index=True
                )
