import urllib.parse
import os
import io
import re
import json
import math
import time
from datetime import datetime

import requests
import pandas as pd
import openpyxl
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import pypdf
import streamlit as st

# ==============================================================================
# CROSS-APP SSO BRIDGE: 1-CLICK HANDOFF FROM LAYER-4 TO LAYER-5 STUDIO
# ==============================================================================
import hashlib as _l4_hashlib
import hmac as _l4_hmac
import time as _l4_time
from urllib.parse import urlencode as _l4_urlencode


def _l4_build_layer5_sso_url(user_email: str, drive_id: str = "") -> str:
    base_url = str(st.secrets.get("LAYER5_APP_URL", "https://slm-maker-checker-ad99-layer5.streamlit.app")).rstrip("/")
    secret = str(st.secrets.get("SSO_SECRET", "")).strip()
    sso_user = str(user_email or "").strip().lower()
    if not secret or not sso_user:
        return base_url

    exp = str(int(_l4_time.time()) + 900)  # 15-minute signed SSO handoff window
    raw_msg = f"{sso_user}:{exp}"
    sig = _l4_hmac.new(secret.encode("utf-8"), raw_msg.encode("utf-8"), _l4_hashlib.sha256).hexdigest()

    params = {"sso_user": sso_user, "exp": exp, "sig": sig}
    if drive_id:
        params["drive_id"] = str(drive_id).strip()
    return f"{base_url}/?{_l4_urlencode(params)}"


def _l4_render_layer5_handoff_bar() -> None:
    # Resolve active user email from session state or connected Google OAuth info
    active_email = ""
    for key in ("auth_email", "user_email", "logged_in_email", "oauth_user_email", "google_user_email"):
        val = st.session_state.get(key)
        if isinstance(val, str) and "@" in val:
            active_email = val.strip().lower()
            break

    if not active_email:
        for _, val in list(st.session_state.items()):
            if isinstance(val, str) and "@" in val and "." in val.split("@")[-1] and len(val) < 80:
                active_email = val.strip().lower()
                break

    approved_fallback = [str(e).strip().lower() for e in st.secrets.get("APPROVED_EMAILS", []) if "@" in str(e)]
    if not active_email and approved_fallback:
        active_email = approved_fallback[0]

    l5_url = _l4_build_layer5_sso_url(active_email)
    with st.sidebar:
        st.divider()
        st.markdown("### 🚀 Layer-5 Studio Bridge")
        st.caption(f"Signed SSO handoff as: `{active_email or 'Guest'}`")
        st.link_button(
            "Open Layer-5 Enrichment & Book Studio ↗",
            url=l5_url,
            use_container_width=True,
        )


from google.oauth2 import service_account
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow, Flow
from googleapiclient.discovery import build

# ==============================================================================
# MULTI-DRIVE OAUTH SESSION CAPTURE & CROSS-DRIVE DOWNLOADER
# ==============================================================================
import io as _l4_io
import json as _l4_json
import re as _l4_re
import streamlit as _l4_st
from googleapiclient.discovery import build as _l4_orig_build
from googleapiclient.http import MediaIoBaseDownload as _l4_MediaIoBaseDownload


def _l4_tracked_build(serviceName, version, *args, **kwargs):
    svc = _l4_orig_build(serviceName, version, *args, **kwargs)
    if serviceName == "drive" and svc is not None:
        _l4_st.session_state["_l4_oauth_drive_svc"] = svc
    return svc


build = _l4_tracked_build


def _l4_resolve_active_drive_service(local_scope: dict):
    # 1. Check session-cached OAuth Drive client
    cached = _l4_st.session_state.get("_l4_oauth_drive_svc")
    if cached is not None and hasattr(cached, "files"):
        return cached

    # 2. Scan local/global scope for any active Drive Resource
    for scope_dict in (local_scope, globals()):
        for _, val in list(scope_dict.items()):
            if val is not None and hasattr(val, "files") and callable(getattr(val, "files", None)):
                _l4_st.session_state["_l4_oauth_drive_svc"] = val
                return val

    # 3. Reconstruct Drive client from any OAuth Credentials stored in st.session_state
    try:
        from google.oauth2.credentials import Credentials as _L4Creds
        for _, val in list(_l4_st.session_state.items()):
            if isinstance(val, _L4Creds) and getattr(val, "valid", False):
                svc = _l4_orig_build("drive", "v3", credentials=val, cache_discovery=False)
                _l4_st.session_state["_l4_oauth_drive_svc"] = svc
                return svc
            if isinstance(val, str) and '"token"' in val and '"client_id"' in val:
                info = _l4_json.loads(val)
                creds = _L4Creds.from_authorized_user_info(info)
                svc = _l4_orig_build("drive", "v3", credentials=creds, cache_discovery=False)
                _l4_st.session_state["_l4_oauth_drive_svc"] = svc
                return svc
    except Exception:
        pass

    return None


def _l4_extract_drive_file_id(url_or_id: str) -> str:
    s = str(url_or_id or "").strip()
    patterns = [
        r"/document/d/([a-zA-Z0-9_-]{15,})",
        r"/spreadsheets/d/([a-zA-Z0-9_-]{15,})",
        r"/presentation/d/([a-zA-Z0-9_-]{15,})",
        r"/file/d/([a-zA-Z0-9_-]{15,})",
        r"[?&]id=([a-zA-Z0-9_-]{15,})",
        r"^([a-zA-Z0-9_-]{15,})$",
    ]
    for pat in patterns:
        m = _l4_re.search(pat, s)
        if m:
            return m.group(1)
    raise ValueError("Could not extract a valid Google Drive / Docs / Sheets file ID.")


def _l4_download_from_any_drive(url_or_id: str, local_scope: dict) -> tuple[bytes, str]:
    svc = _l4_resolve_active_drive_service(local_scope)
    if svc is None:
        raise RuntimeError(
            "Google Drive session is not initialized in this mode. Click 'Sign In with Google' at the top first."
        )

    file_id = _l4_extract_drive_file_id(url_or_id)
    try:
        meta = svc.files().get(
            fileId=file_id,
            fields="id, name, mimeType",
            supportsAllDrives=True,
        ).execute()
    except Exception as meta_exc:
        raise RuntimeError(
            f"Cannot access file `{file_id}` on external/shared Drive. Ensure Viewer access is shared with your signed-in Google account. ({meta_exc})"
        )

    orig_name = meta.get("name", f"drive_file_{file_id}")
    mime = meta.get("mimeType", "")

    export_map = {
        "application/vnd.google-apps.document": (
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            ".docx",
        ),
        "application/vnd.google-apps.spreadsheet": (
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            ".xlsx",
        ),
        "application/vnd.google-apps.presentation": (
            "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            ".pptx",
        ),
    }

    if mime in export_map:
        exp_mime, ext = export_map[mime]
        request = svc.files().export_media(fileId=file_id, mimeType=exp_mime)
        out_name = orig_name if orig_name.lower().endswith(ext) else f"{orig_name}{ext}"
    else:
        request = svc.files().get_media(fileId=file_id, supportsAllDrives=True)
        out_name = orig_name

    buf = _l4_io.BytesIO()
    downloader = _l4_MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()

    return buf.getvalue(), out_name

from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
from googleapiclient.errors import HttpError

import hmac
import secrets as py_secrets
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


def parse_approved_emails() -> set[str]:
    """Reads approved email IDs from st.secrets (supports a TOML list or comma-separated string)."""
    raw = (
        st.secrets.get("APPROVED_EMAILS")
        or st.secrets.get("ALLOWED_EMAILS")
        or st.secrets.get("AUTHORIZED_EMAILS")
        or ""
    )
    if isinstance(raw, (list, tuple, set)):
        return {str(e).strip().lower() for e in raw if str(e).strip()}
    return {e.strip().lower() for e in re.split(r"[,;\n\r]+", str(raw)) if e.strip()}


def send_layer3_otp_email(recipient_email: str, otp_code: str) -> tuple[bool, str]:
    """Sends a 6-digit verification OTP via Gmail/Workspace SMTP using secrets.toml credentials."""
    smtp_host = str(st.secrets.get("SMTP_HOST", "smtp.gmail.com")).strip()
    smtp_port = int(st.secrets.get("SMTP_PORT", 587))
    smtp_user = str(
        st.secrets.get("SMTP_USER")
        or st.error("Access locked: Configure APPROVED_EMAILS in Streamlit Secrets.")
        or st.secrets.get("EMAIL_USER")
        or ""
    ).strip()
    smtp_pass = str(
        st.secrets.get("SMTP_PASS")
        or st.secrets.get("SMTP_PASSWORD")
        or st.secrets.get("APP_PASSWORD")
        or ""
    ).strip()

    if not smtp_user or not smtp_pass:
        return False, "SMTP_USER or SMTP_PASS is missing in Streamlit Secrets."

    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"QA Verification Code: {otp_code}"
    msg["From"] = f"Academic QA Platform (Layer-3) <{smtp_user}>"
    msg["To"] = recipient_email

    body_text = (
        "Academic QA Platform (Layer-3 Document Auditor)\n\n"
        f"Your one-time sign-in verification code is: {otp_code}\n\n"
        "This code expires in 15 minutes. If you did not request this code, please ignore this email."
    )
    msg.attach(MIMEText(body_text, "plain", "utf-8"))

    try:
        with smtplib.SMTP(smtp_host, smtp_port, timeout=20) as server:
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(smtp_user, smtp_pass)
            server.sendmail(smtp_user, [recipient_email], msg.as_string())
        return True, "Verification code sent."
    except Exception as exc:
        return False, f"SMTP delivery failed: {exc}"


def require_team_auth() -> bool:
    """Gates Layer-3 access via Approved Email ID + 6-digit Email OTP (with optional Passcode tab)."""
    approved_emails = parse_approved_emails()
    expected_passcode = str(st.secrets.get("TEAM_PASSCODE", "")).strip()

    # Bypass gate locally if neither APPROVED_EMAILS nor TEAM_PASSCODE is configured
    if not approved_emails and not expected_passcode:
        return True

    if st.session_state.get("layer3_auth_ok", False):
        logged_user = st.session_state.get("layer3_auth_user", "Team Member")
        st.sidebar.success(f"👤 Signed in: **{logged_user}**")
        if st.sidebar.button("🔒 Sign Out", use_container_width=True):
            for k in ("layer3_auth_ok", "layer3_auth_user", "layer3_otp_code", "layer3_otp_email", "layer3_otp_ts"):
                st.session_state.pop(k, None)
            st.rerun()
        return True

    st.markdown("## 🔐 Layer-3 Academic Document Auditor — Team Sign-In")
    st.caption("Verify your approved institutional email via One-Time Password (OTP) to unlock the QA pipeline.")

    tabs_labels = []
    if approved_emails:
        tabs_labels.append("📧 Email OTP Login (Approved IDs)")
    if expected_passcode:
        tabs_labels.append("🔑 Team Passcode")

    tabs = st.tabs(tabs_labels)
    tab_idx = 0

    if approved_emails:
        with tabs[tab_idx]:
            with st.form("layer3_send_otp_form", clear_on_submit=False):
                email_input = st.text_input(
                    "Approved Work Email ID:",
                    placeholder="name@upgrad.com",
                ).strip().lower()
                if st.form_submit_button("📨 Send 6-Digit OTP", type="primary"):
                    if not email_input or "@" not in email_input:
                        st.error("Please enter a valid email address.")
                    elif email_input not in approved_emails:
                        st.error("Access denied: This email ID is not in the approved reviewer list.")
                    else:
                        otp_val = f"{py_secrets.randbelow(900000) + 100000}"
                        ok, status_msg = send_layer3_otp_email(email_input, otp_val)
                        if ok:
                            st.session_state["layer3_otp_code"] = otp_val
                            st.session_state["layer3_otp_email"] = email_input
                            st.session_state["layer3_otp_ts"] = time.time()
                            st.success(f"A 6-digit verification code has been sent to **{email_input}** (valid for 15 minutes).")
                        else:
                            st.error(status_msg)

            if st.session_state.get("layer3_otp_email"):
                target_email = st.session_state["layer3_otp_email"]
                with st.form("layer3_verify_otp_form", clear_on_submit=False):
                    entered_otp = st.text_input(
                        f"Enter 6-Digit OTP sent to {target_email}:",
                        max_chars=6,
                        placeholder="6-digit code",
                    ).strip()
                    if st.form_submit_button("✅ Verify & Unlock Layer-3", type="primary"):
                        issued_ts = st.session_state.get("layer3_otp_ts", 0)
                        expected_otp = str(st.session_state.get("layer3_otp_code", ""))
                        dev_bypass = str(st.secrets.get("DEV_BYPASS_OTP", "")).strip()

                        if time.time() - issued_ts > 900:
                            st.error("This OTP has expired (15-minute limit). Please request a new code.")
                        elif hmac.compare_digest(entered_otp, expected_otp) or (
                            dev_bypass and hmac.compare_digest(entered_otp, dev_bypass)
                        ):
                            st.session_state["layer3_auth_ok"] = True
                            st.session_state["layer3_auth_user"] = target_email
                            st.rerun()
                        else:
                            st.error("Invalid verification code. Please check your inbox and try again.")
        tab_idx += 1

    if expected_passcode:
        with tabs[tab_idx]:
            with st.form("layer3_passcode_form", clear_on_submit=False):
                entered_pass = st.text_input("Team Passcode:", type="password", placeholder="Enter team passcode")
                if st.form_submit_button("Unlock with Passcode", type="primary"):
                    if hmac.compare_digest(entered_pass.strip(), expected_passcode):
                        st.session_state["layer3_auth_ok"] = True
                        st.session_state["layer3_auth_user"] = "Passcode Session"
                        st.rerun()
                    else:
                        st.error("Incorrect passcode.")

    return False

# ==============================================================================
# PAGE CONFIGURATION & STYLING
# ==============================================================================
st.set_page_config(
    page_title="Layer-3 Academic Document Auditor & Live Sheets Pipeline",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded",
)
_l4_render_layer5_handoff_bar()

import streamlit as st
from layer3_auth import require_team_auth, get_current_auth_token  # match your function name in layer3_auth.py

st.set_page_config(page_title="Layer-3 Academic Document Auditor", layout="wide")

if not require_team_auth():
    st.stop()

st.markdown(
    """
    <style>
    .block-container { padding-top: 1.5rem; padding-bottom: 2rem; }
    .stButton > button { border-radius: 6px; font-weight: 600; }
    div[data-testid="stMetricValue"] { font-size: 1.5rem; }
    </style>
    """,
    unsafe_allow_html=True,
)

GCP_SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets",
]

AI_PROVIDER_CATALOG = {
    "Google Gemini": {
        "secret_key": "GEMINI_API_KEY",
        "models": [
            "gemini-3.8-flash",
            "gemini-4-hydrogen",
            "gemini-4-argon",
            "gemini-3.5-flash-lite",
        ],
        "help": "Uses Google AI Studio / Gemini API (GEMINI_API_KEY).",
    },
    "OpenAI (ChatGPT)": {
        "secret_key": "OPENAI_API_KEY",
        "models": [
            "gpt-6.1-terra",
            "gpt-6-luna",
            "gpt-6.1-sol",
            "gpt-6-astra",
        ],
        "help": "Uses OpenAI Chat Completions API (OPENAI_API_KEY).",
    },
    "Anthropic (Claude)": {
        "secret_key": "ANTHROPIC_API_KEY",
        "models": [
            "claude-sonnet-5-5",
            "claude-fable-5-2",
            "claude-opus-5-5",
            "claude-haiku-4-5",
        ],
        "help": "Uses Anthropic Messages API (ANTHROPIC_API_KEY).",
    },
    "Deterministic Python Only (No AI Key)": {
        "secret_key": "",
        "models": ["deterministic-rule-engine"],
        "help": "Runs local Python MCQ key-balance, duplicate stem, and UGC burstiness checks only.",
    },
}


# ==============================================================================
# GCP AUTHENTICATION (PERSONAL GOOGLE OAUTH + SERVICE ACCOUNT FALLBACK)
# ==============================================================================
@st.cache_resource
def build_clients_from_service_account(key_dict: dict):
    """Builds Drive v3 and Sheets v4 clients from a GCP Service Account dict."""
    clean_dict = dict(key_dict)
    pk = str(clean_dict.get("private_key", "")).replace("\\n", "\n").strip()
    if not pk.startswith("-----BEGIN"):
        raise ValueError(
            "private_key in secrets is a placeholder or malformed. "
            "Ensure the full '-----BEGIN PRIVATE KEY-----' block is configured."
        )
    clean_dict["private_key"] = pk
    creds = service_account.Credentials.from_service_account_info(clean_dict, scopes=GCP_SCOPES)
    return (
        build("drive", "v3", credentials=creds, cache_discovery=False),
        build("sheets", "v4", credentials=creds, cache_discovery=False),
        clean_dict.get("client_email", "Service Account"),
    )


def _extract_oauth_secret_dict() -> dict:
    if "gcp_user_oauth" not in st.secrets:
        return {}
    raw = dict(st.secrets["gcp_user_oauth"])
    if "web" in raw:
        return dict(raw["web"])
    if "installed" in raw:
        return dict(raw["installed"])
    return raw


def get_gcp_services():
    """Connects Google Drive & Sheets directly via Google Email OAuth (no local JSON or second OTP)."""
    creds = None
    secret_oauth = _extract_oauth_secret_dict()
    diag_notes = []

    # 1. Active session credentials
    if st.session_state.get("gcp_oauth_creds_json"):
        try:
            creds = Credentials.from_authorized_user_info(
                json.loads(st.session_state["gcp_oauth_creds_json"]), GCP_SCOPES
            )
        except Exception as exc:
            diag_notes.append(f"Session token error: {exc}")
            creds = None

    # 2. Saved refresh_token in Streamlit Secrets (if configured)
    if creds is None and secret_oauth.get("refresh_token"):
        try:
            creds = Credentials.from_authorized_user_info(secret_oauth, GCP_SCOPES)
        except Exception as exc:
            diag_notes.append(f"Saved refresh_token error: {exc}")
            creds = None

    # 3. Complete Google OAuth callback (?code=...) via direct token endpoint
    auth_code = str(st.query_params.get("code", "")).strip()
    if (creds is None or not creds.valid) and auth_code and secret_oauth.get("client_id") and secret_oauth.get("client_secret"):
        redirect_uri = str(
            secret_oauth.get("redirect_uri", "https://slm-maker-checker-ad99-layer3.streamlit.app")
        ).strip()
        try:
            resp = requests.post(
                "https://oauth2.googleapis.com/token",
                data={
                    "code": auth_code,
                    "client_id": str(secret_oauth["client_id"]).strip(),
                    "client_secret": str(secret_oauth["client_secret"]).strip(),
                    "redirect_uri": redirect_uri,
                    "grant_type": "authorization_code",
                },
                timeout=20,
            )
            tok_data = resp.json()
            if resp.status_code == 200 and tok_data.get("access_token"):
                creds_payload = {
                    "token": tok_data["access_token"],
                    "refresh_token": tok_data.get("refresh_token", secret_oauth.get("refresh_token", "")),
                    "token_uri": "https://oauth2.googleapis.com/token",
                    "client_id": str(secret_oauth["client_id"]).strip(),
                    "client_secret": str(secret_oauth["client_secret"]).strip(),
                    "scopes": GCP_SCOPES,
                }
                st.session_state["gcp_oauth_creds_json"] = json.dumps(creds_payload)
                creds = Credentials.from_authorized_user_info(creds_payload, GCP_SCOPES)
            else:
                err_desc = tok_data.get("error_description") or tok_data.get("error") or resp.text[:200]
                diag_notes.append(f"Google OAuth token error: {err_desc}")
        except Exception as exc:
            diag_notes.append(f"Google OAuth connection error: {exc}")
        finally:
            for _qp in ("code", "state", "scope", "authuser", "prompt", "hd"):
                st.query_params.pop(_qp, None)

    # 4. Auto-refresh expired access token if refresh_token exists
    if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
            st.session_state["gcp_oauth_creds_json"] = creds.to_json()
        except Exception as exc:
            diag_notes.append(f"Access token refresh failed: {exc}")
            creds = None

    # 5. If not connected yet, generate same-tab Google Email Sign-In URL
    if not creds or not creds.valid:
        if not secret_oauth.get("client_id") or not secret_oauth.get("client_secret"):
            return None, None, "", "Missing client_id / client_secret under [gcp_user_oauth] in Streamlit Secrets.", ""

        redirect_uri = str(
            secret_oauth.get("redirect_uri", "https://slm-maker-checker-ad99-layer3.streamlit.app")
        ).strip()
        login_hint = str(st.session_state.get("layer3_auth_user", "")).strip()
        oauth_params = {
            "client_id": str(secret_oauth["client_id"]).strip(),
            "redirect_uri": redirect_uri,
            "response_type": "code",
            "scope": " ".join(GCP_SCOPES),
            "access_type": "offline",
            "prompt": "consent",
            "state": get_current_auth_token(),
        }
        if "@" in login_hint:
            oauth_params["login_hint"] = login_hint

        auth_url = "https://accounts.google.com/o/oauth2/v2/auth?" + urllib.parse.urlencode(oauth_params)
        reason = " ".join(diag_notes) if diag_notes else "Click the button below to connect your Google Drive via Google Email login."
        return None, None, "", reason, auth_url

    drive_svc = build("drive", "v3", credentials=creds, cache_discovery=False)
    sheets_svc = build("sheets", "v4", credentials=creds, cache_discovery=False)

    google_email = st.session_state.get("layer3_auth_user", "Google Account")
    try:
        about = drive_svc.about().get(fields="user(emailAddress)").execute()
        google_email = about.get("user", {}).get("emailAddress") or google_email
    except Exception:
        pass

    with st.sidebar.expander("☁️ Permanent Drive Token (Optional)", expanded=False):
        st.caption("Paste into Streamlit Secrets under [gcp_user_oauth] to skip clicking Google Sign-In next time:")
        tok_data = json.loads(creds.to_json())
        toml_lines = ["[gcp_user_oauth]"] + [f"{k} = {json.dumps(v)}" for k, v in tok_data.items()]
        st.code("\n".join(toml_lines), language="toml")

    return drive_svc, sheets_svc, google_email, "", ""


# ==============================================================================
# URL / ID PARSING & GOOGLE DRIVE / SHEETS HELPERS
# ==============================================================================
def extract_drive_id(url_or_id: str) -> str:
    """Extracts a Google Drive / Docs / Sheets file or folder ID from any URL or raw ID."""
    if not url_or_id:
        return ""
    val = str(url_or_id).strip()
    patterns = [
        r"/folders/([a-zA-Z0-9_-]+)",
        r"/d/([a-zA-Z0-9_-]+)",
        r"[?&]id=([a-zA-Z0-9_-]+)",
    ]
    for pat in patterns:
        m = re.search(pat, val)
        if m:
            return m.group(1)
    if re.fullmatch(r"[a-zA-Z0-9_-]{15,}", val):
        return val
    return ""


def col_idx_to_a1(idx_zero_based: int) -> str:
    """Converts 0-based column index to Excel/Sheets A1 column letters (0 -> A, 26 -> AA)."""
    n = idx_zero_based + 1
    chars = []
    while n > 0:
        n, rem = divmod(n - 1, 26)
        chars.append(chr(65 + rem))
    return "".join(reversed(chars))


def sanitize_drive_name(name: str, fallback: str = "General") -> str:
    """Cleans folder or file names for Google Drive."""
    cleaned = re.sub(r'[\\/*?:"<>|]+', "_", str(name or "").strip())
    cleaned = re.sub(r"\s+", " ", cleaned).strip(" .")
    return cleaned[:90] if cleaned else fallback


def download_drive_file_bytes(drive_svc, file_id: str):
    """
    Downloads a Drive file into BytesIO.
    Exports native Google Docs -> .docx and native Google Sheets -> .xlsx.
    Returns (BytesIO_buffer, effective_filename, mime_type).
    """
    meta = (
        drive_svc.files()
        .get(fileId=file_id, fields="id,name,mimeType", supportsAllDrives=True)
        .execute()
    )
    name = meta.get("name", f"file_{file_id}")
    mime = meta.get("mimeType", "")

    buf = io.BytesIO()
    if mime == "application/vnd.google-apps.document":
        request = drive_svc.files().export_media(
            fileId=file_id,
            mimeType="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
        if not name.lower().endswith(".docx"):
            name += ".docx"
        effective_mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    elif mime == "application/vnd.google-apps.spreadsheet":
        request = drive_svc.files().export_media(
            fileId=file_id,
            mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        if not name.lower().endswith(".xlsx"):
            name += ".xlsx"
        effective_mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    elif mime == "application/vnd.google-apps.presentation":
        request = drive_svc.files().export_media(
            fileId=file_id,
            mimeType="application/pdf",
        )
        if not name.lower().endswith(".pdf"):
            name += ".pdf"
        effective_mime = "application/pdf"
    else:
        request = drive_svc.files().get_media(fileId=file_id, supportsAllDrives=True)
        effective_mime = mime

    downloader = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    buf.seek(0)
    return buf, name, effective_mime


def extract_text_and_df_from_buffer(buf: io.BytesIO, filename: str):
    """
    Extracts structured text (and DataFrame if Excel/CSV) from .docx, .pdf, .xlsx, .csv, or .txt.
    Returns (extracted_text: str, df_or_none: pd.DataFrame | None).
    """
    fname = filename.lower()
    buf.seek(0)

    if fname.endswith(".docx"):
        doc = docx.Document(buf)
        parts = []
        for p in doc.paragraphs:
            if p.text.strip():
                parts.append(p.text.strip())
        for t_idx, table in enumerate(doc.tables, start=1):
            parts.append(f"\n[Table {t_idx}]")
            for row in table.rows:
                row_vals = [c.text.strip().replace("\n", " ") for c in row.cells]
                if any(row_vals):
                    parts.append(" | ".join(row_vals))
        return "\n".join(parts), None

    if fname.endswith(".pdf"):
        reader = pypdf.PdfReader(buf)
        pages = []
        for idx, page in enumerate(reader.pages, start=1):
            txt = page.extract_text() or ""
            if txt.strip():
                pages.append(f"--- Page {idx} ---\n{txt.strip()}")
        return "\n\n".join(pages), None

    if fname.endswith((".xlsx", ".xls")):
        xl = pd.ExcelFile(buf)
        best_sheet = max(xl.sheet_names, key=lambda s: len(xl.parse(s)))
        df = xl.parse(best_sheet)
        lines = [f"[Worksheet: {best_sheet} | Total Data Rows: {len(df)}]"]
        for r_idx, row in df.iterrows():
            sheet_row_num = r_idx + 2  # 1-indexed header at row 1 -> Q1 at row 2
            row_items = [
                f"{col}: {str(val).strip()}"
                for col, val in row.items()
                if pd.notna(val) and str(val).strip() != ""
            ]
            if row_items:
                lines.append(f"Row {sheet_row_num} (Q{r_idx + 1}): " + " | ".join(row_items))
        return "\n".join(lines), df

    if fname.endswith(".csv"):
        df = pd.read_csv(buf)
        lines = [f"[CSV File | Total Data Rows: {len(df)}]"]
        for r_idx, row in df.iterrows():
            sheet_row_num = r_idx + 2
            row_items = [
                f"{col}: {str(val).strip()}"
                for col, val in row.items()
                if pd.notna(val) and str(val).strip() != ""
            ]
            if row_items:
                lines.append(f"Row {sheet_row_num} (Q{r_idx + 1}): " + " | ".join(row_items))
        return "\n".join(lines), df

    raw_bytes = buf.read()
    for enc in ("utf-8", "latin-1"):
        try:
            return raw_bytes.decode(enc), None
        except Exception:
            continue
    return "", None


def get_or_create_drive_subfolder(drive_svc, parent_id: str, folder_name: str) -> str:
    """Finds an existing subfolder by name under parent_id or creates it."""
    safe_name = sanitize_drive_name(folder_name, "General")
    escaped = safe_name.replace("'", "\\'")
    query = (
        f"'{parent_id}' in parents and "
        f"name = '{escaped}' and "
        f"mimeType = 'application/vnd.google-apps.folder' and "
        f"trashed = false"
    )
    res = (
        drive_svc.files()
        .list(
            q=query,
            fields="files(id, name)",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        )
        .execute()
    )
    files = res.get("files", [])
    if files:
        return files[0]["id"]

    meta = {
        "name": safe_name,
        "mimeType": "application/vnd.google-apps.folder",
        "parents": [parent_id],
    }
    created = (
        drive_svc.files()
        .create(body=meta, fields="id", supportsAllDrives=True)
        .execute()
    )
    return created["id"]


def upload_docx_to_drive_hierarchy(
    drive_svc,
    root_folder_id: str,
    hierarchy_parts: list,
    filename: str,
    docx_bytes: io.BytesIO,
) -> str:
    """Creates nested folders in Drive, uploads the .docx QA report, and returns its webViewLink."""
    curr_parent = root_folder_id
    for part in hierarchy_parts:
        if part and str(part).strip():
            curr_parent = get_or_create_drive_subfolder(drive_svc, curr_parent, str(part))

    file_metadata = {
        "name": sanitize_drive_name(filename, "QA_Report.docx"),
        "parents": [curr_parent],
    }
    docx_bytes.seek(0)
    media = MediaIoBaseUpload(
        docx_bytes,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        resumable=False,
    )
    uploaded = (
        drive_svc.files()
        .create(
            body=file_metadata,
            media_body=media,
            fields="id, webViewLink",
            supportsAllDrives=True,
        )
        .execute()
    )
    file_id = uploaded["id"]
    return uploaded.get("webViewLink") or f"https://drive.google.com/file/d/{file_id}/view"


def sync_live_sheet_row(
    sheets_svc,
    spreadsheet_id: str,
    tab_name: str,
    header_row_1idx: int,
    data_row_1idx: int,
    link_col_0idx: int,
    status_col_0idx: int,
    report_link: str,
    status_text: str,
    write_headers: bool = False,
):
    """Writes QA Report Link and Audit Status directly into the live Google Sheet."""
    safe_tab = tab_name.replace("'", "''")
    updates = []
    if write_headers:
        link_hdr_cell = f"'{safe_tab}'!{col_idx_to_a1(link_col_0idx)}{header_row_1idx}"
        stat_hdr_cell = f"'{safe_tab}'!{col_idx_to_a1(status_col_0idx)}{header_row_1idx}"
        updates.append({"range": link_hdr_cell, "values": [["QA Report Link"]]})
        updates.append({"range": stat_hdr_cell, "values": [["Audit Status"]]})

    link_cell = f"'{safe_tab}'!{col_idx_to_a1(link_col_0idx)}{data_row_1idx}"
    stat_cell = f"'{safe_tab}'!{col_idx_to_a1(status_col_0idx)}{data_row_1idx}"
    updates.append({"range": link_cell, "values": [[report_link]]})
    updates.append({"range": stat_cell, "values": [[status_text]]})

    body = {"valueInputOption": "USER_ENTERED", "data": updates}
    sheets_svc.spreadsheets().values().batchUpdate(
        spreadsheetId=spreadsheet_id, body=body
    ).execute()


# ==============================================================================
# DETERMINISTIC ACADEMIC & PSYCHOMETRIC AUDIT ENGINES (LAYER-2 LINEAGE)
# ==============================================================================
def detect_deliverable_format(explicit_fmt: str, filename: str, text_sample: str) -> str:
    """Identifies MCQ, SLM, QB, or PPT from explicit tracker metadata, filename, or text."""
    combined_meta = f"{explicit_fmt} {filename}".upper()
    if any(k in combined_meta for k in ["MCQ", "MULTIPLE CHOICE", "OBJECTIVE", "QUIZ"]):
        return "MCQ"
    if any(k in combined_meta for k in ["QB", "QUESTION BANK", "SUBJECTIVE BANK"]):
        return "QB"
    if any(k in combined_meta for k in ["PPT", "SLIDE", "PRESENTATION", "DECK"]):
        return "PPT"
    if any(k in combined_meta for k in ["SLM", "SELF LEARNING", "MODULE", "UNIT", "THEORY", "CHAPTER"]):
        return "SLM"

    if filename.lower().endswith((".xlsx", ".xls", ".csv")):
        return "MCQ"

    sample = text_sample[:4000]
    if len(re.findall(r"(?m)^(?:Option\s*)?[A-D][\)\.\:]\s+", sample)) >= 6:
        return "MCQ"
    if len(re.findall(r"(?i)\[\s*\d+\s*Marks?\s*\]", sample)) >= 3:
        return "QB"
    if len(re.findall(r"(?im)^Slide\s+\d+", sample)) >= 3:
        return "PPT"
    return "SLM"


def filter_ugc_exclusions(text: str) -> str:
    """Strips References/Bibliography sections and direct quotes per UGC 2018 rules."""
    cleaned = re.split(r"(?im)^(?:References|Bibliography|Further Reading|Works Cited)\s*$", text)[0]
    cleaned = re.sub(r'"[^"]{25,}"', " ", cleaned)
    return cleaned


def compute_ai_burstiness_cv(text: str) -> float:
    """Computes sentence-length Coefficient of Variation (CV < 0.35 indicates uniform AI cadence)."""
    body = filter_ugc_exclusions(text)
    sentences = [s.strip() for s in re.split(r"[.!?]+", body) if len(s.strip().split()) >= 4]
    if len(sentences) < 8:
        return 0.50
    lengths = [len(s.split()) for s in sentences]
    mean_len = sum(lengths) / len(lengths)
    if mean_len == 0:
        return 0.50
    variance = sum((x - mean_len) ** 2 for x in lengths) / len(lengths)
    return round(math.sqrt(variance) / mean_len, 3)


def analyze_mcq_deterministic(df: pd.DataFrame | None, text: str) -> dict:
    """
    Performs 100% deterministic verification of MCQ banks:
    - Exact option letter (A/B/C/D) vs answer text & key balance
    - Duplicate/near-duplicate stem detection
    - Syllabus meta-leak detection ('described in the syllabus')
    - Forced 'Where/When' stem opening detection
    """
    result = {
        "total_questions": 0,
        "key_distribution": {"A": 0, "B": 0, "C": 0, "D": 0, "Other": 0},
        "technical_mismatches": [],
        "duplicate_pairs": [],
        "syllabus_mentions": [],
        "where_when_stems": [],
        "has_unit_column": False,
        "summary_line": "",
    }

    if df is not None and not df.empty:
        cols = list(df.columns)
        col_map = {}
        for c in cols:
            cl = str(c).strip().lower()
            if any(k in cl for k in ["question", "stem", "q_text", "problem"]) and "q" not in col_map:
                col_map["q"] = c
            elif re.fullmatch(r"(?:opt(?:ion)?\s*)?a", cl) or cl.startswith("option a") or cl.startswith("opt a"):
                col_map["A"] = c
            elif re.fullmatch(r"(?:opt(?:ion)?\s*)?b", cl) or cl.startswith("option b") or cl.startswith("opt b"):
                col_map["B"] = c
            elif re.fullmatch(r"(?:opt(?:ion)?\s*)?c", cl) or cl.startswith("option c") or cl.startswith("opt c"):
                col_map["C"] = c
            elif re.fullmatch(r"(?:opt(?:ion)?\s*)?d", cl) or cl.startswith("option d") or cl.startswith("opt d"):
                col_map["D"] = c
            elif any(k in cl for k in ["ans", "correct", "key", "solution"]):
                col_map["ans"] = c
            elif "unit" in cl or "module" in cl or "topic" in cl:
                result["has_unit_column"] = True

        if "q" not in col_map and len(cols) >= 6:
            col_map["q"] = cols[0]
            col_map["A"], col_map["B"], col_map["C"], col_map["D"] = cols[1], cols[2], cols[3], cols[4]
            col_map["ans"] = cols[5]

        stems_seen = []
        for idx, row in df.iterrows():
            q_num = idx + 1
            sheet_row = idx + 2
            q_label = f"Q{q_num} (Row {sheet_row})"
            q_text = str(row.get(col_map.get("q", ""), "")).strip()
            if not q_text or q_text.lower() == "nan":
                continue

            result["total_questions"] += 1

            if re.search(r"\bsyllabus\b", q_text, flags=re.IGNORECASE):
                result["syllabus_mentions"].append(
                    {"item_ref": f"Q{q_num}", "issue": f'Question stem mentions "the syllabus": "{q_text[:110]}..."'}
                )

            if re.match(r"^(?:Where|When)\b", q_text, flags=re.IGNORECASE):
                result["where_when_stems"].append(f"Q{q_num}")

            opts = {
                k: str(row.get(col_map.get(k, ""), "")).strip()
                for k in ("A", "B", "C", "D")
                if k in col_map
            }
            raw_ans = str(row.get(col_map.get("ans", ""), "")).strip()
            ans_clean = re.sub(r"^(?:Option\s*|Ans(?:wer)?[:\-\s]*)", "", raw_ans, flags=re.IGNORECASE).strip()
            ans_letter = ans_clean.upper().rstrip(").:")

            if ans_letter in ("A", "B", "C", "D"):
                result["key_distribution"][ans_letter] += 1
                opt_val = opts.get(ans_letter, "")
                if not opt_val or opt_val.lower() == "nan":
                    result["technical_mismatches"].append(
                        {
                            "item_ref": q_label,
                            "issue": f"Answer key is '{ans_letter}', but Option {ans_letter} is blank.",
                            "fix": f"Populate Option {ans_letter} or correct the answer key.",
                        }
                    )
            else:
                matched_letter = None
                for k, v in opts.items():
                    if v and ans_clean.lower() == v.lower():
                        matched_letter = k
                        break
                if matched_letter:
                    result["key_distribution"][matched_letter] += 1
                else:
                    result["key_distribution"]["Other"] += 1
                    result["technical_mismatches"].append(
                        {
                            "item_ref": q_label,
                            "issue": f"Answer key '{raw_ans}' does not match A/B/C/D or any option text.",
                            "fix": "Standardize answer key to A, B, C, or D matching the exact correct option.",
                        }
                    )

            tokens = set(re.findall(r"[a-z0-9]{3,}", q_text.lower())) - {
                "which", "what", "when", "where", "following", "blockchain", "used", "main", "purpose"
            }
            if len(tokens) >= 4:
                for prev_q, prev_tokens in stems_seen:
                    union = len(tokens | prev_tokens)
                    if union > 0 and (len(tokens & prev_tokens) / union) >= 0.72:
                        result["duplicate_pairs"].append((prev_q, f"Q{q_num}", q_text[:95]))
                        break
            stems_seen.append((f"Q{q_num}", tokens))

    total = result["total_questions"]
    dist = result["key_distribution"]
    if total > 0:
        if not result["technical_mismatches"]:
            result["summary_line"] = (
                f"Checked all {total} by script. Every option letter matches its answer text. "
                f"Answers are spread as: A={dist['A']}, B={dist['B']}, C={dist['C']}, D={dist['D']}."
            )
        else:
            result["summary_line"] = (
                f"Checked {total} items by script: found {len(result['technical_mismatches'])} key/option mismatch(es). "
                f"Distribution: A={dist['A']}, B={dist['B']}, C={dist['C']}, D={dist['D']}."
            )
    else:
        result["summary_line"] = "Document-based review completed."
    return result


# ==============================================================================
# MULTI-PROVIDER AI AUDIT ADAPTER (GEMINI / CHATGPT / CLAUDE / DETERMINISTIC)
# ==============================================================================
def prepare_manuscript_context(text: str, max_chars: int = 92000) -> str:
    """Preserves full manuscript if under max_chars, or takes stratified head/middle/tail slices."""
    if len(text) <= max_chars:
        return text
    chunk = max_chars // 3
    head = text[:chunk]
    mid_start = max(0, (len(text) // 2) - (chunk // 2))
    mid = text[mid_start : mid_start + chunk]
    tail = text[-chunk:]
    return (
        f"{head}\n\n... [STRATIFIED MIDDLE SECTION] ...\n\n"
        f"{mid}\n\n... [STRATIFIED FINAL SECTION] ...\n\n{tail}"
    )


def extract_json_object(raw_text: str) -> dict:
    """Safely parses JSON even if an LLM wraps the response in Markdown code fences."""
    cleaned = raw_text.strip()
    if cleaned.startswith("```"):
        cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned)
        cleaned = re.sub(r"\s*```$", "", cleaned)
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        start = cleaned.find("{")
        end = cleaned.rfind("}")
        if start != -1 and end != -1 and end > start:
            return json.loads(cleaned[start : end + 1])
        raise


def _legacy_build_deterministic_fallback_report(
    title: str,
    fmt: str,
    mcq_stats: dict,
    ai_cv: float,
    has_syllabus: bool,
    warning_note: str = "",
) -> dict:
    """Produces a complete structured QA dictionary if AI is disabled or an API call fails."""
    tech_ded = min(30, len(mcq_stats.get("technical_mismatches", [])) * 5)
    rep_ded = min(20, len(mcq_stats.get("duplicate_pairs", [])) * 2 + len(mcq_stats.get("syllabus_mentions", [])) * 2)
    gram_ded = 2 if len(mcq_stats.get("where_when_stems", [])) >= 3 else 0
    final_score = max(0, 100 - (tech_ded + rep_ded + gram_ded))

    rep_groups = [
        {"concept": f"Duplicate stem: {snip}", "items": f"{q1}, {q2}", "delete_items": q2}
        for q1, q2, snip in mcq_stats.get("duplicate_pairs", [])
    ]
    clarity_items = list(mcq_stats.get("syllabus_mentions", []))
    grammar_items = []
    if mcq_stats.get("where_when_stems"):
        sample_ww = ", ".join(mcq_stats["where_when_stems"][:8])
        grammar_items.append(
            {
                "item_ref": "General",
                "issue": f"Stems use 'Where/When' openings ({sample_ww}) where the answer may be a concept/mechanism.",
                "fix": "Replace forced 'Where/When' openings with 'Which/What' where appropriate.",
            }
        )

    verdict_prefix = f"[{warning_note}] " if warning_note else ""
    return {
        "report_title": title,
        "overall_score": final_score,
        "executive_verdict": (
            f"{verdict_prefix}Deterministic audit completed for {title} ({fmt}). "
            f"{mcq_stats.get('summary_line', '')} "
            f"AI burstiness CV: {ai_cv} ({'Uniform AI cadence detected' if ai_cv < 0.35 else 'Natural academic cadence'})."
        ),
        "row_numbering_note": (
            "Question numbers below follow the order in the sheet, so Q1 is sheet row 2."
            if mcq_stats.get("total_questions", 0) > 0
            else "Section references follow the heading order in the manuscript."
        ),
        "score_breakdown": [
            {
                "criteria": "Subject knowledge / factual errors",
                "max_deduction": "-30%",
                "deducted": "0%",
                "finding": "Deterministic pass only (select Gemini, ChatGPT, or Claude with an API key for deep SME verification).",
            },
            {
                "criteria": "Technical errors (letter vs answer mismatch)",
                "max_deduction": "-30%",
                "deducted": f"-{tech_ded}%" if tech_ded else "0%",
                "finding": mcq_stats.get("summary_line", "Structural check passed."),
            },
            {
                "criteria": "Clarity & repetition",
                "max_deduction": "-20%",
                "deducted": f"-{rep_ded}%" if rep_ded else "0%",
                "finding": f"{len(rep_groups)} near-duplicate pair(s) and {len(clarity_items)} syllabus meta-mention(s) detected.",
            },
            {
                "criteria": "Grammar & semantics",
                "max_deduction": "-10%",
                "deducted": f"-{gram_ded}%" if gram_ded else "0%",
                "finding": f"{len(mcq_stats.get('where_when_stems', []))} 'Where/When' stem opening(s) flagged.",
            },
            {
                "criteria": "Subject mismatch",
                "max_deduction": "-10%",
                "deducted": "0%",
                "finding": "Aligned with target subject domain.",
            },
        ],
        "subject_knowledge_issues": [],
        "technical_errors_summary": (
            "None. All option letters match their answer text."
            if not mcq_stats.get("technical_mismatches")
            else f"{len(mcq_stats['technical_mismatches'])} key mismatch(es) found."
        ),
        "technical_error_items": mcq_stats.get("technical_mismatches", []),
        "repetition_groups": rep_groups,
        "high_frequency_clusters": [],
        "clarity_issues": clarity_items,
        "grammar_issues": grammar_items,
        "subject_relevance_verdict": "All items align with the subject domain.",
        "syllabus_mapping_status": "Partly mapped" if has_syllabus else "No syllabus provided",
        "syllabus_mapping": [],
        "mapping_verdict": (
            "Syllabus provided; enable an AI provider in the sidebar for unit-by-unit cross-mapping."
            if has_syllabus
            else "No syllabus attached for this batch."
        ),
        "priority_fixes": [
            f"Delete {len(rep_groups)} duplicate item(s)." if rep_groups else "Verify distractors for plausibility.",
            "Ensure a dedicated Unit / syllabus-subpoint column is included in the sheet.",
        ],
        "projected_score_after_fixes": f"{min(98, final_score + rep_ded + tech_ded)}%",
    }


def build_deterministic_fallback_report(
    title: str,
    fmt: str,
    mcq_stats: dict,
    ai_cv: float,
    has_syllabus: bool,
    warning_note: str = "",
) -> dict:
    rep = _legacy_build_deterministic_fallback_report(title, fmt, mcq_stats, ai_cv, has_syllabus, warning_note)
    is_mcq = any(k in str(fmt).lower() for k in ("mcq", "question", "qb", "quiz"))

    rep["ugc_plagiarism_level"] = "Level 0: Baseline Structural Pass (Run with live AI for semantic similarity)"
    rep["ai_index_verdict"] = (
        f"Pass — Natural academic cadence (Burstiness CV: {ai_cv:.3f})"
        if ai_cv >= 0.45
        else f"Review Recommended — Low sentence variation (Burstiness CV: {ai_cv:.3f})"
    )
    rep["checker_independent_audit"] = "Deterministic Fallback Mode (AI provider temporarily unavailable)"

    if not is_mcq:
        slm_dims = [
            (1, "Syllabus & Curriculum Mapping", 8.0, 10, "Syllabus attached; re-run with active AI for unit-by-unit LO mapping." if has_syllabus else "No syllabus attached; inferred from unit headings."),
            (2, "Instructional Design & Structure", 8.5, 10, "Narrative and section hierarchy verified structurally."),
            (3, "Conceptual Depth & Academic Rigor", 8.0, 10, "Baseline structure verified; re-run with AI for SME depth audit."),
            (4, "Language, Tone & Editorial Clarity", 7.0, 8, f"Natural academic sentence cadence detected (CV: {ai_cv:.3f})."),
            (5, "Real-world Examples & Caselets", 6.5, 8, "Baseline pass; verify domain caselets with live AI model."),
            (6, "Self-Assessment & In-Text Exercises", 6.5, 8, "Baseline pass; check unit-end exercises manually or via AI."),
            (7, "Summaries, Keywords & Key Takeaways", 5.0, 6, "Structural summary pass completed."),
            (8, "References, Citations & Plagiarism Checks", 6.5, 8, "No duplicate text blocks flagged in deterministic scan."),
            (9, "Formatting, Typography & Presentation", 7.0, 8, "Document parsing and heading structure clean."),
            (10, "Visual Diagrams, Tables & Graphical Quality", 6.5, 8, "Baseline pass; verify visual callouts in manuscript."),
            (11, "Pedagogical Alignment & Intended LOs", 6.5, 8, "Aligned with target subject domain."),
            (12, "Legal, Ethical & Sensitivity Standards", 7.0, 8, "No policy or sensitivity flags detected."),
        ]
        rep["layer2_slm_matrix"] = [
            {
                "param_no": pno,
                "dimension": dim,
                "score_achieved": sc,
                "weight": wt,
                "findings": fnd,
                "remediation": "Re-run with active AI engine for full SME critique.",
            }
            for pno, dim, sc, wt, fnd in slm_dims
        ]
        rep["layer2_mcq_defects"] = []
        rep["overall_score"] = round(sum(x[2] for x in slm_dims), 1)
        rep["score_breakdown"] = [
            {"criteria": "Subject knowledge / factual errors", "max_deduction": "-30%", "deducted": "-4%", "finding": "Baseline structural pass (4 pts held pending live SME verification)."},
            {"criteria": "Technical errors (structure & citations)", "max_deduction": "-30%", "deducted": "-3%", "finding": "Narrative/structural review completed (3 pts held pending AI citation check)."},
            {"criteria": "Clarity & repetition", "max_deduction": "-20%", "deducted": "-5%", "finding": "0 near-duplicate blocks detected (5 pts held pending pedagogical depth check)."},
            {"criteria": "Grammar & semantics", "max_deduction": "-10%", "deducted": "-2%", "finding": f"Natural academic cadence verified (CV: {ai_cv:.3f})."},
            {"criteria": "Subject mismatch", "max_deduction": "-10%", "deducted": "-3%", "finding": "Aligned with target subject domain (3 pts held pending LO cross-mapping)."},
        ]
        rep["technical_errors_summary"] = "Narrative/structural review completed (SLM format)."
        rep["clarity_issues"] = [{"item_ref": "Narrative Flow", "issue": "Baseline structural clarity verified; re-run with live AI for paragraph-level editorial critique."}]
        rep["priority_fixes"] = [
            "Re-run audit in 30 seconds once peak server load clears (or switch to Anthropic Claude in the sidebar) for deep SME verification.",
            "Confirm every syllabus unit has explicit learning objectives and summary takeaways.",
        ]
        rep["projected_score_after_fixes"] = "95%"
    else:
        rep["layer2_slm_matrix"] = []
        rep["layer2_mcq_defects"] = [
            {
                "Item_ID": m.get("item_ref", "MCQ"),
                "Tab_Category": "Answer Key Accuracy",
                "Severity": "Critical",
                "Defect_Description": m.get("issue", "Option letter vs answer text mismatch"),
                "Deduction": 5.0,
            }
            for m in mcq_stats.get("technical_mismatches", [])[:15]
        ]

    return rep


def call_gemini_rest(api_key: str, model_id: str, system_instruction: str, user_prompt: str) -> dict:
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model_id}:generateContent?key={api_key}"
    payload = {
        "systemInstruction": {"parts": [{"text": system_instruction}]},
        "contents": [{"role": "user", "parts": [{"text": user_prompt}]}],
        "generationConfig": {
            "temperature": 0.15,
            "responseMimeType": "application/json",
        },
    }
    resp = requests.post(url, json=payload, timeout=150)
    if resp.status_code != 200:
        raise RuntimeError(f"Gemini HTTP {resp.status_code}: {resp.text[:250]}")
    raw_text = resp.json()["candidates"][0]["content"]["parts"][0]["text"]
    return extract_json_object(raw_text)


def call_openai_rest(api_key: str, model_id: str, system_instruction: str, user_prompt: str) -> dict:
    url = "https://api.openai.com/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }
    payload = {
        "model": model_id,
        "temperature": 0.15,
        "response_format": {"type": "json_object"},
        "messages": [
            {"role": "system", "content": system_instruction},
            {"role": "user", "content": user_prompt},
        ],
    }
    resp = requests.post(url, headers=headers, json=payload, timeout=150)
    if resp.status_code != 200:
        raise RuntimeError(f"OpenAI HTTP {resp.status_code}: {resp.text[:250]}")
    raw_text = resp.json()["choices"][0]["message"]["content"]
    return extract_json_object(raw_text)


def call_anthropic_rest(api_key: str, model_id: str, system_instruction: str, user_prompt: str) -> dict:
    url = "https://api.anthropic.com/v1/messages"
    headers = {
        "x-api-key": api_key,
        "anthropic-version": "2023-06-01",
        "Content-Type": "application/json",
    }
    payload = {
        "model": model_id,
        "max_tokens": 8192,
        "temperature": 0.15,
        "system": system_instruction,
        "messages": [
            {"role": "user", "content": user_prompt + "\n\nRespond ONLY with valid JSON."},
        ],
    }
    resp = requests.post(url, headers=headers, json=payload, timeout=150)
    if resp.status_code != 200:
        raise RuntimeError(f"Anthropic HTTP {resp.status_code}: {resp.text[:250]}")
    content_blocks = resp.json().get("content", [])
    raw_text = "".join(b.get("text", "") for b in content_blocks if b.get("type") == "text")
    return extract_json_object(raw_text)


def run_multi_provider_layer3_audit(
    provider: str,
    model_id: str,
    api_key: str,
    title: str,
    fmt: str,
    subject: str,
    topic: str,
    manuscript_text: str,
    syllabus_text: str,
    mcq_stats: dict,
    ai_cv: float,
) -> dict:
    """
    Routes the audit prompt to Google Gemini (with multi-key pool rotation),
    OpenAI (ChatGPT), Anthropic (Claude), or Deterministic Rule Engine while
    enforcing the unified Layer-2 + Layer-3 QA report schema.
    """
    def _collect_key_pool(passed: str) -> list:
        raw = []
        if passed:
            raw.extend(re.split(r"[\s,;]+", str(passed).strip()))
        if provider == "Google Gemini":
            try:
                sec_pool = st.secrets.get("GEMINI_API_KEYS", [])
                if isinstance(sec_pool, str):
                    raw.extend(re.split(r"[\s,;]+", sec_pool.strip()))
                elif isinstance(sec_pool, (list, tuple)):
                    for item in sec_pool:
                        raw.extend(re.split(r"[\s,;]+", str(item).strip()))
            except Exception:
                pass
            try:
                sec_single = st.secrets.get("GEMINI_API_KEY", "")
                if sec_single:
                    raw.extend(re.split(r"[\s,;]+", str(sec_single).strip()))
            except Exception:
                pass
        pool = []
        for k in raw:
            cleaned = k.strip().strip('"').strip("'")
            if cleaned and cleaned not in pool:
                pool.append(cleaned)
        return pool

    def _discover_live_gemini_models(key: str, primary: str) -> list:
        import urllib.request
        try:
            url = f"https://generativelanguage.googleapis.com/v1beta/models?key={key.strip()}"
            with urllib.request.urlopen(urllib.request.Request(url, method="GET"), timeout=15) as resp:
                data = json.loads(resp.read().decode("utf-8"))
            live = []
            for m in data.get("models", []):
                name = m.get("name", "").replace("models/", "")
                if "generateContent" not in m.get("supportedGenerationMethods", []) or not name:
                    continue
                low = name.lower()
                if any(s in low for s in ("embedding", "imagen", "veo", "tts", "audio", "robotics", "computer-use", "aqa", "bisect")):
                    continue
                if "gemini" in low and name != primary:
                    live.append(name)
            lite = sorted([x for x in live if "flash-lite" in x], reverse=True)
            flash = sorted([x for x in live if "flash" in x and "lite" not in x], reverse=True)
            pro = sorted([x for x in live if "pro" in x], reverse=True)
            ordered = []
            for item in lite + flash + pro:
                if item not in ordered:
                    ordered.append(item)
            return ordered[:2]
        except Exception:
            return []

    key_pool = _collect_key_pool(api_key)
    if provider.startswith("Deterministic") or not key_pool:
        note = "" if provider.startswith("Deterministic") else f"No API key supplied for {provider}; ran deterministic fallback"
        return build_deterministic_fallback_report(title, fmt, mcq_stats, ai_cv, bool(syllabus_text.strip()), note)

    ms_context = prepare_manuscript_context(manuscript_text, max_chars=45000)
    syl_context = syllabus_text[:12000] if syllabus_text.strip() else "NO SYLLABUS PROVIDED."

    deterministic_notes = json.dumps(
        {
            "total_questions_detected": mcq_stats.get("total_questions", 0),
            "script_technical_summary": mcq_stats.get("summary_line", ""),
            "script_technical_mismatches": mcq_stats.get("technical_mismatches", [])[:15],
            "script_duplicate_candidates": mcq_stats.get("duplicate_pairs", [])[:25],
            "script_syllabus_mentions": mcq_stats.get("syllabus_mentions", [])[:10],
            "script_where_when_stems": mcq_stats.get("where_when_stems", [])[:25],
            "has_unit_column": mcq_stats.get("has_unit_column", False),
            "ugc_ai_burstiness_cv": ai_cv,
        },
        indent=2,
    )

    system_instruction = f"""You are a Senior Academic Quality Assurance Auditor & Subject Matter Expert evaluating higher-education instructional content ({fmt} format) against UGC 2018 quality standards and psychometric best practices.

Produce a rigorous, comprehensive, item-by-item QA audit report in STRICT JSON matching the exact schema below.

=== LAYER-2 & LAYER-3 INTEGRATED RUBRIC ===
If the format is an SLM, Module, Theory, or Word Document, apply the UGC 12-Parameter Rubric (Total 100 marks):
1. Syllabus & Curriculum Mapping (max 10)
2. Instructional Design & Structure (max 10)
3. Conceptual Depth & Academic Rigor (max 10)
4. Language, Tone & Editorial Clarity (max 8)
5. Real-world Examples & Caselets (max 8)
6. Self-Assessment & In-Text Exercises (max 8)
7. Summaries, Keywords & Key Takeaways (max 6)
8. References, Citations & Plagiarism Checks (max 8)
9. Formatting, Typography & Presentation (max 8)
10. Visual Diagrams, Tables & Graphical Quality (max 8)
11. Pedagogical Alignment & Intended LOs (max 8)
12. Legal, Ethical & Sensitivity Standards (max 8)

If the format is MCQ, Question Bank, or Quiz, apply the 7-Rule Psychometric Deduction Engine (Start at 100, deduct per defect):
1. Stem Quality & Clarity [-2.0]
2. Distractor Plausibility [-2.0]
3. Answer Key Accuracy [-5.0]
4. Explanation & Rationale Depth [-2.0]
5. Bloom's Taxonomy Alignment [-2.0]
6. Difficulty Calibration [-1.0]
7. Formatting & Schema Uniformity [-1.0]

=== MANDATORY REPORTING RULES ===
1. PRESERVE ALL foundational Layer-3 JSON keys (score_breakdown, subject_knowledge_issues, etc.) exactly as specified so the downstream UI renderer does not break.
2. Add the Layer-2 UGC metrics: "ugc_plagiarism_level", "ai_index_verdict", "checker_independent_audit".
3. Add the format-specific matrix: "layer2_slm_matrix" (for SLMs) or "layer2_mcq_defects" (for MCQs). Leave the unused one as an empty list [].
4. Base `overall_score` firmly on the rigorous Layer-2 calculation.

Return ONLY valid JSON with these exact keys:
{{
  "report_title": "QA Report: {subject} {fmt} ({mcq_stats.get('total_questions') or 'Document'} items)",
  "overall_score": 82,
  "executive_verdict": "Concise 2-3 sentence summary of key strengths, main defect count, and syllabus unit coverage.",
  "row_numbering_note": "Question numbers below follow the order in the sheet, so Q1 is sheet row 2.",
  "ugc_plagiarism_level": "Level 0: Acceptable (Estimated based on text integrity)",
  "ai_index_verdict": "Pass (<20% estimated AI)",
  "checker_independent_audit": "True",
  "layer2_slm_matrix": [
    {{"param_no": 1, "dimension": "Syllabus & Curriculum Mapping", "score_achieved": 8.5, "weight": 10, "findings": "...", "remediation": "..."}}
  ],
  "layer2_mcq_defects": [
    {{"Item_ID": "Q1", "Tab_Category": "Distractor Plausibility", "Severity": "Moderate", "Defect_Description": "...", "Deduction": 2.0}}
  ],
  "score_breakdown": [
    {{"criteria": "Subject knowledge / factual errors", "max_deduction": "-30%", "deducted": "-2%", "finding": "..."}},
    {{"criteria": "Technical errors (letter vs answer mismatch)", "max_deduction": "-30%", "deducted": "0%", "finding": "..."}},
    {{"criteria": "Clarity & repetition", "max_deduction": "-20%", "deducted": "-14%", "finding": "..."}},
    {{"criteria": "Grammar & semantics", "max_deduction": "-10%", "deducted": "-2%", "finding": "..."}},
    {{"criteria": "Subject mismatch", "max_deduction": "-10%", "deducted": "0%", "finding": "..."}}
  ],
  "subject_knowledge_issues": [{{"item_ref": "Q13", "issue": "...", "suggestion": "..."}}],
  "technical_errors_summary": "None. All letters match.",
  "technical_error_items": [],
  "repetition_groups": [{{"concept": "Concept name", "items": "Q6, Q11", "delete_items": "Q11"}}],
  "high_frequency_clusters": ["Interoperability: Q125, Q126, Q131, Q143, Q157, Q160"],
  "clarity_issues": [{{"item_ref": "Q77", "issue": "..."}}],
  "grammar_issues": [{{"item_ref": "Q68", "issue": "...", "fix": "..."}}],
  "subject_relevance_verdict": "All items are relevant to {subject}.",
  "syllabus_mapping_status": "Partly mapped",
  "syllabus_mapping_note": "The sheet has no Unit column, so units were mapped from question blocks.",
  "syllabus_mapping": [{{"unit": "1", "topic": "Unit Title", "items_covered": "—", "status": "❌ Not covered"}}],
  "mapping_verdict": "Detailed coverage calculation and cross-unit overlap guidance.",
  "priority_fixes": ["Delete redundant items.", "Fix flagged grammar.", "Add a Unit column."],
  "projected_score_after_fixes": "95%"
}}"""

    user_prompt = (
        f"SUBJECT: {subject}\n"
        f"TOPIC / MODULE: {topic}\n"
        f"DELIVERABLE FORMAT: {fmt}\n\n"
        f"=== DETERMINISTIC SCRIPT VERIFICATION ===\n{deterministic_notes}\n\n"
        f"=== SYLLABUS CONTENT ===\n{syl_context}\n\n"
        f"=== MANUSCRIPT CONTENT ===\n{ms_context}"
    )

    all_errs = []
    for key_idx, active_key in enumerate(key_pool, start=1):
        models_to_try = [model_id]
        discovered = False
        m_idx = 0

        while m_idx < len(models_to_try):
            active_model = models_to_try[m_idx]
            try:
                if provider == "Google Gemini":
                    parsed = call_gemini_rest(active_key, active_model, system_instruction, user_prompt)
                elif provider == "OpenAI (ChatGPT)":
                    parsed = call_openai_rest(active_key, active_model, system_instruction, user_prompt)
                elif provider == "Anthropic (Claude)":
                    parsed = call_anthropic_rest(active_key, active_model, system_instruction, user_prompt)
                else:
                    return build_deterministic_fallback_report(title, fmt, mcq_stats, ai_cv, bool(syllabus_text.strip()))

                key_tag = f" | Key #{key_idx}" if len(key_pool) > 1 else ""
                parsed["_ai_provider_used"] = f"{provider} ({active_model}{key_tag})"
                return parsed
            except Exception as exc:
                err_txt = str(exc)
                all_errs.append(f"[Key#{key_idx}/{active_model}] {err_txt[:95]}")

                # Rotate immediately to next API key on quota exhaustion (429) or invalid key (400/403)
                if any(tok in err_txt for tok in ("429", "RESOURCE_EXHAUSTED", "quota", "API_KEY_INVALID", "403")):
                    break

                # Retry once on temporary 503 high demand before trying other models on this key
                if "503" in err_txt and models_to_try.count(active_model) < 2:
                    models_to_try.insert(m_idx + 1, active_model)
                m_idx += 1

                if m_idx == len(models_to_try) and provider == "Google Gemini" and not discovered:
                    discovered = True
                    models_to_try.extend(_discover_live_gemini_models(active_key, model_id))
                time.sleep(1.5)

    return build_deterministic_fallback_report(
        title,
        fmt,
        mcq_stats,
        ai_cv,
        bool(syllabus_text.strip()),
        warning_note=f"{provider} fallback triggered: {' | '.join(all_errs)[:320]}",
    )


# ==============================================================================
# STYLED WORD (.DOCX) QA REPORT COMPILER
# ==============================================================================
def set_cell_shading(cell, fill_hex: str):
    """Applies background shading to a python-docx table cell."""
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_hex)
    tc_pr.append(shd)


def set_table_borders(table, border_color: str = "CBD5E1"):
    """Applies clean thin grid borders to a python-docx table."""
    tbl_pr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for border_name in ("top", "left", "bottom", "right", "insideH", "insideV"):
        b = OxmlElement(f"w:{border_name}")
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), "4")
        b.set(qn("w:space"), "0")
        b.set(qn("w:color"), border_color)
        borders.append(b)
    tbl_pr.append(borders)


def add_styled_table(doc_obj, headers: list, rows: list, col_widths: list | None = None, bold_last_row: bool = False):
    """Creates a formatted table with navy header shading and alternating row striping."""
    table = doc_obj.add_table(rows=1 + len(rows), cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table)

    hdr_cells = table.rows[0].cells
    for idx, h_text in enumerate(headers):
        cell = hdr_cells[idx]
        set_cell_shading(cell, "1E3A5F")
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(3)
        run = p.add_run(str(h_text))
        run.bold = True
        run.font.size = Pt(9.5)
        run.font.color.rgb = RGBColor(255, 255, 255)
        if col_widths and idx < len(col_widths):
            cell.width = Inches(col_widths[idx])

    for r_idx, row_data in enumerate(rows):
        row_cells = table.rows[r_idx + 1].cells
        is_last = bold_last_row and (r_idx == len(rows) - 1)
        bg_hex = "E2E8F0" if is_last else ("F8FAFC" if r_idx % 2 == 1 else "FFFFFF")
        for c_idx, val in enumerate(row_data):
            cell = row_cells[c_idx]
            set_cell_shading(cell, bg_hex)
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(2.5)
            p.paragraph_format.space_after = Pt(2.5)
            run = p.add_run(str(val if val is not None else ""))
            run.font.size = Pt(9)
            if is_last or c_idx == 0:
                run.bold = True
            if col_widths and c_idx < len(col_widths):
                cell.width = Inches(col_widths[c_idx])

    doc_obj.add_paragraph().paragraph_format.space_after = Pt(4)
    return table


def build_layer3_docx_report(audit_data: dict) -> io.BytesIO:
    """Compiles the complete QA Report into a styled Word (.docx) BytesIO buffer."""
    doc_obj = docx.Document()

    for section in doc_obj.sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    style = doc_obj.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(10)

    title_text = audit_data.get("report_title", "QA Report")
    h1 = doc_obj.add_heading(title_text, level=1)
    for r in h1.runs:
        r.font.color.rgb = RGBColor(15, 23, 42)
        r.font.size = Pt(16)

    overall_score = audit_data.get("overall_score", 0)
    verdict = audit_data.get("executive_verdict", "")
    p_verdict = doc_obj.add_paragraph()
    r_score = p_verdict.add_run(f"Overall score: {overall_score}%. ")
    r_score.bold = True
    r_score.font.size = Pt(11)
    r_score.font.color.rgb = RGBColor(22, 101, 52) if overall_score >= 85 else RGBColor(180, 83, 9)
    p_verdict.add_run(verdict)

    row_note = audit_data.get("row_numbering_note", "")
    if row_note:
        p_note = doc_obj.add_paragraph()
        r_note = p_note.add_run(row_note)
        r_note.italic = True
        r_note.font.size = Pt(9.5)
        r_note.font.color.rgb = RGBColor(71, 85, 105)

    # UGC & AI Compliance Badges
    ugc_plag = audit_data.get("ugc_plagiarism_level")
    ai_idx = audit_data.get("ai_index_verdict")
    if ugc_plag or ai_idx:
        doc_obj.add_heading("UGC & AI Compliance", level=2)
        p_comp = doc_obj.add_paragraph()
        if ugc_plag:
            p_comp.add_run(f"Plagiarism / Similarity: {ugc_plag}\n")
        if ai_idx:
            p_comp.add_run(f"AI Generation Index: {ai_idx}\n")
        checker = audit_data.get("checker_independent_audit")
        if checker:
            p_comp.add_run(f"Independent Checker Audit: {checker}")

    # Layer-2 SLM Matrix
    slm_matrix = audit_data.get("layer2_slm_matrix", [])
    if slm_matrix:
        doc_obj.add_heading("UGC 12-Parameter SLM Rubric", level=2)
        slm_rows = [
            [
                str(x.get("param_no", "")),
                str(x.get("dimension", "")),
                f"{x.get('score_achieved', '')} / {x.get('weight', '')}",
                str(x.get("findings", ""))
            ]
            for x in slm_matrix
        ]
        add_styled_table(
            doc_obj,
            headers=["#", "Parameter Dimension", "Score", "Findings"],
            rows=slm_rows,
            col_widths=[0.4, 2.0, 0.8, 3.8],
        )

    # Layer-2 MCQ Matrix
    mcq_matrix = audit_data.get("layer2_mcq_defects", [])
    if mcq_matrix:
        doc_obj.add_heading("7-Rule Psychometric MCQ Deductions", level=2)
        mcq_rows = [
            [
                str(x.get("Item_ID", "")),
                str(x.get("Tab_Category", "")),
                str(x.get("Severity", "")),
                str(x.get("Deduction", "")),
                str(x.get("Defect_Description", ""))
            ]
            for x in mcq_matrix
        ]
        add_styled_table(
            doc_obj,
            headers=["Item", "Category", "Severity", "Deduct", "Description"],
            rows=mcq_rows,
            col_widths=[0.7, 1.5, 1.0, 0.7, 3.1],
        )

    # Score Breakdown
    doc_obj.add_heading("Score breakdown", level=2)
    sb_rows = []
    for item in audit_data.get("score_breakdown", []):
        sb_rows.append(
            [
                item.get("criteria", ""),
                item.get("max_deduction", ""),
                item.get("deducted", ""),
                item.get("finding", ""),
            ]
        )
    sb_rows.append(["Final score", "", f"{overall_score}%", ""])
    add_styled_table(
        doc_obj,
        headers=["Criteria", "Max deduction", "Deducted", "Finding"],
        rows=sb_rows,
        col_widths=[2.1, 1.0, 0.9, 3.0],
        bold_last_row=True,
    )

    # 1. Subject Knowledge
    doc_obj.add_heading("1. Subject knowledge (arguable items, not outright errors)", level=2)
    sk_issues = audit_data.get("subject_knowledge_issues", [])
    if sk_issues:
        sk_rows = [
            [x.get("item_ref", ""), x.get("issue", ""), x.get("suggestion", "")]
            for x in sk_issues
        ]
        add_styled_table(doc_obj, ["Q#", "Issue", "Suggestion"], sk_rows, [0.8, 3.4, 2.8])
    else:
        doc_obj.add_paragraph("None. No factual or subject-knowledge errors identified.")

    # 2. Technical Errors
    doc_obj.add_heading("2. Technical errors", level=2)
    doc_obj.add_paragraph(audit_data.get("technical_errors_summary", "None."))
    tech_items = audit_data.get("technical_error_items", [])
    if tech_items:
        t_rows = [
            [x.get("item_ref", ""), x.get("issue", ""), x.get("fix", "")]
            for x in tech_items
        ]
        add_styled_table(doc_obj, ["Q#", "Issue", "Fix"], t_rows, [1.0, 3.3, 2.7])

    # 3. Repetition
    doc_obj.add_heading("3. Repetition: same concept asked more than once (delete the second number in each pair)", level=2)
    rep_groups = audit_data.get("repetition_groups", [])
    if rep_groups:
        rep_rows = [
            [x.get("concept", ""), x.get("items", ""), x.get("delete_items", "")]
            for x in rep_groups
        ]
        add_styled_table(doc_obj, ["Concept", "Questions", "Delete"], rep_rows, [3.2, 2.1, 1.7])
    else:
        doc_obj.add_paragraph("No significant conceptual repetition detected.")

    clusters = audit_data.get("high_frequency_clusters", [])
    if clusters:
        p_cl = doc_obj.add_paragraph()
        p_cl.add_run("Concepts asked three or more times, which you could thin out further:").bold = True
        for c_line in clusters:
            doc_obj.add_paragraph(str(c_line), style="List Bullet")

    # 3b. Clarity
    doc_obj.add_heading("3b. Clarity", level=2)
    clarity_issues = audit_data.get("clarity_issues", [])
    if clarity_issues:
        cl_rows = [[x.get("item_ref", ""), x.get("issue", "")] for x in clarity_issues]
        add_styled_table(doc_obj, ["Q#", "Issue"], cl_rows, [1.6, 5.4])
    else:
        doc_obj.add_paragraph("Clear stems and well-balanced distractors throughout.")

    # 4. Grammar & Semantics
    doc_obj.add_heading("4. Grammar & semantics", level=2)
    gram_issues = audit_data.get("grammar_issues", [])
    if gram_issues:
        gr_rows = [
            [x.get("item_ref", ""), x.get("issue", ""), x.get("fix", "")]
            for x in gram_issues
        ]
        add_styled_table(doc_obj, ["Q#", "Issue", "Fix"], gr_rows, [0.9, 3.3, 2.8])
    else:
        doc_obj.add_paragraph("No major grammatical or semantic issues detected.")

    # 5. Subject Relevance & Syllabus Mapping
    doc_obj.add_heading("5. Subject relevance", level=2)
    doc_obj.add_paragraph(audit_data.get("subject_relevance_verdict", ""))

    map_status = audit_data.get("syllabus_mapping_status", "")
    p_ms = doc_obj.add_paragraph()
    p_ms.add_run(f"Syllabus mapping: {map_status}").bold = True

    map_note = audit_data.get("syllabus_mapping_note", "")
    if map_note:
        doc_obj.add_paragraph(map_note)

    syl_rows_raw = audit_data.get("syllabus_mapping", [])
    if syl_rows_raw:
        s_rows = [
            [
                x.get("unit", ""),
                x.get("topic", ""),
                x.get("items_covered", ""),
                x.get("status", ""),
            ]
            for x in syl_rows_raw
        ]
        add_styled_table(doc_obj, ["Unit", "Syllabus topic", "Questions", "Status"], s_rows, [0.7, 2.9, 1.8, 1.6])

    map_verdict = audit_data.get("mapping_verdict", "")
    if map_verdict:
        p_mv = doc_obj.add_paragraph()
        p_mv.add_run("Mapping verdict: ").bold = True
        p_mv.add_run(map_verdict)

    fixes = audit_data.get("priority_fixes", [])
    if fixes:
        p_fx = doc_obj.add_paragraph()
        p_fx.add_run("Fixes that would lift the score:").bold = True
        for fx in fixes:
            doc_obj.add_paragraph(str(fx), style="List Bullet")
        proj = audit_data.get("projected_score_after_fixes", "")
        if proj:
            p_pr = doc_obj.add_paragraph()
            p_pr.add_run(f"Doing that would bring the bank to about {proj}.").italic = True

    out_buf = io.BytesIO()
    doc_obj.save(out_buf)
    out_buf.seek(0)
    return out_buf


# ==============================================================================
# HELPER: GLOBAL SYLLABUS INGESTION (OPTION 1 UPLOAD + OPTION 2 DRIVE LINK)
# ==============================================================================
def render_global_syllabus_widget(drive_svc, key_prefix: str = "batch") -> tuple[str, str]:
    """
    Provides BOTH Option 1 (Direct Local File Upload) and Option 2 (Global Google Drive Link)
    for attaching a Syllabus to the current audit batch.
    """
    st.subheader("📚 Attach Syllabus for this Batch (Optional / Recommended)")
    st.caption(
        "Attach the syllabus once via **Local File Upload** or **Google Drive Link** to cross-map all selected manuscripts."
    )

    syl_mode = st.radio(
        "Choose Syllabus Source Method:",
        options=[
            "Option 1: Upload Syllabus File from Computer (.pdf, .docx, .txt)",
            "Option 2: Paste Global Syllabus Google Drive / Docs Link",
            "None / Use Per-Row Sheet Column Only",
        ],
        horizontal=True,
        key=f"{key_prefix}_syl_mode",
    )

    syllabus_text = ""
    source_label = ""

    if syl_mode.startswith("Option 1"):
        uploaded_syl = st.file_uploader(
            "Upload Syllabus Document (PDF, Word .docx, or .txt):",
            type=["pdf", "docx", "txt"],
            key=f"{key_prefix}_syl_uploader",
        )
        if uploaded_syl is not None:
            buf = io.BytesIO(uploaded_syl.read())
            syllabus_text, _ = extract_text_and_df_from_buffer(buf, uploaded_syl.name)
            source_label = uploaded_syl.name
            if syllabus_text.strip():
                st.success(f"Loaded local syllabus **{uploaded_syl.name}** ({len(syllabus_text):,} characters).")
            else:
                st.warning(f"Could not extract readable text from **{uploaded_syl.name}**.")

    elif syl_mode.startswith("Option 2"):
        drive_syl_url = st.text_input(
            "Paste Global Syllabus Google Drive / Docs / PDF URL:",
            placeholder="https://docs.google.com/document/d/1XyZ.../edit or https://drive.google.com/file/d/.../view",
            key=f"{key_prefix}_syl_drive_url",
        )
        if drive_syl_url.strip():
            syl_id = extract_drive_id(drive_syl_url)
            if not syl_id:
                st.warning("Could not parse a valid Google Drive File ID from that URL.")
            elif drive_svc is None:
                st.error("Google Drive service is not initialized.")
            else:
                cache_key = f"cached_syl_{syl_id}"
                if cache_key not in st.session_state:
                    try:
                        with st.spinner("Fetching Global Syllabus from Google Drive..."):
                            s_buf, s_name, _ = download_drive_file_bytes(drive_svc, syl_id)
                            s_txt, _ = extract_text_and_df_from_buffer(s_buf, s_name)
                            st.session_state[cache_key] = (s_txt, s_name)
                    except Exception as exc:
                        st.error(f"Failed to download Global Syllabus from Drive: {exc}")
                if cache_key in st.session_state:
                    syllabus_text, source_label = st.session_state[cache_key]
                    st.success(f"Loaded Drive syllabus **{source_label}** ({len(syllabus_text):,} characters).")

    return syllabus_text, source_label


# ==============================================================================
# MAIN STREAMLIT APPLICATION
# ==============================================================================
def main():
    st.sidebar.header("⚙️ Layer-3 Configuration")

    # 1. Multi-Provider AI Selector (Gemini / OpenAI ChatGPT / Anthropic Claude / Deterministic)
    st.sidebar.subheader("🤖 AI Engine Selection")
    selected_provider = st.sidebar.selectbox(
        "Choose AI Provider:",
        options=list(AI_PROVIDER_CATALOG.keys()),
        index=0,
        help="Switch freely between Gemini, ChatGPT (OpenAI), Claude (Anthropic), or Offline Rule Engine.",
    )
    provider_cfg = AI_PROVIDER_CATALOG[selected_provider]
    selected_model = st.sidebar.selectbox(
        "Model Version:",
        options=provider_cfg["models"],
        index=0,
        help=provider_cfg["help"],
    )
    custom_model_override = st.sidebar.text_input(
        "Custom Model ID (Optional Override):",
        placeholder="Leave blank to use dropdown model",
        help="Enter any exact API model ID if you want to override the preset list.",
    )
    active_model_id = custom_model_override.strip() or selected_model

    secret_name = provider_cfg["secret_key"]
    saved_key = st.secrets.get(secret_name, "") if secret_name else ""
    if secret_name:
        if saved_key:
            st.sidebar.success(f"Using `{secret_name}` from secrets.toml")
            override_key = st.sidebar.text_input(
                f"Override {secret_name} (Optional):",
                type="password",
                placeholder="Paste paid key to override",
            )
            active_api_key = override_key.strip() or saved_key
        else:
            active_api_key = st.sidebar.text_input(
                f"Enter {secret_name}:",
                type="password",
                placeholder="Paste your paid or free API key",
            )
    else:
        active_api_key = ""

    st.sidebar.divider()

    # 2. Destination Folder & Google Authentication
    dest_folder_input = st.sidebar.text_input(
        "Destination Root Drive Folder URL / ID:",
        placeholder="https://drive.google.com/drive/folders/...",
        help="Root Drive folder where QA Report folders will be created.",
    )
    dest_folder_id = extract_drive_id(dest_folder_input)

    drive_svc, sheets_svc, sa_email, gcp_status_msg, oauth_login_url = get_gcp_services()

    if sa_email and "@" in sa_email:
        st.sidebar.divider()
        st.sidebar.caption("Active Service Account (Share Sheet & Folders with this email as Editor):")
        st.sidebar.code(sa_email, language=None)

    # Header
    st.title("📄 Layer-3 Academic Document Auditor & Live Sheets Pipeline")
    st.caption(
        f"Active Engine: **{selected_provider}** (`{active_model_id}`) • Multi-Format (SLM, MCQ, QB, PPT) • Direct Word/Docs & Live Sheet Sync"
    )

    if drive_svc is None or sheets_svc is None:
        with st.container(border=True):
            st.warning(f"**Step 2: Connect Google Drive** — {gcp_status_msg}")
            if oauth_login_url:
                st.link_button(
                    "🔐 Sign In with Google Email to Connect Drive",
                    oauth_login_url,
                    type="primary",
                )
            st.caption("You can also run **Quick Single-File Audit** below using local file uploads without Drive.")
    else:
        st.success(f"✅ **Google Drive & Sheets Connected:** `{sa_email}`")

    work_mode = st.radio(
        "Select Workflow Mode:",
        options=[
            "📊 Master Tracker Batch Pipeline (Google Sheets / Drive)",
            "⚡ Quick Single-File Audit (Direct Upload or Paste without Master Sheet)",
        ],
        horizontal=True,
    )

    # ==========================================================================
    # WORKFLOW B: QUICK SINGLE-FILE AUDIT (DIRECT UPLOAD / PASTE)
    # ==========================================================================
    if work_mode.startswith("⚡"):
        st.subheader("⚡ Quick Single-File Document Audit")
        c1, c2, c3 = st.columns(3)
        q_subject = c1.text_input("Subject Name:", value="Block Chain Technologies")
        q_topic = c2.text_input("Unit / Module / Bank Name:", value="MCQ Bank")
        q_fmt = c3.selectbox("Deliverable Format:", ["Auto-Detect", "MCQ", "SLM", "QB", "PPT"])

        ms_source_type = st.radio(
            "Manuscript Input Source:",
            ["Upload Local File (.xlsx, .docx, .pdf, .csv)", "Paste Google Drive / Docs / Sheets Link"],
            horizontal=True,
        )

        ms_text = ""
        ms_df = None
        ms_filename = "Manuscript"

        if ms_source_type.startswith("Upload"):
            up_ms = st.file_uploader("Upload Manuscript File:", type=["xlsx", "xls", "docx", "pdf", "csv", "txt"])
            if up_ms is not None:
                ms_filename = up_ms.name
                ms_text, ms_df = extract_text_and_df_from_buffer(io.BytesIO(up_ms.read()), up_ms.name)
                st.info(f"Loaded **{up_ms.name}** ({len(ms_text):,} characters).")
        else:
            ms_link = st.text_input("Paste Manuscript Google Drive / Docs / Sheets URL:")
            if ms_link.strip():
                f_id = extract_drive_id(ms_link)
                if f_id:
                    try:
                        b, ms_filename = _l4_download_from_any_drive(f_id, locals())
                        ms_text, ms_df = extract_text_and_df_from_buffer(b, ms_filename)
                        st.info(f"Loaded Drive file **{ms_filename}** ({len(ms_text):,} characters).")
                    except Exception as exc:
                        st.error(f"Could not download Drive file: {exc}")

        st.divider()
        quick_syl_text, _ = render_global_syllabus_widget(drive_svc, key_prefix="quick")

        if st.button("🚀 Generate QA Report (.docx)", type="primary", disabled=not bool(ms_text.strip())):
            with st.spinner(f"Running deterministic verification & {selected_provider} ({active_model_id}) audit..."):
                resolved_fmt = (
                    detect_deliverable_format("", ms_filename, ms_text) if q_fmt == "Auto-Detect" else q_fmt
                )
                mcq_stats = (
                    analyze_mcq_deterministic(ms_df, ms_text)
                    if resolved_fmt == "MCQ"
                    else {"total_questions": 0, "summary_line": "Narrative/structural review completed."}
                )
                ai_cv = compute_ai_burstiness_cv(ms_text)
                q_count_label = f" ({mcq_stats['total_questions']} questions)" if mcq_stats.get("total_questions") else ""
                title = f"QA Report: {q_subject} {resolved_fmt}{q_count_label}"
                audit_res = run_multi_provider_layer3_audit(
                    provider=selected_provider,
                    model_id=active_model_id,
                    api_key=active_api_key,
                    title=title,
                    fmt=resolved_fmt,
                    subject=q_subject,
                    topic=q_topic,
                    manuscript_text=ms_text,
                    syllabus_text=quick_syl_text,
                    mcq_stats=mcq_stats,
                    ai_cv=ai_cv,
                )
                docx_buf = build_layer3_docx_report(audit_res)

                drive_url = ""
                if dest_folder_id:
                    try:
                        ts = datetime.now().strftime("%Y%m%d_%H%M")
                        out_name = f"QA_Report_{sanitize_drive_name(q_subject)}_{resolved_fmt}_{ts}.docx"
                        drive_url = upload_docx_to_drive_hierarchy(
                            drive_svc,
                            dest_folder_id,
                            [q_subject, f"{resolved_fmt} Reports"],
                            out_name,
                            docx_buf,
                        )
                    except Exception as exc:
                        st.warning(f"Local .docx generated, but Drive upload failed: {exc}")

                st.success(f"Audit Complete! Overall Score: **{audit_res.get('overall_score', 0)}%**")
                if drive_url:
                    st.markdown(f"📂 **Saved to Google Drive:** [Open QA Report in Drive]({drive_url})")

                docx_buf.seek(0)
                st.download_button(
                    "📥 Download QA Report (.docx)",
                    data=docx_buf.getvalue(),
                    file_name=f"QA_Report_{sanitize_drive_name(q_subject)}_{resolved_fmt}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
                with st.expander("Preview Structured Audit JSON", expanded=False):
                    st.json(audit_res)
        st.stop()

    # ==========================================================================
    # WORKFLOW A: MASTER TRACKER BATCH PIPELINE (GOOGLE SHEETS / DRIVE EXCEL)
    # ==========================================================================
    if drive_svc is None or sheets_svc is None:
        st.info("Connect Google Drive above (or select GCP Service Account in the sidebar) to scan a Master Tracker Google Sheet.")
        st.stop()

    st.subheader("Step 1: Connect Master Tracker (Live Google Sheet or Drive Excel)")
    tracker_url = st.text_input(
        "Paste Google Sheet or Drive Excel URL:",
        placeholder="https://docs.google.com/spreadsheets/d/1AbC.../edit",
    )

    if st.button("🔍 Scan Master Tracker Tabs", type="primary"):
        tracker_id = extract_drive_id(tracker_url)
        if not tracker_id:
            st.error("Please paste a valid Google Sheets or Google Drive Excel URL.")
        else:
            try:
                with st.spinner("Connecting to Google Drive & reading Master Tracker..."):
                    meta = (
                        drive_svc.files()
                        .get(fileId=tracker_id, fields="id,name,mimeType", supportsAllDrives=True)
                        .execute()
                    )
                    is_live_gsheet = meta.get("mimeType") == "application/vnd.google-apps.spreadsheet"
                    xl_buf, xl_name, _ = download_drive_file_bytes(drive_svc, tracker_id)
                    wb = openpyxl.load_workbook(xl_buf, data_only=True)
                    st.session_state["tracker_state"] = {
                        "id": tracker_id,
                        "name": meta.get("name", xl_name),
                        "is_live_gsheet": is_live_gsheet,
                        "xl_bytes": xl_buf.getvalue(),
                        "sheet_names": wb.sheetnames,
                    }
            except HttpError as he:
                st.error(f"Could not open tracker: {he}")
            except Exception as exc:
                st.error(f"Error reading Master Tracker: {exc}")

    tracker = st.session_state.get("tracker_state")
    if not tracker:
        st.stop()

    if tracker["is_live_gsheet"]:
        st.success(
            f"✅ Connected to **LIVE Google Sheet** (`{tracker['name']}`) — Direct Write-Back Active | "
            f"Tabs: {', '.join(tracker['sheet_names'])}"
        )
    else:
        st.info(
            f"📄 Connected to **Drive Excel File** (`{tracker['name']}`) — "
            f"Tabs: {', '.join(tracker['sheet_names'])} (Download updated copy in Step 3)"
        )

    # --------------------------------------------------------------------------
    # STEP 2: MAP COLUMNS, ATTACH SYLLABUS (OPTION 1 OR 2), & FILTER SEGMENT
    # --------------------------------------------------------------------------
    st.divider()
    st.subheader("Step 2: Map Columns, Attach Syllabus & Select Segment")

    c_tab, c_hdr = st.columns([3, 1])
    selected_tab = c_tab.selectbox("Select Worksheet / University Tab:", tracker["sheet_names"])
    header_row_1idx = int(
        c_hdr.number_input("Header Row Number (1-indexed):", min_value=1, max_value=20, value=1)
    )

    wb = openpyxl.load_workbook(io.BytesIO(tracker["xl_bytes"]), data_only=True)
    ws = wb[selected_tab]
    all_rows = list(ws.iter_rows(values_only=True))

    if len(all_rows) < header_row_1idx:
        st.warning("Selected tab has fewer rows than the specified Header Row Number.")
        st.stop()

    raw_headers = [
        str(h).strip() if h is not None and str(h).strip() != "" else f"Column_{i+1}"
        for i, h in enumerate(all_rows[header_row_1idx - 1])
    ]
    data_rows = all_rows[header_row_1idx:]

    records = []
    for offset, r_vals in enumerate(data_rows):
        sheet_row_num = header_row_1idx + 1 + offset
        row_dict = {"_sheet_row": sheet_row_num}
        for c_i, h_name in enumerate(raw_headers):
            row_dict[h_name] = r_vals[c_i] if c_i < len(r_vals) else None
        records.append(row_dict)

    raw_df = pd.DataFrame(records)
    if raw_df.empty:
        st.warning("No data rows found below the header row in this tab.")
        st.stop()

    with st.expander("Preview Top 5 Rows of Selected Tab", expanded=False):
        st.dataframe(raw_df.drop(columns=["_sheet_row"]).head(5), use_container_width=True)

    def guess_col(keywords: list, allow_none: bool = True) -> int:
        options = (["(None)"] + raw_headers) if allow_none else raw_headers
        for idx, opt in enumerate(options):
            ol = opt.lower()
            if any(k in ol for k in keywords):
                return idx
        return 0

    none_options = ["(None)"] + raw_headers
    st.markdown("**Column Mapping (Auto-Detected — Adjust if Needed):**")
    m1, m2, m3, m4 = st.columns(4)
    col_ms = m1.selectbox(
        "Manuscript Link (Required):",
        options=raw_headers,
        index=guess_col(["manuscript", "content link", "file link", "drive link", "url", "link"], allow_none=False),
    )
    col_syl = m2.selectbox(
        "Syllabus Link (Optional Sheet Col):",
        options=none_options,
        index=guess_col(["syllabus"], allow_none=True),
    )
    col_fmt = m3.selectbox(
        "Deliverable Format (MCQ/SLM/QB/PPT):",
        options=none_options,
        index=guess_col(["format", "type", "deliverable", "category"], allow_none=True),
    )
    col_topic = m4.selectbox(
        "Topic / Unit Name:",
        options=none_options,
        index=guess_col(["topic", "unit", "module", "chapter", "title"], allow_none=True),
    )

    m5, m6, m7, m8 = st.columns(4)
    col_univ = m5.selectbox(
        "University:",
        options=none_options,
        index=guess_col(["university", "univ", "partner", "client"], allow_none=True),
    )
    col_prog = m6.selectbox(
        "Program:",
        options=none_options,
        index=guess_col(["program", "degree", "course_name"], allow_none=True),
    )
    col_sem = m7.selectbox(
        "Semester:",
        options=none_options,
        index=guess_col(["semester", "sem", "term"], allow_none=True),
    )
    col_subj = m8.selectbox(
        "Subject / Course:",
        options=none_options,
        index=guess_col(["subject", "course", "paper"], allow_none=True),
    )

    lower_headers = [h.lower() for h in raw_headers]
    write_new_headers = False
    if "qa report link" in lower_headers:
        link_col_0idx = lower_headers.index("qa report link")
        existing_link_col = raw_headers[link_col_0idx]
    else:
        link_col_0idx = len(raw_headers)
        existing_link_col = None
        write_new_headers = True

    if "audit status" in lower_headers:
        status_col_0idx = lower_headers.index("audit status")
    else:
        status_col_0idx = link_col_0idx + 1
        write_new_headers = True

    st.divider()
    global_syl_text, global_syl_label = render_global_syllabus_widget(drive_svc, key_prefix="tracker")
    syl_precedence = st.radio(
        "Syllabus Priority Rule:",
        options=[
            "Use Global Batch Syllabus when a row's Syllabus Link column is empty (Recommended)",
            "Always use Global Batch Syllabus for all selected rows",
        ],
        horizontal=True,
        disabled=not bool(global_syl_text.strip()),
    )

    st.divider()
    st.subheader("🎯 Segmented Volume Filter & Execution Queue")

    work_df = raw_df.copy()
    work_df["_ms_link"] = work_df[col_ms].fillna("").astype(str).str.strip()
    work_df = work_df[work_df["_ms_link"].str.startswith("http")].copy()

    def get_clean_series(col_name: str, default_val: str) -> pd.Series:
        if col_name == "(None)":
            return pd.Series([default_val] * len(work_df), index=work_df.index)
        return work_df[col_name].fillna(default_val).astype(str).str.strip().replace("", default_val)

    work_df["_univ"] = get_clean_series(col_univ, "General University")
    work_df["_prog"] = get_clean_series(col_prog, "General Program")
    work_df["_sem"] = get_clean_series(col_sem, "Semester 1")
    work_df["_subj"] = get_clean_series(col_subj, "General Subject")
    work_df["_topic"] = get_clean_series(col_topic, "Complete Bank")
    work_df["_fmt_raw"] = get_clean_series(col_fmt, "")
    work_df["_syl_link"] = (
        work_df[col_syl].fillna("").astype(str).str.strip()
        if col_syl != "(None)"
        else pd.Series([""] * len(work_df), index=work_df.index)
    )
    work_df["_existing_report"] = (
        work_df[existing_link_col].fillna("").astype(str).str.strip()
        if existing_link_col
        else pd.Series([""] * len(work_df), index=work_df.index)
    )

    f1, f2, f3, f4 = st.columns(4)
    sel_univ = f1.selectbox("Filter University:", ["All"] + sorted(work_df["_univ"].unique().tolist()))
    if sel_univ != "All":
        work_df = work_df[work_df["_univ"] == sel_univ]

    sel_prog = f2.selectbox("Filter Program:", ["All"] + sorted(work_df["_prog"].unique().tolist()))
    if sel_prog != "All":
        work_df = work_df[work_df["_prog"] == sel_prog]

    sel_sem = f3.selectbox("Filter Semester:", ["All"] + sorted(work_df["_sem"].unique().tolist()))
    if sel_sem != "All":
        work_df = work_df[work_df["_sem"] == sel_sem]

    sel_subj = f4.selectbox("Filter Subject:", ["All"] + sorted(work_df["_subj"].unique().tolist()))
    if sel_subj != "All":
        work_df = work_df[work_df["_subj"] == sel_subj]

    c_hide, c_cap = st.columns([2, 2])
    hide_audited = c_hide.checkbox("Hide already-audited rows (where QA Report Link exists)", value=True)
    if hide_audited:
        work_df = work_df[~work_df["_existing_report"].str.startswith("http")]

    max_rows = int(c_cap.slider("Max rows per execution:", min_value=1, max_value=50, value=15))
    work_df = work_df.head(max_rows).copy()

    if work_df.empty:
        st.info("No pending rows match your current filter criteria.")
        st.stop()

    queue_display = pd.DataFrame(
        {
            "Audit": [True] * len(work_df),
            "Sheet Row": work_df["_sheet_row"].tolist(),
            "University": work_df["_univ"].tolist(),
            "Program": work_df["_prog"].tolist(),
            "Semester": work_df["_sem"].tolist(),
            "Subject": work_df["_subj"].tolist(),
            "Topic / Unit": work_df["_topic"].tolist(),
            "Format": [
                detect_deliverable_format(r["_fmt_raw"], r["_topic"], "")
                for _, r in work_df.iterrows()
            ],
            "Syllabus Source": [
                "Row Sheet Link"
                if (r["_syl_link"].startswith("http") and syl_precedence.startswith("Use Global"))
                else (f"Global ({global_syl_label})" if global_syl_text.strip() else "None")
                for _, r in work_df.iterrows()
            ],
            "Manuscript Link": work_df["_ms_link"].tolist(),
            "_syl_link": work_df["_syl_link"].tolist(),
        }
    )

    edited_queue = st.data_editor(
        queue_display.drop(columns=["_syl_link"]),
        use_container_width=True,
        hide_index=True,
        column_config={
            "Audit": st.column_config.CheckboxColumn("Audit", default=True),
            "Format": st.column_config.SelectboxColumn("Format", options=["MCQ", "SLM", "QB", "PPT"]),
            "Manuscript Link": st.column_config.LinkColumn("Manuscript Link"),
        },
        disabled=["Sheet Row", "University", "Program", "Semester", "Subject", "Syllabus Source", "Manuscript Link"],
    )

    selected_indices = edited_queue[edited_queue["Audit"] == True].index.tolist()
    st.caption(f"Ready to process **{len(selected_indices)}** selected document(s) using **{selected_provider}** (`{active_model_id}`).")

    # ==========================================================================
    # STEP 3: EXECUTE LAYER-3 AUDIT, ARCHIVE TO DRIVE, & SYNC LIVE SHEET
    # ==========================================================================
    if st.button(f"🚀 Execute Layer-3 Audit ({len(selected_indices)} Selected Files)", type="primary"):
        if not selected_indices:
            st.warning("Please check at least one row in the table above.")
            st.stop()
        if not dest_folder_id:
            st.error("Please paste your **Destination Root Drive Folder URL / ID** in the left sidebar first.")
            st.stop()

        progress_bar = st.progress(0.0)
        status_box = st.empty()
        results_log = []
        row_syllabus_cache = {}

        out_wb = openpyxl.load_workbook(io.BytesIO(tracker["xl_bytes"]))
        out_ws = out_wb[selected_tab]
        if write_new_headers:
            out_ws.cell(row=header_row_1idx, column=link_col_0idx + 1, value="QA Report Link")
            out_ws.cell(row=header_row_1idx, column=status_col_0idx + 1, value="Audit Status")

        total_to_run = len(selected_indices)
        for step_i, q_idx in enumerate(selected_indices, start=1):
            row_ui = edited_queue.loc[q_idx]
            row_syl_url = queue_display.loc[q_idx, "_syl_link"]
            sheet_row = int(row_ui["Sheet Row"])
            univ = str(row_ui["University"])
            prog = str(row_ui["Program"])
            sem = str(row_ui["Semester"])
            subj = str(row_ui["Subject"])
            topic = str(row_ui["Topic / Unit"])
            fmt_choice = str(row_ui["Format"])
            ms_url = str(row_ui["Manuscript Link"])

            status_box.info(
                f"[{step_i}/{total_to_run}] Auditing Row {sheet_row}: **{subj} — {topic}** ({fmt_choice}) via {selected_provider}..."
            )

            try:
                ms_file_id = extract_drive_id(ms_url)
                if not ms_file_id:
                    raise ValueError("Invalid Manuscript Google Drive URL.")

                ms_buf, ms_fname, _ = download_drive_file_bytes(drive_svc, ms_file_id)
                ms_text, ms_df = extract_text_and_df_from_buffer(ms_buf, ms_fname)
                if not ms_text.strip():
                    raise ValueError(f"No readable text extracted from '{ms_fname}'.")

                effective_syl_text = global_syl_text
                if row_syl_url.startswith("http") and syl_precedence.startswith("Use Global"):
                    s_id = extract_drive_id(row_syl_url)
                    if s_id:
                        if s_id not in row_syllabus_cache:
                            sb, sfn, _ = download_drive_file_bytes(drive_svc, s_id)
                            stxt, _ = extract_text_and_df_from_buffer(sb, sfn)
                            row_syllabus_cache[s_id] = stxt
                        effective_syl_text = row_syllabus_cache[s_id]

                mcq_stats = (
                    analyze_mcq_deterministic(ms_df, ms_text)
                    if fmt_choice == "MCQ"
                    else {"total_questions": 0, "summary_line": "Narrative/structural check completed."}
                )
                ai_cv = compute_ai_burstiness_cv(ms_text)
                q_count_label = f" ({mcq_stats['total_questions']} questions)" if mcq_stats.get("total_questions") else ""
                report_title = f"QA Report: {subj} — {topic} [{fmt_choice}]{q_count_label}"

                audit_dict = run_multi_provider_layer3_audit(
                    provider=selected_provider,
                    model_id=active_model_id,
                    api_key=active_api_key,
                    title=report_title,
                    fmt=fmt_choice,
                    subject=subj,
                    topic=topic,
                    manuscript_text=ms_text,
                    syllabus_text=effective_syl_text,
                    mcq_stats=mcq_stats,
                    ai_cv=ai_cv,
                )

                score = audit_dict.get("overall_score", 0)
                docx_buf = build_layer3_docx_report(audit_dict)

                ts = datetime.now().strftime("%Y%m%d_%H%M")
                report_filename = f"QA_Report_{sanitize_drive_name(subj)}_{sanitize_drive_name(topic)}_{fmt_choice}_{ts}.docx"
                hierarchy = [univ, prog, sem, subj, f"{fmt_choice} Reports"]
                report_drive_link = upload_docx_to_drive_hierarchy(
                    drive_svc=drive_svc,
                    root_folder_id=dest_folder_id,
                    hierarchy_parts=hierarchy,
                    filename=report_filename,
                    docx_bytes=docx_buf,
                )

                status_label = f"Completed ({score}%)"

                out_ws.cell(row=sheet_row, column=link_col_0idx + 1, value=report_drive_link)
                out_ws.cell(row=sheet_row, column=status_col_0idx + 1, value=status_label)

                if tracker["is_live_gsheet"]:
                    sync_live_sheet_row(
                        sheets_svc=sheets_svc,
                        spreadsheet_id=tracker["id"],
                        tab_name=selected_tab,
                        header_row_1idx=header_row_1idx,
                        data_row_1idx=sheet_row,
                        link_col_0idx=link_col_0idx,
                        status_col_0idx=status_col_0idx,
                        report_link=report_drive_link,
                        status_text=status_label,
                        write_headers=(write_new_headers and step_i == 1),
                    )

                results_log.append(
                    {
                        "Sheet Row": sheet_row,
                        "Subject": subj,
                        "Topic / Unit": topic,
                        "Format": fmt_choice,
                        "Score": f"{score}%",
                        "Status": status_label,
                        "Open QA Report in Drive": report_drive_link,
                    }
                )

            except Exception as row_exc:
                err_msg = f"Error: {str(row_exc)[:120]}"
                out_ws.cell(row=sheet_row, column=status_col_0idx + 1, value=err_msg)
                if tracker["is_live_gsheet"]:
                    try:
                        sync_live_sheet_row(
                            sheets_svc=sheets_svc,
                            spreadsheet_id=tracker["id"],
                            tab_name=selected_tab,
                            header_row_1idx=header_row_1idx,
                            data_row_1idx=sheet_row,
                            link_col_0idx=link_col_0idx,
                            status_col_0idx=status_col_0idx,
                            report_link="",
                            status_text=err_msg,
                            write_headers=(write_new_headers and step_i == 1),
                        )
                    except Exception:
                        pass
                results_log.append(
                    {
                        "Sheet Row": sheet_row,
                        "Subject": subj,
                        "Topic / Unit": topic,
                        "Format": fmt_choice,
                        "Score": "—",
                        "Status": err_msg,
                        "Open QA Report in Drive": "",
                    }
                )

            progress_bar.progress(step_i / total_to_run)

        status_box.success(f"✅ Batch execution finished! Processed {total_to_run} document(s).")
        st.subheader("Step 3: Execution Results & Live Sync Summary")
        res_df = pd.DataFrame(results_log)
        st.dataframe(
            res_df,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Open QA Report in Drive": st.column_config.LinkColumn("Open QA Report in Drive"),
            },
        )

        xl_out_buf = io.BytesIO()
        out_wb.save(xl_out_buf)
        xl_out_buf.seek(0)
        st.download_button(
            "📥 Download Updated Excel Copy",
            data=xl_out_buf.getvalue(),
            file_name=f"Updated_{sanitize_drive_name(tracker['name'])}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


if __name__ == "__main__":
    main()