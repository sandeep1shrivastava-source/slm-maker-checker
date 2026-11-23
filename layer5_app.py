import hmac
import io
import json
import re
import secrets as py_secrets
import smtplib
import time
import urllib.parse
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import docx
import openpyxl
import pandas as pd
import pypdf
import requests
import streamlit as st
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from pptx import Presentation

from layer5_engine import (
    analyze_document_topology,
    classify_deliverable_category,
    compile_approved_enriched_docx,
    compile_approved_fixit_patch_docx,
    compile_complete_slm_book_docx,
    compile_subject_cumulative_dossier,
    compute_prose_cv_3dec,
    detect_low_cv_windows,
    enrich_pptx_with_native_animations,
    generate_track_a_fixit_items,
    generate_track_b_enrichment_items,
    load_university_profiles,
    parse_and_verify_audit_report,
    save_university_profile,
    sync_layer5_master_tracker_row,
    upload_layer5_artifact_to_drive,
    _extract_prose_paragraphs,
    _generate_brand_crest_png,
)

st.set_page_config(
    page_title="Layer-5 Content Enrichment, Book Compiler & Subject QA Workbench",
    page_icon="✨",
    layout="wide",
    initial_sidebar_state="expanded",
)

GCP_SCOPES = [
    "[https://www.googleapis.com/auth/drive](https://www.googleapis.com/auth/drive)",
    "[https://www.googleapis.com/auth/spreadsheets](https://www.googleapis.com/auth/spreadsheets)",
]

AI_PROVIDERS = {
    "Google Gemini": ["gemini-2.5-flash", "gemini-2.5-flash-lite", "gemini-2.5-pro", "gemini-3.8-flash"],
    "OpenAI (ChatGPT)": ["gpt-4o-mini", "gpt-4o", "gpt-6.1-terra"],
    "Anthropic (Claude)": ["claude-3-5-sonnet-latest", "claude-3-5-haiku-latest", "claude-sonnet-5-5"],
    "Deterministic Offline Scaffolding (No AI Key)": ["deterministic-offline-template"],
}


def _parse_email_list(secret_key: str) -> set[str]:
    raw = st.secrets.get(secret_key, "")
    if isinstance(raw, (list, tuple, set)):
        return {str(e).strip().lower() for e in raw if str(e).strip()}
    return {e.strip().lower() for e in re.split(r"[,;\n\r]+", str(raw)) if e.strip()}


def _is_layer5_authorized(email: str) -> bool:
    clean = (email or "").strip().lower()
    if not clean:
        return False
    blocked = _parse_email_list("LAYER4_ONLY_EMAILS")
    return clean not in blocked


def _send_otp_email(recipient_email: str, otp_code: str) -> tuple[bool, str]:
    smtp_host = str(st.secrets.get("SMTP_HOST", "smtp.gmail.com")).strip()
    smtp_port = int(st.secrets.get("SMTP_PORT", 587))
    smtp_user = str(st.secrets.get("SMTP_USER") or st.secrets.get("EMAIL_USER") or "").strip()
    smtp_pass = str(st.secrets.get("SMTP_PASS") or st.secrets.get("SMTP_PASSWORD") or st.secrets.get("APP_PASSWORD") or "").strip()
    if not smtp_user or not smtp_pass:
        return False, "SMTP credentials missing in st.secrets."

    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"Layer-5 QA Verification Code: {otp_code}"
    msg["From"] = f"Academic QA Platform (Layer-5) <{smtp_user}>"
    msg["To"] = recipient_email
    msg.attach(MIMEText(f"Your Layer-5 sign-in verification code is: {otp_code} (expires in 15 minutes).", "plain", "utf-8"))

    try:
        with smtplib.SMTP(smtp_host, smtp_port, timeout=20) as srv:
            srv.ehlo()
            srv.starttls()
            srv.ehlo()
            srv.login(smtp_user, smtp_pass)
            srv.sendmail(smtp_user, [recipient_email], msg.as_string())
        return True, "Verification code sent."
    except Exception as exc:
        return False, f"SMTP error: {exc}"


def require_layer5_auth() -> bool:
    l5_specific = _parse_email_list("LAYER5_APPROVED_EMAILS") | _parse_email_list("DUAL_LAYER_EMAILS")
    l4_shared = (
        _parse_email_list("APPROVED_EMAILS")
        | _parse_email_list("ALLOWED_EMAILS")
        | _parse_email_list("AUTHORIZED_EMAILS")
    )
    blocked_l4_only = _parse_email_list("LAYER4_ONLY_EMAILS")

    # Enforce strict Layer-5 allowlist when LAYER5_APPROVED_EMAILS or DUAL_LAYER_EMAILS is set;
    # otherwise fall back to shared APPROVED_EMAILS minus explicit LAYER4_ONLY_EMAILS.
    approved_emails = (l5_specific if l5_specific else l4_shared) - blocked_l4_only

    l5_pass = str(st.secrets.get("LAYER5_TEAM_PASSCODE", "")).strip()
    dual_pass = str(st.secrets.get("DUAL_LAYER_PASSCODE", "")).strip()
    shared_pass = str(st.secrets.get("TEAM_PASSCODE", "")).strip()
    valid_passcodes = {p for p in ((l5_pass, dual_pass) if (l5_pass or dual_pass) else (shared_pass,)) if p}

    has_any_auth_config = bool(l5_specific or l4_shared or blocked_l4_only or valid_passcodes)
    if not has_any_auth_config:
        st.session_state.setdefault("l5_auth_user", "local-reviewer@institution.edu")
        return True

    if st.session_state.get("l5_auth_ok"):
        user = st.session_state.get("l5_auth_user", "Reviewer")
        st.sidebar.success(f"👤 Signed in: **{user}**")
        if st.sidebar.button("🔒 Sign Out", use_container_width=True):
            for k in ("l5_auth_ok", "l5_auth_user", "l5_otp_code", "l5_otp_email", "l5_otp_ts"):
                st.session_state.pop(k, None)
            st.rerun()
        return True

    st.markdown("## ✨ Layer-5 Content Enrichment & Book Compiler — Sign-In")
    tabs = st.tabs(["📧 Email OTP Login", "🔑 Team Passcode"])

    with tabs[0]:
        with st.form("l5_send_otp"):
            email_in = st.text_input("Institutional Email ID:").strip().lower()
            if st.form_submit_button("📨 Send 6-Digit OTP", type="primary"):
                if not email_in or email_in not in approved_emails:
                    st.error("Access denied: Email ID is not authorized for Layer-5.")
                else:
                    code = f"{py_secrets.randbelow(900000) + 100000}"
                    ok, msg = _send_otp_email(email_in, code)
                    if ok:
                        st.session_state["l5_otp_code"] = code
                        st.session_state["l5_otp_email"] = email_in
                        st.session_state["l5_otp_ts"] = time.time()
                        st.success(f"OTP sent to **{email_in}**.")
                    else:
                        st.error(msg)

        if st.session_state.get("l5_otp_email"):
            with st.form("l5_verify_otp"):
                otp_in = st.text_input("Enter 6-Digit OTP:", max_chars=6).strip()
                if st.form_submit_button("✅ Verify & Unlock Layer-5", type="primary"):
                    dev_bypass = str(st.secrets.get("DEV_BYPASS_OTP", "")).strip()
                    if time.time() - st.session_state.get("l5_otp_ts", 0) > 900:
                        st.error("OTP expired.")
                    elif hmac.compare_digest(otp_in, str(st.session_state.get("l5_otp_code", ""))) or (
                        dev_bypass and hmac.compare_digest(otp_in, dev_bypass)
                    ):
                        st.session_state["l5_auth_ok"] = True
                        st.session_state["l5_auth_user"] = st.session_state["l5_otp_email"]
                        st.rerun()
                    else:
                        st.error("Invalid OTP.")

    with tabs[1]:
        if valid_passcodes:
            with st.form("l5_pass_form"):
                p_in = st.text_input("Layer-5 / Dual-Access Passcode:", type="password")
                if st.form_submit_button("Unlock", type="primary"):
                    cleaned_p = p_in.strip()
                    if any(hmac.compare_digest(cleaned_p, vp) for vp in valid_passcodes):
                        st.session_state["l5_auth_ok"] = True
                        st.session_state["l5_auth_user"] = "Passcode Reviewer"
                        st.rerun()
                    else:
                        st.error("Incorrect passcode for Layer-5.")
        else:
            st.info("Passcode login is disabled; sign in using your authorized Institutional Email OTP.")
    return False


if not require_layer5_auth():
    st.stop()


def get_gcp_services():
    raw = dict(st.secrets["gcp_user_oauth"]) if "gcp_user_oauth" in st.secrets else {}
    secret_oauth = dict(raw.get("web") or raw.get("installed") or raw)
    creds = None

    if st.session_state.get("l5_gcp_creds_json"):
        creds = Credentials.from_authorized_user_info(json.loads(st.session_state["l5_gcp_creds_json"]), GCP_SCOPES)
    elif secret_oauth.get("refresh_token"):
        creds = Credentials.from_authorized_user_info(secret_oauth, GCP_SCOPES)

    auth_code = str(st.query_params.get("code", "")).strip()
    redirect_uri = str(secret_oauth.get("l5_redirect_uri") or secret_oauth.get("redirect_uri", "")).strip()

    if (creds is None or not creds.valid) and auth_code and secret_oauth.get("client_id"):
        resp = requests.post(
            "[https://oauth2.googleapis.com/token](https://oauth2.googleapis.com/token)",
            data={
                "code": auth_code,
                "client_id": str(secret_oauth["client_id"]).strip(),
                "client_secret": str(secret_oauth["client_secret"]).strip(),
                "redirect_uri": redirect_uri,
                "grant_type": "authorization_code",
            },
            timeout=20,
        )
        tok = resp.json()
        if resp.status_code == 200 and tok.get("access_token"):
            payload = {
                "token": tok["access_token"],
                "refresh_token": tok.get("refresh_token", secret_oauth.get("refresh_token", "")),
                "token_uri": "[https://oauth2.googleapis.com/token](https://oauth2.googleapis.com/token)",
                "client_id": str(secret_oauth["client_id"]).strip(),
                "client_secret": str(secret_oauth["client_secret"]).strip(),
                "scopes": GCP_SCOPES,
            }
            st.session_state["l5_gcp_creds_json"] = json.dumps(payload)
            creds = Credentials.from_authorized_user_info(payload, GCP_SCOPES)
        for qp in ("code", "state", "scope", "authuser", "prompt", "hd"):
            st.query_params.pop(qp, None)

    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
        st.session_state["l5_gcp_creds_json"] = creds.to_json()

    if not creds or not creds.valid:
        if not secret_oauth.get("client_id"):
            return None, None, "Missing [gcp_user_oauth] in st.secrets.", ""
        params = {
            "client_id": str(secret_oauth["client_id"]).strip(),
            "redirect_uri": redirect_uri,
            "response_type": "code",
            "scope": " ".join(GCP_SCOPES),
            "access_type": "offline",
            "prompt": "consent",
        }
        return None, None, "Connect Google Drive & Master Tracker:", "[https://accounts.google.com/o/oauth2/v2/auth](https://accounts.google.com/o/oauth2/v2/auth)?" + urllib.parse.urlencode(params)

    return (
        build("drive", "v3", credentials=creds, cache_discovery=False),
        build("sheets", "v4", credentials=creds, cache_discovery=False),
        "",
        "",
    )


def extract_drive_id(url_or_id: str) -> str:
    val = str(url_or_id or "").strip()
    for pat in (r"/folders/([a-zA-Z0-9_-]+)", r"/d/([a-zA-Z0-9_-]+)", r"[?&]id=([a-zA-Z0-9_-]+)"):
        m = re.search(pat, val)
        if m:
            return m.group(1)
    return val if re.fullmatch(r"[a-zA-Z0-9_-]{15,}", val) else ""


def download_drive_file_readonly(drive_svc, file_id: str) -> tuple[bytes, str]:
    meta = drive_svc.files().get(fileId=file_id, fields="id,name,mimeType", supportsAllDrives=True).execute()
    name = meta.get("name", f"manuscript_{file_id}")
    mime = meta.get("mimeType", "")
    buf = io.BytesIO()

    if mime == "application/vnd.google-apps.document":
        req = drive_svc.files().export_media(fileId=file_id, mimeType="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        if not name.lower().endswith(".docx"):
            name += ".docx"
    elif mime == "application/vnd.google-apps.presentation":
        req = drive_svc.files().export_media(fileId=file_id, mimeType="application/vnd.openxmlformats-officedocument.presentationml.presentation")
        if not name.lower().endswith(".pptx"):
            name += ".pptx"
    else:
        req = drive_svc.files().get_media(fileId=file_id, supportsAllDrives=True)

    downloader = MediaIoBaseDownload(buf, req)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    return buf.getvalue(), name


def extract_text_from_bytes(file_bytes: bytes, filename: str) -> str:
    fn = filename.lower()
    if fn.endswith(".docx"):
        d = docx.Document(io.BytesIO(file_bytes))
        return "\n".join([p.text.strip() for p in d.paragraphs if p.text.strip()])
    if fn.endswith(".pptx"):
        prs = Presentation(io.BytesIO(file_bytes))
        lines = []
        for idx, s in enumerate(prs.slides, start=1):
            lines.append(f"--- Slide {idx} ---")
            for shp in s.shapes:
                if shp.has_text_frame and shp.text_frame.text.strip():
                    lines.append(shp.text_frame.text.strip())
        return "\n".join(lines)
    if fn.endswith(".pdf"):
        reader = pypdf.PdfReader(io.BytesIO(file_bytes))
        return "\n\n".join([p.extract_text() or "" for p in reader.pages])
    return file_bytes.decode("utf-8", errors="ignore")


def render_sidebar_brand_and_config() -> dict:
    st.sidebar.header("🏛️ Dynamic University Brand Profile")
    profiles = load_university_profiles()
    profile_keys = list(profiles.keys()) + ["+ Create / Customize University"]
    sel_key = st.sidebar.selectbox("Active University Profile:", profile_keys, index=0)

    base_prof = profiles.get(sel_key, {
        "university_id": "CUSTOM_UNI",
        "university_name": "CUSTOM UNIVERSITY",
        "tagline": "ADVANCING KNOWLEDGE & INNOVATION",
        "school_name": "Centre for Distance and Online Education",
        "primary_hex": "1E3A8A",
        "accent_hex": "D97706",
        "copyright_line": "All Copyrights reserved with University",
        "statutory_footer": "Established by State Legislature Act | Recognized by UGC, Government of India",
        "enable_mirror_notes_margin": True,
        "default_committee": [],
    })

    with st.sidebar.expander("🎨 Customize University Identity & Logo", expanded=(sel_key == "+ Create / Customize University")):
        u_id = st.text_input("Profile Code:", value=base_prof.get("university_id", "CUSTOM_UNI"))
        u_name = st.text_input("University Name:", value=base_prof.get("university_name", ""))
        u_tag = st.text_input("Tagline:", value=base_prof.get("tagline", ""))
        u_school = st.text_input("School / Centre:", value=base_prof.get("school_name", ""))
        c_col1, c_col2 = st.columns(2)
        p_hex = c_col1.color_picker("Primary Color:", value=f"#{str(base_prof.get('primary_hex', '1E3A8A')).lstrip('#')}").lstrip("#")
        a_hex = c_col2.color_picker("Accent Color:", value=f"#{str(base_prof.get('accent_hex', 'D97706')).lstrip('#')}").lstrip("#")
        u_copy = st.text_input("Copyright Notice:", value=base_prof.get("copyright_line", ""))
        u_stat = st.text_area("Statutory / UGC Line:", value=base_prof.get("statutory_footer", ""), height=68)
        mirror_notes = st.checkbox("Enable Mirror Duplex 'Notes' Margin", value=bool(base_prof.get("enable_mirror_notes_margin", True)))
        logo_up = st.file_uploader("Upload University Logo (PNG/JPG):", type=["png", "jpg", "jpeg"], key="uni_logo_up")

        if st.button("💾 Save University Profile", use_container_width=True):
            updated = {
                "university_id": u_id,
                "university_name": u_name,
                "tagline": u_tag,
                "school_name": u_school,
                "primary_hex": p_hex,
                "accent_hex": a_hex,
                "copyright_line": u_copy,
                "statutory_footer": u_stat,
                "enable_mirror_notes_margin": mirror_notes,
                "default_committee": base_prof.get("default_committee", []),
            }
            save_university_profile(updated)
            st.success(f"Saved profile `{u_id.upper()}`!")

    logo_bytes = logo_up.getvalue() if logo_up is not None else _generate_brand_crest_png(u_name, p_hex, a_hex)
    st.sidebar.image(logo_bytes, width=72, caption=u_id.upper())

    st.sidebar.divider()
    st.sidebar.header("🤖 AI & Workspace Config")
    provider = st.sidebar.selectbox("AI Provider:", list(AI_PROVIDERS.keys()), index=0)
    selected_model = st.sidebar.selectbox("Model:", AI_PROVIDERS[provider], index=0)
    custom_model = st.sidebar.text_input("Custom Model ID (Optional):", placeholder="Leave blank for dropdown")
    model_id = custom_model.strip() or selected_model
    override_key = st.sidebar.text_input("Override API Key (Optional):", type="password")
    dest_root_id = extract_drive_id(st.sidebar.text_input("Destination Root Drive Folder URL / ID:"))

    active_brand = {
        "university_id": u_id.upper(),
        "university_name": u_name,
        "tagline": u_tag,
        "school_name": u_school,
        "primary_hex": p_hex,
        "accent_hex": a_hex,
        "copyright_line": u_copy,
        "statutory_footer": u_stat,
        "enable_mirror_notes_margin": mirror_notes,
        "default_committee": base_prof.get("default_committee", []),
        "logo_bytes": logo_bytes,
    }
    return {"brand": active_brand, "provider": provider, "model_id": model_id, "override_key": override_key, "dest_root_id": dest_root_id}


def main():
    current_user = st.session_state.get("l5_auth_user", "")
    cfg = render_sidebar_brand_and_config()
    brand = cfg["brand"]
    provider = cfg["provider"]
    model_id = cfg["model_id"]
    override_key = cfg["override_key"]
    dest_root_id = cfg["dest_root_id"]

    st.title("✨ Layer-5: Audit-Verified Enrichment, Book Compiler & Subject QA")
    st.caption(
        f"Active Institution: **{brand['university_name']}** (`{brand['university_id']}`)  •  "
        "Mandatory Layer-4+ Review Report Checkpoint  •  Subject-Wide Cumulative Reporting"
    )

    if not _is_layer5_authorized(current_user):
        st.error(
            f"Access Restricted: `{current_user}` is listed in `LAYER4_ONLY_EMAILS` "
            "(read-only Layer-4/4+ auditor scope). Contact QA Lead to unlock Layer-5."
        )
        st.stop()

    drive_svc, sheets_svc, gcp_msg, oauth_url = get_gcp_services()
    if drive_svc is None and oauth_url:
        st.link_button("🔐 Sign In with Google to Connect Master Tracker & Drive", oauth_url, type="primary")

    aspect_mode = st.radio(
        "Select Layer-5 Operational Pillar:",
        [
            "📝 Aspect 1: Unit Remediation & Audit-Verified Enrichment (Track A & Track B)",
            "📚 Aspect 2: 1-to-N Complete Book Compiler (Mirror Duplex & Dynamic Cover)",
            "📊 Aspect 3: Subject-Wide Cumulative QA Report (SLM, MCQ, QB, PPT, FAQ)",
        ],
        horizontal=True,
    )

    # ==========================================================================
    # ASPECT 3: SUBJECT-WIDE CUMULATIVE QA & ENRICHMENT REPORT
    # ==========================================================================
    if aspect_mode.startswith("📊"):
        st.subheader("📊 Subject-Wide Cumulative QA & Enrichment Report (SLM, MCQ, QB, PPT, FAQ)")
        st.info(
            "Upload all Layer-4 / Layer-4+ Audit Review Reports (`.xlsx`, `.pdf`, `.csv`, `.json`) across an entire subject "
            "(**SLM** units, **MCQ** banks, **QB** subjective banks, **PPT** decks, **FAQ** sheets) to aggregate every "
            "evaluated parameter into a unified Subject Readiness Dossier and Release Certificate."
        )

        sc1, sc2, sc3, sc4 = st.columns(4)
        s_prog = sc1.text_input("Program:", value="MBA", key="subj_prog")
        s_code = sc2.text_input("Subject / Course Code:", value="MG502M", key="subj_code")
        s_name = sc3.text_input("Subject / Course Name:", value="Organizational Behavior", key="subj_name")
        s_term = sc4.text_input("Cohort / Academic Year:", value="2026 Cohort", key="subj_term")

        rep_uploads = st.file_uploader(
            "Upload All Subject Audit Review Reports (SLM, MCQ, QB, PPT, FAQ — .xlsx / .pdf / .csv / .json):",
            type=["xlsx", "pdf", "csv", "json"],
            accept_multiple_files=True,
            key="subj_rep_uploader",
        )
        if not rep_uploads:
            st.stop()

        preview_rows = []
        for idx, rf in enumerate(rep_uploads, start=1):
            detected_cat = classify_deliverable_category(rf.name)
            preview_rows.append({
                "#": idx,
                "Audit Report Filename": rf.name,
                "Deliverable Type": detected_cat,
                "Layer-5 Enrichment State": "Auto-Evaluated from Report",
            })

        edited_types = st.data_editor(
            pd.DataFrame(preview_rows),
            column_config={
                "Deliverable Type": st.column_config.SelectboxColumn(
                    "Deliverable Type",
                    options=["SLM", "MCQ", "QB", "PPT", "FAQ", "OTHER"],
                    required=True,
                )
            },
            disabled=["#", "Audit Report Filename"],
            use_container_width=True,
            hide_index=True,
            key="subj_type_editor",
        )

        if st.button("📈 Generate Cumulative Subject QA & Enrichment Dossier", type="primary"):
            type_map = dict(zip(edited_types["Audit Report Filename"], edited_types["Deliverable Type"]))
            items_payload = [
                {
                    "filename": rf.name,
                    "bytes": rf.getvalue(),
                    "override_type": type_map.get(rf.name, "Auto-Detect"),
                }
                for rf in rep_uploads
            ]
            subj_meta = {
                "program_name": s_prog,
                "course_code": s_code,
                "course_name": s_name,
                "academic_term": s_term,
            }
            with st.spinner("Consolidating all evaluated parameters across SLM, MCQ, QB, PPT & FAQ deliverables..."):
                xl_bytes, cert_bytes, summary = compile_subject_cumulative_dossier(
                    subject_meta=subj_meta,
                    report_items=items_payload,
                    brand_profile=brand,
                    reviewer_email=current_user,
                )
                st.session_state["l5_subj_dossier"] = {
                    "xl_bytes": xl_bytes,
                    "cert_bytes": cert_bytes,
                    "summary": summary,
                    "code": re.sub(r"[^A-Za-z0-9_-]+", "_", s_code),
                }

        sdos = st.session_state.get("l5_subj_dossier")
        if sdos:
            sm = sdos["summary"]
            st.divider()
            k1, k2, k3, k4 = st.columns(4)
            k1.metric("Subject Deliverables Audited", f"{sm['total_files']} File(s)")
            k2.metric("Approved & Q2-Compliant", f"{sm['total_approved']} / {sm['total_files']}")
            k3.metric("Subject Composite QA Score", f"{sm['composite_score']:.1f}%")
            k4.metric("Subject Release Verdict", "READY" if sm["subject_ready"] else "HOLD / REMEDIATE")

            st.markdown("### 1. Multi-Deliverable Coverage Matrix (`SLM`, `MCQ`, `QB`, `PPT`, `FAQ`)")
            st.dataframe(pd.DataFrame(sm["category_rows"]), use_container_width=True, hide_index=True)

            with st.expander("📋 View Full Parameter-by-Parameter Evaluation Ledger Across All Assets", expanded=False):
                st.dataframe(
                    pd.DataFrame(
                        sm["parameter_rows"],
                        columns=["Deliverable Type", "Source Report", "Evaluated Parameter", "Score / Value", "Status", "Remarks"],
                    ),
                    use_container_width=True,
                    hide_index=True,
                )

            dl1, dl2 = st.columns(2)
            dl1.download_button(
                f"📊 Download Cumulative Subject Excel Dossier ({sdos['code']}_Subject_QA_Dossier.xlsx)",
                data=sdos["xl_bytes"],
                file_name=f"{sdos['code']}_Subject_Cumulative_QA_Dossier.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )
            dl2.download_button(
                f"📜 Download Subject QA Release Certificate ({sdos['code']}_Release_Certificate.docx)",
                data=sdos["cert_bytes"],
                file_name=f"{sdos['code']}_Subject_QA_Release_Certificate.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
        st.stop()

    # ==========================================================================
    # ASPECT 2: 1-TO-N COMPLETE BOOK COMPILER
    # ==========================================================================
    if aspect_mode.startswith("📚"):
        st.subheader("📚 1-to-N Complete SLM Book Compiler (Dynamic Branding & Mirror-Page Duplex)")
        st.info(
            "Combines multiple approved/enriched `.docx` units into a single print-ready book with **Institutional Cover**, "
            "**Editorial Governance Board**, **Master Syllabus Index (`PAGEREF`)**, **Course Objectives (`CO1–CO5`)**, "
            "**Odd-Page Recto Unit Starts**, and **Recto/Verso Mirror Headers with Outer-Margin 'Notes'**."
        )

        c_m1, c_m2, c_m3 = st.columns(3)
        prog_name = c_m1.text_input("Program Name:", value="MASTERS OF BUSINESS ADMINISTRATION (MBA)")
        crs_name = c_m2.text_input("Course Name:", value="Organizational Behavior")
        crs_code = c_m3.text_input("Course Code:", value="MG502M")

        c_m4, c_m5, c_m6 = st.columns(3)
        credits_val = c_m4.text_input("Credits:", value="04")
        doc_type_val = c_m5.text_input("Document Type:", value="Self-Learning Material (SLM) — Complete Course Book")
        edition_val = c_m6.text_input("Edition / Year:", value="2026 Edition")

        with st.expander("👥 Editorial Governance Board & Course Objectives (CO1–CO5) — Dynamic per Book", expanded=False):
            col_gov, col_co = st.columns(2)
            with col_gov:
                gov_df = pd.DataFrame(brand.get("default_committee") or [{"role": "Vice Chancellor - Chairman", "name": "Dr. Rakesh Kumar Jain"}])
                edited_gov = st.data_editor(gov_df, num_rows="dynamic", use_container_width=True, key="book_gov_ed")
            with col_co:
                co_df = pd.DataFrame([
                    {"co_code": "CO1", "description": "Understand the conceptual and theoretical framework of the discipline."},
                    {"co_code": "CO2", "description": "Analyze core operational mechanisms and behavioral drivers in real-world contexts."},
                    {"co_code": "CO3", "description": "Critically evaluate models, analytical methods, and empirical findings."},
                    {"co_code": "CO4", "description": "Demonstrate structured problem-solving and decision-making techniques."},
                    {"co_code": "CO5", "description": "Apply domain principles ethically to achieve sustainable outcomes."},
                ])
                edited_co = st.data_editor(co_df, num_rows="dynamic", use_container_width=True, key="book_co_ed")

        book_uploads = st.file_uploader(
            "Upload Approved/Enriched Unit Manuscripts in Sequence (.docx):",
            type=["docx"],
            accept_multiple_files=True,
            key="book_unit_uploader",
        )
        if not book_uploads:
            st.stop()

        unit_plan_rows = []
        for idx, u_file in enumerate(book_uploads, start=1):
            clean_stem = u_file.name.rsplit(".", 1)[0].replace("_Enriched_Approved", "")
            unit_plan_rows.append({
                "Sequence": idx,
                "Block Label": f"Block {(idx - 1) // 3 + 1}",
                "Unit Title": clean_stem,
                "Source Filename": u_file.name,
            })

        edited_plan = st.data_editor(pd.DataFrame(unit_plan_rows), num_rows="fixed", use_container_width=True, hide_index=True, key="book_seq_editor")

        if st.button("📖 Compile Complete Mirrored SLM Book (.docx)", type="primary"):
            file_lookup = {u.name: u.getvalue() for u in book_uploads}
            sorted_plan = edited_plan.sort_values(by="Sequence")
            ordered_units = [
                {
                    "filename": r["Source Filename"],
                    "block_label": str(r["Block Label"]),
                    "unit_title": str(r["Unit Title"]),
                    "bytes": file_lookup[r["Source Filename"]],
                }
                for _, r in sorted_plan.iterrows()
            ]
            book_meta = {
                "program_name": prog_name,
                "course_name": crs_name,
                "course_code": crs_code,
                "credits": credits_val,
                "document_type": doc_type_val,
                "edition_year": edition_val,
                "committee": edited_gov.to_dict(orient="records"),
                "course_outcomes": edited_co.to_dict(orient="records"),
            }
            with st.spinner("Assembling Cover, Governance Board, Master Index, Mirror Duplex Headers & Odd-Page Chapters..."):
                book_bytes = compile_complete_slm_book_docx(unit_docs=ordered_units, brand_profile=brand, book_meta=book_meta, reviewer_email=current_user)
                safe_book_name = re.sub(r"[^A-Za-z0-9_-]+", "_", f"{crs_code}_{crs_name}_Complete_Book") + ".docx"
                st.session_state["l5_compiled_book"] = {"bytes": book_bytes, "filename": safe_book_name, "unit_count": len(ordered_units)}

        cbook = st.session_state.get("l5_compiled_book")
        if cbook:
            st.success(f"Complete Book compiled! **{cbook['unit_count']} unit(s)** stitched with mirror duplex headers and **{brand['university_name']}** branding.")
            st.download_button(
                f"📥 Download Complete Book ({cbook['filename']})",
                data=cbook["bytes"],
                file_name=cbook["filename"],
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
            )
        st.stop()

    # ==========================================================================
    # ASPECT 1: UNIT REMEDIATION & AUDIT-VERIFIED CREATIVE ENRICHMENT
    # ==========================================================================
    source_mode = st.radio(
        "Select Manuscript & Audit Verification Source:",
        [
            "📁 Local Manuscript + Layer-4/4+ Review Report Checkpoint",
            "📊 Master Tracker Google Sheet Sync (Live Audited Row)",
        ],
        horizontal=True,
    )

    tracker_ctx = None
    ms_bytes = b""
    ms_filename = ""
    ms_title = ""
    audit_verification: Optional[Dict] = None

    if source_mode.startswith("📁"):
        u_col1, u_col2 = st.columns(2)
        with u_col1:
            up_ms = st.file_uploader("1. Upload Manuscript (.docx or .pptx):", type=["docx", "pptx"], key="l5_ms_up")
        with u_col2:
            up_rep = st.file_uploader(
                "2. Upload Layer-4 / Layer-4+ Audit Review Report (.xlsx, .pdf, .csv, .json):",
                type=["xlsx", "pdf", "csv", "json"],
                key="l5_rep_up",
                help="Mandatory checkpoint to unlock Track B Creative Enrichment. Without a verified Approved report, only Pre-Audit Track A Remediation is permitted.",
            )

        if up_ms is None:
            st.stop()

        ms_bytes = up_ms.getvalue()
        ms_filename = up_ms.name
        ms_title = up_ms.name.rsplit(".", 1)[0]
        unit_text = extract_text_from_bytes(ms_bytes, ms_filename)
        prose_paras = _extract_prose_paragraphs(unit_text)
        cv_metrics = compute_prose_cv_3dec(prose_paras)

        if up_rep is not None:
            audit_verification = parse_and_verify_audit_report(
                report_bytes=up_rep.getvalue(),
                report_filename=up_rep.name,
                manuscript_filename=ms_filename,
                manuscript_text=unit_text,
                manuscript_cv=cv_metrics["prose_cv"],
            )
        else:
            audit_verification = {
                "report_filename": "None Uploaded",
                "deliverable_type": classify_deliverable_category(ms_filename, unit_text),
                "audit_status": "No Review Report Uploaded",
                "audit_score": 0.0,
                "similarity_pct": None,
                "ai_content_pct": None,
                "q2_passed": False,
                "auditor_name": "Pending Layer-4 Review",
                "manuscript_matched": False,
                "match_reason": "No Layer-4 / Layer-4+ audit review report uploaded.",
                "evaluated_parameters": [],
                "flagged_defects": [],
                "verified_for_enrichment": False,
                "block_reasons": [
                    "Missing Layer-4 / Layer-4+ Audit Review Report: upload the corresponding audit dossier (.xlsx / .pdf / .json) to unlock Track B Creative Enrichment."
                ],
            }

    else:
        if drive_svc is None or sheets_svc is None:
            st.warning(gcp_msg or "Connect Google Drive above to load the Master Tracker.")
            st.stop()

        tracker_url = st.text_input("Paste Master Tracker Google Sheet URL:")
        if st.button("🔍 Load Master Tracker", type="primary"):
            tid = extract_drive_id(tracker_url)
            if not tid:
                st.error("Invalid Google Sheet URL.")
            else:
                meta = drive_svc.files().get(fileId=tid, fields="id,name,mimeType", supportsAllDrives=True).execute()
                xl_bytes, _ = download_drive_file_readonly(drive_svc, tid)
                wb = openpyxl.load_workbook(io.BytesIO(xl_bytes), data_only=True)
                st.session_state["l5_tracker"] = {"id": tid, "name": meta["name"], "xl_bytes": xl_bytes, "tabs": wb.sheetnames}

        trk = st.session_state.get("l5_tracker")
        if not trk:
            st.stop()

        c1, c2 = st.columns([3, 1])
        sel_tab = c1.selectbox("Select Worksheet Tab:", trk["tabs"])
        hdr_row = int(c2.number_input("Header Row (1-indexed):", min_value=1, max_value=20, value=1))

        wb = openpyxl.load_workbook(io.BytesIO(trk["xl_bytes"]), data_only=True)
        ws = wb[sel_tab]
        rows = list(ws.iter_rows(values_only=True))
        raw_headers = [str(h).strip() if h is not None and str(h).strip() else f"Col_{i+1}" for i, h in enumerate(rows[hdr_row - 1])]

        records = []
        for idx, r in enumerate(rows[hdr_row:], start=hdr_row + 1):
            rec = {"_sheet_row": idx}
            for ci, hn in enumerate(raw_headers):
                rec[hn] = r[ci] if ci < len(r) else ""
            records.append(rec)

        df = pd.DataFrame(records)
        link_col = next((h for h in raw_headers if any(k in h.lower() for k in ("manuscript", "content link", "file link", "drive link"))), raw_headers[0])
        topic_col = next((h for h in raw_headers if any(k in h.lower() for k in ("topic", "unit", "module", "title"))), link_col)
        status_col = next((h for h in raw_headers if "audit status" in h.lower()), "")
        l5_stat_col = next((h for h in raw_headers if h.lower() == "l5_approval_status"), "")

        df = df[df[link_col].astype(str).str.startswith("http")].copy()
        if df.empty:
            st.info("No manuscript links found in this tab.")
            st.stop()

        display_cols = ["_sheet_row", topic_col] + ([status_col] if status_col else []) + ([l5_stat_col] if l5_stat_col else [])
        st.dataframe(df[display_cols], use_container_width=True, hide_index=True)

        row_choices = {f"Row {r['_sheet_row']} — {str(r[topic_col])[:55]} [{r.get(status_col, 'Unverified')}]": r for _, r in df.iterrows()}
        chosen_label = st.selectbox("Select Audited Row for Layer-5 Staging:", list(row_choices.keys()))
        chosen_row = row_choices[chosen_label]

        if st.button("📥 Clone Source Manuscript Read-Only from Drive"):
            fid = extract_drive_id(str(chosen_row[link_col]))
            with st.spinner("Cloning source file in memory (original remains untouched)..."):
                b_data, f_name = download_drive_file_readonly(drive_svc, fid)
                st.session_state["l5_loaded_doc"] = {
                    "bytes": b_data,
                    "filename": f_name,
                    "title": str(chosen_row[topic_col]).strip() or f_name.rsplit(".", 1)[0],
                    "l4_status": str(chosen_row.get(status_col, "")),
                    "tracker_ctx": {
                        "spreadsheet_id": trk["id"],
                        "tab_name": sel_tab,
                        "header_row": hdr_row,
                        "data_row": int(chosen_row["_sheet_row"]),
                        "raw_headers": raw_headers,
                    },
                }

        loaded = st.session_state.get("l5_loaded_doc")
        if not loaded:
            st.stop()
        ms_bytes = loaded["bytes"]
        ms_filename = loaded["filename"]
        ms_title = loaded["title"]
        tracker_ctx = loaded["tracker_ctx"]

        unit_text = extract_text_from_bytes(ms_bytes, ms_filename)
        prose_paras = _extract_prose_paragraphs(unit_text)
        cv_metrics = compute_prose_cv_3dec(prose_paras)
        synth_rep = f"Audit Status: {loaded['l4_status']}\nTarget Document: {ms_filename}\n"
        audit_verification = parse_and_verify_audit_report(
            report_bytes=synth_rep.encode("utf-8"),
            report_filename="Master_Tracker_Row.csv",
            manuscript_filename=ms_filename,
            manuscript_text=unit_text,
            manuscript_cv=cv_metrics["prose_cv"],
        )

    # ==========================================================================
    # TOPOLOGY & REVIEW CHECKPOINT VERIFICATION PANEL
    # ==========================================================================
    topology = analyze_document_topology(
        unit_text, filename=ms_filename, docx_bytes=ms_bytes if ms_filename.lower().endswith(".docx") else b""
    )
    low_windows = detect_low_cv_windows(unit_text)
    is_l4_approved = bool(audit_verification and audit_verification["verified_for_enrichment"])

    st.divider()
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Self-Recognized Type", topology["archetype_label"][:28], delta=f"{topology['unit_count']} Unit(s)")
    m2.metric(
        "Review Report Checkpoint",
        f"{audit_verification['audit_status']} ({audit_verification['audit_score']:.1f}%)" if audit_verification["audit_score"] > 0 else audit_verification["audit_status"],
        delta="MATCHED" if audit_verification["manuscript_matched"] else "UNVERIFIED",
    )
    m3.metric("Prose Burstiness CV", f"{cv_metrics['prose_cv']:.3f}", delta="PASS >= 0.400" if cv_metrics["passed"] else "FAIL < 0.400")
    m4.metric("Track B Creative Unlock", "UNLOCKED" if is_l4_approved else "LOCKED")

    if is_l4_approved:
        st.success(
            f"**Checkpoint Passed:** Audit report `{audit_verification['report_filename']}` verified "
            f"(Score: **{audit_verification['audit_score']:.1f}%**, Match: {audit_verification['match_reason']}). "
            "Track B Creative Enrichment is unlocked."
        )
    else:
        st.warning(
            "**Track B Creative Enrichment Locked:** " + " | ".join(audit_verification["block_reasons"])
        )

    if audit_verification["flagged_defects"]:
        with st.expander(f"🚩 Layer-4+ Flagged Defects Extracted from Review Report ({len(audit_verification['flagged_defects'])})", expanded=True):
            for df_item in audit_verification["flagged_defects"]:
                st.markdown(f"- {df_item}")

    tab_a, tab_b = st.tabs([
        "🛠️ Track A: AI Auto-Remediation ('Generate Fix-It Draft')",
        "🎨 Track B: Post-Audit Creative Enrichment (Requires Verified Approved Report)",
    ])

    with tab_a:
        if low_windows:
            st.warning(f"Detected **{len(low_windows)}** low-burstiness prose window(s) below `0.400` CV ready for auto-remediation.")
        else:
            st.info("Track A will generate missing CYP boxes, Model Answer Keys, Unit Glossary, and Mermaid specs grounded in your manuscript and audit report.")

        if st.button("🚀 Generate Fix-It Draft (Track A)", type="primary"):
            with st.status("Running Track A Auto-Remediation...", expanded=True) as status_box:
                items = generate_track_a_fixit_items(
                    unit_text=unit_text,
                    provider=provider,
                    model_id=model_id,
                    override_api_key=override_key,
                    progress_cb=lambda msg: status_box.write(msg),
                    topology=topology,
                    report_defects=audit_verification["flagged_defects"],
                )
                status_box.update(label="Track A Fix-It Draft ready for review!", state="complete")
                st.session_state["l5_staged_items"] = items
                st.session_state["l5_active_track"] = "TRACK_A"
                st.session_state["l5_compiled"] = None

    with tab_b:
        if not is_l4_approved:
            st.error(
                "Track B Creative Enrichment cannot run without a verified Layer-4 / Layer-4+ Audit Review Report "
                "confirming `Approved` status (>= 85%), Quadrant-II originality compliance, and Prose CV >= 0.400."
            )
        elif ms_filename.lower().endswith(".pptx"):
            st.info("PPTX Deliverable Detected: Phase 1 (₹0) native OpenXML slide transitions & entrance animations ready.")
            tr_style = st.selectbox("Slide Transition Style:", ["fade", "push"])
            if st.button("✨ Apply Native XML Slide Animations (Track B Phase 1 — ₹0)", type="primary"):
                pptx_out, pptx_meta = enrich_pptx_with_native_animations(ms_bytes, current_user, transition_style=tr_style)
                st.session_state["l5_compiled"] = {
                    "bytes": pptx_out,
                    "filename": f"{ms_title}_Enriched_Approved.pptx",
                    "mime": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    "summary": f"Animated {pptx_meta['animated_shapes']} shapes across {pptx_meta['slides_processed']} slides.",
                }
                st.session_state["l5_staged_items"] = []
                st.session_state["l5_active_track"] = "TRACK_B_PPTX"
        else:
            c_p1, c_p2 = st.columns(2)
            with c_p1:
                st.markdown(
                    f"**Phase 1 (₹0 — Branded to {brand['university_id']}):** Multi-shape infographics (Wheels, Pillars, Flowcharts), "
                    "XML Callout Cards, QR Micro-Quiz hooks, plus optional In-Text CYP & Caselet blocks."
                )
                inc_cyp_case = st.checkbox("Include In-Text Check Your Progress Boxes & Applied Caselets", value=True)
            with c_p2:
                opt_in_p2 = st.checkbox("Opt-In Phase 2: Generative Textbook Illustrations (Budget-Capped)", value=False)
                img_cap = st.slider("Max Illustrations per Manuscript:", 1, 3, 2, disabled=not opt_in_p2)
                img_prov = st.selectbox("Image Engine:", ["OpenAI (DALL-E 3)", "Google Gemini (Imagen 3)"], disabled=not opt_in_p2)

            if st.button("✨ Generate Creative Enrichment Draft (Track B)", type="primary"):
                with st.spinner("Scanning body engagement gaps & staging university-branded enrichment assets..."):
                    items = generate_track_b_enrichment_items(
                        unit_text=unit_text,
                        docx_bytes=ms_bytes if ms_filename.lower().endswith(".docx") else b"",
                        provider=provider,
                        model_id=model_id,
                        override_api_key=override_key,
                        enable_phase2_images=opt_in_p2,
                        max_phase2_images=img_cap,
                        image_provider=img_prov,
                        include_cyp_and_caselets=inc_cyp_case,
                        topology=topology,
                    )
                    st.session_state["l5_staged_items"] = items
                    st.session_state["l5_active_track"] = "TRACK_B_DOCX"
                    st.session_state["l5_compiled"] = None

    staged_items = st.session_state.get("l5_staged_items", [])
    if staged_items:
        st.divider()
        st.subheader("🔍 Staged Reviewer Preview & Item-Level Approval")

        for item in staged_items:
            c_chk, c_body = st.columns([0.08, 0.92])
            with c_chk:
                item["approved"] = st.checkbox("Approve", value=item.get("approved", False), key=f"l5_chk_{item['id']}", label_visibility="collapsed")
            with c_body:
                with st.expander(f"[{item['asset_type']}] {item['summary']} — Anchor: {item['anchor']}", expanded=(item["asset_type"] == "High-Burstiness Prose Rewrite")):
                    data = item.get("data", {})
                    if item["asset_type"] == "High-Burstiness Prose Rewrite":
                        col_orig, col_new = st.columns(2)
                        col_orig.markdown(f"**Original (CV: `{data['before_cv']:.3f}`)**")
                        col_orig.info(data["original"])
                        col_new.markdown(f"**Rewritten (CV: `{data['after_cv']:.3f}`)**")
                        col_new.success(data["rewritten"])
                    elif isinstance(data, dict) and data.get("image_bytes"):
                        st.image(data["image_bytes"], caption=data.get("title", "Generated Illustration"))
                    elif isinstance(data, list):
                        st.table(data)
                    else:
                        st.json({k: v for k, v in data.items() if k != "image_bytes"})

        approved_items = [x for x in staged_items if x.get("approved")]
        if st.button("✅ Compile Approved Staged Items", type="primary"):
            if not approved_items:
                st.error("Check at least one staged item above before compiling.")
            else:
                track = st.session_state.get("l5_active_track", "TRACK_A")
                if track == "TRACK_B_DOCX" and ms_filename.lower().endswith(".docx"):
                    out_bytes = compile_approved_enriched_docx(ms_bytes, approved_items, current_user, brand_profile=brand)
                    out_name = f"{ms_title}_Enriched_Approved.docx"
                else:
                    has_report = audit_verification["report_filename"] != "None Uploaded"
                    out_bytes = compile_approved_fixit_patch_docx(approved_items, ms_title, current_user, brand_profile=brand, audit_verified=has_report)
                    suffix = "_FixIt_Patch_Approved.docx" if has_report else "_FixIt_Patch_PreAudit_Draft.docx"
                    out_name = f"{ms_title}{suffix}"

                st.session_state["l5_compiled"] = {
                    "bytes": out_bytes,
                    "filename": out_name,
                    "mime": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    "summary": f"{len(approved_items)} approved item(s) compiled.",
                }

    compiled = st.session_state.get("l5_compiled")
    if compiled:
        st.divider()
        st.subheader("🔏 Governance Sign-Off & Drive / Master Tracker Archival")
        st.download_button(
            f"📥 Download Cloned Artifact ({compiled['filename']})",
            data=compiled["bytes"],
            file_name=compiled["filename"],
            mime=compiled["mime"],
        )

        with st.form("l5_governance_signoff_form"):
            approver_name = st.text_input("Explicit Human Reviewer Sign-Off (`L5_Approved_By`):", value=current_user).strip()
            confirm_clone = st.checkbox("I confirm this clone-on-write output has been reviewed and preserves original manuscript integrity.")
            submitted = st.form_submit_button("☁️ Approve, Upload to Drive & Sync Master Tracker", type="primary")

            if submitted:
                if not approver_name or not confirm_clone:
                    st.error("Both `L5_Approved_By` sign-off and confirmation checkbox are required.")
                elif drive_svc is None or not dest_root_id:
                    st.error("Connect Google Drive and enter the Destination Root Folder ID in the sidebar to archive.")
                else:
                    with st.spinner("Uploading approved clone to Google Drive & updating Master Tracker..."):
                        drive_link = upload_layer5_artifact_to_drive(
                            drive_svc=drive_svc,
                            root_folder_id=dest_root_id,
                            hierarchy_parts=["Layer-5 Approved Outputs", brand["university_id"]],
                            filename=compiled["filename"],
                            file_bytes=compiled["bytes"],
                            mime_type=compiled["mime"],
                        )
                        if tracker_ctx and sheets_svc is not None:
                            sync_layer5_master_tracker_row(
                                sheets_svc=sheets_svc,
                                spreadsheet_id=tracker_ctx["spreadsheet_id"],
                                tab_name=tracker_ctx["tab_name"],
                                header_row_1idx=tracker_ctx["header_row"],
                                data_row_1idx=tracker_ctx["data_row"],
                                raw_headers=tracker_ctx["raw_headers"],
                                enriched_doc_link=drive_link,
                                l5_approval_status="Approved",
                                l5_approved_by=approver_name,
                            )
                        st.success(f"Archived & synced! **L5_Approved_By:** `{approver_name}` | [Open Approved Artifact]({drive_link})")


if __name__ == "__main__":
    main()
