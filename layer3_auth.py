import hmac
import re
import secrets as py_secrets
import smtplib
import time
from email.mime.text import MIMEText
import streamlit as st


def parse_approved_emails() -> set[str]:
    raw = (
        st.secrets.get("APPROVED_EMAILS")
        or st.secrets.get("ALLOWED_EMAILS")
        or st.secrets.get("AUTHORIZED_EMAILS")
        or ""
    )
    if isinstance(raw, (list, tuple, set)):
        return {str(e).strip().lower() for e in raw if str(e).strip()}
    return {e.strip().lower() for e in re.split(r"[,;\n\r]+", str(raw)) if e.strip()}


def send_otp_email(to_email: str, code: str) -> tuple[bool, str]:
    host = str(st.secrets.get("SMTP_HOST", "smtp.gmail.com")).strip()
    port = int(st.secrets.get("SMTP_PORT", 587))
    user = str(st.secrets.get("SMTP_USER") or st.secrets.get("SENDER_EMAIL") or "").strip()
    pwd = str(st.secrets.get("SMTP_PASS") or st.secrets.get("APP_PASSWORD") or "").strip()
    if not user or not pwd:
        return False, "SMTP_USER or SMTP_PASS is missing in Streamlit Secrets."

    msg = MIMEText(
        f"Academic QA Platform (Layer-3 Document Auditor)\n\n"
        f"Your one-time sign-in verification code is: {code}\n\n"
        "This code expires in 15 minutes.",
        "plain",
        "utf-8",
    )
    msg["Subject"] = f"QA Verification Code: {code}"
    msg["From"] = f"Academic QA Platform <{user}>"
    msg["To"] = to_email

    try:
        with smtplib.SMTP(host, port, timeout=20) as server:
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(user, pwd)
            server.sendmail(user, [to_email], msg.as_string())
        return True, "Sent"
    except Exception as exc:
        return False, f"SMTP error: {exc}"


def require_team_auth() -> bool:
    approved = parse_approved_emails()
    passcode = str(st.secrets.get("TEAM_PASSCODE", "")).strip()

    if not approved and not passcode:
        return True

    if st.session_state.get("layer3_auth_ok", False):
        usr = st.session_state.get("layer3_auth_user", "Team Member")
        st.sidebar.success(f"👤 Signed in: **{usr}**")
        if st.sidebar.button("🔒 Sign Out", use_container_width=True):
            for k in ("layer3_auth_ok", "layer3_auth_user", "layer3_otp_code", "layer3_otp_email", "layer3_otp_ts"):
                st.session_state.pop(k, None)
            st.rerun()
        return True

    st.markdown("## 🔐 Layer-3 Academic Document Auditor — Team Sign-In")
    st.caption("Verify your approved institutional email via One-Time Password (OTP) to unlock the QA pipeline.")

    tab_labels = (["📧 Email OTP Login (Approved IDs)"] if approved else []) + (["🔑 Team Passcode"] if passcode else [])
    tabs = st.tabs(tab_labels)
    idx = 0

    if approved:
        with tabs[idx]:
            with st.form("layer3_send_otp_form", clear_on_submit=False):
                email_in = st.text_input("Approved Work Email ID:", placeholder="name@upgrad.com").strip().lower()
                if st.form_submit_button("📨 Send 6-Digit OTP", type="primary"):
                    if email_in not in approved:
                        st.error("Access denied: This email ID is not in the approved reviewer list.")
                    else:
                        code = f"{py_secrets.randbelow(900000) + 100000}"
                        ok, status = send_otp_email(email_in, code)
                        if ok:
                            st.session_state.update({
                                "layer3_otp_code": code,
                                "layer3_otp_email": email_in,
                                "layer3_otp_ts": time.time(),
                            })
                            st.success(f"A 6-digit OTP has been sent to **{email_in}** (valid for 15 minutes).")
                        else:
                            st.error(status)

            if st.session_state.get("layer3_otp_email"):
                target = st.session_state["layer3_otp_email"]
                with st.form("layer3_verify_otp_form", clear_on_submit=False):
                    entered = st.text_input(f"Enter 6-Digit OTP sent to {target}:", max_chars=6).strip()
                    if st.form_submit_button("✅ Verify & Unlock Layer-3", type="primary"):
                        if time.time() - st.session_state.get("layer3_otp_ts", 0) > 900:
                            st.error("OTP expired (15-minute limit). Please request a new code.")
                        elif hmac.compare_digest(entered, str(st.session_state.get("layer3_otp_code", ""))):
                            st.session_state["layer3_auth_ok"] = True
                            st.session_state["layer3_auth_user"] = target
                            st.rerun()
                        else:
                            st.error("Invalid verification code.")
        idx += 1

    if passcode:
        with tabs[idx]:
            with st.form("layer3_passcode_form", clear_on_submit=False):
                pw = st.text_input("Team Passcode:", type="password", placeholder="Enter team passcode")
                if st.form_submit_button("Unlock with Passcode", type="primary"):
                    if hmac.compare_digest(pw.strip(), passcode):
                        st.session_state["layer3_auth_ok"] = True
                        st.session_state["layer3_auth_user"] = "Passcode Session"
                        st.rerun()
                    else:
                        st.error("Incorrect passcode.")

    return False