import os
import smtplib
import random
import time
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import streamlit as st

# Team Role Directory: Assigns specific rights by email
UPGRAD_ROLE_DIRECTORY = {
    "sandeep1.shrivastava@upgrad.com": "Lead Administrator",
    # Add additional team emails and roles as required:
    # "checker1@upgrad.com": "Checker",
    # "author1@upgrad.com": "Maker",
}

def get_role_for_email(email: str) -> str:
    """Resolves role from directory; defaults unknown @upgrad.com users to Maker."""
    clean_email = email.strip().lower()
    return UPGRAD_ROLE_DIRECTORY.get(clean_email, "Maker")

def send_otp_email(recipient_email: str, otp_code: str) -> bool:
    """Dispatches a branded 6-digit OTP to a registered @upgrad.com address."""
    smtp_server = st.secrets.get("SMTP_SERVER", "smtp.gmail.com")
    smtp_port = int(st.secrets.get("SMTP_PORT", 587))
    smtp_user = st.secrets.get("SMTP_USER", "")
    smtp_password = st.secrets.get("SMTP_PASSWORD", "")
    sender_name = st.secrets.get("SENDER_NAME", "upGrad SLM QA Quality Portal")

    if not smtp_user or not smtp_password:
        print(f"[AUTH CONSOLE] SMTP credentials missing. OTP for {recipient_email}: {otp_code}")
        return False

    subject = f"{otp_code} is your upGrad SLM Quality Portal Verification Code"

    html_body = f"""
    <html>
      <body style="font-family: Arial, sans-serif; color: #222; line-height: 1.6;">
        <div style="max-width: 500px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 8px; padding: 24px;">
          <h2 style="color: #c82333; margin-top: 0;">upGrad SLM Quality Assurance</h2>
          <p>Hello,</p>
          <p>You requested a one-time login passcode for the <strong>SLM Academic Quality Audit Studio</strong>.</p>
          <div style="background: #f8f9fa; border-radius: 6px; padding: 18px; text-align: center; margin: 24px 0;">
            <span style="font-size: 32px; font-weight: bold; letter-spacing: 6px; color: #111;">{otp_code}</span>
          </div>
          <p style="font-size: 13px; color: #666;">This code is valid for <strong>5 minutes</strong>. If you did not request this, please ignore this email.</p>
        </div>
      </body>
    </html>
    """

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = f"{sender_name} <{smtp_user}>"
    msg["To"] = recipient_email
    msg.attach(MIMEText(html_body, "html"))

    try:
        server = smtplib.SMTP(smtp_server, smtp_port, timeout=10)
        server.starttls()
        server.login(smtp_user, smtp_password)
        server.sendmail(smtp_user, recipient_email, msg.as_string())
        server.quit()
        return True
    except Exception as e:
        print(f"[SMTP Error]: {e}")
        return False
