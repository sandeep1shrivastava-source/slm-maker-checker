import ast
import py_compile

target = "content_reviewer_app.py"

with open(target, "r", encoding="utf-8") as f:
    source = f.read()

tree = ast.parse(source)

# Locate the dispatch_email_otp function definition exactly
target_node = None
for node in ast.walk(tree):
    if isinstance(node, ast.FunctionDef) and node.name == "dispatch_email_otp":
        target_node = node
        break

if not target_node:
    print("ERROR: Function 'dispatch_email_otp' not found in content_reviewer_app.py.")
    exit(1)

print(f"Found 'dispatch_email_otp' spanning lines {target_node.lineno} to {target_node.end_lineno}.")

# Production dual-port implementation (Port 465 SSL default with Port 587 fallback)
new_dispatch_func = '''def dispatch_email_otp(recipient_email, otp_code):
    """Dispatches 6-digit OTP using Port 465 SSL or Port 587 STARTTLS."""
    import smtplib, ssl
    from email.mime.text import MIMEText
    from email.mime.multipart import MIMEMultipart

    smtp_cfg = st.secrets.get("smtp", {})
    host = smtp_cfg.get("server", "smtp.gmail.com")
    port = int(smtp_cfg.get("port", 465))
    user = smtp_cfg.get("user", "")
    pwd = smtp_cfg.get("password", "")
    sender = smtp_cfg.get("sender", user)

    if not user or not pwd:
        print(f"[AUTH CONSOLE] SMTP credentials missing. OTP for {recipient_email}: {otp_code}")
        return False

    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"{otp_code} is your upGrad SLM Quality Portal Verification Code"
    msg["From"] = f"upGrad SLM QA <{sender}>"
    msg["To"] = recipient_email

    html = f"""
    <div style='font-family: Arial, sans-serif; padding: 20px; border: 1px solid #e0e0e0; border-radius: 8px; max-width: 480px;'>
        <h3 style='color: #c82333; margin-top: 0;'>upGrad SLM Quality Portal</h3>
        <p>Your one-time login verification code is:</p>
        <div style='background: #f4f4f4; padding: 14px; text-align: center; border-radius: 6px; margin: 18px 0;'>
            <span style='font-size: 30px; font-weight: bold; letter-spacing: 5px; color: #111;'>{otp_code}</span>
        </div>
        <p style='font-size: 12px; color: #777;'>Valid for 5 minutes. If you did not request this, please disregard.</p>
    </div>
    """
    msg.attach(MIMEText(html, "html"))

    try:
        if port == 465:
            ctx = ssl.create_default_context()
            with smtplib.SMTP_SSL(host, port, context=ctx, timeout=12) as server:
                server.login(user, pwd)
                server.sendmail(sender, recipient_email, msg.as_string())
        else:
            with smtplib.SMTP(host, port, timeout=12) as server:
                server.ehlo()
                server.starttls()
                server.ehlo()
                server.login(user, pwd)
                server.sendmail(sender, recipient_email, msg.as_string())
        return True
    except Exception as e:
        print(f"[SMTP Error]: {e}")
        return False'''

lines = source.splitlines(keepends=True)
# Replace only the exact AST line range for this function
new_lines = lines[: target_node.lineno - 1] + [new_dispatch_func + "\n"] + lines[target_node.end_lineno :]

with open(target, "w", encoding="utf-8") as f:
    f.writelines(new_lines)

py_compile.compile(target, doraise=True)
print("SUCCESS: dispatch_email_otp cleanly updated to Port 465 SSL and compiled with 0 errors!")