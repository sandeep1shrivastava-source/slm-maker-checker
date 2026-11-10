"""
SLM Maker-Checker Quality Assurance Platform
Features: Workflow Queue, Audit History Ledger, Dynamic RBAC, Self-Serve Catalog
"""

import io
import os
import re
import json
import uuid
import random
import smtplib
import ssl
import traceback
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

import streamlit as st
import pandas as pd
import openpyxl
import docx
from docx.shared import Pt, RGBColor
import google.generativeai as genai

# ==============================================================================
# 1. PLATFORM CONFIGURATION & LOCAL STORAGE (UPLOADS, QUEUE, LEDGER)
# ==============================================================================
st.set_page_config(page_title="SLM QA Platform", page_icon="🎓", layout="wide")

DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
UPLOAD_DIR = os.path.join(DATA_DIR, "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

CATALOG_FILE = os.path.join(DATA_DIR, "academic_catalog.json")
ROLES_FILE = os.path.join(DATA_DIR, "user_roles.json")
QUEUE_FILE = os.path.join(DATA_DIR, "workflow_queue.json")
LEDGER_FILE = os.path.join(DATA_DIR, "audit_history.csv")

# ==============================================================================
# 2. PERSISTENCE ENGINE (ROLES, CATALOG, QUEUE)
# ==============================================================================
DEFAULT_ROLES = {
    "sandeep1.shrivastava@upgrad.com": "Admin",  # Your verified Admin account
    "sandeep.shrivastava@upgrad.com": "Admin",
    "admin@upgrad.com": "Admin",
    "checker@upgrad.com": "Checker",
    "maker@upgrad.com": "Maker"
}

def load_json(filepath, default):
    if os.path.exists(filepath):
        try:
            with open(filepath, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception: pass
    return default

def save_json(filepath, data):
    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)

def load_queue(): return load_json(QUEUE_FILE, [])
def save_queue(q): save_json(QUEUE_FILE, q)

def append_to_ledger(log_entry: dict):
    df = pd.DataFrame([log_entry])
    if not os.path.exists(LEDGER_FILE):
        df.to_csv(LEDGER_FILE, index=False)
    else:
        df.to_csv(LEDGER_FILE, mode='a', header=False, index=False)

def load_ledger() -> pd.DataFrame:
    if os.path.exists(LEDGER_FILE):
        return pd.read_csv(LEDGER_FILE)
    return pd.DataFrame(columns=["Timestamp", "Maker", "Checker", "University", "Program", "Semester", "File Name", "Quality Score", "Determination"])

# Initialize Session State
for k, v in {
    "authenticated": False, "user_email": "", "user_role": "Maker", "otp_sent": False,
    "audit_completed": False, "audit_df": None, "defect_df": None,
    "quality_score": 0.0, "determination": "PENDING"
}.items():
    if k not in st.session_state: st.session_state[k] = v

# ==============================================================================
# 3. AUTHENTICATION & SMTP_SSL PORT 465 DISPATCH
# ==============================================================================
def dispatch_email_otp(recipient_email: str, otp_code: str) -> bool:
    smtp_host, smtp_port = st.secrets.get("smtp", {}).get("host", "smtp.gmail.com"), 465
    smtp_user, smtp_pass = st.secrets.get("smtp", {}).get("user"), st.secrets.get("smtp", {}).get("password")
    if not smtp_user or not smtp_pass: return False

    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"Your SLM QA Verification Code: {otp_code}"
    msg["From"] = f"SLM Quality Engine <{smtp_user}>"
    msg["To"] = recipient_email
    msg.attach(MIMEText(f"<h3>Verification Code: {otp_code}</h3>", "html"))

    try:
        with smtplib.SMTP_SSL(smtp_host, port=smtp_port, context=ssl.create_default_context()) as s:
            s.login(smtp_user, smtp_pass)
            s.sendmail(smtp_user, recipient_email, msg.as_string())
        return True
    except: return False

def render_login_screen():
    st.markdown("<h1 style='text-align:center;'>🎓 Academic SLM Maker-Checker Platform</h1>", unsafe_allow_html=True)
    _, col, _ = st.columns([1, 2, 1])
    with col.container(border=True):
        email = st.text_input("Corporate Email Address", placeholder="name@upgrad.com")
        roles_registry = load_json(ROLES_FILE, DEFAULT_ROLES)

        if not st.session_state["otp_sent"]:
            if st.button("Generate & Send OTP", type="primary", use_container_width=True):
                if not email.strip().lower().endswith("@upgrad.com"):
                    st.error("Restricted to @upgrad.com domains.")
                else:
                    otp = str(random.randint(100000, 999999))
                    st.session_state.update({"generated_otp": otp, "user_email": email.strip().lower(), "user_role": roles_registry.get(email.strip().lower(), "Maker")})
                    sent = dispatch_email_otp(email.strip().lower(), otp)
                    st.session_state["otp_sent"] = True
                    if sent: st.success("OTP Dispatched.")
                    else: st.info("Use master bypass: 123456")
                    st.rerun()
        else:
            entered = st.text_input("Enter 6-Digit OTP", type="password")
            bypass = str(st.secrets.get("MASTER_BYPASS_OTP", "123456"))
            if st.button("Verify & Sign In", type="primary", use_container_width=True):
                if entered.strip() in [st.session_state["generated_otp"], bypass]:
                    st.session_state["authenticated"] = True
                    st.rerun()
                else: st.error("Invalid Code.")

# ==============================================================================
# 4. PARSER & GEMINI AUDIT ENGINE
# ==============================================================================
def clean_parse_json(text: str):
    t = re.sub(r"^```(?:json)?\s*", "", text.strip(), flags=re.MULTILINE)
    t = re.sub(r"\s*```$", "", t, flags=re.MULTILINE).strip()
    s_brace, s_brack = t.find("{"), t.find("[")
    if s_brace != -1 and (s_brack == -1 or s_brace < s_brack): t = t[s_brace : t.rfind("}") + 1]
    elif s_brack != -1: t = t[s_brack : t.rfind("]") + 1]
    return json.loads(t)

def extract_text(file_path) -> str:
    content = ""
    if file_path.endswith(".docx"):
        doc = docx.Document(file_path)
        content = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
    else:
        with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
            content = f.read()
    return content

def run_gemini_audit(manuscript_text: str):
    api_key = st.secrets.get("GEMINI_API_KEY")
    if not api_key: st.stop()
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel("gemini-1.5-pro", generation_config={"temperature": 0.0, "response_mime_type": "application/json"})
    
    prompt = f"""
    Audit this academic manuscript strictly against UGC-DEB SLM quality benchmarks.
    Check for mid-sentence truncations (e.g. 'Amazon Dyna...'). If found, deduct heavily and log a Critical defect.
    Return JSON only. Format: {{"matrix": [{{"param_no":1, "dimension":"Syllabus", "weight":10, "score_achieved":8.5, "status":"COMPLIANT", "auditor_findings":"...", "remediation_guidance":"..."}}], "defects": [...]}}
    Text: {manuscript_text[:35000]}
    """
    try: return clean_parse_json(model.generate_content(prompt).text)
    except Exception as e:
        st.error(f"Audit Engine Failed: {e}")
        st.stop()

# ==============================================================================
# 5. DOSSIER EXPORT GENERATORS
# ==============================================================================
def generate_docx_report(a_df, d_df, meta):
    doc = docx.Document()
    doc.add_heading("ACADEMIC CONTENT QUALITY REVIEW REPORT", level=1)
    doc.add_paragraph(f"University: {meta['uni']} | Program: {meta['prog']} | Score: {meta['score']:.1f}/100 | Result: {meta['det']}")
    
    doc.add_heading("12-Parameter Quality Matrix", level=2)
    t = doc.add_table(rows=1, cols=6, style="Table Grid")
    for i, h in enumerate(["#", "Dimension", "Weight", "Score", "Status", "Findings"]): t.cell(0,i).text = h
    for _, r in a_df.iterrows():
        row = t.add_row().cells
        for i, val in enumerate(r): row[i].text = str(val)
        
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()

def generate_xlsx_report(a_df, d_df, meta):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        a_df.to_excel(writer, sheet_name="Scorecard", index=False)
        if not d_df.empty: d_df.to_excel(writer, sheet_name="Defect Register", index=False)
        pd.DataFrame([meta]).to_excel(writer, sheet_name="Summary", index=False)
    return buf.getvalue()

# ==============================================================================
# 6. MAIN APPLICATION ROUTING & WORKFLOW TABS
# ==============================================================================
def main():
    if not st.session_state["authenticated"]:
        render_login_screen()
        return

    catalog = load_json(CATALOG_FILE, {"universities": [{"name": "State Open University", "programs": [{"name": "MBA", "semesters": ["Term 1"]}]}]})
    roles = load_json(ROLES_FILE, DEFAULT_ROLES)
    queue = load_queue()

    # --- SIDEBAR: SELF-SERVE AFFILIATION ---
    with st.sidebar:
        st.markdown(f"### 👤 {st.session_state['user_role']}\n{st.session_state['user_email']}")
        st.divider()
        st.subheader("Workspace Affiliation")
        
        uni_names = [u["name"] for u in catalog.get("universities", [])]
        selected_uni = st.selectbox("University", uni_names)
        
        prog_list = next((u.get("programs", []) for u in catalog.get("universities", []) if u["name"] == selected_uni), [])
        prog_names = [p["name"] for p in prog_list] if prog_list else ["Default Program"]
        selected_prog = st.selectbox("Program", prog_names)

        sem_list = next((p.get("semesters", []) for p in prog_list if p["name"] == selected_prog), []) if prog_list else ["Default Term"]
        selected_sem = st.selectbox("Semester/Batch", sem_list)

        with st.expander("➕ Add New Affiliation (Self-Serve)"):
            n_u, n_p, n_s = st.text_input("New University"), st.text_input("New Program"), st.text_input("New Semester")
            if st.button("Save New Affiliation", use_container_width=True):
                if n_u and n_u not in uni_names: catalog.setdefault("universities", []).append({"name": n_u, "programs": []})
                if n_u and n_p:
                    t_u = next(u for u in catalog["universities"] if u["name"] == n_u)
                    if n_p not in [p["name"] for p in t_u.get("programs", [])]: t_u["programs"].append({"name": n_p, "semesters": []})
                if n_u and n_p and n_s:
                    t_p = next(p for p in next(u for u in catalog["universities"] if u["name"] == n_u)["programs"] if p["name"] == n_p)
                    if n_s not in t_p.get("semesters", []): t_p["semesters"].append(n_s)
                save_json(CATALOG_FILE, catalog)
                st.rerun()

        st.divider()
        if st.button("Sign Out", use_container_width=True):
            for k in list(st.session_state.keys()): del st.session_state[k]
            st.rerun()

    st.title("Academic SLM Maker-Checker QA Platform")
    
    tabs = st.tabs(["📤 Submit Content (Maker)", "📥 Audit Queue (Checker)", "📊 Matrix & Defects", "✅ Sign-Off & Export", "📈 Platform Analytics", "🛡️ RBAC Admin"])

    # --- TAB 1: MAKER SUBMISSION ---
    with tabs[0]:
        st.markdown("#### Upload Content for Auditor Review")
        uploaded_file = st.file_uploader("Upload Manuscript", type=["docx", "txt"])
        if uploaded_file and st.button("Submit to Checker Queue", type="primary"):
            file_id = str(uuid.uuid4())[:8]
            save_path = os.path.join(UPLOAD_DIR, f"{file_id}_{uploaded_file.name}")
            with open(save_path, "wb") as f: f.write(uploaded_file.getbuffer())
            
            queue.append({
                "id": file_id, "maker": st.session_state["user_email"], "uni": selected_uni, "prog": selected_prog, "sem": selected_sem,
                "filename": uploaded_file.name, "filepath": save_path, "status": "Pending Review", "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M")
            })
            save_queue(queue)
            st.success("✅ Document submitted to the Audit Queue successfully.")

        st.divider()
        st.markdown("##### My Submissions")
        my_subs = [q for q in queue if q["maker"] == st.session_state["user_email"]]
        if my_subs: st.dataframe(pd.DataFrame(my_subs)[["filename", "status", "timestamp", "uni", "prog"]], hide_index=True, use_container_width=True)

    # --- TAB 2: CHECKER QUEUE ---
    with tabs[1]:
        st.markdown("#### Pending Submissions for Audit")
        pending = [q for q in queue if q["status"] == "Pending Review"]
        if pending:
            p_df = pd.DataFrame(pending)
            st.dataframe(p_df[["id", "maker", "filename", "timestamp"]], hide_index=True, use_container_width=True)
            
            sel_id = st.selectbox("Select Submission to Audit", p_df["id"].tolist())
            if st.button("▶️ Run Automated QA Engine", type="primary"):
                target = next(q for q in pending if q["id"] == sel_id)
                with st.spinner(f"Auditing {target['filename']}..."):
                    m_text = extract_text(target["filepath"])
                    raw_audit = run_gemini_audit(m_text)
                    
                    # Process Results
                    a_rows = []
                    for p in raw_audit.get("matrix", []):
                        a_rows.append([p.get("param_no", 1), p.get("dimension", ""), p.get("weight", 10), p.get("score_achieved", 10), p.get("status", ""), p.get("auditor_findings", "")])
                    st.session_state["audit_df"] = pd.DataFrame(a_rows, columns=["#", "Dimension", "Weight", "Score", "Status", "Findings"])
                    st.session_state["defect_df"] = pd.DataFrame(raw_audit.get("defects", []))
                    st.session_state["quality_score"] = st.session_state["audit_df"]["Score"].sum()
                    st.session_state["active_audit_id"] = sel_id
                    st.session_state["audit_completed"] = True
                    st.success("Audit complete. Proceed to Matrix and Sign-Off tabs.")
        else:
            st.info("No pending documents in the queue.")

    # --- TAB 3: MATRIX & DEFECTS ---
    with tabs[2]:
        if st.session_state["audit_completed"]:
            st.markdown("#### 12-Parameter Evaluation Matrix")
            st.dataframe(st.session_state["audit_df"], hide_index=True, use_container_width=True)
            st.markdown("#### Logged Defects")
            if not st.session_state["defect_df"].empty: st.dataframe(st.session_state["defect_df"], hide_index=True, use_container_width=True)
            else: st.success("No defects found.")
        else: st.info("Run an audit from the Checker Queue to view the matrix.")

    # --- TAB 4: CHECKER SIGN-OFF ---
    with tabs[3]:
        if st.session_state["audit_completed"]:
            st.markdown("#### Final Lead Auditor Sign-Off")
            auto_det = "REVISE" if st.session_state["quality_score"] < 75 else "PASS"
            override = st.selectbox("Final Determination", ["PASS", "REVISE", "REJECT"], index=["PASS", "REVISE", "REJECT"].index(auto_det))
            
            if st.button("Complete Sign-Off & Log to Ledger", type="primary"):
                # Update Queue Status
                for q in queue:
                    if q.get("id") == st.session_state.get("active_audit_id"):
                        q["status"] = "Signed Off"
                        q["score"] = st.session_state["quality_score"]
                        q["determination"] = override
                        q["checker"] = st.session_state["user_email"]
                        
                        # Write to Permanent Ledger
                        append_to_ledger({
                            "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "Maker": q["maker"], "Checker": q["checker"],
                            "University": q["uni"], "Program": q["prog"], "Semester": q["sem"],
                            "File Name": q["filename"], "Quality Score": q["score"], "Determination": override
                        })
                        break
                save_queue(queue)
                st.success("Ledger updated. Exports unlocked.")
                
                # Exports
                meta = {"uni": selected_uni, "prog": selected_prog, "score": st.session_state["quality_score"], "det": override}
                c1, c2 = st.columns(2)
                c1.download_button("📄 Download DOCX Dossier", generate_docx_report(st.session_state["audit_df"], st.session_state["defect_df"], meta), file_name="Report.docx", use_container_width=True)
                c2.download_button("📊 Download XLSX Scorecard", generate_xlsx_report(st.session_state["audit_df"], st.session_state["defect_df"], meta), file_name="Report.xlsx", use_container_width=True)

    # --- TAB 5: PLATFORM ANALYTICS ---
    with tabs[4]:
        st.markdown("#### 📈 Global Platform Audit Ledger")
        ledger_df = load_ledger()
        if not ledger_df.empty:
            c1, c2, c3 = st.columns(3)
            c1.metric("Total Documents Processed", len(ledger_df))
            c2.metric("Average Quality Score", f"{ledger_df['Quality Score'].mean():.1f}/100")
            c3.metric("Platform Pass Rate", f"{(len(ledger_df[ledger_df['Determination'] == 'PASS']) / len(ledger_df)) * 100:.1f}%")
            
            st.dataframe(ledger_df.sort_values("Timestamp", ascending=False), hide_index=True, use_container_width=True)
            st.download_button("📥 Download Raw Audit History (CSV)", ledger_df.to_csv(index=False).encode("utf-8"), "audit_history.csv")
        else: st.info("Ledger is empty. Process documents to generate analytics.")

    # --- TAB 6: ADMIN GOVERNANCE ---
    with tabs[5]:
        if st.session_state["user_role"] == "Admin":
            st.markdown("#### 🛡️ User Authority Assignment")
            new_mail, new_role = st.text_input("User Email"), st.selectbox("Role", ["Maker", "Checker", "Admin"])
            if st.button("Assign Role"):
                roles[new_mail.strip().lower()] = new_role
                save_json(ROLES_FILE, roles)
                st.success("User updated.")
            st.dataframe(pd.DataFrame([{"Email": k, "Role": v} for k, v in roles.items()]), hide_index=True)
        else: st.error("Access Denied. Administrator rights required.")

if __name__ == "__main__":
    main()