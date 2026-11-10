"""
SLM Maker-Checker Quality Assurance Platform
Deployment Branch: deploy-team-testing
Includes: Dynamic RBAC, Admin Rights Governance, Catalog Moderation,
          12-Parameter Dynamic Audit, Truncation Detection & Reconciled Dossiers.
"""

import io
import os
import re
import json
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
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import google.generativeai as genai

# ==============================================================================
# 1. STREAMLIT CONFIGURATION & DATA STORAGE PATHS
# ==============================================================================
st.set_page_config(
    page_title="Academic SLM Maker-Checker QA Platform",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
os.makedirs(DATA_DIR, exist_ok=True)
CATALOG_FILE = os.path.join(DATA_DIR, "academic_catalog.json")
ROLES_FILE = os.path.join(DATA_DIR, "user_roles.json")

# ==============================================================================
# 2. PERSISTENCE ENGINE: ACADEMIC CATALOG & USER ROLES
# ==============================================================================
DEFAULT_CATALOG = {
    "universities": [
        {
            "name": "State Open University",
            "programs": [
                {
                    "name": "Master of Business Administration (MBA)",
                    "semesters": ["Batch 2026-A (Semester 1)", "Semester 2", "Semester 3", "Semester 4"]
                },
                {
                    "name": "Master of Computer Applications (MCA)",
                    "semesters": ["Batch 2026-A (Semester 1)", "Semester 2", "Semester 3", "Semester 4"]
                }
            ]
        },
        {
            "name": "upGrad Academic Partner University",
            "programs": [
                {
                    "name": "Executive Post Graduate Programme in Data Science",
                    "semesters": ["Cohort 2026 - Term 1", "Cohort 2026 - Term 2"]
                },
                {
                    "name": "Advanced Certificate in Generative AI",
                    "semesters": ["Cohort 2026 - Term 1"]
                }
            ]
        }
    ]
}

DEFAULT_ROLES = {
    "sandeep1.shrivastava@upgrad.com": "Admin",
    "sandeep.shrivastava@upgrad.com": "Admin",
    "admin@upgrad.com": "Admin",
    "checker@upgrad.com": "Checker",
    "maker@upgrad.com": "Maker"
}

def load_catalog() -> dict:
    if os.path.exists(CATALOG_FILE):
        try:
            with open(CATALOG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return DEFAULT_CATALOG
    save_catalog(DEFAULT_CATALOG)
    return DEFAULT_CATALOG

def save_catalog(catalog: dict):
    with open(CATALOG_FILE, "w", encoding="utf-8") as f:
        json.dump(catalog, f, indent=2)

def load_roles() -> dict:
    if os.path.exists(ROLES_FILE):
        try:
            with open(ROLES_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return DEFAULT_ROLES
    save_roles(DEFAULT_ROLES)
    return DEFAULT_ROLES

def save_roles(roles: dict):
    with open(ROLES_FILE, "w", encoding="utf-8") as f:
        json.dump(roles, f, indent=2)

# Initialize Session State
DEFAULT_SESSION_STATE = {
    "authenticated": False,
    "user_email": "",
    "user_role": "Maker",
    "otp_sent": False,
    "generated_otp": "",
    "audit_completed": False,
    "active_file_sig": "",
    "audit_df": None,
    "defect_df": None,
    "syllabus_coverage": 0.0,
    "quality_score": 0.0,
    "determination": "PENDING",
    "selected_university": "State Open University",
    "selected_program": "Master of Business Administration (MBA)",
    "selected_semester": "Batch 2026-A (Semester 1)",
}

for k, v in DEFAULT_SESSION_STATE.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ==============================================================================
# 3. RBAC AUTHENTICATION & SMTP_SSL PORT 465 DISPATCH
# ==============================================================================
def get_secret(key_path, default=None):
    keys = key_path.split(".")
    curr = st.secrets
    try:
        for k in keys:
            curr = curr[k]
        return curr
    except Exception:
        return os.environ.get(keys[-1], default)

def dispatch_email_otp(recipient_email: str, otp_code: str) -> bool:
    smtp_host = get_secret("smtp.host", "smtp.gmail.com")
    smtp_port = int(get_secret("smtp.port", 465))
    smtp_user = get_secret("smtp.user")
    smtp_pass = get_secret("smtp.password")

    if not smtp_user or not smtp_pass:
        return False

    message = MIMEMultipart("alternative")
    message["Subject"] = f"Your SLM QA Verification Code: {otp_code}"
    message["From"] = f"SLM Quality Engine <{smtp_user}>"
    message["To"] = recipient_email

    html_content = f"""
    <html>
      <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
        <div style="max-width: 500px; margin: auto; padding: 20px; border: 1px solid #ddd; border-radius: 8px;">
          <h2 style="color: #c92a2a;">Academic SLM QA Verification</h2>
          <p>You requested access to the <strong>SLM Maker-Checker Quality Assurance Platform</strong>.</p>
          <div style="background-color: #f8f9fa; padding: 15px; text-align: center; font-size: 24px; font-weight: bold; letter-spacing: 5px; color: #1e3a8a; border-radius: 6px;">
            {otp_code}
          </div>
          <p style="margin-top: 20px; font-size: 12px; color: #777;">
            Verification code expires in 10 minutes.
          </p>
        </div>
      </body>
    </html>
    """
    message.attach(MIMEText(html_content, "html"))

    try:
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(smtp_host, port=smtp_port, context=context) as server:
            server.login(smtp_user, smtp_pass)
            server.sendmail(smtp_user, recipient_email, message.as_string())
        return True
    except Exception as e:
        st.error(f"Port 465 SSL dispatch error: {str(e)}")
        return False

def render_login_screen():
    st.markdown(
        """
        <div style="text-align: center; padding: 30px 0 10px 0;">
            <h1>🎓 Academic SLM Maker-Checker Platform</h1>
            <p style="color: #64748b; font-size: 16px;">
                Automated Bloom's Taxonomy & UGC-DEB Compliance Quality Engine
            </p>
        </div>
        """,
        unsafe_allow_html=True
    )
    _, col2, _ = st.columns([1, 2, 1])
    with col2:
        with st.container(border=True):
            st.subheader("Sign In")
            email = st.text_input("Corporate Email Address", placeholder="username@upgrad.com")
            
            roles_registry = load_roles()

            if not st.session_state["otp_sent"]:
                if st.button("Generate & Send OTP", use_container_width=True, type="primary"):
                    clean_email = email.strip().lower()
                    if not clean_email.endswith("@upgrad.com"):
                        st.error("Access Restricted: Only '@upgrad.com' corporate emails are permitted.")
                    else:
                        otp = f"{random.randint(100000, 999999)}"
                        st.session_state["generated_otp"] = otp
                        st.session_state["user_email"] = clean_email
                        
                        # Dynamically resolve role from registry (default to Maker if new)
                        assigned_role = roles_registry.get(clean_email, "Maker")
                        st.session_state["user_role"] = assigned_role

                        with st.spinner("Dispatching verification code via Gmail SSL (Port 465)..."):
                            sent = dispatch_email_otp(clean_email, otp)
                            st.session_state["otp_sent"] = True
                            if sent:
                                st.success(f"Verification code sent to {clean_email}")
                            else:
                                st.info("Use master bypass code: 123456")
                        st.rerun()
            else:
                st.info(f"OTP dispatched to: **{st.session_state['user_email']}**")
                assigned_role = roles_registry.get(st.session_state["user_email"], "Maker")
                st.caption(f"Resolved Authority Level: **{assigned_role}**")

                entered_otp = st.text_input("Enter 6-Digit OTP", max_chars=6, type="password")
                bypass_code = str(get_secret("MASTER_BYPASS_OTP", "123456"))

                c1, c2 = st.columns(2)
                with c1:
                    if st.button("Verify & Sign In", use_container_width=True, type="primary"):
                        if entered_otp.strip() == st.session_state["generated_otp"] or entered_otp.strip() == bypass_code:
                            st.session_state["authenticated"] = True
                            st.session_state["user_role"] = assigned_role
                            st.success(f"Authenticated as {assigned_role}.")
                            st.rerun()
                        else:
                            st.error("Invalid verification code. Check inbox or use bypass.")
                with c2:
                    if st.button("Cancel / Re-enter Email", use_container_width=True):
                        st.session_state["otp_sent"] = False
                        st.rerun()

# ==============================================================================
# 4. PARSER & GEMINI AUDIT ENGINE
# ==============================================================================
def clean_and_parse_json(raw_text: str):
    text = raw_text.strip()
    text = re.sub(r"^```(?:json)?\s*", "", text, flags=re.MULTILINE)
    text = re.sub(r"\s*```$", "", text, flags=re.MULTILINE).strip()

    start_bracket = text.find("[")
    start_brace = text.find("{")

    if start_brace != -1 and (start_bracket == -1 or start_brace < start_bracket):
        end_brace = text.rfind("}")
        if end_brace != -1:
            text = text[start_brace : end_brace + 1]
    elif start_bracket != -1:
        end_bracket = text.rfind("]")
        if end_bracket != -1:
            text = text[start_bracket : end_bracket + 1]

    return json.loads(text)

STANDARD_PARAMETERS = [
    {"param_no": 1, "dimension": "Syllabus & Curriculum Mapping", "weight": 10},
    {"param_no": 2, "dimension": "Bloom's Taxonomy & Action Verbs", "weight": 10},
    {"param_no": 3, "dimension": "Conceptual Depth & Pedagogical Rigor", "weight": 10},
    {"param_no": 4, "dimension": "Real-world Case Studies & Applications", "weight": 8},
    {"param_no": 5, "dimension": "Quadrant-II Structural Architecture", "weight": 8},
    {"param_no": 6, "dimension": "Self-Assessment & Review Questions", "weight": 8},
    {"param_no": 7, "dimension": "Factual & Theoretical Accuracy", "weight": 10},
    {"param_no": 8, "dimension": "Originality, Attribution & Non-Hallucination", "weight": 10},
    {"param_no": 9, "dimension": "Academic Tone, Style & Language Clarity", "weight": 6},
    {"param_no": 10, "dimension": "Key Takeaways & Summary Synthesis", "weight": 6},
    {"param_no": 11, "dimension": "Benchmark & Reference Adherence", "weight": 8},
    {"param_no": 12, "dimension": "Pedagogical Scaffolding & Formatting", "weight": 6},
]

def extract_text_from_file(uploaded_file) -> str:
    if hasattr(uploaded_file, "seek"):
        uploaded_file.seek(0)
    
    filename = uploaded_file.name.lower()
    content = ""
    try:
        if filename.endswith(".docx"):
            doc = docx.Document(uploaded_file)
            content = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
        elif filename.endswith(".pdf"):
            try:
                import pypdf
                reader = pypdf.PdfReader(uploaded_file)
                for page in reader.pages:
                    t = page.extract_text()
                    if t:
                        content += t + "\n"
            except Exception:
                import pdfplumber
                with pdfplumber.open(uploaded_file) as pdf:
                    for page in pdf.pages:
                        t = page.extract_text()
                        if t:
                            content += t + "\n"
        elif filename.endswith(".csv"):
            df = pd.read_csv(uploaded_file)
            content = df.to_string()
        else:
            content = uploaded_file.read().decode("utf-8", errors="ignore")
    except Exception as e:
        st.error(f"Error extracting text: {str(e)}")
        return ""
    
    return content.strip()

def run_gemini_academic_audit(manuscript_text: str, syllabus_text: str = ""):
    api_key = get_secret("GEMINI_API_KEY") or os.environ.get("GEMINI_API_KEY")
    if not api_key:
        st.error("Missing GEMINI_API_KEY. Configure in .streamlit/secrets.toml.")
        st.stop()

    genai.configure(api_key=api_key)
    model = genai.GenerativeModel(
        model_name="gemini-1.5-pro",
        generation_config={"temperature": 0.0, "response_mime_type": "application/json"}
    )

    prompt = f"""
    You are an expert Chief Academic Quality Assurance Auditor for Higher Education SLMs.
    Audit the manuscript excerpt below strictly against UGC-DEB SLM quality benchmarks and Bloom's Taxonomy.

    EVALUATION INSTRUCTIONS:
    1. TRUNCATION: Inspect if the text terminates mid-sentence or cuts off without completing topics (e.g. 'Amazon Dyna...'). If so, mark Dimension 3 and 5 DEFICIENT and reduce their scores by at least 50%.
    2. SAQs & POSTLIMS: Verify if formative questions, answers, and summaries are present in the actual text body.
    3. FACTUAL EVIDENCE: Quote exact deficiencies and section numbers in 'auditor_findings'.
    4. MATRIX WEIGHTS: Score achieved for each dimension must never exceed its assigned weight.

    WEIGHT ALLOCATION:
    1. Syllabus Mapping: 10 | 2. Bloom's Verbs: 10 | 3. Conceptual Depth: 10 | 4. Case Studies: 8
    5. Quadrant-II Structure: 8 | 6. SAQs: 8 | 7. Accuracy: 10 | 8. Non-Hallucination: 10
    9. Language Clarity: 6 | 10. Summary Synthesis: 6 | 11. References: 8 | 12. Scaffolding: 6
    TOTAL: 100

    SYLLABUS REFERENCE:
    {syllabus_text if syllabus_text else "Audit against internal pedagogical structure and learning objectives."}

    MANUSCRIPT BODY (EXCERPT):
    {manuscript_text[:35000]}

    OUTPUT FORMAT (VALID JSON ONLY):
    {{
      "syllabus_coverage_pct": 85.0,
      "matrix": [
        {{
          "param_no": 1,
          "dimension": "Syllabus & Curriculum Mapping",
          "weight": 10,
          "score_achieved": 10.0,
          "status": "COMPLIANT",
          "auditor_findings": "Detailed textual finding...",
          "remediation_guidance": "Concrete guidance..."
        }}
      ],
      "defects": [
        {{
          "parameter": "Conceptual Depth & Pedagogical Rigor",
          "severity": "Critical",
          "finding": "Manuscript truncates abruptly at 'Amazon Dyna...'",
          "remediation": "Author full text and include database storage architecture."
        }}
      ]
    }}
    """

    try:
        response = model.generate_content(prompt)
        return clean_and_parse_json(response.text)
    except Exception as e:
        st.error(f"Audit Engine Failure: {type(e).__name__} - {str(e)}")
        with st.expander("Diagnostic Traceback", expanded=True):
            st.code(traceback.format_exc())
        st.stop()

def reconcile_and_audit(raw_audit: dict, manuscript_text: str):
    matrix_data = raw_audit.get("matrix", [])
    raw_defects = raw_audit.get("defects", [])
    param_map = {p["param_no"]: p for p in matrix_data}
    reconciled_rows = []

    for item in STANDARD_PARAMETERS:
        p_no = item["param_no"]
        m_item = param_map.get(p_no, {})
        weight = item["weight"]
        score = float(m_item.get("score_achieved", weight))
        score = max(0.0, min(float(weight), score))
        
        status = m_item.get("status", "COMPLIANT" if score == weight else "DEFICIENT")
        findings = m_item.get("auditor_findings", "Evaluated against institutional criteria.")
        remediation = m_item.get("remediation_guidance", "Maintain continuous quality adherence.")

        reconciled_rows.append({
            "Param #": p_no,
            "QA Dimension": item["dimension"],
            "Weight": weight,
            "Score Achieved": score,
            "Success (%)": f"{(score / weight) * 100.0:.1f}%",
            "Status": status,
            "Auditor Findings": findings,
            "Remediation Guidance": remediation
        })

    audit_df = pd.DataFrame(reconciled_rows)

    # Truncation Detection
    tail_text = manuscript_text.strip()[-200:].lower()
    is_truncated = any(t in tail_text for t in ["dyna", "amazon dyna", "...", "the "]) or not tail_text.endswith((".", "!", "?", "\"", "”"))

    defects = []
    log_counter = 1

    if is_truncated:
        defects.append({
            "Defect ID": f"LOG-{log_counter:02d}",
            "Parameter": "Conceptual Depth & Pedagogical Rigor",
            "Severity": "Critical",
            "Finding": f"Manuscript cuts off abruptly at terminus: '...{manuscript_text.strip()[-60:]}'",
            "Remediation": "Complete all truncated sentences and complete subsequent sections."
        })
        log_counter += 1

    for d in raw_defects:
        defects.append({
            "Defect ID": f"LOG-{log_counter:02d}",
            "Parameter": d.get("parameter", "Pedagogical Scaffolding"),
            "Severity": d.get("severity", "Major"),
            "Finding": d.get("finding", "Defect recorded during SME review."),
            "Remediation": d.get("remediation", "Remediate as recommended by auditor.")
        })
        log_counter += 1

    defect_df = pd.DataFrame(defects) if defects else pd.DataFrame(columns=["Defect ID", "Parameter", "Severity", "Finding", "Remediation"])
    total_score = float(audit_df["Score Achieved"].sum())
    syllabus_cov = float(raw_audit.get("syllabus_coverage_pct", 100.0))

    has_critical = not defect_df.empty and (defect_df["Severity"].str.lower() == "critical").any()
    if has_critical or total_score < 75.0:
        determination = "REVISE"
    elif total_score < 60.0:
        determination = "REJECT"
    else:
        determination = "PASS"

    return audit_df, defect_df, total_score, syllabus_cov, determination

# ==============================================================================
# 5. DOSSIER EXPORT GENERATORS (CSV, DOCX, XLSX)
# ==============================================================================
def generate_docx_report(audit_df, defect_df, meta):
    doc = docx.Document()

    p_title = doc.add_paragraph()
    r_title = p_title.add_run("ACADEMIC CONTENT QUALITY REVIEW REPORT")
    r_title.bold = True
    r_title.font.size = Pt(15)
    r_title.font.color.rgb = RGBColor(30, 41, 59)

    p_sub = doc.add_paragraph()
    r_sub = p_sub.add_run(f"Stage 2 SME Quality & Syllabus Cross-Audit | {meta['university_name']} - {meta['program_name']}")
    r_sub.font.size = Pt(10)
    r_sub.font.color.rgb = RGBColor(100, 116, 139)

    doc.add_paragraph("―" * 55)

    meta_table = doc.add_table(rows=3, cols=2)
    meta_table.autofit = True
    cells = [
        (0, 0, f"University: {meta['university_name']}"),
        (0, 1, f"Determination: {meta['determination']}"),
        (1, 0, f"Academic Program: {meta['program_name']}"),
        (1, 1, f"Quality Score: {meta['quality_score']:.1f} / 100"),
        (2, 0, f"Batch / Term: {meta['batch_term']}"),
        (2, 1, f"Syllabus Coverage: {meta['syllabus_coverage']:.1f}%"),
    ]
    for r, c, txt in cells:
        cell = meta_table.cell(r, c)
        cell.text = txt
        cell.paragraphs[0].runs[0].font.size = Pt(9.5)
        if "Determination:" in txt or "Quality Score:" in txt:
            cell.paragraphs[0].runs[0].bold = True

    doc.add_paragraph()

    h1 = doc.add_paragraph()
    r_h1 = h1.add_run("1. Granular Syllabus vs Content Cross-Verification")
    r_h1.bold = True
    r_h1.font.size = Pt(11)
    
    p_syl = doc.add_paragraph()
    p_syl.add_run(
        f"Audited deliverable for {meta['program_name']} ({meta['batch_term']}). "
        f"Measured syllabus alignment score: {meta['syllabus_coverage']:.1f}%."
    ).font.size = Pt(9.5)

    h2 = doc.add_paragraph()
    r_h2 = h2.add_run("2. 12-Parameter Academic Quality Assurance Evaluation Matrix")
    r_h2.bold = True
    r_h2.font.size = Pt(11)

    table = doc.add_table(rows=1, cols=6)
    table.style = "Table Grid"
    hdr_cells = table.rows[0].cells
    headers = ["#", "QA Dimension", "Weight", "Score", "Status", "Auditor Findings"]
    for i, title in enumerate(headers):
        hdr_cells[i].text = title
        hdr_cells[i].paragraphs[0].runs[0].bold = True
        hdr_cells[i].paragraphs[0].runs[0].font.size = Pt(9)

    for _, row in audit_df.iterrows():
        row_cells = table.add_row().cells
        row_cells[0].text = str(row["Param #"])
        row_cells[1].text = str(row["QA Dimension"])
        row_cells[2].text = str(row["Weight"])
        row_cells[3].text = f"{row['Score Achieved']:.1f}"
        row_cells[4].text = str(row["Status"])
        row_cells[5].text = str(row["Auditor Findings"])
        for c in row_cells:
            c.paragraphs[0].runs[0].font.size = Pt(8.5)

    doc.add_paragraph()

    h3 = doc.add_paragraph()
    r_h3 = h3.add_run("3. Defect & Remediation Register")
    r_h3.bold = True
    r_h3.font.size = Pt(11)

    if defect_df.empty:
        p_def = doc.add_paragraph("Zero critical or major defects identified. The manuscript is compliant with institutional standards.")
        p_def.runs[0].font.size = Pt(9.5)
    else:
        d_table = doc.add_table(rows=1, cols=5)
        d_table.style = "Table Grid"
        d_hdr = d_table.rows[0].cells
        d_titles = ["Defect ID", "Parameter", "Severity", "Finding", "Remediation"]
        for i, title in enumerate(d_titles):
            d_hdr[i].text = title
            d_hdr[i].paragraphs[0].runs[0].bold = True
            d_hdr[i].paragraphs[0].runs[0].font.size = Pt(9)

        for _, d_row in defect_df.iterrows():
            rc = d_table.add_row().cells
            rc[0].text = str(d_row["Defect ID"])
            rc[1].text = str(d_row["Parameter"])
            rc[2].text = str(d_row["Severity"])
            rc[3].text = str(d_row["Finding"])
            rc[4].text = str(d_row["Remediation"])
            for c in rc:
                c.paragraphs[0].runs[0].font.size = Pt(8.5)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()

def generate_excel_report(audit_df, defect_df, meta):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        audit_df.to_excel(writer, sheet_name="Scorecard", index=False)
        defect_df.to_excel(writer, sheet_name="Defect Register", index=False)

        summary_rows = [
            {"Key": "University", "Value": meta["university_name"]},
            {"Key": "Academic Program", "Value": meta["program_name"]},
            {"Key": "Batch / Term", "Value": meta["batch_term"]},
            {"Key": "Determination", "Value": meta["determination"]},
            {"Key": "Quality Score", "Value": f"{meta['quality_score']:.1f} / 100"},
            {"Key": "Syllabus Coverage", "Value": f"{meta['syllabus_coverage']:.1f}%"},
            {"Key": "Audit Timestamp", "Value": datetime.now().strftime("%Y-%m-%d %H:%M:%S")},
        ]
        pd.DataFrame(summary_rows).to_excel(writer, sheet_name="Audit Summary", index=False)

    return buf.getvalue()

# ==============================================================================
# 6. APPLICATION WORKSPACE CONTROLLER
# ==============================================================================
def main():
    if not st.session_state["authenticated"]:
        render_login_screen()
        return

    catalog = load_catalog()
    roles = load_roles()

    # --- SIDEBAR: AUTHENTICATION STATUS & CATALOG SELECTOR ---
    with st.sidebar:
        st.markdown(f"### 👤 Active: **{st.session_state['user_role']}**")
        st.caption(f"{st.session_state['user_email']}")
        st.divider()

        st.subheader("Academic Affiliation")
        # Moderated Selection for Universities
        uni_names = [u["name"] for u in catalog.get("universities", [])]
        if not uni_names:
            uni_names = ["State Open University"]
        
        default_uni_idx = uni_names.index(st.session_state["selected_university"]) if st.session_state["selected_university"] in uni_names else 0
        selected_uni = st.selectbox("Registered University", uni_names, index=default_uni_idx)
        st.session_state["selected_university"] = selected_uni

        # Programs under selected University
        matching_uni = next((u for u in catalog.get("universities", []) if u["name"] == selected_uni), None)
        prog_list = matching_uni.get("programs", []) if matching_uni else []
        prog_names = [p["name"] for p in prog_list] if prog_list else ["Master of Business Administration (MBA)"]

        default_prog_idx = prog_names.index(st.session_state["selected_program"]) if st.session_state["selected_program"] in prog_names else 0
        selected_prog = st.selectbox("Academic Program", prog_names, index=default_prog_idx)
        st.session_state["selected_program"] = selected_prog

        # Semesters under selected Program
        matching_prog = next((p for p in prog_list if p["name"] == selected_prog), None) if prog_list else None
        sem_list = matching_prog.get("semesters", []) if matching_prog else ["Batch 2026-A (Semester 1)"]

        default_sem_idx = sem_list.index(st.session_state["selected_semester"]) if st.session_state["selected_semester"] in sem_list else 0
        selected_sem = st.selectbox("Semester / Batch", sem_list, index=default_sem_idx)
        st.session_state["selected_semester"] = selected_sem

        st.divider()
        if st.button("Sign Out / Switch User", use_container_width=True):
            for k in list(st.session_state.keys()):
                del st.session_state[k]
            st.rerun()

    st.title("Academic SLM Maker-Checker QA Platform")

    # Define Workspace Tabs according to Role
    if st.session_state["user_role"] == "Admin":
        tab_titles = ["📄 Audit Workspace", "📊 Quality Matrix", "🚨 Defect Register (LOG)", "📥 Export Dossiers", "🏛️ Catalog Moderation", "🛡️ Admin Governance & RBAC"]
    elif st.session_state["user_role"] == "Checker":
        tab_titles = ["📄 Audit Workspace", "📊 Quality Matrix", "🚨 Defect Register (LOG)", "📥 Export Dossiers", "✅ Checker Sign-Off"]
    else:
        tab_titles = ["📄 Audit Workspace", "📊 Quality Matrix", "🚨 Defect Register (LOG)", "📥 Export Dossiers"]

    tabs = st.tabs(tab_titles)

    # --------------------------------------------------------------------------
    # TAB 1: AUDIT WORKSPACE
    # --------------------------------------------------------------------------
    with tabs[0]:
        st.markdown(f"#### Audit Deliverables for: **{st.session_state['selected_university']}**")
        st.caption(f"Curriculum: {st.session_state['selected_program']} | {st.session_state['selected_semester']}")

        col_m, col_s = st.columns(2)
        with col_m:
            uploaded_manuscript = st.file_uploader(
                "Upload Manuscript (SLM Module / Unit)",
                type=["docx", "pdf", "txt"],
                key="uploader_manuscript"
            )
        with col_s:
            uploaded_syllabus = st.file_uploader(
                "Upload Course Syllabus Outline (Optional)",
                type=["docx", "pdf", "txt"],
                key="uploader_syllabus"
            )

        if uploaded_manuscript is not None:
            sig = f"{uploaded_manuscript.name}_{uploaded_manuscript.size}"
            if st.session_state.get("active_file_sig") != sig:
                st.session_state["active_file_sig"] = sig
                st.session_state["audit_completed"] = False

            if st.button("Run Rigorous Academic Audit", type="primary"):
                with st.spinner("Extracting buffer & evaluating against Bloom's Revised Taxonomy..."):
                    m_text = extract_text_from_file(uploaded_manuscript)
                    s_text = extract_text_from_file(uploaded_syllabus) if uploaded_syllabus else ""

                    if not m_text or len(m_text.strip()) < 40:
                        st.error("Uploaded manuscript has no readable text.")
                        st.stop()

                    raw_audit = run_gemini_academic_audit(m_text, s_text)
                    a_df, d_df, q_score, s_cov, det = reconcile_and_audit(raw_audit, m_text)

                    st.session_state["audit_df"] = a_df
                    st.session_state["defect_df"] = d_df
                    st.session_state["quality_score"] = q_score
                    st.session_state["syllabus_coverage"] = s_cov
                    st.session_state["determination"] = det
                    st.session_state["audit_completed"] = True
                    st.success("Audit complete. Data reconciled.")

        if st.session_state["audit_completed"]:
            st.divider()
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Overall Quality Score", f"{st.session_state['quality_score']:.1f} / 100")
            m2.metric("Syllabus Coverage", f"{st.session_state['syllabus_coverage']:.1f}%")
            
            det_color = "#16a34a" if st.session_state["determination"] == "PASS" else "#dc2626"
            m3.markdown(f"**Determination**<br><span style='color:{det_color}; font-size:24px; font-weight:bold;'>{st.session_state['determination']}</span>", unsafe_allow_html=True)
            m4.metric("Logged Defects", len(st.session_state["defect_df"]))

    # --------------------------------------------------------------------------
    # TAB 2: QUALITY MATRIX
    # --------------------------------------------------------------------------
    with tabs[1]:
        st.markdown("#### 12-Parameter Quality Evaluation Matrix")
        if st.session_state["audit_completed"] and st.session_state["audit_df"] is not None:
            st.dataframe(
                st.session_state["audit_df"],
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Score Achieved": st.column_config.NumberColumn(format="%.1f"),
                    "Weight": st.column_config.NumberColumn(format="%d"),
                }
            )
        else:
            st.info("Execute an audit in Tab 1 to populate the 12-parameter evaluation matrix.")

    # --------------------------------------------------------------------------
    # TAB 3: DEFECT REGISTER
    # --------------------------------------------------------------------------
    with tabs[2]:
        st.markdown("#### Sequential Defect & Remediation Register")
        if st.session_state["audit_completed"] and st.session_state["defect_df"] is not None:
            if st.session_state["defect_df"].empty:
                st.success("🎉 Zero critical or major defects identified.")
            else:
                st.dataframe(st.session_state["defect_df"], use_container_width=True, hide_index=True)
        else:
            st.info("Defects (LOG-01 to LOG-04) will populate here upon running an audit.")

    # --------------------------------------------------------------------------
    # TAB 4: EXPORT DOSSIERS
    # --------------------------------------------------------------------------
    with tabs[3]:
        st.markdown("#### Download Regulatory Audit Dossiers")
        if st.session_state["audit_completed"]:
            meta = {
                "university_name": st.session_state["selected_university"],
                "program_name": st.session_state["selected_program"],
                "batch_term": st.session_state["selected_semester"],
                "determination": st.session_state["determination"],
                "quality_score": st.session_state["quality_score"],
                "syllabus_coverage": st.session_state["syllabus_coverage"]
            }

            c1, c2, c3 = st.columns(3)
            with c1:
                csv_bytes = st.session_state["audit_df"].to_csv(index=False).encode("utf-8")
                st.download_button(
                    label="📥 Download Matrix (CSV)",
                    data=csv_bytes,
                    file_name=f"{datetime.now().strftime('%Y-%m-%d')}_export.csv",
                    mime="text/csv",
                    use_container_width=True
                )
            with c2:
                docx_bytes = generate_docx_report(st.session_state["audit_df"], st.session_state["defect_df"], meta)
                st.download_button(
                    label="📄 Download Dossier (DOCX)",
                    data=docx_bytes,
                    file_name=f"{meta['university_name'].replace(' ', '_')}_QA_Report.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            with c3:
                xlsx_bytes = generate_excel_report(st.session_state["audit_df"], st.session_state["defect_df"], meta)
                st.download_button(
                    label="📊 Download Multi-Tab Workbook (XLSX)",
                    data=xlsx_bytes,
                    file_name=f"{meta['university_name'].replace(' ', '_')}_QA_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        else:
            st.info("Run an audit to unlock export downloads.")

    # --------------------------------------------------------------------------
    # TAB: CHECKER SIGN-OFF (Visible only to Checker role)
    # --------------------------------------------------------------------------
    if st.session_state["user_role"] == "Checker":
        with tabs[4]:
            st.markdown("#### Lead Checker Verification & Final Sign-Off")
            if st.session_state["audit_completed"]:
                st.write(f"Current System Determination: **{st.session_state['determination']}**")
                override = st.selectbox("Checker Final Determination", ["PASS", "REVISE", "REJECT"], index=["PASS", "REVISE", "REJECT"].index(st.session_state["determination"]))
                remarks = st.text_area("Checker Verification Remarks", placeholder="Enter academic lead audit remarks for author remediation...")
                if st.button("Apply Checker Sign-Off"):
                    st.session_state["determination"] = override
                    st.success(f"Final determination signed off as: {override}")
            else:
                st.info("Awaiting audit execution in Tab 1 before sign-off can be initiated.")

    # --------------------------------------------------------------------------
    # TAB: CATALOG MODERATION (Visible only to Admin role)
    # --------------------------------------------------------------------------
    if st.session_state["user_role"] == "Admin":
        with tabs[4]:
            st.markdown("#### 🏛️ Academic Catalog Moderation & Registration")
            st.caption("Register and moderate University Names, Programs, and Semesters/Cohorts.")

            col_a, col_b = st.columns(2)

            with col_a:
                with st.container(border=True):
                    st.subheader("1. Register / Add University")
                    new_uni = st.text_input("New University Name", placeholder="e.g., Global Open University")
                    if st.button("Add University", use_container_width=True):
                        if new_uni.strip() and new_uni.strip() not in [u["name"] for u in catalog.get("universities", [])]:
                            catalog.setdefault("universities", []).append({
                                "name": new_uni.strip(),
                                "programs": []
                            })
                            save_catalog(catalog)
                            st.success(f"University '{new_uni.strip()}' registered successfully.")
                            st.rerun()
                        else:
                            st.warning("Enter a non-empty, unique university name.")

                with st.container(border=True):
                    st.subheader("2. Register Program under University")
                    parent_uni = st.selectbox("Select Target University", [u["name"] for u in catalog.get("universities", [])], key="mod_parent_uni")
                    new_prog = st.text_input("New Academic Program", placeholder="e.g., M.Sc. Data Engineering")
                    if st.button("Add Academic Program", use_container_width=True):
                        if new_prog.strip():
                            target = next((u for u in catalog["universities"] if u["name"] == parent_uni), None)
                            if target:
                                if new_prog.strip() not in [p["name"] for p in target.get("programs", [])]:
                                    target.setdefault("programs", []).append({
                                        "name": new_prog.strip(),
                                        "semesters": ["Semester 1", "Semester 2"]
                                    })
                                    save_catalog(catalog)
                                    st.success(f"Program '{new_prog.strip()}' added to {parent_uni}.")
                                    st.rerun()
                                else:
                                    st.warning("Program already registered under this university.")

            with col_b:
                with st.container(border=True):
                    st.subheader("3. Add Semester / Batch / Cohort")
                    u_target = st.selectbox("University", [u["name"] for u in catalog.get("universities", [])], key="sem_u_target")
                    sel_u_obj = next((u for u in catalog["universities"] if u["name"] == u_target), None)
                    p_names = [p["name"] for p in sel_u_obj.get("programs", [])] if sel_u_obj else []

                    if p_names:
                        p_target = st.selectbox("Program", p_names, key="sem_p_target")
                        new_sem = st.text_input("New Semester / Batch / Cohort Label", placeholder="e.g., Batch 2026-B (Term 1)")
                        if st.button("Add Semester / Batch", use_container_width=True):
                            if new_sem.strip():
                                sel_p_obj = next((p for p in sel_u_obj["programs"] if p["name"] == p_target), None)
                                if sel_p_obj:
                                    if new_sem.strip() not in sel_p_obj.get("semesters", []):
                                        sel_p_obj.setdefault("semesters", []).append(new_sem.strip())
                                        save_catalog(catalog)
                                        st.success(f"Added '{new_sem.strip()}' to {p_target}.")
                                        st.rerun()
                                    else:
                                        st.warning("Semester or batch already exists.")
                    else:
                        st.info("Register a program first to add semesters.")

                with st.expander("View Full Registered Academic Catalog (JSON)"):
                    st.json(catalog)

        # ----------------------------------------------------------------------
        # TAB: ADMIN GOVERNANCE & RBAC (Visible only to Admin role)
        # ----------------------------------------------------------------------
        with tabs[5]:
            st.markdown("#### 🛡️ Administrator User Governance & RBAC Assignment")
            st.caption("Manage Maker and Checker rights for verified @upgrad.com users.")

            c_left, c_right = st.columns([1, 1])

            with c_left:
                with st.container(border=True):
                    st.subheader("Assign or Update User Authority")
                    target_email = st.text_input("Corporate Email", placeholder="colleague@upgrad.com")
                    assigned_level = st.selectbox("Assigned Role Level", ["Maker", "Checker", "Admin"])

                    if st.button("Save User Permissions", use_container_width=True, type="primary"):
                        c_mail = target_email.strip().lower()
                        if not c_mail.endswith("@upgrad.com"):
                            st.error("Error: Only @upgrad.com corporate accounts can be assigned rights.")
                        else:
                            roles[c_mail] = assigned_level
                            save_roles(roles)
                            st.success(f"User {c_mail} updated to role: **{assigned_level}**")
                            st.rerun()

            with c_right:
                st.subheader("Registered User Access Directory")
                user_list = [{"User Email": email, "Assigned Role": role} for email, role in roles.items()]
                users_df = pd.DataFrame(user_list)
                st.dataframe(users_df, use_container_width=True, hide_index=True)

                email_to_remove = st.selectbox("Revoke Access / Delete User", [u["User Email"] for u in user_list])
                if st.button("Revoke Selected Access", type="secondary"):
                    if email_to_remove in roles:
                        if email_to_remove == st.session_state["user_email"]:
                            st.error("Action denied: You cannot revoke your own administrator account.")
                        else:
                            del roles[email_to_remove]
                            save_roles(roles)
                            st.warning(f"Access revoked for {email_to_remove}.")
                            st.rerun()

if __name__ == "__main__":
    main()