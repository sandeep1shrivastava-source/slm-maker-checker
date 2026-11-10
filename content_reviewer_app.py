"""
Academic SLM Maker-Checker QA Platform
Enterprise Edition with Enhanced Graphical Presentation & Adaptive Model Routing
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
import plotly.express as px
import plotly.graph_objects as go
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import google.generativeai as genai

# ==============================================================================
# 1. STREAMLIT CONFIGURATION & STYLING
# ==============================================================================
st.set_page_config(
    page_title="Academic SLM QA Platform",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .stMetric {
        background-color: #f8fafc;
        padding: 15px;
        border-radius: 8px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    .status-badge-pass {
        background-color: #dcfce7;
        color: #166534;
        padding: 4px 12px;
        border-radius: 9999px;
        font-weight: 600;
        font-size: 0.85rem;
    }
    .status-badge-revise {
        background-color: #fef08a;
        color: #854d0e;
        padding: 4px 12px;
        border-radius: 9999px;
        font-weight: 600;
        font-size: 0.85rem;
    }
    .status-badge-reject {
        background-color: #fee2e2;
        color: #991b1b;
        padding: 4px 12px;
        border-radius: 9999px;
        font-weight: 600;
        font-size: 0.85rem;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        white-space: pre-wrap;
        background-color: #f1f5f9;
        border-radius: 4px 4px 0px 0px;
        gap: 1px;
        padding-top: 10px;
        padding-bottom: 10px;
    }
    .stTabs [aria-selected="true"] {
        background-color: #e2e8f0;
        border-bottom: 2px solid #3b82f6;
    }
</style>
""", unsafe_allow_html=True)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
UPLOAD_DIR = os.path.join(DATA_DIR, "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

CATALOG_FILE = os.path.join(DATA_DIR, "academic_catalog.json")
ROLES_FILE = os.path.join(DATA_DIR, "user_roles.json")
QUEUE_FILE = os.path.join(DATA_DIR, "workflow_queue.json")
LEDGER_FILE = os.path.join(DATA_DIR, "audit_history.csv")

# ==============================================================================
# 2. PERSISTENCE ENGINE
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

def load_json(path, default):
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    save_json(path, default)
    return default

def save_json(path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)

def append_to_ledger(entry: dict):
    df = pd.DataFrame([entry])
    if not os.path.exists(LEDGER_FILE):
        df.to_csv(LEDGER_FILE, index=False)
    else:
        df.to_csv(LEDGER_FILE, mode="a", header=False, index=False)

def load_ledger() -> pd.DataFrame:
    if os.path.exists(LEDGER_FILE):
        return pd.read_csv(LEDGER_FILE)
    return pd.DataFrame(columns=[
        "Timestamp", "Maker", "Checker", "University", "Program",
        "Semester", "File Name", "Quality Score", "Determination"
    ])

SESSION_DEFAULTS = {
    "authenticated": False,
    "user_email": "",
    "user_role": "Maker",
    "otp_sent": False,
    "generated_otp": "",
    "audit_completed": False,
    "active_file_sig": "",
    "active_queue_id": None,
    "audit_df": None,
    "defect_df": None,
    "syllabus_coverage": 0.0,
    "quality_score": 0.0,
    "determination": "PENDING",
    "selected_university": "State Open University",
    "selected_program": "Master of Business Administration (MBA)",
    "selected_semester": "Batch 2026-A (Semester 1)",
}

for key, val in SESSION_DEFAULTS.items():
    if key not in st.session_state:
        st.session_state[key] = val

# ==============================================================================
# 3. RBAC AUTHENTICATION & SMTP DISPATCH
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
          <p>You requested authentication for the <strong>SLM Maker-Checker Platform</strong>.</p>
          <div style="background-color: #f8f9fa; padding: 15px; text-align: center; font-size: 24px; font-weight: bold; letter-spacing: 5px; color: #1e3a8a; border-radius: 6px;">
            {otp_code}
          </div>
          <p style="margin-top: 20px; font-size: 12px; color: #777;">
            Verification code expires in 10 minutes. Use master bypass if testing in offline sandbox.
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
        <div style="text-align: center; padding: 40px 0 20px 0;">
            <h1 style='color: #1e3a8a; font-weight: 800;'>🎓 Academic SLM QA Platform</h1>
            <p style="color: #64748b; font-size: 18px;">
                Intelligent Evaluation & Regulatory Compliance Engine
            </p>
        </div>
        """,
        unsafe_allow_html=True
    )
    _, col2, _ = st.columns([1, 1.5, 1])
    with col2:
        with st.container(border=True):
            st.markdown("### Secure Enterprise Sign In")
            email = st.text_input("Corporate Email Address", placeholder="username@upgrad.com")
            roles_registry = load_json(ROLES_FILE, DEFAULT_ROLES)

            if not st.session_state["otp_sent"]:
                if st.button("Generate & Send OTP", use_container_width=True, type="primary"):
                    clean_email = email.strip().lower()
                    if not clean_email.endswith("@upgrad.com"):
                        st.error("Restricted Access: Only verified '@upgrad.com' corporate addresses are permitted.")
                    else:
                        otp = f"{random.randint(100000, 999999)}"
                        st.session_state["generated_otp"] = otp
                        st.session_state["user_email"] = clean_email
                        st.session_state["user_role"] = roles_registry.get(clean_email, "Maker")

                        with st.spinner("Dispatching verification code via Gmail SSL (Port 465)..."):
                            sent = dispatch_email_otp(clean_email, otp)
                            st.session_state["otp_sent"] = True
                            if sent:
                                st.success(f"Verification code sent to {clean_email}")
                            else:
                                st.info("Notice: SMTP not configured or failed. Use master bypass code: 123456")
                        st.rerun()
            else:
                st.info(f"Verification code sent to: **{st.session_state['user_email']}**")
                assigned_role = roles_registry.get(st.session_state["user_email"], "Maker")
                st.markdown(f"Resolved Authority Level: <span style='color:#2563eb; font-weight:bold;'>{assigned_role}</span>", unsafe_allow_html=True)

                entered_otp = st.text_input("Enter 6-Digit OTP", max_chars=6, type="password")
                bypass_code = str(get_secret("MASTER_BYPASS_OTP", "123456"))

                c1, c2 = st.columns(2)
                with c1:
                    if st.button("Verify & Sign In", use_container_width=True, type="primary"):
                        if entered_otp.strip() == st.session_state["generated_otp"] or entered_otp.strip() == bypass_code:
                            st.session_state["authenticated"] = True
                            st.session_state["user_role"] = assigned_role
                            st.rerun()
                        else:
                            st.error("Invalid verification code. Check inbox or use master bypass.")
                with c2:
                    if st.button("Cancel / Re-enter Email", use_container_width=True):
                        st.session_state["otp_sent"] = False
                        st.rerun()

# ==============================================================================
# 4. TEXT EXTRACTION & ADAPTIVE GEMINI AUDIT ENGINE
# ==============================================================================
def extract_text_from_file(file_obj_or_path) -> str:
    if hasattr(file_obj_or_path, "seek"):
        file_obj_or_path.seek(0)
    filename = getattr(file_obj_or_path, "name", str(file_obj_or_path)).lower()
    content = ""
    try:
        if filename.endswith(".docx"):
            doc = docx.Document(file_obj_or_path)
            content = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
        elif filename.endswith(".pdf"):
            try:
                import pypdf
                reader = pypdf.PdfReader(file_obj_or_path)
                for page in reader.pages:
                    t = page.extract_text()
                    if t: content += t + "\n"
            except Exception:
                import pdfplumber
                with pdfplumber.open(file_obj_or_path) as pdf:
                    for page in pdf.pages:
                        t = page.extract_text()
                        if t: content += t + "\n"
        elif filename.endswith(".csv"):
            df = pd.read_csv(file_obj_or_path)
            content = df.to_string()
        else:
            if hasattr(file_obj_or_path, "read"):
                content = file_obj_or_path.read().decode("utf-8", errors="ignore")
            else:
                with open(file_obj_or_path, "r", encoding="utf-8", errors="ignore") as f:
                    content = f.read()
    except Exception as e:
        st.error(f"Error reading deliverable: {str(e)}")
        return ""
    return content.strip()

def clean_and_parse_json(raw_text: str):
    text = raw_text.strip()
    text = re.sub(r"^```(?:json)?\s*", "", text, flags=re.MULTILINE)
    text = re.sub(r"\s*```$", "", text, flags=re.MULTILINE).strip()
    start_bracket, start_brace = text.find("["), text.find("{")
    if start_brace != -1 and (start_bracket == -1 or start_brace < start_bracket):
        end_brace = text.rfind("}")
        if end_brace != -1: text = text[start_brace : end_brace + 1]
    elif start_bracket != -1:
        end_bracket = text.rfind("]")
        if end_bracket != -1: text = text[start_bracket : end_bracket + 1]
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

def get_available_gemini_models():
    """Queries Google AI API for all models supporting generateContent."""
    priority_order = [
        "gemini-1.5-pro",
        "gemini-1.5-pro-latest",
        "gemini-1.5-flash",
        "gemini-1.5-flash-latest",
        "gemini-2.0-flash",
        "gemini-pro"
    ]
    try:
        supported = [
            m.name.replace("models/", "")
            for m in genai.list_models()
            if "generateContent" in getattr(m, "supported_generation_methods", [])
        ]
        if supported:
            # Sort by priority order
            ordered = [c for c in priority_order if c in supported]
            for s in supported:
                if s not in ordered:
                    ordered.append(s)
            return ordered
    except Exception:
        pass
    return ["gemini-1.5-flash", "gemini-1.5-pro-latest", "gemini-1.5-pro", "gemini-2.0-flash"]

def run_gemini_academic_audit(manuscript_text: str, syllabus_text: str = ""):
    api_key = get_secret("GEMINI_API_KEY") or os.environ.get("GEMINI_API_KEY")
    if not api_key:
        st.error("Missing GEMINI_API_KEY. Configure in .streamlit/secrets.toml.")
        st.stop()

    genai.configure(api_key=api_key)
    candidate_models = get_available_gemini_models()

    prompt = f"""
    You are an expert Chief Academic Auditor for Higher Education Self-Learning Materials (SLM).
    Evaluate the provided manuscript rigorously against UGC-DEB benchmarks, Bloom's Revised Taxonomy, and reference syllabus specifications.

    ### CRITICAL EVALUATION RULES:
    1. TRUNCATION DETECTION: Inspect whether the manuscript cuts off abruptly, ends mid-sentence (e.g., 'Amazon Dyna...'), or terminates without completing required sections. If truncated, assign 'DEFICIENT' to Dimension 3 & 5, deduct at least 50% score for them, and flag a 'Critical' defect.
    2. SAQs & POSTLIMS CHECK: If Section headers (SAQs, Summary, Keywords, References) exist in outline/structure but are missing actual body questions, answers, or text, mark Dimension 6 and 10 as NON-COMPLIANT and log Defect.
    3. NO FAKE PASSES: Do NOT output generic 'Meets pedagogical quality standards.' unless the manuscript is truly flawless. Quote exact issues, section numbers, or sentence fragments in 'Auditor Findings'.
    4. SCORE ARITHMETIC: Each parameter score MUST NOT exceed its weight.

    ### 12-PARAMETER MATRIX WEIGHTS:
    1. Syllabus & Curriculum Mapping (Max: 10) | 2. Bloom's Taxonomy & Action Verbs (Max: 10) | 3. Conceptual Depth & Pedagogical Rigor (Max: 10) | 4. Real-world Case Studies & Applications (Max: 8)
    5. Quadrant-II Structural Architecture (Max: 8) | 6. Self-Assessment & Review Questions (Max: 8) | 7. Factual & Theoretical Accuracy (Max: 10) | 8. Originality, Attribution & Non-Hallucination (Max: 10)
    9. Academic Tone, Style & Language Clarity (Max: 6) | 10. Key Takeaways & Summary Synthesis (Max: 6) | 11. Benchmark & Reference Adherence (Max: 8) | 12. Pedagogical Scaffolding & Formatting (Max: 6)
    TOTAL POSSIBLE SCORE = 100

    ### OUTPUT FORMAT (STRICT JSON ONLY):
    {{
      "syllabus_coverage_pct": 85.0,
      "matrix": [
        {{"param_no": 1, "dimension": "Syllabus & Curriculum Mapping", "weight": 10, "score_achieved": 10.0, "status": "COMPLIANT", "auditor_findings": "Detailed textual finding...", "remediation_guidance": "Concrete guidance..."}}
      ],
      "defects": [
        {{"parameter": "Conceptual Depth & Pedagogical Rigor", "severity": "Critical", "finding": "Manuscript truncates abruptly...", "remediation": "Author full text..."}}
      ]
    }}
    
    ### REFERENCE SYLLABUS:
    {syllabus_text if syllabus_text else "No external syllabus provided. Audit against internal pedagogical structure and stated unit objectives."}
    ### MANUSCRIPT CONTENT (EXCERPT/FULL):
    {manuscript_text[:35000]}
    """

    last_error = None
    for model_name in candidate_models:
        try:
            model = genai.GenerativeModel(
                model_name=model_name,
                generation_config={"temperature": 0.0, "response_mime_type": "application/json"}
            )
            response = model.generate_content(prompt)
            return clean_and_parse_json(response.text)
        except Exception as e:
            last_error = e
            continue

    st.error(f"❌ Gemini Evaluation Engine Error across tested endpoints ({', '.join(candidate_models)}): {str(last_error)}")
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
            "Param #": p_no, "QA Dimension": item["dimension"], "Weight": weight, "Score Achieved": score,
            "Success (%)": f"{(score / weight) * 100.0:.1f}%", "Status": status, "Auditor Findings": findings, "Remediation Guidance": remediation
        })
    audit_df = pd.DataFrame(reconciled_rows)

    # Truncation Detection
    tail_text = manuscript_text.strip()[-200:].lower()
    is_truncated = any(t in tail_text for t in ["dyna", "amazon dyna", "...", "the "]) or not tail_text.endswith((".", "!", "?", "\"", "”"))

    defects, log_counter = [], 1
    if is_truncated:
        defects.append({
            "Defect ID": f"LOG-{log_counter:02d}", "Parameter": "Conceptual Depth & Pedagogical Rigor", "Severity": "Critical",
            "Finding": f"Manuscript cuts off abruptly at terminus: '...{manuscript_text.strip()[-60:]}'", "Remediation": "Complete all truncated sentences and complete subsequent sections."
        })
        log_counter += 1

    for d in raw_defects:
        defects.append({
            "Defect ID": f"LOG-{log_counter:02d}", "Parameter": d.get("parameter", "Pedagogical Scaffolding"), "Severity": d.get("severity", "Major"),
            "Finding": d.get("finding", "Defect identified during SME review."), "Remediation": d.get("remediation", "Remediate as recommended by auditor.")
        })
        log_counter += 1

    defect_df = pd.DataFrame(defects) if defects else pd.DataFrame(columns=["Defect ID", "Parameter", "Severity", "Finding", "Remediation"])
    total_score = float(audit_df["Score Achieved"].sum())
    syllabus_cov = float(raw_audit.get("syllabus_coverage_pct", 100.0))

    has_critical = not defect_df.empty and (defect_df["Severity"].str.lower() == "critical").any()
    if has_critical or total_score < 75.0: determination = "REVISE"
    elif total_score < 60.0: determination = "REJECT"
    else: determination = "PASS"

    return audit_df, defect_df, total_score, syllabus_cov, determination

# ==============================================================================
# 5. DOSSIER EXPORT GENERATORS
# ==============================================================================
def generate_docx_report(audit_df, defect_df, meta):
    doc = docx.Document()
    p_title = doc.add_paragraph()
    r_title = p_title.add_run("ACADEMIC CONTENT QUALITY REVIEW REPORT")
    r_title.bold = True
    r_title.font.size = Pt(16)
    r_title.font.color.rgb = RGBColor(30, 58, 138)
    
    doc.add_paragraph(f"Stage 2 SME Quality & Syllabus Cross-Audit | {meta['university_name']} - {meta['program_name']}")
    doc.add_paragraph("―" * 60)

    meta_table = doc.add_table(rows=3, cols=2)
    meta_table.autofit = True
    cells = [
        (0, 0, f"University: {meta['university_name']}"), (0, 1, f"Determination: {meta['determination']}"),
        (1, 0, f"Academic Program: {meta['program_name']}"), (1, 1, f"Quality Score: {meta['quality_score']:.1f} / 100"),
        (2, 0, f"Batch / Term: {meta['batch_term']}"), (2, 1, f"Syllabus Coverage: {meta['syllabus_coverage']:.1f}%"),
    ]
    for r, c, txt in cells:
        cell = meta_table.cell(r, c)
        cell.text = txt
        if "Determination:" in txt or "Quality Score:" in txt:
            cell.paragraphs[0].runs[0].bold = True

    doc.add_paragraph()
    doc.add_heading("1. 12-Parameter Quality Matrix", level=2)
    table = doc.add_table(rows=1, cols=6, style="Table Grid")
    hdr_cells = table.rows[0].cells
    for i, title in enumerate(["#", "Dimension", "Weight", "Score", "Status", "Findings"]):
        hdr_cells[i].text = title
        hdr_cells[i].paragraphs[0].runs[0].bold = True

    for _, row in audit_df.iterrows():
        row_cells = table.add_row().cells
        for i, val in enumerate(row): row_cells[i].text = str(val)

    doc.add_paragraph()
    doc.add_heading("2. Defect Register", level=2)
    if defect_df.empty:
        doc.add_paragraph("Zero critical or major defects identified. The manuscript is compliant.")
    else:
        d_table = doc.add_table(rows=1, cols=5, style="Table Grid")
        d_hdr = d_table.rows[0].cells
        for i, title in enumerate(["Defect ID", "Parameter", "Severity", "Finding", "Remediation"]):
            d_hdr[i].text = title
            d_hdr[i].paragraphs[0].runs[0].bold = True
        for _, d_row in defect_df.iterrows():
            rc = d_table.add_row().cells
            for i, val in enumerate(d_row): rc[i].text = str(val)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()

def generate_excel_report(audit_df, defect_df, meta):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        audit_df.to_excel(writer, sheet_name="Scorecard", index=False)
        defect_df.to_excel(writer, sheet_name="Defect Register", index=False)
        pd.DataFrame([
            {"Key": "University", "Value": meta["university_name"]},
            {"Key": "Academic Program", "Value": meta["program_name"]},
            {"Key": "Batch / Term", "Value": meta["batch_term"]},
            {"Key": "Determination", "Value": meta["determination"]},
            {"Key": "Quality Score", "Value": f"{meta['quality_score']:.1f} / 100"},
            {"Key": "Audit Timestamp", "Value": datetime.now().strftime("%Y-%m-%d %H:%M:%S")},
        ]).to_excel(writer, sheet_name="Audit Summary", index=False)
    return buf.getvalue()

# ==============================================================================
# 6. VISUALIZATION FUNCTIONS (Plotly)
# ==============================================================================
def render_donut_chart(score):
    fig = go.Figure(go.Indicator(
        mode = "gauge+number",
        value = score,
        domain = {'x': [0, 1], 'y': [0, 1]},
        title = {'text': "Quality Score", 'font': {'size': 24}},
        gauge = {
            'axis': {'range': [None, 100], 'tickwidth': 1},
            'bar': {'color': "#2563eb" if score >= 75 else ("#f59e0b" if score >= 60 else "#ef4444")},
            'bgcolor': "white",
            'steps': [
                {'range': [0, 60], 'color': '#fee2e2'},
                {'range': [60, 75], 'color': '#fef3c7'},
                {'range': [75, 100], 'color': '#dcfce7'}],
        }))
    fig.update_layout(height=250, margin=dict(l=20, r=20, t=50, b=20))
    return fig

def render_radar_chart(df):
    fig = px.line_polar(df, r='Score Achieved', theta='QA Dimension', line_close=True, range_r=[0, 10])
    fig.update_traces(fill='toself', fillcolor='rgba(37, 99, 235, 0.2)', line_color='#2563eb')
    fig.update_layout(
        polar=dict(radialaxis=dict(visible=True, range=[0, 10])),
        showlegend=False,
        height=350, margin=dict(l=40, r=40, t=20, b=20)
    )
    return fig

# ==============================================================================
# 7. MAIN CONTROLLER
# ==============================================================================
def main():
    if not st.session_state["authenticated"]:
        render_login_screen()
        return

    catalog = load_json(CATALOG_FILE, DEFAULT_CATALOG)
    roles = load_json(ROLES_FILE, DEFAULT_ROLES)
    queue = load_json(QUEUE_FILE, [])

    # --- SIDEBAR: AFFILIATION & SELF-SERVE REGISTRATION ---
    with st.sidebar:
        st.markdown(f"### 👤 Role: **{st.session_state['user_role']}**")
        st.caption(f"{st.session_state['user_email']}")
        st.divider()

        st.subheader("Academic Affiliation")
        uni_names = [u["name"] for u in catalog.get("universities", [])]
        if not uni_names: uni_names = ["State Open University"]

        def_u_idx = uni_names.index(st.session_state["selected_university"]) if st.session_state["selected_university"] in uni_names else 0
        selected_uni = st.selectbox("Registered University", uni_names, index=def_u_idx)
        st.session_state["selected_university"] = selected_uni

        matching_uni = next((u for u in catalog.get("universities", []) if u["name"] == selected_uni), None)
        prog_list = matching_uni.get("programs", []) if matching_uni else []
        prog_names = [p["name"] for p in prog_list] if prog_list else ["Master of Business Administration (MBA)"]

        def_p_idx = prog_names.index(st.session_state["selected_program"]) if st.session_state["selected_program"] in prog_names else 0
        selected_prog = st.selectbox("Academic Program", prog_names, index=def_p_idx)
        st.session_state["selected_program"] = selected_prog

        matching_prog = next((p for p in prog_list if p["name"] == selected_prog), None) if prog_list else None
        sem_list = matching_prog.get("semesters", []) if matching_prog else ["Batch 2026-A (Semester 1)"]

        def_s_idx = sem_list.index(st.session_state["selected_semester"]) if st.session_state["selected_semester"] in sem_list else 0
        selected_sem = st.selectbox("Semester / Batch", sem_list, index=def_s_idx)
        st.session_state["selected_semester"] = selected_sem

        with st.expander("➕ Register Affiliation (Self-Serve)"):
            new_u = st.text_input("New University Name")
            new_p = st.text_input("New Degree Program")
            new_s = st.text_input("New Semester / Term")
            if st.button("Save New Affiliation", use_container_width=True):
                if new_u.strip() and new_u.strip() not in [u["name"] for u in catalog.get("universities", [])]:
                    catalog.setdefault("universities", []).append({"name": new_u.strip(), "programs": []})
                target_u = next((u for u in catalog["universities"] if u["name"] == (new_u.strip() or selected_uni)), None)
                if target_u and new_p.strip() and new_p.strip() not in [p["name"] for p in target_u.get("programs", [])]:
                    target_u.setdefault("programs", []).append({"name": new_p.strip(), "semesters": []})
                target_p = next((p for p in target_u["programs"] if p["name"] == (new_p.strip() or selected_prog)), None) if target_u else None
                if target_p and new_s.strip() and new_s.strip() not in target_p.get("semesters", []):
                    target_p.setdefault("semesters", []).append(new_s.strip())
                save_json(CATALOG_FILE, catalog)
                st.success("Affiliation registered.")
                st.rerun()

        st.divider()
        if st.button("Sign Out / Switch User", use_container_width=True):
            for k in list(st.session_state.keys()): del st.session_state[k]
            st.rerun()

    # Main App Header
    st.markdown("""
        <div style='background: linear-gradient(to right, #1e3a8a, #3b82f6); padding: 20px; border-radius: 8px; margin-bottom: 20px;'>
            <h1 style='color: white; margin:0;'>Academic SLM Maker-Checker QA Platform</h1>
            <p style='color: #e2e8f0; margin:0;'>Quality Assurance & Regulatory Compliance Engine</p>
        </div>
    """, unsafe_allow_html=True)

    # Dynamic Tab Routing
    if st.session_state["user_role"] == "Admin":
        tab_titles = ["📤 Submit Deliverable", "📥 Audit Inbox", "📊 Quality Dashboard", "🚨 Defect Register", "✅ Sign-Off & Exports", "📈 Platform Analytics", "🏛️ Catalog", "🛡️ RBAC"]
    elif st.session_state["user_role"] == "Checker":
        tab_titles = ["📤 Submit Deliverable", "📥 Audit Inbox", "📊 Quality Dashboard", "🚨 Defect Register", "✅ Sign-Off & Exports", "📈 Platform Analytics"]
    else:
        tab_titles = ["📤 Submit Deliverable", "📊 Quality Dashboard", "🚨 Defect Register", "✅ Sign-Off & Exports", "📈 Platform Analytics"]

    tabs = st.tabs(tab_titles)

    # --------------------------------------------------------------------------
    # TAB 1: MAKER SUBMISSION
    # --------------------------------------------------------------------------
    with tabs[0]:
        st.markdown(f"### Submit Manuscript for **{st.session_state['selected_university']}**")
        st.caption(f"Curriculum: {st.session_state['selected_program']} | {st.session_state['selected_semester']}")

        with st.container(border=True):
            col_m, col_s = st.columns(2)
            with col_m:
                uploaded_manuscript = st.file_uploader("Upload Manuscript (SLM Module / Unit)", type=["docx", "pdf", "txt"], key="up_maker_manuscript")
            with col_s:
                uploaded_syllabus = st.file_uploader("Upload Syllabus Outline (Optional)", type=["docx", "pdf", "txt"], key="up_maker_syllabus")

            if uploaded_manuscript is not None:
                c_btn1, c_btn2 = st.columns(2)
                with c_btn1:
                    if st.button("🚀 Push to Checker Audit Queue", type="primary", use_container_width=True):
                        file_uuid = str(uuid.uuid4())[:8]
                        saved_filename = f"{file_uuid}_{uploaded_manuscript.name}"
                        save_path = os.path.join(UPLOAD_DIR, saved_filename)
                        with open(save_path, "wb") as f: f.write(uploaded_manuscript.getbuffer())

                        queue.append({
                            "id": file_uuid, "maker": st.session_state["user_email"], "university": st.session_state["selected_university"],
                            "program": st.session_state["selected_program"], "semester": st.session_state["selected_semester"],
                            "filename": uploaded_manuscript.name, "filepath": save_path, "status": "Pending Checker Audit",
                            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        })
                        save_json(QUEUE_FILE, queue)
                        st.success("Deliverable submitted to Checker Queue.")

                with c_btn2:
                    if st.button("⚡ Run Instant Quality Self-Check", use_container_width=True):
                        with st.spinner("Executing 12-Parameter Evaluation Engine..."):
                            m_text = extract_text_from_file(uploaded_manuscript)
                            s_text = extract_text_from_file(uploaded_syllabus) if uploaded_syllabus else ""
                            if not m_text or len(m_text.strip()) < 40:
                                st.error("Uploaded manuscript has no readable text.")
                                st.stop()
                            raw_audit = run_gemini_academic_audit(m_text, s_text)
                            a_df, d_df, q_score, s_cov, det = reconcile_and_audit(raw_audit, m_text)

                            st.session_state.update({
                                "audit_df": a_df, "defect_df": d_df, "quality_score": q_score, "syllabus_coverage": s_cov,
                                "determination": det, "audit_completed": True, "active_queue_id": None
                            })
                            st.success("Quality check complete. View Dashboard.")

        st.markdown("### My Submitted Deliverables")
        my_subs = [q for q in queue if q.get("maker") == st.session_state["user_email"]]
        if my_subs:
            st.dataframe(pd.DataFrame(my_subs)[["filename", "status", "timestamp", "university", "program", "semester"]], use_container_width=True, hide_index=True)

    # --------------------------------------------------------------------------
    # TAB 2: CHECKER AUDIT INBOX
    # --------------------------------------------------------------------------
    if st.session_state["user_role"] in ["Checker", "Admin"]:
        with tabs[1]:
            st.markdown("### Pending Checker Deliverables Inbox")
            pending_items = [q for q in queue if q.get("status") == "Pending Checker Audit"]

            if pending_items:
                df_pending = pd.DataFrame(pending_items)
                st.dataframe(df_pending[["id", "maker", "filename", "timestamp", "university", "program"]], use_container_width=True, hide_index=True)

                target_id = st.selectbox("Select Manuscript ID to Audit", df_pending["id"].tolist())
                if st.button("▶️ Execute Regulatory Audit on Selected Deliverable", type="primary"):
                    target_obj = next(q for q in pending_items if q["id"] == target_id)
                    with st.spinner(f"Auditing '{target_obj['filename']}'..."):
                        m_text = extract_text_from_file(target_obj["filepath"])
                        raw_audit = run_gemini_academic_audit(m_text, "")
                        a_df, d_df, q_score, s_cov, det = reconcile_and_audit(raw_audit, m_text)

                        st.session_state.update({
                            "audit_df": a_df, "defect_df": d_df, "quality_score": q_score, "syllabus_coverage": s_cov, "determination": det,
                            "active_queue_id": target_id, "selected_university": target_obj["university"], "selected_program": target_obj["program"],
                            "selected_semester": target_obj["semester"], "audit_completed": True
                        })
                        st.success("Audit complete. Open Quality Dashboard.")
            else:
                st.info("No submissions currently waiting for Checker audit.")

    # --------------------------------------------------------------------------
    # TAB 3: VISUAL DASHBOARD & MATRIX
    # --------------------------------------------------------------------------
    matrix_tab_idx = 1 if st.session_state["user_role"] == "Maker" else 2
    with tabs[matrix_tab_idx]:
        if st.session_state["audit_completed"] and st.session_state["audit_df"] is not None:
            st.markdown("### Quality Analytics Dashboard")
            
            # Top KPI Row
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Overall Quality Score", f"{st.session_state['quality_score']:.1f} / 100")
            c2.metric("Syllabus Coverage", f"{st.session_state['syllabus_coverage']:.1f}%")
            
            det_class = "status-badge-pass" if st.session_state["determination"] == "PASS" else ("status-badge-revise" if st.session_state["determination"] == "REVISE" else "status-badge-reject")
            c3.markdown(f"**Final Determination**<br><br><span class='{det_class}'>{st.session_state['determination']}</span>", unsafe_allow_html=True)
            c4.metric("Logged Defects", len(st.session_state["defect_df"]))

            st.divider()
            
            # Visual Charts Row
            chart_col1, chart_col2 = st.columns([1, 2])
            with chart_col1:
                st.plotly_chart(render_donut_chart(st.session_state["quality_score"]), use_container_width=True)
            with chart_col2:
                st.plotly_chart(render_radar_chart(st.session_state["audit_df"]), use_container_width=True)

            st.markdown("### 12-Parameter Quality Matrix")
            def highlight_status(val):
                color = '#dcfce7' if val == 'COMPLIANT' else '#fee2e2'
                return f'background-color: {color}'
            
            styled_df = st.session_state["audit_df"].style.map(highlight_status, subset=['Status'])
            st.dataframe(styled_df, use_container_width=True, hide_index=True)
        else:
            st.info("Execute an audit to populate the visual dashboard.")

    # --------------------------------------------------------------------------
    # TAB 4: DEFECT REGISTER
    # --------------------------------------------------------------------------
    defect_tab_idx = 2 if st.session_state["user_role"] == "Maker" else 3
    with tabs[defect_tab_idx]:
        st.markdown("### Sequential Defect & Remediation Register")
        if st.session_state["audit_completed"] and st.session_state["defect_df"] is not None:
            if st.session_state["defect_df"].empty:
                st.success("🎉 Zero critical or major defects identified.")
            else:
                def highlight_severity(val):
                    if val == 'Critical': return 'background-color: #fee2e2; font-weight: bold; color: #991b1b'
                    if val == 'Major': return 'background-color: #fef3c7; color: #92400e'
                    return ''
                styled_defects = st.session_state["defect_df"].style.map(highlight_severity, subset=['Severity'])
                st.dataframe(styled_defects, use_container_width=True, hide_index=True)
        else:
            st.info("Sequential defect logs will display here.")

    # --------------------------------------------------------------------------
    # TAB 5: SIGN-OFF & EXPORTS
    # --------------------------------------------------------------------------
    signoff_tab_idx = 3 if st.session_state["user_role"] == "Maker" else 4
    with tabs[signoff_tab_idx]:
        st.markdown("### Sign-Off & Regulatory Dossier Export")
        if st.session_state["audit_completed"]:
            meta = {
                "university_name": st.session_state["selected_university"],
                "program_name": st.session_state["selected_program"],
                "batch_term": st.session_state["selected_semester"],
                "determination": st.session_state["determination"],
                "quality_score": st.session_state["quality_score"],
                "syllabus_coverage": st.session_state["syllabus_coverage"]
            }

            with st.container(border=True):
                st.subheader("Auditor Determination & Sign-Off")
                st.write(f"System Evaluated Determination: **{st.session_state['determination']}**")

                c_ov1, c_ov2 = st.columns(2)
                with c_ov1:
                    override_det = st.selectbox("Final Regulatory Determination", ["PASS", "REVISE", "REJECT"], index=["PASS", "REVISE", "REJECT"].index(st.session_state["determination"]))
                with c_ov2:
                    st.text_input("Lead Auditor Remarks", value="Reviewed and validated.")

                if st.button("Apply Sign-Off & Commit to Platform Ledger", type="primary"):
                    st.session_state["determination"] = override_det
                    meta["determination"] = override_det

                    if st.session_state.get("active_queue_id"):
                        for q in queue:
                            if q["id"] == st.session_state["active_queue_id"]:
                                q["status"] = f"Signed Off ({override_det})"
                                q["quality_score"] = st.session_state["quality_score"]
                                q["checker"] = st.session_state["user_email"]
                                break
                        save_json(QUEUE_FILE, queue)

                    append_to_ledger({
                        "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "Maker": st.session_state.get("user_email", "Direct Audit"),
                        "Checker": st.session_state["user_email"],
                        "University": meta["university_name"], "Program": meta["program_name"], "Semester": meta["batch_term"],
                        "File Name": f"{meta['program_name']}_Module", "Quality Score": meta["quality_score"], "Determination": override_det
                    })
                    st.success("Sign-off recorded and committed to platform ledger.")

            st.markdown("### Download Regulatory Quality Dossiers")
            c1, c2, c3 = st.columns(3)
            with c1:
                st.download_button("📥 Download Matrix (CSV)", st.session_state["audit_df"].to_csv(index=False).encode("utf-8"), file_name=f"{datetime.now().strftime('%Y-%m-%d')}_export.csv", use_container_width=True)
            with c2:
                st.download_button("📄 Download Dossier (DOCX)", generate_docx_report(st.session_state["audit_df"], st.session_state["defect_df"], meta), file_name=f"{meta['university_name'].replace(' ', '_')}_QA_Report.docx", use_container_width=True)
            with c3:
                st.download_button("📊 Download Scorecard (XLSX)", generate_excel_report(st.session_state["audit_df"], st.session_state["defect_df"], meta), file_name=f"{meta['university_name'].replace(' ', '_')}_QA_Report.xlsx", use_container_width=True)

    # --------------------------------------------------------------------------
    # TAB 6: PLATFORM ANALYTICS DASHBOARD
    # --------------------------------------------------------------------------
    analytics_tab_idx = 4 if st.session_state["user_role"] == "Maker" else 5
    with tabs[analytics_tab_idx]:
        st.markdown("### 📈 Platform Audit Ledger & Usage Analytics")
        ledger_df = load_ledger()
        if not ledger_df.empty:
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Total Audits Recorded", len(ledger_df))
            m2.metric("Universities Serviced", ledger_df["University"].nunique())
            m3.metric("Platform Pass Rate", f"{(len(ledger_df[ledger_df['Determination'] == 'PASS']) / len(ledger_df)) * 100.0:.1f}%")
            m4.metric("Avg Quality Score", f"{ledger_df['Quality Score'].mean():.1f} / 100")

            st.divider()
            
            if len(ledger_df) > 1:
                ledger_df['Date'] = pd.to_datetime(ledger_df['Timestamp']).dt.date
                trend_df = ledger_df.groupby('Date')['Quality Score'].mean().reset_index()
                fig_trend = px.line(trend_df, x='Date', y='Quality Score', title="Average Quality Score Trend")
                fig_trend.update_layout(height=300)
                st.plotly_chart(fig_trend, use_container_width=True)

            st.dataframe(ledger_df.sort_values(by="Timestamp", ascending=False), use_container_width=True, hide_index=True)
            st.download_button("📥 Download Complete Platform Ledger (CSV)", ledger_df.to_csv(index=False).encode("utf-8"), "platform_audit_history.csv")
        else:
            st.info("The platform ledger is currently empty.")

    # --------------------------------------------------------------------------
    # ADMIN TABS
    # --------------------------------------------------------------------------
    if st.session_state["user_role"] == "Admin":
        with tabs[6]:
            st.markdown("### 🏛️ Academic Catalog Moderation")
            col_a, col_b = st.columns(2)
            with col_a:
                with st.container(border=True):
                    st.subheader("1. Add University")
                    n_uni = st.text_input("University Name", key="adm_u")
                    if st.button("Register University", use_container_width=True):
                        if n_uni.strip() and n_uni.strip() not in [u["name"] for u in catalog.get("universities", [])]:
                            catalog.setdefault("universities", []).append({"name": n_uni.strip(), "programs": []})
                            save_json(CATALOG_FILE, catalog)
                            st.success(f"University registered.")
                            st.rerun()

                with st.container(border=True):
                    st.subheader("2. Add Program")
                    p_uni = st.selectbox("Target University", [u["name"] for u in catalog.get("universities", [])], key="adm_pu")
                    n_prg = st.text_input("New Program Title", key="adm_p")
                    if st.button("Register Program", use_container_width=True):
                        t_u = next((u for u in catalog["universities"] if u["name"] == p_uni), None)
                        if t_u and n_prg.strip() and n_prg.strip() not in [p["name"] for p in t_u.get("programs", [])]:
                            t_u.setdefault("programs", []).append({"name": n_prg.strip(), "semesters": ["Semester 1", "Semester 2"]})
                            save_json(CATALOG_FILE, catalog)
                            st.success(f"Program registered.")
                            st.rerun()

            with col_b:
                with st.container(border=True):
                    st.subheader("3. Add Semester / Batch")
                    u_t = st.selectbox("University", [u["name"] for u in catalog.get("universities", [])], key="adm_su")
                    u_obj = next((u for u in catalog["universities"] if u["name"] == u_t), None)
                    p_list = [p["name"] for p in u_obj.get("programs", [])] if u_obj else []

                    if p_list:
                        p_t = st.selectbox("Program", p_list, key="adm_sp")
                        n_sem = st.text_input("Semester Label", key="adm_s")
                        if st.button("Register Semester", use_container_width=True):
                            p_obj = next((p for p in u_obj["programs"] if p["name"] == p_t), None)
                            if p_obj and n_sem.strip() and n_sem.strip() not in p_obj.get("semesters", []):
                                p_obj.setdefault("semesters", []).append(n_sem.strip())
                                save_json(CATALOG_FILE, catalog)
                                st.success(f"Semester registered.")
                                st.rerun()

        with tabs[7]:
            st.markdown("### 🛡️ Administrator User Governance & RBAC")
            c_l, c_r = st.columns(2)
            with c_l:
                with st.container(border=True):
                    st.subheader("Assign User Authority")
                    t_email = st.text_input("User Email Address")
                    t_role = st.selectbox("Assigned Role", ["Maker", "Checker", "Admin"])
                    if st.button("Save Permissions", type="primary", use_container_width=True):
                        c_mail = t_email.strip().lower()
                        if not c_mail.endswith("@upgrad.com"): st.error("Only '@upgrad.com' addresses permitted.")
                        else:
                            roles[c_mail] = t_role
                            save_json(ROLES_FILE, roles)
                            st.success(f"Updated {c_mail} to {t_role}.")
                            st.rerun()
            with c_r:
                st.subheader("User Directory")
                st.dataframe(pd.DataFrame([{"Email": k, "Role": v} for k, v in roles.items()]), use_container_width=True, hide_index=True)

if __name__ == "__main__":
    main()