import io
import json
import streamlit as st
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

# Modular Project Services
from services.access_manager import load_access_config, verify_user, add_university, assign_user_role
from services.unified_engine import run_stage_1_maker, run_stage_2_checker, run_stage_3_prompt_audit

# -------------------------------------------------------------
# PAGE CONFIGURATION
# -------------------------------------------------------------
st.set_page_config(
    page_title="Academic Content Pipeline & Reviewer",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -------------------------------------------------------------
# 12-PARAMETER ACADEMIC QA FRAMEWORK DEFINITION
# -------------------------------------------------------------
QA_12_PARAMETERS = [
    {"id": 1, "name": "Syllabus & Curriculum Mapping", "weight": 10, "desc": "Alignment of all draft units/topics to the prescribed university syllabus."},
    {"id": 2, "name": "Bloom's Taxonomy & Action Verbs", "weight": 10, "desc": "Adequacy of cognitive action verbs across learning outcomes and assessments."},
    {"id": 3, "name": "Conceptual Depth & Pedagogical Rigor", "weight": 10, "desc": "Technical thoroughness, conceptual depth, and academic sufficiency."},
    {"id": 4, "name": "Real-world Case Studies & Applications", "weight": 8, "desc": "Inclusion of illustrative real-world examples, corporate scenarios, or case studies."},
    {"id": 5, "name": "Quadrant-II Structural Architecture", "weight": 8, "desc": "Compliance with UGC/DEB self-learning material pedagogical flow."},
    {"id": 6, "name": "Self-Assessment & Formative Review Questions", "weight": 8, "desc": "Quality of self-assessment questions, MCQs, and reflective exercises."},
    {"id": 7, "name": "Factual & Theoretical Accuracy", "weight": 10, "desc": "Verification of domain formulas, theorems, dates, data, and claims."},
    {"id": 8, "name": "Originality, Attribution & Non-Hallucination", "weight": 10, "desc": "Zero unverified AI assertions, phantom citations, or unattributed blocks."},
    {"id": 9, "name": "Academic Tone, Style & Language Clarity", "weight": 6, "desc": "Objective scholarly voice, clarity of prose, grammatical precision."},
    {"id": 10, "name": "Key Takeaways & Summary Synthesis", "weight": 6, "desc": "Clarity of unit summaries, review checklists, and concept recaps."},
    {"id": 11, "name": "Benchmark & Reference Adherence", "weight": 8, "desc": "Direct correlation to prescribed university benchmark texts and references."},
    {"id": 12, "name": "Pedagogical Scaffolding & Formatting", "weight": 6, "desc": "Use of formatting, tables, highlight callouts, and visual cues."}
]

# -------------------------------------------------------------
# SAFE DOCUMENT EXTRACTION & FORMAT DETECTION
# -------------------------------------------------------------
def extract_text_safely(file_input) -> str:
    if file_input is None:
        return ""
    fname = getattr(file_input, "name", "").lower()
    
    if hasattr(file_input, "getvalue"):
        raw_data = file_input.getvalue()
    elif hasattr(file_input, "read"):
        raw_data = file_input.read()
    else:
        raw_data = file_input

    if fname.endswith(".docx") or (isinstance(raw_data, bytes) and raw_data.startswith(b"PK")):
        try:
            doc = Document(io.BytesIO(raw_data))
            text_blocks = []
            for p in doc.paragraphs:
                if p.text.strip():
                    text_blocks.append(p.text.strip())
            for tbl in doc.tables:
                for row in tbl.rows:
                    row_vals = [c.text.strip() for c in row.cells if c.text.strip()]
                    if row_vals:
                        text_blocks.append(" | ".join(row_vals))
            return "\n".join(text_blocks)
        except Exception:
            pass

    if isinstance(raw_data, bytes):
        try:
            return raw_data.decode("utf-8")
        except Exception:
            return raw_data.decode("latin-1", errors="ignore")
    return str(raw_data)

def extract_styles_from_reference(ref_file):
    font_name = 'Arial'
    font_size = Pt(11)
    alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    
    if not ref_file or not getattr(ref_file, "name", "").lower().endswith(".docx"):
        return font_name, font_size, alignment
        
    try:
        raw_data = ref_file.getvalue() if hasattr(ref_file, "getvalue") else ref_file.read()
        if hasattr(ref_file, "seek"): 
            ref_file.seek(0)
        
        ref_doc = Document(io.BytesIO(raw_data))
        for p in ref_doc.paragraphs:
            if p.text.strip():
                if p.alignment is not None:
                    alignment = p.alignment
                for r in p.runs:
                    if r.font.name:
                        font_name = r.font.name
                    if r.font.size:
                        font_size = r.font.size
                break
    except Exception:
        pass
        
    return font_name, font_size, alignment

# -------------------------------------------------------------
# STAGE 1: DELIVERABLE MANUSCRIPT DOCX BUILDER
# -------------------------------------------------------------
def build_manuscript_docx(title: str, content: str, ref_file=None) -> io.BytesIO:
    doc = Document()
    font_name, font_size, user_alignment = extract_styles_from_reference(ref_file)
    
    style = doc.styles['Normal']
    style.font.name = font_name
    style.font.size = font_size

    for s in doc.sections:
        s.top_margin = Inches(1.0)
        s.bottom_margin = Inches(1.0)
        s.left_margin = Inches(1.0)
        s.right_margin = Inches(1.0)

    p_title = doc.add_paragraph()
    r_title = p_title.add_run(title)
    r_title.bold = True
    r_title.font.size = Pt(16)
    r_title.font.color.rgb = RGBColor(0, 51, 102)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph("―" * 50)
    for block in content.split("\n\n"):
        if block.strip():
            p = doc.add_paragraph(block.strip())
            p.alignment = user_alignment

    stream = io.BytesIO()
    doc.save(stream)
    stream.seek(0)
    return stream

# -------------------------------------------------------------
# REALISTIC QA SCORING LOGIC
# -------------------------------------------------------------
def calculate_realistic_score(weight: float, related_errors: list) -> float:
    score = float(weight)
    for e in related_errors:
        sev = str(e.get("severity", "")).lower()
        if "critical" in sev:
            score -= (weight * 0.40)
        elif "high" in sev or "major" in sev:
            score -= (weight * 0.20)
        else:
            score -= (weight * 0.10)
    return max(0.0, round(score, 1))

# -------------------------------------------------------------
# STAGE 2: 12-QA EXCEL WORKBOOK GENERATOR
# -------------------------------------------------------------
def build_stage2_qa_excel(stage2_data: dict, university: str, program: str, deliverable: str) -> io.BytesIO:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        score = stage2_data.get("score", 92.0)
        coverage = stage2_data.get("coverage_pct", 100.0)
        critical_count = stage2_data.get("critical_count", 0)
        decision = stage2_data.get("decision", "PASS")
        
        summary_rows = [
            ["PARAMETER / FIELD", "METRIC / VALUE"],
            ["Institution / Governing University", university],
            ["Academic Program / Discipline", program],
            ["Deliverable Format Under Review", deliverable],
            ["Audit Framework Standard", "12-Parameter Academic Quality Assurance Framework"],
            ["Overall QA Quality Score", f"{score} / 100"],
            ["Syllabus Topic Coverage", f"{coverage}%"],
            ["Critical Defects Count", critical_count],
            ["Final Audit Determination", decision],
            ["Institutional Certification Status", "ACCREDITED FOR INSTRUCTION" if decision == "PASS" else "REQUIRES SME REVISION"]
        ]
        df_sum = pd.DataFrame(summary_rows[1:], columns=summary_rows[0])
        df_sum.to_excel(writer, sheet_name="Executive QA Summary", index=False)
        
        errs = stage2_data.get("error_register", [])
        matrix_rows = []
        for p in QA_12_PARAMETERS:
            related = [e for e in errs if p["name"].lower() in str(e).lower() or str(e.get("category", "")).lower() in p["name"].lower()]
            achieved = calculate_realistic_score(p["weight"], related)
            percent_achieved = round((achieved / p["weight"]) * 100, 1)

            if related:
                has_crit = any(str(e.get("severity", "")).lower() == "critical" for e in related)
                status = "REVISE" if has_crit else "OBSERVATION"
                findings = "; ".join([str(e.get("description", e.get("finding", "Defect identified"))) for e in related])
                remediation = "; ".join([str(e.get("recommendation", e.get("remediation", "Align with standard"))) for e in related])
            else:
                status = "COMPLIANT"
                findings = "Meets institutional pedagogical quality benchmarks."
                remediation = "None required. Maintain benchmark standard."

            matrix_rows.append({
                "Param #": p["id"],
                "QA Dimension": p["name"],
                "Weight": p["weight"],
                "Score Achieved": achieved,
                "Success (%)": f"{percent_achieved}%",
                "Compliance Status": status,
                "Auditor Findings": findings,
                "Remediation Guidance": remediation,
            })
        df_matrix = pd.DataFrame(matrix_rows)
        df_matrix.to_excel(writer, sheet_name="12 QA Parameter Matrix", index=False)

        if errs:
            df_errs = pd.DataFrame(errs)
        else:
            df_errs = pd.DataFrame([{
                "Issue ID": "LOG-00",
                "Category": "General QA",
                "Severity": "None",
                "Description": "Zero critical defects detected during evaluation.",
                "Remediation": "Content cleared for production."
            }])
        df_errs.to_excel(writer, sheet_name="Error Register", index=False)

        wb = writer.book
        h_fill = PatternFill(start_color="002060", end_color="002060", fill_type="solid")
        h_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        border = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
                        top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))

        for ws in wb.worksheets:
            for col in ws.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = min(max(max_len + 3, 15), 52)
            for cell in ws[1]:
                cell.fill = h_fill
                cell.font = h_font
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            for row in ws.iter_rows(min_row=2):
                for cell in row:
                    cell.border = border
                    cell.alignment = Alignment(vertical="center")

    output.seek(0)
    return output

# -------------------------------------------------------------
# STAGE 2: ACADEMIC QUALITY REVIEW REPORT DOCX BUILDER
# -------------------------------------------------------------
def build_stage2_qa_docx(stage2_data: dict, university: str, program: str, deliverable: str) -> io.BytesIO:
    doc = Document()
    for s in doc.sections:
        s.top_margin = Inches(0.8)
        s.bottom_margin = Inches(0.8)
        s.left_margin = Inches(0.8)
        s.right_margin = Inches(0.8)

    title_p = doc.add_paragraph()
    r_title = title_p.add_run("ACADEMIC CONTENT QUALITY REVIEW REPORT")
    r_title.bold = True
    r_title.font.size = Pt(16)
    r_title.font.color.rgb = RGBColor(0, 51, 102)
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    sub_p = doc.add_paragraph()
    r_sub = sub_p.add_run("Stage 2 SME Quality & 12-Parameter Pedagogical Audit Dossier")
    r_sub.font.size = Pt(10)
    r_sub.font.italic = True
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph("―" * 55)

    meta_tbl = doc.add_table(rows=3, cols=2)
    meta_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    score = stage2_data.get("score", 0.0)
    cov = stage2_data.get("coverage_pct", 0.0)
    dec = stage2_data.get("decision", "PASS")
    crit = stage2_data.get("critical_count", 0)

    meta_tbl.rows[0].cells[0].text = f"University: {university}"
    meta_tbl.rows[0].cells[1].text = f"Verification Status: {dec}"
    meta_tbl.rows[1].cells[0].text = f"Academic Program: {program}"
    meta_tbl.rows[1].cells[1].text = f"Quality Audit Score: {score} / 100"
    meta_tbl.rows[2].cells[0].text = f"Deliverable Format: {deliverable}"
    meta_tbl.rows[2].cells[1].text = f"Syllabus Coverage: {cov}% (Critical Errors: {crit})"

    doc.add_paragraph("")

    h1 = doc.add_heading("1. Executive Quality Scorecard", level=2)
    h1.style.font.color.rgb = RGBColor(0, 51, 102)
    p_card = doc.add_paragraph()
    p_card.add_run(f"• Overall Audit Score: {score}%\n")
    p_card.add_run(f"• Syllabus Topic Coverage: {cov}%\n")
    p_card.add_run(f"• Critical Academic Defects: {crit} detected\n")
    p_card.add_run(f"• Regulatory Status: {'VERIFIED FOR INSTITUTIONAL USE' if dec == 'PASS' else 'ACTION REQUIRED BY SME REVIEWER'}")

    h2 = doc.add_heading("2. 12-Parameter Academic Quality Matrix", level=2)
    h2.style.font.color.rgb = RGBColor(0, 51, 102)

    table = doc.add_table(rows=1, cols=6)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"

    headers = ["#", "QA Dimension", "Weight", "Achieved", "Status", "Auditor Findings"]
    for idx, text in enumerate(headers):
        c = table.rows[0].cells[idx]
        c.text = text
        for p in c.paragraphs:
            for r in p.runs:
                r.bold = True

    errs = stage2_data.get("error_register", [])
    for param in QA_12_PARAMETERS:
        related_errs = [e for e in errs if param["name"].lower() in str(e).lower() or str(e.get("category", "")).lower() in param["name"].lower()]
        achieved = calculate_realistic_score(param["weight"], related_errs)
        pct = round((achieved / param["weight"]) * 100, 1)

        status = "REVISE" if any(str(e.get("severity", "")).lower() == "critical" for e in related_errs) else ("OBSERVATION" if related_errs else "COMPLIANT")
        findings = "; ".join([str(e.get("description", e.get("finding", "Minor variance"))) for e in related_errs]) if related_errs else "Meets institutional pedagogical quality standards."

        r_cells = table.add_row().cells
        r_cells[0].text = str(param["id"])
        r_cells[1].text = param["name"]
        r_cells[2].text = f"{param['weight']}"
        r_cells[3].text = f"{achieved} ({pct}%)"
        r_cells[4].text = status
        r_cells[5].text = findings

    doc.add_paragraph("")

    h3 = doc.add_heading("3. Detailed Error & Remediation Register", level=2)
    h3.style.font.color.rgb = RGBColor(0, 51, 102)
    if errs:
        t_err = doc.add_table(rows=1, cols=4)
        t_err.alignment = WD_TABLE_ALIGNMENT.CENTER
        t_err.style = "Table Grid"
        for idx, text in enumerate(["Issue Category", "Severity", "Finding Description", "Corrective Action"]):
            c = t_err.rows[0].cells[idx]
            c.text = text
            for p in c.paragraphs:
                for r in p.runs:
                    r.bold = True
        for e in errs:
            rc = t_err.add_row().cells
            rc[0].text = str(e.get("category", "General"))
            rc[1].text = str(e.get("severity", "Medium"))
            rc[2].text = str(e.get("description", e.get("finding", "")))
            rc[3].text = str(e.get("recommendation", e.get("remediation", "Revise content")))
    else:
        doc.add_paragraph("Zero critical or major defects identified. Draft is compliant with university authoring guidelines.")

    doc.add_paragraph("\n" + "―" * 55)
    footer = doc.add_paragraph("Certified by University Academic Quality Board & SME Review Panel")
    footer.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    footer.style.font.italic = True
    footer.style.font.size = Pt(9)

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out

# -------------------------------------------------------------
# STAGE 3: PROMPT AUDITOR DOCX BUILDER
# -------------------------------------------------------------
def build_prompt_audit_docx(audit_data: dict, deliverable_type: str = "SLM", university: str = "State Open University", program: str = "MBA") -> io.BytesIO:
    doc = Document()
    for s in doc.sections:
        s.top_margin = Inches(0.8)
        s.bottom_margin = Inches(0.8)
        s.left_margin = Inches(0.8)
        s.right_margin = Inches(0.8)

    title_p = doc.add_paragraph()
    title_run = title_p.add_run("ACADEMIC PROMPT & INSTRUCTION AUDIT REPORT")
    title_run.bold = True
    title_run.font.size = Pt(16)
    title_run.font.color.rgb = RGBColor(0, 51, 102)
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    sub_p = doc.add_paragraph()
    sub_run = sub_p.add_run(f"Stage 3 Deep Meta-Audit Dossier | Deliverable: {deliverable_type}")
    sub_run.font.size = Pt(10)
    sub_run.font.italic = True
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph("―" * 55)

    meta_tbl = doc.add_table(rows=2, cols=2)
    meta_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_tbl.rows[0].cells[0].text = f"University: {university}"
    meta_tbl.rows[0].cells[1].text = f"Audit Status: {audit_data.get('overall_compliance_status', 'COMPLIANT')}"
    meta_tbl.rows[1].cells[0].text = f"Academic Program: {program}"
    meta_tbl.rows[1].cells[1].text = f"Prompt Score: {audit_data.get('prompt_score', 0.0)} / 100"

    doc.add_paragraph("")

    h1 = doc.add_heading("1. Executive Compliance Summary", level=2)
    h1.style.font.color.rgb = RGBColor(0, 51, 102)

    p_sum = doc.add_paragraph()
    p_sum.add_run("• Compliance Status: ").bold = True
    p_sum.add_run(f"{audit_data.get('overall_compliance_status', 'COMPLIANT')}\n")
    p_sum.add_run("• Prompt Fidelity Score: ").bold = True
    p_sum.add_run(f"{audit_data.get('prompt_score', 0.0)}%\n")
    p_sum.add_run("• Negative Constraint Violations: ").bold = True
    p_sum.add_run(f"{audit_data.get('negative_constraint_violations', 0)} detected\n")

    h2 = doc.add_heading("2. Pedagogical & Prompt Compliance Matrix", level=2)
    h2.style.font.color.rgb = RGBColor(0, 51, 102)

    checks = audit_data.get("compliance_checks", [])
    if checks:
        tbl = doc.add_table(rows=1, cols=4)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.style = "Table Grid"

        headers = ["Dimension", "Status", "Severity", "Auditor Findings"]
        for idx, text in enumerate(headers):
            cell = tbl.rows[0].cells[idx]
            cell.text = text
            for p in cell.paragraphs:
                for r in p.runs:
                    r.bold = True

        for item in checks:
            row = tbl.add_row().cells
            row[0].text = str(item.get("dimension", ""))
            row[1].text = str(item.get("status", ""))
            row[2].text = str(item.get("severity", "None"))
            row[3].text = str(item.get("finding", ""))

    doc.add_paragraph("")

    h3 = doc.add_heading("3. Structural Drift & Guardrail Analysis", level=2)
    h3.style.font.color.rgb = RGBColor(0, 51, 102)
    doc.add_paragraph(audit_data.get("structural_drift_analysis", "Zero structural drift detected."))

    doc.add_paragraph("\n" + "―" * 55)
    footer = doc.add_paragraph("Certified by University Academic AI Meta-Auditor")
    footer.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    footer.style.font.italic = True
    footer.style.font.size = Pt(9)

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out

# -------------------------------------------------------------
# DYNAMIC ACCESS CONTROL & SESSION STATE
# -------------------------------------------------------------
access_cfg = load_access_config()

if "stage" not in st.session_state:
    st.session_state.stage = "INPUT"
if "draft_text" not in st.session_state:
    st.session_state.draft_text = ""
if "stage2_results" not in st.session_state:
    st.session_state.stage2_results = None
if "stage3_results" not in st.session_state:
    st.session_state.stage3_results = None

# ==========================================
# SIDEBAR: STRICT AUTHENTICATION GATE (@upgrad.com ENFORCED)
# ==========================================
st.sidebar.header("🔐 User Authentication")
user_email = st.sidebar.text_input(
    "Enter Official Email ID (@upgrad.com)", 
    value="admin@upgrad.com"
).strip().lower()

if not user_email.endswith("@upgrad.com"):
    st.sidebar.error("⛔ Access Denied: Only official '@upgrad.com' email accounts are permitted.")
    st.warning("Please enter your official @upgrad.com faculty or admin email address.")
    st.stop()

is_authorized, user_role = verify_user(user_email)

if not is_authorized:
    st.sidebar.error("⛔ Access Denied: Unauthorized Account")
    st.warning("This @upgrad.com account is not registered. Please contact the Lead Administrator to grant your team access.")
    st.stop()

st.sidebar.success(f"Verified Role: {user_role}")

# -------------------------------------------------------------
# CONSOLIDATED TEAM PERMISSIONS
# -------------------------------------------------------------
can_run_maker = user_role in [
    "Lead Administrator",
    "Content Author (Maker Team)",
    "Content Author (Maker)"
]

can_run_checker = user_role in [
    "Lead Administrator",
    "Academic Reviewer (Checker & Auditor Team)",
    "SME Quality Reviewer (Checker)"
]

can_run_auditor = user_role in [
    "Lead Administrator",
    "Academic Reviewer (Checker & Auditor Team)",
    "Prompt Compliance Auditor"
]

# ==========================================
# SIDEBAR: ADMIN CONSOLE (Lead Administrator Only)
# ==========================================
if user_role == "Lead Administrator":
    with st.sidebar.expander("🛠️ Admin Console (Institutions & Users)", expanded=False):
        st.markdown("##### 🏛️ Add University")
        new_uni_name = st.text_input("University Name", placeholder="e.g., Apex Global University", key="admin_add_uni_input")
        if st.button("➕ Register University", use_container_width=True):
            if new_uni_name.strip():
                try:
                    add_university(new_uni_name.strip())
                    st.success(f"Registered: {new_uni_name.strip()}")
                    st.rerun()
                except Exception as e:
                    st.error(f"Failed to add university: {e}")
            else:
                st.warning("Enter a valid university name.")

        st.divider()
        st.markdown("##### 👤 Manage Team Roles & Logins")
        new_user_email = st.text_input(
            "Official Email (@upgrad.com)", 
            placeholder="e.g., faculty@upgrad.com", 
            key="admin_add_user_email"
        )
        new_user_role = st.selectbox(
            "Assign Team Role",
            [
                "Content Author (Maker Team)",
                "Academic Reviewer (Checker & Auditor Team)",
                "Lead Administrator"
            ],
            key="admin_add_user_role_select"
        )
        if st.button("➕ Save Team Permission", use_container_width=True):
            cleaned_email = new_user_email.strip().lower()
            if cleaned_email:
                if not cleaned_email.endswith("@upgrad.com"):
                    st.error("Invalid Domain: Only @upgrad.com email addresses can be authorized.")
                else:
                    try:
                        assign_user_role(cleaned_email, new_user_role)
                        st.success(f"Role saved for {cleaned_email} as '{new_user_role}'")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Failed to assign role: {e}")
            else:
                st.warning("Enter a valid official email address.")

        if isinstance(access_cfg, dict):
            with st.popover("📋 View Registered Directory"):
                st.markdown("**Registered Universities:**")
                st.write(access_cfg.get("universities", []))
                st.markdown("**Authorized User Directory:**")
                st.write(access_cfg.get("users", {}))

st.sidebar.divider()
st.sidebar.header("🏛️ Institutional Scope")

known_unis = ["State Open University", "National Distance University", "Central Higher Education Institute"]
if isinstance(access_cfg, dict) and "universities" in access_cfg and access_cfg["universities"]:
    known_unis = access_cfg["universities"]

selected_university = st.sidebar.selectbox("Select University", known_unis, index=0)

selected_program = st.sidebar.selectbox(
    "Program / Discipline",
    [
        "Master of Business Administration (MBA)",
        "Master of Science (M.Sc) Data Science",
        "Master of Computer Applications (MCA)",
        "Bachelor of Commerce (B.Com)"
    ],
    index=0
)

selected_deliverable = st.sidebar.selectbox(
    "Deliverable Format",
    ["SLM", "PPT", "Question Bank", "MCQ", "FAQs & Misconceptions Review"],
    index=0
)

selected_level = st.sidebar.selectbox(
    "Academic Level",
    ["Postgraduate (PG)", "Undergraduate (UG)", "Diploma / Executive"],
    index=0
)

# -------------------------------------------------------------
# MAIN APP HEADER & INPUT INGESTION
# -------------------------------------------------------------
st.title("University Unified 3-Tier Academic Content & Audit Pipeline")
st.caption("Layered Execution: AI Maker (1) ➔ Academic QA Checker (2) ➔ Deep Prompt Auditor (3)")

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Syllabus Outline")
    uploaded_syllabus = st.file_uploader("Upload Syllabus Outline (.docx, .txt)", type=["docx", "txt"], key="uploader_syl")
    manual_syllabus = st.text_area("Or Paste Syllabus Manually", placeholder="Paste topics or unit syllabus here...", height=120)
    
    syllabus_text = ""
    if uploaded_syllabus:
        syllabus_text = extract_text_safely(uploaded_syllabus)
        st.success(f"Loaded: {uploaded_syllabus.name}")
    elif manual_syllabus.strip():
        syllabus_text = manual_syllabus.strip()

with col2:
    st.subheader("2. Reference & Benchmark Material")
    uploaded_ref = st.file_uploader("Upload Prescribed References (.docx, .txt)", type=["docx", "txt"], key="uploader_ref")
    manual_ref = st.text_area("Or Paste References Manually", placeholder="Paste source extracts or textbook notes...", height=120)
    
    ref_text = ""
    if uploaded_ref:
        ref_text = extract_text_safely(uploaded_ref)
        st.success(f"Loaded: {uploaded_ref.name}")
    elif manual_ref.strip():
        ref_text = manual_ref.strip()

st.divider()

# -------------------------------------------------------------
# 3. CREATIVE CONTENT INJECTION (Maker Scope)
# -------------------------------------------------------------
st.subheader("3. Creative Direction & Customization (Maker Scope)")
creative_instructions = st.text_area(
    "Add Creative Content, Case Studies, or Custom Guidelines (Optional)",
    placeholder="E.g., Include a real-world startup case study, or structure key takeaways around role-play activities...",
    height=80
)

st.divider()

# -------------------------------------------------------------
# SEQUENTIAL PIPELINE CONTROLS
# -------------------------------------------------------------
st.subheader("Sequential Pipeline Execution")

btn_col1, btn_col2, btn_col3 = st.columns(3)

with btn_col1:
    st.markdown("#### Stage 1: AI Maker")
    run_maker = st.button(
        "Step 1: Run AI Maker",
        type="primary",
        disabled=(not can_run_maker),
        help="Execute AI drafting (Maker Team)" if can_run_maker else "Restricted to Maker Team / Admin",
        use_container_width=True
    )

with btn_col2:
    st.markdown("#### Stage 2: Quality Checker")
    run_checker = st.button(
        "Step 2: Run Academic Audit",
        disabled=(not can_run_checker or not bool(st.session_state.draft_text)),
        help="Execute Academic QA Audit (Reviewer Team)" if can_run_checker else "Restricted to Reviewer Team / Admin",
        use_container_width=True
    )

with btn_col3:
    st.markdown("#### Stage 3: Prompt Auditor")
    run_auditor = st.button(
        "Step 3: Run Prompt Audit",
        disabled=(not can_run_auditor or not bool(st.session_state.draft_text)),
        help="Execute Deep Prompt Meta-Audit (Reviewer Team)" if can_run_auditor else "Restricted to Reviewer Team / Admin",
        use_container_width=True
    )

# --- EXECUTE STAGE 1 ---
if run_maker:
    if not syllabus_text:
        st.error("Please provide an approved Syllabus outline first.")
    else:
        with st.spinner("Stage 1 generation in progress (streaming enabled)..."):
            try:
                final_syllabus = syllabus_text
                if creative_instructions.strip():
                    final_syllabus += f"\n\n[AUTHOR CREATIVE DIRECTIONS TO INTEGRATE]:\n{creative_instructions.strip()}"

                draft = run_stage_1_maker(
                    syllabus=final_syllabus,
                    reference=ref_text,
                    level=selected_level,
                    deliverable=selected_deliverable
                )
                st.session_state.draft_text = draft
                st.session_state.stage2_results = None
                st.session_state.stage3_results = None
                st.session_state.stage = "STAGE_1_COMPLETE"
                st.success("Stage 1 Content Generation Complete!")
            except Exception as e:
                st.error(f"Maker Failed: {e}")

# --- EXECUTE STAGE 2 ---
if run_checker:
    with st.spinner("Stage 2 Academic & Mapping QA Audit running..."):
        try:
            res_checker = run_stage_2_checker(
                draft=st.session_state.draft_text,
                syllabus=syllabus_text,
                reference=ref_text
            )
            st.session_state.stage2_results = res_checker
            st.session_state.stage = "STAGE_2_COMPLETE"
            st.success("Stage 2 Academic Quality Audit Complete!")
        except Exception as e:
            st.error(f"Checker Failed: {e}")

# --- EXECUTE STAGE 3 ---
if run_auditor:
    with st.spinner("Stage 3 Deep Prompt & Guardrail Audit running..."):
        try:
            res_auditor = run_stage_3_prompt_audit(
                draft=st.session_state.draft_text,
                deliverable=selected_deliverable
            )
            st.session_state.stage3_results = res_auditor
            st.session_state.stage = "STAGE_3_COMPLETE"
            st.success("Stage 3 Prompt Compliance Audit Complete!")
        except Exception as e:
            st.error(f"Auditor Failed: {e}")

st.divider()

# -------------------------------------------------------------
# STAGE 1 RESULTS & MANUSCRIPT EXPORT
# -------------------------------------------------------------
if st.session_state.draft_text:
    st.subheader(f"Stage 1 Deliverable: {selected_deliverable} Manuscript")
    with st.expander("📄 Click to Preview Generated Manuscript", expanded=True):
        st.markdown(st.session_state.draft_text)

    manuscript_docx = build_manuscript_docx(
        title=f"{selected_university} - {selected_deliverable}",
        content=st.session_state.draft_text,
        ref_file=uploaded_ref
    )
    st.download_button(
        label=f"📥 Download {selected_deliverable} Manuscript (.DOCX)",
        data=manuscript_docx,
        file_name=f"{selected_program}_{selected_deliverable}_Manuscript.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

# -------------------------------------------------------------
# STAGE 2 RESULTS, 12-QA MATRIX & DUAL EXPORTS (DOCX + XLSX)
# -------------------------------------------------------------
if st.session_state.stage2_results:
    st.divider()
    st.subheader("Stage 2: Academic Quality & Syllabus Verification (12 QA Framework)")
    s2 = st.session_state.stage2_results

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Audit Score", f"{s2.get('score', 0.0)} / 100")
    m2.metric("Syllabus Coverage", f"{s2.get('coverage_pct', 0.0)}%")
    m3.metric("Critical Errors", s2.get("critical_count", 0))
    status_dec = s2.get("decision", "PENDING")
    m4.metric("Verification Status", status_dec)

    if status_dec == "PASS":
        st.success(f"Audit Status: {status_dec} — Meets institutional quality benchmarks.")
    else:
        st.warning(f"Audit Status: {status_dec} — Revision recommended before publishing.")

    with st.expander("📋 View 12-Parameter Academic Quality Matrix", expanded=True):
        matrix_display = []
        errs = s2.get("error_register", [])
        for p in QA_12_PARAMETERS:
            related = [e for e in errs if p["name"].lower() in str(e).lower() or str(e.get("category", "")).lower() in p["name"].lower()]
            
            achieved = calculate_realistic_score(p["weight"], related)
            percent_achieved = round((achieved / p["weight"]) * 100, 1)

            status = "REVISE" if any(str(e.get("severity", "")).lower() == "critical" for e in related) else ("OBSERVATION" if related else "COMPLIANT")
            findings = "; ".join([str(e.get("description", e.get("finding", "Defect identified"))) for e in related]) if related else "Meets quality standards."
            
            matrix_display.append({
                "#": p["id"],
                "QA Dimension": p["name"],
                "Weight": p["weight"],
                "Achieved": achieved,
                "Success (%)": f"{percent_achieved}%",
                "Compliance Status": status,
                "Auditor Findings": findings
            })
        st.dataframe(pd.DataFrame(matrix_display), use_container_width=True)

    if errs:
        st.markdown("##### Quality Checker Error Register")
        st.dataframe(pd.DataFrame(errs), use_container_width=True)

    s2_col1, s2_col2 = st.columns(2)

    with s2_col1:
        s2_docx_stream = build_stage2_qa_docx(
            stage2_data=s2,
            university=selected_university,
            program=selected_program,
            deliverable=selected_deliverable
        )
        st.download_button(
            label="📄 Download Stage 2 QA Review Report (.DOCX)",
            data=s2_docx_stream,
            file_name=f"{selected_program}_{selected_deliverable}_QA_Review_Report.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True
        )

    with s2_col2:
        s2_excel_stream = build_stage2_qa_excel(
            stage2_data=s2,
            university=selected_university,
            program=selected_program,
            deliverable=selected_deliverable
        )
        st.download_button(
            label="📊 Download 12-QA Evaluation Report (.XLSX)",
            data=s2_excel_stream,
            file_name=f"{selected_program}_{selected_deliverable}_12_QA_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

# -------------------------------------------------------------
# STAGE 3 RESULTS & PROMPT AUDIT DOCX EXPORT
# -------------------------------------------------------------
if st.session_state.stage3_results:
    st.divider()
    st.subheader("Stage 3: Deep Prompt & Instruction Compliance Audit")
    s3 = st.session_state.stage3_results

    p1, p2, p3 = st.columns(3)
    p1.metric("Compliance Status", s3.get("overall_compliance_status", "COMPLIANT"))
    p2.metric("Prompt Score", f"{s3.get('prompt_score', 0.0)}%")
    p3.metric("Negative Constraint Violations", s3.get("negative_constraint_violations", 0))

    compliance_matrix = s3.get("compliance_checks", [])
    if compliance_matrix:
        st.markdown("##### Meta-Audit Compliance Matrix")
        st.dataframe(pd.DataFrame(compliance_matrix), use_container_width=True)

    st.markdown("##### Structural Drift Analysis")
    st.info(s3.get("structural_drift_analysis", "Zero structural drift detected."))

    docx_stream = build_prompt_audit_docx(
        audit_data=s3,
        deliverable_type=selected_deliverable,
        university=selected_university,
        program=selected_program
    )
    st.download_button(
        label="📄 Download Stage 3 Prompt Audit Report (.DOCX)",
        data=docx_stream,
        file_name=f"{selected_program}_{selected_deliverable}_Prompt_Audit_Report.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary",
        use_container_width=True
    )