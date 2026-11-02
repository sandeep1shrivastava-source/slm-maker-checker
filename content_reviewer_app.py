import io
import json
import os
import zipfile
import streamlit as st
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

# Modular Project Services
from services.access_manager import load_access_config, verify_user
from services.unified_engine import run_stage_1_maker, run_stage_2_checker, run_stage_3_prompt_audit
from services.depository_manager import list_manuscripts, update_manuscript_review

# -------------------------------------------------------------
# PAGE CONFIGURATION
# -------------------------------------------------------------
st.set_page_config(
    page_title="Academic Reviewer & Quality Audit Pipeline",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -------------------------------------------------------------
# 12-PARAMETER ACADEMIC QA FRAMEWORK
# -------------------------------------------------------------
QA_12_PARAMETERS = [
    {"id": 1, "name": "Syllabus & Curriculum Mapping", "weight": 10, "desc": "Alignment of draft topics to prescribed syllabus."},
    {"id": 2, "name": "Bloom's Taxonomy & Action Verbs", "weight": 10, "desc": "Adequacy of cognitive verbs across outcomes and assessments."},
    {"id": 3, "name": "Conceptual Depth & Pedagogical Rigor", "weight": 10, "desc": "Technical depth and academic sufficiency."},
    {"id": 4, "name": "Real-world Case Studies & Applications", "weight": 8, "desc": "Illustrative real-world examples and corporate scenarios."},
    {"id": 5, "name": "Quadrant-II Structural Architecture", "weight": 8, "desc": "Compliance with UGC/DEB self-learning material flow."},
    {"id": 6, "name": "Self-Assessment & Review Questions", "weight": 8, "desc": "Quality of formative questions, MCQs, and exercises."},
    {"id": 7, "name": "Factual & Theoretical Accuracy", "weight": 10, "desc": "Verification of domain formulas, theorems, dates, and claims."},
    {"id": 8, "name": "Originality, Attribution & Non-Hallucination", "weight": 10, "desc": "Zero unverified AI assertions or phantom citations."},
    {"id": 9, "name": "Academic Tone, Style & Language Clarity", "weight": 6, "desc": "Objective scholarly voice, clarity, and grammatical precision."},
    {"id": 10, "name": "Key Takeaways & Summary Synthesis", "weight": 6, "desc": "Clarity of unit summaries and concept recaps."},
    {"id": 11, "name": "Benchmark & Reference Adherence", "weight": 8, "desc": "Direct correlation to prescribed benchmark texts."},
    {"id": 12, "name": "Pedagogical Scaffolding & Formatting", "weight": 6, "desc": "Structural formatting, callouts, and visual aids."}
]

DEFAULT_PROGRAMS = [
    "Master of Business Administration (MBA)",
    "Master of Science (M.Sc) Data Science",
    "Master of Computer Applications (MCA)",
    "Bachelor of Commerce (B.Com)"
]

# -------------------------------------------------------------
# CONFIG & DIRECTORY MANAGEMENT HELPERS
# -------------------------------------------------------------
def get_config_file_path():
    p1 = os.path.join(os.path.dirname(__file__), "services", "access_config.json")
    if os.path.exists(p1):
        return p1
    p2 = os.path.join(os.path.dirname(__file__), "access_config.json")
    if os.path.exists(p2):
        return p2
    return p1

def save_role_to_config(email: str, role: str):
    path = get_config_file_path()
    cfg = {}
    if os.path.exists(path):
        try:
            with open(path, "r") as f:
                cfg = json.load(f)
        except Exception:
            cfg = {}
    
    users = cfg.setdefault("users", {})
    if isinstance(users, dict):
        users[email.strip().lower()] = role
    
    with open(path, "w") as f:
        json.dump(cfg, f, indent=4)

def save_university_to_config(uni_name: str):
    path = get_config_file_path()
    cfg = {}
    if os.path.exists(path):
        try:
            with open(path, "r") as f:
                cfg = json.load(f)
        except Exception:
            cfg = {}
            
    unis = cfg.setdefault("universities", [])
    if uni_name.strip() not in unis:
        unis.append(uni_name.strip())
        with open(path, "w") as f:
            json.dump(cfg, f, indent=4)

# -------------------------------------------------------------
# DOCUMENT EXTRACTION UTILITIES
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
# REPORT GENERATORS
# -------------------------------------------------------------
def build_stage2_qa_excel(stage2_data: dict, university: str, program: str, deliverable: str, batch: str = "", semester_year: str = "") -> io.BytesIO:
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
            ["Batch / Cohort", batch or "N/A"],
            ["Academic Term / Semester", semester_year or "N/A"],
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

def build_stage2_qa_docx(stage2_data: dict, university: str, program: str, deliverable: str, batch: str = "", semester_year: str = "") -> io.BytesIO:
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
    r_sub = sub_p.add_run(f"Stage 2 SME Quality & 12-Parameter Pedagogical Audit | {batch} ({semester_year})")
    r_sub.font.size = Pt(10)
    r_sub.font.italic = True
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph("―" * 55)

    meta_tbl = doc.add_table(rows=4, cols=2)
    meta_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    score = stage2_data.get("score", 0.0)
    cov = stage2_data.get("coverage_pct", 0.0)
    dec = stage2_data.get("decision", "PASS")
    crit = stage2_data.get("critical_count", 0)

    meta_tbl.rows[0].cells[0].text = f"University: {university}"
    meta_tbl.rows[0].cells[1].text = f"Verification Status: {dec}"
    meta_tbl.rows[1].cells[0].text = f"Academic Program: {program}"
    meta_tbl.rows[1].cells[1].text = f"Quality Audit Score: {score} / 100"
    meta_tbl.rows[2].cells[0].text = f"Batch / Cohort: {batch}"
    meta_tbl.rows[2].cells[1].text = f"Academic Term: {semester_year}"
    meta_tbl.rows[3].cells[0].text = f"Deliverable Format: {deliverable}"
    meta_tbl.rows[3].cells[1].text = f"Syllabus Coverage: {cov}% (Critical Errors: {crit})"

    doc.add_paragraph("")
    h1 = doc.add_heading("1. Executive Quality Scorecard", level=2)
    h1.style.font.color.rgb = RGBColor(0, 51, 102)
    p_card = doc.add_paragraph()
    p_card.add_run(f"• Overall Audit Score: {score}%\n")
    p_card.add_run(f"• Syllabus Topic Coverage: {cov}%\n")
    p_card.add_run(f"• Critical Academic Defects: {crit} detected\n")
    p_card.add_run(f"• Determination: {'VERIFIED FOR INSTRUCTION' if dec == 'PASS' else 'ACTION REQUIRED BY SME REVIEWER'}")

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
        findings = "; ".join([str(e.get("description", e.get("finding", "Minor variance"))) for e in related_errs]) if related_errs else "Meets quality standards."

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
        doc.add_paragraph("Zero critical defects identified. Manuscript approved for delivery.")

    doc.add_paragraph("\n" + "―" * 55)
    footer = doc.add_paragraph("Certified by University Academic Quality Board & SME Review Panel")
    footer.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    footer.style.font.italic = True
    footer.style.font.size = Pt(9)

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out

def build_prompt_audit_docx(audit_data: dict, deliverable_type: str = "SLM", university: str = "State Open University", program: str = "MBA", batch: str = "", semester_year: str = "") -> io.BytesIO:
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
    sub_run = sub_p.add_run(f"Stage 3 Meta-Audit Dossier | {university} - {program} | {batch} ({semester_year})")
    sub_run.font.size = Pt(10)
    sub_run.font.italic = True
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph("―" * 55)

    meta_tbl = doc.add_table(rows=3, cols=2)
    meta_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_tbl.rows[0].cells[0].text = f"University: {university}"
    meta_tbl.rows[0].cells[1].text = f"Audit Status: {audit_data.get('overall_compliance_status', 'COMPLIANT')}"
    meta_tbl.rows[1].cells[0].text = f"Academic Program: {program}"
    meta_tbl.rows[1].cells[1].text = f"Prompt Score: {audit_data.get('prompt_score', 0.0)} / 100"
    meta_tbl.rows[2].cells[0].text = f"Batch: {batch}"
    meta_tbl.rows[2].cells[1].text = f"Term: {semester_year}"

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

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out

def build_zip_audit_package(s2_docx, s2_xlsx, s3_docx, prefix: str) -> io.BytesIO:
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        if s2_docx:
            zf.writestr(f"{prefix}_Stage2_QA_Report.docx", s2_docx.getvalue())
        if s2_xlsx:
            zf.writestr(f"{prefix}_Stage2_12_QA_Evaluation.xlsx", s2_xlsx.getvalue())
        if s3_docx:
            zf.writestr(f"{prefix}_Stage3_Prompt_Audit_Report.docx", s3_docx.getvalue())
    zip_buffer.seek(0)
    return zip_buffer

# -------------------------------------------------------------
# SESSION STATE & AUTHENTICATION
# -------------------------------------------------------------
access_cfg = load_access_config()

if "draft_text" not in st.session_state:
    st.session_state.draft_text = ""
if "stage2_results" not in st.session_state:
    st.session_state.stage2_results = None
if "stage3_results" not in st.session_state:
    st.session_state.stage3_results = None
if "loaded_manuscript_id" not in st.session_state:
    st.session_state.loaded_manuscript_id = None
if "active_batch" not in st.session_state:
    st.session_state.active_batch = "Batch 2026-A"
if "active_term" not in st.session_state:
    st.session_state.active_term = "Semester 1"

st.sidebar.header("🔐 User Authentication")
user_email = st.sidebar.text_input(
    "Enter Official Email ID (@upgrad.com)", 
    value="sandeep1.shrivastava@upgrad.com"
).strip().lower()

if not user_email.endswith("@upgrad.com"):
    st.sidebar.error("⛔ Access Denied: Only official '@upgrad.com' email accounts permitted.")
    st.stop()

is_authorized, user_role = verify_user(user_email)
if not is_authorized:
    st.sidebar.error("⛔ Access Denied: Unauthorized Account.")
    st.stop()

st.sidebar.success(f"Verified Role: **{user_role}**")

can_run_maker = user_role in ["Lead Administrator", "Content Author (Maker Team)", "Content Author (Maker)"]
can_run_checker = user_role in ["Lead Administrator", "Academic Reviewer (Checker & Auditor Team)", "SME Quality Reviewer (Checker)"]
can_run_auditor = user_role in ["Lead Administrator", "Academic Reviewer (Checker & Auditor Team)", "Prompt Compliance Auditor"]

# -------------------------------------------------------------
# RESTORED: LEAD ADMINISTRATOR CONSOLE
# -------------------------------------------------------------
if user_role == "Lead Administrator":
    with st.sidebar.expander("🛠️ Admin Console (Institutions & User Roles)", expanded=True):
        st.markdown("##### 👤 Manage Team Roles & Permissions")
        target_colleague_email = st.text_input("Colleague Email (@upgrad.com)", placeholder="colleague@upgrad.com", key="admin_colleague_email")
        assigned_role = st.selectbox(
            "Assign Team Role",
            [
                "Academic Reviewer (Checker & Auditor Team)",
                "Content Author (Maker Team)",
                "Lead Administrator"
            ],
            key="admin_role_select"
        )
        if st.button("➕ Save Team Permission", use_container_width=True):
            clean_email = target_colleague_email.strip().lower()
            if clean_email.endswith("@upgrad.com"):
                save_role_to_config(clean_email, assigned_role)
                st.success(f"Saved: {clean_email} ➔ {assigned_role}")
                st.rerun()
            else:
                st.warning("Please enter a valid @upgrad.com email.")

        st.divider()
        st.markdown("##### 🏛️ Register University")
        new_uni_name = st.text_input("New University Name", placeholder="e.g., Apex University", key="admin_add_uni")
        if st.button("➕ Add University", use_container_width=True):
            if new_uni_name.strip():
                save_university_to_config(new_uni_name.strip())
                st.success(f"Added: {new_uni_name.strip()}")
                st.rerun()
            else:
                st.warning("Please enter university name.")

        with st.popover("📋 View Registered Directory"):
            cfg_now = load_access_config()
            st.json(cfg_now.get("users", {}))

st.sidebar.divider()
st.sidebar.header("🏛️ Audit Scope Mapping")

known_unis = ["State Open University", "National Distance University", "Central Higher Education Institute"]
if isinstance(access_cfg, dict) and "universities" in access_cfg and access_cfg["universities"]:
    known_unis = access_cfg["universities"]

selected_university = st.sidebar.selectbox("University", known_unis, index=0)

prog_choices = DEFAULT_PROGRAMS + ["➕ Other / Custom Programme..."]
selected_program_choice = st.sidebar.selectbox("Program / Discipline", prog_choices, index=0)

if selected_program_choice == "➕ Other / Custom Programme...":
    custom_prog = st.sidebar.text_input("Enter Programme Name", placeholder="e.g., Master in AI").strip()
    selected_program = custom_prog if custom_prog else "Custom Academic Programme"
else:
    selected_program = selected_program_choice

selected_batch = st.sidebar.text_input("Batch / Cohort", value=st.session_state.active_batch)
selected_term = st.sidebar.selectbox("Semester / Term", ["Semester 1", "Semester 2", "Semester 3", "Semester 4", "Year 1", "Year 2"], index=0)
selected_deliverable = st.sidebar.selectbox("Deliverable Format Under Review", ["SLM", "PPT", "Question Bank", "MCQ", "FAQs & Misconceptions Review"], index=0)

# -------------------------------------------------------------
# MAIN APP: REVIEW-FIRST INTERFACE
# -------------------------------------------------------------
st.title("Academic Quality Assurance & Review Console")
st.caption(f"Reviewer Mode: **{user_email}** ({user_role}) | Target: **{selected_university}** ➔ **{selected_program}**")

# Depository Browser
all_manuscripts = list_manuscripts()
with st.expander("📚 Load Manuscript from Shared Depository", expanded=False):
    if all_manuscripts:
        dep_options = {
            f"[{m.get('status', 'Pending')}] {m.get('university')} ➔ {m.get('program')} ➔ {m.get('batch')} ➔ {m.get('deliverable')} (Author: {m.get('author_email')})": m.get("id")
            for m in all_manuscripts
        }
        sel_label = st.selectbox("Select Manuscript from Storage Queue:", list(dep_options.keys()))
        if st.button("📥 Load Selected Draft into Reviewer Workspace"):
            chosen = next((m for m in all_manuscripts if m["id"] == dep_options[sel_label]), None)
            if chosen:
                st.session_state.draft_text = chosen.get("draft_text", "")
                st.session_state.loaded_manuscript_id = chosen["id"]
                st.session_state.stage2_results = chosen.get("stage2_results")
                st.session_state.stage3_results = chosen.get("stage3_results")
                st.success("Manuscript loaded into reviewer workspace!")
                st.rerun()
    else:
        st.info("No manuscripts in depository yet.")

# MAIN REVIEW WORKSPACE: INGESTION
st.markdown("---")
st.subheader("1. Ingest Content for Review & Audit")

tab_upload, tab_paste, tab_maker = st.tabs([
    "📂 Direct Upload (.DOCX / .TXT)", 
    "✍️ Paste Manuscript Text", 
    "🤖 Run Stage 1 AI Maker (Optional)"
])

with tab_upload:
    uploaded_draft = st.file_uploader("Upload Manuscript File to Audit Directly", type=["docx", "txt"], key="rev_uploader")
    if uploaded_draft:
        extracted = extract_text_safely(uploaded_draft)
        if extracted.strip():
            st.session_state.draft_text = extracted
            st.session_state.loaded_manuscript_id = None
            st.success(f"Loaded '{uploaded_draft.name}' ({len(extracted.split())} words ready for review).")

with tab_paste:
    pasted_draft = st.text_area("Paste Full Manuscript Text Here", value=st.session_state.draft_text, height=180, key="rev_paster")
    if st.button("Apply Pasted Text to Workspace"):
        st.session_state.draft_text = pasted_draft
        st.success("Draft updated.")

with tab_maker:
    if can_run_maker:
        st.caption("Generate a fresh draft from a syllabus outline:")
        maker_syl = st.text_area("Syllabus Outline", placeholder="Enter units and topics...", height=100)
        maker_ref = st.text_area("Reference Material", placeholder="Enter key textbook extracts...", height=80)
        if st.button("Generate Draft (Stage 1 Maker)", type="secondary"):
            if maker_syl.strip():
                with st.spinner("AI Maker generating draft..."):
                    st.session_state.draft_text = run_stage_1_maker(maker_syl, maker_ref, "Postgraduate (PG)", selected_deliverable)
                    st.success("Draft generated and placed into review workspace!")
                    st.rerun()
            else:
                st.warning("Please provide a syllabus outline.")
    else:
        st.info("Stage 1 generation is reserved for Content Authors and Lead Administrators.")

# Reference Syllabus for Verification
st.markdown("#### Prescribed Syllabus & Benchmark Reference (For Mapping Audit)")
col_syl, col_ref = st.columns(2)
with col_syl:
    audit_syl_file = st.file_uploader("Upload Prescribed Syllabus Outline (.docx, .txt)", type=["docx", "txt"], key="rev_syl_file")
    audit_syl_manual = st.text_area("Or Paste Syllabus Topics", placeholder="Unit 1: Concept, Unit 2: Frameworks...", height=100)
    audit_syl = extract_text_safely(audit_syl_file) if audit_syl_file else audit_syl_manual.strip()

with col_ref:
    audit_ref_file = st.file_uploader("Upload Benchmark Textbook Reference (.docx, .txt)", type=["docx", "txt"], key="rev_ref_file")
    audit_ref_manual = st.text_area("Or Paste Benchmark References", placeholder="Source references...", height=100)
    audit_ref = extract_text_safely(audit_ref_file) if audit_ref_file else audit_ref_manual.strip()

# -------------------------------------------------------------
# AUDIT EXECUTION CONTROLS
# -------------------------------------------------------------
st.markdown("---")
st.subheader("2. Execute Academic Review & Audit")

has_draft = bool(st.session_state.draft_text.strip())

col_btn1, col_btn2, col_btn3 = st.columns([1.5, 1, 1])

with col_btn1:
    run_all_audits = st.button(
        "⚡ Run Complete Academic & Prompt Audit (Stages 2 & 3)",
        type="primary",
        disabled=(not can_run_checker or not has_draft),
        use_container_width=True,
        help="Runs both the 12-parameter pedagogical audit and the deep prompt compliance evaluation."
    )

with col_btn2:
    run_stage2_only = st.button(
        "Step 2: 12-QA Checker Only",
        disabled=(not can_run_checker or not has_draft),
        use_container_width=True
    )

with col_btn3:
    run_stage3_only = st.button(
        "Step 3: Prompt Auditor Only",
        disabled=(not can_run_auditor or not has_draft),
        use_container_width=True
    )

# Execution Logic
if run_all_audits or run_stage2_only:
    with st.spinner("Executing Stage 2: 12-Parameter Academic Quality & Syllabus Audit..."):
        try:
            res_s2 = run_stage_2_checker(st.session_state.draft_text, audit_syl, audit_ref)
            st.session_state.stage2_results = res_s2
            if st.session_state.loaded_manuscript_id:
                update_manuscript_review(st.session_state.loaded_manuscript_id, user_email, stage2_results=res_s2)
            st.success("Stage 2 QA Audit complete!")
        except Exception as e:
            st.error(f"Stage 2 Failed: {e}")

if run_all_audits or run_stage3_only:
    with st.spinner("Executing Stage 3: Deep Prompt & Instruction Meta-Audit..."):
        try:
            res_s3 = run_stage_3_prompt_audit(st.session_state.draft_text, selected_deliverable)
            st.session_state.stage3_results = res_s3
            if st.session_state.loaded_manuscript_id:
                update_manuscript_review(st.session_state.loaded_manuscript_id, user_email, stage3_results=res_s3)
            st.success("Stage 3 Prompt Audit complete!")
        except Exception as e:
            st.error(f"Stage 3 Failed: {e}")

# -------------------------------------------------------------
# AUDIT SCORECARDS & DELIVERABLE DOWNLOADS
# -------------------------------------------------------------
s2 = st.session_state.stage2_results
s3 = st.session_state.stage3_results

if s2 or s3:
    st.markdown("---")
    st.subheader("3. Audit Results & Download Reports")
    
    file_prefix = f"{selected_university}_{selected_program}_{selected_batch}_{selected_term}".replace(" ", "_")

    stream_s2_docx = build_stage2_qa_docx(s2, selected_university, selected_program, selected_deliverable, selected_batch, selected_term) if s2 else None
    stream_s2_xlsx = build_stage2_qa_excel(s2, selected_university, selected_program, selected_deliverable, selected_batch, selected_term) if s2 else None
    stream_s3_docx = build_prompt_audit_docx(s3, selected_deliverable, selected_university, selected_program, selected_batch, selected_term) if s3 else None

    # Master Bundle
    if stream_s2_docx and stream_s2_xlsx and stream_s3_docx:
        zip_pkg = build_zip_audit_package(stream_s2_docx, stream_s2_xlsx, stream_s3_docx, file_prefix)
        st.download_button(
            label="📦 Download Complete Audit Package (.ZIP - All 3 Reports)",
            data=zip_pkg,
            file_name=f"{file_prefix}_Complete_Audit_Package.zip",
            mime="application/zip",
            type="primary",
            use_container_width=True
        )
        st.write("")

    col_d1, col_d2, col_d3 = st.columns(3)

    with col_d1:
        if stream_s2_xlsx:
            st.download_button(
                label="📊 12-QA Evaluation Workbook (.XLSX)",
                data=stream_s2_xlsx,
                file_name=f"{file_prefix}_12_QA_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    with col_d2:
        if stream_s2_docx:
            st.download_button(
                label="📄 Stage 2 QA Review Report (.DOCX)",
                data=stream_s2_docx,
                file_name=f"{file_prefix}_Stage2_QA_Report.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
    with col_d3:
        if stream_s3_docx:
            st.download_button(
                label="📄 Stage 3 Prompt Audit Dossier (.DOCX)",
                data=stream_s3_docx,
                file_name=f"{file_prefix}_Stage3_Prompt_Audit.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )

    tab_matrix, tab_s3, tab_preview = st.tabs([
        "📋 12-Parameter Matrix & Error Register", 
        "🔍 Prompt Compliance & Structural Drift", 
        "📄 Manuscript Content Preview"
    ])

    with tab_matrix:
        if s2:
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Audit Score", f"{s2.get('score', 0.0)} / 100")
            m2.metric("Syllabus Coverage", f"{s2.get('coverage_pct', 0.0)}%")
            m3.metric("Critical Defects", s2.get("critical_count", 0))
            m4.metric("Determination", s2.get("decision", "PASS"))

            matrix_data = []
            errs = s2.get("error_register", [])
            for p in QA_12_PARAMETERS:
                related = [e for e in errs if p["name"].lower() in str(e).lower() or str(e.get("category", "")).lower() in p["name"].lower()]
                achieved = calculate_realistic_score(p["weight"], related)
                status = "REVISE" if any(str(e.get("severity", "")).lower() == "critical" for e in related) else ("OBSERVATION" if related else "COMPLIANT")
                findings = "; ".join([str(e.get("description", e.get("finding", "Defect identified"))) for e in related]) if related else "Meets quality standard."
                matrix_data.append({
                    "#": p["id"],
                    "QA Dimension": p["name"],
                    "Weight": p["weight"],
                    "Achieved Score": achieved,
                    "Status": status,
                    "Auditor Findings": findings
                })
            st.dataframe(pd.DataFrame(matrix_data), use_container_width=True)

            if errs:
                st.markdown("##### Defect & Remediation Register")
                st.dataframe(pd.DataFrame(errs), use_container_width=True)

    with tab_s3:
        if s3:
            p1, p2, p3 = st.columns(3)
            p1.metric("Compliance Status", s3.get("overall_compliance_status", "COMPLIANT"))
            p2.metric("Prompt Score", f"{s3.get('prompt_score', 0.0)}%")
            p3.metric("Negative Violations", s3.get("negative_constraint_violations", 0))

            if s3.get("compliance_checks"):
                st.dataframe(pd.DataFrame(s3["compliance_checks"]), use_container_width=True)
            st.info(f"**Structural Drift Analysis:** {s3.get('structural_drift_analysis', 'Zero drift detected.')}")

    with tab_preview:
        st.markdown(st.session_state.draft_text)