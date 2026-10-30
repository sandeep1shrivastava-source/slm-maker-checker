import io
import json
import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

# Project Modular Services
from services.access_manager import load_access_config, verify_user
from services.unified_engine import run_stage_1_maker, run_stage_2_checker, run_stage_3_prompt_audit

try:
    from services.docx_handler import extract_text_from_docx, generate_slm_docx
except ImportError:
    def extract_text_from_docx(file):
        doc = Document(file)
        return "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
    def generate_slm_docx(title, content):
        doc = Document()
        doc.add_heading(title, 0)
        for b in content.split("\n\n"):
            if b.strip():
                doc.add_paragraph(b.strip())
        stream = io.BytesIO()
        doc.save(stream)
        stream.seek(0)
        return stream

try:
    from services.unified_report_generator import build_master_audit_excel
except ImportError:
    build_master_audit_excel = None

# -------------------------------------------------------------
# PAGE CONFIGURATION
# -------------------------------------------------------------
st.set_page_config(
    page_title="Unified Academic Pipeline (Maker-Checker-Prompt)",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -------------------------------------------------------------
# STAGE 3 PROMPT AUDIT DOCX GENERATOR
# -------------------------------------------------------------
def build_prompt_audit_docx(audit_data: dict, deliverable_type: str = "SLM", university: str = "State Open University", program: str = "MBA") -> io.BytesIO:
    """Generates an executive-grade Word document audit dossier for Stage 3 Prompt Compliance."""
    doc = Document()

    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

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
# SIDEBAR: STRICT AUTHENTICATION GATE
# ==========================================
st.sidebar.header("🔐 User Authentication")
user_email = st.sidebar.text_input("Enter Registered Email ID", value="admin@university.edu").strip().lower()

is_authorized, user_role = verify_user(user_email)

if not is_authorized:
    st.sidebar.error("⛔ Access Denied: Unauthorized Email")
    st.warning("Please enter a verified university faculty/admin email address to unlock the pipeline.")
    st.stop()

st.sidebar.success(f"Verified Role: {user_role}")

# Role-Based Permissions
can_run_maker = user_role in ["Lead Administrator", "Content Author (Maker)"]
can_run_checker = user_role in ["Lead Administrator", "SME Quality Reviewer (Checker)"]
can_run_auditor = user_role in ["Lead Administrator", "Prompt Compliance Auditor"]

st.sidebar.divider()
st.sidebar.header("🏛️ Institutional Scope")

# Load registered universities from dynamic config or defaults
known_unis = ["State Open University", "National Distance University", "Central Higher Education Institute"]
if isinstance(access_cfg, dict) and "universities" in access_cfg:
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
        syllabus_text = extract_text_from_docx(uploaded_syllabus) if uploaded_syllabus.name.lower().endswith(".docx") else uploaded_syllabus.read().decode("utf-8", errors="ignore")
        st.success(f"Loaded: {uploaded_syllabus.name}")
    elif manual_syllabus.strip():
        syllabus_text = manual_syllabus.strip()

with col2:
    st.subheader("2. Reference & Benchmark Material")
    uploaded_ref = st.file_uploader("Upload Prescribed References (.docx, .txt)", type=["docx", "txt"], key="uploader_ref")
    manual_ref = st.text_area("Or Paste References Manually", placeholder="Paste source extracts or textbook notes...", height=120)
    
    ref_text = ""
    if uploaded_ref:
        ref_text = extract_text_from_docx(uploaded_ref) if uploaded_ref.name.lower().endswith(".docx") else uploaded_ref.read().decode("utf-8", errors="ignore")
        st.success(f"Loaded: {uploaded_ref.name}")
    elif manual_ref.strip():
        ref_text = manual_ref.strip()

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
        help="Execute AI drafting" if can_run_maker else "Restricted by Role-Based Access Control",
        use_container_width=True
    )

with btn_col2:
    st.markdown("#### Stage 2: Quality Checker")
    run_checker = st.button(
        "Step 2: Run Academic Audit",
        disabled=(not can_run_checker or not bool(st.session_state.draft_text)),
        help="Execute Academic QA Audit" if can_run_checker else "Restricted by Role-Based Access Control",
        use_container_width=True
    )

with btn_col3:
    st.markdown("#### Stage 3: Prompt Auditor")
    run_auditor = st.button(
        "Step 3: Run Prompt Audit",
        disabled=(not can_run_auditor or not bool(st.session_state.draft_text)),
        help="Execute Deep Prompt Meta-Audit" if can_run_auditor else "Restricted by Role-Based Access Control",
        use_container_width=True
    )

# --- EXECUTE STAGE 1 ---
if run_maker:
    if not syllabus_text:
        st.error("Please provide an approved Syllabus outline first.")
    else:
        with st.spinner("Stage 1 generation in progress (streaming enabled)..."):
            try:
                draft = run_stage_1_maker(
                    syllabus=syllabus_text,
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

    manuscript_docx = generate_slm_docx(
        title=f"{selected_university} - {selected_deliverable}",
        content=st.session_state.draft_text
    )
    st.download_button(
        label=f"📥 Download {selected_deliverable} Manuscript (.DOCX)",
        data=manuscript_docx,
        file_name=f"{selected_program}_{selected_deliverable}_Manuscript.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

# -------------------------------------------------------------
# STAGE 2 RESULTS & ERROR REGISTER
# -------------------------------------------------------------
if st.session_state.stage2_results:
    st.divider()
    st.subheader("Stage 2: Academic Quality & Syllabus Verification")
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

    errs = s2.get("error_register", [])
    if errs:
        st.markdown("##### Quality Checker Error Register")
        st.dataframe(pd.DataFrame(errs), use_container_width=True)

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

    dl_col1, dl_col2 = st.columns(2)

    # DOCX Download Trigger
    with dl_col1:
        docx_stream = build_prompt_audit_docx(
            audit_data=s3,
            deliverable_type=selected_deliverable,
            university=selected_university,
            program=selected_program
        )
        st.download_button(
            label="📄 Download Stage 3 Audit Report (.DOCX)",
            data=docx_stream,
            file_name=f"{selected_program}_{selected_deliverable}_Prompt_Audit_Report.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True
        )

    # Optional Excel Master Audit Download Trigger (if available)
    with dl_col2:
        if build_master_audit_excel and st.session_state.stage2_results:
            try:
                excel_stream = build_master_audit_excel(
                    s2_results=st.session_state.stage2_results,
                    s3_results=s3,
                    university=selected_university,
                    program=selected_program,
                    deliverable=selected_deliverable
                )
                st.download_button(
                    label="📊 Download Master Audit Workbook (.XLSX)",
                    data=excel_stream,
                    file_name=f"{selected_program}_{selected_deliverable}_Master_Audit.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            except Exception:
                pass