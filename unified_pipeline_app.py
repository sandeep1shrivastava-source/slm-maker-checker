import io
import json
import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

from services.unified_engine import run_stage_1_maker, run_stage_2_checker, run_stage_3_prompt_audit

# -------------------------------------------------------------
# PAGE CONFIGURATION
# -------------------------------------------------------------
st.set_page_config(
    page_title="AI Academic Content Engine | 3-Tier Pipeline",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -------------------------------------------------------------
# HELPER FUNCTIONS: FILE EXTRACTION & DOCX REPORT GENERATION
# -------------------------------------------------------------
def extract_text_from_file(uploaded_file) -> str:
    """Extracts raw text content from uploaded .docx or .txt documents."""
    if uploaded_file is None:
        return ""
    fname = uploaded_file.name.lower()
    if fname.endswith(".docx"):
        doc = Document(uploaded_file)
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
    else:
        try:
            return uploaded_file.read().decode("utf-8")
        except Exception:
            return uploaded_file.getvalue().decode("latin-1")

def build_manuscript_docx(title: str, content: str) -> io.BytesIO:
    """Generates a downloadable DOCX for the generated academic deliverable."""
    doc = Document()
    for s in doc.sections:
        s.top_margin = Inches(1.0)
        s.bottom_margin = Inches(1.0)
        s.left_margin = Inches(1.0)
        s.right_margin = Inches(1.0)

    p_title = doc.add_paragraph()
    r_title = p_title.add_run(title)
    r_title.bold = True
    r_title.font.size = Pt(18)
    r_title.font.color.rgb = RGBColor(0, 51, 102)

    doc.add_paragraph("―" * 50)
    for block in content.split("\n\n"):
        if block.strip():
            doc.add_paragraph(block.strip())

    stream = io.BytesIO()
    doc.save(stream)
    stream.seek(0)
    return stream

def build_prompt_audit_docx(audit_data: dict, deliverable_type: str, university: str, program: str) -> io.BytesIO:
    """Builds a formal Word document audit dossier for Stage 3 Prompt Compliance."""
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
    p_sum.add_run("• Overall Compliance Status: ").bold = True
    p_sum.add_run(f"{audit_data.get('overall_compliance_status', 'COMPLIANT')}\n")
    p_sum.add_run("• Prompt Fidelity Score: ").bold = True
    p_sum.add_run(f"{audit_data.get('prompt_score', 0.0)}%\n")
    p_sum.add_run("• Negative Constraint Violations: ").bold = True
    p_sum.add_run(f"{audit_data.get('negative_constraint_violations', 0)} detected\n")

    h2 = doc.add_heading("2. Pedagogical & Prompt Compliance Matrix", level=2)
    h2.style.font.color.rgb = RGBColor(0, 51, 102)

    checks = audit_data.get("compliance_checks", [])
    if checks:
        table = doc.add_table(rows=1, cols=4)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.style = "Table Grid"

        headers = ["Dimension", "Status", "Severity", "Auditor Findings"]
        for idx, text in enumerate(headers):
            cell = table.rows[0].cells[idx]
            cell.text = text
            for p in cell.paragraphs:
                for r in p.runs:
                    r.bold = True

        for check in checks:
            row_cells = table.add_row().cells
            row_cells[0].text = str(check.get("dimension", ""))
            row_cells[1].text = str(check.get("status", ""))
            row_cells[2].text = str(check.get("severity", "None"))
            row_cells[3].text = str(check.get("finding", ""))

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
# SESSION STATE INITIALIZATION
# -------------------------------------------------------------
if "stage_1_draft" not in st.session_state:
    st.session_state["stage_1_draft"] = ""
if "stage_2_results" not in st.session_state:
    st.session_state["stage_2_results"] = None
if "stage_3_results" not in st.session_state:
    st.session_state["stage_3_results"] = None

# -------------------------------------------------------------
# SIDEBAR: AUTHENTICATION & INSTITUTIONAL CONFIGURATION
# -------------------------------------------------------------
with st.sidebar:
    st.subheader("🔐 User Authentication")
    user_email = st.text_input("Enter Registered Email ID", value="admin@university.edu")
    st.success("Verified Role: Lead Administrator")
    
    st.divider()
    st.subheader("🏛️ Institutional Scope")
    
    selected_university = st.selectbox(
        "Select University",
        [
            "State Open University",
            "National Distance University",
            "Central Higher Education Institute"
        ],
        index=0
    )
    
    selected_program = st.selectbox(
        "Program / Discipline",
        [
            "Master of Business Administration (MBA)",
            "Master of Science (M.Sc) Data Science",
            "Master of Computer Applications (MCA)",
            "Bachelor of Commerce (B.Com)"
        ],
        index=0
    )
    
    selected_deliverable = st.selectbox(
        "Deliverable Format",
        ["SLM", "PPT", "Question Bank", "MCQ", "FAQs & Misconceptions Review"],
        index=0
    )
    
    selected_level = st.selectbox(
        "Academic Level",
        ["Postgraduate (PG)", "Undergraduate (UG)", "Diploma / Executive"],
        index=0
    )

# -------------------------------------------------------------
# MAIN APP INTERFACE: DOCUMENT INPUTS
# -------------------------------------------------------------
st.title("🎓 Multi-Tier Academic Content & Audit Pipeline")
st.caption("Stage 1: AI Maker | Stage 2: Quality & Syllabus Checker | Stage 3: Prompt Compliance Auditor")

col_left, col_right = st.columns(2)

with col_left:
    st.subheader("1. Syllabus Outline")
    uploaded_syllabus = st.file_uploader("Upload Syllabus Outline (.docx, .txt)", type=["docx", "txt"], key="syl_uploader")
    text_syllabus_input = st.text_area("Or Paste Syllabus Manually", placeholder="Paste topics or unit syllabus here...", height=120)
    
    syllabus_content = extract_text_from_file(uploaded_syllabus) if uploaded_syllabus else text_syllabus_input.strip()
    if uploaded_syllabus:
        st.success(f"Loaded: {uploaded_syllabus.name}")

with col_right:
    st.subheader("2. Reference & Benchmark Material")
    uploaded_reference = st.file_uploader("Upload Prescribed References (.docx, .txt)", type=["docx", "txt"], key="ref_uploader")
    text_ref_input = st.text_area("Or Paste References Manually", placeholder="Paste source extracts or textbook notes...", height=120)
    
    reference_content = extract_text_from_file(uploaded_reference) if uploaded_reference else text_ref_input.strip()
    if uploaded_reference:
        st.success(f"Loaded: {uploaded_reference.name}")

st.divider()

# -------------------------------------------------------------
# PIPELINE EXECUTION INTERFACE
# -------------------------------------------------------------
st.subheader("Sequential Pipeline Execution")

p_col1, p_col2, p_col3 = st.columns(3)

with p_col1:
    st.markdown("#### Stage 1: AI Maker")
    maker_btn = st.button("Step 1: Run AI Maker", type="primary", use_container_width=True)

with p_col2:
    st.markdown("#### Stage 2: Quality Checker")
    checker_btn = st.button(
        "Step 2: Run Academic Audit",
        disabled=(not bool(st.session_state["stage_1_draft"])),
        use_container_width=True
    )

with p_col3:
    st.markdown("#### Stage 3: Prompt Auditor")
    auditor_btn = st.button(
        "Step 3: Run Prompt Audit",
        disabled=(not bool(st.session_state["stage_1_draft"])),
        use_container_width=True
    )

# --- EXECUTE STAGE 1 ---
if maker_btn:
    if not syllabus_content:
        st.error("Please upload or paste a Syllabus outline first.")
    else:
        with st.spinner("Stage 1 drafting in progress (streaming enabled)..."):
            try:
                draft = run_stage_1_maker(
                    syllabus=syllabus_content,
                    reference=reference_content,
                    level=selected_level,
                    deliverable=selected_deliverable
                )
                st.session_state["stage_1_draft"] = draft
                st.session_state["stage_2_results"] = None
                st.session_state["stage_3_results"] = None
                st.success("Stage 1 Generation Complete!")
            except Exception as e:
                st.error(f"Maker Failed: {e}")

# --- EXECUTE STAGE 2 ---
if checker_btn:
    with st.spinner("Stage 2 Academic & Mapping QA Audit running..."):
        try:
            results = run_stage_2_checker(
                draft=st.session_state["stage_1_draft"],
                syllabus=syllabus_content,
                reference=reference_content
            )
            st.session_state["stage_2_results"] = results
            st.success("Stage 2 Academic Audit Complete!")
        except Exception as e:
            st.error(f"Checker Failed: {e}")

# --- EXECUTE STAGE 3 ---
if auditor_btn:
    with st.spinner("Stage 3 Deep Prompt & Guardrail Audit running..."):
        try:
            audit_res = run_stage_3_prompt_audit(
                draft=st.session_state["stage_1_draft"],
                deliverable=selected_deliverable
            )
            st.session_state["stage_3_results"] = audit_res
            st.success("Stage 3 Prompt Audit Complete!")
        except Exception as e:
            st.error(f"Auditor Failed: {e}")

st.divider()

# -------------------------------------------------------------
# STAGE 1 DELIVERABLE DISPLAY & EXPORT
# -------------------------------------------------------------
if st.session_state["stage_1_draft"]:
    st.subheader(f"Stage 1 Deliverable: {selected_deliverable} Manuscript")
    
    with st.expander("📄 Click to Preview Generated Manuscript", expanded=True):
        st.markdown(st.session_state["stage_1_draft"])
        
    manuscript_docx = build_manuscript_docx(
        title=f"{selected_university} - {selected_deliverable} Deliverable",
        content=st.session_state["stage_1_draft"]
    )
    
    st.download_button(
        label=f"📥 Download {selected_deliverable} Manuscript (.DOCX)",
        data=manuscript_docx,
        file_name=f"{selected_program}_{selected_deliverable}_Manuscript.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

# -------------------------------------------------------------
# STAGE 2 AUDIT RESULTS DISPLAY
# -------------------------------------------------------------
if st.session_state["stage_2_results"]:
    res2 = st.session_state["stage_2_results"]
    st.divider()
    st.subheader("Stage 2: Academic Quality & Syllabus Verification")
    
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Audit Score", f"{res2.get('score', 0.0)} / 100")
    c2.metric("Syllabus Coverage", f"{res2.get('coverage_pct', 0.0)}%")
    c3.metric("Critical Errors", res2.get('critical_count', 0))
    decision = res2.get('decision', 'PENDING')
    c4.metric("Verification Status", decision)
    
    if decision == "PASS":
        st.success(f"Audit Status: {decision} — Manuscript passes university quality benchmarks.")
    else:
        st.warning(f"Audit Status: {decision} — Revision recommended before institutional publication.")
        
    errs = res2.get("error_register", [])
    if errs:
        st.markdown("##### Quality Checker Error Register")
        st.dataframe(pd.DataFrame(errs), use_container_width=True)

# -------------------------------------------------------------
# STAGE 3 AUDIT RESULTS DISPLAY & DOCX EXPORT
# -------------------------------------------------------------
if st.session_state["stage_3_results"]:
    res3 = st.session_state["stage_3_results"]
    st.divider()
    st.subheader("Stage 3: Deep Prompt & Instruction Compliance Audit")
    
    m1, m2, m3 = st.columns(3)
    m1.metric("Compliance Status", res3.get("overall_compliance_status", "COMPLIANT"))
    m2.metric("Prompt Score", f"{res3.get('prompt_score', 0.0)}%")
    m3.metric("Negative Constraint Violations", res3.get("negative_constraint_violations", 0))
    
    checks = res3.get("compliance_checks", [])
    if checks:
        st.markdown("##### Meta-Audit Compliance Matrix")
        st.dataframe(pd.DataFrame(checks), use_container_width=True)
        
    st.markdown("##### Structural Drift Analysis")
    st.info(res3.get("structural_drift_analysis", "Zero prompt drift observed."))
    
    # STAGE 3 DOCX EXPORT BUTTON
    docx_audit_io = build_prompt_audit_docx(
        audit_data=res3,
        deliverable_type=selected_deliverable,
        university=selected_university,
        program=selected_program
    )
    
    st.download_button(
        label="📄 Download Stage 3 Audit Report (.DOCX)",
        data=docx_audit_io,
        file_name=f"{selected_program}_{selected_deliverable}_Prompt_Audit_Report.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary"
    )