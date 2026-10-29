import streamlit as st
from services.llm_client import run_maker_generation, run_stage_3_audits
from services.deterministic_qa import run_pre_checks
from services.report_generator import build_audit_excel
from services.docx_handler import extract_text_from_docx, generate_slm_docx

st.set_page_config(page_title="Academic Content Maker-Checker Hub", layout="wide")
st.title("University Academic Content Maker–Checker Hub")
st.caption("Automated Authoring, Bloom's Taxonomy Alignment, and Independent Quality Auditing")

if "stage" not in st.session_state:
    st.session_state.stage = "INPUT"
if "draft_text" not in st.session_state:
    st.session_state.draft_text = ""
if "audit_results" not in st.session_state:
    st.session_state.audit_results = None

# Sidebar: Institution & Workflow Configuration
st.sidebar.header("Institutional Details")
university_name = st.sidebar.text_input("Name of the University", value="State Open University")
program_name = st.sidebar.text_input("Program / Discipline", value="Master of Business Administration (MBA)")

st.sidebar.header("Workflow Settings")
deliverable_type = st.sidebar.selectbox("Deliverable Format", ["Self-Learning Material (SLM)", "Question Bank (QB)"])
academic_level = st.sidebar.selectbox("Academic Level", ["Postgraduate (PG)", "Undergraduate (UG)"])

# Consolidated metadata bundle
app_metadata = {
    "university_name": university_name,
    "program_name": program_name,
    "deliverable_type": deliverable_type,
    "academic_level": academic_level
}

# Input Section: Dual File Uploader + Fallback Text Areas
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Syllabus Input")
    uploaded_syllabus = st.file_uploader("Upload Syllabus (.docx or .txt)", type=["docx", "txt"], key="syl_file")
    syl_text_input = st.text_area("Or Paste Syllabus Manually", height=150, placeholder="Paste topics or syllabus list here...")
    
    if uploaded_syllabus:
        if uploaded_syllabus.name.endswith(".docx"):
            syllabus_data = extract_text_from_docx(uploaded_syllabus.read())
        else:
            syllabus_data = uploaded_syllabus.read().decode("utf-8")
        st.success(f"Loaded syllabus from {uploaded_syllabus.name}")
    else:
        syllabus_data = syl_text_input

with col2:
    st.subheader("2. Reference Materials")
    uploaded_ref = st.file_uploader("Upload Prescribed References (.docx or .txt)", type=["docx", "txt"], key="ref_file")
    ref_text_input = st.text_area("Or Paste References Manually", height=150, placeholder="Paste textbook extracts or authorized notes...")
    
    if uploaded_ref:
        if uploaded_ref.name.endswith(".docx"):
            reference_data = extract_text_from_docx(uploaded_ref.read())
        else:
            reference_data = uploaded_ref.read().decode("utf-8")
        st.success(f"Loaded references from {uploaded_ref.name}")
    else:
        reference_data = ref_text_input

st.divider()

# Action Area
col_btn1, col_btn2 = st.columns(2)

with col_btn1:
    if st.button("Step 1: Run AI Maker (Draft Content)", type="primary", use_container_width=True):
        if not syllabus_data.strip() or not reference_data.strip():
            st.error("Please supply both syllabus and reference sources (via upload or paste).")
        else:
            with st.spinner(f"Authoring {deliverable_type} according to university standards..."):
                try:
                    st.session_state.draft_text = run_maker_generation(syllabus_data, reference_data, academic_level)
                    st.session_state.stage = "DRAFT_READY"
                    st.rerun()
                except Exception as err:
                    st.error(f"Generation Failed: {err}")

# Maker Results Card
if st.session_state.draft_text:
    st.subheader(f"Generated {deliverable_type} Output")
    with st.expander("Preview Content in Markdown", expanded=False):
        st.markdown(st.session_state.draft_text)

    # Independent Maker Download Button with metadata header
    maker_docx_bytes = generate_slm_docx(
        title=f"{deliverable_type} Unit Manuscript",
        content_markdown=st.session_state.draft_text,
        metadata=app_metadata
    )
    st.download_button(
        label=f"Download {deliverable_type} Manuscript (.docx)",
        data=maker_docx_bytes,
        file_name=f"{university_name.replace(' ', '_')}_{deliverable_type.replace(' ', '_')}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

# Checker Audit Action
if st.session_state.stage in ["DRAFT_READY", "AUDITED"]:
    with col_btn2:
        if st.button("Step 2: Run Stage 2 & 3 Quality Audits", use_container_width=True):
            with st.spinner("Executing Deterministic Pre-checks and Parallel AI Checkers..."):
                try:
                    qa_res = run_pre_checks(st.session_state.draft_text)
                    if not qa_res["passed"]:
                        st.warning(f"Deterministic QA Flags: {qa_res['flags']}")

                    audit_res = run_stage_3_audits(st.session_state.draft_text, syllabus_data, reference_data)
                    st.session_state.audit_results = audit_res
                    st.session_state.stage = "AUDITED"
                    st.rerun()
                except Exception as err:
                    st.error(f"Audit Execution Failed: {err}")

# Checker Results & Dashboard
if st.session_state.audit_results:
    res = st.session_state.audit_results
    st.divider()
    st.subheader("Stage 3: Independent Audit Dossier")

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Final Gate Decision", res["decision"])
    m2.metric("Weighted Quality Score", f"{res['score']:.1f}%")
    m3.metric("Syllabus Coverage", f"{res['coverage_pct']:.1f}%")
    m4.metric("Prompt Adherence", res["prompt_status"])

    tab_err, tab_syl, tab_prompt = st.tabs(["Error Register", "Syllabus Matrix", "Prompt Adherence"])
    with tab_err:
        st.dataframe(res["error_register"], use_container_width=True)
    with tab_syl:
        st.dataframe(res["syllabus_mapping"], use_container_width=True)
    with tab_prompt:
        st.dataframe(res["prompt_adherence_logs"], use_container_width=True)

    # Independent Checker Download Button with metadata embedded
    excel_bytes = build_audit_excel(res, metadata=app_metadata)
    st.download_button(
        label="Download Comprehensive Checker Audit Report (.xlsx)",
        data=excel_bytes,
        file_name=f"{university_name.replace(' ', '_')}_Quality_Audit_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )