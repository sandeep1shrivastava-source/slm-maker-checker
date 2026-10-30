import streamlit as st
from services.access_manager import load_access_config, add_university, assign_user_role, verify_user
from services.unified_engine import run_stage_1_maker, run_stage_2_checker, run_stage_3_prompt_audit
from services.deterministic_qa import run_pre_checks
from services.unified_report_generator import build_master_audit_excel
from services.docx_handler import extract_text_from_docx, generate_slm_docx

st.set_page_config(page_title="Unified Academic Pipeline (Maker–Checker–Prompt)", layout="wide")
st.title("University Unified 3-Tier Academic Content & Audit Pipeline")
st.caption("Layered Execution: AI Maker (1) ➔ Academic QA Checker (2) ➔ Deep Prompt Auditor (3)")

# Load dynamic access control & university registry
access_cfg = load_access_config()

# Session State Setup
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

# Verification check
is_authorized, user_role = verify_user(user_email)

if not is_authorized:
    st.sidebar.error("⛔ Access Denied")
    st.error(
        f"**Access Restricted:** The email `{user_email}` is not registered in the institutional access roster. "
        "Please contact the University Administrator to grant access."
    )
    st.stop()  # Halts all pipeline execution for unauthorized users

st.sidebar.success(f"**Verified Role:** {user_role}")

# Dynamic Permission Matrix
is_admin = "Administrator" in user_role
can_run_maker = "Maker" in user_role or is_admin
can_run_checker = "Checker" in user_role or is_admin
can_run_prompt = "Prompt Engineer" in user_role or is_admin

# ==========================================
# SIDEBAR: INSTITUTIONAL SELECTION & FORMATS
# ==========================================
st.sidebar.divider()
st.sidebar.header("🏛 Institutional Scope")

# University Dropdown (Populated dynamically from config)
university_list = access_cfg.get("universities", ["State Open University"])
selected_university = st.sidebar.selectbox("Select University", options=university_list)

program_name = st.sidebar.text_input("Program / Discipline", value="Master of Business Administration (MBA)")

# Exact 5 Review Formats
FORMAT_OPTIONS = [
    "MCQ",
    "SLM",
    "Question Bank",
    "PPT",
    "FAQs & Misconceptions Review"
]
deliverable_type = st.sidebar.selectbox("Deliverable Format", options=FORMAT_OPTIONS)
academic_level = st.sidebar.selectbox("Academic Level", ["Postgraduate (PG)", "Undergraduate (UG)"])

app_metadata = {
    "university_name": selected_university,
    "program_name": program_name,
    "deliverable_type": deliverable_type,
    "academic_level": academic_level,
    "operator_email": user_email,
    "operator_role": user_role
}

# ==========================================
# ADMIN MANAGEMENT CONSOLE (ADMIN ONLY)
# ==========================================
if is_admin:
    with st.expander("⚙️ Admin Management Console (User Whitelist & University Registry)", expanded=False):
        adm_col1, adm_col2 = st.columns(2)

        with adm_col1:
            st.markdown("##### 🏛 Add University to Dropdown")
            new_uni_name = st.text_input("New University Name", placeholder="e.g., Central Technical University")
            if st.button("Add University"):
                if new_uni_name.strip():
                    if add_university(new_uni_name):
                        st.success(f"Added '{new_uni_name}' to dropdown list!")
                        st.rerun()
                    else:
                        st.warning("University already exists in the dropdown list.")
                else:
                    st.error("Please enter a valid university name.")

        with adm_col2:
            st.markdown("##### 👥 Assign User Access & Role")
            new_user_email = st.text_input("User Email ID", placeholder="faculty@university.edu").strip().lower()
            role_options = ["Maker (Author)", "Checker (Auditor)", "Prompt Engineer", "Lead Administrator"]
            selected_role_assignment = st.selectbox("Assign Workflow Role", options=role_options)
            if st.button("Whitelist / Update User"):
                if new_user_email and "@" in new_user_email:
                    assign_user_role(new_user_email, selected_role_assignment)
                    st.success(f"Assigned '{new_user_email}' as '{selected_role_assignment}'!")
                    st.rerun()
                else:
                    st.error("Please provide a valid email address.")

# ==========================================
# INPUT SECTION (DOCX & TEXT FALLBACK)
# ==========================================
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Approved Syllabus")
    up_syl = st.file_uploader("Upload Syllabus (.docx or .txt)", type=["docx", "txt"], key="u_syl")
    syl_text = st.text_area("Or Paste Syllabus Manually", height=140, placeholder="Paste topics or unit syllabus here...")
    if up_syl:
        syllabus_data = extract_text_from_docx(up_syl.read()) if up_syl.name.endswith(".docx") else up_syl.read().decode("utf-8")
        st.success(f"Loaded: {up_syl.name}")
    else:
        syllabus_data = syl_text

with col2:
    st.subheader("2. Prescribed References")
    up_ref = st.file_uploader("Upload References (.docx or .txt)", type=["docx", "txt"], key="u_ref")
    ref_text = st.text_area("Or Paste References Manually", height=140, placeholder="Paste source extracts or textbook notes...")
    if up_ref:
        reference_data = extract_text_from_docx(up_ref.read()) if up_ref.name.endswith(".docx") else up_ref.read().decode("utf-8")
        st.success(f"Loaded: {up_ref.name}")
    else:
        reference_data = ref_text

st.divider()

# ==========================================
# 3-STAGE SEQUENTIAL PIPELINE
# ==========================================
st.subheader("Sequential Pipeline Execution")
c_s1, c_s2, c_s3 = st.columns(3)

# --- STAGE 1: MAKER ---
with c_s1:
    st.markdown("#### Stage 1: AI Maker")
    maker_btn = st.button("Step 1: Run AI Maker", type="primary", disabled=not can_run_maker, use_container_width=True)
    if not can_run_maker:
        st.caption("🔒 Role Restricted: Requires Maker or Admin")

    if maker_btn:
        if not syllabus_data.strip() or not reference_data.strip():
            st.error("Please supply both syllabus and reference data.")
        else:
            with st.spinner(f"Drafting {deliverable_type}..."):
                try:
                    st.session_state.draft_text = run_stage_1_maker(syllabus_data, reference_data, academic_level, deliverable_type)
                    st.session_state.stage = "STAGE_1_COMPLETE"
                    st.session_state.stage2_results = None
                    st.session_state.stage3_results = None
                    st.rerun()
                except Exception as e:
                    st.error(f"Maker Failed: {e}")

# --- STAGE 2: ACADEMIC CHECKER ---
with c_s2:
    st.markdown("#### Stage 2: Quality Checker")
    s2_enabled = (st.session_state.draft_text != "") and can_run_checker
    checker_btn = st.button("Step 2: Run Academic Audit", disabled=not s2_enabled, use_container_width=True)
    if not can_run_checker:
        st.caption("🔒 Role Restricted: Requires Checker or Admin")

    if checker_btn:
        with st.spinner("Auditing 12 QA parameters & syllabus coverage..."):
            try:
                pre = run_pre_checks(st.session_state.draft_text)
                if not pre["passed"]:
                    st.warning(f"Deterministic QA Flags: {pre['flags']}")

                res2 = run_stage_2_checker(st.session_state.draft_text, syllabus_data, reference_data)
                st.session_state.stage2_results = res2
                st.session_state.stage = "STAGE_2_COMPLETE"
                st.rerun()
            except Exception as e:
                st.error(f"Stage 2 Failed: {e}")

# --- STAGE 3: PROMPT AUDITOR ---
with c_s3:
    st.markdown("#### Stage 3: Prompt Auditor")
    s3_enabled = (st.session_state.stage2_results is not None) and can_run_prompt
    prompt_btn = st.button("Step 3: Run Prompt Audit", disabled=not s3_enabled, use_container_width=True)
    if not can_run_prompt:
        st.caption("🔒 Role Restricted: Requires Prompt Auditor or Admin")

    if prompt_btn:
        with st.spinner("Analyzing prompt adherence and negative constraints..."):
            try:
                res3 = run_stage_3_prompt_audit(st.session_state.draft_text, deliverable_type)
                st.session_state.stage3_results = res3
                st.session_state.stage = "STAGE_3_COMPLETE"
                st.rerun()
            except Exception as e:
                st.error(f"Stage 3 Failed: {e}")

# ==========================================
# DELIVERABLE OUTPUTS & REPORTING
# ==========================================
if st.session_state.draft_text:
    st.divider()
    st.subheader(f"Stage 1 Deliverable: {deliverable_type} Manuscript")
    with st.expander("Expand to Preview Manuscript", expanded=False):
        st.markdown(st.session_state.draft_text)

    maker_docx = generate_slm_docx(
        title=f"{deliverable_type} Unit Manuscript",
        content_markdown=st.session_state.draft_text,
        metadata=app_metadata
    )
    st.download_button(
        label=f"Download {deliverable_type} Manuscript (.docx)",
        data=maker_docx,
        file_name=f"{selected_university.replace(' ', '_')}_{deliverable_type.replace(' ', '_')}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

if st.session_state.stage2_results:
    st.divider()
    st.subheader("Audit & Compliance Evaluation")

    r2 = st.session_state.stage2_results
    r3 = st.session_state.stage3_results

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Academic Decision", r2["decision"])
    m2.metric("Academic Score", f"{r2['score']:.1f}%")
    m3.metric("Syllabus Coverage", f"{r2['coverage_pct']:.1f}%")

    if r3:
        m4.metric("Prompt Compliance", f"{r3.get('prompt_score', 0):.1f}%", r3.get("overall_compliance_status"))
    else:
        m4.metric("Prompt Compliance", "Step 3 Pending")

    tab_err, tab_syl, tab_p_audit = st.tabs(["Academic Error Register", "Syllabus Matrix", "Deep Prompt Adherence"])
    with tab_err:
        st.dataframe(r2["error_register"], use_container_width=True)
    with tab_syl:
        st.dataframe(r2["syllabus_mapping"], use_container_width=True)
    with tab_p_audit:
        if r3:
            st.dataframe(r3.get("compliance_checks", []), use_container_width=True)
            st.info(f"**Structural Drift Assessment:** {r3.get('structural_drift_analysis')}")
        else:
            st.info("Execute Step 3 to review negative constraints, boilerplate flags, and Bloom's verb adherence.")

    master_excel = build_master_audit_excel(
        academic_results=r2,
        prompt_results=r3,
        metadata=app_metadata
    )
    st.download_button(
        label="Download Master Consolidated Audit Dossier (.xlsx)",
        data=master_excel,
        file_name=f"{selected_university.replace(' ', '_')}_Master_Audit_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )