import io, json, re, requests, openpyxl
import streamlit as st
from layer3_auth import require_team_auth
from layer5_content_enhancer import (
    auto_clean_mcq_workbook,
    build_surgical_fixit_prompt,
    build_surgical_fixit_docx,
    parse_layer4_docx_report,
    extract_manuscript_text,
)

st.session_state["_app_layer_mode"] = "l5"
st.set_page_config(page_title="Layer-5 Content Enhancement & Remediation Studio", page_icon="Studio", layout="wide", initial_sidebar_state="expanded")

if not require_team_auth():
    st.stop()

active_user = st.session_state.get("layer3_auth_user", "Layer-5 User")
st.markdown("# Layer-5 Content Enhancement & Remediation Studio")
st.caption(f"Signed in as **{active_user}** | Report-Gated Remediation (Requires Layer-4 QA Baseline Report .docx for R1 vs R2 Verification)")

tab_mcq, tab_slm = st.tabs([
    "1. MCQ / QB Bank Remediator (.xlsx + Layer-4 Report)",
    "2. SLM & PPTX Surgical Fix-It Studio (.docx/.pdf/.pptx + Layer-4 Report)",
])

with tab_mcq:
    st.subheader("Report-Gated MCQ / QB Excel Remediator")
    st.info("Mandatory Gate: Upload both the Original MCQ/QB Workbook (.xlsx/.xlsm) AND the Layer-4 QA Audit Report (.docx). Every correction is logged in Audit_Change_Log with the Layer-4 Report v1 baseline for Report v2 re-audit verification.")
    c1, c2 = st.columns(2)
    mcq_file = c1.file_uploader("Step 1 (Required): Upload Original MCQ / QB Workbook (.xlsx, .xlsm)", type=["xlsx", "xlsm"], key="l5_mcq_man")
    mcq_rep = c2.file_uploader("Step 2 (Required): Upload Layer-4 QA Audit Report (.docx)", type=["docx"], key="l5_mcq_rep")
    if mcq_rep:
        r1_data = parse_layer4_docx_report(mcq_rep.getvalue())
        st.success(f"Layer-4 Report v1 verified: {r1_data['total_findings']} findings extracted | Score: {r1_data['overall_score']}")
    elif mcq_file:
        st.warning("Remediation locked: Upload the Layer-4 QA Audit Report (.docx) to unlock MCQ fixes.")
    can_run_mcq = bool(mcq_file and mcq_rep)
    if st.button("Run Layer-4 Report-Anchored MCQ Remediation (Cleaned .xlsx)", type="primary", use_container_width=True, disabled=not can_run_mcq):
        with st.spinner("Remediating MCQ workbook against Layer-4 Report v1 baseline..."):
            r1_data = parse_layer4_docx_report(mcq_rep.getvalue())
            cleaned_bytes, stats = auto_clean_mcq_workbook(mcq_file.getvalue(), audit_res=r1_data)
            wb = openpyxl.load_workbook(io.BytesIO(cleaned_bytes))
            ws = wb.create_sheet(title="L4_Report_v1_Baseline")
            ws.append(["Layer-4 Report File", mcq_rep.name])
            ws.append(["Baseline Score (R1)", str(r1_data.get("overall_score", ""))])
            ws.append(["Baseline Verdict (R1)", str(r1_data.get("verdict", ""))])
            ws.append(["Total L4_Report_v1 Findings", r1_data.get("total_findings", 0)])
            ws.append(["Headers Fixed", stats.get("headers_fixed", 0)])
            ws.append(["Keys Normalized", stats.get("keys_normalized", 0)])
            ws.append(["Metadata Realigned", stats.get("metadata_realigned", 0)])
            ws.append(["Duplicates Removed", stats.get("duplicates_removed", 0)])
            ws.append([])
            ws.append(["Layer-4 Report v1 Extracted Findings (For Report v2 Re-Audit Comparison)"])
            for f in r1_data.get("technical_errors", []) + r1_data.get("clarity_issues", []):
                ws.append([f])
            out_b = io.BytesIO()
            wb.save(out_b)
            cleaned_bytes = out_b.getvalue()
        st.success("Cleaned workbook ready with Audit_Change_Log & L4_Report_v1_Baseline tabs! Re-upload this file into Layer-4 to verify Report v2.")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Headers Fixed", stats.get("headers_fixed", 0))
        m2.metric("Keys Normalized", stats.get("keys_normalized", 0))
        m3.metric("Metadata Realigned", stats.get("metadata_realigned", 0))
        m4.metric("Duplicates Removed", stats.get("duplicates_removed", 0))
        st.download_button("Download Cleaned_" + mcq_file.name, data=cleaned_bytes, file_name="Cleaned_" + mcq_file.name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", use_container_width=True)

with tab_slm:
    st.subheader("SLM & PPTX Surgical Fix-It Studio (Layer-4 Report v1 Gated)")
    st.info("Mandatory Gate: Upload both the Original Manuscript (.docx, .pdf, .pptx, .txt) AND the Layer-4 QA Audit Report (.docx). Layer-5 generates a surgical remediation addendum targeting every Layer-4 Report v1 finding so you can verify corrections in Report v2.")
    col_a, col_b = st.columns(2)
    slm_title = col_a.text_input("Chapter / Deck Title", value="Unit 1: Core Concepts")
    slm_subj = col_b.text_input("Course / Subject Name", value="Data Science & Analytics")
    u1, u2 = st.columns(2)
    slm_file = u1.file_uploader("Step 1 (Required): Upload Original Manuscript (.docx, .pdf, .pptx, .txt)", type=["docx", "pdf", "pptx", "txt"], key="l5_slm_man")
    slm_rep = u2.file_uploader("Step 2 (Required): Upload Layer-4 QA Audit Report (.docx)", type=["docx"], key="l5_slm_rep")
    if slm_rep:
        r1_slm = parse_layer4_docx_report(slm_rep.getvalue())
        st.success(f"Layer-4 Report v1 verified: {r1_slm['total_findings']} findings extracted | Score: {r1_slm['overall_score']}")
        with st.expander("Preview Extracted Layer-4 Report v1 Findings", expanded=False):
            for f in r1_slm.get("technical_errors", []) + r1_slm.get("clarity_issues", []):
                st.write("* " + f)
    elif slm_file:
        st.warning("Remediation locked: Upload the Layer-4 QA Audit Report (.docx) to unlock Surgical Fix-It.")
    can_run_slm = bool(slm_file and slm_rep)
    if st.button("Generate Layer-4 Report-Anchored Surgical Fix-It Addendum (.docx)", type="primary", use_container_width=True, disabled=not can_run_slm):
        ch_text = extract_manuscript_text(slm_file.getvalue(), slm_file.name)
        r1_slm = parse_layer4_docx_report(slm_rep.getvalue())
        api_key = str(st.secrets.get("GEMINI_API_KEY", "")).strip()
        if not api_key:
            st.error("GEMINI_API_KEY is missing in Streamlit Secrets.")
        else:
            with st.spinner("Synthesizing Surgical Fix-It Addendum anchored to Layer-4 Report v1..."):
                prompt = build_surgical_fixit_prompt(slm_title, slm_subj, ch_text[:45000], r1_slm)
                resp = requests.post(f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={api_key}", json={"contents": [{"parts": [{"text": prompt}]}], "generationConfig": {"responseMimeType": "application/json"}}, timeout=120)
                resp.raise_for_status()
                res_txt = resp.json()["candidates"][0]["content"]["parts"][0]["text"]
                payload = json.loads(re.sub(r"^```json|```$", "", res_txt.strip()).strip())
                fix_docx = build_surgical_fixit_docx(slm_title, slm_subj, payload)
            st.success("Layer-5 Surgical Fix-It Addendum generated from Layer-4 Report v1!")
            safe_name = re.sub(r"[^A-Za-z0-9_-]+", "_", slm_title).strip("_") or "Chapter"
            st.download_button(f"Download Layer5_FixIt_{safe_name}.docx", data=fix_docx, file_name=f"Layer5_FixIt_{safe_name}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="primary", use_container_width=True)
