import io, json, re, runpy, requests, pypdf, docx
import streamlit as st
from layer3_auth import require_team_auth
from layer5_content_enhancer import auto_clean_mcq_workbook, build_surgical_fixit_prompt, build_surgical_fixit_docx

st.session_state["_app_layer_mode"] = "l5"
st.set_page_config(page_title="Layer-5 Content Enhancement and Unified Studio", page_icon="Studio", layout="wide", initial_sidebar_state="expanded")

if not require_team_auth():
    st.stop()

active_user = st.session_state.get("layer3_auth_user", "Layer-5 User")
l4_url = str(st.secrets.get("LAYER4_APP_URL", "https://slm-maker-checker-ad99-layer4.streamlit.app")).rstrip("/")
l5_url = str(st.secrets.get("LAYER5_APP_URL", "https://slm-maker-checker-ad99-layer5.streamlit.app")).rstrip("/")

st.sidebar.markdown("---")
st.sidebar.markdown("### Unified Platform Navigation")
st.sidebar.link_button("Open Layer-4 Document Auditor", url=l4_url, use_container_width=True)
st.sidebar.link_button("Layer-5 Enhancement Studio", url=l5_url, use_container_width=True)

st.markdown("# Layer-5 Content Enhancement and Remediation Studio")
st.caption(f"Signed in as **{active_user}** | Automated MCQ Bank Auto-Cleaner, SLM Surgical Fix-It Generator and Unified Cross-Layer Portal")

tab_mcq, tab_slm, tab_portal = st.tabs(["1. MCQ Bank Auto-Cleaner (.xlsx)", "2. SLM Surgical Fix-It Studio (.docx / .pdf)", "3. Unified Layer-4 and Layer-5 Portal"])

with tab_mcq:
    st.subheader("Automated QB / MCQ Excel Workbook Remediator")
    st.info("Fixes corrupted headers (3orre3t Option to Correct Option), normalizes 1-4 keys to A-D, realigns shifted BTL/Difficulty cells, strips repeated banners, deduplicates stems, and generates an Audit_Change_Log tab.")
    mcq_file = st.file_uploader("Upload MCQ / Question Bank Workbook (.xlsx, .xlsm)", type=["xlsx", "xlsm"], key="l5_mcq_up")
    if mcq_file and st.button("Auto-Clean and Remediate MCQ Workbook (.xlsx)", type="primary", use_container_width=True):
        with st.spinner("Running Layer-5 deterministic MCQ remediation..."):
            cleaned_bytes, stats = auto_clean_mcq_workbook(mcq_file.getvalue())
        st.success("Remediation complete! Download your cleaned workbook below.")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Headers Fixed", stats.get("headers_fixed", 0))
        c2.metric("Keys Normalized", stats.get("keys_normalized", 0))
        c3.metric("Metadata Realigned", stats.get("metadata_realigned", 0))
        c4.metric("Duplicates Removed", stats.get("duplicates_removed", 0))
        st.download_button("Download Cleaned_" + mcq_file.name, data=cleaned_bytes, file_name="Cleaned_" + mcq_file.name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", use_container_width=True)

with tab_slm:
    st.subheader("SLM Chapter Surgical Fix-It Addendum Generator")
    col_a, col_b = st.columns(2)
    slm_title = col_a.text_input("Chapter / Unit Title", value="Unit 1: Core Concepts")
    slm_subi = col_b.text_input("Course / Subject Name", value="Data Science and Analytics")
    slm_file = st.file_uploader("Upload SLm Manuscript (.docx, .pdf, .txt)", type=["docx", "pdf", "txt"], key="l5_slm_up")
    qa_notes = st.text_area("Paste Layer-4 Audit Findings / Remediation Instructions (optional)", height=130, placeholder="Paste key gaps, missing Bloom verbs, case study gaps, or formatting issues from Layer-4 QA report...")
    if slm_file and st.button('Generate Layer-5 Surgical Fix-It Addendum (.docx)', type="primary", use_container_width=True):
        raw = slm_file.getvalue()
        fns = slm_file.name.lower()
        if fns.endswith(".docx"):
            d = docx.Document(io.BytesIO(raw))
            ch_text = "\n".join(p.text for p in d.paragraphs if p.text.strip())
        elif fns.endswith(".pdf"):
            reader = pypdf.PdfReader(io.BytesIO(raw))
            ch_text = "\n".join((pg.extract_text() or "") for pg in reader.pages)
        else:
            ch_text = raw.decode("utf-8", errors="ignore")
        api_key = str(st.secrets.get("GEMINI_API_KEY", "")).strip()
        if not api_key:
            st.error("GEMINI_API_KEY is missing in Streamlit Secrets.")
        else:
            with st.spinner("Synthesizing 4-table Surgical Fix-It Addendum via Gemini..."):
                audit_stub = {"summary": qa_notes or "Improve Bloom cognitive depth, real-world corporate examples, Quadrant-II scaffolding, and self-assessment rigor.", "findings": [qa_notes] if qa_notes else []}
                prompt = build_surgical_fixit_prompt(slm_title, slm_subj, ch_text[:45000], audit_stub)
                resp = requests.post(f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={api_key}", json={"contents": [{"parts": [{"text": prompt}]}], "generationConfig": {"responseMimeType": "application/json"}}, timeout=120)
                resp.raise_for_status()
                txt = resp.json()["candidates"][0]["content"]["parts"][0]["text"]
                payload = json.loads(re.sub(rb"^```json|```$", "", txt.strip()).strip())
                fix_docx = build_surgical_fixit_docx(slm_title, slm_subj, payload)
            st.success("Layer-5 Surgical Fix-It Addendum generated!")
            safe_name = re.sub(r"[^A-Za-z0-9_-]+", "_", slm_title).strip("_") or "Chapter"
            st.download_button(f"Download Layer5_FixIt_{safe_name}.docx", data=fix_docx, file_name=f"Layer5_FixIt_{safe_name}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="primary", use_container_width=True)

with tab_portal:
    st.subheader("Unified Cross-Layer Academic QA and Enhancement Portal")
    p1, p2 = st.columns(2)
    with p1:
        st.markdown("### Layer-4 Document Auditor")
        st.write("12-Parameter Forensic QA, Google Sheets Master Tracker Sync and Word Audit Reports.")
        st.link_button("Launch Layer-4 Auditor", url=l4_url, type="primary", use_container_width=True)
    with p2:
        st.markdown("### Layer-5 Content Enhancement Studio")
        st.write("Automated QB/MCQ Excel Auto-Cleaner and SLm Chapter Surgical Fix-It Remediator.")
        st.link_button("Refresh Layer-5 Studio", url=l5_url, use_container_width=True)
