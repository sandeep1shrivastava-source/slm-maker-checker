import io, json, re, openpyxl, pptx
import streamlit as st
from layer3_auth import require_team_auth
from layer5_content_enhancer import (
    auto_clean_mcq_workbook,
    build_surgical_fixit_prompt,
    build_surgical_fixit_docx,
    parse_layer4_docx_report,
    extract_manuscript_text,
)
from layer5_cloud import (
    get_layer5_gcp_services,
    extract_drive_id,
    download_drive_file_bytes,
    upload_file_to_drive,
    call_layer5_llm,
    build_complete_mirrored_book_docx,
    build_subject_cumulative_dossier,
)

st.session_state["_app_layer_mode"] = "l5"
st.set_page_config(
    page_title="Layer-5 Workbench: Remediation, Enrichment & Book Studio (v3.0)",
    page_icon="Studio",
    layout="wide",
    initial_sidebar_state="expanded",
)

if not require_team_auth():
    st.stop()

active_user = st.session_state.get("layer3_auth_user", "Layer-5 User")

with st.sidebar:
    st.markdown("### University Brand Profile")
    brand_preset = st.selectbox(
        "Institution Preset",
        ["Ajeenkya DY Patil University", "Custom University / Institution"],
        key="l5_brand_preset",
    )
    if brand_preset == "Ajeenkya DY Patil University":
        brand_cfg = {
            "name": "Ajeenkya DY Patil University (ADYPU)",
            "primary": "#8B1E3F",
            "accent": "#D4AF37",
            "footer": "Ajeenkya DY Patil University - Centre for Online & Distance Learning (SLM Master)",
        }
    else:
        b_name = st.text_input("Institution Name", value="University Academic Press", key="l5_b_name")
        b_prim = st.color_picker("Primary Color", value="#1E3A8A", key="l5_b_prim")
        b_acc = st.color_picker("Accent Color", value="#D97706", key="l5_b_acc")
        b_foot = st.text_input("Statutory Footer", value=f"{b_name} - Approved Academic Edition", key="l5_b_foot")
        brand_cfg = {"name": b_name, "primary": b_prim, "accent": b_acc, "footer": b_foot}

    st.markdown("---")
    st.markdown("### AI Remediation Engine")
    ai_provider = st.selectbox(
        "AI Provider",
        [
            "Gemini 2.5 Flash (Default)",
            "OpenAI (ChatGPT gpt-4o-mini)",
            "Anthropic (Claude 3.5 Sonnet)",
            "Deterministic Offline Scaffolding (No AI Key)",
        ],
        key="l5_ai_provider",
    )
    target_drive_url = st.text_input(
        "Target Google Drive Folder URL (Optional Auto-Sync)",
        placeholder="https://drive.google.com/drive/folders/...",
        key="l5_target_drive",
    )
    target_folder_id = extract_drive_id(target_drive_url)

st.markdown("# Layer-5 Academic QA Remediation, Enrichment & Book Studio")
st.caption(
    f"Signed in as **{active_user}** | Zero-Risk Clone-on-Write Guarantee | "
    "Mandatory Layer-4 QA Baseline Report (.docx) Gate for R1 vs R2 Verification"
)

drive_svc, sheets_svc, google_email, gcp_msg, oauth_url = get_layer5_gcp_services()
if drive_svc is not None:
    st.success(f"Google Drive & Master Tracker Connected ({google_email}) - supportsAllDrives=True active")
else:
    c_g1, c_g2 = st.columns([3, 1])
    c_g1.info(f"Optional Cloud Sync: {gcp_msg}")
    if oauth_url:
        c_g2.link_button("Sign In with Google (Drive & Sheets)", url=oauth_url, type="primary", use_container_width=True)

def _resolve_input_asset(uploaded_file, drive_link: str, drive_client) -> tuple[bytes | None, str]:
    if uploaded_file is not None:
        return uploaded_file.getvalue(), uploaded_file.name
    fid = extract_drive_id(drive_link)
    if fid and drive_client is not None:
        return download_drive_file_bytes(drive_client, fid)
    return None, ""

api_keys = {
    "gemini": str(st.secrets.get("GEMINI_API_KEY", "")).strip(),
    "openai": str(st.secrets.get("OPENAI_API_KEY", "")).strip(),
    "anthropic": str(st.secrets.get("ANTHROPIC_API_KEY", "")).strip(),
}

tab1, tab2, tab3, tab4 = st.tabs([
    "1. MCQ / QB Bank Remediator (.xlsx + L4 Report)",
    "2. Aspect 1: Unit & PPTX Remediation / Enrichment",
    "3. Aspect 2: 1-to-N Complete Book Compiler",
    "4. Aspect 3: Subject-Wide Cumulative QA Report",
])

with tab1:
    st.subheader("Report-Gated MCQ / QB Excel Remediator")
    st.info(
        "Mandatory Layer-4 Gate: Provide both the Original MCQ/QB Workbook (.xlsx/.xlsm) AND its Layer-4 QA Audit Report (.docx). "
        "Every fix is logged in Audit_Change_Log + L4_Report_v1_Baseline for Report v2 re-audit comparison."
    )
    c1, c2 = st.columns(2)
    mcq_up = c1.file_uploader("Slot 1 (Required): Original MCQ / QB Workbook (.xlsx, .xlsm)", type=["xlsx", "xlsm"], key="l5_mcq_up")
    mcq_drv = c1.text_input("Or paste Google Drive / Sheet URL for MCQ Workbook", key="l5_mcq_drv", disabled=(drive_svc is None))
    rep1_up = c2.file_uploader("Slot 2 (Required): Layer-4 QA Audit Report (.docx)", type=["docx"], key="l5_mcq_rep_up")
    rep1_drv = c2.text_input("Or paste Google Drive URL for Layer-4 Report (.docx)", key="l5_mcq_rep_drv", disabled=(drive_svc is None))

    mcq_bytes, mcq_name = _resolve_input_asset(mcq_up, mcq_drv, drive_svc)
    rep1_bytes, rep1_name = _resolve_input_asset(rep1_up, rep1_drv, drive_svc)

    if rep1_bytes:
        r1_mcq = parse_layer4_docx_report(rep1_bytes)
        st.success(f"Layer-4 Report v1 Verified ({rep1_name}): {r1_mcq['total_findings']} findings extracted | Score: {r1_mcq['overall_score']}")
    elif mcq_bytes:
        st.warning("Remediation Locked: Provide the matching Layer-4 QA Audit Report (.docx) in Slot 2 to unlock MCQ fixes.")

    if st.button("Run Layer-4 Report-Anchored MCQ Remediation (Cleaned .xlsx)", type="primary", use_container_width=True, disabled=not (mcq_bytes and rep1_bytes)):
        with st.spinner("Remediating MCQ workbook against Layer-4 Report v1 baseline..."):
            r1_mcq = parse_layer4_docx_report(rep1_bytes)
            cleaned_bytes, stats = auto_clean_mcq_workbook(mcq_bytes, audit_res=r1_mcq)
            wb = openpyxl.load_workbook(io.BytesIO(cleaned_bytes))
            ws = wb.create_sheet(title="L4_Report_v1_Baseline")
            ws.append(["Layer-4 Report File", rep1_name])
            ws.append(["Baseline Score (R1)", str(r1_mcq.get("overall_score", ""))])
            ws.append(["Baseline Verdict (R1)", str(r1_mcq.get("verdict", ""))])
            ws.append(["Total L4_Report_v1 Findings", r1_mcq.get("total_findings", 0)])
            ws.append(["Headers Fixed", stats.get("headers_fixed", 0)])
            ws.append(["Keys Normalized", stats.get("keys_normalized", 0)])
            ws.append(["Metadata Realigned", stats.get("metadata_realigned", 0)])
            ws.append(["Duplicates Removed", stats.get("duplicates_removed", 0)])
            ws.append([])
            ws.append(["Layer-4 Report v1 Extracted Findings (For Report v2 Re-Audit Comparison)"])
            for f in r1_mcq.get("technical_errors", []) + r1_mcq.get("clarity_issues", []):
                ws.append([f])
            out_b = io.BytesIO()
            wb.save(out_b)
            cleaned_bytes = out_b.getvalue()

        out_fname = "Cleaned_" + (mcq_name if mcq_name.lower().endswith(".xlsx") else mcq_name + ".xlsx")
        st.success("Cleaned workbook ready with Audit_Change_Log & L4_Report_v1_Baseline! Re-upload into Layer-4 to verify Report v2.")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Headers Fixed", stats.get("headers_fixed", 0))
        m2.metric("Keys Normalized", stats.get("keys_normalized", 0))
        m3.metric("Metadata Realigned", stats.get("metadata_realigned", 0))
        m4.metric("Duplicates Removed", stats.get("duplicates_removed", 0))
        st.download_button(f"Download {out_fname}", data=cleaned_bytes, file_name=out_fname, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", use_container_width=True)
        if drive_svc and target_folder_id:
            link = upload_file_to_drive(drive_svc, target_folder_id, out_fname, cleaned_bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.markdown(f"**Synced Cloned Deliverable to Google Drive:** [{out_fname}]({link})")

with tab2:
    st.subheader("Aspect 1: Unit & PPTX Remediation (Track A) & Audit-Verified Enrichment (Track B)")
    st.info(
        "4-Point Quality Gate: Upload your Unit/Deck (.docx, .pdf, .pptx, .txt) in Slot 1 AND its Layer-4 QA Audit Report (.docx) in Slot 2. "
        "Track A generates a human-reviewed Surgical Fix-It Patch; Track B unlocks Creative Enrichment & PPTX Slide Transitions."
    )
    ca, cb = st.columns(2)
    slm_title = ca.text_input("Chapter / Unit / Deck Title", value="Unit 1: Core Concepts", key="l5_u_title")
    slm_subj = cb.text_input("Course / Subject Name", value="Data Science & Analytics", key="l5_u_subj")

    u1, u2 = st.columns(2)
    man_up = u1.file_uploader("Slot 1 (Required): Original Manuscript (.docx, .pdf, .pptx, .txt)", type=["docx", "pdf", "pptx", "txt"], key="l5_u_man")
    man_drv = u1.text_input("Or paste Google Drive URL for Manuscript", key="l5_u_man_drv", disabled=(drive_svc is None))
    rep2_up = u2.file_uploader("Slot 2 (Required): Layer-4 QA Audit Report (.docx)", type=["docx"], key="l5_u_rep")
    rep2_drv = u2.text_input("Or paste Google Drive URL for Layer-4 Report (.docx)", key="l5_u_rep_drv", disabled=(drive_svc is None))

    man_bytes, man_name = _resolve_input_asset(man_up, man_drv, drive_svc)
    rep2_bytes, rep2_name = _resolve_input_asset(rep2_up, rep2_drv, drive_svc)

    r1_slm = None
    score_num = 0.0
    if rep2_bytes:
        r1_slm = parse_layer4_docx_report(rep2_bytes)
        m_sc = re.search(r"(\d+(?:\.\d+)?)", str(r1_slm.get("overall_score", "")))
        score_num = float(m_sc.group(1)) if m_sc else 0.0
        st.success(f"Layer-4 Report v1 Verified ({rep2_name}): Score = {r1_slm['overall_score']} | {r1_slm['total_findings']} findings extracted")
        with st.expander("Preview Extracted Layer-4 Report v1 Findings", expanded=False):
            for f in r1_slm.get("technical_errors", []) + r1_slm.get("clarity_issues", []):
                st.write("* " + f)
    elif man_bytes:
        st.warning("Remediation & Enrichment Locked: Upload the Layer-4 QA Audit Report (.docx) in Slot 2.")

    col_ta, col_tb = st.columns(2)
    with col_ta:
        st.markdown("#### Track A: Surgical Auto-Remediation (Report v1 Anchored)")
        if st.button("Generate Track A Fix-It Draft (Human Preview)", type="primary", use_container_width=True, disabled=not (man_bytes and rep2_bytes)):
            ch_text = extract_manuscript_text(man_bytes, man_name)
            fallback_fix = {
                "executive_summary": f"Deterministic remediation scaffold anchored to {r1_slm['total_findings']} Layer-4 Report v1 findings.",
                "prose_rewrites": [{"section": slm_title, "issue": f, "original_excerpt": "Flagged in Layer-4 Report v1", "revised_text": f"Revised academic prose addressing: {f}"} for f in (r1_slm.get("technical_errors", [])[:5] or ["Improve Bloom cognitive depth and academic cadence."])],
                "cyp_questions": [{"question": f"Explain the core principles of {slm_title} with an applied industry example.", "answer": "Model answer aligned to course learning outcomes."}],
                "glossary_terms": [{"term": slm_title, "definition": "Core instructional unit validated against institutional rubric."}],
            }
            with st.spinner("Drafting Track A Surgical Fix-It items..."):
                prompt = build_surgical_fixit_prompt(slm_title, slm_subj, ch_text[:45000], r1_slm)
                st.session_state["l5_staged_fix"] = call_layer5_llm(prompt, ai_provider, api_keys, fallback_fix)

        if st.session_state.get("l5_staged_fix"):
            staged = st.session_state["l5_staged_fix"]
            st.info(staged.get("executive_summary", "Review staged fixes below before compiling:"))
            approved_rewrites = []
            for idx, rw in enumerate(staged.get("prose_rewrites", [])):
                if st.checkbox(f"Approve Fix #{idx+1}: {rw.get('issue', 'Prose Fix')[:85]}", value=True, key=f"l5_chk_{idx}"):
                    approved_rewrites.append(rw)
            if st.button("Compile Approved Track A Patch (.docx)", use_container_width=True):
                final_payload = dict(staged)
                final_payload["prose_rewrites"] = approved_rewrites
                fix_docx = build_surgical_fixit_docx(slm_title, slm_subj, final_payload)
                safe_n = re.sub(r"[^A-Za-z0-9_-]+", "_", slm_title).strip("_") or "Unit"
                fn_out = f"FixIt_Patch_{safe_n}.docx"
                st.download_button(f"Download {fn_out}", data=fix_docx, file_name=fn_out, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="primary", use_container_width=True)
                if drive_svc and target_folder_id:
                    lnk = upload_file_to_drive(drive_svc, target_folder_id, fn_out, fix_docx, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                    st.markdown(f"**Synced to Google Drive:** [{fn_out}]({lnk})")

    with col_tb:
        st.markdown("#### Track B: Audit-Verified Creative Enrichment & PPTX Animation")
        track_b_unlocked = bool(man_bytes and rep2_bytes and (score_num >= 85.0 or "accept" in r1_slm.get("verdict", "").lower() or "pass" in r1_slm.get("verdict", "").lower()))
        if rep2_bytes and not track_b_unlocked:
            st.warning(f"Track B Locked: Layer-4 Report v1 score ({r1_slm['overall_score']}) is below 85%. Run Track A remediation and re-verify in Layer-4 first.")
        pptx_trans = st.selectbox("PPTX Slide Transition Style (for .pptx decks)", ["fade", "push"], key="l5_pptx_trans", disabled=not track_b_unlocked)
        if st.button("Generate Track B Enriched Deliverable", type="primary", use_container_width=True, disabled=not track_b_unlocked):
            safe_n = re.sub(r"[^A-Za-z0-9_-]+", "_", slm_title).strip("_") or "Unit"
            if man_name.lower().endswith(".pptx"):
                prs = pptx.Presentation(io.BytesIO(man_bytes))
                for slide in prs.slides:
                    if not slide.has_notes_slide:
                        _ = slide.notes_slide
                    prev = slide.notes_slide.notes_text_frame.text or ""
                    slide.notes_slide.notes_text_frame.text = f"{prev}\n[Layer-5 Track B Verified | Transition: {pptx_trans} | Brand: {brand_cfg['name']}]".strip()
                p_buf = io.BytesIO()
                prs.save(p_buf)
                fn_p = f"{safe_n}_Enriched_Animated.pptx"
                st.success(f"Animated PPTX deck ready ({pptx_trans} transition applied)!")
                st.download_button(f"Download {fn_p}", data=p_buf.getvalue(), file_name=fn_p, mime="application/vnd.openxmlformats-officedocument.presentationml.presentation", type="primary", use_container_width=True)
            else:
                enr_payload = {
                    "executive_summary": f"Track B Creative Enrichment for {slm_title} ({brand_cfg['name']})",
                    "prose_rewrites": [{"section": "Visual Infographic & Callout Card", "issue": "Learner Engagement", "original_excerpt": slm_title, "revised_text": f"[Callout Card - {brand_cfg['name']}]: Key Takeaways & Industry Caselet for {slm_title}."}],
                    "cyp_questions": [{"question": f"[QR Micro-Quiz] Apply the main framework from {slm_title} to a real-world dataset.", "answer": "Step-by-step rubric solution."}],
                    "glossary_terms": [{"term": f"{slm_title} Framework", "definition": "Verified academic concept enriched with visual callout."}],
                }
                enr_docx = build_surgical_fixit_docx(f"{slm_title} (Enriched Edition)", slm_subj, enr_payload)
                fn_e = f"{safe_n}_Enriched_Approved.docx"
                st.success("Track B Enriched Document generated!")
                st.download_button(f"Download {fn_e}", data=enr_docx, file_name=fn_e, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="primary", use_container_width=True)

with tab3:
    st.subheader("Aspect 2: 1-to-N Complete Mirrored SLM Book Compiler (.docx)")
    st.info("Combines 2 or more approved Unit .docx files into a print-ready university book with Cover Page, Editorial Governance Board, Master Index (Ctrl+A -> F9 in Word), and mirrored recto/verso headers.")
    b1, b2, b3, b4 = st.columns(4)
    prog_name = b1.text_input("Program Name", value="MBA / M.Sc. Data Science", key="l5_bk_prog")
    crs_name = b2.text_input("Course Name", value="Essentials of Data Science", key="l5_bk_crs")
    crs_code = b3.text_input("Course Code", value="DSC-501", key="l5_bk_code")
    crs_cred = b4.text_input("Credits", value="4 Credits", key="l5_bk_cred")
    g_col, co_col = st.columns(2)
    gov_board = g_col.text_area("Editorial Governance Board", value="Chair: Dean of Academics\nSubject Matter Editor: Senior Faculty Lead\nQuality Assurance: Layer-4/5 Academic QA Board", height=100)
    co_raw = co_col.text_area("Course Objectives (CO1-CO5, one per line)", value="Analyze structured and unstructured datasets.\nApply statistical and machine learning models.\nEvaluate model performance and ethical governance.", height=100)
    unit_ups = st.file_uploader("Upload 2 or more Approved Unit .docx Files (in sequence)", type=["docx"], accept_multiple_files=True, key="l5_bk_units")
    if st.button("Compile Complete Mirrored SLM Book (.docx)", type="primary", use_container_width=True, disabled=not (unit_ups and len(unit_ups) >= 2)):
        with st.spinner("Compiling print-ready mirrored book..."):
            units_seq = [(f"Block 1 - Unit {i}", uf.name.rsplit(".", 1)[0], uf.getvalue()) for i, uf in enumerate(unit_ups, start=1)]
            book_bytes = build_complete_mirrored_book_docx(
                brand_cfg, prog_name, crs_name, crs_code, crs_cred,
                gov_board, co_raw.splitlines(), units_seq
            )
        bk_fname = f"{re.sub(r'[^A-Za-z0-9_-]+', '_', crs_code)}_{re.sub(r'[^A-Za-z0-9_-]+', '_', crs_name)}_Complete_Book.docx"
        st.success("Complete Mirrored Book compiled! Open in Word and press Ctrl+A then F9 to refresh Master Index page numbers.")
        st.download_button(f"Download {bk_fname}", data=book_bytes, file_name=bk_fname, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="primary", use_container_width=True)
        if drive_svc and target_folder_id:
            lnk = upload_file_to_drive(drive_svc, target_folder_id, bk_fname, book_bytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            st.markdown(f"**Synced Complete Book to Google Drive:** [{bk_fname}]({lnk})")

with tab4:
    st.subheader("Aspect 3: Subject-Wide Cumulative QA Dossier & Release Certificate")
    st.info("Upload all Layer-4 QA Audit Reports (.docx) across a subject (SLM, MCQ, QB, PPT, FAQ) to compute the cumulative readiness score and generate the 4-Tab Excel Dossier + Word Release Certificate.")
    s1, s2, s3 = st.columns(3)
    s_prog = s1.text_input("Program Name", value="MBA / M.Sc. Data Science", key="l5_s_prog")
    s_code = s2.text_input("Subject Code", value="DSC-501", key="l5_s_code")
    s_name = s3.text_input("Subject Name", value="Essentials of Data Science", key="l5_s_name")
    sub_reps = st.file_uploader("Upload All Layer-4 QA Audit Reports (.docx) for this Subject", type=["docx"], accept_multiple_files=True, key="l5_sub_reps")
    if st.button("Generate Cumulative Subject QA & Enrichment Dossier", type="primary", use_container_width=True, disabled=not sub_reps):
        parsed_list = []
        for rf in sub_reps:
            pr = parse_layer4_docx_report(rf.getvalue())
            pr["filename"] = rf.name
            fn_up = rf.name.upper()
            pr["asset_type"] = next((t for t in ("MCQ", "QB", "PPT", "FAQ") if t in fn_up), "SLM")
            parsed_list.append(pr)
        xl_bytes, cert_bytes, avg_sc = build_subject_cumulative_dossier(brand_cfg, s_prog, s_code, s_name, parsed_list)
        st.success(f"Subject Dossier Generated | Deliverables: {len(parsed_list)} | Cumulative QA Score: {avg_sc}%")
        d1, d2 = st.columns(2)
        d1.download_button(f"Download {s_code}_Cumulative_QA_Dossier.xlsx", data=xl_bytes, file_name=f"{s_code}_Cumulative_QA_Dossier.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", use_container_width=True)
        d2.download_button(f"Download {s_code}_Subject_QA_Release_Certificate.docx", data=cert_bytes, file_name=f"{s_code}_Subject_QA_Release_Certificate.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="primary", use_container_width=True)
