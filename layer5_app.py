import io, json, re, zipfile, openpyxl, pptx
import pandas as pd
import streamlit as st
from datetime import datetime
from layer3_auth import require_team_auth
from layer5_content_enhancer import (
    auto_clean_mcq_workbook,
    build_surgical_fixit_prompt,
    build_surgical_fixit_docx,
    parse_layer4_docx_report,
    extract_manuscript_text,
    build_updated_manuscript_bytes,
)
from layer5_cloud import (
    get_layer5_gcp_services,
    extract_drive_id,
    download_drive_file_bytes,
    upload_file_to_drive,
    upload_to_drive_hierarchy,
    sanitize_drive_name,
    load_master_tracker_tabs,
    read_master_tracker_tab,
    sync_layer5_sheet_row,
    call_layer5_llm,
    build_complete_mirrored_book_docx,
    build_subject_cumulative_dossier,
)

st.session_state["_app_layer_mode"] = "l5"
st.set_page_config(
    page_title="Layer-5 Workbench: Batch Master Tracker, Remediation & Book Studio",
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
            "Gemini 2.5 Pro (Default)",
            "OpenAI (ChatGPT gpt-4o-mini)",
            "Anthropic (Claude Sonnet 4.6)",
            "Deterministic Offline Scaffolding (No AI Key)",
        ],
        key="l5_ai_provider",
    )
    target_drive_url = st.text_input(
        "Destination Google Drive Root Folder URL (For Batch & Cloud Sync)",
        placeholder="https://drive.google.com/drive/folders/...",
        key="l5_target_drive",
    )
    target_folder_id = extract_drive_id(target_drive_url)

st.markdown("# Layer-5 Academic QA Remediation, Enrichment & Book Studio")
st.caption(
    f"Signed in as **{active_user}** | Zero-Risk Clone-on-Write Guarantee | "
    "Batch Master Tracker Google Sheet & Report-Gated Remediation (R1 -> R2 Verification)"
)

drive_svc, sheets_svc, google_email, gcp_msg, oauth_url = get_layer5_gcp_services()
if drive_svc is not None:
    st.success(f"Google Drive & Master Tracker Connected ({google_email}) - supportsAllDrives=True active")
else:
    c_g1, c_g2 = st.columns([3, 1])
    c_g1.info(f"Google OAuth Required for Master Tracker Batch Sync: {gcp_msg}")
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

tab_batch, tab_mcq, tab_slm, tab_book, tab_subj = st.tabs([
    "1. Batch Master Tracker (Google Sheet Multi-File)",
    "2. Single MCQ / QB Remediator (.xlsx)",
    "3. Aspect 1: Single Unit & PPTX Studio",
    "4. Aspect 2: 1-to-N Complete Book Compiler",
    "5. Aspect 3: Subject-Wide Cumulative QA",
])

with tab_batch:
    st.subheader("Batch Multi-File Remediation via Master Tracker Google Sheet")
    st.info(
        "Connect the same Master Tracker Google Sheet used in Layer-4. Layer-5 reads each selected row's Manuscript Link AND "
        "Layer-4 QA Report Link (.docx), builds the full Updated Manuscript (.docx / .pptx / .xlsx), saves it into Google Drive, "
        "and writes 'Layer-5 Remediated Link' & 'Layer-5 Status' directly into the next columns of the sheet."
    )
    if drive_svc is None or sheets_svc is None:
        st.warning("Click 'Sign In with Google (Drive & Sheets)' at the top of the page to load and update your Master Tracker Google Sheet.")
    else:
        tracker_url = st.text_input(
            "Master Tracker Google Sheet URL (Same Sheet Used in Layer-4)",
            placeholder="https://docs.google.com/spreadsheets/d/1AbC.../edit",
            key="l5_tracker_url",
        )
        tracker_id = extract_drive_id(tracker_url)
        if tracker_id:
            try:
                tabs = load_master_tracker_tabs(sheets_svc, tracker_id)
                selected_tab = st.selectbox("Select Worksheet Tab", tabs, key="l5_tracker_tab")
                hdr_row_1idx, raw_headers, data_rows = read_master_tracker_tab(sheets_svc, tracker_id, selected_tab)

                if not raw_headers or not data_rows:
                    st.warning("No data rows found in the selected sheet tab.")
                else:
                    def guess_idx(keywords, allow_none=False):
                        opts = (["(None)"] + raw_headers) if allow_none else raw_headers
                        for i, o in enumerate(opts):
                            if any(k in o.lower() for k in keywords):
                                return i
                        return 0

                    none_opts = ["(None)"] + raw_headers
                    c_m1, c_m2, c_m3, c_m4 = st.columns(4)
                    col_ms = c_m1.selectbox("Manuscript Link Column (Required)", raw_headers, index=guess_idx(["manuscript", "content link", "file link", "drive link", "url"]), key="l5_col_ms")
                    col_rep = c_m2.selectbox("Layer-4 QA Report Link Column (Required)", raw_headers, index=guess_idx(["qa report", "report link", "audit report", "l4 report"]), key="l5_col_rep")
                    col_subj = c_m3.selectbox("Subject / Course Column", none_opts, index=guess_idx(["subject", "course", "paper"], True), key="l5_col_subj")
                    col_topic = c_m4.selectbox("Topic / Unit Column", none_opts, index=guess_idx(["topic", "unit", "module", "chapter", "title"], True), key="l5_col_topic")

                    c_m5, c_m6, c_m7, c_m8 = st.columns(4)
                    col_univ = c_m5.selectbox("University Column", none_opts, index=guess_idx(["university", "univ", "partner"], True), key="l5_col_univ")
                    col_prog = c_m6.selectbox("Program Column", none_opts, index=guess_idx(["program", "degree"], True), key="l5_col_prog")
                    col_sem = c_m7.selectbox("Semester Column", none_opts, index=guess_idx(["semester", "sem", "term"], True), key="l5_col_sem")
                    col_l4st = c_m8.selectbox("Layer-4 Audit Status Column", none_opts, index=guess_idx(["audit status", "status"], True), key="l5_col_l4st")

                    def _cell(row_vals, col_name):
                        if not col_name or col_name == "(None)":
                            return ""
                        ci = raw_headers.index(col_name)
                        return str(row_vals[ci]).strip() if ci < len(row_vals) else ""

                    lower_h_scan = [h.lower() for h in raw_headers]
                    latest_ms_col = raw_headers[lower_h_scan.index("latest manuscript link")] if "latest manuscript link" in lower_h_scan else (
                        raw_headers[lower_h_scan.index("layer-5 remediated link")] if "layer-5 remediated link" in lower_h_scan else None
                    )
                    latest_rp_col = raw_headers[lower_h_scan.index("latest qa report link")] if "latest qa report link" in lower_h_scan else None
                    hist_col_existing = raw_headers[lower_h_scan.index("audit & remediation history")] if "audit & remediation history" in lower_h_scan else None

                    preview_records = []
                    for offset, rvals in enumerate(data_rows):
                        sheet_r = hdr_row_1idx + 1 + offset
                        orig_ms_u = _cell(rvals, col_ms)
                        prev_l5_u = _cell(rvals, latest_ms_col) if latest_ms_col else ""
                        ms_u = prev_l5_u if extract_drive_id(prev_l5_u) else orig_ms_u
                        prev_rp_u = _cell(rvals, latest_rp_col) if latest_rp_col else ""
                        rp_u = prev_rp_u if extract_drive_id(prev_rp_u) else _cell(rvals, col_rep)
                        hist_txt = _cell(rvals, hist_col_existing) if hist_col_existing else ""
                        cycle_num = hist_txt.count("[Layer-5 v") + 1
                        if not ms_u and not rp_u:
                            continue
                        ready = bool(extract_drive_id(ms_u) and extract_drive_id(rp_u))
                        preview_records.append({
                            "Select": ready,
                            "Sheet Row": sheet_r,
                            "Subject": _cell(rvals, col_subj) or "General",
                            "Topic / Unit": _cell(rvals, col_topic) or f"Row_{sheet_r}",
                            "Layer-4 Status": _cell(rvals, col_l4st),
                            "Gate Check": "READY (Manuscript + L4 Report)" if ready else "LOCKED (Missing L4 Report or Link)",
                            "_ms_url": ms_u,
                            "_rep_url": rp_u,
                            "_cycle_num": cycle_num,
                            "_univ": _cell(rvals, col_univ) or brand_cfg["name"],
                            "_prog": _cell(rvals, col_prog) or "Program",
                            "_sem": _cell(rvals, col_sem) or "Semester",
                        })

                    if preview_records:
                        df_prev = pd.DataFrame(preview_records)
                        edited_df = st.data_editor(
                            df_prev[["Select", "Sheet Row", "Subject", "Topic / Unit", "Layer-4 Status", "Gate Check"]],
                            use_container_width=True,
                            hide_index=True,
                            disabled=["Sheet Row", "Subject", "Topic / Unit", "Layer-4 Status", "Gate Check"],
                            key="l5_batch_editor",
                        )
                        selected_rows = [
                            preview_records[i]
                            for i, sel in enumerate(edited_df["Select"].tolist())
                            if sel and preview_records[i]["Gate Check"].startswith("READY")
                        ]
                        st.caption(f"{len(selected_rows)} gate-verified row(s) selected for batch Layer-5 execution.")
                        if not target_folder_id:
                            st.warning("Paste a Destination Google Drive Root Folder URL in the left sidebar so updated manuscripts can be saved and linked in the next columns of the sheet.")

                        if st.button(f"Run Batch Layer-5 Remediation on {len(selected_rows)} Selected Row(s)", type="primary", use_container_width=True, disabled=not (selected_rows and target_folder_id)):
                            lower_hdrs = [h.lower() for h in raw_headers]
                            write_l5_hdrs = False
                            write_hist_hdr = False
                            next_free_col = len(raw_headers)
                            if "latest manuscript link" in lower_hdrs:
                                l5_link_idx = lower_hdrs.index("latest manuscript link")
                            elif "layer-5 remediated link" in lower_hdrs:
                                l5_link_idx = lower_hdrs.index("layer-5 remediated link")
                            else:
                                l5_link_idx = next_free_col
                                next_free_col += 1
                                write_l5_hdrs = True
                            if "layer-5 status" in lower_hdrs:
                                l5_stat_idx = lower_hdrs.index("layer-5 status")
                            else:
                                l5_stat_idx = next_free_col
                                next_free_col += 1
                                write_l5_hdrs = True
                            if "audit & remediation history" in lower_hdrs:
                                hist_idx = lower_hdrs.index("audit & remediation history")
                            else:
                                hist_idx = next_free_col
                                next_free_col += 1
                                write_hist_hdr = True

                            prog_bar = st.progress(0.0)
                            batch_logs = []
                            zip_buf = io.BytesIO()
                            with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                                for step_i, item in enumerate(selected_rows, start=1):
                                    s_row = item["Sheet Row"]
                                    subj = item["Subject"]
                                    topic = item["Topic / Unit"]
                                    try:
                                        m_bytes, m_fname = download_drive_file_bytes(drive_svc, extract_drive_id(item["_ms_url"]))
                                        r_bytes, r_fname = download_drive_file_bytes(drive_svc, extract_drive_id(item["_rep_url"]))
                                        r1_parsed = parse_layer4_docx_report(r_bytes)
                                        ts = datetime.now().strftime("%Y%m%d_%H%M")
                                        hierarchy = [item["_univ"], item["_prog"], item["_sem"], subj, "Layer-5 Updated Manuscripts"]

                                        cyc_n = item.get("_cycle_num", 1)
                                        if m_fname.lower().endswith((".xlsx", ".xlsm")):
                                            def _batch_mcq_llm_fn(batch_items, rep_dict):
                                                items_payload = [
                                                    {"sr_no": b[2], "question": b[3], "syllabus_subpoint": b[4], "flagged_in_l4": b[5]}
                                                    for b in batch_items
                                                ]
                                                p_mcq = (
                                                    "You are a Psychometric Assessment Editor fixing flagged MCQ items per a Layer-4 QA Audit Report.\n"
                                                    "Rules:\n"
                                                    "1. Replace any forced 'Where/When' stem opening where the answer is a concept/mechanism with 'Which/What'.\n"
                                                    "2. Diversify repetitive 'Identify the / Choose the / Select the' openings into natural academic phrasing without changing the correct answer.\n"
                                                    "3. If a question's syllabus_subpoint is flagged as drifted in the L4 findings below, correct the syllabus_subpoint.\n"
                                                    f"L4 Findings & Rubric Breakdown: {json.dumps((rep_dict.get('technical_errors', []) + rep_dict.get('score_breakdown', []))[:20], ensure_ascii=False)}\n"
                                                    f"Items: {json.dumps(items_payload, ensure_ascii=False)}\n"
                                                    'Return strictly a JSON object: {"items": [{"sr_no": 1, "question": "...", "syllabus_subpoint": "...", "reason": "..."}]}'
                                                )
                                                res_j = call_layer5_llm(p_mcq, ai_provider, api_keys, {"items": []})
                                                raw_items = res_j if isinstance(res_j, list) else (res_j.get("items", []) if isinstance(res_j, dict) else [])
                                                fixed_map = {}
                                                for x in raw_items:
                                                    if isinstance(x, dict) and x.get("sr_no") is not None:
                                                        m_num = re.search(r"\d+", str(x["sr_no"]))
                                                        if m_num:
                                                            fixed_map[int(m_num.group(0))] = x
                                                return [fixed_map.get(b[2], {"question": b[3], "syllabus_subpoint": b[4]}) for b in batch_items]

                                            out_bytes, stats = auto_clean_mcq_workbook(m_bytes, audit_res=r1_parsed, llm_rephrase_fn=_batch_mcq_llm_fn)
                                            wb = openpyxl.load_workbook(io.BytesIO(out_bytes))
                                            # Populate Model Answer Key on Subjective sheet if present and missing answers
                                            subj_sheets = [s for s in wb.sheetnames if "subjective" in s.lower()]
                                            try:
                                              if subj_sheets and not ai_provider.startswith("Deterministic"):
                                                ws_sub = wb[subj_sheets[0]]
                                                s_hdrs = [str(ws_sub.cell(1, c).value or "").strip().lower() for c in range(1, ws_sub.max_column + 1)]
                                                if "question" in s_hdrs and not any("answer" in h or "solution" in h for h in s_hdrs):
                                                    q_c = s_hdrs.index("question") + 1
                                                    sr_c = (s_hdrs.index("sr. no.") + 1) if "sr. no." in s_hdrs else 1
                                                    ans_c = ws_sub.max_column + 1
                                                    ws_sub.cell(1, ans_c, value="Model Answer / Key Points")
                                                    ws_log_ref = wb["Audit_Change_Log"] if "Audit_Change_Log" in wb.sheetnames else None
                                                    for chunk_start in range(2, ws_sub.max_row + 1, 20):
                                                        chunk_q = []
                                                        for r_i in range(chunk_start, min(chunk_start + 20, ws_sub.max_row + 1)):
                                                            q_val = str(ws_sub.cell(r_i, q_c).value or "").strip()
                                                            sr_val = ws_sub.cell(r_i, sr_c).value or (r_i - 1)
                                                            if q_val:
                                                                chunk_q.append({"row": r_i, "sr_no": sr_val, "question": q_val})
                                                        if chunk_q:
                                                            p_sub = (
                                                                "Generate concise, UGC-aligned academic model answer key points (3-4 bullet points per question) for these subjective questions.\n"
                                                                f"Questions: {json.dumps(chunk_q, ensure_ascii=False)}\n"
                                                                'Return strictly JSON: {"answers": [{"row": 2, "model_answer": "..."}]}'
                                                            )
                                                            ans_res = call_layer5_llm(p_sub, ai_provider, api_keys, {"answers": []})
                                                            raw_ans = ans_res if isinstance(ans_res, list) else (ans_res.get("answers", []) if isinstance(ans_res, dict) else [])
                                                            ans_map = {}
                                                            for a in raw_ans:
                                                                if isinstance(a, dict) and a.get("row") is not None:
                                                                    m_r = re.search(r"\d+", str(a["row"]))
                                                                    if m_r:
                                                                        ans_map[int(m_r.group(0))] = str(a.get("model_answer", "")).strip()
                                                            for item_q in chunk_q:
                                                                ma_txt = ans_map.get(item_q["row"], "")
                                                                if ma_txt:
                                                                    ws_sub.cell(item_q["row"], ans_c, value=ma_txt)
                                                                    stats["metadata_realigned"] = stats.get("metadata_realigned", 0) + 1
                                                                    if ws_log_ref is not None:
                                                                        ws_log_ref.append([subj_sheets[0].strip(), item_q["row"], item_q["sr_no"], "Added Model Answer Key per L4 Finding", "(missing)", ma_txt[:110]])
                                            except Exception as _sub_err:
                                                st.warning(f"Subjective model answer generation partial/skipped: {_sub_err}")
                                            ws_b = wb.create_sheet(title="L4_Report_v1_Baseline")
                                            ws_b.append(["Layer-4 Report File", r_fname])
                                            ws_b.append(["Baseline Score (R1)", str(r1_parsed.get("overall_score", ""))])
                                            ws_b.append(["Total L4 Findings", r1_parsed.get("total_findings", 0)])
                                            ws_b.append(["Headers Fixed", stats.get("headers_fixed", 0)])
                                            ws_b.append(["Keys Normalized", stats.get("keys_normalized", 0)])
                                            ws_b.append(["Metadata Realigned", stats.get("metadata_realigned", 0)])
                                            ws_b.append(["Duplicates Removed", stats.get("duplicates_removed", 0)])
                                            b_io = io.BytesIO()
                                            wb.save(b_io)
                                            final_bytes = b_io.getvalue()
                                            out_name = f"Updated_{sanitize_drive_name(subj)}_{sanitize_drive_name(topic)}_v{cyc_n}_{ts}.xlsx"
                                            mime_t = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                            edits_cnt = sum(stats.get(k, 0) for k in ("headers_fixed", "keys_normalized", "metadata_realigned", "duplicates_removed", "banners_stripped", "stems_diversified"))
                                            st_lbl = f"Remediated v{cyc_n} ({edits_cnt} fixes | R1={r1_parsed.get('overall_score', '')})"
                                        else:
                                            ch_txt = extract_manuscript_text(m_bytes, m_fname)
                                            fb = {
                                                "executive_summary": f"Updated manuscript anchored to {r1_parsed['total_findings']} Layer-4 Report v1 findings.",
                                                "prose_rewrites": [{"section": topic, "issue": f, "original_excerpt": "Flagged in Layer-4 Report v1", "revised_text": f"Revised content resolving: {f}"} for f in (r1_parsed.get("technical_errors", [])[:5] or ["Enhance academic rigor and Bloom depth."])],
                                                "cyp_questions": [{"question": f"Explain key concepts in {topic}.", "answer": "Rubric-aligned model answer."}],
                                                "glossary_terms": [{"term": topic, "definition": "Core instructional concept."}],
                                            }
                                            pr = build_surgical_fixit_prompt(topic, subj, ch_txt[:45000], r1_parsed)
                                            pl = call_layer5_llm(pr, ai_provider, api_keys, fb)
                                            final_bytes, ext_out, mime_t = build_updated_manuscript_bytes(m_bytes, m_fname, topic, subj, pl)
                                            out_name = f"Updated_{sanitize_drive_name(subj)}_{sanitize_drive_name(topic)}_v{cyc_n}_{ts}{ext_out}"
                                            st_lbl = f"Remediated v{cyc_n} (R1={r1_parsed.get('overall_score', '')})"

                                        zf.writestr(out_name, final_bytes)
                                        drv_link = upload_to_drive_hierarchy(drive_svc, target_folder_id, hierarchy, out_name, final_bytes, mime_t)
                                        hist_line = f"[Layer-5 v{cyc_n} | {ts} | {st_lbl}] {drv_link}"
                                        sync_layer5_sheet_row(
                                            sheets_svc=sheets_svc,
                                            spreadsheet_id=tracker_id,
                                            tab_name=selected_tab,
                                            header_row_1idx=hdr_row_1idx,
                                            data_row_1idx=s_row,
                                            l5_link_col_0idx=l5_link_idx,
                                            l5_status_col_0idx=l5_stat_idx,
                                            remediated_link=drv_link,
                                            status_text=st_lbl,
                                            write_headers=(write_l5_hdrs and step_i == 1),
                                            history_col_0idx=hist_idx,
                                            history_entry=hist_line,
                                            write_history_header=(write_hist_hdr and step_i == 1),
                                        )
                                        batch_logs.append({"Sheet Row": s_row, "Subject": subj, "Topic / Unit": topic, "Status": st_lbl, "Updated Manuscript Link": drv_link})
                                    except Exception as row_err:
                                        import traceback
                                        st.error(f"Row {s_row} ({topic}) failed: {row_err}")
                                        st.code(traceback.format_exc())
                                        err_txt = f"Error: {str(row_err)[:110]}"
                                        try:
                                            sync_layer5_sheet_row(sheets_svc, tracker_id, selected_tab, hdr_row_1idx, s_row, l5_link_idx, l5_stat_idx, "", err_txt, write_headers=(write_l5_hdrs and step_i == 1))
                                        except Exception:
                                            pass
                                        batch_logs.append({"Sheet Row": s_row, "Subject": subj, "Topic / Unit": topic, "Status": err_txt, "Updated Manuscript Link": ""})
                                    prog_bar.progress(step_i / len(selected_rows))

                            st.success(f"Batch Layer-5 Complete! Saved {len(batch_logs)} updated manuscript(s) to Drive and wrote links into the next columns of Master Tracker.")
                            st.dataframe(pd.DataFrame(batch_logs), use_container_width=True, hide_index=True)
                            st.download_button(
                                "Download All Updated Manuscripts (.zip)",
                                data=zip_buf.getvalue(),
                                file_name=f"Layer5_Updated_Manuscripts_{datetime.now().strftime('%Y%m%d_%H%M')}.zip",
                                mime="application/zip",
                                type="primary",
                                use_container_width=True,
                            )
            except Exception as exc:
                st.error(f"Could not read Master Tracker Google Sheet: {exc}")

with tab_mcq:
    st.subheader("Single MCQ / QB Excel Remediator (Layer-4 Report v1 Gated)")
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
            out_b = io.BytesIO()
            wb.save(out_b)
            cleaned_bytes = out_b.getvalue()

        out_fname = "Updated_" + (mcq_name if mcq_name.lower().endswith(".xlsx") else mcq_name + ".xlsx")
        st.success("Updated workbook ready with Audit_Change_Log & L4_Report_v1_Baseline! Re-upload into Layer-4 to verify Report v2.")
        st.download_button(f"Download {out_fname}", data=cleaned_bytes, file_name=out_fname, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", use_container_width=True)
        if drive_svc and target_folder_id:
            link = upload_file_to_drive(drive_svc, target_folder_id, out_fname, cleaned_bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.markdown(f"**Synced Updated Workbook to Google Drive:** [{out_fname}]({link})")

with tab_slm:
    st.subheader("Aspect 1: Unit & PPTX Remediation (Track A) & Audit-Verified Enrichment (Track B)")
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
    elif man_bytes:
        st.warning("Remediation & Enrichment Locked: Upload the Layer-4 QA Audit Report (.docx) in Slot 2.")

    col_ta, col_tb = st.columns(2)
    with col_ta:
        st.markdown("#### Track A: Surgical Auto-Remediation (Full Updated Manuscript)")
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
            approved_rewrites = []
            for idx, rw in enumerate(staged.get("prose_rewrites", [])):
                if st.checkbox(f"Approve Fix #{idx+1}: {rw.get('issue', 'Prose Fix')[:85]}", value=True, key=f"l5_chk_{idx}"):
                    approved_rewrites.append(rw)
            if st.button("Compile Full Updated Manuscript", use_container_width=True):
                final_payload = dict(staged)
                final_payload["prose_rewrites"] = approved_rewrites
                upd_bytes, ext_o, mime_o = build_updated_manuscript_bytes(man_bytes, man_name, slm_title, slm_subj, final_payload)
                safe_n = sanitize_drive_name(slm_title, "Unit")
                fn_out = f"Updated_{safe_n}{ext_o}"
                st.download_button(f"Download {fn_out}", data=upd_bytes, file_name=fn_out, mime=mime_o, type="primary", use_container_width=True)
                if drive_svc and target_folder_id:
                    lnk = upload_file_to_drive(drive_svc, target_folder_id, fn_out, upd_bytes, mime_o)
                    st.markdown(f"**Synced Updated Manuscript to Google Drive:** [{fn_out}]({lnk})")

    with col_tb:
        st.markdown("#### Track B: Audit-Verified Creative Enrichment & PPTX Animation")
        track_b_unlocked = bool(man_bytes and rep2_bytes and (score_num >= 85.0 or "accept" in r1_slm.get("verdict", "").lower() or "pass" in r1_slm.get("verdict", "").lower()))
        if rep2_bytes and not track_b_unlocked:
            st.warning(f"Track B Locked: Layer-4 Report v1 score ({r1_slm['overall_score']}) is below 85%. Run Track A remediation first.")
        pptx_trans = st.selectbox("PPTX Slide Transition Style", ["fade", "push"], key="l5_pptx_trans", disabled=not track_b_unlocked)
        if st.button("Generate Track B Enriched Deliverable", type="primary", use_container_width=True, disabled=not track_b_unlocked):
            safe_n = sanitize_drive_name(slm_title, "Unit")
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
                st.download_button(f"Download {fn_p}", data=p_buf.getvalue(), file_name=fn_p, mime="application/vnd.openxmlformats-officedocument.presentationml.presentation", type="primary", use_container_width=True)
            else:
                enr_payload = {
                    "executive_summary": f"Track B Creative Enrichment for {slm_title} ({brand_cfg['name']})",
                    "prose_rewrites": [{"section": "Visual Infographic & Callout Card", "issue": "Learner Engagement", "original_excerpt": slm_title, "revised_text": f"[Callout Card - {brand_cfg['name']}]: Key Takeaways & Industry Caselet for {slm_title}."}],
                    "cyp_questions": [{"question": f"[QR Micro-Quiz] Apply the main framework from {slm_title}.", "answer": "Step-by-step rubric solution."}],
                    "glossary_terms": [{"term": f"{slm_title} Framework", "definition": "Verified academic concept enriched with visual callout."}],
                }
                enr_docx = build_surgical_fixit_docx(f"{slm_title} (Enriched Edition)", slm_subj, enr_payload)
                fn_e = f"{safe_n}_Enriched_Approved.docx"
                st.download_button(f"Download {fn_e}", data=enr_docx, file_name=fn_e, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="primary", use_container_width=True)

with tab_book:
    st.subheader("Aspect 2: 1-to-N Complete Mirrored SLM Book Compiler (.docx)")
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
            book_bytes = build_complete_mirrored_book_docx(brand_cfg, prog_name, crs_name, crs_code, crs_cred, gov_board, co_raw.splitlines(), units_seq)
        bk_fname = f"{sanitize_drive_name(crs_code)}_{sanitize_drive_name(crs_name)}_Complete_Book.docx"
        st.success("Complete Mirrored Book compiled! Open in Word and press Ctrl+A then F9 to refresh Master Index page numbers.")
        st.download_button(f"Download {bk_fname}", data=book_bytes, file_name=bk_fname, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="primary", use_container_width=True)

with tab_subj:
    st.subheader("Aspect 3: Subject-Wide Cumulative QA Dossier & Release Certificate")
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
