import io
import re
import os
import json
import urllib.request
import urllib.error
import openpyxl
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import docx
from docx.shared import Inches, Pt, RGBColor

# -----------------------------------------------------------------------------
# 1. EXCEL FORMATTING HELPERS
# -----------------------------------------------------------------------------
NAVY_HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
NAVY_HEADER_FONT = Font(name="Arial", size=10, bold=True, color="FFFFFF")
TITLE_FONT = Font(name="Arial", size=14, bold=True)
SUBTITLE_FONT = Font(name="Arial", size=10, italic=True)
BOLD_FONT = Font(name="Arial", size=10, bold=True)
REGULAR_FONT = Font(name="Arial", size=10)
BLUE_INPUT_FONT = Font(name="Arial", size=10, color="002060", bold=True)

def apply_navy_header(ws, row_idx: int, col_count: int):
    ws.row_dimensions[row_idx].height = 24
    for col in range(1, col_count + 1):
        cell = ws.cell(row=row_idx, column=col)
        cell.fill = NAVY_HEADER_FILL
        cell.font = NAVY_HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

def build_scorecard_sheet(wb, title, subtitle, items):
    ws = wb.create_sheet(title="Scorecard")
    ws.cell(row=1, column=1, value=title).font = TITLE_FONT
    ws.cell(row=2, column=1, value=subtitle).font = SUBTITLE_FONT

    headers = ["Criteria", "Max Deduction", "Deducted", "Finding"]
    for i, h in enumerate(headers, 1):
        ws.cell(row=4, column=i, value=h)
    apply_navy_header(ws, 4, 4)

    for idx, itm in enumerate(items, start=5):
        ws.cell(row=idx, column=1, value=itm.get("criteria", "")).font = REGULAR_FONT
        ws.cell(row=idx, column=2, value=float(itm.get("max", -0.1))).number_format = "0.0%"
        c_ded = ws.cell(row=idx, column=3, value=float(itm.get("deducted", 0.0)))
        c_ded.font = BLUE_INPUT_FONT
        c_ded.number_format = "0.0%"
        ws.cell(row=idx, column=4, value=itm.get("finding", "")).font = REGULAR_FONT

    tot_r = len(items) + 5
    ws.cell(row=tot_r, column=1, value="Total Deduction").font = BOLD_FONT
    ws.cell(row=tot_r, column=3, value='=SUM(C5:C' + str(tot_r - 1) + ')').number_format = "0.0%"
    ws.cell(row=tot_r, column=3).font = BOLD_FONT

    ws.cell(row=tot_r + 1, column=1, value="Final Score").font = BOLD_FONT
    ws.cell(row=tot_r + 1, column=3, value='=1+C' + str(tot_r)).number_format = "0.0%"
    ws.cell(row=tot_r + 1, column=3).font = BOLD_FONT

    ws.cell(row=tot_r + 3, column=1, value="Legend: blue figures in column C are reviewer-assigned deductions (inputs); totals recalculate.").font = SUBTITLE_FONT

    ws.column_dimensions["A"].width = 44
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 95
    return ws, tot_r

def build_syllabus_sheet(wb, syl_list):
    ws = wb.create_sheet(title="Syllabus Mapping")
    headers = ["Unit", "Title", "Coverage Status", "Aligned Sections/Items", "Notes / Gaps"]
    for i, h in enumerate(headers, 1):
        ws.cell(row=1, column=i, value=h)
    apply_navy_header(ws, 1, len(headers))

    for idx, s in enumerate(syl_list, start=2):
        ws.cell(row=idx, column=1, value=s.get("unit", "Unit " + str(idx - 1))).font = REGULAR_FONT
        ws.cell(row=idx, column=2, value=s.get("title", "")).font = REGULAR_FONT
        st_cell = ws.cell(row=idx, column=3, value=s.get("status", "Not covered"))
        st_cell.font = REGULAR_FONT
        if s.get("status") == "Fully mapped":
            st_cell.fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
        ws.cell(row=idx, column=4, value=s.get("aligned", "")).font = REGULAR_FONT
        ws.cell(row=idx, column=5, value=s.get("notes", "")).font = REGULAR_FONT

    s_end = len(syl_list) + 1
    r_cov = s_end + 2
    ws.cell(row=r_cov, column=1, value="Units covered").font = BOLD_FONT
    ws.cell(row=r_cov, column=2, value='=COUNTIF(C2:C' + str(s_end) + ', "Fully mapped") & " of ' + str(len(syl_list)) + ' units"').font = BOLD_FONT

    r_pct = s_end + 3
    ws.cell(row=r_pct, column=1, value="Syllabus coverage").font = BOLD_FONT
    ws.cell(row=r_pct, column=2, value='=COUNTIF(C2:C' + str(s_end) + ', "Fully mapped") / ' + str(max(len(syl_list), 1))).number_format = "0.0%"

    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 40
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 30
    ws.column_dimensions["E"].width = 80
    return ws

# -------------------------------------------------------------
# 2. DELIVERABLE-SPECIFIC WORKBOOK GENERATORS
# -------------------------------------------------------------
def generate_mcq_workbook(data: dict) -> io.BytesIO:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws_sc, tot_r = build_scorecard_sheet(wb, data.get("title", "MCQ Bank Quality Assurance Report"), data.get("subtitle", "Percentage-Wise Outcome Report"), data.get("scorecard", []))
    
    ws_ql = wb.create_sheet(title="Question Log")
    headers_ql = ["Q No.", "Unit", "Section", "Question", "A", "B", "C", "D", "Correct Option", "Correct Answer", "Key Check", "Review Flag"]
    for i, h in enumerate(headers_ql, 1): ws_ql.cell(row=1, column=i, value=h)
    apply_navy_header(ws_ql, 1, 12)

    for idx, q in enumerate(data.get("questions", []), 2):
        ws_ql.cell(row=idx, column=1, value=q.get("q_no", "Q" + str(idx - 1))).font = REGULAR_FONT
        ws_ql.cell(row=idx, column=2, value=q.get("unit", "")).font = REGULAR_FONT
        ws_ql.cell(row=idx, column=3, value=q.get("section", "")).font = REGULAR_FONT
        ws_ql.cell(row=idx, column=4, value=q.get("question", "")).font = REGULAR_FONT
        ws_ql.cell(row=idx, column=5, value=q.get("a", "")).font = REGULAR_FONT
        ws_ql.cell(row=idx, column=6, value=q.get("b", "")).font = REGULAR_FONT
        ws_ql.cell(row=idx, column=7, value=q.get("c", "")).font = REGULAR_FONT
        ws_ql.cell(row=idx, column=8, value=q.get("d", "")).font = REGULAR_FONT
        ws_ql.cell(row=idx, column=9, value=str(q.get("correct_option", "")).strip().upper()).font = BOLD_FONT
        ws_ql.cell(row=idx, column=10, value=q.get("correct_answer", "")).font = REGULAR_FONT
        ws_ql.cell(row=idx, column=11, value='=IF(INDEX(E' + str(idx) + ':H' + str(idx) + ', 1, MATCH(I' + str(idx) + ', {"A","B","C","D"}, 0))=J' + str(idx) + ', "Match", "MISMATCH")').font = REGULAR_FONT
        flag = q.get("review_flag", "OK")
        fl_cell = ws_ql.cell(row=idx, column=12, value=flag)
        if flag and flag.strip().upper() != "OK":
            fl_cell.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    ws_ql.column_dimensions["D"].width = 50
    ws_ql.column_dimensions["L"].width = 40

    ws_sc.cell(row=tot_r + 5, column=1, value="Answer distribution").font = BOLD_FONT
    q_end = max(len(data.get("questions", [])) + 1, 2)
    for idx, opt in enumerate(["A", "B", "C", "D"], start=tot_r + 6):
        ws_sc.cell(row=idx, column=1, value="Option " + opt).font = REGULAR_FONT
        ws_sc.cell(row=idx, column=2, value='=COUNTIF(\'Question Log\'!$I$2:$I$' + str(q_end) + ', "' + opt + '")').font = REGULAR_FONT

    ws_rep = wb.create_sheet(title="Repetition")
    for i, h in enumerate(["#", "Q No.", "Q No.", "Repeated concept", "Suggested action"], 1): ws_rep.cell(row=1, column=i, value=h)
    apply_navy_header(ws_rep, 1, 5)
    for idx, r in enumerate(data.get("repetitions", []), 2):
        ws_rep.cell(row=idx, column=1, value=idx - 1)
        ws_rep.cell(row=idx, column=2, value=r.get("q1", ""))
        ws_rep.cell(row=idx, column=3, value=r.get("q2", ""))
        ws_rep.cell(row=idx, column=4, value=r.get("concept", ""))
        ws_rep.cell(row=idx, column=5, value=r.get("action", ""))
    ws_rep.column_dimensions["D"].width = 65

    ws_dis = wb.create_sheet(title="Distractors")
    for i, h in enumerate(["Q No.", "Implausible option(s) / issue"], 1): ws_dis.cell(row=1, column=i, value=h)
    apply_navy_header(ws_dis, 1, 2)
    for idx, d in enumerate(data.get("distractors", []), 2):
        ws_dis.cell(row=idx, column=1, value=d.get("q_no", ""))
        ws_dis.cell(row=idx, column=2, value=d.get("issue", ""))
    ws_dis.column_dimensions["B"].width = 90

    ws_grm = wb.create_sheet(title="Grammar")
    for i, h in enumerate(["Q No.", "Issue", "Suggested rewrite"], 1): ws_grm.cell(row=1, column=i, value=h)
    apply_navy_header(ws_grm, 1, 3)
    for idx, g in enumerate(data.get("grammar", []), 2):
        ws_grm.cell(row=idx, column=1, value=g.get("q_no", ""))
        ws_grm.cell(row=idx, column=2, value=g.get("issue", ""))
        ws_grm.cell(row=idx, column=3, value=g.get("rewrite", ""))
    ws_grm.column_dimensions["B"].width = 65
    ws_grm.column_dimensions["C"].width = 75

    ws_tn = wb.create_sheet(title="Technical Nuances")
    for i, h in enumerate(["Q No.", "Nuance (no deduction)", "Suggested option wording"], 1): ws_tn.cell(row=1, column=i, value=h)
    apply_navy_header(ws_tn, 1, 3)
    for idx, tn in enumerate(data.get("technical_nuances", []), 2):
        ws_tn.cell(row=idx, column=1, value=tn.get("q_no", ""))
        ws_tn.cell(row=idx, column=2, value=tn.get("nuance", ""))
        ws_tn.cell(row=idx, column=3, value=tn.get("suggested_wording", ""))
    ws_tn.column_dimensions["B"].width = 65
    ws_tn.column_dimensions["C"].width = 75

    build_syllabus_sheet(wb, data.get("syllabus", []))

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out

def generate_slm_workbook(data: dict) -> io.BytesIO:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    build_scorecard_sheet(wb, data.get("title", "SLM Pedagogical Quality Audit Report"), data.get("subtitle", "UGC-DEB SLM Compliance Audit"), data.get("scorecard", []))
    
    ws_st = wb.create_sheet(title="Unit Structure Log")
    headers = ["Element ID", "Pedagogical Element", "Compliance Status", "Specific Finding", "Review Flag"]
    for i, h in enumerate(headers, 1): ws_st.cell(row=1, column=i, value=h)
    apply_navy_header(ws_st, 1, 5)
    for idx, itm in enumerate(data.get("structure_log", []), 2):
        ws_st.cell(row=idx, column=1, value=itm.get("id", "SEC-" + str(idx - 1)))
        ws_st.cell(row=idx, column=2, value=itm.get("element", ""))
        ws_st.cell(row=idx, column=3, value=itm.get("status", "Compliant"))
        ws_st.cell(row=idx, column=4, value=itm.get("finding", ""))
        fl_cell = ws_st.cell(row=idx, column=5, value=itm.get("flag", "OK"))
        if itm.get("flag", "OK") != "OK":
            fl_cell.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    ws_st.column_dimensions["B"].width = 30
    ws_st.column_dimensions["D"].width = 65

    ws_gp = wb.create_sheet(title="Pedagogical & Style Gaps")
    headers_gp = ["Section / Element", "Pedagogical / Style Issue", "Recommended Action / Rewrite"]
    for i, h in enumerate(headers_gp, 1): ws_gp.cell(row=1, column=i, value=h)
    apply_navy_header(ws_gp, 1, 3)
    for idx, gp in enumerate(data.get("gaps", []), 2):
        ws_gp.cell(row=idx, column=1, value=gp.get("section", ""))
        ws_gp.cell(row=idx, column=2, value=gp.get("issue", ""))
        ws_gp.cell(row=idx, column=3, value=gp.get("recommendation", ""))
    ws_gp.column_dimensions["A"].width = 25
    ws_gp.column_dimensions["B"].width = 60
    ws_gp.column_dimensions["C"].width = 70

    build_syllabus_sheet(wb, data.get("syllabus", []))

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out

def generate_qb_workbook(data: dict) -> io.BytesIO:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    build_scorecard_sheet(wb, data.get("title", "Question Bank (QB) Quality Assurance Report"), data.get("subtitle", "Cognitive Rigor & Bloom's Distribution Audit"), data.get("scorecard", []))

    ws_qb = wb.create_sheet(title="Question Bank Log")
    headers = ["Q No.", "Module/Unit", "Question Stem", "Allocated Marks", "Cognitive Tier (LOTS/MOTS/HOTS)", "Marking Scheme", "Review Flag"]
    for i, h in enumerate(headers, 1): ws_qb.cell(row=1, column=i, value=h)
    apply_navy_header(ws_qb, 1, 7)
    for idx, q in enumerate(data.get("qb_log", []), 2):
        ws_qb.cell(row=idx, column=1, value=q.get("q_no", "Q" + str(idx - 1)))
        ws_qb.cell(row=idx, column=2, value=q.get("unit", ""))
        ws_qb.cell(row=idx, column=3, value=q.get("stem", ""))
        ws_qb.cell(row=idx, column=4, value=q.get("marks", "5M"))
        ws_qb.cell(row=idx, column=5, value=q.get("tier", "MOTS"))
        ws_qb.cell(row=idx, column=6, value=q.get("model_key", "Verified"))
        fl = ws_qb.cell(row=idx, column=7, value=q.get("flag", "OK"))
        if q.get("flag", "OK") != "OK":
            fl.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    ws_qb.column_dimensions["C"].width = 55
    ws_qb.column_dimensions["F"].width = 45

    ws_cg = wb.create_sheet(title="Cognitive Balance")
    headers_cg = ["Cognitive Tier", "Target Distribution", "Observed Count", "Observed %", "Status / Remarks"]
    for i, h in enumerate(headers_cg, 1): ws_cg.cell(row=1, column=i, value=h)
    apply_navy_header(ws_cg, 1, 5)
    for idx, c in enumerate(data.get("cognitive_tiers", []), 2):
        ws_cg.cell(row=idx, column=1, value=c.get("tier", ""))
        ws_cg.cell(row=idx, column=2, value=c.get("target", ""))
        ws_cg.cell(row=idx, column=3, value=c.get("count", 0))
        ws_cg.cell(row=idx, column=4, value=c.get("percent", "0%"))
        ws_cg.cell(row=idx, column=5, value=c.get("remarks", ""))
    ws_cg.column_dimensions["A"].width = 25
    ws_cg.column_dimensions["E"].width = 60

    build_syllabus_sheet(wb, data.get("syllabus", []))

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out

def generate_ppt_workbook(data: dict) -> io.BytesIO:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    build_scorecard_sheet(wb, data.get("title", "Lecture Presentation (PPT) Audit Report"), data.get("subtitle", "Rule of Six & Speaker Notes Audit"), data.get("scorecard", []))

    ws_sl = wb.create_sheet(title="Slide Log")
    headers = ["Slide No.", "Slide Title", "Bullet Count", "Max Words/Line", "Rule of 6 Check", "Speaker Notes Status", "Review Flag"]
    for i, h in enumerate(headers, 1): ws_sl.cell(row=1, column=i, value=h)
    apply_navy_header(ws_sl, 1, 7)
    for idx, s in enumerate(data.get("slide_log", []), 2):
        ws_sl.cell(row=idx, column=1, value=s.get("slide_no", "Slide " + str(idx - 1)))
        ws_sl.cell(row=idx, column=2, value=s.get("title", ""))
        ws_sl.cell(row=idx, column=3, value=s.get("bullets", 0))
        ws_sl.cell(row=idx, column=4, value=s.get("max_words", 0))
        ws_sl.cell(row=idx, column=5, value=s.get("rule_of_six", "Passed"))
        ws_sl.cell(row=idx, column=6, value=s.get("notes_status", "Complete"))
        fl = ws_sl.cell(row=idx, column=7, value=s.get("flag", "OK"))
        if s.get("flag", "OK") != "OK":
            fl.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    ws_sl.column_dimensions["B"].width = 40
    ws_sl.column_dimensions["F"].width = 30

    ws_pt = wb.create_sheet(title="Slide Delivery & Notes Gaps")
    headers_pt = ["Slide No.", "Delivery / Visual Issue", "Recommended Action"]
    for i, h in enumerate(headers_pt, 1): ws_pt.cell(row=1, column=i, value=h)
    apply_navy_header(ws_pt, 1, 3)
    for idx, p in enumerate(data.get("notes_gaps", []), 2):
        ws_pt.cell(row=idx, column=1, value=p.get("slide_no", ""))
        ws_pt.cell(row=idx, column=2, value=p.get("issue", ""))
        ws_pt.cell(row=idx, column=3, value=p.get("recommendation", ""))
    ws_pt.column_dimensions["B"].width = 60
    ws_pt.column_dimensions["C"].width = 75

    build_syllabus_sheet(wb, data.get("syllabus", []))

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out

def build_docx_dossier_enhanced(audit_data: dict, institution: str, deliverable: str, topic: str) -> io.BytesIO:
    doc = docx.Document()
    title = doc.add_heading("ACADEMIC QUALITY AUDIT REPORT", level=0)
    if title.runs:
        title.runs[0].font.color.rgb = RGBColor(31, 78, 120)
        
    doc.add_paragraph(f"Institution: {institution} | Deliverable: {deliverable} | Unit / Subject: {topic}")
    doc.add_paragraph("=" * 65)

    scorecard = audit_data.get("scorecard", [])
    total_ded = sum(float(x.get("deducted", 0.0)) for x in scorecard)
    final_score = round(max((1.0 + total_ded) * 100, 0), 1)
    
    doc.add_heading(f"Overall Quality Score: {final_score}/100", level=2)
    doc.add_paragraph(f"Summary / Verdict: {audit_data.get('subtitle', 'Evaluated against syllabus coverage and quality rubrics.')}")

    if scorecard:
        doc.add_heading("Compliance & Deduction Scorecard", level=3)
        table = doc.add_table(rows=1, cols=4)
        table.style = 'Table Grid'
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = "Criteria"
        hdr_cells[1].text = "Max Deduction"
        hdr_cells[2].text = "Deducted"
        hdr_cells[3].text = "Audit Findings"
        
        for c in hdr_cells:
            for p in c.paragraphs:
                for r in p.runs:
                    r.font.bold = True
                    
        for itm in scorecard:
            row_cells = table.add_row().cells
            row_cells[0].text = str(itm.get("criteria", ""))
            row_cells[1].text = f"{itm.get('max', 0):.0%}"
            row_cells[2].text = f"{itm.get('deducted', 0):.0%}"
            row_cells[3].text = str(itm.get("finding", ""))

    syl = audit_data.get("syllabus", [])
    if syl:
        doc.add_heading("Syllabus & Curriculum Coverage", level=3)
        s_table = doc.add_table(rows=1, cols=4)
        s_table.style = 'Table Grid'
        s_hdr = s_table.rows[0].cells
        s_hdr[0].text = "Unit"
        s_hdr[1].text = "Title"
        s_hdr[2].text = "Status"
        s_hdr[3].text = "Notes / Gaps"
        for c in s_hdr:
            for p in c.paragraphs:
                for r in p.runs:
                    r.font.bold = True
        for s in syl:
            r_c = s_table.add_row().cells
            r_c[0].text = str(s.get("unit", ""))
            r_c[1].text = str(s.get("title", ""))
            r_c[2].text = str(s.get("status", ""))
            r_c[3].text = str(s.get("notes", ""))

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out

# -------------------------------------------------------------
# 3. UNIVERSAL AI PROMPT & LLM CALLER (HEADER-BASED AUTH, ZERO KEY IN URL)
# -------------------------------------------------------------
def call_llm(prompt: str, engine: str = "Google Gemini") -> str:
    """Bulletproof LLM caller with dynamic model resolution and current 2026 Gemini models."""
    import os
    import json
    import urllib.request
    import urllib.error

    gemini_key = None
    openai_key = None

    # Retrieve keys from Streamlit Secrets
    try:
        import streamlit as st
        for k in ["GEMINI_API_KEY", "GOOGLE_API_KEY", "gemini_api_key", "google_api_key", "GEMINI_KEY"]:
            if k in st.secrets and str(st.secrets[k]).strip():
                gemini_key = str(st.secrets[k]).strip().strip("'\"")
                break
        for k in ["OPENAI_API_KEY", "openai_api_key", "OPENAI_KEY"]:
            if k in st.secrets and str(st.secrets[k]).strip():
                openai_key = str(st.secrets[k]).strip().strip("'\"")
                break
        if not gemini_key and "gemini" in st.secrets:
            try:
                gemini_key = str(st.secrets["gemini"]["api_key"]).strip().strip("'\"")
            except Exception:
                pass
        if not openai_key and "openai" in st.secrets:
            try:
                openai_key = str(st.secrets["openai"]["api_key"]).strip().strip("'\"")
            except Exception:
                pass
    except Exception:
        pass

    # Environment variable fallback
    if not gemini_key:
        for k in ["GEMINI_API_KEY", "GOOGLE_API_KEY", "gemini_api_key", "google_api_key"]:
            if os.environ.get(k) and os.environ.get(k).strip():
                gemini_key = os.environ.get(k).strip().strip("'\"")
                break
    if not openai_key:
        for k in ["OPENAI_API_KEY", "openai_api_key"]:
            if os.environ.get(k) and os.environ.get(k).strip():
                openai_key = os.environ.get(k).strip().strip("'\"")
                break

    api_errors = []

    # 1. Try Gemini with modern models & auto-discovery
    if gemini_key:
        models_to_try = [
            "gemini-3.8-flash",
            "gemini-3.6-flash",
            "gemini-2.5-flash",
            "gemini-3.5-flash-lite",
            "gemini-2.5-flash-lite"
        ]

        # Dynamically discover active models on this key via ListModels
        try:
            list_url = "https://generativelanguage.googleapis.com/v1beta/models"
            list_req = urllib.request.Request(list_url, headers={"x-goog-api-key": gemini_key})
            with urllib.request.urlopen(list_req, timeout=10) as l_resp:
                m_data = json.loads(l_resp.read().decode("utf-8"))
                discovered = []
                for m in m_data.get("models", []):
                    methods = m.get("supportedGenerationMethods", [])
                    if "generateContent" in methods:
                        clean_name = m.get("name", "").replace("models/", "")
                        if "flash" in clean_name.lower():
                            discovered.append(clean_name)
                if discovered:
                    models_to_try = discovered + [m for m in models_to_try if m not in discovered]
        except Exception:
            pass

        for m_name in models_to_try:
            url = f"https://generativelanguage.googleapis.com/v1beta/models/{m_name}:generateContent"
            payload = {
                "contents": [
                    {
                        "role": "user",
                        "parts": [{"text": prompt}]
                    }
                ],
                "generationConfig": {
                    "responseMimeType": "application/json"
                }
            }
            req = urllib.request.Request(
                url,
                data=json.dumps(payload).encode("utf-8"),
                headers={
                    "Content-Type": "application/json",
                    "x-goog-api-key": gemini_key
                }
            )
            try:
                with urllib.request.urlopen(req, timeout=90) as response:
                    res_obj = json.loads(response.read().decode("utf-8"))
                    text_out = res_obj["candidates"][0]["content"]["parts"][0]["text"]
                    if text_out and text_out.strip():
                        return text_out
            except urllib.error.HTTPError as h_err:
                body = h_err.read().decode("utf-8", errors="ignore")
                api_errors.append(f"Gemini {m_name} (HTTP {h_err.code}): {body[:180]}")
            except Exception as e:
                api_errors.append(f"Gemini {m_name} Error: {str(e)[:180]}")

    # 2. Try OpenAI fallback
    if openai_key:
        for o_model in ["gpt-4o", "gpt-4o-mini"]:
            try:
                from openai import OpenAI
                client = OpenAI(api_key=openai_key)
                resp = client.chat.completions.create(
                    model=o_model,
                    messages=[{"role": "user", "content": prompt}],
                    response_format={"type": "json_object"}
                )
                text_out = resp.choices[0].message.content
                if text_out and text_out.strip():
                    return text_out
            except Exception as o_err:
                api_errors.append(f"OpenAI {o_model} Error: {str(o_err)[:180]}")

    error_summary = "\n".join([f"• {err}" for err in api_errors])
    raise RuntimeError(
        f"LLM API request could not complete across available models.\n\n"
        f"Diagnostic Details:\n{error_summary}\n\n"
        f"Gemini key present: {'Yes' if gemini_key else 'No'} | "
        f"OpenAI key present: {'Yes' if openai_key else 'No'}"
    )

def run_universal_academic_audit(doc_text: str, deliverable: str, syllabus_text: str = "", engine: str = "Google Gemini") -> dict:
    prompts = {
        "MCQ": f"""You are a Principal Academic QA Reviewer auditing a Multiple Choice Question (MCQ) Bank.
Audit against:
1. Subject Knowledge / Factual Errors (Max: -30%)
2. Technical Errors / Option Mismatch (Max: -30%): Option letter must match answer text exactly.
3. Clarity & Repetition (Max: -20%): Repeated concepts across questions and weak/implausible distractors.
4. Grammar & Semantics (Max: -10%): Awkward phrasing or forced question stems.
5. Subject / Syllabus Mismatch (Max: -10%): Questions out of curriculum scope.

MANUSCRIPT: {doc_text[:35000]}
SYLLABUS: {syllabus_text[:10000] if syllabus_text else "Extract units and topics from question labels."}

Return ONLY valid JSON:
{{
  "title": "MCQ Bank Quality Assurance & Deduction Report",
  "subtitle": "Evaluated against syllabus coverage and deduction rubrics.",
  "scorecard": [
    {{"criteria": "Subject Knowledge / Factual Errors", "max": -0.3, "deducted": 0.0, "finding": "..."}},
    {{"criteria": "Technical Errors (Option/Answer Mismatch)", "max": -0.3, "deducted": 0.0, "finding": "..."}},
    {{"criteria": "Clarity & Repetition", "max": -0.2, "deducted": -0.05, "finding": "..."}},
    {{"criteria": "Grammar & Semantics", "max": -0.1, "deducted": 0.0, "finding": "..."}},
    {{"criteria": "Subject Mismatch", "max": -0.1, "deducted": 0.0, "finding": "..."}}
  ],
  "questions": [
    {{"q_no": "Q1", "unit": "Unit 1", "section": "1.1", "question": "...", "a": "...", "b": "...", "c": "...", "d": "...", "correct_option": "A", "correct_answer": "...", "review_flag": "OK"}}
  ],
  "repetitions": [{{"q1": "Q1", "q2": "Q10", "concept": "...", "action": "Delete Q10"}}],
  "distractors": [{{"q_no": "Q3", "issue": "..."}}],
  "grammar": [{{"q_no": "Q3", "issue": "...", "rewrite": "..."}}],
  "technical_nuances": [{{"q_no": "Q5", "nuance": "...", "suggested_wording": "..."}}],
  "syllabus": [{{"unit": "Unit 1", "title": "...", "status": "Fully mapped", "aligned": "Q1-Q10", "notes": "..."}}]
}}""",

        "SLM": f"""You are a Principal Academic QA Reviewer auditing Self-Learning Material (SLM).
Audit against:
1. Pedagogical Architecture & Prelims (Max: -20%): Learning objectives with active Bloom's verbs, unit outline, introduction.
2. Concept Exposition & Depth (Max: -25%): Logical heading progression, examples, real-world case vignettes.
3. Self-Assessment (SAQs) & In-Text Checks (Max: -20%): Formative checks with answer keys after each section.
4. Summary, Glossary & Postlims (Max: -15%): Unit summary, terminal questions, key terms, reference readings.
5. Syllabus Reconciliation & Scope (Max: -20%): Complete coverage without drift.

MANUSCRIPT: {doc_text[:35000]}
SYLLABUS: {syllabus_text[:10000] if syllabus_text else "Extract topics from unit sections."}

Return ONLY valid JSON:
{{
  "title": "SLM Pedagogical Quality Assurance & Deduction Report",
  "subtitle": "Evaluated against UGC-DEB SLM benchmarks and Bloom's Revised Taxonomy.",
  "scorecard": [
    {{"criteria": "Pedagogical Architecture & Prelims", "max": -0.20, "deducted": 0.0, "finding": "..."}},
    {{"criteria": "Concept Exposition & Depth", "max": -0.25, "deducted": -0.05, "finding": "..."}},
    {{"criteria": "Self-Assessment (SAQs) & In-Text Checks", "max": -0.20, "deducted": 0.0, "finding": "..."}},
    {{"criteria": "Summary, Glossary & Postlims", "max": -0.15, "deducted": 0.0, "finding": "..."}},
    {{"criteria": "Syllabus Reconciliation & Scope", "max": -0.20, "deducted": 0.0, "finding": "..."}}
  ],
  "structure_log": [
    {{"id": "SEC-01", "element": "Unit Objectives", "status": "Compliant", "finding": "...", "flag": "OK"}}
  ],
  "gaps": [
    {{"section": "Section 1.2", "issue": "...", "recommendation": "..."}}
  ],
  "syllabus": [
    {{"unit": "Unit 1", "title": "...", "status": "Fully mapped", "aligned": "Sections 1.1-1.4", "notes": "..."}}
  ]
}}""",

        "QB": f"""You are a Principal Academic QA Reviewer auditing an Academic Question Bank (QB).
Audit against:
1. Subject Knowledge & Technical Precision (Max: -30%): Unambiguous questions, correct concepts.
2. Cognitive Rigor & Bloom's Balance (Max: -25%): Target 30% LOTS (Recall), 50% MOTS (Application), 20% HOTS (Analysis/Evaluation).
3. Mark Scheme & Model Answer Guidance (Max: -25%): Proportional marks, clear evaluation rubrics.
4. Syllabus Alignment & Module Balance (Max: -20%): Proportional representation across syllabus units.

MANUSCRIPT: {doc_text[:35000]}
SYLLABUS: {syllabus_text[:10000] if syllabus_text else "Extract module boundaries from question bank."}

Return ONLY valid JSON:
{{
  "title": "Question Bank (QB) Quality Assurance Report",
  "subtitle": "Cognitive Rigor, Mark Scheme & Bloom's Distribution Audit.",
  "scorecard": [
    {{"criteria": "Subject Knowledge & Technical Precision", "max": -0.30, "deducted": 0.0, "finding": "..."}},
    {{"criteria": "Cognitive Rigor & Bloom's Balance", "max": -0.25, "deducted": -0.05, "finding": "..."}},
    {{"criteria": "Mark Scheme & Model Answer Guidance", "max": -0.25, "deducted": 0.0, "finding": "..."}},
    {{"criteria": "Syllabus Alignment & Module Balance", "max": -0.20, "deducted": 0.0, "finding": "..."}}
  ],
  "qb_log": [
    {{"q_no": "Q1", "unit": "Unit 1", "stem": "...", "marks": "5M", "tier": "MOTS", "model_key": "Model guidelines provided", "flag": "OK"}}
  ],
  "cognitive_tiers": [
    {{"tier": "Lower Order (LOTS - Remember/Understand)", "target": "30%", "count": 10, "percent": "33%", "remarks": "Balanced"}},
    {{"tier": "Mid Order (MOTS - Apply/Analyze)", "target": "50%", "count": 15, "percent": "50%", "remarks": "Optimal"}},
    {{"tier": "Higher Order (HOTS - Evaluate/Create)", "target": "20%", "count": 5, "percent": "17%", "remarks": "Slight deficit"}}
  ],
  "syllabus": [
    {{"unit": "Unit 1", "title": "...", "status": "Fully mapped", "aligned": "Q1-Q8", "notes": "..."}}
  ]
}}""",

        "PPT": f"""You are a Principal Academic QA Reviewer auditing a Lecture Slide Presentation (PPT).
Audit against:
1. Rule of Six & Cognitive Load (Max: -25%): Max 6 lines per slide, max 6 words per line. Avoid text-heavy slides.
2. Visual to Text Balance (Max: -25%): Diagrammatic placeholders, clarity, avoiding cognitive overload.
3. Embedded Speaker Notes Completeness (Max: -30%): Comprehensive instructional notes on every slide for faculty delivery.
4. Learning Objective & Syllabus Alignment (Max: -20%): Alignment with module learning goals.

MANUSCRIPT: {doc_text[:35000]}
SYLLABUS: {syllabus_text[:10000] if syllabus_text else "Extract module boundaries from slides."}

Return ONLY valid JSON:
{{
  "title": "Lecture Presentation (PPT) Quality Assurance Report",
  "subtitle": "Rule of Six, Visual Density & Speaker Notes Completeness Audit.",
  "scorecard": [
    {{"criteria": "Rule of Six & Cognitive Load", "max": -0.25, "deducted": 0.0, "finding": "..."}},
    {{"criteria": "Visual to Text Balance", "max": -0.25, "deducted": -0.05, "finding": "..."}},
    {{"criteria": "Embedded Speaker Notes Completeness", "max": -0.30, "deducted": 0.0, "finding": "..."}},
    {{"criteria": "Learning Objective & Syllabus Alignment", "max": -0.20, "deducted": 0.0, "finding": "..."}}
  ],
  "slide_log": [
    {{"slide_no": "Slide 1", "title": "...", "bullets": 4, "max_words": 6, "rule_of_six": "Passed", "notes_status": "Complete", "flag": "OK"}}
  ],
  "notes_gaps": [
    {{"slide_no": "Slide 3", "issue": "...", "recommendation": "..."}}
  ],
  "syllabus": [
    {{"unit": "Unit 1", "title": "...", "status": "Fully mapped", "aligned": "Slides 1-12", "notes": "..."}}
  ]
}}"""
    }

    selected_prompt = prompts.get(deliverable, prompts["SLM"])
    raw_response = call_llm(selected_prompt, engine=engine)

    clean_json = re.sub(r"^```json\s*", "", raw_response.strip())
    clean_json = re.sub(r"\s*```$", "", clean_json.strip())
    return json.loads(clean_json)