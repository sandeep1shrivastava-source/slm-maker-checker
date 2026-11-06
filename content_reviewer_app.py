import io
import json
import os
import zipfile
import time
import random
import smtplib
import re
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import streamlit as st
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from services.universal_qa_builder import (
    run_universal_academic_audit,
    generate_mcq_workbook,
    generate_slm_workbook,
    generate_qb_workbook,
    generate_ppt_workbook,
    build_docx_dossier_enhanced
)

try:
    from pypdf import PdfReader
except ImportError:
    PdfReader = None

# Modular Project Services
# 1. Native Unified Engine tools (Stage 1 and Stage 3 only)
from services.unified_engine import (
    run_stage_1_maker,
    run_stage_3_prompt_audit
)

# 2. Universal Stage 2 QA Engine (Dynamic 7-Tab Reports)
from services.universal_qa_builder import (
    run_universal_academic_audit,
    generate_mcq_workbook,
    generate_slm_workbook,
    generate_qb_workbook,
    generate_ppt_workbook,
    build_docx_dossier_enhanced
)

# 3. Live bridge replacing the hardcoded Stage 2 mock
def run_stage_2_checker(doc_text, chk_syllabus="", engine="Google Gemini"):
    real_text = doc_text if (doc_text and len(str(doc_text).strip()) > 50) else st.session_state.get("draft_text", "")
    deliverable = st.session_state.get("selected_type", "SLM")
    return run_universal_academic_audit(
        doc_text=real_text,
        deliverable=deliverable,
        syllabus_text=chk_syllabus,
        engine=engine
    )
def list_manuscripts():
    """Returns active manuscript queue from session state or uploads."""
    if "manuscript_queue" in st.session_state:
        return st.session_state["manuscript_queue"]
    if "loaded_manuscript_id" in st.session_state and st.session_state["loaded_manuscript_id"]:
        return [{"id": st.session_state["loaded_manuscript_id"], "title": "Current Active Manuscript", "status": "Ready for Audit"}]
    return []
# -------------------------------------------------------------
# PAGE CONFIGURATION
# -------------------------------------------------------------
st.set_page_config(
    page_title="Academic QA & Content Production Pipeline",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -------------------------------------------------------------
# 12-PARAMETER ACADEMIC QA FRAMEWORK & METADATA
# -------------------------------------------------------------
QA_12_PARAMETERS = [
    {"id": 1, "name": "Syllabus & Curriculum Mapping", "weight": 10},
    {"id": 2, "name": "Bloom's Taxonomy & Action Verbs", "weight": 10},
    {"id": 3, "name": "Conceptual Depth & Pedagogical Rigor", "weight": 10},
    {"id": 4, "name": "Real-world Case Studies & Applications", "weight": 8},
    {"id": 5, "name": "Quadrant-II Structural Architecture", "weight": 8},
    {"id": 6, "name": "Self-Assessment & Review Questions", "weight": 8},
    {"id": 7, "name": "Factual & Theoretical Accuracy", "weight": 10},
    {"id": 8, "name": "Originality, Attribution & Non-Hallucination", "weight": 10},
    {"id": 9, "name": "Academic Tone, Style & Language Clarity", "weight": 6},
    {"id": 10, "name": "Key Takeaways & Summary Synthesis", "weight": 6},
    {"id": 11, "name": "Benchmark & Reference Adherence", "weight": 8},
    {"id": 12, "name": "Pedagogical Scaffolding & Formatting", "weight": 6}
]

DEFAULT_PROGRAMS = [
    "Master of Business Administration (MBA)",
    "Master of Science (M.Sc) Data Science",
    "Master of Computer Applications (MCA)",
    "Bachelor of Commerce (B.Com)"
]

# -------------------------------------------------------------
# AUTOMATION FORMATTING GUIDELINES FOR AI MAKER
# -------------------------------------------------------------
FORMATTING_GUIDELINES = """
MANDATORY FORMATTING STANDARDS FOR ACADEMIC GENERATION:
Depending on the deliverable requested, you MUST adhere strictly to the following structural constraints.

1. For SLM (Self-Learning Material):
- Unit Title: Clear, action-oriented.
- Learning Objectives: 3 to 5 objectives explicitly using Bloom's Taxonomy verbs (e.g., Define, Analyze, Evaluate).
- Introduction: Contextual setup bridging previous concepts.
- Core Concepts (Body): Use proper Markdown sub-headings (## and ###).
- Case Studies: Minimum one industry-relevant example, MUST be formatted as a blockquote (>) so the engine can wrap it in a branded callout box.
- Summary / Key Takeaways: Bulleted synthesis.
- Self-Assessment Questions (SAQs): 3 to 5 formative questions.

2. For MCQ & Question Banks:
- Stem Clarity: Unambiguous sentence, no double negatives.
- Options: Exactly four options (A, B, C, D). Avoid 'All of the above' or 'None of the above'.
- Cognitive Distribution: 40% Recall/Understand, 40% Apply/Analyze, 20% Evaluate/Create.
- Mandatory Rationale: Every MCQ must include a 'Correct Answer' key and a detailed 'Explanation/Rationale'.

3. For PPT / Slides:
- Rule of Six: Max 6 bullet points per slide; max 6 words per line.
- No Text Walls: Do not generate dense paragraphs. Use visual cues and bolding.
- Slide Progression: Title Slide -> Unit Objectives -> Concept Slides -> Case Study Slide -> Summary -> Q&A.
- Speaker Notes: Every slide MUST contain detailed, descriptive speaker notes formatted as blockquotes (>).
- Markdown Mapping: Use ## for New Slide (Title and Content), ### for Slide Titles, - for bullet points, and > for Speaker notes.
"""

# -------------------------------------------------------------
# DYNAMIC CONFIG & FEATURE FLAG PERSISTENCE
# -------------------------------------------------------------
def get_config_file_path():
    p1 = os.path.join(os.path.dirname(__file__), "services", "access_config.json")
    if os.path.exists(p1): return p1
    p2 = os.path.join(os.path.dirname(__file__), "access_config.json")
    if os.path.exists(p2): return p2
    return p1

def load_system_config():
    path = get_config_file_path()
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}

def update_system_config(mutator_fn):
    path = get_config_file_path()
    cfg = load_system_config()
    cfg = mutator_fn(cfg)
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=4)
    except Exception as e:
        print(f"Error saving config: {e}")

def save_role_to_config(email: str, role: str):
    def fn(cfg):
        users = cfg.setdefault("users", {})
        if isinstance(users, dict):
            users[email.strip().lower()] = role
        return cfg
    update_system_config(fn)

def save_university_to_config(uni_name: str):
    def fn(cfg):
        unis = cfg.setdefault("universities", [])
        if uni_name.strip() not in unis:
            unis.append(uni_name.strip())
        return cfg
    update_system_config(fn)

def save_program_to_config(prog_name: str):
    def fn(cfg):
        progs = cfg.setdefault("programs", DEFAULT_PROGRAMS.copy())
        if prog_name.strip() not in progs:
            progs.append(prog_name.strip())
        return cfg
    update_system_config(fn)

def save_feature_flags(flags: dict):
    def fn(cfg):
        feat = cfg.setdefault("features", {})
        feat.update(flags)
        return cfg
    update_system_config(fn)

# -------------------------------------------------------------
# EMAIL OTP DISPATCH ENGINE
# -------------------------------------------------------------
def dispatch_email_otp(recipient_email: str, otp_code: str) -> bool:
    """Dispatches a 6-digit OTP via SMTP with an automatic Dev Mode fallback."""
    smtp_cfg = st.secrets.get("smtp", {})
    server_host = smtp_cfg.get("server", "smtp.gmail.com")
    server_port = int(smtp_cfg.get("port", 587))
    sender_email = smtp_cfg.get("sender_email", "")
    sender_password = smtp_cfg.get("sender_password", "")

    # Dev/Test fallback: displays OTP on-screen if SMTP is not yet configured in Secrets
    if not sender_email or not sender_password:
        st.sidebar.warning(f"⚠️ [Dev Mode active]: Test OTP is **{otp_code}**")
        return True

    msg = MIMEMultipart()
    msg["From"] = f"Academic Quality Console <{sender_email}>"
    msg["To"] = recipient_email
    msg["Subject"] = f"Your Login Code: {otp_code}"

    body = f"""Hello,

Your One-Time Password (OTP) for the Academic Publishing & QA Console is:

    ========================================
                    {otp_code}
    ========================================

This code expires in 5 minutes. If you did not request this code, contact your Lead Administrator.

Academic Publishing & QA Pipeline
"""
    msg.attach(MIMEText(body, "plain"))

    try:
        with smtplib.SMTP(server_host, server_port) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.send_message(msg)
        return True
    except Exception as err:
        st.sidebar.error(f"Failed to dispatch OTP email: {err}")
        return False

# -------------------------------------------------------------
# MULTI-FORMAT EXTRACTION (.DOCX, .PDF, .XLSX, .TXT)
# -------------------------------------------------------------
def extract_content_multi_format(uploaded_file) -> str:
    if uploaded_file is None:
        return ""
    fname = getattr(uploaded_file, "name", "").lower()
    raw_bytes = uploaded_file.getvalue() if hasattr(uploaded_file, "getvalue") else uploaded_file.read()
    text_output = []

    if fname.endswith(".docx") or (isinstance(raw_bytes, bytes) and raw_bytes.startswith(b"PK\x03\x04")):
        try:
            doc = Document(io.BytesIO(raw_bytes))
            for p in doc.paragraphs:
                if p.text.strip(): text_output.append(p.text.strip())
            for tbl in doc.tables:
                for row in tbl.rows:
                    row_vals = [c.text.strip() for c in row.cells if c.text.strip()]
                    if row_vals: text_output.append(" | ".join(row_vals))
            return "\n".join(text_output)
        except Exception:
            pass

    if fname.endswith(".pdf") or (isinstance(raw_bytes, bytes) and raw_bytes.startswith(b"%PDF")):
        if PdfReader is None:
            return "[Notice: PDF parser library 'pypdf' is initializing on server. Upload .DOCX or .TXT.]"
        try:
            reader = PdfReader(io.BytesIO(raw_bytes))
            for p_idx, page in enumerate(reader.pages):
                txt = page.extract_text()
                if txt and txt.strip(): text_output.append(f"[Page {p_idx + 1}]\n{txt.strip()}")
            return "\n\n".join(text_output)
        except Exception:
            pass

    if fname.endswith((".xlsx", ".xls")):
        try:
            excel_file = pd.ExcelFile(io.BytesIO(raw_bytes))
            for sheet in excel_file.sheet_names:
                df = pd.read_excel(excel_file, sheet_name=sheet).dropna(how="all")
                text_output.append(f"=== Sheet: {sheet} ===")
                for idx, row in df.iterrows():
                    items = [f"{col}: {val}" for col, val in row.items() if pd.notna(val)]
                    if items: text_output.append(f"[Row {idx + 1}] " + " | ".join(items))
            return "\n".join(text_output)
        except Exception:
            pass

    try:
        return raw_bytes.decode("utf-8")
    except Exception:
        return raw_bytes.decode("latin-1", errors="ignore")

def calculate_realistic_score(weight: float, related_errors: list) -> float:
    score = float(weight)
    for e in related_errors:
        sev = str(e.get("severity", "")).lower()
        if "critical" in sev: score -= (weight * 0.40)
        elif "high" in sev or "major" in sev: score -= (weight * 0.20)
        else: score -= (weight * 0.10)
    return max(0.0, round(score, 1))

def generate_12_param_matrix(stage2_data: dict) -> list:
    """Generates the itemized 12 QA parameter scorecard scored against the active manuscript."""
    raw_params = stage2_data.get("qa_parameters") or stage2_data.get("parameter_matrix") or stage2_data.get("qa_matrix")
    if raw_params and isinstance(raw_params, list) and len(raw_params) > 0:
        return raw_params

    errs = stage2_data.get("error_register", [])
    matrix_rows = []
    for p in QA_12_PARAMETERS:
        p_name = p["name"]
        p_weight = p["weight"]
        related = [
            e for e in errs 
            if p_name.lower() in str(e).lower() 
            or str(e.get("category", "")).lower() in p_name.lower()
            or str(e.get("dimension", "")).lower() in p_name.lower()
        ]
        achieved = calculate_realistic_score(p_weight, related)
        percent_achieved = round((achieved / p_weight) * 100, 1)
        status = "REVISE" if any(str(e.get("severity", "")).lower() == "critical" for e in related) else ("OBSERVATION" if related else "COMPLIANT")
        findings = "; ".join([str(e.get("description", "Observation noted.")) for e in related]) if related else "Meets pedagogical quality standards."
        remediation = "; ".join([str(e.get("recommendation", "None required.")) for e in related]) if related else "No corrective remediation required."
        matrix_rows.append({
            "Param #": p["id"],
            "QA Dimension": p_name,
            "Weight": p_weight,
            "Score Achieved": achieved,
            "Success (%)": f"{percent_achieved}%",
            "Status": status,
            "Auditor Findings": findings,
            "Remediation Guidance": remediation
        })
    return matrix_rows

# -------------------------------------------------------------
# BRANDING & FORMATTING DOCX ENGINE (MAKER STUDIO)
# -------------------------------------------------------------
def hex_to_rgb(hex_code: str) -> RGBColor:
    hex_code = hex_code.lstrip("#")
    return RGBColor(int(hex_code[0:2], 16), int(hex_code[2:4], 16), int(hex_code[4:6], 16))

def build_branded_manuscript_docx(draft_markdown: str, profile: dict) -> io.BytesIO:
    doc = Document()
    primary_rgb = hex_to_rgb(profile.get("primary_color", "002060"))
    accent_hex = profile.get("accent_color", "C09A3E")
    font_name = profile.get("font", "Calibri")
    body_size = profile.get("body_size", 11.0)
    
    # 1. Section Margins & Headers
    for s in doc.sections:
        s.top_margin = Inches(1.0)
        s.bottom_margin = Inches(1.0)
        s.left_margin = Inches(1.0)
        s.right_margin = Inches(1.0)
        
        header_p = s.header.paragraphs[0]
        header_p.text = profile.get("header_text", "")
        header_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        if header_p.runs:
            header_p.runs[0].font.name = font_name
            header_p.runs[0].font.size = Pt(8.5)
            header_p.runs[0].font.color.rgb = RGBColor(120, 120, 120)

    # 2. Cover Page
    if profile.get("logo_file"):
        try:
            doc.add_picture(profile["logo_file"], width=Inches(2.0))
            doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
        except Exception:
            pass
        
    title_p = doc.add_paragraph()
    title_run = title_p.add_run(f"\n{profile.get('university', 'Institutional Manuscript')}\n")
    title_run.bold = True
    title_run.font.name = font_name
    title_run.font.size = Pt(22)
    title_run.font.color.rgb = primary_rgb
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    sub_p = doc.add_paragraph()
    sub_run = sub_p.add_run(f"{profile.get('program', 'Academic Program')} | {profile.get('deliverable', 'SLM')}\n")
    sub_run.font.name = font_name
    sub_run.font.size = Pt(13)
    sub_run.font.italic = True
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_page_break()

    # 3. Content Body
    for line in draft_markdown.split("\n"):
        line_str = line.strip()
        if not line_str: continue

        if line_str.startswith("## "):
            h = doc.add_heading(line_str.replace("## ", ""), level=1)
            h.style.font.name = font_name
            h.style.font.size = Pt(15)
            for r in h.runs:
                r.font.color.rgb = primary_rgb
                r.bold = True
        elif line_str.startswith("### "):
            h = doc.add_heading(line_str.replace("### ", ""), level=2)
            h.style.font.name = font_name
            h.style.font.size = Pt(12.5)
            for r in h.runs:
                r.font.color.rgb = primary_rgb
        elif line_str.startswith(">") and profile.get("callout_boxes"):
            tbl = doc.add_table(rows=1, cols=1)
            tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
            cell = tbl.cell(0, 0)
            shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="F4F6F9"/>')
            borders_elm = parse_xml(
                f'<w:tcBorders {nsdecls("w")}>'
                f'<w:left w:val="single" w:sz="36" w:space="0" w:color="{accent_hex}"/>'
                f'<w:top w:val="none"/><w:right w:val="none"/><w:bottom w:val="none"/>'
                f'</w:tcBorders>'
            )
            cell._tc.get_or_add_tcPr().append(shading_elm)
            cell._tc.get_or_add_tcPr().append(borders_elm)
            p = cell.paragraphs[0]
            p.paragraph_format.left_indent = Inches(0.1)
            r = p.add_run(line_str.lstrip("> ").strip())
            r.font.name = font_name
            r.font.size = Pt(body_size)
            r.font.italic = True
            doc.add_paragraph("")
        else:
            p = doc.add_paragraph()
            r = p.add_run(line_str)
            r.font.name = font_name
            r.font.size = Pt(body_size)

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out

# -------------------------------------------------------------
# AUDIT EXCEL & WORD REPORTS (CHECKER STUDIO)
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
            ["Overall QA Quality Score", f"{score} / 100"],
            ["Syllabus Topic Coverage", f"{coverage}%"],
            ["Critical Defects Count", critical_count],
            ["Final Audit Determination", decision],
            ["Certification Status", "ACCREDITED FOR INSTRUCTION" if decision == "PASS" else "REQUIRES SME REVISION"]
        ]
        pd.DataFrame(summary_rows[1:], columns=summary_rows[0]).to_excel(writer, sheet_name="Executive QA Summary", index=False)
        
        mapping = stage2_data.get("syllabus_mapping", [])
        if mapping:
            map_rows = [{
                "Syllabus Topic / Unit": i.get("topic", "N/A"),
                "Coverage Status": i.get("status", "Covered"),
                "Pedagogical Depth": i.get("depth", "Adequate"),
                "Location": i.get("content_location", "Body"),
                "Gap Analysis": i.get("gap_analysis", "Meets benchmark requirements.")
            } for i in mapping]
            pd.DataFrame(map_rows).to_excel(writer, sheet_name="Syllabus vs Content Mapping", index=False)

        matrix_rows = generate_12_param_matrix(stage2_data)
        pd.DataFrame(matrix_rows).to_excel(writer, sheet_name="12 QA Parameter Matrix", index=False)
        
        errs = stage2_data.get("error_register", [])
        pd.DataFrame(errs if errs else [{"Issue ID": "LOG-00", "Category": "General QA", "Severity": "None", "Description": "Zero defects."}]).to_excel(writer, sheet_name="Error Register", index=False)

        wb = writer.book
        h_fill = PatternFill(start_color="002060", end_color="002060", fill_type="solid")
        h_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        for ws in wb.worksheets:
            for col in ws.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = min(max(max_len + 3, 16), 55)
            for cell in ws[1]:
                cell.fill = h_fill
                cell.font = h_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
    output.seek(0)
    return output

def build_stage2_qa_docx(stage2_data: dict, university: str, program: str, deliverable: str, batch: str = "", semester_year: str = "") -> io.BytesIO:
    doc = Document()
    for s in doc.sections:
        s.top_margin, s.bottom_margin = Inches(0.8), Inches(0.8)
        s.left_margin, s.right_margin = Inches(0.8), Inches(0.8)

    title_p = doc.add_paragraph()
    r_title = title_p.add_run("ACADEMIC CONTENT QUALITY REVIEW REPORT")
    r_title.bold = True
    r_title.font.size = Pt(16)
    r_title.font.color.rgb = RGBColor(0, 51, 102)
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    sub_p = doc.add_paragraph()
    r_sub = sub_p.add_run(f"Stage 2 SME Quality & Syllabus Cross-Audit | {university} - {program}")
    r_sub.font.size = Pt(10)
    r_sub.font.italic = True
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph("―" * 55)

    meta_tbl = doc.add_table(rows=3, cols=2)
    meta_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_tbl.rows[0].cells[0].text = f"University: {university}"
    meta_tbl.rows[0].cells[1].text = f"Determination: {stage2_data.get('decision', 'PASS')}"
    meta_tbl.rows[1].cells[0].text = f"Academic Program: {program}"
    meta_tbl.rows[1].cells[1].text = f"Quality Score: {stage2_data.get('score', 0.0)} / 100"
    meta_tbl.rows[2].cells[0].text = f"Batch / Term: {batch} ({semester_year})"
    meta_tbl.rows[2].cells[1].text = f"Syllabus Coverage: {stage2_data.get('coverage_pct', 0.0)}%"

    # Section 1: Granular Syllabus vs Content Cross-Verification
    h1 = doc.add_heading("\n1. Granular Syllabus vs Content Cross-Verification", level=2)
    h1.style.font.color.rgb = RGBColor(0, 51, 102)
    mapping = stage2_data.get("syllabus_mapping", [])
    if mapping:
        map_tbl = doc.add_table(rows=1, cols=4)
        map_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        map_tbl.style = "Table Grid"
        for idx, text in enumerate(["Syllabus Topic", "Status", "Depth", "Auditor Remarks"]):
            cell = map_tbl.rows[0].cells[idx]
            cell.text = text
            for p in cell.paragraphs:
                for r in p.runs: r.bold = True
        for item in mapping:
            rc = map_tbl.add_row().cells
            rc[0].text = str(item.get("topic", ""))
            rc[1].text = str(item.get("status", ""))
            rc[2].text = str(item.get("depth", ""))
            rc[3].text = str(item.get("gap_analysis", ""))
    else:
        doc.add_paragraph("No external syllabus outline uploaded. Audited against internal pedagogical structure and learning objectives.")

    # Section 2: 12-Parameter Academic QA Evaluation Matrix
    h2 = doc.add_heading("\n2. 12-Parameter Academic Quality Assurance Evaluation Matrix", level=2)
    h2.style.font.color.rgb = RGBColor(0, 51, 102)
    matrix = generate_12_param_matrix(stage2_data)
    matrix_tbl = doc.add_table(rows=1, cols=6)
    matrix_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    matrix_tbl.style = "Table Grid"
    headers_m = ["#", "QA Dimension", "Weight", "Score", "Status", "Auditor Findings"]
    for idx, text in enumerate(headers_m):
        cell = matrix_tbl.rows[0].cells[idx]
        cell.text = text
        for p in cell.paragraphs:
            for r in p.runs: r.bold = True
    for row in matrix:
        rc = matrix_tbl.add_row().cells
        rc[0].text = str(row.get("Param #", ""))
        rc[1].text = str(row.get("QA Dimension", ""))
        rc[2].text = str(row.get("Weight", ""))
        rc[3].text = f"{row.get('Score Achieved', 0.0)}"
        rc[4].text = str(row.get("Status", ""))
        rc[5].text = str(row.get("Auditor Findings", ""))

    # Section 3: Defect & Remediation Register
    h3 = doc.add_heading("\n3. Defect & Remediation Register", level=2)
    h3.style.font.color.rgb = RGBColor(0, 51, 102)
    errs = stage2_data.get("error_register", [])
    if errs:
        err_tbl = doc.add_table(rows=1, cols=5)
        err_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        err_tbl.style = "Table Grid"
        headers_e = ["Issue ID", "Category", "Severity", "Description", "Remediation Guidance"]
        for idx, text in enumerate(headers_e):
            cell = err_tbl.rows[0].cells[idx]
            cell.text = text
            for p in cell.paragraphs:
                for r in p.runs: r.bold = True
        for e in errs:
            rc = err_tbl.add_row().cells
            rc[0].text = str(e.get("issue_id", e.get("id", "LOG-01")))
            rc[1].text = str(e.get("category", "General QA"))
            rc[2].text = str(e.get("severity", "Observation"))
            rc[3].text = str(e.get("description", ""))
            rc[4].text = str(e.get("recommendation", e.get("remediation", "None")))
    else:
        doc.add_paragraph("Zero critical or major defects identified. The manuscript is compliant with institutional standards.")

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out

# -------------------------------------------------------------
# SESSION STATE & OTP AUTHENTICATION GATEWAY
# -------------------------------------------------------------
sys_cfg = load_system_config()
feat_flags = sys_cfg.get("features", {
    "maker_studio_enabled": True,
    "branding_mapper_enabled": True,
    "allow_direct_checker_upload": True
})

if "auth_user" not in st.session_state: st.session_state.auth_user = None
if "auth_role" not in st.session_state: st.session_state.auth_role = None
if "otp_sent" not in st.session_state: st.session_state.otp_sent = False
if "generated_otp" not in st.session_state: st.session_state.generated_otp = None
if "otp_expiry" not in st.session_state: st.session_state.otp_expiry = 0
if "otp_attempts" not in st.session_state: st.session_state.otp_attempts = 0
if "temp_email" not in st.session_state: st.session_state.temp_email = ""
if "temp_role" not in st.session_state: st.session_state.temp_role = None

if "draft_text" not in st.session_state: st.session_state.draft_text = ""
if "stage2_results" not in st.session_state: st.session_state.stage2_results = None
if "stage3_results" not in st.session_state: st.session_state.stage3_results = None
if "loaded_manuscript_id" not in st.session_state: st.session_state.loaded_manuscript_id = None
if "loaded_meta" not in st.session_state: st.session_state.loaded_meta = {}
if "active_batch" not in st.session_state: st.session_state.active_batch = "Batch 2026-A"
if "active_term" not in st.session_state: st.session_state.active_term = "Semester 1"

# --- AUTHENTICATION & ROLE VERIFICATION ---
def verify_user(email):
    """Validates official email and assigns authorization role."""
    email_clean = str(email).strip().lower()
    if not email_clean or "@" not in email_clean:
        return False, None
    if email_clean.endswith("@upgrad.com"):
        return True, "Lead Administrator"
    return True, "Lead Administrator"

st.sidebar.header("🔐 Workspace Login")

if not st.session_state.auth_user:
    # --- PHASE 1: SUBMIT EMAIL FOR OTP ---
    if not st.session_state.otp_sent:
        login_email = st.sidebar.text_input(
            "Enter Official Email ID (@upgrad.com)", 
            placeholder="name@upgrad.com",
            key="login_email_input"
        ).strip().lower()

        if st.sidebar.button("📩 Send Login OTP", type="primary", use_container_width=True):
            if not login_email.endswith("@upgrad.com"):
                st.sidebar.error("⛔ Access Denied: Only official '@upgrad.com' accounts permitted.")
            else:
                is_authorized, user_role = verify_user(login_email)
                if not is_authorized:
                    st.sidebar.error("⛔ Account not registered. Please contact the Lead Administrator.")
                else:
                    new_code = str(random.randint(100000, 999999))
                    st.session_state.generated_otp = new_code
                    st.session_state.otp_expiry = time.time() + 300  # 5-minute validity
                    st.session_state.temp_email = login_email
                    st.session_state.temp_role = user_role
                    st.session_state.otp_attempts = 0

                    with st.sidebar.spinner("Dispatching verification code..."):
                        if dispatch_email_otp(login_email, new_code):
                            st.session_state.otp_sent = True
                            st.sidebar.success(f"Code sent to {login_email}!")
                            st.rerun()

    # --- PHASE 2: VERIFY OTP CODE ---
    else:
        st.sidebar.info(f"Code sent to **{st.session_state.temp_email}** (Valid for 5 mins)")
        entered_code = st.sidebar.text_input("Enter 6-Digit OTP", max_chars=6, key="entered_otp_input")
        
        col_v1, col_v2 = st.sidebar.columns(2)
        with col_v1:
            if st.button("🔓 Verify & Login", type="primary", use_container_width=True):
                if time.time() > st.session_state.otp_expiry:
                    st.sidebar.error("OTP expired. Please request a new code.")
                    st.session_state.otp_sent = False
                    st.session_state.generated_otp = None
                    st.rerun()
                elif st.session_state.otp_attempts >= 3:
                    st.sidebar.error("Too many failed attempts. Please request a new code.")
                    st.session_state.otp_sent = False
                    st.session_state.generated_otp = None
                    st.rerun()
                elif entered_code.strip() == st.session_state.generated_otp:
                    st.session_state.auth_user = st.session_state.temp_email
                    st.session_state.auth_role = st.session_state.temp_role
                    st.session_state.otp_sent = False
                    st.session_state.generated_otp = None
                    st.sidebar.success("Identity verified!")
                    st.rerun()
                else:
                    st.session_state.otp_attempts += 1
                    remaining = 3 - st.session_state.otp_attempts
                    st.sidebar.error(f"Invalid code. {remaining} attempt(s) remaining.")

        with col_v2:
            if st.button("↩️ Cancel", use_container_width=True):
                st.session_state.otp_sent = False
                st.session_state.generated_otp = None
                st.rerun()

    # Stop execution completely if not authenticated
    st.stop()

# --- AUTHENTICATED SESSION CONTEXT ---
user_email = st.session_state.auth_user
user_role = st.session_state.auth_role

st.sidebar.success(f"Verified: **{user_email}**")
st.sidebar.caption(f"Role: **{user_role}**")
if st.sidebar.button("🚪 Logout", use_container_width=True):
    st.session_state.auth_user = None
    st.session_state.auth_role = None
    st.session_state.otp_sent = False
    st.session_state.generated_otp = None
    st.rerun()

# Role Isolation Flags
is_lead_admin = (user_role == "Lead Administrator")
can_maker = is_lead_admin or (user_role in ["Content Author (Maker Team)", "Content Author (Maker)"])
can_checker = is_lead_admin or (user_role in ["Academic Reviewer (Checker & Auditor Team)", "SME Quality Reviewer (Checker)"])

# ADMIN CONSOLE (TEAM, UNIVERSITIES & LIVE FEATURE SWITCHES)
if is_lead_admin:
    with st.sidebar.expander("🛠️ Admin Console & Feature Switches", expanded=False):
        st.markdown("##### 🎛️ Live Feature Switches")
        f_maker = st.toggle("Enable Maker Authoring Studio", value=feat_flags.get("maker_studio_enabled", True))
        f_brand = st.toggle("Enable Branding & Layout Mapper", value=feat_flags.get("branding_mapper_enabled", True))
        f_upload = st.toggle("Allow Direct File Uploads for Checkers", value=feat_flags.get("allow_direct_checker_upload", True))
        
        if st.button("💾 Save Feature Switches", use_container_width=True):
            save_feature_flags({
                "maker_studio_enabled": f_maker,
                "branding_mapper_enabled": f_brand,
                "allow_direct_checker_upload": f_upload
            })
            st.success("Feature flags saved!")
            st.rerun()

        st.divider()
        st.markdown("##### 👤 Manage Team Roles & Permissions")
        tgt_email = st.text_input("Colleague Email (@upgrad.com)", key="adm_email")
        tgt_role = st.selectbox(
            "Assign Team Role",
            ["Content Author (Maker Team)", "Academic Reviewer (Checker & Auditor Team)", "Lead Administrator"],
            key="adm_role"
        )
        if st.button("➕ Save Team Permission", use_container_width=True):
            if tgt_email.strip().endswith("@upgrad.com"):
                save_role_to_config(tgt_email.strip().lower(), tgt_role)
                st.success(f"Updated: {tgt_email} ➔ {tgt_role}")
                st.rerun()
            else:
                st.warning("Please enter a valid @upgrad.com email.")

        st.divider()
        st.markdown("##### 🏛️ Register University")
        new_uni = st.text_input("New University Name", key="adm_uni")
        if st.button("➕ Add University", use_container_width=True):
            if new_uni.strip():
                save_university_to_config(new_uni.strip())
                st.success(f"Added: {new_uni.strip()}")
                st.rerun()
st.divider()
st.markdown("##### 🎓 Register Academic Programme")
new_prog = st.text_input("New Academic Programme Name", key="adm_prog")
if st.button("➕ Add Programme", use_container_width=True):
    if new_prog.strip():
        save_program_to_config(new_prog.strip())
        st.success(f"Added Programme: {new_prog.strip()}")
        st.rerun()

# Audit Scope Mapping
st.sidebar.divider()
st.sidebar.header("🏛️ Operational Scope")
known_unis = ["State Open University", "National Distance University", "Central Higher Education Institute"]
if isinstance(sys_cfg, dict) and "universities" in sys_cfg and sys_cfg["universities"]:
    known_unis = sys_cfg["universities"]

uni_opts = known_unis + ["➕ Register New University..."]
sel_uni = st.sidebar.selectbox("1. University", uni_opts, index=0)
if sel_uni == "➕ Register New University...":
    custom_u = st.sidebar.text_input("New University Name").strip()
    if st.sidebar.button("Save University"):
        save_university_to_config(custom_u); st.rerun()
    selected_university = custom_u or "Custom University"
else:
    selected_university = sel_uni

# Dynamic Programme Selector
known_progs = DEFAULT_PROGRAMS
if isinstance(sys_cfg, dict) and "programs" in sys_cfg and sys_cfg["programs"]:
    known_progs = sys_cfg["programs"]

prog_opts = known_progs + ["➕ Register New Programme..."]
sel_prog = st.sidebar.selectbox("2. Academic Programme", prog_opts, index=0)
if sel_prog == "➕ Register New Programme...":
    custom_p = st.sidebar.text_input("New Programme Name").strip()
    if st.sidebar.button("Save Programme"):
        if custom_p:
            save_program_to_config(custom_p)
            st.rerun()
    selected_program = custom_p or "Custom Programme"
else:
    selected_program = sel_prog
selected_batch = st.sidebar.text_input("3. Batch / Cohort", value=st.session_state.active_batch)
selected_term = st.sidebar.selectbox("4. Academic Term", ["Semester 1", "Semester 2", "Semester 3", "Semester 4"], index=0)
selected_deliverable = st.sidebar.selectbox("5. Deliverable Under Review", ["SLM", "MCQ", "Question Bank", "PPT"], index=0)

# =============================================================
# WORKSPACE ROUTING (STRICT ROLE ISOLATION)
# =============================================================
st.title("Academic Publishing & Quality Assurance Pipeline")
st.caption(f"Authenticated: **{user_email}** | Role: **{user_role}** | Scope: **{selected_university}** ➔ **{selected_program}**")

# Check if Maker Studio is disabled globally by Admin
maker_active = can_maker and feat_flags.get("maker_studio_enabled", True)

# -------------------------------------------------------------
# VIEW 1: MAKER (CONTENT AUTHOR) STUDIO
# -------------------------------------------------------------
if can_maker and not is_lead_admin:
    if not feat_flags.get("maker_studio_enabled", True):
        st.warning("⚠️ The Content Authoring Studio is temporarily under maintenance by the Lead Administrator.")
        st.stop()

    st.header("✍️ Content Authoring Studio (Maker Workspace)")
    tab_write, tab_format, tab_preflight = st.tabs([
        "📝 1. AI Authoring & Drafting",
        "🎨 2. Institutional Brand & Layout Mapper",
        "🧪 3. Pre-Flight Self-Audit & Submission"
    ])

    with tab_write:
        st.markdown("##### Curriculum & Reference Ingestion")
        col_m_syl, col_m_ref = st.columns(2)
        with col_m_syl:
            m_syl_file = st.file_uploader("Upload Syllabus (.docx, .pdf, .txt)", type=["docx", "pdf", "txt"], key="m_syl")
            m_syl_txt = st.text_area("Or Paste Unit Syllabus Topics", placeholder="Unit 1: Foundations\nUnit 2: Frameworks...", height=110)
            m_syllabus = extract_content_multi_format(m_syl_file) if m_syl_file else m_syl_txt

        with col_m_ref:
            m_ref_file = st.file_uploader("Upload Benchmark Textbook Reference", type=["docx", "pdf", "txt"], key="m_ref")
            m_ref_txt = st.text_area("Or Paste Reference Excerpts", placeholder="Prescribed citations and source materials...", height=110)
            m_reference = extract_content_multi_format(m_ref_file) if m_ref_file else m_ref_txt

        if "maker_draft_input" not in st.session_state:
            st.session_state["maker_draft_input"] = st.session_state.get("draft_text", "")

        if st.button("🤖 Generate Draft Manuscript (Stage 1 Maker)", type="primary"):
            if m_syllabus.strip():
                with st.spinner("AI Maker generating academic manuscript according to strict institutional guidelines..."):
                    enhanced_ref = f"{m_reference}\n\n{FORMATTING_GUIDELINES}"
                    gen_m = run_stage_1_maker(m_syllabus, enhanced_ref, "Postgraduate", selected_deliverable)
                    if gen_m and str(gen_m).strip():
                        st.session_state["maker_draft_input"] = str(gen_m)
                        st.session_state.draft_text = str(gen_m)
                        st.success("Draft Generated! Review and refine below.")
                        st.rerun()
                    else:
                        st.error("AI Maker returned empty content. Please verify that GEMINI_API_KEY is active in Streamlit Cloud Secrets.")
            else:
                st.warning("Please supply a syllabus outline.")

        st.text_area(
            "Manuscript Editor (Markdown Supported)", 
            key="maker_draft_input", 
            height=320
        )
        st.session_state.draft_text = st.session_state.get("maker_draft_input", "")

    with tab_format:
        if feat_flags.get("branding_mapper_enabled", True):
            st.markdown("##### 🎨 Institutional Formatting Profile")
            col_f1, col_f2, col_f3 = st.columns(3)
            with col_f1:
                f_font = st.selectbox("Typography Family", ["Calibri", "Times New Roman", "Georgia", "Arial"])
                f_size = st.slider("Body Font Size (pt)", 10.0, 13.0, 11.0, 0.5)
            with col_f2:
                f_primary = st.color_picker("Primary University Color", "#002060")
                f_accent = st.color_picker("Accent Callout Border Color", "#C09A3E")
            with col_f3:
                f_logo = st.file_uploader("Upload University Logo (.PNG, .JPG)", type=["png", "jpg", "jpeg"])
                f_header = st.text_input("Running Header Text", value=f"{selected_university} | {selected_program}")
                f_callouts = st.checkbox("Wrap Case Studies in Branded Callout Boxes", value=True)

            profile = {
                "university": selected_university, "program": selected_program, "deliverable": selected_deliverable,
                "font": f_font, "body_size": f_size, "primary_color": f_primary.replace("#", ""),
                "accent_color": f_accent.replace("#", ""), "header_text": f_header,
                "logo_file": f_logo, "callout_boxes": f_callouts
            }

            if st.session_state.draft_text.strip():
                branded_docx = build_branded_manuscript_docx(st.session_state.draft_text, profile)
                st.download_button(
                    label="📥 Download Branded Word Manuscript (.DOCX)",
                    data=branded_docx,
                    file_name=f"{selected_university}_{selected_program}_Draft.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary"
                )
        else:
            st.info("Institutional Branding Studio is currently locked by the Lead Administrator.")

    with tab_preflight:
        st.markdown("##### 🧪 Pre-Flight Self-Audit & Submission")
        st.caption("Verify your manuscript meets the 12 QA parameters before submitting to the Checker queue.")
        if st.button("Run Instant Self-Audit Check"):
            if st.session_state.draft_text.strip():
                with st.spinner("Running pre-flight quality check..."):
                    res = run_stage_2_checker(st.session_state.draft_text, "", "")
                    st.metric("Estimated QA Score", f"{res.get('score', 0)}/100")
                    st.metric("Critical Defects Detected", res.get("critical_count", 0))
            else:
                st.warning("Workspace draft is empty.")

        st.divider()
        if st.button("📤 Submit Final Draft to Academic Review Queue", type="primary"):
            if st.session_state.draft_text.strip():
                save_manuscript_draft(selected_university, selected_program, selected_batch, selected_deliverable, user_email, st.session_state.draft_text)
                st.success("Draft successfully submitted to the Reviewer Depository! Checkers have been notified.")
            else:
                st.warning("Cannot submit an empty draft.")

# -------------------------------------------------------------
# VIEW 2: CHECKER (QUALITY REVIEWER) WORKSPACE
# -------------------------------------------------------------
elif can_checker and not is_lead_admin:
    st.header("🔍 Academic Quality Assurance Console (Checker Workspace)")

    all_manuscripts = list_manuscripts()
    st.markdown("##### Step 1: Select Manuscript or File for Evaluation")
    
    col_q1, col_q2 = st.columns(2)
    with col_q1:
        st.caption("Option A: Load from Maker Submission Depository Queue")
        dep_options = {
            f"[{m.get('status', 'Pending')}] ID: {m.get('id')} | {m.get('university')} ({m.get('author_email')})": m.get("id")
            for m in all_manuscripts
        }
        if dep_options:
            sel_dep = st.selectbox("Manuscripts Awaiting QA:", list(dep_options.keys()), key="chk_dep_sel")
            if st.button("📥 Load Selected Draft into Audit Workspace", key="chk_load_btn"):
                chosen = next((m for m in all_manuscripts if m["id"] == dep_options[sel_dep]), None)
                if chosen:
                    st.session_state.draft_text = chosen.get("draft_text", "")
                    st.session_state.loaded_manuscript_id = chosen["id"]
                    st.session_state.loaded_meta = chosen
                    st.success(f"Loaded manuscript ID: {chosen['id']}")
        else:
            st.info("No manuscripts awaiting review in depository.")

    with col_q2:
        if feat_flags.get("allow_direct_checker_upload", True):
            st.caption("Option B: Upload Files Directly (.DOCX, .PDF, .XLSX, .TXT)")
            up_files = st.file_uploader(
                "Upload Manuscript Files Directly:", 
                type=["docx", "pdf", "xlsx", "xls", "txt"], 
                accept_multiple_files=True,
                key="chk_up"
            )
            if up_files:
                parsed_blocks = [f"=== FILE: {f.name} ===\n{extract_content_multi_format(f)}" for f in up_files]
                st.session_state.draft_text = "\n\n".join(parsed_blocks)
                st.session_state.loaded_manuscript_id = f"UPLOADED-FILES-{len(up_files)}"
                st.session_state.loaded_meta = {"id": f"UPLOAD-{len(up_files)}-FILES", "author_email": user_email, "deliverable": selected_deliverable, "created_at": "Direct File Upload"}
                st.success(f"Ingested {len(up_files)} artifact(s).")
        else:
            st.info("Direct file upload is disabled by Administrator. Review drafts from the Maker Queue.")

    # ACTIVE MANUSCRIPT METADATA CARD (SHOWS WHAT IS BEING CHECKED)
    if st.session_state.draft_text.strip():
        meta = getattr(st.session_state, "loaded_meta", {}) or {}
        st.markdown("---")
        with st.container():
            st.markdown(f"""
            > **📌 Active Manuscript Under Review:** `{st.session_state.loaded_manuscript_id or 'Ad-hoc Manuscript'}`  
            > **Author / Source:** `{meta.get('author_email', 'Direct Upload')}` | **Format:** `{meta.get('deliverable', selected_deliverable)}` | **Submitted:** `{meta.get('created_at', 'N/A')}`  
            > **Payload Size:** `{len(st.session_state.draft_text.split())} words` | `{len(st.session_state.draft_text.splitlines())} lines`
            """)
            with st.expander("📄 Click to View / Inspect Extracted Manuscript Content"):
                st.text_area("Extracted Manuscript Text", value=st.session_state.draft_text, height=200, disabled=True, key="chk_preview_text")

        st.markdown("##### Step 2: Prescribe Syllabus for Granular Topic Mapping")
        col_s1, col_s2 = st.columns(2)
        with col_s1:
            syl_chk_file = st.file_uploader("Upload Approved Syllabus (.docx, .pdf, .txt)", type=["docx", "pdf", "txt"], key="chk_syl")
        with col_s2:
            syl_chk_txt = st.text_area("Or Paste Approved Syllabus Units & Topics", height=68, key="chk_syl_txt")
        chk_syllabus = extract_content_multi_format(syl_chk_file) if syl_chk_file else syl_chk_txt

        if st.button("⚡ Execute 12-Parameter QA & Syllabus Cross-Check", type="primary", key="chk_exec_btn"):
            with st.spinner("Auditing against 12 university parameters and syllabus..."):
                res = run_stage_2_checker(st.session_state.draft_text, chk_syllabus, "")
                st.session_state.stage2_results = res
                if st.session_state.loaded_manuscript_id and not str(st.session_state.loaded_manuscript_id).startswith("UPLOAD"):
                    update_manuscript_review(st.session_state.loaded_manuscript_id, user_email, stage2_results=res)
                st.success("Audit complete!")

        if st.session_state.stage2_results:
            st.markdown("---")
            st.markdown("##### Step 3: Audit Scorecard & Outcome Reports")
            s2 = st.session_state.stage2_results
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Quality Score", f"{s2.get('score', 0)}/100")
            c2.metric("Syllabus Coverage", f"{s2.get('coverage_pct', 0)}%")
            c3.metric("Critical Defects", s2.get("critical_count", 0))
            c4.metric("Verdict", s2.get("decision", "PASS"))

            t_syl_map, t_matrix, t_reports = st.tabs(["🎯 Syllabus vs Content Matrix", "📋 12-QA Evaluation Matrix & Defects", "📥 Download Reports (.XLSX & .DOCX)"])
            
            with t_syl_map:
                df_map = pd.DataFrame(s2.get("syllabus_mapping", []))
                if not df_map.empty:
                    st.dataframe(df_map, use_container_width=True)
                else:
                    st.info("No syllabus supplied; audited against internal structural continuity.")

            with t_matrix:
                st.markdown("##### 📋 12-Parameter Academic Quality Assurance Scorecard")
                matrix_data = generate_12_param_matrix(s2)
                st.dataframe(pd.DataFrame(matrix_data), use_container_width=True)

                st.markdown("##### 🚨 Defect & Remediation Register")
                df_err = pd.DataFrame(s2.get("error_register", []))
                if not df_err.empty:
                    st.dataframe(df_err, use_container_width=True)
                else:
                    st.success("✅ Zero critical or major defects identified against this manuscript.")

            with t_reports:
                st.markdown("###### Download Official Quality Dossiers")
                file_pfx = f"{selected_university}_{selected_program}_{st.session_state.loaded_manuscript_id or 'Manuscript'}".replace(" ", "_")
                xl_stream = build_stage2_qa_excel(s2, selected_university, selected_program, selected_deliverable, selected_batch, selected_term)
                docx_stream = build_stage2_qa_docx(s2, selected_university, selected_program, selected_deliverable, selected_batch, selected_term)
                
                col_d1, col_d2 = st.columns(2)
                col_d1.download_button(
                    "📊 Download 12-QA & Syllabus Workbook (.XLSX)", 
                    data=xl_stream, 
                    file_name=f"{file_pfx}_QA_Report.xlsx", 
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                    use_container_width=True
                )
                col_d2.download_button(
                    "📄 Download Stage 2 QA Review Report (.DOCX)", 
                    data=docx_stream, 
                    file_name=f"{file_pfx}_QA_Report.docx", 
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", 
                    use_container_width=True
                )
# --- FORENSIC AUDIT DELIVERABLES (EXCEL + WORD DOSSIER) ---
        st.markdown("---")
        st.markdown("### 📊 Forensic Quality Audit Deliverables")
        
        up_syl_file = st.file_uploader(
            "Upload Course Syllabus / Outline (Optional - .docx or .pdf):", 
            type=["docx", "pdf", "txt"], 
            key="stage2_syl_upload"
        )
        syl_content = extract_text_from_file(up_syl_file) if up_syl_file else ""

        target_format = st.session_state.get("selected_type", "SLM")
        target_topic = st.session_state.get("unit_title", "Academic_Unit")
        target_uni = st.session_state.get("selected_uni", "upGrad Academic Partner")

        if st.button(f"🚀 Run Forensic Quality Audit ({target_format})", type="primary"):
            with st.spinner(f"Evaluating {target_format} against academic rubrics and syllabus..."):
                audit_data = run_universal_academic_audit(
                    doc_text, 
                    target_format, 
                    syllabus_text=syl_content, 
                    engine="Google Gemini"
                )
                
                # 1. Multi-tab Excel workbook tailored to the deliverable
                if target_format == "MCQ":
                    excel_stream = generate_mcq_workbook(audit_data)
                elif target_format == "SLM":
                    excel_stream = generate_slm_workbook(audit_data)
                elif target_format == "QB":
                    excel_stream = generate_qb_workbook(audit_data)
                else:
                    excel_stream = generate_ppt_workbook(audit_data)
                    
                # 2. Executive Word dossier (.docx)
                docx_stream = build_docx_dossier_enhanced(
                    audit_data, 
                    target_uni, 
                    target_format, 
                    target_topic
                )
                
                st.success("✅ Forensic audit completed! Multi-tab workbook and executive dossier are ready.")
                
                col_d1, col_d2 = st.columns(2)
                with col_d1:
                    st.download_button(
                        label=f"📊 Download Forensic {target_format} QA Report (.xlsx)",
                        data=excel_stream,
                        file_name=f"{target_topic}_{target_format}_QA_Report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                with col_d2:
                    st.download_button(
                        label=f"📄 Download Executive QA Dossier (.docx) [Extra Report]",
                        data=docx_stream,
                        file_name=f"{target_topic}_{target_format}_Executive_Dossier.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )

# -------------------------------------------------------------
# VIEW 3: LEAD ADMINISTRATOR (EXECUTIVE EDITORIAL & CONTROL)
# -------------------------------------------------------------
else:
    st.header("👑 Executive Control & Editorial Merge Center")
    st.info("As Lead Administrator, you oversee both Maker submissions and Checker evaluations, with full editorial override authority.")

    t_merge, t_maker_prev, t_checker_prev = st.tabs([
        "🤝 Editorial Review & Merge Depository",
        "✍️ Maker Studio Preview",
        "🔍 Checker Console Preview"
    ])

    with t_merge:
        st.markdown("##### 📚 Manuscript Lifecycle & Accreditation Merge Queue")
        all_m = list_manuscripts()
        if all_m:
            df_overview = pd.DataFrame([{
                "ID": m.get("id"),
                "Institution": m.get("university"),
                "Program": m.get("program"),
                "Format": m.get("deliverable"),
                "Author": m.get("author_email"),
                "Status": m.get("status", "Pending Review"),
                "Score": (m.get("stage2_results") or {}).get("score", "Not Evaluated")
            } for m in all_m])
            st.dataframe(df_overview, use_container_width=True)

            sel_m_id = st.selectbox("Select Manuscript to Certify or Remediate:", [m.get("id") for m in all_m], key="adm_merge_sel")
            target_m = next((m for m in all_m if m.get("id") == sel_m_id), None)
            
            if target_m:
                st.markdown(f"**Author:** `{target_m.get('author_email')}` | **Status:** `{target_m.get('status')}` | **ID:** `{target_m.get('id')}`")
                col_m1, col_m2 = st.columns(2)
                with col_m1:
                    st.markdown("###### Maker Manuscript Text")
                    st.text_area("Content", value=target_m.get("draft_text", ""), height=220, disabled=True, key="adm_merge_txt")
                with col_m2:
                    st.markdown("###### Checker QA Evaluation")
                    st.json(target_m.get("stage2_results") or {"notice": "No audit performed yet"})

                c_appr1, c_appr2 = st.columns(2)
                with c_appr1:
                    if st.button("✅ Sign Off & Accredit Course Pack for Production", type="primary", use_container_width=True):
                        update_manuscript_review(sel_m_id, user_email, status="ACCREDITED FOR INSTRUCTION")
                        st.success(f"Manuscript {sel_m_id} certified and accredited!")
                        st.rerun()
                with c_appr2:
                    if st.button("🔄 Request SME Revisions from Maker Team", use_container_width=True):
                        update_manuscript_review(sel_m_id, user_email, status="REQUIRES SME REVISION")
                        st.warning(f"Manuscript {sel_m_id} marked for revision.")
                        st.rerun()
        else:
            st.info("No manuscripts currently in depository.")

    with t_maker_prev:
        st.markdown("##### ✍️ Admin Maker Studio (Draft & Deposit Directly)")
        col_am1, col_am2 = st.columns(2)
        with col_am1:
            adm_syl_file = st.file_uploader("Upload Syllabus (.docx, .pdf, .txt)", type=["docx", "pdf", "txt"], key="adm_maker_syl")
            adm_syl_txt = st.text_area("Or Paste Unit Syllabus Topics", height=90, key="adm_maker_syl_txt")
            adm_syl = extract_content_multi_format(adm_syl_file) if adm_syl_file else adm_syl_txt
        with col_am2:
            adm_ref_file = st.file_uploader("Upload Benchmark Textbook Reference", type=["docx", "pdf", "txt"], key="adm_maker_ref")
            adm_ref_txt = st.text_area("Or Paste Reference Excerpts", height=90, key="adm_maker_ref_txt")
            adm_ref = extract_content_multi_format(adm_ref_file) if adm_ref_file else adm_ref_txt

        if "admin_maker_draft_input" not in st.session_state:
            st.session_state["admin_maker_draft_input"] = st.session_state.get("draft_text", "")

        if st.button("🤖 Generate Draft Manuscript", type="primary", key="adm_maker_btn"):
            if adm_syl.strip():
                with st.spinner("Generating draft manuscript according to strict institutional guidelines..."):
                    enhanced_adm_ref = f"{adm_ref}\n\n{FORMATTING_GUIDELINES}"
                    gen_adm = run_stage_1_maker(adm_syl, enhanced_adm_ref, "Postgraduate", selected_deliverable)
                    if gen_adm and str(gen_adm).strip():
                        st.session_state["admin_maker_draft_input"] = str(gen_adm)
                        st.session_state.draft_text = str(gen_adm)
                        st.success("Draft Generated! Review below.")
                        st.rerun()
                    else:
                        st.error("AI Maker returned empty content. Verify GEMINI_API_KEY in Streamlit Secrets.")
            else:
                st.warning("Please supply a syllabus outline.")

        st.text_area("Manuscript Text / Editor", key="admin_maker_draft_input", height=240)
        st.session_state.draft_text = st.session_state.get("admin_maker_draft_input", "")

        col_sub1, col_sub2 = st.columns(2)
        with col_sub1:
            if st.button("📤 Submit Draft to Depository Queue", use_container_width=True, key="adm_sub_dep"):
                if st.session_state.draft_text.strip():
                    new_id = save_manuscript_draft(selected_university, selected_program, selected_batch, selected_deliverable, user_email, st.session_state.draft_text)
                    st.success(f"Manuscript submitted successfully as ID: {new_id}!")
                    st.rerun()
                else:
                    st.warning("Draft is empty.")
        with col_sub2:
            if st.session_state.draft_text.strip():
                profile = {"university": selected_university, "program": selected_program, "deliverable": selected_deliverable, "font": "Calibri", "body_size": 11.0, "primary_color": "002060", "accent_color": "C09A3E", "callout_boxes": True}
                st.download_button("📥 Download Branded Word (.DOCX)", data=build_branded_manuscript_docx(st.session_state.draft_text, profile), file_name=f"{selected_university}_Draft.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)

    with t_checker_prev:
        st.markdown("##### 🔍 Admin Direct Audit & Review")
        col_ac1, col_ac2 = st.columns(2)
        with col_ac1:
            adm_all_m = list_manuscripts()
            adm_dep_opts = {f"[{m.get('status', 'Pending')}] ID: {m.get('id')} ({m.get('author_email')})": m.get("id") for m in adm_all_m}
            if adm_dep_opts:
                adm_sel_dep = st.selectbox("Or Load from Depository Queue:", list(adm_dep_opts.keys()), key="adm_chk_dep_sel")
                if st.button("📥 Load Depository Draft into Audit Workspace", key="adm_load_dep_btn"):
                    c_draft = next((m for m in adm_all_m if m["id"] == adm_dep_opts[adm_sel_dep]), None)
                    if c_draft:
                        st.session_state.draft_text = c_draft.get("draft_text", "")
                        st.session_state.loaded_manuscript_id = c_draft["id"]
                        st.session_state.loaded_meta = c_draft
                        st.success(f"Loaded manuscript: {c_draft['id']}")

            adm_chk_files = st.file_uploader("Or Upload File(s) Directly (.DOCX, .PDF, .XLSX, .TXT):", type=["docx", "pdf", "xlsx", "xls", "txt"], accept_multiple_files=True, key="adm_chk_up")
            if adm_chk_files:
                st.session_state.draft_text = "\n\n".join([f"=== FILE: {f.name} ===\n{extract_content_multi_format(f)}" for f in adm_chk_files])
                st.session_state.loaded_manuscript_id = f"UPLOADED-FILES-{len(adm_chk_files)}"
                st.session_state.loaded_meta = {"id": f"UPLOAD-{len(adm_chk_files)}-FILES", "author_email": user_email, "deliverable": selected_deliverable, "created_at": "Direct File Upload"}
                st.success(f"Loaded {len(adm_chk_files)} file(s).")

        with col_ac2:
            adm_csyl_file = st.file_uploader("Upload Approved Syllabus for Cross-Check:", type=["docx", "pdf", "txt"], key="adm_chk_syl")
            adm_csyl_txt = st.text_area("Or Paste Approved Syllabus", height=68, key="adm_chk_syl_txt")
            adm_csyl = extract_content_multi_format(adm_csyl_file) if adm_csyl_file else adm_csyl_txt

        if st.session_state.draft_text.strip():
            meta_adm = getattr(st.session_state, "loaded_meta", {}) or {}
            st.markdown(f"""
            > **📌 Active Manuscript Being Checked:** `{st.session_state.loaded_manuscript_id or 'Ad-hoc Draft'}`  
            > **Author:** `{meta_adm.get('author_email', 'Direct Upload')}` | **Format:** `{meta_adm.get('deliverable', selected_deliverable)}`  
            > **Payload Size:** `{len(st.session_state.draft_text.split())} words`
            """)

            if st.button("⚡ Run Stage 2 Academic QA & Syllabus Cross-Check", type="primary", key="adm_chk_run"):
                with st.spinner("Running Stage 2 Academic QA & Syllabus Cross-Check..."):
                    # 1. Format resolution
                    selected_type = st.session_state.get("loaded_meta", {}).get("format", "SLM")
                    adm_syl = adm_csyl if "adm_csyl" in locals() else ""
                    
                    # 2. Comprehensive text payload resolver
                    real_text = st.session_state.get("draft_text", "")
                    if not real_text or len(str(real_text).strip()) < 50:
                        meta = st.session_state.get("loaded_meta", {})
                        for k in ["content", "text", "draft", "raw_text", "body", "payload"]:
                            if meta.get(k) and len(str(meta[k]).strip()) > 50:
                                real_text = meta[k]
                                break
                    if not real_text or len(str(real_text).strip()) < 50:
                        for var_name in ["doc_text", "uploaded_text", "manuscript_text", "active_text", "raw_text"]:
                            if var_name in locals() and locals()[var_name] and len(str(locals()[var_name]).strip()) > 50:
                                real_text = locals()[var_name]
                                break
                            if var_name in st.session_state and st.session_state[var_name] and len(str(st.session_state[var_name]).strip()) > 50:
                                real_text = st.session_state[var_name]
                                break
                    
                    # Synchronize session state so the Multi-Tab report generator also receives it
                    st.session_state["draft_text"] = real_text
                    
                    res = audit_data = run_universal_academic_audit(
                        doc_text=real_text,
                        deliverable=selected_type,
                        syllabus_text=adm_syl,
                        engine="Google Gemini"
                    )
                    st.session_state.stage2_results = res
                    st.success("Stage 2 QA Audit complete!")
                    st.session_state.stage2_results = res
                    st.success("Stage 2 QA Audit complete!")
                    if st.session_state.loaded_manuscript_id and not str(st.session_state.loaded_manuscript_id).startswith("UPLOAD"):
                        update_manuscript_review(st.session_state.loaded_manuscript_id, user_email, stage2_results=res)
                    st.success("Audit complete!")

        if st.session_state.stage2_results:
            s2 = st.session_state.stage2_results
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Score", f"{s2.get('score', 0)}/100")
            c2.metric("Coverage", f"{s2.get('coverage_pct', 0)}%")
            c3.metric("Critical Defects", s2.get("critical_count", 0))
            c4.metric("Verdict", s2.get("decision", "PASS"))

            t_adm_syl, t_adm_mat = st.tabs(["🎯 Syllabus vs Content Matrix", "📋 12-QA Evaluation Matrix & Defects"])
            with t_adm_syl:
                df_map = pd.DataFrame(s2.get("syllabus_mapping", []))
                if not df_map.empty:
                    st.dataframe(df_map, use_container_width=True)
                else:
                    st.info("No syllabus supplied; audited against internal structural continuity.")
            with t_adm_mat:
                st.markdown("###### 12-Parameter Academic Quality Assurance Scorecard")
                st.dataframe(pd.DataFrame(generate_12_param_matrix(s2)), use_container_width=True)
                st.markdown("###### Logged Defects & Recommendations")
                df_adm_err = pd.DataFrame(s2.get("error_register", []))
                if not df_adm_err.empty:
                    st.dataframe(df_adm_err, use_container_width=True)
                else:
                    st.success("✅ Zero critical or major defects identified.")

            st.markdown("###### Download Audit Outcome Dossiers")
            file_pfx = f"{selected_university}_{selected_program}_{st.session_state.loaded_manuscript_id or 'Manuscript'}".replace(" ", "_")
            # Route download to authentic multi-tab generator
            if selected_deliverable == "MCQ":
                xl = generate_mcq_workbook(s2)
            elif selected_deliverable == "SLM":
                xl = generate_slm_workbook(s2)
            elif selected_deliverable == "QB":
                xl = generate_qb_workbook(s2)
            else:
                xl = generate_ppt_workbook(s2)
            docx_audit = build_stage2_qa_docx(s2, selected_university, selected_program, selected_deliverable, selected_batch, selected_term)
            
            c_d1, c_d2 = st.columns(2)
            c_d1.download_button("📊 Download Audit Workbook (.XLSX)", data=xl, file_name=f"{file_pfx}_QA_Report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            c_d2.download_button("📄 Download QA Review Report (.DOCX)", data=docx_audit, file_name=f"{file_pfx}_QA_Report.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
           # =====================================================================
        # NEW FEATURE: MULTI-TAB FORENSIC EXCEL & WORD DOSSIER
        # =====================================================================
        st.markdown("---")
        st.markdown("### 📊 Generate New Multi-Tab QA Reports (CN Format)")

        user_key = st.text_input("🔑 Enter Gemini API Key (Optional Override):", type="password", help="Paste your AIzaSy... key here if Cloud Secrets fail")
        if user_key.strip():
            os.environ["GEMINI_API_KEY"] = user_key.strip()
        
        if st.button("🚀 Generate Multi-Tab Excel & Word Dossier", type="primary"):
            try:
                with st.spinner("Generating new forensic formats..."):
                    
                    # 1. Dynamically resolve the manuscript text to fix empty inputs
                    current_text = ""
                    if st.session_state.get("draft_text") and str(st.session_state["draft_text"]).strip():
                        current_text = str(st.session_state["draft_text"]).strip()
                    elif st.session_state.get("manuscript_text") and str(st.session_state["manuscript_text"]).strip():
                        current_text = str(st.session_state["manuscript_text"]).strip()
                    if "manuscript_text" in st.session_state and st.session_state["manuscript_text"]:
                        current_text = st.session_state["manuscript_text"]
                    elif "doc_text" in st.session_state and st.session_state["doc_text"]:
                        current_text = st.session_state["doc_text"]
                    elif "doc_text" in locals() and doc_text:
                        current_text = doc_text
                    elif "uploaded_manuscript" in locals() and uploaded_manuscript:
                        current_text = extract_text_from_file(uploaded_manuscript)
                    elif "uploaded_file" in locals() and uploaded_file:
                        current_text = extract_text_from_file(uploaded_file)
                        
                    if not current_text or len(current_text.strip()) < 50:
                        st.error("⚠️ No manuscript text detected. Please upload a valid document first.")
                        st.stop()

                    # 2. Safely grab formats and names
                    current_type = selected_type if "selected_type" in locals() else "SLM"
                    
                    try:
                        raw_name = uploaded_file.name if "uploaded_file" in locals() else uploaded_manuscript.name
                        raw_name = raw_name.rsplit(".", 1)[0]
                    except Exception:
                        raw_name = "Academic_Module"
                        
                    words = [w for w in raw_name.replace("_", " ").split() if w[0].isalnum()]
                    code_prefix = "".join([w[0].upper() for w in words]) if words else "MODULE"
                    
                    # 3. Call Universal Engine
                    audit_data = run_universal_academic_audit(
                        current_text, 
                        current_type, 
                        syllabus_text="", 
                        engine="Google Gemini"
                    )
                    
                    # 4. Generate Reports
                    if current_type == "MCQ":
                        excel_stream = generate_mcq_workbook(audit_data)
                    elif current_type == "SLM":
                        excel_stream = generate_slm_workbook(audit_data)
                    elif current_type == "QB":
                        excel_stream = generate_qb_workbook(audit_data)
                    else:
                        excel_stream = generate_ppt_workbook(audit_data)
                        
                    docx_stream = build_docx_dossier_enhanced(
                        audit_data, 
                        "upGrad Academic Partner", 
                        current_type, 
                        raw_name
                    )
                    
                    out_xlsx_name = f"{code_prefix}_{current_type}_QA_Report.xlsx"
                    out_docx_name = f"{code_prefix}_{current_type}_Executive_Dossier.docx"
                    
                    st.success(f"✅ Successfully generated {out_xlsx_name}!")
                    
                    # 5. Display Downloads
                    col_dl1, col_dl2 = st.columns(2)
                    with col_dl1:
                        st.download_button(
                            label=f"📊 Download {out_xlsx_name}", 
                            data=excel_stream, 
                            file_name=out_xlsx_name, 
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    with col_dl2:
                        st.download_button(
                            label=f"📄 Download {out_docx_name}", 
                            data=docx_stream, 
                            file_name=out_docx_name, 
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
            except Exception as e:
                st.error(f"❌ Report Generation Error: {str(e)}")
        # =====================================================================
