import io
import json
import os
import zipfile
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

try:
    from pypdf import PdfReader
except ImportError:
    PdfReader = None

# Modular Project Services
from services.access_manager import load_access_config, verify_user
from services.unified_engine import run_stage_1_maker, run_stage_2_checker, run_stage_3_prompt_audit
try:
    from services.depository_manager import list_manuscripts, update_manuscript_review, save_manuscript_draft
except ImportError:
    def list_manuscripts(): return []
    def update_manuscript_review(*args, **kwargs): pass
    def save_manuscript_draft(*args, **kwargs): pass

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
            with open(path, "r") as f:
                return json.load(f)
        except Exception:
            pass
    return {}

def update_system_config(mutator_fn):
    path = get_config_file_path()
    cfg = load_system_config()
    cfg = mutator_fn(cfg)
    try:
        with open(path, "w") as f:
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

def save_feature_flags(flags: dict):
    def fn(cfg):
        feat = cfg.setdefault("features", {})
        feat.update(flags)
        return cfg
    update_system_config(fn)

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

        errs = stage2_data.get("error_register", [])
        matrix_rows = []
        for p in QA_12_PARAMETERS:
            related = [e for e in errs if p["name"].lower() in str(e).lower() or str(e.get("category", "")).lower() in p["name"].lower()]
            achieved = calculate_realistic_score(p["weight"], related)
            pct = round((achieved / p["weight"]) * 100, 1)
            status = "REVISE" if any(str(e.get("severity", "")).lower() == "critical" for e in related) else ("OBSERVATION" if related else "COMPLIANT")
            findings = "; ".join([str(e.get("description", "Observation")) for e in related]) if related else "Meets quality standards."
            remediation = "; ".join([str(e.get("recommendation", "None")) for e in related]) if related else "None required."
            matrix_rows.append({
                "Param #": p["id"],
                "QA Dimension": p["name"],
                "Weight": p["weight"],
                "Achieved": achieved,
                "Success (%)": f"{pct}%",
                "Status": status,
                "Auditor Findings": findings,
                "Remediation Guidance": remediation
            })
        pd.DataFrame(matrix_rows).to_excel(writer, sheet_name="12 QA Parameter Matrix", index=False)
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

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out

# -------------------------------------------------------------
# SESSION STATE & AUTHENTICATION
# -------------------------------------------------------------
sys_cfg = load_system_config()
feat_flags = sys_cfg.get("features", {
    "maker_studio_enabled": True,
    "branding_mapper_enabled": True,
    "allow_direct_checker_upload": True
})

if "draft_text" not in st.session_state: st.session_state.draft_text = ""
if "stage2_results" not in st.session_state: st.session_state.stage2_results = None
if "stage3_results" not in st.session_state: st.session_state.stage3_results = None
if "loaded_manuscript_id" not in st.session_state: st.session_state.loaded_manuscript_id = None
if "active_batch" not in st.session_state: st.session_state.active_batch = "Batch 2026-A"
if "active_term" not in st.session_state: st.session_state.active_term = "Semester 1"

st.sidebar.header("🔐 Workspace Login")
user_email = st.sidebar.text_input(
    "Enter Official Email ID (@upgrad.com)", 
    value="sandeep1.shrivastava@upgrad.com"
).strip().lower()

if not user_email.endswith("@upgrad.com"):
    st.sidebar.error("⛔ Access Denied: Only official '@upgrad.com' accounts permitted.")
    st.stop()

is_authorized, user_role = verify_user(user_email)
if not is_authorized:
    st.sidebar.error("⛔ Access Denied: Unauthorized Account.")
    st.stop()

st.sidebar.success(f"Verified Role: **{user_role}**")

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

selected_program = st.sidebar.selectbox("2. Academic Programme", DEFAULT_PROGRAMS, index=0)
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

        if st.button("🤖 Generate Draft Manuscript (Stage 1 Maker)", type="primary"):
            if m_syllabus.strip():
                with st.spinner("AI Maker generating academic manuscript..."):
                    st.session_state.draft_text = run_stage_1_maker(m_syllabus, m_reference, "Postgraduate", selected_deliverable)
                    st.success("Draft Generated! Review and refine below.")
            else:
                st.warning("Please supply a syllabus outline.")

        st.session_state.draft_text = st.text_area(
            "Manuscript Editor (Markdown Supported)", 
            value=st.session_state.draft_text, 
            height=320, 
            key="m_editor"
        )

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
    st.markdown("##### Step 1: Ingest Content for Evaluation")
    
    col_q1, col_q2 = st.columns(2)
    with col_q1:
        st.caption("Option A: Load from Maker Submission Queue")
        dep_options = {
            f"[{m.get('status', 'Pending')}] {m.get('university')} ➔ {m.get('deliverable')} (Author: {m.get('author_email')})": m.get("id")
            for m in all_manuscripts
        }
        if dep_options:
            sel_dep = st.selectbox("Manuscripts Ready for QA:", list(dep_options.keys()))
            if st.button("📥 Load Selected Draft into Audit Workspace"):
                chosen = next((m for m in all_manuscripts if m["id"] == dep_options[sel_dep]), None)
                if chosen:
                    st.session_state.draft_text = chosen.get("draft_text", "")
                    st.session_state.loaded_manuscript_id = chosen["id"]
                    st.success("Draft loaded from Maker queue.")
        else:
            st.info("No manuscripts awaiting review in depository.")

    with col_q2:
        if feat_flags.get("allow_direct_checker_upload", True):
            st.caption("Option B: Upload Multi-Format Files Directly")
            up_files = st.file_uploader(
                "Upload .DOCX, .PDF, .XLSX files directly:", 
                type=["docx", "pdf", "xlsx", "xls", "txt"], 
                accept_multiple_files=True,
                key="chk_up"
            )
            if up_files:
                parsed_blocks = [f"=== FILE: {f.name} ===\n{extract_content_multi_format(f)}" for f in up_files]
                st.session_state.draft_text = "\n\n".join(parsed_blocks)
                st.session_state.loaded_manuscript_id = None
                st.success(f"Ingested {len(up_files)} artifact(s).")
        else:
            st.info("Direct file upload is disabled. Please review manuscripts from the Maker Queue.")

    if st.session_state.draft_text.strip():
        st.markdown("---")
        st.markdown("##### Step 2: Prescribe Syllabus for Granular Topic Mapping")
        syl_chk_file = st.file_uploader("Upload Approved Syllabus (.docx, .pdf, .txt)", type=["docx", "pdf", "txt"], key="chk_syl")
        syl_chk_txt = st.text_area("Or Paste Approved Syllabus Units & Topics", height=90, key="chk_syl_txt")
        chk_syllabus = extract_content_multi_format(syl_chk_file) if syl_chk_file else syl_chk_txt

        if st.button("⚡ Execute 12-Parameter QA & Syllabus Cross-Check", type="primary"):
            with st.spinner("Auditing against university parameters and syllabus..."):
                res = run_stage_2_checker(st.session_state.draft_text, chk_syllabus, "")
                st.session_state.stage2_results = res
                if st.session_state.loaded_manuscript_id:
                    update_manuscript_review(st.session_state.loaded_manuscript_id, user_email, stage2_results=res)
                st.success("Audit complete!")

        if st.session_state.stage2_results:
            st.markdown("---")
            st.markdown("##### Step 3: Audit Scorecard & Reports")
            s2 = st.session_state.stage2_results
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Quality Score", f"{s2.get('score', 0)}/100")
            c2.metric("Syllabus Coverage", f"{s2.get('coverage_pct', 0)}%")
            c3.metric("Critical Defects", s2.get("critical_count", 0))
            c4.metric("Verdict", s2.get("decision", "PASS"))

            t_syl_map, t_matrix, t_reports = st.tabs(["🎯 Syllabus vs Content Matrix", "📋 12-QA Matrix & Defects", "📥 Download Reports"])
            
            with t_syl_map:
                df_map = pd.DataFrame(s2.get("syllabus_mapping", []))
                if not df_map.empty:
                    st.dataframe(df_map, use_container_width=True)
                else:
                    st.info("No syllabus supplied; audited against internal structural continuity.")

            with t_matrix:
                st.dataframe(pd.DataFrame(s2.get("error_register", [])), use_container_width=True)

            with t_reports:
                file_pfx = f"{selected_university}_{selected_program}_{selected_batch}".replace(" ", "_")
                xl_stream = build_stage2_qa_excel(s2, selected_university, selected_program, selected_deliverable, selected_batch, selected_term)
                docx_stream = build_stage2_qa_docx(s2, selected_university, selected_program, selected_deliverable, selected_batch, selected_term)
                
                col_d1, col_d2 = st.columns(2)
                col_d1.download_button("📊 Download 12-QA & Syllabus Workbook (.XLSX)", data=xl_stream, file_name=f"{file_pfx}_QA_Report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                col_d2.download_button("📄 Download Stage 2 QA Review Report (.DOCX)", data=docx_stream, file_name=f"{file_pfx}_QA_Report.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)

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
                "Score": m.get("stage2_results", {}).get("score", "Not Evaluated")
            } for m in all_m])
            st.dataframe(df_overview, use_container_width=True)

            sel_m_id = st.selectbox("Select Manuscript to Certify or Remediate:", [m.get("id") for m in all_m])
            target_m = next((m for m in all_m if m.get("id") == sel_m_id), None)
            
            if target_m:
                st.markdown(f"**Author:** {target_m.get('author_email')} | **Status:** {target_m.get('status')}")
                col_m1, col_m2 = st.columns(2)
                with col_m1:
                    st.markdown("###### Maker Manuscript Text")
                    st.text_area("Content", value=target_m.get("draft_text", ""), height=220, disabled=True)
                with col_m2:
                    st.markdown("###### Checker QA Evaluation")
                    st.json(target_m.get("stage2_results", {"notice": "No audit performed yet"}))

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

        if st.button("🤖 Generate Draft Manuscript", type="primary", key="adm_maker_btn"):
            if adm_syl.strip():
                with st.spinner("Generating draft manuscript..."):
                    st.session_state.draft_text = run_stage_1_maker(adm_syl, adm_ref, "Postgraduate", selected_deliverable)
                    st.success("Draft Generated! Review below.")
            else:
                st.warning("Please supply a syllabus outline.")

        st.session_state.draft_text = st.text_area("Manuscript Text / Editor", value=st.session_state.draft_text, height=220, key="adm_maker_editor")

        col_sub1, col_sub2 = st.columns(2)
        with col_sub1:
            if st.button("📤 Submit Draft to Depository Queue", use_container_width=True):
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
            adm_chk_files = st.file_uploader("Upload Manuscript Files (.DOCX, .PDF, .XLSX, .TXT):", type=["docx", "pdf", "xlsx", "xls", "txt"], accept_multiple_files=True, key="adm_chk_up")
            if adm_chk_files:
                st.session_state.draft_text = "\n\n".join([f"=== FILE: {f.name} ===\n{extract_content_multi_format(f)}" for f in adm_chk_files])
                st.success(f"Loaded {len(adm_chk_files)} file(s).")
        with col_ac2:
            adm_csyl_file = st.file_uploader("Upload Approved Syllabus for Cross-Check:", type=["docx", "pdf", "txt"], key="adm_chk_syl")
            adm_csyl_txt = st.text_area("Or Paste Approved Syllabus", height=68, key="adm_chk_syl_txt")
            adm_csyl = extract_content_multi_format(adm_csyl_file) if adm_csyl_file else adm_csyl_txt

        if st.session_state.draft_text.strip():
            if st.button("⚡ Run Stage 2 Academic QA & Syllabus Cross-Check", type="primary", key="adm_chk_run"):
                with st.spinner("Auditing against 12 parameters and syllabus..."):
                    res = run_stage_2_checker(st.session_state.draft_text, adm_csyl, "")
                    st.session_state.stage2_results = res
                    st.success("Audit complete!")

        if st.session_state.stage2_results:
            s2 = st.session_state.stage2_results
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Score", f"{s2.get('score', 0)}/100")
            c2.metric("Coverage", f"{s2.get('coverage_pct', 0)}%")
            c3.metric("Critical Defects", s2.get("critical_count", 0))
            c4.metric("Verdict", s2.get("decision", "PASS"))
            xl = build_stage2_qa_excel(s2, selected_university, selected_program, selected_deliverable, selected_batch, selected_term)
            st.download_button("📊 Download Audit Workbook (.XLSX)", data=xl, file_name="QA_Audit_Report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")