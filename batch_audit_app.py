import streamlit as st
import io
import re
import os
import openpyxl
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

# Reuse your existing Stage 2 AI Audit & Docx generation engines
from services.unified_engine import run_stage_2_checker
from services.depository_manager import get_db

# -----------------------------------------------------------------------------
# 1. PAGE SETUP & SECURITY
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Academic Batch Drive Auditor | upGrad",
    page_icon="📂",
    layout="wide"
)

# Bridge Secrets to Environment
if "GEMINI_API_KEY" in st.secrets:
    os.environ["GEMINI_API_KEY"] = str(st.secrets["GEMINI_API_KEY"])
if "OPENAI_API_KEY" in st.secrets:
    os.environ["OPENAI_API_KEY"] = str(st.secrets["OPENAI_API_KEY"])

st.title("📂 Academic Batch Drive Auditor & Pipeline Extension")
st.caption("Automated ingestion, multi-sheet evaluation (SLM, MCQ, QB, PPT), and dynamic folder hierarchy creation.")

# -----------------------------------------------------------------------------
# 2. GOOGLE DRIVE ENGINE
# -----------------------------------------------------------------------------
@st.cache_resource
def get_drive_service():
    """Initializes Google Drive API using gcp_service_account secret."""
    if "gcp_service_account" not in st.secrets:
        st.error("Missing 'gcp_service_account' in Streamlit Secrets.")
        return None
    key_dict = dict(st.secrets["gcp_service_account"])
    scopes = ["https://www.googleapis.com/auth/drive"]
    creds = service_account.Credentials.from_service_account_info(key_dict, scopes=scopes)
    return build("drive", "v3", credentials=creds)

def extract_drive_id(link_or_id: str) -> str:
    """Extracts File or Folder ID from any Google Drive link."""
    if not link_or_id:
        return ""
    m_file = re.search(r"/d/([a-zA-Z0-9-_]+)", str(link_or_id))
    if m_file:
        return m_file.group(1)
    m_folder = re.search(r"/folders/([a-zA-Z0-9-_]+)", str(link_or_id))
    if m_folder:
        return m_folder.group(1)
    return str(link_or_id).strip()

def download_drive_stream(service, file_id: str) -> io.BytesIO:
    """Downloads a file stream safely into memory, auto-handling native Google Sheets and Docs."""
    meta = service.files().get(
        fileId=file_id, 
        fields="mimeType, name", 
        supportsAllDrives=True
    ).execute()
    mime_type = meta.get("mimeType", "")

    if mime_type == "application/vnd.google-apps.spreadsheet":
        exported_bytes = service.files().export_media(
            fileId=file_id,
            mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ).execute()
        stream = io.BytesIO(exported_bytes)
        stream.seek(0)
        return stream

    elif mime_type == "application/vnd.google-apps.document":
        exported_bytes = service.files().export_media(
            fileId=file_id,
            mimeType="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        ).execute()
        stream = io.BytesIO(exported_bytes)
        stream.seek(0)
        return stream

    else:
        try:
            raw_bytes = service.files().get_media(fileId=file_id, supportsAllDrives=True).execute()
            stream = io.BytesIO(raw_bytes)
        except Exception:
            request = service.files().get_media(fileId=file_id, supportsAllDrives=True)
            stream = io.BytesIO()
            downloader = MediaIoBaseDownload(stream, request)
            done = False
            while not done:
                _, done = downloader.next_chunk()
        stream.seek(0)
        return stream

def get_or_create_subfolder(service, parent_id: str, folder_name: str) -> str:
    """Finds or creates a folder under parent_id without creating duplicates."""
    safe_name = folder_name.replace("'", "\\'")
    query = f"'{parent_id}' in parents and name = '{safe_name}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
    res = service.files().list(q=query, spaces="drive", fields="files(id, name)").execute()
    files = res.get("files", [])
    if files:
        return files[0]["id"]
    
    meta = {
        "name": folder_name,
        "mimeType": "application/vnd.google-apps.folder",
        "parents": [parent_id]
    }
    created = service.files().create(body=meta, fields="id").execute()
    return created.get("id")

def upload_report_to_hierarchy(service, docx_stream, file_name: str, root_id: str, university: str, program: str, deliverable: str) -> str:
    """Creates Root -> University -> Program -> [Deliverable] Reports hierarchy and uploads report."""
    uni_name = university.strip() if university and university.strip() else "General University"
    uni_folder_id = get_or_create_subfolder(service, root_id, uni_name)
    
    prog_name = program.strip() if program and program.strip() else "General Programme"
    prog_folder_id = get_or_create_subfolder(service, uni_folder_id, prog_name)
    
    deliv_name = deliverable.strip().upper() if deliverable and deliverable.strip() else "SLM"
    subfolder_name = f"{deliv_name} Reports"
    target_folder_id = get_or_create_subfolder(service, prog_folder_id, subfolder_name)

    media = MediaIoBaseUpload(
        docx_stream,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        resumable=True
    )
    uploaded = service.files().create(
        body={"name": file_name, "parents": [target_folder_id]},
        media_body=media,
        fields="id, webViewLink"
    ).execute()
    return uploaded.get("webViewLink", "")

# -----------------------------------------------------------------------------
# 3. TEXT EXTRACTOR & REPORT COMPILER
# -----------------------------------------------------------------------------
def extract_text_from_stream(stream: io.BytesIO) -> str:
    """Extracts text content from .docx or .pdf streams."""
    try:
        import docx
        doc = docx.Document(stream)
        return "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
    except Exception:
        try:
            import pypdf
            stream.seek(0)
            reader = pypdf.PdfReader(stream)
            return "\n".join([page.extract_text() or "" for page in reader.pages])
        except Exception:
            return ""

def build_docx_dossier(qa_results: dict, university: str, deliverable: str, topic: str) -> io.BytesIO:
    """Builds a formatted Word QA dossier in memory."""
    import docx
    from docx.shared import Inches, Pt, RGBColor
    
    doc = docx.Document()
    title = doc.add_heading("ACADEMIC QUALITY AUDIT REPORT", level=0)
    title.runs[0].font.color.rgb = RGBColor(192, 0, 0)
    
    doc.add_paragraph(f"Institution: {university} | Deliverable: {deliverable} | Unit/Topic: {topic}")
    doc.add_paragraph("=" * 60)
    
    if isinstance(qa_results, dict):
        score = qa_results.get("overall_score", qa_results.get("quality_score", "N/A"))
        doc.add_heading(f"Overall Audit Score: {score}/100", level=2)
        doc.add_paragraph(f"Summary / Verdict: {qa_results.get('executive_summary', 'Evaluated against compliance rubrics.')}")
        
        matrix = qa_results.get("parameters", qa_results.get("parameter_scores", {}))
        if matrix:
            doc.add_heading("Detailed Parameter Compliance Matrix", level=3)
            table = doc.add_table(rows=1, cols=3)
            hdr_cells = table.rows[0].cells
            hdr_cells[0].text = "Parameter / Dimension"
            hdr_cells[1].text = "Score"
            hdr_cells[2].text = "Findings & Recommendations"
            
            for param, details in matrix.items():
                row_cells = table.add_row().cells
                row_cells[0].text = str(param)
                if isinstance(details, dict):
                    row_cells[1].text = str(details.get("score", details.get("status", "Evaluated")))
                    row_cells[2].text = str(details.get("comments", details.get("feedback", "")))
                else:
                    row_cells[1].text = "Passed"
                    row_cells[2].text = str(details)
    else:
        doc.add_paragraph(str(qa_results))
        
    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out

def detect_deliverable_type(topic_name: str, doc_text: str, explicit_type: str = "") -> str:
    """Dynamically identifies SLM, MCQ, QB, or PPT via metadata and structural text fingerprinting."""
    if explicit_type and str(explicit_type).strip().upper() in ["SLM", "MCQ", "QB", "PPT"]:
        return str(explicit_type).strip().upper()

    metadata = str(topic_name).lower()
    if any(k in metadata for k in ["mcq", "quiz", "objective"]): return "MCQ"
    if any(k in metadata for k in ["question bank", "qb", "q&a"]): return "QB"
    if any(k in metadata for k in ["ppt", "deck", "presentation"]): return "PPT"
    if any(k in metadata for k in ["slm", "unit", "module"]): return "SLM"

    text_snippet = doc_text[:5000]
    if len(re.findall(r"(\([A-Da-d]\)|\bOption\s+[A-D]\b)", text_snippet)) >= 8:
        return "MCQ"
    if len(re.findall(r"(\[\d+\s*Marks?\]|\(\d+\s*Marks?\)|\b\d+\s*Marks\b)", text_snippet, re.IGNORECASE)) >= 3:
        return "QB"
    if len(re.findall(r"(Slide\s+\d+|Presentation:)", text_snippet, re.IGNORECASE)) >= 3:
        return "PPT"

    return "SLM"

# -----------------------------------------------------------------------------
# 4. SMART SHEET INSPECTOR & COLUMN RESOLVER
# -----------------------------------------------------------------------------
def analyze_sheet_structure(ws):
    """Inspects sheet rows to automatically detect header row and column mappings."""
    header_row_idx = 1
    headers = {}
    
    # Check rows 1 to 5 for the most likely header row
    for r in range(1, min(6, ws.max_row + 1)):
        row_vals = [cell.value for cell in ws[r] if cell.value is not None]
        str_count = sum(1 for v in row_vals if isinstance(v, str) and len(v.strip()) > 1)
        if str_count >= 2:
            header_row_idx = r
            for col_idx, cell in enumerate(ws[r], start=1):
                if cell.value is not None:
                    headers[str(cell.value).strip()] = col_idx
            break

    if not headers:
        header_row_idx = 1
        for col_idx, cell in enumerate(ws[1], start=1):
            if cell.value is not None:
                headers[str(cell.value).strip()] = col_idx

    detected_link = None
    detected_uni = None
    detected_prog = None
    detected_topic = None
    detected_deliv = None

    for col_name, idx in headers.items():
        low = col_name.lower().strip()
        if not detected_link and any(k in low for k in ["link", "url", "drive link", "doc link", "file link", "drive"]):
            detected_link = col_name
        if not detected_uni and any(k in low for k in ["university", "univ", "institution", "college"]):
            detected_uni = col_name
        if not detected_prog and any(k in low for k in ["programme", "program", "course", "degree"]):
            detected_prog = col_name
        if not detected_topic and any(k in low for k in ["topic", "unit", "module", "chapter", "title", "subject"]):
            detected_topic = col_name
        if not detected_deliv and any(k in low for k in ["deliverable", "type", "format", "asset"]):
            detected_deliv = col_name

    # Content-based link scanning fallback
    if not detected_link:
        for col_name, col_idx in headers.items():
            for r in range(header_row_idx + 1, min(header_row_idx + 10, ws.max_row + 1)):
                val = str(ws.cell(row=r, column=col_idx).value or "")
                if "drive.google.com" in val or "docs.google.com" in val or "http" in val:
                    detected_link = col_name
                    break
            if detected_link:
                break

    return header_row_idx, headers, detected_link, detected_uni, detected_prog, detected_topic, detected_deliv

# -----------------------------------------------------------------------------
# 5. USER INTERFACE & WORKFLOW
# -----------------------------------------------------------------------------
st.sidebar.header("⚙️ Configuration")
root_folder_input = st.sidebar.text_input(
    "Root QA Reports Folder ID / Link:",
    help="Paste the ID or full link of your 'Master QA Reports Directory'"
)
clean_root_id = extract_drive_id(root_folder_input)

audit_engine = st.sidebar.radio(
    "AI Evaluation Engine:",
    ["Google Gemini", "OpenAI GPT-4o"],
    index=0
)

st.markdown("### Step 1: Connect Master Excel Workbook")
excel_link_input = st.text_input(
    "Paste Google Drive Link to Master Tracker Excel / Google Sheet:",
    placeholder="https://docs.google.com/spreadsheets/d/1AbC.../edit"
)
clean_excel_id = extract_drive_id(excel_link_input)

drive_service = get_drive_service()

if clean_excel_id and drive_service:
    if st.button("🔍 Scan & Inspect University Sheets"):
        with st.spinner("Downloading workbook into memory and scanning sheets..."):
            try:
                xl_stream = download_drive_stream(drive_service, clean_excel_id)
                wb_preview = openpyxl.load_workbook(xl_stream, read_only=True)
                st.session_state["sheet_names"] = wb_preview.sheetnames
                wb_preview.close()
                xl_stream.seek(0)
                st.session_state["raw_excel_bytes"] = xl_stream.getvalue()
                st.success(f"Workbook inspected! Found {len(st.session_state['sheet_names'])} sheets: {', '.join(st.session_state['sheet_names'][:5])}...")
            except Exception as e:
                st.error(f"Failed to access workbook: {e}")

if "sheet_names" in st.session_state and st.session_state["sheet_names"]:
    st.divider()
    st.markdown("### Step 2: Select Sheet & Verify Column Mapping")
    
    target_sheet = st.selectbox("Select Sheet to Audit:", st.session_state["sheet_names"])
    
    # Read active sheet to analyze columns
    active_stream = io.BytesIO(st.session_state["raw_excel_bytes"])
    wb_temp = openpyxl.load_workbook(active_stream, data_only=True)
    ws_temp = wb_temp[target_sheet]
    
    header_row_idx, headers_map, def_link, def_uni, def_prog, def_topic, def_deliv = analyze_sheet_structure(ws_temp)
    available_cols = list(headers_map.keys())
    
    st.info(f"Detected Headers on Row {header_row_idx}: {', '.join(available_cols) if available_cols else 'None'}")
    
    col_m1, col_m2, col_m3 = st.columns(3)
    with col_m1:
        link_idx = available_cols.index(def_link) if def_link in available_cols else 0
        sel_link_col = st.selectbox("🔗 Document Link Column (Required):", available_cols, index=link_idx if available_cols else 0)
    with col_m2:
        uni_options = ["[Use Sheet Name as University]"] + available_cols
        uni_idx = uni_options.index(def_uni) if def_uni in uni_options else 0
        sel_uni_col = st.selectbox("🏛️ University Column:", uni_options, index=uni_idx)
    with col_m3:
        prog_options = ["[Default: General Programme]"] + available_cols
        prog_idx = prog_options.index(def_prog) if def_prog in prog_options else 0
        sel_prog_col = st.selectbox("🎓 Programme Column:", prog_options, index=prog_idx)

    col_m4, col_m5 = st.columns(2)
    with col_m4:
        topic_options = ["[Auto: Row Number]"] + available_cols
        topic_idx = topic_options.index(def_topic) if def_topic in topic_options else 0
        sel_topic_col = st.selectbox("📖 Topic / Unit Column:", topic_options, index=topic_idx)
    with col_m5:
        deliv_options = ["[Auto-Detect via AI / Content]"] + available_cols
        deliv_idx = deliv_options.index(def_deliv) if def_deliv in deliv_options else 0
        sel_deliv_col = st.selectbox("📑 Deliverable Type Column:", deliv_options, index=deliv_idx)

    if st.button(f"🚀 Start Batch Audit for '{target_sheet}'", type="primary"):
        if not clean_root_id:
            st.error("Please enter the 'Root QA Reports Folder ID' in the sidebar first.")
        elif not sel_link_col:
            st.error("Please select the Document Link Column.")
        else:
            prog_bar = st.progress(0)
            status_box = st.empty()
            
            active_stream = io.BytesIO(st.session_state["raw_excel_bytes"])
            wb = openpyxl.load_workbook(active_stream)
            ws = wb[target_sheet]
            
            # Map column indices based on user confirmation
            link_col_idx = headers_map.get(sel_link_col)
            uni_col_idx = headers_map.get(sel_uni_col) if sel_uni_col != "[Use Sheet Name as University]" else None
            prog_col_idx = headers_map.get(sel_prog_col) if sel_prog_col != "[Default: General Programme]" else None
            topic_col_idx = headers_map.get(sel_topic_col) if sel_topic_col != "[Auto: Row Number]" else None
            deliv_col_idx = headers_map.get(sel_deliv_col) if sel_deliv_col != "[Auto-Detect via AI / Content]" else None
            
            # Setup QA Report Link column
            report_col_name = "QA Report Link"
            report_col_idx = None
            for col_idx, cell in enumerate(ws[header_row_idx], start=1):
                if cell.value and str(cell.value).strip().lower() == "qa report link":
                    report_col_idx = col_idx
                    break
            if not report_col_idx:
                report_col_idx = max(headers_map.values()) + 1 if headers_map else ws.max_column + 1
                ws.cell(row=header_row_idx, column=report_col_idx, value=report_col_name)

            total_rows = ws.max_row - header_row_idx
            success_count = 0

            for r in range(header_row_idx + 1, ws.max_row + 1):
                raw_link = ws.cell(row=r, column=link_col_idx).value
                if not raw_link or not str(raw_link).strip():
                    continue
                
                row_uni = str(ws.cell(row=r, column=uni_col_idx).value or target_sheet).strip() if uni_col_idx else target_sheet
                row_prog = str(ws.cell(row=r, column=prog_col_idx).value or "General Programme").strip() if prog_col_idx else "General Programme"
                row_topic = str(ws.cell(row=r, column=topic_col_idx).value or f"Module_{r}").strip() if topic_col_idx else f"Module_{r}"
                explicit_deliv = str(ws.cell(row=r, column=deliv_col_idx).value or "").strip() if deliv_col_idx else ""
                
                try:
                    doc_id = extract_drive_id(str(raw_link))
                    doc_stream = download_drive_stream(drive_service, doc_id)
                    doc_text = extract_text_from_stream(doc_stream)
                    
                    if len(doc_text.strip()) < 50:
                        ws.cell(row=r, column=report_col_idx, value="Skipped: Empty Document")
                        continue
                    
                    detected_deliv = detect_deliverable_type(row_topic, doc_text, explicit_deliv)
                    status_box.info(f"⏳ Auditing [{r - header_row_idx}/{total_rows}]: **{row_uni}** > **{row_prog}** > **{row_topic}** ({detected_deliv})...")
                    
                    qa_result = run_stage_2_checker(doc_text, chk_syllabus="", engine=audit_engine)
                    dossier_stream = build_docx_dossier(qa_result, row_uni, detected_deliv, row_topic)
                    
                    report_name = f"{row_uni}_{row_prog}_{row_topic}_{detected_deliv}_QA_Report.docx"
                    report_url = upload_report_to_hierarchy(
                        drive_service,
                        dossier_stream,
                        report_name,
                        clean_root_id,
                        row_uni,
                        row_prog,
                        detected_deliv
                    )
                    
                    ws.cell(row=r, column=report_col_idx, value=report_url)
                    success_count += 1
                except Exception as err:
                    ws.cell(row=r, column=report_col_idx, value=f"Error: {str(err)[:60]}")
                
                prog_bar.progress(min((r - header_row_idx) / max(total_rows, 1), 1.0))

            out_excel = io.BytesIO()
            wb.save(out_excel)
            out_excel.seek(0)
            st.session_state["raw_excel_bytes"] = out_excel.getvalue()

            status_box.empty()
            st.success(f"Audit completed for '{target_sheet}'. Processed {success_count} item(s) and populated Drive links.")
            
            st.download_button(
                label=f"📥 Download Updated Workbook ({target_sheet})",
                data=out_excel,
                file_name=f"Updated_Master_QA_{target_sheet}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )