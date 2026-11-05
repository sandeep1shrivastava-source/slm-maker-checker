import streamlit as st
import io
import re
import os
import openpyxl
import pandas as pd
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
# 2. GOOGLE DRIVE ENGINE (READ-ONLY FOR SOURCE ASSETS)
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
    """Downloads a file stream safely into memory (Read-Only)."""
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
    clean_name = folder_name.strip() if folder_name and folder_name.strip() else "Untitled"
    safe_name = clean_name.replace("'", "\\'")
    query = (
        f"'{parent_id}' in parents and "
        f"name = '{safe_name}' and "
        f"mimeType = 'application/vnd.google-apps.folder' and "
        f"trashed = false"
    )
    res = service.files().list(
        q=query, 
        spaces="drive", 
        fields="files(id, name)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True
    ).execute()
    files = res.get("files", [])
    if files:
        return files[0]["id"]
    
    meta = {
        "name": clean_name,
        "mimeType": "application/vnd.google-apps.folder",
        "parents": [parent_id]
    }
    created = service.files().create(body=meta, fields="id", supportsAllDrives=True).execute()
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
        fields="id, webViewLink",
        supportsAllDrives=True
    ).execute()
    return uploaded.get("webViewLink", "")

# -----------------------------------------------------------------------------
# 3. TEXT EXTRACTOR & REPORT COMPILER
# -----------------------------------------------------------------------------
def extract_text_from_stream(stream: io.BytesIO) -> str:
    """Extracts text content safely from .docx or .pdf streams."""
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
    from docx.shared import Pt, RGBColor
    
    doc = docx.Document()
    title = doc.add_heading("ACADEMIC QUALITY AUDIT REPORT", level=0)
    if title.runs:
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
    if any(k in metadata for k in ["mcq", "quiz", "objective"]): 
        return "MCQ"
    if any(k in metadata for k in ["question bank", "qb", "q&a"]): 
        return "QB"
    if any(k in metadata for k in ["ppt", "deck", "presentation"]): 
        return "PPT"
    if any(k in metadata for k in ["slm", "unit", "module", "chapter"]): 
        return "SLM"

    text_snippet = doc_text[:5000]
    if len(re.findall(r"(\([A-Da-d]\)| Option\s+[A-D] )", text_snippet)) >= 8:
        return "MCQ"
    if len(re.findall(r"(\[\d+\s*Marks?\]|\(\d+\s*Marks?\)| \d+\s*Marks )", text_snippet, re.IGNORECASE)) >= 3:
        return "QB"
    if len(re.findall(r"(Slide\s+\d+|Presentation:)", text_snippet, re.IGNORECASE)) >= 3:
        return "PPT"

    return "SLM"

# -----------------------------------------------------------------------------
# 4. USER INTERFACE & WORKFLOW
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
                st.success(f"Workbook inspected! Found {len(st.session_state['sheet_names'])} sheets.")
            except Exception as e:
                st.error(f"Failed to access workbook: {e}")

if "sheet_names" in st.session_state and st.session_state["sheet_names"]:
    st.divider()
    st.markdown("### Step 2: Select Sheet & Inspect Layout")
    
    target_sheet = st.selectbox("Select Sheet to Audit:", st.session_state["sheet_names"])
    
    # Load sheet into pandas without header to inspect raw rows
    active_stream = io.BytesIO(st.session_state["raw_excel_bytes"])
    df_raw = pd.read_excel(active_stream, sheet_name=target_sheet, header=None)
    
    st.markdown("#### 📋 Sheet Preview (First 8 Rows)")
    st.dataframe(df_raw.head(8), use_container_width=True)
    
    # Let user pick which row actually has the column names
    col_h1, col_h2 = st.columns([1, 2])
    with col_h1:
        header_row_num = st.number_input(
            "Select Header Row Number (1-indexed):",
            min_value=1,
            max_value=min(15, len(df_raw)),
            value=2 if len(df_raw) > 1 else 1,
            step=1,
            help="Look at the preview table above and pick the row that contains your actual column names."
        )
    
    # Extract headers from the chosen row
    header_row_idx = int(header_row_num)
    raw_header_values = df_raw.iloc[header_row_idx - 1].tolist()
    
    # Build dictionary of {Column_Name: 1-based column index}
    col_mapping = {}
    for col_idx, val in enumerate(raw_header_values, start=1):
        if pd.notna(val) and str(val).strip():
            c_name = str(val).strip()
            if c_name in col_mapping:
                c_name = f"{c_name} (Col {col_idx})"
            col_mapping[c_name] = col_idx
        else:
            col_mapping[f"Column {col_idx}"] = col_idx
            
    header_options = list(col_mapping.keys())
    
    # Auto-detect default selections
    def_link_col = next((c for c in header_options if any(k in c.lower() for k in ["link", "url", "drive"])), header_options[0])
    def_uni_col = next((c for c in header_options if any(k in c.lower() for k in ["university", "univ", "college"])), None)
    def_prog_col = next((c for c in header_options if any(k in c.lower() for k in ["programme", "program", "course", "degree"])), None)
    def_topic_col = next((c for c in header_options if any(k in c.lower() for k in ["topic", "unit", "module", "chapter", "title", "subject"])), None)
    def_deliv_col = next((c for c in header_options if any(k in c.lower() for k in ["deliverable", "type", "format", "asset"])), None)

    st.markdown("#### 🔗 Map Columns for Auditing")
    col_m1, col_m2, col_m3 = st.columns(3)
    with col_m1:
        link_idx = header_options.index(def_link_col) if def_link_col in header_options else 0
        sel_link_col = st.selectbox("🔗 Document Link Column (Required):", header_options, index=link_idx)
    with col_m2:
        uni_opts = ["[Use Sheet Name as University]"] + header_options
        uni_idx = uni_opts.index(def_uni_col) if def_uni_col in uni_opts else 0
        sel_uni_col = st.selectbox("🏛️ University Column:", uni_opts, index=uni_idx)
    with col_m3:
        prog_opts = ["[Default: General Programme]"] + header_options
        prog_idx = prog_opts.index(def_prog_col) if def_prog_col in prog_opts else 0
        sel_prog_col = st.selectbox("🎓 Programme Column:", prog_opts, index=prog_idx)

    col_m4, col_m5 = st.columns(2)
    with col_m4:
        topic_opts = ["[Auto: Row Number]"] + header_options
        topic_idx = topic_opts.index(def_topic_col) if def_topic_col in topic_opts else 0
        sel_topic_col = st.selectbox("📖 Topic / Unit Column:", topic_opts, index=topic_idx)
    with col_m5:
        deliv_opts = ["[Auto-Detect via AI / Content]"] + header_options
        deliv_idx = deliv_opts.index(def_deliv_col) if def_deliv_col in deliv_opts else 0
        sel_deliv_col = st.selectbox("📑 Deliverable Type Column:", deliv_opts, index=deliv_idx)

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
            
            # Map column indices based on user selection
            link_col_idx = col_mapping.get(sel_link_col)
            uni_col_idx = col_mapping.get(sel_uni_col) if sel_uni_col != "[Use Sheet Name as University]" else None
            prog_col_idx = col_mapping.get(sel_prog_col) if sel_prog_col != "[Default: General Programme]" else None
            topic_col_idx = col_mapping.get(sel_topic_col) if sel_topic_col != "[Auto: Row Number]" else None
            deliv_col_idx = col_mapping.get(sel_deliv_col) if sel_deliv_col != "[Auto-Detect via AI / Content]" else None
            
            # Setup QA Report Link column
            report_col_name = "QA Report Link"
            report_col_idx = None
            for col_idx, cell in enumerate(ws[header_row_idx], start=1):
                if cell.value and str(cell.value).strip().lower() == "qa report link":
                    report_col_idx = col_idx
                    break
            if not report_col_idx:
                report_col_idx = max(col_mapping.values()) + 1 if col_mapping else ws.max_column + 1
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