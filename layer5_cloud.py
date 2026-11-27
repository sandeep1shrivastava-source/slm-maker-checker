import io, json, re, urllib.parse, requests, openpyxl
from datetime import datetime
import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
from layer3_auth import get_current_auth_token

GCP_SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets",
]


def _extract_oauth_secret_dict() -> dict:
    if "gcp_user_oauth" not in st.secrets:
        return {}
    raw = dict(st.secrets["gcp_user_oauth"])
    if "web" in raw:
        return dict(raw["web"])
    if "installed" in raw:
        return dict(raw["installed"])
    return raw


def get_layer5_gcp_services():
    """Connects Google Drive & Sheets for Layer-5 via Google Email OAuth (supportsAllDrives=True)."""
    creds = None
    secret_oauth = _extract_oauth_secret_dict()
    diag_notes = []
    l5_default_uri = str(
        st.secrets.get("LAYER5_APP_URL", "https://slm-maker-checker-ad99-layer5.streamlit.app")
    ).strip().rstrip("/")

    if st.session_state.get("gcp_oauth_creds_json"):
        try:
            creds = Credentials.from_authorized_user_info(
                json.loads(st.session_state["gcp_oauth_creds_json"]), GCP_SCOPES
            )
        except Exception as exc:
            diag_notes.append(f"Session token error: {exc}")
            creds = None

    if creds is None and secret_oauth.get("refresh_token"):
        try:
            creds = Credentials.from_authorized_user_info(secret_oauth, GCP_SCOPES)
        except Exception as exc:
            diag_notes.append(f"Saved refresh_token error: {exc}")
            creds = None

    auth_code = str(st.query_params.get("code", "")).strip()
    if (creds is None or not creds.valid) and auth_code and secret_oauth.get("client_id") and secret_oauth.get("client_secret"):
        redirect_uri = str(secret_oauth.get("layer5_redirect_uri") or l5_default_uri).strip()
        try:
            resp = requests.post(
                "https://oauth2.googleapis.com/token",
                data={
                    "code": auth_code,
                    "client_id": str(secret_oauth["client_id"]).strip(),
                    "client_secret": str(secret_oauth["client_secret"]).strip(),
                    "redirect_uri": redirect_uri,
                    "grant_type": "authorization_code",
                },
                timeout=20,
            )
            tok_data = resp.json()
            if resp.status_code == 200 and tok_data.get("access_token"):
                creds_payload = {
                    "token": tok_data["access_token"],
                    "refresh_token": tok_data.get("refresh_token", secret_oauth.get("refresh_token", "")),
                    "token_uri": "https://oauth2.googleapis.com/token",
                    "client_id": str(secret_oauth["client_id"]).strip(),
                    "client_secret": str(secret_oauth["client_secret"]).strip(),
                    "scopes": GCP_SCOPES,
                }
                st.session_state["gcp_oauth_creds_json"] = json.dumps(creds_payload)
                creds = Credentials.from_authorized_user_info(creds_payload, GCP_SCOPES)
            else:
                err_desc = tok_data.get("error_description") or tok_data.get("error") or resp.text[:200]
                diag_notes.append(f"Google OAuth token error: {err_desc}")
        except Exception as exc:
            diag_notes.append(f"Google OAuth connection error: {exc}")
        finally:
            for _qp in ("code", "state", "scope", "authuser", "prompt", "hd"):
                st.query_params.pop(_qp, None)

    if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
            st.session_state["gcp_oauth_creds_json"] = creds.to_json()
        except Exception as exc:
            diag_notes.append(f"Access token refresh failed: {exc}")
            creds = None

    if not creds or not creds.valid:
        if not secret_oauth.get("client_id") or not secret_oauth.get("client_secret"):
            return None, None, "", "Missing client_id / client_secret under [gcp_user_oauth] in Streamlit Secrets.", ""
        redirect_uri = str(secret_oauth.get("layer5_redirect_uri") or l5_default_uri).strip()
        login_hint = str(st.session_state.get("layer3_auth_user", "")).strip()
        oauth_params = {
            "client_id": str(secret_oauth["client_id"]).strip(),
            "redirect_uri": redirect_uri,
            "response_type": "code",
            "scope": " ".join(GCP_SCOPES),
            "access_type": "offline",
            "prompt": "consent",
            "state": get_current_auth_token(),
        }
        if "@" in login_hint:
            oauth_params["login_hint"] = login_hint
        auth_url = "https://accounts.google.com/o/oauth2/v2/auth?" + urllib.parse.urlencode(oauth_params)
        reason = " ".join(diag_notes) if diag_notes else "Connect Google Drive & Master Tracker via Google OAuth."
        return None, None, "", reason, auth_url

    drive_svc = build("drive", "v3", credentials=creds, cache_discovery=False)
    sheets_svc = build("sheets", "v4", credentials=creds, cache_discovery=False)
    google_email = st.session_state.get("layer3_auth_user", "Google Account")
    try:
        about = drive_svc.about().get(fields="user(emailAddress)").execute()
        google_email = about.get("user", {}).get("emailAddress") or google_email
    except Exception:
        pass
    return drive_svc, sheets_svc, google_email, "", ""


def extract_drive_id(url_or_id: str) -> str:
    s = str(url_or_id or "").strip()
    if not s:
        return ""
    for pat in (r"/d/([a-zA-Z0-9_-]{15,})", r"/folders/([a-zA-Z0-9_-]{15,})", r"[?&]id=([a-zA-Z0-9_-]{15,})"):
        m = re.search(pat, s)
        if m:
            return m.group(1)
    return s if re.match(r"^[a-zA-Z0-9_-]{15,}$", s) else ""


def download_drive_file_bytes(drive_svc, file_id: str) -> tuple[bytes, str]:
    meta = drive_svc.files().get(fileId=file_id, fields="id,name,mimeType", supportsAllDrives=True).execute()
    name = meta.get("name", f"{file_id}.bin")
    mime = meta.get("mimeType", "")
    buf = io.BytesIO()
    if mime == "application/vnd.google-apps.document":
        req = drive_svc.files().export_media(fileId=file_id, mimeType="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        if not name.lower().endswith(".docx"):
            name += ".docx"
    elif mime == "application/vnd.google-apps.spreadsheet":
        req = drive_svc.files().export_media(fileId=file_id, mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        if not name.lower().endswith(".xlsx"):
            name += ".xlsx"
    elif mime == "application/vnd.google-apps.presentation":
        req = drive_svc.files().export_media(fileId=file_id, mimeType="application/vnd.openxmlformats-officedocument.presentationml.presentation")
        if not name.lower().endswith(".pptx"):
            name += ".pptx"
    else:
        req = drive_svc.files().get_media(fileId=file_id, supportsAllDrives=True)
    downloader = MediaIoBaseDownload(buf, req)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    return buf.getvalue(), name


def upload_file_to_drive(drive_svc, folder_id: str, filename: str, file_bytes: bytes, mime_type: str) -> str:
    body = {"name": filename}
    if folder_id:
        body["parents"] = [folder_id]
    media = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype=mime_type, resumable=True)
    created = drive_svc.files().create(body=body, media_body=media, fields="id,webViewLink", supportsAllDrives=True).execute()
    return created.get("webViewLink") or f"https://drive.google.com/file/d/{created.get('id', '')}/view"


def _parse_json_lenient(raw_txt: str):
    clean = re.sub(r"^```(?:json)?|```$", "", (raw_txt or "").strip(), flags=re.I).strip()
    try:
        return json.loads(clean)
    except Exception:
        m = re.search(r"(\{.*\}|\[.*\])", clean, flags=re.S)
        if m:
            return json.loads(m.group(1))
        raise

def _resolve_provider_key(api_keys: dict, names: tuple[str, ...]) -> str:
    for n in names:
        v = (api_keys or {}).get(n) or (api_keys or {}).get(n.lower()) or (api_keys or {}).get(n.upper())
        if v and str(v).strip():
            return str(v).strip()
    try:
        import streamlit as st
        for n in names:
            if n in st.secrets and str(st.secrets[n]).strip():
                return str(st.secrets[n]).strip()
            for sec_sect in ("api_keys", "llm", "keys"):
                if sec_sect in st.secrets and n in st.secrets[sec_sect]:
                    return str(st.secrets[sec_sect][n]).strip()
    except Exception:
        pass
    import os
    for n in names:
        if os.environ.get(n, "").strip():
            return os.environ[n].strip()
    return ""

def call_layer5_llm(prompt: str, provider: str, api_keys: dict, fallback_payload):
    """Multi-provider AI router with st.secrets auto-resolution, Pro-first routing, and error reporting."""
    if provider.startswith("Deterministic"):
        return fallback_payload
    gem_key = _resolve_provider_key(api_keys, ("gemini", "GEMINI_API_KEY", "GOOGLE_API_KEY"))
    oai_key = _resolve_provider_key(api_keys, ("openai", "OPENAI_API_KEY"))
    ant_key = _resolve_provider_key(api_keys, ("anthropic", "ANTHROPIC_API_KEY"))
    last_err = None
    for attempt in range(3):
        try:
            if provider.startswith("Gemini"):
                if not gem_key:
                    raise RuntimeError("Missing Gemini API key (set GEMINI_API_KEY in Streamlit Secrets or sidebar).")
                model_id = "gemini-3.1-pro-preview" if attempt == 0 else "gemini-3-flash-preview"
                resp = requests.post(
                    f"https://generativelanguage.googleapis.com/v1beta/models/{model_id}:generateContent?key={gem_key}",
                    json={"contents": [{"parts": [{"text": prompt}]}], "generationConfig": {"responseMimeType": "application/json", "maxOutputTokens": 16384}},
                    timeout=180,
                )
                resp.raise_for_status()
                if not (resp.text or "").strip():
                    raise RuntimeError("Gemini returned empty HTTP 200 body")
                txt = resp.json()["candidates"][0]["content"]["parts"][0]["text"]
                return _parse_json_lenient(txt)
            elif provider.startswith("OpenAI"):
                if not oai_key:
                    raise RuntimeError("Missing OpenAI API key (set OPENAI_API_KEY in Streamlit Secrets or sidebar).")
                model_id = "gpt-4o" if attempt == 0 else "gpt-4o-mini"
                resp = requests.post(
                    "https://api.openai.com/v1/chat/completions",
                    headers={"Authorization": f"Bearer {oai_key}", "Content-Type": "application/json"},
                    json={"model": model_id, "response_format": {"type": "json_object"}, "max_tokens": 16384, "messages": [{"role": "user", "content": prompt}]},
                    timeout=180,
                )
                resp.raise_for_status()
                if not (resp.text or "").strip():
                    raise RuntimeError("OpenAI returned empty HTTP 200 body")
                txt = resp.json()["choices"][0]["message"]["content"]
                return _parse_json_lenient(txt)
            elif provider.startswith("Anthropic"):
                if not ant_key:
                    raise RuntimeError("Missing Anthropic API key (set ANTHROPIC_API_KEY in Streamlit Secrets or sidebar).")
                model_id = "claude-sonnet-4-6" if attempt == 0 else "claude-sonnet-5-5"
                resp = requests.post(
                    "https://api.anthropic.com/v1/messages",
                    headers={"x-api-key": ant_key, "anthropic-version": "2023-06-01", "Content-Type": "application/json"},
                    json={"model": model_id, "max_tokens": 16384, "messages": [{"role": "user", "content": prompt + "\nReturn strictly valid JSON."}]},
                    timeout=180,
                )
                resp.raise_for_status()
                if not (resp.text or "").strip():
                    raise RuntimeError("Anthropic returned empty HTTP 200 body")
                txt = resp.json()["content"][0]["text"]
                return _parse_json_lenient(txt)
        except Exception as _e:
            last_err = _e
            import time
            is_429 = "429" in str(_e) or "Too Many Requests" in str(_e)
            wait_s = (15.0 * (attempt + 1)) if is_429 else (2.0 * (attempt + 1))
            time.sleep(wait_s)
    if last_err is not None:
        raise RuntimeError(f"LLM remediation failed ({provider}): {re.sub(r'key=[^&\s]+', 'key=***', str(last_err))}")
    return fallback_payload


def build_complete_mirrored_book_docx(
    brand: dict, program: str, course_name: str, course_code: str, credits: str,
    governance_board: str, course_objectives: list[str], unit_files: list[tuple[str, str, bytes]]
) -> bytes:
    """Aspect 2: Compiles 2+ approved unit .docx files into a print-ready mirrored university book."""
    doc = Document()
    sec0 = doc.sections[0]
    sec0.different_first_page_header_footer = True
    doc.settings.odd_and_even_pages_header_footer = True

    p_uni = doc.add_paragraph()
    p_uni.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_u = p_uni.add_run(brand.get("name", "University Academic Press").upper())
    r_u.bold = True
    r_u.font.size = Pt(16)

    p_prog = doc.add_paragraph()
    p_prog.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_prog.add_run(f"{program}\n").bold = True

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rt = p_title.add_run(f"\n{course_name}\n")
    rt.bold = True
    rt.font.size = Pt(24)

    p_meta = doc.add_paragraph()
    p_meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_meta.add_run(f"Course Code: {course_code}  |  Credits: {credits}\nSelf-Learning Material (SLM) - Print-Ready Master Edition")
    doc.add_page_break()

    doc.add_heading("Editorial & Academic Governance Board", level=1)
    for line in (governance_board or "").splitlines():
        if line.strip():
            doc.add_paragraph(line.strip(), style="List Bullet")

    doc.add_heading("Course Outcomes (CO1-CO5)", level=1)
    for idx, co in enumerate(course_objectives, start=1):
        if co.strip():
            doc.add_paragraph(f"CO{idx}: {co.strip()}", style="List Bullet")

    doc.add_heading("Master Table of Contents (Press Ctrl+A then F9 in Word to Refresh Page Numbers)", level=1)
    p_toc = doc.add_paragraph()
    fld_begin = parse_xml(r'<w:fldChar %s w:fldCharType="begin"/>' % nsdecls('w'))
    instr = parse_xml(r'<w:instrText %s xml:space="preserve"> TOC \o "1-2" \h \z \u </w:instrText>' % nsdecls('w'))
    fld_sep = parse_xml(r'<w:fldChar %s w:fldCharType="separate"/>' % nsdecls('w'))
    fld_end = parse_xml(r'<w:fldChar %s w:fldCharType="end"/>' % nsdecls('w'))
    r_f = p_toc.add_run()
    r_f._r.append(fld_begin)
    r_f._r.append(instr)
    r_f._r.append(fld_sep)
    r_f._r.append(fld_end)

    for u_idx, (block_lbl, u_name, u_bytes) in enumerate(unit_files, start=1):
        new_sec = doc.add_section()
        new_sec.different_first_page_header_footer = False
        new_sec.header.paragraphs[0].text = f"{course_code}: {course_name}  |  {block_lbl} (Recto - Study Notes)"
        new_sec.even_page_header.paragraphs[0].text = f"Notes Column  |  {brand.get('name', 'University')} - {block_lbl}"
        new_sec.footer.paragraphs[0].text = brand.get("footer", "Approved Academic Publication - Clone-on-Write Master")

        doc.add_heading(f"{block_lbl}: {u_name}", level=1)
        u_doc = Document(io.BytesIO(u_bytes))
        for p in u_doc.paragraphs:
            if p.text.strip():
                st_name = p.style.name if p.style and p.style.name.startswith("Heading") else "Normal"
                if st_name == "Heading 1":
                    st_name = "Heading 2"
                doc.add_paragraph(p.text.strip(), style=st_name)
        for tbl in u_doc.tables:
            if not tbl.rows:
                continue
            nt = doc.add_table(rows=len(tbl.rows), cols=len(tbl.rows[0].cells))
            nt.style = "Table Grid"
            for r_i, row in enumerate(tbl.rows):
                for c_i, cell in enumerate(row.cells):
                    nt.cell(r_i, c_i).text = cell.text.strip()

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def build_subject_cumulative_dossier(
    brand: dict, program: str, subj_code: str, subj_name: str, parsed_reports: list[dict]
) -> tuple[bytes, bytes, float]:
    """Aspect 3: Generates 4-Tab Excel Cumulative QA Dossier + Subject QA Release Certificate (.docx)."""
    wb = openpyxl.Workbook()
    ws_sum = wb.active
    ws_sum.title = "1_Executive_Summary"
    ws_sum.append(["Institution", brand.get("name", "")])
    ws_sum.append(["Program", program])
    ws_sum.append(["Subject Code & Name", f"{subj_code} - {subj_name}"])
    ws_sum.append(["Total Layer-4 Reports Audited", len(parsed_reports)])

    scores = []
    for r in parsed_reports:
        m = re.search(r"(\d+(?:\.\d+)?)", str(r.get("overall_score", "")))
        if m:
            scores.append(float(m.group(1)))
    avg_score = round(sum(scores) / len(scores), 2) if scores else 0.0
    readiness = "APPROVED FOR RELEASE" if avg_score >= 85.0 and parsed_reports else "REMEDIATION REQUIRED"
    ws_sum.append(["Cumulative Subject QA Score", f"{avg_score}%"])
    ws_sum.append(["Subject Release Status", readiness])

    ws_files = wb.create_sheet("2_Deliverable_Matrix")
    ws_files.append(["Deliverable File", "Asset Type", "Layer-4 Score", "Verdict", "Total Findings"])
    ws_defects = wb.create_sheet("3_Consolidated_Defects")
    ws_defects.append(["Source Report", "Severity / Category", "Finding Description"])
    ws_sign = wb.create_sheet("4_QA_SignOff_Log")
    ws_sign.append(["Generated Timestamp", datetime.now().strftime("%Y-%m-%d %H:%M"), "Status", readiness])

    for r in parsed_reports:
        ws_files.append([r.get("filename", ""), r.get("asset_type", "SLM"), r.get("overall_score", ""), r.get("verdict", ""), r.get("total_findings", 0)])
        for f in r.get("technical_errors", []) + r.get("clarity_issues", []):
            ws_defects.append([r.get("filename", ""), "Flagged Defect", f])

    xl_buf = io.BytesIO()
    wb.save(xl_buf)

    cert = Document()
    cert.add_heading(brand.get("name", "University Quality Assurance Board").upper(), level=1)
    cert.add_heading("SUBJECT-WIDE CUMULATIVE QA & RELEASE CERTIFICATE", level=2)
    cert.add_paragraph(f"Program: {program}")
    cert.add_paragraph(f"Course: {subj_code} - {subj_name}")
    cert.add_paragraph(f"Deliverables Audited: {len(parsed_reports)} | Cumulative Score: {avg_score}% | Status: {readiness}")
    t = cert.add_table(rows=1, cols=4)
    t.style = "Table Grid"
    for idx, h in enumerate(["Deliverable", "Type", "Score", "Findings"]):
        t.rows[0].cells[idx].text = h
    for r in parsed_reports:
        row = t.add_row().cells
        row[0].text = str(r.get("filename", ""))
        row[1].text = str(r.get("asset_type", ""))
        row[2].text = str(r.get("overall_score", ""))
        row[3].text = str(r.get("total_findings", 0))
    cert.add_paragraph(f"\nIssued via Layer-5 Academic QA Workbench | {datetime.now().strftime('%d %B %Y')}")
    doc_buf = io.BytesIO()
    cert.save(doc_buf)
    return xl_buf.getvalue(), doc_buf.getvalue(), avg_score



def col_idx_to_a1(idx_0: int) -> str:
    n = idx_0 + 1
    chars = []
    while n > 0:
        n, rem = divmod(n - 1, 26)
        chars.append(chr(65 + rem))
    return "".join(reversed(chars))


def sanitize_drive_name(name: str, fallback: str = "General") -> str:
    cleaned = re.sub(r'[\\/:*?"<>|]+', "_", str(name or "").strip()).strip(" ._")
    return cleaned[:90] or fallback


def get_or_create_drive_subfolder(drive_svc, parent_id: str, folder_name: str) -> str:
    safe_name = sanitize_drive_name(folder_name, "General")
    q = f"name = '{safe_name.replace(chr(39), chr(92) + chr(39))}' and '{parent_id}' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
    res = drive_svc.files().list(q=q, fields="files(id,name)", supportsAllDrives=True, includeItemsFromAllDrives=True).execute()
    files = res.get("files", [])
    if files:
        return files[0]["id"]
    meta = {"name": safe_name, "mimeType": "application/vnd.google-apps.folder", "parents": [parent_id]}
    created = drive_svc.files().create(body=meta, fields="id", supportsAllDrives=True).execute()
    return created["id"]


def upload_to_drive_hierarchy(drive_svc, root_folder_id: str, hierarchy_parts: list[str], filename: str, file_bytes: bytes, mime_type: str) -> str:
    curr_parent = root_folder_id
    for part in hierarchy_parts:
        if str(part or "").strip():
            curr_parent = get_or_create_drive_subfolder(drive_svc, curr_parent, str(part))
    return upload_file_to_drive(drive_svc, curr_parent, sanitize_drive_name(filename, "Layer5_Output.docx"), file_bytes, mime_type)


def load_master_tracker_tabs(sheets_svc, spreadsheet_id: str) -> list[str]:
    meta = sheets_svc.spreadsheets().get(spreadsheetId=spreadsheet_id, fields="sheets.properties.title").execute()
    return [s["properties"]["title"] for s in meta.get("sheets", [])]


def read_master_tracker_tab(sheets_svc, spreadsheet_id: str, tab_name: str) -> tuple[int, list[str], list[list[str]]]:
    safe_tab = tab_name.replace("'", "''")
    res = sheets_svc.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=f"'{safe_tab}'").execute()
    values = res.get("values", [])
    if not values:
        return 1, [], []
    hdr_idx = 0
    for i, row in enumerate(values[:10]):
        joined = " ".join(str(c).lower() for c in row)
        if any(k in joined for k in ("manuscript", "link", "url", "qa report", "subject", "unit", "topic")):
            hdr_idx = i
            break
    raw_headers = [str(h).strip() or f"Col_{c+1}" for c, h in enumerate(values[hdr_idx])]
    data_rows = values[hdr_idx + 1 :]
    return hdr_idx + 1, raw_headers, data_rows


def sync_layer5_sheet_row(
    sheets_svc,
    spreadsheet_id: str,
    tab_name: str,
    header_row_1idx: int,
    data_row_1idx: int,
    l5_link_col_0idx: int,
    l5_status_col_0idx: int,
    remediated_link: str,
    status_text: str,
    write_headers: bool = False,
    history_col_0idx: int | None = None,
    history_entry: str = "",
    write_history_header: bool = False,
):
    resolved_tab = tab_name
    try:
        meta = sheets_svc.spreadsheets().get(spreadsheetId=spreadsheet_id, fields="sheets.properties.title").execute()
        for s in meta.get("sheets", []):
            lt = s.get("properties", {}).get("title", "")
            if lt and (lt == resolved_tab or lt[:31] == resolved_tab[:31]):
                resolved_tab = lt
                break
    except Exception:
        pass
    safe_tab = resolved_tab.replace("'", "''")
    updates = []
    if write_headers:
        updates.append({"range": f"'{safe_tab}'!{col_idx_to_a1(l5_link_col_0idx)}{header_row_1idx}", "values": [["Latest Manuscript Link"]]})
        updates.append({"range": f"'{safe_tab}'!{col_idx_to_a1(l5_status_col_0idx)}{header_row_1idx}", "values": [["Layer-5 Status"]]})
    if write_history_header and history_col_0idx is not None:
        updates.append({"range": f"'{safe_tab}'!{col_idx_to_a1(history_col_0idx)}{header_row_1idx}", "values": [["Audit & Remediation History"]]})
    if remediated_link:
        updates.append({"range": f"'{safe_tab}'!{col_idx_to_a1(l5_link_col_0idx)}{data_row_1idx}", "values": [[remediated_link]]})
    updates.append({"range": f"'{safe_tab}'!{col_idx_to_a1(l5_status_col_0idx)}{data_row_1idx}", "values": [[status_text]]})
    if history_col_0idx is not None and history_entry:
        existing_hist = ""
        try:
            cur = sheets_svc.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=f"'{safe_tab}'!{col_idx_to_a1(history_col_0idx)}{data_row_1idx}",
            ).execute()
            vals = cur.get("values", [])
            if vals and vals[0]:
                existing_hist = str(vals[0][0] or "").strip()
        except Exception:
            pass
        combined_hist = f"{existing_hist}\n{history_entry}".strip() if existing_hist else history_entry
        updates.append({"range": f"'{safe_tab}'!{col_idx_to_a1(history_col_0idx)}{data_row_1idx}", "values": [[combined_hist]]})
    sheets_svc.spreadsheets().values().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"valueInputOption": "USER_ENTERED", "data": updates},
    ).execute()
