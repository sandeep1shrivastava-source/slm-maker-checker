import base64
import csv
import io
import json
import math
import pathlib
import re
import time
import zipfile
from collections import Counter
from datetime import datetime, timezone
from typing import Dict, List, Optional, Tuple

import docx
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pypdf
import qrcode
import requests
import streamlit as st
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from googleapiclient.http import MediaIoBaseUpload
from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn as pptx_qn

try:
    from layer4_rigor_engine import compute_prose_burstiness as _l4_compute_burstiness
except ImportError:
    _l4_compute_burstiness = None

PROFILES_FILE = pathlib.Path("university_profiles.json")


# ==============================================================================
# 1. DYNAMIC UNIVERSITY BRAND & PROFILE ENGINE
# ==============================================================================

def load_university_profiles() -> Dict[str, Dict]:
    """Loads saved university brand profiles from disk, creating defaults if absent."""
    if PROFILES_FILE.exists():
        try:
            data = json.loads(PROFILES_FILE.read_text(encoding="utf-8"))
            if isinstance(data, dict) and data:
                return data
        except Exception:
            pass
    default_profiles = {
        "ADYPU": {
            "university_id": "ADYPU",
            "university_name": "AJEENKYA DY PATIL UNIVERSITY",
            "tagline": "THE INNOVATION UNIVERSITY",
            "school_name": "Centre for Distance and Online Education",
            "primary_hex": "1E3A8A",
            "accent_hex": "D97706",
            "copyright_line": "All Copyrights reserved with Ajeenkya DY Patil University (CDOE)",
            "statutory_footer": "A State Private University established under the Maharashtra State Legislature Act, 2015 | Recognized by the University Grants Commission (UGC), Government of India",
            "enable_mirror_notes_margin": True,
            "default_committee": [
                {"role": "Vice Chancellor - Chairman", "name": "Dr. Rakesh Kumar Jain"},
                {"role": "Dean Academics", "name": "Dr. Sunil Ingole"},
                {"role": "Professor & Director - CDOE", "name": "Dr. Kamaljeet Kaur Siddhu"},
                {"role": "Professor & Dean - IQAC", "name": "Dr. Deepa Joshi"},
                {"role": "Professor & Registrar", "name": "Dr. Sudhakar Shinde"},
            ],
        },
        "GENERIC_UGC": {
            "university_id": "GENERIC_UGC",
            "university_name": "STATE OPEN & DISTANCE LEARNING UNIVERSITY",
            "tagline": "EXCELLENCE IN ACCESSIBLE HIGHER EDUCATION",
            "school_name": "Centre for Distance and Online Education (CDOE)",
            "primary_hex": "0F4C81",
            "accent_hex": "0D9488",
            "copyright_line": "All Copyrights reserved with the University (CDOE)",
            "statutory_footer": "Established by State Legislature Act | Recognized by the University Grants Commission (UGC-DEB), Government of India",
            "enable_mirror_notes_margin": True,
            "default_committee": [
                {"role": "Vice Chancellor - Chairman", "name": "Prof. (Dr.) A. K. Sharma"},
                {"role": "Dean Academics", "name": "Prof. (Dr.) R. M. Kulkarni"},
                {"role": "Director - Distance & Online Education", "name": "Dr. S. P. Deshmukh"},
            ],
        },
    }
    PROFILES_FILE.write_text(json.dumps(default_profiles, indent=2), encoding="utf-8")
    return default_profiles


def save_university_profile(profile: Dict) -> Dict[str, Dict]:
    """Saves or updates a university brand profile dynamically."""
    profiles = load_university_profiles()
    uid = re.sub(r"[^A-Za-z0-9_-]+", "_", str(profile.get("university_id") or "CUSTOM_UNI")).upper()
    profile["university_id"] = uid
    clean_save = {k: v for k, v in profile.items() if k != "logo_bytes"}
    profiles[uid] = clean_save
    PROFILES_FILE.write_text(json.dumps(profiles, indent=2), encoding="utf-8")
    return profiles


def _hex_to_rgb_tuple(hex_str: str, fallback: Tuple[int, int, int] = (30, 58, 138)) -> Tuple[int, int, int]:
    h = str(hex_str or "").strip().lstrip("#")
    if len(h) == 6:
        try:
            return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
        except ValueError:
            pass
    return fallback


def _tint_hex(hex_str: str, factor: float = 0.90) -> str:
    r, g, b = _hex_to_rgb_tuple(hex_str)
    tr = int(r + (255 - r) * factor)
    tg = int(g + (255 - g) * factor)
    tb = int(b + (255 - b) * factor)
    return f"{tr:02X}{tg:02X}{tb:02X}"


def _generate_brand_crest_png(university_name: str, primary_hex: str = "1E3A8A", accent_hex: str = "D97706") -> bytes:
    """Generates a crisp institutional monogram emblem PNG when no custom logo file is uploaded."""
    words = [w for w in re.findall(r"[A-Za-z]+", university_name or "UNIV") if w.lower() not in ("of", "and", "the", "for")]
    initials = "".join(w[0].upper() for w in words[:4]) or "UNIV"

    size = 240
    img = Image.new("RGBA", (size, size), (255, 255, 255, 0))
    draw = ImageDraw.Draw(img)
    p_rgb = _hex_to_rgb_tuple(primary_hex, (30, 58, 138))
    a_rgb = _hex_to_rgb_tuple(accent_hex, (217, 119, 6))

    draw.ellipse([10, 10, size - 10, size - 10], fill=(255, 255, 255, 255), outline=p_rgb, width=10)
    draw.ellipse([26, 26, size - 26, size - 26], outline=a_rgb, width=4)
    font = ImageFont.load_default()
    tx = max(44, (size - len(initials) * 14) // 2)
    draw.text((tx, size // 2 - 8), initials, fill=p_rgb, font=font)

    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


# ==============================================================================
# 2. MANDATORY LAYER-4 / LAYER-4+ AUDIT REVIEW REPORT PARSER & CHECKPOINT GATE
# ==============================================================================

def classify_deliverable_category(filename: str, text_sample: str = "") -> str:
    """Auto-classifies an academic asset or review report into SLM, MCQ, QB, PPT, FAQ, or GENERAL."""
    fn = (filename or "").lower()
    blob = (fn + " " + (text_sample or "")[:1500]).lower()
    if any(k in fn for k in ("mcq", "multiple_choice", "multiple-choice", "quiz")):
        return "MCQ"
    if any(k in fn for k in ("_qb", "qb_", "question_bank", "question-bank", "question bank")):
        return "QB"
    if any(k in fn for k in ("faq", "frequently_asked", "frequently asked")):
        return "FAQ"
    if fn.endswith(".pptx") or any(k in fn for k in ("_ppt", "ppt_", "presentation", "slide_deck", "slides")):
        return "PPT"
    if any(k in fn for k in ("slm", "unit", "chapter", "module", "block", "book")):
        return "SLM"
    if "7 defect" in blob or "distractor" in blob or "psychometric" in blob:
        return "MCQ"
    if "question bank" in blob or "bloom" in blob:
        return "QB"
    if "frequently asked questions" in blob or "faq" in blob:
        return "FAQ"
    if "slide" in blob and "animation" in blob:
        return "PPT"
    if "12 ugc" in blob or "self-learning material" in blob or "learning objectives" in blob:
        return "SLM"
    return "SLM"


def _extract_report_raw_structures(report_bytes: bytes, report_filename: str) -> Tuple[str, List[Dict], Dict[str, str]]:
    """Extracts plain text, tabular rows, and key-value parameter pairs from .xlsx, .pdf, .csv, or .json audit reports."""
    fn = (report_filename or "").lower()
    text_parts: List[str] = []
    rows_out: List[Dict] = []
    kv_params: Dict[str, str] = {}

    if fn.endswith(".docx"):
        doc = docx.Document(io.BytesIO(report_bytes))
        for p in doc.paragraphs:
            txt = re.sub(r"\s+", " ", p.text or "").strip()
            if not txt:
                continue
            text_parts.append(txt)
            if ":" in txt:
                k_part, v_part = txt.split(":", 1)
                if 2 <= len(k_part.strip()) <= 65 and v_part.strip():
                    kv_params[k_part.strip()] = v_part.strip()

        for t_idx, tbl in enumerate(doc.tables, start=1):
            t_rows = []
            for r in tbl.rows:
                row_vals = [re.sub(r"\s+", " ", c.text or "").strip() for c in r.cells]
                t_rows.append(row_vals)
                non_empty = [v for v in row_vals if v]
                if non_empty:
                    text_parts.append(" | ".join(non_empty))
                if len(row_vals) >= 2 and row_vals[0] and row_vals[1] and len(row_vals[0]) <= 65:
                    kv_params[row_vals[0]] = row_vals[1]

            if len(t_rows) >= 2:
                headers = [h if h else f"Col_{i+1}" for i, h in enumerate(t_rows[0])]
                for r_vals in t_rows[1:]:
                    if any(v for v in r_vals):
                        row_dict = {"_sheet": f"Table_{t_idx}"}
                        for ci, h_name in enumerate(headers):
                            row_dict[h_name] = r_vals[ci] if ci < len(r_vals) else ""
                        rows_out.append(row_dict)

    elif fn.endswith(".xlsx"):
        wb = openpyxl.load_workbook(io.BytesIO(report_bytes), data_only=True)
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            sheet_rows = list(ws.iter_rows(values_only=True))
            if not sheet_rows:
                continue
            headers = [str(h).strip() if h is not None else f"Col_{i+1}" for i, h in enumerate(sheet_rows[0])]
            text_parts.append(f"=== Sheet: {sheet_name} ===")
            for r in sheet_rows:
                clean_vals = [str(v).strip() for v in r if v is not None and str(v).strip()]
                if clean_vals:
                    text_parts.append(" | ".join(clean_vals))
                if len(clean_vals) >= 2 and len(clean_vals[0]) <= 65:
                    kv_params[clean_vals[0]] = clean_vals[1]
            for r in sheet_rows[1:]:
                if any(v is not None and str(v).strip() for v in r):
                    row_dict = {"_sheet": sheet_name}
                    for ci, h_name in enumerate(headers):
                        row_dict[h_name] = str(r[ci]).strip() if ci < len(r) and r[ci] is not None else ""
                    rows_out.append(row_dict)

    elif fn.endswith(".csv"):
        decoded = report_bytes.decode("utf-8", errors="ignore")
        reader = csv.DictReader(io.StringIO(decoded))
        for row in reader:
            clean_row = {str(k).strip(): str(v).strip() for k, v in row.items() if k}
            rows_out.append(clean_row)
            text_parts.append(" | ".join(f"{k}: {v}" for k, v in clean_row.items() if v))
            for k, v in clean_row.items():
                if v:
                    kv_params[k] = v

    elif fn.endswith(".json"):
        decoded = report_bytes.decode("utf-8", errors="ignore")
        data = json.loads(decoded)
        text_parts.append(json.dumps(data, indent=2))
        if isinstance(data, dict):
            for k, v in data.items():
                if not isinstance(v, (dict, list)):
                    kv_params[str(k)] = str(v)
                elif isinstance(v, dict):
                    for sk, sv in v.items():
                        kv_params[f"{k}.{sk}"] = str(sv)
            rows_out.append({k: str(v) for k, v in data.items() if not isinstance(v, (dict, list))})
        elif isinstance(data, list):
            for item in data:
                if isinstance(item, dict):
                    rows_out.append({str(k): str(v) for k, v in item.items()})

    elif fn.endswith(".pdf"):
        reader = pypdf.PdfReader(io.BytesIO(report_bytes))
        for p_idx, page in enumerate(reader.pages, start=1):
            p_txt = page.extract_text() or ""
            text_parts.append(f"--- Page {p_idx} ---\n{p_txt}")
            for ln in p_txt.splitlines():
                if ":" in ln:
                    parts = ln.split(":", 1)
                    if 2 <= len(parts[0].strip()) <= 55 and parts[1].strip():
                        kv_params[parts[0].strip()] = parts[1].strip()

    else:
        decoded = report_bytes.decode("utf-8", errors="ignore")
        text_parts.append(decoded)

    return "\n".join(text_parts), rows_out, kv_params


def parse_and_verify_audit_report(
    report_bytes: bytes,
    report_filename: str,
    manuscript_filename: str = "",
    manuscript_text: str = "",
    manuscript_cv: float = 0.0,
) -> Dict:
    """
    Parses a Layer-4 / Layer-4+ Audit Review Report (.xlsx, .pdf, .csv, .json) and verifies:
    1. Target manuscript identity match (prevents cross-file report spoofing).
    2. Audit verdict ('Approved' / 'Pass') and score >= 85.0%.
    3. UGC Quadrant-II gates (Similarity <= 10%, AI Content <= 20%, Boilerplate == 0%).
    4. Extracts all evaluated parameters and specific defect remarks for Track A remediation.
    """
    raw_text, rows_out, kv_params = _extract_report_raw_structures(report_bytes, report_filename)
    deliverable_type = classify_deliverable_category(manuscript_filename or report_filename, raw_text)

    # 1. Extract Score (%)
    score_val: Optional[float] = None
    for k, v in kv_params.items():
        kl = k.lower()
        if any(tok in kl for tok in ("final score", "audit score", "overall score", "total score", "rubric score", "score")):
            m = re.search(r"(\d{1,3}(?:\.\d+)?)\s*%?", str(v))
            if m:
                cand = float(m.group(1))
                if 0.0 <= cand <= 100.0:
                    score_val = cand
                    break
    if score_val is None:
        m_txt = re.search(
            r"(?:Final|Overall|Audit|Total|Compliance)\s*Score\s*[:\-–=]?\s*(\d{1,3}(?:\.\d+)?)\s*%",
            raw_text,
            re.IGNORECASE,
        )
        if m_txt:
            score_val = float(m_txt.group(1))

    # 2. Extract Audit Verdict / Status
    status_norm = "Unverified"
    raw_status = ""
    for k, v in kv_params.items():
        kl = k.lower()
        if any(tok in kl for tok in ("audit status", "verdict", "decision", "outcome", "status")):
            vl = str(v).lower()
            if any(w in vl for w in ("approved", "pass", "major", "minor", "reject", "fail")):
                raw_status = str(v)
                break
    if not raw_status:
        m_st = re.search(
            r"(?:Audit\s*Status|Verdict|Decision|Outcome)\s*[:\-–=]?\s*(Approved|Pass(?:ed)?|Major\s*Changes?|Minor\s*(?:Changes?|Correction)|Rejected|Fail(?:ed)?)",
            raw_text,
            re.IGNORECASE,
        )
        if m_st:
            raw_status = m_st.group(1)

    low_st = raw_status.lower()
    if "reject" in low_st or "fail" in low_st:
        status_norm = "Rejected"
    elif "major" in low_st:
        status_norm = "Major Changes"
    elif "minor" in low_st:
        status_norm = "Minor Changes"
    elif "approved" in low_st or "pass" in low_st:
        status_norm = "Approved"
    elif score_val is not None:
        if score_val >= 85.0:
            status_norm = "Approved"
        elif score_val >= 75.0:
            status_norm = "Minor Changes"
        elif score_val >= 60.0:
            status_norm = "Major Changes"
        else:
            status_norm = "Rejected"

    if score_val is None:
        score_val = 90.0 if status_norm == "Approved" else (78.0 if status_norm == "Minor Changes" else (62.0 if status_norm == "Major Changes" else 0.0))

    # 3. Extract Quadrant-II Originality Metrics (Similarity %, AI %, Boilerplate %)
    sim_pct: Optional[float] = None
    ai_pct: Optional[float] = None
    bp_pct: Optional[float] = None

    m_sim = re.search(r"(?:Similarity|Plagiarism)[^\d\n]{0,25}(\d{1,2}(?:\.\d+)?)\s*%", raw_text, re.IGNORECASE)
    if m_sim:
        sim_pct = float(m_sim.group(1))

    m_ai = re.search(r"(?:AI\s*Content|AI\s*Detection|AI\s*Probability)[^\d\n]{0,25}(\d{1,2}(?:\.\d+)?)\s*%", raw_text, re.IGNORECASE)
    if m_ai:
        ai_pct = float(m_ai.group(1))

    m_bp = re.search(r"(?:Boilerplate)[^\d\n]{0,25}(\d{1,2}(?:\.\d+)?)\s*%", raw_text, re.IGNORECASE)
    if m_bp:
        bp_pct = float(m_bp.group(1))

    q2_ok = True
    q2_notes = []
    if sim_pct is not None and sim_pct > 10.0:
        q2_ok = False
        q2_notes.append(f"Similarity {sim_pct:.1f}% > 10% cap")
    if ai_pct is not None and ai_pct > 20.0:
        q2_ok = False
        q2_notes.append(f"AI Content {ai_pct:.1f}% > 20% cap")
    if bp_pct is not None and bp_pct > 0.0:
        q2_ok = False
        q2_notes.append(f"Boilerplate {bp_pct:.1f}% > 0% cap")

    # 4. Extract Auditor / Checker Identity
    auditor = "Verified Layer-4 Auditor"
    for k, v in kv_params.items():
        if any(tok in k.lower() for tok in ("checker", "auditor", "reviewed by", "approved by", "sign-off")):
            if v and "@" in v or len(v) >= 3:
                auditor = v
                break

    # 5. Extract Evaluated Parameters & Flagged Defects
    evaluated_params: List[Dict] = []
    flagged_defects: List[str] = []

    for r in rows_out:
        param_name = r.get("Parameter") or r.get("Dimension") or r.get("Rule") or r.get("Criterion") or r.get("Particulars") or ""
        param_score = r.get("Score") or r.get("Marks") or r.get("Value") or r.get("Rating") or ""
        param_stat = r.get("Status") or r.get("Result") or r.get("Verdict") or ""
        param_rem = r.get("Remarks") or r.get("Defect") or r.get("Finding") or r.get("Action") or ""
        if param_name:
            evaluated_params.append({
                "Deliverable": deliverable_type,
                "Source_Report": report_filename,
                "Parameter": param_name,
                "Score_or_Value": param_score,
                "Status": param_stat or status_norm,
                "Remarks": param_rem,
            })
        combined_line = f"{param_name} {param_stat} {param_rem}".strip()
        if any(w in combined_line.lower() for w in ("fail", "missing", "defect", "penalty", "flaw", "low burstiness", "below", "flag", "incomplete", "error")):
            if combined_line and combined_line not in flagged_defects:
                flagged_defects.append(combined_line[:220])

    if not evaluated_params:
        for k, v in list(kv_params.items())[:20]:
            evaluated_params.append({
                "Deliverable": deliverable_type,
                "Source_Report": report_filename,
                "Parameter": k,
                "Score_or_Value": v,
                "Status": status_norm,
                "Remarks": "",
            })

    for ln in raw_text.splitlines():
        s_ln = ln.strip()
        if 12 <= len(s_ln) <= 220 and any(w in s_ln.lower() for w in ("defect:", "flag:", "missing ", "penalty", "failed:", "remediate", "action required")):
            if s_ln not in flagged_defects:
                flagged_defects.append(s_ln)

    # 6. Verify Manuscript-to-Report Identity Match
    manuscript_matched = True
    match_reason = "No manuscript filename supplied (standalone report mode)."
    if manuscript_filename:
        ms_stem = re.sub(r"[^a-z0-9]+", " ", manuscript_filename.rsplit(".", 1)[0].lower()).strip()
        rep_stem = re.sub(r"[^a-z0-9]+", " ", report_filename.rsplit(".", 1)[0].lower()).strip()
        ignore_toks = {"audit", "report", "layer", "layer4", "layer2", "qa", "dossier", "review", "approved", "final", "v1", "v2", "docx", "xlsx", "pdf", "pptx"}
        ms_tokens = {t for t in ms_stem.split() if len(t) >= 2 and t not in ignore_toks}
        rep_blob_tokens = set(re.findall(r"[a-z0-9]{2,}", (rep_stem + " " + raw_text[:2500]).lower()))
        overlap = ms_tokens & rep_blob_tokens
        ratio = (len(overlap) / len(ms_tokens)) if ms_tokens else 1.0
        if ms_tokens and ratio < 0.40 and len(overlap) < 2:
            manuscript_matched = False
            match_reason = f"Target mismatch: report does not reference '{manuscript_filename}' (matched tokens: {sorted(overlap)})."
        else:
            manuscript_matched = True
            match_reason = f"Verified identity match ({len(overlap)}/{max(1, len(ms_tokens))} key tokens matched)."

    cv_ok = (manuscript_cv >= 0.400) if manuscript_text else True
    verified_for_enrichment = (
        manuscript_matched
        and status_norm == "Approved"
        and score_val >= 85.0
        and q2_ok
        and cv_ok
    )

    block_reasons = []
    if not manuscript_matched:
        block_reasons.append(match_reason)
    if status_norm != "Approved" or score_val < 85.0:
        block_reasons.append(f"Audit status is '{status_norm}' ({score_val:.1f}% < 85.0% required).")
    if not q2_ok:
        block_reasons.extend(q2_notes)
    if manuscript_text and not cv_ok:
        block_reasons.append(f"Live Prose Burstiness CV is {manuscript_cv:.3f} (< 0.400 threshold).")

    return {
        "report_filename": report_filename,
        "deliverable_type": deliverable_type,
        "audit_status": status_norm,
        "audit_score": round(score_val, 1),
        "similarity_pct": sim_pct,
        "ai_content_pct": ai_pct,
        "boilerplate_pct": bp_pct,
        "q2_passed": q2_ok,
        "auditor_name": auditor,
        "manuscript_matched": manuscript_matched,
        "match_reason": match_reason,
        "evaluated_parameters": evaluated_params,
        "flagged_defects": flagged_defects[:15],
        "verified_for_enrichment": verified_for_enrichment,
        "block_reasons": block_reasons,
    }


# ==============================================================================
# 3. ASPECT 3: SUBJECT-WIDE CUMULATIVE QA & ENRICHMENT REPORT ENGINE
# ==============================================================================

def compile_subject_cumulative_dossier(
    subject_meta: Dict,
    report_items: List[Dict],
    brand_profile: Dict,
    reviewer_email: str,
) -> Tuple[bytes, bytes, Dict]:
    """
    Aggregates all Layer-4 / Layer-4+ review reports across a complete subject
    (SLM, MCQ, QB, PPT, FAQ, etc.) into:
    1. Multi-tab Institutional Excel Dossier (.xlsx)
    2. Printable Subject QA Compliance & Release Certificate (.docx)
    3. Structured summary dict for live Streamlit KPI dashboard
    """
    p_hex = brand_profile.get("primary_hex", "1E3A8A")
    a_hex = brand_profile.get("accent_hex", "D97706")
    uni_name = brand_profile.get("university_name", "UNIVERSITY").upper()
    school_name = brand_profile.get("school_name", "Centre for Distance and Online Education")
    logo_bytes = brand_profile.get("logo_bytes") or _generate_brand_crest_png(uni_name, p_hex, a_hex)

    prog_name = subject_meta.get("program_name", "MBA")
    subj_code = subject_meta.get("course_code", "SUBJ-101")
    subj_name = subject_meta.get("course_name", "Complete Subject")
    acad_term = subject_meta.get("academic_term", "2026 Cohort")

    parsed_reports: List[Dict] = []
    for item in report_items:
        rep = parse_and_verify_audit_report(
            report_bytes=item["bytes"],
            report_filename=item["filename"],
            manuscript_filename="",
        )
        if item.get("override_type") and item["override_type"] != "Auto-Detect":
            rep["deliverable_type"] = item["override_type"]
        rep["l5_enrichment_state"] = item.get("l5_state", "Verified Ready" if rep["audit_status"] == "Approved" else "Track A Remediation Queued")
        parsed_reports.append(rep)

    # Category-wise rollup across SLM, MCQ, QB, PPT, FAQ
    all_categories = ["SLM", "MCQ", "QB", "PPT", "FAQ"]
    cat_summary_rows = []
    for cat in all_categories:
        subset = [r for r in parsed_reports if r["deliverable_type"] == cat]
        if not subset:
            cat_summary_rows.append({
                "Deliverable_Category": cat,
                "Files_Evaluated": 0,
                "Approved_Count": 0,
                "Remediation_Required": 0,
                "Avg_QA_Score_Pct": "N/A",
                "Max_Similarity_Pct": "N/A",
                "Max_AI_Content_Pct": "N/A",
                "Category_Status": "NOT UPLOADED",
            })
            continue
        app_cnt = sum(1 for r in subset if r["audit_status"] == "Approved" and r["q2_passed"])
        rem_cnt = len(subset) - app_cnt
        avg_sc = round(sum(r["audit_score"] for r in subset) / len(subset), 1)
        sims = [r["similarity_pct"] for r in subset if r["similarity_pct"] is not None]
        ais = [r["ai_content_pct"] for r in subset if r["ai_content_pct"] is not None]
        cat_stat = "READY FOR RELEASE" if rem_cnt == 0 else f"ACTION REQUIRED ({rem_cnt} Pending)"
        cat_summary_rows.append({
            "Deliverable_Category": cat,
            "Files_Evaluated": len(subset),
            "Approved_Count": app_cnt,
            "Remediation_Required": rem_cnt,
            "Avg_QA_Score_Pct": f"{avg_sc:.1f}%",
            "Max_Similarity_Pct": f"{max(sims):.1f}%" if sims else "Logged in Sheet",
            "Max_AI_Content_Pct": f"{max(ais):.1f}%" if ais else "Logged in Sheet",
            "Category_Status": cat_stat,
        })

    total_files = len(parsed_reports)
    total_approved = sum(1 for r in parsed_reports if r["audit_status"] == "Approved" and r["q2_passed"])
    composite_score = round(sum(r["audit_score"] for r in parsed_reports) / total_files, 1) if total_files else 0.0
    subject_ready = (total_files > 0 and total_approved == total_files and composite_score >= 85.0)
    subject_verdict = "APPROVED FOR SUBJECT RELEASE" if subject_ready else f"HOLD — {total_files - total_approved} DELIVERABLE(S) REQUIRE REMEDIATION"

    # --------------------------------------------------------------------------
    # 1. BUILD MULTI-SHEET EXCEL DOSSIER (.xlsx)
    # --------------------------------------------------------------------------
    wb = openpyxl.Workbook()
    hdr_fill = PatternFill(start_color=p_hex.lstrip("#"), end_color=p_hex.lstrip("#"), fill_type="solid")
    sub_fill = PatternFill(start_color=_tint_hex(p_hex, 0.88), end_color=_tint_hex(p_hex, 0.88), fill_type="solid")
    hdr_font = Font(color="FFFFFF", bold=True, size=10)
    bold_font = Font(bold=True, size=10)
    thin_side = Side(style="thin", color="CBD5E1")
    grid_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    def _style_table_sheet(ws, start_row: int, headers: List[str], data_rows: List[List]):
        ws.append([])
        for c_i, h_text in enumerate(headers, start=1):
            cell = ws.cell(row=start_row, column=c_i, value=h_text)
            cell.fill = hdr_fill
            cell.font = hdr_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = grid_border
        for r_offset, row_vals in enumerate(data_rows, start=1):
            curr_r = start_row + r_offset
            for c_i, val in enumerate(row_vals, start=1):
                cell = ws.cell(row=curr_r, column=c_i, value=val)
                cell.alignment = Alignment(vertical="center", wrap_text=True)
                cell.border = grid_border
        for col in ws.columns:
            max_len = max((len(str(c.value or "")) for c in col), default=12)
            ws.column_dimensions[col[0].column_letter].width = min(max(max_len + 4, 14), 48)

    # Tab 1: Executive Subject Summary
    ws1 = wb.active
    ws1.title = "1_Subject_Executive_Summary"
    ws1["A1"] = f"{uni_name} — {school_name}"
    ws1["A1"].font = Font(bold=True, size=13, color=p_hex.lstrip("#"))
    ws1["A2"] = f"SUBJECT-WIDE CUMULATIVE QA & ENRICHMENT REPORT: {subj_code} — {subj_name}"
    ws1["A2"].font = Font(bold=True, size=11)
    ws1["A3"] = f"Program: {prog_name}  |  Cohort: {acad_term}  |  Composite Score: {composite_score:.1f}%  |  Verdict: {subject_verdict}"
    ws1["A3"].font = bold_font
    ws1["A4"] = f"Verified By: {reviewer_email}  |  Generated: {datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M UTC')}"

    cat_headers = [
        "Deliverable Type", "Files Evaluated", "Approved Count", "Remediation Needed",
        "Avg QA Score", "Max Similarity", "Max AI Content", "Category Readiness"
    ]
    cat_matrix = [[r[k] for k in [
        "Deliverable_Category", "Files_Evaluated", "Approved_Count", "Remediation_Required",
        "Avg_QA_Score_Pct", "Max_Similarity_Pct", "Max_AI_Content_Pct", "Category_Status"
    ]] for r in cat_summary_rows]
    _style_table_sheet(ws1, start_row=6, headers=cat_headers, data_rows=cat_matrix)

    # Tab 2: Deliverable-by-Deliverable Inventory
    ws2 = wb.create_sheet(title="2_Deliverable_Inventory")
    inv_headers = [
        "Category", "Report Filename", "Audit Status", "Audit Score (%)",
        "Similarity (%)", "AI Content (%)", "Quadrant-II Gate", "Auditor", "Layer-5 Action State"
    ]
    inv_rows = []
    for r in parsed_reports:
        inv_rows.append([
            r["deliverable_type"],
            r["report_filename"],
            r["audit_status"],
            r["audit_score"],
            f"{r['similarity_pct']:.1f}%" if r["similarity_pct"] is not None else "N/A",
            f"{r['ai_content_pct']:.1f}%" if r["ai_content_pct"] is not None else "N/A",
            "PASS" if r["q2_passed"] else "FAIL",
            r["auditor_name"],
            r["l5_enrichment_state"],
        ])
    _style_table_sheet(ws2, start_row=1, headers=inv_headers, data_rows=inv_rows)

    # Tab 3: All Evaluated Parameters Across Subject
    ws3 = wb.create_sheet(title="3_Evaluated_Parameters")
    param_headers = ["Deliverable Type", "Source Report", "Evaluated Parameter / Criterion", "Score / Value", "Status", "Remarks"]
    param_rows = []
    for r in parsed_reports:
        for p in r["evaluated_parameters"]:
            param_rows.append([
                p.get("Deliverable", r["deliverable_type"]),
                p.get("Source_Report", r["report_filename"]),
                p.get("Parameter", ""),
                p.get("Score_or_Value", ""),
                p.get("Status", ""),
                p.get("Remarks", ""),
            ])
    _style_table_sheet(ws3, start_row=1, headers=param_headers, data_rows=param_rows or [["ALL", "N/A", "All Subject Parameters", f"{composite_score:.1f}%", subject_verdict, ""]])

    # Tab 4: Consolidated Remediation Backlog
    ws4 = wb.create_sheet(title="4_Defect_Remediation_Queue")
    def_headers = ["Deliverable Type", "Source Report", "Flagged Parameter / Defect", "Required Layer-5 Track"]
    def_rows = []
    for r in parsed_reports:
        for d in r["flagged_defects"]:
            def_rows.append([r["deliverable_type"], r["report_filename"], d, "Track A: AI Auto-Remediation"])
    _style_table_sheet(
        ws4,
        start_row=1,
        headers=def_headers,
        data_rows=def_rows or [["ALL", "All Uploaded Reports", "Zero open defects flagged across subject deliverables.", "Track B: Creative Enrichment Ready"]],
    )

    xl_buf = io.BytesIO()
    wb.save(xl_buf)

    # --------------------------------------------------------------------------
    # 2. BUILD PRINTABLE SUBJECT CUMULATIVE CERTIFICATE (.docx)
    # --------------------------------------------------------------------------
    p_rgb = RGBColor(*_hex_to_rgb_tuple(p_hex))
    a_rgb = RGBColor(*_hex_to_rgb_tuple(a_hex))

    cert = docx.Document()
    p_lg = cert.add_paragraph()
    p_lg.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_lg.add_run().add_picture(io.BytesIO(logo_bytes), width=Inches(1.1))

    p_u = cert.add_paragraph()
    p_u.alignment = WD_ALIGN_PARAGRAPH.CENTER
    ru = p_u.add_run(uni_name)
    ru.bold = True
    ru.font.size = Pt(16)
    ru.font.color.rgb = p_rgb

    p_s = cert.add_paragraph()
    p_s.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rs = p_s.add_run(school_name)
    rs.bold = True
    rs.font.size = Pt(11)

    h_main = cert.add_heading(f"Subject-Wide Cumulative QA & Enrichment Dossier", level=1)
    h_main.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h_main.runs[0].font.color.rgb = p_rgb

    meta_tbl = cert.add_table(rows=3, cols=2)
    meta_tbl.style = "Table Grid"
    meta_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_pairs = [
        ("Program & Cohort", f"{prog_name} ({acad_term})"),
        ("Course Code & Subject", f"{subj_code} — {subj_name}"),
        ("Composite Score & Verdict", f"{composite_score:.1f}%  |  {subject_verdict}"),
    ]
    for idx_m, (mk, mv) in enumerate(meta_pairs):
        _set_cell_shading_and_border(meta_tbl.cell(idx_m, 0), _tint_hex(p_hex, 0.88))
        meta_tbl.cell(idx_m, 0).paragraphs[0].add_run(mk).bold = True
        meta_tbl.cell(idx_m, 1).paragraphs[0].add_run(mv).bold = (idx_m == 2)

    h_cat = cert.add_heading("1. Multi-Deliverable Readiness Matrix (SLM, MCQ, QB, PPT, FAQ)", level=2)
    h_cat.runs[0].font.color.rgb = p_rgb
    ctbl = cert.add_table(rows=1, cols=6)
    ctbl.style = "Table Grid"
    ctbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_hdrs = ["Asset Type", "Files", "Approved", "Avg Score", "Max Similarity", "Readiness Status"]
    for ci, ch in enumerate(c_hdrs):
        _set_cell_shading_and_border(ctbl.cell(0, ci), _tint_hex(p_hex, 0.85))
        ctbl.cell(0, ci).paragraphs[0].add_run(ch).bold = True
    for crow in cat_summary_rows:
        rc = ctbl.add_row().cells
        rc[0].paragraphs[0].add_run(crow["Deliverable_Category"]).bold = True
        rc[1].text = str(crow["Files_Evaluated"])
        rc[2].text = str(crow["Approved_Count"])
        rc[3].text = str(crow["Avg_QA_Score_Pct"])
        rc[4].text = str(crow["Max_Similarity_Pct"])
        rc[5].paragraphs[0].add_run(crow["Category_Status"]).bold = True

    h_inv = cert.add_heading("2. Evaluated Subject Deliverables & Layer-5 Routing Log", level=2)
    h_inv.runs[0].font.color.rgb = p_rgb
    itbl = cert.add_table(rows=1, cols=5)
    itbl.style = "Table Grid"
    itbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    i_hdrs = ["Type", "Audit Report", "Score", "Verdict", "Layer-5 State"]
    for ci, ih in enumerate(i_hdrs):
        _set_cell_shading_and_border(itbl.cell(0, ci), _tint_hex(p_hex, 0.85))
        itbl.cell(0, ci).paragraphs[0].add_run(ih).bold = True
    for r in parsed_reports:
        rc = itbl.add_row().cells
        rc[0].paragraphs[0].add_run(r["deliverable_type"]).bold = True
        rc[1].text = r["report_filename"]
        rc[2].text = f"{r['audit_score']:.1f}%"
        rc[3].text = r["audit_status"]
        rc[4].text = r["l5_enrichment_state"]

    p_sig = cert.add_paragraph()
    p_sig.paragraph_format.space_before = Pt(14)
    r_sg = p_sig.add_run(
        f"IQAC / Layer-5 Sign-Off: {reviewer_email}  |  "
        f"{datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M UTC')}"
    )
    r_sg.bold = True
    r_sg.font.color.rgb = a_rgb

    doc_buf = io.BytesIO()
    cert.save(doc_buf)

    summary_dict = {
        "total_files": total_files,
        "total_approved": total_approved,
        "composite_score": composite_score,
        "subject_ready": subject_ready,
        "subject_verdict": subject_verdict,
        "category_rows": cat_summary_rows,
        "deliverable_rows": inv_rows,
        "parameter_rows": param_rows,
        "defect_rows": def_rows,
    }
    return xl_buf.getvalue(), doc_buf.getvalue(), summary_dict


# ==============================================================================
# 4. DETERMINISTIC PROSE BURSTINESS & SLIDING-WINDOW LOW-CV DETECTOR
# ==============================================================================

def _extract_prose_paragraphs(text: str) -> List[str]:
    """Extracts running prose paragraphs (>= 12 words), excluding references and lists."""
    body = re.split(r"(?im)^(?:References|Bibliography|Further Reading|Works Cited)\s*$", text)[0]
    body = re.sub(r'"[^"]{25,}"', " ", body)
    paragraphs = []
    for raw_line in body.splitlines():
        ln = raw_line.strip()
        if len(ln.split()) < 12:
            continue
        if re.match(r"^(\d+[\.\)]|[a-dA-D][\.\)]|[\-\*\u2022])\s+", ln):
            continue
        paragraphs.append(ln)
    return paragraphs


def compute_prose_cv_3dec(prose_paragraphs: List[str]) -> Dict:
    """Computes prose-only Burstiness CV rounded to 3 decimals (>= 0.400 gate)."""
    sentence_lengths = []
    for para in prose_paragraphs:
        for sent in re.split(r"(?<=[.!?])\s+", para.strip()):
            words = re.findall(r"\b[A-Za-z0-9'-]+\b", sent)
            if len(words) >= 3:
                sentence_lengths.append(len(words))

    count = len(sentence_lengths)
    if count < 2:
        return {"sentence_count": count, "mean_words": 0.0, "sd_words": 0.0, "prose_cv": 0.0, "passed": False}

    mean_len = sum(sentence_lengths) / count
    variance = sum((x - mean_len) ** 2 for x in sentence_lengths) / (count - 1)
    sd_len = math.sqrt(variance)
    cv_3 = round(sd_len / mean_len, 3) if mean_len > 0 else 0.0

    l4_cv = cv_3
    if _l4_compute_burstiness is not None:
        try:
            l4_res = _l4_compute_burstiness(prose_paragraphs)
            l4_cv = float(l4_res.get("prose_cv", cv_3))
        except Exception:
            l4_cv = cv_3

    return {
        "sentence_count": count,
        "mean_words": round(mean_len, 1),
        "sd_words": round(sd_len, 1),
        "prose_cv": cv_3,
        "l4_core_cv": l4_cv,
        "passed": (cv_3 >= 0.400 and l4_cv >= 0.40),
    }


def detect_low_cv_windows(unit_text: str, window_size: int = 4, min_cv: float = 0.400) -> List[Dict]:
    """Scans prose paragraphs in sliding windows to locate the lowest-CV passages (< 0.400)."""
    paras = _extract_prose_paragraphs(unit_text)
    if not paras:
        return []

    if len(paras) <= window_size:
        metrics = compute_prose_cv_3dec(paras)
        excerpt = "\n\n".join(paras)
        return [{
            "window_id": "W1",
            "para_range": f"Paragraphs 1–{len(paras)}",
            "cv": metrics["prose_cv"],
            "sentence_count": metrics["sentence_count"],
            "excerpt": excerpt,
        }] if metrics["prose_cv"] < min_cv else []

    flagged = []
    step = max(1, window_size // 2)
    for start in range(0, len(paras) - window_size + 1, step):
        chunk = paras[start : start + window_size]
        metrics = compute_prose_cv_3dec(chunk)
        if metrics["sentence_count"] >= 4 and metrics["prose_cv"] < min_cv:
            flagged.append({
                "window_id": f"W{len(flagged) + 1}",
                "para_range": f"Paragraphs {start + 1}–{start + len(chunk)}",
                "cv": metrics["prose_cv"],
                "sentence_count": metrics["sentence_count"],
                "excerpt": "\n\n".join(chunk),
            })

    flagged.sort(key=lambda x: x["cv"])
    return flagged[:3]


# ==============================================================================
# 5. SELF-RECOGNIZING DOCUMENT TOPOLOGY & MULTI-UNIT SEGMENTER
# ==============================================================================

_BASE_BOILERPLATE_WORDS = {
    "the", "this", "that", "these", "those", "however", "therefore", "although",
    "after", "before", "during", "because", "since", "when", "where", "what",
    "which", "another", "some", "many", "various", "several", "such", "each",
    "every", "both", "unlike", "despite", "through", "without", "within", "under",
    "over", "between", "among", "into", "from", "with", "about", "above", "below",
    "example", "examples", "note", "notes", "table", "figure", "chapter", "section",
    "unit", "module", "block", "lesson", "part", "slide", "step", "basis", "source",
    "nature", "methods", "types", "factors", "sources", "importance", "advantages",
    "limitations", "characteristics", "applications", "properties", "procedure",
    "interpretation", "background", "situation", "analysis", "outcome", "summary",
    "keywords", "glossary", "references", "bibliography", "learning", "objectives",
    "outcomes", "structure", "suppose", "consider", "researchers", "research",
    "definition", "purpose", "meaning", "formula", "components", "features",
    "benefits", "challenges", "overview", "introduction", "conclusion", "significance",
    "self-assessment", "assessment", "questions", "case", "study", "caselet", "its",
    "key", "points", "explanation", "detailed", "overall", "major", "common",
    "important", "basic", "advanced", "working", "process", "principle", "strengths",
    "recommendations", "discussion",
}

_UNIT_HEADER_RE = re.compile(
    r"^(?:Unit|Module|Chapter|Block|Lesson|Part)\s+(?:\d+|[IVXLCDM]+)(?:\s*[:\-–—.]\s*.*)?$",
    re.IGNORECASE,
)
_SLIDE_HEADER_RE = re.compile(r"^Slide\s+(\d+)\s*[:\-–—]\s*(.+)$", re.IGNORECASE)
_SEC_HEADER_RE = re.compile(r"^(\d+(?:\.\d+)+)\s+([A-Za-z][^\n]{2,80})$")
_BACK_MATTER_RE = re.compile(
    r"^(?:Slide\s+\d+\s*[:\-–—]\s*)?(?:\d+(?:\.\d+)*\s+)?(?:Detailed\s+Summary|Summary|Keywords|Glossary|Self-Assessment|Check Your Progress|Case Study|Caselet|References|Bibliography)\b",
    re.IGNORECASE,
)


def _detect_dynamic_boilerplate_labels(lines: List[str]) -> set:
    short_counts = Counter()
    for raw in lines:
        clean = raw.strip().rstrip(":")
        words = clean.split()
        if 1 <= len(words) <= 5 and not clean.endswith((".", "?", "!")) and not re.search(r"\d", clean):
            short_counts[clean.lower()] += 1
    learned = set(_BASE_BOILERPLATE_WORDS)
    for phrase, cnt in short_counts.items():
        if cnt >= 3:
            for tok in re.split(r"[\s\-/]+", phrase):
                if len(tok) >= 2:
                    learned.add(tok.lower())
    return learned


def _extract_dynamic_domain_concepts(text: str, extra_stopwords: set, limit: int = 8) -> List[str]:
    clean_lines = []
    for raw in text.splitlines():
        ln = raw.strip()
        if not ln:
            continue
        if _UNIT_HEADER_RE.match(ln) or _SLIDE_HEADER_RE.match(ln) or _SEC_HEADER_RE.match(ln) or _BACK_MATTER_RE.match(ln):
            continue
        clean_lines.append(ln)

    joined = "\n".join(clean_lines)
    phrases = re.findall(r"\b[A-Z][a-zA-Z0-9]+(?:[ -]+[A-Z][a-zA-Z0-9]+){1,2}\b", joined)
    concepts = []
    for ph in phrases:
        tokens = [t.lower() for t in re.split(r"[ -]+", ph) if t]
        if not tokens:
            continue
        if tokens[0] in extra_stopwords or tokens[-1] in extra_stopwords or all(t in extra_stopwords for t in tokens):
            continue
        if ph not in concepts:
            concepts.append(ph)
        if len(concepts) >= limit:
            break

    if len(concepts) < limit:
        singles = re.findall(r"\b[A-Z][a-zA-Z0-9]{4,}\b", joined)
        for tok in singles:
            if tok.lower() not in extra_stopwords and tok not in concepts:
                concepts.append(tok)
            if len(concepts) >= limit:
                break

    return concepts or ["Core Framework", "System Architecture", "Analytical Method", "Validation Process"]


def analyze_document_topology(unit_text: str, filename: str = "", docx_bytes: bytes = b"") -> Dict:
    fn = (filename or "").lower()
    if fn.endswith(".pptx"):
        lines = [ln.strip() for ln in unit_text.splitlines() if ln.strip()]
        dyn_stop = _detect_dynamic_boilerplate_labels(lines)
        concepts = _extract_dynamic_domain_concepts(unit_text, dyn_stop, limit=8)
        slide_headers = [ln for ln in lines if ln.startswith("--- Slide ")]
        return {
            "archetype": "NATIVE_PPTX",
            "archetype_label": "Native PowerPoint Deck (.pptx)",
            "unit_count": max(1, len(slide_headers)),
            "units": [{"unit_index": 1, "unit_label": "Slide Deck", "title": filename, "start_para": 0, "body_start_para": 0, "body_end_para": 0, "end_para": 0}],
            "has_slides": True,
            "domain_concepts": concepts,
            "section_titles": slide_headers[:8],
            "dynamic_stopwords": dyn_stop,
        }

    if docx_bytes:
        doc = docx.Document(io.BytesIO(docx_bytes))
        para_texts = [(p.text or "").strip() for p in doc.paragraphs]
    else:
        para_texts = [ln.strip() for ln in unit_text.splitlines()]

    non_empty = [t for t in para_texts if t]
    dyn_stop = _detect_dynamic_boilerplate_labels(non_empty)
    has_slides = sum(1 for t in non_empty if _SLIDE_HEADER_RE.match(t)) >= 3

    raw_unit_idxs = []
    for idx, txt in enumerate(para_texts):
        if _UNIT_HEADER_RE.match(txt) and len(txt.split()) <= 14:
            non_empty_after = [para_texts[j] for j in range(idx + 1, min(idx + 6, len(para_texts))) if para_texts[j]]
            toc_like = sum(1 for x in non_empty_after[:3] if _UNIT_HEADER_RE.match(x) or _SEC_HEADER_RE.match(x))
            if toc_like < 2:
                raw_unit_idxs.append(idx)

    units = []
    if len(raw_unit_idxs) >= 2:
        for u_pos, start_idx in enumerate(raw_unit_idxs):
            end_idx = (raw_unit_idxs[u_pos + 1] - 1) if (u_pos + 1 < len(raw_unit_idxs)) else (len(para_texts) - 1)
            raw_lbl = para_texts[start_idx]
            m_lbl = re.match(r"^((?:Unit|Module|Chapter|Block|Lesson|Part)\s+(?:\d+|[IVXLCDM]+))", raw_lbl, re.IGNORECASE)
            short_lbl = m_lbl.group(1).title() if m_lbl else f"Unit {u_pos + 1}"
            units.append({"unit_index": u_pos + 1, "unit_label": short_lbl, "title": raw_lbl, "start_para": start_idx, "end_para": end_idx})
    else:
        first_idx = raw_unit_idxs[0] if raw_unit_idxs else 0
        raw_lbl = para_texts[first_idx] if first_idx < len(para_texts) and para_texts[first_idx] else (filename or "Document Body")
        m_lbl = re.match(r"^((?:Unit|Module|Chapter|Block|Lesson|Part)\s+(?:\d+|[IVXLCDM]+))", raw_lbl, re.IGNORECASE)
        short_lbl = m_lbl.group(1).title() if m_lbl else "Unit 1"
        units.append({"unit_index": 1, "unit_label": short_lbl, "title": raw_lbl[:70], "start_para": first_idx, "end_para": max(0, len(para_texts) - 1)})

    for u in units:
        u_start, u_end = u["start_para"], u["end_para"]
        b_start = u_start
        for idx in range(u_start + 1, u_end + 1):
            txt = para_texts[idx]
            if not txt:
                continue
            if _SEC_HEADER_RE.match(txt):
                following = [para_texts[j] for j in range(idx + 1, min(idx + 5, u_end + 1)) if para_texts[j]]
                if any(len(f.split()) >= 15 and not _SEC_HEADER_RE.match(f) for f in following):
                    b_start = idx
                    break
            elif _SLIDE_HEADER_RE.match(txt):
                b_start = idx
                break
            elif len(txt.split()) >= 25 and not txt.startswith(("•", "-", "*")):
                b_start = max(u_start + 1, idx - 1)
                break

        b_end = u_end
        for idx in range(b_start + 3, u_end + 1):
            txt = para_texts[idx]
            if _BACK_MATTER_RE.match(txt) and len(txt.split()) <= 12:
                b_end = max(b_start + 1, idx - 1)
                break

        u["body_start_para"] = b_start
        u["body_end_para"] = b_end

    if len(units) >= 2 and has_slides:
        archetype, archetype_label = "MULTI_UNIT_SLIDE_COMPENDIUM", f"Multi-Unit Slide Compendium ({len(units)} Units)"
    elif len(units) >= 2:
        archetype, archetype_label = "MULTI_UNIT_SLM", f"Multi-Unit SLM Manuscript ({len(units)} Units)"
    elif has_slides:
        archetype, archetype_label = "SLIDE_SCRIPT_DOCX", "Slide-Deck Script (.docx)"
    elif any(_SEC_HEADER_RE.match(t) for t in non_empty):
        archetype, archetype_label = "SINGLE_UNIT_SLM", "Structured Single-Unit SLM"
    else:
        archetype, archetype_label = "GENERAL_ACADEMIC_DOC", "Academic / Technical Manuscript"

    sec_titles = []
    for t in non_empty:
        m_sec = _SEC_HEADER_RE.match(t)
        m_sld = _SLIDE_HEADER_RE.match(t)
        if m_sec:
            lbl = f"{m_sec.group(1)} {m_sec.group(2).strip()}"
            if lbl not in sec_titles:
                sec_titles.append(lbl)
        elif m_sld and not _BACK_MATTER_RE.match(t):
            lbl = m_sld.group(2).strip()
            if lbl not in sec_titles:
                sec_titles.append(lbl)

    domain_concepts = _extract_dynamic_domain_concepts(unit_text, dyn_stop, limit=8)
    return {
        "archetype": archetype,
        "archetype_label": archetype_label,
        "unit_count": len(units),
        "units": units,
        "has_slides": has_slides,
        "domain_concepts": domain_concepts,
        "section_titles": sec_titles[:12],
        "dynamic_stopwords": dyn_stop,
    }


def _build_balanced_context_sample(unit_text: str, topology: Dict, max_chars: int = 14000) -> str:
    if len(unit_text) <= max_chars:
        return unit_text
    lines = unit_text.splitlines()
    units = topology.get("units", [])
    if len(units) >= 2 and len(lines) > 20:
        per_unit_budget = max(800, max_chars // min(len(units), 8))
        step = max(1, len(units) // min(len(units), 8))
        sampled_parts = []
        for u in units[::step][:8]:
            chunk = "\n".join(lines[u["start_para"] : u["end_para"] + 1])
            sampled_parts.append(chunk[:per_unit_budget])
        return "\n\n[...]\n\n".join(sampled_parts)[:max_chars]
    head = int(max_chars * 0.70)
    tail = max_chars - head
    return unit_text[:head] + "\n\n[... middle section sampled ...]\n\n" + unit_text[-tail:]


# ==============================================================================
# 6. MULTI-PROVIDER REST LLM HELPER + DYNAMIC OFFLINE SCAFFOLDING
# ==============================================================================

def _collect_secret_keys(provider: str, override_key: str = "") -> List[str]:
    raw = []
    if override_key:
        raw.extend(re.split(r"[\s,;]+", override_key.strip()))
    secret_map = {
        "Google Gemini": ("GEMINI_API_KEYS", "GEMINI_API_KEY"),
        "OpenAI (ChatGPT)": ("OPENAI_API_KEYS", "OPENAI_API_KEY"),
        "Anthropic (Claude)": ("ANTHROPIC_API_KEYS", "ANTHROPIC_API_KEY"),
    }
    if provider in secret_map:
        for sk in secret_map[provider]:
            try:
                val = st.secrets.get(sk, "")
                if isinstance(val, (list, tuple)):
                    for item in val:
                        raw.extend(re.split(r"[\s,;]+", str(item).strip()))
                elif val:
                    raw.extend(re.split(r"[\s,;]+", str(val).strip()))
            except Exception:
                pass
    deduped = []
    for k in raw:
        c = k.strip().strip('"').strip("'")
        if c and c not in deduped:
            deduped.append(c)
    return deduped


def _extract_json_safe(raw_text: str) -> Dict:
    cleaned = raw_text.strip()
    if cleaned.startswith("```"):
        cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned)
        cleaned = re.sub(r"\s*```$", "", cleaned)
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        start, end = cleaned.find("{"), cleaned.rfind("}")
        if start != -1 and end > start:
            return json.loads(cleaned[start : end + 1])
        raise


def _discover_live_gemini_models(api_key: str, primary_model: str) -> List[str]:
    try:
        url = f"https://generativelanguage.googleapis.com/v1beta/models?key={api_key.strip()}"
        resp = requests.get(url, timeout=15)
        if resp.status_code != 200:
            return []
        live = []
        for m in resp.json().get("models", []):
            name = m.get("name", "").replace("models/", "")
            if "generateContent" not in m.get("supportedGenerationMethods", []) or not name:
                continue
            low = name.lower()
            if any(s in low for s in ("embedding", "imagen", "veo", "tts", "audio", "robotics", "aqa")):
                continue
            if "gemini" in low and name != primary_model:
                live.append(name)
        flash = sorted([x for x in live if "flash" in x and "lite" not in x], reverse=True)
        lite = sorted([x for x in live if "flash-lite" in x], reverse=True)
        pro = sorted([x for x in live if "pro" in x], reverse=True)
        ordered = []
        for item in flash + lite + pro:
            if item not in ordered:
                ordered.append(item)
        return ordered[:2]
    except Exception:
        return []


def _build_dynamic_offline_payload(user_prompt: str, topology: Optional[Dict] = None) -> Dict:
    ms_split = re.split(r"MANUSCRIPT (?:EXCERPT|CONTEXT):\s*", user_prompt, maxsplit=1)
    ms_body = (ms_split[1] if len(ms_split) > 1 else user_prompt).split("Return ONLY valid JSON")[0]

    if topology:
        domain_terms = list(topology.get("domain_concepts", []))
        sec_titles = list(topology.get("section_titles", []))
        units = topology.get("units", [])
    else:
        dyn_stop = _detect_dynamic_boilerplate_labels(ms_body.splitlines())
        domain_terms = _extract_dynamic_domain_concepts(ms_body, dyn_stop, limit=6)
        sec_titles, units = [], []

    if len(domain_terms) < 4:
        domain_terms.extend(["Core Architecture", "Analytical Framework", "Operational Workflow", "Evaluation Criteria"])

    t0, t1, t2, t3 = domain_terms[0], domain_terms[1], domain_terms[2], domain_terms[3]
    s1 = sec_titles[0] if len(sec_titles) > 0 else (units[0]["unit_label"] if units else "Section 1")
    s2 = sec_titles[1] if len(sec_titles) > 1 else (units[1]["unit_label"] if len(units) > 1 else s1)
    s3 = sec_titles[2] if len(sec_titles) > 2 else s2

    anchor_matches = re.findall(r"Anchor #(\d+)\s+\[([^\]]+)\]", user_prompt)
    if anchor_matches:
        anchors = [(int(idx), ctx) for idx, ctx in anchor_matches]
    else:
        raw_nums = [int(x) for x in re.findall(r"Anchor #(\d+)", user_prompt)]
        anchors = [(n, s1) for n in raw_nums] or [(18, s1), (36, s2), (54, s3)]

    callouts, flowcharts, qr_hooks = [], [], []
    styles = ["wheel", "pillars", "flowchart"]
    for i, (a_idx, a_ctx) in enumerate(anchors[:9]):
        ca = domain_terms[i % len(domain_terms)]
        cb = domain_terms[(i + 1) % len(domain_terms)]
        cc = domain_terms[(i + 2) % len(domain_terms)]
        cd = domain_terms[(i + 3) % len(domain_terms)]
        if i % 3 == 0:
            callouts.append({
                "anchor_index": a_idx,
                "context_label": a_ctx,
                "category": "Key Takeaway",
                "title": f"Core Principle — {ca}",
                "body": f"Verify the foundational assumptions linking {ca} and {cb} before scaling to complex scenarios.",
            })
        elif i % 3 == 1:
            d_style = styles[(i // 3) % len(styles)]
            flowcharts.append({
                "anchor_index": a_idx,
                "context_label": a_ctx,
                "diagram_style": d_style,
                "title": f"{ca} Conceptual Framework",
                "mermaid_code": f"graph TD; A[{ca[:24]}] --> B[{cb[:24]}]; B --> C[{cc[:24]}]; C --> D[{cd[:24]}];",
            })
        else:
            qr_hooks.append({
                "anchor_index": a_idx,
                "context_label": a_ctx,
                "hook_type": "Micro-Quiz Recap",
                "title": f"60-Second Check: {ca[:34]}",
                "prompt_payload": f"Self-Check ({a_ctx}): Compare {ca} and {cb} and state one practical trade-off in two sentences.",
            })

    if not callouts:
        callouts.append({
            "anchor_index": anchors[0][0],
            "context_label": anchors[0][1],
            "category": "Key Takeaway",
            "title": f"Core Principle — {t0}",
            "body": f"Verify the foundational assumptions linking {t0} and {t1} before applying advanced procedures.",
        })
    if not flowcharts:
        af = anchors[min(1, len(anchors) - 1)]
        flowcharts.append({
            "anchor_index": af[0],
            "context_label": af[1],
            "diagram_style": "wheel",
            "title": f"Components of {t0}",
            "mermaid_code": f"graph TD; A[{t0[:24]}] --> B[{t1[:24]}]; B --> C[{t2[:24]}]; C --> D[{t3[:24]}];",
        })
    if not qr_hooks:
        aq = anchors[min(2, len(anchors) - 1)]
        qr_hooks.append({
            "anchor_index": aq[0],
            "context_label": aq[1],
            "hook_type": "Micro-Quiz Recap",
            "title": f"60-Second Check: {t2[:34]}",
            "prompt_payload": f"Self-Check ({aq[1]}): Summarize the core role of {t2} and its practical impact in two sentences.",
        })

    a_mid = anchors[min(1, len(anchors) - 1)]
    a_end = anchors[-1]
    return {
        "rewritten_text": "",
        "editorial_notes": "Offline deterministic mode: supply an active API key for AI prose rewriting.",
        "cyp_boxes": [
            {
                "anchor_index": a_mid[0],
                "context_label": a_mid[1],
                "title": f"Check Your Progress 1 ({s1[:42]})",
                "bloom_level": "Analyze",
                "questions": [
                    f"Distinguish between {t0} and {t1} using a concrete domain scenario.",
                    f"Analyze how {t2} affects accuracy, trade-offs, and decision reliability in practice.",
                ],
                "answers": [
                    f"Model Answer 1: {t0} establishes the foundational structure, whereas {t1} operationalizes execution and evaluation.",
                    f"Model Answer 2: Rigorous handling of {t2} minimizes systematic errors and ensures consistent, verifiable outcomes.",
                ],
            }
        ],
        "caselets": [
            {
                "anchor_index": a_end[0],
                "context_label": a_end[1],
                "title": f"Applied Caselet: Integrating {t0} & {t1}",
                "background": f"An enterprise team scaling operations encountered structural bottlenecks when aligning {t0} with {t1} across distributed workflows.",
                "questions": [
                    f"Which core principles of {t0} explain the initial operational friction?",
                    f"How can {t1} and {t2} be combined to improve measurable performance?",
                ],
                "recommendations": [
                    f"Establish clear baseline metrics grounded in {t0} before deploying changes.",
                    f"Conduct iterative reviews linking {t1} to continuous feedback loops.",
                ],
                "conclusion": f"Aligning {t0} with {t1} transforms ad-hoc practices into a repeatable, auditable capability.",
            }
        ],
        "glossary": [
            {"term": term, "definition": f"Core domain concept in this module governing the structure, application, or evaluation of {term.lower()}."}
            for term in domain_terms[:6]
        ],
        "diagrams": [
            {
                "figure_id": "Figure 1.1",
                "title": f"{t0} & {t1} Conceptual Workflow",
                "anchor": f"After {s1[:45]}",
                "alt_text": f"Structured workflow connecting {t0}, {t1}, and {t2}",
                "mermaid_code": f"graph TD; A[{t0[:24]}] --> B[{t1[:24]}]; B --> C[{t2[:24]}]; C --> D[Verified Outcome];",
            }
        ],
        "callouts": callouts,
        "flowcharts": flowcharts,
        "qr_hooks": qr_hooks,
        "illustration_prompts": [],
    }


def call_layer5_llm_json(
    provider: str,
    model_id: str,
    system_instruction: str,
    user_prompt: str,
    override_api_key: str = "",
    topology: Optional[Dict] = None,
) -> Dict:
    if provider.startswith("Deterministic"):
        return _build_dynamic_offline_payload(user_prompt, topology=topology)

    keys = _collect_secret_keys(provider, override_api_key)
    if not keys:
        raise RuntimeError(f"No API key configured in st.secrets for {provider}.")

    delays = [8, 16]
    last_err = None

    for api_key in keys:
        models_to_try = [model_id]
        if provider == "Google Gemini":
            models_to_try.extend(_discover_live_gemini_models(api_key, model_id))

        for active_model in models_to_try:
            for attempt in range(len(delays) + 1):
                try:
                    if provider == "Google Gemini":
                        url = f"https://generativelanguage.googleapis.com/v1beta/models/{active_model}:generateContent?key={api_key}"
                        payload = {
                            "systemInstruction": {"parts": [{"text": system_instruction}]},
                            "contents": [{"role": "user", "parts": [{"text": user_prompt}]}],
                            "generationConfig": {"temperature": 0.45, "maxOutputTokens": 4096, "responseMimeType": "application/json"},
                        }
                        resp = requests.post(url, json=payload, timeout=(15, 60))
                        if resp.status_code != 200:
                            raise RuntimeError(f"Gemini HTTP {resp.status_code}: {resp.text[:220]}")
                        raw_txt = resp.json()["candidates"][0]["content"]["parts"][0]["text"]
                        return _extract_json_safe(raw_txt)

                    elif provider == "OpenAI (ChatGPT)":
                        url = "https://api.openai.com/v1/chat/completions"
                        headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
                        payload = {
                            "model": active_model,
                            "temperature": 0.45,
                            "max_tokens": 4096,
                            "response_format": {"type": "json_object"},
                            "messages": [{"role": "system", "content": system_instruction}, {"role": "user", "content": user_prompt}],
                        }
                        resp = requests.post(url, headers=headers, json=payload, timeout=(15, 60))
                        if resp.status_code != 200:
                            raise RuntimeError(f"OpenAI HTTP {resp.status_code}: {resp.text[:220]}")
                        raw_txt = resp.json()["choices"][0]["message"]["content"]
                        return _extract_json_safe(raw_txt)

                    elif provider == "Anthropic (Claude)":
                        url = "https://api.anthropic.com/v1/messages"
                        headers = {"x-api-key": api_key, "anthropic-version": "2023-06-01", "Content-Type": "application/json"}
                        payload = {
                            "model": active_model,
                            "max_tokens": 4096,
                            "temperature": 0.45,
                            "system": system_instruction,
                            "messages": [{"role": "user", "content": user_prompt + "\n\nRespond ONLY with valid JSON."}],
                        }
                        resp = requests.post(url, headers=headers, json=payload, timeout=(15, 60))
                        if resp.status_code != 200:
                            raise RuntimeError(f"Anthropic HTTP {resp.status_code}: {resp.text[:220]}")
                        blocks = resp.json().get("content", [])
                        raw_txt = "".join(b.get("text", "") for b in blocks if b.get("type") == "text")
                        return _extract_json_safe(raw_txt)

                    raise ValueError(f"Unsupported provider: {provider}")

                except (requests.exceptions.Timeout, requests.exceptions.ConnectionError) as net_exc:
                    last_err = f"[{active_model}] Timeout/Network stall: {net_exc}"
                    if attempt == 0:
                        time.sleep(2)
                        continue
                    break

                except Exception as exc:
                    last_err = exc
                    msg = str(exc)
                    if any(tok in msg for tok in ("401", "403", "API_KEY_INVALID", "invalid_api_key")):
                        break
                    if "exceeded your current quota" in msg.lower() or "limit: 0" in msg.lower():
                        break
                    if any(tok in msg for tok in ("429", "RESOURCE_EXHAUSTED", "rate_limit", "quota")) and attempt < len(delays):
                        time.sleep(delays[attempt])
                        continue
                    break

    raise RuntimeError(f"Layer-5 LLM call failed across all models/keys: {last_err}")


# ==============================================================================
# 7. ASPECT 1: CONTENT / TOPIC-WISE REMEDIATION & ENRICHMENT (TRACK A & B)
# ==============================================================================

def _rewrite_passage_with_cv_verification(
    excerpt: str,
    before_cv: float,
    provider: str,
    model_id: str,
    override_api_key: str = "",
    max_retries: int = 2,
    topology: Optional[Dict] = None,
) -> Dict:
    sys_msg = (
        "You are an expert Academic & Instructional Editor. Rewrite prose so it preserves 100% "
        "of technical accuracy while achieving high natural human sentence-length burstiness (CV >= 0.45). "
        "Mix crisp 5-9 word sentences with rich 22-32 word analytical sentences. Eliminate robotic transitions "
        "('Moreover', 'Furthermore', 'In conclusion', 'It is important to note'). Return ONLY JSON."
    )
    feedback_note = ""
    best_text = excerpt
    best_metrics = compute_prose_cv_3dec(_extract_prose_paragraphs(excerpt) or [excerpt])
    best_notes = ""

    for _ in range(max_retries + 1):
        usr_msg = f"""Rewrite the academic passage below to achieve Prose Burstiness CV >= 0.450.
{feedback_note}

ORIGINAL PASSAGE (Current CV: {before_cv:.3f}):
{excerpt}

Return ONLY valid JSON:
{{
  "rewritten_text": "Full rewritten prose paragraphs...",
  "editorial_notes": "What structural and cadence changes were made."
}}"""
        res = call_layer5_llm_json(provider, model_id, sys_msg, usr_msg, override_api_key, topology=topology)
        candidate = str(res.get("rewritten_text", "")).strip()
        if not candidate:
            best_notes = res.get("editorial_notes", best_notes)
            break

        cand_paras = _extract_prose_paragraphs(candidate) or [candidate]
        cand_metrics = compute_prose_cv_3dec(cand_paras)
        if cand_metrics["prose_cv"] >= best_metrics["prose_cv"]:
            best_text = candidate
            best_metrics = cand_metrics
            best_notes = res.get("editorial_notes", "")

        if cand_metrics["passed"]:
            return {
                "original": excerpt,
                "rewritten": candidate,
                "before_cv": before_cv,
                "after_cv": cand_metrics["prose_cv"],
                "qa_passed": True,
                "notes": res.get("editorial_notes", ""),
            }

        feedback_note = (
            f"PREVIOUS ATTEMPT FAILED QA GATE (CV = {cand_metrics['prose_cv']:.3f} < 0.400). "
            "Increase sentence length contrast sharply: alternate very short 4-8 word sentences with "
            "complex 24-34 word multi-clause sentences."
        )

    return {
        "original": excerpt,
        "rewritten": best_text,
        "before_cv": before_cv,
        "after_cv": best_metrics["prose_cv"],
        "qa_passed": best_metrics["passed"],
        "notes": best_notes or "Completed maximum rewrite verification passes.",
    }


def generate_track_a_fixit_items(
    unit_text: str,
    provider: str,
    model_id: str,
    override_api_key: str = "",
    progress_cb=None,
    topology: Optional[Dict] = None,
    report_defects: Optional[List[str]] = None,
) -> List[Dict]:
    """Track A: Rewrites low-CV windows and remediates specific defects cited in the Layer-4+ Audit Report."""
    topo = topology or analyze_document_topology(unit_text)
    low_windows = detect_low_cv_windows(unit_text, window_size=4, min_cv=0.400)
    sampled_text = _build_balanced_context_sample(unit_text, topo, max_chars=14000)

    if progress_cb:
        progress_cb(f"Step 1/2 [{topo['archetype_label']}]: Drafting CYP boxes, Model Answer Keys, Glossary & Diagram specs...")

    defect_directive = ""
    if report_defects:
        defect_directive = "\nSPECIFIC LAYER-4+ AUDIT DEFECTS TO REMEDIATE:\n- " + "\n- ".join(report_defects[:8]) + "\n"

    sys_msg = (
        "You are a Senior Academic & Instructional Designer. "
        "Generate concise, high-accuracy pedagogical remediation assets strictly grounded in the manuscript's actual domain "
        "and directly resolving any Layer-4+ audit defect flags. Return ONLY valid JSON."
    )
    usr_msg = f"""Analyze this {topo['archetype_label']} ({topo['unit_count']} unit(s)) and generate missing pedagogical components:
{defect_directive}
1. Two 'Check Your Progress' (CYP) self-assessment boxes aligned to Bloom's Taxonomy (Apply/Analyze) with concise Model Answer Keys.
2. A Unit Glossary table with 6 domain-specific technical terms and crisp standalone definitions.
3. Two Mermaid.js instructional diagram specifications (Figure ID, Title, Placement Anchor, Alt-Text, and valid Mermaid graph/flowchart syntax).

MANUSCRIPT EXCERPT:
{sampled_text}

Return ONLY valid JSON:
{{
  "cyp_boxes": [
    {{
      "title": "Check Your Progress 1",
      "bloom_level": "Analyze",
      "questions": ["Q1...", "Q2..."],
      "answers": ["A1...", "A2..."]
    }}
  ],
  "glossary": [
    {{"term": "Domain Term", "definition": "Crisp definition"}}
  ],
  "diagrams": [
    {{
      "figure_id": "Figure 1.1",
      "title": "Conceptual Overview",
      "anchor": "After Section 1",
      "alt_text": "Accessible description of diagram",
      "mermaid_code": "graph TD; A[Input] --> B[Process]; B --> C[Output];"
    }}
  ]
}}"""

    payload = call_layer5_llm_json(provider, model_id, sys_msg, usr_msg, override_api_key, topology=topo)
    staged: List[Dict] = []

    targets = low_windows[:2]
    for idx, win in enumerate(targets, start=1):
        if progress_cb:
            progress_cb(f"Step 2/2: Rewriting & verifying low-CV window {idx}/{len(targets)} ({win['para_range']})...")
        rw = _rewrite_passage_with_cv_verification(
            excerpt=win["excerpt"],
            before_cv=win["cv"],
            provider=provider,
            model_id=model_id,
            override_api_key=override_api_key,
            max_retries=1,
            topology=topo,
        )
        badge = "PASS >= 0.400" if rw["qa_passed"] else "NEEDS MANUAL EDIT (< 0.400)"
        staged.append({
            "id": f"rewrite_{idx}",
            "asset_type": "High-Burstiness Prose Rewrite",
            "anchor": win["para_range"],
            "summary": f"CV Delta: {rw['before_cv']:.3f} -> {rw['after_cv']:.3f} ({badge})",
            "data": rw,
            "qa_passed": rw["qa_passed"],
            "approved": rw["qa_passed"],
        })

    for idx, box in enumerate(payload.get("cyp_boxes", []), start=1):
        staged.append({
            "id": f"cyp_{idx}",
            "asset_type": "CYP Box + Model Answer Key",
            "anchor": f"End of Section {idx}",
            "summary": f"{box.get('title', f'CYP {idx}')} ({len(box.get('questions', []))} Qs + Keys)",
            "data": box,
            "qa_passed": True,
            "approved": True,
        })

    if payload.get("glossary"):
        staged.append({
            "id": "glossary_1",
            "asset_type": "Unit Glossary Table",
            "anchor": "End of Unit",
            "summary": f"{len(payload['glossary'])} domain terms defined",
            "data": payload["glossary"],
            "qa_passed": True,
            "approved": True,
        })

    for idx, diag in enumerate(payload.get("diagrams", []), start=1):
        staged.append({
            "id": f"diag_{idx}",
            "asset_type": "Mermaid.js Diagram Spec",
            "anchor": diag.get("anchor", f"Section {idx}"),
            "summary": f"{diag.get('figure_id', f'Figure {idx}')}: {diag.get('title', '')}",
            "data": diag,
            "qa_passed": True,
            "approved": True,
        })

    return staged


def scan_docx_engagement_gaps(
    docx_bytes: bytes,
    word_threshold: int = 400,
    topology: Optional[Dict] = None,
) -> List[Dict]:
    if not docx_bytes:
        return []
    doc = docx.Document(io.BytesIO(docx_bytes))
    paras = doc.paragraphs
    if not paras:
        return []

    topo = topology or analyze_document_topology(
        "\n".join((p.text or "").strip() for p in paras if (p.text or "").strip()),
        docx_bytes=docx_bytes,
    )
    units = topo.get("units", [])
    gaps: List[Dict] = []

    for u in units:
        b_start = u.get("body_start_para", 0)
        b_end = min(u.get("body_end_para", len(paras) - 1), len(paras) - 1)
        u_label = u.get("unit_label", "Body")
        curr_ctx = u_label
        running_words = 0
        unit_gaps = []
        narrative_candidates = []

        for idx in range(b_start, b_end + 1):
            para = paras[idx]
            txt = (para.text or "").strip()
            if not txt:
                continue

            m_sld = _SLIDE_HEADER_RE.match(txt)
            m_sec = _SEC_HEADER_RE.match(txt)
            if m_sld:
                curr_ctx = f"{u_label} — Slide {m_sld.group(1)}"
                running_words = 0
                continue
            if (para.style and para.style.name.startswith("Heading")) or (m_sec and len(txt.split()) <= 14):
                curr_ctx = f"{u_label} — {txt[:32]}"
                running_words = 0
                continue

            has_media = len(para._p.xpath(".//*[local-name()='blip' or local-name()='imagedata']")) > 0
            if has_media:
                running_words = 0
                continue

            wc = len(txt.split())
            if wc >= 16:
                narrative_candidates.append((idx, wc, txt, curr_ctx))

            running_words += wc
            if running_words >= word_threshold and wc >= 14:
                unit_gaps.append({"anchor_index": idx, "word_span": running_words, "context_label": curr_ctx, "preview": txt[:150]})
                running_words = 0

        if not unit_gaps and narrative_candidates:
            pick_idx, pick_wc, pick_txt, pick_ctx = narrative_candidates[len(narrative_candidates) // 2]
            unit_gaps.append({"anchor_index": pick_idx, "word_span": max(word_threshold, pick_wc), "context_label": pick_ctx, "preview": pick_txt[:150]})
        elif len(units) == 1 and len(unit_gaps) < 3 and len(narrative_candidates) >= 3:
            existing = {g["anchor_index"] for g in unit_gaps}
            for frac in (0.25, 0.55, 0.80):
                pick_idx, pick_wc, pick_txt, pick_ctx = narrative_candidates[int(len(narrative_candidates) * frac)]
                if not any(abs(pick_idx - ex) < 4 for ex in existing):
                    unit_gaps.append({"anchor_index": pick_idx, "word_span": max(word_threshold, pick_wc), "context_label": pick_ctx, "preview": pick_txt[:150]})
                    existing.add(pick_idx)
                if len(unit_gaps) >= 3:
                    break

        gaps.extend(unit_gaps)

    if len(units) >= 2 and len(gaps) > 9:
        step = len(gaps) / 9.0
        gaps = [gaps[int(i * step)] for i in range(9)]

    gaps.sort(key=lambda x: x["anchor_index"])
    return gaps


def _render_branded_infographic_png(
    mermaid_code: str,
    title: str = "Conceptual Framework",
    diagram_style: str = "flowchart",
    primary_hex: str = "1E3A8A",
    accent_hex: str = "D97706",
) -> bytes:
    labels = [lbl.strip() for lbl in re.findall(r"\[([^\]]+)\]", mermaid_code) if lbl.strip()]
    nodes = list(dict.fromkeys(labels)) or ["Core Framework", "Analytical Process", "Strategic Outcome"]

    p_rgb = _hex_to_rgb_tuple(primary_hex, (30, 58, 138))
    a_rgb = _hex_to_rgb_tuple(accent_hex, (217, 119, 6))
    tint_rgb = _hex_to_rgb_tuple(_tint_hex(primary_hex, 0.92), (239, 246, 255))
    accent_tint = _hex_to_rgb_tuple(_tint_hex(accent_hex, 0.88), (254, 243, 199))
    font = ImageFont.load_default()
    style = (diagram_style or "flowchart").lower().strip()

    if style == "wheel" and len(nodes) >= 3:
        w, h = 960, 560
        img = Image.new("RGB", (w, h), "#FFFFFF")
        draw = ImageDraw.Draw(img)
        draw.text((36, 20), title.upper()[:72], fill=p_rgb, font=font)
        draw.line([(36, 44), (w - 36, 44)], fill=(203, 213, 225), width=2)

        cx, cy = w // 2, h // 2 + 16
        hub_r = 76
        spokes = nodes[1:6] if len(nodes) > 3 else nodes[:4]
        hub_label = nodes[0][:22] if len(nodes) > 3 else "CORE MODEL"

        radius = 185
        for i, sp_lbl in enumerate(spokes):
            ang = (2 * math.pi * i / len(spokes)) - (math.pi / 2)
            sx = int(cx + radius * 1.45 * math.cos(ang))
            sy = int(cy + radius * 0.92 * math.sin(ang))
            draw.line([(cx, cy), (sx, sy)], fill=a_rgb if i % 2 else p_rgb, width=4)
            bw, bh = 210, 62
            draw.rounded_rectangle(
                [sx - bw // 2, sy - bh // 2, sx + bw // 2, sy + bh // 2],
                radius=10,
                fill=accent_tint if i % 2 else tint_rgb,
                outline=a_rgb if i % 2 else p_rgb,
                width=2,
            )
            draw.text((sx - bw // 2 + 14, sy - 6), f"{i + 1}. {sp_lbl[:24]}", fill=(15, 23, 42), font=font)

        draw.ellipse([cx - hub_r, cy - hub_r, cx + hub_r, cy + hub_r], fill=p_rgb, outline=a_rgb, width=4)
        draw.text((cx - 52, cy - 6), hub_label[:18], fill=(255, 255, 255), font=font)

        buf = io.BytesIO()
        img.save(buf, format="PNG")
        return buf.getvalue()

    if style == "pillars" and len(nodes) >= 3:
        w, h = 960, 440
        img = Image.new("RGB", (w, h), "#FFFFFF")
        draw = ImageDraw.Draw(img)
        draw.text((36, 20), title.upper()[:72], fill=p_rgb, font=font)
        draw.line([(36, 44), (w - 36, 44)], fill=(203, 213, 225), width=2)

        pillars = nodes[:5]
        margin, gap = 44, 18
        col_w = (w - 2 * margin - gap * (len(pillars) - 1)) // len(pillars)
        draw.rounded_rectangle([margin, 66, w - margin, 114], radius=8, fill=p_rgb)
        draw.text((margin + 24, 84), f"INTEGRATED FRAMEWORK: {title[:56]}", fill=(255, 255, 255), font=font)

        for i, p_lbl in enumerate(pillars):
            x0 = margin + i * (col_w + gap)
            x1 = x0 + col_w
            is_alt = (i % 2 == 1)
            draw.rounded_rectangle([x0, 134, x1, 360], radius=8, fill=accent_tint if is_alt else tint_rgb, outline=a_rgb if is_alt else p_rgb, width=2)
            draw.rectangle([x0, 134, x1, 168], fill=a_rgb if is_alt else p_rgb)
            draw.text((x0 + 12, 145), f"PILLAR 0{i + 1}", fill=(255, 255, 255), font=font)
            draw.text((x0 + 12, 195), p_lbl[:22], fill=(15, 23, 42), font=font)
            if len(p_lbl) > 22:
                draw.text((x0 + 12, 215), p_lbl, fill=(51, 65, 85), font=font)

        draw.rounded_rectangle([margin, 376, w - margin, 412], radius=6, fill=(241, 245, 249), outline=p_rgb, width=1)
        draw.text((margin + 20, 388), "APPLIED OUTCOME: ALIGNED INSTITUTIONAL & LEARNER PERFORMANCE", fill=p_rgb, font=font)

        buf = io.BytesIO()
        img.save(buf, format="PNG")
        return buf.getvalue()

    width, box_h, gap_h, top_pad = 920, 68, 48, 70
    height = top_pad + len(nodes[:5]) * box_h + max(0, len(nodes[:5]) - 1) * gap_h + 40
    img = Image.new("RGB", (width, height), "#FFFFFF")
    draw = ImageDraw.Draw(img)
    draw.text((36, 22), title.upper()[:72], fill=p_rgb, font=font)
    draw.line([(36, 46), (width - 36, 46)], fill=(203, 213, 225), width=2)

    box_w = 600
    x0 = (width - box_w) // 2
    x1 = x0 + box_w
    for idx, label in enumerate(nodes[:5]):
        y0 = top_pad + idx * (box_h + gap_h)
        y1 = y0 + box_h
        draw.rounded_rectangle([x0, y0, x1, y1], radius=8, fill=tint_rgb, outline=p_rgb, width=2)
        draw.text((x0 + 20, y0 + 26), f"{idx + 1}. {label[:66]}", fill=(15, 23, 42), font=font)
        if idx < len(nodes[:5]) - 1:
            mx = width // 2
            draw.line([(mx, y1 + 3), (mx, y1 + gap_h - 5)], fill=a_rgb, width=3)
            draw.polygon([(mx - 7, y1 + gap_h - 12), (mx + 7, y1 + gap_h - 12), (mx, y1 + gap_h - 4)], fill=a_rgb)

    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def _render_mermaid_300dpi_png(
    mermaid_code: str,
    title: str = "Conceptual Flowchart",
    diagram_style: str = "flowchart",
    primary_hex: str = "1E3A8A",
    accent_hex: str = "D97706",
) -> bytes:
    clean = mermaid_code.replace("```mermaid", "").replace("```", "").strip()
    if (diagram_style or "").lower() in ("wheel", "pillars"):
        return _render_branded_infographic_png(clean, title=title, diagram_style=diagram_style, primary_hex=primary_hex, accent_hex=accent_hex)
    try:
        b64 = base64.urlsafe_b64encode(clean.encode("utf-8")).decode("ascii")
        resp = requests.get(f"https://mermaid.ink/img/{b64}?type=png&bgColor=!white&scale=3", timeout=8)
        if resp.status_code == 200 and resp.content:
            return resp.content
    except Exception:
        pass
    return _render_branded_infographic_png(clean, title=title, diagram_style=diagram_style, primary_hex=primary_hex, accent_hex=accent_hex)


def _generate_qr_png(payload_text: str, primary_hex: str = "0F172A") -> bytes:
    qr = qrcode.QRCode(version=2, error_correction=qrcode.constants.ERROR_CORRECT_M, box_size=8, border=2)
    qr.add_data(payload_text)
    qr.make(fit=True)
    fill_hex = f"#{str(primary_hex or '0F172A').lstrip('#')}"
    img = qr.make_image(fill_color=fill_hex, back_color="#FFFFFF")
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def _generate_phase2_illustration(prompt_text: str, image_provider: str = "OpenAI (DALL-E 3)") -> bytes:
    safe_prompt = (
        "Clean flat-vector academic textbook educational diagram on pure white background, "
        "high contrast, clear visual hierarchy, no garbled text, publication quality: "
        + prompt_text[:500]
    )
    if image_provider == "Google Gemini (Imagen 3)":
        keys = _collect_secret_keys("Google Gemini")
        if not keys:
            raise RuntimeError("GEMINI_API_KEY missing in st.secrets for Imagen 3.")
        url = f"https://generativelanguage.googleapis.com/v1beta/models/imagen-3.0-generate-002:predict?key={keys[0]}"
        payload = {"instances": [{"prompt": safe_prompt}], "parameters": {"sampleCount": 1, "aspectRatio": "16:9"}}
        resp = requests.post(url, json=payload, timeout=90)
        if resp.status_code != 200:
            raise RuntimeError(f"Imagen HTTP {resp.status_code}: {resp.text[:220]}")
        return base64.b64decode(resp.json()["predictions"][0]["bytesBase64Encoded"])

    keys = _collect_secret_keys("OpenAI (ChatGPT)")
    if not keys:
        raise RuntimeError("OPENAI_API_KEY missing in st.secrets for DALL-E 3.")
    url = "https://api.openai.com/v1/images/generations"
    headers = {"Authorization": f"Bearer {keys[0]}", "Content-Type": "application/json"}
    payload = {"model": "dall-e-3", "prompt": safe_prompt, "n": 1, "size": "1792x1024", "response_format": "b64_json"}
    resp = requests.post(url, headers=headers, json=payload, timeout=90)
    if resp.status_code != 200:
        raise RuntimeError(f"DALL-E 3 HTTP {resp.status_code}: {resp.text[:220]}")
    return base64.b64decode(resp.json()["data"][0]["b64_json"])


def generate_track_b_enrichment_items(
    unit_text: str,
    docx_bytes: bytes = b"",
    provider: str = "Google Gemini",
    model_id: str = "gemini-2.5-flash",
    override_api_key: str = "",
    enable_phase2_images: bool = False,
    max_phase2_images: int = 2,
    image_provider: str = "OpenAI (DALL-E 3)",
    include_cyp_and_caselets: bool = True,
    topology: Optional[Dict] = None,
) -> List[Dict]:
    topo = topology or analyze_document_topology(unit_text, docx_bytes=docx_bytes)
    gaps = scan_docx_engagement_gaps(docx_bytes, word_threshold=400, topology=topo)
    if not gaps:
        gaps = [
            {"anchor_index": 18, "word_span": 420, "context_label": "Section 1", "preview": unit_text[:150]},
            {"anchor_index": 36, "word_span": 420, "context_label": "Section 2", "preview": unit_text[150:300]},
            {"anchor_index": 54, "word_span": 420, "context_label": "Section 3", "preview": unit_text[300:450]},
        ]

    gap_list_txt = "\n".join(
        [f"- Anchor #{g['anchor_index']} [{g.get('context_label', 'Body')}] ({g['word_span']}w prose span): {g['preview']}" for g in gaps[:9]]
    )
    sampled_text = _build_balanced_context_sample(unit_text, topo, max_chars=12000)
    a1 = gaps[0]["anchor_index"]
    a2 = gaps[min(1, len(gaps) - 1)]["anchor_index"]
    a3 = gaps[min(2, len(gaps) - 1)]["anchor_index"]

    sys_msg = (
        "You are an Academic Visual & Interactive Learning Designer. "
        "Design high-engagement visual callouts, multi-shape infographics (wheel, pillars, flowchart), "
        "QR micro-quiz hooks, in-text Check Your Progress boxes, and applied Caselets at the exact "
        "detected engagement gap anchor indices. Return ONLY valid JSON."
    )
    usr_msg = f"""Document Archetype: {topo['archetype_label']} ({topo['unit_count']} unit(s)).
Propose Track B creative enrichment assets distributed across these detected body engagement gaps:

DETECTED ENGAGEMENT GAPS:
{gap_list_txt}

MANUSCRIPT CONTEXT:
{sampled_text}

Return ONLY valid JSON:
{{
  "callouts": [
    {{
      "anchor_index": {a1},
      "category": "Key Takeaway",
      "title": "Applied Concept",
      "body": "Concise 2-3 sentence real-world applied insight grounded in this section."
    }}
  ],
  "flowcharts": [
    {{
      "anchor_index": {a2},
      "diagram_style": "wheel",
      "title": "Core Components Framework",
      "mermaid_code": "graph TD; A[Hub] --> B[Component 1]; B --> C[Component 2]; C --> D[Component 3];"
    }}
  ],
  "qr_hooks": [
    {{
      "anchor_index": {a3},
      "hook_type": "Micro-Quiz Recap",
      "title": "60-Second Concept Check",
      "prompt_payload": "Quick Self-Check: Explain the core trade-off discussed in this section in 2 sentences."
    }}
  ],
  "cyp_boxes": [
    {{
      "anchor_index": {a2},
      "title": "Check Your Progress",
      "bloom_level": "Analyze",
      "questions": ["Q1...", "Q2..."],
      "answers": ["A1...", "A2..."]
    }}
  ],
  "caselets": [
    {{
      "anchor_index": {a3},
      "title": "Applied Industry Caselet",
      "background": "Concise organizational scenario...",
      "questions": ["Analytical Q1?", "Analytical Q2?"],
      "recommendations": ["Recommendation 1", "Recommendation 2"],
      "conclusion": "Key takeaway synthesis."
    }}
  ],
  "illustration_prompts": [
    {{
      "anchor_index": {a2},
      "title": "System Architecture Diagram",
      "visual_prompt": "Clean flat-vector schematic showing..."
    }}
  ]
}}"""

    payload = call_layer5_llm_json(provider, model_id, sys_msg, usr_msg, override_api_key, topology=topo)
    anchor_ctx_map = {g["anchor_index"]: g.get("context_label", "") for g in gaps}
    staged: List[Dict] = []

    for idx, c in enumerate(payload.get("callouts", []), start=1):
        a_idx = int(c.get("anchor_index", a1))
        ctx = c.get("context_label") or anchor_ctx_map.get(a_idx, f"Paragraph #{a_idx}")
        staged.append({
            "id": f"callout_{idx}",
            "asset_type": "Branded Callout Card (₹0)",
            "anchor": f"{ctx} (Para #{a_idx})",
            "summary": f"[{c.get('category', 'Key Takeaway')}] {c.get('title', '')}",
            "data": {"kind": "callout", **c},
            "qa_passed": True,
            "approved": True,
        })

    for idx, f in enumerate(payload.get("flowcharts", []), start=1):
        a_idx = int(f.get("anchor_index", a2))
        ctx = f.get("context_label") or anchor_ctx_map.get(a_idx, f"Paragraph #{a_idx}")
        d_style = f.get("diagram_style", "flowchart").title()
        staged.append({
            "id": f"mermaid_{idx}",
            "asset_type": f"Branded {d_style} Infographic (₹0)",
            "anchor": f"{ctx} (Para #{a_idx})",
            "summary": f"{d_style}: {f.get('title', '')}",
            "data": {"kind": "mermaid", **f},
            "qa_passed": True,
            "approved": True,
        })

    for idx, q in enumerate(payload.get("qr_hooks", []), start=1):
        a_idx = int(q.get("anchor_index", a3))
        ctx = q.get("context_label") or anchor_ctx_map.get(a_idx, f"Paragraph #{a_idx}")
        staged.append({
            "id": f"qr_{idx}",
            "asset_type": "QR Micro-Quiz Hook (₹0)",
            "anchor": f"{ctx} (Para #{a_idx})",
            "summary": f"QR Hook: {q.get('title', '')}",
            "data": {"kind": "qr", **q},
            "qa_passed": True,
            "approved": True,
        })

    if include_cyp_and_caselets:
        for idx, box in enumerate(payload.get("cyp_boxes", []), start=1):
            a_idx = int(box.get("anchor_index", a2))
            ctx = box.get("context_label") or anchor_ctx_map.get(a_idx, f"Paragraph #{a_idx}")
            staged.append({
                "id": f"tb_cyp_{idx}",
                "asset_type": "In-Text Check Your Progress Box (₹0)",
                "anchor": f"{ctx} (Para #{a_idx})",
                "summary": f"{box.get('title', 'Check Your Progress')} ({len(box.get('questions', []))} Qs)",
                "data": {"kind": "cyp_inline", **box},
                "qa_passed": True,
                "approved": True,
            })

        for idx, cs in enumerate(payload.get("caselets", []), start=1):
            a_idx = int(cs.get("anchor_index", a3))
            ctx = cs.get("context_label") or anchor_ctx_map.get(a_idx, f"Paragraph #{a_idx}")
            staged.append({
                "id": f"tb_caselet_{idx}",
                "asset_type": "Structured Applied Caselet (₹0)",
                "anchor": f"{ctx} (Para #{a_idx})",
                "summary": f"Caselet: {cs.get('title', 'Applied Caselet')}",
                "data": {"kind": "caselet_inline", **cs},
                "qa_passed": True,
                "approved": True,
            })

    if enable_phase2_images:
        capped_limit = min(max(1, int(max_phase2_images)), 3)
        for idx, ill in enumerate(payload.get("illustration_prompts", [])[:capped_limit], start=1):
            img_bytes = _generate_phase2_illustration(ill.get("visual_prompt", ""), image_provider=image_provider)
            a_idx = int(ill.get("anchor_index", a2))
            ctx = anchor_ctx_map.get(a_idx, f"Paragraph #{a_idx}")
            staged.append({
                "id": f"img_p2_{idx}",
                "asset_type": f"Phase 2 Illustration ({image_provider})",
                "anchor": f"{ctx} (Para #{a_idx})",
                "summary": f"Illustration {idx}/{capped_limit}: {ill.get('title', '')}",
                "data": {"kind": "illustration", "image_bytes": img_bytes, **ill},
                "qa_passed": True,
                "approved": True,
            })

    return staged


# ==============================================================================
# 8. CLONE-ON-WRITE ARTIFACT COMPILERS (BRAND-AWARE .DOCX & NATIVE .PPTX)
# ==============================================================================

def _set_cell_shading_and_border(cell, fill_hex: str, left_border_hex: Optional[str] = None):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_hex.lstrip("#"))
    tc_pr.append(shd)

    if left_border_hex:
        borders = OxmlElement("w:tcBorders")
        left_b = OxmlElement("w:left")
        left_b.set(qn("w:val"), "single")
        left_b.set(qn("w:sz"), "24")
        left_b.set(qn("w:space"), "0")
        left_b.set(qn("w:color"), left_border_hex.lstrip("#"))
        borders.append(left_b)
        tc_pr.append(borders)


def _inject_xml_callout_after(
    anchor_para,
    category: str,
    title: str,
    body_text: str,
    primary_hex: str = "1E3A8A",
    accent_hex: str = "D97706",
):
    doc = anchor_para.part.document
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.cell(0, 0)

    is_cyp = "progress" in category.lower() or "caselet" in category.lower()
    border_hex = accent_hex if is_cyp else primary_hex
    fill_hex = _tint_hex(border_hex, 0.91)
    _set_cell_shading_and_border(cell, fill_hex, left_border_hex=border_hex)

    p_rgb = RGBColor(*_hex_to_rgb_tuple(border_hex))
    p0 = cell.paragraphs[0]
    p0.paragraph_format.space_before = Pt(4)
    p0.paragraph_format.space_after = Pt(2)
    r_cat = p0.add_run(f"{category.upper()}: {title}")
    r_cat.bold = True
    r_cat.font.size = Pt(10)
    r_cat.font.color.rgb = p_rgb

    for line in body_text.splitlines():
        if not line.strip():
            continue
        p1 = cell.add_paragraph(line.strip())
        p1.paragraph_format.space_before = Pt(2)
        p1.paragraph_format.space_after = Pt(3)
        for r in p1.runs:
            r.font.size = Pt(9.5)

    anchor_para._p.addnext(tbl._tbl)


def compile_approved_fixit_patch_docx(
    approved_items: List[Dict],
    manuscript_title: str,
    reviewer_email: str,
    brand_profile: Optional[Dict] = None,
    audit_verified: bool = True,
) -> bytes:
    bp = brand_profile or {}
    p_hex = bp.get("primary_hex", "1E3A8A")
    p_rgb = RGBColor(*_hex_to_rgb_tuple(p_hex))

    doc = docx.Document()
    doc.core_properties.title = f"Layer-5 Remediation Patch - {manuscript_title}"
    doc.core_properties.author = reviewer_email
    doc.core_properties.subject = "Academic QA Staged Content Remediation"

    if bp.get("university_name"):
        p_uni = doc.add_paragraph()
        r_uni = p_uni.add_run(f"{bp['university_name']} — {bp.get('school_name', '')}")
        r_uni.bold = True
        r_uni.font.size = Pt(10)
        r_uni.font.color.rgb = p_rgb

    status_label = "APPROVED (AUDIT VERIFIED)" if audit_verified else "PRE-AUDIT REMEDIATION DRAFT (SUBMIT TO LAYER-4+)"
    doc.add_heading(f"Layer-5 Remediation Patch — {manuscript_title}", level=1)
    p_gov = doc.add_paragraph()
    p_gov.add_run(
        f"L5_Approval_Status: {status_label}  |  L5_Approved_By: {reviewer_email}  |  "
        f"{datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M UTC')}"
    ).bold = True

    for item in approved_items:
        atype = item.get("asset_type", "")
        data = item.get("data", {})

        if atype == "High-Burstiness Prose Rewrite":
            doc.add_heading(f"Approved High-Burstiness Rewrite ({item.get('anchor', '')})", level=2)
            doc.add_paragraph(f"Verified Prose Burstiness CV: {data.get('before_cv', 0):.3f} -> {data.get('after_cv', 0):.3f}").italic = True
            doc.add_paragraph(data.get("rewritten", ""))

        elif atype == "CYP Box + Model Answer Key":
            doc.add_heading(f"{data.get('title', 'Check Your Progress')} [Bloom: {data.get('bloom_level', 'Analyze')}]", level=2)
            for i, q in enumerate(data.get("questions", []), start=1):
                doc.add_paragraph(f"Q{i}. {q}")
            doc.add_heading("Model Answer Key", level=3)
            for i, a in enumerate(data.get("answers", []), start=1):
                doc.add_paragraph(f"A{i}. {a}")

        elif atype == "Unit Glossary Table":
            doc.add_heading("Unit Glossary", level=2)
            tbl = doc.add_table(rows=1, cols=2)
            tbl.style = "Table Grid"
            _set_cell_shading_and_border(tbl.cell(0, 0), _tint_hex(p_hex, 0.85))
            _set_cell_shading_and_border(tbl.cell(0, 1), _tint_hex(p_hex, 0.85))
            tbl.cell(0, 0).paragraphs[0].add_run("Term").bold = True
            tbl.cell(0, 1).paragraphs[0].add_run("Definition").bold = True
            for entry in data:
                row = tbl.add_row().cells
                row[0].paragraphs[0].add_run(str(entry.get("term", ""))).bold = True
                row[1].text = str(entry.get("definition", ""))

        elif atype == "Mermaid.js Diagram Spec":
            doc.add_heading(f"{data.get('figure_id', 'Figure')}: {data.get('title', '')}", level=2)
            doc.add_paragraph(f"Placement Anchor: {data.get('anchor', '')}")
            doc.add_paragraph(f"Alt-Text: {data.get('alt_text', '')}")
            doc.add_paragraph(f"Mermaid.js Specification:\n{data.get('mermaid_code', '')}")

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def compile_approved_enriched_docx(
    original_docx_bytes: bytes,
    approved_items: List[Dict],
    reviewer_email: str,
    brand_profile: Optional[Dict] = None,
) -> bytes:
    bp = brand_profile or {}
    p_hex = bp.get("primary_hex", "1E3A8A")
    a_hex = bp.get("accent_hex", "D97706")

    doc = docx.Document(io.BytesIO(original_docx_bytes))
    doc.core_properties.comments = f"Layer-5 Enriched Clone approved by {reviewer_email}"

    sorted_items = sorted(approved_items, key=lambda x: int(x.get("data", {}).get("anchor_index", 0)), reverse=True)

    for item in sorted_items:
        data = item.get("data", {})
        kind = data.get("kind", "")
        idx = min(int(data.get("anchor_index", 0)), max(0, len(doc.paragraphs) - 1))
        anchor_para = doc.paragraphs[idx]

        if kind == "callout":
            _inject_xml_callout_after(anchor_para, data.get("category", "Key Takeaway"), data.get("title", "Applied Note"), data.get("body", ""), primary_hex=p_hex, accent_hex=a_hex)
        elif kind == "cyp_inline":
            q_lines = [f"Q{i}. {q}" for i, q in enumerate(data.get("questions", []), start=1)]
            _inject_xml_callout_after(anchor_para, f"Check Your Progress [{data.get('bloom_level', 'Analyze')}]", data.get("title", "Self-Check"), "\n".join(q_lines), primary_hex=p_hex, accent_hex=a_hex)
        elif kind == "caselet_inline":
            c_lines = [
                f"Background: {data.get('background', '')}",
                "Discussion Questions: " + " | ".join(data.get("questions", [])),
                "Recommendations: " + " ; ".join(data.get("recommendations", [])),
                f"Conclusion: {data.get('conclusion', '')}",
            ]
            _inject_xml_callout_after(anchor_para, "Applied Caselet", data.get("title", "Case Study"), "\n".join(c_lines), primary_hex=p_hex, accent_hex=a_hex)
        elif kind == "mermaid" and data.get("mermaid_code"):
            png_bytes = _render_mermaid_300dpi_png(data["mermaid_code"], title=data.get("title", "Conceptual Framework"), diagram_style=data.get("diagram_style", "flowchart"), primary_hex=p_hex, accent_hex=a_hex)
            new_p_el = OxmlElement("w:p")
            anchor_para._p.addnext(new_p_el)
            new_p = docx.text.paragraph.Paragraph(new_p_el, anchor_para._parent)
            new_p.add_run().add_picture(io.BytesIO(png_bytes), width=Inches(5.6))
            cap = new_p.insert_paragraph_before(f"Figure: {data.get('title', 'Conceptual Framework')}")
            cap.runs[0].bold = True
        elif kind == "qr":
            qr_png = _generate_qr_png(data.get("prompt_payload", data.get("title", "Self-Check")), primary_hex=p_hex)
            new_p_el = OxmlElement("w:p")
            anchor_para._p.addnext(new_p_el)
            new_p = docx.text.paragraph.Paragraph(new_p_el, anchor_para._parent)
            new_p.add_run().add_picture(io.BytesIO(qr_png), width=Inches(1.25))
            r_lbl = new_p.add_run(f"  [{data.get('hook_type', 'QR Micro-Quiz')}] {data.get('title', '')}: {data.get('prompt_payload', '')}")
            r_lbl.italic = True
            r_lbl.font.size = Pt(9)
        elif kind == "illustration" and data.get("image_bytes"):
            new_p_el = OxmlElement("w:p")
            anchor_para._p.addnext(new_p_el)
            new_p = docx.text.paragraph.Paragraph(new_p_el, anchor_para._parent)
            new_p.add_run().add_picture(io.BytesIO(data["image_bytes"]), width=Inches(5.6))
            cap = new_p.insert_paragraph_before(f"Illustration: {data.get('title', 'Academic Figure')}")
            cap.runs[0].bold = True

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def enrich_pptx_with_native_animations(
    original_pptx_bytes: bytes,
    reviewer_email: str,
    transition_style: str = "fade",
) -> Tuple[bytes, Dict]:
    prs = Presentation(io.BytesIO(original_pptx_bytes))
    prs.core_properties.comments = f"Layer-5 Native XML Animated Clone approved by {reviewer_email}"

    dense_slides = []
    animated_shapes_total = 0

    for s_idx, slide in enumerate(prs.slides, start=1):
        sld_el = slide.element
        existing_tr = sld_el.find(pptx_qn("p:transition"))
        if existing_tr is not None:
            sld_el.remove(existing_tr)

        tr_tag = "p:push" if transition_style == "push" else "p:fade"
        transition_xml = parse_xml(f'<p:transition {nsdecls("p")} spd="med"><{tr_tag}/></p:transition>')

        clr_map = sld_el.find(pptx_qn("p:clrMapOvr"))
        c_sld = sld_el.find(pptx_qn("p:cSld"))
        if clr_map is not None:
            clr_map.addnext(transition_xml)
        elif c_sld is not None:
            c_sld.addnext(transition_xml)
        else:
            sld_el.insert(0, transition_xml)

        shape_ids = []
        slide_words = 0
        for shp in slide.shapes:
            if shp.has_text_frame:
                txt = shp.text_frame.text or ""
                slide_words += len(txt.split())
                if txt.strip():
                    shape_ids.append(shp.shape_id)

        if slide_words > 400:
            dense_slides.append({"slide_no": s_idx, "word_count": slide_words})

        if len(shape_ids) > 1:
            existing_timing = sld_el.find(pptx_qn("p:timing"))
            if existing_timing is not None:
                sld_el.remove(existing_timing)

            anim_nodes = []
            node_id = 3
            for spid in shape_ids[1:]:
                anim_nodes.append(f"""
                <p:par>
                  <p:cTn id="{node_id}" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" nodeType="clickEffect">
                    <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                    <p:childTnLst>
                      <p:set>
                        <p:cBhvr>
                          <p:cTn id="{node_id + 1}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>
                          <p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>
                          <p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>
                        </p:cBhvr>
                        <p:to><p:strVal val="visible"/></p:to>
                      </p:set>
                      <p:animEffect transition="in" filter="fade">
                        <p:cBhvr>
                          <p:cTn id="{node_id + 2}" dur="500"/>
                          <p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>
                        </p:cBhvr>
                      </p:animEffect>
                    </p:childTnLst>
                  </p:cTn>
                </p:par>""")
                node_id += 3
                animated_shapes_total += 1

            timing_xml = parse_xml(f"""
            <p:timing {nsdecls("p")}>
              <p:tnLst>
                <p:par>
                  <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">
                    <p:childTnLst>
                      <p:seq concurrent="1" nextAc="seek">
                        <p:cTn id="2" dur="indefinite" nodeType="mainSeq">
                          <p:childTnLst>{''.join(anim_nodes)}</p:childTnLst>
                        </p:cTn>
                        <p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>
                        <p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>
                      </p:seq>
                    </p:childTnLst>
                  </p:cTn>
                </p:par>
              </p:tnLst>
            </p:timing>""")
            sld_el.append(timing_xml)

    out_buf = io.BytesIO()
    prs.save(out_buf)
    summary = {"slides_processed": len(prs.slides), "animated_shapes": animated_shapes_total, "dense_slides_over_400w": dense_slides}
    return out_buf.getvalue(), summary


# ==============================================================================
# 9. ASPECT 2: 1-TO-N COMPLETE BOOK COMPILER (MIRROR DUPLEX & DYNAMIC BRANDING)
# ==============================================================================

def _append_word_field(paragraph, instr_str: str, placeholder_text: str = "1"):
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), instr_str)
    run = OxmlElement("w:r")
    txt = OxmlElement("w:t")
    txt.text = placeholder_text
    run.append(txt)
    fld.append(run)
    paragraph._p.append(fld)


def _add_bookmark_to_paragraph(paragraph, bookmark_name: str, bm_id: int):
    bm_start = OxmlElement("w:bookmarkStart")
    bm_start.set(qn("w:id"), str(bm_id))
    bm_start.set(qn("w:name"), bookmark_name)
    bm_end = OxmlElement("w:bookmarkEnd")
    bm_end.set(qn("w:id"), str(bm_id))
    paragraph._p.insert(0, bm_start)
    paragraph._p.append(bm_end)


def _remove_table_borders(table):
    tbl_pr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for b_name in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{b_name}")
        el.set(qn("w:val"), "none")
        borders.append(el)
    tbl_pr.append(borders)


def _apply_mirror_duplex_headers_footers(
    section,
    logo_bytes: bytes,
    course_name: str,
    primary_hex: str,
    accent_hex: str,
    enable_notes_margin: bool = True,
):
    section.different_first_page_header_footer = False
    p_rgb = RGBColor(*_hex_to_rgb_tuple(primary_hex))
    a_rgb = RGBColor(*_hex_to_rgb_tuple(accent_hex))

    odd_hdr = section.header
    odd_hdr.is_linked_to_previous = False
    odd_hdr.paragraphs[0].paragraph_format.space_after = Pt(0)
    odd_tbl = odd_hdr.add_table(rows=1, cols=3, width=Inches(6.5))
    odd_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    _remove_table_borders(odd_tbl)
    odd_tbl.columns[0].width, odd_tbl.columns[1].width, odd_tbl.columns[2].width = Inches(1.1), Inches(4.4), Inches(1.0)

    odd_tbl.cell(0, 0).paragraphs[0].add_run().add_picture(io.BytesIO(logo_bytes), width=Inches(0.65))
    c1 = odd_tbl.cell(0, 1).paragraphs[0]
    c1.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r_cn = c1.add_run(course_name)
    r_cn.bold = True
    r_cn.font.size = Pt(9.5)
    r_cn.font.color.rgb = p_rgb

    c2 = odd_tbl.cell(0, 2).paragraphs[0]
    c2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    if enable_notes_margin:
        r_nt = c2.add_run("Notes")
        r_nt.bold = True
        r_nt.font.size = Pt(9)
        r_nt.font.color.rgb = a_rgb

    even_hdr = section.even_page_header
    even_hdr.is_linked_to_previous = False
    even_hdr.paragraphs[0].paragraph_format.space_after = Pt(0)
    even_tbl = even_hdr.add_table(rows=1, cols=3, width=Inches(6.5))
    even_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    _remove_table_borders(even_tbl)
    even_tbl.columns[0].width, even_tbl.columns[1].width, even_tbl.columns[2].width = Inches(1.0), Inches(4.4), Inches(1.1)

    ec0 = even_tbl.cell(0, 0).paragraphs[0]
    ec0.alignment = WD_ALIGN_PARAGRAPH.LEFT
    if enable_notes_margin:
        er_nt = ec0.add_run("Notes")
        er_nt.bold = True
        er_nt.font.size = Pt(9)
        er_nt.font.color.rgb = a_rgb

    ec1 = even_tbl.cell(0, 1).paragraphs[0]
    ec1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    er_cn = ec1.add_run(course_name)
    er_cn.bold = True
    er_cn.font.size = Pt(9.5)
    er_cn.font.color.rgb = p_rgb

    ec2 = even_tbl.cell(0, 2).paragraphs[0]
    ec2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    ec2.add_run().add_picture(io.BytesIO(logo_bytes), width=Inches(0.65))

    odd_ftr = section.footer
    odd_ftr.is_linked_to_previous = False
    of_p = odd_ftr.paragraphs[0]
    of_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    _append_word_field(of_p, "PAGE", "1")

    even_ftr = section.even_page_footer
    even_ftr.is_linked_to_previous = False
    ef_p = even_ftr.paragraphs[0]
    ef_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    _append_word_field(ef_p, "PAGE", "2")


def _clone_docx_body_into_book(src_doc: docx.Document, target_doc: docx.Document, unit_num: int):
    src_part = src_doc.part
    for child in src_doc.element.body:
        tag = child.tag.split("}")[-1]
        if tag == "p":
            src_p = docx.text.paragraph.Paragraph(child, src_doc)
            txt = src_p.text or ""
            blips = child.xpath(".//*[local-name()='blip']")
            if not txt.strip() and not blips:
                continue

            style_name = src_p.style.name if src_p.style else "Normal"
            if style_name.startswith("Heading 1"):
                dst_p = target_doc.add_heading("", level=1)
            elif style_name.startswith("Heading 2"):
                dst_p = target_doc.add_heading("", level=2)
            elif style_name.startswith("Heading 3"):
                dst_p = target_doc.add_heading("", level=3)
            elif re.match(r"^\d+\.\d+\s+[A-Za-z]", txt.strip()) and len(txt.split()) <= 14:
                dst_p = target_doc.add_heading("", level=2)
            elif re.match(r"^\d+\.\d+\.\d+\s+[A-Za-z]", txt.strip()) and len(txt.split()) <= 16:
                dst_p = target_doc.add_heading("", level=3)
            else:
                dst_p = target_doc.add_paragraph()

            dst_p.alignment = src_p.alignment
            if blips:
                for bl in blips:
                    r_id = bl.get(qn("r:embed"))
                    if r_id and r_id in src_part.related_parts:
                        img_blob = src_part.related_parts[r_id].blob
                        w_inches = 1.25 if len(img_blob) < 8000 else 5.4
                        dst_p.add_run().add_picture(io.BytesIO(img_blob), width=Inches(w_inches))

            for r in src_p.runs:
                if r.text:
                    nr = dst_p.add_run(r.text)
                    nr.bold = r.bold
                    nr.italic = r.italic
                    nr.underline = r.underline
                    if r.font.size:
                        nr.font.size = r.font.size
                    if r.font.color and r.font.color.rgb:
                        nr.font.color.rgb = r.font.color.rgb

        elif tag == "tbl":
            src_tbl = docx.table.Table(child, src_doc)
            rows_cnt, cols_cnt = len(src_tbl.rows), len(src_tbl.columns)
            if rows_cnt == 0 or cols_cnt == 0:
                continue
            dst_tbl = target_doc.add_table(rows=rows_cnt, cols=cols_cnt)
            dst_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
            if rows_cnt > 1 or cols_cnt > 1:
                dst_tbl.style = "Table Grid"

            for r_i, s_row in enumerate(src_tbl.rows):
                for c_i, s_cell in enumerate(s_row.cells):
                    d_cell = dst_tbl.cell(r_i, c_i)
                    shd_els = s_cell._tc.xpath(".//*[local-name()='shd']")
                    left_els = s_cell._tc.xpath(".//*[local-name()='tcBorders']/*[local-name()='left']")
                    fill_hex = shd_els[0].get(qn("w:fill")) if shd_els else None
                    l_hex = left_els[0].get(qn("w:color")) if left_els else None
                    if fill_hex and fill_hex != "auto":
                        _set_cell_shading_and_border(d_cell, fill_hex, left_border_hex=l_hex)

                    d_cell.text = ""
                    for p_i, sp in enumerate(s_cell.paragraphs):
                        dp = d_cell.paragraphs[0] if p_i == 0 else d_cell.add_paragraph()
                        for sr in sp.runs:
                            if sr.text:
                                dr = dp.add_run(sr.text)
                                dr.bold = sr.bold
                                dr.italic = sr.italic
                                if sr.font.size:
                                    dr.font.size = sr.font.size
                                if sr.font.color and sr.font.color.rgb:
                                    dr.font.color.rgb = sr.font.color.rgb


def compile_complete_slm_book_docx(
    unit_docs: List[Dict],
    brand_profile: Dict,
    book_meta: Dict,
    reviewer_email: str,
) -> bytes:
    p_hex = brand_profile.get("primary_hex", "1E3A8A")
    a_hex = brand_profile.get("accent_hex", "D97706")
    p_rgb = RGBColor(*_hex_to_rgb_tuple(p_hex))
    a_rgb = RGBColor(*_hex_to_rgb_tuple(a_hex))

    uni_name = brand_profile.get("university_name", "UNIVERSITY").upper()
    tagline = brand_profile.get("tagline", "")
    school_name = brand_profile.get("school_name", "Centre for Distance and Online Education")
    logo_bytes = brand_profile.get("logo_bytes") or _generate_brand_crest_png(uni_name, p_hex, a_hex)
    enable_notes = bool(brand_profile.get("enable_mirror_notes_margin", True))

    program_name = book_meta.get("program_name", "MASTERS OF BUSINESS ADMINISTRATION (MBA)")
    course_name = book_meta.get("course_name", "Course Textbook")
    course_code = book_meta.get("course_code", "COURSE-101")
    credits = book_meta.get("credits", "04")
    doc_type = book_meta.get("document_type", "Self-Learning Material (SLM) — Complete Course Book")
    edition = book_meta.get("edition_year", f"{datetime.now().year} Edition")
    committee_rows = book_meta.get("committee", brand_profile.get("default_committee", []))
    co_rows = book_meta.get("course_outcomes", [])

    book = docx.Document()
    book.settings.odd_and_even_pages_header_footer = True
    book.core_properties.title = f"{course_code} - {course_name}"
    book.core_properties.author = reviewer_email
    book.core_properties.subject = f"{uni_name} | {doc_type}"

    sec0 = book.sections[0]
    sec0.top_margin, sec0.bottom_margin = Inches(0.85), Inches(0.85)
    sec0.left_margin, sec0.right_margin, sec0.gutter = Inches(0.95), Inches(0.85), Inches(0.20)

    # Page 1: Cover
    p_logo = book.add_paragraph()
    p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_logo.paragraph_format.space_before = Pt(18)
    p_logo.add_run().add_picture(io.BytesIO(logo_bytes), width=Inches(1.55))

    p_uname = book.add_paragraph()
    p_uname.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_uname.paragraph_format.space_before = Pt(10)
    r_un = p_uname.add_run(uni_name)
    r_un.bold = True
    r_un.font.size = Pt(20)
    r_un.font.color.rgb = p_rgb

    if tagline:
        p_tag = book.add_paragraph()
        p_tag.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_tg = p_tag.add_run(tagline.upper())
        r_tg.bold = True
        r_tg.font.size = Pt(10.5)
        r_tg.font.color.rgb = a_rgb

    p_sch = book.add_paragraph()
    p_sch.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sch.paragraph_format.space_before = Pt(6)
    r_sc = p_sch.add_run(school_name)
    r_sc.bold = True
    r_sc.font.size = Pt(13)

    p_dtype = book.add_paragraph()
    p_dtype.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_dtype.paragraph_format.space_before = Pt(24)
    r_dt = p_dtype.add_run(doc_type.upper())
    r_dt.bold = True
    r_dt.font.size = Pt(10)
    r_dt.font.color.rgb = a_rgb

    p_prog = book.add_paragraph()
    p_prog.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_prog.paragraph_format.space_before = Pt(8)
    r_pr = p_prog.add_run(program_name.upper())
    r_pr.bold = True
    r_pr.font.size = Pt(13.5)
    r_pr.font.color.rgb = p_rgb

    cov_tbl = book.add_table(rows=1, cols=1)
    cov_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    ccell = cov_tbl.cell(0, 0)
    _set_cell_shading_and_border(ccell, _tint_hex(p_hex, 0.90), left_border_hex=p_hex)
    cp1 = ccell.paragraphs[0]
    cp1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cp1.paragraph_format.space_before = Pt(12)
    cp1.paragraph_format.space_after = Pt(6)
    rc_title = cp1.add_run(course_name)
    rc_title.bold = True
    rc_title.font.size = Pt(22)
    rc_title.font.color.rgb = p_rgb

    cp2 = ccell.add_paragraph()
    cp2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cp2.paragraph_format.space_after = Pt(12)
    rc_meta = cp2.add_run(f"Course Code: {course_code}   |   Credits: {credits}   |   {edition}")
    rc_meta.bold = True
    rc_meta.font.size = Pt(11.5)

    # Page 2: Governance & Statutory Page
    book.add_page_break()
    h_gov = book.add_heading("Program Design & Editorial Governance Board", level=1)
    h_gov.runs[0].font.color.rgb = p_rgb

    if committee_rows:
        gtbl = book.add_table(rows=1, cols=2)
        gtbl.style = "Table Grid"
        gtbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        _set_cell_shading_and_border(gtbl.cell(0, 0), _tint_hex(p_hex, 0.85))
        _set_cell_shading_and_border(gtbl.cell(0, 1), _tint_hex(p_hex, 0.85))
        gtbl.cell(0, 0).paragraphs[0].add_run("Role / Editorial Designation").bold = True
        gtbl.cell(0, 1).paragraphs[0].add_run("Faculty / Academic Member").bold = True
        for row_item in committee_rows:
            r_cells = gtbl.add_row().cells
            r_cells[0].paragraphs[0].add_run(str(row_item.get("role", ""))).bold = True
            r_cells[1].text = str(row_item.get("name", ""))

    p_qa = book.add_paragraph()
    p_qa.paragraph_format.space_before = Pt(14)
    p_qa.add_run(
        f"Academic QA Sign-Off (Layer-5): APPROVED  |  Verified By: {reviewer_email}  |  "
        f"{datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M UTC')}"
    ).bold = True

    stat_tbl = book.add_table(rows=1, cols=1)
    stat_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    scell = stat_tbl.cell(0, 0)
    _set_cell_shading_and_border(scell, _tint_hex(a_hex, 0.92), left_border_hex=a_hex)
    sp1 = scell.paragraphs[0]
    sp1.add_run(f"Published: {edition}  |  {brand_profile.get('copyright_line', '')}").bold = True
    sp2 = scell.add_paragraph(brand_profile.get("statutory_footer", ""))
    sp2.runs[0].font.size = Pt(9)

    # Page 3: Master Syllabus Index & CO1-CO5
    book.add_page_break()
    h_idx = book.add_heading(f"{course_name} — Master Book Index", level=1)
    h_idx.runs[0].font.color.rgb = p_rgb

    idx_tbl = book.add_table(rows=1, cols=4)
    idx_tbl.style = "Table Grid"
    idx_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    headers_idx = ["Block / Module", "Unit", "Unit Title & Key Topics", "Start Page"]
    for ci, h_txt in enumerate(headers_idx):
        _set_cell_shading_and_border(idx_tbl.cell(0, ci), _tint_hex(p_hex, 0.85))
        idx_tbl.cell(0, ci).paragraphs[0].add_run(h_txt).bold = True

    for u_i, u_item in enumerate(unit_docs, start=1):
        rc = idx_tbl.add_row().cells
        rc[0].text = str(u_item.get("block_label", f"Block {(u_i - 1) // 3 + 1}"))
        rc[1].paragraphs[0].add_run(f"Unit {u_i}").bold = True
        rc[2].text = str(u_item.get("unit_title", u_item.get("filename", f"Unit {u_i}")))
        p_pg = rc[3].paragraphs[0]
        p_pg.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _append_word_field(p_pg, f"PAGEREF _L5_Unit_{u_i} \\h", str(u_i * 12 - 11))

    if co_rows:
        h_co = book.add_heading("Course Objectives (CO Mapping)", level=1)
        h_co.paragraph_format.space_before = Pt(16)
        h_co.runs[0].font.color.rgb = p_rgb
        cotbl = book.add_table(rows=1, cols=2)
        cotbl.style = "Table Grid"
        cotbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        _set_cell_shading_and_border(cotbl.cell(0, 0), _tint_hex(p_hex, 0.85))
        _set_cell_shading_and_border(cotbl.cell(0, 1), _tint_hex(p_hex, 0.85))
        cotbl.cell(0, 0).paragraphs[0].add_run("CO Code").bold = True
        cotbl.cell(0, 1).paragraphs[0].add_run("Course Objective Description").bold = True
        for co in co_rows:
            r_c = cotbl.add_row().cells
            r_c[0].paragraphs[0].add_run(str(co.get("co_code", ""))).bold = True
            r_c[1].text = str(co.get("description", ""))

    # Chapters 1..N: Odd-Page Starts + Mirror Duplex Headers
    for u_i, u_item in enumerate(unit_docs, start=1):
        u_sec = book.add_section(WD_SECTION.ODD_PAGE)
        u_sec.top_margin, u_sec.bottom_margin = Inches(0.90), Inches(0.80)
        u_sec.left_margin, u_sec.right_margin, u_sec.gutter = Inches(0.95), Inches(0.85), Inches(0.20)

        _apply_mirror_duplex_headers_footers(
            section=u_sec,
            logo_bytes=logo_bytes,
            course_name=f"{course_code}: {course_name}",
            primary_hex=p_hex,
            accent_hex=a_hex,
            enable_notes_margin=enable_notes,
        )

        block_lbl = str(u_item.get("block_label", f"Block {(u_i - 1) // 3 + 1}"))
        unit_title = str(u_item.get("unit_title", f"Unit {u_i}"))

        p_blk = book.add_paragraph()
        r_bl = p_blk.add_run(f"{block_lbl.upper()}  |  {school_name}")
        r_bl.bold = True
        r_bl.font.size = Pt(9.5)
        r_bl.font.color.rgb = a_rgb

        h_unit = book.add_heading(f"Unit {u_i}: {unit_title}", level=1)
        h_unit.runs[0].font.color.rgb = p_rgb
        _add_bookmark_to_paragraph(h_unit, f"_L5_Unit_{u_i}", bm_id=100 + u_i)

        src_doc = docx.Document(io.BytesIO(u_item["bytes"]))
        _clone_docx_body_into_book(src_doc, book, unit_num=u_i)

    out_buf = io.BytesIO()
    book.save(out_buf)
    return out_buf.getvalue()


# ==============================================================================
# 10. DRIVE ARCHIVAL & MASTER TRACKER WRITE-BACK (NON-DESTRUCTIVE SYNC)
# ==============================================================================

def _col_idx_to_a1(idx_zero_based: int) -> str:
    n = idx_zero_based + 1
    chars = []
    while n > 0:
        n, rem = divmod(n - 1, 26)
        chars.append(chr(65 + rem))
    return "".join(reversed(chars))


def _get_or_create_subfolder(drive_svc, parent_id: str, folder_name: str) -> str:
    safe_name = re.sub(r'[\\/*?:"<>|]+', "_", folder_name).strip() or "Layer5_Approved"
    escaped = safe_name.replace("'", "\\'")
    q = f"'{parent_id}' in parents and name = '{escaped}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
    res = drive_svc.files().list(q=q, fields="files(id)", supportsAllDrives=True, includeItemsFromAllDrives=True).execute()
    if res.get("files"):
        return res["files"][0]["id"]
    meta = {"name": safe_name, "mimeType": "application/vnd.google-apps.folder", "parents": [parent_id]}
    created = drive_svc.files().create(body=meta, fields="id", supportsAllDrives=True).execute()
    return created["id"]


def upload_layer5_artifact_to_drive(
    drive_svc,
    root_folder_id: str,
    hierarchy_parts: List[str],
    filename: str,
    file_bytes: bytes,
    mime_type: str,
) -> str:
    curr_parent = root_folder_id
    for part in hierarchy_parts:
        if part and str(part).strip():
            curr_parent = _get_or_create_subfolder(drive_svc, curr_parent, str(part).strip())

    safe_fn = re.sub(r'[\\/*?:"<>|]+', "_", filename).strip()
    meta = {"name": safe_fn, "parents": [curr_parent]}
    media = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype=mime_type, resumable=False)
    uploaded = drive_svc.files().create(body=meta, media_body=media, fields="id, webViewLink", supportsAllDrives=True).execute()
    return uploaded.get("webViewLink") or f"https://drive.google.com/file/d/{uploaded['id']}/view"


def sync_layer5_master_tracker_row(
    sheets_svc,
    spreadsheet_id: str,
    tab_name: str,
    header_row_1idx: int,
    data_row_1idx: int,
    raw_headers: List[str],
    enriched_doc_link: str,
    l5_approval_status: str,
    l5_approved_by: str,
):
    safe_tab = tab_name.replace("'", "''")
    lower_hdrs = [str(h).strip().lower() for h in raw_headers]
    next_col_idx = len(raw_headers)
    updates = []

    def _resolve_col(col_label: str) -> int:
        nonlocal next_col_idx
        key = col_label.lower()
        if key in lower_hdrs:
            return lower_hdrs.index(key)
        assigned = next_col_idx
        next_col_idx += 1
        lower_hdrs.append(key)
        updates.append({"range": f"'{safe_tab}'!{_col_idx_to_a1(assigned)}{header_row_1idx}", "values": [[col_label]]})
        return assigned

    c_link = _resolve_col("Enriched_Doc_Link")
    c_stat = _resolve_col("L5_Approval_Status")
    c_by = _resolve_col("L5_Approved_By")

    updates.extend([
        {"range": f"'{safe_tab}'!{_col_idx_to_a1(c_link)}{data_row_1idx}", "values": [[enriched_doc_link]]},
        {"range": f"'{safe_tab}'!{_col_idx_to_a1(c_stat)}{data_row_1idx}", "values": [[l5_approval_status]]},
        {"range": f"'{safe_tab}'!{_col_idx_to_a1(c_by)}{data_row_1idx}", "values": [[l5_approved_by]]},
    ])

    sheets_svc.spreadsheets().values().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"valueInputOption": "USER_ENTERED", "data": updates},
    ).execute()
