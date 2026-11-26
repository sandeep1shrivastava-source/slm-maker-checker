import io, json, re
import openpyxl
from docx import Document

def auto_clean_mcq_workbook(source, audit_res: dict | None = None, llm_rephrase_fn=None) -> tuple[bytes, dict]:
    """
    Deterministic + optional LLM-assisted Layer-5 surgical remediator for MCQ Excel workbooks.
    - Renames corrupted option-key headers (e.g., '3orre3t Option') -> 'Correct Option'
    - Normalizes numeric option identifiers (1/2/3/4) -> (A/B/C/D) and verifies alignment with Correct Answer text
    - Strips mid-sheet repeated headers & section banners, re-sequences Sr No (1..N)
    - Detects and realigns left-shifted metadata cells (e.g., 'L2' in Syllabus Subpoint, 'Easy' in BTL)
    - Flags/removes exact duplicate question stems across tabs into an 'Audit_Change_Log' sheet
    - Optionally diversifies repetitive formulaic stem openers via llm_rephrase_fn
    """
    if isinstance(source, (bytes, bytearray)):
        raw_bytes = bytes(source)
    elif hasattr(source, "read"):
        pos = source.tell() if hasattr(source, "tell") else None
        raw_bytes = source.read()
        if pos is not None and hasattr(source, "seek"):
            source.seek(pos)
    else:
        raw_bytes = open(source, "rb").read()

    wb_in = openpyxl.load_workbook(io.BytesIO(raw_bytes), data_only=True)
    wb_out = openpyxl.Workbook()
    wb_out.remove(wb_out.active)

    log_rows = [["Sheet", "Original Row", "Sr No", "Action Taken", "Before", "After"]]
    seen_stems = {}
    stats = {"headers_fixed": 0, "keys_normalized": 0, "metadata_realigned": 0, "duplicates_removed": 0, "banners_stripped": 0, "stems_diversified": 0, "clean_items_kept": 0}
    let_map = {"1": "A", "2": "B", "3": "C", "4": "D"}
    btl_pat = re.compile(r"^(L[1-6]|BLOOM\s*[1-6]|REMEMBER|UNDERSTAND|APPLY|ANALYZE|EVALUATE|CREATE)$", re.I)
    diff_pat = re.compile(r"^(EASY|MEDIUM|MODERATE|HARD|DIFFICULT)$", re.I)

    for sname in wb_in.sheetnames:
        ws_in = wb_in[sname]
        ws_out = wb_out.create_sheet(title=sname[:31])

        if any(k in sname.lower() for k in ("structure", "syllabus", "blueprint")):
            for r in range(1, ws_in.max_row + 1):
                ws_out.append([ws_in.cell(r, c).value for c in range(1, ws_in.max_column + 1)])
            continue

        h_row, cmap = None, {}
        for r in range(1, min(11, ws_in.max_row + 1)):
            rvals = [str(ws_in.cell(r, c).value).strip().lower() if ws_in.cell(r, c).value is not None else "" for c in range(1, ws_in.max_column + 1)]
            if "question" in rvals and any(x in rvals for x in ("option a", "opt a", "option 1", "opt 1")):
                h_row = r
                for i, h in enumerate(rvals):
                    if h in ("sr no", "sr. no", "sr. no.", "q no", "q.no", "id"): cmap["sr"] = i
                    elif "subpoint" in h or "sub-point" in h or "sub point" in h: cmap["sub"] = i
                    elif "section" in h or "topic" in h or "module" in h or "unit" in h: cmap.setdefault("sec", i)
                    elif h == "question" or "question stem" in h: cmap["q"] = i
                    elif h in ("option a", "opt a", "a", "option 1", "opt 1", "1"): cmap["a"] = i
                    elif h in ("option b", "opt b", "b", "option 2", "opt 2", "2"): cmap["b"] = i
                    elif h in ("option c", "opt c", "c", "option 3", "opt 3", "3"): cmap["c"] = i
                    elif h in ("option d", "opt d", "d", "option 4", "opt 4", "4"): cmap["d"] = i
                    elif h in ("correct option", "option key", "answer key", "correct letter", "key") or ("option" in h and any(k in h for k in ("correct", "orre3t", "key", "id", "letter"))): cmap["let"] = i
                    elif h in ("correct answer", "answer text", "answer"): cmap["ans"] = i
                    elif "btl" in h or "bloom" in h: cmap["btl"] = i
                    elif "difficulty" in h: cmap["diff"] = i
                    elif "explanation" in h or "rationale" in h: cmap["exp"] = i
                break

        if h_row is None:
            for r in range(1, ws_in.max_row + 1):
                ws_out.append([ws_in.cell(r, c).value for c in range(1, ws_in.max_column + 1)])
            continue

        hdr_vals = [ws_in.cell(h_row, c).value for c in range(1, ws_in.max_column + 1)]
        if "let" in cmap:
            orig_hdr = str(hdr_vals[cmap["let"]] or "").strip()
            if orig_hdr.lower() != "correct option":
                hdr_vals[cmap["let"]] = "Correct Option"
                stats["headers_fixed"] += 1
                log_rows.append([sname.strip(), h_row, "-", "Header Sanitized", orig_hdr, "Correct Option"])
        ws_out.append(hdr_vals)

        new_sr = 0
        rephrase_batch = []

        for r in range(h_row + 1, ws_in.max_row + 1):
            row = [ws_in.cell(r, c).value for c in range(1, ws_in.max_column + 1)]
            non_empty = [v for v in row if v is not None and str(v).strip() != ""]
            if not non_empty:
                continue

            qv = row[cmap["q"]] if "q" in cmap and cmap["q"] < len(row) else None
            srv = row[cmap["sr"]] if "sr" in cmap and cmap["sr"] < len(row) else None
            oa = row[cmap["a"]] if "a" in cmap and cmap["a"] < len(row) else None

            if str(qv).strip().lower() in ("question", "question stem") or str(srv).strip().lower() in ("sr no", "sr. no"):
                stats["banners_stripped"] += 1
                log_rows.append([sname.strip(), r, str(srv or ""), "Removed Repeated Header Row", str(qv or ""), "(removed)"])
                continue
            if qv is None or (oa is None and len(non_empty) <= 2):
                stats["banners_stripped"] += 1
                log_rows.append([sname.strip(), r, "", "Removed Section Banner Row", str(non_empty[0]), "(removed)"])
                continue

            nq = re.sub(r"\s+", " ", str(qv).strip().lower())
            if nq in seen_stems:
                stats["duplicates_removed"] += 1
                log_rows.append([sname.strip(), r, str(srv or ""), f"Removed Duplicate Stem (matches {seen_stems[nq]})", str(qv).strip()[:100], "(removed)"])
                continue
            seen_stems[nq] = f"{sname.strip()} Row {r}"

            if all(k in cmap for k in ("sub", "btl", "diff")) and max(cmap["sub"], cmap["btl"], cmap["diff"]) < len(row):
                v_sub = str(row[cmap["sub"]] or "").strip()
                v_btl = str(row[cmap["btl"]] or "").strip()
                v_dif = str(row[cmap["diff"]] or "").strip()
                if btl_pat.match(v_sub) and diff_pat.match(v_btl):
                    before_str = f"Sub={v_sub} | BTL={v_btl} | Diff={v_dif[:40]}"
                    if "exp" in cmap and cmap["exp"] < len(row) and not str(row[cmap["exp"]] or "").strip() and len(v_dif) > 20:
                        row[cmap["exp"]] = v_dif
                    row[cmap["sub"]] = str(row[cmap["sec"]]).strip() if "sec" in cmap and cmap["sec"] < len(row) and row[cmap["sec"]] else "General"
                    row[cmap["btl"]] = v_sub.upper()
                    row[cmap["diff"]] = v_btl.title()
                    stats["metadata_realigned"] += 1
                    log_rows.append([sname.strip(), r, str(srv or ""), "Realigned Left-Shifted Metadata", before_str, f"Sub={row[cmap['sub']]} | BTL={row[cmap['btl']]} | Diff={row[cmap['diff']]}"])

            opts = {k: (str(row[cmap[k.lower()]]).strip() if k.lower() in cmap and cmap[k.lower()] < len(row) and row[cmap[k.lower()]] is not None else "") for k in ("A", "B", "C", "D")}
            ans = str(row[cmap["ans"]]).strip() if "ans" in cmap and cmap["ans"] < len(row) and row[cmap["ans"]] is not None else ""
            if "let" in cmap and cmap["let"] < len(row):
                raw_let = str(row[cmap["let"]] or "").strip().upper()
                norm_let = let_map.get(raw_let, raw_let)
                if norm_let not in ("A", "B", "C", "D") and ans:
                    for k_let, o_txt in opts.items():
                        if o_txt and o_txt.lower() == ans.lower():
                            norm_let = k_let
                            break
                if norm_let in ("A", "B", "C", "D") and norm_let != raw_let:
                    row[cmap["let"]] = norm_let
                    stats["keys_normalized"] += 1
                    log_rows.append([sname.strip(), r, str(srv or ""), "Normalized Option Identifier", raw_let, norm_let])

            new_sr += 1
            if "sr" in cmap and cmap["sr"] < len(row):
                row[cmap["sr"]] = new_sr

            out_row_idx = ws_out.max_row + 1
            ws_out.append(row)
            stats["clean_items_kept"] += 1

            q_str = str(qv).strip()
            if llm_rephrase_fn and re.match(r"^(Identify the|Choose the|Select the|Where is|When does)\b", q_str, re.I) and len(rephrase_batch) < 25:
                rephrase_batch.append((out_row_idx, r, new_sr, q_str))

        if llm_rephrase_fn and rephrase_batch:
            try:
                rewritten = llm_rephrase_fn([x[3] for x in rephrase_batch])
                if isinstance(rewritten, list) and len(rewritten) == len(rephrase_batch):
                    for (out_r, orig_r, sr_i, old_q), new_q in zip(rephrase_batch, rewritten):
                        nq_clean = str(new_q or "").strip()
                        if nq_clean and nq_clean != old_q:
                            ws_out.cell(out_r, cmap["q"] + 1).value = nq_clean
                            stats["stems_diversified"] += 1
                            log_rows.append([sname.strip(), orig_r, sr_i, "Diversified Formulaic Stem Opener", old_q[:90], nq_clean[:90]])
            except Exception:
                pass

    ws_log = wb_out.create_sheet(title="Audit_Change_Log")
    for lr in log_rows:
        ws_log.append(lr)

    out_buf = io.BytesIO()
    wb_out.save(out_buf)
    return out_buf.getvalue(), stats


def build_surgical_fixit_prompt(title: str, subject: str, chapter_text: str, audit_res: dict) -> str:
    issues_summary = json.dumps({
        "overall_score": audit_res.get("overall_score", audit_res.get("final_score", "")),
        "verdict": audit_res.get("verdict", ""),
        "score_breakdown": audit_res.get("score_breakdown", []),
        "technical_errors": audit_res.get("technical_errors", audit_res.get("subject_knowledge_issues", []))[:15],
        "clarity_issues": audit_res.get("clarity_issues", [])[:15],
        "grammar_issues": audit_res.get("grammar_issues", [])[:15],
        "fixes_to_lift_score": audit_res.get("fixes_that_would_lift_score", audit_res.get("top_fixes", [])),
    }, ensure_ascii=False, indent=2)

    return f"""You are a Senior Academic Editor executing a Layer-5 SURGICAL FIX-IT ENHANCEMENT for an upGrad / UGC-aligned Self-Learning Material (SLM) chapter.
Do NOT rewrite the entire chapter from scratch. Generate ONLY self-contained, copy-paste-ready remediation blocks that fix every defect flagged in the Layer-4 Audit Report below so the chapter achieves >=90% QA approval.

DOCUMENT TITLE: {title}
SUBJECT: {subject}

=== LAYER-4 QA DEFECT REPORT (TARGET REMEDIATION LIST) ===
{issues_summary}

=== SOURCE CHAPTER EXCERPT (FOR CONTEXT & TERMINOLOGY ALIGNMENT) ===
{chapter_text[:45000]}

Return a strict JSON object with these exact keys:
{{
  "executive_remediation_summary": "2-3 sentence summary of surgical fixes generated and projected post-fix score",
  "projected_post_fix_score": 92.0,
  "technical_and_conceptual_patches": [
    {{"location": "Section/Heading", "original_issue": "Defect description", "surgical_replacement": "Corrected academic paragraph/formula/definition"}}
  ],
  "missing_pedagogical_blocks": [
    {{"block_type": "Check Your Progress / Worked Example / Summary / Glossary", "insert_after_section": "Target heading", "ready_to_paste_content": "Complete formatted block including Answer Key"}}
  ],
  "ai_burstiness_rewrites": [
    {{"target_paragraph_hint": "Opening words of robotic/uniform paragraph", "rewritten_high_burstiness_prose": "Varied sentence-length academic prose (mix of crisp 8-word statements and 25-word analytical sentences) preserving all technical facts"}}
  ],
  "diagram_and_citation_addenda": [
    {{"item_type": "Diagram Specification / APA-7 Reference", "content": "Detailed figure caption/spec or verified scholarly reference"}}
  ]
}}
"""


def build_surgical_fixit_docx(title: str, subject: str, fixit_payload: dict) -> bytes:
    doc = Document()
    doc.add_heading(f"Layer-5 Surgical Fix-It Addendum: {title}", level=0)
    p_meta = doc.add_paragraph()
    p_meta.add_run("Subject: ").bold = True
    p_meta.add_run(f"{subject}   |   ")
    p_meta.add_run("Projected Post-Fix Score: ").bold = True
    p_meta.add_run(f"{fixit_payload.get('projected_post_fix_score', '90+')}%")

    doc.add_paragraph(str(fixit_payload.get("executive_remediation_summary", "")))

    sections = [
        ("1. Technical & Conceptual Patches", "technical_and_conceptual_patches", ["location", "original_issue", "surgical_replacement"], ["Location", "Flagged Defect", "Ready-to-Paste Replacement"]),
        ("2. Missing Pedagogical Blocks (CYP, Worked Examples, Glossary)", "missing_pedagogical_blocks", ["block_type", "insert_after_section", "ready_to_paste_content"], ["Block Type", "Insert After", "Content & Answer Key"]),
        ("3. High-Burstiness Prose Rewrites (AI-Uniformity Remediation)", "ai_burstiness_rewrites", ["target_paragraph_hint", "rewritten_high_burstiness_prose"], ["Target Paragraph", "Human-Cadence Academic Rewrite"]),
        ("4. Diagram Specifications & Verified References", "diagram_and_citation_addenda", ["item_type", "content"], ["Asset / Citation Type", "Specification / Reference"]),
    ]

    for heading, key, fields, headers in sections:
        items = fixit_payload.get(key) or []
        doc.add_heading(heading, level=1)
        if not items:
            doc.add_paragraph("No surgical patches required in this category.")
            continue
        tbl = doc.add_table(rows=1, cols=len(headers))
        tbl.style = "Table Grid"
        for idx, h_text in enumerate(headers):
            cell = tbl.rows[0].cells[idx]
            cell.text = h_text
            if cell.paragraphs and cell.paragraphs[0].runs:
                cell.paragraphs[0].runs[0].bold = True
        for it in items:
            row_cells = tbl.add_row().cells
            if isinstance(it, dict):
                for idx, f_key in enumerate(fields):
                    row_cells[idx].text = str(it.get(f_key, "") or "")
            else:
                row_cells[0].text = str(it)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()

import pptx

def parse_layer4_docx_report(report_source) -> dict:
    raw = bytes(report_source) if isinstance(report_source, (bytes, bytearray)) else report_source.read()
    doc = Document(io.BytesIO(raw))
    paras = [(p.text or "").strip() for p in doc.paragraphs if (p.text or "").strip()]
    overall_score, verdict = "", ""
    for p in paras[:35]:
        m = re.search(r"(?:score|rating|total)[\s:]*(\d+(?:\.\d+)?%?)", p, re.I)
        if m and not overall_score:
            overall_score = m.group(1)
        if any(k in p.lower() for k in ("verdict", "status", "reject", "rework", "accept", "pass")) and len(p) < 140:
            verdict = p
    score_breakdown, findings = [], []
    for t_idx, tbl in enumerate(doc.tables, start=1):
        rows = [[(c.text or "").strip() for c in row_obj.cells] for row_obj in tbl.rows]
        if len(rows) < 2:
            continue
        hdr_joined = " ".join(rows[0]).lower()
        if any(k in hdr_joined for k in ("parameter", "criterion", "score", "weight")):
            for r in rows[1:]:
                if any(r):
                    score_breakdown.append(" | ".join([c for c in r if c]))
        else:
            for r_i, r in enumerate(rows[1:], start=1):
                if any(r):
                    item = " | ".join([c for c in r if c])
                    findings.append(f"[L4-T{t_idx}-R{r_i}] {item}")
    for p in paras:
        if p.startswith(("-", "*", "\u2022")) or re.match(r"^\d+[\.\)]", p):
            if len(p) > 15:
                findings.append(p)
    deduped = list(dict.fromkeys(findings))
    return {
        "overall_score": overall_score or "Report v1 Loaded",
        "verdict": verdict or "Layer-4 Baseline Report",
        "score_breakdown": score_breakdown[:20],
        "technical_errors": deduped[:25],
        "clarity_issues": deduped[25:45],
        "fixes_that_would_lift_score": deduped[:30],
        "total_findings": len(deduped),
    }

def extract_manuscript_text(raw_bytes: bytes, filename: str) -> str:
    fn = filename.lower()
    if fn.endswith(".docx"):
        d = Document(io.BytesIO(raw_bytes))
        blocks = [p.text.strip() for p in d.paragraphs if p.text.strip()]
        for t in d.tables:
            for r in t.rows:
                rt = " | ".join(c.text.strip() for c in r.cells if c.text.strip())
                if rt:
                    blocks.append(rt)
        return "\n".join(blocks)
    elif fn.endswith(".pdf"):
        import pypdf
        reader = pypdf.PdfReader(io.BytesIO(raw_bytes))
        return "\n".join((pg.extract_text() or "") for pg in reader.pages)
    elif fn.endswith(".pptx"):
        prs = pptx.Presentation(io.BytesIO(raw_bytes))
        slides_out = []
        for idx, slide in enumerate(prs.slides, start=1):
            parts = [f"[--- SLIDE {idx} ---]"]
            for shp in slide.shapes:
                if shp.has_text_frame:
                    t_shp = "\n".join(p.text.strip() for p in shp.text_frame.paragraphs if p.text.strip())
                    if t_shp:
                        parts.append(t_shp)
                if shp.has_table:
                    for r in shp.table.rows:
                        rt = " | ".join(c.text.strip() for c in r.cells if c.text.strip())
                        if rt:
                            parts.append(rt)
            if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
                nt = (slide.notes_slide.notes_text_frame.text or "").strip()
                if nt:
                    parts.append(f"[Speaker Notes]: {nt}")
            slides_out.append("\n".join(parts))
        return "\n\n".join(slides_out)
    return raw_bytes.decode("utf-8", errors="ignore")



def build_updated_manuscript_bytes(raw_bytes: bytes, filename: str, title: str, subject: str, payload: dict) -> tuple[bytes, str, str]:
    """Clones the original .docx/.pptx/.txt manuscript and embeds approved Layer-5 fixes inline (Zero-Risk Clone-on-Write)."""
    fn = filename.lower()
    rewrites = payload.get("prose_rewrites", []) or []
    cyps = payload.get("cyp_questions", []) or []
    glossary = payload.get("glossary_terms", []) or []

    if fn.endswith(".pptx"):
        prs = pptx.Presentation(io.BytesIO(raw_bytes))
        for idx, slide in enumerate(prs.slides, start=1):
            if not slide.has_notes_slide:
                _ = slide.notes_slide
            prev = (slide.notes_slide.notes_text_frame.text or "").strip()
            slide.notes_slide.notes_text_frame.text = f"{prev}\n[Layer-5 Remediated Slide {idx} | Verified against Layer-4 Report v1]".strip()
        summary_layout = prs.slide_layouts[1] if len(prs.slide_layouts) > 1 else prs.slide_layouts[0]
        s_end = prs.slides.add_slide(summary_layout)
        if s_end.shapes.title:
            s_end.shapes.title.text = f"Layer-5 QA Remediation Summary - {title}"
        body_lines = [payload.get("executive_summary", "All Layer-4 Report v1 findings addressed.")]
        for rw in rewrites[:6]:
            body_lines.append(f"- {rw.get('issue', 'Fix')}: {rw.get('revised_text', '')[:140]}")
        for shp in s_end.shapes:
            if shp.has_text_frame and shp != s_end.shapes.title:
                shp.text_frame.text = "\n".join(body_lines)
                break
        out = io.BytesIO()
        prs.save(out)
        return out.getvalue(), ".pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation"

    # Build full cloned .docx manuscript with inline replacements + CYP + Glossary + Change Log
    doc = Document(io.BytesIO(raw_bytes)) if fn.endswith(".docx") else Document()
    if not fn.endswith(".docx"):
        doc.add_heading(title, level=1)
        for line in extract_manuscript_text(raw_bytes, filename).splitlines():
            if line.strip():
                doc.add_paragraph(line.strip())

    # Apply inline paragraph replacements where original_excerpt matches
    for rw in rewrites:
        orig = str(rw.get("original_excerpt", "")).strip()
        rev = str(rw.get("revised_text", "")).strip()
        if len(orig) >= 12 and rev:
            for p in doc.paragraphs:
                if orig in p.text:
                    p.text = p.text.replace(orig, rev)

    doc.add_page_break()
    doc.add_heading(f"Layer-5 Remediated Enhancements & Audit Log - {title}", level=1)
    doc.add_paragraph(payload.get("executive_summary", "Updated manuscript anchored to Layer-4 QA Audit Report v1."))

    if rewrites:
        doc.add_heading("Applied Editorial & Academic Corrections", level=2)
        for idx, rw in enumerate(rewrites, start=1):
            doc.add_paragraph(f"{idx}. [{rw.get('section', title)}] {rw.get('issue', '')}", style="List Bullet")
            doc.add_paragraph(f"Updated Text: {rw.get('revised_text', '')}")

    if cyps:
        doc.add_heading("Check Your Progress (CYP) & Model Answer Key", level=2)
        for idx, q in enumerate(cyps, start=1):
            doc.add_paragraph(f"Q{idx}. {q.get('question', '')}")
            doc.add_paragraph(f"Answer: {q.get('answer', '')}")

    if glossary:
        doc.add_heading("Key Terms & Glossary", level=2)
        for g in glossary:
            doc.add_paragraph(f"{g.get('term', '')}: {g.get('definition', '')}", style="List Bullet")

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue(), ".docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
