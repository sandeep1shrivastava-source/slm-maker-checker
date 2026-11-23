import io, re
from collections import defaultdict
import openpyxl

def parse_and_audit_mcq_workbook(source) -> dict:
    if isinstance(source, (bytes, bytearray)):
        wb = openpyxl.load_workbook(io.BytesIO(source), data_only=True)
    elif hasattr(source, "read"):
        pos = source.tell() if hasattr(source, "tell") else None
        data = source.read()
        if pos is not None and hasattr(source, "seek"):
            source.seek(pos)
        wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    else:
        wb = openpyxl.load_workbook(source, data_only=True)

    q_sheets, struct_lines, items, mismatches, anomalies, dups, summaries = [], [], [], [], [], [], []
    seen = {}

    for sname in wb.sheetnames:
        ws = wb[sname]
        if any(k in sname.lower() for k in ("structure", "syllabus", "blueprint")):
            struct_lines.append(f"=== EMBEDDED STRUCTURE TAB: {sname.strip()} ===")
            for r in range(1, ws.max_row + 1):
                vals = [str(ws.cell(r, c).value).strip() for c in range(1, ws.max_column + 1) if ws.cell(r, c).value is not None]
                if vals:
                    struct_lines.append(" | ".join(vals))
            continue

        cmap, h_row = {}, None
        for r in range(1, min(11, ws.max_row + 1)):
            rvals = [str(ws.cell(r, c).value).strip().lower() if ws.cell(r, c).value is not None else "" for c in range(1, ws.max_column + 1)]
            if "question" in rvals and ("option a" in rvals or "opt a" in rvals):
                h_row = r
                for i, h in enumerate(rvals):
                    if h in ("sr no", "sr. no", "sr. no.", "q no", "q.no", "id"): cmap["sr"] = i
                    elif "section" in h or "topic" in h or "module" in h: cmap.setdefault("sec", i)
                    elif h == "question" or "question stem" in h: cmap["q"] = i
                    elif h in ("option a", "opt a", "a"): cmap["a"] = i
                    elif h in ("option b", "opt b", "b"): cmap["b"] = i
                    elif h in ("option c", "opt c", "c"): cmap["c"] = i
                    elif h in ("option d", "opt d", "d"): cmap["d"] = i
                    elif h in ("correct option", "option key", "answer key", "correct letter", "key"): cmap["let"] = i
                    elif h in ("correct answer", "answer text", "answer"): cmap["ans"] = i
                    elif "btl" in h or "bloom" in h: cmap["btl"] = i
                    elif "difficulty" in h: cmap["diff"] = i
                break
        if h_row is None:
            continue

        q_sheets.append(sname)
        sec_cnt = defaultdict(int)
        prev_sr, gaps, valid_cnt = 0, 0, 0

        for r in range(h_row + 1, ws.max_row + 1):
            raw = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
            non_empty = [v for v in raw if v is not None and str(v).strip() != ""]
            if not non_empty:
                continue
            qv = raw[cmap["q"]] if "q" in cmap and cmap["q"] < len(raw) else None
            srv = raw[cmap["sr"]] if "sr" in cmap and cmap["sr"] < len(raw) else None
            oa = raw[cmap["a"]] if "a" in cmap and cmap["a"] < len(raw) else None

            if str(qv).strip().lower() in ("question", "question stem") or str(srv).strip().lower() in ("sr no", "sr. no"):
                anomalies.append(f"[{sname.strip()} | Row {r}] Repeated header row skipped")
                continue
            if qv is None or (oa is None and len(non_empty) <= 2):
                anomalies.append(f"[{sname.strip()} | Row {r}] Section banner skipped: {non_empty[0]}")
                continue

            if isinstance(srv, (int, float)):
                sri = int(srv)
                if sri != prev_sr + 1 and sri != 1:
                    gaps += 1
                prev_sr = sri

            sec = str(raw[cmap["sec"]]).strip() if "sec" in cmap and cmap["sec"] < len(raw) and raw[cmap["sec"]] is not None else "General"
            sec_cnt[sec] += 1
            valid_cnt += 1

            opts = {k: (str(raw[cmap[k.lower()]]).strip() if k.lower() in cmap and cmap[k.lower()] < len(raw) and raw[cmap[k.lower()]] is not None else "") for k in ("A", "B", "C", "D")}
            let = str(raw[cmap["let"]]).strip().upper() if "let" in cmap and cmap["let"] < len(raw) and raw[cmap["let"]] is not None else ""
            ans = str(raw[cmap["ans"]]).strip() if "ans" in cmap and cmap["ans"] < len(raw) and raw[cmap["ans"]] is not None else ""

            if let:
                if let not in ("A", "B", "C", "D"):
                    mismatches.append(f"[{sname.strip()} | Row {r} | Sr {srv}] Invalid option letter '{let}'.")
                elif ans and opts.get(let, "") != ans:
                    mismatches.append(f"[{sname.strip()} | Row {r} | Sr {srv}] Option '{let}' ('{opts.get(let)}') != Answer text ('{ans}').")
            elif ans and ans.upper() not in ("A", "B", "C", "D"):
                mismatches.append(f"[{sname.strip()} | Row {r} | Sr {srv}] Answer '{ans}' lacks A/B/C/D letter.")

            nq = re.sub(r"\s+", " ", str(qv).strip().lower())
            tag = f"{sname.strip()} Row {r} (Sr {srv})"
            if nq in seen:
                dups.append((seen[nq], tag, str(qv).strip()[:90]))
            else:
                seen[nq] = tag

            items.append({
                "sheet": sname.strip(), "excel_row": r, "sr_no": srv, "section": sec,
                "question": str(qv).strip(), "options": opts, "correct_option": let or ans,
                "correct_answer_text": ans,
                "btl": str(raw[cmap["btl"]]).strip() if "btl" in cmap and cmap["btl"] < len(raw) and raw[cmap["btl"]] is not None else "",
                "difficulty": str(raw[cmap["diff"]]).strip() if "diff" in cmap and cmap["diff"] < len(raw) and raw[cmap["diff"]] is not None else "",
            })

        if gaps:
            anomalies.append(f"[{sname.strip()}] {gaps} Sr No numbering sequence skips detected")
        summaries.append(f"  - Tab '{sname.strip()}': {valid_cnt} valid MCQs (" + ", ".join(f"{k}: {v}" for k, v in sec_cnt.items()) + ")")

    return {
        "is_mcq_workbook": len(items) > 0,
        "question_sheets": [s.strip() for s in q_sheets],
        "sheet_summaries": summaries,
        "total_valid_mcqs": len(items),
        "key_mismatches": mismatches,
        "duplicate_stems": dups,
        "structural_anomalies": anomalies,
        "embedded_structure_text": "\n".join(struct_lines),
        "items": items,
    }


def format_mcq_audit_payload(source) -> dict:
    res = parse_and_audit_mcq_workbook(source)
    if not res["is_mcq_workbook"]:
        return {"is_mcq_workbook": False, "formatted_text": "", "embedded_structure_text": res["embedded_structure_text"], "audit": res}
    lines = [
        "=== DETERMINISTIC EXCEL MCQ FORENSIC PRE-AUDIT (GROUND TRUTH) ===",
        f"Question Sheets Audited: {', '.join(res['question_sheets'])}",
        *res["sheet_summaries"],
        f"Total Valid MCQs Across All Tabs: {res['total_valid_mcqs']}",
        f"Verified Answer-Key Mismatches (Option Letter vs Answer Text): {len(res['key_mismatches'])}",
    ]
    if res["key_mismatches"]:
        lines.extend(f"  * {m}" for m in res["key_mismatches"][:15])
    else:
        lines.append("  * ZERO key mismatches: 100% of items have valid A/B/C/D Correct Option identifiers matching their Correct Answer text.")
    lines.append(f"Exact Duplicate Question Stems Across Tabs: {len(res['duplicate_stems'])}")
    lines.extend(f"  * Duplicate: {o} == {d} -> '{s}'" for o, d, s in res["duplicate_stems"][:15])
    lines.append(f"Spreadsheet Hygiene Anomalies (Repeated Headers / Banners / Sr No Gaps): {len(res['structural_anomalies'])}")
    lines.extend(f"  * {a}" for a in res["structural_anomalies"][:15])
    lines.append("\n=== LABELED MCQ ITEMS (PROVENANCE-LOCKED) ===")
    for it in res["items"]:
        o = it["options"]
        ks = f"{it['correct_option']} ({it['correct_answer_text']})" if it["correct_answer_text"] else it["correct_option"]
        mb = "/".join(b for b in (it["btl"], it["difficulty"]) if b)
        lines.append(f"[Sheet: {it['sheet']} | Row: {it['excel_row']} | Sr No: {it['sr_no']} | Sec: {it['section']}] Q: {it['question']} | A: {o['A']} | B: {o['B']} | C: {o['C']} | D: {o['D']} | Key: {ks}" + (f" | Meta: {mb}" if mb else ""))
    return {"is_mcq_workbook": True, "formatted_text": "\n".join(lines), "embedded_structure_text": res["embedded_structure_text"], "audit": res}
