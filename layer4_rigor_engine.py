import io
import math
import re
from collections import Counter
from typing import Dict, List
from docx import Document


def compute_prose_burstiness(prose_paragraphs: List[str]) -> Dict:
    """Compute Burstiness CV strictly on running prose sentences (>= 3 words)."""
    sentence_lengths = []
    for para in prose_paragraphs:
        for sent in re.split(r"(?<=[.!?])\s+", para.strip()):
            words = re.findall(r"\b[A-Za-z0-9'-]+\b", sent)
            if len(words) >= 3:
                sentence_lengths.append(len(words))

    count = len(sentence_lengths)
    if count < 2:
        return {
            "sentence_count": count,
            "mean_words": 0.0,
            "sd_words": 0.0,
            "prose_cv": 0.0,
        }

    mean_len = sum(sentence_lengths) / count
    variance = sum((x - mean_len) ** 2 for x in sentence_lengths) / (count - 1)
    sd_len = math.sqrt(variance)
    cv = round(sd_len / mean_len, 2) if mean_len > 0 else 0.0

    return {
        "sentence_count": count,
        "mean_words": round(mean_len, 1),
        "sd_words": round(sd_len, 1),
        "prose_cv": cv,
    }


def run_text_forensic_precheck(
    raw_text: str, doc_type: str = "SLM", min_cv: float = 0.40
) -> Dict:
    """Forensic pre-check on extracted manuscript text (prose burstiness + UGC elements)."""
    doc_type_upper = doc_type.strip().upper()
    lines = [ln.strip() for ln in raw_text.splitlines() if ln.strip()]
    prose_paragraphs = [
        ln
        for ln in lines
        if len(ln.split()) >= 12
        and not re.match(r"^(\d+[\.\)]|[a-dA-D][\.\)]|[\-\*\u2022])\s+", ln)
    ]

    full_lower = raw_text.lower()
    tokens = re.findall(r"\b[a-z]{4,}\b", full_lower)
    freq = Counter(tokens)
    buzzwords = {
        w: freq[w]
        for w in ("modern", "essential", "organizations", "crucial", "dynamic")
        if freq[w] >= 8
    }

    burstiness = compute_prose_burstiness(prose_paragraphs)
    failures = []

    if burstiness["prose_cv"] < min_cv:
        failures.append(
            f"Prose Burstiness CV is {burstiness['prose_cv']} (below minimum {min_cv})."
        )
    if doc_type_upper == "SLM":
        if "check your progress" not in full_lower and "self-assessment" not in full_lower:
            failures.append("Missing in-text 'Check Your Progress' / Self-Assessment exercises.")
        if (
            "answer key" not in full_lower
            and "answers to" not in full_lower
            and "model answer" not in full_lower
        ):
            failures.append("Missing Answer Key / model solutions for exercises.")

    return {
        "doc_type": doc_type_upper,
        "word_count": len(raw_text.split()),
        "heading_style_count": None,
        "normal_style_count": None,
        "image_count": None,
        "table_count": None,
        "metadata_complete": None,
        "stray_artifacts_count": 0,
        "buzzword_spikes": buzzwords,
        "burstiness": burstiness,
        "part_c_failures": failures,
        "force_major_changes": len(failures) > 0,
    }


def run_docx_forensic_precheck(
    file_bytes: bytes, doc_type: str = "SLM", min_cv: float = 0.40
) -> Dict:
    """Inspect .docx XML structure, styles, images, metadata, and prose burstiness."""
    doc = Document(io.BytesIO(file_bytes))
    doc_type_upper = doc_type.strip().upper()

    heading_style_count = 0
    normal_style_count = 0
    prose_paragraphs = []
    all_lines = []
    stray_artifacts = []

    for p in doc.paragraphs:
        txt = p.text.strip()
        if not txt:
            continue
        all_lines.append(txt)

        if "top of form" in txt.lower() or "bottom of form" in txt.lower():
            stray_artifacts.append(txt)

        style_name = (p.style.name if p.style else "").lower()
        is_list = "list" in style_name or bool(
            re.match(r"^(\d+[\.\)]|[a-dA-D][\.\)]|[\-\*\u2022])\s+", txt)
        )

        if style_name.startswith("heading"):
            heading_style_count += 1
        else:
            if style_name == "normal":
                normal_style_count += 1
            if not is_list and len(txt.split()) >= 12:
                prose_paragraphs.append(txt)

    full_text = "\n".join(all_lines)
    full_lower = full_text.lower()
    image_count = len(doc.inline_shapes)
    table_count = len(doc.tables)

    title_meta = (doc.core_properties.title or "").strip()
    subject_meta = (doc.core_properties.subject or "").strip()

    tokens = re.findall(r"\b[a-z]{4,}\b", full_lower)
    freq = Counter(tokens)
    buzzwords = {
        w: freq[w]
        for w in ("modern", "essential", "organizations", "crucial", "dynamic")
        if freq[w] >= 8
    }

    burstiness = compute_prose_burstiness(prose_paragraphs)
    failures = []

    if burstiness["prose_cv"] < min_cv:
        failures.append(
            f"Prose Burstiness CV is {burstiness['prose_cv']} (below minimum {min_cv})."
        )
    if heading_style_count == 0:
        failures.append(
            "Zero Word Heading styles (Heading 1/2/3) used; all paragraphs formatted as Normal."
        )
    if stray_artifacts:
        failures.append(
            f"Stray form artifacts detected: {len(stray_artifacts)} occurrence(s)."
        )
    if doc_type_upper == "SLM":
        if image_count == 0:
            failures.append("Zero embedded figures/diagrams found in SLM.")
        if "check your progress" not in full_lower and "self-assessment" not in full_lower:
            failures.append("Missing in-text 'Check Your Progress' / Self-Assessment exercises.")
        if (
            "answer key" not in full_lower
            and "answers to" not in full_lower
            and "model answer" not in full_lower
        ):
            failures.append("Missing Answer Key / model solutions for exercises.")

    return {
        "doc_type": doc_type_upper,
        "word_count": len(full_text.split()),
        "heading_style_count": heading_style_count,
        "normal_style_count": normal_style_count,
        "image_count": image_count,
        "table_count": table_count,
        "metadata_complete": bool(title_meta and subject_meta),
        "stray_artifacts_count": len(stray_artifacts),
        "buzzword_spikes": buzzwords,
        "burstiness": burstiness,
        "part_c_failures": failures,
        "force_major_changes": len(failures) > 0,
    }


def build_strict_audit_prompt(
    doc_type: str,
    manuscript_text: str,
    syllabus_text: str,
    precheck: Dict,
    thresholds: Dict,
) -> str:
    """Return isolated, non-bleeding audit prompt for SLM vs MCQ/QB."""
    dtype = doc_type.strip().upper()
    approved_min = thresholds.get("approved_min", 85.0)
    minor_min = thresholds.get("minor_min", 70.0)
    max_ai = thresholds.get("max_ai_pct", 20.0)

    common_header = f"""You are a Forensic Academic Quality Assurance Auditor.
Audit the manuscript strictly against the provided syllabus and deterministic pre-check evidence.

ACTIVE THRESHOLDS:
- Approved >= {approved_min}% | Minor Changes >= {minor_min}% | Max Allowed AI Content <= {max_ai}%

DETERMINISTIC PRE-CHECK EVIDENCE (VERIFIED BY PYTHON):
{precheck}

EPISTEMIC HONESTY & ANTI-INFLATION RULES:
1. Never award >70% on any parameter unless you cite specific section numbers and concrete evidence from the text.
2. Plagiarism / Similarity MUST be reported as: "NOT RUN (Requires Institutional Similarity Tool)". Never guess "Level 0".
3. Verify technical math, formulas, domain mappings, and bibliographic citations critically; flag any inconsistency.
4. If 'force_major_changes' is True in the pre-check evidence, final verdict MUST be "Major Changes".
"""

    if dtype in ("MCQ", "QB"):
        rubric_block = """
RUBRIC MODE: 7-POINT MCQ / QUESTION BANK DEDUCTION MATRIX (Start at 100%, subtract penalties; NEVER use SLM rubric):
1. Subject Knowledge / Factual Errors (Max -30%)
2. Technical Errors / Letter-vs-Answer Key Mismatch (Max -30%)
3. Distractor Plausibility & Length Bias (Max -15%)
4. Clarity & Duplicate/Repeated Questions (Max -10%)
5. Bloom's Taxonomy & Difficulty Spread Mismatch (Max -5%)
6. Grammar, Semantics & Negative Stem Formatting (Max -5%)
7. Syllabus / Unit Coverage Drift (Max -5%)
"""
    else:
        rubric_block = """
RUBRIC MODE: 12-PARAMETER UGC-DEB SLM RUBRIC (Exact 100-Point Weightage; NEVER include MCQ deduction rows):
1. Syllabus & Curriculum Mapping (Max 10)
2. Instructional Design & Structure (Max 10)
3. Conceptual Depth & Academic Rigor (Max 10)
4. Language, Tone & Editorial Clarity (Max 10)
5. Real-World Examples & Caselets (Max 10)
6. Self-Assessment & In-Text Exercises (Max 8)
7. Summaries, Keywords & Key Takeaways (Max 8)
8. Formatting, Typography & Presentation (Max 8)
9. Visual Diagrams, Tables & Graphical Quality (Max 8)
10. References, Citations & Plagiarism Transparency (Max 8)
11. Pedagogical Alignment & Intended LOs (Max 5)
12. Legal, Ethical & Sensitivity Standards (Max 5)
"""

    parts = [common_header, rubric_block]
    if syllabus_text.strip():
        parts.append(f"REFERENCE SYLLABUS:\n{syllabus_text}")
    if manuscript_text.strip():
        parts.append(f"MANUSCRIPT TEXT:\n{manuscript_text}")
    # Dual-Mode Syllabus Scoping Safeguard (Single-Unit vs. Full-Course)
    _unit_scope_rule = (
        "\n=== SYLLABUS SCOPING RULE: SINGLE-UNIT VS. FULL-COURSE (MANDATORY) ===\n"
        "- Detect Manuscript Scope First: Check the Report Title, Manuscript Title, and document headings to determine whether this file is (A) a Single Chapter / Unit / Module manuscript, or (B) a Full-Course / Complete Bank manuscript covering all units.\n"
        "- Case A (Single Chapter / Unit File): Locate ONLY the matching chapter/unit in the full-course syllabus and grade Parameter 1 strictly on that unit's topics. Do NOT deduct points or flag missing topics for other chapters/units that belong in separate files.\n"
        "- Case B (Full-Course / Complete Bank File, e.g., 'Complete Bank', 'Full Syllabus', or multi-unit compilation): Evaluate Parameter 1 across ALL units/modules in the syllabus and flag any uncovered units normally.\n"
    )
    _mcq_rule = (
        "\n=== DETERMINISTIC EXCEL MCQ GROUNDING RULE (MANDATORY) ===\n"
        "- Treat '=== DETERMINISTIC EXCEL MCQ FORENSIC PRE-AUDIT (GROUND TRUTH) ===' counts as binding ground truth.\n"
        "- If Verified Answer-Key Mismatches is 0, award 0% deduction for Technical Errors (letter vs answer mismatch) and do NOT invent key mismatches or misread Option C as the answer key.\n"
        "- Cite exact [Sheet | Row | Sr No] tags and report verified duplicate stems and Sr No skips under Clarity & Repetition.\n"
    )
    return (("\n\n".join(parts)) + _unit_scope_rule) + _mcq_rule


def enforce_post_audit_gates(audit_result: Dict, precheck: Dict) -> Dict:
    """Lock final verdict and plagiarism honesty rules after LLM response."""
    if precheck.get("force_major_changes"):
        audit_result["verdict"] = "Major Changes"
    audit_result["plagiarism_status"] = "NOT RUN (Requires Institutional Similarity Tool)"
    return audit_result

def run_mcq_xlsx_forensic_precheck(source) -> dict:
    from mcq_forensic_parser import format_mcq_audit_payload
    return format_mcq_audit_payload(source)
