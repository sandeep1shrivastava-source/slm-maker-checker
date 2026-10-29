import json
from concurrent.futures import ThreadPoolExecutor
import google.generativeai as genai
import streamlit as st

MODEL_ID = "gemini-3.8-flash"

def get_gemini_model(system_instruction: str, json_mode: bool = False, temp: float = 0.0):
    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
    cfg = {"temperature": temp}
    if json_mode:
        cfg["response_mime_type"] = "application/json"
    return genai.GenerativeModel(
        model_name=MODEL_ID,
        system_instruction=system_instruction,
        generation_config=cfg
    )

# -------------------------------------------------------------
# STAGE 1: AI MAKER ENGINE (5 Exact Review Formats)
# -------------------------------------------------------------
def run_stage_1_maker(syllabus: str, reference: str, level: str, deliverable: str) -> str:
    """Routes generation to the exact prompt architecture matching the selected review format."""
    from prompts.academic_formats import MCQ_PROMPT, PPT_PROMPT, FAQS_MISCONCEPTIONS_PROMPT

    user_payload = (
        f"Academic Level: {level}\n\n"
        f"Approved Syllabus Outline:\n{syllabus}\n\n"
        f"Prescribed Reference Text:\n{reference}"
    )

    if deliverable == "MCQ":
        sys_prompt = MCQ_PROMPT
    elif deliverable == "Question Bank":
        from prompts.qb_prompts import QB_MAKER_SYSTEM_PROMPT
        sys_prompt = QB_MAKER_SYSTEM_PROMPT
    elif deliverable == "PPT":
        sys_prompt = PPT_PROMPT
    elif deliverable == "FAQs & Misconceptions Review":
        sys_prompt = FAQS_MISCONCEPTIONS_PROMPT
    else:
        # Default: SLM (Self-Learning Material Quadrant-II)
        from prompts.maker_prompts import MAKER_SYSTEM_PROMPT, MAKER_USER_TEMPLATE
        sys_prompt = MAKER_SYSTEM_PROMPT
        user_payload = MAKER_USER_TEMPLATE.format(academic_level=level, syllabus=syllabus, reference=reference)

    model = get_gemini_model(sys_prompt, json_mode=False, temp=0.35)
    response = model.generate_content(user_payload)
    return response.text

# -------------------------------------------------------------
# STAGE 2: ACADEMIC QUALITY & SYLLABUS CHECKER
# -------------------------------------------------------------
def run_stage_2_checker(draft: str, syllabus: str, reference: str) -> dict:
    from prompts.checker_prompts import CHECKER_SYSTEM_PROMPT
    from prompts.mapping_prompts import MAPPING_SYSTEM_PROMPT

    def audit_quality():
        payload = f"Syllabus:\n{syllabus}\n\nReferences:\n{reference}\n\nDraft:\n{draft}"
        model = get_gemini_model(CHECKER_SYSTEM_PROMPT, json_mode=True, temp=0.0)
        return json.loads(model.generate_content(payload).text)

    def audit_mapping():
        payload = f"Syllabus:\n{syllabus}\n\nDraft:\n{draft}"
        model = get_gemini_model(MAPPING_SYSTEM_PROMPT, json_mode=True, temp=0.0)
        return json.loads(model.generate_content(payload).text)

    with ThreadPoolExecutor(max_workers=2) as executor:
        f_checker = executor.submit(audit_quality)
        f_mapping = executor.submit(audit_mapping)
        checker_data = f_checker.result()
        mapping_data = f_mapping.result()

    score = checker_data.get("audit_metadata", {}).get("overall_score", 0.0)
    errors = checker_data.get("error_register", [])
    critical_errors = [e for e in errors if e.get("severity", "").lower() == "critical"]
    coverage_pct = mapping_data.get("coverage_percentage", 0.0)

    decision = "PASS"
    if score < 95.0 or len(critical_errors) > 0 or coverage_pct < 100.0:
        decision = "REVISE" if score >= 85.0 else "SME REVIEW REQUIRED"

    return {
        "decision": decision,
        "score": score,
        "coverage_pct": coverage_pct,
        "critical_count": len(critical_errors),
        "error_register": errors,
        "syllabus_mapping": mapping_data.get("syllabus_mapping", [])
    }

# -------------------------------------------------------------
# STAGE 3: DEEP PROMPT AUDITOR & INSTRUCTION COMPLIANCE
# -------------------------------------------------------------
PROMPT_AUDITOR_SYSTEM = """You are a Principal AI Prompt Compliance Auditor for University Educational Content.
Perform a strict Meta-Audit on the provided manuscript against generation prompt constraints and guardrails.

EVALUATION PARAMETERS:
1. Negative Constraints: Detect conversational filler ("Certainly!", "As an AI...", "Here is the unit..."), bracket placeholders ("[Insert diagram]"), or ungrounded claims.
2. Section Architecture: Confirm all structural sections are present in sequence without omission.
3. Bloom's Taxonomy Verbs: Verify whether objectives and question prompts utilize measurable action verbs.
4. Academic Tone Calibration: Assess whether the delivery is calibrated to higher-education self-learning standards.

OUTPUT FORMAT:
Return ONLY a valid JSON object matching this schema:
{
  "prompt_score": 96.5,
  "overall_compliance_status": "COMPLIANT",
  "negative_constraint_violations": 0,
  "compliance_checks": [
    {"dimension": "Negative Constraints & Boilerplate", "status": "PASS", "finding": "Zero conversational boilerplate or placeholder tokens detected.", "severity": "None"},
    {"dimension": "Structural Section Architecture", "status": "PASS", "finding": "All mandatory pedagogical sections present in valid order.", "severity": "None"},
    {"dimension": "Bloom's Taxonomy Verbs", "status": "PASS", "finding": "Measurable cognitive action verbs applied consistently.", "severity": "None"},
    {"dimension": "Pedagogical Tone Calibration", "status": "PASS", "finding": "Academic tone strictly adheres to university Quadrant-II specifications.", "severity": "None"}
  ],
  "structural_drift_analysis": "Zero prompt drift observed. Manuscript adhered strictly to university authoring parameters."
}"""

def run_stage_3_prompt_audit(draft: str, deliverable: str) -> dict:
    model = get_gemini_model(PROMPT_AUDITOR_SYSTEM, json_mode=True, temp=0.0)
    user_payload = f"Deliverable Format: {deliverable}\n\nDraft Text to Audit:\n{draft}"
    response = model.generate_content(user_payload)
    return json.loads(response.text)