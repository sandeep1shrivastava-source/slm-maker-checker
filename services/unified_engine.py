import json
import google.generativeai as genai
import streamlit as st

# High-speed flash models for large-document generation
FALLBACK_MODELS = [
    "gemini-2.5-flash",
    "gemini-2.0-flash",
    "gemini-1.5-flash",
    "gemini-3.8-flash"
]

def get_api_key_pool() -> list:
    """Safely extracts API keys whether defined as a single string, a list, or comma-separated values."""
    raw_keys = []
    for secret_name in ["GEMINI_API_KEYS", "GEMINI_API_KEY"]:
        if secret_name in st.secrets:
            val = st.secrets[secret_name]
            if isinstance(val, (list, tuple)):
                raw_keys.extend(val)
            elif isinstance(val, str):
                if "," in val:
                    raw_keys.extend(val.split(","))
                else:
                    raw_keys.append(val)

    # Flatten nested lists and ensure each key is a stripped string
    clean_keys = []
    for k in raw_keys:
        if isinstance(k, (list, tuple)):
            for sub_k in k:
                if isinstance(sub_k, str) and sub_k.strip():
                    clean_keys.append(sub_k.strip())
        elif isinstance(k, str) and k.strip():
            clean_keys.append(k.strip())

    # Remove duplicates while preserving priority order
    deduped = []
    for k in clean_keys:
        if k not in deduped:
            deduped.append(k)

    if not deduped:
        raise ValueError("No valid Gemini API key found in Streamlit secrets.")
    return deduped

def generate_with_resilience(system_instruction: str, user_payload: str, json_mode: bool = False, temp: float = 0.0) -> str:
    """
    Executes generation with streaming to prevent 504 Deadline Exceeded idle timeouts,
    paired with multi-model and multi-key failover for 429 quota resilience.
    """
    key_pool = get_api_key_pool()
    cfg = {"temperature": temp}
    if json_mode:
        cfg["response_mime_type"] = "application/json"

    last_error = None
    retryable_triggers = [
        "429", "quota", "resourceexhausted",
        "504", "deadline", "timeout", "timed out",
        "503", "unavailable"
    ]

    for model_name in FALLBACK_MODELS:
        for active_key in key_pool:
            try:
                genai.configure(api_key=active_key)
                model = genai.GenerativeModel(
                    model_name=model_name,
                    system_instruction=system_instruction,
                    generation_config=cfg
                )
                
                # Streaming keeps the HTTP connection active during long generations
                if not json_mode:
                    response = model.generate_content(
                        user_payload,
                        stream=True,
                        request_options={"timeout": 600.0}
                    )
                    text_accumulator = []
                    for chunk in response:
                        try:
                            if chunk.text:
                                text_accumulator.append(chunk.text)
                        except Exception:
                            continue
                    if text_accumulator:
                        return "".join(text_accumulator)
                else:
                    response = model.generate_content(
                        user_payload,
                        stream=False,
                        request_options={"timeout": 300.0}
                    )
                    return response.text

            except Exception as e:
                err_str = str(e).lower()
                last_error = e
                if any(trigger in err_str for trigger in retryable_triggers):
                    continue
                else:
                    raise e

    raise RuntimeError(f"All models and keys exhausted. Last error: {last_error}")

# -------------------------------------------------------------
# STAGE 1: AI MAKER ENGINE (5 Exact Review Formats)
# -------------------------------------------------------------
def run_stage_1_maker(syllabus: str, reference: str, level: str, deliverable: str) -> str:
    from prompts.academic_formats import MCQ_PROMPT, PPT_PROMPT, FAQS_MISCONCEPTIONS_PROMPT

    # Guard against list inputs from document parsers
    if isinstance(syllabus, list):
        syllabus = "\n".join(str(s) for s in syllabus)
    if isinstance(reference, list):
        reference = "\n".join(str(r) for r in reference)

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
        # Default: SLM
        from prompts.maker_prompts import MAKER_SYSTEM_PROMPT, MAKER_USER_TEMPLATE
        sys_prompt = MAKER_SYSTEM_PROMPT
        user_payload = MAKER_USER_TEMPLATE.format(academic_level=level, syllabus=syllabus, reference=reference)

    return generate_with_resilience(sys_prompt, user_payload, json_mode=False, temp=0.35)

# -------------------------------------------------------------
# STAGE 2: ACADEMIC QUALITY & SYLLABUS CHECKER
# -------------------------------------------------------------
def run_stage_2_checker(draft: str, syllabus: str, reference: str) -> dict:
    from prompts.checker_prompts import CHECKER_SYSTEM_PROMPT
    from prompts.mapping_prompts import MAPPING_SYSTEM_PROMPT

    if isinstance(syllabus, list):
        syllabus = "\n".join(str(s) for s in syllabus)
    if isinstance(reference, list):
        reference = "\n".join(str(r) for r in reference)

    def audit_quality():
        payload = f"Syllabus:\n{syllabus}\n\nReferences:\n{reference}\n\nDraft:\n{draft}"
        res_text = generate_with_resilience(CHECKER_SYSTEM_PROMPT, payload, json_mode=True, temp=0.0)
        return json.loads(res_text)

    def audit_mapping():
        payload = f"Syllabus:\n{syllabus}\n\nDraft:\n{draft}"
        res_text = generate_with_resilience(MAPPING_SYSTEM_PROMPT, payload, json_mode=True, temp=0.0)
        return json.loads(res_text)

    checker_data = audit_quality()
    mapping_data = audit_mapping()

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
1. Negative Constraints: Detect conversational filler ("Certainly!", "As an AI..."), bracket placeholders ("[Insert diagram]"), or ungrounded claims.
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
    user_payload = f"Deliverable Format: {deliverable}\n\nDraft Text to Audit:\n{draft}"
    res_text = generate_with_resilience(PROMPT_AUDITOR_SYSTEM, user_payload, json_mode=True, temp=0.0)
    return json.loads(res_text)