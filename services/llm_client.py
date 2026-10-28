import json
from concurrent.futures import ThreadPoolExecutor
import google.generativeai as genai
import streamlit as st

# Configure Gemini API client
genai.configure(api_key=st.secrets["GEMINI_API_KEY"])

MODEL_ID = "gemini-3.8-flash"

def call_llm(system_prompt: str, user_content: str, json_mode: bool = False) -> str:
    generation_config = {"temperature": 0.0}
    if json_mode:
        generation_config["response_mime_type"] = "application/json"

    model = genai.GenerativeModel(
        model_name=MODEL_ID,
        system_instruction=system_prompt,
        generation_config=generation_config
    )
    response = model.generate_content(user_content)
    return response.text

def run_maker_generation(syllabus: str, reference: str, level: str) -> str:
    from prompts.maker_prompts import MAKER_SYSTEM_PROMPT, MAKER_USER_TEMPLATE
    user_payload = MAKER_USER_TEMPLATE.format(
        academic_level=level, syllabus=syllabus, reference=reference
    )
    model = genai.GenerativeModel(
        model_name=MODEL_ID,
        system_instruction=MAKER_SYSTEM_PROMPT,
        generation_config={"temperature": 0.35}
    )
    response = model.generate_content(user_payload)
    return response.text

def run_stage_3_audits(draft: str, syllabus: str, reference: str) -> dict:
    from prompts.checker_prompts import CHECKER_SYSTEM_PROMPT
    from prompts.mapping_prompts import MAPPING_SYSTEM_PROMPT
    from prompts.meta_prompts import META_SYSTEM_PROMPT

    def run_checker():
        payload = f"Syllabus:\n{syllabus}\n\nReferences:\n{reference}\n\nDraft:\n{draft}"
        res = call_llm(CHECKER_SYSTEM_PROMPT, payload, json_mode=True)
        return json.loads(res)

    def run_mapping():
        payload = f"Syllabus:\n{syllabus}\n\nDraft:\n{draft}"
        res = call_llm(MAPPING_SYSTEM_PROMPT, payload, json_mode=True)
        return json.loads(res)

    def run_meta():
        payload = f"Draft to analyze:\n{draft}"
        res = call_llm(META_SYSTEM_PROMPT, payload, json_mode=True)
        return json.loads(res)

    # Concurrently execute all three audit streams
    with ThreadPoolExecutor(max_workers=3) as executor:
        f_checker = executor.submit(run_checker)
        f_mapping = executor.submit(run_mapping)
        f_meta = executor.submit(run_meta)

        checker_data = f_checker.result()
        mapping_data = f_mapping.result()
        meta_data = f_meta.result()

    score = checker_data.get("audit_metadata", {}).get("overall_score", 0.0)
    errors = checker_data.get("error_register", [])
    critical_errors = [e for e in errors if e.get("severity", "").lower() == "critical"]
    coverage_pct = mapping_data.get("coverage_percentage", 0.0)
    prompt_status = meta_data.get("prompt_adherence_status", "FAIL")

    decision = "PASS"
    if score < 95.0 or len(critical_errors) > 0 or coverage_pct < 100.0 or prompt_status == "FAIL":
        decision = "REVISE" if score >= 85.0 else "SME REVIEW REQUIRED"

    return {
        "decision": decision,
        "score": score,
        "coverage_pct": coverage_pct,
        "prompt_status": prompt_status,
        "critical_count": len(critical_errors),
        "error_register": errors,
        "syllabus_mapping": mapping_data.get("syllabus_mapping", []),
        "prompt_adherence_logs": meta_data.get("adherence_logs", [])
    }