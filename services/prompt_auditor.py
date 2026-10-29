import json
import google.generativeai as genai
import streamlit as st

MODEL_ID = "gemini-3.8-flash"

DEEP_PROMPT_AUDIT_PROMPT = """You are an Academic AI Prompt Adherence and Instruction Auditor.
Your task is to analyze the generated academic manuscript against prompt constraints and university standards.

EVALUATE THE FOLLOWING 4 DIMENSIONS:
1. Negative Constraints: Did the model output any conversational boilerplate ("As an AI...", "Sure, here is..."), placeholder tags ("[Insert diagram]"), or ungrounded claims?
2. Mandatory Section Hierarchy: Are all required sections present in correct logical sequence without omission?
3. Bloom's Taxonomy Verbs: Are learning objectives and question items constructed with explicit, measurable cognitive verbs?
4. Tone & Level Calibration: Does the writing maintain a formal, self-paced academic distance-learning tone calibrated to the target level?

Output strictly a valid JSON object matching this schema:
{
  "prompt_score": 96.0,
  "status": "COMPLIANT",
  "negative_constraint_violations": 0,
  "checks": [
    {"dimension": "Negative Constraints & Boilerplate", "status": "PASS", "detail": "No conversational filler or placeholder brackets detected."},
    {"dimension": "Mandatory Section Hierarchy", "status": "PASS", "detail": "All prescribed sections present in correct sequence."},
    {"dimension": "Bloom's Taxonomy Verbs", "status": "PASS", "detail": "Objectives and questions use measurable academic action verbs."},
    {"dimension": "Academic Tone Calibration", "status": "PASS", "detail": "Pedagogical tone calibrated correctly to university standards."}
  ],
  "drift_notes": "Manuscript strictly followed system instructions with zero instruction drift."
}"""

def run_deep_prompt_audit(content: str, deliverable_type: str) -> dict:
    """Executes standalone prompt compliance verification without disturbing main checker."""
    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
    model = genai.GenerativeModel(
        model_name=MODEL_ID,
        system_instruction=DEEP_PROMPT_AUDIT_PROMPT,
        generation_config={"temperature": 0.0, "response_mime_type": "application/json"}
    )
    user_payload = f"Deliverable Type: {deliverable_type}\n\nDraft Content to Audit:\n{content}"
    response = model.generate_content(user_payload)
    return json.loads(response.text)