def call_llm(prompt: str, engine: str = "Google Gemini") -> str:
    """Bulletproof LLM caller: header-based auth, model fallbacks, clean diagnostics."""
    import os
    import json
    import urllib.request
    import urllib.error

    gemini_key = None
    openai_key = None

    # Retrieve keys from Streamlit Secrets
    try:
        import streamlit as st
        for k in ["GEMINI_API_KEY", "GOOGLE_API_KEY", "gemini_api_key", "google_api_key", "GEMINI_KEY"]:
            if k in st.secrets and str(st.secrets[k]).strip():
                gemini_key = str(st.secrets[k]).strip().strip("'\"")
                break
        for k in ["OPENAI_API_KEY", "openai_api_key", "OPENAI_KEY"]:
            if k in st.secrets and str(st.secrets[k]).strip():
                openai_key = str(st.secrets[k]).strip().strip("'\"")
                break
        if not gemini_key and "gemini" in st.secrets:
            try:
                gemini_key = str(st.secrets["gemini"]["api_key"]).strip().strip("'\"")
            except Exception:
                pass
        if not openai_key and "openai" in st.secrets:
            try:
                openai_key = str(st.secrets["openai"]["api_key"]).strip().strip("'\"")
            except Exception:
                pass
    except Exception:
        pass

    # Environment variable fallback
    if not gemini_key:
        for k in ["GEMINI_API_KEY", "GOOGLE_API_KEY", "gemini_api_key", "google_api_key"]:
            if os.environ.get(k) and os.environ.get(k).strip():
                gemini_key = os.environ.get(k).strip().strip("'\"")
                break
    if not openai_key:
        for k in ["OPENAI_API_KEY", "openai_api_key"]:
            if os.environ.get(k) and os.environ.get(k).strip():
                openai_key = os.environ.get(k).strip().strip("'\"")
                break

    api_errors = []

    # 1. Try Gemini with header-based auth (key is NEVER in URL)
    if gemini_key:
        models_to_try = ["gemini-1.5-flash", "gemini-2.0-flash", "gemini-1.5-pro"]
        for m_name in models_to_try:
            url = f"https://generativelanguage.googleapis.com/v1beta/models/{m_name}:generateContent"
            payload = {
                "contents": [
                    {
                        "role": "user",
                        "parts": [{"text": prompt}]
                    }
                ],
                "generationConfig": {
                    "responseMimeType": "application/json"
                }
            }
            req = urllib.request.Request(
                url,
                data=json.dumps(payload).encode("utf-8"),
                headers={
                    "Content-Type": "application/json",
                    "x-goog-api-key": gemini_key
                }
            )
            try:
                with urllib.request.urlopen(req, timeout=90) as response:
                    res_obj = json.loads(response.read().decode("utf-8"))
                    text_out = res_obj["candidates"][0]["content"]["parts"][0]["text"]
                    if text_out and text_out.strip():
                        return text_out
            except urllib.error.HTTPError as h_err:
                body = h_err.read().decode("utf-8", errors="ignore")
                api_errors.append(f"Gemini {m_name} (HTTP {h_err.code}): {body[:180]}")
            except Exception as e:
                api_errors.append(f"Gemini {m_name} Error: {str(e)[:180]}")

    # 2. Try OpenAI fallback
    if openai_key:
        for o_model in ["gpt-4o", "gpt-4o-mini"]:
            try:
                from openai import OpenAI
                client = OpenAI(api_key=openai_key)
                resp = client.chat.completions.create(
                    model=o_model,
                    messages=[{"role": "user", "content": prompt}],
                    response_format={"type": "json_object"}
                )
                text_out = resp.choices[0].message.content
                if text_out and text_out.strip():
                    return text_out
            except Exception as o_err:
                api_errors.append(f"OpenAI {o_model} Error: {str(o_err)[:180]}")

    error_summary = "\n".join([f"• {err}" for err in api_errors])
    raise RuntimeError(
        f"LLM API request could not complete across available models.\n\n"
        f"Diagnostic Details:\n{error_summary}\n\n"
        f"Gemini key present: {'Yes' if gemini_key else 'No'} | "
        f"OpenAI key present: {'Yes' if openai_key else 'No'}"
    )