import runpy
import streamlit as st

# Make st.set_page_config idempotent so both Layer-4 and Layer-5 run cleanly in one app
_ORIG_SET_PAGE_CONFIG = st.set_page_config
_PAGE_CONFIG_DONE = False


def _idempotent_set_page_config(*args, **kwargs):
    global _PAGE_CONFIG_DONE
    if not _PAGE_CONFIG_DONE:
        _PAGE_CONFIG_DONE = True
        return _ORIG_SET_PAGE_CONFIG(*args, **kwargs)
    return None


st.set_page_config = _idempotent_set_page_config
st.set_page_config(
    page_title="SLM Unified Studio — Layer-4 QA & Layer-5 Book Pipeline",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded",
)


def _sync_unified_auth_state() -> None:
    # Share authenticated email and OAuth session state seamlessly across both layers
    email_keys = ("auth_email", "user_email", "logged_in_email", "oauth_user_email", "google_user_email")
    active_email = ""
    for k in email_keys:
        v = st.session_state.get(k)
        if isinstance(v, str) and "@" in v:
            active_email = v.strip().lower()
            break
    if active_email:
        for k in email_keys:
            st.session_state.setdefault(k, active_email)


_sync_unified_auth_state()

st.sidebar.markdown("## 🎓 SLM Unified Workspace")
st.sidebar.caption("Combined Layer-4 QA + Layer-5 Studio")

if hasattr(st, "navigation") and hasattr(st, "Page"):
    pg = st.navigation(
        {
            "Unified Pipeline": [
                st.Page(
                    "layer3_doc_auditor.py",
                    title="Layer-4: Document QA Auditor",
                    icon="🛡️",
                    default=True,
                ),
                st.Page(
                    "layer5_app.py",
                    title="Layer-5: Enrichment & Book Studio",
                    icon="📚",
                ),
            ]
        }
    )
    pg.run()
else:
    selected_layer = st.sidebar.radio(
        "Active Layer",
        options=[
            "🛡️ Layer-4: Document QA Auditor",
            "📚 Layer-5: Enrichment & Book Studio",
        ],
        key="_unified_active_layer",
    )
    st.sidebar.divider()
    target_script = "layer3_doc_auditor.py" if "Layer-4" in selected_layer else "layer5_app.py"
    runpy.run_path(target_script, run_name="__main__")
