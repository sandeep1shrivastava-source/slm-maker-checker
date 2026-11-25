import runpy, streamlit as st
st.session_state["_app_layer_mode"]="l5"
runpy.run_path("layer3_doc_auditor.py",run_name="__main__")
