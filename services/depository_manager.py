import streamlit as st
import time
from google.cloud import firestore
from google.oauth2 import service_account

# -------------------------------------------------------------
# 1. DATABASE CONNECTION (CACHED)
# -------------------------------------------------------------
@st.cache_resource
def get_db():
    """Authenticates with Google Cloud Firestore using Streamlit Secrets."""
    if "gcp_service_account" not in st.secrets:
        st.error("⚠️ Firestore credentials missing: 'gcp_service_account' block not found in Streamlit Secrets.")
        return None

    try:
        key_dict = dict(st.secrets["gcp_service_account"])
        creds = service_account.Credentials.from_service_account_info(key_dict)
        db = firestore.Client(credentials=creds, project=key_dict.get("project_id"))
        return db
    except Exception as e:
        st.error(f"Failed to connect to Google Firestore: {e}")
        return None

# -------------------------------------------------------------
# 2. SAVE DRAFT MANUSCRIPT (MAKER ACTION)
# -------------------------------------------------------------
def save_manuscript_draft(university: str, program: str, batch: str, deliverable: str, author_email: str, draft_text: str) -> str:
    """Saves a newly generated manuscript draft to the Google Firestore depository."""
    db = get_db()
    
    # Unique timestamped manuscript identifier
    doc_id = f"MANU-{time.strftime('%Y%m%d')}-{int(time.time()) % 100000:05d}"
    
    payload = {
        "id": doc_id,
        "university": university,
        "program": program,
        "batch": batch or "Default Cohort",
        "deliverable": deliverable,
        "author_email": author_email,
        "draft_text": draft_text,
        "status": "Pending Review",
        "created_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        "stage2_results": {}
    }
    
    if db is not None:
        try:
            db.collection("manuscripts").document(doc_id).set(payload)
            return doc_id
        except Exception as err:
            st.error(f"Error persisting draft to Firestore: {err}")
            return doc_id
    else:
        st.warning("Database unavailable. Manuscript not saved permanently.")
        return doc_id

# -------------------------------------------------------------
# 3. LIST ALL MANUSCRIPTS (CHECKER & ADMIN QUEUE)
# -------------------------------------------------------------
def list_manuscripts() -> list:
    """Retrieves all submitted manuscripts from Firestore sorted by most recent."""
    db = get_db()
    if db is None:
        return []

    try:
        docs = db.collection("manuscripts").order_by("created_at", direction=firestore.Query.DESCENDING).stream()
        manuscripts = []
        for doc in docs:
            data = doc.to_dict()
            manuscripts.append(data)
        return manuscripts
    except Exception as err:
        st.sidebar.error(f"Could not load manuscript queue: {err}")
        return []

# -------------------------------------------------------------
# 4. UPDATE REVIEW STATUS & AUDIT DOSSIER (CHECKER ACTION)
# -------------------------------------------------------------
def update_manuscript_review(manuscript_id: str, reviewer_email: str, status: str = None, stage2_results: dict = None):
    """Updates a manuscript record with the Checker's evaluation and status."""
    db = get_db()
    if db is None:
        return

    doc_ref = db.collection("manuscripts").document(manuscript_id)
    
    update_data = {
        "last_reviewed_by": reviewer_email,
        "updated_at": time.strftime("%Y-%m-%d %H:%M:%S")
    }
    
    if status:
        update_data["status"] = status
    if stage2_results:
        update_data["stage2_results"] = stage2_results
        
    try:
        doc_ref.set(update_data, merge=True)
    except Exception as err:
        st.error(f"Error updating manuscript review record in Firestore: {err}")