import os
import json
import uuid
from datetime import datetime

DEP_FILE = os.path.join(os.path.dirname(__file__), "manuscript_depository.json")

def _load_depository():
    if not os.path.exists(DEP_FILE):
        return []
    try:
        with open(DEP_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []

def _save_depository(data):
    try:
        with open(DEP_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4)
    except Exception as e:
        print(f"Error saving depository: {e}")

def list_manuscripts():
    return _load_depository()

def save_manuscript_draft(university, program, batch, deliverable, author_email, draft_text):
    data = _load_depository()
    manuscript_id = f"MANU-{datetime.now().strftime('%Y%m%d')}-{uuid.uuid4().hex[:6].upper()}"
    entry = {
        "id": manuscript_id,
        "university": university,
        "program": program,
        "batch": batch,
        "deliverable": deliverable,
        "author_email": author_email,
        "draft_text": draft_text,
        "status": "Pending Review",
        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "stage2_results": None,
        "stage3_results": None
    }
    data.insert(0, entry)
    _save_depository(data)
    return manuscript_id

def update_manuscript_review(manuscript_id, reviewer_email, stage2_results=None, stage3_results=None, status="In Review"):
    data = _load_depository()
    for item in data:
        if item.get("id") == manuscript_id:
            if stage2_results is not None:
                item["stage2_results"] = stage2_results
            if stage3_results is not None:
                item["stage3_results"] = stage3_results
            item["reviewer_email"] = reviewer_email
            item["status"] = status
            item["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            break
    _save_depository(data)