import json
import os
import uuid
from datetime import datetime

DEPOSITORY_FILE = os.path.join(os.path.dirname(__file__), "depository.json")

def _init_depository():
    if not os.path.exists(DEPOSITORY_FILE):
        try:
            with open(DEPOSITORY_FILE, "w") as f:
                json.dump([], f)
        except Exception:
            pass

def list_manuscripts():
    _init_depository()
    try:
        with open(DEPOSITORY_FILE, "r") as f:
            data = json.load(f)
            # Ensure backward compatibility for any existing records
            for item in data:
                item.setdefault("batch", "General Cohort")
                item.setdefault("semester_year", "Semester 1")
            return data
    except Exception:
        return []

def save_manuscripts(data):
    try:
        with open(DEPOSITORY_FILE, "w") as f:
            json.dump(data, f, indent=4)
    except Exception as e:
        print(f"Error saving depository: {e}")

def park_manuscript(title, author_email, university, program, batch, semester_year, deliverable, level, syllabus_text, ref_text, draft_text):
    data = list_manuscripts()
    item_id = str(uuid.uuid4())[:8]
    record = {
        "id": item_id,
        "title": title or f"{university} | {program} | {batch} | {semester_year} - {deliverable}",
        "author_email": author_email,
        "university": university,
        "program": program,
        "batch": batch,
        "semester_year": semester_year,
        "deliverable": deliverable,
        "level": level,
        "syllabus_text": syllabus_text,
        "ref_text": ref_text,
        "draft_text": draft_text,
        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "status": "Pending Review",
        "stage2_results": None,
        "stage3_results": None,
        "reviewer_email": None,
        "reviewed_at": None
    }
    data.insert(0, record)
    save_manuscripts(data)
    return item_id

def update_manuscript_review(item_id, reviewer_email, stage2_results=None, stage3_results=None):
    data = list_manuscripts()
    for item in data:
        if item["id"] == item_id:
            if stage2_results:
                item["stage2_results"] = stage2_results
                item["status"] = stage2_results.get("decision", "Reviewed")
            if stage3_results:
                item["stage3_results"] = stage3_results
            item["reviewer_email"] = reviewer_email
            item["reviewed_at"] = datetime.now().strftime("%Y-%m-%d %H:%M")
            break
    save_manuscripts(data)