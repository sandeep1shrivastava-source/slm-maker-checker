import json
import os

CONFIG_FILE = "access_control.json"

DEFAULT_CONFIG = {
    "universities": [
        "State Open University",
        "National Institute of Distance Education",
        "Apex Central University"
    ],
    "users": {
        "admin@university.edu": {"role": "Lead Administrator", "status": "Active"},
        "maker@university.edu": {"role": "Maker (Author)", "status": "Active"},
        "checker@university.edu": {"role": "Checker (Auditor)", "status": "Active"},
        "prompt-auditor@university.edu": {"role": "Prompt Engineer", "status": "Active"}
    }
}

def load_access_config() -> dict:
    """Loads access configuration or creates default JSON if not found."""
    if not os.path.exists(CONFIG_FILE):
        save_access_config(DEFAULT_CONFIG)
        return DEFAULT_CONFIG
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return DEFAULT_CONFIG

def save_access_config(config_data: dict):
    """Saves updated universities and user assignments."""
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config_data, f, indent=4)

def add_university(university_name: str) -> bool:
    """Appends a new university to the dropdown list."""
    config = load_access_config()
    clean_name = university_name.strip()
    if clean_name and clean_name not in config["universities"]:
        config["universities"].append(clean_name)
        save_access_config(config)
        return True
    return False

def assign_user_role(email: str, role: str) -> bool:
    """Adds or updates an authorized user's email and role."""
    config = load_access_config()
    clean_email = email.strip().lower()
    if clean_email:
        config["users"][clean_email] = {"role": role, "status": "Active"}
        save_access_config(config)
        return True
    return False

def verify_user(email: str):
    """Verifies if an email is authorized and returns their assigned role."""
    config = load_access_config()
    clean_email = email.strip().lower()
    user_record = config["users"].get(clean_email)
    if user_record and user_record.get("status") == "Active":
        return True, user_record.get("role")
    return False, None