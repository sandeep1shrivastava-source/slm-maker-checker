import hashlib
import os
import json

CONFIG_PATH_V2 = os.path.join(os.path.dirname(__file__), "access_config_v2.json")
ALLOWED_DOMAIN = "upgrad.com"
PRIMARY_ADMIN = "sandeep1.shrivastava@upgrad.com"
DEFAULT_ADMIN_PASS = "Admin@upgrad2026"

def hash_password(password: str, salt: str = None) -> tuple:
    """Generates a secure SHA-256 hash with an independent 16-byte hex salt."""
    if not salt:
        salt = os.urandom(16).hex()
    hash_obj = hashlib.sha256((salt + password).encode("utf-8"))
    return hash_obj.hexdigest(), salt

def verify_password(stored_hash: str, salt: str, password_attempt: str) -> bool:
    """Verifies user password attempt against the stored salt and hash."""
    hash_obj = hashlib.sha256((salt + password_attempt).encode("utf-8"))
    return hash_obj.hexdigest() == stored_hash

admin_h, admin_s = hash_password(DEFAULT_ADMIN_PASS)

DEFAULT_CONFIG_V2 = {
    "allowed_domain": ALLOWED_DOMAIN,
    "universities": [
        "State Open University",
        "National Distance University",
        "Central Higher Education Institute"
    ],
    "university_programs": {
        "State Open University": [
            {
                "name": "Master of Business Administration (MBA)",
                "code": "MBA-GEN",
                "specialization": "General Management & Strategy",
                "department": "School of Business"
            },
            {
                "name": "Master of Science (M.Sc) Data Science",
                "code": "MSC-DS",
                "specialization": "Applied AI & Data Analytics",
                "department": "Department of Computer Science"
            }
        ],
        "National Distance University": [
            {
                "name": "Master of Computer Applications (MCA)",
                "code": "MCA-SE",
                "specialization": "Software Architecture",
                "department": "School of Information Technology"
            }
        ],
        "Central Higher Education Institute": [
            {
                "name": "Bachelor of Commerce (B.Com)",
                "code": "BCOM-AF",
                "specialization": "Accounting & Finance",
                "department": "School of Commerce"
            }
        ]
    },
    "users": {
        PRIMARY_ADMIN: {
            "role": "Lead Administrator",
            "password_hash": admin_h,
            "salt": admin_s
        }
    }
}

def load_access_config_v2():
    if not os.path.exists(CONFIG_PATH_V2):
        with open(CONFIG_PATH_V2, "w") as f:
            json.dump(DEFAULT_CONFIG_V2, f, indent=4)
        return DEFAULT_CONFIG_V2
    try:
        with open(CONFIG_PATH_V2, "r") as f:
            data = json.load(f)
            # Ensure primary admin is always present and updated
            if PRIMARY_ADMIN not in data.get("users", {}):
                data.setdefault("users", {})[PRIMARY_ADMIN] = {
                    "role": "Lead Administrator",
                    "password_hash": admin_h,
                    "salt": admin_s
                }
            return data
    except Exception:
        return DEFAULT_CONFIG_V2

def save_access_config_v2(cfg):
    try:
        with open(CONFIG_PATH_V2, "w") as f:
            json.dump(cfg, f, indent=4)
    except Exception as e:
        print(f"Error saving config v2: {e}")

def authenticate_user_v2(email: str, password: str):
    email = email.strip().lower()
    if not email.endswith(f"@{ALLOWED_DOMAIN}"):
        return False, None, "Invalid Domain: Only official @upgrad.com email addresses are allowed."
    
    cfg = load_access_config_v2()
    users = cfg.get("users", {})
    if email not in users:
        return False, None, "Account not found. Please contact the Lead Administrator to register your account."
    
    user_info = users[email]
    stored_hash = user_info.get("password_hash")
    salt = user_info.get("salt")
    role = user_info.get("role", "Content Author (Maker Team)")
    
    if verify_password(stored_hash, salt, password):
        return True, role, "Login successful."
    else:
        return False, None, "Incorrect password. Please verify and try again."

def register_user_with_password(email: str, password: str, role: str):
    email = email.strip().lower()
    if not email.endswith(f"@{ALLOWED_DOMAIN}"):
        raise ValueError("Email must end with @upgrad.com")
    if len(password.strip()) < 6:
        raise ValueError("Password must be at least 6 characters.")
    
    cfg = load_access_config_v2()
    pw_hash, salt = hash_password(password.strip())
    cfg.setdefault("users", {})[email] = {
        "role": role,
        "password_hash": pw_hash,
        "salt": salt
    }
    save_access_config_v2(cfg)

def add_university_v2(uni_name: str):
    cfg = load_access_config_v2()
    unis = cfg.get("universities", [])
    if uni_name not in unis:
        unis.append(uni_name)
        cfg["universities"] = unis
    cfg.setdefault("university_programs", {}).setdefault(uni_name, [])
    save_access_config_v2(cfg)

def add_program_with_fields(uni_name: str, prog_name: str, prog_code: str = "", specialization: str = "", department: str = ""):
    cfg = load_access_config_v2()
    uni_progs = cfg.setdefault("university_programs", {}).setdefault(uni_name, [])
    
    # Update existing entry if present
    for p in uni_progs:
        if p["name"].lower() == prog_name.strip().lower():
            p["code"] = prog_code.strip() or p.get("code", "N/A")
            p["specialization"] = specialization.strip() or p.get("specialization", "General")
            p["department"] = department.strip() or p.get("department", "Academic Affairs")
            save_access_config_v2(cfg)
            return
            
    uni_progs.append({
        "name": prog_name.strip(),
        "code": prog_code.strip() or "N/A",
        "specialization": specialization.strip() or "General",
        "department": department.strip() or "Academic Affairs"
    })
    save_access_config_v2(cfg)

def get_programs_for_university_v2(uni_name: str) -> list:
    cfg = load_access_config_v2()
    return cfg.get("university_programs", {}).get(uni_name, [])