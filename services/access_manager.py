import json
import os

CONFIG_PATH = os.path.join(os.path.dirname(__file__), "access_config.json")
ALLOWED_DOMAIN = "upgrad.com"

# Primary Lead Administrator
PRIMARY_ADMIN = "sandeep1.shrivastava@upgrad.com"

DEFAULT_CONFIG = {
    "allowed_domain": ALLOWED_DOMAIN,
    "universities": [
        "State Open University",
        "National Distance University",
        "Central Higher Education Institute"
    ],
    "users": {
        PRIMARY_ADMIN: "Lead Administrator",
        "admin@upgrad.com": "Lead Administrator"
    }
}

def load_access_config():
    if not os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "w") as f:
                json.dump(DEFAULT_CONFIG, f, indent=4)
        except Exception:
            pass
        return DEFAULT_CONFIG
    try:
        with open(CONFIG_PATH, "r") as f:
            data = json.load(f)
            if "users" not in data:
                data["users"] = {}
            # Ensure primary admin is always registered
            data["users"][PRIMARY_ADMIN] = "Lead Administrator"
            return data
    except Exception:
        return DEFAULT_CONFIG

def save_access_config(cfg):
    try:
        with open(CONFIG_PATH, "w") as f:
            json.dump(cfg, f, indent=4)
    except Exception as e:
        print(f"Error saving config: {e}")

def verify_user(email: str):
    email = email.strip().lower()
    # 1. Strictly enforce @upgrad.com domain
    if not email.endswith(f"@{ALLOWED_DOMAIN}"):
        return False, "Invalid Domain"

    # 2. Permanent Lead Administrator Access
    if email == PRIMARY_ADMIN.lower() or email == f"admin@{ALLOWED_DOMAIN}":
        return True, "Lead Administrator"

    # 3. Check authorized team members in registry
    cfg = load_access_config()
    users = cfg.get("users", {})

    if email in users:
        return True, users[email]

    return False, None

def add_university(uni_name: str):
    cfg = load_access_config()
    unis = cfg.get("universities", [])
    if uni_name not in unis:
        unis.append(uni_name)
        cfg["universities"] = unis
        save_access_config(cfg)

def assign_user_role(email: str, role: str):
    email = email.strip().lower()
    if not email.endswith(f"@{ALLOWED_DOMAIN}"):
        raise ValueError(f"Email must end with @{ALLOWED_DOMAIN}")
    cfg = load_access_config()
    if "users" not in cfg:
        cfg["users"] = {}
    cfg["users"][email] = role
    save_access_config(cfg)