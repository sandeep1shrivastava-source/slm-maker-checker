import re

def run_pre_checks(text: str) -> dict:
    results = {"passed": True, "flags": []}

    prohibited_patterns = [
        r"as an ai",
        r"certainly,? here is",
        r"i hope this helps",
        r"\[insert .*? here\]"
    ]
    for pattern in prohibited_patterns:
        if re.search(pattern, text, re.IGNORECASE):
            results["passed"] = False
            results["flags"].append(f"AI boilerplate/placeholder detected matching: '{pattern}'")

    required_sections = ["Introduction", "Objectives", "Summary", "References"]
    for sec in required_sections:
        if not re.search(rf"\b{sec}\b", text, re.IGNORECASE):
            results["passed"] = False
            results["flags"].append(f"Missing mandatory section keyword: '{sec}'")

    return results