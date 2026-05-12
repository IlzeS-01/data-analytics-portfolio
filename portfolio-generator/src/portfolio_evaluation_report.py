from pathlib import Path
import re

# CONFIG
PROJECT_ROOT = Path.home() / "OneDrive - Demo Company/path/to/file"

DATA_KEYWORDS = ["pandas", "numpy", "read_csv", "DataFrame"]
SQL_KEYWORDS = ["SELECT", "JOIN", "FROM", "WHERE"]
OUTPUT_KEYWORDS = ["to_csv", "to_excel", "print", "save", "write"]
AUTOMATION_KEYWORDS = ["os.", "shutil", "schedule", "subprocess"]

EXCLUDE_HINTS = ["copy", "move", "rename", "tmp"]

# ANALYSIS FUNCTIONS
def analyze_file(file_path: Path) -> dict:
    """Analyze a single script"""

    try:
        content = file_path.read_text(encoding="utf-8")
    except:
        return {"score": 0, "reason": "Unreadable file"}

    score = 0
    reasons = []

    # --- Length (complexity proxy)
    lines = len(content.splitlines())
    if lines > 50:
        score += 1
        reasons.append("Moderate/path/to/file logic")

    if lines > 150:
        score += 1
        reasons.append("High complexity")

    # --- Data usage
    if any(k in content for k in DATA_KEYWORDS):
        score += 1
        reasons.append("Uses data processing (pandas/path/to/file")

    # --- SQL usage
    if any(k in content.upper() for k in SQL_KEYWORDS):
        score += 1
        reasons.append("Uses SQL/path/to/file extraction")

    # --- Output generation
    if any(k in content for k in OUTPUT_KEYWORDS):
        score += 1
        reasons.append("Produces output/path/to/file")

    # --- Automation signals
    if any(k in content for k in AUTOMATION_KEYWORDS):
        score += 1
        reasons.append("Includes automation logic")

    # --- Weak signals
    if any(h in file_path.name.lower() for h in EXCLUDE_HINTS):
        score -= 1
        reasons.append("Likely utility/path/to/file operation script")

    return {"score": score, "reasons": reasons}

def classify(score: int) -> str:
    """Classify script based on score"""
    if score >= 4:
        return "STRONG (Feature as standalone project)"
    elif score >= 2:
        return "MEDIUM (Combine into larger project)"
    else:
        return "WEAK (Exclude from portfolio)"

# PROJECT SCAN
def evaluate_projects():
    print("\nPortfolio Evaluation Report\n")

    scripts = list(PROJECT_ROOT.glob("*.py"))

    if not scripts:
        print("No scripts found.")
        return

    total_score = 0

    for file in scripts:
        print(f"\nScript: {file.name}")
        print("-" * 40)

        result = analyze_file(file)
        classification = classify(result["score"])

        print(f"Score: {result['score']} → {classification}")

        for r in result["reasons"]:
            print(f"  - {r}")

        total_score += result["score"]

    avg_score = total_score / len(scripts)

    print("\nOverall Summary")
    print(f"Average Score: {avg_score:.2f}")
    print("Top candidates = scripts with score ≥ 4")

    print("\nEvaluation Complete")

# RUN
if __name__ == "__main__":
    evaluate_projects()