import ast
import re
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
FORBIDDEN_FILENAMES = {"scheduler_data.json", ".env"}
FORBIDDEN_SUFFIXES = {".csv", ".xls", ".xlsx"}
FORBIDDEN_DIRECTORIES = {"exports", "private_archive"}
GENERIC_LOCATION_IDS = {f"location-{letter}" for letter in "abcdefghi"}
GENERIC_LOCATION_NAMES = {f"Location {letter}" for letter in "ABCDEFGHI"}
SENSITIVE_TEXT_PATTERNS = {
    "absolute Windows user path": re.compile(r"C:\\Users\\", re.IGNORECASE),
    "OneDrive path": re.compile(r"OneDrive", re.IGNORECASE),
    "OpenAI-style secret": re.compile(r"\bsk-[A-Za-z0-9_-]{16,}\b"),
}


def get_tracked_files():
    result = subprocess.run(
        ["git", "ls-files", "-z", "--cached", "--others", "--exclude-standard"],
        cwd=ROOT,
        check=True,
        capture_output=True,
    )
    return [Path(item.decode("utf-8")) for item in result.stdout.split(b"\0") if item]


def get_literal_assignment(name):
    source = (ROOT / "app.py").read_text(encoding="utf-8-sig")
    tree = ast.parse(source)
    for node in tree.body:
        if not isinstance(node, ast.Assign):
            continue
        if any(isinstance(target, ast.Name) and target.id == name for target in node.targets):
            return ast.literal_eval(node.value)
    raise RuntimeError(f"{name} not found")


def main():
    failures = []
    tracked_files = get_tracked_files()

    for relative_path in tracked_files:
        lower_parts = {part.lower() for part in relative_path.parts}
        if relative_path.name.lower() in FORBIDDEN_FILENAMES:
            failures.append(f"forbidden tracked file: {relative_path}")
        if relative_path.suffix.lower() in FORBIDDEN_SUFFIXES:
            failures.append(f"private data format is tracked: {relative_path}")
        if lower_parts & FORBIDDEN_DIRECTORIES:
            failures.append(f"private directory is tracked: {relative_path}")

        path = ROOT / relative_path
        if path.suffix.lower() not in {".py", ".html", ".md", ".yml", ".yaml", ".txt", ".json"}:
            continue
        if relative_path == Path("scripts/privacy_check.py"):
            continue
        text = path.read_text(encoding="utf-8-sig", errors="replace")
        for label, pattern in SENSITIVE_TEXT_PATTERNS.items():
            if pattern.search(text):
                failures.append(f"{label} found in {relative_path}")

    configs = get_literal_assignment("LOCATION_CONFIGS")
    location_ids = {item.get("id") for item in configs}
    location_names = {item.get("name") for item in configs}
    if location_ids != GENERIC_LOCATION_IDS:
        failures.append("LOCATION_CONFIGS must use only location-a through location-i")
    if location_names != GENERIC_LOCATION_NAMES:
        failures.append("LOCATION_CONFIGS must use only Location A through Location I")

    preferred_worker_names = set(get_literal_assignment("PREFERRED_WORKER_FILL_COLORS"))
    if any(not re.fullmatch(r"worker [a-z]", name) for name in preferred_worker_names):
        failures.append("preferred export colors must use only generic Worker labels")

    gitignore = (ROOT / ".gitignore").read_text(encoding="utf-8").splitlines()
    for required_entry in (".env", "exports/", "scheduler_data.json"):
        if required_entry not in gitignore:
            failures.append(f"missing .gitignore privacy rule: {required_entry}")

    if failures:
        print("Public privacy check failed:")
        for failure in failures:
            print(f"- {failure}")
        return 1

    print(f"Public privacy check passed ({len(tracked_files)} tracked files inspected).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
