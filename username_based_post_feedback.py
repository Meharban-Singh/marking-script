# Posts feedback from feedback.txt to the corresponding GitHub repositories' Feedback PRs

import os
import re
import sys
from typing import Optional
from github import Github, Auth  # Run import PyGithub first using: pip install PyGithub

# Token from: Github profile settings > Developer settings > Personal Access Tokens > Create New
GITHUB_TOKEN = os.getenv("GITHUB_TOKEN", "ghp_XXXXXXXXXXXXXXXXXXXXXXXXXXXX")
ORG_NAME = "ACS-3913-main"
ASSIGNMENT_SLUG = os.getenv("ASSIGNMENT_SLUG", "assignment-2")
TEMPLATE_REPO_FULL_NAME = "acs-3913-001-f2025-assignment-2-assignment-3-f25"
LOG_FILE = "log.txt"


def log(message: str):
    """log to console and append to log file."""
    print(message)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(message + "\n")
    except Exception:
        pass  # Silent fail if logging fails


def parse_feedback_file(file_path):
    """
    Parses feedback.txt into {username: feedback_text}.
    
    Expected header per entry:
    <Last, First> [github_username]
    
    feedback
    
    TOTAL and separator lines follow.
    """
    with open(file_path, "r", encoding="utf-8") as f:
        content = f.read().strip()

    # Split entries on any line with 3 or more '=' characters
    entries = re.split(r"\n\s*={3,}\s*\n", content)
    feedback_data = {}

    for entry in entries:
        entry = entry.strip()
        if not entry:
            continue

        if entry.lower().startswith("<template"):
            log("INFO: Skipping template entry.")
            continue

        # Match "<Last, First> [github_username]" and capture username then the body
        match = re.match(r"^<[^>]+>\s*\[([^\]]+)\]\s*\n+\s*(.+)$", entry, re.DOTALL)
        if not match:
            log(f"ERROR: Could not parse entry:\n{entry[:80]}...")
            continue

        username = match.group(1).strip()
        feedback_text = match.group(2).strip()

        if username == "unknown_repo":
            log(f"ERROR: No repo found for \n{entry}...")

        feedback_data[username] = feedback_text

    log(f"INFO: Parsed {len(feedback_data)} feedback entries from {file_path}")
    return feedback_data

def is_repo_from_template(repo, template_full_name: str) -> bool:
    try:
        if not getattr(repo, "fork", False):
            return False
        parent = getattr(repo, "parent", None)
        if not parent:
            return False
        # Compare by full_name when possible; fall back to name match
        return parent.full_name.lower() == template_full_name.lower() or parent.name.lower() == template_full_name.split("/", 1)[-1].lower()
    except Exception:
        return False


def resolve_repo_for_username(username: str, gh: Github) -> Optional[str]:
    """
    Resolve a student's repo name from their GitHub username by:
    1) Trying the standard classroom pattern: <ASSIGNMENT_SLUG>-<username>
    2) Searching the org for repos containing the username and ASSIGNMENT_SLUG
       and verifying they are forks of TEMPLATE_REPO_FULL_NAME
    Returns the repo name (without org) or None if not found/validated.
    """
    # 1) Try standard naming
    candidate = f"{ASSIGNMENT_SLUG}-{username}".strip()
    try:
        repo = gh.get_repo(f"{ORG_NAME}/{candidate}")
        if is_repo_from_template(repo, TEMPLATE_REPO_FULL_NAME):
            return repo.name
    except Exception:
        pass

    # 2) Search within org
    try:
        q = f"org:{ORG_NAME} in:name {ASSIGNMENT_SLUG} {username} fork:true"
        results = gh.search_repositories(q)
        matches = []
        for r in results:
            if r.owner.login.lower() != ORG_NAME.lower():
                continue
            if not is_repo_from_template(r, TEMPLATE_REPO_FULL_NAME):
                continue
            matches.append(r)

        if not matches or len(matches) != 1:
            return None
        
        return matches[0].name
    except Exception:
        return None


def post_feedback(repo_name, feedback_text, gh):
    """Finds the Feedback PR in a repo and posts feedback as a comment."""
    try:
        repo = gh.get_repo(f"{ORG_NAME}/{repo_name}")
        pulls = repo.get_pulls(state="open")
        feedback_pr = None

        for pr in pulls:
            if "feedback" in pr.title.lower():
                feedback_pr = pr
                break

        if not feedback_pr:
            pulls = repo.get_pulls(state="closed")
            for pr in pulls:
                if "feedback" in pr.title.lower():
                    feedback_pr = pr
                    break

        if feedback_pr:
            feedback_pr.create_issue_comment(feedback_text)
            log(f"INFO: Feedback posted to {repo_name} (PR #{feedback_pr.number})")
        else:
            log(f"WARN: No Feedback PR found in {repo_name}")

    except Exception as e:
        log(f"ERROR: posting to {repo_name}: {e}")


# -------------------------------
# MAIN SCRIPT
# -------------------------------

if __name__ == "__main__":
    gh = Github(auth=Auth.Token(GITHUB_TOKEN))

    if len(sys.argv) < 2:
        log('ERROR: File name not provided!') 
        sys.exit(1)
    
    input_file = sys.argv[1]

    feedback_dict = parse_feedback_file(input_file)

    log("INFO: Resolving repositories and posting feedback...")

    for username, feedback_text in feedback_dict.items():
        repo_name = resolve_repo_for_username(username, gh)
        if not repo_name:
            log(f"ERROR: Could not resolve repo for username '{username}'. "
                  f"Check ASSIGNMENT_SLUG='{ASSIGNMENT_SLUG}' and TEMPLATE_REPO_FULL_NAME='{TEMPLATE_REPO_FULL_NAME}'.")
            continue

        log(f"INFO: Post for {username} to {repo_name}...")
        post_feedback(repo_name, feedback_text, gh)