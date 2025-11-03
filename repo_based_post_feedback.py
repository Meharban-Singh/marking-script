# Posts feedback from feedback.txt to the corresponding GitHub repositories' Feedback PRs

import os
import re
import sys
from github import Github, Auth # Run import PyGithub first using: pip install PyGithub

# Token from: Github profile settings > Developer settings > Personal Access Tokens > Create New
GITHUB_TOKEN = os.getenv("GITHUB_TOKEN", "ghp_XXXXXXXXXXXXXXXXXXXXXXXXXXXX")
ORG_NAME = "ACS-3913-main"


def parse_feedback_file(file_path):
    """
    Parses feedback.txt into {repo_name: feedback_text}.
    
    Expected format:
    <Last, First> [repo_name]
    
    feedback
    
    # TOTAL: 50/50
    
    ===
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
            print("INFO: Skipping template entry.")
            continue

        # Match "<Last, First> [repo_name]" and capture repo_name only
        match = re.match(r"^<[^>]+>\s*\[([^\]]+)\]\s*\n+\s*(.+)$", entry, re.DOTALL)
        if not match:
            print(f"ERROR: Could not parse entry:\n{entry[:80]}...")
            continue

        repo_name = match.group(1).strip()
        feedback_text = match.group(2).strip()

        feedback_data[repo_name] = feedback_text

    print(f"INFO: Parsed {len(feedback_data)} feedback entries from {file_path}")
    return feedback_data

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
            print(f"INFO: Feedback posted to {repo_name} (PR #{feedback_pr.number})")
        else:
            print(f"WARN: No Feedback PR found in {repo_name}")

    except Exception as e:
        print(f"ERROR: posting to {repo_name}: {e}")


# -------------------------------
# MAIN SCRIPT
# -------------------------------

if __name__ == "__main__":
    gh = Github(auth=Auth.Token(GITHUB_TOKEN))

    if len(sys.argv) < 2:
        print('ERROR: File name not provided!') 
        sys.exit(1)
    
    input_file = sys.argv[1]

    feedback_dict = parse_feedback_file(input_file)

    print(f"INFO: Posting feedback now....")

    for repo_name, feedback_text in feedback_dict.items():
        post_feedback(repo_name, feedback_text, gh)