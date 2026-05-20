#!/usr/bin/env python3
"""
GitHub Push Script
Auto-commits and pushes HTML files to GitHub
Run this after task_tracker_FINAL.py generates new dashboards
"""

import subprocess
from pathlib import Path
from datetime import datetime

def run_command(cmd, description):
    """Run a shell command and print output"""
    print(f"\n{description}...")
    try:
        result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
        if result.returncode == 0:
            print(f"✓ {description}")
            return True
        else:
            print(f"ERROR: {result.stderr}")
            return False
    except Exception as e:
        print(f"ERROR: {e}")
        return False

def main():
    cwd = Path.cwd()
    print(f"\n=== GitHub Push ===")
    print(f"Working directory: {cwd}")
    
    # Check if we're in a git repo
    if not (cwd / ".git").exists():
        print("ERROR: Not a git repository. Make sure you're in ~/Desktop/task-tracker/")
        return
    
    # Stage HTML files
    print(f"\nStaging files...")
    run_command("git add dashboard.html approval_ui.html", "Staging HTML files")
    
    # Commit
    timestamp = datetime.now().strftime("%A, %B %d, %Y at %I:%M %p")
    commit_msg = f"Update dashboards - {timestamp}"
    
    if run_command(f'git commit -m "{commit_msg}"', "Committing changes"):
        # Push
        if run_command("git push origin main", "Pushing to GitHub"):
            print(f"\n✓ COMPLETE! Dashboard updated on GitHub")
            print(f"Dashboard: https://aashnagoel.github.io/task-tracker/dashboard.html")
        else:
            print("ERROR: Failed to push to GitHub")
    else:
        print("No changes to commit")

if __name__ == "__main__":
    main()
