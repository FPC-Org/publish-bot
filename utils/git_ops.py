from __future__ import annotations

import subprocess
from pathlib import Path


def push_report_changes(repo_root: Path, message: str) -> None:
    """Commit and push the current members-repo changes to the remote branch.

    Parameters:
    - `repo_root`: Root directory of the git repository that should receive the
      commit. In this project that should always be the `members` repo, never
      the publish-bot repo itself.
    - `message`: Commit message to use for the deployment commit.

    Commands executed:
    - `git add .`
    - `git commit -m <message>`
    - `git push origin HEAD`

    Any failure raises `RuntimeError` with the captured stderr/stdout payload so
    the caller can surface a useful error message.
    """
    commands = [
        ["git", "add", "."],
        ["git", "commit", "-m", message],
        ["git", "push", "origin", "HEAD"],
    ]
    for command in commands:
        result = subprocess.run(command, cwd=repo_root, capture_output=True, text=True)
        if result.returncode != 0:
            stderr = result.stderr.strip() or result.stdout.strip()
            raise RuntimeError(f"Git command failed: {' '.join(command)}\n{stderr}")
