"""Git operations and utilities."""

import subprocess


def get_changed_files(commit1, commit2):
    """Get list of changed files between two commits."""
    result = subprocess.run(
        ["git", "diff", "--name-only", commit1, commit2],
        capture_output=True, text=True, encoding="utf-8"
    )
    return result.stdout.splitlines()


def get_first_commit():
    """Get the very first commit hash in the repository."""
    result = subprocess.run(
        ["git", "rev-list", "--max-parents=0", "HEAD"],
        capture_output=True, text=True, encoding="utf-8"
    )
    return result.stdout.strip()


def get_head_commit():
    """Get the HEAD commit hash."""
    result = subprocess.run(
        ["git", "rev-parse", "HEAD"],
        capture_output=True, text=True, encoding="utf-8"
    )
    return result.stdout.strip()


def get_file_at_commit(commit, file_path):
    """Get file content at a specific commit."""
    try:
        result = subprocess.run(
            ["git", "show", f"{commit}:{file_path}"],
            capture_output=True
        )
        return result.stdout
    except:
        return b""
