"""Interactive input prompts."""

import os
from gitdiff2docx.utils.console import ask_yes_no, print_red


def prompt_target_directory(lang):
    """Prompt for target git directory."""
    while True:
        target_dir = input(lang["enter_target_dir"] + " ").strip()
        if not os.path.isdir(target_dir):
            print_red(lang["invalid_target_dir"])
            continue
        if ".git" not in os.listdir(target_dir):
            print_red(lang["no_git_repo_found"].format(target_dir=target_dir))
            if ask_yes_no(lang["still_continue"], lang):
                break
            continue
        break
    return target_dir


def prompt_commits(lang):
    """Prompt for commit hashes."""
    from gitdiff2docx.git.operations import get_first_commit, get_head_commit
    
    # FIRST COMMIT
    commit1 = input(lang["enter_commit1"] + " ").strip()
    commit1_specified = bool(commit1)
    if not commit1_specified:
        commit1 = get_first_commit()

    # Special case: if commit1 is the very first commit
    very_first_commit_hash = get_first_commit()
    is_very_first_commit = (commit1 == very_first_commit_hash)

    if not (commit1.endswith("^") or "~" in commit1) and commit1 != "HEAD":
        # Use config later for include_first_commit decision
        pass

    if not commit1_specified:
        print(lang["using_first_commit"].format(commit1=commit1))

    # LAST COMMIT
    commit2 = input(lang["enter_commit2"] + " ").strip()
    if not commit2:
        commit2 = get_head_commit()
        print(lang["using_last_commit"].format(commit2=commit2))

    return commit1, commit2, is_very_first_commit


def prompt_output_path(lang, script_dir):
    """Prompt for output DOCX file path."""
    output_docx = input(lang["enter_output_docx"] + " ").strip()
    if not output_docx:
        output_docx = os.path.join(script_dir, "output.docx")
        print(lang["using_default_output"].format(output_docx=output_docx))
    return output_docx


def handle_first_commit_logic(commit1, is_very_first_commit, config):
    """Handle logic for including the first commit."""
    if not (commit1.endswith("^") or "~" in commit1) and commit1 != "HEAD":
        include_first_commit = config.get("include_first_commit", False)

        if is_very_first_commit and include_first_commit:
            # Special revision number for an empty tree (state before any commit)
            empty_tree = "4b825dc642cb6eb9a060e54bf8d69288fbee4904"
            commit1 = empty_tree
        elif include_first_commit:
            commit1 = f"{commit1}^"
    
    return commit1
