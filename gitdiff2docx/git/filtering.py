"""Git file filtering utilities."""

import os
import pathspec


def load_ignore_spec(target_dir, gdd_ignore_filename):
    """Load .gddignore file if it exists."""
    gdd_ignore_path = os.path.join(target_dir, gdd_ignore_filename)
    if os.path.exists(gdd_ignore_path):
        with open(gdd_ignore_path, "r", encoding="utf-8") as f:
            return pathspec.PathSpec.from_lines("gitwildmatch", f)
    return None


def filter_ignored_files(changed_files, ignore_spec, gdd_ignore_filename):
    """Filter files based on .gddignore patterns."""
    # Remove the ignore file itself
    changed_files = [f for f in changed_files if f != gdd_ignore_filename]
    
    # Apply ignore patterns
    if ignore_spec:
        changed_files = [f for f in changed_files if not ignore_spec.match_file(f)]
    
    return changed_files
