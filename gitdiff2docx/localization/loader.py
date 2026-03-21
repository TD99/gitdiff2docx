"""Language file loading utilities."""

import os
import json
from gitdiff2docx.utils.console import print_red


def load_language(lang_choice, lang_dir):
    """Load language file with fallback to English."""
    lang_file = os.path.join(lang_dir, f"{lang_choice}.json")
    if not os.path.exists(lang_file):
        print_red(f"Language '{lang_choice}' not found, defaulting to English if available.")
        lang_file = os.path.join(lang_dir, "en.json")

    with open(lang_file, "r", encoding="utf-8") as f:
        return json.load(f)
