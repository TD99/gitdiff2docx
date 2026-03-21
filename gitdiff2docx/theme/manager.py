"""Theme manager for loading and validating themes."""

import os
import json
from gitdiff2docx.utils.console import print_red
from gitdiff2docx.utils.dict import deep_merge_dict
from gitdiff2docx.theme.creator import is_valid_theme_filename


def list_available_themes(themes_dir, excluded_filenames=None):
    """List all available theme files in the themes directory."""
    if not os.path.isdir(themes_dir):
        return []
    excluded = {name.lower() for name in (excluded_filenames or [])}
    return sorted(
        os.path.splitext(filename)[0]
        for filename in os.listdir(themes_dir)
        if filename.lower().endswith(".json") and filename.lower() not in excluded
    )


def load_theme_overrides(overrides_file_path):
    """Load theme overrides from _overrides.json file."""
    if not overrides_file_path or not os.path.exists(overrides_file_path):
        return {}

    try:
        with open(overrides_file_path, "r", encoding="utf-8") as f:
            overrides_data = json.load(f)
    except Exception as e:
        print_red(f"Error: Failed to read theme override file '{overrides_file_path}': {e}")
        exit(1)

    if not isinstance(overrides_data, dict):
        print_red(f"Error: Theme override file '{overrides_file_path}' must contain a JSON object.")
        exit(1)

    return overrides_data


def load_theme(theme_name, themes_dir, overrides_data=None, excluded_filenames=None):
    """Load a theme by name with optional overrides."""
    available_themes = list_available_themes(themes_dir, excluded_filenames=excluded_filenames)
    if not available_themes:
        print_red(f"Error: No theme files found in '{themes_dir}'.")
        exit()

    if not is_valid_theme_filename(theme_name):
        print_red(
            f"Error: Invalid theme name '{theme_name}'. Available themes: {', '.join(available_themes)}"
        )
        exit()

    if theme_name not in available_themes:
        print_red(
            f"Error: Theme '{theme_name}' not found in '{themes_dir}'. Available themes: {', '.join(available_themes)}"
        )
        exit()

    themes_dir_abs = os.path.abspath(themes_dir)
    theme_path = os.path.abspath(os.path.normpath(os.path.join(themes_dir_abs, f"{theme_name}.json")))
    if os.path.commonpath([themes_dir_abs, theme_path]) != themes_dir_abs:
        print_red(f"Error: Invalid theme path for theme '{theme_name}'.")
        exit()

    try:
        with open(theme_path, "r", encoding="utf-8") as f:
            theme_data = json.load(f)
    except Exception as e:
        print_red(f"Error: Failed to read theme file '{theme_path}': {e}")
        exit()

    if not isinstance(theme_data, dict):
        print_red(f"Error: Theme file '{theme_path}' must contain a JSON object.")
        exit()

    merged_theme = dict(theme_data)
    if overrides_data:
        merged_theme = deep_merge_dict(merged_theme, overrides_data)

    merged_theme["name"] = theme_name
    return merged_theme
