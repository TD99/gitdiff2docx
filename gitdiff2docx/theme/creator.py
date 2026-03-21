"""Theme creation and validation utilities."""

import os
import json
from gitdiff2docx.config.schema import build_defaults_from_schema_node
from gitdiff2docx.config.loader import load_json_from_path
from gitdiff2docx.utils.console import print_red, print_green


def build_default_theme_template(script_dir):
    """Build a default theme template from schema."""
    schema_path = os.path.join(script_dir, "schemas", "theme.schema.json")
    schema_root, abs_schema_path = load_json_from_path(schema_path)
    return build_defaults_from_schema_node(
        schema_root,
        abs_schema_path,
        include_optional_defaults=True,
    )


def is_valid_theme_filename(theme_name):
    """Validate theme filename."""
    invalid_chars = set('<>:"/\\|?*')
    if not theme_name:
        return False
    if theme_name.startswith("."):
        return False
    if theme_name.lower() == "_overrides":
        return False
    return not any(char in invalid_chars for char in theme_name)


def create_theme_interactive(themes_dir, script_dir):
    """Create a new theme interactively."""
    while True:
        theme_name = input("Enter new theme name (without .json): ").strip()
        if not is_valid_theme_filename(theme_name):
            print_red("Invalid theme name. Avoid empty names, leading dots, reserved '_overrides', and filename-invalid characters.")
            continue

        theme_path = os.path.join(themes_dir, f"{theme_name}.json")
        if os.path.exists(theme_path):
            print_red(f"Theme already exists: {theme_path}")
            continue

        theme_data = build_default_theme_template(script_dir)
        with open(theme_path, "w", encoding="utf-8") as f:
            json.dump(theme_data, f, indent=4)
            f.write("\n")

        print_green(f"Created theme: {theme_path}")
        return
