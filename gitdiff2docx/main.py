"""Main CLI entry point using Click."""

import os
import json
import click
from pygments.styles import get_style_by_name

from gitdiff2docx.config.constants import SCRIPT_DIR
from gitdiff2docx.config.loader import load_json_from_path
from gitdiff2docx.theme.creator import create_theme_interactive
from gitdiff2docx.theme.manager import load_theme, load_theme_overrides
from gitdiff2docx.localization.loader import load_language
from gitdiff2docx.git.operations import get_changed_files, get_first_commit, get_head_commit
from gitdiff2docx.git.filtering import load_ignore_spec, filter_ignored_files
from gitdiff2docx.document.builder import create_document
from gitdiff2docx.diff.processor import process_file_diff
from gitdiff2docx.utils.console import print_green, print_yellow, print_red, ask_yes_no
from gitdiff2docx.cli.interactive import (
    prompt_target_directory, prompt_commits, prompt_output_path,
    handle_first_commit_logic
)


@click.command()
@click.option('--create-theme', is_flag=True, help='Create a new theme file and exit.')
@click.option('--target-dir', '-d', type=click.Path(exists=True), help='Target Git directory.')
@click.option('--commit1', '-c1', help='First commit hash.')
@click.option('--commit2', '-c2', help='Last commit hash.')
@click.option('--output', '-o', type=click.Path(), help='Output .docx file path.')
@click.option('--config-file', '-c', type=click.Path(exists=True), help='Path to config.json file.')
@click.option('--verbose/--no-verbose', default=None, help='Enable verbose output.')
@click.option('--theme', help='Theme name to use.')
@click.option('--language', '-l', help='Language code (e.g., en, de).')
def main(create_theme, target_dir, commit1, commit2, output, config_file, verbose, theme, language):
    """GitDiff2Docx - Convert git diffs to formatted Word documents."""
    
    # Handle theme creation
    if create_theme:
        themes_dir_for_creation = os.path.join(SCRIPT_DIR, "themes")
        if not os.path.isdir(themes_dir_for_creation):
            print_red(f"Error: Themes directory not found: {themes_dir_for_creation}")
            return 1
        create_theme_interactive(themes_dir_for_creation, SCRIPT_DIR)
        return 0

    # Load configuration
    if not config_file:
        config_file = os.path.join(SCRIPT_DIR, "config.json")
    
    if not os.path.exists(config_file):
        print_red(f"Error: Configuration file not found: {config_file}")
        return 1

    with open(config_file, "r", encoding="utf-8") as f:
        config = json.load(f)

    # Override config with CLI arguments
    if verbose is not None:
        config['verbose'] = verbose
    if theme:
        config['theme'] = theme
    if language:
        config['language'] = language

    file_encoding = config.get("file_encoding", "utf-8")
    
    # Load theme
    themes_dir = os.path.join(SCRIPT_DIR, "themes")
    theme_overrides_path = os.path.join(themes_dir, "_overrides.json")
    theme_overrides = load_theme_overrides(theme_overrides_path)
    excluded_theme_files = [os.path.basename(theme_overrides_path)]
    theme_name = str(config.get("theme", "classic")).strip() or "classic"
    theme_obj = load_theme(
        theme_name,
        themes_dir,
        overrides_data=theme_overrides,
        excluded_filenames=excluded_theme_files,
    )

    # Load Pygments style
    pygments_style = config.get("pygments_style", "default")
    try:
        pygments_style_obj = get_style_by_name(pygments_style)
        token_styles = pygments_style_obj.styles
    except Exception:
        token_styles = get_style_by_name("default").styles

    # Load language
    lang_dir = os.path.join(SCRIPT_DIR, "lang")
    if not os.path.exists(lang_dir):
        print_red(f"Error: Language directory not found: {lang_dir}")
        return 1

    lang_choice = config.get("language", "en")
    lang = load_language(lang_choice, lang_dir)

    # Show banner
    print_green(lang["title"])
    print_yellow(f"Using theme: {theme_name}")

    # Get target directory (interactive or from CLI)
    if not target_dir:
        target_dir = prompt_target_directory(lang)
    
    os.chdir(target_dir)

    # Load .gddignore
    gdd_ignore_filename = config.get("gdd_ignore_file_name", ".gddignore")
    ignore_spec = load_ignore_spec(target_dir, gdd_ignore_filename)

    # Get commits (interactive or from CLI)
    if not commit1 or not commit2:
        commit1_prompt, commit2_prompt, is_very_first = prompt_commits(lang)
        if not commit1:
            commit1 = commit1_prompt
        if not commit2:
            commit2 = commit2_prompt
        # Handle first commit logic
        commit1 = handle_first_commit_logic(commit1, is_very_first, config)
    else:
        # When both are provided via CLI, still need to handle first commit logic
        very_first_commit_hash = get_first_commit()
        is_very_first = (commit1 == very_first_commit_hash)
        commit1 = handle_first_commit_logic(commit1, is_very_first, config)

    # Get output path (interactive or from CLI)
    if not output:
        output = prompt_output_path(lang, SCRIPT_DIR)

    # Get changed files
    changed_files = get_changed_files(commit1, commit2)
    changed_files = filter_ignored_files(changed_files, ignore_spec, gdd_ignore_filename)

    if not changed_files:
        print_yellow(lang["no_changes_found"].format(commit1=commit1, commit2=commit2))
        return 0

    # Create document
    doc = create_document(commit1, commit2, lang, config, theme_obj)

    # Check if output file exists
    if os.path.exists(output):
        if not ask_yes_no(lang["output_exists"].format(output_docx=output), lang):
            print(lang["exiting"])
            return 0
        else:
            while True:
                try:
                    with open(output, "a", encoding="utf-8"):
                        break
                except Exception as e:
                    print_red(lang["error_removing_file"].format(output_docx=output, error=str(e)))
                    input(lang["press_enter_to_retry"])

    verbose_mode = config.get("verbose", False)

    # Process each file
    for index, file in enumerate(changed_files):
        if not index == 0 and config.get("insert_page_breaks", True):
            doc.add_page_break()

        doc.add_heading(f"{lang['file']}: {file}", level=config.get("heading_level", 2) + 1)

        if verbose_mode:
            print(lang["processing_file"].format(file=file))

        process_file_diff(doc, file, commit1, commit2, config, theme_obj, token_styles, lang, file_encoding)

        if verbose_mode:
            print_green(lang["processing_done"].format(file=file))

    # Save document
    try:
        doc.save(output)
    except Exception as e:
        print_red(lang["error_saving_file"].format(output_docx=output, error=str(e)))
        return 1

    print_green(lang["saving_report"].format(output_docx=output))

    # Open file if configured
    if config.get("open_after_creation", False):
        try:
            os.startfile(output)
        except Exception as e:
            print_red(lang["error_opening_file"].format(output_docx=output, error=str(e)))

    return 0


if __name__ == '__main__':
    exit(main())
