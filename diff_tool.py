import os
import json
import subprocess
import io
import mimetypes

import pathspec

from datetime import datetime

from PIL import Image

from difflib import SequenceMatcher

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

from pygments import lex
from pygments.lexers import guess_lexer_for_filename, guess_lexer
from pygments.styles import get_style_by_name

# The directory where this script is located
script_dir = os.path.dirname(os.path.abspath(__file__))

# ------------------------------------------------------------------------------
# Utility functions

def ask_yes_no(prompt: str, lang: dict) -> bool:
    while True:
        answer = input(prompt + " ").strip().lower()
        if answer in [lang["yes"], lang["no"]]:
            return answer == lang["yes"]

def print_green(text: str):
    print(f"\033[92m{text}\033[0m")

def print_yellow(text: str):
    print(f"\033[93m{text}\033[0m")

def print_red(text: str):
    print(f"\033[91m{text}\033[0m")

def rgb_from_hex(hex_color, fallback=(0, 0, 0)):
    try:
        cleaned = hex_color.lstrip("#")
        return RGBColor(
            int(cleaned[0:2], 16),
            int(cleaned[2:4], 16),
            int(cleaned[4:6], 16),
        )
    except Exception:
        return RGBColor(*fallback)

def normalize_hex_color(hex_color, fallback="auto"):
    if isinstance(hex_color, str):
        cleaned = hex_color.strip().lstrip("#")
        if len(cleaned) == 6:
            try:
                int(cleaned, 16)
                return cleaned.upper()
            except ValueError:
                pass
    return fallback

def is_binary_string(bytes_data):
    textchars = bytearray({7, 8, 9, 10, 12, 13, 27}
                          | set(range(0x20, 0x100)) - {0x7f})
    return bool(bytes_data.translate(None, textchars))

def is_image_file(filename):
    mimetype, _ = mimetypes.guess_type(filename)
    return mimetype and mimetype.startswith("image/")

def get_usable_width(document):
    section = document.sections[0]
    page_width = section.page_width
    left_margin = section.left_margin
    right_margin = section.right_margin
    return page_width - left_margin - right_margin  # in EMUs

def set_table_borders(table, border_spec):
    tbl_pr = table._tbl.tblPr
    existing = tbl_pr.find(qn("w:tblBorders"))
    if existing is not None:
        tbl_pr.remove(existing)

    tbl_borders = OxmlElement("w:tblBorders")
    for border, attrs in border_spec.items():
        border_element = OxmlElement(f"w:{border}")
        for key, value in attrs.items():
            border_element.set(qn(f"w:{key}"), str(value))
        tbl_borders.append(border_element)
    tbl_pr.append(tbl_borders)

def set_table_cell_margins(table, top=40, left=80, bottom=40, right=80):
    tbl_pr = table._tbl.tblPr
    existing = tbl_pr.find(qn("w:tblCellMar"))
    if existing is not None:
        tbl_pr.remove(existing)

    tbl_cell_mar = OxmlElement("w:tblCellMar")
    for side, value in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        side_element = OxmlElement(f"w:{side}")
        side_element.set(qn("w:w"), str(value))
        side_element.set(qn("w:type"), "dxa")
        tbl_cell_mar.append(side_element)
    tbl_pr.append(tbl_cell_mar)

def deep_merge_dict(base, override):
    merged = dict(base)
    for key, value in override.items():
        if isinstance(value, dict) and isinstance(merged.get(key), dict):
            merged[key] = deep_merge_dict(merged[key], value)
        else:
            merged[key] = value
    return merged

def get_theme_font(theme, config):
    theme_font = theme.get("font", {})

    font_name = theme_font.get("name") or config.get("diff_font", "Courier New")

    theme_size = theme_font.get("size")
    if theme_size is None:
        resolved_size = int(config.get("diff_font_size", 8))
    else:
        resolved_size = int(theme_size)
    resolved_size = max(resolved_size, int(theme_font.get("min_size", 0)))

    return font_name, resolved_size

def build_border_attrs(side_config, default_config):
    cfg = deep_merge_dict(default_config, side_config if isinstance(side_config, dict) else {})
    visible = bool(cfg.get("visible", False))
    if not visible:
        return {"val": "nil"}

    style = str(cfg.get("style", "single")).strip() or "single"
    weight_pt = float(cfg.get("weight_pt", 0.5))
    sz = int(round(weight_pt * 8))
    sz = max(2, min(96, sz))
    color = normalize_hex_color(cfg.get("color", "auto"), fallback="auto")
    space = max(0, int(cfg.get("space", 0)))

    return {
        "val": style,
        "sz": sz,
        "space": space,
        "color": color,
    }

def build_table_border_spec(theme):
    table_borders_cfg = theme.get("table_borders", {})
    default_side_cfg = table_borders_cfg.get("default", {})
    default_side = default_side_cfg if isinstance(default_side_cfg, dict) else {}

    return {
        "top": build_border_attrs(table_borders_cfg.get("top", {}), default_side),
        "left": build_border_attrs(table_borders_cfg.get("left", {}), default_side),
        "bottom": build_border_attrs(table_borders_cfg.get("bottom", {}), default_side),
        "right": build_border_attrs(table_borders_cfg.get("right", {}), default_side),
        "insideH": build_border_attrs(table_borders_cfg.get("inside_h", {}), default_side),
        "insideV": build_border_attrs(table_borders_cfg.get("inside_v", {}), default_side),
    }

def list_available_themes(themes_dir, excluded_filenames=None):
    if not os.path.isdir(themes_dir):
        return []
    excluded = {name.lower() for name in (excluded_filenames or [])}
    return sorted(
        os.path.splitext(filename)[0]
        for filename in os.listdir(themes_dir)
        if filename.lower().endswith(".json") and filename.lower() not in excluded
    )

def load_theme_overrides(overrides_file_path):
    if not overrides_file_path or not os.path.exists(overrides_file_path):
        return {}

    try:
        with open(overrides_file_path, "r", encoding="utf-8") as f:
            overrides_data = json.load(f)
    except Exception as e:
        print_red(f"Error: Failed to read theme override file '{overrides_file_path}': {e}")
        exit()

    if not isinstance(overrides_data, dict):
        print_red(f"Error: Theme override file '{overrides_file_path}' must contain a JSON object.")
        exit()

    return overrides_data

def load_theme(theme_name, themes_dir, overrides_data=None, excluded_filenames=None):
    available_themes = list_available_themes(themes_dir, excluded_filenames=excluded_filenames)
    if not available_themes:
        print_red(f"Error: No theme files found in '{themes_dir}'.")
        exit()

    theme_path = os.path.join(themes_dir, f"{theme_name}.json")
    if not os.path.exists(theme_path):
        print_red(
            f"Error: Theme '{theme_name}' not found in '{themes_dir}'. Available themes: {', '.join(available_themes)}"
        )
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

def apply_table_theme_style(table, theme):
    table.autofit = False

    margins = theme.get("table_cell_margins", {})
    set_table_cell_margins(
        table,
        top=int(margins.get("top", 40)),
        left=int(margins.get("left", 80)),
        bottom=int(margins.get("bottom", 40)),
        right=int(margins.get("right", 80)),
    )
    set_table_borders(table, build_table_border_spec(theme))

# ------------------------------------------------------------------------------
# Configuration

config_file = os.path.join(script_dir, "config.json")
if not os.path.exists(config_file):
    print_red(f"Error: Configuration file not found: {config_file}")
    exit()

with open(config_file, "r", encoding="utf-8") as f:
    config = json.load(f)

file_encoding = config.get("file_encoding", "utf-8")
themes_dir = os.path.join(script_dir, "themes")
theme_overrides_path = os.path.join(themes_dir, "_overrides.json")
theme_overrides = load_theme_overrides(theme_overrides_path)
excluded_theme_files = [os.path.basename(theme_overrides_path)]
theme_name = str(config.get("theme", "old")).strip() or "old"
theme = load_theme(
    theme_name,
    themes_dir,
    overrides_data=theme_overrides,
    excluded_filenames=excluded_theme_files,
)
# ------------------------------------------------------------------------------
# Pygments style configuration

pygments_style = config.get("pygments_style", "default")
try:
    pygments_style_obj = get_style_by_name(pygments_style)
    token_styles = pygments_style_obj.styles
except Exception:
    token_styles = get_style_by_name("default").styles

# ------------------------------------------------------------------------------
# Localization

lang_dir = os.path.join(script_dir, "lang")
if not os.path.exists(lang_dir):
    print_red(f"Error: Language directory not found: {lang_dir}")
    exit()

lang_choice = config.get("language", "en")
lang_file = os.path.join(lang_dir, f"{lang_choice}.json")
if not os.path.exists(lang_file):
    print_red(f"Language '{lang_choice}' not found, defaulting to English if available.")
    lang_file = os.path.join(lang_dir, "en.json")

with open(lang_file, "r", encoding="utf-8") as f:
    lang = json.load(f)

# ------------------------------------------------------------------------------
# Show banner

print_green(lang["title"])
print_yellow(f"Using theme: {theme_name}")

# ------------------------------------------------------------------------------
# Input prompts

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
    os.chdir(target_dir)
    break

# GDDIgnore
gdd_ignore_filename = config.get("gdd_ignore_file_name", ".gddignore")
gdd_ignore_path = os.path.join(target_dir, gdd_ignore_filename)
ignore_spec = None

if os.path.exists(gdd_ignore_path):
    with open(gdd_ignore_path, "r", encoding="utf-8") as f:
        ignore_spec = pathspec.PathSpec.from_lines("gitwildmatch", f)

# FIRST COMMIT
commit1 = input(lang["enter_commit1"] + " ").strip()
commit1_specified = bool(commit1)
if not commit1_specified:
    commit1 = subprocess.run(
        ["git", "rev-list", "--max-parents=0", "HEAD"],
        capture_output=True, text=True, encoding="utf-8"
    ).stdout.strip()

# Special case: if commit1 is the very first commit, we cannot add a caret
very_first_commit_hash = subprocess.run(
    ["git", "rev-list", "--max-parents=0", "HEAD"],
    capture_output=True, text=True, encoding="utf-8"
).stdout.strip()
is_very_first_commit = (commit1 == very_first_commit_hash)


if not (commit1.endswith("^") or "~" in commit1) and commit1 != "HEAD":
    include_first_commit = config.get("include_first_commit", False)

    if is_very_first_commit and include_first_commit:
        # Special revision number for an empty tree (state before any commit)
        empty_tree = "4b825dc642cb6eb9a060e54bf8d69288fbee4904"
        commit1 = empty_tree
    elif include_first_commit:
        commit1 = f"{commit1}^" if include_first_commit else commit1

if not commit1_specified:
    print(lang["using_first_commit"].format(commit1=commit1))

# LAST COMMIT
commit2 = input(lang["enter_commit2"] + " ").strip()
if not commit2:
    commit2 = subprocess.run(
        ["git", "rev-parse", "HEAD"],
        capture_output=True, text=True, encoding="utf-8"
    ).stdout.strip()
    print(lang["using_last_commit"].format(commit2=commit2))

output_docx = input(lang["enter_output_docx"] + " ").strip()
if not output_docx:
    output_docx = os.path.join(script_dir, "output.docx")
    print(lang["using_default_output"].format(output_docx=output_docx))

# ------------------------------------------------------------------------------
# Diff generation
changed_files = subprocess.run(
    ["git", "diff", "--name-only", commit1, commit2],
    capture_output=True, text=True, encoding="utf-8"
).stdout.splitlines()

changed_files = [f for f in changed_files if f != gdd_ignore_filename]

if ignore_spec:
    changed_files = [f for f in changed_files if not ignore_spec.match_file(f)]

if not changed_files:
    print_yellow(lang["no_changes_found"].format(commit1=commit1, commit2=commit2))
    exit()

# ------------------------------------------------------------------------------
# Word document generation

doc = Document()
doc.add_heading(f"{lang['git_changes_report']} ({commit1} → {commit2})", level=1)
doc.add_paragraph(lang["report_generated_on"].format(
    date=datetime.now().strftime("%Y-%m-%d %H:%M:%S")
))

# Add legend table
def add_legend_table(document):
    legend_table = document.add_table(rows=0, cols=2)
    legend_table.style = "Table Grid"
    apply_table_theme_style(legend_table, theme)

    theme_colors = theme.get("colors", {})
    theme_symbols = theme.get("symbols", {})
    theme_font = theme.get("font", {})

    legend_add_color = theme_colors.get("add_fill", "D0FFD0")
    legend_remove_color = theme_colors.get("remove_fill", "FFD0D0")
    legend_neutral_color = theme_colors.get("neutral_fill", "F5F5F5")

    legend_add_symbol = theme_symbols.get("add", "+")
    legend_remove_symbol = theme_symbols.get("remove", "-")
    legend_neutral_symbol = theme_symbols.get("neutral", "=")

    diff_font_name, diff_font_size = get_theme_font(theme, config)
    bold_symbols = bool(theme_font.get("bold_symbols", False))

    legend_data = [
        (lang["legend_add"], legend_add_color, legend_add_symbol),
        (lang["legend_remove"], legend_remove_color, legend_remove_symbol),
    ]

    # Only include neutral/unchanged lines in the legend if they're being shown in the diff
    if config.get("include_unchanged_lines", True):
        legend_data.append(
            (lang["legend_neutral"], legend_neutral_color, legend_neutral_symbol)
        )

    for label, color, symbol in legend_data:
        column = legend_table.add_row().cells
        column[0].text = label

        shading = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls("w"), color))
        column[1]._element.get_or_add_tcPr().append(shading)
        column[1].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        p = column[1].paragraphs[0]
        run = p.add_run(symbol)
        font = run.font
        font.name = diff_font_name
        font.size = Pt(diff_font_size)
        if bold_symbols:
            run.bold = True
        if symbol == legend_add_symbol:
            run.font.color.rgb = rgb_from_hex(theme_colors.get("add_symbol", "000000"))
        elif symbol == legend_remove_symbol:
            run.font.color.rgb = rgb_from_hex(theme_colors.get("remove_symbol", "000000"))
        else:
            run.font.color.rgb = rgb_from_hex(theme_colors.get("neutral_symbol", "000000"))

doc.add_heading(lang["legend"], level=config.get("heading_level", 2))
add_legend_table(doc)

doc.add_page_break()
doc.add_heading(lang["diffs"], level=config.get("heading_level", 2))

# Extract line numbers from git diff
def extract_line_numbers(diff_lines):
    line_numbers = []
    current_line = 0
    for line in diff_lines:
        if line.startswith("@@"):
            parts = line.split(" ")
            new_file_info = parts[2]
            start_line = int(new_file_info.split(",")[0][1:])
            current_line = start_line
        elif not line.startswith("-"):
            line_numbers.append(current_line)
            current_line += 1
    return line_numbers

# Add a formatted and syntax-highlighted code diff table
def add_diff_table(document, diff_lines, line_numbers, lexer):
    table = document.add_table(rows=0, cols=2)
    table.style = "Table Grid"
    apply_table_theme_style(table, theme)

    theme_colors = theme.get("colors", {})
    theme_symbols = theme.get("symbols", {})
    theme_font = theme.get("font", {})

    symbol_col_width = Cm(float(theme.get("symbol_column_width_cm", 0.57)))
    table.columns[0].width = symbol_col_width
    table.columns[1].width = get_usable_width(document) - symbol_col_width

    diff_font_name, diff_font_size = get_theme_font(theme, config)
    diff_font_size_pt = Pt(diff_font_size)

    add_symbol = theme_symbols.get("add", "+")
    remove_symbol = theme_symbols.get("remove", "-")
    neutral_symbol = theme_symbols.get("neutral", "=")

    bold_symbols = bool(theme_font.get("bold_symbols", False))
    center_symbols = bool(theme_font.get("center_symbols", False))
    line_spacing = float(theme_font.get("line_spacing", 1.0))
    use_syntax_highlighting = bool(theme.get("use_syntax_highlighting", True))

    # Skip unchanged lines if configured to do so
    include_unchanged = config.get("include_unchanged_lines", True)
    if not include_unchanged:
        filtered_diff_lines = []
        filtered_line_numbers = []
        for line, line_num in zip(diff_lines, line_numbers):
            if not line.startswith(" "):
                filtered_diff_lines.append(line)
                filtered_line_numbers.append(line_num)
        diff_lines = filtered_diff_lines
        line_numbers = filtered_line_numbers

    for line in diff_lines:
        row_cells = table.add_row().cells
        symbol_cell = row_cells[0]
        code_cell = row_cells[1]

        if line.startswith("+"):
            fill = theme_colors.get("add_fill", "D0FFD0")
            symbol_color = theme_colors.get("add_symbol", "000000")
            symbol = add_symbol
        elif line.startswith("-"):
            fill = theme_colors.get("remove_fill", "FFD0D0")
            symbol_color = theme_colors.get("remove_symbol", "000000")
            symbol = remove_symbol
        else:
            fill = theme_colors.get("neutral_fill", "F5F5F5")
            symbol_color = theme_colors.get("neutral_symbol", "000000")
            symbol = neutral_symbol

        for cell in (symbol_cell, code_cell):
            shading = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls("w"), fill))
            cell._element.get_or_add_tcPr().append(shading)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        symbol_paragraph = symbol_cell.paragraphs[0]
        symbol_paragraph.clear()
        symbol_paragraph.paragraph_format.space_before = Pt(0)
        symbol_paragraph.paragraph_format.space_after = Pt(0)
        if center_symbols:
            symbol_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_sym = symbol_paragraph.add_run(symbol)
        run_sym.font.name = diff_font_name
        run_sym.font.size = diff_font_size_pt
        run_sym.bold = bold_symbols
        run_sym.font.color.rgb = rgb_from_hex(symbol_color)

        paragraph = code_cell.paragraphs[0]
        paragraph.clear()
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)
        if line_spacing > 0:
            paragraph.paragraph_format.line_spacing = line_spacing

        code_content = line[1:]

        for ttype, value in lex(code_content, lexer):
            value = value.rstrip('\n')
            if not value:
                value = "\u00A0"

            run = paragraph.add_run(value)
            run.font.name = diff_font_name
            run.font.size = diff_font_size_pt

            style_str = token_styles.get(ttype) if use_syntax_highlighting else None
            if style_str:
                for part in style_str.split():
                    if part == "bold":
                        run.bold = True
                    elif part == "italic":
                        run.italic = True
                    elif part.startswith("#") and len(part) == 7:
                        hexcode = part.lstrip("#")
                        try:
                            r = int(hexcode[0:2], 16)
                            g = int(hexcode[2:4], 16)
                            b = int(hexcode[4:6], 16)
                            run.font.color.rgb = RGBColor(r, g, b)
                        except ValueError:
                            pass
            else:
                run.font.color.rgb = rgb_from_hex(theme_colors.get("default_code", "000000"))

# Add images to the document
def add_image(document, file_bytes, image_name):
    image_stream = io.BytesIO(file_bytes)
    try:
        image = Image.open(image_stream)
        width, _ = image.size
        image_stream.seek(0)  # rewind for docx
        max_width_inches = 6  # ~75% of page width (8 inches)
        width_inches = min(max_width_inches, width / image.info.get('dpi', (96, 96))[0])
        document.add_picture(image_stream, width=Inches(width_inches))
        document.paragraphs[-1].alignment = 1  # center
    except Exception as e:
        document.add_paragraph(lang["error_inserting_image"].format(image_name=image_name))

# ------------------------------------------------------------------------------
# Main loop

if os.path.exists(output_docx):
    if not ask_yes_no(lang["output_exists"].format(output_docx=output_docx), lang):
        print(lang["exiting"])
        exit()
    else:
        while True:
            try:
                with open(output_docx, "a", encoding="utf-8"):
                    break
            except Exception as e:
                print_red(lang["error_removing_file"].format(output_docx=output_docx, error=str(e)))
                input(lang["press_enter_to_retry"])

verbose = config.get("verbose", False)

for index, file in enumerate(changed_files):
    if not index == 0 and config.get("insert_page_breaks", True):
        doc.add_page_break()

    doc.add_heading(f"{lang['file']}: {file}", level=config.get("heading_level", 2) + 1)

    if verbose:
        print(lang["processing_file"].format(file=file))

    # Try to get binary contents
    try:
        old_bytes = subprocess.run(
            ["git", "show", f"{commit1}:{file}"],
            capture_output=True
        ).stdout
    except:
        old_bytes = b""

    try:
        new_bytes = subprocess.run(
            ["git", "show", f"{commit2}:{file}"],
            capture_output=True
        ).stdout
    except:
        new_bytes = b""

    # If binary but not image → skip
    if is_binary_string(new_bytes) or is_binary_string(old_bytes):
        if is_image_file(file):
            # Insert only if the image changed
            if old_bytes != new_bytes:
                paragraph = doc.add_paragraph()
                run = paragraph.add_run(lang['image_changed'])
                run.italic = True

                if config.get("include_images", True):
                    add_image(doc, new_bytes, file)
        else:
            paragraph = doc.add_paragraph()
            run = paragraph.add_run(lang['binary_file_skipped'])
            run.italic = True
        continue

    # Fallback for text diff
    old_content = old_bytes.decode(file_encoding, errors="ignore").splitlines()
    new_content = new_bytes.decode(file_encoding, errors="ignore").splitlines()

    # Choose lexer based on filename and content
    sample = "\n".join(new_content or old_content)
    try:
        lexer = guess_lexer_for_filename(file, sample)
    except:
        lexer = guess_lexer(sample or "")

    matcher = SequenceMatcher(None, old_content, new_content)
    diff_lines = []
    line_nums = []

    old_idx = new_idx = 0

    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            # Only add unchanged lines if configured to do so
            if config.get("include_unchanged_lines", True):
                for line in new_content[j1:j2]:
                    diff_lines.append(f" {line}")
                    line_nums.append(new_idx + 1)
                    new_idx += 1
                    old_idx += 1
            else:
                # Still need to update indices even if we don't add the lines
                new_idx += (j2 - j1)
                old_idx += (i2 - i1)
        elif tag == "replace":
            for line in old_content[i1:i2]:
                diff_lines.append(f"-{line}")
                line_nums.append(old_idx + 1)
                old_idx += 1
            for line in new_content[j1:j2]:
                diff_lines.append(f"+{line}")
                line_nums.append(new_idx + 1)
                new_idx += 1
        elif tag == "delete":
            for line in old_content[i1:i2]:
                diff_lines.append(f"-{line}")
                line_nums.append(old_idx + 1)
                old_idx += 1
        elif tag == "insert":
            for line in new_content[j1:j2]:
                diff_lines.append(f"+{line}")
                line_nums.append(new_idx + 1)
                new_idx += 1

    # Add Table if there are significant changes
    if diff_lines:
        add_diff_table(doc, diff_lines, line_nums, lexer)
    else:
        doc.add_paragraph(lang["no_significant_changes"], style="Italic")

    if verbose:
        print_green(lang["processing_done"].format(file=file))

try:
    doc.save(output_docx)
except Exception as e:
    print_red(lang["error_saving_file"].format(output_docx=output_docx, error=str(e)))
    exit()

print_green(lang["saving_report"].format(output_docx=output_docx))

if config.get("open_after_creation", False):
    try:
        os.startfile(output_docx)
    except Exception as e:
        print_red(lang["error_opening_file"].format(output_docx=output_docx, error=str(e)))
        exit()
