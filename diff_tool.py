import os
import json
import subprocess

from datetime import datetime

from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

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

# ------------------------------------------------------------------------------
# Configuration

config_file = os.path.join(script_dir, "config.json")
if not os.path.exists(config_file):
    print_red(f"Error: Configuration file not found: {config_file}")
    exit()

with open(config_file, "r", encoding="utf-8") as f:
    config = json.load(f)

file_encoding = config.get("file_encoding", "utf-8")

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
    print_red(f"Language '{lang_choice}' not found, defaulting to English.")
    lang_file = os.path.join(lang_dir, "en.json")

with open(lang_file, "r", encoding="utf-8") as f:
    lang = json.load(f)

# ------------------------------------------------------------------------------
# Show banner

print_green(lang["title"])

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

commit1 = input(lang["enter_commit1"] + " ").strip()
if not commit1:
    commit1 = subprocess.run(
        ["git", "rev-list", "--max-parents=0", "HEAD"],
        capture_output=True, text=True, encoding="utf-8"
    ).stdout.strip()
    print(lang["using_first_commit"].format(commit1=commit1))

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

    legend_data = [
        (lang["legend_add"], config.get("add_color", "D0FFD0"), "+"),
        (lang["legend_remove"], config.get("remove_color", "FFD0D0"), "-"),
        (lang["legend_neutral"], config.get("neutral_color", "F5F5F5"), " "),
    ]

    for label, color, symbol in legend_data:
        column = legend_table.add_row().cells
        column[0].text = label
        shading = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls("w"), color))
        column[1]._element.get_or_add_tcPr().append(shading)
        column[1].text = symbol

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
    table = document.add_table(rows=0, cols=1)
    table.style = "Table Grid"

    for line, line_number in zip(diff_lines, line_numbers):
        row_cells = table.add_row().cells

        code_cell = row_cells[0]

        # Apply background shading
        if line.startswith("+"):
            fill = config.get("add_color", "D0FFD0")
        elif line.startswith("-"):
            fill = config.get("remove_color", "FFD0D0")
        else:
            fill = config.get("neutral_color", "F5F5F5")
        shading = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls("w"), fill))
        code_cell._element.get_or_add_tcPr().append(shading)

        # Prepare the paragraph for syntax-highlighted runs
        paragraph = code_cell.paragraphs[0]
        paragraph.clear()  # remove any auto-inserted text

        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)

        # Optionally include the diff symbol
        symbol = ""
        code_content = line
        if line.startswith(("+", "-")):
            symbol = line[0]
            code_content = line[1:]

        # Add symbol run
        if symbol:
            run_sym = paragraph.add_run(symbol)
            run_sym.font.name = config.get("diff_font", "Courier New")
            run_sym.font.size = Pt(int(config.get("diff_font_size", 12)))

        # Lex and style each token
        for ttype, value in lex(code_content, lexer):
            value = value.rstrip('\n')

            run = paragraph.add_run(value)
            run.font.name = config.get("diff_font", "Courier New")
            run.font.size = Pt(int(config.get("diff_font_size", 12)))

            style_str = token_styles.get(ttype)
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

# ------------------------------------------------------------------------------
# Main loop

if os.path.exists(output_docx):
    if not ask_yes_no(lang["output_exists"].format(output_docx=output_docx), lang):
        print(lang["exiting"])
        input("Press Enter to exit...")
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

for file in changed_files:
    doc.add_page_break()
    doc.add_heading(f"{lang['file']}: {file}", level=config.get("heading_level", 2) + 1)

    if verbose:
        print(lang["processing_file"].format(file=file))

    # Get file versions
    try:
        old_content = subprocess.run(
            ["git", "show", f"{commit1}:{file}"], capture_output=True, text=True, encoding=file_encoding
        ).stdout.splitlines()
    except:
        old_content = []
    try:
        new_content = subprocess.run(
            ["git", "show", f"{commit2}:{file}"], capture_output=True, text=True, encoding=file_encoding
        ).stdout.splitlines()
    except:
        new_content = []

    # Choose lexer based on filename and content
    sample = "\n".join(new_content or old_content)
    try:
        lexer = guess_lexer_for_filename(file, sample)
    except:
        lexer = guess_lexer(sample or "")

    import difflib
    diff_lines = list(difflib.unified_diff(old_content, new_content, lineterm=""))

    if diff_lines:
        filtered = [
            l for l in diff_lines
            if (l.startswith("+") and not l.startswith("+++")) or
               (l.startswith("-") and not l.startswith("---")) or
               (not l.startswith(("+", "-", "@", "diff", "index")))
        ]
        line_nums = extract_line_numbers(diff_lines)
        add_diff_table(doc, filtered, line_nums, config.get("include_line_numbers", False), lexer)
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
