import os
import json
import subprocess
from docx import Document
from docx.shared import Pt
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from datetime import datetime

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

# Ask for the target directory
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

# Ask for optional commit hashes
commit1 = input(lang["enter_commit1"] + " ").strip()
if not commit1:
    commit1 = subprocess.run(["git", "rev-list", "--max-parents=0", "HEAD"], capture_output=True, text=True, encoding="utf-8").stdout.strip()
    print(lang["using_first_commit"].format(commit1=commit1))

commit2 = input(lang["enter_commit2"] + " ").strip()
if not commit2:
    commit2 = subprocess.run(["git", "rev-parse", "HEAD"], capture_output=True, text=True, encoding="utf-8").stdout.strip()
    print(lang["using_last_commit"].format(commit2=commit2))

# Ask for the output file
output_docx = input(lang["enter_output_docx"] + " ")
if not output_docx:
    output_docx = os.path.join(script_dir, "output.docx")
    print(lang["using_default_output"].format(output_docx=output_docx))

# ------------------------------------------------------------------------------
# Diff generation

changed_files = subprocess.run(
    ["git", "diff", "--name-only", commit1, commit2], capture_output=True, text=True, encoding="utf-8"
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
        (lang["legend_add"], "D0FFD0"),
        (lang["legend_remove"], "FFD0D0"),
        (lang["legend_context"], "F5F5F5")
    ]

    for label, color in legend_data:
        row = legend_table.add_row().cells
        row[0].text = label
        run = row[1].paragraphs[0].add_run(" ")
        shading = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls("w"), color))
        row[1]._element.get_or_add_tcPr().append(shading)

doc.add_heading(lang["legend"], level=2)
add_legend_table(doc)

# Extract line numbers from git diff
def extract_line_numbers(diff_lines):
    line_numbers = []
    current_line = 0
    for line in diff_lines:
        if line.startswith("@@"):
            parts = line.split(" ")
            new_file_info = parts[2]
            start_line = int(new_file_info.split(",")[0][1:])  # Line number
            current_line = start_line
        elif not line.startswith("-"):
            line_numbers.append(current_line)
            current_line += 1
    return line_numbers

# Add a formatted code diff table
def add_diff_table(document, diff_lines, line_numbers, include_numbers):
    table = document.add_table(rows=0, cols=2 if include_numbers else 1)
    table.style = "Table Grid"

    for line, line_number in zip(diff_lines, line_numbers):
        row_cells = table.add_row().cells

        if include_numbers:
            row_cells[0].text = str(line_number)
            row_cells[0].paragraphs[0].runs[0].font.size = Pt(9)

        row_cells[-1].text = line.strip()

        # Apply font
        run = row_cells[-1].paragraphs[0].runs[0]
        run.font.name = config.get("diff_font", "Courier New")
        run.font.size = Pt(int(config.get("diff_font_size", 12)))

        # Apply background color based on type of change
        if line.startswith("+"):
            shading_elm = parse_xml(r'<w:shd {} w:fill="D0FFD0"/>'.format(nsdecls("w")))  # Light Green: Additions
        elif line.startswith("-"):
            shading_elm = parse_xml(r'<w:shd {} w:fill="FFD0D0"/>'.format(nsdecls("w")))  # Light Red: Deletions
        else:
            shading_elm = parse_xml(r'<w:shd {} w:fill="F5F5F5"/>'.format(nsdecls("w")))  # Light Gray: Context
        row_cells[-1]._element.get_or_add_tcPr().append(shading_elm)

# ------------------------------------------------------------------------------
# Main loop

# Check if the document can be created or overwritten
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

# Process each changed file
verbose = config.get("verbose", False)

for file in changed_files:
    doc.add_page_break()
    doc.add_heading(f"{lang['file']}: {file}", level=2)

    if verbose:
        print(lang["processing_file"].format(file=file))

    
    # Get the full content of the file in both commits
    try:
        old_content = subprocess.run(
            ["git", "show", f"{commit1}:{file}"], capture_output=True, text=True, encoding="utf-8"
        ).stdout.splitlines()
    except Exception:
        old_content = []

    try:
        new_content = subprocess.run(
            ["git", "show", f"{commit2}:{file}"], capture_output=True, text=True, encoding="utf-8"
        ).stdout.splitlines()
    except Exception:
        new_content = []

    import difflib
    diff_lines = list(difflib.unified_diff(old_content, new_content, lineterm=""))

    if diff_lines:
        doc.add_paragraph(lang["code_changes"] + ":", style="Heading 3")
        # Remove hunk headers (lines starting with @@)
        filtered_diff_lines = [
            line for line in diff_lines
            if (line.startswith("+") and not line.startswith("+++")) or
               (line.startswith("-") and not line.startswith("---")) or
               (not line.startswith(("+", "-", "@", "diff", "index")))
        ]
        line_numbers = extract_line_numbers(diff_lines)
        add_diff_table(doc, filtered_diff_lines, line_numbers, config.get("include_line_numbers", False))
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