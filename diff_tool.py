import os
import json
import subprocess
from docx import Document
from docx.shared import Pt
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

script_dir = os.path.dirname(os.path.abspath(__file__))

# ------------------------------------------------------------------------------
# Localization

lang_dir = os.path.join(script_dir, "lang")

if not os.path.exists(lang_dir):
    print(f"Error: Language directory not found: {lang_dir}")
    exit()

available_languages = [f.split(".json")[0] for f in os.listdir(lang_dir) if f.endswith(".json")]

print(f"Available languages: {', '.join(available_languages)}")
lang_choice = input("Select a language: ").strip().lower()

lang_file = os.path.join(lang_dir, f"{lang_choice}.json")
if not os.path.exists(lang_file):
    print(f"Language '{lang_choice}' not found, defaulting to English.")
    lang_file = os.path.join(lang_dir, "en.json")

with open(lang_file, "r", encoding="utf-8") as f:
    lang = json.load(f)

# ------------------------------------------------------------------------------
# Input prompts

# Ask for commit hashes and output file
commit1 = input(lang["enter_commit1"] + " ")
commit2 = input(lang["enter_commit2"] + " ")
output_docx = input(lang["enter_output_docx"] + " ")

# Ask if line numbers should be included
while True:
    include_line_numbers_input = input(lang["include_line_numbers"] + " ").strip().lower()
    if include_line_numbers_input in [lang["yes"], lang["no"]]:
        include_line_numbers = include_line_numbers_input == lang["yes"]
        break

# Ask which font should be used for diff content
diff_font = input(lang["enter_diff_font"] + " ").strip()
diff_font_size = input(lang["enter_diff_font_size"] + " ").strip()

# ------------------------------------------------------------------------------
# Diff generation

changed_files = subprocess.run(
    ["git", "diff", "--name-only", commit1, commit2], capture_output=True, text=True
).stdout.splitlines()

if not changed_files:
    print(lang["no_changes_found"].format(commit1=commit1, commit2=commit2))
    exit()

# ------------------------------------------------------------------------------
# Word document generation

doc = Document()
doc.add_heading(f"{lang['git_changes_report']} ({commit1} → {commit2})", level=1)

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
        run.font.name = diff_font
        run.font.size = Pt(int(diff_font_size))

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

# Process each changed file
for file in changed_files:
    doc.add_page_break()
    doc.add_heading(f"{lang['file']}: {file}", level=2)

    # Get the git diff for this file (only modified lines)
    diff_output = subprocess.run(
        ["git", "diff", "-U0", commit1, commit2, "--", file], capture_output=True, text=True
    ).stdout

    if diff_output:
        doc.add_paragraph(lang["code_changes"] + ":", style="Heading 3")
        diff_lines = [line for line in diff_output.splitlines() if line.startswith("+") or line.startswith("-") or line.startswith("@@")]
        line_numbers = extract_line_numbers(diff_lines)
        add_diff_table(doc, diff_lines, line_numbers, include_line_numbers)
    else:
        doc.add_paragraph(lang["no_significant_changes"], style="Italic")

doc.save(output_docx)
print(lang["saving_report"].format(output_docx=output_docx))
