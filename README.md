# GitDiff2Docx
A Python utility that converts git diffs into formatted Word documents. The tool helps developers and teams document code changes by creating reports from git commit differences.

## Features
- Convert git diffs to formatted DOCX files
- Multilingual support with JSON-based localization
- Customizable fonts and styling for diff content
- Optional line number display
- Color-coded changes (green for additions, red for deletions, gray for context)
- Support for multiple file changes in a single report

## Installation
1. Clone the repository
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage
Run the script and follow the interactive prompts:
```bash
python diff_tool.py
```

The tool will ask for:
- Source and target commit hashes
- Output DOCX filename
- Whether to include line numbers
- Font preferences for diff content

## Localization
Add language files in the `lang` directory using JSON format. Available languages are automatically detected from JSON files.

Example language structure:
```json
{
  "enter_commit1": "Enter the first commit hash:",
  "enter_commit2": "Enter the last commit hash:",
  "enter_output_docx": "Enter the output .docx file path (e.g., output.docx):",
  "include_line_numbers": "Include line numbers? (yes/no):",
  "yes": "yes",
  "no": "no",
  "select_language": "Select a language (available: {languages}):",
  "no_changes_found": "No changed files found between {commit1} and {commit2}.",
  "saving_report": "Report saved to {output_docx}",
  "file": "File",
  "code_changes": "Code Changes",
  "no_significant_changes": "No significant changes found.",
  "git_changes_report": "Git Changes Report",
  "enter_diff_font": "Enter the font for the code (e.g., Courier New):",
  "enter_diff_font_size": "Enter the font size for the code (e.g., 12):"
}
```

## Output Format
The generated DOCX includes:
- Title with commit range
- Section for each changed file
- Color-coded diff tables
- Optional line numbers
- Custom font styling

## Co-created with AI

This project was co-created with the assistance of GitHub Copilot and ChatGPT, two AI programming assistants.

## Disclaimer
Use this tool at your own risk. Always verify the generated reports for accuracy.
