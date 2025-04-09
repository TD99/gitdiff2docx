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
  "code_changes": "Code Changes",
  "diffs": "Code Changes",
  "enter_commit1": "Enter the first commit hash (optional):",
  "enter_commit2": "Enter the last commit hash (optional):",
  "enter_output_docx": "Enter the output .docx file path (e.g., output.docx):",
  "enter_target_dir": "Enter the target Git directory:",
  "error_opening_file": "Error opening file {file}: {error}",
  "error_removing_file": "Error removing file {output_docx}: {error}",
  "error_saving_file": "Error saving file {output_docx}: {error}",
  "exiting": "Exiting...",
  "file": "File",
  "git_changes_report": "Git Changes Report",
  "invalid_target_dir": "Invalid directory. Please enter a valid path.",
  "legend": "Legend",
  "legend_add": "Added line",
  "legend_context": "Unchanged line",
  "legend_remove": "Removed line",
  "no": "no",
  "no_changes_found": "No changed files found between {commit1} and {commit2}.",
  "no_git_repo_found": "No .git folder found in {target_dir}.",
  "no_significant_changes": "No significant changes found.",
  "output_exists": "The output file {output_docx} already exists. Do you want to overwrite it? (yes/no):",
  "press_enter_to_retry": "Press Enter to retry...",
  "processing_done": "Processing done for file: {file}",
  "processing_file": "Processing file: {file}",
  "report_generated_on": "Report generated on {date}",
  "saving_report": "Report saved to {output_docx}",
  "select_language": "Select a language (available: {languages}):",
  "still_continue": "Do you still want to continue? (yes/no):",
  "title": "Git Diff to Word Document Generator",
  "using_default_output": "Using default output file: {output_docx}",
  "using_first_commit": "Using first commit: {commit1}",
  "using_last_commit": "Using last commit: {commit2}",
  "yes": "yes"
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
