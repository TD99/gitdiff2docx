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

## Output Format
The generated DOCX includes:
- Title with commit range
- Section for each changed file
- Color-coded diff tables
- Custom font styling
- Image support
- Syntax highlighting

## Disclaimer
Use this tool at your own risk. Always verify the generated reports for accuracy.
