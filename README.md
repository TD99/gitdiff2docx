# GitDiff2Docx
A Python utility that converts git diffs into formatted Word documents. The tool helps developers and teams document code changes by creating reports from git commit differences.

## Features
- Convert git diffs to formatted DOCX files
- Multilingual support with JSON-based localization
- Customizable fonts and styling for diff content
- Optional line number display
- Color-coded changes (green for additions, red for deletions, gray for context)
- Syntax highlighting
- Support for multiple file changes in a single report
- Configurable defaults via config.json
- Ignore specific files or directories using .gddignore

## Installation

### Option 1: Install from PyPI (when published)
```bash
pip install gitdiff2docx
```

### Option 2: Install from Source
1. Clone the repository:
   ```bash
   git clone https://github.com/TD99/gitdiff2docx.git
   cd gitdiff2docx
   ```

2. Install the package:
   ```bash
   pip install .
   ```

   Or for development mode:
   ```bash
   pip install -e .
   ```

Once installed, you can use the `gdd` command from anywhere in your terminal.

## Usage
Run the tool using the `gdd` command and follow the interactive prompts:
```bash
gdd
```

Alternatively, you can still run it as a Python module:
```bash
python -m gitdiff2docx
```

Or if running from source without installation:
```bash
python diff_tool.py
```

The tool will ask for:
- Path of the target git repository
- Source and target commit hashes (if none entered previously)
- Path of the output DOCX file

Example interaction:
```
< Enter the target Git directory: . 
< Enter the first commit hash (optional): <none> 
< Enter the last commit hash (optional): <none> 
< Enter the output .docx file path (e.g., output.docx): ./test.docx
```
If no commit hashes are provided, the tool automatically compares the latest commit (HEAD) with its predecessor.

## Output Format
The generated DOCX includes:
- Title with commit range
- Section for each changed file
- Color-coded diff tables
- Custom font styling
- Image support
- Syntax highlighting

## Configuration
GitDiff2Docx can be customized using a `config.json` file. When installed via pip, you can create this file in your home directory or current working directory. The tool will look for configuration in the following order:
1. `config.json` in the current directory
2. Default configuration from the package installation

When running from source, place `config.json` in the same directory as `diff_tool.py`.

## Localization
- The default interface and output language is German (`de`).
- To change it, set the `language` value in your `config.json` file to a language code available in the `lang` directory.
   - Default languages: `de` (German), `en` (English)
- Additional languages can be added by creating new JSON files in the `lang` directory.
   - Each file should be named according to its language code (e.g. `fr.json` for French).
   - GitDiff2Docx automatically detects all available languages based on the JSON files present in the `lang` directory.

## Ignoring Files and Directories
You can exclude files or directories from being included in the diff report using a `.gddignore` file located in the target repository.  
This file works exactly like a `.gitignore`, supporting the same syntax and matching rules.

## Disclaimer
Use this tool at your own risk. Always verify the generated reports for accuracy.
