# GitDiff2Docx
A professional Python utility that converts git diffs into formatted Word documents. The tool helps developers and teams document code changes by creating reports from git commit differences.

## Features
- Convert git diffs to formatted DOCX files
- **Three modes of operation**: CLI (non-interactive), CLI (interactive), and GUI
- Multilingual support with JSON-based localization
- Customizable fonts and styling for diff content
- Color-coded changes (green for additions, red for deletions, gray for context)
- Syntax highlighting with Pygments
- Support for multiple file changes in a single report
- Configurable defaults via config.json
- Ignore specific files or directories using .gddignore
- Built-in and custom themes
- Professional modular Python project structure

## Installation

### Option 1: Using pip (Recommended)
```bash
git clone https://github.com/TD99/gitdiff2docx.git
cd gitdiff2docx
pip install -e .
```

After installation, you can use the `gitdiff2docx` command from anywhere.

### Option 2: Direct usage
1. Clone the repository
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

### GUI Mode (Graphical Interface)
Launch the graphical user interface for easy interaction:
```bash
python gitdiff2docx-gui.py
```

The GUI provides:
- Directory and file browsers
- Visual theme and language selection
- Real-time output logging
- All configuration options in one window

### CLI Mode (Non-Interactive)
Use command-line arguments to run without prompts:
```bash
python gitdiff2docx-cli.py -d /path/to/repo -c1 abc123 -c2 def456 -o output.docx
```

**Available options:**
```
  --create-theme            Create a new theme file and exit
  -d, --target-dir PATH     Target Git directory
  -c1, --commit1 TEXT       First commit hash
  -c2, --commit2 TEXT       Last commit hash
  -o, --output PATH         Output .docx file path
  -c, --config-file PATH    Path to config.json file
  --verbose / --no-verbose  Enable verbose output
  --theme TEXT              Theme name to use
  -l, --language TEXT       Language code (e.g., en, de)
  --help                    Show this message and exit
```

**Examples:**
```bash
# Generate diff between two commits
python gitdiff2docx-cli.py -d . -c1 abc123 -c2 HEAD -o report.docx

# Use specific theme and language
python gitdiff2docx-cli.py -d /repo --theme modern --language en -o diff.docx

# Verbose output with custom config
python gitdiff2docx-cli.py -d . --verbose -c config.json -o out.docx
```

### CLI Mode (Interactive - Legacy)
Run the script and follow the interactive prompts:
```bash
python gitdiff2docx-cli.py
```

The tool will ask for:
- Path of the target git repository
- Source and target commit hashes (if none entered, uses defaults)
- Path of the output DOCX file

Example interaction:
```
< Enter the target Git directory: . 
< Enter the first commit hash (optional): <none> 
< Enter the last commit hash (optional): <none> 
< Enter the output .docx file path (e.g., output.docx): ./test.docx
```
If no commit hashes are provided, the tool automatically compares all commits.

### Using as a Python Package
After installing with pip, you can also use it as a module:
```python
from gitdiff2docx.main import main
import sys

# Set up arguments as if from command line
sys.argv = ['gitdiff2docx', '-d', '.', '-o', 'output.docx']
main()
```

## Project Structure
```
gitdiff2docx/
├── gitdiff2docx/           # Main package
│   ├── cli/                # CLI utilities and interactive prompts
│   ├── config/             # Configuration and schema handling
│   ├── diff/               # Diff processing
│   ├── document/           # Word document generation
│   ├── git/                # Git operations
│   ├── localization/       # Language file handling
│   ├── theme/              # Theme management
│   ├── utils/              # Utility functions
│   └── main.py             # Main CLI entry point
├── gitdiff2docx-cli.py     # CLI wrapper script
├── gitdiff2docx-gui.py     # GUI application
├── diff_tool.py            # Legacy script (preserved for compatibility)
├── config.json             # Configuration file
├── themes/                 # Theme files
├── lang/                   # Language files
├── schemas/                # JSON schemas
└── setup.py                # Package setup
```

## Output Format
The generated DOCX includes:
- Title with commit range
- Section for each changed file
- Color-coded diff tables
- Custom font styling
- Image support
- Syntax highlighting

## Configuration
GitDiff2Docx can be customized using a `config.json` file placed in the repository root.

### Themes
- Theme files are loaded from the `themes/` folder.
- Set `"theme": "<name>"` in `config.json` (without `.json`).
- Built-in themes are `classic`, `modern`, `modern-atlas`, `modern-carbon`, `modern-sand`, and `modern-dark`.
- Add your own theme by creating `themes/<your-theme-name>.json`.
- Optional global overrides can be applied by creating `themes/_overrides.json`.
- Overrides are deep-merged into the selected theme, so only provided subkeys are replaced.
- Theme-level font settings are available via `font.name` and `font.size` (applies to code tables only).
- Theme-level border settings are available via `table_borders` (`visible`, `style`, `weight_pt`, `color`, `space`).

#### Creating a New Theme
Create a new theme template with all required fields prefilled:
```bash
python gitdiff2docx-cli.py --create-theme
```

Select the theme in `config.json`:
```json
"theme": "<your-theme-name>"
```

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

## Development
The project follows a professional Python package structure with:
- Modular design for easy testing and maintenance
- Separation of concerns (config, theme, git, document generation)
- Clean import structure
- Type hints where applicable
- Comprehensive docstrings

## Disclaimer
Use this tool at your own risk. Always verify the generated reports for accuracy.
