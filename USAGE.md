# GitDiff2Docx Usage Guide

This guide covers all three modes of operation for GitDiff2Docx.

## Quick Start

### GUI Mode (Easiest)
```bash
python gitdiff2docx-gui.py
```
- Point and click interface
- File browsers for easy selection
- Visual theme/language selection
- Real-time output logging

### CLI Non-Interactive Mode (Automation)
```bash
python gitdiff2docx-cli.py -d /path/to/repo -c1 abc123 -c2 HEAD -o report.docx
```

### CLI Interactive Mode (Legacy)
```bash
python gitdiff2docx-cli.py
```
Follow the prompts to enter repository path, commits, and output file.

## Detailed Examples

### 1. Compare Two Specific Commits
```bash
python gitdiff2docx-cli.py \
  -d /path/to/myproject \
  -c1 a1b2c3d \
  -c2 e4f5g6h \
  -o changes.docx \
  --theme modern \
  --language en
```

### 2. Compare First Commit to HEAD
```bash
# Automatically uses first commit and HEAD
python gitdiff2docx-cli.py -d . -o full-history.docx
```

### 3. Compare with Custom Theme
```bash
# First create a theme
python gitdiff2docx-cli.py --create-theme
# Name it: my-theme

# Then use it
python gitdiff2docx-cli.py -d . -o output.docx --theme my-theme
```

### 4. Verbose Output for Debugging
```bash
python gitdiff2docx-cli.py -d . -o debug.docx --verbose
```

### 5. Force Overwrite Existing File
```bash
python gitdiff2docx-cli.py -d . -o existing.docx --force
```

### 6. Use Custom Config File
```bash
python gitdiff2docx-cli.py -d . -o out.docx -c my-config.json
```

## Configuration Options

Edit `config.json` to customize defaults:

```json
{
  "language": "en",              // UI language (en, de)
  "verbose": false,              // Show detailed progress
  "theme": "classic",            // Theme name
  "diff_font": "Courier New",    // Code font
  "diff_font_size": 8,           // Font size in points
  "open_after_creation": false,  // Auto-open document
  "include_unchanged_lines": false,  // Show context lines
  "include_images": true,        // Embed changed images
  "insert_page_breaks": true,    // Page break between files
  "file_encoding": "utf-8"       // Source file encoding
}
```

## Ignoring Files

Create a `.gddignore` file in your repository:
```
# Ignore build artifacts
*.min.js
*.min.css
build/
dist/

# Ignore dependencies
node_modules/
vendor/
```

## Themes

### Built-in Themes
- `classic` - Traditional diff colors
- `modern` - Clean, minimalist
- `modern-atlas` - Blue accent
- `modern-carbon` - Dark theme
- `modern-sand` - Warm colors
- `modern-dark` - High contrast

### Create Custom Theme
```bash
python gitdiff2docx-cli.py --create-theme
```

This generates a template with all required fields. Edit `themes/your-theme.json` to customize colors, fonts, borders, etc.

## CLI Options Reference

```
Options:
  --create-theme              Create a new theme file and exit
  -d, --target-dir PATH       Target Git directory
  -c1, --commit1 TEXT         First commit hash
  -c2, --commit2 TEXT         Last commit hash
  -o, --output PATH           Output .docx file path
  -c, --config-file PATH      Path to config.json file
  --verbose / --no-verbose    Enable verbose output
  --theme TEXT                Theme name to use
  -l, --language TEXT         Language code (e.g., en, de)
  -f, --force                 Force overwrite output file
  --help                      Show help and exit
```

## Tips and Tricks

1. **Auto Commits**: Leave commit fields empty to use defaults (first commit → HEAD)
2. **Relative Paths**: Use `.` for current directory
3. **Commit Shortcuts**: Use `HEAD`, `HEAD~1`, `main`, branch names
4. **Batch Processing**: Use CLI mode in shell scripts for automation
5. **CI/CD Integration**: Generate reports automatically in your pipeline

## Troubleshooting

### "No changes found"
- Check commit hashes are valid
- Verify you're in a git repository
- Ensure commits exist in current branch

### "Language not found"
- Check `lang/` directory for available languages
- Verify language code in config.json
- Fallback to English if language missing

### "Theme not found"
- List themes: `ls themes/`
- Verify theme name (without .json)
- Check for typos in config.json

### GUI Not Working
- Ensure tkinter is installed: `python -m tkinter`
- On Linux: `sudo apt-get install python3-tk`

## Integration Examples

### Git Alias
Add to `.gitconfig`:
```ini
[alias]
    diff2doc = "!f() { python /path/to/gitdiff2docx-cli.py -d . -c1 $1 -c2 ${2:-HEAD} -o diff.docx; }; f"
```

Usage: `git diff2doc abc123 def456`

### Shell Script
```bash
#!/bin/bash
# generate-report.sh
REPO_PATH="${1:-.}"
OUTPUT="${2:-report.docx}"

python gitdiff2docx-cli.py \
  -d "$REPO_PATH" \
  -o "$OUTPUT" \
  --theme modern \
  --language en \
  --verbose
```

### Python Script
```python
import subprocess
import sys

def generate_diff_report(repo, commit1, commit2, output):
    cmd = [
        'python', 'gitdiff2docx-cli.py',
        '-d', repo,
        '-c1', commit1,
        '-c2', commit2,
        '-o', output,
        '--force'
    ]
    subprocess.run(cmd, check=True)

if __name__ == '__main__':
    generate_diff_report('.', 'abc123', 'HEAD', 'report.docx')
```

## Package Installation

Install as a package for global access:
```bash
pip install -e .
```

Then use anywhere:
```bash
gitdiff2docx -d /any/repo -o report.docx
```

## Support

For issues, feature requests, or contributions:
- GitHub: https://github.com/TD99/gitdiff2docx
- Create an issue with details about your problem
- Include: OS, Python version, error messages, config
