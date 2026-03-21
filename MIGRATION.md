# Migration Guide: diff_tool.py → GitDiff2Docx 2.0

## Overview

GitDiff2Docx has been completely refactored with a professional Python package structure. This guide helps you migrate from the old `diff_tool.py` to the new modular architecture.

## What Changed?

### Old Structure (v1.x)
```
.
├── diff_tool.py          # Single 889-line file
├── config.json
├── themes/
└── lang/
```

### New Structure (v2.0)
```
.
├── gitdiff2docx/         # Professional package
│   ├── cli/
│   ├── config/
│   ├── diff/
│   ├── document/
│   ├── git/
│   ├── theme/
│   └── utils/
├── gitdiff2docx-cli.py   # CLI entry point
├── gitdiff2docx-gui.py   # GUI application
├── diff_tool.py          # Legacy (preserved)
└── setup.py              # Package installer
```

## Backward Compatibility

**Good news!** The old `diff_tool.py` is still included for backward compatibility. Your existing workflows will continue to work.

### Option 1: Keep Using diff_tool.py (No Changes)
```bash
# This still works exactly as before
python diff_tool.py
```

### Option 2: Migrate to New CLI (Recommended)
```bash
# Same interactive mode
python gitdiff2docx-cli.py

# Or use new non-interactive mode
python gitdiff2docx-cli.py -d . -o output.docx
```

## Migration Steps

### Step 1: Test Compatibility
Run your existing commands with the new CLI:
```bash
# Old way (still works)
python diff_tool.py

# New way (equivalent)
python gitdiff2docx-cli.py
```

### Step 2: Update Scripts
If you have automation scripts, update them to use the new CLI:

**Before:**
```bash
# Manual input required
python diff_tool.py
```

**After:**
```bash
# Fully automated
python gitdiff2docx-cli.py \
  -d /path/to/repo \
  -c1 abc123 \
  -c2 HEAD \
  -o report.docx
```

### Step 3: Install as Package (Optional)
For system-wide access:
```bash
pip install -e .
```

Then use anywhere:
```bash
gitdiff2docx -d /any/repo -o report.docx
```

## Feature Comparison

| Feature | Old (v1.x) | New (v2.0) |
|---------|------------|------------|
| Interactive mode | ✅ | ✅ |
| Non-interactive CLI | ❌ | ✅ |
| GUI application | ❌ | ✅ |
| Theme support | ✅ | ✅ |
| Multi-language | ✅ | ✅ |
| Package install | ❌ | ✅ |
| Modular code | ❌ | ✅ |
| CI/CD friendly | ⚠️ | ✅ |

## Configuration Migration

### No Changes Required!
Your existing `config.json` works without modification.

```json
{
  "language": "en",
  "theme": "classic",
  "verbose": true,
  "open_after_creation": false
}
```

All options remain the same:
- ✅ Themes work identically
- ✅ Language files unchanged
- ✅ All settings supported
- ✅ `.gddignore` still works

## Command Equivalents

### Creating a Theme
**Old:**
```bash
python diff_tool.py --create-theme
```

**New:**
```bash
python gitdiff2docx-cli.py --create-theme
```

### Interactive Mode
**Old:**
```bash
python diff_tool.py
```

**New:**
```bash
python gitdiff2docx-cli.py
```

### Specifying Output
**Old:**
```
# Prompted during execution
Enter the output .docx file path: output.docx
```

**New:**
```bash
# CLI argument
python gitdiff2docx-cli.py -o output.docx

# Or still prompted in interactive mode
python gitdiff2docx-cli.py
```

## New Capabilities

### 1. Non-Interactive Mode
Perfect for automation and CI/CD:
```bash
python gitdiff2docx-cli.py \
  -d . \
  -c1 main \
  -c2 feature-branch \
  -o comparison.docx \
  --force
```

### 2. GUI Application
Easy point-and-click interface:
```bash
python gitdiff2docx-gui.py
```

### 3. Advanced Options
```bash
# Override config theme
python gitdiff2docx-cli.py --theme modern -o out.docx

# Override language
python gitdiff2docx-cli.py --language en -o out.docx

# Force overwrite
python gitdiff2docx-cli.py -o existing.docx --force

# Custom config file
python gitdiff2docx-cli.py -c custom-config.json -o out.docx
```

## Troubleshooting

### "Module not found" Error
Install dependencies:
```bash
pip install -r requirements.txt
```

### GUI Won't Start
Install tkinter:
```bash
# Ubuntu/Debian
sudo apt-get install python3-tk

# macOS (usually pre-installed)
brew install python-tk

# Windows (usually pre-installed)
```

### Import Errors
Make sure you're in the repository root:
```bash
cd /path/to/gitdiff2docx
python gitdiff2docx-cli.py
```

## Best Practices

### For Development
Use the modular structure:
```python
from gitdiff2docx.config import load_theme
from gitdiff2docx.document import create_document
```

### For Automation
Use CLI with explicit parameters:
```bash
python gitdiff2docx-cli.py \
  -d "$REPO_PATH" \
  -c1 "$COMMIT1" \
  -c2 "$COMMIT2" \
  -o "$OUTPUT" \
  --force \
  --no-verbose
```

### For End Users
Use the GUI:
```bash
python gitdiff2docx-gui.py
```

## Rollback Plan

If you encounter issues, you can always use the old `diff_tool.py`:
```bash
# Original version still works
python diff_tool.py
```

The old file is preserved for backward compatibility and will continue to function.

## Support

For questions or issues:
1. Check [USAGE.md](USAGE.md) for detailed examples
2. Review [README.md](README.md) for features
3. Create a GitHub issue with:
   - Python version
   - Operating system
   - Error messages
   - Steps to reproduce

## Summary

✅ **No breaking changes** - Old workflows still work  
✅ **Backward compatible** - Keep using diff_tool.py if needed  
✅ **New features** - CLI, GUI, packaging  
✅ **Easy migration** - Simple command updates  
✅ **Better maintainability** - Modular code structure
