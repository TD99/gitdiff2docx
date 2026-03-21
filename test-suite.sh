#!/bin/bash
# Test script for GitDiff2Docx functionality

echo "=== GitDiff2Docx Test Suite ==="
echo ""

# Test 1: CLI Help
echo "Test 1: CLI Help"
python gitdiff2docx-cli.py --help > /dev/null 2>&1
if [ $? -eq 0 ]; then
    echo "✓ CLI help works"
else
    echo "✗ CLI help failed"
fi
echo ""

# Test 2: Non-interactive mode
echo "Test 2: Non-interactive mode"
rm -f /tmp/test-cli.docx
python gitdiff2docx-cli.py -d . -c1 HEAD~3 -c2 HEAD -o /tmp/test-cli.docx --language en --no-verbose > /dev/null 2>&1
if [ -f /tmp/test-cli.docx ]; then
    echo "✓ Non-interactive mode works"
    echo "  Generated: $(ls -lh /tmp/test-cli.docx | awk '{print $5}')"
else
    echo "✗ Non-interactive mode failed"
fi
echo ""

# Test 3: Theme switching
echo "Test 3: Theme switching"
rm -f /tmp/test-theme.docx
python gitdiff2docx-cli.py -d . -c1 HEAD~3 -c2 HEAD -o /tmp/test-theme.docx --theme modern --language en --no-verbose > /dev/null 2>&1
if [ -f /tmp/test-theme.docx ]; then
    echo "✓ Theme switching works"
else
    echo "✗ Theme switching failed"
fi
echo ""

# Test 4: Force overwrite
echo "Test 4: Force overwrite"
python gitdiff2docx-cli.py -d . -c1 HEAD~3 -c2 HEAD -o /tmp/test-cli.docx --force --language en --no-verbose > /dev/null 2>&1
if [ $? -eq 0 ]; then
    echo "✓ Force overwrite works"
else
    echo "✗ Force overwrite failed"
fi
echo ""

# Test 5: Package structure
echo "Test 5: Package structure"
if [ -d "gitdiff2docx/cli" ] && [ -d "gitdiff2docx/config" ] && [ -d "gitdiff2docx/document" ]; then
    echo "✓ Package structure is correct"
else
    echo "✗ Package structure is incorrect"
fi
echo ""

# Test 6: Module imports
echo "Test 6: Module imports"
python -c "from gitdiff2docx.main import main; from gitdiff2docx.config.constants import SCRIPT_DIR" 2>&1
if [ $? -eq 0 ]; then
    echo "✓ Module imports work"
else
    echo "✗ Module imports failed"
fi
echo ""

# Summary
echo "=== Test Summary ==="
echo "All tests completed. Check output above for results."
echo ""
echo "Generated files:"
ls -lh /tmp/test-*.docx 2>/dev/null | awk '{print "  " $9 " - " $5}'
