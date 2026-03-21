#!/usr/bin/env python3
"""
GitDiff2Docx CLI entry point (backward compatibility wrapper).

This script maintains backward compatibility with the original diff_tool.py.
"""

import sys
import os

# Add the parent directory to the path so we can import gitdiff2docx
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from gitdiff2docx.main import main

if __name__ == '__main__':
    sys.exit(main())
