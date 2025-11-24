"""Entry point for running gitdiff2docx as a module."""

import sys

# Import and run the main function from main.py
if __name__ == "__main__":
    # Import here to avoid circular imports
    from gitdiff2docx.main import main
    sys.exit(main())
