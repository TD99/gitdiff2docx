"""JSON loading and caching utilities."""

import os
import json

# Global cache for loaded JSON files
_schema_cache = {}


def load_json_from_path(file_path):
    """Load JSON file with caching."""
    abs_path = os.path.abspath(file_path)
    if abs_path not in _schema_cache:
        with open(abs_path, "r", encoding="utf-8") as f:
            _schema_cache[abs_path] = json.load(f)
    return _schema_cache[abs_path], abs_path
