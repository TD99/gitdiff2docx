"""JSON Schema processing utilities."""

import os
from copy import deepcopy
from .loader import load_json_from_path


def resolve_json_pointer(document, pointer):
    """Resolve JSON Pointer (RFC 6901) in a document."""
    if not pointer:
        return document
    if pointer.startswith("/"):
        node = document
        for raw_part in pointer.split("/")[1:]:
            part = raw_part.replace("~1", "/").replace("~0", "~")
            node = node[part]
        return node
    return document


def resolve_schema_ref(ref_value, current_schema_path):
    """Resolve a $ref value in a JSON schema."""
    if "#" in ref_value:
        path_part, pointer_part = ref_value.split("#", 1)
        pointer = pointer_part if pointer_part.startswith("/") else f"/{pointer_part}" if pointer_part else ""
    else:
        path_part, pointer = ref_value, ""

    if not path_part:
        doc, doc_path = load_json_from_path(current_schema_path)
    else:
        target_path = os.path.normpath(os.path.join(os.path.dirname(current_schema_path), path_part))
        doc, doc_path = load_json_from_path(target_path)

    return resolve_json_pointer(doc, pointer), doc_path


def merge_schema_fragments(base, override):
    """Merge two schema fragments with special handling for required/properties."""
    result = deepcopy(base)
    for key, value in override.items():
        if key == "required":
            existing = result.get("required", [])
            for required_key in value:
                if required_key not in existing:
                    existing.append(required_key)
            result["required"] = existing
        elif key in ("properties", "definitions"):
            current = result.get(key, {})
            merged = deepcopy(current)
            for prop_key, prop_val in value.items():
                if prop_key in merged and isinstance(merged[prop_key], dict) and isinstance(prop_val, dict):
                    merged[prop_key] = merge_schema_fragments(merged[prop_key], prop_val)
                else:
                    merged[prop_key] = deepcopy(prop_val)
            result[key] = merged
        else:
            result[key] = deepcopy(value)
    return result


def expand_schema_node(schema_node, current_schema_path):
    """Recursively expand a schema node by resolving $ref and allOf."""
    expanded = {}

    if "$ref" in schema_node:
        ref_node, ref_path = resolve_schema_ref(schema_node["$ref"], current_schema_path)
        expanded = merge_schema_fragments(expanded, expand_schema_node(ref_node, ref_path))

    if "allOf" in schema_node:
        for sub_schema in schema_node["allOf"]:
            expanded = merge_schema_fragments(expanded, expand_schema_node(sub_schema, current_schema_path))

    inline_schema = {}
    for key, value in schema_node.items():
        if key in ("$ref", "allOf"):
            continue
        if key in ("properties", "definitions") and isinstance(value, dict):
            expanded_map = {}
            for nested_key, nested_schema in value.items():
                if isinstance(nested_schema, dict):
                    expanded_map[nested_key] = expand_schema_node(nested_schema, current_schema_path)
                else:
                    expanded_map[nested_key] = deepcopy(nested_schema)
            inline_schema[key] = expanded_map
        else:
            inline_schema[key] = deepcopy(value)

    expanded = merge_schema_fragments(expanded, inline_schema)
    return expanded


def build_defaults_from_schema_node(schema_node, current_schema_path, include_optional_defaults=False):
    """Build default values from a JSON schema node."""
    expanded = expand_schema_node(schema_node, current_schema_path)

    if "default" in expanded:
        return deepcopy(expanded["default"])

    if expanded.get("type") == "object" or "properties" in expanded:
        result = {}
        properties = expanded.get("properties", {})

        for key in expanded.get("required", []):
            prop_schema = properties.get(key)
            if prop_schema is None:
                continue
            result[key] = build_defaults_from_schema_node(prop_schema, current_schema_path)

        if include_optional_defaults:
            for key, prop_schema in properties.items():
                if key in result:
                    continue
                prop_expanded = expand_schema_node(prop_schema, current_schema_path)
                if "default" in prop_expanded:
                    result[key] = deepcopy(prop_expanded["default"])

        return result

    if "default" in expanded:
        return deepcopy(expanded["default"])

    schema_type = expanded.get("type")
    if schema_type == "string":
        return ""
    if schema_type == "boolean":
        return False
    if schema_type in ("integer", "number"):
        return 0
    if schema_type == "array":
        return []
    return None
