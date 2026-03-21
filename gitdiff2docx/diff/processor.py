"""Diff processing and calculation."""

from difflib import SequenceMatcher
from pygments.lexers import guess_lexer_for_filename, guess_lexer
from gitdiff2docx.utils.file import is_binary_string, is_image_file
from gitdiff2docx.git.operations import get_file_at_commit
from gitdiff2docx.document.tables import add_diff_table
from gitdiff2docx.document.media import add_image


def process_file_diff(doc, file, commit1, commit2, config, theme, token_styles, lang, file_encoding):
    """Process a single file and add its diff to the document."""
    # Get file bytes at both commits
    old_bytes = get_file_at_commit(commit1, file)
    new_bytes = get_file_at_commit(commit2, file)

    # If binary but not image → skip
    if is_binary_string(new_bytes) or is_binary_string(old_bytes):
        if is_image_file(file):
            # Insert only if the image changed
            if old_bytes != new_bytes:
                paragraph = doc.add_paragraph()
                run = paragraph.add_run(lang['image_changed'])
                run.italic = True

                if config.get("include_images", True):
                    add_image(doc, new_bytes, file, lang)
        else:
            paragraph = doc.add_paragraph()
            run = paragraph.add_run(lang['binary_file_skipped'])
            run.italic = True
        return

    # Fallback for text diff
    old_content = old_bytes.decode(file_encoding, errors="ignore").splitlines()
    new_content = new_bytes.decode(file_encoding, errors="ignore").splitlines()

    # Choose lexer based on filename and content
    sample = "\n".join(new_content or old_content)
    try:
        lexer = guess_lexer_for_filename(file, sample)
    except:
        lexer = guess_lexer(sample or "")

    matcher = SequenceMatcher(None, old_content, new_content)
    diff_lines = []
    old_idx = new_idx = 0

    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            # Only add unchanged lines if configured to do so
            if config.get("include_unchanged_lines", True):
                for line in new_content[j1:j2]:
                    diff_lines.append(f" {line}")
                    new_idx += 1
                    old_idx += 1
            else:
                # Still need to update indices even if we don't add the lines
                new_idx += (j2 - j1)
                old_idx += (i2 - i1)
        elif tag == "replace":
            for line in old_content[i1:i2]:
                diff_lines.append(f"-{line}")
                old_idx += 1
            for line in new_content[j1:j2]:
                diff_lines.append(f"+{line}")
                new_idx += 1
        elif tag == "delete":
            for line in old_content[i1:i2]:
                diff_lines.append(f"-{line}")
                old_idx += 1
        elif tag == "insert":
            for line in new_content[j1:j2]:
                diff_lines.append(f"+{line}")
                new_idx += 1

    # Add Table if there are significant changes
    if diff_lines:
        add_diff_table(doc, diff_lines, lexer, theme, config, token_styles)
    else:
        doc.add_paragraph(lang["no_significant_changes"], style="Italic")
