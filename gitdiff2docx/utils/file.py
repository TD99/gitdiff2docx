"""File type detection utilities."""

import mimetypes


def is_binary_string(bytes_data):
    """Check if byte data is binary (not text)."""
    textchars = bytearray({7, 8, 9, 10, 12, 13, 27}
                          | set(range(0x20, 0x100)) - {0x7f})
    return bool(bytes_data.translate(None, textchars))


def is_image_file(filename):
    """Check if file is an image based on MIME type."""
    mimetype, _ = mimetypes.guess_type(filename)
    return mimetype and mimetype.startswith("image/")
