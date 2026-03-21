"""Media handling for documents (images, etc.)."""

import io
from PIL import Image
from docx.shared import Inches


def add_image(document, file_bytes, image_name, lang):
    """Add an image to the document with auto-sizing."""
    image_stream = io.BytesIO(file_bytes)
    try:
        image = Image.open(image_stream)
        width, _ = image.size
        image_stream.seek(0)  # rewind for docx
        max_width_inches = 6  # ~75% of page width (8 inches)
        width_inches = min(max_width_inches, width / image.info.get('dpi', (96, 96))[0])
        document.add_picture(image_stream, width=Inches(width_inches))
        document.paragraphs[-1].alignment = 1  # center
    except Exception:
        document.add_paragraph(lang["error_inserting_image"].format(image_name=image_name))
