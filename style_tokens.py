from __future__ import annotations

from docx.enum.text import WD_COLOR_INDEX
from docx.shared import RGBColor


# Shared semantic tokens for reference/source-like content.
BODY_TEXT_SIZE_PT = 12
REVIEW_TEXT_SIZE_PT = 10
REVIEW_NOTES_TEXT_SIZE_PT = 9
REFERENCE_TEXT_SIZE_PT = 10
REFERENCE_HIGHLIGHT_DEFAULT = WD_COLOR_INDEX.TURQUOISE
REFERENCE_HIGHLIGHT_MARKED = WD_COLOR_INDEX.BRIGHT_GREEN
REFERENCE_LINK_RGB = RGBColor(0x05, 0x63, 0xC1)

# Shared semantic tokens for section labels.
SECTION_LABEL_BLUE_RGB = RGBColor(0x00, 0x70, 0xC0)
DEFAULT_DOCX_ASCII_FONT_NAME = "細明體"
DEFAULT_DOCX_EAST_ASIA_FONT_NAME = "新細明體"
