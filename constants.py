"""
UWM brand colors used in the marked document output.

Hex values are used directly in OOXML shading elements; RGBColor values
are assigned to python-docx run.font.color.rgb.
"""

from docx.shared import RGBColor

# ---------------------------------------------------------------------------
# Section-banner background colors (hex, no leading #)
# ---------------------------------------------------------------------------
COLOR_ACCESSIBILITY_HEX = '035056'   # dark teal  — white text
COLOR_MISSING_HEX       = 'F79651'   # orange     — black text
COLOR_TONE_HEX          = '14364D'   # dark navy  — white text
COLOR_QUALITY_HEX       = 'CFD25B'   # lime green — black text

# ---------------------------------------------------------------------------
# Matching foreground (text) colors
# ---------------------------------------------------------------------------
COLOR_ACCESSIBILITY_TEXT = RGBColor(255, 255, 255)
COLOR_MISSING_TEXT       = RGBColor(0,   0,   0)
COLOR_TONE_TEXT          = RGBColor(255, 255, 255)
COLOR_QUALITY_TEXT       = RGBColor(0,   0,   0)
