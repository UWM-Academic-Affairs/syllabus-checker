#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Shared data classes and utility functions used across syllabus checker modules.
"""

from typing import Tuple, Optional
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def calculate_relative_luminance(rgb: Tuple[int, int, int]) -> float:
    """Calculate relative luminance for a color (WCAG formula)"""
    def adjust_channel(channel: int) -> float:
        c = channel / 255.0
        if c <= 0.03928:
            return c / 12.92
        else:
            return ((c + 0.055) / 1.055) ** 2.4

    r, g, b = rgb
    r_adj = adjust_channel(r)
    g_adj = adjust_channel(g)
    b_adj = adjust_channel(b)

    return 0.2126 * r_adj + 0.7152 * g_adj + 0.0722 * b_adj


def calculate_contrast_ratio(color1: Tuple[int, int, int], color2: Tuple[int, int, int]) -> float:
    """Calculate contrast ratio between two colors (WCAG formula)"""
    lum1 = calculate_relative_luminance(color1)
    lum2 = calculate_relative_luminance(color2)

    lighter = max(lum1, lum2)
    darker = min(lum1, lum2)

    return (lighter + 0.05) / (darker + 0.05)


# ---------------------------------------------------------------------------
# Word comment helpers (XML-level, no high-level python-docx support exists)
# ---------------------------------------------------------------------------

def _ensure_comments_part(doc):
    """Get or create the comments.xml part of a docx document."""
    from docx.opc.part import Part
    from docx.opc.packuri import PackURI

    COMMENTS_URI      = '/word/comments.xml'
    COMMENTS_CT       = 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml'
    COMMENTS_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments'

    try:
        return doc.part.part_related_by(COMMENTS_REL_TYPE)
    except KeyError:
        pass

    empty_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'
    ).encode('utf-8')
    part = Part(PackURI(COMMENTS_URI), COMMENTS_CT, empty_xml, doc.part.package)
    doc.part.relate_to(part, COMMENTS_REL_TYPE)
    return part


def _add_word_comment(doc, para, comment_id, comment_text, author='Syllabus Checker'):
    """Attach a Word review comment to *para* (a python-docx Paragraph)."""
    from lxml import etree
    from datetime import datetime, timezone

    W   = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    XML = 'http://www.w3.org/XML/1998/namespace'

    comments_part = _ensure_comments_part(doc)
    root = etree.fromstring(comments_part._blob)

    # Build the <w:comment> element
    dt = datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
    comment_el = etree.SubElement(root, f'{{{W}}}comment')
    comment_el.set(f'{{{W}}}id',       str(comment_id))
    comment_el.set(f'{{{W}}}author',   author)
    comment_el.set(f'{{{W}}}date',     dt)
    comment_el.set(f'{{{W}}}initials', 'SC')

    for line in comment_text.split('\n'):
        p_el = etree.SubElement(comment_el, f'{{{W}}}p')
        r_el = etree.SubElement(p_el, f'{{{W}}}r')
        t_el = etree.SubElement(r_el, f'{{{W}}}t')
        t_el.text = line
        if line != line.strip():
            t_el.set(f'{{{XML}}}space', 'preserve')

    comments_part._blob = etree.tostring(
        root, xml_declaration=True, encoding='UTF-8', standalone=True
    )

    # Wire the comment into the paragraph XML
    p = para._p
    cs = OxmlElement('w:commentRangeStart')
    cs.set(qn('w:id'), str(comment_id))
    p.insert(0, cs)

    ce = OxmlElement('w:commentRangeEnd')
    ce.set(qn('w:id'), str(comment_id))
    p.append(ce)

    cr_run = OxmlElement('w:r')
    cr_rpr = OxmlElement('w:rPr')
    cr_style = OxmlElement('w:rStyle')
    cr_style.set(qn('w:val'), 'CommentReference')
    cr_rpr.append(cr_style)
    cr_run.append(cr_rpr)
    cr_ref = OxmlElement('w:commentReference')
    cr_ref.set(qn('w:id'), str(comment_id))
    cr_run.append(cr_ref)
    p.append(cr_run)


class ParagraphInfo:
    """Stores information about a paragraph and its location"""
    def __init__(self, paragraph, index: int, location: str,
                 table_idx: Optional[int] = None,
                 row_idx: Optional[int] = None,
                 cell_idx: Optional[int] = None):
        self.paragraph = paragraph
        self.index = index  # Global paragraph index
        self.location = location  # Human-readable location
        self.table_idx = table_idx
        self.row_idx = row_idx
        self.cell_idx = cell_idx


class AccessibilityIssue:
    """Represents a specific accessibility issue in the document"""
    def __init__(self, issue_type: str, description: str, location: str,
                 para_info: Optional[ParagraphInfo] = None,
                 table_index: Optional[int] = None):
        self.issue_type = issue_type
        self.description = description
        self.location = location
        self.para_info = para_info
        self.table_index = table_index
