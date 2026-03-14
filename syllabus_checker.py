#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Syllabus Checker - Assesses syllabi against UWM template standards
Checks for missing sections and accessibility issues
Analyzes content both in document body AND within table cells
Generates both a text report and a marked-up Word document with tracked issues
"""

import argparse
import html
import re
import os
from typing import List, Dict, Tuple, Optional
from docx import Document
from docx.shared import RGBColor, Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from openai import OpenAI


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


class SyllabusChecker:
    """Analyzes a syllabus document for completeness and accessibility"""

    # User-friendly labels for Word comments and end-of-doc sections.
    # Internal issue_type codes stay unchanged; only the display string changes.
    ISSUE_DISPLAY_LABELS = {
        # Heading structure
        'ALL_CAPS_HEADING':                   'Heading Uses All Caps',
        'LONG_HEADING':                        'Heading Is Too Long',
        'H1_NOT_AT_TOP':                       'Move Heading 1 to Top of Document',
        'MULTIPLE_H1':                         'Multiple Heading 1s Found',
        # Font usage
        'SMALL_FONT':                          'Font Size Too Small',
        'DECORATIVE_FONT':                     'Replace Decorative Font',
        'INCONSISTENT_FONTS':                  'Multiple Fonts Detected',
        'BROKEN_STYLE_COPIED_CONTENT':         'Copied Content — Clean Up Styles',
        # Text formatting
        'EXCESSIVE_BOLD':                      'Use a Heading Style Instead of Manual Bold',
        'EXCESSIVE_ITALIC':                    'Limit Continuous Italic Text',
        'EXCESSIVE_UNDERLINE':                 'Underline Is Reserved for Hyperlinks',
        'INCONSISTENT_FORMATTING':             'Inconsistent Inline Formatting',
        'ALL_CAPS_BLOCK':                      'Avoid All-Caps Text Blocks',
        'FULL_JUSTIFICATION':                  'Change to Left-Aligned Text',
        'LOW_LINE_SPACING':                    'Increase Line Spacing',
        'UNDERLINE_NON_LINK':                  'Underline Is Reserved for Hyperlinks',
        'PSEUDO_TABLE':                        'Use a Word Table Instead of Tabs/Spaces',
        # Color & contrast
        'LOW_CONTRAST':                        'Low Color Contrast',
        'COLOR_ONLY_MEANING':                  'Color Used as the Only Visual Indicator',
        'COLOR_CODED_TABLE':                   'Table Uses Color to Convey Meaning',
        'TEXT_OVER_BACKGROUND':                'Low Contrast: Text Over Colored Background',
        # Tables
        'TABLE_NO_HEADER':                     'Table Missing a Header Row',
        'TABLE_MISSING_SCOPE':                 "Enable 'Repeat Header Rows' on This Table",
        'TABLE_MISSING_CAPTION':               'Table Missing Alt Text Description',
        'LAYOUT_TABLE':                        'Table Used for Layout Instead of Data',
        'EMPTY_TABLE_ROW':                     'Delete Empty Table Row',
        'EMPTY_TABLE_COLUMN':                  'Delete Empty Table Column',
        'TABLE_MERGED_CELLS':                  'Merged Cells Disrupt Reading Order',
        'TABLE_INCONSISTENT_NUMERIC_ALIGNMENT':'Inconsistent Number Alignment in Column',
        'TABLE_COMPLEX_READING_ORDER':         'Complex Table Structure',
        'TABLE_COLOR_CODED_MEANING':           'Table Uses Color to Convey Meaning',
        'TABLE_EMBEDDED_IMAGES':               'Images Embedded in Table Cells',
        # Links & navigation
        'NON_DESCRIPTIVE_LINK':                'Add Descriptive Link Text',
        'LONG_URL':                            'Replace URL with Descriptive Link Text',
        'UNSTYLED_LINK':                       'Hyperlink Looks Like Regular Text',
        'MISSING_TOC':                         'Add a Table of Contents',
        'MISSING_BOOKMARKS':                   'Add Internal Navigation Links',
        # Lists
        'INCONSISTENT_LIST_HIERARCHY':         'Fix List Indentation Levels',
        'LAYOUT_LIST':                         'List Style Used for Visual Layout',
        # Images
        'IMAGE_MISSING_ALT':                   'Image Missing Alt Text',
        'DECORATIVE_IMAGE_QUESTIONABLE':       'Verify: Image Marked as Decorative',
        'IMAGE_TEXT_CONTENT':                  'Image Appears to Contain Text',
        # Document properties
        'MISSING_TITLE':                       'Add a Document Title (File › Info › Properties)',
        'INVALID_LANGUAGE':                    'Invalid Document Language Setting',
        'MULTIPLE_LANGUAGES':                  'Multiple Language Tags Found',
        'UNTAGGED_LANGUAGE':                   'Possible Untagged Foreign-Language Content',
        # Readability & content
        'NUMERIC_DATE_FORMAT':                 'Use a Spelled-Out Date Format',
        'LONG_SENTENCE':                       'Sentence Is Too Long',
        'FOOTNOTE_USAGE':                      'Consider Moving Footnote Inline',
        'VISUAL_INDICATOR_NO_TEXT':            'Visual Symbol Needs a Text Explanation',
        'MATH_NO_ACCESSIBLE_MARKUP':           'Use Equation Editor for Math Expressions',
    }

    # Required sections based on UWM template (aligned with Quality Matters & AAC&U)
    REQUIRED_SECTIONS = {
        # Basic Course Info
        "Instructor Name & Contact": ["instructor", "professor", "contact information"],
        "Office Hours": ["office hours", "drop-in hours", "student hours"],
        "Course Title & Credits": ["course title", "credits", "credit hours"],
        "Meeting Times & Location": ["meeting time", "location", "class meets", "when and where"],

        # Welcome Statement
        "Welcome Statement": ["welcome"],

        # Course Overview
        "Course Objectives": ["course objectives", "objectives"],
        "Student Learning Outcomes": ["student learning outcomes", "learning outcomes", "outcomes"],
        "Assessment": ["assessment"],
        "Prerequisites": ["prerequisites", "prerequisite", "special skills required"],
        "Course Modality and Format": ["course modality", "modality", "format"],
        "Course Materials": ["course materials", "materials", "textbook", "required materials"],
        "Time Investment": ["time investment", "workload", "time commitment"],

        # Assignments & Grading
        "Assignment and Grading Policies": ["assignment", "grading polic"],
        "Grading Scheme": ["grading scheme", "grade breakdown", "breakdown"],
        "Grading Scale": ["grading scale", "grade scale", "letter grade"],

        # Policies
        "Course Policies": ["course policies", "course policy"],
        "Department Policies": ["department policies", "department policy"],
        "University Policies": ["university policies", "uwm policies", "institutional policies"],

        # Resources
        "Resources": ["resources", "student resources", "support services"],

        # Important Dates
        "Important Dates": ["important dates", "key dates"],

        # Instructor Statement of Support
        "Instructor Statement of Support": ["statement of support", "support statement", "here to help", "want you to succeed"],

        # Calendar
        "Calendar": ["calendar", "schedule", "course schedule"],
    }

    def __init__(self, template_path: str, target_path: str):
        """Initialize checker with template and target syllabus paths"""
        self.template_doc = Document(template_path)
        self.target_doc = Document(target_path)
        self.target_path = target_path
        self.template_sections = self._extract_sections(self.template_doc)
        self.target_sections = self._extract_sections(self.target_doc)
        self.all_paragraphs = self._get_all_paragraphs()  # NEW: Get all paragraphs including from tables
        self.issues = []  # Store all detected issues
        self.standard_headings = self._extract_standard_headings()  # Extract from template

    @staticmethod
    def _wrap_list_items(content_parts: List[str]) -> str:
        """Join content parts, wrapping consecutive <li> items in <ul> tags."""
        result = []
        in_list = False
        for part in content_parts:
            if part.startswith('<li>'):
                if not in_list:
                    result.append('<ul>')
                    in_list = True
                result.append(part)
            else:
                if in_list:
                    result.append('</ul>')
                    in_list = False
                if result:
                    result.append('<br>')
                result.append(part)
        if in_list:
            result.append('</ul>')
        return ''.join(result)

    def _para_to_html(self, para) -> str:
        """Extract paragraph content as HTML, preserving hyperlinks as <a> tags."""
        ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
        parts = []
        for child in para._element:
            tag = child.tag
            if tag == f'{ns}r':
                # Regular text run
                text = ''.join(t.text or '' for t in child.findall(f'{ns}t'))
                parts.append(html.escape(text))
            elif tag == f'{ns}hyperlink':
                # Extract text only from <w:t> elements to avoid lxml itertext() duplication
                link_text = html.escape(''.join(
                    t.text or '' for t in child.findall(f'.//{ns}t')
                ))
                r_id = child.get(qn('r:id'))
                url = None
                if r_id:
                    try:
                        url = para.part.rels[r_id].target_ref
                    except (KeyError, AttributeError):
                        pass
                if url:
                    parts.append(f'<a href="{html.escape(url)}" target="_blank">{link_text}</a>')
                else:
                    parts.append(link_text)
        return ''.join(parts).strip()

    def _extract_standard_headings(self) -> Dict[str, Dict]:
        """Extract standard section headings and their descriptions from the template.

        Returns a dict mapping heading text to {level, description, keywords}
        """
        standard = {}

        # Get all headings from the template
        current_heading = None
        current_content = []

        for para in self.template_doc.paragraphs:
            if para.style.name.startswith('Heading'):
                # Save previous heading with its content
                if current_heading:
                    description = self._wrap_list_items(current_content)
                    # Extract keywords from heading and content
                    heading_lower = current_heading['text'].lower()
                    keywords = [w for w in heading_lower.split() if len(w) > 3]
                    standard[current_heading['text']] = {
                        'level': current_heading['level'],
                        'description': description,
                        'keywords': keywords
                    }

                # Start new heading
                level = int(para.style.name.split()[-1]) if para.style.name[-1].isdigit() else 1
                current_heading = {
                    'text': para.text.strip(),
                    'level': level
                }
                current_content = []
            elif current_heading and para.text.strip():
                is_list = para.style.name == 'List Paragraph'
                para_html = self._para_to_html(para)
                if is_list:
                    para_html = f'<li>{para_html}</li>'
                current_content.append(para_html)

        # Don't forget the last heading
        if current_heading:
            description = self._wrap_list_items(current_content)
            heading_lower = current_heading['text'].lower()
            keywords = [w for w in heading_lower.split() if len(w) > 3]
            standard[current_heading['text']] = {
                'level': current_heading['level'],
                'description': description,
                'keywords': keywords
            }

        return standard

    def _get_section_description(self, section_name: str) -> str:
        """Get the template's description for a section to help users understand what to include.

        Returns a cleaned-up description from the template, or a generic message if not found.
        """
        section_lower = section_name.lower()

        # Try exact match first
        for heading, info in self.standard_headings.items():
            if heading.lower() == section_lower:
                desc = info.get('description', '')
                if desc:
                    # Clean up the description - remove placeholder markers
                    desc = re.sub(r'&lt;[\s-]*-', '', desc)
                    desc = re.sub(r'-[\s-]*&gt;', '', desc)
                    desc = desc.replace('< - -', '').replace('- - >', '')
                    desc = desc.replace('<--', '').replace('-->', '')
                    desc = desc.strip()
                    return desc

        # Try partial match
        for heading, info in self.standard_headings.items():
            heading_lower = heading.lower()
            # Check if significant words match
            section_words = set(w for w in section_lower.split() if len(w) > 3)
            heading_words = set(w for w in heading_lower.split() if len(w) > 3)
            if section_words & heading_words:  # If there's overlap
                desc = info.get('description', '')
                if desc:
                    desc = re.sub(r'&lt;[\s-]*-', '', desc)
                    desc = re.sub(r'-[\s-]*&gt;', '', desc)
                    desc = desc.replace('< - -', '').replace('- - >', '')
                    desc = desc.replace('<--', '').replace('-->', '')
                    desc = desc.strip()
                    return desc

        # Generic fallback
        return f"This section helps students understand {section_name.lower()}."

    def _get_all_paragraphs(self) -> List[ParagraphInfo]:
        """Extract ALL paragraphs from document, including those inside table cells and nested tables"""
        all_paras = []
        global_index = 0

        # Get top-level paragraphs
        for para in self.target_doc.paragraphs:
            para_info = ParagraphInfo(
                paragraph=para,
                index=global_index,
                location=f"Paragraph {global_index + 1}"
            )
            all_paras.append(para_info)
            global_index += 1

        # Get paragraphs from within tables (including nested tables)
        w_ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

        def process_table(table, table_idx, depth=0):
            nonlocal global_index
            for row_idx, row in enumerate(table.rows):
                for cell_idx, cell in enumerate(row.cells):
                    for para in cell.paragraphs:
                        location = f"Table {table_idx + 1}, Row {row_idx + 1}, Cell {cell_idx + 1}"
                        if depth > 0:
                            location = f"Nested {location}"
                        para_info = ParagraphInfo(
                            paragraph=para,
                            index=global_index,
                            location=location,
                            table_idx=table_idx,
                            row_idx=row_idx,
                            cell_idx=cell_idx
                        )
                        all_paras.append(para_info)
                        global_index += 1
                    # Recurse into any nested tables within this cell
                    for nested_tbl_xml in cell._tc.findall(f'./{{{w_ns}}}tbl'):
                        from docx.table import Table as DocxTable
                        nested = DocxTable(nested_tbl_xml, table._parent)
                        process_table(nested, table_idx, depth + 1)

        for table_idx, table in enumerate(self.target_doc.tables):
            process_table(table, table_idx)

        return all_paras

    def _extract_all_text(self) -> str:
        """Extract ALL text from document, including nested tables.

        Word documents can have nested tables (tables inside table cells),
        which aren't captured by iterating through cell.paragraphs.
        This method extracts text from all XML text elements.
        """
        from docx.oxml.ns import qn

        all_texts = []

        # Get text from top-level paragraphs
        for para in self.target_doc.paragraphs:
            text = para.text.strip()
            if text:
                all_texts.append(text)

        # Get text from all tables (including nested tables via XML)
        for table in self.target_doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    # Get direct paragraph text
                    for para in cell.paragraphs:
                        text = para.text.strip()
                        if text:
                            all_texts.append(text)

                    # Check for nested tables using XML
                    nested_tables = cell._tc.findall('.//w:tbl',
                        {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
                    for nested_tbl in nested_tables:
                        # Extract all text nodes from nested table
                        text_nodes = nested_tbl.findall('.//w:t',
                            {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
                        nested_text = ' '.join([t.text for t in text_nodes if t.text])
                        if nested_text.strip():
                            all_texts.append(nested_text.strip())

        return "\n".join(all_texts)

    def _table_to_markdown(self, table) -> str:
        """Convert a docx Table to markdown pipe-delimited format, including nested tables."""
        w_ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        rows_md = []
        for row in table.rows:
            cells_text = []
            for cell in row.cells:
                # Check for nested tables in this cell
                nested_tbls = cell._tc.findall(f'./{{{w_ns}}}tbl')
                if nested_tbls:
                    # Collect paragraph text and nested table text in document order
                    parts = []
                    for child in cell._tc:
                        tag = child.tag.split('}')[-1]
                        if tag == 'p':
                            t = ''.join(
                                node.text for node in child.findall(f'.//{{{w_ns}}}t') if node.text
                            ).strip()
                            if t:
                                parts.append(t)
                        elif tag == 'tbl':
                            from docx.table import Table as DocxTable
                            nested = DocxTable(child, table._parent)
                            parts.append(self._table_to_markdown(nested))
                    cells_text.append(' / '.join(parts))
                else:
                    # Join multi-paragraph cells with " / "
                    cell_text = ' / '.join(
                        p.text.strip() for p in cell.paragraphs if p.text.strip()
                    )
                    cells_text.append(cell_text)
            rows_md.append('| ' + ' | '.join(cells_text) + ' |')
        if len(rows_md) >= 1:
            # Insert separator after first row (header)
            num_cols = len(table.rows[0].cells) if table.rows else 0
            separator = '| ' + ' | '.join(['---'] * num_cols) + ' |'
            rows_md.insert(1, separator)
        return '\n'.join(rows_md)

    def _document_to_structured_text(self) -> str:
        """Extract document text preserving table structure as markdown tables.

        Walks through document body elements in order, rendering paragraphs
        as plain text and tables as markdown pipe-delimited tables. This
        preserves row/column relationships so the LLM can understand context
        like calendar entries and grading breakdowns.
        """
        from docx.table import Table as DocxTable
        w_ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
        parts = []

        for child in self.target_doc.element.body:
            tag = child.tag
            if tag == f'{w_ns}p':
                text = ''.join(
                    node.text for node in child.findall(f'.//{w_ns}t') if node.text
                ).strip()
                if text:
                    parts.append(text)
            elif tag == f'{w_ns}tbl':
                table = DocxTable(child, self.target_doc)
                parts.append(self._table_to_markdown(table))

        return '\n\n'.join(parts)

    def _extract_headings_with_positions(self, full_text: str) -> List[Dict]:
        """Extract headings with their positions in the full text.

        Uses actual Word heading styles from the document plus recognizable heading patterns.
        Returns list of {text: str, position: int, standard_name: str or None}
        """
        headings = []
        text_lower = full_text.lower()

        # Build heading patterns from template's standard headings
        # Each standard heading has keywords we can match against
        heading_patterns = []
        for heading_name, info in self.standard_headings.items():
            keywords = info.get('keywords', [])
            # Add the heading name words as keywords too
            heading_words = [w.lower() for w in heading_name.split() if len(w) > 3]
            all_keywords = list(set(keywords + heading_words))
            if all_keywords:
                heading_patterns.append((all_keywords, heading_name))

        # First, get headings from Word styles (most reliable)
        for section in self.target_sections:
            heading_text = section['text']
            # Find position in full_text
            pos = full_text.find(heading_text)
            if pos >= 0:
                heading_lower = heading_text.lower()
                matched_standard = None

                # Check for exact match with template headings first
                for std_heading in self.standard_headings.keys():
                    if std_heading.lower() == heading_lower or std_heading.lower() in heading_lower:
                        matched_standard = std_heading
                        break

                # If no exact match, try keyword matching
                if not matched_standard:
                    for patterns, standard_name in heading_patterns:
                        if any(p in heading_lower for p in patterns):
                            matched_standard = standard_name
                            break

                headings.append({
                    'text': heading_text,
                    'position': pos,
                    'standard_name': matched_standard,
                    'is_word_heading': True,
                    'is_top_level': section.get('location', '') == 'top-level'
                })

        # Also look for table cell headings (single words followed by content in tables)
        # These are often used as labels like "Instructor", "Materials", etc.
        lines = full_text.split('\n')
        current_pos = 0

        for line in lines:
            line_stripped = line.strip()
            line_lower = line_stripped.lower()

            # Only consider very short lines (< 40 chars) that look like headings
            # and aren't already captured as Word headings
            if line_stripped and len(line_stripped) < 40:
                # Check if this matches a standard heading pattern
                matched_standard = None
                for patterns, standard_name in heading_patterns:
                    if any(p in line_lower for p in patterns):
                        matched_standard = standard_name
                        break

                # Only add if it matches a pattern or ends with colon
                # and isn't already in our headings list
                if matched_standard or (line_stripped.endswith(':') and len(line_stripped) < 30):
                    # Check if we already have this heading
                    already_exists = any(h['text'] == line_stripped for h in headings)
                    if not already_exists:
                        headings.append({
                            'text': line_stripped,
                            'position': current_pos,
                            'standard_name': matched_standard,
                            'is_word_heading': False
                        })

            current_pos += len(line) + 1  # +1 for newline

        # Sort by position
        headings.sort(key=lambda h: h['position'])
        return headings

    def _find_heading_for_content(self, content_position: int, headings: List[Dict]) -> Dict:
        """Find the nearest heading that precedes the content position.

        Priority order:
        1. Top-level Word-styled headings (most reliable — actual document section headings)
        2. Table-cell Word-styled headings (Heading style but inside a table)
        3. Pseudo-headings (short plain-text lines that look like headings)

        This prevents table cell row labels (e.g. "Revision", "Draft") from overriding
        the real section heading that governs the surrounding content.
        """
        nearest_toplevel = None
        nearest_word_heading = None
        nearest_any_heading = None
        for heading in headings:
            if heading['position'] < content_position:
                if nearest_any_heading is None or heading['position'] > nearest_any_heading['position']:
                    nearest_any_heading = heading
                if heading.get('is_word_heading'):
                    if nearest_word_heading is None or heading['position'] > nearest_word_heading['position']:
                        nearest_word_heading = heading
                    if heading.get('is_top_level'):
                        if nearest_toplevel is None or heading['position'] > nearest_toplevel['position']:
                            nearest_toplevel = heading
        return nearest_toplevel or nearest_word_heading or nearest_any_heading

    def _extract_sections(self, doc: Document) -> List[Dict]:
        """Extract headings and their content from document (including tables and nested tables)"""
        sections = []

        # Check top-level paragraphs
        for para in doc.paragraphs:
            if para.style.name.startswith('Heading'):
                level = int(para.style.name.split()[-1])
                sections.append({
                    'text': para.text.strip(),
                    'level': level,
                    'style': para.style.name,
                    'location': 'top-level'
                })

        # Check paragraphs in tables (including nested tables)
        w_ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

        def process_table(table, table_idx, depth=0):
            for row_idx, row in enumerate(table.rows):
                for cell_idx, cell in enumerate(row.cells):
                    for para in cell.paragraphs:
                        if para.style.name.startswith('Heading'):
                            level = int(para.style.name.split()[-1])
                            location = f'Table {table_idx + 1}, Row {row_idx + 1}, Cell {cell_idx + 1}'
                            if depth > 0:
                                location = f'Nested {location}'
                            sections.append({
                                'text': para.text.strip(),
                                'level': level,
                                'style': para.style.name,
                                'location': location
                            })
                    # Recurse into any nested tables within this cell
                    for nested_tbl_xml in cell._tc.findall(f'./{{{w_ns}}}tbl'):
                        from docx.table import Table as DocxTable
                        nested = DocxTable(nested_tbl_xml, table._parent)
                        process_table(nested, table_idx, depth + 1)

        for table_idx, table in enumerate(doc.tables):
            process_table(table, table_idx)

        return sections

    def _find_section(self, section_name: str, search_terms: List[str],
                     sections: List[Dict]) -> bool:
        """Check if a section exists in the document"""
        # Check in extracted heading sections
        for section in sections:
            section_text = section['text'].lower()
            if any(term in section_text for term in search_terms):
                return True

        # Also check in ALL paragraph text (using our comprehensive list)
        for para_info in self.all_paragraphs:
            para_text = para_info.paragraph.text.lower()
            if any(term in para_text for term in search_terms):
                # Only count if it's at the start of a line/paragraph (likely a heading)
                if len(para_info.paragraph.text.strip()) < 100:  # Likely a heading, not body text
                    return True
        return False

    def check_missing_sections(self) -> List[str]:
        """Identify missing required sections"""
        missing = []
        for section, search_terms in self.REQUIRED_SECTIONS.items():
            if not self._find_section(section, search_terms, self.target_sections):
                missing.append(section)
        return missing

    def _detect_content_patterns(self, full_text: str) -> Dict[str, Dict]:
        """
        Use pattern matching to reliably detect content that exists in the document.
        Returns a dict of content_type -> {found: str, heading_ok: bool, current_heading: str} for detected items.
        """
        import re
        detected = {}
        text_lower = full_text.lower()

        # Extract headings with positions for heading detection
        headings = self._extract_headings_with_positions(full_text)

        def is_placeholder(text: str) -> bool:
            """Return True if text contains template placeholder markers, not real content."""
            text_lower = text.lower()
            return any(m in text_lower for m in ['< -', '- >', '<--', '-->', 'canvas insert'])

        def get_heading_info(content_position: int, standard_name: str) -> tuple:
            """Get the current heading and whether it matches the standard.

            Returns (current_heading, heading_ok) tuple.
            """
            heading = self._find_heading_for_content(content_position, headings)
            if heading:
                current_heading = heading['text']
                heading_lower = current_heading.lower()

                # Find the best matching template heading for this standard_name
                best_match = get_standard_heading(standard_name)

                # Check if the current heading matches the expected standard
                heading_ok = False
                best_match_lower = best_match.lower()

                # Exact or partial match with template heading
                if best_match_lower == heading_lower or best_match_lower in heading_lower or heading_lower in best_match_lower:
                    heading_ok = True
                else:
                    # Check keywords from template
                    template_info = self.standard_headings.get(best_match, {})
                    keywords = template_info.get('keywords', [])
                    if any(kw in heading_lower for kw in keywords if len(kw) > 3):
                        heading_ok = True
                    else:
                        # Fallback: check if heading contains key words from standard_name
                        standard_words = [w.lower() for w in standard_name.split() if len(w) > 3]
                        heading_ok = any(word in heading_lower for word in standard_words) if standard_words else False

                # If the content lands under a recognized *required* course-section heading
                # (even a different one than expected), don't suggest renaming it — the
                # content legitimately belongs to that section.  We intentionally exclude
                # template meta-sections (e.g. "Questions, comments…", "Tips from CETL")
                # from this check by only looking at REQUIRED_SECTIONS keys.
                if not heading_ok:
                    for std in self.REQUIRED_SECTIONS.keys():
                        std_lower = std.lower()
                        if std_lower == heading_lower or std_lower in heading_lower or heading_lower in std_lower:
                            heading_ok = True
                            break

                return current_heading, heading_ok
            return None, False

        def get_standard_heading(content_type: str) -> str:
            """Find the template's standard heading that matches this content type."""
            content_lower = content_type.lower()
            content_words = set(w for w in content_lower.split() if len(w) > 3)

            best_match = None
            best_score = 0

            for template_heading in self.standard_headings.keys():
                template_lower = template_heading.lower()
                template_words = set(w for w in template_lower.split() if len(w) > 3)

                # Exact match
                if content_lower == template_lower:
                    return template_heading

                # Calculate word overlap score
                if content_words and template_words:
                    overlap = content_words & template_words
                    score = len(overlap) / max(len(content_words), 1)
                    if score > best_score:
                        best_score = score
                        best_match = template_heading

                # Check if content_type is substring of template heading (but not too short)
                if len(content_lower) > 5 and content_lower in template_lower:
                    return template_heading

            # Only return best match if score is reasonable (at least one word match)
            if best_match and best_score >= 0.5:
                return best_match

            return content_type

        # Instructor Contact Info - look for email addresses, phone numbers.
        # Search all occurrences and prefer those found under an instructor-related heading.
        # This prevents picking up an author/contact email in template preamble sections.
        email_matches = list(re.finditer(r'[\w\.-]+@[\w\.-]+\.\w+', full_text))
        phone_matches = list(re.finditer(r'\d{3}[-.\s]?\d{4}|\(\d{3}\)\s*\d{3}[-.\s]?\d{4}', full_text))
        if email_matches or phone_matches:
            content_type = 'Instructor'
            # Try to find any occurrence that sits under a good (instructor-related) heading
            best_pos = None
            best_heading_ok = False
            all_matches = email_matches + phone_matches
            for m in all_matches:
                _, ok = get_heading_info(m.start(), content_type)
                if ok:
                    best_pos = m.start()
                    best_heading_ok = True
                    break
            # Fall back to first occurrence only if it's under a heading that isn't
            # clearly a non-course section (e.g. template tips / feedback sections).
            if not best_heading_ok:
                first_match = all_matches[0]
                first_heading, _ = get_heading_info(first_match.start(), content_type)
                non_course_keywords = ['question', 'suggest', 'comment', 'improve', 'feedback',
                                       'template', 'tip', 'cetl']
                if first_heading and any(kw in first_heading.lower() for kw in non_course_keywords):
                    # The only hit is in template preamble — don't report a false heading warning.
                    pass  # leave detected['Instructor'] unset
                else:
                    best_pos = first_match.start()
            if best_pos is not None:
                found_parts = []
                if email_matches:
                    found_parts.append(email_matches[0].group())
                if phone_matches:
                    found_parts.append(phone_matches[0].group())
                current_heading, heading_ok = get_heading_info(best_pos, content_type)
                detected[content_type] = {
                    'found': ', '.join(found_parts),
                    'heading_ok': heading_ok,
                    'current_heading': current_heading,
                    'standard_heading': get_standard_heading(content_type)
                }

        # Office Hours - look for "office hours" followed by times
        # Note: Office hours is typically part of the Instructor section in the template
        office_hours_match = re.search(r'office\s*hours[:\s]+([^\n]{10,80})', text_lower)
        if office_hours_match:
            idx = office_hours_match.start()
            content_type = 'Office Hours'
            current_heading, heading_ok = get_heading_info(idx, content_type)
            # Office hours is OK if under Instructor heading or has "office hours" label
            heading_ok = heading_ok or 'office hours' in (current_heading or '').lower() or 'instructor' in (current_heading or '').lower()
            detected[content_type] = {
                'found': office_hours_match.group(1).strip()[:80],
                'heading_ok': heading_ok,
                'current_heading': current_heading,
                'standard_heading': 'Instructor'  # Per template, office hours goes under Instructor
            }
        elif 'office hours' in text_lower or 'student hours' in text_lower:
            idx = text_lower.find('office hours')
            snippet = full_text[max(0, idx):idx+100]
            content_type = 'Office Hours'
            current_heading, heading_ok = get_heading_info(idx, content_type)
            heading_ok = heading_ok or 'office hours' in (current_heading or '').lower() or 'instructor' in (current_heading or '').lower()
            detected[content_type] = {
                'found': snippet.strip()[:80],
                'heading_ok': heading_ok,
                'current_heading': current_heading,
                'standard_heading': 'Instructor'
            }

        # Course Materials - look for "materials", "textbook", book titles
        if any(term in text_lower for term in ['materials', 'textbook', 'required text', 'course pack']):
            for term in ['materials', 'textbook', 'required text']:
                if term in text_lower:
                    idx = text_lower.find(term)
                    snippet = full_text[max(0, idx):idx+150]
                    content_type = 'Course Materials'
                    current_heading, heading_ok = get_heading_info(idx, content_type)
                    detected[content_type] = {
                        'found': snippet.strip()[:100],
                        'heading_ok': heading_ok,
                        'current_heading': current_heading,
                        'standard_heading': get_standard_heading(content_type)
                    }
                    break

        # Grading Scale - look for letter grade patterns.
        # Matches inline format (A=93, A-:90) or table-cell format where the letter
        # grade and number land in separate cells and are joined by whitespace (A 93).
        # Restrict table-format match to 60-99 to avoid false positives with point values.
        grade_pattern = re.search(r'[A-F][+-]?\s*[=:]\s*\d{2}|[A-F][+-]?\s+[6-9]\d\b', full_text)
        if grade_pattern:
            idx = grade_pattern.start()
            snippet = full_text[max(0, idx-20):idx+100]
            content_type = 'Grading Scale'
            current_heading, heading_ok = get_heading_info(idx, content_type)
            detected[content_type] = {
                'found': snippet.strip()[:100],
                'heading_ok': heading_ok,
                'current_heading': current_heading,
                'standard_heading': get_standard_heading(content_type)
            }

        # Grading Scheme / Point System - look for point values
        point_pattern = re.search(r'\d+\s*points?|\d+\s*pts?|worth\s*\d+', text_lower)
        if point_pattern or 'point system' in text_lower:
            if 'point system' in text_lower:
                idx = text_lower.find('point system')
            elif point_pattern:
                idx = point_pattern.start()
            else:
                idx = 0
            snippet = full_text[max(0, idx):idx+150]
            content_type = 'Grading Scheme'
            current_heading, heading_ok = get_heading_info(idx, content_type)
            detected[content_type] = {
                'found': snippet.strip()[:100],
                'heading_ok': heading_ok,
                'current_heading': current_heading,
                'standard_heading': get_standard_heading(content_type)
            }

        # Calendar/Schedule - look for dates and weekly structure
        date_pattern = re.search(r'(january|february|march|april|may|june|july|august|september|october|november|december)\s+\d{1,2}', text_lower)
        week_pattern = re.search(r'week\s*\d|week\s*one|week\s*two', text_lower)
        # Only trigger on concrete date/week evidence, not just the word "schedule" appearing
        # anywhere in the document (that's too broad and causes false positives in tips sections).
        if date_pattern or week_pattern:
            idx = date_pattern.start() if date_pattern else week_pattern.start()
            content_type = 'Calendar'
            current_heading, heading_ok = get_heading_info(idx, content_type)
            detected[content_type] = {
                'found': 'Course schedule/calendar detected',
                'heading_ok': heading_ok,
                'current_heading': current_heading,
                'standard_heading': get_standard_heading(content_type)
            }

        # Assignment Policies - look for policy-related content
        policy_terms = ['late work', 'late assignment', 'revision', 'plagiarism', 'academic dishonesty', 'extension']
        for term in policy_terms:
            if term in text_lower:
                idx = text_lower.find(term)
                snippet = full_text[max(0, idx):idx+100]
                content_type = 'Assignment and Grading Policies'
                current_heading, heading_ok = get_heading_info(idx, content_type)
                detected[content_type] = {
                    'found': snippet.strip()[:100],
                    'heading_ok': heading_ok,
                    'current_heading': current_heading,
                    'standard_heading': get_standard_heading(content_type)
                }
                break

        # Course Objectives/Overview
        if any(term in text_lower for term in ['overview', 'objective', 'this course', 'course is designed']):
            for term in ['overview', 'objective']:
                if term in text_lower:
                    idx = text_lower.find(term)
                    snippet = full_text[max(0, idx):idx+150]
                    content_type = 'Course Objectives'
                    current_heading, heading_ok = get_heading_info(idx, content_type)
                    detected[content_type] = {
                        'found': snippet.strip()[:100],
                        'heading_ok': heading_ok,
                        'current_heading': current_heading,
                        'standard_heading': get_standard_heading(content_type)
                    }
                    break

        # Course Title & Credits - look for credit patterns (avoid "extra credit", "creative", etc.)
        credit_pattern = re.search(r'(?<![a-z])(\d)\s*(credits?|cr\.?|credit\s*hours?)(?![a-z])', text_lower)
        if credit_pattern:
            idx = credit_pattern.start()
            context_before = text_lower[max(0, idx-10):idx].strip()
            if not context_before.endswith('extra') and not context_before.endswith('xtra'):
                snippet = full_text[max(0, idx-50):idx+50]
                content_type = 'Course Title'
                current_heading, heading_ok = get_heading_info(idx, content_type)
                detected[content_type] = {
                    'found': snippet.strip()[:100],
                    'heading_ok': heading_ok,
                    'current_heading': current_heading,
                    'standard_heading': get_standard_heading(content_type)
                }

        # Meeting Times & Location - look for days/times.
        # The day+time pattern requires an actual HH:MM time (not bare colons) so that
        # schedule entries like "Workshop:" or table cells ending in "t:" don't match.
        time_pattern = re.search(
            r'\b(monday|tuesday|wednesday|thursday|friday|m[wf]f?|t[r]?h?|mwf)\b'
            r'[,\s]*(and\s+)?\b(monday|tuesday|wednesday|thursday|friday|m[wf]f?|t[r]?h?|mwf\b)?'
            r'\s*\d{1,2}:\d{2}',
            text_lower)
        if not time_pattern:
            time_pattern = re.search(r'\d{1,2}:\d{2}\s*(am|pm|a\.m\.|p\.m\.)', text_lower)
        if time_pattern:
            idx = time_pattern.start()
            snippet = full_text[max(0, idx):idx+80]
            content_type = 'Meeting Times'
            current_heading, heading_ok = get_heading_info(idx, content_type)
            detected[content_type] = {
                'found': snippet.strip()[:80],
                'heading_ok': heading_ok,
                'current_heading': current_heading,
                'standard_heading': get_standard_heading(content_type)
            }

        # Student Learning Outcomes - look for outcome language
        outcome_terms = ['students will be able', 'learning outcomes', 'course outcomes', 'by the end of this course', 'will learn to', 'will understand']
        for term in outcome_terms:
            if term in text_lower:
                idx = text_lower.find(term)
                snippet = full_text[max(0, idx):idx+150]
                content_type = 'Student Learning Outcomes'
                current_heading, heading_ok = get_heading_info(idx, content_type)
                detected[content_type] = {
                    'found': snippet.strip()[:100],
                    'heading_ok': heading_ok,
                    'current_heading': current_heading,
                    'standard_heading': get_standard_heading(content_type)
                }
                break

        # Prerequisites - look for prerequisite mentions
        prereq_terms = ['prerequisite', 'prereq', 'prior course', 'required course', 'must have completed']
        for term in prereq_terms:
            if term in text_lower:
                idx = text_lower.find(term)
                snippet = full_text[max(0, idx):idx+150]
                content_type = 'Prerequisites'
                current_heading, heading_ok = get_heading_info(idx, content_type)
                detected[content_type] = {
                    'found': snippet.strip()[:100],
                    'heading_ok': heading_ok,
                    'current_heading': current_heading,
                    'standard_heading': get_standard_heading(content_type)
                }
                break

        # University Policies - look for policy references
        uni_policy_terms = ['academic integrity', 'academic honesty', 'academic misconduct', 'title ix', 'ada', 'disability', 'accommodation', 'ferpa']
        for term in uni_policy_terms:
            if term in text_lower:
                idx = text_lower.find(term)
                snippet = full_text[max(0, idx):idx+150]
                content_type = 'University Policies'
                current_heading, heading_ok = get_heading_info(idx, content_type)
                detected[content_type] = {
                    'found': snippet.strip()[:100],
                    'heading_ok': heading_ok,
                    'current_heading': current_heading,
                    'standard_heading': get_standard_heading(content_type)
                }
                break

        # Department Policies - look for section heading or keyword
        for term in ['department polic', 'department policy']:
            if term in text_lower:
                idx = text_lower.find(term)
                snippet = full_text[max(0, idx):idx+150]
                content_type = 'Department Policies'
                current_heading, heading_ok = get_heading_info(idx, content_type)
                detected[content_type] = {
                    'found': snippet.strip()[:100],
                    'heading_ok': heading_ok,
                    'current_heading': current_heading,
                    'standard_heading': get_standard_heading(content_type)
                }
                break

        # Important Dates - look for section heading or keyword
        for term in ['important dates', 'key dates']:
            if term in text_lower:
                idx = text_lower.find(term)
                snippet = full_text[max(0, idx):idx+150]
                content_type = 'Important Dates'
                current_heading, heading_ok = get_heading_info(idx, content_type)
                detected[content_type] = {
                    'found': snippet.strip()[:100],
                    'heading_ok': heading_ok,
                    'current_heading': current_heading,
                    'standard_heading': get_standard_heading(content_type)
                }
                break

        # Instructor Statement of Support - look for support language
        for term in ['statement of support', 'here to help', 'want you to succeed',
                     'do not hesitate to reach out', 'help you succeed', 'instructor statement']:
            if term in text_lower:
                idx = text_lower.find(term)
                snippet = full_text[max(0, idx):idx+150]
                content_type = 'Instructor Statement of Support'
                current_heading, heading_ok = get_heading_info(idx, content_type)
                detected[content_type] = {
                    'found': snippet.strip()[:100],
                    'heading_ok': heading_ok,
                    'current_heading': current_heading,
                    'standard_heading': get_standard_heading(content_type)
                }
                break

        # Suppress heading-standardization false positives for placeholder text:
        # if the detected snippet is template placeholder content (not real student content),
        # don't flag the nearby heading as wrong.
        for info in detected.values():
            if not info.get('heading_ok') and is_placeholder(info.get('found', '')):
                info['heading_ok'] = True

        return detected

    def analyze_sections_with_llm(self) -> Dict:
        """
        Hybrid approach: Use pattern matching first for reliable detection,
        then LLM for nuanced content analysis.
        """
        # Extract full document content (including nested tables)
        full_text = self._extract_all_text()

        # STEP 1: Pattern-based detection (reliable baseline)
        detected = self._detect_content_patterns(full_text)

        # Build present list from pattern detection
        present = []
        for content_type, info in detected.items():
            present.append({
                'content_type': content_type,
                'found': info['found'],
                'heading_ok': info.get('heading_ok', True),
                'current_heading': info.get('current_heading'),
                'standard_heading': info.get('standard_heading', content_type)
            })

        # Content types that need LLM analysis (subjective/nuanced)
        llm_content_types = [
            'Welcome Statement',
            'Student Learning Outcomes',
            'Prerequisites',
            'Resources',
            'University Policies',
            'Course Policies',
            'Meeting Times & Location',
            'Course Modality and Format',
            'Time Investment',
            'Department Policies',
            'Important Dates',
            'Instructor Statement of Support',
            'Assessment'
        ]

        # Items already detected by pattern matching
        detected_types = set(detected.keys())

        # Get required sections list
        required_sections = list(self.REQUIRED_SECTIONS.keys())

        # Truncate for LLM if needed
        llm_text = full_text[:12000] if len(full_text) > 12000 else full_text

        # Get API key
        api_key = os.environ.get('OPENAI_API_KEY')
        if not api_key:
            # Use pattern-detected results, plus basic keyword matching for the rest
            missing_sections = self.check_missing_sections()

            # Start with pattern-detected items
            result_present = list(present)  # From our pattern detection
            pattern_detected_types = detected_types

            # Add any other sections found via keyword matching (that weren't already detected)
            for s in required_sections:
                if s not in missing_sections and s not in pattern_detected_types:
                    result_present.append({
                        'content_type': s,
                        'found': 'Detected via keyword matching',
                        'heading_ok': True,
                        'standard_heading': s
                    })

            # Determine what's actually missing (not detected by patterns AND not found by keyword matching)
            actually_missing = [s for s in required_sections if s in missing_sections and s not in pattern_detected_types]

            return {
                'present': result_present,
                'missing': [{'content_type': s,
                            'description': self._get_section_description(s),
                            'recommendation': f'Add a "{s}" section to your syllabus'}
                           for s in actually_missing],
                'error': 'OPENAI_API_KEY not set - using pattern matching + basic keyword matching'
            }

        # STEP 2: LLM analysis for nuanced content (only items not detected by patterns)
        remaining_to_check = [ct for ct in llm_content_types if ct not in detected_types]
        llm_present = []
        llm_missing = []

        try:
            client = OpenAI(api_key=api_key)

            # Only ask LLM about items not already detected
            if remaining_to_check:
                content_descriptions = {
                    'Welcome Statement': 'welcoming language to students, phrases like "welcome to" or expressing care for student success',
                    'Student Learning Outcomes': 'what students will learn or be able to do by end of course',
                    'Prerequisites': 'required prior courses, skills, or knowledge',
                    'Resources': 'support services, tutoring, disability services, counseling',
                    'University Policies': 'institutional policies like academic integrity, Title IX',
                    'Course Policies': 'attendance requirements, participation expectations, classroom behavior',
                    'Meeting Times & Location': 'when and where the class meets (days, times, room)',
                    'Course Modality and Format': 'whether online, in-person, hybrid; how class is conducted',
                    'Time Investment': 'expected hours per week, workload expectations',
                    'Department Policies': 'department-specific policies',
                    'Important Dates': 'key deadlines, exam dates, holidays',
                    'Instructor Statement of Support': 'statement about caring for student success',
                    'Assessment': 'how students will be evaluated, types of assessments'
                }

                items_to_find = "\n".join([f"- {ct}: {content_descriptions.get(ct, ct)}"
                                          for ct in remaining_to_check])

                prompt = f"""Search this syllabus for the following content. Quote what you find.

SYLLABUS:
{llm_text}

FIND THESE (quote actual text if found):
{items_to_find}

Return JSON:
{{"found": [{{"type": "X", "quote": "actual text from document"}}], "not_found": ["Y", "Z"]}}

Be generous - if content relates to the topic, include it in "found"."""

                response = client.chat.completions.create(
                    model="gpt-4o",
                    messages=[
                        {"role": "system", "content": "Search the document and quote what you find. Be generous in matching."},
                        {"role": "user", "content": prompt}
                    ],
                    max_tokens=2000,
                    temperature=0.1
                )

                response_text = response.choices[0].message.content.strip()
                if response_text.startswith('```'):
                    response_text = response_text.split('```')[1]
                    if response_text.startswith('json'):
                        response_text = response_text[4:]
                    response_text = response_text.strip()

                import json
                llm_result = json.loads(response_text)

                # Process LLM findings
                for item in llm_result.get('found', []):
                    llm_present.append({
                        'content_type': item.get('type', ''),
                        'found': item.get('quote', 'Found')[:100],
                        'heading_ok': True,
                        'standard_heading': item.get('type', '')
                    })

                for item in llm_result.get('not_found', []):
                    if isinstance(item, str):
                        llm_missing.append({
                            'content_type': item,
                            'description': self._get_section_description(item),
                            'recommendation': f'Add a "{item}" section to your syllabus'
                        })

        except Exception as e:
            # If LLM fails, mark remaining items as needing review
            for ct in remaining_to_check:
                llm_missing.append({
                    'content_type': ct,
                    'description': f'Could not analyze (LLM error: {str(e)[:50]})',
                    'recommendation': f'Please review if {ct} is included'
                })

        # STEP 3: Merge results - pattern detection takes precedence
        all_present = present + llm_present

        # Determine what's truly missing (not in pattern detection OR LLM found)
        found_types = {p['content_type'] for p in all_present}
        missing = []
        for item in llm_missing:
            if item['content_type'] not in found_types:
                missing.append(item)

        # Also check required sections not covered by either
        all_checked = found_types | {m['content_type'] for m in missing}
        for section in required_sections:
            if section not in all_checked and section not in detected_types:
                # Check if this is a pattern-detectable item that just wasn't found
                if section not in llm_content_types:
                    # Fall back to keyword matching before declaring missing — prevents false
                    # positives when pattern detection uses a different key name (e.g.
                    # 'Instructor' vs 'Instructor Name & Contact') or the content lives in
                    # a nested table that pattern detection already handled via XML.
                    search_terms = self.REQUIRED_SECTIONS.get(section, [])
                    if self._find_section(section, search_terms, self.target_sections):
                        all_present.append({
                            'content_type': section,
                            'found': 'Detected via keyword matching',
                            'heading_ok': True,
                            'standard_heading': section
                        })
                    else:
                        missing.append({
                            'content_type': section,
                            'description': self._get_section_description(section),
                            'recommendation': f'Add a "{section}" section to your syllabus'
                        })

        return {
            'present': all_present,
            'missing': missing
        }

    def check_unstyled_headings(self) -> List[AccessibilityIssue]:
        """Detect text that looks like headings but doesn't use heading styles"""
        unstyled_headings = []

        # Track last heading level to suggest appropriate level
        last_heading_level = 0
        last_heading_location = None

        for para_info in self.all_paragraphs:
            para = para_info.paragraph

            # Track actual heading styles we encounter
            if para.style.name.startswith('Heading'):
                last_heading_level = int(para.style.name.split()[-1])
                last_heading_location = para_info.location
                continue

            text = para.text.strip()

            # Skip empty paragraphs or very long paragraphs
            if not text or len(text) > 150:
                continue

            # FILTER OUT FALSE POSITIVES

            # Skip list items (likely just emphasized list text, not headings)
            if para.style.name and 'List' in para.style.name:
                continue

            # Skip if starts with bullet/number (manual list items)
            if text and text[0] in ['•', '·', '-', '*', '○', '■']:
                continue

            # Skip if looks like a numbered item
            if len(text) > 0 and text[0].isdigit() and '.' in text[:5]:
                continue

            # Skip very short text (1-2 chars) - likely labels, not headings
            if len(text) <= 2:
                continue

            # Skip if only PART of the text is bold (inline emphasis, not heading)
            if para.runs:
                bold_chars = sum(len(run.text) for run in para.runs if run.bold)
                total_chars = len(text)
                bold_ratio = bold_chars / total_chars if total_chars > 0 else 0

                # If less than 80% bold, it's probably inline emphasis
                if 0 < bold_ratio < 0.8:
                    continue

            # Skip if ends with certain punctuation (likely emphasized text, not heading)
            if text.endswith(('.', '!', ',', ';', ':')):
                continue

            # Skip text that starts with deadline/assignment indicators (not headings)
            # Examples: "Due by January 15", "Submit your assignment", "Complete the reading"
            text_lower = text.lower()
            deadline_prefixes = [
                'due by', 'due on', 'due date', 'due:', 'submit', 'turn in',
                'complete', 'finish', 'read', 'prepare', 'bring to class'
            ]
            if any(text_lower.startswith(prefix) for prefix in deadline_prefixes):
                continue

            # Skip "Label: Description" patterns (activity labels, not headings)
            # Examples: "In-class writing: Business packet", "Reading: Chapter 5", "Homework: Read chapter 5"
            if ':' in text:
                # Check if this is a label:description pattern (not a section heading)
                parts = text.split(':', 1)
                if len(parts) == 2:
                    label = parts[0].strip()
                    description = parts[1].strip()

                    # Common activity/assignment labels (not major section headings)
                    activity_labels = [
                        'in-class', 'homework', 'assignment', 'due', 'reading', 'activity',
                        'discussion', 'group work', 'lab', 'quiz', 'exam', 'project',
                        'presentation', 'review', 'workshop', 'lecture', 'tutorial'
                    ]

                    # If the label is short and matches activity patterns, skip it
                    label_lower = label.lower()
                    if len(label) < 30 and description:  # Has both label and description
                        # Check if it's an activity label
                        if any(activity_word in label_lower for activity_word in activity_labels):
                            continue

                        # Also skip if the label is short and description contains specific items
                        # (e.g., "Reading: Chapter 5", "Due: Assignment 3")
                        if len(label.split()) <= 4 and len(description.split()) <= 10:
                            # This is likely a label:description, not a major section heading
                            continue

            # NOW CHECK FOR HEADING CHARACTERISTICS
            is_likely_heading = False
            reasons = []

            # Check for bold text (full paragraph bold)
            if para.runs:
                bold_chars = sum(len(run.text) for run in para.runs if run.bold)
                total_chars = len(text)
                if total_chars > 0 and bold_chars / total_chars >= 0.8:
                    # Additional check: is it standalone and short?
                    words = text.split()
                    if len(words) <= 10:  # Headings are typically short
                        is_likely_heading = True
                        reasons.append("bold text")

            # Check for all caps (but not single words like "I" or "A")
            if text.isupper() and len(text) > 3:
                is_likely_heading = True
                reasons.append("all caps")

            # Check for title case + no ending punctuation + short
            if len(text) < 100:
                words = text.split()
                if 2 <= len(words) <= 8:  # Multi-word but not too long
                    # Check for larger font size
                    has_large_font = False
                    for run in para.runs:
                        if run.font.size and run.font.size.pt > 12:
                            has_large_font = True
                            reasons.append("large font")
                            break

                    if has_large_font or (bold_chars > 0 and len(words) <= 6):
                        is_likely_heading = True

            if is_likely_heading and reasons:
                # Determine recommended heading level based on context
                recommended_level = self._recommend_heading_level(
                    para_info, text, last_heading_level, last_heading_location
                )

                issue = AccessibilityIssue(
                    issue_type=f"SHOULD BE HEADING {recommended_level}",
                    description=f"Text appears to be a heading ({', '.join(reasons)}) but uses manual formatting instead of a Heading style. Apply Heading {recommended_level} from the Styles panel (Home tab) so screen readers and navigation tools can recognize it: \"{text[:60]}...\"",
                    location=para_info.location,
                    para_info=para_info
                )
                unstyled_headings.append(issue)

        return unstyled_headings

    def _recommend_heading_level(self, para_info: ParagraphInfo, text: str,
                                  last_heading_level: int, last_heading_location: str) -> int:
        """Recommend appropriate heading level based on context"""

        # Check if text looks like a major section title
        major_section_keywords = [
            'overview', 'introduction', 'objectives', 'materials', 'assignments',
            'grading', 'policies', 'schedule', 'calendar', 'resources'
        ]

        text_lower = text.lower()
        is_major_section = any(keyword in text_lower for keyword in major_section_keywords)

        # If in a table (likely nested content)
        if para_info.table_idx is not None:
            # If it's a recognized major section, could be H1
            if is_major_section and len(text.split()) <= 3:
                return 1
            # Otherwise likely a subsection
            return 2

        # Top-level paragraph
        else:
            # First heading in document or major section
            if last_heading_level == 0 or is_major_section:
                return 1
            # Following an H1, likely also H1 or H2
            elif last_heading_level == 1:
                # If it looks like a subsection, suggest H2
                if len(text.split()) > 3:
                    return 2
                return 1
            # Following an H2 or deeper
            else:
                return last_heading_level

        return 1  # Default to H1 if unsure

    def _table_has_header_row(self, table) -> bool:
        """Helper: Check if a table has a proper header row"""
        if not table.rows:
            return False

        first_row = table.rows[0]
        num_cells = len(first_row.cells)

        # FIRST: Check if the first row has the official "Repeat Header Row" property
        # This is the most reliable indicator that Word recognizes this as a header
        try:
            tr = first_row._element
            trPr = tr.find(qn('w:trPr'))
            if trPr is not None:
                tblHeader = trPr.find(qn('w:tblHeader'))
                if tblHeader is not None:
                    # Official header row designation found!
                    return True
        except:
            pass

        # FALLBACK: Check for visual indicators of header rows
        bold_count = 0
        shaded_count = 0
        large_font_count = 0

        for cell in first_row.cells:
            # Check for bold text
            has_bold = False
            has_large_font = False

            for para in cell.paragraphs:
                for run in para.runs:
                    # Check direct formatting
                    if run.bold == True:
                        has_bold = True
                    # Also check if bold is inherited from style
                    elif run.bold is None and run.style and hasattr(run.style, 'font'):
                        try:
                            if run.style.font.bold == True:
                                has_bold = True
                        except:
                            pass

                    # Check the XML directly for bold
                    try:
                        rPr = run._element.find(qn('w:rPr'))
                        if rPr is not None:
                            b_elem = rPr.find(qn('w:b'))
                            if b_elem is not None:
                                has_bold = True
                    except:
                        pass

                    # Check for larger font (>11pt suggests header)
                    if run.font.size and run.font.size.pt > 11:
                        has_large_font = True

            if has_bold:
                bold_count += 1
            if has_large_font:
                large_font_count += 1

            # Check for cell shading (background color)
            # Access the cell's XML to check for shading
            try:
                tc = cell._element
                tcPr = tc.get_or_add_tcPr()
                shd = tcPr.find(qn('w:shd'), tcPr.nsmap)
                if shd is not None:
                    fill = shd.get(qn('w:fill'))
                    # If there's a fill color (not 'auto'), likely a header
                    if fill and fill != 'auto' and fill != 'FFFFFF':
                        shaded_count += 1
            except:
                pass  # If we can't check shading, skip it

        # Header row if most cells have ANY of these indicators:
        # - Bold text (50%+)
        # - Shading (50%+)
        # - Large font (50%+)
        if bold_count >= num_cells * 0.5:
            return True
        if shaded_count >= num_cells * 0.5:
            return True
        if large_font_count >= num_cells * 0.5:
            return True

        # FALLBACK: Content-based heuristic for tables where formatting isn't detected
        # (common issue with Mac Word where styles don't save to XML)
        # If the first row has concise text in all/most cells AND the table has multiple rows,
        # it's very likely a header row
        if len(table.rows) > 2:  # Only for multi-row tables (not layout tables)
            filled_cells = 0
            concise_cells = 0
            for cell in first_row.cells:
                text = cell.text.strip()
                if text:
                    filled_cells += 1
                    # Headers are typically concise (< 50 chars)
                    if len(text) < 50:
                        concise_cells += 1

            # If most cells are filled with concise text, likely a header row
            if filled_cells >= num_cells * 0.75 and concise_cells >= num_cells * 0.75:
                return True

        return False

    def check_table_headers(self) -> List[AccessibilityIssue]:
        """Check if tables have proper header rows"""
        table_issues = []

        for i, table in enumerate(self.target_doc.tables):
            has_header = self._table_has_header_row(table)

            if not has_header and len(table.rows) > 1:
                # Get preview of table content
                preview = " | ".join([cell.text.strip()[:20] for cell in table.rows[0].cells])

                issue = AccessibilityIssue(
                    issue_type="TABLE_NO_HEADER",
                    description=f"Table {i+1} ({preview}) — no header row designated. Select the first row, then check 'Repeat Header Rows' under Table Layout > Properties so assistive technology knows which row is the header.",
                    location=f"Table {i+1}",
                    table_index=i
                )
                table_issues.append(issue)

        return table_issues

    def check_layout_tables(self) -> List[AccessibilityIssue]:
        """Identify tables used for layout rather than tabular data"""
        layout_tables = []

        for i, table in enumerate(self.target_doc.tables):
            is_layout_table = False
            reasons = []

            rows = len(table.rows)
            cols = len(table.columns)

            # Check if table has headers (data tables with headers are generally OK)
            has_header = self._table_has_header_row(table)

            # Check for single column/row (almost always layout)
            if cols == 1 or rows == 1:
                is_layout_table = True
                reasons.append("single column/row")

            # Check for cells with excessive text
            cell_lengths = []
            empty_cells = 0
            very_long_cells = 0

            for row in table.rows:
                for cell in row.cells:
                    cell_text = cell.text.strip()
                    cell_len = len(cell_text)
                    cell_lengths.append(cell_len)

                    if cell_len == 0:
                        empty_cells += 1
                    elif cell_len > 500:
                        very_long_cells += 1

            total_cells = len(cell_lengths)

            # Only flag empty cells if table lacks headers (data tables can have sparse data)
            if not has_header and total_cells > 0 and empty_cells / total_cells > 0.3:
                is_layout_table = True
                reasons.append(f"{empty_cells}/{total_cells} empty cells")

            # If cells have very long text, likely layout (not data)
            # But only if no headers (tables with headers might legitimately have long cells)
            if not has_header and very_long_cells > 0:
                is_layout_table = True
                reasons.append(f"{very_long_cells} cells with >500 chars")

            # Check for inconsistent cell lengths only if no headers
            # (data tables can have varying cell content)
            if not has_header and len(cell_lengths) > 4:
                import statistics
                try:
                    stdev = statistics.stdev(cell_lengths)
                    mean = statistics.mean(cell_lengths)
                    if mean > 0 and stdev / mean > 2:  # High variance
                        is_layout_table = True
                        reasons.append("inconsistent cell content lengths")
                except:
                    pass

            # Check for two-column "label: value" pattern
            # BUT: Don't flag if the table has proper headers (first row bold)
            if cols == 2 and rows > 2:
                # Use helper to check if table has header row
                has_header_row = self._table_has_header_row(table)

                # Only check for label-value pattern if NO header row
                if not has_header_row:
                    label_value_pattern = 0
                    for row in table.rows:
                        if len(row.cells) >= 2:
                            cell1_len = len(row.cells[0].text.strip())
                            cell2_len = len(row.cells[1].text.strip())
                            if 0 < cell1_len < 50 and cell2_len > cell1_len:
                                label_value_pattern += 1

                    if label_value_pattern / rows > 0.6:
                        is_layout_table = True
                        reasons.append("label-value pattern detected")

            if is_layout_table:
                issue = AccessibilityIssue(
                    issue_type="LAYOUT_TABLE",
                    description=f"Table {i+1} ({rows} rows × {cols} cols) — used for layout, not data ({'; '.join(reasons)})",
                    location=f"Table {i+1}",
                    table_index=i
                )
                layout_tables.append(issue)

        return layout_tables

    def check_heading_hierarchy(self) -> List[AccessibilityIssue]:
        """Check for improper heading hierarchy (e.g., too many H1s, H1s under H1s)"""
        hierarchy_issues = []

        # Track heading sequence to detect hierarchy problems
        last_heading_level = 0
        h1_count = 0
        first_h1_index = None

        for idx, para_info in enumerate(self.all_paragraphs):
            para = para_info.paragraph
            if para.style.name.startswith('Heading'):
                level = int(para.style.name.split()[-1])
                text = para.text.strip()

                # Check for ALL CAPS headings
                if text.isupper() and len(text) > 3:  # Ignore very short text like "FAQ"
                    issue = AccessibilityIssue(
                        issue_type="ALL_CAPS_HEADING",
                        description=f"Heading '{text}' is in ALL CAPS. Use title case for better readability and accessibility.",
                        location=para_info.location,
                        para_info=para_info
                    )
                    hierarchy_issues.append(issue)

                # Check for overly long headings
                if len(text) > 100:
                    issue = AccessibilityIssue(
                        issue_type="LONG_HEADING",
                        description=f"Heading '{text[:50]}...' is {len(text)} characters long. Keep headings concise (under 100 characters).",
                        location=para_info.location,
                        para_info=para_info
                    )
                    hierarchy_issues.append(issue)

                # Count H1s and track first H1
                if level == 1:
                    h1_count += 1
                    if first_h1_index is None:
                        first_h1_index = idx

                    # If we've seen an H1 before and now see another H1, check context
                    # In table cells, multiple H1s might indicate subsections that should be H2
                    if h1_count > 3 and para_info.table_idx is not None:
                        # This H1 is inside a table and there are many H1s
                        # It's likely a subsection that should be H2
                        issue = AccessibilityIssue(
                            issue_type="SHOULD BE HEADING 2",
                            description=f"'{text}' uses Heading 1 but appears to be a subsection (should be Heading 2)",
                            location=para_info.location,
                            para_info=para_info
                        )
                        hierarchy_issues.append(issue)

                # Check for skipped levels (e.g., H1 -> H3, skipping H2)
                if last_heading_level > 0 and level > last_heading_level + 1:
                    # It should be one level below the previous heading
                    recommended_level = last_heading_level + 1
                    issue = AccessibilityIssue(
                        issue_type=f"SHOULD BE HEADING {recommended_level}",
                        description=f"'{text}' jumps from Heading {last_heading_level} to Heading {level} (should be Heading {recommended_level})",
                        location=para_info.location,
                        para_info=para_info
                    )
                    hierarchy_issues.append(issue)

                last_heading_level = level

        # Check if H1 appears near the top (within first 10 paragraphs)
        if first_h1_index is not None and first_h1_index > 10:
            issue = AccessibilityIssue(
                issue_type="H1_NOT_AT_TOP",
                description=f"First Heading 1 appears at paragraph {first_h1_index + 1}. Move it or add a Heading 1 near the top so readers and screen readers immediately recognize the document title.",
                location="Document start",
                para_info=None
            )
            hierarchy_issues.append(issue)

        # Check for multiple H1s (should typically only have one main title)
        if h1_count > 1:
            issue = AccessibilityIssue(
                issue_type="MULTIPLE_H1",
                description=f"Document has {h1_count} Heading 1 elements. Keep one Heading 1 as the main title and convert the rest to Heading 2 (for major sections) or lower.",
                location="Throughout document",
                para_info=None
            )
            hierarchy_issues.append(issue)

        return hierarchy_issues

    def check_heading_structure(self) -> Dict[str, any]:
        """Analyze heading usage and structure"""
        issues = []
        heading_counts = {}

        # Count headings by level (now checking all paragraphs)
        for para_info in self.all_paragraphs:
            para = para_info.paragraph
            if para.style.name.startswith('Heading'):
                level = para.style.name
                heading_counts[level] = heading_counts.get(level, 0) + 1

        # Check for proper hierarchy
        if not heading_counts:
            issues.append("CRITICAL: No heading styles used. Document should use Heading 1, Heading 2, etc.")
        elif len(heading_counts) < 2:
            issues.append("WARNING: Limited heading hierarchy. Use multiple heading levels for better structure.")

        # Check if Heading 1 exists
        if 'Heading 1' not in heading_counts:
            issues.append("WARNING: No Heading 1 found. Major sections should use Heading 1.")

        # Check for proper nesting
        levels_used = sorted([int(h.split()[-1]) for h in heading_counts.keys()])
        if levels_used and levels_used[0] != 1:
            issues.append("WARNING: Document should start with Heading 1, not Heading {}.".format(levels_used[0]))

        # Warn about excessive H1s only if there's evidence of hierarchy problems
        if 'Heading 1' in heading_counts:
            h1_count = heading_counts['Heading 1']
            h2_count = heading_counts.get('Heading 2', 0)

            # Only warn if:
            # 1. Many H1s (>10) with very few H2s (suggests bad hierarchy)
            # 2. Or many H1s (>15) total (likely overuse regardless)
            if h1_count > 10 and h2_count < 3:
                issues.append(
                    f"WARNING: Document has {h1_count} Heading 1s but only {h2_count} Heading 2s. "
                    "Many H1s should likely be H2 (subsections)."
                )
            elif h1_count > 15:
                issues.append(
                    f"WARNING: Document has {h1_count} Heading 1s. This is excessive - "
                    "consider using Heading 2 for subsections."
                )

        return {
            'counts': heading_counts,
            'issues': issues
        }

    def check_table_usage(self) -> Dict[str, any]:
        """Analyze table usage for accessibility"""
        issues = []
        tables = self.target_doc.tables

        if len(tables) == 0:
            return {'count': 0, 'issues': []}

        complex_tables = []
        for i, table in enumerate(tables):
            for row in table.rows:
                for cell in row.cells:
                    cell_text = cell.text.strip()
                    if len(cell_text) > 500:
                        complex_tables.append(i)
                        issues.append(
                            f"WARNING: Table {i+1} contains cells with excessive text ({len(cell_text)} chars). "
                            "Consider using regular paragraphs with headings instead."
                        )
                        break

        if len(tables) > 5:  # Increased threshold from 2 to 5
            issues.append(
                f"WARNING: Document contains {len(tables)} tables. "
                "Excessive table usage can harm accessibility. Consider using headings and lists."
            )

        # NOTE: Large table check removed - most syllabi have large assignment tables,
        # and table size alone doesn't indicate layout misuse. Will add LLM-based
        # semantic check for layout tables in the future.

        return {
            'count': len(tables),
            'complex_tables': complex_tables,
            'issues': issues
        }

    def check_hyperlinks(self) -> Dict[str, any]:
        """Check for accessible hyperlink practices"""
        issues = []
        hyperlinks = []

        # Check all paragraphs (including in tables)
        for para_info in self.all_paragraphs:
            para = para_info.paragraph
            for run in para.runs:
                if run.font.underline or run.font.color.rgb:
                    if re.search(r'https?://', run.text):
                        hyperlinks.append(run.text)
                        if run.text.startswith('http'):
                            issues.append(
                                f"WARNING: Hyperlink displays URL directly: '{run.text[:50]}...'. "
                                "Use descriptive link text instead."
                            )

        return {
            'count': len(hyperlinks),
            'issues': issues
        }

    def check_images(self) -> Dict[str, any]:
        """Check for image alt text"""
        issues = []
        images_without_alt = []

        for rel in self.target_doc.part.rels.values():
            if "image" in rel.target_ref:
                images_without_alt.append(rel.target_ref)

        if images_without_alt:
            issues.append(
                f"INFO: Document contains {len(images_without_alt)} image(s). "
                "Verify all images have descriptive alt text."
            )

        return {
            'count': len(images_without_alt),
            'issues': issues
        }

    def check_text_formatting(self) -> Dict[str, any]:
        """Check for proper use of styles vs. manual formatting"""
        issues = []
        manual_formatting_count = 0

        for para_info in self.all_paragraphs:
            para = para_info.paragraph
            text = para.text.strip()

            # Only check Normal styled paragraphs
            if para.style.name == 'Normal' and para.runs and text:
                # Skip very short text (likely labels, not headings)
                if len(text) <= 3:
                    continue

                # Skip text that ends with sentence punctuation (likely inline emphasis)
                if text.endswith(('.', '!', '?', ',', ';', ':')):
                    continue

                # Check if MOST of the paragraph is bold (>80% = likely a heading, not inline emphasis)
                bold_chars = sum(len(run.text) for run in para.runs if run.bold)
                total_chars = len(text)

                if total_chars > 0:
                    bold_ratio = bold_chars / total_chars

                    # Only flag if most of the text is bold AND it's relatively short (heading-like)
                    if bold_ratio > 0.8 and len(text) < 100:
                        # Additional check: skip if it's in the middle of a sentence
                        # (look for lowercase start or mid-sentence indicators)
                        if text and text[0].isupper():
                            manual_formatting_count += 1

        if manual_formatting_count > 5:
            issues.append(
                f"WARNING: Found {manual_formatting_count} instances of manual bold formatting that look like headings. "
                "Use Heading styles instead of bold text for section headers."
            )

        return {
            'manual_formatting_count': manual_formatting_count,
            'issues': issues
        }

    def check_list_usage(self) -> Dict[str, any]:
        """Check for proper list formatting"""
        issues = []
        list_count = 0
        manual_list_count = 0

        for para_info in self.all_paragraphs:
            para = para_info.paragraph
            if 'List' in para.style.name:
                list_count += 1
            elif para.style.name == 'Normal':
                text = para.text.strip()
                if text and text[0] in ['-', '*', '•', '·']:
                    manual_list_count += 1

        if manual_list_count > 3:
            issues.append(
                f"WARNING: Found {manual_list_count} manually formatted list items. "
                "Use built-in list styles for better accessibility."
            )

        return {
            'list_count': list_count,
            'manual_list_count': manual_list_count,
            'issues': issues
        }

    def check_manual_alignment(self) -> List[AccessibilityIssue]:
        """Check for manual alignment using tabs and multiple spaces (pseudo-tables)"""
        issues = []

        # Get list of paragraphs already flagged as heading issues to avoid double-flagging
        heading_issue_paras = set()
        for issue in self.issues:
            if issue.issue_type.startswith("SHOULD BE HEADING") and issue.para_info:
                heading_issue_paras.add(id(issue.para_info.paragraph))

        for para_info in self.all_paragraphs:
            para = para_info.paragraph
            text = para.text

            # Skip if already flagged as a heading issue
            if id(para) in heading_issue_paras:
                continue

            # Skip very short text (unlikely to be pseudo-tables)
            if len(text.strip()) < 10:
                continue

            is_pseudo_table = False
            reasons = []

            # Check for tabs used for columnar alignment
            # Only flag if there's substantive content (5+ chars) on both sides of a tab
            if '\t' in text:
                tab_parts = text.split('\t')
                for i in range(len(tab_parts) - 1):
                    left_side = tab_parts[i].strip()
                    right_side = tab_parts[i + 1].strip()
                    # If both sides have content, it's likely a pseudo-table
                    if len(left_side) >= 5 and len(right_side) >= 3:
                        is_pseudo_table = True
                        reasons.append("tabs for columnar alignment")
                        break

            # Check for multiple spaces used for alignment
            # Only flag if: 4+ consecutive spaces OR multiple instances of 2+ spaces
            space_matches = list(re.finditer(r'  +', text))
            if space_matches:
                # Flag if any match has 4+ spaces (clearly intentional alignment, not just sentence spacing)
                # But ignore trailing spaces at the end
                significant_matches = [m for m in space_matches if m.end() < len(text.rstrip())]

                if any(len(match.group()) >= 4 for match in significant_matches):
                    is_pseudo_table = True
                    reasons.append("multiple spaces for alignment")
                # Or if there are multiple occurrences of 2+ spaces (suggesting columns)
                elif len(significant_matches) >= 2:
                    # Additional check: make sure they're not all just sentence spacing
                    # If all matches are exactly after periods, it's likely just double-spacing between sentences
                    non_sentence_spacing = []
                    for match in significant_matches:
                        # Check if this is sentence/clause spacing (punctuation before the spaces)
                        before_pos = match.start() - 1
                        if before_pos >= 0 and text[before_pos] not in '.!?:;)"\'"\u201d\u2019':
                            non_sentence_spacing.append(match)

                    if len(non_sentence_spacing) >= 3:
                        is_pseudo_table = True
                        reasons.append("multiple spacing gaps (columnar layout)")

            if is_pseudo_table:
                # Create a preview of the problematic text
                preview = text[:60] + "..." if len(text) > 60 else text

                issue = AccessibilityIssue(
                    issue_type="PSEUDO_TABLE",
                    description=f"Uses {' and '.join(reasons)} to simulate a table. Replace with an actual Word table (Insert > Table) so screen readers can navigate rows and columns properly: \"{preview}\"",
                    location=para_info.location,
                    para_info=para_info
                )
                issues.append(issue)

        return issues

    def check_font_sizes(self) -> List[AccessibilityIssue]:
        """Check for font sizes below 11pt (inaccessible for body text)"""
        issues = []

        for para_info in self.all_paragraphs:
            para = para_info.paragraph
            text = para.text.strip()

            # Skip empty paragraphs
            if not text:
                continue

            # Check each run for small font sizes
            for run in para.runs:
                if run.font.size and run.font.size.pt < 11:
                    # Get a preview of the text with small font
                    run_text = run.text.strip()
                    if run_text:  # Only flag if there's actual text content
                        preview = text[:60] + "..." if len(text) > 60 else text

                        issue = AccessibilityIssue(
                            issue_type="SMALL_FONT",
                            description=f"Font size {run.font.size.pt}pt is below the 11pt minimum. Increase it to at least 11pt for readability: \"{preview}\"",
                            location=para_info.location,
                            para_info=para_info
                        )
                        issues.append(issue)
                        break  # Only flag paragraph once

        return issues

    def check_decorative_fonts(self) -> List[AccessibilityIssue]:
        """Detect decorative or inaccessible fonts that are hard to read"""
        issues = []

        # List of decorative/display fonts that are hard to read or inaccessible
        # These include script fonts, decorative fonts, and fonts known for poor readability
        decorative_fonts = {
            # Script/Handwriting fonts
            'brush script', 'lucida handwriting', 'mistral', 'monotype corsiva',
            'palace script', 'vivaldi', 'edwardian script', 'freestyle script',
            'french script', 'kunstler script', 'script', 'cursive',

            # Display/Decorative fonts
            'curlz', 'jokerman', 'ravie', 'showcard gothic', 'snap itc',
            'stencil', 'algerian', 'broadway', 'old english', 'chiller',
            'harrington', 'papyrus', 'comic sans', 'impact',

            # Gothic/Blackletter fonts
            'blackadder', 'fraktur', 'old english text', 'blackletter',

            # Comic/Novelty fonts
            'comic sans ms', 'comic sans', 'kristen itc', 'tempus sans itc',
            'gigi', 'marker', 'jester'
        }

        # Track fonts we've already flagged to avoid duplicate warnings
        flagged_fonts = set()

        for para_info in self.all_paragraphs:
            para = para_info.paragraph
            text = para.text.strip()

            # Skip empty paragraphs
            if not text:
                continue

            # Check each run for decorative fonts
            for run in para.runs:
                run_text = run.text.strip()
                if not run_text:
                    continue

                font_name = run.font.name
                if font_name:
                    font_name_lower = font_name.lower()

                    # Check if this font is in our decorative list
                    is_decorative = False
                    matched_font = None
                    for dec_font in decorative_fonts:
                        if dec_font in font_name_lower:
                            is_decorative = True
                            matched_font = dec_font
                            break

                    if is_decorative and font_name not in flagged_fonts:
                        flagged_fonts.add(font_name)
                        preview = text[:60] + "..." if len(text) > 60 else text

                        issue = AccessibilityIssue(
                            issue_type="DECORATIVE_FONT",
                            description=f"Font '{font_name}' is decorative/hard to read (use Arial, Calibri, or Times New Roman): \"{preview}\"",
                            location=para_info.location,
                            para_info=para_info
                        )
                        issues.append(issue)
                        break  # Only flag paragraph once

        return issues

    def check_inconsistent_fonts(self) -> List[AccessibilityIssue]:
        """Identify inconsistent font families across the document"""
        issues = []

        # Track font usage across the document
        font_usage = {}  # font_name -> count
        font_locations = {}  # font_name -> list of locations

        # Collect all fonts used
        for para_info in self.all_paragraphs:
            para = para_info.paragraph
            text = para.text.strip()

            # Skip empty paragraphs
            if not text:
                continue

            # Skip headings (they may legitimately use different fonts)
            if para.style.name.startswith('Heading'):
                continue

            # Check each run for font
            for run in para.runs:
                run_text = run.text.strip()
                if not run_text:
                    continue

                font_name = run.font.name
                if font_name:
                    # Count this font
                    font_usage[font_name] = font_usage.get(font_name, 0) + 1

                    # Track location
                    if font_name not in font_locations:
                        font_locations[font_name] = []
                    if para_info.location not in font_locations[font_name]:
                        font_locations[font_name].append(para_info.location)

        # Identify the primary font (most commonly used)
        if not font_usage:
            return issues

        primary_font = max(font_usage, key=font_usage.get)
        primary_count = font_usage[primary_font]
        total_count = sum(font_usage.values())

        # Check for font inconsistency
        # If there are multiple fonts and the primary font isn't dominant (>80%), flag it
        if len(font_usage) > 2:
            # Multiple fonts detected
            primary_ratio = primary_count / total_count if total_count > 0 else 0

            # List secondary fonts (fonts that aren't the primary)
            secondary_fonts = {font: count for font, count in font_usage.items() if font != primary_font}

            # Only flag if there are significant secondary fonts (not just a few instances)
            significant_secondary = {font: count for font, count in secondary_fonts.items()
                                    if count > total_count * 0.1}  # More than 10% usage

            if significant_secondary:
                font_summary = []
                font_summary.append(f"{primary_font} ({primary_count} instances, {primary_ratio*100:.1f}%)")
                for font, count in sorted(significant_secondary.items(), key=lambda x: x[1], reverse=True):
                    ratio = count / total_count * 100
                    font_summary.append(f"{font} ({count} instances, {ratio:.1f}%)")

                issue = AccessibilityIssue(
                    issue_type="INCONSISTENT_FONTS",
                    description=f"Document uses {len(font_usage)} different fonts. Consider using a single font for consistency. Fonts: {'; '.join(font_summary[:3])}{'...' if len(font_summary) > 3 else ''}",
                    location="Document-wide"
                )
                issues.append(issue)

        return issues

    def check_color_contrast(self) -> List[AccessibilityIssue]:
        """Check for insufficient color contrast based on WCAG thresholds"""
        issues = []

        for para_info in self.all_paragraphs:
            para = para_info.paragraph
            text = para.text.strip()

            # Skip empty paragraphs
            if not text:
                continue

            # Check each run for color contrast
            for run in para.runs:
                run_text = run.text.strip()
                if not run_text:
                    continue

                # Get text color (foreground)
                text_color = None
                try:
                    if run.font.color.rgb:
                        text_color = (run.font.color.rgb[0], run.font.color.rgb[1], run.font.color.rgb[2])
                    elif run.font.color.theme_color:
                        # Default to black for theme colors (conservative)
                        text_color = (0, 0, 0)
                    else:
                        # No explicit color set, assume black (default)
                        text_color = (0, 0, 0)
                except:
                    # If we can't get color, assume black
                    text_color = (0, 0, 0)

                # Get background color - check paragraph shading
                bg_color = None
                try:
                    # Check paragraph shading
                    pPr = para._element.get_or_add_pPr()
                    shd = pPr.find(qn('w:shd'))
                    if shd is not None:
                        fill = shd.get(qn('w:fill'))
                        if fill and fill != 'auto' and len(fill) == 6:
                            # Convert hex to RGB
                            bg_color = tuple(int(fill[i:i+2], 16) for i in (0, 2, 4))
                except:
                    pass

                # If no paragraph background, assume white (default)
                if bg_color is None:
                    bg_color = (255, 255, 255)

                # Calculate contrast ratio
                if text_color and bg_color:
                    contrast_ratio = calculate_contrast_ratio(text_color, bg_color)

                    # Determine if text is "large" (18pt+ or 14pt+ bold)
                    is_large_text = False
                    font_size = run.font.size
                    if font_size:
                        size_pt = font_size.pt
                        if run.bold and size_pt >= 14:
                            is_large_text = True
                        elif size_pt >= 18:
                            is_large_text = True

                    # WCAG AA thresholds
                    required_ratio = 3.0 if is_large_text else 4.5

                    if contrast_ratio < required_ratio:
                        preview = text[:60] + "..." if len(text) > 60 else text
                        text_color_hex = '#{:02x}{:02x}{:02x}'.format(*text_color)
                        bg_color_hex = '#{:02x}{:02x}{:02x}'.format(*bg_color)

                        issue = AccessibilityIssue(
                            issue_type="LOW_CONTRAST",
                            description=f"Insufficient color contrast {contrast_ratio:.2f}:1 (minimum {required_ratio}:1). Text: {text_color_hex}, Background: {bg_color_hex}. \"{preview}\"",
                            location=para_info.location,
                            para_info=para_info
                        )
                        issues.append(issue)
                        break  # Only flag paragraph once

        return issues

    def check_color_only_meaning(self) -> List[AccessibilityIssue]:
        """Flag color used as the sole means of conveying meaning"""
        issues = []

        # Track colored text runs to identify patterns
        colored_runs = []

        for para_info in self.all_paragraphs:
            para = para_info.paragraph
            text = para.text.strip()

            if not text:
                continue

            # Check for runs with different colors but no other distinguishing features
            run_colors = []
            for run in para.runs:
                run_text = run.text.strip()
                if not run_text:
                    continue

                try:
                    # Check if run has explicit color
                    has_color = run.font.color.rgb is not None or run.font.color.theme_color is not None

                    # Check for other formatting
                    has_bold = run.bold
                    has_italic = run.italic
                    has_underline = run.font.underline

                    if has_color and not (has_bold or has_italic or has_underline):
                        # Color is used without other formatting
                        color_rgb = None
                        try:
                            if run.font.color.rgb:
                                color_rgb = (run.font.color.rgb[0], run.font.color.rgb[1], run.font.color.rgb[2])
                        except:
                            pass

                        run_colors.append({
                            'text': run_text,
                            'color': color_rgb,
                            'para_info': para_info
                        })
                except:
                    pass

            # If we found multiple colored runs in the same paragraph, flag it
            if len(run_colors) >= 2:
                # Check if they have different colors
                unique_colors = set()
                for rc in run_colors:
                    if rc['color']:
                        unique_colors.add(rc['color'])

                if len(unique_colors) > 1:
                    preview = text[:60] + "..." if len(text) > 60 else text
                    issue = AccessibilityIssue(
                        issue_type="COLOR_ONLY_MEANING",
                        description=f"Color appears to be the only way to distinguish text (add bold, italic, or other formatting): \"{preview}\"",
                        location=para_info.location,
                        para_info=para_info
                    )
                    issues.append(issue)

        # Check tables for color-coded cells
        for table_idx, table in enumerate(self.target_doc.tables):
            # Track cell background colors
            cell_colors = {}

            for row_idx, row in enumerate(table.rows):
                for cell_idx, cell in enumerate(row.cells):
                    try:
                        tc = cell._element
                        tcPr = tc.get_or_add_tcPr()
                        shd = tcPr.find(qn('w:shd'))
                        if shd is not None:
                            fill = shd.get(qn('w:fill'))
                            if fill and fill != 'auto' and fill != 'FFFFFF':
                                # This cell has a background color
                                cell_text = cell.text.strip()
                                if fill not in cell_colors:
                                    cell_colors[fill] = []
                                cell_colors[fill].append({
                                    'row': row_idx,
                                    'col': cell_idx,
                                    'text': cell_text[:30]
                                })
                    except:
                        pass

            # If table has multiple background colors, it might be using color for meaning
            if len(cell_colors) > 1:
                issue = AccessibilityIssue(
                    issue_type="COLOR_CODED_TABLE",
                    description=f"Table {table_idx + 1} — uses {len(cell_colors)} background colors to convey meaning",
                    location=f"Table {table_idx + 1}",
                    table_index=table_idx
                )
                issues.append(issue)

        return issues

    def check_text_over_backgrounds(self) -> List[AccessibilityIssue]:
        """Identify text placed over backgrounds/images without readable contrast"""
        issues = []

        for para_info in self.all_paragraphs:
            para = para_info.paragraph
            text = para.text.strip()

            # Skip empty paragraphs
            if not text:
                continue

            # Check for paragraph background/shading
            has_background = False
            bg_color = None

            try:
                pPr = para._element.get_or_add_pPr()
                shd = pPr.find(qn('w:shd'))
                if shd is not None:
                    fill = shd.get(qn('w:fill'))
                    if fill and fill != 'auto':
                        has_background = True
                        # Convert hex to RGB
                        if len(fill) == 6:
                            bg_color = tuple(int(fill[i:i+2], 16) for i in (0, 2, 4))
            except:
                pass

            if has_background:
                # Check contrast of all text runs against this background
                for run in para.runs:
                    run_text = run.text.strip()
                    if not run_text:
                        continue

                    # Get text color
                    text_color = None
                    try:
                        if run.font.color.rgb:
                            text_color = (run.font.color.rgb[0], run.font.color.rgb[1], run.font.color.rgb[2])
                        else:
                            # Default to black
                            text_color = (0, 0, 0)
                    except:
                        text_color = (0, 0, 0)

                    if text_color and bg_color:
                        contrast_ratio = calculate_contrast_ratio(text_color, bg_color)

                        # Check if text is large
                        is_large_text = False
                        if run.font.size:
                            size_pt = run.font.size.pt
                            if run.bold and size_pt >= 14:
                                is_large_text = True
                            elif size_pt >= 18:
                                is_large_text = True

                        required_ratio = 3.0 if is_large_text else 4.5

                        if contrast_ratio < required_ratio:
                            preview = text[:60] + "..." if len(text) > 60 else text
                            bg_color_hex = '#{:02x}{:02x}{:02x}'.format(*bg_color)

                            issue = AccessibilityIssue(
                                issue_type="TEXT_OVER_BACKGROUND",
                                description=f"Text over colored background (#{bg_color_hex}) has low contrast {contrast_ratio:.2f}:1 (minimum {required_ratio}:1): \"{preview}\"",
                                location=para_info.location,
                                para_info=para_info
                            )
                            issues.append(issue)
                            break  # Only flag paragraph once

        # Check table cells with backgrounds
        for table_idx, table in enumerate(self.target_doc.tables):
            for row_idx, row in enumerate(table.rows):
                for cell_idx, cell in enumerate(row.cells):
                    # Get cell background color
                    bg_color = None
                    try:
                        tc = cell._element
                        tcPr = tc.get_or_add_tcPr()
                        shd = tcPr.find(qn('w:shd'))
                        if shd is not None:
                            fill = shd.get(qn('w:fill'))
                            if fill and fill != 'auto' and len(fill) == 6:
                                bg_color = tuple(int(fill[i:i+2], 16) for i in (0, 2, 4))
                    except:
                        pass

                    if bg_color:
                        # Check text in this cell
                        for para in cell.paragraphs:
                            para_text = para.text.strip()
                            if not para_text:
                                continue

                            for run in para.runs:
                                run_text = run.text.strip()
                                if not run_text:
                                    continue

                                # Get text color
                                text_color = None
                                try:
                                    if run.font.color.rgb:
                                        text_color = (run.font.color.rgb[0], run.font.color.rgb[1], run.font.color.rgb[2])
                                    else:
                                        text_color = (0, 0, 0)
                                except:
                                    text_color = (0, 0, 0)

                                if text_color:
                                    contrast_ratio = calculate_contrast_ratio(text_color, bg_color)

                                    # Check if large text
                                    is_large_text = False
                                    if run.font.size:
                                        size_pt = run.font.size.pt
                                        if run.bold and size_pt >= 14:
                                            is_large_text = True
                                        elif size_pt >= 18:
                                            is_large_text = True

                                    required_ratio = 3.0 if is_large_text else 4.5

                                    if contrast_ratio < required_ratio:
                                        preview = para_text[:60] + "..." if len(para_text) > 60 else para_text
                                        bg_color_hex = '#{:02x}{:02x}{:02x}'.format(*bg_color)

                                        issue = AccessibilityIssue(
                                            issue_type="TEXT_OVER_BACKGROUND",
                                            description=f"Text in table cell over colored background (#{bg_color_hex}) has low contrast {contrast_ratio:.2f}:1 (minimum {required_ratio}:1): \"{preview}\"",
                                            location=f"Table {table_idx + 1}, Row {row_idx + 1}, Cell {cell_idx + 1}",
                                            table_index=table_idx
                                        )
                                        issues.append(issue)
                                        break  # Only flag cell once

        return issues

    def check_empty_table_rows_columns(self) -> List[AccessibilityIssue]:
        """Detect empty rows/columns in tables used for spacing"""
        issues = []

        for table_idx, table in enumerate(self.target_doc.tables):
            # Check for completely empty rows
            for row_idx, row in enumerate(table.rows):
                all_empty = True
                for cell in row.cells:
                    cell_text = cell.text.strip()
                    if cell_text:
                        all_empty = False
                        break

                if all_empty:
                    issue = AccessibilityIssue(
                        issue_type="EMPTY_TABLE_ROW",
                        description=f"Table {table_idx + 1}, Row {row_idx + 1} — empty row used for spacing. Delete the row and use cell padding or paragraph spacing for visual spacing instead (Table Properties > Cell > Options).",
                        location=f"Table {table_idx + 1}, Row {row_idx + 1}",
                        table_index=table_idx
                    )
                    issues.append(issue)

            # Check for completely empty columns
            num_rows = len(table.rows)
            num_cols = len(table.columns) if table.rows else 0

            if num_rows > 0 and num_cols > 0:
                for col_idx in range(num_cols):
                    all_empty = True
                    for row in table.rows:
                        if col_idx < len(row.cells):
                            cell_text = row.cells[col_idx].text.strip()
                            if cell_text:
                                all_empty = False
                                break

                    if all_empty:
                        issue = AccessibilityIssue(
                            issue_type="EMPTY_TABLE_COLUMN",
                            description=f"Table {table_idx + 1}, Column {col_idx + 1} — empty column used for spacing. Delete the column and use cell padding for visual spacing instead (Table Properties > Cell > Options).",
                            location=f"Table {table_idx + 1}, Column {col_idx + 1}",
                            table_index=table_idx
                        )
                        issues.append(issue)

        return issues

    def check_table_scope_declarations(self) -> List[AccessibilityIssue]:
        """Flag tables missing proper scope declarations on header cells"""
        issues = []

        for table_idx, table in enumerate(self.target_doc.tables):
            # Check if table has a header row first
            if not self._table_has_header_row(table):
                continue  # Skip tables without headers

            # Check if the first row is marked as a "repeat header row"
            # This is Word's way of indicating header rows for accessibility
            first_row = table.rows[0]
            has_repeat_header = False

            try:
                # Check the trPr (table row properties) for the tblHeader element
                tr = first_row._element
                trPr = tr.find(qn('w:trPr'))
                if trPr is not None:
                    tblHeader = trPr.find(qn('w:tblHeader'))
                    if tblHeader is not None:
                        # The presence of w:tblHeader means "repeat as header row" is enabled
                        has_repeat_header = True
            except:
                pass

            # If table has a header row but "repeat header row" is not enabled
            if not has_repeat_header:
                # Get preview of header content for context
                header_preview = " | ".join([cell.text.strip()[:20] for cell in table.rows[0].cells[:4]])
                issue = AccessibilityIssue(
                    issue_type="TABLE_MISSING_SCOPE",
                    description=f"Table {table_idx + 1} ({header_preview}) — header row not marked as repeating. Select the header row, then enable 'Repeat Header Rows' under Table Layout > Properties so it repeats across pages and assistive technology identifies it correctly.",
                    location=f"Table {table_idx + 1}",
                    table_index=table_idx
                )
                issues.append(issue)

        return issues

    def check_table_captions(self) -> List[AccessibilityIssue]:
        """Detect missing captions or table descriptions"""
        issues = []

        for table_idx, table in enumerate(self.target_doc.tables):
            # Check if table has a caption or title
            # In Word, captions are typically added as paragraphs before/after the table
            # or in the table properties (Alt Text)

            has_caption = False

            # Method 1: Check for Alt Text (Title/Description)
            try:
                # Access table properties to check for title/description
                tbl = table._element
                tblPr = tbl.find(qn('w:tblPr'))
                if tblPr is not None:
                    # Look for table description in Alt Text
                    tblDescription = tblPr.find(qn('w:tblDescription'))
                    tblTitle = tblPr.find(qn('w:tblCaption'))

                    if tblDescription is not None:
                        desc_val = tblDescription.get(qn('w:val'))
                        if desc_val and desc_val.strip():
                            has_caption = True

                    if tblTitle is not None:
                        title_val = tblTitle.get(qn('w:val'))
                        if title_val and title_val.strip():
                            has_caption = True
            except:
                pass

            if not has_caption:
                # Get table preview for context
                preview = ""
                if table.rows:
                    preview = " | ".join([cell.text.strip()[:15] for cell in table.rows[0].cells[:3]])

                issue = AccessibilityIssue(
                    issue_type="TABLE_MISSING_CAPTION",
                    description=f"Table {table_idx + 1} ({preview}) — missing Alt Text description. Right-click the table, choose Table Properties > Alt Text, and add a brief description of what the table shows.",
                    location=f"Table {table_idx + 1}",
                    table_index=table_idx
                )
                issues.append(issue)

        return issues

    def check_table_merged_cells(self) -> List[AccessibilityIssue]:
        """Detect merged or split cells that may impact reading order"""
        issues = []

        for table_idx, table in enumerate(self.target_doc.tables):
            merged_cells_found = False
            merged_locations = []

            for row_idx, row in enumerate(table.rows):
                for cell_idx, cell in enumerate(row.cells):
                    try:
                        tc = cell._element
                        tcPr = tc.find(qn('w:tcPr'))

                        if tcPr is not None:
                            # Check for horizontal merge (gridSpan)
                            gridSpan = tcPr.find(qn('w:gridSpan'))
                            if gridSpan is not None:
                                span_val = gridSpan.get(qn('w:val'))
                                if span_val and int(span_val) > 1:
                                    merged_cells_found = True
                                    merged_locations.append(f"Row {row_idx + 1}, Col {cell_idx + 1}")

                            # Check for vertical merge (vMerge)
                            vMerge = tcPr.find(qn('w:vMerge'))
                            if vMerge is not None:
                                merged_cells_found = True
                                if f"Row {row_idx + 1}, Col {cell_idx + 1}" not in merged_locations:
                                    merged_locations.append(f"Row {row_idx + 1}, Col {cell_idx + 1}")
                    except:
                        pass

            if merged_cells_found:
                locations_str = "; ".join(merged_locations[:3])  # Show first 3 locations
                if len(merged_locations) > 3:
                    locations_str += f" (and {len(merged_locations) - 3} more)"

                issue = AccessibilityIssue(
                    issue_type="TABLE_MERGED_CELLS",
                    description=f"Table {table_idx + 1} — merged cells at {locations_str}. Merged cells disrupt screen reader reading order. Unmerge where possible and restructure the table so each row and column has a consistent number of cells.",
                    location=f"Table {table_idx + 1}",
                    table_index=table_idx
                )
                issues.append(issue)

        return issues

    def check_table_numeric_alignment(self) -> List[AccessibilityIssue]:
        """Identify non-left-aligned numeric data inconsistently applied"""
        issues = []

        for table_idx, table in enumerate(self.target_doc.tables):
            # Track numeric columns and their alignment
            num_rows = len(table.rows)
            num_cols = len(table.columns) if table.rows else 0

            if num_rows == 0 or num_cols == 0:
                continue

            # Skip header row if present
            start_row = 1 if self._table_has_header_row(table) else 0

            # Check each column for numeric data
            for col_idx in range(num_cols):
                numeric_count = 0
                total_count = 0
                alignments = []

                for row_idx in range(start_row, num_rows):
                    if col_idx < len(table.rows[row_idx].cells):
                        cell = table.rows[row_idx].cells[col_idx]
                        cell_text = cell.text.strip()

                        if cell_text:
                            total_count += 1

                            # Check if cell contains numeric data
                            # Remove common formatting (currency symbols, commas, %)
                            cleaned = cell_text.replace('$', '').replace(',', '').replace('%', '').strip()
                            try:
                                float(cleaned)
                                numeric_count += 1

                                # Check alignment of this cell
                                try:
                                    for para in cell.paragraphs:
                                        if para.text.strip():
                                            alignment = para.alignment
                                            if alignment is not None:
                                                alignments.append(str(alignment))
                                            break  # Only check first paragraph
                                except:
                                    pass
                            except ValueError:
                                pass

                # If column is mostly numeric (>70%) and has inconsistent alignment
                if total_count > 0 and numeric_count / total_count > 0.7:
                    # Check if alignments are inconsistent
                    if len(set(alignments)) > 1:
                        issue = AccessibilityIssue(
                            issue_type="TABLE_INCONSISTENT_NUMERIC_ALIGNMENT",
                            description=f"Table {table_idx + 1}, Column {col_idx + 1} — numbers have {len(set(alignments))} different alignments. Align all numeric cells consistently (right-align is standard for numbers) by selecting the column and setting alignment in the Home tab.",
                            location=f"Table {table_idx + 1}, Column {col_idx + 1}",
                            table_index=table_idx
                        )
                        issues.append(issue)

        return issues

    def check_table_reading_order(self) -> List[AccessibilityIssue]:
        """Flag inconsistent or illogical table reading order"""
        issues = []

        for table_idx, table in enumerate(self.target_doc.tables):
            # Check for tables with complex structures that might confuse reading order

            # Factor 1: Check if table has both merged cells AND multiple header patterns
            has_merged = False
            try:
                for row in table.rows:
                    for cell in row.cells:
                        tc = cell._element
                        tcPr = tc.find(qn('w:tcPr'))
                        if tcPr is not None:
                            if tcPr.find(qn('w:gridSpan')) is not None or tcPr.find(qn('w:vMerge')) is not None:
                                has_merged = True
                                break
                    if has_merged:
                        break
            except:
                pass

            # Factor 2: Check if table has multiple "header-like" rows (bold rows beyond first)
            header_like_rows = []
            for row_idx, row in enumerate(table.rows):
                bold_count = 0
                total_runs = 0

                for cell in row.cells:
                    for para in cell.paragraphs:
                        for run in para.runs:
                            total_runs += 1
                            if run.bold:
                                bold_count += 1

                if total_runs > 0 and bold_count / total_runs > 0.7:
                    header_like_rows.append(row_idx)

            # If table has merged cells AND multiple header rows, reading order may be confusing
            if has_merged and len(header_like_rows) > 1:
                issue = AccessibilityIssue(
                    issue_type="TABLE_COMPLEX_READING_ORDER",
                    description=f"Table {table_idx + 1} — complex structure with merged cells and multiple header-like rows (rows {', '.join(str(r+1) for r in header_like_rows)})",
                    location=f"Table {table_idx + 1}",
                    table_index=table_idx
                )
                issues.append(issue)

        return issues

    def check_table_color_coded_meaning(self) -> List[AccessibilityIssue]:
        """Detect color-coded meaning in cells without textual explanation"""
        issues = []

        for table_idx, table in enumerate(self.target_doc.tables):
            # Track cells with background colors
            colored_cells = {}  # color -> list of (row, col, text)

            for row_idx, row in enumerate(table.rows):
                for cell_idx, cell in enumerate(row.cells):
                    cell_text = cell.text.strip()

                    # Check for cell background color
                    try:
                        tc = cell._element
                        tcPr = tc.find(qn('w:tcPr'))
                        if tcPr is not None:
                            shd = tcPr.find(qn('w:shd'))
                            if shd is not None:
                                fill = shd.get(qn('w:fill'))
                                # Ignore white/auto backgrounds
                                if fill and fill != 'auto' and fill.upper() not in ['FFFFFF', 'FFFF', 'FFF']:
                                    if fill not in colored_cells:
                                        colored_cells[fill] = []
                                    colored_cells[fill].append((row_idx, cell_idx, cell_text[:30]))
                    except:
                        pass

            # If table has multiple distinct colors used across cells
            if len(colored_cells) >= 2:
                # Check if there's explanatory text (legend) in the table or nearby
                # Simple heuristic: if cells with same color have very similar content,
                # color might be conveying meaning

                potentially_meaningful_colors = []
                for color, cells in colored_cells.items():
                    if len(cells) >= 2:  # Color used multiple times
                        # Check if cells with this color have varying content
                        texts = [c[2] for c in cells]
                        unique_texts = len(set(texts))
                        if unique_texts > 1:  # Different content, same color
                            potentially_meaningful_colors.append(color)

                if len(potentially_meaningful_colors) >= 2:
                    issue = AccessibilityIssue(
                        issue_type="TABLE_COLOR_CODED_MEANING",
                        description=f"Table {table_idx + 1} — {len(potentially_meaningful_colors)} background colors may convey meaning without text labels",
                        location=f"Table {table_idx + 1}",
                        table_index=table_idx
                    )
                    issues.append(issue)

        return issues

    def check_table_embedded_images(self) -> List[AccessibilityIssue]:
        """Identify images embedded inside table cells when text should be used"""
        issues = []

        for table_idx, table in enumerate(self.target_doc.tables):
            cells_with_images = []

            for row_idx, row in enumerate(table.rows):
                for cell_idx, cell in enumerate(row.cells):
                    # Check if cell contains images
                    has_image = False
                    cell_text = cell.text.strip()

                    for para in cell.paragraphs:
                        # Check for inline shapes (images)
                        try:
                            # Look for drawing elements in the paragraph
                            para_element = para._element
                            drawings = para_element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing')
                            if drawings:
                                has_image = True
                                break

                            # Also check for pict elements (older image format)
                            picts = para_element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pict')
                            if picts:
                                has_image = True
                                break
                        except:
                            pass

                    if has_image:
                        # Flag if image is in cell with little/no text
                        if len(cell_text) < 10:  # Arbitrary threshold
                            cells_with_images.append(f"Row {row_idx + 1}, Col {cell_idx + 1}")

            if cells_with_images:
                locations_str = "; ".join(cells_with_images[:3])
                if len(cells_with_images) > 3:
                    locations_str += f" (and {len(cells_with_images) - 3} more)"

                issue = AccessibilityIssue(
                    issue_type="TABLE_EMBEDDED_IMAGES",
                    description=f"Table {table_idx + 1} — images in cells with minimal text at {locations_str}",
                    location=f"Table {table_idx + 1}",
                    table_index=table_idx
                )
                issues.append(issue)

        return issues

    def check_document_metadata(self) -> List[AccessibilityIssue]:
        """Check for missing document title"""
        issues = []

        # Check document title
        core_properties = self.target_doc.core_properties
        title = core_properties.title

        if not title or title.strip() == "":
            issue = AccessibilityIssue(
                issue_type="MISSING_TITLE",
                description="Document is missing a title in metadata (File > Info > Properties)",
                location="Document metadata"
            )
            issues.append(issue)

        return issues

    def check_document_language(self) -> List[AccessibilityIssue]:
        """Check for missing or incorrect document language setting - only flags if multilingual content detected"""
        issues = []

        # Note: We no longer flag missing language settings for monolingual documents
        # because many properly-formatted Word documents (especially from Mac) don't have
        # explicit language tags, and screen readers will use their default language correctly.
        # We only care about language tagging when there's actual multilingual content,
        # which is handled by check_multilingual_content()

        # Still check for explicitly invalid language codes if they exist
        has_language = False
        document_language = None

        try:
            # Check the document's default paragraph style for language
            for style in self.target_doc.styles:
                if style.name in ['Normal', 'Default Paragraph Font', 'No Spacing', 'Body Text']:
                    try:
                        style_element = style.element
                        rPr = style_element.find(qn('w:rPr'))
                        if rPr is not None:
                            lang = rPr.find(qn('w:lang'))
                            if lang is not None:
                                val = lang.get(qn('w:val'))
                                if not val:
                                    val = lang.get(qn('w:eastAsia'))
                                if not val:
                                    val = lang.get(qn('w:bidi'))
                                if val:
                                    has_language = True
                                    document_language = val
                                    break
                    except:
                        pass
        except:
            pass

        # Check first few paragraphs for explicit language settings
        if not has_language:
            for para_info in self.all_paragraphs[:10]:
                para = para_info.paragraph
                if para.runs:
                    for run in para.runs:
                        try:
                            rPr = run._element.find(qn('w:rPr'))
                            if rPr is not None:
                                lang = rPr.find(qn('w:lang'))
                                if lang is not None:
                                    val = lang.get(qn('w:val'))
                                    if not val:
                                        val = lang.get(qn('w:eastAsia'))
                                    if not val:
                                        val = lang.get(qn('w:bidi'))
                                    if val:
                                        has_language = True
                                        document_language = val
                                        break
                        except:
                            pass
                if has_language:
                    break

        # Only flag invalid language codes if they exist
        if has_language and document_language:
            if len(document_language) < 2:
                issue = AccessibilityIssue(
                    issue_type="INVALID_LANGUAGE",
                    description=f"Document language setting '{document_language}' appears invalid. Use standard language codes (e.g., en-US, es-ES).",
                    location="Document metadata"
                )
                issues.append(issue)

        return issues

    def check_multilingual_content(self) -> List[AccessibilityIssue]:
        """Identify multiple languages without proper tagging"""
        issues = []

        # Track different languages used in the document
        languages_found = set()
        language_locations = {}  # language -> list of locations

        for para_info in self.all_paragraphs:
            para = para_info.paragraph
            text = para.text.strip()

            if not text:
                continue

            # Check each run for language tags
            for run in para.runs:
                run_text = run.text.strip()
                if not run_text:
                    continue

                try:
                    # Check if run has explicit language setting
                    rPr = run._element.get_or_add_rPr()
                    lang = rPr.find(qn('w:lang'))

                    lang_code = None
                    if lang is not None:
                        val = lang.get(qn('w:val'))
                        if val:
                            lang_code = val
                            languages_found.add(lang_code)

                            if lang_code not in language_locations:
                                language_locations[lang_code] = []
                            if para_info.location not in language_locations[lang_code]:
                                language_locations[lang_code].append(para_info.location)
                except:
                    pass

        # If we found multiple languages, check if they're properly tagged
        if len(languages_found) > 1:
            # Multiple languages detected - this is fine if properly tagged
            # Just inform the user
            lang_summary = []
            for lang in sorted(languages_found):
                count = len(language_locations.get(lang, []))
                lang_summary.append(f"{lang} ({count} locations)")

            issue = AccessibilityIssue(
                issue_type="MULTIPLE_LANGUAGES",
                description=f"Document contains {len(languages_found)} different language tags: {', '.join(lang_summary)}. Verify all content is properly tagged with correct language for screen reader pronunciation.",
                location="Document-wide"
            )
            issues.append(issue)

        # Detect common non-English phrases that might need language tagging
        # This is a heuristic check for potential multilingual content
        non_english_patterns = {
            'spanish': [r'\bseñor\b', r'\bseñora\b', r'\bhola\b', r'\bgracias\b', r'\bpor favor\b',
                       r'\b¿\w+\?', r'\b¡\w+!', r'\baño\b', r'\bniño\b', r'\bniña\b'],
            'french': [r'\bmonsieur\b', r'\bmadame\b', r'\bbonjour\b', r'\bmerci\b',
                      r'\bs\'il vous plaît\b', r'\bça\b', r'\bà\b', r'\bélève\b'],
            'german': [r'\bherr\b', r'\bfrau\b', r'\bguten tag\b', r'\bdanke\b',
                      r'\bbitte\b', r'\bstraße\b', r'\bschön\b', r'\büber\b'],
        }

        potential_multilingual = []

        for para_info in self.all_paragraphs:
            para = para_info.paragraph
            text = para.text.lower()

            if not text or len(text) < 10:
                continue

            # Check for non-English patterns
            for language, patterns in non_english_patterns.items():
                for pattern in patterns:
                    if re.search(pattern, text, re.IGNORECASE):
                        # Check if this paragraph has proper language tagging
                        has_proper_tag = False
                        for run in para.runs:
                            try:
                                rPr = run._element.get_or_add_rPr()
                                lang = rPr.find(qn('w:lang'))
                                if lang is not None:
                                    val = lang.get(qn('w:val'))
                                    if val and not val.startswith('en-'):
                                        has_proper_tag = True
                                        break
                            except:
                                pass

                        if not has_proper_tag:
                            preview = text[:60] + "..." if len(text) > 60 else text
                            potential_multilingual.append({
                                'location': para_info.location,
                                'language': language,
                                'text': preview,
                                'para_info': para_info
                            })
                            break  # Only flag once per paragraph
                    if potential_multilingual and potential_multilingual[-1].get('location') == para_info.location:
                        break

        if potential_multilingual:
            # Group by language
            by_language = {}
            for item in potential_multilingual:
                lang = item['language']
                if lang not in by_language:
                    by_language[lang] = []
                by_language[lang].append(item)

            for language, items in by_language.items():
                issue = AccessibilityIssue(
                    issue_type="UNTAGGED_LANGUAGE",
                    description=f"Found {len(items)} paragraph(s) with potential {language.title()} content that may need language tagging. Example: \"{items[0]['text']}\" at {items[0]['location']}",
                    location="Document-wide",
                    para_info=items[0]['para_info']
                )
                issues.append(issue)

        return issues

    def check_underline_non_links(self) -> List[AccessibilityIssue]:
        """Detect underlined text that is not a hyperlink (confusing for navigation)"""
        issues = []

        for para_info in self.all_paragraphs:
            para = para_info.paragraph
            text = para.text.strip()

            if not text:
                continue

            # Check each run for underline
            for run in para.runs:
                run_text = run.text.strip()

                # Skip if no text
                if not run_text:
                    continue

                # Check if underlined
                if run.font.underline:
                    # Check if it's a hyperlink by checking for hyperlink relationship
                    # In python-docx, hyperlinks are tricky to detect
                    # We'll use a heuristic: if it contains http:// or www., it's likely a link
                    is_likely_link = 'http://' in run_text or 'https://' in run_text or 'www.' in run_text

                    # Also check if the run has a hyperlink style
                    # Hyperlinks often have specific formatting
                    has_link_color = False
                    if run.font.color.rgb:
                        # Common link colors are blue (0, 0, 255) or similar
                        rgb = run.font.color.rgb
                        # Check if it's bluish (rgb[0]=red, rgb[1]=green, rgb[2]=blue)
                        if rgb[2] > 200 and rgb[0] < 100 and rgb[1] < 100:
                            has_link_color = True

                    if not is_likely_link and not has_link_color:
                        preview = text[:60] + "..." if len(text) > 60 else text

                        issue = AccessibilityIssue(
                            issue_type="UNDERLINE_NON_LINK",
                            description=f"Underlined text that is not a hyperlink — readers may try to click it. Use bold or italic for emphasis instead, and reserve underline for actual links: \"{preview}\"",
                            location=para_info.location,
                            para_info=para_info
                        )
                        issues.append(issue)
                        break  # Only flag paragraph once

        return issues

    def check_line_spacing(self) -> List[AccessibilityIssue]:
        """Check for line spacing below 1.15 (WCAG recommendation)"""
        issues = []

        for para_info in self.all_paragraphs:
            para = para_info.paragraph
            text = para.text.strip()

            # Skip empty paragraphs
            if not text:
                continue

            # Get line spacing
            # In python-docx, line spacing is in paragraph_format.line_spacing
            # Can be a float (multiple of single spacing) or None (uses default)
            if para.paragraph_format.line_spacing is not None:
                line_spacing = para.paragraph_format.line_spacing

                # Line spacing can be specified as a float (1.0, 1.15, etc.)
                # or as a Pt value - we need to handle both
                if hasattr(line_spacing, '__float__'):
                    spacing_value = float(line_spacing)
                    if spacing_value < 1.15:
                        preview = text[:60] + "..." if len(text) > 60 else text

                        issue = AccessibilityIssue(
                            issue_type="LOW_LINE_SPACING",
                            description=f"Line spacing {spacing_value:.2f} is below the 1.15 minimum for readability. Increase it via Home > Line and Paragraph Spacing (or Paragraph Settings > Line Spacing): \"{preview}\"",
                            location=para_info.location,
                            para_info=para_info
                        )
                        issues.append(issue)

        return issues

    def check_full_justification(self) -> List[AccessibilityIssue]:
        """Check for full text justification (hard to read, especially for dyslexic readers)"""
        issues = []

        for para_info in self.all_paragraphs:
            para = para_info.paragraph
            text = para.text.strip()

            # Skip empty paragraphs
            if not text:
                continue

            # Check alignment - WD_ALIGN_PARAGRAPH.JUSTIFY = 3
            if para.paragraph_format.alignment == 3:  # JUSTIFY
                preview = text[:60] + "..." if len(text) > 60 else text

                issue = AccessibilityIssue(
                    issue_type="FULL_JUSTIFICATION",
                    description=f"Text is fully justified, which creates uneven word spacing that is harder to read. Change alignment to left-aligned (Home > Align Left or Ctrl+L): \"{preview}\"",
                    location=para_info.location,
                    para_info=para_info
                )
                issues.append(issue)

        return issues

    def check_all_caps_blocks(self) -> List[AccessibilityIssue]:
        """Check for large blocks of ALL CAPS text (hard to read)"""
        issues = []

        for para_info in self.all_paragraphs:
            para = para_info.paragraph
            text = para.text.strip()

            # Skip empty paragraphs
            if not text:
                continue

            # Skip very short text (headings already handled elsewhere)
            # We're looking for BLOCKS of caps text, not short headings
            if len(text) < 50:  # Less than 50 chars is likely a heading/label
                continue

            # Check if most of the text is uppercase
            # Count letters only (ignore numbers, punctuation, spaces)
            letters = [c for c in text if c.isalpha()]
            if len(letters) > 0:
                uppercase_letters = [c for c in letters if c.isupper()]
                caps_ratio = len(uppercase_letters) / len(letters)

                # If more than 80% uppercase and it's a substantial block, flag it
                if caps_ratio > 0.8:
                    preview = text[:60] + "..." if len(text) > 60 else text

                    issue = AccessibilityIssue(
                        issue_type="ALL_CAPS_BLOCK",
                        description=f"Large block of ALL CAPS text, which is harder to read and can be misread by some screen readers. Convert to sentence or title case and use bold for emphasis if needed: \"{preview}\"",
                        location=para_info.location,
                        para_info=para_info
                    )
                    issues.append(issue)

        return issues

    def check_non_descriptive_links(self) -> List[AccessibilityIssue]:
        """Check for non-descriptive link text like 'click here', 'here', 'read more'"""
        issues = []

        # Common non-descriptive link text patterns
        non_descriptive_patterns = [
            'click here', 'here', 'read more', 'more', 'link', 'this link',
            'click', 'download', 'more info', 'more information', 'learn more',
            'see more', 'view more', 'details', 'continue', 'next', 'previous'
        ]

        for para_info in self.all_paragraphs:
            para = para_info.paragraph

            # Check hyperlinks in this paragraph
            # In python-docx, hyperlinks are tricky - we'll check run text
            for run in para.runs:
                run_text = run.text.strip().lower()

                # Skip if no text
                if not run_text:
                    continue

                # Check if underlined or has link color (heuristic for links)
                is_likely_link = False
                if run.font.underline:
                    is_likely_link = True
                elif run.font.color.rgb:
                    try:
                        rgb = run.font.color.rgb
                        # RGBColor is a tuple-like (r, g, b)
                        # Check if it's bluish (common link color)
                        if len(rgb) >= 3 and rgb[2] > 200 and rgb[0] < 100 and rgb[1] < 100:
                            is_likely_link = True
                    except:
                        pass  # If we can't check color, skip it

                # If it looks like a link, check if text is descriptive
                if is_likely_link:
                    # Check against non-descriptive patterns
                    if run_text in non_descriptive_patterns:
                        preview = para.text[:60] + "..." if len(para.text) > 60 else para.text

                        issue = AccessibilityIssue(
                            issue_type="NON_DESCRIPTIVE_LINK",
                            description=f"Link text '{run_text}' is not descriptive — screen reader users who navigate by links won't know where it goes. Replace with text that describes the destination, e.g. 'Course Schedule (PDF)' or 'UWM Writing Center': \"{preview}\"",
                            location=para_info.location,
                            para_info=para_info
                        )
                        issues.append(issue)
                        break  # Only flag paragraph once

        return issues

    def check_unstyled_links(self) -> List[AccessibilityIssue]:
        """Identify links styled as normal text (missing underline/color)"""
        issues = []

        for para_info in self.all_paragraphs:
            para = para_info.paragraph

            # Check for hyperlinks in the paragraph XML
            try:
                # Access paragraph element to find hyperlinks
                para_element = para._element
                hyperlinks = para_element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hyperlink')

                for hyperlink in hyperlinks:
                    # Get the text content of the hyperlink
                    link_text = ''.join(hyperlink.itertext())

                    # Check the formatting of runs within the hyperlink
                    runs_in_link = hyperlink.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')

                    has_underline = False
                    has_distinctive_color = False
                    has_hyperlink_style = False

                    for run_elem in runs_in_link:
                        rPr = run_elem.find(qn('w:rPr'))
                        if rPr is not None:
                            # Check for "Hyperlink" character style (default Word hyperlink styling)
                            rStyle = rPr.find(qn('w:rStyle'))
                            if rStyle is not None:
                                style_val = rStyle.get(qn('w:val'))
                                # Common hyperlink style names
                                if style_val and 'hyperlink' in style_val.lower():
                                    has_hyperlink_style = True

                            # Check for explicit underline
                            u = rPr.find(qn('w:u'))
                            if u is not None:
                                has_underline = True

                            # Check for color (non-black)
                            color_elem = rPr.find(qn('w:color'))
                            if color_elem is not None:
                                color_val = color_elem.get(qn('w:val'))
                                # Check if color is not black or auto
                                if color_val and color_val.upper() not in ['000000', '000', 'AUTO', 'FFFFFF']:
                                    has_distinctive_color = True

                    # Only flag if link has NO styling indicators (no hyperlink style, no underline, no distinctive color)
                    if not has_hyperlink_style and not has_underline and not has_distinctive_color:
                        preview = link_text[:40] + "..." if len(link_text) > 40 else link_text

                        issue = AccessibilityIssue(
                            issue_type="UNSTYLED_LINK",
                            description=f"Link is styled as normal text with no underline or color, so readers can't tell it's clickable. Apply the built-in 'Hyperlink' character style, or manually add underline and a blue color: \"{preview}\"",
                            location=para_info.location,
                            para_info=para_info
                        )
                        issues.append(issue)
            except:
                pass  # If we can't parse hyperlinks, skip

        return issues

    def check_long_urls(self) -> List[AccessibilityIssue]:
        """Detect excessively long URLs that should be shortened or use descriptive text"""
        issues = []

        for para_info in self.all_paragraphs:
            para = para_info.paragraph

            # Check for hyperlinks in the paragraph XML
            try:
                para_element = para._element
                hyperlinks = para_element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hyperlink')

                for hyperlink in hyperlinks:
                    # Get the URL from the relationship ID
                    r_id = hyperlink.get(qn('r:id'))
                    if r_id:
                        try:
                            # Get the relationship to find the actual URL
                            url = para.part.rels[r_id].target_ref
                            link_text = ''.join(hyperlink.itertext()).strip()

                            # Only flag if URL is excessively long AND the link text IS the URL
                            # (Good accessibility practice is to hide long URLs behind descriptive text)
                            if len(url) > 100:
                                # Check if the link text is the URL itself (or looks like a URL)
                                link_text_is_url = (
                                    link_text.startswith(('http://', 'https://', 'ftp://', 'file://', 'www.')) or
                                    url.endswith(link_text) or  # Link text is end of URL
                                    link_text in url  # Link text is part of URL
                                )

                                # Only flag if the visible text IS the URL (bad practice)
                                # Don't flag if it's descriptive text (good practice!)
                                if link_text_is_url:
                                    preview = link_text[:40] + "..." if len(link_text) > 40 else link_text

                                    issue = AccessibilityIssue(
                                        issue_type="LONG_URL",
                                        description=f"Excessively long URL displayed as link text ({len(url)} chars): \"{preview}\". "
                                                   f"Replace the URL with descriptive link text (e.g., 'Course Resources' instead of showing the full URL).",
                                        location=para_info.location,
                                        para_info=para_info
                                    )
                                    issues.append(issue)
                        except:
                            pass
            except:
                pass

        return issues

    def check_pdf_links(self) -> List[AccessibilityIssue]:
        """Identify links to external PDF files which may not be accessible"""
        issues = []

        for para_info in self.all_paragraphs:
            para = para_info.paragraph

            # Check for hyperlinks in the paragraph XML
            try:
                para_element = para._element
                hyperlinks = para_element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hyperlink')

                for hyperlink in hyperlinks:
                    # Get the URL from the relationship ID
                    r_id = hyperlink.get(qn('r:id'))
                    if r_id:
                        try:
                            # Get the relationship to find the actual URL
                            url = para.part.rels[r_id].target_ref

                            # Check if URL points to a PDF file
                            if url.lower().endswith('.pdf') or '.pdf?' in url.lower() or '.pdf#' in url.lower():
                                link_text = ''.join(hyperlink.itertext())
                                preview = link_text[:40] + "..." if len(link_text) > 40 else link_text

                                issue = AccessibilityIssue(
                                    issue_type="PDF_LINK",
                                    description=f"Link to external PDF (verify PDF is accessible): \"{preview}\" -> {url[:60]}...",
                                    location=para_info.location,
                                    para_info=para_info
                                )
                                issues.append(issue)
                        except:
                            pass
            except:
                pass

        return issues

    def check_missing_toc(self) -> List[AccessibilityIssue]:
        """Detect missing table of contents in longer documents"""
        issues = []

        # Count paragraphs and check document length
        num_paragraphs = len(self.all_paragraphs)
        num_pages_estimate = num_paragraphs / 30  # Rough estimate: ~30 paragraphs per page

        # Check if document is "long" (>5 pages or >150 paragraphs)
        is_long_document = num_paragraphs > 150 or num_pages_estimate > 5

        if is_long_document:
            # Check if document has a table of contents
            has_toc = False

            # Method 1: Check for TOC field
            try:
                for para in self.target_doc.paragraphs:
                    # Look for TOC field code
                    para_element = para._element
                    fldChar_elements = para_element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fldChar')
                    if fldChar_elements:
                        # Check if this is a TOC field
                        instrText = para_element.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}instrText')
                        if instrText is not None and 'TOC' in instrText.text:
                            has_toc = True
                            break
            except:
                pass

            # Method 2: Check for text that looks like a ToC heading
            if not has_toc:
                for para_info in self.all_paragraphs[:20]:  # Check first 20 paragraphs
                    para_text = para_info.paragraph.text.strip().lower()
                    if para_text in ['table of contents', 'contents', 'table of content']:
                        has_toc = True
                        break

            if not has_toc:
                issue = AccessibilityIssue(
                    issue_type="MISSING_TOC",
                    description=f"Long document (~{int(num_pages_estimate)} pages, {num_paragraphs} paragraphs) is missing a Table of Contents. Add one via References > Table of Contents so readers can jump directly to sections.",
                    location="Document structure"
                )
                issues.append(issue)

        return issues

    def check_missing_bookmarks(self) -> List[AccessibilityIssue]:
        """Detect missing internal navigation/bookmarks in longer documents"""
        issues = []

        # Count internal bookmarks
        bookmark_count = 0
        internal_link_count = 0

        try:
            # Check for bookmarks in the document
            doc_element = self.target_doc.element
            bookmarks = doc_element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}bookmarkStart')
            bookmark_count = len(bookmarks)

            # Filter out built-in bookmarks (like _Toc, _Hlk, _GoBack)
            named_bookmarks = 0
            for bookmark in bookmarks:
                name = bookmark.get(qn('w:name'))
                if name and not name.startswith('_'):
                    named_bookmarks += 1

            # Check for internal hyperlinks (links to bookmarks)
            for para_info in self.all_paragraphs:
                para_element = para_info.paragraph._element
                hyperlinks = para_element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hyperlink')

                for hyperlink in hyperlinks:
                    # Check if it's an internal link (has anchor attribute)
                    anchor = hyperlink.get(qn('w:anchor'))
                    if anchor:
                        internal_link_count += 1

        except:
            pass

        # Check if document is long and lacks bookmarks/internal navigation
        num_paragraphs = len(self.all_paragraphs)
        is_long_document = num_paragraphs > 150

        if is_long_document and named_bookmarks == 0 and internal_link_count == 0:
            issue = AccessibilityIssue(
                issue_type="MISSING_BOOKMARKS",
                description=f"Long document ({num_paragraphs} paragraphs) lacks internal bookmarks or navigation links. Add bookmarks to major sections (Insert > Bookmark) and link to them so readers can jump around without scrolling.",
                location="Document structure"
            )
            issues.append(issue)

        return issues

    def check_nested_list_hierarchy(self) -> List[AccessibilityIssue]:
        """Flag incorrect or inconsistent nested list hierarchy"""
        issues = []

        # Track list levels and patterns
        prev_level = 0
        list_started = False

        for para_info in self.all_paragraphs:
            para = para_info.paragraph
            style_name = para.style.name

            # Check if this is a list paragraph
            if 'List' in style_name:
                list_started = True

                # Try to determine the level from the style name
                # Common patterns: "List Bullet", "List Bullet 2", "List Bullet 3", etc.
                # Or: "List Number", "List Number 2", etc.
                level = 1  # Default to level 1

                # Extract level from style name
                import re
                level_match = re.search(r'(\d+)$', style_name)
                if level_match:
                    level = int(level_match.group(1))

                # Check for level skipping (e.g., going from level 1 to level 3)
                if list_started and level > prev_level + 1:
                    issue = AccessibilityIssue(
                        issue_type="INCONSISTENT_LIST_HIERARCHY",
                        description=f"List jumps from level {prev_level} to level {level}, skipping a level. Change the indentation to increment by one level at a time using the Increase Indent button or Tab key within the list. Text: \"{para.text[:50]}...\"",
                        location=para_info.location,
                        para_info=para_info
                    )
                    issues.append(issue)

                prev_level = level
            else:
                # Reset when we exit list context
                if list_started:
                    prev_level = 0
                    list_started = False

        return issues

    def check_layout_lists(self) -> List[AccessibilityIssue]:
        """Identify lists used for layout or indentation rather than semantic grouping"""
        issues = []

        # Track consecutive list items
        consecutive_lists = []
        current_list_group = []

        for para_info in self.all_paragraphs:
            para = para_info.paragraph
            style_name = para.style.name

            if 'List' in style_name:
                current_list_group.append(para_info)
            else:
                # End of list group
                if current_list_group:
                    consecutive_lists.append(current_list_group)
                    current_list_group = []

        # Add final group if exists
        if current_list_group:
            consecutive_lists.append(current_list_group)

        # Check each list group for potential layout misuse
        for list_group in consecutive_lists:
            # Check if items are very dissimilar in length (indicates layout use)
            # Only check lists with 3+ items where we can detect patterns
            if len(list_group) >= 3:
                lengths = [len(p.paragraph.text.strip()) for p in list_group]
                avg_length = sum(lengths) / len(lengths)

                # Calculate variance
                variance = sum((l - avg_length) ** 2 for l in lengths) / len(lengths)
                std_dev = variance ** 0.5

                # If high variance and some items are very long, likely layout
                if avg_length > 0 and std_dev / avg_length > 2 and max(lengths) > 300:
                    issue = AccessibilityIssue(
                        issue_type="LAYOUT_LIST",
                        description=f"List with highly inconsistent item lengths — this may be using a list style for visual layout rather than actual list content. If it's a list, edit items for consistent length. If it's layout, replace with a table or paragraph spacing instead.",
                        location=list_group[0].location,
                        para_info=list_group[0]
                    )
                    issues.append(issue)

        return issues

    def check_long_sentences(self) -> List[AccessibilityIssue]:
        """Detect overly long sentences (>50 words)"""
        issues = []

        for para_info in self.all_paragraphs:
            para = para_info.paragraph
            text = para.text.strip()

            # Skip empty paragraphs
            if not text:
                continue

            # Split into sentences (simple split by . ! ?)
            # Use a more sophisticated sentence splitter
            import re
            sentences = re.split(r'[.!?]+\s+', text)

            for sentence in sentences:
                sentence = sentence.strip()
                if not sentence:
                    continue

                # Count words (split by whitespace)
                words = sentence.split()
                word_count = len(words)

                # Flag if more than 50 words
                if word_count > 50:
                    preview = sentence[:60] + "..." if len(sentence) > 60 else sentence

                    issue = AccessibilityIssue(
                        issue_type="LONG_SENTENCE",
                        description=f"Sentence has {word_count} words, which is harder to follow. Break it into two or more shorter sentences: \"{preview}\"",
                        location=para_info.location,
                        para_info=para_info
                    )
                    issues.append(issue)
                    # Only flag first long sentence in paragraph to avoid overwhelming output
                    break

        return issues

    def check_excessive_formatting(self) -> List[AccessibilityIssue]:
        """Detect excessive or inconsistent use of bold, italics, or underline"""
        issues = []

        for para_info in self.all_paragraphs:
            para = para_info.paragraph
            text = para.text.strip()

            # Skip empty paragraphs
            if not text or len(text) < 10:
                continue

            # Skip headings (which are often bold)
            if 'Heading' in para.style.name:
                continue

            # Count characters with bold, italic, and underline
            total_chars = len(text)
            bold_chars = 0
            italic_chars = 0
            underline_chars = 0
            formatting_switches = 0
            italic_only_switches = 0   # switches where ONLY italic changed (title-style toggling)
            prev_bold = False
            prev_italic = False
            prev_underline = False

            for run in para.runs:
                run_text = run.text
                run_len = len(run_text)

                if run.bold:
                    bold_chars += run_len
                if run.italic:
                    italic_chars += run_len
                if run.font.underline:
                    underline_chars += run_len

                # Track formatting switches (inconsistent use)
                italic_changed    = run.italic        != prev_italic
                bold_changed      = run.bold          != prev_bold
                underline_changed = run.font.underline != prev_underline
                if italic_changed or bold_changed or underline_changed:
                    formatting_switches += 1
                    if italic_changed and not bold_changed and not underline_changed:
                        italic_only_switches += 1

                prev_bold = run.bold
                prev_italic = run.italic
                prev_underline = run.font.underline

            # Flag if >50% of paragraph is bold, italic, or underlined
            if total_chars > 0:
                bold_ratio = bold_chars / total_chars
                italic_ratio = italic_chars / total_chars
                underline_ratio = underline_chars / total_chars

                # Flag excessive bold (>50%)
                if bold_ratio > 0.5:
                    preview = text[:60] + "..." if len(text) > 60 else text
                    issue = AccessibilityIssue(
                        issue_type="EXCESSIVE_BOLD",
                        description=f"Paragraph has {int(bold_ratio*100)}% bold text. If this is a heading, use a Heading style (Heading 1, 2, etc.) instead of bold. If intended for emphasis, use bold more selectively: \"{preview}\"",
                        location=para_info.location,
                        para_info=para_info
                    )
                    issues.append(issue)

                # Flag excessive italic (>50%), but skip paragraphs where every
                # italic run is short (≤50 chars) — those are almost certainly
                # inline titles (books, journals) rather than problematic emphasis.
                if italic_ratio > 0.5:
                    longest_italic_run = max(
                        (len(r.text) for r in para.runs if r.italic and r.text),
                        default=0
                    )
                    if longest_italic_run > 50:
                        preview = text[:60] + "..." if len(text) > 60 else text
                        issue = AccessibilityIssue(
                            issue_type="EXCESSIVE_ITALIC",
                            description=f"Paragraph has {int(italic_ratio*100)}% italic text. If this is a heading or section title, use a Heading style instead. If intended for emphasis, use italics more selectively (large blocks of italic text are harder to read): \"{preview}\"",
                            location=para_info.location,
                            para_info=para_info
                        )
                        issues.append(issue)

                # Flag excessive underline (>30%) - lower threshold since underline should be rare
                if underline_ratio > 0.3:
                    preview = text[:60] + "..." if len(text) > 60 else text
                    issue = AccessibilityIssue(
                        issue_type="EXCESSIVE_UNDERLINE",
                        description=f"Paragraph has {int(underline_ratio*100)}% underlined text. In digital documents, underlining signals a hyperlink — readers may try to click it. Use bold or italic for emphasis instead: \"{preview}\"",
                        location=para_info.location,
                        para_info=para_info
                    )
                    issues.append(issue)

                # Flag inconsistent formatting (many switches relative to paragraph length)
                # More than 8 formatting switches in a paragraph suggests inconsistent use
                if formatting_switches > 8 and len(para.runs) > 10:
                    # Suppress for reading lists / bibliographies:
                    # nearly all switches are pure italic toggling (e.g. author + italic title)
                    # and the paragraph is comma-separated (many items in a list)
                    comma_count = text.count(',')
                    is_formatted_list = (
                        italic_only_switches >= formatting_switches - 1
                        and comma_count >= 5
                    )
                    if not is_formatted_list:
                        preview = text[:60] + "..." if len(text) > 60 else text
                        issue = AccessibilityIssue(
                            issue_type="INCONSISTENT_FORMATTING",
                            description=f"Paragraph has inconsistent formatting ({formatting_switches} format changes). Pick one emphasis style (bold or italic) and apply it consistently — avoid switching between multiple styles within the same paragraph: \"{preview}\"",
                            location=para_info.location,
                            para_info=para_info
                        )
                        issues.append(issue)

        return issues

    def check_images_alt_text(self) -> List[AccessibilityIssue]:
        """Detect images missing alt text"""
        issues = []

        # Find all inline shapes (images) in the document
        for para_info in self.all_paragraphs:
            para = para_info.paragraph

            # Check for inline shapes (images) in the paragraph
            try:
                para_element = para._element
                # Look for drawing elements (modern images)
                drawings = para_element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing')

                for drawing in drawings:
                    # Look for alt text in the docPr element
                    has_alt_text = False
                    docPr = drawing.find('.//{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}docPr')

                    if docPr is not None:
                        # Check for title and description
                        title = docPr.get('title', '')
                        descr = docPr.get('descr', '')

                        if title or descr:
                            has_alt_text = True

                    if not has_alt_text:
                        issue = AccessibilityIssue(
                            issue_type="IMAGE_MISSING_ALT",
                            description="Image is missing alt text (right-click image > Edit Alt Text)",
                            location=para_info.location,
                            para_info=para_info
                        )
                        issues.append(issue)

                # Also check for older pict elements
                picts = para_element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pict')
                for pict in picts:
                    # Old format images often lack alt text
                    # Check for title attribute
                    has_alt_text = False
                    shape = pict.find('.//{urn:schemas-microsoft-com:vml}shape')
                    if shape is not None:
                        alt = shape.get('alt', '')
                        title = shape.get('title', '')
                        if alt or title:
                            has_alt_text = True

                    if not has_alt_text:
                        issue = AccessibilityIssue(
                            issue_type="IMAGE_MISSING_ALT",
                            description="Image is missing alt text (right-click image > Edit Alt Text)",
                            location=para_info.location,
                            para_info=para_info
                        )
                        issues.append(issue)

            except:
                pass  # If we can't parse images, skip

        return issues

    def check_decorative_images(self) -> List[AccessibilityIssue]:
        """Identify decorative images missing decorative marking"""
        issues = []

        # Find all inline shapes (images) in the document
        for para_info in self.all_paragraphs:
            para = para_info.paragraph

            try:
                para_element = para._element
                drawings = para_element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing')

                for drawing in drawings:
                    # Check if image has alt text
                    has_alt_text = False
                    is_marked_decorative = False

                    docPr = drawing.find('.//{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}docPr')

                    if docPr is not None:
                        title = docPr.get('title', '')
                        descr = docPr.get('descr', '')

                        # Check if marked as decorative (empty title AND empty description)
                        # In Word, decorative images should have both set to empty string
                        if title == '' and descr == '':
                            is_marked_decorative = True
                        elif title or descr:
                            has_alt_text = True

                    # Heuristic: if image has no alt text and is not marked decorative,
                    # it might be decorative but not marked properly
                    # We can't definitively say it's decorative, so we'll flag for review
                    # Only flag if it doesn't have alt text (already covered by other check)
                    # This check is for images that ARE marked decorative but shouldn't be

                    # Actually, let's flag images that have empty alt text but might not be decorative
                    # Check if the image is likely meaningful by its context
                    if is_marked_decorative:
                        # Check surrounding text for context that suggests image is meaningful
                        para_text = para.text.strip()

                        # Keywords that suggest the image might be meaningful
                        meaningful_keywords = ['figure', 'diagram', 'chart', 'graph', 'screenshot',
                                             'illustration', 'photo', 'image', 'picture', 'logo']

                        # If paragraph mentions the image, it's probably not decorative
                        if any(keyword in para_text.lower() for keyword in meaningful_keywords):
                            issue = AccessibilityIssue(
                                issue_type="DECORATIVE_IMAGE_QUESTIONABLE",
                                description=f"Image marked as decorative but context suggests it may be meaningful: \"{para_text[:50]}...\"",
                                location=para_info.location,
                                para_info=para_info
                            )
                            issues.append(issue)

            except:
                pass

        return issues

    def check_image_text_content(self) -> List[AccessibilityIssue]:
        """Detect images that might contain text content (screenshots, schedules, tables)"""
        issues = []

        # Look for images with suspiciously large file sizes (likely screenshots)
        # or images in contexts that suggest they contain text
        for para_info in self.all_paragraphs:
            para = para_info.paragraph
            para_text = para.text.strip().lower()

            try:
                para_element = para._element
                drawings = para_element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing')

                for drawing in drawings:
                    # Check context - if nearby text mentions schedule, calendar, table, etc.
                    # this image might be a screenshot
                    context_keywords = ['schedule', 'calendar', 'table', 'dates', 'timeline',
                                       'grading scale', 'rubric', 'assignment', 'due date']

                    # Check alt text for hints
                    docPr = drawing.find('.//{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}docPr')
                    alt_text = ""
                    if docPr is not None:
                        alt_text = docPr.get('title', '') + " " + docPr.get('descr', '')

                    # Check if image appears in a context suggesting it contains text
                    suspicious = False
                    reason = ""

                    # Check nearby text (current paragraph)
                    if any(keyword in para_text for keyword in context_keywords):
                        suspicious = True
                        reason = f"Image appears near text mentioning '{next(k for k in context_keywords if k in para_text)}'"

                    # Check alt text for keywords suggesting text content
                    alt_lower = alt_text.lower()
                    if any(keyword in alt_lower for keyword in ['screenshot', 'table', 'schedule', 'calendar']):
                        suspicious = True
                        reason = "Alt text suggests image contains structured data"

                    if suspicious:
                        issue = AccessibilityIssue(
                            issue_type="IMAGE_TEXT_CONTENT",
                            description=f"Image may contain text content that should be provided as actual text. {reason}. Consider using a real table or text list instead.",
                            location=para_info.location,
                            para_info=para_info
                        )
                        issues.append(issue)

            except:
                pass

        return issues

    def check_date_formats(self) -> List[AccessibilityIssue]:
        """Detect unclear date formats (numeric-only dates)"""
        issues = []

        import re

        # Patterns for numeric-only dates (ambiguous)
        # Examples: 12/5/2024, 5-12-24, 12.5.24
        numeric_date_pattern = r'\b\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}\b'

        for para_info in self.all_paragraphs:
            text = para_info.paragraph.text

            # Find all numeric date patterns
            matches = re.finditer(numeric_date_pattern, text)
            found_numeric_dates = []

            for match in matches:
                date_str = match.group()
                # Skip if it looks like a version number (e.g., 1.2.3)
                if not any(sep in date_str for sep in ['-', '/']):
                    continue
                found_numeric_dates.append(date_str)

            if found_numeric_dates:
                # Remove duplicates
                unique_dates = list(set(found_numeric_dates))
                if unique_dates:
                    example_dates = ', '.join(unique_dates[:3])
                    if len(unique_dates) > 3:
                        example_dates += f" (and {len(unique_dates) - 3} more)"

                    issue = AccessibilityIssue(
                        issue_type="NUMERIC_DATE_FORMAT",
                        description=f"Numeric date format found: {example_dates}. Use clear formats like 'January 15, 2024' or '15 Jan 2024' to avoid confusion.",
                        location=para_info.location,
                        para_info=para_info
                    )
                    issues.append(issue)

        return issues

    def check_merged_cells_layout(self) -> List[AccessibilityIssue]:
        """Identify merged table cells used to create visual layouts (already covered by check_table_merged_cells)"""
        # This is already implemented in check_table_merged_cells()
        # That function detects merged cells that may impact reading order
        # We'll enhance the existing function's description to emphasize layout issues
        return []

    def check_copied_content_broken_styles(self) -> List[AccessibilityIssue]:
        """Identify copied content with broken or inconsistent style application"""
        issues = []

        for para_info in self.all_paragraphs:
            para = para_info.paragraph
            text = para.text.strip()

            if not text:
                continue

            # Check for runs with different fonts within a single paragraph
            # that might indicate copied content with broken styles
            runs = para.runs
            if len(runs) > 1:
                fonts_in_para = []
                sizes_in_para = []

                for run in runs:
                    if run.text.strip():  # Only consider non-empty runs
                        font_name = run.font.name if run.font.name else 'Default'
                        font_size = run.font.size.pt if run.font.size else 0
                        fonts_in_para.append(font_name)
                        sizes_in_para.append(font_size)

                # Check if there are multiple different fonts/sizes in the paragraph
                unique_fonts = set(fonts_in_para)
                unique_sizes = set(sizes_in_para)

                # Flag if there are 3+ different fonts or unusual font mixing
                if len(unique_fonts) >= 3:
                    issue = AccessibilityIssue(
                        issue_type="BROKEN_STYLE_COPIED_CONTENT",
                        description=f"Paragraph contains {len(unique_fonts)} different fonts ({', '.join(list(unique_fonts)[:3])}), suggesting copied content with inconsistent formatting.",
                        location=para_info.location,
                        para_info=para_info
                    )
                    issues.append(issue)
                # Also flag if there's significant font size variation (3+ sizes)
                elif len(unique_sizes) >= 3:
                    size_list = [str(int(s)) + 'pt' for s in unique_sizes if s > 0]
                    if len(size_list) >= 3:
                        issue = AccessibilityIssue(
                            issue_type="BROKEN_STYLE_COPIED_CONTENT",
                            description=f"Paragraph contains {len(size_list)} different font sizes ({', '.join(size_list[:3])}), suggesting copied content with inconsistent formatting.",
                            location=para_info.location,
                            para_info=para_info
                        )
                        issues.append(issue)

        return issues

    def check_footnotes(self) -> List[AccessibilityIssue]:
        """Flag footnotes used instead of inline explanations"""
        issues = []

        import re

        for para_info in self.all_paragraphs:
            para = para_info.paragraph
            text = para.text.strip()

            if not text:
                continue

            has_footnote_marker = False

            # Check for actual footnote markers (asterisks, daggers, etc.)
            # These are rarely used except for footnotes
            if re.search(r'[\*†‡§¶#](?!\w)', text):
                has_footnote_marker = True

            # Check for bracketed/parenthesized numbers at the END of sentences or paragraphs
            # (typical footnote position), NOT in the middle of comma-separated lists
            # Pattern: [1] or (1) appearing after period/exclamation/question mark, or at end of text
            # Do NOT match after commas or semicolons (used in lists)
            if re.search(r'[.!?]\s*[\[\(]\d+[\]\)]|[\[\(]\d+[\]\)]\s*$', text):
                has_footnote_marker = True

            # Check for superscript numbers that look like footnote references
            # (superscript single digit or double digit at end of text)
            for run in para.runs:
                if run.font.superscript and run.text.strip():
                    superscript_text = run.text.strip()
                    # Only flag if it's a standalone number (likely footnote reference)
                    if superscript_text.isdigit() and len(superscript_text) <= 2:
                        has_footnote_marker = True
                        break

            if has_footnote_marker:
                preview = text[:80] + "..." if len(text) > 80 else text
                issue = AccessibilityIssue(
                    issue_type="FOOTNOTE_USAGE",
                    description=f"Footnote marker detected: \"{preview}\". Consider using inline explanations or endnotes for better accessibility.",
                    location=para_info.location,
                    para_info=para_info
                )
                issues.append(issue)

        return issues

    def check_visual_indicators(self) -> List[AccessibilityIssue]:
        """Detect visual indicators of due dates or significance without text equivalents"""
        issues = []

        import re

        for para_info in self.all_paragraphs:
            para = para_info.paragraph
            text = para.text.strip()

            if not text:
                continue

            # Check for symbols that might indicate importance/due dates without text
            # Common patterns: ★, ☆, •, ►, ►, ⚠, ⚡, ✓, ✗, etc.
            symbol_pattern = r'[★☆•►▸⚠⚡✓✗✘✔✖❗❢❣⭐🔴🟡🟢⬤●○◆◇■□▪▫]+'

            if re.search(symbol_pattern, text):
                # Check if there's explanatory text nearby
                words_after_symbol = re.split(symbol_pattern, text)

                for segment in words_after_symbol:
                    segment = segment.strip()
                    # If the segment is very short or just a date/number, it's likely missing explanation
                    if segment and len(segment.split()) < 3:
                        preview = text[:80] + "..." if len(text) > 80 else text
                        issue = AccessibilityIssue(
                            issue_type="VISUAL_INDICATOR_NO_TEXT",
                            description=f"Visual symbol used without clear text explanation: \"{preview}\". Add descriptive text (e.g., 'Important:', 'Due date:', 'Required:').",
                            location=para_info.location,
                            para_info=para_info
                        )
                        issues.append(issue)
                        break

        return issues

    def check_math_expressions(self) -> List[AccessibilityIssue]:
        """Identify math expressions not using accessible markup (MathML or equivalent)"""
        issues = []

        import re

        for para_info in self.all_paragraphs:
            para = para_info.paragraph
            text = para.text.strip()

            if not text:
                continue

            # Skip common false positives
            # Skip if it looks like a grade (e.g., "86/100", "90/100")
            if re.search(r'\b\d{1,3}/\d{2,3}\b', text):
                # But only skip if it's in a grading context
                if re.search(r'\b(grade|score|point|percent|%)\b', text, re.IGNORECASE):
                    continue

            # Skip chapter/page ranges (e.g., "chapters 1-28", "pages 10-20")
            if re.search(r'\b(chapter|page|pp?\.?)\s*\d+\s*[-–]\s*\d+\b', text, re.IGNORECASE):
                continue

            # Remove phone numbers before checking math patterns so that
            # NXX-NXX-XXXX formats don't look like arithmetic chains.
            text = re.sub(r'\(?\d{3}\)?[\s.\-]\d{3}[\s.\-]\d{4}', '', text)

            is_math = False

            # Check for variable equations: x = 2 + 3
            if re.search(r'\b[a-z]\s*=\s*[0-9a-z\+\-\*/\(\)]+', text, re.IGNORECASE):
                is_math = True

            # Check for math functions/symbols (high confidence)
            if re.search(r'\b(sin|cos|tan|log|ln|sqrt|sum|integral|∫|∑|π|α|β|γ|θ|μ|σ|Δ)\b', text, re.IGNORECASE):
                is_math = True

            # Check for subscripts/superscripts (Unicode)
            if re.search(r'[₀₁₂₃₄₅₆₇₈₉]{2,}|[⁰¹²³⁴⁵⁶⁷⁸⁹]{2,}', text):
                is_math = True

            # Check for ASCII superscript/subscript notation (e.g., x^2, H_2O with multiple uses)
            if re.search(r'(\^[0-9]+|_[0-9]+).*(\^[0-9]+|_[0-9]+)', text):
                is_math = True

            # Check for math relation symbols
            if re.search(r'[≤≥≠≈∞∝∈∉⊂⊃∪∩]', text):
                is_math = True

            # Check for complex arithmetic operations (more than simple division)
            # Must have multiple operations or parentheses to be flagged
            arithmetic_match = re.search(r'[0-9]+\s*[\+\-\*÷×]\s*[0-9]+\s*[\+\-\*÷×]\s*[0-9]+', text)
            if arithmetic_match:
                is_math = True

            # Check for expressions with parentheses (e.g., "(2 + 3) * 4")
            if re.search(r'\([0-9\s\+\-\*/]+\)\s*[\+\-\*÷×]', text):
                is_math = True

            if is_math:
                preview = text[:80] + "..." if len(text) > 80 else text
                issue = AccessibilityIssue(
                    issue_type="MATH_NO_ACCESSIBLE_MARKUP",
                    description=f"Mathematical expression detected without accessible markup: \"{preview}\". Use equation editors or MathML for complex math.",
                    location=para_info.location,
                    para_info=para_info
                )
                issues.append(issue)

        return issues

    def run_all_checks(self):
        """Run all accessibility checks and collect issues"""
        self.issues = []
        # Heading checks disabled - better suited for LLM analysis
        # self.issues.extend(self.check_unstyled_headings())
        # self.issues.extend(self.check_heading_hierarchy())

        # Algorithmic checks
        self.issues.extend(self.check_table_headers())
        self.issues.extend(self.check_layout_tables())
        self.issues.extend(self.check_empty_table_rows_columns())
        self.issues.extend(self.check_table_scope_declarations())
        self.issues.extend(self.check_table_captions())
        self.issues.extend(self.check_table_merged_cells())
        self.issues.extend(self.check_table_numeric_alignment())
        self.issues.extend(self.check_table_reading_order())
        self.issues.extend(self.check_table_color_coded_meaning())
        self.issues.extend(self.check_table_embedded_images())
        self.issues.extend(self.check_manual_alignment())
        self.issues.extend(self.check_font_sizes())
        self.issues.extend(self.check_decorative_fonts())
        self.issues.extend(self.check_inconsistent_fonts())
        self.issues.extend(self.check_color_contrast())
        self.issues.extend(self.check_color_only_meaning())
        self.issues.extend(self.check_text_over_backgrounds())
        self.issues.extend(self.check_document_metadata())
        self.issues.extend(self.check_document_language())
        self.issues.extend(self.check_multilingual_content())
        self.issues.extend(self.check_underline_non_links())
        self.issues.extend(self.check_line_spacing())
        self.issues.extend(self.check_full_justification())
        self.issues.extend(self.check_all_caps_blocks())
        self.issues.extend(self.check_non_descriptive_links())
        # Removed checks (pared back to critical issues):
        # self.issues.extend(self.check_unstyled_links())  # Unstyled links
        self.issues.extend(self.check_long_urls())
        # PDF link check removed - users typically can't control accessibility of institutional PDFs
        # self.issues.extend(self.check_pdf_links())
        # self.issues.extend(self.check_missing_toc())  # Missing table of contents
        # self.issues.extend(self.check_missing_bookmarks())  # Missing bookmarks/internal navigation
        # self.issues.extend(self.check_nested_list_hierarchy())  # Inconsistent list hierarchy
        # self.issues.extend(self.check_layout_lists())  # Lists used for layout
        # self.issues.extend(self.check_long_sentences())  # Overly long sentences
        self.issues.extend(self.check_excessive_formatting())
        self.issues.extend(self.check_images_alt_text())
        self.issues.extend(self.check_decorative_images())
        self.issues.extend(self.check_image_text_content())
        # self.issues.extend(self.check_date_formats())  # Numeric date formats
        self.issues.extend(self.check_copied_content_broken_styles())
        self.issues.extend(self.check_footnotes())
        self.issues.extend(self.check_visual_indicators())
        self.issues.extend(self.check_math_expressions())

    def create_marked_document(self, output_path: str, missing_sections: List[str] = None,
                               growth_mindset_analysis: Dict[str, str] = None,
                               quality_analysis: Dict[str, str] = None):
        """Create a copy of the document with issues marked via comments and highlighting,
        plus sections for missing required content and growth mindset recommendations"""
        marked_doc = Document(self.target_path)

        # UWM brand colors for each area — used as paragraph background shading
        # on section headings, so we can use the exact brand hex values without
        # worrying about Word's limited highlight palette.
        COLOR_ACCESSIBILITY_HEX  = '035056'   # dark teal  — white text
        COLOR_MISSING_HEX        = 'F79651'   # orange     — black text
        COLOR_TONE_HEX           = '14364D'   # dark navy  — white text
        COLOR_QUALITY_HEX        = 'CFD25B'   # lime       — black text

        COLOR_ACCESSIBILITY_TEXT = RGBColor(255, 255, 255)
        COLOR_MISSING_TEXT       = RGBColor(0,   0,   0)
        COLOR_TONE_TEXT          = RGBColor(255, 255, 255)
        COLOR_QUALITY_TEXT       = RGBColor(0,   0,   0)

        def add_markdown_runs(para, text, font_size=11, base_color=None):
            """Split text on **...** markers and add runs with proper bold formatting."""
            import re
            parts = re.split(r'\*\*', text)
            for i, part in enumerate(parts):
                if not part:
                    continue
                run = para.add_run(part)
                run.font.size = Pt(font_size)
                if i % 2 == 1:   # odd index → inside ** ** → bold
                    run.bold = True
                if base_color:
                    run.font.color.rgb = base_color

        def shade_paragraph(para, hex_color):
            """Apply a solid background fill to a paragraph using XML shading."""
            from docx.oxml.ns import qn as _qn
            from docx.oxml import OxmlElement
            pPr = para._p.get_or_add_pPr()
            existing = pPr.find(_qn('w:shd'))
            if existing is not None:
                pPr.remove(existing)
            shd = OxmlElement('w:shd')
            shd.set(_qn('w:val'),   'clear')
            shd.set(_qn('w:color'), 'auto')
            shd.set(_qn('w:fill'),  hex_color.upper())
            pPr.append(shd)

        def add_hyperlink(para, text, url):
            """Insert a live, clickable hyperlink run into a paragraph."""
            r_id = para.part.relate_to(
                url,
                'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
                is_external=True,
            )
            hl = OxmlElement('w:hyperlink')
            hl.set(qn('r:id'), r_id)
            r = OxmlElement('w:r')
            rPr = OxmlElement('w:rPr')
            rStyle = OxmlElement('w:rStyle')
            rStyle.set(qn('w:val'), 'Hyperlink')
            rPr.append(rStyle)
            r.append(rPr)
            t = OxmlElement('w:t')
            t.text = text
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            r.append(t)
            hl.append(r)
            para._p.append(hl)

        def add_html_line_runs(para, html_line, font_size=10, base_color=None):
            """Add runs to *para* from an HTML snippet, turning <a href="..."> into live links."""
            segments = re.split(r'(<a\s[^>]*>.*?</a>)', html_line, flags=re.DOTALL)
            for seg in segments:
                m = re.match(r'<a\s+href="([^"]+)"[^>]*>(.*?)</a>', seg, re.DOTALL)
                if m:
                    url = m.group(1)
                    link_text = html.unescape(re.sub(r'<[^>]+>', '', m.group(2)))
                    if link_text:
                        add_hyperlink(para, link_text, url)
                else:
                    text = html.unescape(re.sub(r'<[^>]+>', '', seg))
                    if text:
                        run = para.add_run(text)
                        run.font.size = Pt(font_size)
                        if base_color:
                            run.font.color.rgb = base_color

        # Store parameters for later use
        self._missing_sections = missing_sections or []
        self._growth_mindset_analysis = growth_mindset_analysis or {}

        # Running counter for Word comment IDs (must be unique integers)
        comment_id_counter = [0]

        def next_comment_id():
            cid = comment_id_counter[0]
            comment_id_counter[0] += 1
            return cid

        # Group issues by location
        para_issues_by_location = {}
        table_issues = {}

        for issue in self.issues:
            if issue.para_info is not None:
                # For paragraph issues, group by table location if applicable
                key = (issue.para_info.table_idx, issue.para_info.row_idx, issue.para_info.cell_idx)
                if key not in para_issues_by_location:
                    para_issues_by_location[key] = []
                para_issues_by_location[key].append(issue)

            if issue.table_index is not None:
                if issue.table_index not in table_issues:
                    table_issues[issue.table_index] = []
                table_issues[issue.table_index].append(issue)

        # Mark top-level paragraphs (not in tables) — highlight + Word comment
        for (table_idx, row_idx, cell_idx), issues_list in para_issues_by_location.items():
            if table_idx is None:
                for issue in issues_list:
                    original_text = issue.para_info.paragraph.text.strip()

                    for para in marked_doc.paragraphs:
                        para_text = para.text.strip()

                        if para_text and (para_text in original_text or original_text in para_text):
                            # Add a Word comment — the speech-bubble annotation is
                            # sufficient to mark the issue; no body-text highlight needed.
                            label = self.ISSUE_DISPLAY_LABELS.get(issue.issue_type, issue.issue_type.replace('_', ' ').title())
                            comment_text = f"{label}\n{issue.description}"
                            _add_word_comment(marked_doc, para, next_comment_id(), comment_text)
                            break

        # Mark paragraphs in table cells — Word comment only
        for (table_idx, row_idx, cell_idx), issues_list in para_issues_by_location.items():
            if table_idx is not None and table_idx < len(marked_doc.tables):
                table = marked_doc.tables[table_idx]
                if row_idx < len(table.rows) and cell_idx < len(table.rows[row_idx].cells):
                    cell = table.rows[row_idx].cells[cell_idx]

                    for issue in issues_list:
                        original_text = issue.para_info.paragraph.text.strip()

                        for para in cell.paragraphs:
                            para_text = para.text.strip()

                            if para_text and (para_text in original_text or original_text in para_text):
                                label = self.ISSUE_DISPLAY_LABELS.get(issue.issue_type, issue.issue_type.replace('_', ' ').title())
                                comment_text = f"{label}\n{issue.description}"
                                _add_word_comment(marked_doc, para, next_comment_id(), comment_text)
                                break

        # Mark tables with issues — one comment on the first cell per table
        for table_idx, issues_list in table_issues.items():
            if table_idx < len(marked_doc.tables):
                table = marked_doc.tables[table_idx]

                if table.rows and table.rows[0].cells:
                    first_cell = table.rows[0].cells[0]

                    # Group by type to build a concise comment
                    issue_summary = {}
                    for issue in issues_list:
                        issue_summary.setdefault(issue.issue_type, []).append(issue.description)

                    issue_count = len(issues_list)
                    if issue_count == 1:
                        label = self.ISSUE_DISPLAY_LABELS.get(issues_list[0].issue_type, issues_list[0].issue_type.replace('_', ' ').title())
                        comment_text = f"TABLE ISSUE\n{label}: {issues_list[0].description}"
                    else:
                        lines = [f"TABLE HAS {issue_count} ISSUES:"]
                        for itype, descs in issue_summary.items():
                            label = self.ISSUE_DISPLAY_LABELS.get(itype, itype.replace('_', ' ').title())
                            extra = f" (+{len(descs)-1} more)" if len(descs) > 1 else ""
                            lines.append(f"{label}: {descs[0]}{extra}")
                        comment_text = '\n'.join(lines)

                    anchor_para = first_cell.paragraphs[0] if first_cell.paragraphs else first_cell.add_paragraph()
                    _add_word_comment(marked_doc, anchor_para, next_comment_id(), comment_text)

        # Add document-level accessibility issues (no paragraph/table anchor)
        doc_level_issues = [
            issue for issue in self.issues
            if issue.para_info is None and issue.table_index is None
        ]
        if not self.issues:
            marked_doc.add_paragraph()
            banner_para = marked_doc.add_paragraph()
            banner_run = banner_para.add_run("  ACCESSIBILITY — NO ISSUES DETECTED  ")
            banner_run.bold = True
            banner_run.font.size = Pt(14)
            banner_run.font.color.rgb = RGBColor(0, 100, 0)
            shade_paragraph(banner_para, 'D4EDDA')  # light green
            desc_para = marked_doc.add_paragraph()
            desc_run = desc_para.add_run("Your syllabus passed all accessibility checks.")
            desc_run.font.size = Pt(10)
        elif doc_level_issues:
            marked_doc.add_paragraph()
            banner_para = marked_doc.add_paragraph()
            banner_run = banner_para.add_run("  ACCESSIBILITY — DOCUMENT-LEVEL ISSUES  ")
            banner_run.bold = True
            banner_run.font.size = Pt(14)
            banner_run.font.color.rgb = COLOR_ACCESSIBILITY_TEXT
            shade_paragraph(banner_para, COLOR_ACCESSIBILITY_HEX)

            for issue in doc_level_issues:
                marked_doc.add_paragraph()

                # Issue type as subheading
                sub_para = marked_doc.add_paragraph()
                label = self.ISSUE_DISPLAY_LABELS.get(issue.issue_type, issue.issue_type.replace('_', ' ').title())
                sub_run = sub_para.add_run(f"  {label}  ")
                sub_run.bold = True
                sub_run.font.size = Pt(11)
                sub_run.font.color.rgb = COLOR_ACCESSIBILITY_TEXT
                shade_paragraph(sub_para, COLOR_ACCESSIBILITY_HEX)

                # Description
                desc_para = marked_doc.add_paragraph()
                desc_run = desc_para.add_run(issue.description)
                desc_run.font.size = Pt(10)

        # Add missing sections block at the end of the document
        if self._missing_sections:
            # Primary "MISSING SECTIONS" banner
            marked_doc.add_paragraph()
            banner_para = marked_doc.add_paragraph()
            banner_run = banner_para.add_run("  MISSING SECTIONS  ")
            banner_run.bold = True
            banner_run.font.size = Pt(14)
            banner_run.font.color.rgb = COLOR_MISSING_TEXT
            shade_paragraph(banner_para, COLOR_MISSING_HEX)

            for section in self._missing_sections:
                marked_doc.add_paragraph()

                # Section name as subheading (smaller, same orange background)
                heading_para = marked_doc.add_paragraph()
                heading_run = heading_para.add_run(f"  {section}  ")
                heading_run.bold = True
                heading_run.font.size = Pt(11)
                heading_run.font.color.rgb = COLOR_MISSING_TEXT
                shade_paragraph(heading_para, COLOR_MISSING_HEX)

                # Convert HTML description to structured Word paragraphs
                raw_desc = self._get_section_description(section)
                # Convert structural HTML to line breaks / bullets (keep <a> tags for live links)
                desc = re.sub(r'<br\s*/?>', '\n', raw_desc)
                desc = re.sub(r'</?ul>', '\n', desc)
                desc = re.sub(r'<li>', '• ', desc)
                desc = re.sub(r'</li>', '\n', desc)

                for line in desc.split('\n'):
                    line = line.strip()
                    if not line:
                        continue
                    p = marked_doc.add_paragraph()
                    if line.startswith('• '):
                        p.paragraph_format.left_indent = Pt(18)
                    add_html_line_runs(p, line, font_size=10)

        # Add full Belonging & Growth Mindset analysis (mirrors the web app display)
        if self._growth_mindset_analysis.get('status') == 'success' and 'analysis' in self._growth_mindset_analysis:
            analysis_text = self._growth_mindset_analysis.get('analysis', '').strip()
            if analysis_text:
                marked_doc.add_paragraph()

                # Section banner
                tone_header_para = marked_doc.add_paragraph()
                tone_header_run = tone_header_para.add_run("  BELONGING & GROWTH MINDSET SUGGESTIONS  ")
                tone_header_run.bold = True
                tone_header_run.font.size = Pt(14)
                tone_header_run.font.color.rgb = COLOR_TONE_TEXT
                shade_paragraph(tone_header_para, COLOR_TONE_HEX)

                subtitle_para = marked_doc.add_paragraph()
                sub_run = subtitle_para.add_run(
                    "AI-powered analysis of how well your syllabus promotes student belonging and growth mindset, "
                    "with ready-to-use suggested language for each area."
                )
                sub_run.italic = True
                sub_run.font.size = Pt(10)
                sub_run.font.color.rgb = RGBColor(80, 80, 80)

                marked_doc.add_paragraph()

                in_code_block = False
                for line in analysis_text.split('\n'):
                    line = line.rstrip()

                    # Toggle code block (copy-paste ready text)
                    if line.startswith('```'):
                        in_code_block = not in_code_block
                        continue

                    if in_code_block:
                        # Render copy-paste text with light teal background
                        para = marked_doc.add_paragraph()
                        run = para.add_run(line if line else ' ')
                        run.font.size = Pt(10)
                        para.paragraph_format.left_indent = Pt(18)
                        shade_paragraph(para, 'D6EEEE')  # very light teal
                        continue

                    if not line:
                        marked_doc.add_paragraph()
                        continue

                    para = marked_doc.add_paragraph()

                    if line.startswith('## '):
                        # Question heading with navy background
                        run = para.add_run(f"  {line[3:].strip()}  ")
                        run.bold = True
                        run.font.size = Pt(12)
                        run.font.color.rgb = COLOR_TONE_TEXT
                        shade_paragraph(para, COLOR_TONE_HEX)
                    elif line.startswith('    - ') or line.startswith('    • '):
                        try:
                            para.style = marked_doc.styles['List Bullet 2']
                        except KeyError:
                            para.paragraph_format.left_indent = Pt(36)
                            para.paragraph_format.first_line_indent = Pt(-18)
                        add_markdown_runs(para, '\u2022 ' + line[6:], font_size=10)
                    elif line.startswith('- ') or line.startswith('• '):
                        try:
                            para.style = marked_doc.styles['List Bullet']
                        except KeyError:
                            para.paragraph_format.left_indent = Pt(18)
                            para.paragraph_format.first_line_indent = Pt(-18)
                        add_markdown_runs(para, '\u2022 ' + line[2:], font_size=10)
                    else:
                        add_markdown_runs(para, line, font_size=10)

        # Add Clarity & Quality analysis as an appendix
        if quality_analysis and quality_analysis.get('status') == 'success':
            analysis_text = quality_analysis.get('analysis', '').strip()
            if analysis_text:
                marked_doc.add_paragraph()

                # Section header with UWM lime background
                header_para = marked_doc.add_paragraph()
                header_run = header_para.add_run("  CLARITY & QUALITY SUGGESTIONS  ")
                header_run.bold = True
                header_run.font.size = Pt(14)
                header_run.font.color.rgb = COLOR_QUALITY_TEXT
                shade_paragraph(header_para, COLOR_QUALITY_HEX)

                subtitle_para = marked_doc.add_paragraph()
                sub_run = subtitle_para.add_run(
                    "AI-identified clarity and quality issues. Review each item and revise as appropriate."
                )
                sub_run.italic = True
                sub_run.font.size = Pt(10)
                sub_run.font.color.rgb = RGBColor(80, 80, 80)

                marked_doc.add_paragraph()

                for line in analysis_text.split('\n'):
                    line = line.rstrip()
                    if not line:
                        continue

                    para = marked_doc.add_paragraph()

                    if line.startswith('## '):
                        # Sub-section heading with UWM lime background
                        run = para.add_run(f"  {line[3:].strip()}  ")
                        run.bold = True
                        run.font.size = Pt(12)
                        run.font.color.rgb = COLOR_QUALITY_TEXT
                        shade_paragraph(para, COLOR_QUALITY_HEX)
                    elif line.startswith('✓'):
                        # All-clear line — green text, no background
                        add_markdown_runs(para, line, base_color=RGBColor(0, 120, 0))
                    else:
                        # Content line — use Word list styles for proper bullet alignment,
                        # falling back to manual indent if the style isn't in this document.
                        if line.startswith('    - ') or line.startswith('    • '):
                            try:
                                para.style = marked_doc.styles['List Bullet 2']
                            except KeyError:
                                para.paragraph_format.left_indent = Pt(36)
                                para.paragraph_format.first_line_indent = Pt(-18)
                            text = '\u2022 ' + line[6:]
                        elif line.startswith('- ') or line.startswith('• '):
                            try:
                                para.style = marked_doc.styles['List Bullet']
                            except KeyError:
                                para.paragraph_format.left_indent = Pt(18)
                                para.paragraph_format.first_line_indent = Pt(-18)
                            text = '\u2022 ' + line[2:]
                        else:
                            text = line
                        add_markdown_runs(para, text)

        marked_doc.save(output_path)

    def analyze_growth_mindset_and_belonging(self) -> Dict[str, str]:
        """
        Use OpenAI GPT API to analyze syllabus for growth mindset and belonging cues.
        Returns a dictionary with analysis for each of the 6 core questions.
        """
        # Extract full syllabus text
        syllabus_text = []

        # Get text from all paragraphs
        for para_info in self.all_paragraphs:
            text = para_info.paragraph.text.strip()
            if text:
                syllabus_text.append(text)

        # Join all text
        full_text = "\n\n".join(syllabus_text)

        # Get API key from environment
        api_key = os.environ.get('OPENAI_API_KEY')
        if not api_key:
            return {
                'error': 'OPENAI_API_KEY environment variable not set. Please configure your API key in .env file to enable growth mindset analysis.',
                'status': 'error'
            }

        try:
            client = OpenAI(api_key=api_key)

            # Create comprehensive prompt based on the provided guidance
            prompt = f"""You are an expert in educational equity and student-centered course design. Your task is to review a course syllabus and evaluate how well it communicates growth mindset and promotes student belonging.

SYLLABUS TEXT:
{full_text}

Please analyze this syllabus according to the following 6 core questions. For each question, provide:
1. A brief assessment (Strong / Moderate / Weak / Not Addressed)
2. Specific evidence from the syllabus (quote relevant passages or note absence)
3. Concrete suggestions for improvement with example language
4. Ready-to-use text that can be copy-pasted directly into the syllabus, with clear guidance on WHERE to place it

**CORE QUESTION 1: Growth Mindset**
Does this syllabus communicate that the instructor has a "growth mindset" rather than a "fixed mindset" about students' abilities?

Growth mindset conveys that students can grow their abilities through effort, strategies, and help-seeking. Look for:
- Language about developing skills and abilities over time
- Assurances that students can succeed with effort and support
- References to learning as a process vs. innate talent
- Avoid fixed mindset language like "this course is only for quick learners"

**CORE QUESTION 2: Normalizing Challenge**
Do the messages communicate that it is normal to be challenged by course material, and that this is not a sign that a student is not capable or does not belong?

Look for:
- Acknowledgment that students may struggle or face setbacks
- Framing challenges as part of learning, not indicators of inability
- Instructor/TA stories about overcoming their own challenges
- Guidance on what to do when struggling (e.g., seek help, use resources)

**CORE QUESTION 3: Instructor Care**
Does the syllabus communicate that the instructor and instructional team care about students' success?

Look for:
- Personal language about caring for student success
- Accessible office hours (called "student drop-in hours")
- Responsive email policies
- Personal info about TAs to make them more approachable
- Acknowledgment of large class sizes with assurances of care

**CORE QUESTION 4: Valuing Diversity**
Does the syllabus communicate that diversity is valued in the classroom?

Look for:
- Land acknowledgments
- Inclusive course conduct policies
- Recognition of diverse student backgrounds and experiences
- Caregiver/family responsibility policies
- Respectful disagreement and civil discourse expectations
- Statements about creating welcoming environment for all

**CORE QUESTION 5: Normalizing Student Challenges**
Does the syllabus normalize challenges that students often face in college and connect students with resources?

Look for:
- Acknowledgment of financial strain (textbook affordability, food insecurity)
- Mental health and counseling services presented without stigma
- Disability/accessibility services normalized
- Work-life balance and nontraditional student support
- Resources presented as standard tools for success, not remedial

**CORE QUESTION 6: Normalizing Academic Support Use**
Does the syllabus communicate that utilizing academic resources is a standard part of succeeding?

Look for:
- Office hours framed as "student drop-in hours"
- Tutoring/academic support presented as normal, not remedial
- Multiple access points (in-person, virtual, flexible timing)
- Course chat/discussion forums for questions
- Language like "most students access these supports"

IMPORTANT FORMATTING INSTRUCTIONS:
- Structure your response with clear headings for each of the 6 core questions
- Use this exact format for each section:

## QUESTION [NUMBER]: [Brief Title]
**Assessment:** [Strong/Moderate/Weak/Not Addressed]

**Evidence:**
[Quote specific passages or note absence]

**Suggestions:**
[Provide 2-3 concrete, actionable improvements with example language]

**Ready-to-Use Text:**
**Where to place it:** [Specify the section of the syllabus - e.g., "Add a new section titled 'Student Support Resources' after the grading policy" or "Insert at the end of the 'Course Description' section" or "Create a new 'Campus Resources' section near the end of the syllabus"]

**Text to insert:**
```
[Provide 2-4 paragraphs of polished, ready-to-use text that the instructor can copy and paste directly into their syllabus. This should be written in first person from the instructor's perspective, be warm and supportive in tone, and incorporate the best practices from the research. Make it specific to the course subject if possible based on the syllabus content.]
```

---

Please provide a thorough, evidence-based analysis that will help the instructor improve their syllabus to better support all students, especially those from historically marginalized groups."""

            # Call OpenAI API
            response = client.chat.completions.create(
                model="gpt-4o",
                messages=[
                    {
                        "role": "system",
                        "content": "You are an expert in educational equity and student-centered course design, specializing in analyzing syllabi for growth mindset and belonging cues based on research from the Student Experience Project."
                    },
                    {
                        "role": "user",
                        "content": prompt
                    }
                ],
                max_tokens=4000,
                temperature=0.7
            )

            # Extract the response text
            analysis_text = response.choices[0].message.content

            return {
                'analysis': analysis_text,
                'status': 'success'
            }

        except Exception as e:
            return {
                'error': f'Error calling OpenAI API: {str(e)}',
                'status': 'error'
            }

    def analyze_syllabus_quality(self) -> Dict[str, str]:
        """
        Use OpenAI GPT API to analyze syllabus for quality issues:
        - Undefined course terminology
        - Tone and inclusivity issues in policy language
        - Policies that may violate accessibility guidance or confuse students
        - Inconsistent formatting across repeated elements
        """
        # Extract full syllabus text with table structure preserved
        full_text = self._document_to_structured_text()

        # Get API key from environment
        api_key = os.environ.get('OPENAI_API_KEY')
        if not api_key:
            return {
                'error': 'OPENAI_API_KEY environment variable not set. Please configure your API key in .env file to enable quality analysis.',
                'status': 'error'
            }

        try:
            client = OpenAI(api_key=api_key)

            prompt = f"""You are an expert in syllabus design, educational accessibility, and inclusive pedagogy. Analyze this course syllabus for quality and clarity issues across four key dimensions.

SYLLABUS TEXT:
{full_text}

Please provide a comprehensive analysis covering these four areas:

## 1. UNDEFINED COURSE TERMINOLOGY

Identify any technical terms, jargon, or course-specific vocabulary that students (especially first-time students in this field) might not understand. Look for:
- Field-specific terminology used without definition
- Acronyms that aren't spelled out on first use
- Assignment types or formats that may be unfamiliar (e.g., "response paper", "think-pair-share", "flipped classroom")
- Institutional terminology that may not be universally understood

For each undefined term found, use this NESTED bullet format:
- **Term**: [the term or phrase]
    - **Why it might be unclear**: [brief explanation]
    - **Suggested fix**: [how to define or clarify it]

If no issues found, state: "✓ All course terminology appears to be clearly defined or would be familiar to students."

## 2. TONE AND INCLUSIVITY IN POLICY LANGUAGE

Examine the tone of policies, rules, and requirements. Look for:
- Punitive or threatening language (focus on consequences rather than support)
- Deficit-based framing (assuming students will fail, cheat, or not try)
- Language that may alienate students from marginalized groups
- Overly legalistic or bureaucratic tone that creates distance
- Policies that assume all students have the same resources, circumstances, or prior knowledge

For each issue found, use this NESTED bullet format:
- **Issue**: [quote the problematic language]
    - **Problem**: [why this language may be problematic]
    - **Better alternative**: [rewritten version with more inclusive, supportive tone]

If no issues found, state: "✓ Policy language uses an inclusive, supportive tone throughout."

## 3. POLICIES THAT MAY CONFUSE OR EXCLUDE

Identify policies or statements that:
- Contradict accessibility best practices or may create barriers for students with disabilities
- Conflict with institutional policies (if identifiable)
- Are vague or ambiguous in ways that could confuse students
- Make assumptions about student circumstances (e.g., access to technology, transportation, childcare)
- May inadvertently exclude certain student populations

For each policy issue, use this NESTED bullet format:
- **Policy**: [quote the policy]
    - **Potential problem**: [what confusion or exclusion this might cause]
    - **Recommendation**: [how to revise or clarify]

If no issues found, state: "✓ All policies appear clear and inclusive."

## 4. INCONSISTENT FORMATTING

A FORMAT is the template or pattern used to express something — NOT the specific values filled in.
A FORMAT INCONSISTENCY exists only when the TEMPLATE differs, not when the values differ.

To test whether something is a format inconsistency, ask: "If I replaced the specific values with placeholders, would the two examples use a different template?" If no — it is NOT an inconsistency.

CRITICAL RULE FOR DATES: Dates like "September 6", "September 13", and "September 28" all use the template [Month Name] [Day Number]. They are NOT inconsistent. The day number having 1 or 2 digits is NOT a format difference — numbers naturally vary in length. Only flag dates if the template itself differs, e.g., "September 6" vs "9/6" vs "Sept. 6th" (different templates).

Examples of TRUE inconsistencies to flag:
- Date FORMATS differ: "Jan 15" vs "1/15" vs "January 15th" (different templates: abbreviated-name vs numeric vs ordinal)
- Headings at the same level using different styles (some bold, some underlined, different sizes)
- Assignment names using inconsistent capitalization (some Title Case, some lowercase)
- Times formatted differently: "2pm" vs "2:00 PM" vs "14:00"

Examples that are NOT inconsistencies (do NOT flag these, and do NOT mention them at all):
- "September 6" and "September 28" — same template [Month Day], different values
- "September 6" and "September 13" — same template [Month Day], different values
- "September 6" appearing in multiple places in the document — same template, same value
- "Assignment 1" and "Assignment 2" — same template, different numbers
- Multiple headings all using the same bold style

If you are unsure whether something is a true format inconsistency, omit it entirely. Only include items that are unambiguously inconsistent.

For each TRUE formatting inconsistency, use this NESTED bullet format:
- **Element type**: [e.g., "dates", "headings", "assignment names"]
    - **Inconsistency**: [describe the actual template variation, showing contrasting examples]
    - **Recommendation**: [suggest a consistent format to use throughout]

If no issues found, state: "✓ Formatting appears consistent throughout the syllabus."

---

**IMPORTANT FORMATTING NOTE**:
- Use NESTED bullets: main item as top-level bullet, sub-items indented with 4 spaces
- Use clear markdown formatting with ## for main headings and **bold** for labels
- Be specific and quote actual text from the syllabus
- Prioritize the most significant issues in each category
- If you find 5+ issues in a category, focus on the most important ones
- DO NOT include a section on heading structure - that is covered separately in accessibility checks
"""

            response = client.chat.completions.create(
                model="gpt-4o",
                messages=[
                    {"role": "system", "content": "You are an expert in syllabus design, educational accessibility, and inclusive pedagogy. Provide specific, actionable feedback."},
                    {"role": "user", "content": prompt}
                ],
                max_tokens=3000,
                temperature=0.2
            )

            analysis_text = response.choices[0].message.content

            return {
                'analysis': analysis_text,
                'status': 'success'
            }

        except Exception as e:
            return {
                'error': f'Error calling OpenAI API: {str(e)}',
                'status': 'error'
            }

    def analyze_image_alt_text(self) -> Dict[str, str]:
        """
        Use OpenAI GPT API to analyze image alt text quality:
        - Flag institutional logos without meaningful alt text
        - Identify alt text that is too short, irrelevant, or redundant (e.g., "Image of...")
        """
        # Extract all images with their alt text
        images_data = []

        for para_info in self.all_paragraphs:
            para = para_info.paragraph

            try:
                para_element = para._element

                # Look for modern drawing elements
                drawings = para_element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing')

                for drawing in drawings:
                    docPr = drawing.find('.//{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}docPr')

                    if docPr is not None:
                        title = docPr.get('title', '')
                        descr = docPr.get('descr', '')
                        name = docPr.get('name', 'Unnamed image')

                        # Include images even if they have alt text (we want to check quality)
                        images_data.append({
                            'name': name,
                            'title': title,
                            'description': descr,
                            'location': para_info.location
                        })

                # Also check for older pict elements
                picts = para_element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pict')
                for pict in picts:
                    shape = pict.find('.//{urn:schemas-microsoft-com:vml}shape')
                    if shape is not None:
                        alt = shape.get('alt', '')
                        title = shape.get('title', '')

                        images_data.append({
                            'name': 'Legacy image',
                            'title': title,
                            'description': alt,
                            'location': para_info.location
                        })

            except:
                pass  # Skip if we can't parse images

        # If no images found, return early
        if not images_data:
            return {
                'analysis': 'No images found in the document.',
                'status': 'success'
            }

        # Get API key from environment
        api_key = os.environ.get('OPENAI_API_KEY')
        if not api_key:
            return {
                'error': 'OPENAI_API_KEY environment variable not set. Please configure your API key in .env file to enable image alt text analysis.',
                'status': 'error'
            }

        try:
            client = OpenAI(api_key=api_key)

            # Format images data for the prompt
            images_list = []
            for idx, img in enumerate(images_data, 1):
                alt_text = img['description'] or img['title'] or '(no alt text)'
                images_list.append(f"{idx}. **{img['name']}** - Alt text: \"{alt_text}\" - Location: {img['location']}")

            images_text = "\n".join(images_list)

            prompt = f"""You are an expert in web accessibility and educational document design. Analyze the alt text quality for images in this syllabus.

IMAGES AND THEIR ALT TEXT:
{images_text}

Please analyze these images for accessibility issues in two key areas:

## 1. INSTITUTIONAL LOGOS WITHOUT MEANINGFUL ALT TEXT

Look for images that appear to be institutional logos, university seals, or organizational branding. These often have alt text like:
- Empty or missing alt text
- Just the filename (e.g., "logo.png")
- Generic text like "logo" or "image"
- The institution name without context (e.g., just "Harvard" instead of "Harvard University logo")

**Why this matters**: Screen reader users need to know what the logo represents and its purpose in the document. Simply stating the institution name without context can be confusing.

For each logo with inadequate alt text:
- **Image**: [image name/number]
- **Current alt text**: [what it currently says]
- **Problem**: [why this is insufficient]
- **Recommended alt text**: [specific suggestion, e.g., "University of Michigan logo" or "Michigan State University seal"]

If no logo issues found, state: "✓ All institutional logos have meaningful alt text."

## 2. POOR QUALITY ALT TEXT

Identify alt text that is:
- **Too short**: Single words or very brief text that doesn't convey meaning (e.g., "chart", "graph", "picture")
- **Redundant**: Starts with unnecessary phrases like "Image of...", "Picture of...", "Graphic showing..."
- **Irrelevant**: Describes the image format rather than content (e.g., "JPEG image", "PNG file")
- **Non-descriptive**: Vague text that doesn't convey the image's purpose (e.g., "illustration", "diagram")

**Note**: Decorative images should have empty alt text, but this is rare in syllabi. Focus on informational images.

For each poor quality alt text:
- **Image**: [image name/number]
- **Current alt text**: [what it currently says]
- **Problem**: [specific issue - too short, redundant, irrelevant, or non-descriptive]
- **Recommended alt text**: [improved version that clearly describes the image's content and purpose]

If no issues found, state: "✓ All alt text is clear, concise, and descriptive."

---

**IMPORTANT FORMATTING NOTES**:
- Use clear markdown formatting with ## for main headings and **bold** for labels
- Be specific and reference actual alt text from the images
- Prioritize the most significant issues
- If you find more than 5 issues in a category, focus on the most problematic ones
- Remember: Good alt text is concise but informative, typically 1-2 sentences maximum
"""

            response = client.chat.completions.create(
                model="gpt-4o",
                messages=[
                    {"role": "system", "content": "You are an expert in web accessibility, WCAG guidelines, and alt text best practices. Provide specific, actionable feedback."},
                    {"role": "user", "content": prompt}
                ],
                max_tokens=2000,
                temperature=0.7
            )

            analysis_text = response.choices[0].message.content

            return {
                'analysis': analysis_text,
                'status': 'success',
                'image_count': len(images_data)
            }

        except Exception as e:
            return {
                'error': f'Error calling OpenAI API: {str(e)}',
                'status': 'error'
            }

    def extract_course_metadata(self) -> Dict[str, str]:
        """
        Use OpenAI GPT API to extract course metadata from the syllabus.
        Returns instructor name, department, course subject, and course number.
        """
        # Extract first portion of syllabus text (metadata is usually at the top)
        syllabus_text = []
        for para_info in self.all_paragraphs[:50]:  # First 50 paragraphs should contain header info
            text = para_info.paragraph.text.strip()
            if text:
                syllabus_text.append(text)

        # Also check tables (course info is often in tables)
        for table in self.target_doc.tables[:3]:  # Check first 3 tables
            for row in table.rows:
                for cell in row.cells:
                    text = cell.text.strip()
                    if text:
                        syllabus_text.append(text)

        full_text = "\n".join(syllabus_text)

        # Limit text length
        if len(full_text) > 4000:
            full_text = full_text[:4000]

        # Get API key from environment
        api_key = os.environ.get('OPENAI_API_KEY')
        if not api_key:
            return {
                'status': 'error',
                'error': 'OPENAI_API_KEY not set'
            }

        try:
            client = OpenAI(api_key=api_key)

            prompt = f"""Extract the following information from this course syllabus. Return ONLY a JSON object with these exact keys. If a field cannot be determined, use null.

SYLLABUS TEXT:
{full_text}

Extract:
1. "instructor_name": The instructor's full name (e.g., "Dr. Jane Smith", "Professor John Doe", "Sarah Johnson, PhD")
2. "department": The academic department (e.g., "Computer Science", "English", "Biology", "School of Business")
3. "course_subject": The course subject code/prefix (e.g., "CS", "ENGL", "BIO", "MATH", "PSYCH")
4. "course_number": The course number (e.g., "101", "350", "4400")
5. "course_title": The full course title (e.g., "Introduction to Programming", "American Literature")
6. "semester": The semester/term (e.g., "Fall 2024", "Spring 2025")

Return ONLY valid JSON, no other text. Example:
{{"instructor_name": "Dr. Jane Smith", "department": "Computer Science", "course_subject": "CS", "course_number": "101", "course_title": "Introduction to Programming", "semester": "Fall 2024"}}
"""

            response = client.chat.completions.create(
                model="gpt-4o-mini",  # Use faster/cheaper model for simple extraction
                messages=[
                    {"role": "system", "content": "You are a data extraction assistant. Extract course metadata from syllabi and return only valid JSON."},
                    {"role": "user", "content": prompt}
                ],
                max_tokens=300,
                temperature=0.1  # Low temperature for consistent extraction
            )

            response_text = response.choices[0].message.content.strip()

            # Parse JSON response
            import json
            # Clean up response if needed (remove markdown code blocks)
            if response_text.startswith('```'):
                response_text = response_text.split('```')[1]
                if response_text.startswith('json'):
                    response_text = response_text[4:]
                response_text = response_text.strip()

            metadata = json.loads(response_text)
            metadata['status'] = 'success'
            return metadata

        except json.JSONDecodeError as e:
            return {
                'status': 'error',
                'error': f'Failed to parse metadata response: {str(e)}'
            }
        except Exception as e:
            return {
                'status': 'error',
                'error': f'Error extracting metadata: {str(e)}'
            }

    def generate_report(self) -> str:
        """Generate comprehensive assessment report organized by category"""
        report_lines = []
        report_lines.append("=" * 80)
        report_lines.append("SYLLABUS ASSESSMENT REPORT")
        report_lines.append("=" * 80)
        report_lines.append("")

        # Document stats
        report_lines.append("DOCUMENT STATISTICS")
        report_lines.append("-" * 80)
        report_lines.append(f"Total paragraphs analyzed: {len(self.all_paragraphs)}")
        top_level = sum(1 for p in self.all_paragraphs if p.table_idx is None)
        in_tables = sum(1 for p in self.all_paragraphs if p.table_idx is not None)
        report_lines.append(f"  - Top-level paragraphs: {top_level}")
        report_lines.append(f"  - Paragraphs in tables: {in_tables}")
        report_lines.append(f"Total tables: {len(self.target_doc.tables)}")
        report_lines.append("")

        # Section completeness
        report_lines.append("SECTION COMPLETENESS")
        report_lines.append("-" * 80)
        missing = self.check_missing_sections()
        if missing:
            report_lines.append(f"[X] MISSING {len(missing)} REQUIRED SECTIONS:")
            for section in missing:
                report_lines.append(f"  - {section}")
        else:
            report_lines.append("[OK] All required sections present")
        report_lines.append("")

        # Heading structure
        report_lines.append("ACCESSIBILITY: HEADING STRUCTURE")
        report_lines.append("-" * 80)
        heading_check = self.check_heading_structure()
        if heading_check['counts']:
            report_lines.append("Heading usage:")
            for level, count in sorted(heading_check['counts'].items()):
                report_lines.append(f"  - {level}: {count}")
        if heading_check['issues']:
            for issue in heading_check['issues']:
                report_lines.append(f"[X] {issue}")
        else:
            report_lines.append("[OK] Heading structure looks good")
        report_lines.append("")

        # Heading level recommendations
        report_lines.append("ACCESSIBILITY: HEADING LEVEL RECOMMENDATIONS")
        report_lines.append("-" * 80)
        heading_recommendations = [issue for issue in self.issues
                    if issue.issue_type.startswith("SHOULD BE HEADING")]
        if heading_recommendations:
            report_lines.append(f"[X] Found {len(heading_recommendations)} headings that need to be corrected:")
            by_level = {}
            for issue in heading_recommendations:
                level = issue.issue_type
                if level not in by_level:
                    by_level[level] = []
                by_level[level].append(issue)
            for level in sorted(by_level.keys()):
                report_lines.append(f"\n  {level} ({len(by_level[level])} items):")
                for issue in by_level[level][:5]:
                    report_lines.append(f"    - {issue.location}: {issue.description}")
                if len(by_level[level]) > 5:
                    report_lines.append(f"    ... and {len(by_level[level]) - 5} more")
        else:
            report_lines.append("[OK] All headings are properly styled")
        report_lines.append("")

        # ============================================================
        # CATEGORY 1: FONT USAGE
        # ============================================================
        report_lines.append("=" * 80)
        report_lines.append("CATEGORY: FONT USAGE")
        report_lines.append("=" * 80)
        report_lines.append("")

        # Font sizes
        report_lines.append("ACCESSIBILITY: FONT SIZES")
        report_lines.append("-" * 80)
        font_issues = [issue for issue in self.issues if issue.issue_type == "SMALL_FONT"]
        if font_issues:
            report_lines.append(f"[X] Found {len(font_issues)} paragraph(s) with font size below 11pt:")
            report_lines.append("    (Minimum recommended size is 11pt for accessibility)")
            for issue in font_issues[:5]:
                report_lines.append(f"  - {issue.location}: {issue.description}")
            if len(font_issues) > 5:
                report_lines.append(f"  ... and {len(font_issues) - 5} more")
        else:
            report_lines.append("[OK] All text meets minimum font size requirements")
        report_lines.append("")

        # Decorative fonts
        report_lines.append("ACCESSIBILITY: DECORATIVE/INACCESSIBLE FONTS")
        report_lines.append("-" * 80)
        decorative_font_issues = [issue for issue in self.issues if issue.issue_type == "DECORATIVE_FONT"]
        if decorative_font_issues:
            report_lines.append(f"[X] Found {len(decorative_font_issues)} instance(s) of decorative/hard-to-read fonts:")
            report_lines.append("    (Use simple, readable fonts like Arial, Calibri, or Times New Roman)")
            for issue in decorative_font_issues[:5]:
                report_lines.append(f"  - {issue.location}: {issue.description}")
            if len(decorative_font_issues) > 5:
                report_lines.append(f"  ... and {len(decorative_font_issues) - 5} more")
        else:
            report_lines.append("[OK] All fonts are accessible and readable")
        report_lines.append("")

        # Font consistency
        report_lines.append("ACCESSIBILITY: FONT CONSISTENCY")
        report_lines.append("-" * 80)
        inconsistent_font_issues = [issue for issue in self.issues if issue.issue_type == "INCONSISTENT_FONTS"]
        if inconsistent_font_issues:
            for issue in inconsistent_font_issues:
                report_lines.append(f"[X] {issue.description}")
        else:
            report_lines.append("[OK] Font usage is consistent throughout the document")
        report_lines.append("")

        # ============================================================
        # CATEGORY 2: TABLE STRUCTURE
        # ============================================================
        report_lines.append("=" * 80)
        report_lines.append("CATEGORY: TABLE STRUCTURE")
        report_lines.append("=" * 80)
        report_lines.append("")

        # Table usage
        table_check = self.check_table_usage()
        report_lines.append("ACCESSIBILITY: TABLE USAGE")
        report_lines.append("-" * 80)
        report_lines.append(f"Total tables: {table_check['count']}")
        if table_check['issues']:
            for issue in table_check['issues']:
                report_lines.append(f"[X] {issue}")
        else:
            if table_check['count'] == 0:
                report_lines.append("[OK] No tables (good for accessibility)")
            else:
                report_lines.append("[OK] Table usage appears reasonable")
        report_lines.append("")

        # Empty table rows/columns
        report_lines.append("ACCESSIBILITY: EMPTY TABLE ROWS/COLUMNS")
        report_lines.append("-" * 80)
        empty_row_issues = [issue for issue in self.issues if issue.issue_type == "EMPTY_TABLE_ROW"]
        empty_col_issues = [issue for issue in self.issues if issue.issue_type == "EMPTY_TABLE_COLUMN"]
        if empty_row_issues or empty_col_issues:
            if empty_row_issues:
                report_lines.append(f"[X] Found {len(empty_row_issues)} empty table row(s) (used for spacing):")
                for issue in empty_row_issues:
                    report_lines.append(f"  - {issue.description}")
            if empty_col_issues:
                report_lines.append(f"[X] Found {len(empty_col_issues)} empty table column(s) (used for spacing):")
                for issue in empty_col_issues:
                    report_lines.append(f"  - {issue.description}")
            report_lines.append("    (Use proper table spacing/padding instead of empty rows/columns)")
        else:
            if len(self.target_doc.tables) > 0:
                report_lines.append("[OK] No empty table rows or columns detected")
            else:
                report_lines.append("[OK] No tables to check")
        report_lines.append("")

        # Layout tables
        report_lines.append("ACCESSIBILITY: LAYOUT vs. DATA TABLES")
        report_lines.append("-" * 80)
        layout_issues = [issue for issue in self.issues if issue.issue_type == "LAYOUT_TABLE"]
        if layout_issues:
            report_lines.append(f"[X] Found {len(layout_issues)} tables used for layout (should use headings/paragraphs):")
            for issue in layout_issues:
                report_lines.append(f"  - {issue.description}")
        else:
            if table_check['count'] > 0:
                report_lines.append("[OK] Tables appear to contain tabular data (not layout)")
            else:
                report_lines.append("[OK] No tables to check")
        report_lines.append("")

        # Table headers
        report_lines.append("ACCESSIBILITY: TABLE HEADERS")
        report_lines.append("-" * 80)
        table_header_issues = [issue for issue in self.issues if issue.issue_type == "TABLE_NO_HEADER"]
        if table_header_issues:
            report_lines.append(f"[X] Found {len(table_header_issues)} tables without proper headers:")
            for issue in table_header_issues:
                report_lines.append(f"  - {issue.description}")
        else:
            if table_check['count'] > 0:
                report_lines.append("[OK] All tables appear to have headers")
            else:
                report_lines.append("[OK] No tables to check")
        report_lines.append("")

        # Table scope declarations
        report_lines.append("ACCESSIBILITY: TABLE SCOPE DECLARATIONS")
        report_lines.append("-" * 80)
        scope_issues = [issue for issue in self.issues if issue.issue_type == "TABLE_MISSING_SCOPE"]
        if scope_issues:
            report_lines.append(f"[X] Found {len(scope_issues)} table(s) missing proper scope declarations:")
            for issue in scope_issues:
                report_lines.append(f"  - {issue.description}")
        else:
            if len(self.target_doc.tables) > 0:
                report_lines.append("[OK] Table scope declarations are in place")
            else:
                report_lines.append("[OK] No tables to check")
        report_lines.append("")

        # Table captions
        report_lines.append("ACCESSIBILITY: TABLE CAPTIONS/DESCRIPTIONS")
        report_lines.append("-" * 80)
        caption_issues = [issue for issue in self.issues if issue.issue_type == "TABLE_MISSING_CAPTION"]
        if caption_issues:
            report_lines.append(f"[X] Found {len(caption_issues)} table(s) missing captions or descriptions:")
            for issue in caption_issues:
                report_lines.append(f"  - {issue.description}")
        else:
            if len(self.target_doc.tables) > 0:
                report_lines.append("[OK] All tables have captions or descriptions")
            else:
                report_lines.append("[OK] No tables to check")
        report_lines.append("")

        # Table merged cells
        report_lines.append("ACCESSIBILITY: TABLE MERGED CELLS")
        report_lines.append("-" * 80)
        merged_cell_issues = [issue for issue in self.issues if issue.issue_type == "TABLE_MERGED_CELLS"]
        if merged_cell_issues:
            report_lines.append(f"[X] Found {len(merged_cell_issues)} table(s) with merged cells:")
            for issue in merged_cell_issues:
                report_lines.append(f"  - {issue.description}")
        else:
            if len(self.target_doc.tables) > 0:
                report_lines.append("[OK] No merged cells detected")
            else:
                report_lines.append("[OK] No tables to check")
        report_lines.append("")

        # Table numeric alignment
        report_lines.append("ACCESSIBILITY: TABLE NUMERIC ALIGNMENT")
        report_lines.append("-" * 80)
        numeric_align_issues = [issue for issue in self.issues if issue.issue_type == "TABLE_INCONSISTENT_NUMERIC_ALIGNMENT"]
        if numeric_align_issues:
            report_lines.append(f"[X] Found {len(numeric_align_issues)} table column(s) with inconsistent numeric alignment:")
            for issue in numeric_align_issues:
                report_lines.append(f"  - {issue.description}")
        else:
            if len(self.target_doc.tables) > 0:
                report_lines.append("[OK] Numeric data alignment is consistent")
            else:
                report_lines.append("[OK] No tables to check")
        report_lines.append("")

        # Table reading order
        report_lines.append("ACCESSIBILITY: TABLE READING ORDER")
        report_lines.append("-" * 80)
        reading_order_issues = [issue for issue in self.issues if issue.issue_type == "TABLE_COMPLEX_READING_ORDER"]
        if reading_order_issues:
            report_lines.append(f"[X] Found {len(reading_order_issues)} table(s) with complex reading order:")
            for issue in reading_order_issues:
                report_lines.append(f"  - {issue.description}")
        else:
            if len(self.target_doc.tables) > 0:
                report_lines.append("[OK] Table reading order appears logical")
            else:
                report_lines.append("[OK] No tables to check")
        report_lines.append("")

        # Table embedded images
        report_lines.append("ACCESSIBILITY: TABLE EMBEDDED IMAGES")
        report_lines.append("-" * 80)
        embedded_image_issues = [issue for issue in self.issues if issue.issue_type == "TABLE_EMBEDDED_IMAGES"]
        if embedded_image_issues:
            report_lines.append(f"[X] Found {len(embedded_image_issues)} table(s) with embedded images:")
            for issue in embedded_image_issues:
                report_lines.append(f"  - {issue.description}")
        else:
            if len(self.target_doc.tables) > 0:
                report_lines.append("[OK] No embedded images in table cells")
            else:
                report_lines.append("[OK] No tables to check")
        report_lines.append("")

        # ============================================================
        # CATEGORY 3: COLOR & CONTRAST
        # ============================================================
        report_lines.append("=" * 80)
        report_lines.append("CATEGORY: COLOR & CONTRAST")
        report_lines.append("=" * 80)
        report_lines.append("")

        # Color contrast
        report_lines.append("ACCESSIBILITY: COLOR CONTRAST")
        report_lines.append("-" * 80)
        contrast_issues = [issue for issue in self.issues if issue.issue_type == "LOW_CONTRAST"]
        if contrast_issues:
            report_lines.append(f"[X] Found {len(contrast_issues)} instance(s) of insufficient color contrast:")
            report_lines.append("    (WCAG AA requires 4.5:1 for normal text, 3:1 for large text)")
            for issue in contrast_issues[:5]:
                report_lines.append(f"  - {issue.location}: {issue.description}")
            if len(contrast_issues) > 5:
                report_lines.append(f"  ... and {len(contrast_issues) - 5} more")
        else:
            report_lines.append("[OK] All text meets WCAG AA color contrast requirements")
        report_lines.append("")

        # Color as sole indicator
        report_lines.append("ACCESSIBILITY: COLOR AS SOLE INDICATOR")
        report_lines.append("-" * 80)
        color_only_issues = [issue for issue in self.issues if issue.issue_type == "COLOR_ONLY_MEANING"]
        color_coded_table_issues = [issue for issue in self.issues if issue.issue_type == "COLOR_CODED_TABLE"]
        if color_only_issues or color_coded_table_issues:
            if color_only_issues:
                report_lines.append(f"[X] Found {len(color_only_issues)} instance(s) where color may be the only indicator:")
                report_lines.append("    (Add bold, italic, or other formatting to supplement color)")
                for issue in color_only_issues[:5]:
                    report_lines.append(f"  - {issue.location}: {issue.description}")
                if len(color_only_issues) > 5:
                    report_lines.append(f"  ... and {len(color_only_issues) - 5} more")
            if color_coded_table_issues:
                report_lines.append(f"[X] Found {len(color_coded_table_issues)} color-coded table(s):")
                for issue in color_coded_table_issues:
                    report_lines.append(f"  - {issue.description}")
        else:
            report_lines.append("[OK] No issues with color as sole indicator detected")
        report_lines.append("")

        # Text over backgrounds
        report_lines.append("ACCESSIBILITY: TEXT OVER COLORED BACKGROUNDS")
        report_lines.append("-" * 80)
        text_bg_issues = [issue for issue in self.issues if issue.issue_type == "TEXT_OVER_BACKGROUND"]
        if text_bg_issues:
            report_lines.append(f"[X] Found {len(text_bg_issues)} instance(s) of text over colored backgrounds with low contrast:")
            report_lines.append("    (Ensure sufficient contrast between text and background colors)")
            for issue in text_bg_issues[:5]:
                report_lines.append(f"  - {issue.location}: {issue.description}")
            if len(text_bg_issues) > 5:
                report_lines.append(f"  ... and {len(text_bg_issues) - 5} more")
        else:
            report_lines.append("[OK] No issues with text over colored backgrounds detected")
        report_lines.append("")

        # ============================================================
        # CATEGORY 4: LINKS & NAVIGATION
        # ============================================================
        report_lines.append("=" * 80)
        report_lines.append("CATEGORY: LINKS & NAVIGATION")
        report_lines.append("=" * 80)
        report_lines.append("")

        # Non-descriptive links
        report_lines.append("ACCESSIBILITY: NON-DESCRIPTIVE LINKS")
        report_lines.append("-" * 80)
        non_desc_link_issues = [issue for issue in self.issues if issue.issue_type == "NON_DESCRIPTIVE_LINK"]
        if non_desc_link_issues:
            report_lines.append(f"[X] Found {len(non_desc_link_issues)} link(s) with non-descriptive text:")
            report_lines.append("    (Use descriptive text like 'Course Syllabus' instead of 'click here')")
            for issue in non_desc_link_issues[:5]:
                report_lines.append(f"  - {issue.location}: {issue.description}")
            if len(non_desc_link_issues) > 5:
                report_lines.append(f"  ... and {len(non_desc_link_issues) - 5} more")
        else:
            report_lines.append("[OK] No non-descriptive links detected")
        report_lines.append("")

        # Unstyled links check removed (pared back to critical issues)

        # Long URLs
        report_lines.append("ACCESSIBILITY: LONG URLs")
        report_lines.append("-" * 80)
        long_url_issues = [issue for issue in self.issues if issue.issue_type == "LONG_URL"]
        if long_url_issues:
            report_lines.append(f"[X] Found {len(long_url_issues)} excessively long URL(s):")
            report_lines.append("    (Consider using descriptive link text instead of displaying long URLs)")
            for issue in long_url_issues[:5]:
                report_lines.append(f"  - {issue.location}: {issue.description}")
            if len(long_url_issues) > 5:
                report_lines.append(f"  ... and {len(long_url_issues) - 5} more")
        else:
            report_lines.append("[OK] No excessively long URLs detected")
        report_lines.append("")

        # PDF link check removed - users typically can't control accessibility of institutional PDFs
        # Table of contents and bookmarks checks removed (pared back to critical issues)

        # ============================================================
        # CATEGORY 5: LISTS
        # ============================================================
        report_lines.append("=" * 80)
        report_lines.append("CATEGORY: LISTS")
        report_lines.append("=" * 80)
        report_lines.append("")

        # List formatting
        report_lines.append("ACCESSIBILITY: LIST FORMATTING")
        report_lines.append("-" * 80)
        list_check = self.check_list_usage()
        report_lines.append(f"Proper list items: {list_check['list_count']}")
        report_lines.append(f"Manual list items: {list_check['manual_list_count']}")
        if list_check['issues']:
            for issue in list_check['issues']:
                report_lines.append(f"[X] {issue}")
        else:
            report_lines.append("[OK] List formatting looks good")
        report_lines.append("")

        # Nested list hierarchy and layout lists checks removed (pared back to critical issues)

        # ============================================================
        # CATEGORY 6: TEXT FORMATTING
        # ============================================================
        report_lines.append("=" * 80)
        report_lines.append("CATEGORY: TEXT FORMATTING")
        report_lines.append("=" * 80)
        report_lines.append("")

        # Manual alignment (pseudo-tables)
        report_lines.append("ACCESSIBILITY: MANUAL ALIGNMENT (PSEUDO-TABLES)")
        report_lines.append("-" * 80)
        pseudo_table_issues = [issue for issue in self.issues if issue.issue_type == "PSEUDO_TABLE"]
        if pseudo_table_issues:
            report_lines.append(f"[X] Found {len(pseudo_table_issues)} paragraph(s) using tabs or spaces for alignment:")
            report_lines.append("    (This creates pseudo-tables that are difficult for screen readers to parse)")
            for issue in pseudo_table_issues[:5]:
                report_lines.append(f"  - {issue.location}: {issue.description}")
            if len(pseudo_table_issues) > 5:
                report_lines.append(f"  ... and {len(pseudo_table_issues) - 5} more")
        else:
            report_lines.append("[OK] No manual alignment detected")
        report_lines.append("")

        # Underlined text
        report_lines.append("ACCESSIBILITY: UNDERLINED TEXT")
        report_lines.append("-" * 80)
        underline_issues = [issue for issue in self.issues if issue.issue_type == "UNDERLINE_NON_LINK"]
        if underline_issues:
            report_lines.append(f"[X] Found {len(underline_issues)} underlined text(s) that are not hyperlinks:")
            report_lines.append("    (Underlined text is often confused with links by users)")
            for issue in underline_issues[:5]:
                report_lines.append(f"  - {issue.location}: {issue.description}")
            if len(underline_issues) > 5:
                report_lines.append(f"  ... and {len(underline_issues) - 5} more")
        else:
            report_lines.append("[OK] No non-link underlined text detected")
        report_lines.append("")

        # Line spacing
        report_lines.append("ACCESSIBILITY: LINE SPACING")
        report_lines.append("-" * 80)
        spacing_issues = [issue for issue in self.issues if issue.issue_type == "LOW_LINE_SPACING"]
        if spacing_issues:
            report_lines.append(f"[X] Found {len(spacing_issues)} paragraph(s) with line spacing below 1.15:")
            report_lines.append("    (WCAG recommends minimum 1.15 for readability)")
            for issue in spacing_issues[:5]:
                report_lines.append(f"  - {issue.location}: {issue.description}")
            if len(spacing_issues) > 5:
                report_lines.append(f"  ... and {len(spacing_issues) - 5} more")
        else:
            report_lines.append("[OK] Line spacing meets minimum requirements")
        report_lines.append("")

        # Text justification
        report_lines.append("ACCESSIBILITY: TEXT JUSTIFICATION")
        report_lines.append("-" * 80)
        justify_issues = [issue for issue in self.issues if issue.issue_type == "FULL_JUSTIFICATION"]
        if justify_issues:
            report_lines.append(f"[X] Found {len(justify_issues)} paragraph(s) with full justification:")
            report_lines.append("    (Creates uneven spacing, harder to read especially for dyslexic readers)")
            for issue in justify_issues[:5]:
                report_lines.append(f"  - {issue.location}: {issue.description}")
            if len(justify_issues) > 5:
                report_lines.append(f"  ... and {len(justify_issues) - 5} more")
        else:
            report_lines.append("[OK] No full justification detected")
        report_lines.append("")

        # ALL CAPS blocks
        report_lines.append("ACCESSIBILITY: ALL CAPS BLOCKS")
        report_lines.append("-" * 80)
        caps_issues = [issue for issue in self.issues if issue.issue_type == "ALL_CAPS_BLOCK"]
        if caps_issues:
            report_lines.append(f"[X] Found {len(caps_issues)} large block(s) of ALL CAPS text:")
            report_lines.append("    (ALL CAPS text is harder to read than mixed case)")
            for issue in caps_issues[:5]:
                report_lines.append(f"  - {issue.location}: {issue.description}")
            if len(caps_issues) > 5:
                report_lines.append(f"  ... and {len(caps_issues) - 5} more")
        else:
            report_lines.append("[OK] No large ALL CAPS blocks detected")
        report_lines.append("")

        # Excessive formatting
        report_lines.append("ACCESSIBILITY: EXCESSIVE/INCONSISTENT FORMATTING")
        report_lines.append("-" * 80)
        excessive_bold = [issue for issue in self.issues if issue.issue_type == "EXCESSIVE_BOLD"]
        excessive_italic = [issue for issue in self.issues if issue.issue_type == "EXCESSIVE_ITALIC"]
        excessive_underline = [issue for issue in self.issues if issue.issue_type == "EXCESSIVE_UNDERLINE"]
        inconsistent_formatting = [issue for issue in self.issues if issue.issue_type == "INCONSISTENT_FORMATTING"]
        all_formatting_issues = excessive_bold + excessive_italic + excessive_underline + inconsistent_formatting
        if all_formatting_issues:
            if excessive_bold:
                report_lines.append(f"[X] Found {len(excessive_bold)} paragraph(s) with excessive bold:")
                for issue in excessive_bold[:3]:
                    report_lines.append(f"  - {issue.location}: {issue.description}")
            if excessive_italic:
                report_lines.append(f"[X] Found {len(excessive_italic)} paragraph(s) with excessive italics:")
                for issue in excessive_italic[:3]:
                    report_lines.append(f"  - {issue.location}: {issue.description}")
            if excessive_underline:
                report_lines.append(f"[X] Found {len(excessive_underline)} paragraph(s) with excessive underline:")
                for issue in excessive_underline[:3]:
                    report_lines.append(f"  - {issue.location}: {issue.description}")
            if inconsistent_formatting:
                report_lines.append(f"[X] Found {len(inconsistent_formatting)} paragraph(s) with inconsistent formatting:")
                for issue in inconsistent_formatting[:3]:
                    report_lines.append(f"  - {issue.location}: {issue.description}")
            report_lines.append("    (Excessive or inconsistent formatting reduces effectiveness and readability)")
        else:
            report_lines.append("[OK] Formatting usage is appropriate")
        report_lines.append("")

        # CATEGORY 7: READABILITY removed (long sentences and date formats checks pared back)

        # ============================================================
        # CATEGORY 7: IMAGES
        # ============================================================
        report_lines.append("=" * 80)
        report_lines.append("CATEGORY: IMAGES")
        report_lines.append("=" * 80)
        report_lines.append("")

        # Image alt text
        report_lines.append("ACCESSIBILITY: IMAGE ALT TEXT")
        report_lines.append("-" * 80)
        image_alt_issues = [issue for issue in self.issues if issue.issue_type == "IMAGE_MISSING_ALT"]
        if image_alt_issues:
            report_lines.append(f"[X] Found {len(image_alt_issues)} image(s) missing alt text:")
            report_lines.append("    (All images need alt text or must be marked as decorative)")
            for issue in image_alt_issues[:5]:
                report_lines.append(f"  - {issue.location}: {issue.description}")
            if len(image_alt_issues) > 5:
                report_lines.append(f"  ... and {len(image_alt_issues) - 5} more")
        else:
            report_lines.append("[OK] All images have alt text or no images detected")
        report_lines.append("")

        # Decorative images
        report_lines.append("ACCESSIBILITY: DECORATIVE IMAGE MARKING")
        report_lines.append("-" * 80)
        decorative_issues = [issue for issue in self.issues if issue.issue_type == "DECORATIVE_IMAGE_QUESTIONABLE"]
        if decorative_issues:
            report_lines.append(f"[X] Found {len(decorative_issues)} image(s) marked as decorative but may be meaningful:")
            report_lines.append("    (Verify that decorative marking is appropriate for these images)")
            for issue in decorative_issues[:5]:
                report_lines.append(f"  - {issue.location}: {issue.description}")
            if len(decorative_issues) > 5:
                report_lines.append(f"  ... and {len(decorative_issues) - 5} more")
        else:
            report_lines.append("[OK] Decorative image marking appears appropriate")
        report_lines.append("")

        # Images containing text content
        report_lines.append("ACCESSIBILITY: IMAGES CONTAINING TEXT/SCHEDULES")
        report_lines.append("-" * 80)
        image_text_issues = [issue for issue in self.issues if issue.issue_type == "IMAGE_TEXT_CONTENT"]
        if image_text_issues:
            report_lines.append(f"[X] Found {len(image_text_issues)} image(s) that may contain text content:")
            report_lines.append("    (Text in images is inaccessible to screen readers - use actual text/tables instead)")
            for issue in image_text_issues[:5]:
                report_lines.append(f"  - {issue.location}: {issue.description}")
            if len(image_text_issues) > 5:
                report_lines.append(f"  ... and {len(image_text_issues) - 5} more")
        else:
            report_lines.append("[OK] No images appear to contain text content")
        report_lines.append("")

        # General images info
        report_lines.append("ACCESSIBILITY: IMAGES")
        report_lines.append("-" * 80)
        image_check = self.check_images()
        if image_check['issues']:
            for issue in image_check['issues']:
                report_lines.append(f"[i] {issue}")
        else:
            report_lines.append("[OK] No images detected")
        report_lines.append("")

        # ============================================================
        # CATEGORY 9: DOCUMENT PROPERTIES
        # ============================================================
        report_lines.append("=" * 80)
        report_lines.append("CATEGORY: DOCUMENT PROPERTIES")
        report_lines.append("=" * 80)
        report_lines.append("")

        # Document metadata
        report_lines.append("ACCESSIBILITY: DOCUMENT METADATA")
        report_lines.append("-" * 80)
        metadata_issues = [issue for issue in self.issues if issue.issue_type == "MISSING_TITLE"]
        if metadata_issues:
            for issue in metadata_issues:
                report_lines.append(f"[X] {issue.description}")
        else:
            report_lines.append("[OK] Document has title metadata")
        report_lines.append("")

        # Document language
        report_lines.append("ACCESSIBILITY: DOCUMENT LANGUAGE SETTING")
        report_lines.append("-" * 80)
        language_issues = [issue for issue in self.issues if issue.issue_type in ["MISSING_LANGUAGE", "INVALID_LANGUAGE"]]
        if language_issues:
            for issue in language_issues:
                report_lines.append(f"[X] {issue.description}")
        else:
            report_lines.append("[OK] Document has valid language setting")
        report_lines.append("")

        # Multilingual content
        report_lines.append("ACCESSIBILITY: MULTILINGUAL CONTENT")
        report_lines.append("-" * 80)
        multilingual_issues = [issue for issue in self.issues if issue.issue_type == "MULTIPLE_LANGUAGES"]
        untagged_language_issues = [issue for issue in self.issues if issue.issue_type == "UNTAGGED_LANGUAGE"]
        if multilingual_issues or untagged_language_issues:
            if multilingual_issues:
                for issue in multilingual_issues:
                    report_lines.append(f"[i] {issue.description}")
            if untagged_language_issues:
                report_lines.append(f"[X] Found potential multilingual content without proper language tagging:")
                for issue in untagged_language_issues:
                    report_lines.append(f"  - {issue.description}")
        else:
            report_lines.append("[OK] No multilingual content issues detected")
        report_lines.append("")

        # ============================================================
        # CATEGORY 10: CONTENT QUALITY & FORMATTING
        # ============================================================
        report_lines.append("=" * 80)
        report_lines.append("CATEGORY: CONTENT QUALITY & FORMATTING")
        report_lines.append("=" * 80)
        report_lines.append("")

        # Copied content with broken styles
        report_lines.append("ACCESSIBILITY: COPIED CONTENT WITH INCONSISTENT STYLES")
        report_lines.append("-" * 80)
        broken_style_issues = [issue for issue in self.issues if issue.issue_type == "BROKEN_STYLE_COPIED_CONTENT"]
        if broken_style_issues:
            report_lines.append(f"[X] Found {len(broken_style_issues)} paragraph(s) with inconsistent formatting (possible copied content):")
            report_lines.append("    (Mixed fonts/sizes within paragraphs suggest pasted content with broken styles)")
            for issue in broken_style_issues[:5]:
                report_lines.append(f"  - {issue.location}: {issue.description}")
            if len(broken_style_issues) > 5:
                report_lines.append(f"  ... and {len(broken_style_issues) - 5} more")
        else:
            report_lines.append("[OK] No inconsistent formatting detected")
        report_lines.append("")

        # Footnotes
        report_lines.append("ACCESSIBILITY: FOOTNOTE USAGE")
        report_lines.append("-" * 80)
        footnote_issues = [issue for issue in self.issues if issue.issue_type == "FOOTNOTE_USAGE"]
        if footnote_issues:
            report_lines.append(f"[X] Found {len(footnote_issues)} footnote marker(s):")
            report_lines.append("    (Footnotes can be difficult for screen readers; consider inline explanations or endnotes)")
            for issue in footnote_issues[:5]:
                report_lines.append(f"  - {issue.location}: {issue.description}")
            if len(footnote_issues) > 5:
                report_lines.append(f"  ... and {len(footnote_issues) - 5} more")
        else:
            report_lines.append("[OK] No footnotes detected")
        report_lines.append("")

        # Visual indicators
        report_lines.append("ACCESSIBILITY: VISUAL INDICATORS WITHOUT TEXT")
        report_lines.append("-" * 80)
        visual_indicator_issues = [issue for issue in self.issues if issue.issue_type == "VISUAL_INDICATOR_NO_TEXT"]
        if visual_indicator_issues:
            report_lines.append(f"[X] Found {len(visual_indicator_issues)} visual symbol(s) without clear text explanation:")
            report_lines.append("    (Symbols like ★, •, ⚠ may not be understood by screen readers)")
            for issue in visual_indicator_issues[:5]:
                report_lines.append(f"  - {issue.location}: {issue.description}")
            if len(visual_indicator_issues) > 5:
                report_lines.append(f"  ... and {len(visual_indicator_issues) - 5} more")
        else:
            report_lines.append("[OK] No unexplained visual indicators detected")
        report_lines.append("")

        # Math expressions
        report_lines.append("ACCESSIBILITY: MATHEMATICAL EXPRESSIONS")
        report_lines.append("-" * 80)
        math_issues = [issue for issue in self.issues if issue.issue_type == "MATH_NO_ACCESSIBLE_MARKUP"]
        if math_issues:
            report_lines.append(f"[X] Found {len(math_issues)} mathematical expression(s) without accessible markup:")
            report_lines.append("    (Use equation editors or MathML for complex math to ensure screen reader compatibility)")
            for issue in math_issues[:5]:
                report_lines.append(f"  - {issue.location}: {issue.description}")
            if len(math_issues) > 5:
                report_lines.append(f"  ... and {len(math_issues) - 5} more")
        else:
            report_lines.append("[OK] No inaccessible math expressions detected")
        report_lines.append("")

        # ============================================================
        # SUMMARY
        # ============================================================
        report_lines.append("=" * 80)
        report_lines.append("SUMMARY")
        report_lines.append("=" * 80)

        # Count total issues - use self.issues which contains all algorithmic checks
        # Plus add heading_check, table_check, and list_check issues which are returned separately
        total_issues = (
            len(missing) +
            len(heading_check['issues']) +
            len(table_check['issues']) +
            len(list_check['issues']) +
            len(self.issues)  # This contains all issues from run_all_checks()
        )

        if total_issues == 0:
            report_lines.append("[OK] Excellent! This syllabus meets all standards.")
        elif total_issues < 5:
            report_lines.append(f"[!] This syllabus has {total_issues} issues that should be addressed.")
        else:
            report_lines.append(f"[X] This syllabus has {total_issues} issues that need attention.")

        report_lines.append("")
        report_lines.append("RECOMMENDATIONS:")
        report_lines.append("1. Add all missing required sections from the template")
        report_lines.append("2. Use Heading 1, Heading 2, etc. styles (not bold text) for section headers")
        report_lines.append("3. Avoid using tables for layout; use headings and paragraphs instead")
        report_lines.append("4. Ensure all data tables have proper header rows")
        report_lines.append("5. Use descriptive hyperlink text instead of displaying URLs")
        report_lines.append("6. Ensure all images have descriptive alt text")
        report_lines.append("7. Use built-in list formatting instead of manual bullets")
        report_lines.append("8. Don't use tabs or multiple spaces for alignment; use tables or proper formatting")
        report_lines.append("")

        return "\n".join(report_lines)


def main():
    """Main entry point for command-line usage"""
    import os

    parser = argparse.ArgumentParser(
        description='Check a syllabus against UWM template standards',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s my_syllabus.docx
  %(prog)s english215.docx
        """
    )
    parser.add_argument('syllabus', help='Path to the syllabus to check (DOCX)')
    parser.add_argument('-t', '--template',
                       default='Uniform-Syllabus-Template-1.docx',
                       help='Path to template (default: Uniform-Syllabus-Template-1.docx)')

    args = parser.parse_args()

    # Auto-generate output file names based on input syllabus
    base_name = os.path.splitext(args.syllabus)[0]
    output_report = f"{base_name}_report.txt"
    output_marked = f"{base_name}_marked.docx"

    print(f"Analyzing: {args.syllabus}")
    print(f"Template: {args.template}")
    print()

    checker = SyllabusChecker(args.template, args.syllabus)

    # Show document stats
    print(f"Found {len(checker.all_paragraphs)} total paragraphs")
    top_level = sum(1 for p in checker.all_paragraphs if p.table_idx is None)
    in_tables = sum(1 for p in checker.all_paragraphs if p.table_idx is not None)
    print(f"  - {top_level} top-level paragraphs")
    print(f"  - {in_tables} paragraphs in tables")
    print()

    print("Running accessibility checks...")
    checker.run_all_checks()

    # Generate report
    report = checker.generate_report()

    # Always save text report
    with open(output_report, 'w') as f:
        f.write(report)
    print(f"✓ Text report saved to: {output_report}")

    # Always create marked document
    print(f"Creating marked-up document...")
    checker.create_marked_document(output_marked)
    print(f"✓ Marked document saved to: {output_marked}")

    print()
    print("=" * 60)
    print("DONE! Review the following files:")
    print(f"  1. {output_report} - Detailed text report")
    print(f"  2. {output_marked} - Word doc with issues highlighted")
    print("=" * 60)


if __name__ == '__main__':
    main()
