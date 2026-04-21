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
from constants import (
    COLOR_ACCESSIBILITY_HEX, COLOR_MISSING_HEX, COLOR_TONE_HEX, COLOR_QUALITY_HEX,
    COLOR_ACCESSIBILITY_TEXT, COLOR_MISSING_TEXT, COLOR_TONE_TEXT, COLOR_QUALITY_TEXT,
)

# Import shared types and utilities
from _checker_types import (
    calculate_relative_luminance,
    calculate_contrast_ratio,
    _ensure_comments_part,
    _add_word_comment,
    ParagraphInfo,
    AccessibilityIssue,
)

# Import mixin classes
from checks.patterns import PatternsMixin
from checks.accessibility import AccessibilityMixin
from llm.analysis import LLMMixin
from output.marked_doc import OutputMixin


class SyllabusChecker(PatternsMixin, AccessibilityMixin, LLMMixin, OutputMixin):
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


def main():
    """Main entry point for command-line usage"""

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
