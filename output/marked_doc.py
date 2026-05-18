#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
OutputMixin — document marking and report generation for SyllabusChecker.
"""

import re
import html
from typing import List, Dict
from docx import Document
from docx.shared import RGBColor, Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from constants import (
    COLOR_ACCESSIBILITY_HEX, COLOR_MISSING_HEX, COLOR_TONE_HEX, COLOR_QUALITY_HEX,
    COLOR_ACCESSIBILITY_TEXT, COLOR_MISSING_TEXT, COLOR_TONE_TEXT, COLOR_QUALITY_TEXT,
)
from _checker_types import _add_word_comment


class OutputMixin:
    """Mixin providing document marking and report generation methods."""

    def create_marked_document(self, output_path: str, missing_sections: List[str] = None,
                               growth_mindset_analysis: Dict[str, str] = None,
                               quality_analysis: Dict[str, str] = None):
        """Create a copy of the document with issues marked via comments and highlighting,
        plus sections for missing required content and growth mindset recommendations"""
        marked_doc = Document(self.target_path)

        # UWM brand colors are defined in constants.py and imported at module level.

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

        def insert_callout(para, label, description, fill_hex='FFF0E8'):
            """Insert a small shaded note paragraph immediately after *para* in the body."""
            from lxml import etree
            W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
            XML = 'http://www.w3.org/XML/1998/namespace'
            new_p = etree.SubElement(para._p.getparent(), f'{{{W}}}p')
            pPr = etree.SubElement(new_p, f'{{{W}}}pPr')
            shd = etree.SubElement(pPr, f'{{{W}}}shd')
            shd.set(f'{{{W}}}val', 'clear')
            shd.set(f'{{{W}}}color', 'auto')
            shd.set(f'{{{W}}}fill', fill_hex)
            ind = etree.SubElement(pPr, f'{{{W}}}ind')
            ind.set(f'{{{W}}}left', '360')
            rPr_p = etree.SubElement(pPr, f'{{{W}}}rPr')
            sz_p = etree.SubElement(rPr_p, f'{{{W}}}sz')
            sz_p.set(f'{{{W}}}val', '18')
            r = etree.SubElement(new_p, f'{{{W}}}r')
            rPr = etree.SubElement(r, f'{{{W}}}rPr')
            etree.SubElement(rPr, f'{{{W}}}b')
            sz = etree.SubElement(rPr, f'{{{W}}}sz')
            sz.set(f'{{{W}}}val', '18')
            t = etree.SubElement(r, f'{{{W}}}t')
            t.text = f'▶ {label}: '
            t.set(f'{{{XML}}}space', 'preserve')
            r2 = etree.SubElement(new_p, f'{{{W}}}r')
            rPr2 = etree.SubElement(r2, f'{{{W}}}rPr')
            sz2 = etree.SubElement(rPr2, f'{{{W}}}sz')
            sz2.set(f'{{{W}}}val', '18')
            t2 = etree.SubElement(r2, f'{{{W}}}t')
            t2.text = description
            t2.set(f'{{{XML}}}space', 'preserve')
            # Move new_p to immediately follow para._p
            para._p.addnext(new_p)

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

        # Mark top-level paragraphs — inline callout + Word comment
        for (table_idx, row_idx, cell_idx), issues_list in para_issues_by_location.items():
            if table_idx is None:
                for issue in issues_list:
                    original_text = issue.para_info.paragraph.text.strip()

                    for para in marked_doc.paragraphs:
                        para_text = para.text.strip()

                        if para_text and (para_text in original_text or original_text in para_text):
                            label = self.ISSUE_DISPLAY_LABELS.get(issue.issue_type, issue.issue_type.replace('_', ' ').title())
                            _add_word_comment(marked_doc, para, next_comment_id(), f"{label}\n{issue.description}")
                            insert_callout(para, label, issue.description)
                            break

        # Mark paragraphs in table cells — inline callout + Word comment
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
                                _add_word_comment(marked_doc, para, next_comment_id(), f"{label}\n{issue.description}")
                                insert_callout(para, label, issue.description)
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
                report_lines.append("[X] Found potential multilingual content without proper language tagging:")
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

