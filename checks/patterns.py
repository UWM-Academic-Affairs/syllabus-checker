#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PatternsMixin — section/pattern detection methods for SyllabusChecker.
"""

import re
import os
import html
from typing import List, Dict
import sys
from docx import Document
from docx.oxml.ns import qn
from _checker_types import AccessibilityIssue, ParagraphInfo


def _get_openai():
    """Return OpenAI class, resolving through syllabus_checker so mocks work."""
    sc = sys.modules.get('syllabus_checker')
    if sc is not None and hasattr(sc, 'OpenAI'):
        return sc.OpenAI
    from openai import OpenAI
    return OpenAI


class PatternsMixin:
    """Mixin providing section/pattern detection capabilities."""

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

        Walks document body children in order (paragraphs AND tables) so that
        sample tables embedded between instruction paragraphs are included in
        the section description.

        Returns a dict mapping heading text to {level, description, keywords}
        """
        from docx.text.paragraph import Paragraph as DocxParagraph

        w_ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

        standard = {}
        current_heading = None
        current_content = []

        def save_heading():
            if current_heading:
                description = self._wrap_list_items(current_content)
                heading_lower = current_heading['text'].lower()
                keywords = [w for w in heading_lower.split() if len(w) > 3]
                standard[current_heading['text']] = {
                    'level': current_heading['level'],
                    'description': description,
                    'keywords': keywords
                }

        def table_to_html(tbl_element) -> str:
            """Render a template table as a simple pipe-delimited HTML string."""
            rows = tbl_element.findall(f'{{{w_ns}}}tr')
            if not rows:
                return ''
            html_rows = []
            for row in rows:
                cells = row.findall(f'.//{{{w_ns}}}tc')
                cell_texts = []
                for cell in cells:
                    # Join runs directly (no added spaces) — Word text runs are
                    # already properly spaced within the original text.
                    text = ''.join(
                        t.text for t in cell.findall(f'.//{{{w_ns}}}t') if t.text
                    ).strip()
                    cell_texts.append(text)
                html_rows.append(' | '.join(cell_texts))
            return '<br>' + '<br>'.join(html_rows)

        for child in self.template_doc.element.body:
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag

            if tag == 'p':
                para = DocxParagraph(child, self.template_doc)
                style_name = para.style.name if para.style else ''
                text = para.text.strip()

                if style_name.startswith('Heading'):
                    save_heading()
                    level = int(style_name.split()[-1]) if style_name[-1].isdigit() else 1
                    current_heading = {'text': text, 'level': level}
                    current_content = []
                elif current_heading and text:
                    is_list = style_name == 'List Paragraph'
                    para_html = self._para_to_html(para)
                    if is_list:
                        para_html = f'<li>{para_html}</li>'
                    current_content.append(para_html)

            elif tag == 'tbl' and current_heading:
                # Include sample tables in the section description
                tbl_html = table_to_html(child)
                if tbl_html:
                    current_content.append(tbl_html)

        save_heading()
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

        Priority rules:
        - Table-cell Word headings never override a top-level Word heading (prevents row
          labels like "Revision" from overriding the real section heading).
        - A pseudo-heading (plain-text informal heading) that is *closer* to the content
          than the nearest top-level Word heading wins — this handles documents where all
          content lives inside a table whose row labels are not styled as headings (e.g.
          a two-column table with section labels in column 0).
        """
        nearest_toplevel = None
        nearest_word_heading = None  # table-cell Word heading (is_word_heading, not is_top_level)
        nearest_pseudo = None        # plain-text pseudo-heading (not is_word_heading)
        for heading in headings:
            if heading['position'] < content_position:
                pos = heading['position']
                if heading.get('is_word_heading'):
                    if heading.get('is_top_level'):
                        if nearest_toplevel is None or pos > nearest_toplevel['position']:
                            nearest_toplevel = heading
                    else:
                        if nearest_word_heading is None or pos > nearest_word_heading['position']:
                            nearest_word_heading = heading
                else:
                    if nearest_pseudo is None or pos > nearest_pseudo['position']:
                        nearest_pseudo = heading

        # If a pseudo-heading is closer than the top-level heading, prefer it —
        # the top-level heading is likely a document title covering the whole doc.
        if nearest_toplevel and nearest_pseudo:
            if nearest_pseudo['position'] > nearest_toplevel['position']:
                return nearest_pseudo
            return nearest_toplevel

        return nearest_toplevel or nearest_word_heading or nearest_pseudo

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
                    heading_words = set(w for w in heading_lower.split() if len(w) > 3)
                    for std in self.REQUIRED_SECTIONS.keys():
                        std_lower = std.lower()
                        if std_lower == heading_lower or std_lower in heading_lower or heading_lower in std_lower:
                            heading_ok = True
                            break
                        # Word-overlap: heading shares 2+ meaningful words with a required section
                        # Use prefix matching to handle plurals (assignment/assignments)
                        std_words = set(w for w in std_lower.split() if len(w) > 3)
                        def _words_overlap(ws1, ws2):
                            count = 0
                            for w1 in ws1:
                                for w2 in ws2:
                                    if w1 == w2 or w1.startswith(w2) or w2.startswith(w1):
                                        count += 1
                                        break
                            return count
                        if _words_overlap(heading_words, std_words) >= 2:
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
            client = _get_openai()(api_key=api_key)

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

