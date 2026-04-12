#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AccessibilityMixin — all check_* accessibility methods for SyllabusChecker.
"""

import re
from typing import List, Dict
from docx.oxml.ns import qn
from docx.shared import RGBColor, Pt
from _checker_types import AccessibilityIssue, calculate_relative_luminance, calculate_contrast_ratio


class AccessibilityMixin:
    """Mixin providing accessibility check methods."""

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

