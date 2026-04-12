"""
Unit tests for SyllabusChecker accessibility check methods:
  check_table_headers()
  check_images_alt_text()
  check_color_contrast()
  check_non_descriptive_links()
  check_all_caps_blocks()
  check_full_justification()
  check_font_sizes()

Each test builds a minimal DOCX programmatically, wraps it in a
SyllabusChecker, and calls only the relevant check method.
"""

import os
import pytest
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from lxml import etree

from syllabus_checker import SyllabusChecker


# ---------------------------------------------------------------------------
# Shared helper
# ---------------------------------------------------------------------------

def _checker(tmp_path_factory, name, builder_fn):
    """Build a DOCX via builder_fn, return a SyllabusChecker instance."""
    d = tmp_path_factory.mktemp(name)
    path = str(d / f'{name}.docx')
    builder_fn(path)
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    template = os.path.join(root, 'assets', 'Uniform-Syllabus-Template-1.docx')
    return SyllabusChecker(template, path)


# ---------------------------------------------------------------------------
# check_table_headers
# ---------------------------------------------------------------------------

class TestCheckTableHeaders:

    @pytest.fixture(scope='class')
    def checker_no_header(self, tmp_path_factory):
        """2-row table, first row has plain (non-bold) text — no header detected."""
        def build(path):
            doc = Document()
            table = doc.add_table(rows=2, cols=2)
            table.rows[0].cells[0].text = 'Col A'
            table.rows[0].cells[1].text = 'Col B'
            table.rows[1].cells[0].text = 'Data 1'
            table.rows[1].cells[1].text = 'Data 2'
            doc.save(path)
        return _checker(tmp_path_factory, 'no_header', build)

    @pytest.fixture(scope='class')
    def checker_bold_header(self, tmp_path_factory):
        """2-row table, first row cells have bold text — heuristic header detected."""
        def build(path):
            doc = Document()
            table = doc.add_table(rows=2, cols=2)
            for cell in table.rows[0].cells:
                run = cell.paragraphs[0].add_run(cell.text or 'Header')
                run.bold = True
            table.rows[1].cells[0].text = 'Data 1'
            table.rows[1].cells[1].text = 'Data 2'
            doc.save(path)
        return _checker(tmp_path_factory, 'bold_header', build)

    def test_missing_header_flagged(self, checker_no_header):
        issues = checker_no_header.check_table_headers()
        types = [i.issue_type for i in issues]
        assert 'TABLE_NO_HEADER' in types

    def test_bold_header_not_flagged(self, checker_bold_header):
        issues = checker_bold_header.check_table_headers()
        types = [i.issue_type for i in issues]
        assert 'TABLE_NO_HEADER' not in types


# ---------------------------------------------------------------------------
# check_images_alt_text
# ---------------------------------------------------------------------------

class TestCheckImagesAltText:

    @pytest.fixture(scope='class')
    def checker_no_alt(self, tmp_path_factory):
        """Inject a w:drawing element with a docPr that has no title or descr."""
        def build(path):
            doc = Document()
            doc.add_paragraph('A paragraph with an embedded image (no alt text).')
            # Inject minimal w:drawing XML with empty docPr into the paragraph
            para = doc.add_paragraph()
            drawing_xml = (
                '<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                '<w:drawing>'
                '<wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">'
                '<wp:docPr id="1" name="Image1"/>'
                '</wp:inline>'
                '</w:drawing>'
                '</w:r>'
            )
            run_elem = etree.fromstring(drawing_xml)
            para._element.append(run_elem)
            doc.save(path)
        return _checker(tmp_path_factory, 'no_alt', build)

    @pytest.fixture(scope='class')
    def checker_with_alt(self, tmp_path_factory):
        """Inject a w:drawing element with a docPr that has a non-empty descr."""
        def build(path):
            doc = Document()
            doc.add_paragraph('A paragraph with a described image.')
            para = doc.add_paragraph()
            drawing_xml = (
                '<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                '<w:drawing>'
                '<wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">'
                '<wp:docPr id="2" name="Image2" descr="A photograph of the campus"/>'
                '</wp:inline>'
                '</w:drawing>'
                '</w:r>'
            )
            run_elem = etree.fromstring(drawing_xml)
            para._element.append(run_elem)
            doc.save(path)
        return _checker(tmp_path_factory, 'with_alt', build)

    def test_image_without_alt_flagged(self, checker_no_alt):
        issues = checker_no_alt.check_images_alt_text()
        types = [i.issue_type for i in issues]
        assert 'IMAGE_MISSING_ALT' in types

    def test_image_with_alt_not_flagged(self, checker_with_alt):
        issues = checker_with_alt.check_images_alt_text()
        types = [i.issue_type for i in issues]
        assert 'IMAGE_MISSING_ALT' not in types


# ---------------------------------------------------------------------------
# check_color_contrast
# ---------------------------------------------------------------------------

class TestCheckColorContrast:

    @pytest.fixture(scope='class')
    def checker_low_contrast(self, tmp_path_factory):
        """Light gray (#AAAAAA) text on white — well below WCAG AA 4.5:1."""
        def build(path):
            doc = Document()
            para = doc.add_paragraph()
            run = para.add_run('This text has very low contrast against the white background.')
            run.font.color.rgb = RGBColor(0xAA, 0xAA, 0xAA)
            doc.save(path)
        return _checker(tmp_path_factory, 'low_contrast', build)

    def test_low_contrast_flagged(self, checker_low_contrast):
        issues = checker_low_contrast.check_color_contrast()
        types = [i.issue_type for i in issues]
        assert 'LOW_CONTRAST' in types


# ---------------------------------------------------------------------------
# check_non_descriptive_links
# ---------------------------------------------------------------------------

class TestCheckNonDescriptiveLinks:

    @pytest.fixture(scope='class')
    def checker_bad_link(self, tmp_path_factory):
        """Underlined run with text 'click here' — non-descriptive."""
        def build(path):
            doc = Document()
            para = doc.add_paragraph('For more information ')
            run = para.add_run('click here')
            run.underline = True
            para.add_run(' to continue.')
            doc.save(path)
        return _checker(tmp_path_factory, 'bad_link', build)

    @pytest.fixture(scope='class')
    def checker_good_link(self, tmp_path_factory):
        """Underlined run with descriptive text — should not be flagged."""
        def build(path):
            doc = Document()
            para = doc.add_paragraph('See the ')
            run = para.add_run('UWM Academic Integrity Policy')
            run.underline = True
            para.add_run(' for details.')
            doc.save(path)
        return _checker(tmp_path_factory, 'good_link', build)

    def test_click_here_flagged(self, checker_bad_link):
        issues = checker_bad_link.check_non_descriptive_links()
        types = [i.issue_type for i in issues]
        assert 'NON_DESCRIPTIVE_LINK' in types

    def test_descriptive_link_not_flagged(self, checker_good_link):
        issues = checker_good_link.check_non_descriptive_links()
        types = [i.issue_type for i in issues]
        assert 'NON_DESCRIPTIVE_LINK' not in types


# ---------------------------------------------------------------------------
# check_all_caps_blocks
# ---------------------------------------------------------------------------

class TestCheckAllCapsBlocks:

    @pytest.fixture(scope='class')
    def checker_all_caps(self, tmp_path_factory):
        """Paragraph ≥50 chars that is >80% uppercase."""
        def build(path):
            doc = Document()
            # 60 uppercase characters — well over the threshold
            doc.add_paragraph('THIS IS A VERY LONG ALL CAPS SENTENCE THAT SHOULD BE FLAGGED HERE.')
            doc.save(path)
        return _checker(tmp_path_factory, 'all_caps', build)

    @pytest.fixture(scope='class')
    def checker_mixed_case(self, tmp_path_factory):
        """Normal mixed-case paragraph — should not be flagged."""
        def build(path):
            doc = Document()
            doc.add_paragraph('This is a normal sentence with mixed case and proper punctuation.')
            doc.save(path)
        return _checker(tmp_path_factory, 'mixed_case', build)

    def test_all_caps_paragraph_flagged(self, checker_all_caps):
        issues = checker_all_caps.check_all_caps_blocks()
        types = [i.issue_type for i in issues]
        assert 'ALL_CAPS_BLOCK' in types

    def test_mixed_case_not_flagged(self, checker_mixed_case):
        issues = checker_mixed_case.check_all_caps_blocks()
        types = [i.issue_type for i in issues]
        assert 'ALL_CAPS_BLOCK' not in types


# ---------------------------------------------------------------------------
# check_full_justification
# ---------------------------------------------------------------------------

class TestCheckFullJustification:

    @pytest.fixture(scope='class')
    def checker_justified(self, tmp_path_factory):
        """Paragraph with WD_ALIGN_PARAGRAPH.JUSTIFY (value 3)."""
        def build(path):
            doc = Document()
            para = doc.add_paragraph(
                'This paragraph is fully justified, which can make it harder '
                'to read for people with dyslexia or other reading difficulties.'
            )
            para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            doc.save(path)
        return _checker(tmp_path_factory, 'justified', build)

    @pytest.fixture(scope='class')
    def checker_left_aligned(self, tmp_path_factory):
        """Paragraph with left alignment — should not be flagged."""
        def build(path):
            doc = Document()
            para = doc.add_paragraph('This paragraph uses left alignment, which is easier to read.')
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            doc.save(path)
        return _checker(tmp_path_factory, 'left_aligned', build)

    def test_justified_paragraph_flagged(self, checker_justified):
        issues = checker_justified.check_full_justification()
        types = [i.issue_type for i in issues]
        assert 'FULL_JUSTIFICATION' in types

    def test_left_aligned_not_flagged(self, checker_left_aligned):
        issues = checker_left_aligned.check_full_justification()
        types = [i.issue_type for i in issues]
        assert 'FULL_JUSTIFICATION' not in types


# ---------------------------------------------------------------------------
# check_font_sizes
# ---------------------------------------------------------------------------

class TestCheckFontSizes:

    @pytest.fixture(scope='class')
    def checker_small_font(self, tmp_path_factory):
        """Run with 9pt font — below the 11pt threshold."""
        def build(path):
            doc = Document()
            para = doc.add_paragraph()
            run = para.add_run('This text is uncomfortably small.')
            run.font.size = Pt(9)
            doc.save(path)
        return _checker(tmp_path_factory, 'small_font', build)

    @pytest.fixture(scope='class')
    def checker_normal_font(self, tmp_path_factory):
        """Run with 12pt font — at or above the threshold."""
        def build(path):
            doc = Document()
            para = doc.add_paragraph()
            run = para.add_run('This text is a comfortable size.')
            run.font.size = Pt(12)
            doc.save(path)
        return _checker(tmp_path_factory, 'normal_font', build)

    def test_small_font_flagged(self, checker_small_font):
        issues = checker_small_font.check_font_sizes()
        types = [i.issue_type for i in issues]
        assert 'SMALL_FONT' in types

    def test_normal_font_not_flagged(self, checker_normal_font):
        issues = checker_normal_font.check_font_sizes()
        types = [i.issue_type for i in issues]
        assert 'SMALL_FONT' not in types


# ---------------------------------------------------------------------------
# Clean document — all individual checks return empty lists
# ---------------------------------------------------------------------------

class TestCleanDocumentNoIssues:
    """
    A purpose-built clean DOCX with no accessibility problems.
    Each check method should return an empty list.
    """

    @pytest.fixture(scope='class')
    def clean_checker(self, tmp_path_factory):
        """DOCX carefully crafted to avoid triggering any accessibility check."""
        def build(path):
            doc = Document()
            # Left-aligned normal paragraphs, mixed case, 12pt, black text
            for i in range(3):
                para = doc.add_paragraph()
                para.alignment = WD_ALIGN_PARAGRAPH.LEFT
                run = para.add_run(f'This is a normal paragraph number {i + 1} with regular text.')
                run.font.size = Pt(12)
            doc.save(path)
        return _checker(tmp_path_factory, 'clean_doc', build)

    def test_no_table_header_issues(self, clean_checker):
        assert clean_checker.check_table_headers() == []

    def test_no_image_alt_issues(self, clean_checker):
        assert clean_checker.check_images_alt_text() == []

    def test_no_non_descriptive_link_issues(self, clean_checker):
        assert clean_checker.check_non_descriptive_links() == []

    def test_no_all_caps_issues(self, clean_checker):
        assert clean_checker.check_all_caps_blocks() == []

    def test_no_full_justification_issues(self, clean_checker):
        assert clean_checker.check_full_justification() == []
