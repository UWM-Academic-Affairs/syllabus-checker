"""
Tests for Anthropic-backed SyllabusChecker methods (all API calls mocked):
  extract_course_metadata()
  analyze_sections_with_llm()
  analyze_growth_mindset_and_belonging()
  analyze_syllabus_quality()
"""

import os
import pytest
from unittest.mock import MagicMock

from syllabus_checker import SyllabusChecker


# ---------------------------------------------------------------------------
# Shared fixture
# ---------------------------------------------------------------------------

@pytest.fixture(scope='module')
def checker(fixture_dir):
    """SyllabusChecker backed by minimal_valid.docx."""
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    template = os.path.join(root, 'assets', 'Uniform-Syllabus-Template-1.docx')
    return SyllabusChecker(template, str(fixture_dir / 'minimal_valid.docx'))


def _make_response(content: str) -> MagicMock:
    """Build a minimal object that mimics an Anthropic Messages API response."""
    block = MagicMock()
    block.type = 'text'
    block.text = content
    resp = MagicMock()
    resp.content = [block]
    resp.stop_reason = 'end_turn'
    return resp


# ---------------------------------------------------------------------------
# extract_course_metadata
# ---------------------------------------------------------------------------

class TestExtractCourseMetadata:

    VALID_JSON = (
        '{"instructor_name": "Dr. Test Instructor", '
        '"department": "Computer Science", '
        '"course_subject": "CS", '
        '"course_number": "101", '
        '"course_title": "Introduction to Testing", '
        '"semester": "Fall 2025"}'
    )

    def test_parses_valid_json_into_correct_fields(self, checker, monkeypatch, mocker):
        monkeypatch.setenv('ANTHROPIC_API_KEY', 'test-key')
        mock_client = mocker.patch('llm.client.get_client').return_value
        mock_client.messages.create.return_value = (
            _make_response(self.VALID_JSON)
        )

        result = checker.extract_course_metadata()

        assert result['status'] == 'success'
        assert result['instructor_name'] == 'Dr. Test Instructor'
        assert result['department'] == 'Computer Science'
        assert result['course_subject'] == 'CS'
        assert result['course_number'] == '101'
        assert result['course_title'] == 'Introduction to Testing'
        assert result['semester'] == 'Fall 2025'

    def test_strips_markdown_code_fence_before_parsing(self, checker, monkeypatch, mocker):
        """Response wrapped in ```json ... ``` must still be parsed correctly."""
        monkeypatch.setenv('ANTHROPIC_API_KEY', 'test-key')
        mock_client = mocker.patch('llm.client.get_client').return_value
        fenced = f'```json\n{self.VALID_JSON}\n```'
        mock_client.messages.create.return_value = (
            _make_response(fenced)
        )

        result = checker.extract_course_metadata()

        assert result['status'] == 'success'
        assert result['course_subject'] == 'CS'

    def test_handles_malformed_json_without_crashing(self, checker, monkeypatch, mocker):
        """Garbage JSON must return status='error', not raise an exception."""
        monkeypatch.setenv('ANTHROPIC_API_KEY', 'test-key')
        mock_client = mocker.patch('llm.client.get_client').return_value
        mock_client.messages.create.return_value = (
            _make_response('This is not JSON at all!!!')
        )

        result = checker.extract_course_metadata()

        assert result['status'] == 'error'
        assert 'error' in result

    def test_returns_error_when_api_key_absent(self, checker, monkeypatch):
        monkeypatch.delenv('ANTHROPIC_API_KEY', raising=False)

        result = checker.extract_course_metadata()

        assert result['status'] == 'error'

    def test_returns_error_on_api_exception(self, checker, monkeypatch, mocker):
        monkeypatch.setenv('ANTHROPIC_API_KEY', 'test-key')
        mock_client = mocker.patch('llm.client.get_client').return_value
        mock_client.messages.create.side_effect = (
            RuntimeError('network error')
        )

        result = checker.extract_course_metadata()

        assert result['status'] == 'error'


# ---------------------------------------------------------------------------
# analyze_sections_with_llm
# ---------------------------------------------------------------------------

class TestAnalyzeSectionsWithLlm:

    def test_uses_pattern_detection_when_api_key_absent(self, checker, monkeypatch):
        """Without an API key the method falls back to pattern + keyword matching."""
        monkeypatch.delenv('ANTHROPIC_API_KEY', raising=False)

        result = checker.analyze_sections_with_llm()

        # Must return the expected keys without calling the API
        assert 'present' in result
        assert 'missing' in result
        # The error field signals that pattern-matching fallback was used
        assert 'error' in result
        assert 'pattern' in result['error'].lower() or 'anthropic' in result['error'].lower()

    def test_present_list_is_non_empty_for_complete_syllabus(self, checker, monkeypatch):
        """minimal_valid.docx has many detectable sections; present list must be non-empty."""
        monkeypatch.delenv('ANTHROPIC_API_KEY', raising=False)

        result = checker.analyze_sections_with_llm()

        assert isinstance(result['present'], list)
        assert len(result['present']) > 0

    def test_missing_list_is_a_list(self, checker, monkeypatch):
        monkeypatch.delenv('ANTHROPIC_API_KEY', raising=False)

        result = checker.analyze_sections_with_llm()

        assert isinstance(result['missing'], list)


# ---------------------------------------------------------------------------
# analyze_growth_mindset_and_belonging
# ---------------------------------------------------------------------------

class TestAnalyzeGrowthMindset:

    def test_returns_error_status_when_api_key_absent(self, checker, monkeypatch):
        monkeypatch.delenv('ANTHROPIC_API_KEY', raising=False)

        result = checker.analyze_growth_mindset_and_belonging()

        assert result['status'] == 'error'
        assert 'error' in result

    def test_returns_error_status_when_api_call_raises(self, checker, monkeypatch, mocker):
        """Any exception from the Anthropic call must produce status='error'."""
        monkeypatch.setenv('ANTHROPIC_API_KEY', 'test-key')
        mock_client = mocker.patch('llm.client.get_client').return_value
        mock_client.messages.create.side_effect = (
            Exception('timeout')
        )

        result = checker.analyze_growth_mindset_and_belonging()

        assert result['status'] == 'error'
        assert 'error' in result

    def test_returns_success_with_analysis_text(self, checker, monkeypatch, mocker):
        monkeypatch.setenv('ANTHROPIC_API_KEY', 'test-key')
        mock_client = mocker.patch('llm.client.get_client').return_value
        mock_client.messages.create.return_value = (
            _make_response('Great syllabus with strong growth mindset language.')
        )

        result = checker.analyze_growth_mindset_and_belonging()

        assert result['status'] == 'success'
        assert 'analysis' in result
        assert 'growth mindset' in result['analysis'].lower()


# ---------------------------------------------------------------------------
# analyze_syllabus_quality
# ---------------------------------------------------------------------------

class TestAnalyzeSyllabusQuality:

    def test_returns_error_status_when_api_key_absent(self, checker, monkeypatch):
        monkeypatch.delenv('ANTHROPIC_API_KEY', raising=False)

        result = checker.analyze_syllabus_quality()

        assert result['status'] == 'error'
        assert 'error' in result

    def test_returns_error_gracefully_on_timeout(self, checker, monkeypatch, mocker):
        """A timeout (or any exception) must produce status='error', not a 500."""
        monkeypatch.setenv('ANTHROPIC_API_KEY', 'test-key')
        mock_client = mocker.patch('llm.client.get_client').return_value
        mock_client.messages.create.side_effect = (
            TimeoutError('request timed out')
        )

        result = checker.analyze_syllabus_quality()

        assert result['status'] == 'error'
        assert 'error' in result

    def test_returns_success_with_analysis_text(self, checker, monkeypatch, mocker):
        monkeypatch.setenv('ANTHROPIC_API_KEY', 'test-key')
        mock_client = mocker.patch('llm.client.get_client').return_value
        mock_client.messages.create.return_value = (
            _make_response('No significant quality issues found.')
        )

        result = checker.analyze_syllabus_quality()

        assert result['status'] == 'success'
        assert 'analysis' in result


# ---------------------------------------------------------------------------
# analyze_image_alt_text (image path)
#
# NOTE: this method sends the *alt-text strings* found in the document to the
# model; it does not upload the image bytes. These tests cover that real path:
# drawings are found in the docx, their alt text lands in the prompt, and the
# result keeps its shape (including image_count).
# ---------------------------------------------------------------------------

def _tiny_png(path):
    """Write a valid 2x2 PNG without needing Pillow."""
    import struct
    import zlib

    def chunk(tag, data):
        c = struct.pack('>I', len(data)) + tag + data
        return c + struct.pack('>I', zlib.crc32(tag + data) & 0xFFFFFFFF)

    raw = b''.join(b'\x00' + b'\xff\x00\x00' * 2 for _ in range(2))
    png = (b'\x89PNG\r\n\x1a\n'
           + chunk(b'IHDR', struct.pack('>IIBBBBB', 2, 2, 8, 2, 0, 0, 0))
           + chunk(b'IDAT', zlib.compress(raw))
           + chunk(b'IEND', b''))
    with open(path, 'wb') as f:
        f.write(png)


@pytest.fixture
def image_checker(tmp_path):
    """SyllabusChecker over a docx that contains one picture with alt text 'logo'."""
    from docx import Document
    from docx.shared import Inches

    png = tmp_path / 'logo.png'
    _tiny_png(str(png))
    doc = Document()
    doc.add_paragraph('Course syllabus')
    doc.add_picture(str(png), width=Inches(1))
    for el in doc.element.body.iter(
            '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}docPr'):
        el.set('descr', 'logo')
    path = tmp_path / 'with_image.docx'
    doc.save(str(path))

    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    template = os.path.join(root, 'assets', 'Uniform-Syllabus-Template-1.docx')
    return SyllabusChecker(template, str(path))


class TestAnalyzeImageAltText:

    def test_success_includes_alt_text_in_prompt_and_image_count(
            self, image_checker, monkeypatch, mocker):
        monkeypatch.setenv('ANTHROPIC_API_KEY', 'test-key')
        mock_client = mocker.patch('llm.client.get_client').return_value
        mock_client.messages.create.return_value = _make_response(
            'Image 1 needs better alt text.')

        result = image_checker.analyze_image_alt_text()

        assert result == {
            'analysis': 'Image 1 needs better alt text.',
            'status': 'success',
            'image_count': 1,
        }
        kwargs = mock_client.messages.create.call_args.kwargs
        assert 'Alt text: "logo"' in kwargs['messages'][0]['content']

    def test_no_images_returns_early_without_calling_api(self, checker, monkeypatch, mocker):
        monkeypatch.setenv('ANTHROPIC_API_KEY', 'test-key')
        mock_get_client = mocker.patch('llm.client.get_client')

        result = checker.analyze_image_alt_text()

        assert result == {'analysis': 'No images found in the document.', 'status': 'success'}
        mock_get_client.assert_not_called()

    def test_returns_error_when_api_key_absent(self, image_checker, monkeypatch):
        monkeypatch.delenv('ANTHROPIC_API_KEY', raising=False)

        result = image_checker.analyze_image_alt_text()

        assert result['status'] == 'error'
        assert 'ANTHROPIC_API_KEY' in result['error']

    def test_returns_error_on_api_exception(self, image_checker, monkeypatch, mocker):
        monkeypatch.setenv('ANTHROPIC_API_KEY', 'test-key')
        mock_client = mocker.patch('llm.client.get_client').return_value
        mock_client.messages.create.side_effect = RuntimeError('boom')

        result = image_checker.analyze_image_alt_text()

        assert result['status'] == 'error'
        assert 'Claude API' in result['error']


# ---------------------------------------------------------------------------
# Request shape: system prompt, max_tokens, model, structured output
# ---------------------------------------------------------------------------

class TestRequestShape:

    def _call(self, monkeypatch, mocker, fn, text):
        monkeypatch.setenv('ANTHROPIC_API_KEY', 'test-key')
        monkeypatch.delenv('CLAUDE_MODEL', raising=False)
        monkeypatch.delenv('CLAUDE_FAST_MODEL', raising=False)
        mock_client = mocker.patch('llm.client.get_client').return_value
        mock_client.messages.create.return_value = _make_response(text)
        fn()
        return mock_client.messages.create.call_args.kwargs

    @pytest.mark.parametrize('method', ['analyze_growth_mindset_and_belonging',
                                        'analyze_syllabus_quality'])
    def test_analysis_calls_use_system_param_and_primary_model(
            self, checker, monkeypatch, mocker, method):
        kw = self._call(monkeypatch, mocker, getattr(checker, method), 'ok')
        assert kw['model'] == 'claude-sonnet-5-5'
        assert isinstance(kw['system'], str) and kw['system']
        assert all(m['role'] != 'system' for m in kw['messages'])
        assert kw['max_tokens'] >= 4000
        assert 'temperature' not in kw and 'output_config' not in kw

    def test_metadata_uses_fast_model_and_json_schema(self, checker, monkeypatch, mocker):
        kw = self._call(monkeypatch, mocker, checker.extract_course_metadata,
                        TestExtractCourseMetadata.VALID_JSON)
        assert kw['model'] == 'claude-haiku-5-5'
        assert kw['output_config']['format']['type'] == 'json_schema'
        props = kw['output_config']['format']['schema']['properties']
        assert set(props) == {'instructor_name', 'department', 'course_subject',
                              'course_number', 'course_title', 'semester'}

    def test_model_env_vars_override_defaults(self, checker, monkeypatch, mocker):
        monkeypatch.setenv('CLAUDE_MODEL', 'my-main')
        monkeypatch.setenv('CLAUDE_FAST_MODEL', 'my-fast')
        monkeypatch.setenv('ANTHROPIC_API_KEY', 'test-key')
        mock_client = mocker.patch('llm.client.get_client').return_value
        mock_client.messages.create.return_value = _make_response(
            TestExtractCourseMetadata.VALID_JSON)

        checker.extract_course_metadata()
        assert mock_client.messages.create.call_args.kwargs['model'] == 'my-fast'
        checker.analyze_syllabus_quality()
        assert mock_client.messages.create.call_args.kwargs['model'] == 'my-main'

    def test_truncated_metadata_json_is_reported_as_error(self, checker, monkeypatch, mocker):
        monkeypatch.setenv('ANTHROPIC_API_KEY', 'test-key')
        mock_client = mocker.patch('llm.client.get_client').return_value
        resp = _make_response('{"instructor_name": "Dr. T')
        resp.stop_reason = 'max_tokens'
        mock_client.messages.create.return_value = resp

        result = checker.extract_course_metadata()

        assert result['status'] == 'error'
