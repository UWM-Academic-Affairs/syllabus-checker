"""
Tests for OpenAI-backed SyllabusChecker methods (all API calls mocked):
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
    """Build a minimal object that mimics an OpenAI chat completion."""
    resp = MagicMock()
    resp.choices = [MagicMock()]
    resp.choices[0].message.content = content
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
        monkeypatch.setenv('OPENAI_API_KEY', 'test-key')
        mock_cls = mocker.patch('syllabus_checker.OpenAI')
        mock_cls.return_value.chat.completions.create.return_value = (
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
        monkeypatch.setenv('OPENAI_API_KEY', 'test-key')
        mock_cls = mocker.patch('syllabus_checker.OpenAI')
        fenced = f'```json\n{self.VALID_JSON}\n```'
        mock_cls.return_value.chat.completions.create.return_value = (
            _make_response(fenced)
        )

        result = checker.extract_course_metadata()

        assert result['status'] == 'success'
        assert result['course_subject'] == 'CS'

    def test_handles_malformed_json_without_crashing(self, checker, monkeypatch, mocker):
        """Garbage JSON must return status='error', not raise an exception."""
        monkeypatch.setenv('OPENAI_API_KEY', 'test-key')
        mock_cls = mocker.patch('syllabus_checker.OpenAI')
        mock_cls.return_value.chat.completions.create.return_value = (
            _make_response('This is not JSON at all!!!')
        )

        result = checker.extract_course_metadata()

        assert result['status'] == 'error'
        assert 'error' in result

    def test_returns_error_when_api_key_absent(self, checker, monkeypatch):
        monkeypatch.delenv('OPENAI_API_KEY', raising=False)

        result = checker.extract_course_metadata()

        assert result['status'] == 'error'

    def test_returns_error_on_api_exception(self, checker, monkeypatch, mocker):
        monkeypatch.setenv('OPENAI_API_KEY', 'test-key')
        mock_cls = mocker.patch('syllabus_checker.OpenAI')
        mock_cls.return_value.chat.completions.create.side_effect = (
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
        monkeypatch.delenv('OPENAI_API_KEY', raising=False)

        result = checker.analyze_sections_with_llm()

        # Must return the expected keys without calling the API
        assert 'present' in result
        assert 'missing' in result
        # The error field signals that pattern-matching fallback was used
        assert 'error' in result
        assert 'pattern' in result['error'].lower() or 'openai' in result['error'].lower()

    def test_present_list_is_non_empty_for_complete_syllabus(self, checker, monkeypatch):
        """minimal_valid.docx has many detectable sections; present list must be non-empty."""
        monkeypatch.delenv('OPENAI_API_KEY', raising=False)

        result = checker.analyze_sections_with_llm()

        assert isinstance(result['present'], list)
        assert len(result['present']) > 0

    def test_missing_list_is_a_list(self, checker, monkeypatch):
        monkeypatch.delenv('OPENAI_API_KEY', raising=False)

        result = checker.analyze_sections_with_llm()

        assert isinstance(result['missing'], list)


# ---------------------------------------------------------------------------
# analyze_growth_mindset_and_belonging
# ---------------------------------------------------------------------------

class TestAnalyzeGrowthMindset:

    def test_returns_error_status_when_api_key_absent(self, checker, monkeypatch):
        monkeypatch.delenv('OPENAI_API_KEY', raising=False)

        result = checker.analyze_growth_mindset_and_belonging()

        assert result['status'] == 'error'
        assert 'error' in result

    def test_returns_error_status_when_api_call_raises(self, checker, monkeypatch, mocker):
        """Any exception from the OpenAI call must produce status='error'."""
        monkeypatch.setenv('OPENAI_API_KEY', 'test-key')
        mock_cls = mocker.patch('syllabus_checker.OpenAI')
        mock_cls.return_value.chat.completions.create.side_effect = (
            Exception('timeout')
        )

        result = checker.analyze_growth_mindset_and_belonging()

        assert result['status'] == 'error'
        assert 'error' in result

    def test_returns_success_with_analysis_text(self, checker, monkeypatch, mocker):
        monkeypatch.setenv('OPENAI_API_KEY', 'test-key')
        mock_cls = mocker.patch('syllabus_checker.OpenAI')
        mock_cls.return_value.chat.completions.create.return_value = (
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
        monkeypatch.delenv('OPENAI_API_KEY', raising=False)

        result = checker.analyze_syllabus_quality()

        assert result['status'] == 'error'
        assert 'error' in result

    def test_returns_error_gracefully_on_timeout(self, checker, monkeypatch, mocker):
        """A timeout (or any exception) must produce status='error', not a 500."""
        monkeypatch.setenv('OPENAI_API_KEY', 'test-key')
        mock_cls = mocker.patch('syllabus_checker.OpenAI')
        mock_cls.return_value.chat.completions.create.side_effect = (
            TimeoutError('request timed out')
        )

        result = checker.analyze_syllabus_quality()

        assert result['status'] == 'error'
        assert 'error' in result

    def test_returns_success_with_analysis_text(self, checker, monkeypatch, mocker):
        monkeypatch.setenv('OPENAI_API_KEY', 'test-key')
        mock_cls = mocker.patch('syllabus_checker.OpenAI')
        mock_cls.return_value.chat.completions.create.return_value = (
            _make_response('No significant quality issues found.')
        )

        result = checker.analyze_syllabus_quality()

        assert result['status'] == 'success'
        assert 'analysis' in result
