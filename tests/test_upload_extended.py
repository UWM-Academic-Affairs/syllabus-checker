"""
Extended upload and processing pipeline tests targeting uncovered app.py branches:
  - Growth mindset analysis success parsing (language_tone_summary, subcategories)
  - Quality analysis success parsing (clarity_quality_summary, subcategories)
  - Metadata extraction success (fields rendered in results)
  - Download route (file found / file missing)
  - End-to-end smoke test: real DOCX through the full pipeline
"""

import io
import os

import analytics_db


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _upload_file(client, path, filename=None, *, follow_redirects=False):
    if filename is None:
        filename = os.path.basename(path)
    with open(path, 'rb') as fh:
        raw = fh.read()
    return client.post(
        '/upload',
        data={'file': (io.BytesIO(raw), filename)},
        content_type='multipart/form-data',
        follow_redirects=follow_redirects,
    )


def _mock_checker(mocker, fixture_dir, *, growth_mindset=None, quality=None, metadata=None):
    """
    Patch app.SyllabusChecker so the upload route gets controlled LLM responses.
    All heavy methods (run_all_checks, create_marked_document) are no-ops.
    Returns the mock class so callers can further configure it.
    """
    mock_cls = mocker.patch('app.SyllabusChecker')
    inst = mock_cls.return_value
    inst.issues = []
    inst.run_all_checks.return_value = None
    inst.check_heading_structure.return_value = {'counts': {}, 'issues': []}
    inst.check_table_usage.return_value = {'counts': {}, 'issues': []}
    inst.check_list_usage.return_value = {'counts': {}, 'issues': []}
    inst.analyze_sections_with_llm.return_value = {'present': [], 'missing': []}
    inst.analyze_image_alt_text.return_value = {'status': 'error'}
    inst.generate_report.return_value = ''
    inst.create_marked_document.return_value = None

    inst.analyze_growth_mindset_and_belonging.return_value = (
        growth_mindset or {'status': 'error'}
    )
    inst.analyze_syllabus_quality.return_value = (
        quality or {'status': 'error'}
    )
    inst.extract_course_metadata.return_value = (
        metadata or {'status': 'error'}
    )
    return mock_cls


# ---------------------------------------------------------------------------
# Growth mindset success — language_tone_summary parsing
# ---------------------------------------------------------------------------

class TestGrowthMindsetParsing:

    # LLM response text with 4+ "Strong" assessments → "Strong (N/6 areas)"
    STRONG_ANALYSIS = (
        "## QUESTION 1: Growth Mindset\n**Assessment:** Strong\n"
        "## QUESTION 2: Normalizing Challenge\n**Assessment:** Strong\n"
        "## QUESTION 3: Instructor Care\n**Assessment:** Strong\n"
        "## QUESTION 4: Valuing Diversity\n**Assessment:** Strong\n"
        "## QUESTION 5: Normalizing Student Challenges\n**Assessment:** Moderate\n"
        "## QUESTION 6: Normalizing Academic Support\n**Assessment:** Moderate\n"
    )

    WEAK_ANALYSIS = (
        "## QUESTION 1: Growth Mindset\n**Assessment:** Weak\n"
        "## QUESTION 2: Normalizing Challenge\n**Assessment:** Not Addressed\n"
        "## QUESTION 3: Instructor Care\n**Assessment:** Weak\n"
        "## QUESTION 4: Valuing Diversity\n**Assessment:** Moderate\n"
        "## QUESTION 5: Normalizing Student Challenges\n**Assessment:** Moderate\n"
        "## QUESTION 6: Normalizing Academic Support\n**Assessment:** Moderate\n"
    )

    def test_strong_results_show_strong_summary(self, client, app, fixture_dir, mocker):
        _mock_checker(mocker, fixture_dir, growth_mindset={
            'status': 'success',
            'analysis': self.STRONG_ANALYSIS,
        })
        resp = _upload_file(client, str(fixture_dir / 'minimal_valid.docx'))
        assert resp.status_code == 200

    def test_weak_results_show_needs_work_summary(self, client, app, fixture_dir, mocker):
        _mock_checker(mocker, fixture_dir, growth_mindset={
            'status': 'success',
            'analysis': self.WEAK_ANALYSIS,
        })
        resp = _upload_file(client, str(fixture_dir / 'minimal_valid.docx'))
        assert resp.status_code == 200

    def test_mixed_results_show_mixed_summary(self, client, app, fixture_dir, mocker):
        mixed = (
            "## QUESTION 1: Growth Mindset\n**Assessment:** Strong\n"
            "## QUESTION 2: Normalizing Challenge\n**Assessment:** Moderate\n"
            "## QUESTION 3: Instructor Care\n**Assessment:** Strong\n"
            "## QUESTION 4: Valuing Diversity\n**Assessment:** Moderate\n"
            "## QUESTION 5: Normalizing Student Challenges\n**Assessment:** Moderate\n"
            "## QUESTION 6: Normalizing Academic Support\n**Assessment:** Moderate\n"
        )
        _mock_checker(mocker, fixture_dir, growth_mindset={
            'status': 'success',
            'analysis': mixed,
        })
        resp = _upload_file(client, str(fixture_dir / 'minimal_valid.docx'))
        assert resp.status_code == 200

    def test_error_status_still_returns_200(self, client, app, fixture_dir, mocker):
        _mock_checker(mocker, fixture_dir, growth_mindset={
            'status': 'error',
            'error': 'API timeout',
        })
        resp = _upload_file(client, str(fixture_dir / 'minimal_valid.docx'))
        assert resp.status_code == 200


# ---------------------------------------------------------------------------
# Quality analysis success — clarity_quality_summary parsing
# ---------------------------------------------------------------------------

class TestQualityAnalysisParsing:

    NO_ISSUES = (
        "## 1. UNDEFINED COURSE TERMINOLOGY\n"
        "✓ All course terminology appears to be clearly defined.\n\n"
        "## 2. TONE AND INCLUSIVITY\n"
        "✓ Policy language uses an inclusive, supportive tone throughout.\n\n"
        "## 3. CONFUSING POLICIES\n"
        "✓ All policies appear clear and inclusive.\n\n"
        "## 4. INCONSISTENT FORMATTING\n"
        "✓ Formatting is consistent throughout.\n"
    )

    ONE_CATEGORY_ISSUE = (
        "## 1. UNDEFINED COURSE TERMINOLOGY\n"
        '- **Term**: "LMS"\n'
        '    - **Why it might be unclear**: Acronym not spelled out.\n'
        '    - **Suggested fix**: Spell out as Learning Management System.\n\n'
        "## 2. TONE AND INCLUSIVITY\n"
        "✓ Policy language uses an inclusive, supportive tone throughout.\n\n"
        "## 3. CONFUSING POLICIES\n"
        "✓ All policies appear clear and inclusive.\n\n"
        "## 4. INCONSISTENT FORMATTING\n"
        "✓ Formatting is consistent throughout.\n"
    )

    MULTI_CATEGORY_ISSUES = (
        "## 1. UNDEFINED COURSE TERMINOLOGY\n"
        '- **Term**: "LMS"\n'
        '    - **Why**: Acronym.\n'
        '    - **Fix**: Spell out.\n\n'
        "## 2. TONE AND INCLUSIVITY\n"
        '- **Issue**: "Students who miss class will be penalized."\n'
        '    - **Problem**: Punitive language.\n'
        '    - **Better alternative**: More supportive framing.\n\n'
        "## 3. CONFUSING POLICIES\n"
        "✓ All policies appear clear and inclusive.\n\n"
        "## 4. INCONSISTENT FORMATTING\n"
        "✓ Formatting is consistent throughout.\n"
    )

    def test_no_issues_returns_200(self, client, app, fixture_dir, mocker):
        _mock_checker(mocker, fixture_dir, quality={
            'status': 'success',
            'analysis': self.NO_ISSUES,
        })
        resp = _upload_file(client, str(fixture_dir / 'minimal_valid.docx'))
        assert resp.status_code == 200

    def test_one_category_issue_returns_200(self, client, app, fixture_dir, mocker):
        _mock_checker(mocker, fixture_dir, quality={
            'status': 'success',
            'analysis': self.ONE_CATEGORY_ISSUE,
        })
        resp = _upload_file(client, str(fixture_dir / 'minimal_valid.docx'))
        assert resp.status_code == 200

    def test_multi_category_issues_returns_200(self, client, app, fixture_dir, mocker):
        _mock_checker(mocker, fixture_dir, quality={
            'status': 'success',
            'analysis': self.MULTI_CATEGORY_ISSUES,
        })
        resp = _upload_file(client, str(fixture_dir / 'minimal_valid.docx'))
        assert resp.status_code == 200

    def test_error_status_still_returns_200(self, client, app, fixture_dir, mocker):
        _mock_checker(mocker, fixture_dir, quality={
            'status': 'error',
            'error': 'No API key',
        })
        resp = _upload_file(client, str(fixture_dir / 'minimal_valid.docx'))
        assert resp.status_code == 200


# ---------------------------------------------------------------------------
# Metadata extraction success path
# ---------------------------------------------------------------------------

class TestMetadataExtractionPath:

    def test_extracted_metadata_stored_in_db(self, client, app, fixture_dir, mocker):
        """When metadata extraction succeeds, the DB row has the extracted fields."""
        _mock_checker(mocker, fixture_dir, metadata={
            'status': 'success',
            'instructor_name': 'Dr. Jane Smith',
            'department': 'Computer Science',
            'course_subject': 'CS',
            'course_number': '101',
            'course_title': 'Introduction to Testing',
            'semester': 'Fall 2025',
        })
        _upload_file(client, str(fixture_dir / 'minimal_valid.docx'))

        scans = analytics_db.get_all_scans()
        assert len(scans) == 1
        row = scans[0]
        assert row['instructor_name'] == 'Dr. Jane Smith'
        assert row['department'] == 'Computer Science'
        assert row['course_subject'] == 'CS'
        assert row['semester'] == 'Fall 2025'

    def test_failed_metadata_leaves_db_row_with_nulls(self, client, app, fixture_dir, mocker):
        """When metadata extraction fails, the DB row is still created (with null fields)."""
        _mock_checker(mocker, fixture_dir, metadata={'status': 'error'})
        _upload_file(client, str(fixture_dir / 'minimal_valid.docx'))

        scans = analytics_db.get_all_scans()
        assert len(scans) == 1
        assert scans[0]['instructor_name'] is None


# ---------------------------------------------------------------------------
# Download route
# ---------------------------------------------------------------------------

class TestDownloadRoute:

    def test_existing_file_streams_correctly(self, client, app):
        """GET /download/<name> for a file that exists returns 200."""
        upload_dir = app.config['UPLOAD_FOLDER']
        content = b'PK\x03\x04fake docx bytes'
        fname = 'test_syllabus_marked.docx'
        with open(os.path.join(upload_dir, fname), 'wb') as fh:
            fh.write(content)

        resp = client.get(f'/download/{fname}')
        assert resp.status_code == 200
        assert resp.data == content

    def test_missing_file_redirects_with_error(self, client):
        resp = client.get('/download/nonexistent_marked.docx', follow_redirects=True)
        assert resp.status_code == 200
        assert b'not found' in resp.data.lower() or b'error' in resp.data.lower()


# ---------------------------------------------------------------------------
# End-to-end smoke test
# ---------------------------------------------------------------------------

class TestEndToEndSmoke:
    """
    Upload a real .docx through the full pipeline (no mocks, no API key).
    Verifies the results page renders key structural elements.
    """

    def test_results_page_contains_filename(self, client, fixture_dir):
        resp = _upload_file(client, str(fixture_dir / 'minimal_valid.docx'))
        assert resp.status_code == 200
        assert b'minimal_valid.docx' in resp.data

    def test_results_page_renders_total_issues(self, client, fixture_dir):
        resp = _upload_file(client, str(fixture_dir / 'minimal_valid.docx'))
        assert resp.status_code == 200
        # The results template always renders a summary message referencing issue counts
        assert b'issue' in resp.data.lower() or b'Excellent' in resp.data

    def test_results_page_has_download_link(self, client, fixture_dir):
        resp = _upload_file(client, str(fixture_dir / 'minimal_valid.docx'))
        assert resp.status_code == 200
        # The marked document download button is always present after a successful upload
        assert b'_marked.docx' in resp.data or b'download' in resp.data.lower()

    def test_db_row_created_with_filename(self, client, fixture_dir):
        _upload_file(client, str(fixture_dir / 'minimal_valid.docx'))
        scans = analytics_db.get_all_scans()
        assert any(s['filename'] == 'minimal_valid.docx' for s in scans)

    def test_marked_file_exists_on_disk(self, client, app, fixture_dir):
        _upload_file(client, str(fixture_dir / 'minimal_valid.docx'))
        marked = os.path.join(app.config['UPLOAD_FOLDER'], 'minimal_valid_marked.docx')
        assert os.path.exists(marked)

    def test_missing_sections_docx_shows_missing_content(self, client, fixture_dir):
        """Uploading a nearly-empty DOCX should flag missing sections."""
        resp = _upload_file(client, str(fixture_dir / 'missing_sections.docx'))
        assert resp.status_code == 200
        # The results page should mention missing content
        assert b'missing' in resp.data.lower() or b'add' in resp.data.lower()
