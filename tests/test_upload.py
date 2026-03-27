"""
Tests for the file upload and processing pipeline (/upload route).
"""

import io
import os

import pytest
import analytics_db


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _upload(client, data, *, follow_redirects=False):
    """POST multipart form data to /upload."""
    return client.post(
        '/upload',
        data=data,
        content_type='multipart/form-data',
        follow_redirects=follow_redirects,
    )


def _upload_file(client, path, filename=None, *, follow_redirects=False):
    """POST a file from disk to /upload."""
    if filename is None:
        filename = os.path.basename(path)
    with open(path, 'rb') as fh:
        raw = fh.read()
    return _upload(
        client,
        {'file': (io.BytesIO(raw), filename)},
        follow_redirects=follow_redirects,
    )


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------

class TestUploadValidation:
    """Input-validation checks that never reach the checker."""

    def test_no_file_part_redirects_with_error(self, client):
        resp = _upload(client, {}, follow_redirects=True)
        assert resp.status_code == 200
        assert b'No file uploaded' in resp.data

    def test_empty_filename_redirects_with_error(self, client):
        resp = _upload(
            client,
            {'file': (io.BytesIO(b''), '')},
            follow_redirects=True,
        )
        assert resp.status_code == 200
        assert b'No file selected' in resp.data

    @pytest.mark.parametrize('bad_name', [
        'report.pdf',
        'notes.txt',
        'script.exe',
        'archive.zip',
    ])
    def test_non_docx_rejected_with_user_friendly_error(self, client, bad_name):
        resp = _upload(
            client,
            {'file': (io.BytesIO(b'dummy content'), bad_name)},
            follow_redirects=True,
        )
        assert resp.status_code == 200
        assert b'.docx' in resp.data  # error message mentions .docx

    def test_oversized_upload_rejected(self, client, app):
        """Files over MAX_CONTENT_LENGTH (16 MB) must be refused."""
        # Flask returns 413 before the view function is ever called.
        big = io.BytesIO(b'x' * (16 * 1024 * 1024 + 1))
        resp = _upload(client, {'file': (big, 'huge.docx')})
        assert resp.status_code == 413


class TestUploadProcessing:
    """End-to-end tests that exercise the SyllabusChecker pipeline."""

    def test_valid_docx_returns_200_and_results_page(self, client, fixture_dir):
        resp = _upload_file(client, str(fixture_dir / 'minimal_valid.docx'))
        assert resp.status_code == 200
        # Results template renders the filename and key sections
        assert b'minimal_valid.docx' in resp.data

    def test_marked_document_created_after_upload(self, client, app, fixture_dir):
        _upload_file(client, str(fixture_dir / 'minimal_valid.docx'))
        marked = os.path.join(
            app.config['UPLOAD_FOLDER'], 'minimal_valid_marked.docx'
        )
        assert os.path.exists(marked), 'Marked document was not created'

    def test_scan_record_written_to_database(self, client, fixture_dir):
        assert analytics_db.get_scan_count() == 0
        _upload_file(client, str(fixture_dir / 'minimal_valid.docx'))
        assert analytics_db.get_scan_count() == 1

    def test_corrupted_docx_flashes_error_not_500(self, client, fixture_dir):
        """A file that cannot be parsed must redirect to index, not crash."""
        resp = _upload_file(
            client,
            str(fixture_dir / 'malformed.docx'),
            follow_redirects=True,
        )
        assert resp.status_code == 200
        assert b'error occurred' in resp.data.lower()
