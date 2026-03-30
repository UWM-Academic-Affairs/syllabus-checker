"""
Tests for admin routes:
  GET  /admin/analytics
  GET  /admin/analytics?department=X
  GET  /admin/analytics/csv
  POST /admin/analytics/delete
"""

import csv
import io
import pytest
import analytics_db


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _store(filename='test.docx', **kwargs):
    return analytics_db.store_scan_result(filename=filename, **kwargs)


# ---------------------------------------------------------------------------
# GET /admin/analytics
# ---------------------------------------------------------------------------

class TestAdminAnalytics:

    def test_returns_200_with_empty_db(self, client):
        resp = client.get('/admin/analytics')
        assert resp.status_code == 200

    def test_contains_scan_statistics(self, client):
        _store('a.docx', total_issues=3)
        _store('b.docx', total_issues=7)

        resp = client.get('/admin/analytics')

        assert resp.status_code == 200
        # Template renders filenames and issue counts in the recent-scans table
        assert b'a.docx' in resp.data
        assert b'b.docx' in resp.data

    def test_department_filter_shows_only_matching_scans(self, client):
        _store('eng.docx', department='English')
        _store('math.docx', department='Math')

        resp = client.get('/admin/analytics?department=English')

        assert resp.status_code == 200
        assert b'eng.docx' in resp.data
        assert b'math.docx' not in resp.data

    def test_department_filter_unknown_returns_200(self, client):
        _store('x.docx', department='Biology')

        resp = client.get('/admin/analytics?department=Chemistry')

        assert resp.status_code == 200
        assert b'x.docx' not in resp.data


# ---------------------------------------------------------------------------
# GET /admin/analytics/csv
# ---------------------------------------------------------------------------

class TestAdminAnalyticsCsv:

    EXPECTED_HEADERS = [
        'Date', 'Course Subject', 'Course Number', 'Course Title',
        'Instructor', 'Department', 'Semester', 'Missing Sections',
        'Language & Tone Score', 'Quality Issues', 'Accessibility Issues',
        'Total Issues', 'Filename',
    ]

    def test_returns_csv_content_type(self, client):
        resp = client.get('/admin/analytics/csv')
        assert resp.status_code == 200
        assert 'text/csv' in resp.content_type

    def test_has_correct_headers(self, client):
        resp = client.get('/admin/analytics/csv')
        reader = csv.reader(io.StringIO(resp.data.decode()))
        headers = next(reader)
        assert headers == self.EXPECTED_HEADERS

    def test_empty_db_returns_header_row_only(self, client):
        resp = client.get('/admin/analytics/csv')
        reader = csv.reader(io.StringIO(resp.data.decode()))
        rows = list(reader)
        assert len(rows) == 1  # header only

    def test_data_rows_match_stored_scans(self, client):
        _store('syllabus.docx', instructor_name='Dr. Smith',
               department='English', total_issues=5)

        resp = client.get('/admin/analytics/csv')
        reader = csv.reader(io.StringIO(resp.data.decode()))
        rows = list(reader)

        assert len(rows) == 2  # header + 1 data row
        data = rows[1]
        assert data[self.EXPECTED_HEADERS.index('Instructor')] == 'Dr. Smith'
        assert data[self.EXPECTED_HEADERS.index('Department')] == 'English'
        assert data[self.EXPECTED_HEADERS.index('Filename')] == 'syllabus.docx'

    def test_department_filter_limits_csv_rows(self, client):
        _store('eng.docx', department='English')
        _store('math.docx', department='Math')

        resp = client.get('/admin/analytics/csv?department=English')
        reader = csv.reader(io.StringIO(resp.data.decode()))
        rows = list(reader)

        assert len(rows) == 2  # header + English row only
        assert rows[1][self.EXPECTED_HEADERS.index('Filename')] == 'eng.docx'

    def test_content_disposition_is_attachment(self, client):
        resp = client.get('/admin/analytics/csv')
        assert 'attachment' in resp.headers.get('Content-Disposition', '')
        assert '.csv' in resp.headers.get('Content-Disposition', '')


# ---------------------------------------------------------------------------
# POST /admin/analytics/delete
# ---------------------------------------------------------------------------

class TestAdminAnalyticsDelete:

    def test_valid_scan_ids_are_removed(self, client):
        id1 = _store('a.docx')
        id2 = _store('b.docx')
        id3 = _store('c.docx')

        resp = client.post(
            '/admin/analytics/delete',
            data={'scan_ids': [str(id1), str(id3)]},
            follow_redirects=True,
        )

        assert resp.status_code == 200
        assert analytics_db.get_scan_count() == 1
        assert analytics_db.get_all_scans()[0]['id'] == id2

    def test_non_integer_id_is_skipped_not_500(self, client):
        _store('x.docx')

        resp = client.post(
            '/admin/analytics/delete',
            data={'scan_ids': ['not-an-int', 'abc', '0', '-5']},
            follow_redirects=True,
        )

        # Must not crash; must show a warning flash, not a 500
        assert resp.status_code == 200
        assert analytics_db.get_scan_count() == 1  # nothing deleted

    def test_no_ids_selected_shows_warning(self, client):
        resp = client.post(
            '/admin/analytics/delete',
            data={},
            follow_redirects=True,
        )
        assert resp.status_code == 200
        assert b'selected' in resp.data.lower()

    def test_success_flash_mentions_deleted_count(self, client):
        id1 = _store('a.docx')
        id2 = _store('b.docx')

        resp = client.post(
            '/admin/analytics/delete',
            data={'scan_ids': [str(id1), str(id2)]},
            follow_redirects=True,
        )

        assert resp.status_code == 200
        assert b'2' in resp.data  # "Successfully deleted 2 scan(s)"


# ---------------------------------------------------------------------------
# Repository tab — saved syllabi shown with correct counts
# ---------------------------------------------------------------------------

class TestAdminRepositoryTab:

    def test_saved_syllabi_appear_in_analytics_page(self, client, app):
        scan_id = _store('saved.docx', instructor_name='Prof. Test',
                         department='History')
        # Drop a file in the repo folder so the record is realistic
        repo_dir = app.config['REPO_FOLDER']
        stored = '20250101_120000_saved.docx'
        open(f'{repo_dir}/{stored}', 'wb').close()
        analytics_db.save_to_repository(scan_id, 'saved.docx', stored)

        resp = client.get('/admin/analytics')

        assert resp.status_code == 200
        assert b'saved.docx' in resp.data

    def test_repository_count_reflects_saved_syllabi(self, client, app):
        repo_dir = app.config['REPO_FOLDER']
        for i in range(3):
            sid = _store(f's{i}.docx')
            stored = f'2025010{i+1}_120000_s{i}.docx'
            open(f'{repo_dir}/{stored}', 'wb').close()
            analytics_db.save_to_repository(sid, f's{i}.docx', stored)

        resp = client.get('/admin/analytics')

        assert resp.status_code == 200
        # The template renders the count of saved syllabi somewhere on the page
        assert b'3' in resp.data
