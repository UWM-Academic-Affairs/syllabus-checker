"""
Tests for the repository routes:
  POST /repository/save
  GET  /repository/download/<name>
  POST /repository/delete/<id>
"""

import os
import analytics_db


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _put_file(folder, filename, content=b'fake docx content'):
    """Write a dummy file into folder and return its full path."""
    os.makedirs(folder, exist_ok=True)
    path = os.path.join(folder, filename)
    with open(path, 'wb') as fh:
        fh.write(content)
    return path


def _save(client, *, filename, marked_filename='', scan_id=None):
    """POST to /repository/save."""
    data = {'filename': filename, 'marked_filename': marked_filename}
    if scan_id is not None:
        data['scan_id'] = str(scan_id)
    return client.post('/repository/save', data=data)


# ---------------------------------------------------------------------------
# POST /repository/save
# ---------------------------------------------------------------------------

class TestRepositorySave:

    def test_copies_original_file_to_repo_folder(self, client, app):
        upload_dir = app.config['UPLOAD_FOLDER']
        repo_dir = app.config['REPO_FOLDER']
        _put_file(upload_dir, 'syllabus.docx')

        resp = _save(client, filename='syllabus.docx')

        assert resp.status_code == 204
        stored = os.listdir(repo_dir)
        assert len(stored) == 1
        assert stored[0].endswith('syllabus.docx')

    def test_copies_marked_file_when_present(self, client, app):
        upload_dir = app.config['UPLOAD_FOLDER']
        repo_dir = app.config['REPO_FOLDER']
        _put_file(upload_dir, 'syllabus.docx')
        _put_file(upload_dir, 'syllabus_marked.docx')

        resp = _save(
            client,
            filename='syllabus.docx',
            marked_filename='syllabus_marked.docx',
        )

        assert resp.status_code == 204
        stored = os.listdir(repo_dir)
        originals = [f for f in stored if not f.endswith('_marked.docx')]
        marked = [f for f in stored if f.endswith('_marked.docx')]
        assert len(originals) == 1
        assert len(marked) == 1

    def test_missing_filename_returns_400(self, client):
        resp = _save(client, filename='')
        assert resp.status_code == 400

    def test_original_file_gone_returns_400(self, client, app):
        # File is NOT placed in upload dir — simulates expired/swept file
        resp = _save(client, filename='ghost.docx')
        assert resp.status_code == 400

    def test_db_row_created_with_correct_filenames(self, client, app):
        upload_dir = app.config['UPLOAD_FOLDER']
        scan_id = analytics_db.store_scan_result('syllabus.docx')
        _put_file(upload_dir, 'syllabus.docx')
        _put_file(upload_dir, 'syllabus_marked.docx')

        _save(
            client,
            filename='syllabus.docx',
            marked_filename='syllabus_marked.docx',
            scan_id=scan_id,
        )

        assert analytics_db.is_saved(scan_id)
        rows = analytics_db.get_saved_syllabi()
        assert len(rows) == 1
        assert rows[0]['original_filename'] == 'syllabus.docx'
        assert rows[0]['stored_marked_filename'].endswith('syllabus_marked.docx')


# ---------------------------------------------------------------------------
# GET /repository/download/<name>
# ---------------------------------------------------------------------------

class TestRepositoryDownload:

    def test_streams_correct_file(self, client, app):
        repo_dir = app.config['REPO_FOLDER']
        content = b'PK\x03\x04fake docx bytes'
        _put_file(repo_dir, '20250101_120000_syllabus.docx', content=content)

        resp = client.get('/repository/download/20250101_120000_syllabus.docx')

        assert resp.status_code == 200
        assert resp.data == content

    def test_missing_file_redirects_with_error(self, client, app):
        resp = client.get(
            '/repository/download/nonexistent.docx',
            follow_redirects=True,
        )
        assert resp.status_code == 200
        assert b'not found' in resp.data.lower()

    def test_path_traversal_is_neutralized(self, client, app):
        """A ../../../etc/passwd style traversal must not stream a file."""
        # secure_filename() strips path separators, so the resolved path
        # will be inside the repo folder but the file won't exist there.
        resp = client.get('/repository/download/../../../etc/passwd')
        # Must redirect, not stream file content
        assert resp.status_code in (301, 302, 200)
        # If 200, it must be the admin page (redirect followed), not the file
        if resp.status_code == 200:
            assert b'PK' not in resp.data  # no ZIP/DOCX magic bytes

    def test_path_traversal_does_not_expose_sensitive_file(self, client, app):
        """Even if /etc/passwd exists on the system, it must not be returned."""
        resp = client.get(
            '/repository/download/../../../etc/passwd',
            follow_redirects=False,
        )
        # Route must redirect, never stream the sensitive file
        assert resp.status_code in (301, 302)


# ---------------------------------------------------------------------------
# POST /repository/delete/<id>
# ---------------------------------------------------------------------------

class TestRepositoryDelete:

    def _seed(self, app, original='orig.docx', marked='orig_marked.docx'):
        """Insert a scan + repo row and put matching files in the repo folder."""
        repo_dir = app.config['REPO_FOLDER']
        scan_id = analytics_db.store_scan_result(original)
        stored = f'20250101_120000_{original}'
        stored_marked = f'20250101_120000_{marked}'
        _put_file(repo_dir, stored)
        _put_file(repo_dir, stored_marked)
        row_id = analytics_db.save_to_repository(
            scan_id, original, stored, stored_marked
        )
        return row_id, stored, stored_marked

    def test_removes_db_row_and_both_files(self, client, app):
        repo_dir = app.config['REPO_FOLDER']
        row_id, stored, stored_marked = self._seed(app)

        resp = client.post(f'/repository/delete/{row_id}')

        assert resp.status_code == 204
        assert analytics_db.get_saved_syllabi() == []
        assert not os.path.exists(os.path.join(repo_dir, stored))
        assert not os.path.exists(os.path.join(repo_dir, stored_marked))

    def test_nonexistent_id_returns_gracefully(self, client):
        resp = client.post('/repository/delete/99999')
        assert resp.status_code == 204  # no-op, not a 404 or 500

    def test_delete_leaves_other_repo_entries_intact(self, client, app):
        row_id_a, _, _ = self._seed(app, 'a.docx', 'a_marked.docx')
        row_id_b, _, _ = self._seed(app, 'b.docx', 'b_marked.docx')

        client.post(f'/repository/delete/{row_id_a}')

        remaining = analytics_db.get_saved_syllabi()
        assert len(remaining) == 1
        assert remaining[0]['id'] == row_id_b
