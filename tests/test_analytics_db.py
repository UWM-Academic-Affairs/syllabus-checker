"""
Unit tests for analytics_db.py — all functions that read/write the SQLite DB.
Each test gets a fresh in-memory DB via the `db` fixture from conftest.
"""

import sqlite3
import analytics_db


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _store(filename='test.docx', **kwargs):
    """Call store_scan_result with sensible defaults."""
    return analytics_db.store_scan_result(filename=filename, **kwargs)


def _set_dates(db_path, date_by_id: dict):
    """Directly set scan_date for specific rows so ordering tests are stable."""
    conn = sqlite3.connect(db_path)
    for scan_id, date_str in date_by_id.items():
        conn.execute(
            "UPDATE scans SET scan_date = ? WHERE id = ?", (date_str, scan_id)
        )
    conn.commit()
    conn.close()


# ---------------------------------------------------------------------------
# store_scan_result
# ---------------------------------------------------------------------------

class TestStoreScanResult:
    def test_returns_positive_integer_scan_id(self, db):
        scan_id = _store()
        assert isinstance(scan_id, int)
        assert scan_id > 0

    def test_each_call_returns_unique_id(self, db):
        ids = [_store(f'f{i}.docx') for i in range(3)]
        assert len(set(ids)) == 3

    def test_stored_values_are_correct(self, db):
        _store(
            'syllabus.docx',
            instructor_name='Dr. Smith',
            department='English',
            total_issues=4,
            missing_sections=['Grading Scale', 'Calendar'],
        )
        rows = analytics_db.get_all_scans()
        assert len(rows) == 1
        row = rows[0]
        assert row['filename'] == 'syllabus.docx'
        assert row['instructor_name'] == 'Dr. Smith'
        assert row['department'] == 'English'
        assert row['total_issues'] == 4
        assert row['missing_sections_count'] == 2


# ---------------------------------------------------------------------------
# get_all_scans / get_scan_count
# ---------------------------------------------------------------------------

class TestGetAllScans:
    def test_returns_rows_in_descending_date_order(self, db):
        id1 = _store('a.docx')
        id2 = _store('b.docx')
        id3 = _store('c.docx')
        _set_dates(db, {
            id1: '2025-01-01 10:00:00',
            id2: '2025-02-01 10:00:00',
            id3: '2025-03-01 10:00:00',
        })

        scans = analytics_db.get_all_scans()
        assert [s['id'] for s in scans] == [id3, id2, id1]

    def test_department_filter_returns_only_matching_rows(self, db):
        _store('eng.docx', department='English')
        _store('math.docx', department='Math')
        _store('eng2.docx', department='English')

        results = analytics_db.get_all_scans(department='English')
        assert len(results) == 2
        assert all(r['department'] == 'English' for r in results)

    def test_department_filter_no_match_returns_empty(self, db):
        _store('x.docx', department='Biology')
        assert analytics_db.get_all_scans(department='Chemistry') == []

    def test_empty_db_returns_empty_list(self, db):
        assert analytics_db.get_all_scans() == []


class TestGetScanCount:
    def test_returns_correct_total(self, db):
        assert analytics_db.get_scan_count() == 0
        _store()
        _store()
        assert analytics_db.get_scan_count() == 2

    def test_department_filter(self, db):
        _store('a.docx', department='Art')
        _store('b.docx', department='Art')
        _store('c.docx', department='Biology')

        assert analytics_db.get_scan_count(department='Art') == 2
        assert analytics_db.get_scan_count(department='Biology') == 1
        assert analytics_db.get_scan_count() == 3


# ---------------------------------------------------------------------------
# delete_scans
# ---------------------------------------------------------------------------

class TestDeleteScans:
    def test_removes_exactly_the_specified_rows(self, db):
        id1 = _store('a.docx')
        id2 = _store('b.docx')
        id3 = _store('c.docx')

        deleted = analytics_db.delete_scans([id1, id3])

        assert deleted == 2
        remaining = analytics_db.get_all_scans()
        assert len(remaining) == 1
        assert remaining[0]['id'] == id2

    def test_empty_list_deletes_nothing(self, db):
        _store()
        analytics_db.delete_scans([])
        assert analytics_db.get_scan_count() == 1

    def test_nonexistent_id_is_a_no_op(self, db):
        _store()
        deleted = analytics_db.delete_scans([99999])
        assert deleted == 0
        assert analytics_db.get_scan_count() == 1


# ---------------------------------------------------------------------------
# Repository: save_to_repository / is_saved / get_saved_syllabi / delete_from_repository
# ---------------------------------------------------------------------------

class TestRepository:
    def test_save_to_repository_inserts_row_with_correct_filenames(self, db):
        scan_id = _store('orig.docx')
        row_id = analytics_db.save_to_repository(
            scan_id=scan_id,
            original_filename='orig.docx',
            stored_filename='stored_orig.docx',
            stored_marked_filename='stored_orig_marked.docx',
        )
        assert isinstance(row_id, int) and row_id > 0

        syllabi = analytics_db.get_saved_syllabi()
        assert len(syllabi) == 1
        s = syllabi[0]
        assert s['original_filename'] == 'orig.docx'
        assert s['stored_filename'] == 'stored_orig.docx'
        assert s['stored_marked_filename'] == 'stored_orig_marked.docx'

    def test_is_saved_returns_false_before_save(self, db):
        scan_id = _store()
        assert analytics_db.is_saved(scan_id) is False

    def test_is_saved_returns_true_after_save(self, db):
        scan_id = _store()
        analytics_db.save_to_repository(scan_id, 'o.docx', 'stored.docx')
        assert analytics_db.is_saved(scan_id) is True

    def test_is_saved_with_none_returns_false(self, db):
        assert analytics_db.is_saved(None) is False

    def test_delete_from_repository_removes_row_and_returns_filenames(self, db):
        scan_id = _store()
        row_id = analytics_db.save_to_repository(
            scan_id, 'o.docx', 'stored.docx', 'stored_marked.docx'
        )

        returned = analytics_db.delete_from_repository(row_id)

        assert returned['stored_filename'] == 'stored.docx'
        assert returned['stored_marked_filename'] == 'stored_marked.docx'
        assert analytics_db.get_saved_syllabi() == []

    def test_delete_from_repository_nonexistent_returns_none(self, db):
        assert analytics_db.delete_from_repository(99999) is None

    def test_get_saved_syllabi_joins_scan_metadata(self, db):
        scan_id = _store(
            'join_test.docx',
            instructor_name='Prof. Test',
            department='History',
        )
        analytics_db.save_to_repository(scan_id, 'join_test.docx', 'stored.docx')

        rows = analytics_db.get_saved_syllabi()
        assert len(rows) == 1
        row = rows[0]
        assert row['instructor_name'] == 'Prof. Test'
        assert row['department'] == 'History'

    def test_get_saved_syllabi_without_marked_file(self, db):
        scan_id = _store()
        analytics_db.save_to_repository(scan_id, 'o.docx', 'stored.docx')
        rows = analytics_db.get_saved_syllabi()
        assert rows[0]['stored_marked_filename'] is None


# ---------------------------------------------------------------------------
# Schema migration
# ---------------------------------------------------------------------------

class TestSchemaMigration:
    def test_init_db_on_existing_db_adds_missing_columns(self, tmp_path, monkeypatch):
        """
        Simulate an older DB that lacks the newer columns.
        init_db() must add them without raising.
        """
        db_path = str(tmp_path / 'old.db')
        monkeypatch.setattr(analytics_db, 'DATABASE_PATH', db_path)

        # Simulate an older schema: has the original core columns but lacks
        # the later additions (course_title, semester, growth_mindset_score, …).
        conn = sqlite3.connect(db_path)
        conn.execute('''
            CREATE TABLE scans (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                scan_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                filename TEXT,
                instructor_name TEXT,
                department TEXT,
                course_subject TEXT,
                course_number TEXT,
                total_issues INTEGER DEFAULT 0,
                missing_sections_count INTEGER DEFAULT 0,
                accessibility_issues_count INTEGER DEFAULT 0,
                category_counts JSON,
                missing_sections_summary TEXT,
                growth_mindset_status TEXT,
                language_tone_summary TEXT,
                quality_analysis_status TEXT,
                clarity_quality_summary TEXT,
                accessibility_summary TEXT
            )
        ''')
        conn.commit()
        conn.close()

        # Should not raise even though newer columns are absent
        analytics_db.init_db()

        # All newer columns should now exist
        conn = sqlite3.connect(db_path)
        cursor = conn.execute("PRAGMA table_info(scans)")
        cols = {row[1] for row in cursor.fetchall()}
        conn.close()

        for expected in ('course_title', 'semester', 'growth_mindset_score',
                         'quality_issues_count', 'metadata_extraction_status',
                         'language_tone_subcategories', 'clarity_quality_subcategories'):
            assert expected in cols, f"Column '{expected}' not added by migration"

    def test_init_db_twice_does_not_raise(self, db):
        """Calling init_db() on an already-initialised DB must be idempotent."""
        analytics_db.init_db()  # second call
        assert analytics_db.get_scan_count() == 0  # no data corrupted


# ---------------------------------------------------------------------------
# get_trends_by_month
# ---------------------------------------------------------------------------

class TestGetTrendsByMonth:
    def test_returns_one_row_per_month(self, db):
        id1 = _store('jan_a.docx', total_issues=2, missing_sections=['A'])
        id2 = _store('jan_b.docx', total_issues=4, missing_sections=['B', 'C'])
        id3 = _store('feb.docx', total_issues=6, missing_sections=['D', 'E', 'F'])
        _set_dates(db, {
            id1: '2025-01-10 09:00:00',
            id2: '2025-01-20 09:00:00',
            id3: '2025-02-05 09:00:00',
        })

        trends = analytics_db.get_trends_by_month(months=6)

        months_found = [t['month'] for t in trends]
        assert '2025-01' in months_found
        assert '2025-02' in months_found
        # Each month appears exactly once
        assert len([m for m in months_found if m == '2025-01']) == 1
        assert len([m for m in months_found if m == '2025-02']) == 1

    def test_returns_correct_scan_counts_per_month(self, db):
        ids = [_store(f'jan{i}.docx') for i in range(3)]
        id_feb = _store('feb.docx')
        for i, scan_id in enumerate(ids):
            _set_dates(db, {scan_id: f'2025-01-{i+1:02d} 09:00:00'})
        _set_dates(db, {id_feb: '2025-02-01 09:00:00'})

        trends = analytics_db.get_trends_by_month(months=6)
        by_month = {t['month']: t for t in trends}

        assert by_month['2025-01']['scan_count'] == 3
        assert by_month['2025-02']['scan_count'] == 1

    def test_most_recent_month_is_first(self, db):
        id1 = _store()
        id2 = _store()
        _set_dates(db, {id1: '2025-01-01 09:00:00', id2: '2025-03-01 09:00:00'})

        trends = analytics_db.get_trends_by_month(months=12)
        assert trends[0]['month'] == '2025-03'


# ---------------------------------------------------------------------------
# get_department_summary
# ---------------------------------------------------------------------------

class TestGetDepartmentSummary:
    def test_aggregates_correctly_across_multiple_scans(self, db):
        for _ in range(3):
            _store(department='English', total_issues=4)
        for _ in range(2):
            _store(department='Math', total_issues=2)

        summary = analytics_db.get_department_summary()
        by_dept = {row['department']: row for row in summary}

        assert by_dept['English']['scan_count'] == 3
        assert by_dept['Math']['scan_count'] == 2
        assert abs(by_dept['English']['avg_issues'] - 4.0) < 0.01
        assert abs(by_dept['Math']['avg_issues'] - 2.0) < 0.01

    def test_excludes_rows_with_null_department(self, db):
        _store(department=None)
        _store(department='')
        _store(department='Biology')

        summary = analytics_db.get_department_summary()
        depts = [r['department'] for r in summary]
        assert None not in depts
        assert '' not in depts
        assert 'Biology' in depts

    def test_ordered_by_scan_count_descending(self, db):
        for _ in range(5):
            _store(department='English')
        for _ in range(2):
            _store(department='Math')

        summary = analytics_db.get_department_summary()
        assert summary[0]['department'] == 'English'
