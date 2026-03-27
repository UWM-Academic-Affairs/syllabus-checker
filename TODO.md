# Syllabus Checker — TODO: Tests & Security

## Security

### High Priority
- [x] **Add authentication to admin routes** — `@require_admin` HTTP Basic Auth decorator applied to all `/admin/*` routes; credentials set via `ADMIN_USERNAME`/`ADMIN_PASSWORD` env vars (no-op if unset, for local dev).
- [x] **Add authentication to repository routes** — `@require_admin` applied to `/repository/delete` and `/repository/download`; `/repository/save` stays public (any uploader may save their own syllabus).
- [x] **Add CSRF protection** — Flask-WTF `CSRFProtect` enabled globally; `csrf_token()` added to upload and admin-delete forms; AJAX-only routes (`/repository/save`, `/repository/delete`) exempted with `@csrf.exempt`.
- [x] **Verify `.env` is in `.gitignore`** — confirmed present; `.env` and `.env.local` both excluded.
- [x] **Harden `SECRET_KEY`** — startup warning logged if default key detected outside development mode.

### Medium Priority
- [x] **Sanitize error messages shown to users** — exceptions logged server-side via `logging`; generic messages shown to users.
- [x] **Add rate limiting to upload route** — Flask-Limiter added (10 uploads/hour per IP); `Flask-Limiter==3.5.0` in requirements.txt.
- [x] **Validate `repository/download` path parameter** — `os.path.realpath` check added to ensure resolved path stays within `REPO_FOLDER`.
- [x] **Validate department filter input** — capped at 100 characters and stripped before use.
- [x] **Set `SESSION_COOKIE_SECURE` and `SESSION_COOKIE_HTTPONLY`** — both set; `SECURE` disabled only when `FLASK_ENV=development`.
- [x] **Add `Content-Security-Policy` header** — set via `after_request` hook along with `X-Content-Type-Options` and `X-Frame-Options`.

### Low Priority
- [x] **Periodic temp-file cleanup** — `_sweep_old_uploads()` runs at startup and removes files older than 24 hours from the upload staging folder.
- [x] **Limit uploaded filename length** — `safe_filename()` wrapper caps filenames at 100 characters (preserving extension) before saving.
- [x] **Audit `scan_ids` deletion endpoint** — input capped at 50 IDs per request; each must be a positive integer; non-integer values are silently skipped.

---

## Automated Tests

### Setup
- [x] Add `pytest` and `pytest-flask` to `requirements.txt` (or a separate `requirements-dev.txt`)
- [x] Create `tests/` directory with `conftest.py` that sets up a Flask test client, a temp upload folder, and a temp SQLite DB
- [x] Add sample DOCX fixtures to `tests/fixtures/` (minimal valid, malformed, missing sections, template itself)
- [x] Add a mock for OpenAI API calls (e.g., `unittest.mock.patch` or `pytest-mock`) so tests run without API keys

### File Upload & Processing Pipeline
- [x] Valid `.docx` upload returns 200 and renders results page
- [x] Non-`.docx` upload (`.pdf`, `.txt`, `.exe`) is rejected with a user-friendly error
- [x] Upload larger than 16 MB is rejected
- [x] Empty file upload is rejected
- [x] Upload with no file part in the request is rejected
- [x] Marked document is created at the expected path after upload
- [x] Scan record is written to the database after a successful upload
- [x] Processing error (corrupted DOCX) is caught and flashes an error without a 500

### Analytics Database (`analytics_db.py`)
- [ ] `store_scan_result()` inserts a row and returns a valid integer scan_id
- [ ] `get_all_scans()` returns rows in descending date order
- [ ] `get_all_scans(department='X')` filters correctly
- [ ] `delete_scans([id1, id2])` removes exactly those rows and leaves others intact
- [ ] `get_scan_count()` returns the correct total
- [ ] `save_to_repository()` inserts a row with the correct filenames
- [ ] `delete_from_repository()` removes the row and returns stored filenames
- [ ] `is_saved(scan_id)` returns True/False correctly
- [ ] `get_saved_syllabi()` returns rows joined with scan metadata
- [ ] Database schema migration (new column added) does not fail on an existing DB
- [ ] `get_trends_by_month()` returns one row per month and correct issue counts
- [ ] `get_department_summary()` aggregates correctly across multiple scans

### Repository Routes
- [ ] `POST /repository/save` copies original file to repo folder
- [ ] `POST /repository/save` copies marked file to repo folder when present
- [ ] `POST /repository/save` with missing filename returns 400
- [ ] `POST /repository/save` when original file is gone returns 400
- [ ] `GET /repository/download/<name>` streams the correct file
- [ ] `GET /repository/download/<name>` for missing file redirects with error
- [ ] `POST /repository/delete/<id>` removes the DB row and both files
- [ ] `POST /repository/delete/<id>` for non-existent id returns gracefully
- [ ] Path traversal attempt in download URL is neutralized (e.g., `../../../etc/passwd`)

### Admin Routes
- [ ] `GET /admin/analytics` returns 200 and contains scan statistics
- [ ] `GET /admin/analytics?department=X` filters the displayed scans
- [ ] `GET /admin/analytics/csv` returns a valid CSV with correct headers
- [ ] `POST /admin/analytics/delete` with valid scan ids removes them from DB
- [ ] `POST /admin/analytics/delete` with non-integer id returns an error, not a 500
- [ ] Repository tab shows saved syllabi with correct counts

### SyllabusChecker — Pattern Detection (`syllabus_checker.py`)
- [ ] `_detect_content_patterns()` detects Instructor section when email present
- [ ] `_detect_content_patterns()` detects Grading Scale from letter-grade table
- [ ] `_detect_content_patterns()` detects Calendar when date/week patterns present
- [ ] `_detect_content_patterns()` does NOT flag template placeholders as real content
- [ ] `_find_heading_for_content()` returns closest pseudo-heading over distant document title
- [ ] `_find_heading_for_content()` does NOT let table-cell Word headings override top-level headings
- [ ] `_extract_headings_with_positions()` correctly sets `is_top_level` vs table-cell headings
- [ ] `check_missing_sections()` lists sections absent from a minimal syllabus
- [ ] `check_missing_sections()` does NOT flag present sections as missing
- [ ] `_get_section_description('Grading Scale')` includes sample grade table rows

### SyllabusChecker — Accessibility Checks
- [ ] Missing table header flagged as an issue
- [ ] Table with proper headers not flagged
- [ ] Image without alt text flagged
- [ ] Image with descriptive alt text not flagged
- [ ] Color contrast below WCAG AA threshold flagged
- [ ] Non-descriptive link text ("click here", "here") flagged
- [ ] All-caps paragraph of sufficient length flagged
- [ ] Fully-justified text block flagged
- [ ] Font size below 11pt flagged
- [ ] Document with no accessibility issues returns empty issue list for those checks

### SyllabusChecker — Marked Document Generation
- [ ] `create_marked_document()` produces a valid `.docx` file
- [ ] Missing sections appear in the output document's appended section
- [ ] No crash when `missing_sections` is empty
- [ ] No crash when `growth_mindset_analysis` is absent or errored

### OpenAI Integration (mocked)
- [ ] `extract_course_metadata()` parses valid JSON response into correct fields
- [ ] `extract_course_metadata()` handles malformed JSON response without crashing
- [ ] `analyze_sections_with_llm()` uses pattern detection when API key is absent
- [ ] `analyze_growth_mindset_and_belonging()` returns `status: error` when API call fails
- [ ] `analyze_syllabus_quality()` returns `status: error` gracefully on timeout

---

## Code Quality & Maintenance

- [ ] Add `requirements-dev.txt` separating test dependencies from runtime dependencies
- [ ] Configure `pytest` in `pyproject.toml` or `pytest.ini` (test discovery paths, coverage settings)
- [ ] Add GitHub Actions CI workflow to run `pytest` on push to `main`
- [ ] Add linting step (e.g., `ruff` or `flake8`) to CI
- [ ] Document environment variables in `README.md` or a `CONFIGURATION.md` file
- [ ] Move hardcoded UWM brand colors from `syllabus_checker.py` to a single constants module
- [ ] Add `__version__` to `app.py` and surface it in the admin footer
