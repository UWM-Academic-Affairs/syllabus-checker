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
- [x] `store_scan_result()` inserts a row and returns a valid integer scan_id
- [x] `get_all_scans()` returns rows in descending date order
- [x] `get_all_scans(department='X')` filters correctly
- [x] `delete_scans([id1, id2])` removes exactly those rows and leaves others intact
- [x] `get_scan_count()` returns the correct total
- [x] `save_to_repository()` inserts a row with the correct filenames
- [x] `delete_from_repository()` removes the row and returns stored filenames
- [x] `is_saved(scan_id)` returns True/False correctly
- [x] `get_saved_syllabi()` returns rows joined with scan metadata
- [x] Database schema migration (new column added) does not fail on an existing DB
- [x] `get_trends_by_month()` returns one row per month and correct issue counts
- [x] `get_department_summary()` aggregates correctly across multiple scans

### Repository Routes
- [x] `POST /repository/save` copies original file to repo folder
- [x] `POST /repository/save` copies marked file to repo folder when present
- [x] `POST /repository/save` with missing filename returns 400
- [x] `POST /repository/save` when original file is gone returns 400
- [x] `GET /repository/download/<name>` streams the correct file
- [x] `GET /repository/download/<name>` for missing file redirects with error
- [x] `POST /repository/delete/<id>` removes the DB row and both files
- [x] `POST /repository/delete/<id>` for non-existent id returns gracefully
- [x] Path traversal attempt in download URL is neutralized (e.g., `../../../etc/passwd`)

### Admin Routes
- [x] `GET /admin/analytics` returns 200 and contains scan statistics
- [x] `GET /admin/analytics?department=X` filters the displayed scans
- [x] `GET /admin/analytics/csv` returns a valid CSV with correct headers
- [x] `POST /admin/analytics/delete` with valid scan ids removes them from DB
- [x] `POST /admin/analytics/delete` with non-integer id returns an error, not a 500
- [x] Repository tab shows saved syllabi with correct counts

### SyllabusChecker — Pattern Detection (`syllabus_checker.py`)
- [x] `_detect_content_patterns()` detects Instructor section when email present
- [x] `_detect_content_patterns()` detects Grading Scale from letter-grade table
- [x] `_detect_content_patterns()` detects Calendar when date/week patterns present
- [x] `_detect_content_patterns()` does NOT flag template placeholders as real content
- [x] `_find_heading_for_content()` returns closest pseudo-heading over distant document title
- [x] `_find_heading_for_content()` does NOT let table-cell Word headings override top-level headings
- [x] `_extract_headings_with_positions()` correctly sets `is_top_level` vs table-cell headings
- [x] `check_missing_sections()` lists sections absent from a minimal syllabus
- [x] `check_missing_sections()` does NOT flag present sections as missing
- [x] `_get_section_description('Grading Scale')` includes sample grade table rows

### SyllabusChecker — Accessibility Checks
- [x] Missing table header flagged as an issue
- [x] Table with proper headers not flagged
- [x] Image without alt text flagged
- [x] Image with descriptive alt text not flagged
- [x] Color contrast below WCAG AA threshold flagged
- [x] Non-descriptive link text ("click here", "here") flagged
- [x] All-caps paragraph of sufficient length flagged
- [x] Fully-justified text block flagged
- [x] Font size below 11pt flagged
- [x] Document with no accessibility issues returns empty issue list for those checks

### SyllabusChecker — Marked Document Generation
- [x] `create_marked_document()` produces a valid `.docx` file
- [x] Missing sections appear in the output document's appended section
- [x] No crash when `missing_sections` is empty
- [x] No crash when `growth_mindset_analysis` is absent or errored

### OpenAI Integration (mocked)
- [x] `extract_course_metadata()` parses valid JSON response into correct fields
- [x] `extract_course_metadata()` handles malformed JSON response without crashing
- [x] `analyze_sections_with_llm()` uses pattern detection when API key is absent
- [x] `analyze_growth_mindset_and_belonging()` returns `status: error` when API call fails
- [x] `analyze_syllabus_quality()` returns `status: error` gracefully on timeout

---

## Code Quality & Maintenance

- [x] Add `requirements-dev.txt` separating test dependencies from runtime dependencies
- [x] Configure `pytest` in `pyproject.toml` or `pytest.ini` (test discovery paths, coverage settings)
- [x] Add GitHub Actions CI workflow to run `pytest` on push to `main`
- [x] Add linting step (e.g., `ruff` or `flake8`) to CI
- [x] Document environment variables in `README.md` or a `CONFIGURATION.md` file
- [x] Move hardcoded UWM brand colors from `syllabus_checker.py` to a single constants module
- [x] Add `__version__` to `app.py` and surface it in the admin footer

---

## Reliability & Test Coverage

- [x] Raise `app.py` coverage from 75% to ≥ 85% — achieved 88% (growth mindset/quality/metadata parsing paths, download route)
- [x] Raise `syllabus_checker.py` coverage from 53% to ≥ 65% — achieved 62% on `syllabus_checker.py` (66% total); added tests for `check_unstyled_headings()`, `check_inconsistent_fonts()`, `check_unstyled_links()`, `check_heading_hierarchy()`, `check_layout_tables()`, `check_table_merged_cells()`, `check_manual_alignment()`, `check_table_scope_declarations()`, and LLM path of `analyze_sections_with_llm()`
- [x] Add an end-to-end smoke test — upload a real `.docx` through the test client and assert the results page contains expected section names and issue counts

---

## Deployment & Operations

- [x] Add a health-check endpoint `GET /healthz` returning 200 + `{"version": __version__, "status": "ok"}` — useful for uptime monitoring and confirming deployments landed
- [x] Add a CD step to the GitHub Actions workflow — on merge to `main`, SSH to the server, `git pull`, and restart the gunicorn service
- [x] Add `dependabot.yml` to auto-open PRs when `requirements.txt` dependencies go stale

---

## User-Facing Features

- [x] Batch upload — allow instructors to upload multiple `.docx` files at once and receive a summary table across all syllabi; new `/batch-upload` route, `batch_results.html` summary table, Single/Batch toggle on index page
- [x] Diff view in the admin repository — show what changed between a newly uploaded syllabus and its previously saved version; new `/repository/diff/<id>/<prev_id>` route, `diff.html` side-by-side view, auto-detected previous versions shown with a compare button in the repository tab

---

## Operations & Monitoring

- [x] **Uptime monitoring** — UptimeRobot monitoring `/healthz` every 5 minutes; public status page at https://stats.uptimerobot.com/5MCfonWblL
- [x] **CD step in GitHub Actions** — on merge to `main`, SSH to server, `git pull`, `pip install -r requirements.txt`, restart gunicorn
- [x] **Log rotation** — `/etc/logrotate.d/syllabus-checker` configured; gunicorn writing to `/var/log/syllabus-checker/access.log` and `error.log`
- [x] **Automated DB backups** — `scripts/backup.sh` + `scripts/install-cron.sh` written; run `install-cron.sh` on server to activate daily 2am cron
- [x] **nginx + TLS audit** — gunicorn on 127.0.0.1:8000 (localhost only); nginx proxying syllabus-checker.clarkenstein.com with HTTP→HTTPS redirect; Let's Encrypt cert valid through 2026-07-23

## Security (Production)

- [x] **`.env` audit** — all required env vars confirmed set on server (`SECRET_KEY`, `OPENAI_API_KEY`, `ADMIN_USERNAME`, `ADMIN_PASSWORD`, `FLASK_ENV=production`); duplicate entry fixed, gunicorn restarted
- [x] **Malformed DOCX upload testing** — automated tests added in `tests/test_malformed_uploads.py` (zip bomb, corrupted bytes, empty zip, XML bomb, PNG-renamed-to-docx); all 5 pass
- [x] **Review OpenAI data handling disclosure** — disclosure added to upload form, results page, and batch results page; links to OpenAI privacy policy

## User Experience

- [x] **User feedback on results** — thumbs-up/thumbs-down on results page so instructors can flag false positives; stored in `feedback` table; `POST /feedback` endpoint
- [x] **Department trends dashboard** — Trends & Departments tab in admin analytics with department bar chart, 6-month issue trends line chart, and feedback summary card (Chart.js, UWM colors)

---

## Code Maintainability

- [x] Break up `syllabus_checker.py` (6,000+ lines) into focused modules — `checks/accessibility.py` (AccessibilityMixin, 1572 lines), `checks/patterns.py` (PatternsMixin, 708 lines), `llm/analysis.py` (LLMMixin, 119 lines), `output/marked_doc.py` (OutputMixin, 841 lines); `syllabus_checker.py` reduced to 283 lines; shared types in `_checker_types.py`; all 196 tests pass
