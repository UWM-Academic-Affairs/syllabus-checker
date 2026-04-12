# Configuration Reference

All configuration is done through environment variables.
For local development, create a `.env` file in the project root (it is gitignored).
For production, set these in your systemd unit file or hosting platform's environment dashboard.

---

## Required in Production

### `SECRET_KEY`
Flask uses this to cryptographically sign session cookies and CSRF tokens.

**Generate a strong key:**
```bash
python -c "import secrets; print(secrets.token_hex(32))"
```

If this variable is unset or left at the default value while `FLASK_ENV` is not `development`,
the application logs a warning at startup.

---

## Optional — OpenAI Features

### `OPENAI_API_KEY`
Enables the three LLM-powered analysis features:

| Feature | Route | Behavior without key |
|---|---|---|
| Course metadata extraction | `POST /` (upload) | Skipped; metadata fields left blank |
| Belonging & Growth Mindset analysis | `POST /` (upload) | Returns `status: error`, skipped in marked doc |
| Syllabus quality analysis | `POST /` (upload) | Returns `status: error`, skipped in marked doc |

All 40+ rule-based accessibility checks run regardless of whether this key is set.

---

## File Storage

### `UPLOAD_FOLDER`
- **Default:** `/tmp/syllabus_checker_uploads`
- Temporary staging area for uploaded `.docx` files.
- Files older than 24 hours are swept at startup.
- Must be writable by the web process user.

### `REPO_FOLDER`
- **Default:** `<project_root>/repository`
- Permanent storage for syllabi saved via the repository feature.
- Falls back to `/tmp/syllabus_checker_repository` if the default path is not writable.
- Must be writable by the web process user.

---

## Admin Dashboard

The `/admin/analytics` dashboard and the `/repository/delete` and `/repository/download`
routes are protected by HTTP Basic Auth when both of the following are set.
If either is absent, the routes are unprotected (suitable for local development).

### `ADMIN_USERNAME`
Username for HTTP Basic Auth on admin and restricted repository routes.

### `ADMIN_PASSWORD`
Password for HTTP Basic Auth on admin and restricted repository routes.

---

## Flask / Development

### `FLASK_ENV`
- Set to `development` to:
  - Disable `SESSION_COOKIE_SECURE` (allows HTTP in local dev)
  - Suppress the `SECRET_KEY` default-value warning
  - Enable Flask's built-in debugger and reloader (`python app.py`)
- Any other value (including unset) is treated as production mode.

---

## Example `.env` (local development)

```dotenv
FLASK_ENV=development
SECRET_KEY=replace-me-with-a-real-secret
OPENAI_API_KEY=sk-...
UPLOAD_FOLDER=/tmp/syllabus_checker_uploads
REPO_FOLDER=/tmp/syllabus_checker_repository
# ADMIN_USERNAME and ADMIN_PASSWORD intentionally omitted for local dev
```

## Example `.env` (production)

```dotenv
SECRET_KEY=<output of secrets.token_hex(32)>
OPENAI_API_KEY=sk-...
UPLOAD_FOLDER=/var/www/syllabus-checker/uploads
REPO_FOLDER=/var/www/syllabus-checker/repository
ADMIN_USERNAME=admin
ADMIN_PASSWORD=<strong-random-password>
```
