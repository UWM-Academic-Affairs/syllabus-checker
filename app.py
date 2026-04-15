"""
Syllabus Accessibility Checker Web Application
Flask app for uploading and checking Word documents for accessibility issues
"""

__version__ = '1.1.0'

from flask import Flask, render_template, request, redirect, url_for, send_file, flash, Response
import os
import re
import logging
import secrets
import traceback
from functools import wraps
from werkzeug.utils import secure_filename
from syllabus_checker import SyllabusChecker
import shutil
from datetime import datetime
from dotenv import load_dotenv
import markdown
import analytics_db
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address
from flask_wtf.csrf import CSRFProtect

# Load environment variables from .env file
load_dotenv()

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)

_DEFAULT_SECRET = 'syllabus-checker-secret-key-change-in-production'
_secret_key = os.environ.get('SECRET_KEY', _DEFAULT_SECRET)
if _secret_key == _DEFAULT_SECRET and not os.environ.get('FLASK_ENV') == 'development':
    logger.warning(
        'SECRET_KEY is using the insecure default value. '
        'Set a strong random SECRET_KEY environment variable in production.'
    )
app.secret_key = _secret_key
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# Session cookie security
app.config['SESSION_COOKIE_HTTPONLY'] = True
app.config['SESSION_COOKIE_SECURE'] = os.environ.get('FLASK_ENV') != 'development'
app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'
app.config['WTF_CSRF_TIME_LIMIT'] = 3600  # 1 hour

csrf = CSRFProtect(app)

# ── Admin authentication ──────────────────────────────────────────────────────
# Set ADMIN_USERNAME and ADMIN_PASSWORD in .env to enable Basic Auth protection.
# If neither is set, admin routes are unprotected (useful for local dev).

def require_admin(f):
    """HTTP Basic Auth guard for admin and restricted repository routes."""
    @wraps(f)
    def decorated(*args, **kwargs):
        admin_user = os.environ.get('ADMIN_USERNAME', '')
        admin_pass = os.environ.get('ADMIN_PASSWORD', '')
        if not admin_user or not admin_pass:
            # No credentials configured — allow through (dev mode)
            return f(*args, **kwargs)
        auth = request.authorization
        if auth and \
           secrets.compare_digest(auth.username, admin_user) and \
           secrets.compare_digest(auth.password, admin_pass):
            return f(*args, **kwargs)
        return Response(
            'Admin access required.',
            401,
            {'WWW-Authenticate': 'Basic realm="Syllabus Checker Admin"'}
        )
    return decorated

# Rate limiter — prevents abuse of the upload route (each request hits the OpenAI API)
limiter = Limiter(
    get_remote_address,
    app=app,
    default_limits=[],
    storage_uri='memory://'
)

# Security headers
@app.after_request
def set_security_headers(response):
    response.headers['Content-Security-Policy'] = (
        "default-src 'self'; "
        "script-src 'self' cdn.jsdelivr.net 'unsafe-inline'; "
        "style-src 'self' cdn.jsdelivr.net 'unsafe-inline'; "
        "font-src 'self' cdn.jsdelivr.net; "
        "img-src 'self' data:; "
        "frame-ancestors 'none';"
    )
    response.headers['X-Content-Type-Options'] = 'nosniff'
    response.headers['X-Frame-Options'] = 'DENY'
    return response
app.config['UPLOAD_FOLDER'] = os.environ.get('UPLOAD_FOLDER', '/tmp/syllabus_checker_uploads')
app.config['REPO_FOLDER'] = os.environ.get('REPO_FOLDER',
    os.path.join(os.path.dirname(os.path.abspath(__file__)), 'repository'))
# Fall back to a /tmp path if the default location isn't writable (e.g. restricted service user)
try:
    os.makedirs(app.config['REPO_FOLDER'], exist_ok=True)
except PermissionError:
    app.config['REPO_FOLDER'] = os.path.join('/tmp', 'syllabus_checker_repository')
    os.makedirs(app.config['REPO_FOLDER'], exist_ok=True)

# Ensure upload folder exists
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

ALLOWED_EXTENSIONS = {'docx'}

def allowed_file(filename):
    """Check if file has allowed extension"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def safe_filename(filename, max_length=100):
    """Sanitize and cap filename length while preserving the extension."""
    name = secure_filename(filename)
    if len(name) > max_length:
        base, ext = os.path.splitext(name)
        name = base[:max_length - len(ext)] + ext
    return name

def _sweep_old_uploads(folder, max_age_hours=24):
    """Remove files older than max_age_hours from the upload staging folder."""
    if not os.path.isdir(folder):
        return
    cutoff = datetime.now().timestamp() - max_age_hours * 3600
    for fname in os.listdir(folder):
        fpath = os.path.join(folder, fname)
        try:
            if os.path.isfile(fpath) and os.path.getmtime(fpath) < cutoff:
                os.remove(fpath)
                logger.info('Swept stale upload: %s', fname)
        except OSError as e:
            logger.warning('Could not remove stale upload %s: %s', fname, e)

# Sweep leftover uploads from any previous unclean shutdown
_sweep_old_uploads(app.config['UPLOAD_FOLDER'])

# ── Module-level constants ─────────────────────────────────────────────────────

CATEGORY_MAPPING = {
    'Font Usage': ['SMALL_FONT', 'DECORATIVE_FONT', 'INCONSISTENT_FONTS'],
    'Table Headers': ['TABLE_NO_HEADER'],
    'Table Scope (Repeat Header Rows)': ['TABLE_MISSING_SCOPE'],
    'Table Captions': ['TABLE_MISSING_CAPTION'],
    'Layout Tables': ['LAYOUT_TABLE'],
    'Table Empty Rows/Columns': ['EMPTY_TABLE_ROW', 'EMPTY_TABLE_COLUMN'],
    'Table Merged Cells': ['TABLE_MERGED_CELLS'],
    'Table Reading Order': ['TABLE_COMPLEX_READING_ORDER'],
    'Table Color Coding': ['COLOR_CODED_TABLE', 'TABLE_COLOR_CODED_MEANING'],
    'Table Embedded Images': ['TABLE_EMBEDDED_IMAGES'],
    'Table Numeric Alignment': ['TABLE_INCONSISTENT_NUMERIC_ALIGNMENT'],
    'Color & Contrast': ['LOW_CONTRAST', 'COLOR_ONLY_MEANING', 'TEXT_OVER_BACKGROUND'],
    'Links & Navigation': ['NON_DESCRIPTIVE_LINK', 'UNSTYLED_LINK', 'LONG_URL',
                           'MISSING_TOC', 'MISSING_BOOKMARKS'],
    'Lists': ['INCONSISTENT_LIST_HIERARCHY', 'LAYOUT_LIST'],
    'Text Formatting': ['PSEUDO_TABLE', 'UNDERLINE_NON_LINK', 'INSUFFICIENT_LINE_SPACING',
                       'FULL_JUSTIFICATION', 'ALL_CAPS_BLOCK', 'EXCESSIVE_BOLD',
                       'EXCESSIVE_ITALIC', 'EXCESSIVE_UNDERLINE', 'INCONSISTENT_FORMATTING'],
    'Readability': ['LONG_SENTENCE', 'NUMERIC_DATE_FORMAT'],
    'Images': ['IMAGE_MISSING_ALT', 'DECORATIVE_IMAGE_QUESTIONABLE', 'IMAGE_TEXT_CONTENT'],
    'Document Properties': ['MISSING_TITLE', 'MISSING_LANGUAGE', 'MULTIPLE_LANGUAGES'],
    'Heading Structure': ['ALL_CAPS_HEADING', 'LONG_HEADING', 'H1_NOT_AT_TOP', 'MULTIPLE_H1',
                         'SHOULD BE HEADING 2', 'SHOULD BE HEADING 3', 'SHOULD BE HEADING 4'],
    'Content Quality': ['BROKEN_STYLE_COPIED_CONTENT', 'FOOTNOTE_USAGE', 'VISUAL_INDICATOR_NO_TEXT', 'MATH_NO_ACCESSIBLE_MARKUP'],
}

CATEGORY_HELP = {
    'Font Usage': {
        'summary': 'Some text may be hard to read due to small or decorative fonts.',
        'why_matters': 'Students with visual impairments or reading difficulties need adequate font sizes (11pt+) and clear, standard fonts to read your syllabus.',
        'how_to_fix': 'Select the text → Home tab → Change font size to 11pt or larger. Use Arial, Calibri, or Times New Roman.',
        'help_link': 'https://support.microsoft.com/en-us/office/change-the-font-size-dc5b3f47-9e78-4af6-9734-ffc69175811c',
        'help_text': 'How to change fonts'
    },
    'Table Headers': {
        'summary': 'Your table is missing a designated header row — the top row that labels each column (e.g., "Week", "Topic", "Reading").',
        'why_matters': 'Screen readers announce table contents cell-by-cell. Without a header row, students using assistive technology hear raw data with no context for what each column means.',
        'what_is_it': '<strong>Table header</strong> = the top row that labels each column. When marked correctly, a screen reader says "Week: 1, Topic: Introduction" instead of just "1, Introduction."',
        'how_to_fix': 'Click anywhere in your table &rarr; go to the <strong>Table Design</strong> tab &rarr; check the <strong>"Header Row"</strong> box. This tells Word (and screen readers) that the first row contains column labels.',
        'help_link': 'https://support.microsoft.com/en-us/office/video-create-accessible-tables-in-word-cb464015-59dc-46a0-ac01-6217c62210e5',
        'help_text': 'Video: Create accessible tables'
    },
    'Table Scope (Repeat Header Rows)': {
        'summary': 'Your table\'s header row isn\'t marked as a repeating header. This means screen readers can\'t associate column labels with cell data.',
        'why_matters': 'When a table spans multiple pages, sighted readers see headers repeated at the top of each page. Screen reader users need the "Repeat Header Rows" setting so their software knows which row contains the column labels.',
        'what_is_it': '<strong>Repeat Header Rows</strong> = a Word setting that marks the first row as the official header and repeats it on every printed page. This is how Word communicates table structure to assistive technology.',
        'how_to_fix': '<strong>Windows:</strong> Click in the header row &rarr; Table Tools &gt; Layout tab &rarr; click <strong>"Repeat Header Rows"</strong>.<br><strong>Mac:</strong> Click in the header row &rarr; Table menu &rarr; <strong>"Repeat Header Row"</strong>.',
        'help_link': 'https://support.microsoft.com/en-us/office/video-create-accessible-tables-in-word-cb464015-59dc-46a0-ac01-6217c62210e5',
        'help_text': 'Video: Create accessible tables'
    },
    'Table Captions': {
        'summary': 'Your table doesn\'t have a caption or description, so screen reader users won\'t know what the table is about before they start reading it.',
        'why_matters': 'Screen readers announce the caption before reading table contents. Without one, users must listen to several rows of data before they can figure out what the table shows — imagine listening to "Week 1, Monday, Room 101..." with no idea what the table is for.',
        'what_is_it': '<strong>Table caption</strong> = a short title and description stored in the table\'s properties (Alt Text). Example: Title "Assignment Schedule", Description "Lists all assignments with due dates and point values for the semester."',
        'how_to_fix': 'Right-click anywhere in the table &rarr; select <strong>"Table Properties"</strong> &rarr; go to the <strong>"Alt Text"</strong> tab &rarr; fill in the <strong>Title</strong> (brief name) and <strong>Description</strong> (what the table shows).',
        'help_link': 'https://support.microsoft.com/en-us/office/add-alternative-text-to-a-shape-picture-chart-smartart-graphic-or-other-object-44989b2a-903c-4d9a-b742-6a75b451c669',
        'help_text': 'How to add alt text'
    },
    'Layout Tables': {
        'summary': 'This table is being used for page layout rather than to present data. Screen readers will try to read it as a data table, which confuses students.',
        'why_matters': 'When a screen reader encounters a table, it announces "Table with X rows and Y columns" and reads cell-by-cell. If the table is just arranging text on the page, this forces students to listen to disjointed content in a confusing order.',
        'how_to_fix': 'Replace the layout table with <strong>headings and paragraphs</strong>. Select the content inside the table, cut it, delete the table, then paste the content back and apply proper heading styles (Heading 2, Heading 3) to organize it.',
        'help_link': 'https://support.microsoft.com/en-us/office/make-your-word-documents-accessible-d9bf3683-87ac-47ea-b91a-78dcacb3c66d',
        'help_text': 'Accessibility best practices'
    },
    'Table Empty Rows/Columns': {
        'summary': 'Your table has empty rows or columns that are being used for visual spacing. Screen readers read these as blank cells, which is confusing.',
        'why_matters': 'A screen reader user navigating through a table will hear "blank" for every empty cell, making them think the table has missing data. Empty rows also break the logical flow of the table.',
        'how_to_fix': 'Delete the empty rows/columns. For spacing, right-click the table &rarr; <strong>"Table Properties"</strong> &rarr; <strong>"Cell"</strong> tab &rarr; adjust cell <strong>margins/padding</strong> instead of using empty rows.',
        'help_link': 'https://support.microsoft.com/en-us/office/video-create-accessible-tables-in-word-cb464015-59dc-46a0-ac01-6217c62210e5',
        'help_text': 'Video: Create accessible tables'
    },
    'Table Merged Cells': {
        'summary': 'Your table has merged cells, which can make it difficult for screen readers to navigate correctly.',
        'why_matters': 'Screen readers navigate tables by moving between cells using arrow keys. Merged cells break the expected grid structure, causing the reader to skip cells or announce the wrong column header for a cell.',
        'how_to_fix': 'Simplify the table structure — split merged cells back into individual cells, or reorganize the table so merging isn\'t needed. If you need a section header spanning the table, use a heading above the table instead of a merged row.',
        'help_link': 'https://support.microsoft.com/en-us/office/video-create-accessible-tables-in-word-cb464015-59dc-46a0-ac01-6217c62210e5',
        'help_text': 'Video: Create accessible tables'
    },
    'Table Reading Order': {
        'summary': 'Your table has a complex structure (merged cells combined with multiple header-like rows) that may create a confusing reading order.',
        'why_matters': 'Screen readers read tables left-to-right, top-to-bottom. Complex structures with nested headers and merged cells can cause the reader to announce cells in an unexpected order, making the data incomprehensible.',
        'how_to_fix': 'Simplify the table — use a single header row, avoid merged cells, and consider splitting complex tables into multiple simpler tables. Each table should have one clear, simple structure.',
        'help_link': 'https://support.microsoft.com/en-us/office/video-create-accessible-tables-in-word-cb464015-59dc-46a0-ac01-6217c62210e5',
        'help_text': 'Video: Create accessible tables'
    },
    'Table Color Coding': {
        'summary': 'Your table uses multiple background colors which may be conveying meaning that isn\'t available to colorblind or screen reader users.',
        'why_matters': 'About 8% of men have some form of color blindness, and screen readers don\'t announce cell background colors at all. If color indicates something (like "green = complete, red = overdue"), those students will miss that information entirely.',
        'how_to_fix': 'Add a <strong>text label</strong> or <strong>icon description</strong> alongside any color coding. For example, instead of just coloring a cell red, also add the text "Overdue" or a column with status labels.',
        'help_link': 'https://webaim.org/resources/contrastchecker/',
        'help_text': 'Check your colors'
    },
    'Table Embedded Images': {
        'summary': 'Your table contains images inside cells with little or no accompanying text.',
        'why_matters': 'Screen readers may not properly read alt text for images embedded in table cells, and the image-to-cell association can be lost. Students using assistive technology may miss important visual information.',
        'how_to_fix': 'Replace images with text where possible. If images are necessary (logos, icons), ensure each has descriptive alt text: right-click the image &rarr; <strong>"Edit Alt Text"</strong> &rarr; describe what the image conveys.',
        'help_link': 'https://support.microsoft.com/en-us/office/add-alternative-text-to-a-shape-picture-chart-smartart-graphic-or-other-object-44989b2a-903c-4d9a-b742-6a75b451c669',
        'help_text': 'How to add alt text'
    },
    'Table Numeric Alignment': {
        'summary': 'A column of numbers in your table has inconsistent alignment, making it harder to scan and compare values.',
        'why_matters': 'Inconsistent alignment of numeric data (some left-aligned, some centered, some right-aligned) makes it harder for all students — especially those with cognitive or visual disabilities — to quickly compare values in a column.',
        'how_to_fix': 'Select the numeric column &rarr; right-click &rarr; <strong>"Cell Alignment"</strong> &rarr; choose a consistent alignment (right-aligned is standard for numbers). Apply the same alignment to all data cells in the column.',
        'help_link': 'https://support.microsoft.com/en-us/office/video-create-accessible-tables-in-word-cb464015-59dc-46a0-ac01-6217c62210e5',
        'help_text': 'Video: Create accessible tables'
    },
    'Color & Contrast': {
        'summary': 'Some colors may be hard to see or distinguish.',
        'why_matters': 'About 8% of men have some form of color blindness. Low contrast text is also difficult for students reading on phones or in bright environments.',
        'how_to_fix': 'Use dark text (black or dark blue) on light backgrounds. If you use color to convey meaning (like red for "late"), also add a text label.',
        'help_link': 'https://webaim.org/resources/contrastchecker/',
        'help_text': 'Check your colors'
    },
    'Links & Navigation': {
        'summary': 'Some links use vague text like "click here" that doesn\'t describe the destination.',
        'why_matters': 'Screen reader users often navigate by jumping between links. If all links say "click here," they can\'t tell where each one goes.',
        'how_to_fix': 'Right-click the link → Edit Hyperlink → Change "Text to display" to something descriptive like "view the grading rubric" instead of "click here".',
        'help_link': 'https://support.microsoft.com/en-us/office/create-accessible-links-in-word-28305cc8-3be2-417f-a789-67c5a8d14393',
        'help_text': 'How to fix link text'
    },
    'Text Formatting': {
        'summary': 'Some formatting choices can make text harder to read.',
        'why_matters': 'ALL CAPS text is harder to read (we recognize words by their shape). Justified text creates uneven spacing. Underlined text looks like links.',
        'how_to_fix': 'For emphasis, use <strong>bold</strong> instead of ALL CAPS or underline. Use left alignment instead of justified. Use Word\'s built-in Heading styles.',
        'help_link': 'https://support.microsoft.com/en-us/office/use-headings-and-styles-in-word-c0d409b5-5c0a-4a5c-8c5e-001e76072e51',
        'help_text': 'Using Word styles'
    },
    'Images': {
        'summary': 'Some images are missing descriptions (alt text).',
        'why_matters': 'Students who are blind or have low vision use screen readers that read alt text aloud. Without it, they miss the image content entirely.',
        'what_is_it': '<strong>Alt text</strong> = a brief description of what\'s in the image. It should convey the same information a sighted person would get.',
        'how_to_fix': 'Right-click the image → "View Alt Text" (or "Edit Alt Text") → Type a brief description like "UWM logo" or "Chart showing enrollment trends 2020-2024".',
        'help_link': 'https://support.microsoft.com/en-us/office/add-alternative-text-to-a-shape-picture-chart-smartart-graphic-or-other-object-44989b2a-903c-4d9a-b742-6a75b451c669',
        'help_text': 'Video: Add alt text'
    },
    'Document Properties': {
        'summary': 'The document is missing a title or language setting.',
        'why_matters': 'Screen readers announce the document title when opening a file. The language setting ensures text-to-speech uses correct pronunciation.',
        'how_to_fix': 'File → Info → Properties (on the right) → Add a Title. For language: Review tab → Language → Set Proofing Language.',
        'help_link': 'https://support.microsoft.com/en-us/office/make-your-word-documents-accessible-d9bf3683-87ac-47ea-b91a-78dcacb3c66d',
        'help_text': 'Full accessibility guide'
    },
    'Heading Structure': {
        'summary': 'The document structure could be improved with proper headings.',
        'why_matters': 'Screen reader users navigate long documents by jumping between headings. Proper heading levels (Heading 1, 2, 3) create a logical outline.',
        'what_is_it': '<strong>Headings</strong> = structural markers that create an outline. Use Heading 1 for the title, Heading 2 for main sections, Heading 3 for subsections.',
        'how_to_fix': 'Select your heading text → Home tab → Styles group → Click "Heading 1", "Heading 2", etc. Don\'t just make text big and bold—use actual heading styles.',
        'help_link': 'https://support.microsoft.com/en-us/office/use-headings-and-styles-in-word-c0d409b5-5c0a-4a5c-8c5e-001e76072e51',
        'help_text': 'Using heading styles'
    },
    'Content Quality': {
        'summary': 'Some content may have formatting issues or accessibility concerns.',
        'why_matters': 'Pasted content often brings hidden formatting problems. Footnotes can be hard to navigate. Visual symbols need text alternatives.',
        'how_to_fix': 'When pasting, use Ctrl+Shift+V (or Paste Special → Keep Text Only) to avoid formatting issues. Replace footnotes with inline explanations when possible.',
        'help_link': 'https://support.microsoft.com/en-us/office/make-your-word-documents-accessible-d9bf3683-87ac-47ea-b91a-78dcacb3c66d',
        'help_text': 'Accessibility best practices'
    },
    'Lists': {
        'summary': 'List formatting could be improved.',
        'why_matters': 'Properly formatted lists help screen readers announce "list of 5 items" and navigate item-by-item.',
        'how_to_fix': 'Use Word\'s built-in bullet or numbered list buttons instead of manually typing dashes or numbers.',
        'help_link': 'https://support.microsoft.com/en-us/office/create-a-bulleted-or-numbered-list-9ff81241-58a8-4d88-8d8c-acab3006a23e',
        'help_text': 'Creating proper lists'
    },
    'Readability': {
        'summary': 'Some content may be difficult to read or understand.',
        'why_matters': 'Long sentences and ambiguous dates can confuse all students, especially non-native English speakers and those with cognitive disabilities.',
        'how_to_fix': 'Break long sentences into shorter ones. Use clear date formats like "January 15, 2024" instead of "1/15/24".',
        'help_link': 'https://support.microsoft.com/en-us/office/get-your-document-s-readability-and-level-statistics-85b4969e-e80a-4777-8dd3-f7fc3c8b3fd2',
        'help_text': 'Check readability'
    },
}

# ── Routes ─────────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    """Main upload page"""
    return render_template('index.html')

def _process_docx(filename: str, upload_path: str):
    """Process a single .docx file; return a results dict (with scan_id) or None on error."""
    try:
        _template_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                      'assets', 'Uniform-Syllabus-Template-1.docx')
        checker = SyllabusChecker(_template_path, upload_path)

        # Get section analysis using LLM (content-based analysis)
        section_analysis = checker.analyze_sections_with_llm()

        # Extract missing content types for backward compatibility with analytics
        # New format: missing is [{content_type: "X", description: "...", recommendation: "..."}]
        missing_items = section_analysis.get('missing', [])
        if missing_items and isinstance(missing_items[0], dict):
            missing_sections = [item.get('content_type', str(item)) for item in missing_items]
        else:
            missing_sections = missing_items  # Fallback for old format

        # Run all checks
        checker.run_all_checks()

        # Get additional checks that return dictionaries (not added to self.issues)
        heading_check = checker.check_heading_structure()
        table_check = checker.check_table_usage()
        list_check = checker.check_list_usage()

        # Run growth mindset and belonging analysis
        growth_mindset_analysis = checker.analyze_growth_mindset_and_belonging()

        # Run syllabus quality analysis
        quality_analysis = checker.analyze_syllabus_quality()

        # Extract course metadata using LLM
        course_metadata = checker.extract_course_metadata()

        # Run image alt text analysis
        image_alt_analysis = checker.analyze_image_alt_text()

        # Generate report (generate_report already returns a joined string)
        report_text = checker.generate_report()

        # Create marked document with missing sections and growth mindset recommendations
        marked_filename = f"{os.path.splitext(filename)[0]}_marked.docx"
        marked_path = os.path.join(os.path.dirname(upload_path), marked_filename)
        checker.create_marked_document(marked_path, missing_sections, growth_mindset_analysis, quality_analysis)

        # Count issues by category and collect structured issue data
        category_counts = {}
        category_issues = {}

        for issue in checker.issues:
            # Find which category this issue belongs to
            issue_category = None
            for category, issue_types in CATEGORY_MAPPING.items():
                if issue.issue_type in issue_types:
                    issue_category = category
                    break

            if issue_category:
                if issue_category not in category_counts:
                    category_counts[issue_category] = 0
                    category_issues[issue_category] = []
                category_counts[issue_category] += 1
                category_issues[issue_category].append({
                    'location': issue.location,
                    'description': issue.description
                })

        # Extract the SUMMARY section from the report
        summary_section = ""
        report_lines = report_text.split('\n')
        in_summary = False
        summary_lines = []

        for i, line in enumerate(report_lines):
            # Look for the SUMMARY header (it's between two lines of equals signs)
            if line.strip() == "SUMMARY":
                # Skip the next line (which should be another line of equals signs)
                in_summary = True
                continue
            elif in_summary:
                # Skip the line of equals signs right after SUMMARY
                if line.strip().startswith("=" * 40) and len(summary_lines) == 0:
                    continue
                # Stop at the next major section (another line of equals signs after we've started collecting)
                if line.strip().startswith("=" * 40) and len(summary_lines) > 0:
                    break
                summary_lines.append(line)

        summary_section = '\n'.join(summary_lines).strip()

        # Calculate total issues (must match text report calculation)
        total_web_issues = (
            len(missing_sections) +
            len(heading_check.get('issues', [])) +
            len(table_check.get('issues', [])) +
            len(list_check.get('issues', [])) +
            len(checker.issues)
        )

        # Prepare summary
        summary_message = ""
        if total_web_issues == 0:
            summary_message = "Excellent! This syllabus meets all standards."
            summary_status = "success"
        elif total_web_issues < 5:
            summary_message = f"This syllabus has {total_web_issues} issue(s) that should be addressed."
            summary_status = "warning"
        else:
            summary_message = f"This syllabus has {total_web_issues} issue(s) that need attention."
            summary_status = "danger"

        # Add structural issues to appropriate categories
        if table_check.get('issues'):
            if 'Layout Tables' not in category_counts:
                category_counts['Layout Tables'] = 0
                category_issues['Layout Tables'] = []
            category_counts['Layout Tables'] += len(table_check.get('issues', []))
            for issue_text in table_check['issues']:
                category_issues['Layout Tables'].append({
                    'location': 'Document',
                    'description': issue_text.replace('WARNING: ', '')
                })

        if heading_check.get('issues'):
            if 'Document Structure' not in category_counts:
                category_counts['Document Structure'] = 0
                category_issues['Document Structure'] = []
            category_counts['Document Structure'] += len(heading_check.get('issues', []))
            for issue_text in heading_check['issues']:
                category_issues['Document Structure'].append({
                    'location': 'Document',
                    'description': issue_text.replace('WARNING: ', '')
                })

        if list_check.get('issues'):
            if 'Lists' not in category_counts:
                category_counts['Lists'] = 0
            category_counts['Lists'] += len(list_check.get('issues', []))

        # Convert growth mindset analysis markdown to HTML
        if growth_mindset_analysis.get('status') == 'success' and 'analysis' in growth_mindset_analysis:
            growth_mindset_analysis['analysis_html'] = markdown.markdown(
                growth_mindset_analysis['analysis'],
                extensions=['extra', 'nl2br', 'sane_lists']
            )

        # Convert quality analysis markdown to HTML
        if quality_analysis.get('status') == 'success' and 'analysis' in quality_analysis:
            quality_analysis['analysis_html'] = markdown.markdown(
                quality_analysis['analysis'],
                extensions=['extra', 'nl2br', 'sane_lists']
            )

        # Convert image alt text analysis markdown to HTML
        if image_alt_analysis.get('status') == 'success' and 'analysis' in image_alt_analysis:
            image_alt_analysis['analysis_html'] = markdown.markdown(
                image_alt_analysis['analysis'],
                extensions=['extra', 'nl2br', 'sane_lists']
            )
            analysis_lower = image_alt_analysis['analysis'].lower()
            image_alt_analysis['has_issues'] = not (
                'all institutional logos have meaningful alt text' in analysis_lower
                and 'all alt text is clear, concise, and descriptive' in analysis_lower
            )

        results = {
            'filename': filename,
            'total_issues': total_web_issues,
            'missing_sections': missing_sections,
            'section_analysis': section_analysis,
            'category_counts': category_counts,
            'category_issues': category_issues,
            'category_help': CATEGORY_HELP,
            'report_text': report_text,
            'summary_section': summary_section,
            'marked_file': marked_filename,
            'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'summary_message': summary_message,
            'summary_status': summary_status,
            'growth_mindset_analysis': growth_mindset_analysis,
            'quality_analysis': quality_analysis,
            'image_alt_analysis': image_alt_analysis
        }

        # Store scan results in analytics database
        try:
            instructor_name = None
            department = None
            course_subject = None
            course_number = None
            course_title = None
            semester = None
            metadata_extraction_status = course_metadata.get('status', 'error')

            if course_metadata.get('status') == 'success':
                instructor_name = course_metadata.get('instructor_name')
                department = course_metadata.get('department')
                course_subject = course_metadata.get('course_subject')
                course_number = course_metadata.get('course_number')
                course_title = course_metadata.get('course_title')
                semester = course_metadata.get('semester')

            # ===== LANGUAGE & TONE (Growth Mindset) - 6 subcategories =====
            language_tone_summary = None
            growth_mindset_score = None
            language_tone_subcategories = {}

            if growth_mindset_analysis.get('status') == 'success':
                analysis_text = growth_mindset_analysis.get('analysis', '')

                subcategory_patterns = [
                    (r'(?:QUESTION\s*1|Growth\s*Mindset)[^\n]*\n\**Assessment:?\**\s*(Strong|Moderate|Weak|Not\s*Addressed)', 'Growth Mindset'),
                    (r'(?:QUESTION\s*2|Normalizing\s*Challenge)[^\n]*\n\**Assessment:?\**\s*(Strong|Moderate|Weak|Not\s*Addressed)', 'Normalizing Challenge'),
                    (r'(?:QUESTION\s*3|Instructor\s*Care)[^\n]*\n\**Assessment:?\**\s*(Strong|Moderate|Weak|Not\s*Addressed)', 'Instructor Care'),
                    (r'(?:QUESTION\s*4|Valuing\s*Diversity)[^\n]*\n\**Assessment:?\**\s*(Strong|Moderate|Weak|Not\s*Addressed)', 'Valuing Diversity'),
                    (r'(?:QUESTION\s*5|Normalizing\s*Student\s*Challenges)[^\n]*\n\**Assessment:?\**\s*(Strong|Moderate|Weak|Not\s*Addressed)', 'Normalizing Student Challenges'),
                    (r'(?:QUESTION\s*6|Normalizing\s*Academic\s*Support)[^\n]*\n\**Assessment:?\**\s*(Strong|Moderate|Weak|Not\s*Addressed)', 'Normalizing Academic Support'),
                ]

                for pattern, category in subcategory_patterns:
                    match = re.search(pattern, analysis_text, re.IGNORECASE)
                    if match:
                        assessment = match.group(1).strip().title()
                        if 'not' in assessment.lower():
                            assessment = 'Not Addressed'
                        language_tone_subcategories[category] = assessment

                strong_count = sum(1 for v in language_tone_subcategories.values() if v == 'Strong')
                weak_count = sum(1 for v in language_tone_subcategories.values() if v in ['Weak', 'Not Addressed'])

                if strong_count >= 4:
                    language_tone_summary = f'Strong ({strong_count}/6 areas)'
                elif weak_count >= 3:
                    language_tone_summary = f'Needs work ({weak_count}/6 areas weak)'
                else:
                    language_tone_summary = 'Mixed results'

                score_match = re.search(r'(?:score|rating|overall)[:\s]*(\d+)(?:/10|\s*out of\s*10)?', analysis_text, re.IGNORECASE)
                if score_match:
                    growth_mindset_score = int(score_match.group(1))

            elif growth_mindset_analysis.get('error'):
                language_tone_summary = 'Analysis unavailable'

            # ===== CLARITY & QUALITY - 4 subcategories =====
            clarity_quality_summary = None
            quality_issues_count = 0
            clarity_quality_subcategories = {}

            if quality_analysis.get('status') == 'success':
                analysis_text = quality_analysis.get('analysis', '')

                sections = [
                    (r'1\.?\s*UNDEFINED\s*(?:COURSE\s*)?TERMINOLOGY(.*?)(?=\n\s*(?:##\s*)?2\.|$)', 'Undefined Terminology', r'\*?\*?Term\*?\*?\s*:\s*[""\'\']'),
                    (r'2\.?\s*TONE\s*(?:AND\s*INCLUSIVITY|ISSUES)(.*?)(?=\n\s*(?:##\s*)?3\.|$)', 'Tone Issues', r'\*?\*?Issue\*?\*?\s*:\s*[""\'\']'),
                    (r'3\.?\s*(?:POLICIES\s*THAT\s*MAY\s*)?CONFUS(?:ING|E)(.*?)(?=\n\s*(?:##\s*)?4\.|$)', 'Confusing Policies', r'\*?\*?Policy\*?\*?\s*:\s*[""\'\']'),
                    (r'4\.?\s*(?:INCONSISTENT\s*)?FORMAT(?:TING)?(?:\s*INCONSISTENC(?:IES|Y))?(.*?)(?=\n\s*(?:##\s*)?5\.|$)', 'Formatting Inconsistencies', r'\*?\*?Element\s*type\*?\*?\s*:\s*'),
                ]

                for section_pattern, category, issue_pattern in sections:
                    match = re.search(section_pattern, analysis_text, re.IGNORECASE | re.DOTALL)
                    if match:
                        section_text = match.group(1)
                        issues = len(re.findall(issue_pattern, section_text, re.IGNORECASE))
                        clarity_quality_subcategories[category] = issues
                    else:
                        clarity_quality_subcategories[category] = 0

                quality_issues_count = sum(clarity_quality_subcategories.values())
                categories_with_issues = sum(1 for v in clarity_quality_subcategories.values() if v > 0)

                if quality_issues_count == 0:
                    clarity_quality_summary = 'Excellent clarity'
                elif categories_with_issues == 1:
                    problem_cat = [k for k, v in clarity_quality_subcategories.items() if v > 0][0]
                    clarity_quality_summary = f'{problem_cat}: {quality_issues_count} issue(s)'
                else:
                    clarity_quality_summary = f'{quality_issues_count} issues in {categories_with_issues} areas'

            elif quality_analysis.get('error'):
                clarity_quality_summary = 'Analysis unavailable'

            # Generate accessibility summary from top issues
            accessibility_summary = None
            if category_counts:
                top_issues = sorted(category_counts.items(), key=lambda x: x[1], reverse=True)[:3]
                if top_issues:
                    accessibility_summary = '; '.join([f"{cat}: {count}" for cat, count in top_issues])

            scan_id = analytics_db.store_scan_result(
                filename=filename,
                instructor_name=instructor_name,
                department=department,
                course_subject=course_subject,
                course_number=course_number,
                course_title=course_title,
                semester=semester,
                total_issues=total_web_issues,
                missing_sections=missing_sections,
                category_counts=category_counts,
                growth_mindset_status=growth_mindset_analysis.get('status'),
                growth_mindset_score=growth_mindset_score,
                quality_analysis_status=quality_analysis.get('status'),
                quality_issues_count=quality_issues_count,
                language_tone_summary=language_tone_summary,
                clarity_quality_summary=clarity_quality_summary,
                accessibility_summary=accessibility_summary,
                metadata_extraction_status=metadata_extraction_status,
                language_tone_subcategories=language_tone_subcategories,
                clarity_quality_subcategories=clarity_quality_subcategories
            )
            results['scan_id'] = scan_id
        except Exception as db_error:
            # Don't fail the request if analytics storage fails
            logger.error('Analytics storage error: %s', db_error)

        return results

    except Exception as e:
        logger.error('Error in _process_docx for %s: %s\n%s', filename, e, traceback.format_exc())
        return None


@app.route('/upload', methods=['POST'])
@limiter.limit('10 per hour')
def upload_file():
    """Handle file upload and run accessibility checks"""
    if 'file' not in request.files:
        flash('No file uploaded', 'error')
        return redirect(url_for('index'))
    file = request.files['file']
    if file.filename == '':
        flash('No file selected', 'error')
        return redirect(url_for('index'))
    if not allowed_file(file.filename):
        flash('Invalid file type. Please upload a .docx file', 'error')
        return redirect(url_for('index'))
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    filename = safe_filename(file.filename)
    upload_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(upload_path)
    results = _process_docx(filename, upload_path)
    if results is None:
        flash('An error occurred while processing your file. Please try again.', 'error')
        return redirect(url_for('index'))
    return render_template('results.html', results=results)


@app.route('/batch-upload', methods=['POST'])
@limiter.limit('5 per hour')
def batch_upload():
    """Handle multiple .docx uploads and return a summary table."""
    files = request.files.getlist('files[]')
    if not files or all(f.filename == '' for f in files):
        flash('No files selected', 'error')
        return redirect(url_for('index'))

    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    batch_results = []

    for file in files[:10]:   # cap at 10 files per batch
        if not file.filename or not allowed_file(file.filename):
            batch_results.append({'filename': file.filename or '(unnamed)', 'error': 'Invalid file type — must be .docx'})
            continue
        if file.content_length and file.content_length > app.config['MAX_CONTENT_LENGTH']:
            batch_results.append({'filename': file.filename, 'error': 'File exceeds 16 MB limit'})
            continue
        filename = safe_filename(file.filename)
        upload_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        try:
            file.save(upload_path)
        except Exception as save_err:
            logger.error('Batch upload save error: %s', save_err)
            batch_results.append({'filename': filename, 'error': 'Could not save file'})
            continue
        result = _process_docx(filename, upload_path)
        if result is None:
            batch_results.append({'filename': filename, 'error': 'Processing failed'})
        else:
            batch_results.append(result)

    return render_template('batch_results.html', batch_results=batch_results)


@app.route('/download/<filename>')
def download_file(filename):
    """Download marked document"""
    try:
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(filename))
        if os.path.exists(file_path):
            return send_file(
                file_path,
                as_attachment=True,
                download_name=filename,
                mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            )
        else:
            flash('File not found', 'error')
            return redirect(url_for('index'))
    except Exception as e:
        logger.error('Error downloading file %s: %s', filename, e)
        flash('An error occurred while downloading the file.', 'error')
        return redirect(url_for('index'))

@app.route('/repository/save', methods=['POST'])
@csrf.exempt   # called via fetch from results page; rate-limited at /upload
def repository_save():
    """Copy the original and marked syllabi into the persistent repository."""
    scan_id = request.form.get('scan_id', type=int)
    filename = request.form.get('filename', '')
    marked_filename = request.form.get('marked_filename', '')

    if not filename:
        return ('No file specified.', 400)

    safe_name = secure_filename(filename)
    src = os.path.join(app.config['UPLOAD_FOLDER'], safe_name)

    if not os.path.exists(src):
        return ('Original file is no longer available.', 400)

    # Store with a timestamp prefix so duplicates are kept distinct
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    stored_name = f"{timestamp}_{safe_name}"
    shutil.copy2(src, os.path.join(app.config['REPO_FOLDER'], stored_name))

    # Also copy the marked version if it exists
    stored_marked_name = None
    if marked_filename:
        safe_marked = secure_filename(marked_filename)
        marked_src = os.path.join(app.config['UPLOAD_FOLDER'], safe_marked)
        if os.path.exists(marked_src):
            stored_marked_name = f"{timestamp}_{safe_marked}"
            shutil.copy2(marked_src, os.path.join(app.config['REPO_FOLDER'], stored_marked_name))

    analytics_db.save_to_repository(scan_id, filename, stored_name, stored_marked_name)
    return ('', 204)


@app.route('/repository/download/<path:stored_filename>')
@require_admin
def repository_download(stored_filename):
    """Download a syllabus from the repository."""
    safe_name = secure_filename(stored_filename)
    repo_folder = os.path.realpath(app.config['REPO_FOLDER'])
    file_path = os.path.realpath(os.path.join(repo_folder, safe_name))
    # Ensure the resolved path stays inside the repo folder (path traversal guard)
    if not file_path.startswith(repo_folder + os.sep):
        logger.warning('Path traversal attempt blocked: %s', stored_filename)
        flash('Invalid file path.', 'error')
        return redirect(url_for('admin_analytics'))
    if not os.path.exists(file_path):
        flash('File not found in repository.', 'error')
        return redirect(url_for('admin_analytics'))
    # Strip the timestamp prefix for the download name
    display_name = safe_name[16:] if len(safe_name) > 16 and safe_name[8] == '_' else safe_name
    return send_file(file_path, as_attachment=True, download_name=display_name,
                     mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')


@app.route('/repository/delete/<int:repo_id>', methods=['POST'])
@require_admin
@csrf.exempt   # called via fetch from admin page; admin auth provides equivalent protection
def repository_delete(repo_id):
    """Delete a repository entry and its stored files."""
    row = analytics_db.delete_from_repository(repo_id)
    if row:
        for fname in [row.get('stored_filename'), row.get('stored_marked_filename')]:
            if fname:
                path = os.path.join(app.config['REPO_FOLDER'], secure_filename(fname))
                if os.path.exists(path):
                    os.remove(path)
    return ('', 204)


def _extract_docx_text(path: str) -> list:
    """Extract non-empty paragraph texts from a .docx file for diffing."""
    from docx import Document as DocxDocument
    try:
        doc = DocxDocument(path)
        lines = []
        for para in doc.paragraphs:
            text = para.text.strip()
            if text:
                lines.append(text)
        return lines
    except Exception:
        return []


@app.route('/repository/diff/<int:repo_id>/<int:prev_id>')
@require_admin
def repository_diff(repo_id, prev_id):
    """Show a text diff between two saved syllabus versions."""
    import difflib
    current = analytics_db.get_repo_entry(repo_id)
    previous = analytics_db.get_repo_entry(prev_id)
    if not current or not previous:
        flash('One or both repository entries not found.', 'error')
        return redirect(url_for('admin_analytics'))

    repo_folder = app.config['REPO_FOLDER']
    current_path = os.path.realpath(os.path.join(repo_folder, secure_filename(current['stored_filename'])))
    prev_path = os.path.realpath(os.path.join(repo_folder, secure_filename(previous['stored_filename'])))

    # Path traversal guard
    real_repo = os.path.realpath(repo_folder)
    if not current_path.startswith(real_repo + os.sep) or not prev_path.startswith(real_repo + os.sep):
        flash('Invalid file path.', 'error')
        return redirect(url_for('admin_analytics'))

    current_lines = _extract_docx_text(current_path) if os.path.exists(current_path) else []
    prev_lines = _extract_docx_text(prev_path) if os.path.exists(prev_path) else []

    diff_html = difflib.HtmlDiff(wrapcolumn=80).make_table(
        prev_lines,
        current_lines,
        fromdesc=f"Previous ({previous['saved_at'][:10] if previous['saved_at'] else 'unknown'})",
        todesc=f"Current ({current['saved_at'][:10] if current['saved_at'] else 'unknown'})",
        context=True,
        numlines=3,
    )

    return render_template('diff.html',
        current=current,
        previous=previous,
        diff_html=diff_html,
    )


@app.route('/about')
def about():
    """About page"""
    return render_template('about.html')

# ============================================================
# ADMIN ANALYTICS ROUTES
# ============================================================

@app.route('/admin/analytics')
@require_admin
def admin_analytics():
    """Admin analytics dashboard"""
    # Get filter parameters — cap length and strip to prevent oversized inputs
    department_filter = request.args.get('department', '')[:100].strip()
    dept_param = department_filter if department_filter else None

    # Get overview statistics
    stats = analytics_db.get_overview_stats(dept_param)

    # Get four areas summary
    four_areas = analytics_db.get_four_areas_summary(dept_param)

    # Get common issues (accessibility)
    common_issues = analytics_db.get_common_issues(limit=10, department=dept_param)

    # Get trends by month
    trends = analytics_db.get_trends_by_month(months=6, department=dept_param)

    # Get growth mindset trends
    growth_mindset_trends = analytics_db.get_growth_mindset_trends(months=6, department=dept_param)

    # Get quality analysis trends
    quality_trends = analytics_db.get_quality_analysis_trends(months=6, department=dept_param)

    # Get language & tone breakdown
    language_tone_breakdown = analytics_db.get_language_tone_breakdown(dept_param)

    # Get clarity & quality breakdown
    clarity_quality_breakdown = analytics_db.get_clarity_quality_breakdown(dept_param)

    # Get subcategory statistics for detailed breakdowns
    language_tone_subcategories = analytics_db.get_language_tone_subcategory_stats(dept_param)
    clarity_quality_subcategories = analytics_db.get_clarity_quality_subcategory_stats(dept_param)

    # Get department summary
    department_summary = analytics_db.get_department_summary()

    # Get missing sections frequency
    missing_sections_freq = analytics_db.get_missing_sections_frequency()

    # Get list of departments for filter dropdown
    departments = analytics_db.get_departments_list()

    # Get recent scans (show 20 most recent)
    recent_scans = analytics_db.get_all_scans(limit=20, department=dept_param)

    # Get total scan count for CSV download info
    total_scans = analytics_db.get_scan_count(dept_param)

    # Repository data
    saved_syllabi = analytics_db.get_saved_syllabi()
    saved_scan_ids = analytics_db.get_saved_scan_ids()
    feedback_summary = analytics_db.get_feedback_summary()

    # Group by original_filename to find versions for diff
    from collections import defaultdict
    filename_versions = defaultdict(list)
    for item in saved_syllabi:
        filename_versions[item['original_filename']].append(item)
    # Assign prev_repo_id to each item that has an older version
    for item in saved_syllabi:
        versions = filename_versions[item['original_filename']]
        # versions are newest first (from DB query)
        idx = next((i for i, v in enumerate(versions) if v['id'] == item['id']), None)
        item['prev_repo_id'] = versions[idx + 1]['id'] if idx is not None and idx + 1 < len(versions) else None

    return render_template('admin_analytics.html',
                           stats=stats,
                           four_areas=four_areas,
                           common_issues=common_issues,
                           trends=trends,
                           growth_mindset_trends=growth_mindset_trends,
                           quality_trends=quality_trends,
                           language_tone_breakdown=language_tone_breakdown,
                           clarity_quality_breakdown=clarity_quality_breakdown,
                           language_tone_subcategories=language_tone_subcategories,
                           clarity_quality_subcategories=clarity_quality_subcategories,
                           department_summary=department_summary,
                           missing_sections_freq=missing_sections_freq,
                           departments=departments,
                           department_filter=department_filter,
                           recent_scans=recent_scans,
                           total_scans=total_scans,
                           saved_syllabi=saved_syllabi,
                           saved_scan_ids=saved_scan_ids,
                           feedback_summary=feedback_summary,
                           app_version=__version__)

@app.route('/admin/analytics/delete', methods=['POST'])
@require_admin
def admin_analytics_delete():
    """Delete selected scan entries"""
    try:
        # Get scan IDs from form — cap at 50 per request, require positive integers
        raw_ids = request.form.getlist('scan_ids')[:50]

        if not raw_ids:
            flash('No scans selected for deletion', 'warning')
            return redirect(url_for('admin_analytics'))

        scan_ids = []
        for raw in raw_ids:
            try:
                val = int(raw)
                if val > 0:
                    scan_ids.append(val)
            except (ValueError, TypeError):
                pass  # silently skip non-integer values

        if not scan_ids:
            flash('No valid scan IDs provided.', 'warning')
            return redirect(url_for('admin_analytics'))

        # Delete the scans
        deleted_count = analytics_db.delete_scans(scan_ids)

        flash(f'Successfully deleted {deleted_count} scan(s)', 'success')

    except Exception as e:
        logger.error('Error deleting scans: %s', e)
        flash('An error occurred while deleting the selected scans.', 'error')

    # Preserve department filter if present
    department = request.form.get('department_filter', '')
    if department:
        return redirect(url_for('admin_analytics', department=department))
    return redirect(url_for('admin_analytics'))

@app.route('/admin/analytics/csv')
@require_admin
def admin_analytics_csv():
    """Download all scans as CSV"""
    import csv
    import io

    # Get optional department filter
    dept_param = request.args.get('department', None)

    # Get all scans (no limit)
    all_scans = analytics_db.get_all_scans(limit=10000, department=dept_param)

    # Create CSV in memory
    output = io.StringIO()
    writer = csv.writer(output)

    # Write header
    writer.writerow([
        'Date',
        'Course Subject',
        'Course Number',
        'Course Title',
        'Instructor',
        'Department',
        'Semester',
        'Missing Sections',
        'Language & Tone Score',
        'Quality Issues',
        'Accessibility Issues',
        'Total Issues',
        'Filename'
    ])

    # Write data rows
    for scan in all_scans:
        writer.writerow([
            scan.get('scan_date', '')[:10] if scan.get('scan_date') else '',
            scan.get('course_subject', ''),
            scan.get('course_number', ''),
            scan.get('course_title', ''),
            scan.get('instructor_name', ''),
            scan.get('department', ''),
            scan.get('semester', ''),
            scan.get('missing_sections_count', 0),
            scan.get('growth_mindset_score', '') if scan.get('growth_mindset_status') == 'success' else '',
            scan.get('quality_issues_count', '') if scan.get('quality_analysis_status') == 'success' else '',
            scan.get('accessibility_issues_count', 0),
            scan.get('total_issues', 0),
            scan.get('filename', '')
        ])

    # Prepare response
    output.seek(0)
    filename = f"syllabus_scans_{datetime.now().strftime('%Y%m%d')}"
    if dept_param:
        filename += f"_{dept_param}"
    filename += ".csv"

    return Response(
        output.getvalue(),
        mimetype='text/csv',
        headers={'Content-Disposition': f'attachment; filename={filename}'}
    )

@app.route('/healthz')
def healthz():
    from flask import jsonify
    return jsonify({'version': __version__, 'status': 'ok'})


@app.route('/feedback', methods=['POST'])
@csrf.exempt
def submit_feedback():
    """Store thumbs-up/down feedback for a scan result."""
    scan_id = request.form.get('scan_id', type=int)
    rating = request.form.get('rating', '').strip()
    comment = request.form.get('comment', '').strip()[:500]  # cap at 500 chars

    if not scan_id or rating not in ('up', 'down'):
        return ('Invalid feedback.', 400)

    analytics_db.store_feedback(scan_id, rating, comment or None)
    return ('', 204)


if __name__ == '__main__':
    # Clean up temp folder on exit
    import atexit
    atexit.register(lambda: shutil.rmtree(app.config['UPLOAD_FOLDER'], ignore_errors=True))

    # Run the app (using port 5001 to avoid macOS AirPlay conflict)
    app.run(debug=True, host='0.0.0.0', port=5001)
