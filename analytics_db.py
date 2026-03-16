"""
Analytics Database Module for Syllabus Checker
Stores scan results for reporting and trend analysis
"""

import sqlite3
import os
from datetime import datetime
from typing import List, Dict, Optional
import json

DATABASE_PATH = os.environ.get('ANALYTICS_DB_PATH', 'syllabus_analytics.db')

def get_db_connection():
    """Get a database connection"""
    conn = sqlite3.connect(DATABASE_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    """Initialize the database schema"""
    conn = get_db_connection()
    cursor = conn.cursor()

    # Create scans table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS scans (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            scan_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            filename TEXT,
            instructor_name TEXT,
            department TEXT,
            course_subject TEXT,
            course_number TEXT,
            course_title TEXT,
            semester TEXT,

            -- Issue counts by category
            total_issues INTEGER DEFAULT 0,
            missing_sections_count INTEGER DEFAULT 0,
            accessibility_issues_count INTEGER DEFAULT 0,

            -- Detailed counts by issue type (stored as JSON)
            category_counts JSON,

            -- AREA 1: Missing Required Sections
            missing_sections_summary TEXT,

            -- AREA 2: Language & Tone (Growth Mindset) - 6 subcategories
            growth_mindset_status TEXT,
            growth_mindset_score INTEGER,
            language_tone_summary TEXT,
            -- Subcategory assessments stored as JSON: {category: "Strong/Moderate/Weak/Not Addressed"}
            language_tone_subcategories JSON,

            -- AREA 3: Clarity & Quality - 4 subcategories
            quality_analysis_status TEXT,
            quality_issues_count INTEGER DEFAULT 0,
            clarity_quality_summary TEXT,
            -- Subcategory issue counts stored as JSON: {category: count}
            clarity_quality_subcategories JSON,

            -- AREA 4: Accessibility
            accessibility_summary TEXT,

            -- Metadata extraction status
            metadata_extraction_status TEXT
        )
    ''')

    # Add new columns if they don't exist (for existing databases)
    cursor.execute("PRAGMA table_info(scans)")
    existing_columns = [col[1] for col in cursor.fetchall()]

    new_columns = [
        ("course_title", "TEXT"),
        ("semester", "TEXT"),
        ("growth_mindset_score", "INTEGER"),
        ("quality_issues_count", "INTEGER DEFAULT 0"),
        ("metadata_extraction_status", "TEXT"),
        ("language_tone_subcategories", "JSON"),
        ("clarity_quality_subcategories", "JSON")
    ]

    for col_name, col_type in new_columns:
        if col_name not in existing_columns:
            try:
                cursor.execute(f"ALTER TABLE scans ADD COLUMN {col_name} {col_type}")
            except:
                pass  # Column might already exist

    # Create index for common queries
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_scan_date ON scans(scan_date)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_department ON scans(department)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_course_subject ON scans(course_subject)')

    # Create repository table for saved syllabi
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS saved_syllabi (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            scan_id INTEGER REFERENCES scans(id) ON DELETE SET NULL,
            original_filename TEXT NOT NULL,
            stored_filename TEXT NOT NULL,
            saved_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_saved_at ON saved_syllabi(saved_at)')

    conn.commit()
    conn.close()

def store_scan_result(
    filename: str,
    instructor_name: str = None,
    department: str = None,
    course_subject: str = None,
    course_number: str = None,
    course_title: str = None,
    semester: str = None,
    total_issues: int = 0,
    missing_sections: List[str] = None,
    category_counts: Dict[str, int] = None,
    growth_mindset_status: str = None,
    growth_mindset_score: int = None,
    quality_analysis_status: str = None,
    quality_issues_count: int = None,
    language_tone_summary: str = None,
    clarity_quality_summary: str = None,
    accessibility_summary: str = None,
    metadata_extraction_status: str = None,
    language_tone_subcategories: Dict[str, str] = None,
    clarity_quality_subcategories: Dict[str, int] = None
) -> int:
    """Store a scan result in the database"""
    conn = get_db_connection()
    cursor = conn.cursor()

    missing_sections = missing_sections or []
    category_counts = category_counts or {}

    # Calculate accessibility issues (total minus missing sections)
    accessibility_issues_count = total_issues - len(missing_sections)

    # Create summaries
    missing_sections_summary = ', '.join(missing_sections) if missing_sections else 'None'

    # Generate accessibility summary from category counts
    if not accessibility_summary and category_counts:
        top_issues = sorted(category_counts.items(), key=lambda x: x[1], reverse=True)[:3]
        if top_issues:
            accessibility_summary = '; '.join([f"{cat}: {count}" for cat, count in top_issues])
        else:
            accessibility_summary = 'No accessibility issues'

    language_tone_subcategories = language_tone_subcategories or {}
    clarity_quality_subcategories = clarity_quality_subcategories or {}

    cursor.execute('''
        INSERT INTO scans (
            filename, instructor_name, department, course_subject, course_number,
            course_title, semester,
            total_issues, missing_sections_count, accessibility_issues_count,
            category_counts, missing_sections_summary, language_tone_summary,
            clarity_quality_summary, accessibility_summary,
            growth_mindset_status, growth_mindset_score,
            quality_analysis_status, quality_issues_count,
            metadata_extraction_status,
            language_tone_subcategories, clarity_quality_subcategories
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (
        filename,
        instructor_name,
        department,
        course_subject,
        course_number,
        course_title,
        semester,
        total_issues,
        len(missing_sections),
        accessibility_issues_count,
        json.dumps(category_counts),
        missing_sections_summary,
        language_tone_summary,
        clarity_quality_summary,
        accessibility_summary,
        growth_mindset_status,
        growth_mindset_score,
        quality_analysis_status,
        quality_issues_count,
        metadata_extraction_status,
        json.dumps(language_tone_subcategories),
        json.dumps(clarity_quality_subcategories)
    ))

    scan_id = cursor.lastrowid
    conn.commit()
    conn.close()

    return scan_id

def delete_scans(scan_ids: List[int]) -> int:
    """Delete multiple scans by their IDs. Returns number of deleted rows."""
    if not scan_ids:
        return 0

    conn = get_db_connection()
    cursor = conn.cursor()

    # Use parameterized query with placeholders for each ID
    placeholders = ','.join('?' * len(scan_ids))
    cursor.execute(f'DELETE FROM scans WHERE id IN ({placeholders})', scan_ids)

    deleted_count = cursor.rowcount
    conn.commit()
    conn.close()

    return deleted_count

def get_all_scans(limit: int = 100, offset: int = 0, department: str = None) -> List[Dict]:
    """Get all scans with optional department filter"""
    conn = get_db_connection()
    cursor = conn.cursor()

    if department:
        cursor.execute('''
            SELECT * FROM scans
            WHERE department = ?
            ORDER BY scan_date DESC
            LIMIT ? OFFSET ?
        ''', (department, limit, offset))
    else:
        cursor.execute('''
            SELECT * FROM scans
            ORDER BY scan_date DESC
            LIMIT ? OFFSET ?
        ''', (limit, offset))

    rows = cursor.fetchall()
    conn.close()

    return [dict(row) for row in rows]

def get_scan_count(department: str = None) -> int:
    """Get total number of scans"""
    conn = get_db_connection()
    cursor = conn.cursor()

    if department:
        cursor.execute('SELECT COUNT(*) FROM scans WHERE department = ?', (department,))
    else:
        cursor.execute('SELECT COUNT(*) FROM scans')

    count = cursor.fetchone()[0]
    conn.close()
    return count

def get_common_issues(limit: int = 10, department: str = None) -> List[Dict]:
    """Get most common issues across all scans"""
    conn = get_db_connection()
    cursor = conn.cursor()

    if department:
        cursor.execute('SELECT category_counts FROM scans WHERE department = ?', (department,))
    else:
        cursor.execute('SELECT category_counts FROM scans')

    rows = cursor.fetchall()
    conn.close()

    # Aggregate issue counts
    issue_totals = {}
    for row in rows:
        if row['category_counts']:
            counts = json.loads(row['category_counts'])
            for category, count in counts.items():
                issue_totals[category] = issue_totals.get(category, 0) + count

    # Sort by count and return top issues
    sorted_issues = sorted(issue_totals.items(), key=lambda x: x[1], reverse=True)
    return [{'category': cat, 'count': count} for cat, count in sorted_issues[:limit]]

def get_trends_by_month(months: int = 6, department: str = None) -> List[Dict]:
    """Get issue trends by month"""
    conn = get_db_connection()
    cursor = conn.cursor()

    if department:
        cursor.execute('''
            SELECT
                strftime('%Y-%m', scan_date) as month,
                COUNT(*) as scan_count,
                AVG(total_issues) as avg_issues,
                AVG(missing_sections_count) as avg_missing_sections,
                AVG(accessibility_issues_count) as avg_accessibility_issues
            FROM scans
            WHERE department = ?
            GROUP BY month
            ORDER BY month DESC
            LIMIT ?
        ''', (department, months))
    else:
        cursor.execute('''
            SELECT
                strftime('%Y-%m', scan_date) as month,
                COUNT(*) as scan_count,
                AVG(total_issues) as avg_issues,
                AVG(missing_sections_count) as avg_missing_sections,
                AVG(accessibility_issues_count) as avg_accessibility_issues
            FROM scans
            GROUP BY month
            ORDER BY month DESC
            LIMIT ?
        ''', (months,))

    rows = cursor.fetchall()
    conn.close()

    return [dict(row) for row in rows]

def get_department_summary() -> List[Dict]:
    """Get summary statistics by department"""
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute('''
        SELECT
            department,
            COUNT(*) as scan_count,
            AVG(total_issues) as avg_issues,
            AVG(missing_sections_count) as avg_missing_sections,
            AVG(accessibility_issues_count) as avg_accessibility_issues
        FROM scans
        WHERE department IS NOT NULL AND department != ''
        GROUP BY department
        ORDER BY scan_count DESC
    ''')

    rows = cursor.fetchall()
    conn.close()

    return [dict(row) for row in rows]

def get_missing_sections_frequency() -> List[Dict]:
    """Get frequency of each missing section"""
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute('SELECT missing_sections_summary FROM scans')
    rows = cursor.fetchall()
    conn.close()

    # Count occurrences of each missing section
    section_counts = {}
    for row in rows:
        if row['missing_sections_summary'] and row['missing_sections_summary'] != 'None':
            sections = [s.strip() for s in row['missing_sections_summary'].split(',')]
            for section in sections:
                if section:
                    section_counts[section] = section_counts.get(section, 0) + 1

    sorted_sections = sorted(section_counts.items(), key=lambda x: x[1], reverse=True)
    return [{'section': section, 'count': count} for section, count in sorted_sections]

def get_departments_list() -> List[str]:
    """Get list of all departments"""
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute('''
        SELECT DISTINCT department FROM scans
        WHERE department IS NOT NULL AND department != ''
        ORDER BY department
    ''')

    rows = cursor.fetchall()
    conn.close()

    return [row['department'] for row in rows]

def get_overview_stats(department: str = None) -> Dict:
    """Get overview statistics for all four assessment areas"""
    conn = get_db_connection()
    cursor = conn.cursor()

    base_query = '''
        SELECT
            COUNT(*) as total_scans,
            AVG(total_issues) as avg_issues,
            MIN(total_issues) as min_issues,
            MAX(total_issues) as max_issues,
            SUM(CASE WHEN total_issues = 0 THEN 1 ELSE 0 END) as perfect_scans,
            AVG(missing_sections_count) as avg_missing_sections,
            AVG(accessibility_issues_count) as avg_accessibility_issues,
            -- Growth Mindset / Language & Tone stats
            SUM(CASE WHEN growth_mindset_status = 'success' THEN 1 ELSE 0 END) as growth_mindset_analyzed,
            AVG(CASE WHEN growth_mindset_score IS NOT NULL THEN growth_mindset_score END) as avg_growth_mindset_score,
            -- Clarity & Quality stats
            SUM(CASE WHEN quality_analysis_status = 'success' THEN 1 ELSE 0 END) as quality_analyzed,
            AVG(CASE WHEN quality_issues_count IS NOT NULL THEN quality_issues_count END) as avg_quality_issues
        FROM scans
    '''

    if department:
        cursor.execute(base_query + ' WHERE department = ?', (department,))
    else:
        cursor.execute(base_query)

    row = cursor.fetchone()
    conn.close()

    return dict(row) if row else {}

def get_growth_mindset_trends(months: int = 6, department: str = None) -> List[Dict]:
    """Get growth mindset score trends by month"""
    conn = get_db_connection()
    cursor = conn.cursor()

    base_query = '''
        SELECT
            strftime('%Y-%m', scan_date) as month,
            COUNT(*) as scan_count,
            SUM(CASE WHEN growth_mindset_status = 'success' THEN 1 ELSE 0 END) as analyzed_count,
            AVG(CASE WHEN growth_mindset_score IS NOT NULL THEN growth_mindset_score END) as avg_score,
            SUM(CASE WHEN language_tone_summary LIKE '%Strong%' THEN 1 ELSE 0 END) as strong_count,
            SUM(CASE WHEN language_tone_summary LIKE '%improvement%' OR language_tone_summary LIKE '%opportunities%' THEN 1 ELSE 0 END) as needs_improvement_count
        FROM scans
    '''

    if department:
        cursor.execute(base_query + '''
            WHERE department = ?
            GROUP BY month
            ORDER BY month DESC
            LIMIT ?
        ''', (department, months))
    else:
        cursor.execute(base_query + '''
            GROUP BY month
            ORDER BY month DESC
            LIMIT ?
        ''', (months,))

    rows = cursor.fetchall()
    conn.close()

    return [dict(row) for row in rows]

def get_quality_analysis_trends(months: int = 6, department: str = None) -> List[Dict]:
    """Get quality analysis trends by month"""
    conn = get_db_connection()
    cursor = conn.cursor()

    base_query = '''
        SELECT
            strftime('%Y-%m', scan_date) as month,
            COUNT(*) as scan_count,
            SUM(CASE WHEN quality_analysis_status = 'success' THEN 1 ELSE 0 END) as analyzed_count,
            AVG(CASE WHEN quality_issues_count IS NOT NULL THEN quality_issues_count END) as avg_issues,
            SUM(CASE WHEN clarity_quality_summary LIKE '%Excellent%' THEN 1 ELSE 0 END) as excellent_count
        FROM scans
    '''

    if department:
        cursor.execute(base_query + '''
            WHERE department = ?
            GROUP BY month
            ORDER BY month DESC
            LIMIT ?
        ''', (department, months))
    else:
        cursor.execute(base_query + '''
            GROUP BY month
            ORDER BY month DESC
            LIMIT ?
        ''', (months,))

    rows = cursor.fetchall()
    conn.close()

    return [dict(row) for row in rows]

def get_four_areas_summary(department: str = None) -> Dict:
    """Get summary statistics for all four assessment areas"""
    conn = get_db_connection()
    cursor = conn.cursor()

    base_query = '''
        SELECT
            -- Area 1: Missing Sections
            AVG(missing_sections_count) as avg_missing_sections,
            SUM(CASE WHEN missing_sections_count = 0 THEN 1 ELSE 0 END) as complete_syllabi,

            -- Area 2: Language & Tone (Growth Mindset)
            SUM(CASE WHEN growth_mindset_status = 'success' THEN 1 ELSE 0 END) as growth_mindset_analyzed,
            AVG(CASE WHEN growth_mindset_score IS NOT NULL THEN growth_mindset_score END) as avg_growth_score,

            -- Area 3: Clarity & Quality
            SUM(CASE WHEN quality_analysis_status = 'success' THEN 1 ELSE 0 END) as quality_analyzed,
            AVG(CASE WHEN quality_issues_count IS NOT NULL THEN quality_issues_count END) as avg_quality_issues,

            -- Area 4: Accessibility
            AVG(accessibility_issues_count) as avg_accessibility_issues,
            SUM(CASE WHEN accessibility_issues_count = 0 THEN 1 ELSE 0 END) as fully_accessible,

            -- Total
            COUNT(*) as total_scans
        FROM scans
    '''

    if department:
        cursor.execute(base_query + ' WHERE department = ?', (department,))
    else:
        cursor.execute(base_query)

    row = cursor.fetchone()
    conn.close()

    return dict(row) if row else {}

def get_language_tone_breakdown(department: str = None) -> List[Dict]:
    """Get breakdown of language & tone analysis results"""
    conn = get_db_connection()
    cursor = conn.cursor()

    base_query = '''
        SELECT
            language_tone_summary as summary,
            COUNT(*) as count
        FROM scans
        WHERE language_tone_summary IS NOT NULL
    '''

    if department:
        cursor.execute(base_query + '''
            AND department = ?
            GROUP BY language_tone_summary
            ORDER BY count DESC
        ''', (department,))
    else:
        cursor.execute(base_query + '''
            GROUP BY language_tone_summary
            ORDER BY count DESC
        ''')

    rows = cursor.fetchall()
    conn.close()

    return [dict(row) for row in rows]

def get_clarity_quality_breakdown(department: str = None) -> List[Dict]:
    """Get breakdown of clarity & quality analysis results"""
    conn = get_db_connection()
    cursor = conn.cursor()

    base_query = '''
        SELECT
            clarity_quality_summary as summary,
            COUNT(*) as count
        FROM scans
        WHERE clarity_quality_summary IS NOT NULL
    '''

    if department:
        cursor.execute(base_query + '''
            AND department = ?
            GROUP BY clarity_quality_summary
            ORDER BY count DESC
        ''', (department,))
    else:
        cursor.execute(base_query + '''
            GROUP BY clarity_quality_summary
            ORDER BY count DESC
        ''')

    rows = cursor.fetchall()
    conn.close()

    return [dict(row) for row in rows]

def get_language_tone_subcategory_stats(department: str = None) -> Dict[str, Dict]:
    """
    Get aggregated statistics for Language & Tone subcategories.
    Returns counts of Strong/Moderate/Weak/Not Addressed for each of the 6 categories:
    - Growth Mindset
    - Normalizing Challenge
    - Instructor Care
    - Valuing Diversity
    - Normalizing Student Challenges
    - Normalizing Academic Support
    """
    conn = get_db_connection()
    cursor = conn.cursor()

    if department:
        cursor.execute('''
            SELECT language_tone_subcategories FROM scans
            WHERE language_tone_subcategories IS NOT NULL AND department = ?
        ''', (department,))
    else:
        cursor.execute('''
            SELECT language_tone_subcategories FROM scans
            WHERE language_tone_subcategories IS NOT NULL
        ''')

    rows = cursor.fetchall()
    conn.close()

    # Define the 6 subcategories
    subcategories = [
        'Growth Mindset',
        'Normalizing Challenge',
        'Instructor Care',
        'Valuing Diversity',
        'Normalizing Student Challenges',
        'Normalizing Academic Support'
    ]

    # Initialize stats
    stats = {cat: {'Strong': 0, 'Moderate': 0, 'Weak': 0, 'Not Addressed': 0, 'total': 0}
             for cat in subcategories}

    # Aggregate data
    for row in rows:
        if row['language_tone_subcategories']:
            data = json.loads(row['language_tone_subcategories'])
            for cat in subcategories:
                if cat in data:
                    assessment = data[cat]
                    if assessment in stats[cat]:
                        stats[cat][assessment] += 1
                        stats[cat]['total'] += 1

    return stats

def get_clarity_quality_subcategory_stats(department: str = None) -> Dict[str, Dict]:
    """
    Get aggregated statistics for Clarity & Quality subcategories.
    Returns issue counts for each of the 4 categories:
    - Undefined Terminology
    - Tone Issues
    - Confusing Policies
    - Formatting Inconsistencies
    """
    conn = get_db_connection()
    cursor = conn.cursor()

    if department:
        cursor.execute('''
            SELECT clarity_quality_subcategories FROM scans
            WHERE clarity_quality_subcategories IS NOT NULL AND department = ?
        ''', (department,))
    else:
        cursor.execute('''
            SELECT clarity_quality_subcategories FROM scans
            WHERE clarity_quality_subcategories IS NOT NULL
        ''')

    rows = cursor.fetchall()
    conn.close()

    # Define the 4 subcategories
    subcategories = [
        'Undefined Terminology',
        'Tone Issues',
        'Confusing Policies',
        'Formatting Inconsistencies'
    ]

    # Initialize stats
    stats = {cat: {'total_issues': 0, 'scans_with_issues': 0, 'scans_clean': 0, 'total_scans': 0}
             for cat in subcategories}

    # Aggregate data
    for row in rows:
        if row['clarity_quality_subcategories']:
            data = json.loads(row['clarity_quality_subcategories'])
            for cat in subcategories:
                if cat in data:
                    count = data[cat]
                    stats[cat]['total_issues'] += count
                    stats[cat]['total_scans'] += 1
                    if count > 0:
                        stats[cat]['scans_with_issues'] += 1
                    else:
                        stats[cat]['scans_clean'] += 1

    return stats

def get_saved_scan_ids() -> set:
    """Return the set of scan IDs that have been saved to the repository."""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('SELECT scan_id FROM saved_syllabi WHERE scan_id IS NOT NULL')
    ids = {row['scan_id'] for row in cursor.fetchall()}
    conn.close()
    return ids


def save_to_repository(scan_id: Optional[int], original_filename: str, stored_filename: str) -> int:
    """Record a syllabus as saved to the repository. Returns the new row id."""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO saved_syllabi (scan_id, original_filename, stored_filename)
        VALUES (?, ?, ?)
    ''', (scan_id, original_filename, stored_filename))
    row_id = cursor.lastrowid
    conn.commit()
    conn.close()
    return row_id


def is_saved(scan_id: int) -> bool:
    """Return True if this scan has already been saved to the repository."""
    if not scan_id:
        return False
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('SELECT 1 FROM saved_syllabi WHERE scan_id = ?', (scan_id,))
    found = cursor.fetchone() is not None
    conn.close()
    return found


def get_saved_syllabi(limit: int = 200) -> List[Dict]:
    """Return saved syllabi joined with scan metadata, newest first."""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('''
        SELECT
            ss.id, ss.scan_id, ss.original_filename, ss.stored_filename, ss.saved_at,
            s.instructor_name, s.department, s.course_subject, s.course_number,
            s.course_title, s.semester, s.total_issues
        FROM saved_syllabi ss
        LEFT JOIN scans s ON ss.scan_id = s.id
        ORDER BY ss.saved_at DESC
        LIMIT ?
    ''', (limit,))
    rows = cursor.fetchall()
    conn.close()
    return [dict(row) for row in rows]


# Initialize database on module load
init_db()
