#!/usr/bin/env python3
"""
Real-API smoke test: one live Claude call per LLM-backed function, using a
generated sample syllabus and a generated sample image. Costs a few cents.

Run from the project root (ANTHROPIC_API_KEY must be in your environment):

    set -a; source .env; set +a; venv/bin/python scripts/smoke_test_llm.py
"""
import os
import struct
import sys
import tempfile
import zlib

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

if not os.environ.get('ANTHROPIC_API_KEY'):
    sys.exit('ANTHROPIC_API_KEY is not set in your environment')

from docx import Document  # noqa: E402
from docx.shared import Inches  # noqa: E402

from syllabus_checker import SyllabusChecker  # noqa: E402

DRAWING_PR = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}docPr'


def _png(path):
    def chunk(tag, data):
        c = struct.pack('>I', len(data)) + tag + data
        return c + struct.pack('>I', zlib.crc32(tag + data) & 0xFFFFFFFF)
    raw = b''.join(b'\x00' + b'\x00\x55\xaa' * 8 for _ in range(8))
    with open(path, 'wb') as f:
        f.write(b'\x89PNG\r\n\x1a\n'
                + chunk(b'IHDR', struct.pack('>IIBBBBB', 8, 8, 8, 2, 0, 0, 0))
                + chunk(b'IDAT', zlib.compress(raw)) + chunk(b'IEND', b''))


def build_sample(tmp):
    """Sparse syllabus (so the section-finder has gaps to ask the model about) + one image."""
    png = os.path.join(tmp, 'logo.png')
    _png(png)
    doc = Document()
    doc.add_heading('CS 101: Introduction to Programming', 0)
    doc.add_paragraph('Instructor: Dr. Jane Smith, Department of Computer Science. Fall 2025.')
    doc.add_paragraph('Welcome to the course! I want every student to succeed, and I am glad you are here.')
    doc.add_paragraph('Students will learn to write and debug small Python programs.')
    doc.add_paragraph('Late work loses 10% per day. Cheating will result in a failing grade.')
    doc.add_picture(png, width=Inches(1))
    for el in doc.element.body.iter(DRAWING_PR):
        el.set('descr', 'logo')
    path = os.path.join(tmp, 'sample_syllabus.docx')
    doc.save(path)
    return path


def main():
    template = os.path.join(ROOT, 'assets', 'Uniform-Syllabus-Template-1.docx')
    failures = []
    with tempfile.TemporaryDirectory() as tmp:
        checker = SyllabusChecker(template, build_sample(tmp))

        def run(name, fn, check):
            try:
                result = fn()
                ok = check(result)
            except Exception as e:  # noqa: BLE001
                result, ok = f'raised {e!r}', False
            print(f"{'OK  ' if ok else 'FAIL'} {name}")
            if not ok:
                print('     ->', str(result)[:300])
                failures.append(name)

        run('extract_course_metadata (fast model, structured output)',
            checker.extract_course_metadata,
            lambda r: r.get('status') == 'success' and 'instructor_name' in r)
        run('analyze_growth_mindset_and_belonging',
            checker.analyze_growth_mindset_and_belonging,
            lambda r: r.get('status') == 'success' and len(r.get('analysis', '')) > 200)
        run('analyze_syllabus_quality',
            checker.analyze_syllabus_quality,
            lambda r: r.get('status') == 'success' and len(r.get('analysis', '')) > 200)
        run('analyze_image_alt_text',
            checker.analyze_image_alt_text,
            lambda r: r.get('status') == 'success' and r.get('image_count') == 1
            and len(r.get('analysis', '')) > 50)
        run('analyze_sections_with_llm (structured output)',
            checker.analyze_sections_with_llm,
            lambda r: 'error' not in r and r.get('present') is not None
            and not any('LLM error' in m.get('description', '') for m in r.get('missing', [])))

    print('ALL OK' if not failures else f'{len(failures)} FAILED: {failures}')
    sys.exit(1 if failures else 0)


if __name__ == '__main__':
    main()
