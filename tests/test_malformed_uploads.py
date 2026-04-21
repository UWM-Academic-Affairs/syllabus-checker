"""
Tests for malformed DOCX uploads that pass the extension check but have
corrupted or adversarial content.

These tests verify the app returns a graceful error (no unhandled 500) when
it encounters:
  - A zip bomb (valid zip, highly compressible content, large uncompressed size)
  - A corrupted DOCX (random bytes, not a valid zip)
  - An empty zip (valid zip but no word/document.xml entry)
  - An XML bomb (billion-laughs-style deeply nested entities in document.xml)
  - A file of the wrong type renamed to .docx (e.g. a PNG)

All assertions require either:
  - HTTP 200 with error text in the response body, OR
  - An HTTP redirect (3xx) that resolves to a 200 error page (follow_redirects=True)
  - No 500 status code in any case.
"""

import io
import zipfile

import pytest


# ---------------------------------------------------------------------------
# Shared upload helper
# ---------------------------------------------------------------------------

def _post_docx(client, raw_bytes, filename='test.docx'):
    """POST raw bytes as a .docx upload and follow any redirect."""
    return client.post(
        '/upload',
        data={'file': (io.BytesIO(raw_bytes), filename)},
        content_type='multipart/form-data',
        follow_redirects=True,
    )


def _assert_graceful_error(resp):
    """Assert the response is a non-500 page that signals an error to the user."""
    assert resp.status_code != 500, 'App returned an unhandled 500 error'
    body = resp.data.lower()
    assert (
        b'error' in body or b'invalid' in body or b'could not' in body
    ), f'Expected an error message in the response body, got:\n{resp.data[:500]}'


# ---------------------------------------------------------------------------
# Fixtures — all built in-memory (no disk files needed)
# ---------------------------------------------------------------------------

@pytest.fixture
def zip_bomb_docx():
    """
    A .docx that is a syntactically valid zip, but word/document.xml
    contains ~10 MB of uncompressed zeros (compresses to almost nothing).
    The app's 16 MB Flask limit applies to the upload size, not the
    decompressed size, so this fixture fits well under the limit on disk.
    """
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
        # Minimal [Content_Types].xml so python-docx at least gets past the
        # zip-open stage before hitting the oversized document.xml.
        zf.writestr('[Content_Types].xml',
                    '<?xml version="1.0" encoding="UTF-8"?>'
                    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                    '<Override PartName="/word/document.xml"'
                    ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
                    '</Types>')
        # 12 MB of zeros — compresses to a few KB inside the zip.
        zf.writestr('word/document.xml', b'\x00' * 12 * 1024 * 1024)
    return buf.getvalue()


@pytest.fixture
def corrupted_docx():
    """Random bytes with a .docx extension — not a valid zip at all."""
    return bytes(range(256)) * 40 + b'\xde\xad\xbe\xef' * 100


@pytest.fixture
def empty_zip_docx():
    """A valid zip archive that contains no word/document.xml entry."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('README.txt', 'This zip has no Word content.')
    return buf.getvalue()


@pytest.fixture
def xml_bomb_docx():
    """
    A .docx where word/document.xml contains a deeply nested XML structure
    intended to stress XML parsers.  We use repeated element nesting rather
    than DOCTYPE entity expansion (which defusedxml blocks outright) so the
    fixture actually reaches the XML-parsing layer.
    """
    # Build 1 000 levels of nested <w:p> tags — deep enough to be a stress
    # test without being a true denial-of-service payload.
    open_tags = '<w:p>' * 1000
    close_tags = '</w:p>' * 1000
    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"'
        ' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:body>'
        + open_tags
        + '<w:r><w:t>bomb</w:t></w:r>'
        + close_tags
        + '</w:body></w:document>'
    )
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml',
                    '<?xml version="1.0" encoding="UTF-8"?>'
                    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                    '<Override PartName="/word/document.xml"'
                    ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
                    '</Types>')
        zf.writestr('word/document.xml', document_xml.encode('utf-8'))
    return buf.getvalue()


@pytest.fixture
def png_renamed_to_docx():
    """Minimal valid PNG file bytes, uploaded with a .docx filename."""
    # A 1x1 transparent PNG (89 bytes, well-known magic bytes).
    return (
        b'\x89PNG\r\n\x1a\n'          # PNG signature
        b'\x00\x00\x00\rIHDR'         # IHDR chunk length + type
        b'\x00\x00\x00\x01'           # width = 1
        b'\x00\x00\x00\x01'           # height = 1
        b'\x08\x02'                    # bit depth 8, colour type 2 (RGB)
        b'\x00\x00\x00'               # compression, filter, interlace
        b'\x90wS\xde'                 # CRC
        b'\x00\x00\x00\x0cIDATx\x9c'  # IDAT chunk
        b'b\xf8\x0f\x00\x00\x01\x01\x00\x05\x18\xd8N'
        b'\x00\x00\x00\x00IEND\xaeB`\x82'  # IEND
    )


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------

class TestMalformedUploads:

    def test_zip_bomb_does_not_crash_app(self, client, zip_bomb_docx):
        """A zip with a very large decompressed entry must not cause a 500."""
        resp = _post_docx(client, zip_bomb_docx, 'bomb.docx')
        _assert_graceful_error(resp)

    def test_corrupted_docx_returns_error_not_500(self, client, corrupted_docx):
        """Random bytes with a .docx extension must be handled without crashing."""
        resp = _post_docx(client, corrupted_docx, 'corrupt.docx')
        _assert_graceful_error(resp)

    def test_empty_zip_returns_error_not_500(self, client, empty_zip_docx):
        """A valid zip with no word/document.xml must not produce a 500."""
        resp = _post_docx(client, empty_zip_docx, 'empty.docx')
        _assert_graceful_error(resp)

    def test_xml_bomb_returns_error_not_500(self, client, xml_bomb_docx):
        """Deeply nested XML in document.xml must be handled without crashing."""
        resp = _post_docx(client, xml_bomb_docx, 'xmlbomb.docx')
        _assert_graceful_error(resp)

    def test_png_renamed_to_docx_returns_error_not_500(self, client, png_renamed_to_docx):
        """A PNG file uploaded with a .docx extension must not produce a 500."""
        resp = _post_docx(client, png_renamed_to_docx, 'image.docx')
        _assert_graceful_error(resp)
