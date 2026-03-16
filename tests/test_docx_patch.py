"""Tests for docx_patch — XML well-formedness validation."""

import zipfile
from pathlib import Path

import pytest

from docx_patch import validate_xml_wellformedness, patch_docx


# ── validate_xml_wellformedness ─────────────────────────────────────────────


_W_NS = b' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


class TestValidateXmlWellformedness:
    def test_valid_xml_passes(self):
        parts = {
            "word/styles.xml": b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles' + _W_NS + b"></w:styles>",
            "word/document.xml": b"<w:document" + _W_NS + b"><w:body/></w:document>",
        }
        assert validate_xml_wellformedness(parts) == []

    def test_unclosed_tag_detected(self):
        parts = {
            "word/styles.xml": b"<styles><style>",
        }
        errors = validate_xml_wellformedness(parts)
        assert len(errors) == 1
        assert "word/styles.xml" in errors[0]
        assert "XML parse error" in errors[0]

    def test_empty_bytes_detected(self):
        parts = {"word/document.xml": b""}
        errors = validate_xml_wellformedness(parts)
        assert len(errors) == 1
        assert "word/document.xml" in errors[0]

    def test_xml_declaration_with_valid_content_passes(self):
        content = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>'
        assert validate_xml_wellformedness({"[Content_Types].xml": content}) == []

    def test_mixed_valid_and_invalid(self):
        parts = {
            "word/styles.xml": b"<root/>",
            "word/settings.xml": b"<broken><",
            "word/document.xml": b"<root/>",
        }
        errors = validate_xml_wellformedness(parts)
        assert len(errors) == 1
        assert "word/settings.xml" in errors[0]


# ── patch_docx integration ──────────────────────────────────────────────────


class TestPatchDocxXmlValidation:
    def _make_minimal_docx(self, path: Path) -> None:
        """Create a minimal valid DOCX (ZIP with one XML entry)."""
        with zipfile.ZipFile(path, "w") as zf:
            zf.writestr("word/document.xml", b"<w:document" + _W_NS + b"/>")

    def test_malformed_xml_prevents_repack(self, tmp_path):
        src = tmp_path / "input.docx"
        out = tmp_path / "output.docx"
        self._make_minimal_docx(src)

        with pytest.raises(RuntimeError, match="XML well-formedness check failed"):
            patch_docx(
                src_docx=src,
                out_docx=out,
                replacements={"word/document.xml": b"<broken><"},
            )
        assert not out.exists()

    def test_valid_xml_allows_repack(self, tmp_path):
        src = tmp_path / "input.docx"
        out = tmp_path / "output.docx"
        self._make_minimal_docx(src)

        patch_docx(
            src_docx=src,
            out_docx=out,
            replacements={"word/document.xml": b"<w:document" + _W_NS + b"><w:body/></w:document>"},
        )
        assert out.exists()


class TestPatchDocxHeaderFooterSupport:
    def _make_minimal_docx(self, path: Path) -> None:
        with zipfile.ZipFile(path, "w") as zf:
            zf.writestr("word/document.xml", b"<w:document" + _W_NS + b"/>")

    def _make_docx_with_header(self, path: Path) -> None:
        with zipfile.ZipFile(path, "w") as zf:
            zf.writestr("word/document.xml", b"<w:document" + _W_NS + b"/>")
            zf.writestr("word/header1.xml", b"<w:hdr" + _W_NS + b"/>")
            zf.writestr("word/_rels/header1.xml.rels", b"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"/>")

    def test_allows_media_binary_replacement(self, tmp_path):
        src = tmp_path / "input.docx"
        out = tmp_path / "output.docx"
        self._make_minimal_docx(src)

        patch_docx(
            src_docx=src,
            out_docx=out,
            replacements={
                "word/document.xml": b"<w:document" + _W_NS + b"><w:body/></w:document>",
                "word/media/image1.png": b"\x89PNG",
            },
        )
        with zipfile.ZipFile(out, "r") as zf:
            assert zf.read("word/media/image1.png") == b"\x89PNG"

    def test_exclude_parts_drops_old_header_parts(self, tmp_path):
        src = tmp_path / "input.docx"
        out = tmp_path / "output.docx"
        self._make_docx_with_header(src)

        patch_docx(
            src_docx=src,
            out_docx=out,
            replacements={"word/document.xml": b"<w:document" + _W_NS + b"><w:body/></w:document>"},
            exclude_parts={"word/header1.xml", "word/_rels/header1.xml.rels"},
        )
        with zipfile.ZipFile(out, "r") as zf:
            assert "word/header1.xml" not in zf.namelist()
            assert "word/_rels/header1.xml.rels" not in zf.namelist()
