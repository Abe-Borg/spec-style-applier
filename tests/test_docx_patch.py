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


def _make_minimal_docx(path: Path) -> None:
    """Create a minimal valid DOCX (ZIP with one XML entry)."""
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("word/document.xml", b"<w:document" + _W_NS + b"/>")


class TestPatchDocxXmlValidation:
    def _make_minimal_docx(self, path: Path) -> None:
        _make_minimal_docx(path)

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


# ── Sync mode allowlist tests ─────────────────────────────────────────────


class TestPatchDocxSyncModes:
    """Tests for body_only vs template_sync patch target allowlists."""

    def test_body_only_rejects_header_patch(self, tmp_path):
        src = tmp_path / "input.docx"
        out = tmp_path / "output.docx"
        _make_minimal_docx(src)

        with pytest.raises(RuntimeError, match="Forbidden patch target"):
            patch_docx(
                src_docx=src,
                out_docx=out,
                replacements={"word/header1.xml": b"<w:hdr" + _W_NS + b"/>"},
                sync_mode="body_only",
            )

    def test_body_only_rejects_footer_patch(self, tmp_path):
        src = tmp_path / "input.docx"
        out = tmp_path / "output.docx"
        _make_minimal_docx(src)

        with pytest.raises(RuntimeError, match="Forbidden patch target"):
            patch_docx(
                src_docx=src,
                out_docx=out,
                replacements={"word/footer1.xml": b"<w:ftr" + _W_NS + b"/>"},
                sync_mode="body_only",
            )

    def test_template_sync_allows_header_patch(self, tmp_path):
        src = tmp_path / "input.docx"
        out = tmp_path / "output.docx"
        _make_minimal_docx(src)

        patch_docx(
            src_docx=src,
            out_docx=out,
            replacements={"word/header1.xml": b"<w:hdr" + _W_NS + b"/>"},
            sync_mode="template_sync",
        )
        assert out.exists()

    def test_template_sync_allows_footer_patch(self, tmp_path):
        src = tmp_path / "input.docx"
        out = tmp_path / "output.docx"
        _make_minimal_docx(src)

        patch_docx(
            src_docx=src,
            out_docx=out,
            replacements={"word/footer1.xml": b"<w:ftr" + _W_NS + b"/>"},
            sync_mode="template_sync",
        )
        assert out.exists()

    def test_body_only_allows_styles_patch(self, tmp_path):
        src = tmp_path / "input.docx"
        out = tmp_path / "output.docx"
        _make_minimal_docx(src)

        patch_docx(
            src_docx=src,
            out_docx=out,
            replacements={"word/styles.xml": b"<w:styles" + _W_NS + b"/>"},
            sync_mode="body_only",
        )
        assert out.exists()

    def test_illegal_target_rejected_in_both_modes(self, tmp_path):
        src = tmp_path / "input.docx"
        out = tmp_path / "output.docx"
        _make_minimal_docx(src)

        for mode in ("body_only", "template_sync"):
            if out.exists():
                out.unlink()
            with pytest.raises(RuntimeError, match="Illegal patch target"):
                patch_docx(
                    src_docx=src,
                    out_docx=out,
                    replacements={"word/media/image1.png": b"\x89PNG"},
                    sync_mode=mode,
                )
