"""Tests for arch_env_applier.py — settings, font table, and sync mode tests."""

import pytest
from pathlib import Path

from arch_env_applier import (
    apply_settings, apply_font_table,
    _apply_page_layout, _apply_headers_footers,
    apply_environment_to_target,
)


# ---------------------------------------------------------------------------
# Helpers to build minimal extract directories
# ---------------------------------------------------------------------------

_CONTENT_TYPES_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '</Types>'
)

_RELS_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '  <Relationship Id="rId1" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
    'Target="document.xml"/>'
    '</Relationships>'
)

_SETTINGS_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '</w:settings>'
)

_SETTINGS_WITH_COMPAT_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '  <w:compat><w:useFELayout/></w:compat>'
    '</w:settings>'
)

_FONT_TABLE_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '  <w:font w:name="Calibri"><w:panose1 w:val="020F0502020204030204"/></w:font>'
    '</w:fonts>'
)


def _setup_extract_dir(tmp_path, content_types=True, rels=True):
    """Create a minimal extract directory with optional plumbing files."""
    (tmp_path / "[Content_Types].xml").write_text(
        _CONTENT_TYPES_XML if content_types else "", encoding="utf-8"
    )
    rels_dir = tmp_path / "word" / "_rels"
    rels_dir.mkdir(parents=True, exist_ok=True)
    if rels:
        (rels_dir / "document.xml.rels").write_text(_RELS_XML, encoding="utf-8")
    (tmp_path / "word").mkdir(exist_ok=True)
    return tmp_path


# ---------------------------------------------------------------------------
# Part A — apply_settings() tests
# ---------------------------------------------------------------------------

class TestApplySettingsRejectsMalformed:
    """Malformed compat_xml must be rejected without mutating anything."""

    def test_missing_close_tag(self, tmp_path):
        extract = _setup_extract_dir(tmp_path)
        settings_path = extract / "word" / "settings.xml"
        settings_path.write_text(_SETTINGS_XML, encoding="utf-8")

        registry = {
            "settings": {
                "compat": {"compat_xml": "<w:compat><w:useFELayout/>"}  # no close
            }
        }
        log = []
        apply_settings(extract, registry, log)

        assert any("WARNING" in m and "Skipping compat" in m for m in log)
        # settings.xml must be untouched
        assert settings_path.read_text(encoding="utf-8") == _SETTINGS_XML

    def test_missing_open_tag(self, tmp_path):
        extract = _setup_extract_dir(tmp_path)
        settings_path = extract / "word" / "settings.xml"
        settings_path.write_text(_SETTINGS_XML, encoding="utf-8")

        registry = {
            "settings": {
                "compat": {"compat_xml": "</w:compat>"}
            }
        }
        log = []
        apply_settings(extract, registry, log)

        assert any("WARNING" in m and "Skipping compat" in m for m in log)
        assert settings_path.read_text(encoding="utf-8") == _SETTINGS_XML

    def test_empty_string_compat_skips(self, tmp_path):
        """Empty string compat_xml is falsy — skipped before validation."""
        extract = _setup_extract_dir(tmp_path)
        settings_path = extract / "word" / "settings.xml"
        settings_path.write_text(_SETTINGS_XML, encoding="utf-8")

        registry = {"settings": {"compat": {"compat_xml": ""}}}
        log = []
        apply_settings(extract, registry, log)

        assert any("No compat" in m or "skipping" in m.lower() for m in log)
        assert settings_path.read_text(encoding="utf-8") == _SETTINGS_XML

    def test_none_compat_skips(self, tmp_path):
        """None compat_xml is falsy — skipped before validation."""
        extract = _setup_extract_dir(tmp_path)
        settings_path = extract / "word" / "settings.xml"
        settings_path.write_text(_SETTINGS_XML, encoding="utf-8")

        registry = {"settings": {"compat": {"compat_xml": None}}}
        log = []
        apply_settings(extract, registry, log)

        assert any("No compat" in m or "skipping" in m.lower() for m in log)
        assert settings_path.read_text(encoding="utf-8") == _SETTINGS_XML


class TestApplySettingsCreatesWhenMissing:
    """When target has no settings.xml, create one and wire plumbing."""

    def test_creates_settings_with_compat(self, tmp_path):
        extract = _setup_extract_dir(tmp_path)
        compat = "<w:compat><w:useFELayout/></w:compat>"
        registry = {"settings": {"compat": {"compat_xml": compat}}}

        log = []
        apply_settings(extract, registry, log)

        settings_path = extract / "word" / "settings.xml"
        assert settings_path.exists()
        content = settings_path.read_text(encoding="utf-8")
        assert "<w:useFELayout/>" in content
        assert "<w:settings" in content
        assert "</w:settings>" in content

    def test_wires_content_types(self, tmp_path):
        extract = _setup_extract_dir(tmp_path)
        compat = "<w:compat><w:useFELayout/></w:compat>"
        registry = {"settings": {"compat": {"compat_xml": compat}}}

        log = []
        apply_settings(extract, registry, log)

        ct = (extract / "[Content_Types].xml").read_text(encoding="utf-8")
        assert 'PartName="/word/settings.xml"' in ct
        assert "wordprocessingml.settings+xml" in ct

    def test_wires_rels(self, tmp_path):
        extract = _setup_extract_dir(tmp_path)
        compat = "<w:compat><w:useFELayout/></w:compat>"
        registry = {"settings": {"compat": {"compat_xml": compat}}}

        log = []
        apply_settings(extract, registry, log)

        rels = (extract / "word" / "_rels" / "document.xml.rels").read_text(encoding="utf-8")
        assert 'Target="settings.xml"' in rels
        assert "relationships/settings" in rels


class TestApplySettingsReplacesExisting:
    """Existing compat block should be replaced."""

    def test_replaces_compat(self, tmp_path):
        extract = _setup_extract_dir(tmp_path)
        settings_path = extract / "word" / "settings.xml"
        settings_path.write_text(_SETTINGS_WITH_COMPAT_XML, encoding="utf-8")

        new_compat = "<w:compat><w:compatSetting w:name=\"test\" w:val=\"1\"/></w:compat>"
        registry = {"settings": {"compat": {"compat_xml": new_compat}}}

        log = []
        apply_settings(extract, registry, log)

        content = settings_path.read_text(encoding="utf-8")
        assert "w:compatSetting" in content
        assert "<w:useFELayout/>" not in content
        assert any("Replaced" in m for m in log)


class TestApplySettingsInserts:
    """When settings.xml exists but has no compat, insert it."""

    def test_inserts_compat(self, tmp_path):
        extract = _setup_extract_dir(tmp_path)
        settings_path = extract / "word" / "settings.xml"
        settings_path.write_text(_SETTINGS_XML, encoding="utf-8")

        compat = "<w:compat><w:useFELayout/></w:compat>"
        registry = {"settings": {"compat": {"compat_xml": compat}}}

        log = []
        apply_settings(extract, registry, log)

        content = settings_path.read_text(encoding="utf-8")
        assert "<w:useFELayout/>" in content
        assert "</w:settings>" in content
        assert any("Inserted" in m for m in log)


class TestApplySettingsIdempotent:
    """Calling twice should not duplicate plumbing entries."""

    def test_idempotent_content_types_and_rels(self, tmp_path):
        extract = _setup_extract_dir(tmp_path)
        compat = "<w:compat><w:useFELayout/></w:compat>"
        registry = {"settings": {"compat": {"compat_xml": compat}}}

        apply_settings(extract, registry, [])
        apply_settings(extract, registry, [])

        ct = (extract / "[Content_Types].xml").read_text(encoding="utf-8")
        assert ct.count('PartName="/word/settings.xml"') == 1

        rels = (extract / "word" / "_rels" / "document.xml.rels").read_text(encoding="utf-8")
        assert rels.count('Target="settings.xml"') == 1


# ---------------------------------------------------------------------------
# Part B — apply_font_table() tests
# ---------------------------------------------------------------------------

class TestApplyFontTableCreatesWithPlumbing:
    """New fontTable.xml should be wired into content types and rels."""

    def test_creates_font_table(self, tmp_path):
        extract = _setup_extract_dir(tmp_path)
        arch_font_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '  <w:font w:name="Arial"><w:panose1 w:val="020B0604020202020204"/></w:font>'
            '</w:fonts>'
        )
        registry = {"fonts": {"font_table_xml": arch_font_xml}}

        log = []
        apply_font_table(extract, registry, log)

        font_path = extract / "word" / "fontTable.xml"
        assert font_path.exists()
        assert "Arial" in font_path.read_text(encoding="utf-8")

    def test_wires_content_types(self, tmp_path):
        extract = _setup_extract_dir(tmp_path)
        arch_font_xml = (
            '<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '  <w:font w:name="Arial"><w:panose1 w:val="020B0604020202020204"/></w:font>'
            '</w:fonts>'
        )
        registry = {"fonts": {"font_table_xml": arch_font_xml}}

        log = []
        apply_font_table(extract, registry, log)

        ct = (extract / "[Content_Types].xml").read_text(encoding="utf-8")
        assert 'PartName="/word/fontTable.xml"' in ct
        assert "wordprocessingml.fontTable+xml" in ct

    def test_wires_rels(self, tmp_path):
        extract = _setup_extract_dir(tmp_path)
        arch_font_xml = (
            '<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '  <w:font w:name="Arial"><w:panose1 w:val="020B0604020202020204"/></w:font>'
            '</w:fonts>'
        )
        registry = {"fonts": {"font_table_xml": arch_font_xml}}

        log = []
        apply_font_table(extract, registry, log)

        rels = (extract / "word" / "_rels" / "document.xml.rels").read_text(encoding="utf-8")
        assert 'Target="fontTable.xml"' in rels
        assert "relationships/fontTable" in rels

    def test_content_type_mime(self, tmp_path):
        """Content type must use the correct OOXML MIME type."""
        extract = _setup_extract_dir(tmp_path)
        arch_font_xml = (
            '<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '  <w:font w:name="Arial"><w:panose1 w:val="020B0604020202020204"/></w:font>'
            '</w:fonts>'
        )
        registry = {"fonts": {"font_table_xml": arch_font_xml}}
        apply_font_table(extract, registry, [])

        ct = (extract / "[Content_Types].xml").read_text(encoding="utf-8")
        assert "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml" in ct

    def test_rels_relationship_type_uri(self, tmp_path):
        """Rels must use the correct OOXML relationship type URI."""
        extract = _setup_extract_dir(tmp_path)
        arch_font_xml = (
            '<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '  <w:font w:name="Arial"><w:panose1 w:val="020B0604020202020204"/></w:font>'
            '</w:fonts>'
        )
        registry = {"fonts": {"font_table_xml": arch_font_xml}}
        apply_font_table(extract, registry, [])

        rels = (extract / "word" / "_rels" / "document.xml.rels").read_text(encoding="utf-8")
        assert "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" in rels


class TestApplyFontTableMerge:
    """When fontTable exists, only missing fonts are added."""

    def test_adds_missing_fonts(self, tmp_path):
        extract = _setup_extract_dir(tmp_path)
        font_path = extract / "word" / "fontTable.xml"
        font_path.write_text(_FONT_TABLE_XML, encoding="utf-8")

        arch_font_xml = (
            '<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '  <w:font w:name="Calibri"><w:panose1 w:val="020F0502020204030204"/></w:font>'
            '  <w:font w:name="Arial"><w:panose1 w:val="020B0604020202020204"/></w:font>'
            '</w:fonts>'
        )
        registry = {"fonts": {"font_table_xml": arch_font_xml}}

        log = []
        apply_font_table(extract, registry, log)

        content = font_path.read_text(encoding="utf-8")
        assert "Arial" in content
        assert content.count('w:name="Calibri"') == 1  # not duplicated
        assert any("1 font" in m for m in log)

    def test_multiple_existing_fonts_only_missing_added(self, tmp_path):
        """Target with multiple fonts — only truly missing ones added."""
        extract = _setup_extract_dir(tmp_path)
        font_path = extract / "word" / "fontTable.xml"
        multi_font_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '  <w:font w:name="Calibri"><w:panose1 w:val="020F0502020204030204"/></w:font>'
            '  <w:font w:name="Times New Roman"><w:panose1 w:val="02020603050405020304"/></w:font>'
            '</w:fonts>'
        )
        font_path.write_text(multi_font_xml, encoding="utf-8")

        arch_font_xml = (
            '<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '  <w:font w:name="Calibri"><w:panose1 w:val="020F0502020204030204"/></w:font>'
            '  <w:font w:name="Times New Roman"><w:panose1 w:val="02020603050405020304"/></w:font>'
            '  <w:font w:name="Arial"><w:panose1 w:val="020B0604020202020204"/></w:font>'
            '</w:fonts>'
        )
        registry = {"fonts": {"font_table_xml": arch_font_xml}}

        log = []
        apply_font_table(extract, registry, log)

        content = font_path.read_text(encoding="utf-8")
        assert "Arial" in content
        assert content.count('w:name="Calibri"') == 1
        assert content.count('w:name="Times New Roman"') == 1

    def test_skips_when_all_present(self, tmp_path):
        extract = _setup_extract_dir(tmp_path)
        font_path = extract / "word" / "fontTable.xml"
        font_path.write_text(_FONT_TABLE_XML, encoding="utf-8")

        arch_font_xml = (
            '<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '  <w:font w:name="Calibri"><w:panose1 w:val="020F0502020204030204"/></w:font>'
            '</w:fonts>'
        )
        registry = {"fonts": {"font_table_xml": arch_font_xml}}

        log = []
        apply_font_table(extract, registry, log)

        assert any("already present" in m for m in log)


class TestApplyFontTableIdempotent:
    """Calling twice should not duplicate plumbing entries."""

    def test_idempotent_plumbing(self, tmp_path):
        extract = _setup_extract_dir(tmp_path)
        arch_font_xml = (
            '<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '  <w:font w:name="Arial"><w:panose1 w:val="020B0604020202020204"/></w:font>'
            '</w:fonts>'
        )
        registry = {"fonts": {"font_table_xml": arch_font_xml}}

        apply_font_table(extract, registry, [])
        apply_font_table(extract, registry, [])

        ct = (extract / "[Content_Types].xml").read_text(encoding="utf-8")
        assert ct.count('PartName="/word/fontTable.xml"') == 1

        rels = (extract / "word" / "_rels" / "document.xml.rels").read_text(encoding="utf-8")
        assert rels.count('Target="fontTable.xml"') == 1


class TestApplyFontTableValidation:
    """Post-mutation validation catches malformed results."""

    def test_valid_result_no_warning(self, tmp_path):
        extract = _setup_extract_dir(tmp_path)
        arch_font_xml = (
            '<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '  <w:font w:name="Arial"><w:panose1 w:val="020B0604020202020204"/></w:font>'
            '</w:fonts>'
        )
        registry = {"fonts": {"font_table_xml": arch_font_xml}}

        log = []
        apply_font_table(extract, registry, log)

        assert not any("WARNING" in m and "malformed" in m for m in log)


# ── P2-010: Template sync environment extension ──────────────────────────


_W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

_DOC_XML_WITH_SECTPR = (
    f'<w:document {_W_NS}><w:body>'
    '<w:p><w:r><w:t>Hello</w:t></w:r></w:p>'
    f'<w:sectPr {_W_NS}><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>'
    '</w:body></w:document>'
)


class TestApplyPageLayout:
    def test_replaces_final_sectpr(self, tmp_path):
        extract = tmp_path / "extracted"
        word_dir = extract / "word"
        word_dir.mkdir(parents=True)
        (word_dir / "document.xml").write_text(_DOC_XML_WITH_SECTPR, encoding="utf-8")

        new_sectpr = f'<w:sectPr {_W_NS}><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>'
        log = []
        _apply_page_layout(extract, new_sectpr, log)

        result = (word_dir / "document.xml").read_text(encoding="utf-8")
        assert 'w:w="11906"' in result
        assert 'w:w="12240"' not in result

    def test_no_sectpr_warns(self, tmp_path):
        extract = tmp_path / "extracted"
        word_dir = extract / "word"
        word_dir.mkdir(parents=True)
        (word_dir / "document.xml").write_text(
            f'<w:document {_W_NS}><w:body/></w:document>', encoding="utf-8"
        )

        log = []
        _apply_page_layout(extract, "<w:sectPr/>", log)
        assert any("No sectPr found" in m for m in log)


class TestApplyHeadersFooters:
    def test_writes_header_files(self, tmp_path):
        extract = tmp_path / "extracted"
        word_dir = extract / "word"
        word_dir.mkdir(parents=True)

        hf = {
            "word/header1.xml": '<w:hdr xmlns:w="http://example.com"><w:p/></w:hdr>',
            "word/footer1.xml": '<w:ftr xmlns:w="http://example.com"><w:p/></w:ftr>',
        }
        log = []
        _apply_headers_footers(extract, hf, log)

        assert (word_dir / "header1.xml").exists()
        assert (word_dir / "footer1.xml").exists()
        assert any("2 header/footer" in m for m in log)

    def test_skips_non_header_paths(self, tmp_path):
        extract = tmp_path / "extracted"
        word_dir = extract / "word"
        word_dir.mkdir(parents=True)

        hf = {"word/styles.xml": "<data/>"}
        log = []
        _apply_headers_footers(extract, hf, log)
        assert any("Skipping non-header/footer" in m for m in log)


class TestTemplateSyncMode:
    def test_body_only_skips_page_artifacts(self, tmp_path):
        extract = tmp_path / "extracted"
        word_dir = extract / "word"
        word_dir.mkdir(parents=True)
        (word_dir / "document.xml").write_text(_DOC_XML_WITH_SECTPR, encoding="utf-8")
        (word_dir / "styles.xml").write_text(
            f'<w:styles {_W_NS}></w:styles>', encoding="utf-8"
        )

        registry = {
            "page_layout": '<w:sectPr><w:pgSz w:w="999"/></w:sectPr>',
            "headers_footers": {"word/header1.xml": "<hdr/>"},
        }
        log = []
        apply_environment_to_target(extract, registry, log, sync_mode="body_only",
                                    apply_theme_flag=False, apply_settings_flag=False,
                                    apply_fonts_flag=False, apply_doc_defaults_flag=False)

        # sectPr should be unchanged
        doc = (word_dir / "document.xml").read_text(encoding="utf-8")
        assert 'w:w="12240"' in doc
        # No header file should be written
        assert not (word_dir / "header1.xml").exists()

    def test_template_sync_applies_page_artifacts(self, tmp_path):
        extract = tmp_path / "extracted"
        word_dir = extract / "word"
        word_dir.mkdir(parents=True)
        (word_dir / "document.xml").write_text(_DOC_XML_WITH_SECTPR, encoding="utf-8")
        (word_dir / "styles.xml").write_text(
            f'<w:styles {_W_NS}></w:styles>', encoding="utf-8"
        )

        new_sectpr = f'<w:sectPr {_W_NS}><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>'
        registry = {
            "page_layout": new_sectpr,
            "headers_footers": {
                "word/header1.xml": '<w:hdr xmlns:w="http://example.com"><w:p/></w:hdr>',
            },
        }
        log = []
        apply_environment_to_target(extract, registry, log, sync_mode="template_sync",
                                    apply_theme_flag=False, apply_settings_flag=False,
                                    apply_fonts_flag=False, apply_doc_defaults_flag=False)

        # sectPr should be replaced
        doc = (word_dir / "document.xml").read_text(encoding="utf-8")
        assert 'w:w="11906"' in doc
        assert 'w:w="12240"' not in doc
        # Header file should be written
        assert (word_dir / "header1.xml").exists()
