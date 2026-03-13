"""Tests for core.xml_helpers — paragraph-level XML manipulation."""

import pytest
from core.xml_helpers import (
    apply_pstyle_to_paragraph_block,
    strip_run_font_formatting,
    iter_paragraph_xml_blocks,
    paragraph_text_from_block,
    paragraph_contains_sectpr,
    paragraph_pstyle_from_block,
    paragraph_numpr_from_block,
    strip_conflicting_direct_ppr,
)


# ── apply_pstyle_to_paragraph_block ──────────────────────────────────────────

class TestApplyPstyle:
    def test_replace_existing_pstyle(self):
        p = '<w:p><w:pPr><w:pStyle w:val="OldStyle"/></w:pPr><w:r><w:t>Hello</w:t></w:r></w:p>'
        result = apply_pstyle_to_paragraph_block(p, "NewStyle")
        assert 'w:val="NewStyle"' in result
        assert 'w:val="OldStyle"' not in result

    def test_self_closing_ppr(self):
        p = '<w:p><w:pPr/><w:r><w:t>Hello</w:t></w:r></w:p>'
        result = apply_pstyle_to_paragraph_block(p, "MyStyle")
        assert '<w:pPr><w:pStyle w:val="MyStyle"/></w:pPr>' in result

    def test_open_ppr_no_pstyle(self):
        p = '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>Hello</w:t></w:r></w:p>'
        result = apply_pstyle_to_paragraph_block(p, "MyStyle")
        assert '<w:pStyle w:val="MyStyle"/>' in result
        assert '<w:jc w:val="center"/>' in result

    def test_no_ppr_at_all(self):
        p = '<w:p><w:r><w:t>Hello</w:t></w:r></w:p>'
        result = apply_pstyle_to_paragraph_block(p, "MyStyle")
        assert '<w:pPr><w:pStyle w:val="MyStyle"/></w:pPr>' in result

    def test_sectpr_unchanged(self):
        p = '<w:p><w:pPr><w:sectPr><w:pgSz/></w:sectPr></w:pPr></w:p>'
        result = apply_pstyle_to_paragraph_block(p, "MyStyle")
        assert result == p  # Must not modify sectPr paragraphs


# ── strip_run_font_formatting ────────────────────────────────────────────────

class TestStripRunFontFormatting:
    def test_strip_rfonts_sz_szcs(self):
        p = (
            '<w:p><w:r><w:rPr>'
            '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>'
            '<w:sz w:val="20"/>'
            '<w:szCs w:val="20"/>'
            '</w:rPr><w:t>Hello</w:t></w:r></w:p>'
        )
        result = strip_run_font_formatting(p)
        assert '<w:rFonts' not in result
        assert '<w:sz' not in result
        assert '<w:szCs' not in result
        assert '<w:t>Hello</w:t>' in result

    def test_preserve_bold(self):
        p = (
            '<w:p><w:r><w:rPr>'
            '<w:rFonts w:ascii="Arial"/>'
            '<w:b/>'
            '</w:rPr><w:t>Bold</w:t></w:r></w:p>'
        )
        result = strip_run_font_formatting(p)
        assert '<w:rFonts' not in result
        assert '<w:b/>' in result

    def test_no_font_formatting_unchanged(self):
        p = '<w:p><w:r><w:rPr><w:b/><w:i/></w:rPr><w:t>BI</w:t></w:r></w:p>'
        result = strip_run_font_formatting(p)
        assert result == p


class TestStripConflictingDirectPpr:
    def test_removes_jc_ind_spacing_but_keeps_numpr(self):
        p = (
            '<w:p><w:pPr>'
            '<w:numPr><w:numId w:val="4"/><w:ilvl w:val="1"/></w:numPr>'
            '<w:jc w:val="center"/><w:ind w:left="720"/><w:spacing w:before="120"/>'
            '</w:pPr><w:r><w:t>X</w:t></w:r></w:p>'
        )
        result = strip_conflicting_direct_ppr(p)
        assert '<w:jc' not in result
        assert '<w:ind' not in result
        assert '<w:spacing' not in result
        assert '<w:numPr>' in result

    def test_sectpr_unchanged(self):
        p = '<w:p><w:pPr><w:sectPr/><w:jc w:val="center"/></w:pPr></w:p>'
        result = strip_conflicting_direct_ppr(p)
        assert result == p

    def test_multiple_runs(self):
        p = (
            '<w:p>'
            '<w:r><w:rPr><w:rFonts w:ascii="Arial"/><w:sz w:val="20"/></w:rPr><w:t>A</w:t></w:r>'
            '<w:r><w:rPr><w:rFonts w:ascii="Times"/><w:b/></w:rPr><w:t>B</w:t></w:r>'
            '</w:p>'
        )
        result = strip_run_font_formatting(p)
        assert '<w:rFonts' not in result
        assert '<w:sz' not in result
        assert '<w:b/>' in result
        assert '<w:t>A</w:t>' in result
        assert '<w:t>B</w:t>' in result

    def test_sectpr_unchanged(self):
        p = '<w:p><w:pPr><w:sectPr/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial"/></w:rPr><w:t>X</w:t></w:r></w:p>'
        result = strip_run_font_formatting(p)
        assert result == p


# ── iter_paragraph_xml_blocks ────────────────────────────────────────────────

class TestIterParagraphXmlBlocks:
    def test_basic_document(self):
        doc = (
            '<?xml version="1.0"?>'
            '<w:document><w:body>'
            '<w:p><w:r><w:t>First</w:t></w:r></w:p>'
            '<w:p><w:r><w:t>Second</w:t></w:r></w:p>'
            '<w:p><w:r><w:t>Third</w:t></w:r></w:p>'
            '</w:body></w:document>'
        )
        blocks = list(iter_paragraph_xml_blocks(doc))
        assert len(blocks) == 3

    def test_correct_positions(self):
        doc = '<w:body><w:p><w:t>A</w:t></w:p></w:body>'
        blocks = list(iter_paragraph_xml_blocks(doc))
        assert len(blocks) == 1
        start, end, p_xml = blocks[0]
        assert doc[start:end] == p_xml
        assert '<w:t>A</w:t>' in p_xml

    def test_paragraph_count(self):
        paras = ''.join(f'<w:p><w:r><w:t>P{i}</w:t></w:r></w:p>' for i in range(10))
        doc = f'<w:body>{paras}</w:body>'
        blocks = list(iter_paragraph_xml_blocks(doc))
        assert len(blocks) == 10


# ── paragraph_text_from_block ────────────────────────────────────────────────

class TestParagraphTextFromBlock:
    def test_basic_text(self):
        p = '<w:p><w:r><w:t>Hello World</w:t></w:r></w:p>'
        assert paragraph_text_from_block(p) == "Hello World"

    def test_multiple_runs(self):
        p = '<w:p><w:r><w:t>Hello </w:t></w:r><w:r><w:t>World</w:t></w:r></w:p>'
        assert paragraph_text_from_block(p) == "Hello World"

    def test_empty_paragraph(self):
        p = '<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr></w:p>'
        assert paragraph_text_from_block(p) == ""

    def test_html_entities(self):
        p = '<w:p><w:r><w:t>A &amp; B</w:t></w:r></w:p>'
        assert paragraph_text_from_block(p) == "A & B"


# ── paragraph_contains_sectpr ────────────────────────────────────────────────

class TestParagraphContainsSectpr:
    def test_with_sectpr(self):
        p = '<w:p><w:pPr><w:sectPr><w:pgSz/></w:sectPr></w:pPr></w:p>'
        assert paragraph_contains_sectpr(p) is True

    def test_without_sectpr(self):
        p = '<w:p><w:r><w:t>Normal</w:t></w:r></w:p>'
        assert paragraph_contains_sectpr(p) is False


# ── paragraph_pstyle_from_block ──────────────────────────────────────────────

class TestParagraphPstyleFromBlock:
    def test_with_pstyle(self):
        p = '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>H</w:t></w:r></w:p>'
        assert paragraph_pstyle_from_block(p) == "Heading1"

    def test_without_pstyle(self):
        p = '<w:p><w:r><w:t>Normal</w:t></w:r></w:p>'
        assert paragraph_pstyle_from_block(p) is None


# ── paragraph_numpr_from_block ───────────────────────────────────────────────

class TestParagraphNumprFromBlock:
    def test_with_numpr(self):
        p = '<w:p><w:pPr><w:numPr><w:numId w:val="5"/><w:ilvl w:val="2"/></w:numPr></w:pPr></w:p>'
        result = paragraph_numpr_from_block(p)
        assert result["numId"] == "5"
        assert result["ilvl"] == "2"

    def test_without_numpr(self):
        p = '<w:p><w:r><w:t>No list</w:t></w:r></w:p>'
        result = paragraph_numpr_from_block(p)
        assert result["numId"] is None
        assert result["ilvl"] is None
