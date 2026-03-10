"""Tests for core.style_import — style extraction and materialization."""

import pytest
from core.style_import import (
    ensure_explicit_numpr_from_current_style,
    materialize_arch_style_block,
    _extract_style_block,
    _extract_basedOn,
    _find_style_numpr_in_chain,
)


# ── Sample styles.xml fragments ─────────────────────────────────────────────

STYLES_WITH_NUMPR = '''
<w:styles>
  <w:style w:type="paragraph" w:styleId="ListBullet">
    <w:name w:val="List Bullet"/>
    <w:pPr>
      <w:numPr><w:numId w:val="1"/><w:ilvl w:val="0"/></w:numPr>
    </w:pPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="ListBullet2">
    <w:name w:val="List Bullet 2"/>
    <w:basedOn w:val="ListBullet"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="ListBullet3">
    <w:name w:val="List Bullet 3"/>
    <w:basedOn w:val="ListBullet2"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
</w:styles>
'''

STYLES_FOR_MATERIALIZE = '''
<w:docDefaults>
  <w:rPrDefault>
    <w:rPr>
      <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>
      <w:sz w:val="22"/>
      <w:szCs w:val="22"/>
      <w:lang w:val="en-US"/>
    </w:rPr>
  </w:rPrDefault>
  <w:pPrDefault>
    <w:pPr>
      <w:spacing w:after="160"/>
    </w:pPr>
  </w:pPrDefault>
</w:docDefaults>
<w:styles>
  <w:style w:type="paragraph" w:styleId="CSI-Part">
    <w:name w:val="CSI-Part"/>
    <w:rPr>
      <w:b/>
      <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="CSI-Article">
    <w:name w:val="CSI-Article"/>
    <w:basedOn w:val="CSI-Part"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="NoRpr">
    <w:name w:val="NoRpr"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="CompleteRpr">
    <w:name w:val="CompleteRpr"/>
    <w:rPr>
      <w:rFonts w:ascii="Times" w:hAnsi="Times"/>
      <w:sz w:val="24"/>
      <w:szCs w:val="24"/>
      <w:lang w:val="en-GB"/>
    </w:rPr>
  </w:style>
</w:styles>
'''


# ── ensure_explicit_numpr_from_current_style ─────────────────────────────────

class TestEnsureExplicitNumpr:
    def test_already_has_numpr(self):
        p = '<w:p><w:pPr><w:pStyle w:val="ListBullet"/><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>A</w:t></w:r></w:p>'
        result = ensure_explicit_numpr_from_current_style(p, STYLES_WITH_NUMPR)
        assert result == p  # Unchanged

    def test_no_pstyle(self):
        p = '<w:p><w:r><w:t>Plain</w:t></w:r></w:p>'
        result = ensure_explicit_numpr_from_current_style(p, STYLES_WITH_NUMPR)
        assert result == p  # Unchanged

    def test_style_has_numpr_depth_1(self):
        p = '<w:p><w:pPr><w:pStyle w:val="ListBullet"/></w:pPr><w:r><w:t>Item</w:t></w:r></w:p>'
        result = ensure_explicit_numpr_from_current_style(p, STYLES_WITH_NUMPR)
        assert '<w:numPr>' in result
        assert '<w:numId w:val="1"/>' in result

    def test_style_has_numpr_depth_2(self):
        """ListBullet2 -> basedOn ListBullet which has numPr."""
        p = '<w:p><w:pPr><w:pStyle w:val="ListBullet2"/></w:pPr><w:r><w:t>Item</w:t></w:r></w:p>'
        result = ensure_explicit_numpr_from_current_style(p, STYLES_WITH_NUMPR)
        assert '<w:numPr>' in result

    def test_style_has_numpr_depth_3(self):
        """ListBullet3 -> ListBullet2 -> ListBullet which has numPr."""
        p = '<w:p><w:pPr><w:pStyle w:val="ListBullet3"/></w:pPr><w:r><w:t>Item</w:t></w:r></w:p>'
        result = ensure_explicit_numpr_from_current_style(p, STYLES_WITH_NUMPR)
        assert '<w:numPr>' in result

    def test_style_no_numpr(self):
        p = '<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr><w:r><w:t>Text</w:t></w:r></w:p>'
        result = ensure_explicit_numpr_from_current_style(p, STYLES_WITH_NUMPR)
        assert '<w:numPr>' not in result

    def test_sectpr_unchanged(self):
        p = '<w:p><w:pPr><w:pStyle w:val="ListBullet"/><w:sectPr/></w:pPr></w:p>'
        result = ensure_explicit_numpr_from_current_style(p, STYLES_WITH_NUMPR)
        assert result == p


# ── materialize_arch_style_block ─────────────────────────────────────────────

class TestMaterializeArchStyleBlock:
    def test_no_rpr_gets_effective_rpr(self):
        """Style with no rPr should get effective rPr injected from docDefaults."""
        style = '<w:style w:type="paragraph" w:styleId="NoRpr"><w:name w:val="NoRpr"/></w:style>'
        result = materialize_arch_style_block(style, "NoRpr", STYLES_FOR_MATERIALIZE)
        assert '<w:rPr>' in result
        assert '<w:rFonts' in result
        assert '<w:sz' in result

    def test_partial_rpr_gets_missing_filled(self):
        """CSI-Part has rFonts but no sz/szCs/lang — those should be filled from docDefaults."""
        style = _extract_style_block(STYLES_FOR_MATERIALIZE, "CSI-Part")
        assert style is not None
        result = materialize_arch_style_block(style, "CSI-Part", STYLES_FOR_MATERIALIZE)
        assert '<w:sz' in result
        assert '<w:szCs' in result
        assert '<w:lang' in result
        # Original rFonts should still be there
        assert 'w:ascii="Arial"' in result

    def test_complete_rpr_unchanged(self):
        """CompleteRpr already has all FORCE tags — should not be modified."""
        style = _extract_style_block(STYLES_FOR_MATERIALIZE, "CompleteRpr")
        assert style is not None
        result = materialize_arch_style_block(style, "CompleteRpr", STYLES_FOR_MATERIALIZE)
        # All original values preserved
        assert 'w:ascii="Times"' in result
        assert 'w:val="24"' in result
        assert 'w:val="en-GB"' in result
