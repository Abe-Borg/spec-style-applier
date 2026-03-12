"""Tests for numbering_importer.py — fail-fast validation and no font injection."""

import pytest
import re
from pathlib import Path

from numbering_importer import (
    build_numbering_import_plan,
    inject_numbering_into_xml,
    import_numbering,
    extract_used_num_ids_from_styles,
)


# ---------------------------------------------------------------------------
# Helpers — minimal XML fixtures
# ---------------------------------------------------------------------------

def _make_arch_styles_xml(styles):
    """Build a minimal styles.xml from a list of (styleId, numId_or_None)."""
    blocks = []
    for sid, num_id in styles:
        numpr = ""
        if num_id is not None:
            numpr = f'<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="{num_id}"/></w:numPr></w:pPr>'
        blocks.append(
            f'<w:style w:type="paragraph" w:styleId="{sid}">'
            f'{numpr}'
            f'</w:style>'
        )
    return f'<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">{"".join(blocks)}</w:styles>'


def _make_registry(abstract_nums=None, nums=None):
    """Build a minimal arch_template_registry with numbering data."""
    reg = {}
    if abstract_nums is not None or nums is not None:
        reg["numbering"] = {
            "abstract_nums": abstract_nums or [],
            "nums": nums or [],
        }
    return reg


def _make_abstract_num(abstract_num_id, rpr_xml=""):
    """Build a minimal abstractNum dict entry."""
    xml = (
        f'<w:abstractNum w:abstractNumId="{abstract_num_id}">'
        f'<w:nsid w:val="AABB0011"/>'
        f'<w:lvl w:ilvl="0">{rpr_xml}<w:start w:val="1"/>'
        f'<w:numFmt w:val="decimal"/></w:lvl>'
        f'</w:abstractNum>'
    )
    return {"abstractNumId": abstract_num_id, "xml": xml}


def _make_num(num_id, abstract_num_id):
    """Build a minimal num dict entry."""
    xml = (
        f'<w:num w:numId="{num_id}">'
        f'<w:abstractNumId w:val="{abstract_num_id}"/>'
        f'</w:num>'
    )
    return {"numId": num_id, "abstractNumId": abstract_num_id, "xml": xml}


MINIMAL_TARGET_NUMBERING = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:abstractNum w:abstractNumId="0"><w:nsid w:val="00000001"/>'
    '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/></w:lvl>'
    '</w:abstractNum>'
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
    '</w:numbering>'
)


# ---------------------------------------------------------------------------
# Test: no font injection in inject_numbering_into_xml
# ---------------------------------------------------------------------------

class TestNoFontInjection:
    def test_inject_numbering_preserves_architect_rpr(self):
        """Verify that inject_numbering_into_xml does NOT modify <w:rPr> blocks."""
        rpr = '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr>'
        abstract_xml = (
            f'<w:abstractNum w:abstractNumId="10">'
            f'<w:nsid w:val="AABB0011"/>'
            f'<w:lvl w:ilvl="0">{rpr}<w:start w:val="1"/>'
            f'<w:numFmt w:val="decimal"/></w:lvl>'
            f'</w:abstractNum>'
        )
        num_xml = (
            '<w:num w:numId="10">'
            '<w:abstractNumId w:val="10"/>'
            '</w:num>'
        )

        result = inject_numbering_into_xml(
            MINIMAL_TARGET_NUMBERING,
            [{"xml": abstract_xml}],
            [{"xml": num_xml}],
        )

        # The original rPr must appear verbatim — no Arial injection
        assert rpr in result
        assert "Arial" not in result

    def test_inject_numbering_preserves_rpr_without_fonts(self):
        """rPr blocks without rFonts should NOT get Arial injected."""
        rpr = '<w:rPr><w:b/></w:rPr>'
        abstract_xml = (
            f'<w:abstractNum w:abstractNumId="10">'
            f'<w:nsid w:val="AABB0011"/>'
            f'<w:lvl w:ilvl="0">{rpr}<w:start w:val="1"/>'
            f'<w:numFmt w:val="decimal"/></w:lvl>'
            f'</w:abstractNum>'
        )

        result = inject_numbering_into_xml(
            MINIMAL_TARGET_NUMBERING,
            [{"xml": abstract_xml}],
            [],
        )

        assert rpr in result
        assert "Arial" not in result


# ---------------------------------------------------------------------------
# Test: fail-fast on missing numId in registry
# ---------------------------------------------------------------------------

class TestBuildPlanFailFast:
    def test_raises_on_missing_num_id(self):
        """build_numbering_import_plan raises when a referenced numId is missing."""
        styles_xml = _make_arch_styles_xml([("CSILevel1", 99)])
        registry = _make_registry(
            abstract_nums=[_make_abstract_num(5)],
            nums=[_make_num(2, 5)],  # has numId=2 but style references 99
        )

        with pytest.raises(ValueError, match="missing required numId"):
            build_numbering_import_plan(
                registry,
                styles_xml,
                MINIMAL_TARGET_NUMBERING,
                ["CSILevel1"],
            )

    def test_raises_on_missing_abstract_num(self):
        """build_numbering_import_plan raises when a referenced abstractNumId is missing."""
        styles_xml = _make_arch_styles_xml([("CSILevel1", 2)])
        registry = _make_registry(
            abstract_nums=[],  # no abstractNums at all
            nums=[_make_num(2, 5)],  # numId=2 references abstractNum 5
        )

        with pytest.raises(ValueError, match="missing required abstractNum"):
            build_numbering_import_plan(
                registry,
                styles_xml,
                MINIMAL_TARGET_NUMBERING,
                ["CSILevel1"],
            )

    def test_succeeds_when_no_numbering_needed(self):
        """Empty plan returned when styles don't reference numbering."""
        styles_xml = _make_arch_styles_xml([("PlainStyle", None)])
        registry = _make_registry(
            abstract_nums=[_make_abstract_num(5)],
            nums=[_make_num(2, 5)],
        )

        plan = build_numbering_import_plan(
            registry,
            styles_xml,
            MINIMAL_TARGET_NUMBERING,
            ["PlainStyle"],
        )
        assert plan["abstract_nums_to_import"] == []
        assert plan["nums_to_import"] == []
        assert plan["style_numid_remap"] == {}


# ---------------------------------------------------------------------------
# Test: import_numbering fail-fast on missing numbering.xml
# ---------------------------------------------------------------------------

class TestImportNumberingFailFast:
    def test_raises_when_no_numbering_xml_but_needed(self, tmp_path):
        """import_numbering raises when target lacks numbering.xml but styles need it."""
        # Create a target dir without numbering.xml
        word_dir = tmp_path / "word"
        word_dir.mkdir()

        styles_xml = _make_arch_styles_xml([("CSILevel1", 2)])
        registry = _make_registry(
            abstract_nums=[_make_abstract_num(5)],
            nums=[_make_num(2, 5)],
        )
        log = []

        with pytest.raises(ValueError, match="no numbering.xml"):
            import_numbering(
                target_extract_dir=tmp_path,
                arch_template_registry=registry,
                arch_styles_xml=styles_xml,
                style_ids_to_import=["CSILevel1"],
                log=log,
            )

    def test_succeeds_when_no_numbering_xml_not_needed(self, tmp_path):
        """import_numbering returns {} when no styles need numbering, even without numbering.xml."""
        word_dir = tmp_path / "word"
        word_dir.mkdir()

        styles_xml = _make_arch_styles_xml([("PlainStyle", None)])
        registry = _make_registry(
            abstract_nums=[_make_abstract_num(5)],
            nums=[_make_num(2, 5)],
        )
        log = []

        result = import_numbering(
            target_extract_dir=tmp_path,
            arch_template_registry=registry,
            arch_styles_xml=styles_xml,
            style_ids_to_import=["PlainStyle"],
            log=log,
        )
        assert result == {}

    def test_raises_when_registry_has_no_numbering_but_needed(self, tmp_path):
        """import_numbering raises when registry has no numbering data but styles need it."""
        word_dir = tmp_path / "word"
        word_dir.mkdir()
        (word_dir / "numbering.xml").write_text(MINIMAL_TARGET_NUMBERING, encoding="utf-8")

        styles_xml = _make_arch_styles_xml([("CSILevel1", 2)])
        registry = {}  # no numbering key at all
        log = []

        with pytest.raises(ValueError, match="no numbering data"):
            import_numbering(
                target_extract_dir=tmp_path,
                arch_template_registry=registry,
                arch_styles_xml=styles_xml,
                style_ids_to_import=["CSILevel1"],
                log=log,
            )


# ---------------------------------------------------------------------------
# Test: happy path — full remap with no font injection
# ---------------------------------------------------------------------------

class TestHappyPath:
    def test_import_happy_path_no_font_injection(self, tmp_path):
        """Full remap works correctly and no font XML is injected."""
        word_dir = tmp_path / "word"
        word_dir.mkdir()
        (word_dir / "numbering.xml").write_text(MINIMAL_TARGET_NUMBERING, encoding="utf-8")

        rpr = '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="18"/></w:rPr>'
        styles_xml = _make_arch_styles_xml([("CSILevel1", 2)])
        registry = _make_registry(
            abstract_nums=[_make_abstract_num(5, rpr_xml=rpr)],
            nums=[_make_num(2, 5)],
        )
        log = []

        remap = import_numbering(
            target_extract_dir=tmp_path,
            arch_template_registry=registry,
            arch_styles_xml=styles_xml,
            style_ids_to_import=["CSILevel1"],
            log=log,
        )

        # Verify remap was produced
        assert "CSILevel1" in remap
        assert remap["CSILevel1"]["old_numId"] == 2
        assert remap["CSILevel1"]["new_numId"] > 1  # remapped to avoid collision

        # Verify the written numbering.xml has no Arial injection
        result_xml = (word_dir / "numbering.xml").read_text(encoding="utf-8")
        assert "Arial" not in result_xml
        # Original font should be preserved
        assert "Calibri" in result_xml
