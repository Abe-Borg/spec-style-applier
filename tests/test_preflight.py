"""Tests for core.registry.preflight_validate_registries()."""

import pytest
from core.registry import preflight_validate_registries


# ---------------------------------------------------------------------------
# Helpers to build minimal valid registries
# ---------------------------------------------------------------------------

def _minimal_style_registry():
    """role -> styleId mapping (output of load_arch_style_registry)."""
    return {"PART": "CSIPart", "ARTICLE": "CSIArticle"}


def _minimal_template_registry():
    """Minimal valid arch_template_registry dict."""
    return {
        "styles": {
            "style_defs": [
                {
                    "style_id": "CSIPart",
                    "type": "paragraph",
                    "name": "CSI Part",
                },
                {
                    "style_id": "CSIArticle",
                    "type": "paragraph",
                    "name": "CSI Article",
                },
            ]
        },
        "page_layout": {
            "default_section": {
                "sectPr": "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1800\" w:right=\"1080\" w:bottom=\"1440\" w:left=\"2160\" w:header=\"900\" w:footer=\"720\"/></w:sectPr>"
            },
            "section_chain": [],
        },
    }


# ---------------------------------------------------------------------------
# 1. Valid registries → no errors
# ---------------------------------------------------------------------------

def test_valid_registries_no_errors():
    errors = preflight_validate_registries(
        _minimal_style_registry(), _minimal_template_registry()
    )
    assert errors == []


# ---------------------------------------------------------------------------
# 2. Wrong section types
# ---------------------------------------------------------------------------

def test_wrong_section_types():
    tmpl = _minimal_template_registry()
    tmpl["theme"] = "not-a-dict"
    errors = preflight_validate_registries(_minimal_style_registry(), tmpl)
    assert any("'theme' must be dict" in e for e in errors)


# ---------------------------------------------------------------------------
# 3. style_defs not a list
# ---------------------------------------------------------------------------

def test_style_defs_not_list():
    tmpl = {
        "styles": {"style_defs": {"bad": True}},
        "page_layout": {"default_section": {"sectPr": "<w:sectPr/>"}},
    }
    errors = preflight_validate_registries({}, tmpl)
    assert any("style_defs must be a list" in e for e in errors)


# ---------------------------------------------------------------------------
# 4. Duplicate style_id
# ---------------------------------------------------------------------------

def test_duplicate_style_ids():
    tmpl = {
        "styles": {
            "style_defs": [
                {"style_id": "Dup", "type": "paragraph", "name": "A"},
                {"style_id": "Dup", "type": "paragraph", "name": "B"},
            ]
        },
        "page_layout": {"default_section": {"sectPr": "<w:sectPr/>"}},
    }
    errors = preflight_validate_registries({}, tmpl)
    assert any("Duplicate style_id 'Dup'" in e for e in errors)


# ---------------------------------------------------------------------------
# 5. Malformed XML fragment in style_def
# ---------------------------------------------------------------------------

def test_malformed_xml_fragment():
    tmpl = {
        "styles": {
            "style_defs": [
                {
                    "style_id": "Bad",
                    "type": "paragraph",
                    "name": "Bad Style",
                    "pPr": "<w:pPr><w:spacing w:after='200'/>",  # missing close
                },
            ]
        },
        "page_layout": {"default_section": {"sectPr": "<w:sectPr/>"}},
    }
    errors = preflight_validate_registries({}, tmpl)
    assert any("w:pPr" in e and "malformed" in e for e in errors)


def test_self_closing_xml_fragment_is_valid():
    tmpl = {
        "styles": {
            "style_defs": [
                {
                    "style_id": "Good",
                    "type": "paragraph",
                    "name": "Good Style",
                    "pPr": '<w:pPr w:val="x"/>',
                },
            ]
        },
        "page_layout": {"default_section": {"sectPr": "<w:sectPr/>"}},
    }
    errors = preflight_validate_registries({}, tmpl)
    assert errors == []


# ---------------------------------------------------------------------------
# 6. Invalid compat_xml
# ---------------------------------------------------------------------------

def test_invalid_compat_xml():
    tmpl = _minimal_template_registry()
    tmpl["settings"] = {"compat": {"compat_xml": "<w:compat><w:useFELayout/>"}}
    errors = preflight_validate_registries(_minimal_style_registry(), tmpl)
    assert any("compat_xml" in e and "malformed" in e for e in errors)


def test_valid_compat_xml():
    tmpl = _minimal_template_registry()
    tmpl["settings"] = {
        "compat": {"compat_xml": "<w:compat><w:useFELayout/></w:compat>"}
    }
    errors = preflight_validate_registries(_minimal_style_registry(), tmpl)
    assert errors == []


# ---------------------------------------------------------------------------
# 7. Style ID missing from template
# ---------------------------------------------------------------------------

def test_style_id_missing_from_template():
    style_reg = {"PART": "CSIPart", "ARTICLE": "MissingStyle"}
    tmpl = {
        "styles": {
            "style_defs": [
                {"style_id": "CSIPart", "type": "paragraph", "name": "Part"},
            ]
        },
        "page_layout": {"default_section": {"sectPr": "<w:sectPr/>"}},
    }
    errors = preflight_validate_registries(style_reg, tmpl)
    assert any("MissingStyle" in e and "ARTICLE" in e for e in errors)


# ---------------------------------------------------------------------------
# 8. Numbering abstractNumId ref missing
# ---------------------------------------------------------------------------

def test_numbering_abstract_ref_missing():
    tmpl = _minimal_template_registry()
    tmpl["numbering"] = {
        "abstract_nums": [{"abstractNumId": 1, "xml": "<w:abstractNum/>"}],
        "nums": [{"numId": 5, "abstractNumId": 99, "xml": "<w:num/>"}],
    }
    errors = preflight_validate_registries(_minimal_style_registry(), tmpl)
    assert any("numId=5" in e and "abstractNumId=99" in e for e in errors)


def test_numbering_consistent():
    tmpl = _minimal_template_registry()
    tmpl["numbering"] = {
        "abstract_nums": [{"abstractNumId": 1, "xml": "<w:abstractNum/>"}],
        "nums": [{"numId": 5, "abstractNumId": 1, "xml": "<w:num/>"}],
    }
    errors = preflight_validate_registries(_minimal_style_registry(), tmpl)
    assert errors == []


# ---------------------------------------------------------------------------
# 9. Empty style_defs list is valid
# ---------------------------------------------------------------------------

def test_empty_style_defs_is_ok():
    tmpl = {
        "styles": {"style_defs": []},
        "page_layout": {"default_section": {"sectPr": "<w:sectPr/>"}},
    }
    errors = preflight_validate_registries({}, tmpl)
    assert errors == []


# ---------------------------------------------------------------------------
# 10. Missing optional sections is valid
# ---------------------------------------------------------------------------

def test_missing_page_layout_is_error():
    tmpl = {"styles": _minimal_template_registry()["styles"]}
    errors = preflight_validate_registries({}, tmpl)
    assert any("missing page_layout" in e for e in errors)


# ---------------------------------------------------------------------------
# Additional edge cases
# ---------------------------------------------------------------------------

def test_style_def_not_dict():
    tmpl = {
        "styles": {"style_defs": ["not-a-dict"]},
        "page_layout": {"default_section": {"sectPr": "<w:sectPr/>"}},
    }
    errors = preflight_validate_registries({}, tmpl)
    assert any("must be a dict" in e for e in errors)


def test_malformed_theme_xml():
    tmpl = _minimal_template_registry()
    tmpl["theme"] = {"theme1_xml": "<a:theme><a:themeElements/>"}
    errors = preflight_validate_registries(_minimal_style_registry(), tmpl)
    assert any("theme1_xml" in e and "malformed" in e for e in errors)


def test_malformed_font_table_xml():
    tmpl = _minimal_template_registry()
    tmpl["fonts"] = {"font_table_xml": "<w:fonts>incomplete"}
    errors = preflight_validate_registries(_minimal_style_registry(), tmpl)
    assert any("font_table_xml" in e and "malformed" in e for e in errors)


def test_numbering_nums_not_list():
    tmpl = _minimal_template_registry()
    tmpl["numbering"] = {"abstract_nums": [], "nums": "bad"}
    errors = preflight_validate_registries(_minimal_style_registry(), tmpl)
    assert any("numbering.nums must be a list" in e for e in errors)


def test_collects_all_errors_at_once():
    """Verify that multiple errors are collected in a single pass."""
    style_reg = {"PART": "MissingA", "ARTICLE": "MissingB"}
    tmpl = {
        "styles": "not-a-dict",  # wrong type
        "theme": 42,             # wrong type
    }
    errors = preflight_validate_registries(style_reg, tmpl)
    # At minimum: styles wrong type + theme wrong type + 2 missing cross-refs
    assert len(errors) >= 3
