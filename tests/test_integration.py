"""
P2-014: Integration tests for the Phase 2 pipeline.

These tests verify cross-module interactions using minimal synthetic XML
fixtures rather than real .docx files.
"""

import re
import zipfile
from pathlib import Path

import pytest

from core.preclassifier import preclassify_paragraphs
from core.classification import detect_marker_class
from core.llm_classifier import (
    _validate_classifications,
    merge_classifications,
)
from core.style_import import (
    import_arch_styles_into_target,
    _replace_style_block_in_xml,
)
from core.stability import snapshot_stability, verify_stability
from docx_patch import patch_docx
from numbering_importer import (
    build_numbering_import_plan,
    _generate_deterministic_nsid,
)


# ── Fixtures ─────────────────────────────────────────────────────────────

_W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
_W_NS_B = b'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


def _make_bundle(paragraphs):
    """Build a minimal slim bundle dict from paragraph text list."""
    entries = []
    for i, text in enumerate(paragraphs):
        entries.append({
            "paragraph_index": i,
            "text": text,
            "marker_class": detect_marker_class(text),
            "is_all_caps": text == text.upper() and any(c.isalpha() for c in text),
        })
    return {"paragraphs": entries, "discipline": "mechanical"}


def _make_extract_dir(tmp_path):
    """Create a minimal extracted DOCX directory."""
    extract = tmp_path / "extracted"
    word_dir = extract / "word"
    word_dir.mkdir(parents=True)

    doc_xml = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document {_W_NS}><w:body>'
        f'<w:p><w:r><w:t>Hello</w:t></w:r></w:p>'
        f'<w:sectPr {_W_NS}><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>'
        f'</w:body></w:document>'
    )
    (word_dir / "document.xml").write_text(doc_xml, encoding="utf-8")

    styles_xml = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:styles {_W_NS}>'
        f'<w:style w:type="paragraph" w:styleId="ExistingStyle">'
        f'<w:name w:val="Existing"/>'
        f'<w:pPr><w:spacing w:after="120"/></w:pPr>'
        f'</w:style>'
        f'</w:styles>'
    )
    (word_dir / "styles.xml").write_text(styles_xml, encoding="utf-8")

    # Header and footer
    (word_dir / "header1.xml").write_text(
        f'<w:hdr {_W_NS}><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:hdr>',
        encoding="utf-8",
    )
    (word_dir / "footer1.xml").write_text(
        f'<w:ftr {_W_NS}><w:p><w:r><w:t>Footer</w:t></w:r></w:p></w:ftr>',
        encoding="utf-8",
    )

    return extract


def _make_minimal_docx(path):
    """Create a minimal valid DOCX for patch testing."""
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("word/document.xml", f'<w:document {_W_NS}/>'.encode())
        zf.writestr("word/header1.xml", f'<w:hdr {_W_NS}><w:p/></w:hdr>'.encode())
        zf.writestr("word/footer1.xml", f'<w:ftr {_W_NS}><w:p/></w:ftr>'.encode())


# ── 1. Preclassifier accuracy ────────────────────────────────────────────


class TestPreclassifierAccuracy:
    """Feed known CSI text through preclassifier, verify correct role assignment."""

    def test_full_csi_document(self):
        paragraphs = [
            "SECTION 23 05 13",       # SectionID
            "COMMON MOTOR REQUIREMENTS",  # SectionTitle (all-caps after SectionID)
            "",                        # empty
            "Some boilerplate text",
            "PART 1 GENERAL",         # PART
            "1.01 SUMMARY",           # ARTICLE
            "A. This section...",      # PARAGRAPH
            "1. First item",          # SUBPARAGRAPH
            "a. Sub item",            # SUBSUBPARAGRAPH
            "B. Another paragraph",   # PARAGRAPH
            "2.01 PRODUCTS",          # ARTICLE
        ]
        bundle = _make_bundle(paragraphs)
        all_roles = ["SectionID", "SectionTitle", "PART", "ARTICLE",
                      "PARAGRAPH", "SUBPARAGRAPH", "SUBSUBPARAGRAPH"]

        preclassified, ambiguous = preclassify_paragraphs(bundle, all_roles)

        assert preclassified.get(0) == "SectionID"
        assert preclassified.get(1) == "SectionTitle"
        assert preclassified.get(4) == "PART"
        assert preclassified.get(5) == "ARTICLE"
        assert preclassified.get(6) == "PARAGRAPH"
        assert preclassified.get(7) == "SUBPARAGRAPH"
        assert preclassified.get(9) == "PARAGRAPH"
        assert preclassified.get(10) == "ARTICLE"

    def test_preclassifier_skips_unavailable_roles(self):
        paragraphs = [
            "SECTION 23 05 13",
            "COMMON MOTOR REQUIREMENTS",
        ]
        bundle = _make_bundle(paragraphs)
        # Only PART is available
        preclassified, ambiguous = preclassify_paragraphs(bundle, ["PART"])
        assert 0 not in preclassified  # SectionID not available
        assert 1 not in preclassified  # SectionTitle not available


# ── 2. Ambiguous-only LLM shaping ────────────────────────────────────────


class TestAmbiguousOnlyShaping:
    """Verify LLM only receives unresolved paragraphs."""

    def test_preclassified_excluded_from_ambiguous(self):
        paragraphs = [
            "SECTION 23 05 13",       # preclassified
            "Some text",              # ambiguous
            "PART 1 GENERAL",         # preclassified
            "Another text",           # ambiguous
        ]
        bundle = _make_bundle(paragraphs)
        all_roles = ["SectionID", "SectionTitle", "PART", "ARTICLE",
                      "PARAGRAPH", "SUBPARAGRAPH", "SUBSUBPARAGRAPH"]

        preclassified, ambiguous = preclassify_paragraphs(bundle, all_roles)

        # Preclassified paragraphs should not be in ambiguous
        for idx in preclassified:
            assert idx not in ambiguous

        # Ambiguous paragraphs should be in ambiguous
        assert 1 in ambiguous or 3 in ambiguous


# ── 3. Post-LLM validation ───────────────────────────────────────────────


class TestPostLLMValidation:
    """Verify malformed LLM output triggers appropriate errors/warnings."""

    _ALL_ROLES = ["SectionID", "SectionTitle", "PART", "ARTICLE",
                  "PARAGRAPH", "SUBPARAGRAPH", "SUBSUBPARAGRAPH"]

    def test_unknown_role_rejected(self):
        raw = {
            "classifications": [
                {"paragraph_index": 0, "csi_role": "INVALID_ROLE"},
            ]
        }
        result = _validate_classifications(raw, self._ALL_ROLES, total_paragraphs=10)
        assert len(result["classifications"]) == 0

    def test_out_of_range_rejected(self):
        raw = {
            "classifications": [
                {"paragraph_index": 999, "csi_role": "PART"},
            ]
        }
        result = _validate_classifications(raw, self._ALL_ROLES, total_paragraphs=10)
        assert len(result["classifications"]) == 0

    def test_duplicates_keep_last(self):
        raw = {
            "classifications": [
                {"paragraph_index": 0, "csi_role": "PART"},
                {"paragraph_index": 0, "csi_role": "ARTICLE"},
            ]
        }
        result = _validate_classifications(raw, self._ALL_ROLES, total_paragraphs=10)
        assert len(result["classifications"]) == 1
        assert result["classifications"][0]["csi_role"] == "ARTICLE"


# ── 4. Style overwrite sync ──────────────────────────────────────────────


class TestStyleOverwriteSync:
    """Verify conflicting styles are replaced in template_sync mode."""

    def test_overwrite_replaces_existing(self, tmp_path):
        extract = _make_extract_dir(tmp_path)

        arch_styles_xml = (
            f'<w:styles {_W_NS}>'
            f'<w:style w:type="paragraph" w:styleId="ExistingStyle">'
            f'<w:name w:val="Existing"/>'
            f'<w:pPr><w:spacing w:after="240"/></w:pPr>'
            f'<w:rPr><w:b/></w:rPr>'
            f'</w:style>'
            f'</w:styles>'
        )

        log = []
        import_arch_styles_into_target(
            target_extract_dir=extract,
            arch_styles_xml=arch_styles_xml,
            needed_style_ids=["ExistingStyle"],
            log=log,
            overwrite_existing=True,
        )

        result = (extract / "word" / "styles.xml").read_text(encoding="utf-8")
        assert 'w:after="240"' in result
        assert '<w:b/>' in result


# ── 5. body_only invariants ──────────────────────────────────────────────


class TestBodyOnlyInvariants:
    """Verify headers/footers/sectPr unchanged in body_only mode."""

    def test_stability_passes_when_unchanged(self, tmp_path):
        extract = _make_extract_dir(tmp_path)
        snap = snapshot_stability(extract)
        # No modifications
        verify_stability(extract, snap, sync_mode="body_only")

    def test_stability_fails_when_header_changed(self, tmp_path):
        extract = _make_extract_dir(tmp_path)
        snap = snapshot_stability(extract)

        # Modify header
        (extract / "word" / "header1.xml").write_text("CHANGED", encoding="utf-8")

        with pytest.raises(ValueError, match="Header/footer stability"):
            verify_stability(extract, snap, sync_mode="body_only")

    def test_stability_skipped_in_template_sync(self, tmp_path):
        extract = _make_extract_dir(tmp_path)
        snap = snapshot_stability(extract)

        # Modify header
        (extract / "word" / "header1.xml").write_text("CHANGED", encoding="utf-8")

        # Should not raise in template_sync mode
        verify_stability(extract, snap, sync_mode="template_sync")


# ── 6. Numbering stability ──────────────────────────────────────────────


class TestNumberingStability:
    """Verify repeated runs produce identical numbering imports."""

    def test_deterministic_ids_across_runs(self):
        nsid1 = _generate_deterministic_nsid("abstractNum", 5)
        nsid2 = _generate_deterministic_nsid("abstractNum", 5)
        assert nsid1 == nsid2

    def test_import_plan_deterministic(self):
        styles_xml = (
            f'<w:styles {_W_NS}>'
            '<w:style w:type="paragraph" w:styleId="S1">'
            '<w:pPr><w:numPr><w:ilvl w:val="0"/>'
            '<w:numId w:val="1"/></w:numPr></w:pPr>'
            '</w:style></w:styles>'
        )
        target_xml = (
            f'<w:numbering {_W_NS}></w:numbering>'
        )
        arch_abstract_xml = (
            '<w:abstractNum w:abstractNumId="0">'
            '<w:nsid w:val="AABB0011"/>'
            '<w:lvl w:ilvl="0"><w:start w:val="1"/>'
            '<w:numFmt w:val="upperLetter"/></w:lvl>'
            '</w:abstractNum>'
        )
        registry = {
            "numbering": {
                "abstract_nums": [{"abstractNumId": 0, "xml": arch_abstract_xml}],
                "nums": [{"numId": 1, "abstractNumId": 0,
                          "xml": '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'}],
            }
        }

        plan1 = build_numbering_import_plan(registry, styles_xml, target_xml, ["S1"])
        plan2 = build_numbering_import_plan(registry, styles_xml, target_xml, ["S1"])

        assert plan1["abstract_nums_to_import"] == plan2["abstract_nums_to_import"]
        assert plan1["nums_to_import"] == plan2["nums_to_import"]


# ── 7. Patch target allowlist by mode ────────────────────────────────────


class TestPatchTargetAllowlist:
    """Verify body_only rejects header patches, template_sync allows them."""

    def test_body_only_rejects_header(self, tmp_path):
        src = tmp_path / "input.docx"
        out = tmp_path / "output.docx"
        _make_minimal_docx(src)

        with pytest.raises(RuntimeError, match="Forbidden patch target"):
            patch_docx(
                src_docx=src,
                out_docx=out,
                replacements={"word/header1.xml": f'<w:hdr {_W_NS}/>'.encode()},
                sync_mode="body_only",
            )

    def test_template_sync_allows_header(self, tmp_path):
        src = tmp_path / "input.docx"
        out = tmp_path / "output.docx"
        _make_minimal_docx(src)

        patch_docx(
            src_docx=src,
            out_docx=out,
            replacements={"word/header1.xml": f'<w:hdr {_W_NS}/>'.encode()},
            sync_mode="template_sync",
        )
        assert out.exists()


# ── 8. Classification merge ──────────────────────────────────────────────


class TestClassificationMerge:
    """Verify preclassified + LLM merge works correctly."""

    def test_merge_combines_both_sources(self):
        preclassified = {0: "SectionID", 2: "PART"}
        llm_classified = {
            "classifications": [
                {"paragraph_index": 1, "csi_role": "SectionTitle"},
                {"paragraph_index": 3, "csi_role": "ARTICLE"},
            ]
        }
        result = merge_classifications(preclassified, llm_classified, total_paragraphs=5)
        classes = {c["paragraph_index"]: c["csi_role"] for c in result["classifications"]}
        assert classes[0] == "SectionID"
        assert classes[1] == "SectionTitle"
        assert classes[2] == "PART"
        assert classes[3] == "ARTICLE"
