"""Tests for core.preclassifier — deterministic CSI role assignment."""

import pytest
from core.preclassifier import preclassify_paragraphs


ALL_ROLES = [
    "SectionID", "SectionTitle", "PART", "ARTICLE",
    "PARAGRAPH", "SUBPARAGRAPH", "SUBSUBPARAGRAPH",
]


def _bundle(texts):
    """Build a minimal slim bundle from a list of text strings."""
    paragraphs = [
        {"paragraph_index": i, "text": t} for i, t in enumerate(texts)
    ]
    return {
        "paragraphs": paragraphs,
        "available_roles": ALL_ROLES,
    }


class TestPreclassifyBasicPatterns:
    """Each obvious CSI pattern should be detected."""

    def test_section_id(self):
        pre, amb = preclassify_paragraphs(
            _bundle(["SECTION 23 05 13"]), ALL_ROLES
        )
        assert pre[0] == "SectionID"

    def test_part(self):
        pre, amb = preclassify_paragraphs(
            _bundle(["PART 1 GENERAL"]), ALL_ROLES
        )
        assert pre[0] == "PART"

    def test_article(self):
        pre, amb = preclassify_paragraphs(
            _bundle(["1.01 SUMMARY"]), ALL_ROLES
        )
        assert pre[0] == "ARTICLE"

    def test_paragraph(self):
        pre, amb = preclassify_paragraphs(
            _bundle(["A. Provide valves"]), ALL_ROLES
        )
        assert pre[0] == "PARAGRAPH"

    def test_subparagraph_with_context(self):
        """1. is SUBPARAGRAPH when preceded by a PARAGRAPH."""
        pre, amb = preclassify_paragraphs(
            _bundle(["A. Parent paragraph", "1. Child item"]), ALL_ROLES
        )
        assert pre[0] == "PARAGRAPH"
        assert pre[1] == "SUBPARAGRAPH"

    def test_subsubparagraph_with_context(self):
        """a. is SUBSUBPARAGRAPH when preceded by a SUBPARAGRAPH."""
        pre, amb = preclassify_paragraphs(
            _bundle([
                "A. Parent paragraph",
                "1. Sub item",
                "a. Sub-sub item",
            ]),
            ALL_ROLES
        )
        assert pre[2] == "SUBSUBPARAGRAPH"


class TestPreclassifySectionTitle:
    """All-caps text after SectionID is SectionTitle."""

    def test_section_title_after_section_id(self):
        pre, amb = preclassify_paragraphs(
            _bundle([
                "SECTION 23 05 13",
                "COMMON MOTOR REQUIREMENTS FOR HVAC EQUIPMENT",
            ]),
            ALL_ROLES
        )
        assert pre[0] == "SectionID"
        assert pre[1] == "SectionTitle"

    def test_mixed_case_after_section_id_not_title(self):
        pre, amb = preclassify_paragraphs(
            _bundle([
                "SECTION 23 05 13",
                "Some mixed case text here",
            ]),
            ALL_ROLES
        )
        assert pre[0] == "SectionID"
        assert 1 in amb  # not preclassified


class TestPreclassifyContextAware:
    """Context-aware disambiguation."""

    def test_subparagraph_without_context_is_ambiguous(self):
        """1. without a preceding PARAGRAPH should be ambiguous."""
        pre, amb = preclassify_paragraphs(
            _bundle(["1. Some item"]), ALL_ROLES
        )
        assert 0 not in pre
        assert 0 in amb

    def test_subsubparagraph_without_context_is_ambiguous(self):
        """a. without a preceding SUBPARAGRAPH should be ambiguous."""
        pre, amb = preclassify_paragraphs(
            _bundle(["a. Some detail"]), ALL_ROLES
        )
        assert 0 not in pre
        assert 0 in amb


class TestPreclassifyRoleAvailability:
    """If a role is not in available_roles, don't preclassify it."""

    def test_missing_role_leaves_ambiguous(self):
        # Only PART and ARTICLE are available — no SectionID
        pre, amb = preclassify_paragraphs(
            _bundle(["SECTION 23 05 13"]),
            ["PART", "ARTICLE"],
        )
        assert 0 not in pre
        assert 0 in amb

    def test_part_still_works_when_available(self):
        pre, amb = preclassify_paragraphs(
            _bundle(["PART 2 PRODUCTS"]),
            ["PART"],
        )
        assert pre[0] == "PART"


class TestPreclassifyBypass:
    """force_llm_all=True skips all pre-classification."""

    def test_force_llm_all(self):
        pre, amb = preclassify_paragraphs(
            _bundle(["PART 1 GENERAL", "1.01 SUMMARY"]),
            ALL_ROLES,
            force_llm_all=True,
        )
        assert len(pre) == 0
        assert len(amb) == 2


class TestPreclassifyFullDocument:
    """Simulate a small CSI document."""

    def test_typical_structure(self):
        texts = [
            "SECTION 23 05 13",                    # 0: SectionID
            "COMMON MOTOR REQUIREMENTS",            # 1: SectionTitle
            "PART 1 GENERAL",                       # 2: PART
            "1.01 SUMMARY",                         # 3: ARTICLE
            "A. This Section includes motor req.",   # 4: PARAGRAPH
            "B. Related Sections:",                  # 5: PARAGRAPH
            "1. Section 23 09 00",                  # 6: SUBPARAGRAPH
            "2. Section 26 05 00",                  # 7: SUBPARAGRAPH
            "a. Low voltage motors",                # 8: SUBSUBPARAGRAPH
            "PART 2 PRODUCTS",                      # 9: PART
            "2.01 MOTORS",                          # 10: ARTICLE
        ]
        pre, amb = preclassify_paragraphs(_bundle(texts), ALL_ROLES)

        assert pre[0] == "SectionID"
        assert pre[1] == "SectionTitle"
        assert pre[2] == "PART"
        assert pre[3] == "ARTICLE"
        assert pre[4] == "PARAGRAPH"
        assert pre[5] == "PARAGRAPH"
        assert pre[6] == "SUBPARAGRAPH"
        assert pre[7] == "SUBPARAGRAPH"
        assert pre[8] == "SUBSUBPARAGRAPH"
        assert pre[9] == "PART"
        assert pre[10] == "ARTICLE"
        assert len(amb) == 0
