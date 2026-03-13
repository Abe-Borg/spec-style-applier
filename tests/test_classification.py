"""Tests for core.classification — boilerplate filtering, marker detection, and bundle building."""

import pytest
from core.classification import strip_boilerplate_with_report, detect_marker_class


class TestEndOfSectionFiltering:
    """END OF SECTION lines must be stripped and tagged 'end_of_section'."""

    def test_plain_end_of_section(self):
        cleaned, tags = strip_boilerplate_with_report("END OF SECTION")
        assert cleaned == ""
        assert "end_of_section" in tags

    def test_end_of_section_with_number(self):
        cleaned, tags = strip_boilerplate_with_report("END OF SECTION 211300")
        assert cleaned == ""
        assert "end_of_section" in tags

    def test_end_of_section_with_spaced_number(self):
        cleaned, tags = strip_boilerplate_with_report("END OF SECTION 23 05 13")
        assert cleaned == ""
        assert "end_of_section" in tags

    def test_lowercase(self):
        cleaned, tags = strip_boilerplate_with_report("end of section")
        assert cleaned == ""
        assert "end_of_section" in tags

    def test_mixed_case(self):
        cleaned, tags = strip_boilerplate_with_report("End Of Section")
        assert cleaned == ""
        assert "end_of_section" in tags

    def test_leading_trailing_whitespace(self):
        cleaned, tags = strip_boilerplate_with_report("  END OF SECTION  ")
        assert cleaned == ""
        assert "end_of_section" in tags

    def test_real_content_unchanged(self):
        text = "A. Provide valves as specified."
        cleaned, tags = strip_boilerplate_with_report(text)
        assert cleaned == text
        assert tags == []

    def test_article_heading_unchanged(self):
        text = "1.01 SUMMARY"
        cleaned, tags = strip_boilerplate_with_report(text)
        assert cleaned == text
        assert tags == []

    def test_part_heading_unchanged(self):
        text = "PART 1 GENERAL"
        cleaned, tags = strip_boilerplate_with_report(text)
        assert cleaned == text
        assert tags == []


class TestEndOfSectionEdgeCases:
    """Additional END OF SECTION edge cases."""

    def test_trailing_period(self):
        cleaned, tags = strip_boilerplate_with_report("END OF SECTION.")
        assert cleaned == ""
        assert "end_of_section" in tags

    def test_with_discipline_suffix(self):
        cleaned, tags = strip_boilerplate_with_report("END OF SECTION - MECHANICAL")
        assert cleaned == ""
        assert "end_of_section" in tags

    def test_not_matched_when_embedded_in_sentence(self):
        """Regex anchors at start of line — embedded phrase should NOT match."""
        text = "THE END OF SECTION DESCRIBES THE SCOPE"
        cleaned, tags = strip_boilerplate_with_report(text)
        assert cleaned == text
        assert tags == []

    def test_dashes_before_prevent_match(self):
        """Leading dashes prevent the '^\\s*END' anchor from matching."""
        text = "--- END OF SECTION ---"
        cleaned, tags = strip_boilerplate_with_report(text)
        assert cleaned == text
        assert tags == []

    def test_empty_string_no_tags(self):
        cleaned, tags = strip_boilerplate_with_report("")
        assert cleaned == ""
        assert tags == []


# ---------------------------------------------------------------------------
# P2-003: Marker class detection
# ---------------------------------------------------------------------------

class TestDetectMarkerClass:
    """detect_marker_class assigns obvious CSI markers."""

    def test_section_id(self):
        assert detect_marker_class("SECTION 23 05 13") == "SectionID"

    def test_section_id_lowercase(self):
        assert detect_marker_class("section 23 05 13") == "SectionID"

    def test_part_1(self):
        assert detect_marker_class("PART 1 GENERAL") == "PART"

    def test_part_2(self):
        assert detect_marker_class("PART 2 PRODUCTS") == "PART"

    def test_part_3(self):
        assert detect_marker_class("PART 3 EXECUTION") == "PART"

    def test_article(self):
        assert detect_marker_class("1.01 SUMMARY") == "ARTICLE"

    def test_article_2(self):
        assert detect_marker_class("2.03 VALVES") == "ARTICLE"

    def test_paragraph(self):
        assert detect_marker_class("A. Provide valves as specified.") == "PARAGRAPH"

    def test_paragraph_b(self):
        assert detect_marker_class("B. Submit shop drawings.") == "PARAGRAPH"

    def test_subparagraph(self):
        assert detect_marker_class("1. Type A valve") == "SUBPARAGRAPH"

    def test_subsubparagraph(self):
        assert detect_marker_class("a. 150 psi minimum") == "SUBSUBPARAGRAPH"

    def test_plain_text_none(self):
        assert detect_marker_class("Provide valves as specified.") is None

    def test_empty_none(self):
        assert detect_marker_class("") is None

    def test_all_caps_title_none(self):
        """All-caps text without a CSI marker is not classified."""
        assert detect_marker_class("COMMON MOTOR REQUIREMENTS") is None
