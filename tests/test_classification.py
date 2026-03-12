"""Tests for core.classification — boilerplate filtering, especially END OF SECTION."""

import pytest
from core.classification import strip_boilerplate_with_report


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
