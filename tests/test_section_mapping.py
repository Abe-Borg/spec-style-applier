from core.section_mapping import choose_section_sources


def test_mismatched_sections_with_default_fill_use_expected_per_index_mapping():
    page_layout = {
        "section_chain": [{"name": "s0"}, {"name": "s1"}],
        "default_section": {"name": "d"},
    }
    out = choose_section_sources(4, page_layout, require_default=True, log=[])
    assert [x["name"] for x in out] == ["s0", "s1", "d", "d"]
