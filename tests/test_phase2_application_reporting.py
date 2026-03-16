from pathlib import Path

import pytest

from core.classification import apply_phase2_classifications


DOC_XML = (
    '<?xml version="1.0" encoding="UTF-8"?>'
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:body>'
    '<w:p><w:pPr><w:spacing w:after="120"/></w:pPr><w:r><w:t>A</w:t></w:r></w:p>'
    '<w:p><w:pPr><w:spacing w:after="120"/></w:pPr><w:r><w:t>B</w:t></w:r></w:p>'
    '</w:body></w:document>'
)


STYLE_WITHOUT_PPR = (
    '<?xml version="1.0" encoding="UTF-8"?>'
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:style w:type="paragraph" w:styleId="Body"><w:name w:val="Body"/></w:style>'
    '</w:styles>'
)


STYLE_WITH_PPR = (
    '<?xml version="1.0" encoding="UTF-8"?>'
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:style w:type="paragraph" w:styleId="Body"><w:name w:val="Body"/>'
    '<w:pPr><w:spacing w:after="240"/></w:pPr></w:style>'
    '</w:styles>'
)


def _seed_extract(tmp_path: Path, styles_xml: str) -> Path:
    (tmp_path / "word").mkdir(parents=True, exist_ok=True)
    (tmp_path / "word" / "document.xml").write_text(DOC_XML, encoding="utf-8")
    (tmp_path / "word" / "styles.xml").write_text(styles_xml, encoding="utf-8")
    return tmp_path


def test_invalid_index_is_fatal(tmp_path):
    extract = _seed_extract(tmp_path, STYLE_WITHOUT_PPR)
    with pytest.raises(ValueError, match="Invalid paragraph indices"):
        apply_phase2_classifications(
            extract,
            {"classifications": [{"paragraph_index": 99, "csi_role": "PARAGRAPH"}]},
            {"PARAGRAPH": "Body"},
            [],
        )


def test_preserve_direct_ppr_when_style_lacks_replacement(tmp_path):
    extract = _seed_extract(tmp_path, STYLE_WITHOUT_PPR)
    report = apply_phase2_classifications(
        extract,
        {"classifications": [{"paragraph_index": 0, "csi_role": "PARAGRAPH"}]},
        {"PARAGRAPH": "Body"},
        [],
    )
    out = (extract / "word" / "document.xml").read_text(encoding="utf-8")
    assert '<w:spacing w:after="120"/>' in out
    assert report.preserved_direct_ppr == 1
    assert report.stripped_direct_ppr == 0


def test_strip_direct_ppr_when_style_has_replacement(tmp_path):
    extract = _seed_extract(tmp_path, STYLE_WITH_PPR)
    report = apply_phase2_classifications(
        extract,
        {"classifications": [{"paragraph_index": 0, "csi_role": "PARAGRAPH"}]},
        {"PARAGRAPH": "Body"},
        [],
    )
    out = (extract / "word" / "document.xml").read_text(encoding="utf-8")
    assert out.count('<w:spacing w:after="120"/>') == 1
    assert report.stripped_direct_ppr == 1
    assert report.preserved_direct_ppr == 0
