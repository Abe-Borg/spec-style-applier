import json
from pathlib import Path

import pytest

from core.classification import (
    build_phase2_slim_bundle,
    preclassify_paragraphs,
    validate_phase2_classification_contract,
    PHASE2_MASTER_PROMPT,
    PHASE2_RUN_INSTRUCTION,
)


def _write_document_xml(tmp_path: Path, body_xml: str) -> Path:
    word = tmp_path / "word"
    word.mkdir(parents=True)
    doc_xml = (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">\n'
        f'<w:body>{body_xml}</w:body></w:document>'
    )
    (word / "document.xml").write_text(doc_xml, encoding="utf-8")
    return tmp_path


def test_prompts_loaded_from_files():
    assert "CSI STRUCTURE CLASSIFIER" in PHASE2_MASTER_PROMPT
    assert "Output schema" in PHASE2_RUN_INSTRUCTION


def test_bundle_enrichment_and_table_filtering(tmp_path: Path):
    extract_dir = _write_document_xml(
        tmp_path,
        (
            '<w:p><w:pPr><w:pStyle w:val="Body"/><w:ind w:left="720"/><w:spacing w:after="120"/></w:pPr>'
            '<w:r><w:rPr><w:b/></w:rPr><w:t>PART 1 GENERAL</w:t></w:r></w:p>'
            '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>A. In table</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
            '<w:p><w:r><w:t>1.01 SUMMARY</w:t></w:r></w:p>'
        ),
    )

    bundle = build_phase2_slim_bundle(extract_dir, "mechanical")
    assert bundle["deterministic_classifications"]
    assert len(bundle["paragraphs"]) == 0


def test_preclassify_markers():
    paragraphs = [
        {"paragraph_index": 0, "text": "PART 1 GENERAL", "in_table": False, "marker_type": None},
        {"paragraph_index": 1, "text": "1.01 SUMMARY", "in_table": False, "marker_type": None},
        {"paragraph_index": 2, "text": "A. Scope", "in_table": False, "marker_type": "upper_alpha"},
        {"paragraph_index": 3, "text": "1. Sub item", "in_table": False, "marker_type": "number"},
        {"paragraph_index": 4, "text": "a. Lower", "in_table": False, "marker_type": "lower_alpha"},
    ]
    roles = ["PART", "ARTICLE", "PARAGRAPH", "SUBPARAGRAPH", "SUBSUBPARAGRAPH"]
    out = preclassify_paragraphs(paragraphs, roles)
    assert out[0] == "PART"
    assert out[1] == "ARTICLE"
    assert out[2] == "PARAGRAPH"
    assert out[3] == "SUBPARAGRAPH"
    assert out[4] == "SUBSUBPARAGRAPH"


def test_validation_fails_on_duplicate_and_missing_coverage():
    bundle = {
        "paragraphs": [{"paragraph_index": 10}, {"paragraph_index": 11}],
        "deterministic_classifications": [],
    }
    with pytest.raises(ValueError, match="duplicate"):
        validate_phase2_classification_contract(
            bundle,
            {"classifications": [
                {"paragraph_index": 10, "csi_role": "PART"},
                {"paragraph_index": 10, "csi_role": "ARTICLE"},
            ]},
            ["PART", "ARTICLE"],
        )

    with pytest.raises(ValueError, match="missing coverage"):
        validate_phase2_classification_contract(
            bundle,
            {"classifications": [{"paragraph_index": 10, "csi_role": "PART"}]},
            ["PART", "ARTICLE"],
        )
