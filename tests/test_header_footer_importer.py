import base64
from pathlib import Path

from core.xml_helpers import iter_paragraph_xml_blocks
from header_footer_importer import import_headers_footers, patch_footer_tokens


def _seed_extract(tmp_path: Path) -> Path:
    (tmp_path / "word" / "_rels").mkdir(parents=True, exist_ok=True)
    (tmp_path / "word" / "document.xml").write_text(
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<w:body><w:p><w:r><w:t>x</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>',
        encoding="utf-8",
    )
    (tmp_path / "word" / "_rels" / "document.xml.rels").write_text(
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header9.xml"/>'
        '</Relationships>',
        encoding="utf-8",
    )
    (tmp_path / "[Content_Types].xml").write_text(
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '</Types>',
        encoding="utf-8",
    )
    (tmp_path / "word" / "header9.xml").write_text("<old/>", encoding="utf-8")
    return tmp_path


def test_import_headers_footers_replaces_parts_and_refs(tmp_path):
    extract = _seed_extract(tmp_path)
    registry = {
        "headers_footers": {
            "headers": [
                {
                    "part_name": "word/header1.xml",
                    "rid": "rId10",
                    "xml": '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
                    "media": [
                        {
                            "path": "media/logo.png",
                            "content_base64": base64.b64encode(b"png").decode("ascii"),
                        }
                    ],
                }
            ],
            "footers": [
                {
                    "part_name": "word/footer1.xml",
                    "rid": "rId11",
                    "xml": '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
                }
            ],
        },
        "page_layout": {
            "section_chain": [
                {
                    "header_refs": {"default": "rId10"},
                    "footer_refs": {"default": "rId11"},
                }
            ]
        },
    }

    log = []
    import_headers_footers(extract, registry, log)

    assert not (extract / "word" / "header9.xml").exists()
    assert (extract / "word" / "header1.xml").exists()
    assert (extract / "word" / "footer1.xml").exists()
    assert any(p.read_bytes() == b"png" for p in (extract / "word" / "media").iterdir())

    rels_xml = (extract / "word" / "_rels" / "document.xml.rels").read_text(encoding="utf-8")
    assert "relationships/header" in rels_xml
    assert "relationships/footer" in rels_xml
    assert "header9.xml" not in rels_xml

    doc_xml = (extract / "word" / "document.xml").read_text(encoding="utf-8")
    assert "headerReference" in doc_xml
    assert "footerReference" in doc_xml
    assert "<ns0:" not in doc_xml
    assert "<w:p" in doc_xml
    assert len(list(iter_paragraph_xml_blocks(doc_xml))) == 1

    ct_xml = (extract / "[Content_Types].xml").read_text(encoding="utf-8")
    assert "/word/header1.xml" in ct_xml
    assert "/word/footer1.xml" in ct_xml
    assert 'Extension="png"' in ct_xml

def test_hf_rewire_preserves_unknown_document_prefixes(tmp_path):
    extract = _seed_extract(tmp_path)
    doc = (
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
        'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14">'
        '<w:body><w:p><w:r><w14:paraId w14:val="1234"/><w:t>x</w:t></w:r></w:p>'
        '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>'
    )
    (extract / "word" / "document.xml").write_text(doc, encoding="utf-8")
    registry = {
        "headers_footers": {"headers": [{"part_name": "word/header1.xml", "rid": "rId10", "xml": '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'}]},
        "page_layout": {"default_section": {"header_refs": {"default": "rId10"}}},
    }
    log = []
    import_headers_footers(extract, registry, log)
    out = (extract / "word" / "document.xml").read_text(encoding="utf-8")
    assert "xmlns:mc" in out and "xmlns:w14" in out
    assert 'mc:Ignorable="w14"' in out
    assert "<w14:paraId" in out


def test_hf_media_import_does_not_overwrite_existing_body_media(tmp_path):
    extract = _seed_extract(tmp_path)
    media_dir = extract / "word" / "media"
    media_dir.mkdir(parents=True, exist_ok=True)
    (media_dir / "image1.png").write_bytes(b"body")

    registry = {
        "headers_footers": {
            "headers": [{
                "part_name": "word/header1.xml",
                "rid": "rId10",
                "xml": '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
                "rels_xml": '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/></Relationships>',
                "media": [{"path": "media/image1.png", "content_base64": base64.b64encode(b"header").decode("ascii")}],
            }]
        },
        "page_layout": {"default_section": {"header_refs": {"default": "rId10"}}},
    }

    result = import_headers_footers(extract, registry, [])
    assert (media_dir / "image1.png").read_bytes() == b"body"
    assert result.media_names
    assert all(name.startswith("word/media/hf_") for name in result.media_names)


def test_patch_footer_tokens_handles_split_wt_nodes_and_case_mirroring(tmp_path):
    word_dir = tmp_path / "word"
    word_dir.mkdir(parents=True, exist_ok=True)
    footer_path = word_dir / "footer1.xml"
    footer_path.write_text(
        '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:tbl><w:tr><w:tc><w:p>'
        '<w:r><w:t>Metal </w:t></w:r><w:r><w:t>Ducts</w:t></w:r>'
        '<w:r><w:t xml:space="preserve"> </w:t></w:r>'
        '<w:r><w:t>23 31 00</w:t></w:r>'
        '</w:p></w:tc></w:tr></w:tbl>'
        '</w:ftr>',
        encoding="utf-8",
    )
    log = []
    patch_footer_tokens(
        target_extract_dir=tmp_path,
        source_tokens={"SectionTitle": "METAL DUCTS", "SectionID": "SECTION 23 31 00"},
        target_tokens={"SectionTitle": "Direct-Digital Control System for HVAC", "SectionID": "SECTION 23 09 00"},
        log=log,
    )

    out = footer_path.read_text(encoding="utf-8")
    assert "Direct-Digital Control System for HVAC" in out
    assert "23 09 00" in out
    assert "Metal " not in out
    assert any("Patched tokens in footer1.xml" in line for line in log)
