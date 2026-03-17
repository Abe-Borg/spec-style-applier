from __future__ import annotations

import re
from typing import Dict, List, Optional


CANONICAL_SECTPR_ORDER = [
    "headerReference", "footerReference", "type", "pgSz", "pgMar", "paperSrc",
    "pgBorders", "lnNumType", "pgNumType", "cols", "formProt", "vAlign",
    "noEndnote", "titlePg", "textDirection", "bidi", "rtlGutter", "docGrid",
    "printerSettings", "sectPrChange",
]


def extract_all_sectpr_blocks(document_xml: str) -> List[str]:
    return re.findall(r"<w:sectPr\b[\s\S]*?</w:sectPr>", document_xml)


def extract_tag_block(xml: str, tag: str) -> Optional[str]:
    self_closing = re.search(rf'(<w:{tag}\b[^>]*/>)', xml)
    if self_closing:
        return self_closing.group(1)
    paired = re.search(rf'(<w:{tag}\b[^>]*>[\s\S]*?</w:{tag}>)', xml, flags=re.S)
    return paired.group(1) if paired else None


def strip_tag_block(xml: str, tag: str) -> str:
    xml = re.sub(rf'<w:{tag}\b[^>]*/>', '', xml)
    return re.sub(rf'<w:{tag}\b[^>]*>[\s\S]*?</w:{tag}>', '', xml, flags=re.S)


def child_tag_name(child_xml: str) -> Optional[str]:
    m = re.match(r'<w:([A-Za-z0-9]+)\b', child_xml)
    return m.group(1) if m else None


def extract_sectpr_children(inner: str) -> List[str]:
    return [
        m.group(0)
        for m in re.finditer(r'<w:[A-Za-z0-9]+\b[^>]*(?:/>|>[\s\S]*?</w:[A-Za-z0-9]+>)', inner)
    ]


def replace_nth_sectpr_block(document_xml: str, idx: int, replacement: str) -> str:
    matches = list(re.finditer(r"<w:sectPr\b[\s\S]*?</w:sectPr>", document_xml))
    if idx < 0 or idx >= len(matches):
        return document_xml
    m = matches[idx]
    return document_xml[:m.start()] + replacement + document_xml[m.end():]


def canonical_sectpr_order_index() -> Dict[str, int]:
    return {tag: i for i, tag in enumerate(CANONICAL_SECTPR_ORDER)}

