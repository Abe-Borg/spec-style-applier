"""
XML helper functions for paragraph-level DOCX manipulation.

All functions use regex-based XML processing (not DOM/ElementTree)
to preserve byte-level fidelity.
"""

import re
import html
from typing import Dict, Any, Optional, Tuple, Generator


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def iter_paragraph_xml_blocks(document_xml_text: str) -> Generator[Tuple[int, int, str], None, None]:
    # Non-greedy paragraph blocks. Works well for DOCX document.xml.
    # NOTE: This intentionally avoids parsing full XML to keep indices aligned with raw text.
    for m in re.finditer(r"(<w:p\b[\s\S]*?</w:p>)", document_xml_text):
        yield m.start(), m.end(), m.group(1)


def paragraph_text_from_block(p_xml: str) -> str:
    texts = re.findall(r"<w:t\b[^>]*>([\s\S]*?)</w:t>", p_xml)
    if not texts:
        return ""
    joined = html.unescape("".join(texts))
    joined = re.sub(r"\s+", " ", joined).strip()
    return joined


def paragraph_contains_sectpr(p_xml: str) -> bool:
    return "<w:sectPr" in p_xml


def paragraph_pstyle_from_block(p_xml: str) -> Optional[str]:
    m = re.search(r"<w:pStyle\b[^>]*w:val=\"([^\"]+)\"", p_xml)
    return m.group(1) if m else None


def paragraph_numpr_from_block(p_xml: str) -> Dict[str, Optional[str]]:
    numId = None
    ilvl = None
    m1 = re.search(r"<w:numId\b[^>]*w:val=\"([^\"]+)\"", p_xml)
    m2 = re.search(r"<w:ilvl\b[^>]*w:val=\"([^\"]+)\"", p_xml)
    if m1: numId = m1.group(1)
    if m2: ilvl = m2.group(1)
    return {"numId": numId, "ilvl": ilvl}


def paragraph_ppr_hints_from_block(p_xml: str) -> Dict[str, Any]:
    # lightweight hints (alignment + ind + spacing)
    hints: Dict[str, Any] = {}
    m = re.search(r"<w:jc\b[^>]*w:val=\"([^\"]+)\"", p_xml)
    if m:
        hints["jc"] = m.group(1)
    ind = {}
    for k in ["left", "right", "firstLine", "hanging"]:
        m2 = re.search(rf"<w:ind\b[^>]*w:{k}=\"([^\"]+)\"", p_xml)
        if m2:
            ind[k] = m2.group(1)
    if ind:
        hints["ind"] = ind
    spacing = {}
    for k in ["before", "after", "line"]:
        m3 = re.search(rf"<w:spacing\b[^>]*w:{k}=\"([^\"]+)\"", p_xml)
        if m3:
            spacing[k] = m3.group(1)
    if spacing:
        hints["spacing"] = spacing
    return hints


def apply_pstyle_to_paragraph_block(p_xml: str, styleId: str) -> str:
    # refuse to touch sectPr paragraph
    if "<w:sectPr" in p_xml:
        return p_xml

    # If pStyle already exists, replace its value
    if re.search(r"<w:pStyle\b", p_xml):
        p_xml = re.sub(
            r'(<w:pStyle\b[^>]*w:val=")([^"]+)(")',
            rf'\g<1>{styleId}\g<3>',
            p_xml,
            count=1
        )
        return p_xml

    # Handle self-closing pPr: <w:pPr/> or <w:pPr />
    if re.search(r"<w:pPr\b[^>]*/>", p_xml):
        p_xml = re.sub(
            r"<w:pPr\b[^>]*/>",
            rf'<w:pPr><w:pStyle w:val="{styleId}"/></w:pPr>',
            p_xml,
            count=1
        )
        return p_xml

    # If pPr exists as a normal open/close element, insert pStyle right after opening tag
    if "<w:pPr" in p_xml:
        p_xml = re.sub(
            r'(<w:pPr\b[^>]*>)',
            rf'\1<w:pStyle w:val="{styleId}"/>',
            p_xml,
            count=1
        )
        return p_xml

    # No pPr at all: create one right after <w:p ...>
    p_xml = re.sub(
        r'(<w:p\b[^>]*>)',
        rf'\1<w:pPr><w:pStyle w:val="{styleId}"/></w:pPr>',
        p_xml,
        count=1
    )
    return p_xml


def strip_run_font_formatting(p_xml: str) -> str:
    """
    Strip font-related formatting from all runs in a paragraph.

    This allows the paragraph style's font definitions to take effect,
    overriding hardcoded run-level fonts (common in MasterSpec/ARCOM docs).

    Strips from <w:rPr> inside <w:r>:
    - <w:rFonts .../> (font family)
    - <w:sz .../> (font size)
    - <w:szCs .../> (complex script font size)

    Preserves:
    - Bold, italic, underline, strikethrough
    - Colors, highlighting
    - Character styles (<w:rStyle>)
    - Everything else
    """
    # Don't touch sectPr paragraphs
    if "<w:sectPr" in p_xml:
        return p_xml

    def strip_font_from_rpr_text(rpr_text: str) -> str:
        """Process a raw rPr string."""
        result = rpr_text
        # Strip rFonts (self-closing or with content)
        result = re.sub(r'<w:rFonts\b[^>]*/>', '', result)
        result = re.sub(r'<w:rFonts\b[^>]*>[\s\S]*?</w:rFonts>', '', result, flags=re.S)
        # Strip sz (font size)
        result = re.sub(r'<w:sz\b[^>]*/>', '', result)
        # Strip szCs (complex script font size)
        result = re.sub(r'<w:szCs\b[^>]*/>', '', result)

        # Check if empty - remove entirely if so
        inner = re.sub(r'<w:rPr\b[^>]*>([\s\S]*)</w:rPr>', r'\1', result, flags=re.S)
        if not inner.strip():
            return ''
        return result

    def process_run(run_match):
        """Process a single <w:r>...</w:r> block."""
        run_block = run_match.group(0)

        # Find and replace rPr inside this run
        run_block = re.sub(
            r'<w:rPr\b[^>]*>[\s\S]*?</w:rPr>',
            lambda m: strip_font_from_rpr_text(m.group(0)),
            run_block,
            count=1,
            flags=re.S
        )
        return run_block

    # Process each run in the paragraph
    result = re.sub(
        r'<w:r\b[^>]*>[\s\S]*?</w:r>',
        process_run,
        p_xml,
        flags=re.S
    )

    return result


_DIRECT_PPR_OVERRIDE_TAGS = ("jc", "ind", "spacing")


def strip_conflicting_direct_ppr(p_xml: str) -> str:
    """
    Remove direct paragraph-layout overrides that commonly win over paragraph styles.

    Strips these tags from paragraph-level <w:pPr> only:
    - <w:jc>
    - <w:ind>
    - <w:spacing>

    Preserves numbering, section properties, and other pPr children.
    """
    if "<w:sectPr" in p_xml:
        return p_xml

    def _strip_from_ppr(match):
        ppr = match.group(0)
        for tag in _DIRECT_PPR_OVERRIDE_TAGS:
            ppr = re.sub(rf'<w:{tag}\b[^>]*/>', '', ppr)
            ppr = re.sub(rf'<w:{tag}\b[^>]*>[\s\S]*?</w:{tag}>', '', ppr, flags=re.S)
        return ppr

    return re.sub(r'<w:pPr\b[^>]*>[\s\S]*?</w:pPr>', _strip_from_ppr, p_xml, count=1, flags=re.S)


def _paragraph_style_id(p_xml: str) -> Optional[str]:
    m = re.search(r'<w:pStyle\b[^>]*w:val="([^"]+)"', p_xml)
    return m.group(1) if m else None


def _paragraph_has_numpr(p_xml: str) -> bool:
    return "<w:numPr" in p_xml
