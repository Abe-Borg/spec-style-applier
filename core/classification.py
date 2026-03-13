"""
Phase 2 classification: applying LLM classifications to paragraphs,
building slim bundles for LLM input, and boilerplate filtering.
"""

import re
import difflib
from pathlib import Path
from typing import Dict, Any, List, Optional, Set, Tuple

from core.xml_helpers import (
    iter_paragraph_xml_blocks,
    paragraph_text_from_block,
    paragraph_contains_sectpr,
    paragraph_numpr_from_block,
    paragraph_pstyle_from_block,
    paragraph_ppr_hints_from_block,
    apply_pstyle_to_paragraph_block,
    strip_run_font_formatting,
)
from core.style_import import ensure_explicit_numpr_from_current_style


def _load_prompt_text(filename: str) -> str:
    prompt_path = Path(__file__).parent / "prompts" / filename
    return prompt_path.read_text(encoding="utf-8")


PHASE2_MASTER_PROMPT = _load_prompt_text("phase2_master_prompt.txt")
PHASE2_RUN_INSTRUCTION = _load_prompt_text("phase2_run_instruction.txt")


# -------------------------------
# Phase 2: Boilerplate filtering (LLM input only)
# -------------------------------

BOILERPLATE_PATTERNS = [
    # END OF SECTION structural markers
    (r'(?i)^\s*END\s+OF\s+SECTION\s*.*$', 'end_of_section'),

    # Specifier notes - bracketed formats
    (r'\[Note to [Ss]pecifier[:\s][^\]]*\]', 'specifier_note'),
    (r'\[Specifier[:\s][^\]]*\]', 'specifier_note'),
    (r'\[SPECIFIER[:\s][^\]]*\]', 'specifier_note'),
    (r'(?i)\*\*\s*note to specifier\s*\*\*[^\n]*(?:\n(?!\n)[^\n]*)*', 'specifier_note'),
    (r'(?i)<<\s*note to specifier[^>]*>>', 'specifier_note'),
    (r'(?i)^\s*note to specifier:.*$', 'specifier_note'),

    # MasterSpec / AIA / ARCOM editorial instructions
    (r'(?i)^Retain or delete this article.*$', 'masterspec_instruction'),
    (r'(?i)^Retain [^\n]*paragraph[^\n]*below.*$', 'masterspec_instruction'),
    (r'(?i)^Retain [^\n]*subparagraph[^\n]*below.*$', 'masterspec_instruction'),
    (r'(?i)^Retain [^\n]*article[^\n]*below.*$', 'masterspec_instruction'),
    (r'(?i)^Retain [^\n]*section[^\n]*below.*$', 'masterspec_instruction'),
    (r'(?i)^Retain [^\n]*if .*$', 'masterspec_instruction'),
    (r'(?i)^Retain one of.*$', 'masterspec_instruction'),
    (r'(?i)^Retain one or more of.*$', 'masterspec_instruction'),
    (r'(?i)^Revise this Section by deleting.*$', 'masterspec_instruction'),
    (r'(?i)^Revise [^\n]*to suit [Pp]roject.*$', 'masterspec_instruction'),
    (r'(?i)^This Section uses the term.*$', 'masterspec_instruction'),
    (r'(?i)^Verify that Section titles.*$', 'masterspec_instruction'),
    (r'(?i)^Coordinate [^\n]*paragraph[^\n]* with.*$', 'masterspec_instruction'),
    (r'(?i)^Coordinate [^\n]*revision[^\n]* with.*$', 'masterspec_instruction'),
    (r'(?i)^The list below matches.*$', 'masterspec_instruction'),
    (r'(?i)^See [^\n]*Evaluations?[^\n]* for .*$', 'masterspec_instruction'),
    (r'(?i)^See [^\n]*Article[^\n]* in the Evaluations.*$', 'masterspec_instruction'),
    (r'(?i)^If retaining [^\n]*paragraph.*$', 'masterspec_instruction'),
    (r'(?i)^If retaining [^\n]*subparagraph.*$', 'masterspec_instruction'),
    (r'(?i)^If retaining [^\n]*article.*$', 'masterspec_instruction'),
    (r'(?i)^When [^\n]*characteristics are important.*$', 'masterspec_instruction'),
    (r'(?i)^Inspections in this article are.*$', 'masterspec_instruction'),
    (r'(?i)^Materials and thicknesses in schedules below.*$', 'masterspec_instruction'),
    (r'(?i)^Insulation materials and thicknesses are identified below.*$', 'masterspec_instruction'),
    (r'(?i)^Do not duplicate requirements.*$', 'masterspec_instruction'),
    (r'(?i)^Not all materials and thicknesses may be suitable.*$', 'masterspec_instruction'),
    (r'(?i)^Consider the exposure of installed insulation.*$', 'masterspec_instruction'),
    (r'(?i)^Flexible elastomeric and polyolefin thicknesses are limited.*$', 'masterspec_instruction'),
    (r'(?i)^To comply with ASHRAE.*insulation should have.*$', 'masterspec_instruction'),
    (r'(?i)^Architect should be prepared to reject.*$', 'masterspec_instruction'),

    # Copyright notices
    (r'(?i)^Copyright\s*©?\s*\d{4}.*$', 'copyright'),
    (r'(?i)^©\s*\d{4}.*$', 'copyright'),
    (r'(?i)^Exclusively published and distributed by.*$', 'copyright'),
    (r'(?i)all rights reserved.*$', 'copyright'),
    (r'(?i)proprietary\s+information.*$', 'copyright'),

    # Separator lines
    (r'^[\*]{4,}\s*$', 'separator'),
    (r'^[-]{4,}\s*$', 'separator'),
    (r'^[=]{4,}\s*$', 'separator'),

    # Page artifacts
    (r'(?i)^page\s+\d+\s*(?:of\s*\d+)?\s*$', 'page_number'),

    # Revision marks
    (r'(?i)\{revision[^\}]*\}', 'revision_mark'),

    # Hidden text markers
    (r'(?i)<<[^>]*hidden[^>]*>>', 'hidden_text'),
]

# Pre-compile for speed and to avoid repeated regex compilation
_BOILERPLATE_RX = [(re.compile(pat, flags=re.MULTILINE), tag) for pat, tag in BOILERPLATE_PATTERNS]

_TABLE_BLOCK_RX = re.compile(r"<w:tbl\b[\s\S]*?</w:tbl>")
_PART_RX = re.compile(r"^\s*PART\s+[123]\b", re.IGNORECASE)
_ARTICLE_RX = re.compile(r"^\s*\d+\.\d{2}\b")
_SECTION_ID_RX = re.compile(r"^\s*SECTION\s+\d{2}(?:\s+\d{2}){2,}\b", re.IGNORECASE)
_ALL_CAPS_RX = re.compile(r"^[^a-z]*[A-Z][^a-z]*$")
_MARKER_RX = [
    (re.compile(r"^\s*[A-Z]\.\s+"), "upper_alpha"),
    (re.compile(r"^\s*\d+\.\s+"), "number"),
    (re.compile(r"^\s*[a-z]\.\s+"), "lower_alpha"),
]


def _table_ranges(document_xml_text: str) -> List[Tuple[int, int]]:
    return [(m.start(), m.end()) for m in _TABLE_BLOCK_RX.finditer(document_xml_text)]


def _in_any_range(pos: int, ranges: List[Tuple[int, int]]) -> bool:
    for start, end in ranges:
        if start <= pos < end:
            return True
    return False


def _extract_rpr_hints(p_xml: str) -> Dict[str, Any]:
    hints: Dict[str, Any] = {}
    if re.search(r"<w:b\b", p_xml):
        hints["bold"] = True
    if re.search(r"<w:i\b", p_xml):
        hints["italic"] = True
    if re.search(r"<w:u\b", p_xml):
        hints["underline"] = True
    return hints


def _detect_marker_type(text: str, numpr: Dict[str, Optional[str]]) -> Optional[str]:
    for rx, marker_type in _MARKER_RX:
        if rx.match(text):
            return marker_type
    if numpr.get("numId"):
        return "list_numpr"
    return None


def _resolve_role(preferred: str, available_roles: List[str]) -> Optional[str]:
    fallback_chain = {
        "SectionID": ["SectionID", "SectionTitle"],
        "SUBSUBPARAGRAPH": ["SUBSUBPARAGRAPH", "SUBPARAGRAPH", "PARAGRAPH"],
        "SUBPARAGRAPH": ["SUBPARAGRAPH", "PARAGRAPH"],
        "PARAGRAPH": ["PARAGRAPH"],
        "PART": ["PART"],
        "ARTICLE": ["ARTICLE"],
        "SectionTitle": ["SectionTitle"],
    }
    for candidate in fallback_chain.get(preferred, [preferred]):
        if candidate in available_roles:
            return candidate
    return None


def _deterministic_role_for_paragraph(paragraph: Dict[str, Any], prev_text: str = "") -> Optional[str]:
    text = paragraph.get("text", "")
    if not text or paragraph.get("in_table"):
        return None
    if _SECTION_ID_RX.match(text):
        return "SectionID"
    if _PART_RX.match(text):
        return "PART"
    if _ARTICLE_RX.match(text):
        return "ARTICLE"
    if prev_text and _SECTION_ID_RX.match(prev_text) and _ALL_CAPS_RX.match(text):
        return "SectionTitle"

    marker_type = paragraph.get("marker_type")
    if marker_type == "upper_alpha":
        return "PARAGRAPH"
    if marker_type in {"number", "list_numpr"}:
        return "SUBPARAGRAPH"
    if marker_type == "lower_alpha":
        return "SUBSUBPARAGRAPH"
    return None


def preclassify_paragraphs(paragraphs: List[Dict[str, Any]], available_roles: List[str]) -> Dict[int, str]:
    out: Dict[int, str] = {}
    prev_text = ""
    for paragraph in paragraphs:
        preferred = _deterministic_role_for_paragraph(paragraph, prev_text=prev_text)
        resolved = _resolve_role(preferred, available_roles) if preferred else None
        if resolved:
            out[paragraph["paragraph_index"]] = resolved
        prev_text = paragraph.get("text", "")
    return out


def strip_boilerplate_with_report(content: str) -> tuple:
    """
    Strip boilerplate from a paragraph string and return (cleaned_text, matched_tags).
    Placeholders are NOT stripped here (your patterns do not remove generic [ ... ] placeholders).
    """
    cleaned = content
    hits: list = []

    for rx, tag in _BOILERPLATE_RX:
        if rx.search(cleaned):
            hits.append(tag)
            cleaned = rx.sub('', cleaned)

    # Clean up whitespace
    cleaned = re.sub(r'\n{3,}', '\n\n', cleaned)
    cleaned = re.sub(r'[ \t]+\n', '\n', cleaned)
    cleaned = cleaned.strip()

    # Deduplicate tags (stable order)
    if hits:
        seen = set()
        hits = [t for t in hits if not (t in seen or seen.add(t))]

    return cleaned, hits


def build_phase2_slim_bundle(
    extract_dir: Path,
    discipline: str,
    available_roles: Optional[List[str]] = None
) -> Dict[str, Any]:
    """
    Build the slim bundle for Phase 2 LLM classification.

    Args:
        extract_dir: Path to extracted DOCX folder
        discipline: "mechanical" or "plumbing"
        available_roles: List of role names available in the architect template.
                        If None, all standard roles are allowed.

    Returns:
        Dict containing document_meta, available_roles, filter_report, and paragraphs
    """
    doc_path = extract_dir / "word" / "document.xml"
    doc_text = doc_path.read_text(encoding="utf-8")

    paragraphs = []
    filter_report = {
        "paragraphs_removed_entirely": [],
        "paragraphs_stripped": []
    }

    table_ranges = _table_ranges(doc_text)
    raw_paragraphs = list(iter_paragraph_xml_blocks(doc_text))

    for idx, (start, _e, p_xml) in enumerate(raw_paragraphs):
        if paragraph_contains_sectpr(p_xml):
            continue

        raw_text = paragraph_text_from_block(p_xml)
        if not raw_text:
            continue

        cleaned_text, tags = strip_boilerplate_with_report(raw_text)

        if not cleaned_text:
            if tags:
                filter_report["paragraphs_removed_entirely"].append({
                    "paragraph_index": idx,
                    "tags": tags,
                    "original_text_preview": raw_text[:120]
                })
            continue

        if tags:
            filter_report["paragraphs_stripped"].append({
                "paragraph_index": idx,
                "tags": tags
            })

        numpr = paragraph_numpr_from_block(p_xml)
        in_table = _in_any_range(start, table_ranges)
        pstyle = paragraph_pstyle_from_block(p_xml)
        ppr_hints = paragraph_ppr_hints_from_block(p_xml)
        rpr_hints = _extract_rpr_hints(p_xml)
        marker_type = _detect_marker_type(cleaned_text, numpr)

        paragraphs.append({
            "paragraph_index": idx,
            "text": cleaned_text[:200],
            "prev_text": paragraphs[-1]["text"][:80] if paragraphs else "",
            "next_text": "",
            "pStyle": pstyle,
            "pPr_hints": ppr_hints,
            "rPr_hints": rpr_hints,
            "in_table": in_table,
            "marker_type": marker_type,
            "numPr": numpr if (numpr.get("numId") or numpr.get("ilvl")) else None,
            "contains_sectPr": False
        })

    for i in range(len(paragraphs) - 1):
        paragraphs[i]["next_text"] = paragraphs[i + 1]["text"][:80]

    # Default roles if none specified
    if available_roles is None:
        available_roles = [
            "SectionID",
            "SectionTitle",
            "PART",
            "ARTICLE",
            "PARAGRAPH",
            "SUBPARAGRAPH",
            "SUBSUBPARAGRAPH"
        ]

    deterministic = preclassify_paragraphs(paragraphs, available_roles)
    unresolved_paragraphs = [p for p in paragraphs if not p.get("in_table") and p["paragraph_index"] not in deterministic]

    return {
        "document_meta": {
            "discipline": discipline
        },
        "available_roles": available_roles,
        "filter_report": filter_report,
        "paragraphs": unresolved_paragraphs,
        "deterministic_classifications": [
            {"paragraph_index": idx, "csi_role": role}
            for idx, role in sorted(deterministic.items())
        ]
    }


def validate_phase2_classification_contract(bundle: Dict[str, Any], classifications: Dict[str, Any], allowed_roles: List[str]) -> None:
    if not isinstance(classifications, dict):
        raise ValueError("classifications payload must be an object")
    items = classifications.get("classifications")
    if not isinstance(items, list):
        raise ValueError("classifications payload missing classifications list")

    allowed = set(allowed_roles)
    classifiable_indices = {p["paragraph_index"] for p in bundle.get("paragraphs", [])}
    seen: Dict[int, str] = {}
    for item in items:
        if not isinstance(item, dict):
            raise ValueError("all classification entries must be objects")
        idx = item.get("paragraph_index")
        role = item.get("csi_role")
        if not isinstance(idx, int):
            raise ValueError(f"invalid paragraph_index: {idx!r}")
        if idx in seen:
            raise ValueError(f"duplicate classification for paragraph_index={idx}")
        if idx not in classifiable_indices:
            raise ValueError(f"classification index not classifiable: {idx}")
        if role not in allowed:
            raise ValueError(f"invalid csi_role for paragraph_index={idx}: {role!r}")
        seen[idx] = role

    missing = sorted(classifiable_indices - set(seen.keys()))
    if missing:
        raise ValueError(f"missing coverage for paragraph indices: {missing[:20]}")


def _normalize_paragraph_for_contract(p_xml: str) -> str:
    """
    Normalize paragraph for contract comparison.
    Strips elements we're allowed to change: pStyle, numPr, and
    run-level font formatting (rFonts, sz, szCs).  Also removes
    empty pPr / rPr shells so paragraphs that originally lacked
    these blocks compare equal after stripping.
    """
    out = p_xml
    # Strip pStyle (we change this)
    out = re.sub(r"<w:pStyle\b[^>]*/>", "", out)
    # Strip numPr (we may materialize this)
    out = re.sub(r"<w:numPr\b[^>]*>[\s\S]*?</w:numPr>", "", out, flags=re.S)
    # Strip run-level font formatting (we now strip this too)
    out = re.sub(r"<w:rFonts\b[^>]*/>", "", out)
    out = re.sub(r"<w:rFonts\b[^>]*>[\s\S]*?</w:rFonts>", "", out, flags=re.S)
    out = re.sub(r"<w:sz\b[^>]*/>", "", out)
    out = re.sub(r"<w:szCs\b[^>]*/>", "", out)
    # Clean up empty rPr blocks that might result
    out = re.sub(r"<w:rPr>\s*</w:rPr>", "", out)
    out = re.sub(r"<w:rPr\s*/>", "", out)
    # Clean up empty pPr blocks that might result
    out = re.sub(r"<w:pPr>\s*</w:pPr>", "", out)
    out = re.sub(r"<w:pPr\s*/>", "", out)
    return out


def apply_phase2_classifications(
    extract_dir: Path,
    classifications: Dict[str, Any],
    arch_style_registry: Dict[str, str],
    log: List[str]
) -> None:
    """
    Apply CSI role classifications to paragraphs by setting pStyle.

    Also strips run-level font formatting so the style's fonts take effect.
    This handles MasterSpec/ARCOM documents that have hardcoded fonts in every run.
    """
    doc_path = extract_dir / "word" / "document.xml"
    doc_text = doc_path.read_text(encoding="utf-8")

    # Load styles once so we can preserve style-linked numbering before swapping styles
    styles_xml_text = (extract_dir / "word" / "styles.xml").read_text(encoding="utf-8")
    style_ids_in_styles = set(re.findall(r'w:styleId="([^"]+)"', styles_xml_text))

    blocks = list(iter_paragraph_xml_blocks(doc_text))
    para_blocks = [b[2] for b in blocks]

    # Track which paragraphs we modify (for logging)
    modified_indices = set()

    # Contract check: normalize paragraphs for comparison
    contract_before = [_normalize_paragraph_for_contract(p) for p in para_blocks]

    items = classifications.get("classifications", [])
    if not isinstance(items, list):
        raise ValueError("phase2 classifications: 'classifications' must be a list")

    for item in items:
        if not isinstance(item, dict):
            log.append(f"Invalid classification entry (not object): {item!r}")
            continue

        idx = item.get("paragraph_index")
        role = item.get("csi_role")

        if not isinstance(idx, int) or idx < 0 or idx >= len(para_blocks):
            log.append(f"Invalid paragraph_index in classifications: {idx!r}")
            continue

        if not isinstance(role, str):
            log.append(f"Invalid csi_role type at paragraph {idx}: {role!r}")
            continue

        style_id = arch_style_registry.get(role)
        if not style_id:
            log.append(f"Unmapped CSI role '{role}' at paragraph {idx} (skipped)")
            continue

        if style_id not in style_ids_in_styles:
            raise ValueError(
                f"Phase 2 needs styleId '{style_id}' for role '{role}' at paragraph {idx}, "
                "but that styleId is not present in target word/styles.xml. "
                "Import failed or registry mismatch."
            )

        if paragraph_contains_sectpr(para_blocks[idx]):
            log.append(f"Skipped sectPr paragraph at index {idx}")
            continue

        # Preserve list continuation by materializing style-linked numPr *before* swapping styles.
        pb = para_blocks[idx]
        pb = ensure_explicit_numpr_from_current_style(pb, styles_xml_text)

        # Strip run-level font formatting so style fonts take effect
        pb = strip_run_font_formatting(pb)

        # Now safely swap pStyle
        para_blocks[idx] = apply_pstyle_to_paragraph_block(pb, style_id)
        modified_indices.add(idx)

    # Log summary
    log.append(f"Applied styles to {len(modified_indices)} paragraphs")
    log.append(f"Stripped run-level font formatting from modified paragraphs")

    # Enforce the diff contract.
    contract_after = [_normalize_paragraph_for_contract(p) for p in para_blocks]
    if len(contract_before) != len(contract_after):
        raise RuntimeError("Internal error: paragraph count changed during Phase 2 application")

    for i, (b, a) in enumerate(zip(contract_before, contract_after)):
        if b != a:
            diff = "\n".join(difflib.unified_diff(
                b.splitlines(),
                a.splitlines(),
                fromfile=f"before:p[{i}]",
                tofile=f"after:p[{i}]",
                lineterm=""
            ))
            raise ValueError(
                "Phase 2 invariant violation: paragraph content changed outside allowed edits "
                f"(pStyle/numPr/run fonts) at paragraph index {i}.\n" + diff[:4000]
            )

    # Rebuild document.xml
    out = []
    last = 0
    for (s, e, _), pb in zip(blocks, para_blocks):
        out.append(doc_text[last:s])
        out.append(pb)
        last = e
    out.append(doc_text[last:])
    doc_path.write_text("".join(out), encoding="utf-8")
