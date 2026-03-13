"""
Phase 2 classification: applying LLM classifications to paragraphs,
building slim bundles for LLM input, and boilerplate filtering.
"""

import re
import difflib
from pathlib import Path
from typing import Dict, Any, List, Optional

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


PHASE2_MASTER_PROMPT = r'''
You are a CSI STRUCTURE CLASSIFIER for AEC specifications.

Some paragraphs have already been pre-classified deterministically by a rule-based
engine.  You will ONLY receive paragraphs that need your judgment — the ambiguous
ones.  Pre-classified paragraphs are included as read-only context (marked with
"preclassified": true) to help you understand surrounding structure.

Your job:
- Classify EVERY ambiguous paragraph into its CSI semantic role
- ONLY use roles from the available_roles list
- If a paragraph's natural role is not in available_roles, use the closest parent role

CRITICAL: Every ambiguous paragraph_index you receive MUST appear exactly once in
your output.  Any paragraph you omit will remain unstyled in the output document.

Paragraphs to SKIP (do not include in classifications):
- Paragraphs marked "preclassified": true (already handled)
- Empty or blank paragraphs
- Paragraphs containing section/page break markup (w:sectPr)
- "END OF SECTION" lines
- Boilerplate lines that survived filtering (e.g., spec headers/footers)

CSI Hierarchy (for reference):
- SectionID: Section number line (e.g., "SECTION 23 05 13")
- SectionTitle: Section name line (e.g., "COMMON MOTOR REQUIREMENTS FOR HVAC EQUIPMENT")
- PART: Part headings (PART 1, PART 2, PART 3)
- ARTICLE: Article numbers (1.01, 2.03, etc.)
- PARAGRAPH: Lettered paragraphs (A., B., C.)
- SUBPARAGRAPH: Numbered under letters (1., 2., 3.)
- SUBSUBPARAGRAPH: Lettered under numbers (a., b., c.)

Fallback rules when a role is not available:
- If SectionID not available but SectionTitle is -> classify section numbers as SectionTitle
- If SUBSUBPARAGRAPH not available -> classify as SUBPARAGRAPH
- If SUBPARAGRAPH not available -> classify as PARAGRAPH

Rules:
- Do NOT create new roles outside of available_roles
- Do NOT reference formatting
- Do NOT re-classify paragraphs marked as preclassified
- Use context (surrounding paragraphs, including preclassified ones) to resolve ambiguity

Return strict JSON only — no markdown, no explanation.
'''

PHASE2_RUN_INSTRUCTION = r'''
Task:
Classify CSI roles for EVERY ambiguous paragraph using ONLY the roles listed in
available_roles.  Paragraphs marked "preclassified": true are already handled —
do NOT include them in your output.

Output schema:
{
  "classifications": [
    { "paragraph_index": 12, "csi_role": "PART" }
  ],
  "notes": []
}

IMPORTANT:
- Your classifications array MUST contain an entry for every ambiguous paragraph
- Every csi_role value MUST be one of the strings in available_roles
- Do not invent roles that aren't in available_roles
- When uncertain between two roles, prefer the more specific one (e.g., SUBPARAGRAPH over PARAGRAPH)
- Return strict JSON only — no markdown code blocks, no explanation text
'''


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


def detect_marker_class(text: str) -> Optional[str]:
    """Detect obvious CSI marker patterns in paragraph text.

    Returns a role name string if the text starts with a recognisable CSI
    marker, or ``None`` if ambiguous.  This is intentionally conservative —
    context-aware disambiguation happens in the preclassifier (P2-004).
    """
    t = text.strip()
    if not t:
        return None
    # SECTION 23 05 13
    if re.match(r'(?i)^\s*SECTION\s+\d{2}\s*\d{2}\s*\d{2}', t):
        return "SectionID"
    # PART 1 / PART 2 / PART 3
    if re.match(r'(?i)^\s*PART\s+[123]\b', t):
        return "PART"
    # 1.01, 2.03  (article numbering — digit.two-digits at start)
    if re.match(r'^\s*\d+\.\d{2}\b', t):
        return "ARTICLE"
    # A. B. C. (uppercase letter + dot + space)
    if re.match(r'^\s*[A-Z]\.\s', t):
        return "PARAGRAPH"
    # 1. 2. 3. (digit + dot + space, NOT article pattern)
    if re.match(r'^\s*\d+\.\s', t):
        return "SUBPARAGRAPH"
    # a. b. c. (lowercase letter + dot + space)
    if re.match(r'^\s*[a-z]\.\s', t):
        return "SUBSUBPARAGRAPH"
    return None


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
    skipped_paragraphs: List[Dict[str, Any]] = []
    filter_report = {
        "paragraphs_removed_entirely": [],
        "paragraphs_stripped": []
    }

    for idx, (_s, _e, p_xml) in enumerate(iter_paragraph_xml_blocks(doc_text)):
        if paragraph_contains_sectpr(p_xml):
            skipped_paragraphs.append({"paragraph_index": idx, "reason": "sectPr"})
            continue

        raw_text = paragraph_text_from_block(p_xml)
        if not raw_text:
            skipped_paragraphs.append({"paragraph_index": idx, "reason": "empty"})
            continue

        cleaned_text, tags = strip_boilerplate_with_report(raw_text)

        if not cleaned_text:
            if tags:
                filter_report["paragraphs_removed_entirely"].append({
                    "paragraph_index": idx,
                    "tags": tags,
                    "original_text_preview": raw_text[:120]
                })
            skipped_paragraphs.append({"paragraph_index": idx, "reason": "boilerplate"})
            continue

        if tags:
            filter_report["paragraphs_stripped"].append({
                "paragraph_index": idx,
                "tags": tags
            })

        numpr = paragraph_numpr_from_block(p_xml)
        pstyle = paragraph_pstyle_from_block(p_xml)
        ppr_hints = paragraph_ppr_hints_from_block(p_xml)
        marker = detect_marker_class(cleaned_text)
        is_all_caps = cleaned_text == cleaned_text.upper() and any(c.isalpha() for c in cleaned_text)

        entry: Dict[str, Any] = {
            "paragraph_index": idx,
            "text": cleaned_text[:200],
            "numPr": numpr if (numpr.get("numId") or numpr.get("ilvl")) else None,
            "contains_sectPr": False,
            "pStyle": pstyle,
            "ppr_hints": ppr_hints if ppr_hints else None,
            "marker_class": marker,
            "is_all_caps": is_all_caps,
        }
        paragraphs.append(entry)

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

    return {
        "document_meta": {
            "discipline": discipline
        },
        "available_roles": available_roles,
        "filter_report": filter_report,
        "skipped_paragraphs": skipped_paragraphs,
        "paragraphs": paragraphs
    }


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
