"""
Deterministic pre-classifier for Phase 2.

Assigns obvious CSI roles using regex patterns and context-aware state
tracking, leaving only ambiguous paragraphs for the LLM.
"""

import re
from typing import Dict, List, Tuple, Any, Optional


# ---------------------------------------------------------------------------
# Pattern definitions
# ---------------------------------------------------------------------------

_SECTION_ID_RE = re.compile(r'(?i)^\s*SECTION\s+\d{2}\s*\d{2}\s*\d{2}')
_PART_RE = re.compile(r'(?i)^\s*PART\s+[123]\b')
_ARTICLE_RE = re.compile(r'^\s*\d+\.\d{2}\b')
_PARAGRAPH_RE = re.compile(r'^\s*[A-Z]\.\s')
_SUBPARAGRAPH_RE = re.compile(r'^\s*\d+\.\s')
_SUBSUBPARAGRAPH_RE = re.compile(r'^\s*[a-z]\.\s')


def _is_all_caps_title(text: str) -> bool:
    """True when text is all-uppercase and contains alphabetic chars."""
    return text == text.upper() and any(c.isalpha() for c in text)


def preclassify_paragraphs(
    slim_bundle: dict,
    available_roles: List[str],
    force_llm_all: bool = False,
) -> Tuple[Dict[int, str], Dict[int, dict]]:
    """Deterministically classify obvious CSI paragraphs.

    Args:
        slim_bundle: Output of ``build_phase2_slim_bundle()``.
        available_roles: Roles present in the architect registry.
        force_llm_all: When True, skip all pre-classification and mark
            everything as ambiguous (useful for debugging).

    Returns:
        (preclassified, ambiguous)
        - preclassified: ``{paragraph_index: csi_role}`` for deterministically
          resolved paragraphs.
        - ambiguous: ``{paragraph_index: paragraph_entry}`` for paragraphs that
          need LLM judgment.
    """
    paragraphs = slim_bundle.get("paragraphs", [])
    role_set = set(available_roles)

    if force_llm_all:
        return {}, {p["paragraph_index"]: p for p in paragraphs}

    preclassified: Dict[int, str] = {}
    ambiguous: Dict[int, dict] = {}

    # State machine: track current CSI hierarchy level for context-aware
    # disambiguation (e.g. "1." is SUBPARAGRAPH only under a PARAGRAPH).
    last_role: Optional[str] = None

    for para in paragraphs:
        idx = para["paragraph_index"]
        text = para.get("text", "").strip()

        if not text:
            ambiguous[idx] = para
            continue

        role = _match_role(text, last_role, role_set)

        if role is not None:
            preclassified[idx] = role
            last_role = role
        else:
            # SectionTitle heuristic: all-caps paragraph immediately after SectionID
            if (
                last_role == "SectionID"
                and "SectionTitle" in role_set
                and _is_all_caps_title(text)
                and not _PART_RE.match(text)
                and not _ARTICLE_RE.match(text)
            ):
                preclassified[idx] = "SectionTitle"
                last_role = "SectionTitle"
            else:
                ambiguous[idx] = para

    return preclassified, ambiguous


def _match_role(
    text: str,
    last_role: Optional[str],
    available: set,
) -> Optional[str]:
    """Attempt to match a single paragraph text to a CSI role.

    Returns the role string or None if ambiguous.
    """
    # Order matters: check more specific patterns first.

    # SECTION 23 05 13
    if "SectionID" in available and _SECTION_ID_RE.match(text):
        return "SectionID"

    # PART 1 / PART 2 / PART 3
    if "PART" in available and _PART_RE.match(text):
        return "PART"

    # 1.01, 2.03 (article numbering)
    if "ARTICLE" in available and _ARTICLE_RE.match(text):
        return "ARTICLE"

    # A. B. C. (uppercase letter + dot + space)
    if "PARAGRAPH" in available and _PARAGRAPH_RE.match(text):
        return "PARAGRAPH"

    # 1. 2. 3. (digit + dot + space) — only when we are under a PARAGRAPH
    # (otherwise ambiguous)
    if "SUBPARAGRAPH" in available and _SUBPARAGRAPH_RE.match(text):
        if last_role in ("PARAGRAPH", "SUBPARAGRAPH", "SUBSUBPARAGRAPH"):
            return "SUBPARAGRAPH"
        # Could also be SUBPARAGRAPH at start of doc; leave for LLM
        return None

    # a. b. c. (lowercase letter + dot + space) — only when context is right
    if "SUBSUBPARAGRAPH" in available and _SUBSUBPARAGRAPH_RE.match(text):
        if last_role in ("SUBPARAGRAPH", "SUBSUBPARAGRAPH"):
            return "SUBSUBPARAGRAPH"
        return None

    return None
