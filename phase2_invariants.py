import re
import hashlib
import zipfile
from pathlib import Path
from typing import List, Dict, Any
import xml.etree.ElementTree as ET

def _sha256(b: bytes) -> str:
    return hashlib.sha256(b).hexdigest()

def _read_docx_part(docx: Path, internal_path: str) -> bytes:
    with zipfile.ZipFile(docx, "r") as z:
        return z.read(internal_path)

def _extract_all_sectpr_blocks(document_xml: str) -> List[str]:
    return re.findall(r"<w:sectPr\b[\s\S]*?</w:sectPr>", document_xml)


def _normalize_sectpr_for_comparison(sectpr: str) -> str:
    """Strip managed layout tags so only non-layout section semantics are compared."""
    out = sectpr
    for tag in (
        "pgSz", "pgMar", "cols", "docGrid",
        "headerReference", "footerReference", "titlePg",
        "type", "pgNumType", "vAlign", "textDirection",
        "bidi", "rtlGutter", "lnNumType", "pgBorders",
        "paperSrc", "formProt", "noEndnote",
    ):
        out = re.sub(rf'<w:{tag}\b[^>]*/>', '', out)
        out = re.sub(rf'<w:{tag}\b[^>]*>[\s\S]*?</w:{tag}>', '', out, flags=re.S)
    # reduce whitespace noise introduced by stripping
    out = re.sub(r'>\s+<', '><', out)
    return out.strip()


def _extract_hf_relationship_subset(rels_xml: str) -> List[str]:
    rels = re.findall(r'<Relationship\b[^>]*/>', rels_xml)
    subset = [
        rel for rel in rels
        if 'relationships/header' in rel or 'relationships/footer' in rel
    ]
    return sorted(subset)


def _normalize_rpr_for_comparison(rpr_block: str) -> str:
    """
    Normalize an rPr block for comparison by stripping font-related elements.
    
    We allow changes to:
    - <w:rFonts .../> (font family)
    - <w:sz .../> (font size)  
    - <w:szCs .../> (complex script font size)
    
    Everything else (bold, italic, color, etc.) must remain unchanged.
    """
    result = rpr_block
    # Strip rFonts
    result = re.sub(r'<w:rFonts\b[^>]*/>', '', result)
    result = re.sub(r'<w:rFonts\b[^>]*>[\s\S]*?</w:rFonts>', '', result, flags=re.S)
    # Strip sz
    result = re.sub(r'<w:sz\b[^>]*/>', '', result)
    # Strip szCs
    result = re.sub(r'<w:szCs\b[^>]*/>', '', result)
    return result


def _extract_and_normalize_rpr_blocks(document_xml: str) -> List[str]:
    """
    Extract all rPr blocks from document.xml and normalize them.
    This allows us to check that non-font formatting is preserved.
    """
    rpr_blocks = re.findall(r"<w:rPr\b[\s\S]*?</w:rPr>", document_xml)
    return [_normalize_rpr_for_comparison(b) for b in rpr_blocks]


def verify_phase2_invariants(
    src_docx: Path,
    new_document_xml: bytes,
    new_docx: Path | None = None,
    arch_template_registry: Dict[str, Any] | None = None,
) -> None:
    """
    Verify Phase 2 invariants:
    1. sectPr non-layout semantics unchanged (managed layout tags may change)
    2. If architect header/footer data is present, output header/footer parts match architect set
       and sectPr references resolve to valid document rels IDs
    3. Run properties unchanged EXCEPT for font-related elements (rFonts, sz, szCs)
    
    The font exception allows us to strip hardcoded fonts from MasterSpec docs
    so that style-level fonts take effect.
    """
    # 1) sectPr non-layout semantics unchanged
    before_doc = _read_docx_part(src_docx, "word/document.xml").decode("utf-8", errors="strict")
    after_doc = new_document_xml.decode("utf-8", errors="strict")

    before_sectprs = _extract_all_sectpr_blocks(before_doc)
    after_sectprs = _extract_all_sectpr_blocks(after_doc)
    if len(before_sectprs) != len(after_sectprs):
        raise RuntimeError("INVARIANT FAIL: sectPr block count changed")

    before_norm = [_normalize_sectpr_for_comparison(s) for s in before_sectprs]
    after_norm = [_normalize_sectpr_for_comparison(s) for s in after_sectprs]
    if before_norm != after_norm:
        raise RuntimeError("INVARIANT FAIL: non-layout sectPr semantics changed")

    # 2) header/footer invariants (against architect registry, when provided)
    hf_data = (arch_template_registry or {}).get("headers_footers", {}) if isinstance(arch_template_registry, dict) else {}
    arch_headers = hf_data.get("headers", []) if isinstance(hf_data, dict) else []
    arch_footers = hf_data.get("footers", []) if isinstance(hf_data, dict) else []

    if new_docx is not None:
        before_rels = _read_docx_part(src_docx, "word/_rels/document.xml.rels").decode("utf-8", errors="strict")
        after_rels = _read_docx_part(new_docx, "word/_rels/document.xml.rels").decode("utf-8", errors="strict")
        if _extract_hf_relationship_subset(before_rels) != _extract_hf_relationship_subset(after_rels):
            raise RuntimeError("INVARIANT FAIL: relationship subset changed")

    if new_docx is not None and (arch_headers or arch_footers):
        expected_parts = {
            item.get("part_name")
            for item in [*arch_headers, *arch_footers]
            if isinstance(item, dict) and isinstance(item.get("part_name"), str)
        }
        with zipfile.ZipFile(new_docx, "r") as z_after:
            actual_parts = {
                n for n in z_after.namelist()
                if (n.startswith("word/header") or n.startswith("word/footer")) and n.endswith(".xml")
            }
            if expected_parts != actual_parts:
                raise RuntimeError(
                    "INVARIANT FAIL: output header/footer part set does not match architect registry\n"
                    f"Expected: {sorted(expected_parts)}\nActual: {sorted(actual_parts)}"
                )

            rels_xml = z_after.read("word/_rels/document.xml.rels")
            rels_root = ET.fromstring(rels_xml)
            rel_ids = {rel.attrib.get("Id") for rel in rels_root.findall('.//{*}Relationship') if rel.attrib.get("Id")}

            out_doc_xml = z_after.read("word/document.xml").decode("utf-8", errors="strict")
            refs = re.findall(r'<w:(?:headerReference|footerReference)\b[^>]*\br:id="([^"]+)"', out_doc_xml)
            unresolved = sorted({rid for rid in refs if rid not in rel_ids})
            if unresolved:
                raise RuntimeError(
                    "INVARIANT FAIL: document.xml contains header/footer refs with missing rel IDs: "
                    + ", ".join(unresolved)
                )

    # 3) no run-level formatting edits EXCEPT font-related (rFonts, sz, szCs)
    # We normalize rPr blocks by stripping font elements, then compare
    before_rpr_normalized = _extract_and_normalize_rpr_blocks(before_doc)
    after_rpr_normalized = _extract_and_normalize_rpr_blocks(after_doc)
    
    # Note: The number of rPr blocks might change if we remove empty ones,
    # so we compare the non-empty normalized blocks
    before_rpr_filtered = [b for b in before_rpr_normalized if b.strip() and b.strip() != '<w:rPr></w:rPr>']
    after_rpr_filtered = [b for b in after_rpr_normalized if b.strip() and b.strip() != '<w:rPr></w:rPr>']
    
    # Instead of strict equality (which fails if rPr blocks are removed),
    # we check that no NON-FONT formatting was changed.
    # This is a relaxed check - we're mainly guarding against accidental changes.
    
    # Verify that no non-font formatting was lost.
    # We can't do a strict count comparison because stripping fonts can remove
    # entire rPr blocks (when rFonts/sz/szCs were the only children).
    # Instead, check that every non-empty normalized "before" block still appears
    # somewhere in the "after" set. This catches accidental bold/italic/color changes.
    before_set = {}
    for b in before_rpr_filtered:
        before_set[b] = before_set.get(b, 0) + 1

    after_set = {}
    for a in after_rpr_filtered:
        after_set[a] = after_set.get(a, 0) + 1

    for block, count in before_set.items():
        after_count = after_set.get(block, 0)
        if after_count < count:
            raise RuntimeError(
                f"INVARIANT FAIL: non-font run formatting was lost. "
                f"A normalized rPr block appeared {count}x before but {after_count}x after.\n"
                f"Block: {block[:200]}"
            )
