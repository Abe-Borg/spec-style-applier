"""
Registry loading and preflight reporting for Phase 2.

Handles loading the architect style registry and generating
preflight reports before classification application.
"""

import re
import json
import xml.etree.ElementTree as _ET
from pathlib import Path
from typing import Dict, Any, List, Optional, Set
from xml.sax.saxutils import escape as _sax_escape


def _xml_escape_attr(value: str) -> str:
    """Escape a string for safe use inside an XML attribute value (double-quoted)."""
    return _sax_escape(str(value), {'"': "&quot;"})


def build_arch_styles_xml_from_registry(registry: Dict[str, Any]) -> str:
    """
    Reconstruct a synthetic styles.xml string from arch_template_registry.json.

    This allows Phase 2 to operate entirely from the two JSON contract files
    without needing the architect's extracted word/styles.xml on disk.

    The output is a well-formed XML string containing <w:docDefaults> and all
    <w:style> blocks. The existing regex-based functions (extract_style_block_raw,
    _extract_basedOn, _find_style_numpr_in_chain, etc.) work on this string
    identically to how they work on a real styles.xml file.
    """
    style_defs = registry.get("styles", {}).get("style_defs", [])
    doc_defaults = registry.get("doc_defaults", {})

    default_rpr = doc_defaults.get("default_run_props", {}).get("rPr") or ""
    default_ppr = doc_defaults.get("default_paragraph_props", {}).get("pPr") or ""

    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
        ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">',
    ]

    # docDefaults
    parts.append("<w:docDefaults>")
    if default_rpr:
        parts.append(f"<w:rPrDefault>{default_rpr}</w:rPrDefault>")
    else:
        parts.append("<w:rPrDefault><w:rPr/></w:rPrDefault>")
    if default_ppr:
        parts.append(f"<w:pPrDefault>{default_ppr}</w:pPrDefault>")
    else:
        parts.append("<w:pPrDefault><w:pPr/></w:pPrDefault>")
    parts.append("</w:docDefaults>")

    # All style definitions
    for sd in style_defs:
        sid = sd.get("style_id", "")
        if not sid:
            continue

        stype = sd.get("type", "paragraph")
        name = sd.get("name") or sid
        based_on = sd.get("based_on")
        next_style = sd.get("next")
        link = sd.get("link")

        # XML-escape all attribute values (NOT raw XML property fragments)
        e_sid = _xml_escape_attr(sid)
        e_stype = _xml_escape_attr(stype)
        e_name = _xml_escape_attr(name)

        parts.append(f'<w:style w:type="{e_stype}" w:styleId="{e_sid}">')
        parts.append(f'<w:name w:val="{e_name}"/>')
        if based_on:
            parts.append(f'<w:basedOn w:val="{_xml_escape_attr(based_on)}"/>')
        if next_style:
            parts.append(f'<w:next w:val="{_xml_escape_attr(next_style)}"/>')
        if link:
            parts.append(f'<w:link w:val="{_xml_escape_attr(link)}"/>')
        if sd.get("ui_priority") is not None:
            parts.append(f'<w:uiPriority w:val="{_xml_escape_attr(sd["ui_priority"])}"/>')
        if sd.get("semi_hidden"):
            parts.append("<w:semiHidden/>")
        if sd.get("unhide_when_used"):
            parts.append("<w:unhideWhenUsed/>")
        if sd.get("qformat"):
            parts.append("<w:qFormat/>")

        # Raw XML fragments — inserted verbatim, never escaped
        if sd.get("pPr"):
            parts.append(sd["pPr"])
        if sd.get("rPr"):
            parts.append(sd["rPr"])
        if sd.get("tblPr"):
            parts.append(sd["tblPr"])
        if sd.get("trPr"):
            parts.append(sd["trPr"])
        if sd.get("tcPr"):
            parts.append(sd["tcPr"])

        parts.append("</w:style>")

    parts.append("</w:styles>")
    result = "\n".join(parts)

    # Validate that the generated XML is well-formed
    try:
        _ET.fromstring(result.encode("utf-8"))
    except _ET.ParseError as exc:
        raise ValueError(
            f"Synthetic styles.xml failed XML well-formedness check: {exc}"
        ) from exc

    return result


def resolve_arch_extract_root(p: Path) -> Path:
    """
    Resolve the architect template directory.

    Accepts a path that contains arch_style_registry.json and
    arch_template_registry.json. These two JSON files are the
    complete interface between Phase 1 and Phase 2 — no other
    files from the extracted folder are needed.

    Returns the directory path.
    """
    p = Path(p)

    # If they passed a file, use its parent directory
    if p.is_file():
        p = p.parent

    # Check for required contract files
    style_reg = p / "arch_style_registry.json"
    template_reg = p / "arch_template_registry.json"

    if not style_reg.exists():
        raise FileNotFoundError(
            f"arch_style_registry.json not found at: {style_reg}\n"
            "Point Phase 2 to the folder containing both JSON files from Phase 1."
        )
    if not template_reg.exists():
        raise FileNotFoundError(
            f"arch_template_registry.json not found at: {template_reg}\n"
            "Point Phase 2 to the folder containing both JSON files from Phase 1."
        )

    return p


def load_available_roles_from_registry(registry_path: Path) -> Optional[List[str]]:
    """
    Load the list of available role names from arch_style_registry.json.

    Args:
        registry_path: Path to arch_style_registry.json or the extracted folder containing it

    Returns:
        List of role names (e.g., ["SectionTitle", "PART", "ARTICLE", ...])
        Returns None if registry not found.
    """
    registry_path = Path(registry_path)

    # Handle both direct JSON path and folder path
    if registry_path.is_dir():
        registry_path = registry_path / "arch_style_registry.json"

    if not registry_path.exists():
        return None

    registry = json.loads(registry_path.read_text(encoding="utf-8"))
    roles = registry.get("roles", {})

    return sorted(roles.keys())


def load_arch_style_registry(arch_extract_dir: Path) -> Dict[str, str]:
    """
    Phase 2 contract (STRICT):
    - arch_style_registry.json must exist (emitted by Phase 1).
    - NO inference / NO heuristics.
    Returns: { role: styleId }
    """
    arch_extract_dir = Path(arch_extract_dir)

    # Allow passing the registry JSON directly
    if arch_extract_dir.is_file() and arch_extract_dir.suffix.lower() == ".json":
        reg_path = arch_extract_dir
        root_dir = arch_extract_dir.parent
    else:
        root_dir = resolve_arch_extract_root(arch_extract_dir)
        reg_path = root_dir / "arch_style_registry.json"

    if not reg_path.exists():
        raise FileNotFoundError(
            f"arch_style_registry.json not found at {reg_path}. "
            f"Run Phase 1 on the architect template and copy the extracted folder here."
        )

    reg = json.loads(reg_path.read_text(encoding="utf-8"))
    if not isinstance(reg, dict):
        raise ValueError("arch_style_registry.json must be a JSON object")

    # Expected shape:
    # { "version": 1, "source_docx": "...", "roles": { "PART": { "style_id": "X", ... }, ... } }
    roles = reg.get("roles")
    if not isinstance(roles, dict):
        raise ValueError("arch_style_registry.json missing 'roles' object")

    out: Dict[str, str] = {}
    for role, info in roles.items():
        if not isinstance(role, str) or not isinstance(info, dict):
            continue
        sid = info.get("style_id") or info.get("styleId")
        if isinstance(sid, str) and sid.strip():
            out[role.strip()] = sid.strip()

    if not out:
        raise ValueError("arch_style_registry.json contained no usable role->style mappings")

    return out


def write_phase2_preflight(
    extract_dir: Path,
    arch_root: Path,
    arch_registry: Dict[str, str],
    classifications: Dict[str, Any],
    out_path: Path
) -> Dict[str, Any]:
    # Count classifications per role
    role_counts: Dict[str, int] = {}
    for item in classifications.get("classifications", []):
        r = item.get("csi_role")
        if isinstance(r, str):
            role_counts[r] = role_counts.get(r, 0) + 1

    # Identify which roles are unmapped
    needed_roles = sorted(role_counts.keys())
    unmapped_roles = [r for r in needed_roles if r not in arch_registry]

    report = {
        "arch_extract_root": str(arch_root),
        "target_extract_root": str(extract_dir),
        "roles_in_classifications": role_counts,
        "arch_style_registry": arch_registry,
        "unmapped_roles": unmapped_roles,
    }

    out_path.write_text(json.dumps(report, indent=2), encoding="utf-8")
    return report


# ---------------------------------------------------------------------------
# Phase 2 preflight contract validation
# ---------------------------------------------------------------------------

_EXPECTED_TEMPLATE_SECTIONS = {
    "theme": dict,
    "settings": dict,
    "fonts": dict,
    "doc_defaults": dict,
    "styles": dict,
    "numbering": dict,
}

# Maps style_def property keys to the Word XML tag they should contain.
_STYLE_PR_TAG_MAP = {
    "pPr": "w:pPr",
    "rPr": "w:rPr",
    "tblPr": "w:tblPr",
    "trPr": "w:trPr",
    "tcPr": "w:tcPr",
}


def _check_xml_fragment(fragment: str, expected_tag: str) -> Optional[str]:
    """Return an error message if *fragment* lacks matching open/close tags for *expected_tag*."""
    if not isinstance(fragment, str) or not fragment.strip():
        return None  # empty/absent is fine — caller decides whether the key is required
    escaped = re.escape(expected_tag)
    # Self-closing is valid: <w:pPr/>
    if re.search(r"<" + escaped + r"(?:\s[^>]*)?\s*/\s*>", fragment):
        return None
    has_open = bool(re.search(r"<" + escaped + r"[\s>/]", fragment))
    has_close = bool(re.search(r"</" + escaped + r"\s*>", fragment))
    if not has_open or not has_close:
        return f"XML fragment for <{expected_tag}> is malformed: missing open or close tag"
    return None


def _validate_template_sections(
    template_registry: Dict[str, Any], errors: List[str]
) -> None:
    """Check 1: top-level section types."""
    for key, expected_type in _EXPECTED_TEMPLATE_SECTIONS.items():
        if key in template_registry:
            val = template_registry[key]
            if not isinstance(val, expected_type):
                errors.append(
                    f"Template registry section '{key}' must be {expected_type.__name__}, "
                    f"got {type(val).__name__}"
                )


def _validate_style_defs(
    template_registry: Dict[str, Any], errors: List[str]
) -> Set[str]:
    """Checks 2-4: style_defs is list, style_ids unique/usable, XML fragments parseable.

    Returns the set of known style IDs for cross-reference validation.
    """
    known_ids: Set[str] = set()
    styles_section = template_registry.get("styles")
    if styles_section is None:
        return known_ids
    if not isinstance(styles_section, dict):
        return known_ids  # already caught by _validate_template_sections

    style_defs = styles_section.get("style_defs")
    if style_defs is None:
        return known_ids
    if not isinstance(style_defs, list):
        errors.append(
            f"styles.style_defs must be a list, got {type(style_defs).__name__}"
        )
        return known_ids

    seen_ids: Dict[str, int] = {}  # style_id -> first index
    for idx, sd in enumerate(style_defs):
        if not isinstance(sd, dict):
            errors.append(f"styles.style_defs[{idx}] must be a dict, got {type(sd).__name__}")
            continue

        sid = sd.get("style_id")
        if not isinstance(sid, str) or not sid.strip():
            errors.append(f"styles.style_defs[{idx}] missing or empty 'style_id'")
            continue

        sid = sid.strip()
        if sid in seen_ids:
            errors.append(
                f"Duplicate style_id '{sid}' in styles.style_defs "
                f"(indices {seen_ids[sid]} and {idx})"
            )
        else:
            seen_ids[sid] = idx
        known_ids.add(sid)

        # Validate XML property fragments
        for pr_key, tag in _STYLE_PR_TAG_MAP.items():
            val = sd.get(pr_key)
            if val:
                err = _check_xml_fragment(val, tag)
                if err:
                    errors.append(f"styles.style_defs[{idx}] ('{sid}'): {err}")

    return known_ids


def _validate_compat_xml(
    template_registry: Dict[str, Any], errors: List[str]
) -> None:
    """Check 5: settings.compat.compat_xml is a full valid block."""
    settings = template_registry.get("settings")
    if not isinstance(settings, dict):
        return
    compat = settings.get("compat")
    if not isinstance(compat, dict):
        return
    compat_xml = compat.get("compat_xml")
    if not compat_xml:
        return
    if not isinstance(compat_xml, str):
        errors.append(
            f"settings.compat.compat_xml must be a string, got {type(compat_xml).__name__}"
        )
        return
    err = _check_xml_fragment(compat_xml, "w:compat")
    if err:
        errors.append(f"settings.compat.compat_xml: {err}")


def _validate_top_level_xml_fragments(
    template_registry: Dict[str, Any], errors: List[str]
) -> None:
    """Check 4 (continued): validate top-level XML fragments (theme, fonts)."""
    theme = template_registry.get("theme")
    if isinstance(theme, dict):
        xml = theme.get("theme1_xml")
        if xml:
            err = _check_xml_fragment(xml, "a:theme")
            if err:
                errors.append(f"theme.theme1_xml: {err}")

    fonts = template_registry.get("fonts")
    if isinstance(fonts, dict):
        xml = fonts.get("font_table_xml")
        if xml:
            err = _check_xml_fragment(xml, "w:fonts")
            if err:
                errors.append(f"fonts.font_table_xml: {err}")


def _validate_style_cross_ref(
    style_registry: Dict[str, str],
    known_style_ids: Set[str],
    errors: List[str],
) -> None:
    """Check 6: every style ID referenced by arch_style_registry exists in template style_defs."""
    for role, sid in sorted(style_registry.items()):
        if sid not in known_style_ids:
            errors.append(
                f"Style ID '{sid}' (mapped from role '{role}') in "
                "arch_style_registry.json not found in template registry style_defs"
            )


def _validate_numbering_consistency(
    template_registry: Dict[str, Any], errors: List[str]
) -> None:
    """Check 7: numbering num -> abstractNum references are internally consistent."""
    numbering = template_registry.get("numbering")
    if not isinstance(numbering, dict):
        return

    # Collect known abstractNumIds
    abstract_nums = numbering.get("abstract_nums", [])
    if not isinstance(abstract_nums, list):
        errors.append(
            f"numbering.abstract_nums must be a list, got {type(abstract_nums).__name__}"
        )
        return

    abstract_ids: Set[int] = set()
    for idx, an in enumerate(abstract_nums):
        if not isinstance(an, dict):
            errors.append(f"numbering.abstract_nums[{idx}] must be a dict")
            continue
        aid = an.get("abstractNumId")
        if not isinstance(aid, int):
            errors.append(
                f"numbering.abstract_nums[{idx}] missing or non-integer 'abstractNumId'"
            )
            continue
        abstract_ids.add(aid)

    # Validate num references
    nums = numbering.get("nums", [])
    if not isinstance(nums, list):
        errors.append(f"numbering.nums must be a list, got {type(nums).__name__}")
        return

    for idx, num in enumerate(nums):
        if not isinstance(num, dict):
            errors.append(f"numbering.nums[{idx}] must be a dict")
            continue
        nid = num.get("numId")
        if not isinstance(nid, int):
            errors.append(f"numbering.nums[{idx}] missing or non-integer 'numId'")
            continue
        ref_aid = num.get("abstractNumId")
        if not isinstance(ref_aid, int):
            errors.append(
                f"numbering.nums[{idx}] (numId={nid}) missing or non-integer 'abstractNumId'"
            )
            continue
        if ref_aid not in abstract_ids:
            errors.append(
                f"numbering.nums[{idx}] (numId={nid}) references "
                f"abstractNumId={ref_aid} which is not defined in abstract_nums"
            )


def preflight_validate_registries(
    style_registry: Dict[str, str],
    template_registry: Dict[str, Any],
) -> List[str]:
    """
    Validate both Phase 2 contract files before any mutation.

    Runs all checks and collects every error so the caller sees the full
    picture in a single pass.  An empty return list means validation passed.

    Args:
        style_registry:    role -> styleId mapping (from load_arch_style_registry).
        template_registry: full dict from arch_template_registry.json.

    Returns:
        List of error strings.  Empty list means the contract is valid.
    """
    errors: List[str] = []

    _validate_template_sections(template_registry, errors)
    known_style_ids = _validate_style_defs(template_registry, errors)
    _validate_compat_xml(template_registry, errors)
    _validate_top_level_xml_fragments(template_registry, errors)
    _validate_style_cross_ref(style_registry, known_style_ids, errors)
    _validate_numbering_consistency(template_registry, errors)

    return errors
