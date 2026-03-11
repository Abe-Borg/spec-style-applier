"""
Registry loading and preflight reporting for Phase 2.

Handles loading the architect style registry and generating
preflight reports before classification application.
"""

import json
from pathlib import Path
from typing import Dict, Any, List, Optional


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
        name = sd.get("name", sid)
        based_on = sd.get("based_on")
        next_style = sd.get("next")
        link = sd.get("link")

        style_attrs = f'w:type="{stype}" w:styleId="{sid}"'

        parts.append(f'<w:style {style_attrs}>')
        parts.append(f'<w:name w:val="{name}"/>')
        if based_on:
            parts.append(f'<w:basedOn w:val="{based_on}"/>')
        if next_style:
            parts.append(f'<w:next w:val="{next_style}"/>')
        if link:
            parts.append(f'<w:link w:val="{link}"/>')
        if sd.get("ui_priority") is not None:
            parts.append(f'<w:uiPriority w:val="{sd["ui_priority"]}"/>')
        if sd.get("semi_hidden"):
            parts.append("<w:semiHidden/>")
        if sd.get("unhide_when_used"):
            parts.append("<w:unhideWhenUsed/>")
        if sd.get("qformat"):
            parts.append("<w:qFormat/>")

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
    return "\n".join(parts)


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
