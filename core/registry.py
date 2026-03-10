"""
Registry loading and preflight reporting for Phase 2.

Handles loading the architect style registry and generating
preflight reports before classification application.
"""

import json
from pathlib import Path
from typing import Dict, Any, List, Optional


def resolve_arch_extract_root(p: Path) -> Path:
    """
    Accepts either:
      - extracted root folder (contains word/styles.xml)
      - word folder itself
    Returns the extracted root folder.
    """
    p = Path(p)

    # If they pass .../word, go up one
    if p.name.lower() == "word":
        p = p.parent

    styles_path = p / "word" / "styles.xml"
    if not styles_path.exists():
        raise FileNotFoundError(f"Architect styles.xml not found at: {styles_path}")

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
