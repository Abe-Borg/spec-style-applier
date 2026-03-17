from __future__ import annotations

from typing import Any, Dict, List


def choose_section_sources(
    target_count: int,
    page_layout: Dict[str, Any],
    *,
    require_default: bool,
    log: List[str],
) -> List[Dict[str, Any]]:
    chain_raw = page_layout.get("section_chain", []) if isinstance(page_layout, dict) else []
    chain = [item for item in chain_raw if isinstance(item, dict)]
    default_raw = page_layout.get("default_section") if isinstance(page_layout, dict) else None
    default_section = default_raw if isinstance(default_raw, dict) else None

    if target_count == len(chain) and target_count > 0:
        return chain

    if default_section is not None:
        mapped: List[Dict[str, Any]] = []
        for idx in range(target_count):
            mapped.append(chain[idx] if idx < len(chain) else default_section)
        if target_count != len(chain):
            log.append(
                f"target sections={target_count}, architect sections={len(chain)}; "
                "using index-aligned sources + default_section for overflow"
            )
        return mapped

    if require_default and target_count != len(chain):
        raise ValueError(
            "Template registry missing usable page_layout.default_section for section-count mismatch"
        )

    return chain[:target_count]

