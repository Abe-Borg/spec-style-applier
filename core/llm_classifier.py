"""
LLM-based classification for Phase 2.

Sends paragraph bundles to the Anthropic API for CSI role classification,
with retry logic, chunking for large documents, and coverage reporting.
"""

import json
import time
import re
from typing import Dict, Any, List, Optional

from core.classification import PHASE2_MASTER_PROMPT, PHASE2_RUN_INSTRUCTION, detect_marker_class


# Approximate tokens per character for English text + JSON overhead
_CHARS_PER_TOKEN = 4
_MAX_BUNDLE_TOKENS = 80_000
_MAX_BUNDLE_CHARS = _MAX_BUNDLE_TOKENS * _CHARS_PER_TOKEN
_CHUNK_OVERLAP = 20  # paragraphs of overlap between chunks


def _estimate_tokens(text: str) -> int:
    """Rough token estimate for a string."""
    return len(text) // _CHARS_PER_TOKEN


def _build_user_message(slim_bundle: dict, available_roles: list) -> str:
    """Build the user message combining bundle and roles."""
    return (
        PHASE2_RUN_INSTRUCTION.strip()
        + "\n\navailable_roles: " + json.dumps(available_roles)
        + "\n\n" + json.dumps(slim_bundle, indent=2)
    )


def _build_api_kwargs(
    model: str,
    system_prompt: str,
    user_message: str,
    max_tokens: int = 128000,
) -> dict:
    """Build the kwargs dict for client.messages.stream().

    Centralised so the request shape is testable without hitting the API.
    """
    return dict(
        model=model,
        max_tokens=max_tokens,
        temperature=1,
        thinking={"type": "adaptive"},
        output_config={"effort": "high"},
        system=system_prompt,
        messages=[{"role": "user", "content": user_message}],
    )


def _parse_classification_response(response_text: str) -> dict:
    """Parse JSON from LLM response, handling markdown code blocks."""
    text = response_text.strip()

    # Strip markdown code block if present
    if text.startswith("```"):
        # Remove opening ```json or ``` line
        text = re.sub(r'^```\w*\s*\n?', '', text)
        # Remove closing ```
        text = re.sub(r'\n?```\s*$', '', text)

    return json.loads(text)


def _validate_classifications(
    classifications: dict,
    available_roles: list,
    total_paragraphs: Optional[int] = None,
) -> dict:
    """Validate and filter classification results.

    Args:
        classifications: Parsed JSON from the LLM.
        available_roles: Allowed CSI role strings.
        total_paragraphs: If provided, paragraph_index values must be in
            [0, total_paragraphs).  Out-of-range entries are dropped with
            a warning printed to stdout.

    Returns:
        Validated dict with 'classifications' (deduplicated) and 'notes'.
    """
    if not isinstance(classifications, dict):
        raise ValueError("LLM response is not a JSON object")

    items = classifications.get("classifications", [])
    if not isinstance(items, list):
        raise ValueError("LLM response missing 'classifications' array")

    valid_roles = set(available_roles)
    seen_indices: dict = {}  # paragraph_index -> item (last wins)
    dropped_roles: list = []
    dropped_range: list = []

    for item in items:
        if not isinstance(item, dict):
            continue
        idx = item.get("paragraph_index")
        role = item.get("csi_role")
        if not isinstance(idx, int) or not isinstance(role, str):
            continue
        if role not in valid_roles:
            dropped_roles.append((idx, role))
            continue
        if total_paragraphs is not None and (idx < 0 or idx >= total_paragraphs):
            dropped_range.append(idx)
            continue
        if idx in seen_indices:
            print(f"  WARNING: duplicate paragraph_index {idx} — keeping last occurrence")
        seen_indices[idx] = {"paragraph_index": idx, "csi_role": role}

    if dropped_roles:
        print(f"  WARNING: dropped {len(dropped_roles)} entries with unknown roles: "
              f"{dropped_roles[:5]}{'...' if len(dropped_roles) > 5 else ''}")
    if dropped_range:
        print(f"  WARNING: dropped {len(dropped_range)} out-of-range indices: "
              f"{dropped_range[:5]}{'...' if len(dropped_range) > 5 else ''}")

    validated = sorted(seen_indices.values(), key=lambda x: x["paragraph_index"])

    return {
        "classifications": validated,
        "notes": classifications.get("notes", [])
    }


def validate_and_repair_classifications(
    llm_result: dict,
    available_roles: list,
    ambiguous_indices: set,
    slim_bundle: dict,
) -> dict:
    """Stricter post-LLM validation with a repair pass for small gaps.

    Args:
        llm_result: Raw validated LLM output (from ``_validate_classifications``).
        available_roles: Allowed CSI role strings.
        ambiguous_indices: Set of paragraph indices the LLM was asked to classify.
        slim_bundle: The original slim bundle (for text lookup during repair).

    Returns:
        Validated+repaired dict with 'classifications' and 'notes'.

    Raises:
        ValueError: When coverage after repair is below 95% of ambiguous set.
    """
    classified = {c["paragraph_index"]: c["csi_role"] for c in llm_result.get("classifications", [])}
    role_set = set(available_roles)

    # Find missing ambiguous indices
    missing = ambiguous_indices - set(classified.keys())

    if missing:
        # Build text lookup from slim bundle
        text_by_idx = {p["paragraph_index"]: p.get("text", "") for p in slim_bundle.get("paragraphs", [])}

        repaired = []
        still_missing = []
        for idx in sorted(missing):
            text = text_by_idx.get(idx, "")
            marker = detect_marker_class(text) if text else None
            if marker and marker in role_set:
                classified[idx] = marker
                repaired.append(idx)
            else:
                still_missing.append(idx)

        if repaired:
            print(f"  Repaired {len(repaired)} missing classifications via marker detection")

        # Check coverage threshold
        if ambiguous_indices:
            coverage = (len(ambiguous_indices) - len(still_missing)) / len(ambiguous_indices) * 100
            if coverage < 95:
                raise ValueError(
                    f"Post-LLM coverage too low: {coverage:.1f}% of ambiguous paragraphs classified. "
                    f"Missing indices: {still_missing[:20]}{'...' if len(still_missing) > 20 else ''}"
                )
            if still_missing:
                print(f"  WARNING: {len(still_missing)} ambiguous paragraphs remain unclassified: "
                      f"{still_missing[:10]}{'...' if len(still_missing) > 10 else ''}")

    items = [{"paragraph_index": idx, "csi_role": role} for idx, role in sorted(classified.items())]
    return {
        "classifications": items,
        "notes": llm_result.get("notes", []),
    }


def merge_classifications(
    preclassified: Dict[int, str],
    llm_classified: dict,
    total_paragraphs: int,
) -> dict:
    """Merge deterministic and LLM classifications into a single result.

    Args:
        preclassified: ``{paragraph_index: role}`` from preclassifier.
        llm_classified: Validated LLM output dict with 'classifications' list.
        total_paragraphs: Total paragraph count for range validation.

    Returns:
        Merged dict with 'classifications' and 'notes'.
    """
    merged: Dict[int, str] = {}

    # Preclassified first
    for idx, role in preclassified.items():
        if 0 <= idx < total_paragraphs:
            merged[idx] = role

    # LLM results override (shouldn't conflict, but LLM wins if it does)
    for item in llm_classified.get("classifications", []):
        idx = item.get("paragraph_index")
        role = item.get("csi_role")
        if isinstance(idx, int) and isinstance(role, str) and 0 <= idx < total_paragraphs:
            if idx in merged and merged[idx] != role:
                print(f"  WARNING: LLM overrides preclassified role at paragraph {idx}: "
                      f"{merged[idx]} -> {role}")
            merged[idx] = role

    items = [{"paragraph_index": idx, "csi_role": role} for idx, role in sorted(merged.items())]
    return {
        "classifications": items,
        "notes": llm_classified.get("notes", []),
    }


def _split_bundle_into_chunks(
    slim_bundle: dict,
    max_chars: int = _MAX_BUNDLE_CHARS,
    preclassified: Optional[Dict[int, str]] = None,
    context_window: int = 3,
) -> List[dict]:
    """Split a large bundle into overlapping chunks.

    When *preclassified* is provided, preclassified paragraphs are included as
    read-only context (marked ``"preclassified": true``) rather than as items
    to classify.  Up to *context_window* surrounding preclassified paragraphs
    are included at chunk boundaries for disambiguation.
    """
    paragraphs = slim_bundle.get("paragraphs", [])
    meta = slim_bundle.get("document_meta", {})
    roles = slim_bundle.get("available_roles", [])
    filter_report = slim_bundle.get("filter_report", {})
    pre = preclassified or {}

    # Check if splitting is needed
    full_json = json.dumps(slim_bundle)
    if len(full_json) <= max_chars and len(paragraphs) <= 300:
        return [slim_bundle]

    # Calculate chunk size (number of paragraphs per chunk)
    overhead = len(json.dumps({
        "document_meta": meta,
        "available_roles": roles,
        "filter_report": {"paragraphs_removed_entirely": [], "paragraphs_stripped": []},
        "paragraphs": []
    }))
    avg_para_size = (len(full_json) - overhead) / max(len(paragraphs), 1)
    paras_per_chunk = max(10, int((max_chars - overhead) / max(avg_para_size, 1)))

    # Build an index for fast lookup
    para_by_idx = {p["paragraph_index"]: p for p in paragraphs}

    chunks = []
    start = 0
    while start < len(paragraphs):
        end = min(start + paras_per_chunk, len(paragraphs))
        chunk_paras = list(paragraphs[start:end])

        # Add preclassified context paragraphs at boundaries
        if pre:
            first_idx = chunk_paras[0]["paragraph_index"] if chunk_paras else 0
            last_idx = chunk_paras[-1]["paragraph_index"] if chunk_paras else 0

            context_entries = []
            for p_idx, p_role in sorted(pre.items()):
                # Include if within context_window of chunk boundaries
                if (first_idx - context_window <= p_idx < first_idx or
                        last_idx < p_idx <= last_idx + context_window):
                    if p_idx in para_by_idx:
                        ctx = dict(para_by_idx[p_idx])
                        ctx["preclassified"] = True
                        ctx["csi_role"] = p_role
                        context_entries.append(ctx)

            # Merge and sort by paragraph_index
            all_paras = chunk_paras + context_entries
            all_paras.sort(key=lambda p: p["paragraph_index"])
            chunk_paras = all_paras

        chunk = {
            "document_meta": meta,
            "available_roles": roles,
            "filter_report": {"paragraphs_removed_entirely": [], "paragraphs_stripped": []},
            "paragraphs": chunk_paras,
            "_chunk_info": {
                "chunk_index": len(chunks),
                "paragraph_range": [chunk_paras[0]["paragraph_index"], chunk_paras[-1]["paragraph_index"]] if chunk_paras else [0, 0]
            }
        }
        chunks.append(chunk)
        # Advance with overlap
        start = end - _CHUNK_OVERLAP if end < len(paragraphs) else end

    return chunks


def _merge_chunk_results(chunk_results: List[dict]) -> dict:
    """Merge classification results from multiple chunks, deduplicating by paragraph_index.

    When two chunks disagree on the same paragraph, the later chunk wins
    (it has more surrounding context) and a warning is printed.
    """
    seen: Dict[int, Any] = {}
    conflicts: list = []
    all_notes = []

    for chunk_idx, result in enumerate(chunk_results):
        for item in result.get("classifications", []):
            idx = item.get("paragraph_index")
            if idx is None:
                continue
            if idx in seen and seen[idx]["csi_role"] != item.get("csi_role"):
                conflicts.append((idx, seen[idx]["csi_role"], item.get("csi_role"), chunk_idx))
            seen[idx] = item  # Later chunks override
        all_notes.extend(result.get("notes", []))

    if conflicts:
        print(f"  WARNING: {len(conflicts)} chunk overlap conflict(s):")
        for idx, old_role, new_role, chunk_idx in conflicts[:5]:
            print(f"    paragraph {idx}: {old_role} -> {new_role} (chunk {chunk_idx} wins)")

    return {
        "classifications": sorted(seen.values(), key=lambda x: x.get("paragraph_index", 0)),
        "notes": all_notes
    }


def classify_target_document(
    slim_bundle: dict,
    available_roles: list,
    api_key: str,
    model: str = "claude-opus-4-6",
    preclassified: Optional[Dict[int, str]] = None,
) -> dict:
    """
    Classify paragraphs in a slim bundle using the Anthropic API.

    Args:
        slim_bundle: The slim bundle dict from build_phase2_slim_bundle()
        available_roles: List of available CSI role names
        api_key: Anthropic API key
        model: Model ID to use
        preclassified: If provided, dict of {paragraph_index: role} for
            paragraphs already deterministically classified.  These are
            merged into the final result and excluded from LLM input.

    Returns:
        Dict with 'classifications' list and 'notes' list
    """
    import anthropic

    client = anthropic.Anthropic(api_key=api_key)

    chunks = _split_bundle_into_chunks(slim_bundle)
    chunk_results = []

    for i, chunk in enumerate(chunks):
        if len(chunks) > 1:
            print(f"  Processing chunk {i + 1}/{len(chunks)}...")

        user_message = _build_user_message(chunk, available_roles)
        max_retries = 2
        last_error = None

        for attempt in range(max_retries + 1):
            try:
                api_kwargs = _build_api_kwargs(
                    model=model,
                    system_prompt=PHASE2_MASTER_PROMPT.strip(),
                    user_message=user_message,
                )
                with client.messages.stream(**api_kwargs) as stream:
                    response_text = stream.get_final_text()
                parsed = _parse_classification_response(response_text)
                validated = _validate_classifications(parsed, available_roles)
                chunk_results.append(validated)
                break

            except json.JSONDecodeError as e:
                last_error = e
                if attempt < max_retries:
                    print(f"  JSON parse error, retrying ({attempt + 1}/{max_retries})...")
                    time.sleep(2 ** attempt)
                else:
                    raise ValueError(f"Failed to parse LLM response as JSON after {max_retries + 1} attempts: {e}")

            except Exception as e:
                last_error = e
                if attempt < max_retries:
                    wait = 2 ** (attempt + 1)
                    print(f"  API error: {e}, retrying in {wait}s ({attempt + 1}/{max_retries})...")
                    time.sleep(wait)
                else:
                    raise RuntimeError(f"LLM classification failed after {max_retries + 1} attempts: {e}")

    # Merge chunks
    if len(chunk_results) == 1:
        result = chunk_results[0]
    else:
        result = _merge_chunk_results(chunk_results)

    # Merge preclassified results
    if preclassified:
        pre_items = [
            {"paragraph_index": idx, "csi_role": role}
            for idx, role in sorted(preclassified.items())
        ]
        # LLM results take precedence for any overlap (shouldn't happen)
        llm_indices = {c["paragraph_index"] for c in result.get("classifications", [])}
        merged_items = [c for c in pre_items if c["paragraph_index"] not in llm_indices]
        merged_items.extend(result.get("classifications", []))
        merged_items.sort(key=lambda x: x["paragraph_index"])
        result["classifications"] = merged_items

    # Coverage check
    total_paragraphs = len(slim_bundle.get("paragraphs", []))
    classified_count = len(result.get("classifications", []))
    if total_paragraphs > 0:
        coverage = classified_count / total_paragraphs * 100
        print(f"Classification coverage: {classified_count}/{total_paragraphs} ({coverage:.1f}%)")
        if coverage < 85:
            print(f"WARNING: Coverage below 85% — {total_paragraphs - classified_count} paragraphs unclassified")

    return result
