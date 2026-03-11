"""
LLM-based classification for Phase 2.

Sends paragraph bundles to the Anthropic API for CSI role classification,
with retry logic, chunking for large documents, and coverage reporting.
"""

import json
import time
import re
from typing import Dict, Any, List, Optional

from core.classification import PHASE2_MASTER_PROMPT, PHASE2_RUN_INSTRUCTION


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


def _validate_classifications(classifications: dict, available_roles: list) -> dict:
    """Validate and filter classification results."""
    if not isinstance(classifications, dict):
        raise ValueError("LLM response is not a JSON object")

    items = classifications.get("classifications", [])
    if not isinstance(items, list):
        raise ValueError("LLM response missing 'classifications' array")

    valid_roles = set(available_roles)
    validated = []
    for item in items:
        if not isinstance(item, dict):
            continue
        idx = item.get("paragraph_index")
        role = item.get("csi_role")
        if not isinstance(idx, int) or not isinstance(role, str):
            continue
        if role not in valid_roles:
            continue
        validated.append({"paragraph_index": idx, "csi_role": role})

    return {
        "classifications": validated,
        "notes": classifications.get("notes", [])
    }


def _split_bundle_into_chunks(slim_bundle: dict, max_chars: int = _MAX_BUNDLE_CHARS) -> List[dict]:
    """Split a large bundle into overlapping chunks."""
    paragraphs = slim_bundle.get("paragraphs", [])
    meta = slim_bundle.get("document_meta", {})
    roles = slim_bundle.get("available_roles", [])
    filter_report = slim_bundle.get("filter_report", {})

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

    chunks = []
    start = 0
    while start < len(paragraphs):
        end = min(start + paras_per_chunk, len(paragraphs))
        chunk_paras = paragraphs[start:end]
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
    """Merge classification results from multiple chunks, deduplicating by paragraph_index."""
    seen = {}
    all_notes = []

    for result in chunk_results:
        for item in result.get("classifications", []):
            idx = item.get("paragraph_index")
            if idx is not None:
                seen[idx] = item  # Later chunks override (they have more context)
        all_notes.extend(result.get("notes", []))

    return {
        "classifications": sorted(seen.values(), key=lambda x: x.get("paragraph_index", 0)),
        "notes": all_notes
    }


def classify_target_document(
    slim_bundle: dict,
    available_roles: list,
    api_key: str,
    model: str = "claude-sonnet-4-20250514"
) -> dict:
    """
    Classify paragraphs in a slim bundle using the Anthropic API.

    Args:
        slim_bundle: The slim bundle dict from build_phase2_slim_bundle()
        available_roles: List of available CSI role names
        api_key: Anthropic API key
        model: Model ID to use

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
                with client.messages.stream(
                    model=model,
                    max_tokens=16384,
                    temperature=0,
                    system=PHASE2_MASTER_PROMPT.strip(),
                    messages=[{"role": "user", "content": user_message}]
                ) as stream:
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

    # Coverage check
    total_paragraphs = len(slim_bundle.get("paragraphs", []))
    classified_count = len(result.get("classifications", []))
    if total_paragraphs > 0:
        coverage = classified_count / total_paragraphs * 100
        print(f"Classification coverage: {classified_count}/{total_paragraphs} ({coverage:.1f}%)")
        if coverage < 85:
            print(f"WARNING: Coverage below 85% — {total_paragraphs - classified_count} paragraphs unclassified")

    return result
