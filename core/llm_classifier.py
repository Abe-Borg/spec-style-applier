"""
LLM-based classification for Phase 2.

Sends paragraph bundles to the Anthropic API for CSI role classification,
with retry logic, chunking for large documents, and coverage reporting.
"""

import json
import time
import re
from typing import Dict, Any, List

from core.classification import PHASE2_MASTER_PROMPT, PHASE2_RUN_INSTRUCTION


_CHARS_PER_TOKEN = 4
_MAX_BUNDLE_TOKENS = 80_000
_MAX_BUNDLE_CHARS = _MAX_BUNDLE_TOKENS * _CHARS_PER_TOKEN
_CHUNK_OVERLAP = 20


def _build_user_message(slim_bundle: dict, available_roles: list) -> str:
    return (
        PHASE2_RUN_INSTRUCTION.strip()
        + "\n\navailable_roles: " + json.dumps(available_roles)
        + "\n\n" + json.dumps(slim_bundle, indent=2)
    )


def _parse_classification_response(response_text: str) -> dict:
    text = response_text.strip()
    if text.startswith("```"):
        text = re.sub(r'^```\w*\s*\n?', '', text)
        text = re.sub(r'\n?```\s*$', '', text)
    return json.loads(text)


def _validate_classifications(classifications: dict, available_roles: list) -> dict:
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
        if isinstance(idx, int) and isinstance(role, str) and role in valid_roles:
            validated.append({"paragraph_index": idx, "csi_role": role})

    return {"classifications": validated, "notes": classifications.get("notes", [])}


def _split_bundle_into_chunks(slim_bundle: dict, max_chars: int = _MAX_BUNDLE_CHARS) -> List[dict]:
    paragraphs = slim_bundle.get("paragraphs", [])
    meta = slim_bundle.get("document_meta", {})
    roles = slim_bundle.get("available_roles", [])
    filter_report = slim_bundle.get("filter_report", {})

    full_json = json.dumps(slim_bundle)
    if len(full_json) <= max_chars and len(paragraphs) <= 300:
        return [slim_bundle]

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
        start = end - _CHUNK_OVERLAP if end < len(paragraphs) else end
    return chunks


def _merge_chunk_results(chunk_results: List[dict]) -> dict:
    seen: Dict[int, str] = {}
    conflicts: List[Dict[str, Any]] = []
    all_notes: List[Any] = []

    for result in chunk_results:
        for item in result.get("classifications", []):
            idx = item.get("paragraph_index")
            role = item.get("csi_role")
            if idx is None or role is None:
                continue
            prior = seen.get(idx)
            if prior is not None and prior != role:
                conflicts.append({"paragraph_index": idx, "existing_role": prior, "conflicting_role": role})
            seen[idx] = role
        all_notes.extend(result.get("notes", []))

    if conflicts:
        examples = ", ".join([f"{c['paragraph_index']}:{c['existing_role']}|{c['conflicting_role']}" for c in conflicts[:10]])
        raise ValueError(f"Chunk merge conflicts detected ({len(conflicts)} total): {examples}")

    return {
        "classifications": [{"paragraph_index": idx, "csi_role": role} for idx, role in sorted(seen.items())],
        "notes": all_notes,
    }


def _merge_deterministic_with_llm(slim_bundle: dict, llm_result: dict) -> dict:
    merged: Dict[int, str] = {
        item["paragraph_index"]: item["csi_role"]
        for item in slim_bundle.get("deterministic_classifications", [])
        if isinstance(item, dict) and isinstance(item.get("paragraph_index"), int) and isinstance(item.get("csi_role"), str)
    }
    for item in llm_result.get("classifications", []):
        idx = item.get("paragraph_index")
        role = item.get("csi_role")
        if isinstance(idx, int) and isinstance(role, str):
            merged[idx] = role
    return {
        "classifications": [{"paragraph_index": idx, "csi_role": role} for idx, role in sorted(merged.items())],
        "notes": llm_result.get("notes", []),
    }


def classify_target_document(slim_bundle: dict, available_roles: list, api_key: str, model: str = "claude-opus-4-6") -> dict:
    import anthropic

    client = anthropic.Anthropic(api_key=api_key)
    chunks = _split_bundle_into_chunks(slim_bundle)
    chunk_results = []

    for i, chunk in enumerate(chunks):
        if len(chunks) > 1:
            print(f"  Processing chunk {i + 1}/{len(chunks)}...")

        user_message = _build_user_message(chunk, available_roles)
        max_retries = 2

        for attempt in range(max_retries + 1):
            try:
                with client.messages.stream(
                    model=model,
                    max_tokens=128000,
                    temperature=1,
                    thinking={"type": "adaptive"},
                    output_config={"effort": "high"},
                    system=PHASE2_MASTER_PROMPT.strip(),
                    messages=[{"role": "user", "content": user_message}],
                ) as stream:
                    response_text = stream.get_final_text()
                parsed = _parse_classification_response(response_text)
                chunk_results.append(_validate_classifications(parsed, available_roles))
                break
            except json.JSONDecodeError as e:
                if attempt < max_retries:
                    print(f"  JSON parse error, retrying ({attempt + 1}/{max_retries})...")
                    time.sleep(2 ** attempt)
                else:
                    raise ValueError(f"Failed to parse LLM response as JSON after {max_retries + 1} attempts: {e}")
            except Exception as e:
                if attempt < max_retries:
                    wait = 2 ** (attempt + 1)
                    print(f"  API error: {e}, retrying in {wait}s ({attempt + 1}/{max_retries})...")
                    time.sleep(wait)
                else:
                    raise RuntimeError(f"LLM classification failed after {max_retries + 1} attempts: {e}")

    llm_only = chunk_results[0] if len(chunk_results) == 1 else _merge_chunk_results(chunk_results)
    result = _merge_deterministic_with_llm(slim_bundle, llm_only)

    total_expected = len(slim_bundle.get("paragraphs", [])) + len(slim_bundle.get("deterministic_classifications", []))
    classified_count = len(result.get("classifications", []))
    if total_expected > 0 and classified_count != total_expected:
        raise ValueError(
            f"Classification coverage incomplete: {classified_count}/{total_expected}. "
            "All classifiable paragraphs must be classified."
        )

    print(f"Classification coverage: {classified_count}/{total_expected} (100.0%)")
    return result
