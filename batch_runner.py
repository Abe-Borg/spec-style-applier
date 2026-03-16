"""Shared Phase 2 file pipeline and concurrent batch runner."""

from __future__ import annotations

import hashlib
import json
import re
import time
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional

from arch_env_applier import apply_environment_to_target
from core.classification import apply_phase2_classifications, build_phase2_slim_bundle
from core.batch_classifier import (
    BatchClassificationError,
    build_batch_requests,
    reassemble_file_classifications,
    submit_and_poll,
)
from core.llm_classifier import classify_target_document
from core.registry import (
    build_arch_styles_xml_from_registry,
    load_arch_style_registry,
    load_available_roles_from_registry,
    preflight_validate_registries,
    resolve_arch_extract_root,
)
from core.stability import snapshot_stability, verify_stability
from core.style_import import import_arch_styles_into_target
from docx_decomposer import DocxDecomposer
from docx_patch import patch_docx

try:
    from numbering_importer import import_numbering

    HAS_NUMBERING_IMPORTER = True
except ImportError:
    HAS_NUMBERING_IMPORTER = False


@dataclass
class BatchResult:
    filename: str
    success: bool
    output_path: Optional[Path]
    log: List[str]
    error: Optional[str]
    duration_seconds: float


@dataclass(frozen=True)
class SharedConfig:
    arch_registry: Dict[str, str]
    env_registry: Dict[str, Any]
    arch_styles_xml: str
    available_roles: List[str]


@dataclass(frozen=True)
class PreparedFile:
    docx_path: Path
    extract_dir: Path
    bundle: Dict[str, Any]
    prep_log: List[str]


def _coverage_counts(bundle: Dict[str, Any], classifications: Dict[str, Any]) -> tuple[int, int, int]:
    total = len(bundle.get("paragraphs", [])) + len(bundle.get("deterministic_classifications", []))
    classified = len(classifications.get("classifications", []))
    return classified, total, len(bundle.get("paragraphs", []))


def _check_numbering_module_needed(arch_styles_xml: str, needed_style_ids: List[str]) -> None:
    """Raise if styles need numbering but numbering_importer is unavailable."""
    for sid in needed_style_ids:
        pat = r'<w:style[^>]*w:styleId="' + re.escape(sid) + r'"[^>]*>[\s\S]*?</w:style>'
        m = re.search(pat, arch_styles_xml)
        if m and '<w:numId' in m.group(0):
            raise ImportError(
                "numbering_importer module is not available but imported styles "
                f"require numbering definitions (e.g. style '{sid}'). "
                "Ensure numbering_importer.py is on the Python path."
            )


def load_and_validate_shared_config(arch_path: Path) -> SharedConfig:
    arch_registry = load_arch_style_registry(arch_path)
    arch_root = resolve_arch_extract_root(arch_path)
    available_roles = load_available_roles_from_registry(arch_path)
    if not available_roles:
        raise ValueError("Could not load architect registry")

    arch_template_registry_path = arch_root / "arch_template_registry.json"
    env_registry = json.loads(arch_template_registry_path.read_text(encoding="utf-8"))

    preflight_errors = preflight_validate_registries(arch_registry, env_registry)
    if preflight_errors:
        error_report = "\n".join(f"  - {e}" for e in preflight_errors)
        raise ValueError(
            f"Preflight validation failed ({len(preflight_errors)} error(s)):\n{error_report}"
        )

    arch_styles_xml = build_arch_styles_xml_from_registry(env_registry)
    return SharedConfig(
        arch_registry=arch_registry,
        env_registry=env_registry,
        arch_styles_xml=arch_styles_xml,
        available_roles=available_roles,
    )


def process_single_file(
    docx_path: Path,
    arch_registry: Dict[str, str],
    env_registry: Dict[str, Any],
    arch_styles_xml: str,
    available_roles: List[str],
    api_key: str,
    output_dir: Path,
    extract_base_dir: Path = Path("output"),
    model: str = "claude-opus-4-6",
) -> BatchResult:
    start = time.monotonic()
    per_file_log: List[str] = []
    filename = docx_path.name
    output_path: Optional[Path] = None

    try:
        digest = hashlib.sha256(str(docx_path.resolve()).encode("utf-8")).hexdigest()[:8]
        extract_dir_name = f"{docx_path.stem}_{digest}_extracted"

        per_file_log.append("Extracting DOCX...")
        decomposer = DocxDecomposer(str(docx_path))
        extract_dir = decomposer.extract(output_dir=extract_base_dir / extract_dir_name)

        per_file_log.append("Building slim bundle...")
        bundle = build_phase2_slim_bundle(extract_dir, available_roles=available_roles)
        unresolved = len(bundle.get("paragraphs", []))
        deterministic = len(bundle.get("deterministic_classifications", []))
        per_file_log.append(
            f"Built slim bundle: {unresolved} unresolved + {deterministic} deterministic"
        )

        if unresolved > 0 and not api_key:
            raise ValueError("Anthropic API key is required when unresolved paragraphs exist.")

        per_file_log.append("Classifying with LLM...")
        classifications = classify_target_document(
            slim_bundle=bundle,
            available_roles=available_roles,
            api_key=api_key,
            model=model,
        )

        classifications_path = extract_dir / "phase2_classifications.json"
        classifications_path.write_text(json.dumps(classifications, indent=2), encoding="utf-8")
        per_file_log.append(f"Classifications saved: {classifications_path}")

        apply_environment_to_target(target_extract_dir=extract_dir, registry=env_registry, log=per_file_log)
        per_file_log.append("Applied environment")

        used_roles = {
            item.get("csi_role")
            for item in classifications.get("classifications", [])
            if isinstance(item, dict) and isinstance(item.get("csi_role"), str)
        }
        needed_style_ids = sorted({arch_registry[r] for r in used_roles if r in arch_registry})

        style_numid_remap = {}
        if HAS_NUMBERING_IMPORTER:
            style_numid_remap = import_numbering(
                target_extract_dir=extract_dir,
                arch_template_registry=env_registry,
                arch_styles_xml=arch_styles_xml,
                style_ids_to_import=needed_style_ids,
                log=per_file_log,
            )
        else:
            _check_numbering_module_needed(arch_styles_xml, needed_style_ids)

        import_arch_styles_into_target(
            target_extract_dir=extract_dir,
            arch_styles_xml=arch_styles_xml,
            needed_style_ids=needed_style_ids,
            log=per_file_log,
            style_numid_remap=style_numid_remap,
        )
        per_file_log.append(f"Imported {len(needed_style_ids)} styles")

        snap = snapshot_stability(extract_dir)
        apply_report = apply_phase2_classifications(
            extract_dir=extract_dir,
            classifications=classifications,
            arch_style_registry=arch_registry,
            log=per_file_log,
        )
        verify_stability(extract_dir, snap)
        per_file_log.append("Applied classifications, stability verified")

        output_dir.mkdir(parents=True, exist_ok=True)
        output_path = output_dir / (docx_path.stem + "_PHASE2_FORMATTED.docx")
        replacements = {
            "word/document.xml": (extract_dir / "word" / "document.xml").read_bytes(),
            "word/styles.xml": (extract_dir / "word" / "styles.xml").read_bytes(),
        }

        for rel_path, local_path in [
            ("word/theme/theme1.xml", extract_dir / "word" / "theme" / "theme1.xml"),
            ("word/settings.xml", extract_dir / "word" / "settings.xml"),
            ("word/fontTable.xml", extract_dir / "word" / "fontTable.xml"),
            ("word/numbering.xml", extract_dir / "word" / "numbering.xml"),
            ("[Content_Types].xml", extract_dir / "[Content_Types].xml"),
            ("word/_rels/document.xml.rels", extract_dir / "word" / "_rels" / "document.xml.rels"),
        ]:
            if local_path.exists():
                replacements[rel_path] = local_path.read_bytes()

        for hf_path in sorted((extract_dir / "word").glob("header*.xml")):
            replacements[f"word/{hf_path.name}"] = hf_path.read_bytes()
        for hf_path in sorted((extract_dir / "word").glob("footer*.xml")):
            replacements[f"word/{hf_path.name}"] = hf_path.read_bytes()

        rels_dir = extract_dir / "word" / "_rels"
        if rels_dir.exists():
            for rels_path in sorted(rels_dir.glob("header*.xml.rels")):
                replacements[f"word/_rels/{rels_path.name}"] = rels_path.read_bytes()
            for rels_path in sorted(rels_dir.glob("footer*.xml.rels")):
                replacements[f"word/_rels/{rels_path.name}"] = rels_path.read_bytes()

        media_dir = extract_dir / "word" / "media"
        if media_dir.exists():
            for media_path in sorted(media_dir.iterdir()):
                if media_path.is_file():
                    replacements[f"word/media/{media_path.name}"] = media_path.read_bytes()

        with zipfile.ZipFile(docx_path, "r") as z:
            old_hf_parts = {
                n
                for n in z.namelist()
                if (n.startswith("word/header") or n.startswith("word/footer")) and n.endswith(".xml")
            }
            old_hf_rels = {
                n
                for n in z.namelist()
                if (n.startswith("word/_rels/header") or n.startswith("word/_rels/footer")) and n.endswith(".rels")
            }
        exclude_parts = (old_hf_parts | old_hf_rels) - set(replacements.keys())

        patch_docx(
            src_docx=docx_path,
            out_docx=output_path,
            replacements=replacements,
            exclude_parts=exclude_parts,
        )

        classified, total, unresolved = _coverage_counts(bundle, classifications)
        class_coverage = (classified / total * 100) if total > 0 else 100.0
        expected_targetable = apply_report.requested - len(apply_report.skipped_sectpr)
        app_coverage = (apply_report.modified / expected_targetable * 100) if expected_targetable > 0 else 100.0
        per_file_log.append(f"Output: {output_path}")
        per_file_log.append(f"Classification coverage: {classified}/{total} ({class_coverage:.1f}%)")
        per_file_log.append(
            f"Application coverage: {apply_report.modified}/{expected_targetable} ({app_coverage:.1f}%)"
        )

        issues_path = extract_dir / "phase2_issues.log"
        issues_path.write_text("\n".join(per_file_log) + "\n", encoding="utf-8")

        return BatchResult(
            filename=filename,
            success=True,
            output_path=output_path,
            log=per_file_log,
            error=None,
            duration_seconds=time.monotonic() - start,
        )
    except Exception as exc:
        per_file_log.append(f"FAILED: {exc}")
        return BatchResult(
            filename=filename,
            success=False,
            output_path=output_path,
            log=per_file_log,
            error=str(exc),
            duration_seconds=time.monotonic() - start,
        )


def _prepare_file_for_batch(
    docx_path: Path,
    available_roles: List[str],
    extract_base_dir: Path,
) -> PreparedFile:
    per_file_log: List[str] = []
    digest = hashlib.sha256(str(docx_path.resolve()).encode("utf-8")).hexdigest()[:8]
    extract_dir_name = f"{docx_path.stem}_{digest}_extracted"

    per_file_log.append("Extracting DOCX...")
    decomposer = DocxDecomposer(str(docx_path))
    extract_dir = decomposer.extract(output_dir=extract_base_dir / extract_dir_name)

    per_file_log.append("Building slim bundle...")
    bundle = build_phase2_slim_bundle(extract_dir, available_roles=available_roles)
    unresolved = len(bundle.get("paragraphs", []))
    deterministic = len(bundle.get("deterministic_classifications", []))
    per_file_log.append(
        f"Built slim bundle: {unresolved} unresolved + {deterministic} deterministic"
    )
    return PreparedFile(docx_path=docx_path, extract_dir=extract_dir, bundle=bundle, prep_log=per_file_log)


def _apply_batch_result(
    prepared: PreparedFile,
    classifications: Dict[str, Any],
    arch_registry: Dict[str, str],
    env_registry: Dict[str, Any],
    arch_styles_xml: str,
    output_dir: Path,
) -> BatchResult:
    start = time.monotonic()
    per_file_log = list(prepared.prep_log)
    output_path: Optional[Path] = None
    filename = prepared.docx_path.name

    try:
        classifications_path = prepared.extract_dir / "phase2_classifications.json"
        classifications_path.write_text(json.dumps(classifications, indent=2), encoding="utf-8")
        per_file_log.append(f"Classifications saved: {classifications_path}")

        apply_environment_to_target(target_extract_dir=prepared.extract_dir, registry=env_registry, log=per_file_log)
        per_file_log.append("Applied environment")

        used_roles = {
            item.get("csi_role")
            for item in classifications.get("classifications", [])
            if isinstance(item, dict) and isinstance(item.get("csi_role"), str)
        }
        needed_style_ids = sorted({arch_registry[r] for r in used_roles if r in arch_registry})

        style_numid_remap = {}
        if HAS_NUMBERING_IMPORTER:
            style_numid_remap = import_numbering(
                target_extract_dir=prepared.extract_dir,
                arch_template_registry=env_registry,
                arch_styles_xml=arch_styles_xml,
                style_ids_to_import=needed_style_ids,
                log=per_file_log,
            )
        else:
            _check_numbering_module_needed(arch_styles_xml, needed_style_ids)

        import_arch_styles_into_target(
            target_extract_dir=prepared.extract_dir,
            arch_styles_xml=arch_styles_xml,
            needed_style_ids=needed_style_ids,
            log=per_file_log,
            style_numid_remap=style_numid_remap,
        )
        per_file_log.append(f"Imported {len(needed_style_ids)} styles")

        snap = snapshot_stability(prepared.extract_dir)
        apply_report = apply_phase2_classifications(
            extract_dir=prepared.extract_dir,
            classifications=classifications,
            arch_style_registry=arch_registry,
            log=per_file_log,
        )
        verify_stability(prepared.extract_dir, snap)
        per_file_log.append("Applied classifications, stability verified")

        output_dir.mkdir(parents=True, exist_ok=True)
        output_path = output_dir / (prepared.docx_path.stem + "_PHASE2_FORMATTED.docx")
        replacements = {
            "word/document.xml": (prepared.extract_dir / "word" / "document.xml").read_bytes(),
            "word/styles.xml": (prepared.extract_dir / "word" / "styles.xml").read_bytes(),
        }

        for rel_path, local_path in [
            ("word/theme/theme1.xml", prepared.extract_dir / "word" / "theme" / "theme1.xml"),
            ("word/settings.xml", prepared.extract_dir / "word" / "settings.xml"),
            ("word/fontTable.xml", prepared.extract_dir / "word" / "fontTable.xml"),
            ("word/numbering.xml", prepared.extract_dir / "word" / "numbering.xml"),
            ("[Content_Types].xml", prepared.extract_dir / "[Content_Types].xml"),
            ("word/_rels/document.xml.rels", prepared.extract_dir / "word" / "_rels" / "document.xml.rels"),
        ]:
            if local_path.exists():
                replacements[rel_path] = local_path.read_bytes()

        for hf_path in sorted((prepared.extract_dir / "word").glob("header*.xml")):
            replacements[f"word/{hf_path.name}"] = hf_path.read_bytes()
        for hf_path in sorted((prepared.extract_dir / "word").glob("footer*.xml")):
            replacements[f"word/{hf_path.name}"] = hf_path.read_bytes()

        rels_dir = prepared.extract_dir / "word" / "_rels"
        if rels_dir.exists():
            for rels_path in sorted(rels_dir.glob("header*.xml.rels")):
                replacements[f"word/_rels/{rels_path.name}"] = rels_path.read_bytes()
            for rels_path in sorted(rels_dir.glob("footer*.xml.rels")):
                replacements[f"word/_rels/{rels_path.name}"] = rels_path.read_bytes()

        media_dir = prepared.extract_dir / "word" / "media"
        if media_dir.exists():
            for media_path in sorted(media_dir.iterdir()):
                if media_path.is_file():
                    replacements[f"word/media/{media_path.name}"] = media_path.read_bytes()

        with zipfile.ZipFile(prepared.docx_path, "r") as z:
            old_hf_parts = {
                n
                for n in z.namelist()
                if (n.startswith("word/header") or n.startswith("word/footer")) and n.endswith(".xml")
            }
            old_hf_rels = {
                n
                for n in z.namelist()
                if (n.startswith("word/_rels/header") or n.startswith("word/_rels/footer")) and n.endswith(".rels")
            }
        exclude_parts = (old_hf_parts | old_hf_rels) - set(replacements.keys())

        patch_docx(
            src_docx=prepared.docx_path,
            out_docx=output_path,
            replacements=replacements,
            exclude_parts=exclude_parts,
        )

        classified, total, unresolved = _coverage_counts(prepared.bundle, classifications)
        class_coverage = (classified / total * 100) if total > 0 else 100.0
        expected_targetable = apply_report.requested - len(apply_report.skipped_sectpr)
        app_coverage = (apply_report.modified / expected_targetable * 100) if expected_targetable > 0 else 100.0
        per_file_log.append(f"Output: {output_path}")
        per_file_log.append(f"Classification coverage: {classified}/{total} ({class_coverage:.1f}%)")
        per_file_log.append(
            f"Application coverage: {apply_report.modified}/{expected_targetable} ({app_coverage:.1f}%)"
        )

        issues_path = prepared.extract_dir / "phase2_issues.log"
        issues_path.write_text("\n".join(per_file_log) + "\n", encoding="utf-8")

        return BatchResult(
            filename=filename,
            success=True,
            output_path=output_path,
            log=per_file_log,
            error=None,
            duration_seconds=time.monotonic() - start,
        )
    except Exception as exc:
        per_file_log.append(f"FAILED: {exc}")
        return BatchResult(
            filename=filename,
            success=False,
            output_path=output_path,
            log=per_file_log,
            error=str(exc),
            duration_seconds=time.monotonic() - start,
        )


def run_batch_concurrent(
    docx_paths: List[Path],
    arch_registry: Dict[str, str],
    env_registry: Dict[str, Any],
    arch_styles_xml: str,
    available_roles: List[str],
    api_key: str,
    output_dir: Path,
    max_workers: int = 3,
    on_file_complete: Optional[Callable[[BatchResult], None]] = None,
) -> List[BatchResult]:
    if not docx_paths:
        return []

    workers = max(1, min(max_workers, len(docx_paths)))
    results: List[BatchResult] = []

    with ThreadPoolExecutor(max_workers=workers) as executor:
        futures = {
            executor.submit(
                process_single_file,
                docx_path,
                arch_registry,
                env_registry,
                arch_styles_xml,
                available_roles,
                api_key,
                output_dir,
            ): docx_path
            for docx_path in docx_paths
        }

        for future in as_completed(futures):
            result = future.result()
            results.append(result)
            if on_file_complete:
                on_file_complete(result)

    return sorted(results, key=lambda item: item.filename)


def run_batch_api(
    docx_paths: List[Path],
    arch_registry: Dict[str, str],
    env_registry: Dict[str, Any],
    arch_styles_xml: str,
    available_roles: List[str],
    api_key: str,
    output_dir: Path,
    max_workers: int = 3,
    poll_interval: int = 30,
    on_file_complete: Optional[Callable[[BatchResult], None]] = None,
    on_batch_poll: Optional[Callable[[str, str, Any], None]] = None,
    extract_base_dir: Path = Path("output"),
    model: str = "claude-opus-4-6",
) -> List[BatchResult]:
    if not docx_paths:
        return []

    workers = max(1, min(max_workers, len(docx_paths)))
    prepared_files: Dict[str, PreparedFile] = {}

    with ThreadPoolExecutor(max_workers=workers) as executor:
        futures = {
            executor.submit(_prepare_file_for_batch, docx_path, available_roles, extract_base_dir): docx_path
            for docx_path in docx_paths
        }
        for future in as_completed(futures):
            prepared = future.result()
            prepared_files[prepared.docx_path.name] = prepared

    file_bundles = {name: prepared.bundle for name, prepared in prepared_files.items()}
    requests = build_batch_requests(file_bundles, available_roles, model)

    raw_results = submit_and_poll(
        requests=requests,
        api_key=api_key,
        poll_interval=poll_interval,
        on_poll=on_batch_poll,
    )

    try:
        per_file_classifications = reassemble_file_classifications(raw_results, file_bundles, available_roles)
    except BatchClassificationError:
        raise

    results: List[BatchResult] = []
    with ThreadPoolExecutor(max_workers=workers) as executor:
        futures = {
            executor.submit(
                _apply_batch_result,
                prepared,
                per_file_classifications[filename],
                arch_registry,
                env_registry,
                arch_styles_xml,
                output_dir,
            ): filename
            for filename, prepared in prepared_files.items()
        }
        for future in as_completed(futures):
            result = future.result()
            results.append(result)
            if on_file_complete:
                on_file_complete(result)

    return sorted(results, key=lambda item: item.filename)
