#!/usr/bin/env python3
"""
Word Document Decomposer and Phase 2 Styling Engine

CLI entry point for the Phase 2 MEP Specification Styling Engine.
Extracts DOCX files, builds LLM classification bundles, and applies
architect-defined CSI paragraph styles to MEP specification documents.
"""

import zipfile
import shutil
from pathlib import Path
from typing import List

from arch_env_applier import apply_environment_to_target

try:
    from numbering_importer import import_numbering, extract_used_num_ids_from_styles
    HAS_NUMBERING_IMPORTER = True
except ImportError:
    HAS_NUMBERING_IMPORTER = False

from core.classification import (
    PHASE2_MASTER_PROMPT,
    PHASE2_RUN_INSTRUCTION,
    build_phase2_slim_bundle,
    apply_phase2_classifications,
    coerce_to_final_classifications,
)
from core.stability import snapshot_stability, verify_stability
from core.style_import import import_arch_styles_into_target
from core.registry import (
    resolve_arch_extract_root,
    load_available_roles_from_registry,
    load_arch_style_registry,
    write_phase2_preflight,
    build_arch_styles_xml_from_registry,
)


def _check_numbering_module_needed(arch_styles_xml: str, needed_style_ids: list) -> None:
    """Raise if styles need numbering but numbering_importer is unavailable."""
    # Lightweight regex check: look for numId references in any needed style
    import re
    for sid in needed_style_ids:
        pat = r'<w:style[^>]*w:styleId="' + re.escape(sid) + r'"[^>]*>[\s\S]*?</w:style>'
        m = re.search(pat, arch_styles_xml)
        if m and '<w:numId' in m.group(0):
            raise ImportError(
                "numbering_importer module is not available but imported styles "
                f"require numbering definitions (e.g. style '{sid}'). "
                "Ensure numbering_importer.py is on the Python path."
            )


class DocxDecomposer:
    def __init__(self, docx_path):
        """
        Initialize the decomposer with a path to a .docx file.

        Args:
            docx_path: Path to the input .docx file
        """
        self.docx_path = Path(docx_path)
        self.extract_dir = None

    def extract(self, output_dir=None):
        """
        Extract the .docx file to a directory.

        Args:
            output_dir: Directory to extract to. If None, creates a directory
                    based on the docx filename.

        Returns:
            Path to the extraction directory
        """
        if output_dir is None:
            base_name = self.docx_path.stem
            output_dir = Path(f"{base_name}_extracted")
        else:
            output_dir = Path(output_dir)

        # Remove existing directory if it exists (OneDrive-safe)
        if output_dir.exists():
            import time
            import uuid
            max_retries = 3
            for attempt in range(max_retries):
                try:
                    shutil.rmtree(output_dir)
                    break
                except PermissionError as e:
                    if attempt < max_retries - 1:
                        print(f"Folder locked (OneDrive?), retrying in 2s... ({attempt + 1}/{max_retries})")
                        time.sleep(2)
                    else:
                        # Last resort: rename instead of delete
                        backup = output_dir.with_name(f"{output_dir.name}_old_{uuid.uuid4().hex[:8]}")
                        print(f"Cannot delete {output_dir}, renaming to {backup}")
                        output_dir.rename(backup)

        output_dir.mkdir(parents=True, exist_ok=True)

        # Extract the ZIP archive
        print(f"Extracting {self.docx_path} to {output_dir}...")
        with zipfile.ZipFile(self.docx_path, 'r') as zip_ref:
            zip_ref.extractall(output_dir)

        self.extract_dir = output_dir
        print(f"Extraction complete: {len(list(output_dir.rglob('*')))} items extracted")
        return output_dir


def main():
    import argparse
    import sys
    import os
    from pathlib import Path
    import json

    parser = argparse.ArgumentParser(description="DOCX decomposer + Phase 2 styling engine")
    parser.add_argument("docx_path", help="Path to input .docx")
    parser.add_argument("--extract-dir", default=None, help="Optional extraction directory")
    parser.add_argument("--output-docx", default=None, help="Output .docx path")
    parser.add_argument("--use-extract-dir", default=None, help="Use an existing extracted folder (skip extract/delete)")

    # Phase 2
    parser.add_argument("--phase2-arch-extract", help="Architect extracted folder")
    parser.add_argument("--phase2-classifications", help="Phase 2 LLM output JSON")
    parser.add_argument(
        "--phase2-build-bundle",
        action="store_true",
        help="Write Phase 2 slim bundle for LLM classification"
    )

    # Automated classification
    parser.add_argument(
        "--classify",
        action="store_true",
        help="Run full automated pipeline: extract -> classify -> apply -> format"
    )
    parser.add_argument("--api-key", default=None, help="Anthropic API key (default: ANTHROPIC_API_KEY env var)")
    parser.add_argument("--model", default="claude-opus-4-6", help="LLM model for classification")

    args = parser.parse_args()

    # Validate input path
    if not os.path.exists(args.docx_path):
        print(f"Error: File not found: {args.docx_path}")
        sys.exit(1)

    # Validate mutually exclusive flags
    if getattr(args, 'phase2_build_bundle', False) and getattr(args, 'classify', False):
        print("Error: --phase2-build-bundle and --classify are mutually exclusive.")
        print("  --phase2-build-bundle: build slim bundle for manual LLM step")
        print("  --classify: run full automated pipeline")
        sys.exit(1)

    input_docx_path = Path(args.docx_path)

    # Create decomposer
    decomposer = DocxDecomposer(args.docx_path)

    # Use existing extraction folder or extract fresh
    if args.use_extract_dir:
        extract_dir = Path(args.use_extract_dir)
        if not extract_dir.exists():
            print(f"Error: extract dir not found: {extract_dir}")
            sys.exit(1)
        decomposer.extract_dir = extract_dir
    else:
        extract_dir = decomposer.extract(output_dir=args.extract_dir)

    # -------------------------------
    # PHASE 2: BUILD SLIM BUNDLE
    # -------------------------------
    if args.phase2_build_bundle:
        available_roles = None
        if args.phase2_arch_extract:
            arch_path = Path(args.phase2_arch_extract)
            available_roles = load_available_roles_from_registry(arch_path)
            if available_roles:
                print(f"Available roles from architect template: {available_roles}")
            else:
                print("WARNING: Could not load architect registry, using all standard roles")

        bundle = build_phase2_slim_bundle(
            extract_dir,
            available_roles=available_roles
        )

        out_path = extract_dir / "phase2_slim_bundle.json"
        out_path.write_text(json.dumps(bundle, indent=2), encoding="utf-8")

        # Also write the prompts for convenience
        prompts_dir = extract_dir / "phase2_prompts"
        prompts_dir.mkdir(exist_ok=True)
        (prompts_dir / "master_prompt.txt").write_text(PHASE2_MASTER_PROMPT.strip(), encoding="utf-8")
        (prompts_dir / "run_instruction.txt").write_text(PHASE2_RUN_INSTRUCTION.strip(), encoding="utf-8")

        print(f"Phase 2 slim bundle written: {out_path}")
        print(f"Phase 2 prompts written to: {prompts_dir}")
        print("")
        print("NEXT STEPS:")
        print("1. Open your LLM (Claude/ChatGPT)")
        print("2. Paste the content of: master_prompt.txt")
        print("3. Paste the content of: phase2_slim_bundle.json")
        print("4. Paste the content of: run_instruction.txt")
        print("5. Save LLM JSON output as: phase2_classifications.json")
        print("   - Classify only unresolved paragraphs from bundle['paragraphs']")
        print("   - Apply step auto-merges deterministic classifications")
        print("   - Apply also accepts already-merged final classifications")
        print("6. Run Phase 2 apply:")
        print(f'   python docx_decomposer.py {args.docx_path} --phase2-arch-extract <arch_folder> --phase2-classifications phase2_classifications.json')
        return

    # -------------------------------
    # FULL AUTOMATED PIPELINE
    # -------------------------------
    if args.classify:
        if not args.phase2_arch_extract:
            print("Error: --classify requires --phase2-arch-extract")
            sys.exit(1)

        from core.llm_classifier import classify_target_document

        arch_path = Path(args.phase2_arch_extract)
        available_roles = load_available_roles_from_registry(arch_path)
        if not available_roles:
            print("ERROR: Could not load architect registry")
            sys.exit(1)

        print(f"Available roles: {available_roles}")

        # Build slim bundle
        bundle = build_phase2_slim_bundle(
            extract_dir,
            available_roles=available_roles
        )
        unresolved = len(bundle.get("paragraphs", []))
        deterministic = len(bundle.get("deterministic_classifications", []))
        print(f"Built slim bundle: {unresolved} unresolved + {deterministic} deterministic = {unresolved + deterministic} total")

        api_key = args.api_key or os.environ.get("ANTHROPIC_API_KEY")
        if unresolved > 0 and not api_key:
            print("Error: --classify requires --api-key or ANTHROPIC_API_KEY environment variable when unresolved paragraphs exist")
            sys.exit(1)

        # Classify via LLM
        print(f"Classifying with {args.model}...")
        classifications = classify_target_document(
            slim_bundle=bundle,
            available_roles=available_roles,
            api_key=api_key,
            model=args.model
        )

        # Save classifications for auditability
        classifications_path = extract_dir / "phase2_classifications.json"
        classifications_path.write_text(json.dumps(classifications, indent=2), encoding="utf-8")
        print(f"Classifications saved: {classifications_path}")

        # Now apply (fall through to the apply block below)
        args.phase2_classifications = str(classifications_path)

    # -------------------------------
    # PHASE 2: APPLY CLASSIFICATIONS
    # -------------------------------
    if args.phase2_arch_extract and args.phase2_classifications:
        from docx_patch import patch_docx

        log: List[str] = []

        arch_input = Path(args.phase2_arch_extract)
        arch_registry = load_arch_style_registry(arch_input)

        if arch_input.is_file() and arch_input.suffix.lower() == ".json":
            arch_root = resolve_arch_extract_root(arch_input.parent)
        else:
            arch_root = resolve_arch_extract_root(arch_input)

        classifications = json.loads(Path(args.phase2_classifications).read_text(encoding="utf-8"))

        available_roles = load_available_roles_from_registry(arch_root) or sorted(arch_registry.keys())
        validation_bundle = build_phase2_slim_bundle(
            extract_dir,
            available_roles=available_roles,
        )
        classifications = coerce_to_final_classifications(
            validation_bundle,
            classifications,
            available_roles,
        )
        normalized_path = extract_dir / "phase2_classifications.normalized.json"
        normalized_path.write_text(json.dumps(classifications, indent=2), encoding="utf-8")
        print(f"Normalized classifications saved: {normalized_path}")

        # Preflight report
        preflight_path = extract_dir / "phase2_preflight.json"
        preflight = write_phase2_preflight(
            extract_dir=extract_dir,
            arch_root=arch_root,
            arch_registry=arch_registry,
            classifications=classifications,
            out_path=preflight_path
        )
        print(f"Phase 2 preflight written: {preflight_path}")
        if preflight.get("unmapped_roles"):
            print(f"WARNING: Unmapped roles: {preflight['unmapped_roles']}")

        # Apply formatting environment
        arch_template_registry_path = arch_root / "arch_template_registry.json"
        if not arch_template_registry_path.exists():
            raise FileNotFoundError(
                f"arch_template_registry.json not found at {arch_template_registry_path}. "
                "Phase 2 cannot proceed without the template registry."
            )
        env_registry = json.loads(arch_template_registry_path.read_text(encoding="utf-8"))

        # Preflight contract validation — abort before any mutation
        from core.registry import preflight_validate_registries
        preflight_errors = preflight_validate_registries(arch_registry, env_registry)
        if preflight_errors:
            error_report = "\n".join(f"  - {e}" for e in preflight_errors)
            raise ValueError(
                f"Phase 2 preflight validation failed with {len(preflight_errors)} error(s):\n"
                f"{error_report}\n"
                "Fix the contract files and retry."
            )

        apply_environment_to_target(
            target_extract_dir=extract_dir,
            registry=env_registry,
            log=log
        )
        print(f"Applied environment from: {arch_template_registry_path}")

        # Build synthetic styles.xml from registry (no disk dependency on arch extracted folder)
        arch_styles_xml = build_arch_styles_xml_from_registry(env_registry)

        # Import styles
        used_roles = {
            item.get("csi_role")
            for item in classifications.get("classifications", [])
            if isinstance(item, dict) and isinstance(item.get("csi_role"), str)
        }
        needed_style_ids = sorted({arch_registry[r] for r in used_roles if r in arch_registry})

        # Import numbering definitions
        style_numid_remap = {}
        if HAS_NUMBERING_IMPORTER and arch_template_registry_path.exists():
            log.append("")
            log.append("=" * 60)
            log.append("IMPORTING NUMBERING DEFINITIONS")
            log.append("=" * 60)

            style_numid_remap = import_numbering(
                target_extract_dir=extract_dir,
                arch_template_registry=env_registry,
                arch_styles_xml=arch_styles_xml,
                style_ids_to_import=needed_style_ids,
                log=log
            )
        elif not HAS_NUMBERING_IMPORTER:
            # Check whether numbering is actually needed before silently skipping
            _check_numbering_module_needed(arch_styles_xml, needed_style_ids)

        log.append("")
        log.append("=" * 60)
        log.append("IMPORTING STYLE DEFINITIONS")
        log.append("=" * 60)

        import_arch_styles_into_target(
            target_extract_dir=extract_dir,
            arch_styles_xml=arch_styles_xml,
            needed_style_ids=needed_style_ids,
            log=log,
            style_numid_remap=style_numid_remap
        )

        if not needed_style_ids:
            log.append("No architect styles needed for this doc (no mapped roles used).")

        # Snapshot invariants after environment baseline is established
        # (including managed sectPr page-layout sync) and before classification writes.
        snap = snapshot_stability(extract_dir)

        apply_phase2_classifications(
            extract_dir=extract_dir,
            classifications=classifications,
            arch_style_registry=arch_registry,
            log=log
        )

        # Stability checks
        verify_stability(extract_dir, snap)

        # Write formatted DOCX
        output_docx_path = Path(args.output_docx) if args.output_docx else (
            input_docx_path.with_name(input_docx_path.stem + "_PHASE2_FORMATTED.docx")
        )

        replacements = {
            "word/document.xml": (extract_dir / "word" / "document.xml").read_bytes(),
            "word/styles.xml":   (extract_dir / "word" / "styles.xml").read_bytes(),
        }

        # Add environment parts if they were modified
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

        patch_docx(
            src_docx=input_docx_path,
            out_docx=output_docx_path,
            replacements=replacements,
        )

        # Optional invariant verification
        try:
            from phase2_invariants import verify_phase2_invariants
            new_doc_xml_bytes = (extract_dir / "word" / "document.xml").read_bytes()
            verify_phase2_invariants(
                src_docx=input_docx_path,
                new_document_xml=new_doc_xml_bytes,
                new_docx=output_docx_path,
                arch_template_registry=env_registry,
            )
        except ModuleNotFoundError:
            pass

        issues_path = extract_dir / "phase2_issues.log"
        issues_path.write_text("\n".join(log) + "\n", encoding="utf-8")

        print(f"Phase 2 output written: {output_docx_path}")
        print(f"Phase 2 log written:    {issues_path}")
        return

    # -------------------------------
    # DEFAULT: no action specified
    # -------------------------------
    print("No action specified.")
    print("Use one of:")
    print("  --phase2-build-bundle")
    print("  --phase2-arch-extract <arch_extract> --phase2-classifications <json> [--output-docx out.docx]")
    print("  --classify --phase2-arch-extract <arch_extract> [--api-key KEY]")
    print(f"Extracted to: {extract_dir}")


if __name__ == "__main__":
    main()
