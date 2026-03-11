#!/usr/bin/env python3
"""
Phase 2 MEP Specification Styling Engine — GUI

Tkinter-based graphical interface for applying architect-defined CSI
paragraph styles to MEP specification documents. Supports single-file
and batch processing modes.
"""

import os
import sys
import json
import threading
import subprocess
from pathlib import Path
from typing import Optional, List

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext


class Phase2GUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Phase 2 — MEP Spec Styling Engine")
        self.root.geometry("780x680")
        self.root.minsize(600, 500)

        self.processing = False
        self.output_path: Optional[Path] = None
        self.log_path: Optional[Path] = None

        self._build_ui()

    def _build_ui(self):
        # Main frame with padding
        main = ttk.Frame(self.root, padding=10)
        main.pack(fill=tk.BOTH, expand=True)

        # ── Input Section ────────────────────────────────────────────────
        input_frame = ttk.LabelFrame(main, text="Input", padding=8)
        input_frame.pack(fill=tk.X, pady=(0, 8))

        # Mode toggle
        mode_frame = ttk.Frame(input_frame)
        mode_frame.pack(fill=tk.X, pady=(0, 6))
        self.batch_var = tk.BooleanVar(value=False)
        ttk.Radiobutton(mode_frame, text="Single File", variable=self.batch_var,
                        value=False, command=self._on_mode_change).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Radiobutton(mode_frame, text="Batch Mode (folder)", variable=self.batch_var,
                        value=True, command=self._on_mode_change).pack(side=tk.LEFT)

        # Target DOCX / folder
        row1 = ttk.Frame(input_frame)
        row1.pack(fill=tk.X, pady=2)
        self.target_label = ttk.Label(row1, text="Target Spec (.docx):", width=22, anchor="w")
        self.target_label.pack(side=tk.LEFT)
        self.target_var = tk.StringVar()
        ttk.Entry(row1, textvariable=self.target_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
        self.target_btn = ttk.Button(row1, text="Browse...", command=self._browse_target)
        self.target_btn.pack(side=tk.RIGHT)

        # Architect template folder
        row2 = ttk.Frame(input_frame)
        row2.pack(fill=tk.X, pady=2)
        ttk.Label(row2, text="Architect Template Folder:", width=22, anchor="w").pack(side=tk.LEFT)
        self.arch_var = tk.StringVar()
        ttk.Entry(row2, textvariable=self.arch_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
        ttk.Button(row2, text="Browse...", command=self._browse_arch).pack(side=tk.RIGHT)

        # API key
        row3 = ttk.Frame(input_frame)
        row3.pack(fill=tk.X, pady=2)
        ttk.Label(row3, text="Anthropic API Key:", width=22, anchor="w").pack(side=tk.LEFT)
        self.api_key_var = tk.StringVar(value=os.environ.get("ANTHROPIC_API_KEY", ""))
        ttk.Entry(row3, textvariable=self.api_key_var, show="*").pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Discipline
        row4 = ttk.Frame(input_frame)
        row4.pack(fill=tk.X, pady=2)
        ttk.Label(row4, text="Discipline:", width=22, anchor="w").pack(side=tk.LEFT)
        self.discipline_var = tk.StringVar(value="mechanical")
        disc_combo = ttk.Combobox(row4, textvariable=self.discipline_var,
                                  values=["mechanical", "plumbing", "fire protection"],
                                  state="readonly", width=20)
        disc_combo.pack(side=tk.LEFT)

        # ── Action Section ───────────────────────────────────────────────
        action_frame = ttk.Frame(main)
        action_frame.pack(fill=tk.X, pady=(0, 8))

        self.run_btn = ttk.Button(action_frame, text="Run Phase 2", command=self._run)
        self.run_btn.pack(side=tk.LEFT, padx=(0, 8))

        self.progress = ttk.Progressbar(action_frame, mode="indeterminate", length=200)
        self.progress.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))

        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(action_frame, textvariable=self.status_var, width=30, anchor="w").pack(side=tk.RIGHT)

        # ── Log Section ──────────────────────────────────────────────────
        log_frame = ttk.LabelFrame(main, text="Log", padding=4)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        self.log_text = scrolledtext.ScrolledText(log_frame, height=15, state=tk.DISABLED,
                                                   font=("Consolas", 9), wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # ── Bottom Buttons ───────────────────────────────────────────────
        btn_frame = ttk.Frame(main)
        btn_frame.pack(fill=tk.X)

        self.open_output_btn = ttk.Button(btn_frame, text="Open Output DOCX",
                                           command=self._open_output, state=tk.DISABLED)
        self.open_output_btn.pack(side=tk.LEFT, padx=(0, 8))

        self.open_log_btn = ttk.Button(btn_frame, text="Open Log",
                                        command=self._open_log, state=tk.DISABLED)
        self.open_log_btn.pack(side=tk.LEFT)

    def _on_mode_change(self):
        if self.batch_var.get():
            self.target_label.config(text="Target Folder:")
        else:
            self.target_label.config(text="Target Spec (.docx):")
        self.target_var.set("")

    def _browse_target(self):
        if self.batch_var.get():
            path = filedialog.askdirectory(title="Select folder with .docx files")
        else:
            path = filedialog.askopenfilename(
                title="Select Target Spec",
                filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")]
            )
        if path:
            self.target_var.set(path)

    def _browse_arch(self):
        path = filedialog.askdirectory(title="Select Architect Template Folder")
        if path:
            self.arch_var.set(path)

    def _log(self, msg: str):
        self.root.after(0, self._append_log, msg)

    def _append_log(self, msg: str):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)

    def _set_status(self, status: str):
        self.root.after(0, lambda: self.status_var.set(status))

    def _validate_inputs(self) -> bool:
        if not self.target_var.get():
            messagebox.showerror("Error", "Please select a target spec or folder.")
            return False
        if not self.arch_var.get():
            messagebox.showerror("Error", "Please select an architect template folder.")
            return False
        if not self.api_key_var.get():
            messagebox.showerror("Error", "Please provide an Anthropic API key.")
            return False

        arch_path = Path(self.arch_var.get())
        if not (arch_path / "arch_style_registry.json").exists():
            messagebox.showerror("Error",
                f"arch_style_registry.json not found in {arch_path}.\n"
                "Make sure you selected the correct architect template folder.")
            return False

        return True

    def _run(self):
        if self.processing:
            return
        if not self._validate_inputs():
            return

        self.processing = True
        self.run_btn.config(state=tk.DISABLED)
        self.open_output_btn.config(state=tk.DISABLED)
        self.open_log_btn.config(state=tk.DISABLED)
        self.progress.start(10)

        # Clear log
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete("1.0", tk.END)
        self.log_text.config(state=tk.DISABLED)

        thread = threading.Thread(target=self._process, daemon=True)
        thread.start()

    def _process(self):
        try:
            if self.batch_var.get():
                self._process_batch()
            else:
                self._process_single(Path(self.target_var.get()))
        except Exception as e:
            self._log(f"\nERROR: {e}")
            self._set_status("Failed")
            import traceback
            self._log(traceback.format_exc())
        finally:
            self.root.after(0, self._finish_processing)

    def _process_single(self, docx_path: Path) -> Optional[Path]:
        """Process a single DOCX file. Returns output path on success."""
        self._set_status(f"Processing: {docx_path.name}")
        self._log(f"Processing: {docx_path}")

        from docx_decomposer import DocxDecomposer
        from docx_patch import patch_docx
        from arch_env_applier import apply_environment_to_target
        from core.classification import (
            PHASE2_MASTER_PROMPT, PHASE2_RUN_INSTRUCTION,
            build_phase2_slim_bundle, apply_phase2_classifications,
        )
        from core.stability import snapshot_stability, verify_stability
        from core.style_import import import_arch_styles_into_target
        from core.registry import (
            resolve_arch_extract_root, load_available_roles_from_registry,
            load_arch_style_registry, write_phase2_preflight,
        )
        from core.llm_classifier import classify_target_document

        try:
            from numbering_importer import import_numbering
            has_numbering = True
        except ImportError:
            has_numbering = False

        arch_path = Path(self.arch_var.get())
        api_key = self.api_key_var.get()
        discipline = self.discipline_var.get()

        log: List[str] = []

        # Extract
        self._log("  Extracting DOCX...")
        decomposer = DocxDecomposer(str(docx_path))
        extract_dir = decomposer.extract(output_dir=Path("output") / f"{docx_path.stem}_extracted")

        # Load registry
        arch_registry = load_arch_style_registry(arch_path)
        arch_root = resolve_arch_extract_root(arch_path)
        available_roles = load_available_roles_from_registry(arch_path)
        if not available_roles:
            raise ValueError("Could not load architect registry")
        self._log(f"  Available roles: {available_roles}")

        # Build slim bundle
        self._log("  Building slim bundle...")
        bundle = build_phase2_slim_bundle(extract_dir, discipline, available_roles=available_roles)
        self._log(f"  Bundle: {len(bundle.get('paragraphs', []))} paragraphs")

        # Classify
        self._log(f"  Classifying with LLM...")
        classifications = classify_target_document(
            slim_bundle=bundle,
            available_roles=available_roles,
            api_key=api_key,
        )

        # Save classifications
        classifications_path = extract_dir / "phase2_classifications.json"
        classifications_path.write_text(json.dumps(classifications, indent=2), encoding="utf-8")
        self._log(f"  Classifications saved: {classifications_path}")

        # Apply environment
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
                f"Preflight validation failed ({len(preflight_errors)} error(s)):\n{error_report}"
            )

        apply_environment_to_target(target_extract_dir=extract_dir, registry=env_registry, log=log)
        self._log("  Applied environment")

        # Build synthetic styles.xml from registry (no disk dependency on arch extracted folder)
        from core.registry import build_arch_styles_xml_from_registry
        arch_styles_xml = build_arch_styles_xml_from_registry(env_registry)

        # Import styles
        used_roles = {
            item.get("csi_role")
            for item in classifications.get("classifications", [])
            if isinstance(item, dict) and isinstance(item.get("csi_role"), str)
        }
        needed_style_ids = sorted({arch_registry[r] for r in used_roles if r in arch_registry})

        # Import numbering
        style_numid_remap = {}
        if has_numbering:
            try:
                style_numid_remap = import_numbering(
                    target_extract_dir=extract_dir,
                    arch_template_registry=env_registry,
                    arch_styles_xml=arch_styles_xml,
                    style_ids_to_import=needed_style_ids,
                    log=log
                )
            except Exception as e:
                log.append(f"WARNING: Numbering import failed: {e}")

        import_arch_styles_into_target(
            target_extract_dir=extract_dir,
            arch_styles_xml=arch_styles_xml,
            needed_style_ids=needed_style_ids,
            log=log,
            style_numid_remap=style_numid_remap
        )
        self._log(f"  Imported {len(needed_style_ids)} styles")

        # Snapshot + apply + verify
        snap = snapshot_stability(extract_dir)
        apply_phase2_classifications(
            extract_dir=extract_dir,
            classifications=classifications,
            arch_style_registry=arch_registry,
            log=log
        )
        verify_stability(extract_dir, snap)
        self._log("  Applied classifications, stability verified")

        # Patch output
        output_docx_path = Path("output") / (docx_path.stem + "_PHASE2_FORMATTED.docx")
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

        patch_docx(src_docx=docx_path, out_docx=output_docx_path, replacements=replacements)

        # Coverage
        total = len(bundle.get("paragraphs", []))
        classified = len(classifications.get("classifications", []))
        coverage = (classified / total * 100) if total > 0 else 0

        self._log(f"  Output: {output_docx_path}")
        self._log(f"  Coverage: {classified}/{total} ({coverage:.1f}%)")

        issues_path = extract_dir / "phase2_issues.log"
        issues_path.write_text("\n".join(log) + "\n", encoding="utf-8")

        self.output_path = output_docx_path
        self.log_path = issues_path
        return output_docx_path

    def _process_batch(self):
        """Process all .docx files in a folder."""
        folder = Path(self.target_var.get())
        docx_files = sorted(folder.glob("*.docx"))

        if not docx_files:
            self._log(f"No .docx files found in {folder}")
            self._set_status("No files found")
            return

        self._log(f"Batch mode: {len(docx_files)} files in {folder}")
        self._log("")

        results = []
        for i, docx_path in enumerate(docx_files, 1):
            self._set_status(f"Batch {i}/{len(docx_files)}: {docx_path.name}")
            try:
                output = self._process_single(docx_path)
                results.append((docx_path.name, "OK", output))
                self._log("")
            except Exception as e:
                results.append((docx_path.name, f"FAILED: {e}", None))
                self._log(f"  FAILED: {e}\n")

        # Summary
        self._log("=" * 60)
        self._log("BATCH SUMMARY")
        self._log("=" * 60)
        ok_count = sum(1 for _, s, _ in results if s == "OK")
        for name, status, _ in results:
            self._log(f"  {name}: {status}")
        self._log(f"\n  {ok_count}/{len(results)} files processed successfully")
        self._set_status(f"Batch done: {ok_count}/{len(results)} OK")

    def _finish_processing(self):
        self.processing = False
        self.progress.stop()
        self.run_btn.config(state=tk.NORMAL)
        if self.output_path:
            self.open_output_btn.config(state=tk.NORMAL)
        if self.log_path:
            self.open_log_btn.config(state=tk.NORMAL)
        if not self.batch_var.get() and self.output_path:
            self._set_status("Done")

    def _open_output(self):
        if self.output_path and self.output_path.exists():
            self._open_file(self.output_path)

    def _open_log(self):
        if self.log_path and self.log_path.exists():
            self._open_file(self.log_path)

    @staticmethod
    def _open_file(path: Path):
        """Cross-platform file open."""
        path = str(path)
        if sys.platform == "win32":
            os.startfile(path)
        elif sys.platform == "darwin":
            subprocess.run(["open", path])
        else:
            subprocess.run(["xdg-open", path])


def main():
    root = tk.Tk()
    Phase2GUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
