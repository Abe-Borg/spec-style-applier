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
import re
import threading
import subprocess
from pathlib import Path
from typing import Optional, List

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext


HOW_IT_WORKS_TEXT = """# How It Works

## The Problem It Solves

When a project's architect provides a specification template, that template defines a specific visual style for every level of the CSI spec hierarchy — section headers, part headings, article numbers, paragraphs, subparagraphs, and so on. Each level has its own font, size, spacing, and indentation.

MEP (mechanical, plumbing, fire protection) specifications are written separately, usually from an in-house master spec or a purchased library like MasterSpec. These documents have their own formatting — sometimes minimal, sometimes inconsistent — that rarely matches what the architect specified.

Manually reformatting a spec section to match an architect's template means selecting every heading, every article number, every paragraph line, and applying the correct Word paragraph style one by one. On a single spec section that might be manageable. Across a full set of MEP sections with dozens of CSI divisions, it becomes a significant hours-long task, and it's easy to miss paragraphs or apply the wrong style.

This tool automates that reformatting process.

---

## What the Tool Does, Step by Step

### Step 1 — Read the Architect's Template

Before this tool can do anything with your spec, someone on the team needs to run **Phase 1** on the architect's Word template. Phase 1 reads that template and extracts all the paragraph style definitions — fonts, spacing, numbering, indentation — and saves them to two JSON files in a folder. That folder is what you point this tool at.

You only need to run Phase 1 once per architect template. The extracted folder can be reused for every spec section on that project.

### Step 2 — Open and Read Your Spec Document

The tool opens your MEP spec `.docx` file and reads every paragraph. It does this at the raw XML level — the underlying structure that Word uses internally — rather than through Word itself. This lets it work in the background without opening Word.

### Step 3 — Classify Every Paragraph

This is the core step. The tool needs to know what *kind* of content each paragraph is before it can apply the right style. Is a given line a PART heading? An article number like `1.01`? A lettered paragraph like `A.`? A numbered subparagraph?

The tool uses two methods:

**Deterministic classification** handles the easy cases automatically. A line that starts with `PART 1`, `PART 2`, or `PART 3` is always a PART heading. A line matching the pattern `1.01`, `2.03`, etc. is always an article number. Lettered list items (`A.`, `B.`) and numbered items (`1.`, `2.`) are identified by their formatting patterns. These don't need AI — the patterns are unambiguous.

**AI classification** handles everything else. Paragraphs that aren't deterministically identifiable — body content, section headers, ambiguous lines — are sent to an AI model (Claude) via the Anthropic API. The AI reads each paragraph in context and assigns it a CSI role from the list of roles available in the architect's template.

This two-step approach means the AI only has to work on the paragraphs that actually need judgment. The deterministic ones never leave your machine.

### Step 4 — Apply the Styles

With every paragraph classified, the tool applies the matching Word paragraph style from the architect's template to each paragraph. A paragraph classified as `ARTICLE` gets the architect's Article style applied. A paragraph classified as `PARAGRAPH` gets the architect's Paragraph style. And so on.

The tool also imports the architect's numbering definitions (the list formatting rules) and formatting environment (fonts, spacing defaults, theme) into the document, so the imported styles render exactly as they would in the architect's own template.

### Step 5 — Output a Formatted Document

The result is a new `.docx` file — your original spec with all formatting replaced by the architect's styles. The content of every paragraph is untouched. Only the styling changes. Headers, footers, and page layout from your original document are preserved.

---

## What the API Key Is For

The AI classification step (Step 3) requires a call to Anthropic's API, which is a paid cloud service. The API key in the interface authenticates those calls. The key is never stored by this tool — it's only used during the active processing run.

If your document's paragraphs can all be classified deterministically (which is rare for a full spec section), the API call is skipped entirely and no key is needed.

---

## Batch Mode

Batch mode runs the same process on every `.docx` file in a folder, one at a time. This is useful when you have a full set of spec sections to reformat for a project. Each file is processed independently and gets its own formatted output.

---

## What This Tool Does Not Do

- It does not change any spec content — no words, no requirements, no product references
- It does not open Word
- It does not modify your original file — it always creates a new output document
- It does not create or design styles — it only applies styles that already exist in the architect's template
- It does not format tables, embedded objects, or content inside text boxes
"""


HOW_TO_USE_TEXT = """# How to Use

## Before You Start

You need two things ready before running this tool:

1. **The architect's extracted template folder** — a folder containing `arch_style_registry.json` and `arch_template_registry.json`. These are produced by running Phase 1 on the architect's Word template. Get this folder from whoever ran Phase 1 on your project, or run it yourself. You only need to do this once per architect template.

2. **An Anthropic API key** — required for the AI classification step. Get this from your firm's Anthropic account or your own at [console.anthropic.com](https://console.anthropic.com). If your `ANTHROPIC_API_KEY` environment variable is already set on your machine, the field will be pre-filled.

---

## Single File — Step by Step

**1. Select the target spec**
Click **Browse...** next to *Target Spec (.docx)* and select the MEP spec section you want to reformat.

**2. Select the architect template folder**
Click **Browse...** next to *Architect Template Folder* and select the folder containing the two registry JSON files.

**3. Enter your API key**
Paste your Anthropic API key into the *Anthropic API Key* field. It will appear masked.

**4. Set the discipline**
Select *mechanical* or *plumbing* from the dropdown to match the spec you're formatting.

**5. Click Run Phase 2**
The progress bar will animate while the tool works. The log window shows live status. Classification typically takes 15–60 seconds depending on document length.

**6. Open the output**
When complete, click **Open Output DOCX** to open the formatted document directly. The output file is saved in the same folder as your input file with `_PHASE2_FORMATTED` appended to the name.

If anything went wrong, click **Open Log** to see a detailed record of what the tool did and where it stopped.

---

## Batch Mode

**1.** Select *Batch Mode (folder)* at the top of the Input section.

**2.** Click **Browse...** and select a folder containing the `.docx` spec files you want to process.

**3.** Fill in the architect template folder, API key, and discipline as above.

**4.** Click **Run Phase 2**. The tool processes each file in sequence. The log shows per-file results and a summary at the end.

Output files are placed in the selected output folder (defaults to the input folder), each named `<original_filename>_PHASE2_FORMATTED.docx`.

---

## Notes

- Your **original file is never modified**. The tool always writes a new output document.
- Run the tool again on the same file any time — it will overwrite the previous output.
- If you see a preflight validation error, the architect's template folder may be from a different Phase 1 run than expected. Confirm you have the right folder for this project.
- The API key field is masked and is not saved to disk by this tool.
"""


def _check_numbering_module_needed(arch_styles_xml: str, needed_style_ids: list) -> None:
    """Raise if styles need numbering but numbering_importer is unavailable."""
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


class Phase2GUI:
    def __init__(self, root: tk.Tk):
        self.colors = {
            "bg_dark": "#0D0D0D",
            "bg_card": "#1A1A1A",
            "bg_input": "#252525",
            "border": "#333333",
            "text_primary": "#FFFFFF",
            "text_secondary": "#B0B0B0",
            "text_muted": "#707070",
            "accent": "#3B82F6",
            "accent_hover": "#2563EB",
            "success": "#22C55E",
            "error": "#EF4444",
        }

        self.root = root
        self.root.title("Phase 2 — MEP Spec Styling Engine")
        self.root.geometry("900x850")
        self.root.minsize(750, 700)
        self.root.configure(bg=self.colors["bg_dark"])

        self.processing = False
        self.output_path: Optional[Path] = None
        self.log_path: Optional[Path] = None

        self._build_ui()

    def _build_ui(self):
        self._apply_styles()

        main = tk.Frame(self.root, bg=self.colors["bg_dark"], padx=24, pady=24)
        main.pack(fill=tk.BOTH, expand=True)

        header = tk.Frame(main, bg=self.colors["bg_dark"])
        header.pack(fill=tk.X, pady=(0, 14))

        header_left = tk.Frame(header, bg=self.colors["bg_dark"])
        header_left.pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Label(
            header_left,
            text="PHASE 2 MEP STYLING ENGINE",
            bg=self.colors["bg_dark"],
            fg=self.colors["text_primary"],
            font=("Segoe UI", 20, "bold"),
        ).pack(anchor="w")
        tk.Label(
            header_left,
            text="Apply architect style language to target specs",
            bg=self.colors["bg_dark"],
            fg=self.colors["text_secondary"],
            font=("Segoe UI", 11),
        ).pack(anchor="w", pady=(2, 0))

        header_right = tk.Frame(header, bg=self.colors["bg_dark"])
        header_right.pack(side=tk.RIGHT, anchor="ne")
        tk.Button(
            header_right,
            text="How It Works",
            command=lambda: self._show_info_popup("How It Works", HOW_IT_WORKS_TEXT),
            **self.secondary_btn_style,
        ).pack(side=tk.LEFT, padx=(0, 8))
        tk.Button(
            header_right,
            text="How to Use",
            command=lambda: self._show_info_popup("How to Use", HOW_TO_USE_TEXT),
            **self.secondary_btn_style,
        ).pack(side=tk.LEFT)

        input_frame = tk.Frame(main, bg=self.colors["bg_card"], padx=16, pady=12)
        input_frame.pack(fill=tk.X, pady=(0, 12))
        tk.Label(input_frame, text="INPUTS", bg=self.colors["bg_card"], fg=self.colors["text_muted"],
                 font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 8))

        mode_frame = tk.Frame(input_frame, bg=self.colors["bg_card"])
        mode_frame.pack(fill=tk.X, pady=(0, 8))
        self.batch_var = tk.BooleanVar(value=False)
        tk.Radiobutton(
            mode_frame,
            text="Single File",
            variable=self.batch_var,
            value=False,
            command=self._on_mode_change,
            bg=self.colors["bg_card"],
            activebackground=self.colors["bg_card"],
            fg=self.colors["text_secondary"],
            activeforeground=self.colors["text_primary"],
            selectcolor=self.colors["bg_input"],
            font=("Segoe UI", 10),
            highlightthickness=0,
        ).pack(side=tk.LEFT, padx=(0, 12))
        tk.Radiobutton(
            mode_frame,
            text="Batch Mode (folder)",
            variable=self.batch_var,
            value=True,
            command=self._on_mode_change,
            bg=self.colors["bg_card"],
            activebackground=self.colors["bg_card"],
            fg=self.colors["text_secondary"],
            activeforeground=self.colors["text_primary"],
            selectcolor=self.colors["bg_input"],
            font=("Segoe UI", 10),
            highlightthickness=0,
        ).pack(side=tk.LEFT)

        # Target DOCX / folder
        row1 = tk.Frame(input_frame, bg=self.colors["bg_card"])
        row1.pack(fill=tk.X, pady=2)
        self.target_label = tk.Label(row1, text="Target Spec (.docx):", width=22, anchor="w",
                                     bg=self.colors["bg_card"], fg=self.colors["text_secondary"],
                                     font=("Segoe UI", 11))
        self.target_label.pack(side=tk.LEFT)
        self.target_var = tk.StringVar()
        tk.Entry(
            row1,
            textvariable=self.target_var,
            bg=self.colors["bg_input"],
            fg=self.colors["text_primary"],
            insertbackground=self.colors["text_primary"],
            relief=tk.FLAT,
            highlightthickness=1,
            highlightbackground=self.colors["border"],
            highlightcolor=self.colors["accent"],
            font=("Consolas", 11),
        ).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6), ipady=7)
        self.target_btn = tk.Button(row1, text="Browse...", command=self._browse_target, **self.secondary_btn_style)
        self.target_btn.pack(side=tk.RIGHT)

        # Architect template folder
        row2 = tk.Frame(input_frame, bg=self.colors["bg_card"])
        row2.pack(fill=tk.X, pady=2)
        tk.Label(row2, text="Architect Template Folder:", width=22, anchor="w",
                 bg=self.colors["bg_card"], fg=self.colors["text_secondary"], font=("Segoe UI", 11)).pack(side=tk.LEFT)
        self.arch_var = tk.StringVar()
        tk.Entry(
            row2,
            textvariable=self.arch_var,
            bg=self.colors["bg_input"],
            fg=self.colors["text_primary"],
            insertbackground=self.colors["text_primary"],
            relief=tk.FLAT,
            highlightthickness=1,
            highlightbackground=self.colors["border"],
            highlightcolor=self.colors["accent"],
            font=("Consolas", 11),
        ).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6), ipady=7)
        tk.Button(row2, text="Browse...", command=self._browse_arch, **self.secondary_btn_style).pack(side=tk.RIGHT)

        # API key
        row3 = tk.Frame(input_frame, bg=self.colors["bg_card"])
        row3.pack(fill=tk.X, pady=2)
        tk.Label(row3, text="Anthropic API Key:", width=22, anchor="w",
                 bg=self.colors["bg_card"], fg=self.colors["text_secondary"], font=("Segoe UI", 11)).pack(side=tk.LEFT)
        self.api_key_var = tk.StringVar(value=os.environ.get("ANTHROPIC_API_KEY", ""))
        tk.Entry(
            row3,
            textvariable=self.api_key_var,
            show="•",
            bg=self.colors["bg_input"],
            fg=self.colors["text_primary"],
            insertbackground=self.colors["text_primary"],
            relief=tk.FLAT,
            highlightthickness=1,
            highlightbackground=self.colors["border"],
            highlightcolor=self.colors["accent"],
            font=("Consolas", 11),
        ).pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=7)

        # Output folder
        row4 = tk.Frame(input_frame, bg=self.colors["bg_card"])
        row4.pack(fill=tk.X, pady=2)
        tk.Label(row4, text="Output Folder:", width=22, anchor="w",
                 bg=self.colors["bg_card"], fg=self.colors["text_secondary"], font=("Segoe UI", 11)).pack(side=tk.LEFT)
        self.output_dir_var = tk.StringVar()
        tk.Entry(
            row4,
            textvariable=self.output_dir_var,
            bg=self.colors["bg_input"],
            fg=self.colors["text_primary"],
            insertbackground=self.colors["text_primary"],
            relief=tk.FLAT,
            highlightthickness=1,
            highlightbackground=self.colors["border"],
            highlightcolor=self.colors["accent"],
            font=("Consolas", 11),
        ).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6), ipady=7)
        tk.Button(row4, text="Browse...", command=self._browse_output_dir, **self.secondary_btn_style).pack(side=tk.RIGHT)

        # ── Action Section ───────────────────────────────────────────────
        action_frame = tk.Frame(main, bg=self.colors["bg_dark"])
        action_frame.pack(fill=tk.X, pady=(0, 12))

        self.run_btn = tk.Button(action_frame, text="Run Phase 2", command=self._run, **self.primary_btn_style)
        self.run_btn.pack(side=tk.LEFT, padx=(0, 12), ipadx=18, ipady=10)

        self.run_btn.bind("<Enter>", lambda _e: self.run_btn.config(bg=self.colors["accent_hover"]))
        self.run_btn.bind("<Leave>", lambda _e: self.run_btn.config(bg=self.colors["accent"]))

        self.progress = ttk.Progressbar(action_frame, mode="indeterminate", length=200, style=self.progress_style)
        self.progress.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))

        self.status_var = tk.StringVar(value="Ready")
        tk.Label(action_frame, textvariable=self.status_var, width=30, anchor="w",
                 bg=self.colors["bg_dark"], fg=self.colors["text_secondary"],
                 font=("Segoe UI", 11)).pack(side=tk.RIGHT)

        # ── Log Section ──────────────────────────────────────────────────
        log_frame = tk.Frame(main, bg=self.colors["bg_card"], padx=16, pady=12)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 12))
        tk.Label(log_frame, text="ACTIVITY LOG", bg=self.colors["bg_card"], fg=self.colors["text_muted"],
                 font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 8))

        self.log_text = scrolledtext.ScrolledText(log_frame, height=15, state=tk.DISABLED,
                                                   font=("Consolas", 11), wrap=tk.WORD,
                                                   bg=self.colors["bg_input"],
                                                   fg=self.colors["text_secondary"],
                                                   insertbackground=self.colors["text_primary"],
                                                   relief=tk.FLAT,
                                                   highlightthickness=1,
                                                   highlightbackground=self.colors["border"],
                                                   highlightcolor=self.colors["accent"])
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # ── Bottom Buttons ───────────────────────────────────────────────
        btn_frame = tk.Frame(main, bg=self.colors["bg_dark"])
        btn_frame.pack(fill=tk.X)

        self.open_output_btn = tk.Button(btn_frame, text="Open Output DOCX",
                                         command=self._open_output, state=tk.DISABLED,
                                         **self.secondary_btn_style)
        self.open_output_btn.pack(side=tk.LEFT, padx=(0, 8))

        self.open_log_btn = tk.Button(btn_frame, text="Open Log",
                                      command=self._open_log, state=tk.DISABLED,
                                      **self.secondary_btn_style)
        self.open_log_btn.pack(side=tk.LEFT)

    def _apply_styles(self):
        self.primary_btn_style = {
            "bg": self.colors["accent"],
            "fg": self.colors["text_primary"],
            "activebackground": self.colors["accent_hover"],
            "activeforeground": self.colors["text_primary"],
            "relief": tk.FLAT,
            "bd": 0,
            "font": ("Segoe UI", 12, "bold"),
            "cursor": "hand2",
            "disabledforeground": self.colors["text_muted"],
        }
        self.secondary_btn_style = {
            "bg": self.colors["bg_input"],
            "fg": self.colors["text_secondary"],
            "activebackground": self.colors["border"],
            "activeforeground": self.colors["text_primary"],
            "relief": tk.FLAT,
            "bd": 1,
            "font": ("Segoe UI", 10),
            "cursor": "hand2",
            "disabledforeground": self.colors["text_muted"],
        }

        style = ttk.Style()
        style.theme_use("default")
        style.configure(
            "Dark.Horizontal.TProgressbar",
            background=self.colors["accent"],
            troughcolor=self.colors["bg_input"],
            bordercolor=self.colors["border"],
            lightcolor=self.colors["accent"],
            darkcolor=self.colors["accent"],
            thickness=4,
        )
        self.progress_style = "Dark.Horizontal.TProgressbar"

    def _on_mode_change(self):
        if self.batch_var.get():
            self.target_label.config(text="Target Folder:")
        else:
            self.target_label.config(text="Target Spec (.docx):")
        self.target_var.set("")
        self.output_dir_var.set("")

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
            if self.batch_var.get():
                default_output = path
            else:
                default_output = str(Path(path).parent)
            self.output_dir_var.set(default_output)

    def _browse_arch(self):
        path = filedialog.askdirectory(title="Select Architect Template Folder")
        if path:
            self.arch_var.set(path)

    def _browse_output_dir(self):
        path = filedialog.askdirectory(title="Select Output Folder")
        if path:
            self.output_dir_var.set(path)

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
        if not self.output_dir_var.get():
            messagebox.showerror("Error", "Please select an output folder.")
            return False
        arch_path = Path(self.arch_var.get())
        if not (arch_path / "arch_style_registry.json").exists():
            messagebox.showerror("Error",
                f"arch_style_registry.json not found in {arch_path}.\n"
                "Make sure you selected the correct architect template folder.")
            return False
        if not (arch_path / "arch_template_registry.json").exists():
            messagebox.showerror("Error",
                f"arch_template_registry.json not found in {arch_path}.\n"
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
        bundle = build_phase2_slim_bundle(extract_dir, available_roles=available_roles)
        unresolved = len(bundle.get("paragraphs", []))
        deterministic = len(bundle.get("deterministic_classifications", []))
        self._log(f"  Bundle: {unresolved} unresolved + {deterministic} deterministic = {unresolved + deterministic} total")

        if unresolved > 0 and not api_key:
            raise ValueError("Anthropic API key is required when unresolved paragraphs exist.")

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
            style_numid_remap = import_numbering(
                target_extract_dir=extract_dir,
                arch_template_registry=env_registry,
                arch_styles_xml=arch_styles_xml,
                style_ids_to_import=needed_style_ids,
                log=log
            )
        else:
            # Check whether numbering is actually needed before silently skipping
            _check_numbering_module_needed(arch_styles_xml, needed_style_ids)

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
        output_dir = Path(self.output_dir_var.get())
        output_dir.mkdir(parents=True, exist_ok=True)
        output_docx_path = output_dir / (docx_path.stem + "_PHASE2_FORMATTED.docx")
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
        total = len(bundle.get("paragraphs", [])) + len(bundle.get("deterministic_classifications", []))
        classified = len(classifications.get("classifications", []))
        coverage = (classified / total * 100) if total > 0 else 100.0

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

    def _show_info_popup(self, title: str, content: str):
        popup = tk.Toplevel(self.root)
        popup.title(title)
        popup.geometry("760x650")
        popup.minsize(640, 480)
        popup.configure(bg=self.colors["bg_card"])
        popup.transient(self.root)

        frame = tk.Frame(popup, bg=self.colors["bg_card"], padx=16, pady=16)
        frame.pack(fill=tk.BOTH, expand=True)

        text = scrolledtext.ScrolledText(
            frame,
            wrap=tk.WORD,
            font=("Segoe UI", 10),
            bg=self.colors["bg_input"],
            fg=self.colors["text_primary"],
            insertbackground=self.colors["text_primary"],
            relief=tk.FLAT,
            highlightthickness=1,
            highlightbackground=self.colors["border"],
            highlightcolor=self.colors["accent"],
            padx=12,
            pady=12,
        )
        text.pack(fill=tk.BOTH, expand=True)
        self._render_markdown(text, content)
        text.config(state=tk.DISABLED)

        tk.Button(frame, text="Close", command=popup.destroy, **self.secondary_btn_style).pack(pady=(12, 0))

    def _render_markdown(self, widget: scrolledtext.ScrolledText, markdown_text: str) -> None:
        """Render a focused subset of Markdown into a Tk text widget."""
        widget.tag_configure("h1", font=("Segoe UI", 16, "bold"), spacing1=10, spacing3=6)
        widget.tag_configure("h2", font=("Segoe UI", 13, "bold"), spacing1=8, spacing3=4)
        widget.tag_configure("h3", font=("Segoe UI", 11, "bold"), spacing1=6, spacing3=3)
        widget.tag_configure("bold", font=("Segoe UI", 10, "bold"))
        widget.tag_configure("italic", font=("Segoe UI", 10, "italic"))
        widget.tag_configure("code", font=("Consolas", 10), background="#202020")
        widget.tag_configure("hr", foreground=self.colors["text_muted"])

        for raw_line in markdown_text.splitlines():
            line = raw_line.rstrip("\n")
            if not line.strip():
                widget.insert(tk.END, "\n")
                continue

            if line.startswith("### "):
                self._insert_inline_markdown(widget, f"{line[4:]}\n", base_tag="h3")
                continue
            if line.startswith("## "):
                self._insert_inline_markdown(widget, f"{line[3:]}\n", base_tag="h2")
                continue
            if line.startswith("# "):
                self._insert_inline_markdown(widget, f"{line[2:]}\n", base_tag="h1")
                continue

            if re.fullmatch(r"\s*[-*_]{3,}\s*", line):
                widget.insert(tk.END, "─" * 48 + "\n", ("hr",))
                continue

            bullet_match = re.match(r"^(\s*)([-*]|\d+\.)\s+(.*)$", line)
            if bullet_match:
                indent, marker, content = bullet_match.groups()
                is_numbered = marker.endswith(".") and marker[:-1].isdigit()
                bullet = f"{marker if is_numbered else '•'} "
                widget.insert(tk.END, indent + bullet)
                self._insert_inline_markdown(widget, f"{content}\n")
                continue

            self._insert_inline_markdown(widget, f"{line}\n")

    def _insert_inline_markdown(self, widget: scrolledtext.ScrolledText, text: str, base_tag: Optional[str] = None) -> None:
        """Insert inline markdown (**bold**, *italic*, `code`) into widget."""
        combined_re = re.compile(r"(`[^`]+`|\*\*[^*]+\*\*|\*[^*]+\*)")
        idx = 0
        for match in combined_re.finditer(text):
            if match.start() > idx:
                tags = (base_tag,) if base_tag else ()
                widget.insert(tk.END, text[idx:match.start()], tags)

            token = match.group(0)
            if token.startswith("**"):
                token_text = token[2:-2]
                tags = ("bold",) if not base_tag else (base_tag, "bold")
            elif token.startswith("`"):
                token_text = token[1:-1]
                tags = ("code",) if not base_tag else (base_tag, "code")
            else:
                token_text = token[1:-1]
                tags = ("italic",) if not base_tag else (base_tag, "italic")
            widget.insert(tk.END, token_text, tags)
            idx = match.end()

        if idx < len(text):
            tags = (base_tag,) if base_tag else ()
            widget.insert(tk.END, text[idx:], tags)

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
