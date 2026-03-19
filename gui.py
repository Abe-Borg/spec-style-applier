#!/usr/bin/env python3
"""
Phase 2 MEP Specification Styling Engine — GUI

Tkinter-based graphical interface for applying architect-defined CSI
paragraph styles to MEP specification documents. Supports single-file
and batch processing modes.
"""

import os
import sys
import re
import threading
import subprocess
from pathlib import Path
from typing import Optional

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

import customtkinter as ctk

from batch_runner import (
    BatchResult,
    load_and_validate_shared_config,
    process_single_file,
    run_batch_api,
    run_batch_concurrent,
)

COLORS = {
    "bg_dark": "#0D0D0D",
    "bg_card": "#1A1A1A",
    "bg_input": "#252525",
    "border": "#333333",
    "text_primary": "#FFFFFF",
    "text_secondary": "#B0B0B0",
    "text_muted": "#707070",
    "accent": "#3B82F6",
    "accent_hover": "#2563EB",
    "accent_glow": "#60A5FA",
    "success": "#22C55E",
    "success_glow": "#4ADE80",
    "warning": "#F59E0B",
    "error": "#EF4444",
    "critical": "#DC2626",
    "high": "#F97316",
    "medium": "#EAB308",
    "gripe": "#A855F7",
    "coordination": "#06B6D4",
}


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

Batch mode runs the same process on every `.docx` file in a folder with configurable concurrency. This is useful when you have a full set of spec sections to reformat for a project. Each file is processed independently and gets its own formatted output.

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

**4.** Click **Run Phase 2**. The tool processes files concurrently based on the worker count. The log shows per-file results and a summary at the end.

Output files are placed in the selected output folder (defaults to the input folder), each named `<original_filename>_PHASE2_FORMATTED.docx`.

---

## Notes

- Your **original file is never modified**. The tool always writes a new output document.
- Run the tool again on the same file any time — it will overwrite the previous output.
- If you see a preflight validation error, the architect's template folder may be from a different Phase 1 run than expected. Confirm you have the right folder for this project.
- The API key field is masked and is not saved to disk by this tool.
"""




class Phase2GUI(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Phase 2 — MEP Spec Styling Engine")
        self.geometry("900x850")
        self.minsize(750, 700)
        self.configure(fg_color=COLORS["bg_dark"])

        self.processing = False
        self.output_path: Optional[Path] = None
        self.log_path: Optional[Path] = None

        self._mode_var = ctk.StringVar(value="Single File")
        self._workers_var = ctk.StringVar(value="3")
        self._use_batch_api_var = ctk.BooleanVar(value=False)
        self.status_var = ctk.StringVar(value="Ready")
        self._inputs_expanded = True
        self._log_expanded = True

        self._build_ui()

    @staticmethod
    def _resolve_initial_dir(current_value: str) -> Optional[str]:
        """Return a valid starting directory from a field value, or None."""
        if not current_value:
            return None
        p = Path(current_value)
        if p.is_dir():
            return str(p)
        if p.is_file():
            return str(p.parent)
        # Path typed but not yet created — walk up until we find a real ancestor
        for ancestor in p.parents:
            if ancestor.is_dir():
                return str(ancestor)
        return None

    def _build_ui(self):
        main = ctk.CTkFrame(self, fg_color="transparent")
        main.pack(fill="both", expand=True, padx=24, pady=24)

        header = ctk.CTkFrame(main, fg_color="transparent")
        header.pack(fill="x", pady=(0, 14))

        header_left = ctk.CTkFrame(header, fg_color="transparent")
        header_left.pack(side="left", fill="x", expand=True)
        ctk.CTkLabel(
            header_left,
            text="PHASE 2 MEP STYLING ENGINE",
            text_color=COLORS["text_primary"],
            font=ctk.CTkFont(family="Segoe UI", size=20, weight="bold"),
        ).pack(anchor="w")
        ctk.CTkLabel(
            header_left,
            text="Apply architect style language to target specs",
            text_color=COLORS["text_secondary"],
            font=ctk.CTkFont(family="Segoe UI", size=11),
        ).pack(anchor="w", pady=(2, 0))

        header_right = ctk.CTkFrame(header, fg_color="transparent")
        header_right.pack(side="right", anchor="ne")
        self._create_secondary_button(
            header_right,
            text="How It Works",
            command=lambda: self._show_info_popup("How It Works", HOW_IT_WORKS_TEXT),
        ).pack(side="left", padx=(0, 8))
        self._create_secondary_button(
            header_right,
            text="How to Use",
            command=lambda: self._show_info_popup("How to Use", HOW_TO_USE_TEXT),
        ).pack(side="left")

        input_card = ctk.CTkFrame(main, fg_color=COLORS["bg_card"], corner_radius=8)
        input_card.pack(fill="x", pady=(0, 12))

        input_header = ctk.CTkFrame(input_card, fg_color="transparent")
        input_header.pack(fill="x", padx=16, pady=12)
        input_header.bind("<Button-1>", self._toggle_inputs)

        self._inputs_arrow = ctk.CTkLabel(
            input_header,
            text="▼",
            width=20,
            text_color=COLORS["text_muted"],
            font=ctk.CTkFont(family="Consolas", size=12),
        )
        self._inputs_arrow.pack(side="left")
        self._inputs_arrow.bind("<Button-1>", self._toggle_inputs)

        inputs_title = ctk.CTkLabel(
            input_header,
            text="INPUTS",
            text_color=COLORS["text_muted"],
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
        )
        inputs_title.pack(side="left", padx=(4, 0))
        inputs_title.bind("<Button-1>", self._toggle_inputs)

        self._inputs_content = ctk.CTkFrame(input_card, fg_color="transparent")
        self._inputs_content.pack(fill="x", padx=16, pady=(0, 16))
        self._inputs_content.columnconfigure(1, weight=1)

        ctk.CTkLabel(self._inputs_content, text="Mode", width=120, anchor="w", text_color=COLORS["text_secondary"], font=ctk.CTkFont(family="Segoe UI", size=11)).grid(row=0, column=0, sticky="w", pady=8)

        mode_frame = ctk.CTkFrame(self._inputs_content, fg_color="transparent")
        mode_frame.grid(row=0, column=1, sticky="w", padx=(8, 0), pady=8)
        self.mode_selector = ctk.CTkSegmentedButton(
            mode_frame,
            values=["Single File", "Batch (folder)"],
            variable=self._mode_var,
            command=self._on_mode_change,
            font=ctk.CTkFont(family="Segoe UI", size=11),
            selected_color=COLORS["accent"],
            selected_hover_color=COLORS["accent_hover"],
            unselected_color=COLORS["bg_input"],
            unselected_hover_color=COLORS["border"],
            fg_color=COLORS["bg_input"],
            text_color=COLORS["text_secondary"],
            height=32,
        )
        self.mode_selector.pack(anchor="w")

        self.target_label = ctk.CTkLabel(self._inputs_content, text="Target Spec (.docx):", width=120, anchor="w", text_color=COLORS["text_secondary"], font=ctk.CTkFont(family="Segoe UI", size=11))
        self.target_label.grid(row=1, column=0, sticky="w", pady=8)
        self.target_var = ctk.StringVar()
        target_frame = ctk.CTkFrame(self._inputs_content, fg_color="transparent")
        target_frame.grid(row=1, column=1, sticky="ew", padx=(8, 0), pady=8)
        target_frame.columnconfigure(0, weight=1)
        self.target_entry = ctk.CTkEntry(target_frame, textvariable=self.target_var, fg_color=COLORS["bg_input"], border_color=COLORS["border"], text_color=COLORS["text_primary"], font=ctk.CTkFont(family="Consolas", size=11), height=36)
        self.target_entry.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        self.target_btn = self._create_secondary_button(target_frame, text="Browse...", command=self._browse_target, width=90)
        self.target_btn.grid(row=0, column=1)

        ctk.CTkLabel(self._inputs_content, text="Architect Template Folder:", width=120, anchor="w", text_color=COLORS["text_secondary"], font=ctk.CTkFont(family="Segoe UI", size=11)).grid(row=2, column=0, sticky="w", pady=8)
        self.arch_var = ctk.StringVar()
        arch_frame = ctk.CTkFrame(self._inputs_content, fg_color="transparent")
        arch_frame.grid(row=2, column=1, sticky="ew", padx=(8, 0), pady=8)
        arch_frame.columnconfigure(0, weight=1)
        ctk.CTkEntry(arch_frame, textvariable=self.arch_var, fg_color=COLORS["bg_input"], border_color=COLORS["border"], text_color=COLORS["text_primary"], font=ctk.CTkFont(family="Consolas", size=11), height=36).grid(row=0, column=0, sticky="ew", padx=(0, 6))
        self._create_secondary_button(arch_frame, text="Browse...", command=self._browse_arch, width=90).grid(row=0, column=1)

        ctk.CTkLabel(self._inputs_content, text="Anthropic API Key:", width=120, anchor="w", text_color=COLORS["text_secondary"], font=ctk.CTkFont(family="Segoe UI", size=11)).grid(row=3, column=0, sticky="w", pady=8)
        self.api_key_var = ctk.StringVar(value=os.environ.get("ANTHROPIC_API_KEY", ""))
        ctk.CTkEntry(self._inputs_content, textvariable=self.api_key_var, show="•", placeholder_text="sk-ant-...", fg_color=COLORS["bg_input"], border_color=COLORS["border"], text_color=COLORS["text_primary"], font=ctk.CTkFont(family="Consolas", size=11), height=36).grid(row=3, column=1, sticky="ew", padx=(8, 0), pady=8)

        ctk.CTkLabel(self._inputs_content, text="Output Folder:", width=120, anchor="w", text_color=COLORS["text_secondary"], font=ctk.CTkFont(family="Segoe UI", size=11)).grid(row=4, column=0, sticky="w", pady=8)
        self.output_dir_var = ctk.StringVar()
        output_frame = ctk.CTkFrame(self._inputs_content, fg_color="transparent")
        output_frame.grid(row=4, column=1, sticky="ew", padx=(8, 0), pady=8)
        output_frame.columnconfigure(0, weight=1)
        ctk.CTkEntry(output_frame, textvariable=self.output_dir_var, fg_color=COLORS["bg_input"], border_color=COLORS["border"], text_color=COLORS["text_primary"], font=ctk.CTkFont(family="Consolas", size=11), height=36).grid(row=0, column=0, sticky="ew", padx=(0, 6))
        self._create_secondary_button(output_frame, text="Browse...", command=self._browse_output_dir, width=90).grid(row=0, column=1)

        self.workers_label = ctk.CTkLabel(self._inputs_content, text="Concurrent Files:", width=120, anchor="w", text_color=COLORS["text_secondary"], font=ctk.CTkFont(family="Segoe UI", size=11))
        self.workers_menu = ctk.CTkOptionMenu(
            self._inputs_content,
            values=[str(i) for i in range(1, 7)],
            variable=self._workers_var,
            width=100,
            fg_color=COLORS["bg_input"],
            button_color=COLORS["border"],
            button_hover_color=COLORS["accent_hover"],
            text_color=COLORS["text_primary"],
            dropdown_fg_color=COLORS["bg_input"],
            dropdown_text_color=COLORS["text_primary"],
        )

        self.batch_api_checkbox = ctk.CTkCheckBox(
            self._inputs_content,
            text="Use Batch API (50% cost savings)",
            variable=self._use_batch_api_var,
            command=self._update_run_button_text,
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=COLORS["text_secondary"],
            fg_color=COLORS["accent"],
            hover_color=COLORS["accent_hover"],
        )

        self.run_btn = ctk.CTkButton(
            main, text="Run Phase 2", command=self._run,
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            height=44, corner_radius=8,
            fg_color=COLORS["accent"], hover_color=COLORS["accent_hover"],
        )
        self.run_btn.pack(fill="x", pady=(0, 8))

        self.progress_bar = ctk.CTkProgressBar(
            main,
            height=4,
            corner_radius=2,
            fg_color=COLORS["bg_input"],
            progress_color=COLORS["accent"],
            indeterminate_speed=0.5,
        )
        self.progress_bar.configure(mode="indeterminate")

        self.status_label = ctk.CTkLabel(
            main,
            textvariable=self.status_var,
            text_color=COLORS["text_secondary"],
            anchor="e",
            font=ctk.CTkFont(family="Segoe UI", size=11),
        )
        self.status_label.pack(fill="x", pady=(0, 12))

        log_card = ctk.CTkFrame(main, fg_color=COLORS["bg_card"], corner_radius=8)
        log_card.pack(fill="both", expand=True, pady=(0, 12))

        log_header = ctk.CTkFrame(log_card, fg_color="transparent")
        log_header.pack(fill="x", padx=16, pady=12)
        log_header.bind("<Button-1>", self._toggle_log)

        self._log_arrow = ctk.CTkLabel(log_header, text="▼", width=20, text_color=COLORS["text_muted"], font=ctk.CTkFont(family="Consolas", size=12))
        self._log_arrow.pack(side="left")
        self._log_arrow.bind("<Button-1>", self._toggle_log)

        log_title = ctk.CTkLabel(log_header, text="ACTIVITY LOG", text_color=COLORS["text_muted"], font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"))
        log_title.pack(side="left", padx=(4, 0))
        log_title.bind("<Button-1>", self._toggle_log)

        self.clear_log_btn = self._create_secondary_button(log_header, text="Clear", command=self._clear_log, width=80, height=30)
        self.clear_log_btn.pack(side="right")

        self._log_content = ctk.CTkFrame(log_card, fg_color="transparent")
        self._log_content.pack(fill="both", expand=True, padx=16, pady=(0, 16))

        self.log_text = ctk.CTkTextbox(
            self._log_content,
            fg_color=COLORS["bg_input"],
            corner_radius=4,
            font=ctk.CTkFont(family="Consolas", size=12),
            text_color=COLORS["text_secondary"],
            wrap="word",
            state="disabled",
            activate_scrollbars=True,
        )
        self.log_text.pack(fill="both", expand=True)

        btn_frame = ctk.CTkFrame(self._log_content, fg_color="transparent")
        btn_frame.pack(fill="x", pady=(10, 0))

        self.open_output_btn = self._create_secondary_button(btn_frame, text="Open Output DOCX", command=self._open_output, state="disabled", width=160, height=32)
        self.open_output_btn.pack(side="left", padx=(0, 8))

        self.open_log_btn = self._create_secondary_button(btn_frame, text="Open Log", command=self._open_log, state="disabled", width=120, height=32)
        self.open_log_btn.pack(side="left")

        self._on_mode_change()

    def _create_secondary_button(self, parent, text: str, command, width: int = 100, height: int = 36, state: str = "normal"):
        return ctk.CTkButton(
            parent,
            text=text,
            command=command,
            width=width,
            height=height,
            state=state,
            font=ctk.CTkFont(family="Segoe UI", size=12),
            fg_color=COLORS["bg_input"],
            hover_color=COLORS["border"],
            border_width=1,
            border_color=COLORS["border"],
            text_color=COLORS["text_secondary"],
        )

    def _toggle_inputs(self, event=None):
        if self._inputs_expanded:
            self._inputs_content.pack_forget()
            self._inputs_arrow.configure(text="▶")
            self._inputs_expanded = False
        else:
            self._inputs_content.pack(fill="x", padx=16, pady=(0, 16))
            self._inputs_arrow.configure(text="▼")
            self._inputs_expanded = True

    def _toggle_log(self, event=None):
        if self._log_expanded:
            self._log_content.pack_forget()
            self._log_arrow.configure(text="▶")
            self._log_expanded = False
        else:
            self._log_content.pack(fill="both", expand=True, padx=16, pady=(0, 16))
            self._log_arrow.configure(text="▼")
            self._log_expanded = True

    def _clear_log(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def _on_mode_change(self, _value=None):
        is_batch = self._mode_var.get() == "Batch (folder)"
        self.target_label.configure(text="Target Folder:" if is_batch else "Target Spec (.docx):")
        if is_batch:
            self.workers_label.grid(row=5, column=0, sticky="w", pady=8)
            self.workers_menu.grid(row=5, column=1, sticky="w", padx=(8, 0), pady=8)
            self.batch_api_checkbox.grid(row=6, column=1, sticky="w", padx=(8, 0), pady=(0, 8))
        else:
            self.workers_label.grid_remove()
            self.workers_menu.grid_remove()
            self.batch_api_checkbox.grid_remove()
            self._use_batch_api_var.set(False)
        self.target_var.set("")
        self.output_dir_var.set("")
        if is_batch and self._use_batch_api_var.get():
            self.run_btn.configure(text="Submit Batch")
        else:
            self.run_btn.configure(text="Run Phase 2")

    def _update_run_button_text(self):
        if self.processing:
            return
        if self._mode_var.get() == "Batch (folder)" and self._use_batch_api_var.get():
            self.run_btn.configure(text="Submit Batch")
        else:
            self.run_btn.configure(text="Run Phase 2")

    def _browse_target(self):
        initial = self._resolve_initial_dir(self.target_var.get())
        if self._mode_var.get() == "Batch (folder)":
            path = filedialog.askdirectory(title="Select folder with .docx files", initialdir = initial,)
        else:
            path = filedialog.askopenfilename(
                title="Select Target Spec",
                filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")],
                initialdir = initial,
            )
        if path:
            self.target_var.set(path)
            default_output = path if self._mode_var.get() == "Batch (folder)" else str(Path(path).parent)
            self.output_dir_var.set(default_output)

    def _browse_arch(self):
        initial = self._resolve_initial_dir(self.arch_var.get())
        path = filedialog.askdirectory(title="Select Architect Template Folder", initialdir = initial)
        if path:
            self.arch_var.set(path)

    def _browse_output_dir(self):
        initial = self._resolve_initial_dir(self.output_dir_var.get())
        path = filedialog.askdirectory(title="Select Output Folder", initialdir = initial,)
        if path:
            self.output_dir_var.set(path)

    def _log(self, msg: str):
        self.after(0, self._append_log, msg)

    def _append_log(self, msg: str):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _set_status(self, status: str):
        self.after(0, lambda: self.status_var.set(status))

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
            messagebox.showerror("Error", f"arch_style_registry.json not found in {arch_path}.\nMake sure you selected the correct architect template folder.")
            return False
        if not (arch_path / "arch_template_registry.json").exists():
            messagebox.showerror("Error", f"arch_template_registry.json not found in {arch_path}.\nMake sure you selected the correct architect template folder.")
            return False
        return True

    def _run(self):
        if self.processing:
            return
        if not self._validate_inputs():
            return

        self.processing = True
        self.output_path = None
        self.log_path = None
        self.run_btn.configure(text="Processing...", state="disabled", text_color_disabled="#FFFFFF")
        self.open_output_btn.configure(state="disabled")
        self.open_log_btn.configure(state="disabled")
        self.progress_bar.pack(fill="x", pady=(0, 8), after=self.run_btn)
        if self._mode_var.get() == "Batch (folder)":
            self.progress_bar.configure(mode="determinate")
            self.progress_bar.set(0)
        else:
            self.progress_bar.configure(mode="indeterminate")
            self.progress_bar.start()

        self._clear_log()

        thread = threading.Thread(target=self._process, daemon=True)
        thread.start()

    def _process(self):
        try:
            if self._mode_var.get() == "Batch (folder)":
                self._process_batch()
            else:
                self._process_single(Path(self.target_var.get()))
        except Exception as e:
            self._log(f"\nERROR: {e}")
            self._set_status("Failed")
            import traceback
            self._log(traceback.format_exc())
        finally:
            self.after(0, self._finish_processing)

    def _process_single(self, docx_path: Path) -> Optional[Path]:
        """Process a single DOCX file. Returns output path on success."""
        self._set_status(f"Processing: {docx_path.name}")
        self._log(f"Processing: {docx_path}")

        shared = load_and_validate_shared_config(Path(self.arch_var.get()))
        result = process_single_file(
            docx_path=docx_path,
            arch_registry=shared.arch_registry,
            env_registry=shared.env_registry,
            arch_styles_xml=shared.arch_styles_xml,
            available_roles=shared.available_roles,
            api_key=self.api_key_var.get(),
            output_dir=Path(self.output_dir_var.get()),
        )

        for line in result.log:
            self._log(f"  {line}")

        if not result.success:
            raise RuntimeError(result.error or "Unknown processing error")

        self.output_path = result.output_path
        return result.output_path

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

        shared = load_and_validate_shared_config(Path(self.arch_var.get()))
        total = len(docx_files)
        state = {"completed": 0}

        def handle_result(result: BatchResult):
            def _on_main_thread():
                state["completed"] += 1
                completed = state["completed"]
                active = max(total - completed, 0)
                self.progress_bar.set(completed / total)
                self._set_status(f"Batch: {completed}/{total} complete, {active} active")

                status_text = "OK" if result.success else "FAILED"
                self._log(f"═══ [{completed}/{total}] {result.filename} — {result.duration_seconds:.1f}s — {status_text} ═══")
                for line in result.log:
                    self._log(f"  {line}")
                self._log("")

            self.after(0, _on_main_thread)

        use_batch_api = self._use_batch_api_var.get()
        if use_batch_api:
            self._log("Phase A: Preparing files...")
            self._set_status("Phase A: preparing files")
            self.progress_bar.configure(mode="indeterminate")
            self.progress_bar.start()

            def on_batch_poll(batch_id: str, processing_status: str, request_counts):
                completed = getattr(request_counts, "succeeded", 0) + getattr(request_counts, "errored", 0) + getattr(request_counts, "canceled", 0) + getattr(request_counts, "expired", 0)
                total_requests = getattr(request_counts, "processing", 0) + completed
                msg = f"Phase B: polling {batch_id} ({completed}/{total_requests} complete)"
                self._set_status(msg)
                self._log(msg)

            self._log("Phase B: submitting Anthropic batch...")
            results = run_batch_api(
                docx_paths=docx_files,
                arch_registry=shared.arch_registry,
                env_registry=shared.env_registry,
                arch_styles_xml=shared.arch_styles_xml,
                available_roles=shared.available_roles,
                api_key=self.api_key_var.get(),
                output_dir=Path(self.output_dir_var.get()),
                max_workers=int(self._workers_var.get()),
                on_file_complete=handle_result,
                on_batch_poll=on_batch_poll,
            )
            self.progress_bar.stop()
            self.progress_bar.configure(mode="determinate")
            self.progress_bar.set(0)
            self._log("Phase C: applying results...")
        else:
            self.progress_bar.configure(mode="determinate")
            self.progress_bar.set(0)
            results = run_batch_concurrent(
                docx_paths=docx_files,
                arch_registry=shared.arch_registry,
                env_registry=shared.env_registry,
                arch_styles_xml=shared.arch_styles_xml,
                available_roles=shared.available_roles,
                api_key=self.api_key_var.get(),
                output_dir=Path(self.output_dir_var.get()),
                max_workers=int(self._workers_var.get()),
                on_file_complete=handle_result,
            )

        ok_count = sum(1 for item in results if item.success)
        self._log("=" * 60)
        self._log("BATCH SUMMARY")
        self._log("=" * 60)
        for item in results:
            status = "OK" if item.success else f"FAILED: {item.error}"
            self._log(f"  {item.filename}: {status}")
        self._log(f"\n  {ok_count}/{len(results)} files processed successfully")
        self._set_status(f"Batch done: {ok_count}/{len(results)} OK")

    def _show_info_popup(self, title: str, content: str):
        popup = ctk.CTkToplevel(self)
        popup.title(title)
        popup.geometry("760x650")
        popup.minsize(640, 480)
        popup.configure(fg_color=COLORS["bg_dark"])
        popup.transient(self)
        popup.grab_set()
        popup.lift()
        popup.focus_force()

        frame = ctk.CTkFrame(popup, fg_color=COLORS["bg_card"], corner_radius=8)
        frame.pack(fill="both", expand=True, padx=16, pady=16)

        text_frame = ctk.CTkFrame(frame, fg_color=COLORS["bg_input"], corner_radius=4)
        text_frame.pack(fill="both", expand=True)

        text = scrolledtext.ScrolledText(
            text_frame,
            wrap=tk.WORD,
            font=("Segoe UI", 10),
            bg=COLORS["bg_input"],
            fg=COLORS["text_primary"],
            insertbackground=COLORS["text_primary"],
            relief=tk.FLAT,
            highlightthickness=1,
            highlightbackground=COLORS["border"],
            highlightcolor=COLORS["accent"],
            padx=12,
            pady=12,
        )
        text.pack(fill=tk.BOTH, expand=True)
        self._render_markdown(text, content)
        text.config(state=tk.DISABLED)

        self._create_secondary_button(frame, text="Close", command=popup.destroy, width=100).pack(pady=(12, 0))

    def _render_markdown(self, widget: scrolledtext.ScrolledText, markdown_text: str) -> None:
        """Render a focused subset of Markdown into a Tk text widget."""
        widget.tag_configure("h1", font=("Segoe UI", 16, "bold"), spacing1=10, spacing3=6)
        widget.tag_configure("h2", font=("Segoe UI", 13, "bold"), spacing1=8, spacing3=4)
        widget.tag_configure("h3", font=("Segoe UI", 11, "bold"), spacing1=6, spacing3=3)
        widget.tag_configure("bold", font=("Segoe UI", 10, "bold"))
        widget.tag_configure("italic", font=("Segoe UI", 10, "italic"))
        widget.tag_configure("code", font=("Consolas", 10), background="#202020")
        widget.tag_configure("hr", foreground=COLORS["text_muted"])

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
        self.progress_bar.stop()
        self.progress_bar.set(0)
        self.progress_bar.pack_forget()
        if self.status_var.get() == "Failed":
            self._reset_run_button()
        else:
            self._on_processing_complete()
        if self.output_path:
            self.open_output_btn.configure(state="normal")
        if self.log_path:
            self.open_log_btn.configure(state="normal")
        if self._mode_var.get() != "Batch (folder)" and self.output_path:
            self._set_status("Done")

    def _on_processing_complete(self):
        self.run_btn.configure(text="✓ Complete", fg_color=COLORS["success"], hover_color=COLORS["success_glow"], state="disabled")
        self.after(2500, self._reset_run_button)

    def _reset_run_button(self):
        if self.processing:
            return
        self.run_btn.configure(fg_color=COLORS["accent"], hover_color=COLORS["accent_hover"], state="normal")
        self._update_run_button_text()

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
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")
    Phase2GUI().mainloop()


if __name__ == "__main__":
    main()
