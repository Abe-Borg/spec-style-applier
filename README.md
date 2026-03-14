# Phase 2: MEP Specification Styling Engine

## Overview

Phase 2 applies architect-defined CSI paragraph styles to mechanical and plumbing specifications while preserving exact Word behavior and appearance.

This phase does not design styles. It consumes styles produced by Phase 1 and applies them deterministically.

**Core principle:** Change as little Word XML as possible while achieving exact visual and behavioral alignment with the architect's template.

## Quick Start

### Automated Pipeline (Recommended)

```bash
python docx_decomposer.py target.docx \
  --phase2-arch-extract NVES_extracted/ \
  --classify \
  --api-key YOUR_API_KEY
```

This runs the full pipeline: extract → build bundle → LLM classify → apply → format.

### GUI

```bash
python gui.py
```

The GUI provides file/folder pickers, batch processing, and real-time progress logging. Batch mode processes all `.docx` files in a folder sequentially.

### Manual Pipeline (Two-Step)

**Step 1: Build LLM input bundle**
```bash
python docx_decomposer.py target.docx \
  --phase2-arch-extract NVES_extracted/ \
  --phase2-build-bundle
```

**Step 2: Apply LLM classifications**
```bash
python docx_decomposer.py target.docx \
  --phase2-arch-extract NVES_extracted/ \
  --phase2-classifications phase2_classifications.json \
  --output-docx output.docx
```

### CLI Flags

| Flag | Description |
|------|-------------|
| `--phase2-arch-extract DIR` | Architect extracted template folder |
| `--phase2-build-bundle` | Build slim bundle for manual LLM classification |
| `--phase2-classifications JSON` | Apply pre-computed LLM classifications |
| `--classify` | Run full automated pipeline (requires API key) |
| `--api-key KEY` | Anthropic API key (or set `ANTHROPIC_API_KEY` env var) |
| `--model MODEL` | LLM model (default: `claude-sonnet-4-20250514`) |
| `--output-docx PATH` | Output DOCX path |
| `--use-extract-dir DIR` | Reuse existing extracted folder |
| `--extract-dir DIR` | Specify extraction location |

## Installation

```bash
pip install -r requirements.txt
```

Runtime dependency: `anthropic==0.84.0` (for automated classification, pinned with all transitive dependencies).

For development:
```bash
pip install -r requirements-dev.txt
```

For PyInstaller packaging:
```bash
pip install -r requirements-build.txt
```

## Inputs

1. **Target MEP DOCX** — A mechanical or plumbing specification document
2. **Architect Registry Folder** — Folder containing the two JSON files from Phase 1:
   - `arch_style_registry.json` — CSI role → style ID mapping
   - `arch_template_registry.json` — Complete formatting environment (styles, numbering, theme, fonts, settings, docDefaults)

   Only these two files are needed. Phase 2 does not require the architect's extracted `word/` folder.
3. **API Key** (for automated classification) — Anthropic API key

## What Phase 2 Does

1. **Extracts** the target DOCX safely
2. **Applies formatting environment** (theme, fonts, settings, docDefaults) from the architect template
3. **Imports numbering definitions** with ID collision avoidance
4. **Imports architect styles** with property materialization for cross-document portability
5. **Builds a slim bundle** of paragraph text for LLM classification
6. **Classifies paragraphs** via LLM into CSI semantic roles
7. **Applies styles** using only `w:pStyle` — no run-level formatting
8. **Strips run-level fonts** so paragraph styles take effect
9. **Verifies stability** of headers, footers, sectPr, and relationships
10. **Patches output DOCX** via surgical ZIP replacement

## CSI Roles

| Role | Description |
|------|-------------|
| `SectionID` | Section number line (e.g., "SECTION 23 05 13") |
| `SectionTitle` | Section name line |
| `END_OF_SECTION` | End of section marker (e.g., "END OF SECTION") |
| `PART` | Part headings (PART 1, PART 2, PART 3) |
| `ARTICLE` | Article numbers (1.01, 2.03) |
| `PARAGRAPH` | Lettered paragraphs (A., B., C.) |
| `SUBPARAGRAPH` | Numbered under letters (1., 2., 3.) |
| `SUBSUBPARAGRAPH` | Lettered under numbers (a., b., c.) |

## Hard Invariants

1. **Headers/footers unchanged** — no XML drift, no relationship changes
2. **`w:sectPr` untouched** — no page setup or section break changes
3. **Numbering definitions untouched** — `numbering.xml` only modified by the numbering importer
4. **Registry-only styling** — no guessing style IDs; missing role = skip + log

## Testing

```bash
python -m pytest tests/ -v
```

## Known Failure Modes

1. **Numbering stops on Enter** — Fixed by `ensure_explicit_numpr_from_current_style()` before restyling
2. **Fonts change after styling** — Fixed by materializing effective properties during import
3. **Some paragraphs not styled** — Expected when role is missing from registry (logged in preflight)
4. **Word "Repair" warning** — Check: style blocks intact, all dependencies imported, no artifacts in DOCX
5. **LLM output truncated** — If the spec has 300+ paragraphs, the classification JSON may exceed the model's output limit. The classifier automatically chunks large documents. If coverage is still low on a large spec, try reducing the chunk size.

## Copyright Notice

**Copyright (c) 2025 Abraham Borg. All Rights Reserved.**

This software and associated documentation files (the "Software") are the proprietary property of Abraham Borg.

**Unauthorized copying, modification, distribution, or use of this Software, via any medium, is strictly prohibited without express written permission from the copyright holder.**

This Software is provided for review and reference purposes only. No license or right to use, copy, modify, or distribute this Software for any purpose, commercial or non-commercial, is granted.
