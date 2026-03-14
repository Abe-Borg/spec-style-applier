# CLAUDE.md — AI Assistant Guide for Claude_Spec_Auto_Formatting

## Project Overview

This is a **Phase 2 MEP Specification Styling Engine** for the AEC (Architecture/Engineering/Construction) industry. It applies architect-defined CSI (Construction Specifications Institute) paragraph styles to MEP (Mechanical, Electrical, Plumbing) specification documents (.docx) while preserving exact Word behavior.

**Core principle:** Change as little Word XML as possible while achieving exact visual and behavioral alignment with the architect's template.

Phase 1 (external, not in this repo) extracts and catalogs styles from an architect's template. Phase 2 (this repo) applies those styles to target MEP specifications deterministically.

## Repository Structure

```
spec-style-applier/
├── docx_decomposer.py        # CLI entry point + DocxDecomposer class (~440 lines)
├── arch_env_applier.py        # Formatting environment, style materialization, Content-Types/rels provisioning (~800 lines)
├── numbering_importer.py      # Numbering definition import with collision avoidance (~360 lines)
├── docx_patch.py              # Surgical ZIP patching with XML well-formedness validation (~115 lines)
├── phase2_invariants.py       # Post-processing invariant verification (118 lines)
├── gui.py                     # Tkinter GUI with batch mode (~460 lines)
├── core/                      # Core logic package
│   ├── __init__.py            # Re-exports public interface
│   ├── xml_helpers.py         # Paragraph XML iteration and manipulation
│   ├── stability.py           # Stability snapshots and verification
│   ├── style_import.py        # Style extraction, materialization, and import
│   ├── classification.py      # Classification application, slim bundle, boilerplate, prompts
│   ├── registry.py            # Registry loading, resolve, preflight
│   └── llm_classifier.py      # Anthropic API integration for automated classification
├── tests/                     # Unit tests (pytest)
│   ├── test_xml_helpers.py    # Tests for XML manipulation functions
│   ├── test_style_import.py   # Tests for style import and materialization
│   ├── test_env_applier.py    # Tests for environment application and style materialization
│   ├── test_registry.py       # Tests for registry loading and preflight validation
│   ├── test_numbering_importer.py # Tests for numbering import and ID collision avoidance
│   ├── test_preflight.py      # Tests for preflight reporting
│   ├── test_classification.py # Tests for classification and boilerplate filtering
│   ├── test_docx_patch.py     # Tests for DOCX patching and XML validation
│   └── test_normalize_contract.py # Tests for paragraph normalization contract
├── requirements.txt           # Full pinned dependencies (anthropic==0.84.0 + transitive deps)
├── requirements-build.txt     # PyInstaller packaging dependencies
├── requirements-dev.txt       # Development: pytest>=7.0
├── README.md                  # Technical documentation
├── CLAUDE.md                  # This file
├── .gitignore                 # Standard Python + project-specific ignores
└── phase2_classifications.json # Example LLM classification output
```

## Technology Stack

- **Python 3.7+** — all source code
- **Standard library** for core functionality: `zipfile`, `re`, `json`, `pathlib`, `hashlib`, `dataclasses`
- **anthropic==0.84.0** — sole external runtime dependency (for automated LLM classification)
- **No external XML libraries** — regex-based XML manipulation for byte-level fidelity
- **pytest** — development dependency for unit tests
- **tkinter** — GUI (stdlib, no extra install)

## How to Run

### Automated Pipeline (Recommended)
```bash
python docx_decomposer.py <target.docx> \
  --phase2-arch-extract <arch_extracted_folder> \
  --classify \
  --api-key <YOUR_API_KEY>
```

### GUI
```bash
python gui.py
```

### Manual Pipeline
**Step 1: Build LLM input bundle**
```bash
python docx_decomposer.py <target.docx> \
  --phase2-arch-extract <arch_extracted_folder> \
  --phase2-build-bundle
```

**Step 2: Apply LLM classifications**
```bash
python docx_decomposer.py <target.docx> \
  --phase2-arch-extract <arch_extracted_folder> \
  --phase2-classifications <phase2_classifications.json> \
  [--output-docx <output.docx>]
```

### CLI Flags
- `--classify` — run full automated pipeline (extract → classify → apply → format)
- `--api-key KEY` — Anthropic API key (or set `ANTHROPIC_API_KEY` env var)
- `--model MODEL` — LLM model (default: `claude-sonnet-4-20250514`)
- `--phase2-build-bundle` — build slim bundle only (manual LLM step)
- `--phase2-classifications JSON` — apply pre-computed classifications
- `--phase2-arch-extract DIR` — architect extracted folder
- `--output-docx PATH` — output DOCX path
- `--use-extract-dir DIR` — reuse existing extracted folder
- `--extract-dir DIR` — specify extraction location

### Running Tests
```bash
python -m pytest tests/ -v
```

## Architecture and Data Flow

```
INPUT: target.docx + arch_template_registry.json [+ API key for --classify]
  │
  ├─→ Extract DOCX to folder (DocxDecomposer.extract)  [docx_decomposer.py]
  │
  ├─→ apply_environment_to_target()                     [arch_env_applier.py]
  │   ├─→ apply_theme()                                 theme/theme1.xml
  │   ├─→ apply_settings()                              compat flags
  │   ├─→ apply_font_table()                            fontTable.xml
  │   └─→ apply_doc_defaults()                          baseline rPr/pPr
  │
  ├─→ import_numbering()                                [numbering_importer.py]
  │   ├─→ find_max_ids_in_numbering()                   collision detection
  │   ├─→ build_numbering_import_plan()                 remapping strategy
  │   └─→ inject_numbering_into_xml()                   merge abstractNums + nums
  │
  ├─→ import_arch_styles_into_target()                  [core/style_import.py]
  │   ├─→ _collect_style_deps_from_arch()               expand basedOn chain
  │   ├─→ materialize_arch_style_block()                harden for portability
  │   └─→ insert_styles_into_styles_xml()               merge into target
  │
  ├─→ build_phase2_slim_bundle()                        [core/classification.py]
  │   └─→ strip_boilerplate_with_report()               filter noise
  │
  ├─→ classify_target_document()                        [core/llm_classifier.py]
  │   ├─→ Anthropic API call (with chunking + retry)
  │   └─→ Coverage check (warn if < 85%)
  │
  ├─→ snapshot_stability()                              [core/stability.py]
  │
  ├─→ apply_phase2_classifications()                    [core/classification.py]
  │   ├─→ ensure_explicit_numpr_from_current_style()    [core/style_import.py]
  │   ├─→ strip_run_font_formatting()                   [core/xml_helpers.py]
  │   └─→ apply_pstyle_to_paragraph_block()             [core/xml_helpers.py]
  │
  ├─→ verify_stability()                                [core/stability.py]
  │
  ├─→ patch_docx()                                      [docx_patch.py]
  │
  └─→ verify_phase2_invariants()                        [phase2_invariants.py]

OUTPUT: <target>_PHASE2_FORMATTED.docx
```

## Module Responsibilities

| Module | Role |
|---|---|
| `docx_decomposer.py` | CLI entry point, orchestrator, DOCX extraction (`DocxDecomposer` class) |
| `core/xml_helpers.py` | Paragraph XML iteration, text extraction, pStyle application, run font stripping |
| `core/stability.py` | SHA256 snapshots, header/footer/sectPr/rels stability verification |
| `core/style_import.py` | Style extraction, basedOn chain walking, property materialization, style import |
| `core/classification.py` | LLM prompts, boilerplate filtering, slim bundle building, classification application |
| `core/registry.py` | Architect registry loading, preflight reporting, path resolution |
| `core/llm_classifier.py` | Anthropic API integration (streaming), retry logic, chunking, coverage metrics |
| `arch_env_applier.py` | Imports formatting environment (theme, settings, fonts, docDefaults), Content-Types/rels provisioning, effective rPr resolution, style materialization for import |
| `numbering_importer.py` | Imports numbering definitions (abstractNum + num) with ID collision avoidance |
| `docx_patch.py` | Creates output DOCX via surgical ZIP entry replacement with XML well-formedness validation |
| `phase2_invariants.py` | Post-processing validation: sectPr, headers/footers, run properties |
| `gui.py` | Tkinter GUI with single-file and batch processing modes |

## Key Conventions and Patterns

### XML Processing via Regex (not DOM)
The codebase uses regex for all Word XML parsing instead of DOM/ElementTree. This is intentional — it preserves byte-level fidelity and avoids namespace normalization issues. When modifying code, continue this pattern:
```python
m = re.search(r'<w:pStyle\b[^>]*w:val="([^"]+)"', p_xml)
```

### Iterator-Based Paragraph Processing
Paragraphs are processed via character-position iterators that maintain byte alignment for precise in-place replacement. The `iter_paragraph_xml_blocks()` function yields `(start, end, p_xml)` tuples.

### Style Chain Walking
Word styles use `basedOn` inheritance. The codebase resolves effective properties by walking the chain:
```python
cur = style_id
while cur and cur not in seen:
    block = _extract_style_block(styles_xml, cur)
    cur = _extract_basedOn(block)
```

### Property Materialization
When importing styles across documents, inherited properties are materialized (copied explicitly) so styles are self-contained and don't depend on the target document's inheritance chain.

### Defensive Contract Checking
All external inputs (LLM classifications, registries) are validated defensively. Invalid entries are logged and skipped, never causing crashes.

### Stability Snapshots
Before and after document.xml modifications, SHA256 hashes of headers/footers and sectPr blocks are compared to verify invariant preservation.

## Hard Invariants (Must Never Break)

1. **Headers/footers unchanged** — no XML drift, no relationship changes
2. **`w:sectPr` untouched** — no page setup or section break changes
3. **Numbering definitions untouched** — `numbering.xml` is never edited by style application; only the numbering importer adds new definitions
4. **No run-level formatting** — no `<w:rPr>` edits inside document.xml (except font stripping); all formatting through paragraph styles only
5. **Registry-only styling** — no guessing style IDs; missing role = skip + log

## CSI Roles (Classification Vocabulary)

The LLM classifies paragraphs into these semantic roles:
- `SectionID` — section number line (e.g., "SECTION 23 05 13")
- `SectionTitle` — section name line
- `END_OF_SECTION` — end of section marker (e.g., "END OF SECTION")
- `PART` — part headings (PART 1, PART 2, PART 3)
- `ARTICLE` — article numbers (1.01, 2.03)
- `PARAGRAPH` — lettered paragraphs (A., B., C.)
- `SUBPARAGRAPH` — numbered under letters (1., 2., 3.)
- `SUBSUBPARAGRAPH` — lettered under numbers (a., b., c.)

## Key Input/Output Files

### Inputs (from Phase 1 / LLM)
- **`arch_style_registry.json`** — Maps CSI roles to Word style IDs. Sole source of truth for styling.
- **`arch_template_registry.json`** — Complete formatting environment: theme, styles, numbering, docDefaults, fonts, settings.
- **`phase2_classifications.json`** — LLM output: `{ "classifications": [{ "paragraph_index": N, "csi_role": "ROLE" }] }`

Phase 2 operates entirely from these two JSON files. No access to the architect's extracted `word/` folder is needed. The style definitions in `arch_template_registry.json` are reconstructed into a synthetic `styles.xml` at runtime via `build_arch_styles_xml_from_registry()` for regex-based processing.

### Outputs
- **`<target>_PHASE2_FORMATTED.docx`** — The restyled document

## Coding Standards

### Naming
- `snake_case` for functions and variables
- `UPPERCASE` for constants
- `_prefix` for internal/private helper functions
- Word XML namespace prefix `w:` preserved in string literals

### Docstrings
Present on all major public functions. Include parameter and return descriptions where complexity warrants it.

### Type Hints
Used via `typing` module (Dict, List, Optional, Tuple, etc.) on function signatures.

### Error Handling
Mix of try/except and explicit validation. Errors are logged into `log: List[str]` arrays passed through the call chain, and fatal invariant violations raise `RuntimeError` or `ValueError`.

### Imports
- Standard library imports first, then local module imports
- `numbering_importer` is imported conditionally with a `HAS_NUMBERING_IMPORTER` flag fallback
- Core logic lives in `core/` package; `docx_decomposer.py` imports from it

## What NOT to Do

- **Do not use DOM/ElementTree for modifying Word XML** — the codebase intentionally uses regex to preserve byte-level fidelity
- **Do not create styles** — Phase 2 only applies styles defined in the architect registry
- **Do not modify headers, footers, or sectPr** — these are protected by hard invariants
- **Do not apply run-level formatting** — all formatting is through `w:pStyle` only
- **Do not infer or guess style IDs** — if a role is missing from the registry, skip it
- **Do not modify `numbering.xml` during style application** — numbering preservation happens at the paragraph level via `<w:numPr>` materialization
- **Do not add heavy external dependencies** — core uses stdlib + anthropic only
- **Do not merge Phase 1 logic into this codebase** — Phase 1 and Phase 2 are separate concerns
- **Do not read the architect's extracted `word/` folder directly** — all architect data must come through the two JSON contract files via `build_arch_styles_xml_from_registry()`

## Allowed Patch Targets (docx_patch.py)

Only these ZIP entries may be modified in the output DOCX:
- `word/document.xml`
- `word/styles.xml`
- `word/theme/theme1.xml`
- `word/numbering.xml`
- `word/settings.xml`
- `word/fontTable.xml`
- `[Content_Types].xml`
- `word/_rels/document.xml.rels`

Headers (`word/header*`) and footers (`word/footer*`) are **explicitly forbidden** from patching.

## Known Failure Modes

1. **Numbering stops on Enter** — numbering was style-linked; `w:pStyle` swapped without materializing `<w:numPr>`. Fix: ensure `ensure_explicit_numpr_from_current_style()` runs before restyling.
2. **Fonts change after styling** — imported style lacked explicit `<w:rPr>` due to inheritance. Fix: materialize effective properties when importing.
3. **Some paragraphs not styled** — role missing from registry or intentionally skipped (`SKIP`, `END_OF_SECTION`). Expected behavior, logged in preflight.
4. **Word opens with "Repair" warning** — invalid XML or broken basedOn chain. Check: style blocks are intact, all dependencies imported, no workspace artifacts in DOCX.
5. **LLM output truncated** — If the spec has 300+ paragraphs, the classification JSON may exceed the model's output limit. The classifier automatically chunks large documents. If coverage is still low on a large spec, try reducing the chunk size.
