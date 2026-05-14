# CLAUDE.md — AI Assistant Guide for Claude_Spec_Auto_Formatting

## Project Overview

This is a **Phase 2 MEP Specification Styling Engine** for the AEC (Architecture/Engineering/Construction) industry. It applies architect-defined CSI (Construction Specifications Institute) paragraph styles to MEP (Mechanical, Electrical, Plumbing) specification documents (.docx) while preserving exact Word behavior.

**Core principle:** Change as little Word XML as possible while achieving exact visual and behavioral alignment with the architect's template.

Phase 1 (external, not in this repo) extracts and catalogs styles from an architect's template. Phase 2 (this repo) applies those styles to target MEP specifications deterministically.

## Repository Structure

```
spec-style-applier/
├── docx_decomposer.py        # DocxDecomposer class — DOCX extraction only (~70 lines)
├── batch_runner.py           # Pipeline orchestrator: single-file and batch processing
├── arch_env_applier.py       # Formatting environment, style materialization, Content-Types/rels provisioning
├── header_footer_importer.py # Header/footer import from architect template
├── numbering_importer.py     # Numbering definition import with collision avoidance
├── docx_patch.py             # Surgical ZIP patching with XML well-formedness validation
├── phase2_invariants.py      # Post-processing invariant verification
├── gui.py                    # customtkinter GUI with single-file and batch modes
├── core/                     # Core logic package
│   ├── __init__.py           # Re-exports public interface
│   ├── xml_helpers.py        # Paragraph XML iteration and manipulation
│   ├── stability.py          # Stability snapshots and verification
│   ├── style_import.py       # Style extraction, materialization, and import
│   ├── classification.py     # Classification application, slim bundle, boilerplate, prompts
│   ├── registry.py           # Registry loading, resolve, preflight
│   ├── llm_classifier.py     # Anthropic API integration for automated classification
│   ├── batch_classifier.py   # Anthropic Batch API helpers for folder-level classification
│   ├── token_utils.py        # SectionID/SectionTitle token extraction and case utilities
│   ├── ooxml_namespaces.py   # OOXML namespace constants and ElementTree serialization helpers
│   ├── section_mapping.py    # Section source selection for sectPr alignment
│   ├── sectpr_tools.py       # sectPr block extraction, child ordering, and tag manipulation
│   └── prompts/              # LLM prompt text files (phase2_master_prompt.txt, phase2_run_instruction.txt)
├── tests/                    # Unit tests (pytest)
│   ├── test_xml_helpers.py
│   ├── test_style_import.py
│   ├── test_env_applier.py
│   ├── test_registry.py
│   ├── test_numbering_importer.py
│   ├── test_preflight.py
│   ├── test_classification.py
│   ├── test_docx_patch.py
│   ├── test_normalize_contract.py
│   ├── test_batch_classifier.py
│   ├── test_batch_runner.py
│   ├── test_boilerplate_patterns.py
│   ├── test_header_footer_importer.py
│   ├── test_llm_classifier.py
│   ├── test_numpr_cascade.py
│   ├── test_page_layout_sync.py
│   ├── test_phase2_application_reporting.py
│   ├── test_phase2_bundle_and_validation.py
│   └── test_section_mapping.py
├── requirements.txt          # Full pinned dependencies (anthropic + customtkinter + transitive deps)
├── requirements-build.txt    # PyInstaller packaging dependencies
├── requirements-dev.txt      # Development: pytest>=7.0
├── DESIGN_SYSTEM.md          # GUI design system reference
├── README.md                 # Technical documentation
├── CLAUDE.md                 # This file
├── .gitignore                # Standard Python + project-specific ignores
└── phase2_classifications.json # Example LLM classification output
```

## Technology Stack

- **Python 3.7+** — all source code
- **Standard library** for core functionality: `zipfile`, `re`, `json`, `pathlib`, `hashlib`, `dataclasses`
- **anthropic==0.84.0** — external runtime dependency (for automated LLM classification)
- **customtkinter==5.2.2** — external runtime dependency (GUI)
- **No external XML libraries** — regex-based XML manipulation for byte-level fidelity
- **pytest** — development dependency for unit tests

## How to Run

### GUI (primary interface)
```bash
python gui.py
```

The pipeline has no standalone CLI. All user-facing processing (single-file and batch) is launched from the GUI. Programmatic callers should use `batch_runner.process_single_file()` or `batch_runner.run_batch_concurrent()` / `run_batch_api()` directly.

Default model: `claude-opus-4-6`

### Running Tests
```bash
python -m pytest tests/ -v
```

## Architecture and Data Flow

```
INPUT: target.docx + arch_template_registry.json + arch_style_registry.json + API key
  │
  ├─→ load_and_validate_shared_config()                 [batch_runner.py]
  │   └─→ preflight_validate_registries()               [core/registry.py]
  │
  ├─→ process_single_file() / run_batch_concurrent()    [batch_runner.py]
  │   │   (batch uses ThreadPoolExecutor; Batch API path uses run_batch_api())
  │   │
  │   ├─→ DocxDecomposer.extract()                      [docx_decomposer.py]
  │   │
  │   ├─→ apply_environment_to_target()                 [arch_env_applier.py]
  │   │   ├─→ apply_theme()                             theme/theme1.xml
  │   │   ├─→ apply_settings()                          compat flags
  │   │   ├─→ apply_font_table()                        fontTable.xml
  │   │   └─→ apply_doc_defaults()                      baseline rPr/pPr
  │   │
  │   ├─→ patch_footer_tokens()                         [header_footer_importer.py]
  │   │   └─→ imports architect headers/footers + media into extract dir
  │   │
  │   ├─→ import_numbering()                            [numbering_importer.py]
  │   │
  │   ├─→ import_arch_styles_into_target()              [core/style_import.py]
  │   │
  │   ├─→ build_phase2_slim_bundle()                    [core/classification.py]
  │   │   └─→ strip_boilerplate_with_report()           filter noise
  │   │
  │   ├─→ classify_target_document()                    [core/llm_classifier.py]
  │   │   ├─→ Anthropic API call (chunking + retry)
  │   │   └─→ Coverage check (warn if < 85%)
  │   │
  │   ├─→ snapshot_stability()                          [core/stability.py]
  │   │
  │   ├─→ apply_phase2_classifications()                [core/classification.py]
  │   │
  │   ├─→ verify_stability()                            [core/stability.py]
  │   │
  │   ├─→ _build_and_patch_output()                     [batch_runner.py]
  │   │   └─→ patch_docx()                              [docx_patch.py]
  │   │
  │   └─→ verify_phase2_invariants()                    [phase2_invariants.py]

OUTPUT: <target>_PHASE2_FORMATTED.docx
```

## Module Responsibilities

| Module | Role |
|---|---|
| `batch_runner.py` | Pipeline orchestrator: config loading, per-file processing, concurrent batch and Batch API runners |
| `docx_decomposer.py` | `DocxDecomposer` class — DOCX ZIP extraction only |
| `header_footer_importer.py` | Imports architect headers/footers (XML + media + rels) into the target extract directory |
| `arch_env_applier.py` | Imports formatting environment (theme, settings, fonts, docDefaults), Content-Types/rels provisioning, effective rPr resolution, style materialization for import |
| `numbering_importer.py` | Imports numbering definitions (abstractNum + num) with ID collision avoidance |
| `docx_patch.py` | Creates output DOCX via surgical ZIP entry replacement with XML well-formedness validation |
| `phase2_invariants.py` | Post-processing validation: sectPr, headers/footers, run properties |
| `gui.py` | customtkinter GUI with single-file and batch processing modes |
| `core/xml_helpers.py` | Paragraph XML iteration, text extraction, pStyle application, run font stripping |
| `core/stability.py` | SHA256 snapshots, header/footer/sectPr/rels stability verification |
| `core/style_import.py` | Style extraction, basedOn chain walking, property materialization, style import |
| `core/classification.py` | LLM prompts, boilerplate filtering, slim bundle building, classification application |
| `core/registry.py` | Architect registry loading, preflight reporting, path resolution |
| `core/llm_classifier.py` | Anthropic API integration (streaming), retry logic, chunking, coverage metrics |
| `core/batch_classifier.py` | Anthropic Batch API: request building, polling, result reassembly |
| `core/token_utils.py` | SectionID/SectionTitle token extraction and smart title-case utilities |
| `core/ooxml_namespaces.py` | OOXML namespace constants and ElementTree serialization helpers |
| `core/section_mapping.py` | Maps architect sectPr chain to target section count |
| `core/sectpr_tools.py` | sectPr block extraction, canonical child ordering, and tag-level manipulation |

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

1. **Headers/footers** — the target document's original headers/footers are replaced wholesale by the architect's headers/footers via `header_footer_importer.py`. After that replacement, no further drift is permitted.
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
- Core logic lives in `core/` package; `batch_runner.py` imports from it and from the top-level modules

## What NOT to Do

- **Do not use DOM/ElementTree for modifying Word XML** — the codebase intentionally uses regex to preserve byte-level fidelity
- **Do not create styles** — Phase 2 only applies styles defined in the architect registry
- **Do not modify headers, footers, or sectPr outside the designated import steps** — headers/footers are replaced by `header_footer_importer.py` during environment application, and sectPr is managed by `sectpr_tools.py`; do not touch them elsewhere
- **Do not apply run-level formatting** — all formatting is through `w:pStyle` only
- **Do not infer or guess style IDs** — if a role is missing from the registry, skip it
- **Do not modify `numbering.xml` during style application** — numbering preservation happens at the paragraph level via `<w:numPr>` materialization
- **Do not add heavy external dependencies** — core uses stdlib + anthropic only
- **Do not merge Phase 1 logic into this codebase** — Phase 1 and Phase 2 are separate concerns
- **Do not read the architect's extracted `word/` folder directly** — all architect data must come through the two JSON contract files via `build_arch_styles_xml_from_registry()`

## Allowed Patch Targets (docx_patch.py)

Core ZIP entries replaced in the output DOCX:
- `word/document.xml`
- `word/styles.xml`
- `word/theme/theme1.xml`
- `word/numbering.xml`
- `word/settings.xml`
- `word/fontTable.xml`
- `[Content_Types].xml`
- `word/_rels/document.xml.rels`

When `header_footer_importer.py` runs, the architect's header/footer parts, their `.rels` files, and any embedded media are also added to the replacement set. The target document's original header/footer entries are excluded from the output. Style application code must not touch headers or footers independently.

## Known Failure Modes

1. **Numbering stops on Enter** — numbering was style-linked; `w:pStyle` swapped without materializing `<w:numPr>`. Fix: ensure `ensure_explicit_numpr_from_current_style()` runs before restyling.
2. **Fonts change after styling** — imported style lacked explicit `<w:rPr>` due to inheritance. Fix: materialize effective properties when importing.
3. **Some paragraphs not styled** — role missing from registry or intentionally skipped (`SKIP`, `END_OF_SECTION`). Expected behavior, logged in preflight.
4. **Word opens with "Repair" warning** — invalid XML or broken basedOn chain. Check: style blocks are intact, all dependencies imported, no workspace artifacts in DOCX.
5. **LLM output truncated** — If the spec has 300+ paragraphs, the classification JSON may exceed the model's output limit. The classifier automatically chunks large documents. If coverage is still low on a large spec, try reducing the chunk size.
