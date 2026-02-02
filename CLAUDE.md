# CLAUDE.md — AI Assistant Guide for Claude_Spec_Auto_Formatting

## Project Overview

This is a **Phase 2 MEP Specification Styling Engine** for the AEC (Architecture/Engineering/Construction) industry. It applies architect-defined CSI (Construction Specifications Institute) paragraph styles to MEP (Mechanical, Electrical, Plumbing) specification documents (.docx) while preserving exact Word behavior.

**Core principle:** Change as little Word XML as possible while achieving exact visual and behavioral alignment with the architect's template.

Phase 1 (external, not in this repo) extracts and catalogs styles from an architect's template. Phase 2 (this repo) applies those styles to target MEP specifications deterministically.

## Repository Structure

```
Claude_Spec_Auto_Formatting/
├── docx_decomposer.py        # Main orchestrator (1569 lines) — CLI entry point
├── arch_env_applier.py        # Formatting environment application (657 lines)
├── numbering_importer.py      # Numbering definition import with collision avoidance (393 lines)
├── docx_patch.py              # Surgical ZIP patching for output DOCX (93 lines)
├── phase2_invariants.py       # Post-processing invariant verification (105 lines)
├── README.md                  # Technical documentation
├── requirements.txt           # Python dependencies (PyInstaller packaging only)
├── .gitignore                 # Standard Python + project-specific ignores
├── FIRE_SPEC.docx             # Example specification document
├── MECH_SPEC.docx             # Example specification document
├── PLUMB_SPEC.docx            # Example specification document
├── phase2_classifications.json # Example LLM classification output
└── NVES_extracted/            # Example extracted architect template
    ├── [Content_Types].xml
    ├── _rels/
    ├── docProps/
    ├── word/
    │   ├── document.xml
    │   ├── styles.xml
    │   ├── numbering.xml
    │   ├── theme/theme1.xml
    │   ├── fontTable.xml
    │   ├── settings.xml
    │   ├── header1.xml, footer1.xml
    │   ├── footnotes.xml, endnotes.xml
    │   ├── webSettings.xml
    │   └── _rels/document.xml.rels
    ├── arch_style_registry.json     # CSI role → style mapping
    ├── arch_template_registry.json  # Complete environment data
    └── slim_bundle.json             # LLM input bundle
```

## Technology Stack

- **Python 3.7+** — all source code
- **Standard library only** for core functionality: `zipfile`, `re`, `json`, `pathlib`, `xml.etree.ElementTree`, `hashlib`, `dataclasses`
- **No external XML libraries** — regex-based XML manipulation for performance and minimal dependencies
- **requirements.txt** contains PyInstaller dependencies for optional .exe packaging, not runtime dependencies

## How to Run

### Phase 2, Mode 1: Build LLM input bundle
```bash
python docx_decomposer.py <target.docx> \
  --phase2-arch-extract <arch_extracted_folder> \
  --phase2-build-bundle
```

### Phase 2, Mode 2: Apply LLM classifications
```bash
python docx_decomposer.py <target.docx> \
  --phase2-arch-extract <arch_extracted_folder> \
  --phase2-classifications <phase2_classifications.json> \
  [--output-docx <output.docx>]
```

### Optional flags
- `--use-extract-dir <dir>` — reuse an existing extracted folder (skip extraction)
- `--phase2-discipline mechanical|plumbing` — discipline (default: mechanical)
- `--extract-dir <dir>` — specify extraction location

There are no automated tests, linting, or CI/CD pipelines. Validation is done manually by opening output DOCX files in Word and verifying against the acceptance checklist in the README.

## Architecture and Data Flow

```
INPUT: target.docx + arch_template_registry.json + phase2_classifications.json
  │
  ├─→ Extract DOCX to folder (DocxDecomposer.extract)
  │
  ├─→ apply_environment_to_target()        [arch_env_applier.py]
  │   ├─→ apply_theme()                    theme/theme1.xml
  │   ├─→ apply_settings()                 compat flags
  │   ├─→ apply_font_table()               fontTable.xml
  │   └─→ apply_doc_defaults()             baseline rPr/pPr
  │
  ├─→ import_numbering()                   [numbering_importer.py]
  │   ├─→ find_max_ids_in_numbering()      collision detection
  │   ├─→ build_numbering_import_plan()    remapping strategy
  │   └─→ inject_numbering_into_xml()      merge abstractNums + nums
  │
  ├─→ import_arch_styles_into_target()     [docx_decomposer.py]
  │   ├─→ get_styles_with_dependencies()   expand basedOn chain
  │   ├─→ materialize_arch_style_block()   harden for portability
  │   └─→ insert into target styles.xml
  │
  ├─→ snapshot_stability()                 [before document.xml edits]
  │
  ├─→ apply_phase2_classifications()       [docx_decomposer.py]
  │   ├─→ ensure_explicit_numpr_from_current_style()
  │   ├─→ strip_run_font_formatting()
  │   └─→ apply_pstyle_to_paragraph_block()
  │
  ├─→ verify_stability()                   [after document.xml edits]
  │
  ├─→ patch_docx()                         [docx_patch.py]
  │   └─→ surgical ZIP replacement (only allowed files)
  │
  └─→ verify_phase2_invariants()           [phase2_invariants.py]

OUTPUT: <target>_PHASE2_FORMATTED.docx
```

## Module Responsibilities

| Module | Role |
|---|---|
| `docx_decomposer.py` | CLI entry point, orchestrator, DOCX extraction, slim bundle generation, style application, paragraph processing, boilerplate detection |
| `arch_env_applier.py` | Imports formatting environment (theme, settings, fonts, docDefaults) from architect template into target |
| `numbering_importer.py` | Imports numbering definitions (abstractNum + num) with ID collision avoidance |
| `docx_patch.py` | Creates output DOCX via surgical ZIP entry replacement with strict allow-list |
| `phase2_invariants.py` | Post-processing validation: sectPr unchanged, headers/footers byte-identical, run properties preserved |

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
- `PART` — part headings (PART 1, PART 2, PART 3)
- `ARTICLE` — article numbers (1.01, 2.03)
- `PARAGRAPH` — lettered paragraphs (A., B., C.)
- `SUBPARAGRAPH` — numbered under letters (1., 2., 3.)
- `SUBSUBPARAGRAPH` — lettered under numbers (a., b., c.)

## Key Input/Output Files

### Inputs (from Phase 1 / LLM)
- **`arch_style_registry.json`** — Maps CSI roles to Word style IDs. This is the sole source of truth for styling decisions. No heuristics or guessing allowed.
- **`arch_template_registry.json`** — Complete formatting environment: theme, all styles with materialized properties, numbering definitions, docDefaults, font declarations, settings.
- **`phase2_classifications.json`** — LLM output: `{ "classifications": [{ "paragraph_index": N, "csi_role": "ROLE" }] }`
- **`slim_bundle.json`** — LLM input: paragraph text + indices only (no XML or formatting).

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
Used via `typing` module (Dict, List, Optional, Tuple, etc.) on function signatures. Not comprehensive throughout.

### Error Handling
Mix of try/except and explicit validation. Errors are logged into `log: List[str]` arrays passed through the call chain, and fatal invariant violations raise `RuntimeError`.

### Imports
- Standard library imports first, then local module imports
- `numbering_importer` is imported conditionally with a `HAS_NUMBERING_IMPORTER` flag fallback
- Imports inside `main()` for argparse-only dependencies

## What NOT to Do

- **Do not use DOM/ElementTree for modifying Word XML** — the codebase intentionally uses regex to preserve byte-level fidelity
- **Do not create styles** — Phase 2 only applies styles defined in the architect registry
- **Do not modify headers, footers, or sectPr** — these are protected by hard invariants
- **Do not apply run-level formatting** — all formatting is through `w:pStyle` only
- **Do not infer or guess style IDs** — if a role is missing from the registry, skip it
- **Do not modify `numbering.xml` during style application** — numbering preservation happens at the paragraph level via `<w:numPr>` materialization
- **Do not add external dependencies** — the core functionality uses stdlib only
- **Do not merge Phase 1 logic into this codebase** — Phase 1 and Phase 2 are separate concerns

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
