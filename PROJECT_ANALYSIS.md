# Claude_Spec_Auto_Formatting — Project Analysis

**Analyzed:** March 7, 2026
**Project dates:** December 12–14, 2025 (3 days, 29 commits)
**Main file:** `docx_decomposer.py` — 2,663 lines, 89 functions

---

## Pattern 1: God-File Monolith

The entire project lives in a single 2,663-line Python file containing 89 functions. This file handles at least six distinct responsibilities:

- **DOCX extraction/reconstruction** (ZIP handling, file inventory)
- **Markdown analysis report generation** (~15 `_add_*` methods on the class)
- **LLM bundle building** (slim bundles, normalize bundles)
- **Architect style import/materialization** (style chain resolution, `basedOn` deps)
- **Phase 2 classification application** (paragraph restyling, numPr preservation)
- **Stability verification** (SHA-256 snapshots, sectPr/header/footer guards)

The two utility modules (`docx_patch.py` at 52 lines, `phase2_invariants.py` at 41 lines) are well-scoped and clean. They prove you *know* how to separate concerns — you just didn't finish doing it for the main file.

**Specific examples:**

1. `build_phase2_slim_bundle()` (line 2150) and `build_slim_bundle()` (line 1864) are two separate bundle builders in the same file with overlapping logic but different purposes. Neither calls the other.
2. The `DocxDecomposer` class (line 89) handles extraction, analysis, reconstruction, *and* LLM workflow orchestration — it's a class that does everything.
3. Boilerplate filtering (line 2057–2146) is a fully self-contained feature mixed into the middle of the file with no class boundary.

---

## Pattern 2: Aggressive Regex-Over-XML-Parsing

The codebase uses two fundamentally different strategies for reading Word XML:

- **ElementTree parsing** — used in the `DocxDecomposer` class's analysis methods (for generating markdown reports)
- **Raw regex on XML strings** — used for *all the production-critical operations*: style extraction, paragraph manipulation, numPr materialization, pStyle swapping

The regex approach is intentional (the code comments say it's to avoid ET rewriting/reformatting), and in the DOCX context this is defensible — Word is extremely sensitive to whitespace and attribute ordering, and ET will silently normalize those. But the regex patterns are fragile in specific ways:

**Specific examples:**

1. `iter_paragraph_xml_blocks()` (line 1670) uses `(<w:p\b[\s\S]*?</w:p>)` — a non-greedy match that will fail on nested `<w:p>` elements if they ever appear (e.g., inside structured document tags or text boxes).
2. `_extract_style_block()` (line 1462) uses `(<w:style\b[^>]*w:styleId="..."[\s\S]*?</w:style>)` — also non-greedy, so a style containing a nested element that ends with `</w:style>` could cause a short match.
3. The `ensure_explicit_numpr_from_current_style()` function (line 1501) does 4 different regex substitutions to handle different XML shapes, which is correct but makes it a prime candidate for a unit test suite.

---

## Pattern 3: Commit History Tells a Story of Frustration

The commit messages are raw and honest. Reading them chronologically reveals a clear emotional arc:

**Day 1 (Dec 12) — 12 commits: Optimism → Hitting the Wall**
- Starts confident: "first commit", "mostly working", "ready to run program"
- Hits trouble: "starting over", "didn't work, try again tomorrow"

**Day 2 (Dec 13) — 9 commits: Grinding Through It**
- Progress: "finished adding new funcs", "mech spec bundle extracted"
- LLM integration works: "llm returned classification schema"
- Then collapse: "unsuccessful run", "trying to unfuck this mess", "taking a break from this shit"

**Day 3 (Dec 14) — 8 commits: Pushing Through → Giving Up**
- New energy: "starting new test", "trim legacy code"
- Escalating frustration: "let's fucking finish already", "super lost", "idk"

The last two commits ("super lost", "idk") are on a branch `12-14-2025`, not merged to master. This project was set aside mid-struggle.

---

## Pattern 4: Legacy Code Kept But Disabled

There's a significant amount of code that's been intentionally disabled but not removed:

**Specific examples:**

1. `--normalize`, `--apply-edits`, `--normalize-slim`, `--apply-instructions` CLI flags (lines 1188–1191) are explicitly marked as `(LEGACY) disabled` but still parsed and handled with error messages.
2. `write_normalize_bundle()` (line 1057) and `write_slim_normalize_bundle()` (line 1118) are methods on the class that appear unreachable from `main()` since the legacy modes are disabled.
3. `apply_edits_and_rebuild()` (line 1090) and `apply_instructions_and_rebuild()` (line 1140) are also orphaned.
4. `MASTER_PROMPT` and `RUN_INSTRUCTION_DEFAULT` are referenced at line 1083–1084 but never defined in the file (they must have existed in a previous version). The prompts that DO exist are `PHASE2_MASTER_PROMPT` and `PHASE2_RUN_INSTRUCTION` (lines 47–84).
5. The `validate_instructions()` function (line 2452) and `apply_instructions()` (line 2513) are part of the legacy slim-bundle workflow. `apply_instructions()` has a bug on line 2570 — it references `style_id` (from the loop in `apply_instructions`) but should reference `sid` from `idx_map`. This was likely never caught because this code path is disabled.

---

## Pattern 5: Strong Design Philosophy, Incomplete Execution

The README is one of the best-written technical design docs I've seen for a personal project. It clearly defines:

- Hard invariants (headers/footers unchanged, sectPr untouched, no run-level formatting)
- The "change as little XML as possible" philosophy
- Explicit non-goals (no style creation, no numbering.xml edits)
- Known failure modes with root causes and fixes

The code actually implements these guardrails well — `verify_stability()`, `verify_phase2_invariants()`, the `_DOCX_ALLOWED_TOP_LEVEL_DIRS` filter, and the `ALLOWED_EDIT_PATHS` whitelist are all solid defensive patterns.

But the gap between the README's vision and the code's state is visible:

1. The README mentions a "Phase 1" that produces `arch_style_registry.json`, but Phase 1 doesn't exist in this repo. Phase 2 requires it as input.
2. The "Acceptance Checklist" in the README has no automated implementation — it's manual checkboxes.
3. The cleanup guidance ("Remove: heuristic code paths, legacy analysis reports, unused CLI modes") hasn't been done.

---

## Completion Status

**What works:**
- DOCX extraction and safe reconstruction via ZIP patching
- Phase 2 slim bundle generation (for sending to LLM)
- LLM classification consumption and paragraph restyling
- Architect style import with full `basedOn` dependency resolution
- Style materialization (inherited pPr/rPr made explicit)
- Dynamic numbering preservation (numPr materialization before style swap)
- Multi-layer stability verification (headers, footers, sectPr, rels, rPr)

**What's broken or incomplete:**
- Phase 1 (style extraction from architect templates) — doesn't exist here
- Legacy code paths have a bug at line 2570 (`style_id` should be `sid`)
- No test suite at all — every run is manual
- The `mech_extracted_analysis.md` is 838KB — this bloated report generation was the original approach before the slim bundle strategy
- The `12-14-2025` branch was never merged back; unclear what diverged

**What's left per the README:**
- Remove legacy code paths (normalize, apply-edits, normalize-slim, apply-instructions)
- Remove the massive analysis report generator (the `_add_*` methods)
- Build Phase 1 (or document the manual process)
- Write tests for the invariant checks and paragraph manipulation functions

---

## Recommendations (Prioritized)

### 1. Split `docx_decomposer.py` into modules
Suggested structure:
```
core/
  extraction.py       — DocxDecomposer class (extract + reconstruct only)
  xml_helpers.py      — iter_paragraph_xml_blocks, paragraph_text_from_block, etc.
  stability.py        — StabilitySnapshot, verify_stability, verify_phase2_invariants
  style_import.py     — import_arch_styles_into_target, materialize_arch_style_block
  classification.py   — apply_phase2_classifications, build_phase2_slim_bundle
  boilerplate.py      — BOILERPLATE_PATTERNS, strip_boilerplate_with_report
docx_patch.py         — (keep as-is)
main.py               — CLI entry point
```

### 2. Delete dead code
Remove: `write_normalize_bundle`, `write_slim_normalize_bundle`, `apply_edits_and_rebuild`, `apply_instructions_and_rebuild`, `apply_instructions`, `validate_instructions`, `build_llm_bundle`, `build_slim_bundle`, `apply_llm_edits`, `SLIM_MASTER_PROMPT`/`SLIM_RUN_INSTRUCTION_DEFAULT` references, all `(LEGACY)` CLI args, and the 15+ `_add_*` analysis methods. This would cut the file roughly in half.

### 3. Fix the bug in `apply_instructions`
Line 2570: `para_blocks[idx] = apply_pstyle_to_paragraph_block(pb, style_id)` should be `apply_pstyle_to_paragraph_block(pb, sid)`. Even though this code path is disabled, fix it before you forget.

### 4. Write tests for the regex-heavy functions
Priority targets: `ensure_explicit_numpr_from_current_style`, `apply_pstyle_to_paragraph_block`, `iter_paragraph_xml_blocks`, `_find_style_numpr_in_chain`. These are the functions where a subtle XML variation could cause silent corruption.

### 5. Decide on Phase 1
Either build it as a separate tool in this repo, or document the manual process for creating `arch_style_registry.json`. Right now this is a hidden dependency.

### 6. Add PyInstaller config or build script
Your `requirements.txt` includes `pyinstaller`, suggesting you planned to ship this as an executable. No build config exists yet.
