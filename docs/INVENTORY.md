# Module inventory

26 working modules in `src/`, plus 2 non-functional ones in `src/legacy/`.
Grouped by role rather than alphabetically — the grouping reflects how data
actually flows through the system.

---

## Core pipeline

### `pipeline_pdf_to_excel.py` (266 lines)
**Orchestrator.** Runs OCR → name correction → statement generation as one
command, streaming each stage's log.
- **In**: folder of PDFs, output path, Excel template
- **Out**: `.xlsx`/`.xlsm` statement, plus two intermediate folders
  (`*_txt_raw`, `*_txt_fixed`) kept deliberately so a bad row can be traced to
  the stage that produced it
- **Key decisions**: a thin dispatcher, not a merge — each stage stays runnable
  on its own (the common case is re-running only one). Child stages are launched
  as direct `python.exe` processes, never through a shell. Stops the chain and
  names the failing stage rather than continuing with partial data.

### `pdf_to_text_multithread.py` (1,253 lines)
**OCR stage.** Renders PDF pages, runs three recognition passes, merges them.
- **In**: folder of PDFs; flags `--preprocess`, `--no-magick`, `--min-conf`,
  `--fuzzy-threshold`, `--save-training`
- **Out**: one `.txt` per report that passes the section filter
- **Key decisions**:
  - **Line-level consensus** across the three passes. Character-positional
    voting was measurably destructive — see README §1.
  - **Sequential PDF→PNG phase before parallel OCR.** `pdftoppm` competing with
    three OCR workers hit 240 s timeouts on files that render in ~7 s alone.
  - **Per-page timeout 300 s.** At 30 s, whole pages were being dropped with
    only a warning — 19 % vs 85 % measured accuracy (README §2).
  - **Safe DPI** computed from page geometry, capping the long edge at ~3300 px,
    so iOS-scan PDFs (1782×2586 pt) don't render into 80-megapixel images.
  - TSV word-coordinate reconstruction rebuilds reading order from geometry
    instead of trusting `--psm`; large horizontal gaps become tabs so column
    structure survives.

### `ocr_text_corrector.py` (400 lines)
**Name correction between OCR and statement generation.**
- **In**: folder of `.txt`, Excel template (for the catalogue); flag `--llm`
- **Out**: folder of corrected `.txt`
- **Key decisions**: three escalating tiers — deterministic glyph/unit-code
  fixes → catalogue match at a **high** threshold (92) → optional LLM. The
  threshold is high on purpose: at 72 it produced semantically wrong
  substitutions that *looked* successful in the log (README §3). The LLM is
  lazily started, so if the catalogue handles everything the 4 GB model is never
  loaded. Results cached in `Data/ocr_corrections.json`.

### `excel_report_generator.py` (1,803 lines)
**Statement generation** — the largest and most domain-specific module.
- **In**: folder of `.txt`, output path, Excel template with sheets
  `Макет відомості на списання`, `Словник`, `синоніми`, `Словник_підрозділи`,
  `Ціни`; flags `--threshold`, `--history-folder`
- **Out**: workbook with one sheet per report, one block per episode, filled
  form fields, matched names/units/prices, colour-coded confidence
- **Key decisions**:
  - **Content-based section boundary** (README §4) with page-furniture skipping
    and a two-miss rule.
  - **Three-level price confidence** — green / yellow+comment / red, and
    uncertain matches are not cached so the flag can't vanish.
  - **Synonym sheet outranks the catalogue** on a near-exact variant match; a
    curated human mapping should beat a fuzzy hit. Without this, a shortened
    name that also exists verbatim in the catalogue matched itself at 100 % and
    the mapping never fired.
  - **Strict-name list** for entries too easily confused with others.
  - Quantity glyph decoding (`|`→1, `б`→6) — `|` used to make the whole line
    unparseable and drop the item; `б` silently became 1 instead of 6, which is
    worse because the error is invisible.
  - Block height read from the template's **print area**, not `max_row`
    (openpyxl counts stray formatting, which inflated blocks to hundreds of rows).

### `image_preprocessor.py` (1,158 lines)
**Image conditioning** for photographed documents; used standalone for
handwritten record cards and as `--preprocess` inside the OCR stage.
- **In**: folder of images (recursive), output folder; flags `--preset`,
  `--no-probe`, `--workers`
- **Out**: 8-bit grayscale PNGs, optional gamma variant and row-strip crops
- **Key decisions**:
  - Geometry **first**: content-based rotation (EXIF as hypothesis, OCR
    orientation as verification), perspective dewarp from table corners,
    cylindrical dewarp, then fine deskew.
  - **Every geometric step must earn its place** — accepted only if it raises a
    grid-regularity score *and* doesn't crop the table away. Without those two
    guards a mis-fitted quadrilateral either shredded the page or silently cut
    17–41 % of the grid while the regularity score *improved*.
  - **Flat-field** background division rather than global `normalize` (which
    blows out the centre and leaves dark corners); paper normalised to ~245, not
    255, so faint pencil survives.
  - No binarisation, no median blur — both destroy faint strokes and punctuation.
  - Measures digit stroke height and warns when it is below the readable
    threshold, i.e. tells the operator to re-shoot instead of pretending filters
    can recover it.

---

## Catalogue maintenance

### `collect_unique_items.py` (492 lines)
Mines unique item names from a folder of completed statements and appends the
missing ones to the catalogue.
- **In**: folder of `.xlsx` statements, Excel template, optional official
  nomenclature text (`--nakaz`); `--apply` to write
- **Out**: catalogue rows appended **at the end** (so existing unit/service-life
  columns stay aligned with their names), colour-coded by provenance — confirmed
  by official source / frequent / rare — plus a comment column and a CSV report
- **Key decisions**: name normalisation strips re-valuation year suffixes and
  merges case/whitespace variants (85 collapsed on real data; without it 30
  duplicates entered the catalogue, and a duplicate is harmful because fuzzy
  matching then picks between them arbitrarily). Unit of measure is taken from
  the official source when confirmed, otherwise from the statement. Service life
  is **never** invented. The 1,700-file scan is cached so re-runs are instant.

### `cleanup_dictionary.py` (254 lines)
Removes non-nomenclature content and canonicalises the catalogue.
- **In**: Excel template; `--apply`
- **Out**: catalogue edited in place, timestamped backup alongside
- **Key decisions**: edits **names only**, in place, then deletes duplicate rows
  bottom-up so row numbers don't shift; when merging duplicates the surviving
  row is the one with the most complete data and missing unit/service-life is
  carried over from the rows being dropped. Explicitly does **not** treat "short
  name" as "incomplete" — verified against real data where several correct
  official names are two words long.

### `universal_dict_extractor.py` (399 lines)
Recursively pulls dictionary entries from mixed `.xlsx`/`.xlsm`/`.docx` sources
into a template sheet, with a configurable match threshold and a backup of the
original.

---

## Verification and cross-checking

### `doc_processor.py` (2,243 lines)
Processes property-loss documents against a units dictionary and a balance
table; flags shortages and supports an interactive disambiguation mode where the
operator picks the right candidate (protocol-driven, so the GUI can render it as
a dialog and answer over stdin).
- **In**: dictionary `.xlsx`, balances `.xlsx`, input folder, output folder;
  `--no-check-remains`, `--manual-match`
- **Out**: processed `.docx` with colour-coded issues

### `excel_to_word_transfer.py` (320 lines)
Python replacement for a VBA `TransferByPageBreaks` macro: fills a Word template
from statement sheets. Imported directly (not spawned) by the GUI, which
therefore honestly reports that it cannot be force-stopped mid-write rather than
faking a cancel and risking a half-written `.docx`.

### `vidomist_report.py` / `run_vidomist.py` (786 / 268 lines)
Collects data across statement (`.xlsm`) and extract (`.doc`) files, recursively;
used to reconcile what was issued against what was written off.

### `validator.py` (214 lines)
Validates the JSON reports produced by the extraction path — root-element and
unit-of-measure checks.

### `price_matcher.py` (658 lines)
Standalone price matching from a separate price workbook into a chosen sheet.
Its logic was later folded into `excel_report_generator.py` as a final pass; the
module remains for one-off use against an external price file.

---

## Accounting-side utilities

| Module | Purpose |
|---|---|
| `oblik_common.py` (466) | shared helpers for the inventory scripts: normalisation, workbook conventions, journal formats |
| `build_master_workbook.py` (133) | creates the master inventory workbook structure from scratch (`Склади`, `Підрозділи`, balances, reconciliation) |
| `build_unit_map.py` (502) | maps subdivision designations between documents; handles the abbreviation variants used in practice |
| `import_balances_docx.py` (105) | imports one `.docx` balance statement |
| `import_pidrozdily_xlsm.py` (122) | imports a monthly `.xlsm` balance file (one sheet per subdivision) |
| `cards_photo_to_excel.py` (596) | converts photographed **handwritten** record cards to a flat Excel journal via a vision model. Separate on purpose: Tesseract does not read handwriting regardless of preprocessing, and the text-only local LLM cannot see images. Keys come from the environment, never hardcoded. |
| `main_processor.py` (805) | earlier end-to-end driver, superseded by `pipeline_pdf_to_excel.py` |
| `sort_vidomosti_gui.py` (213) | sorts a mixed folder into the statement structure |

---

## Interfaces

| Module | Purpose |
|---|---|
| `run_gui_unified.py` (1,962) | tabbed Tkinter launcher for all tools. Child scripts run as direct `python.exe` processes — no `cmd`/`powershell`/`bash` anywhere. Live log streaming per tab; the stop button kills the **entire process tree** (`taskkill /T`), because an orphaned model server otherwise keeps gigabytes of RAM. Includes an interactive dialog for the manual-match protocol. |
| `launcher.py` (337) | simple menu-style launcher |
| `run_gui_image_preprocessor.py` (310) | single-purpose window for the image preprocessor |
| `run_vidomist.py` (268) | launcher for the statement collector |

---

## `src/legacy/` — not functional

| Module | Missing import | Status |
|---|---|---|
| `gui_launcher.py` | `python_src.main_processor` | fails on import |
| `batch_processor.py` | `excel_utils.excel_writer` | fails on import |

Both reference an older package layout that no longer exists. Kept for
completeness, excluded from every documented entry point.
