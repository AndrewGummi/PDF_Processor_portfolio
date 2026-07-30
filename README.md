# Document Digitization Pipeline — scanned reports → structured Excel

A production pipeline that turns **photographed / scanned paper reports** into a
**structured Excel statement** with matched catalogue names, units and prices.

Built for a real internal workflow where the same task was done by hand:
open a scan, read item names off the page, find each one in a ~1,200-entry
catalogue, copy the price, type it into a statement form. Roughly 10–15 minutes
per document, several dozen documents per batch, and transcription errors that
surface weeks later during audit.

The pipeline does it in about **30 seconds per document**, and — more
importantly — **marks its own uncertainty** instead of silently guessing.

```
scanned PDF ──▶ OCR ──▶ name correction ──▶ statement generation ──▶ Word cross-check
                (Tesseract,   (dictionary +    (Excel form fill,       (docx export /
                 3 variants +  synonyms +       fuzzy price match,      verification)
                 voting)       optional LLM)    3-level confidence)
```

---

## Why this is not "just run OCR on it"

Every naive step in this problem has a failure mode that only shows up on real
documents. The interesting engineering is in the guards, and each one below was
added in response to a measured failure, not a hypothetical:

### 1. Character-level voting across OCR variants destroys correct text

The pipeline runs three OCR passes per page (Tesseract `--psm 3`, a TSV
layout-reconstruction pass, and an ImageMagick-preprocessed pass) and merges
them. The original merge compared them **positionally, character by character**.

But the variants are not aligned — on one page they were 1150 / 1137 / 1139
characters long, because each pass emits different line breaks. One extra
character early on shifts everything after it, and "voting" then stitches
letters from unrelated parts of the document into words that appeared in *no*
variant:

| fragment | each variant alone | after positional voting |
|---|---|---|
| `газовий балон 11 кг` | correct in all three | `балон 11 к` |
| `(50 м)` | correct in all three | `( 0 м)` |
| `продовольча` | correct in all three | `поооовольча` |

Positional voting scored **82.8 %** against ground truth — *below every
individual variant*. It was actively destroying data, and it silently corrupted
the digits inside item names, which are exactly the values that end up in the
statement.

Replaced with **line-level consensus**: the highest-confidence variant is the
base, each of its lines is matched against the most similar line in the others,
and a majority wins. The key property is that **every output line is a line some
OCR pass actually produced** — never a synthetic character blend. Accuracy:
**85.5 %**, on par with the best single pass and immune to the corruption class.

### 2. A 30-second OCR timeout was silently dropping entire pages

Per-page OCR had a 30 s limit. On a 4-core CPU under three parallel workers, an
A4 page at ~2330×3300 px does not always finish in 30 s. Tesseract was killed,
the code logged a warning, and **the page simply never appeared in the output** —
a two-page report yielded 183 characters instead of ~1400.

Measured against a PDF with a real text layer as ground truth:

| | accuracy |
|---|---|
| 30 s timeout (page silently lost) | **19 %** |
| 300 s timeout | **85 %** |

Accuracy is the priority here, not throughput, so the limit is now generous and
page loss is reported loudly.

### 3. String similarity is not semantic identity

The catalogue contains eight `Чоботи гумові …` ("rubber boots") entries, each
with a qualifier, and one bare `Чоботи хромові` ("chrome-leather boots"). A
report that says just `чоботи гумові` scores:

```
Чоботи хромові            81.5 %   ← wrong product, but shortest edit distance
Чоботи гумові КАНАДА      78.8 %
Чоботи гумові утеплені    74.3 %   ← correct family, penalised for being longer
```

Levenshtein measures **shared characters**; the meaning lives in the one word
that differs. Every extra qualifier word makes the *correct* entry score
*worse*. Raising or lowering a single global threshold cannot fix this.

The resolution is three-tiered:

- **≥ 92 %** — accept automatically (that range is genuinely just OCR noise).
- **45–92 %** — hand to a human-curated **synonym sheet**, or leave untouched.
  Guessing here is how a backpack turned into a cooking pot: an OCR-mangled
  `рюкзак індив дуальтий` matched an unrelated `Казанок ...` entry at 77 %.
- **strict list** — specific entries that are too easily confused with others
  are only eligible on a near-exact match, never on "looks similar".

The synonym sheet is the important part: it is explicit, instant, auditable, and
**cannot invent**. One human decision fixes that name permanently, for every
future document.

### 4. Where the section ends has to be decided by content, not headers

Only one service's items belong in the statement. The original code stopped at a
hardcoded list of other-service headings — which fails when OCR mangles the
heading, or when there is no heading at all (closing paragraph, signature block,
page footer). Items from other services leaked into the statement.

Now each line's leading words are scored against the item catalogue. Measured
across 10 real reports:

| | similarity to catalogue |
|---|---|
| lines inside the section | **78.4 – 100** |
| lines outside it | **29.6 – 57.9** |

A 20.5-point gap, so the cutoff sits safely at 68. Two additional guards were
needed once this was live:

- **page furniture is skipped, not treated as a boundary.** When a list spans a
  page break, the footer of one page and the header of the next sit between its
  halves. The first version read those two lines as "section over" and lost
  **11 real items** from a 25-item list.
- **one unrecognised line is not a boundary** — two consecutive are. A single
  badly-mangled item line shouldn't truncate the rest of the section.

### 5. The local LLM earns its keep in exactly one narrow role

A local `llama-server` + 7B GGUF model is wired in as an optional last resort.
It was benchmarked on eight real failure cases rather than assumed useful:

| outcome | count | example |
|---|---|---|
| genuinely correct where fuzzy matching failed | **1** | OCR-mangled `рюкзак індив дуальтий` → the correct `Рюкзак ...` entry (80 s) |
| honest refusal (`NONE`) | 1 | two items merged into one line |
| partial — found the items but **fabricated** a quantity | 2 | invented "1 pc." that was not in the input |
| fabricated or made it worse | 4 | a bare "protective goggles" gained an invented country of origin; a mangled Latin standard code came back worse than the input |

Per-call cost: **80–673 seconds** on CPU.

The pattern is consistent and worth stating plainly: the model is useful when
**the answer exists but is corrupted**, and harmful when **the document is
genuinely ambiguous** — because it cannot answer "insufficient data", so it
guesses, and a guess is indistinguishable from an answer.

Consequently the LLM stage is **opt-in** (`--llm`), its output is always marked
for review, and every result is cached so repeat names cost nothing. Expanding
the catalogue from 757 to 1,236 entries — mined automatically from 1,741
already-completed statements — fixed **21 names instantly and with zero
fabrication**, which is far more leverage than the model provides.

---

## Three-level confidence, instead of a boolean

Prices are matched fuzzily against a price sheet whose entries carry qualifiers
the reports omit (`Матрац` vs `Матрац, розмір 1850*650*60` → 63 %). A single
threshold either leaves most cells empty or fills them with wrong numbers, so
the output encodes confidence:

| fill | match | meaning |
|---|---|---|
| green | ≥ 90 % | trust it |
| **yellow** | 75–90 % | a price *is* filled in, plus an Excel comment naming the exact source row — check it |
| red | < 75 % | left empty on purpose |

Yellow matters. In the demo run below, `чоботи гумові` was filled with the price
of `Чоботи хромові` at 81 % — the wrong product — and the cell carries the
comment `НЕТОЧНИЙ збіг 81% — перевір | Взято: Чоботи хромові`. The mistake is
visible in a second instead of hiding in a signed document. Uncertain matches
are also deliberately **not cached**, so the flag cannot quietly disappear on
the next run.

---

## Demo on synthetic data

`sample_data/make_sample_data.py` generates everything from scratch: an Excel
template, a catalogue with the ambiguity traps described above, and two
"scanned" reports rendered as **image-only PDFs** with tilt, uneven lighting and
sensor noise, so the OCR stage does real work.

```bash
python sample_data/make_sample_data.py

set TESSERACT_EXE=C:\Program Files\Tesseract-OCR\tesseract.exe
set POPPLER_BIN=C:\poppler\Library\bin

python src/pipeline_pdf_to_excel.py ^
    sample_data/reports_pdf ^
    sample_data/out/statement.xlsx ^
    sample_data/template_sample.xlsx --preprocess
```

**Input** — what OCR actually read off the degraded page (note the mangled
site code `АОООО`, and that three different sections are present):

```
речова служба складу об'єкта АОООО:        ← the section we want
знищено:
казанок туристичний - 1 шт.;
мішок спальний літній - 2 шт.;
окуляри захисні - 4 к-т;          ← ambiguous short form
чоботи гумові - 4 пари;           ← no exact catalogue entry exists
матрац - 1 шт.;                   ← catalogue name carries a size suffix
чохол інструментальний - 1 шт.;
медична служба складу об'єкта АО0О0О:      ← different section - must be excluded
знищено:
аптечка - 2 шт.;
служба зв'язку складу об'єкта АОООО:       ← different section - must be excluded
знищено:
антена виносна - 1 шт.;
```

**Output** — sheet `sample_report_1`:

| № | Item | Unit | Qty | Price | |
|---|------|------|-----|-------|--|
| 1 | Казанок туристичний | шт. | 1 | 487.80 | |
| 2 | Мішок спальний літній | шт. | 2 | 910.92 | |
| 3 | **Окуляри захисні прозорі** | к-т | 4 | 1566.65 | resolved via synonym sheet — *not* `Окуляри світлозахисні` (640.00) |
| 4 | чоботи гумові | пари | 4 | 980.00 | 🟡 `81% — Взято: Чоботи хромові` — flagged, wrong product |
| 5 | **Матрац, розмір 1850\*650\*60** | шт | 1 | 924.78 | short form expanded to the priced catalogue entry |
| 6 | **Чохол інструментальний універсальний** | шт | 1 | 621.30 | idem |

Items from the other two sections were correctly excluded by the
content-based section boundary.

Run: **2 PDFs → 63 s** total (58 s OCR, 2 s correction, 3 s statement),
13 items, 13 prices matched, 1 flagged for review.

---

## Architecture

| Stage | Module | In → Out |
|---|---|---|
| 1. Image conditioning *(optional)* | `image_preprocessor.py` | photo/scan → deskewed, flat-fielded, CLAHE-enhanced PNG |
| 2. OCR | `pdf_to_text_multithread.py` | PDF → `.txt` (3 variants + line-level consensus) |
| 3. Name correction | `ocr_text_corrector.py` | `.txt` → `.txt` (dictionary → synonyms → optional LLM) |
| 4. Statement generation | `excel_report_generator.py` | `.txt` + template → `.xlsx` / `.xlsm` |
| 5. Orchestration | `pipeline_pdf_to_excel.py` | runs 2→3→4 as one command |
| 6. Word cross-check | `excel_to_word_transfer.py`, `doc_processor.py` | statement → `.docx`, verification against balance sheets |

**Catalogue maintenance** (the highest-leverage part of the system):

| Module | Purpose |
|---|---|
| `collect_unique_items.py` | mines unique item names from completed statements, cross-checks them against an official nomenclature PDF, appends the missing ones with units and colour-coded provenance |
| `cleanup_dictionary.py` | strips non-nomenclature noise (annotations, double spaces), canonicalises variant spellings, merges duplicates while preserving service-life data |
| `universal_dict_extractor.py` | pulls dictionary entries out of mixed `.xlsx` / `.docx` sources |

**Interfaces**: `run_gui_unified.py` is a tabbed Tkinter launcher for all tools
(live log streaming, hard-stop that kills the whole child process tree so an
orphaned model server cannot keep gigabytes of RAM). `launcher.py` and the
`run_*.py` scripts are single-purpose windows.

---

## Tech stack

- **Python 3.11+**
- **OCR**: Tesseract (`--psm 3` + TSV word-coordinate reconstruction), Poppler
  (`pdftoppm`/`pdfinfo`), ImageMagick
- **Imaging**: OpenCV, NumPy — perspective/cylindrical dewarp, flat-field
  correction, CLAHE, threshold-aware unsharp
- **Fuzzy matching**: RapidFuzz, `pymorphy3` for Ukrainian lemmatisation
- **Excel/Word**: openpyxl (VBA-preserving `.xlsm` round-trip), python-docx
- **Local LLM**: `llama-server` + GGUF, called over HTTP on localhost; Levenshtein
  pre-filter cuts ~1,200 candidates to ~10 before any model call, so the
  expensive step sees a shortlist rather than the whole catalogue
- **Concurrency**: `ThreadPoolExecutor`; PDF→PNG runs sequentially *before* the
  parallel OCR phase, because `pdftoppm` competing with three OCR workers caused
  timeouts on files that render in seconds when run alone

### Configuration

External binaries are located via environment variables, with conventional
defaults — no code changes needed per machine:

| Variable | Default |
|---|---|
| `TESSERACT_EXE` | `C:\Program Files\Tesseract-OCR\tesseract.exe` |
| `POPPLER_BIN` | bundled `poppler/.../bin` next to the project |
| `MAGICK_EXE` | `magick` (from `PATH`) |
| `LLAMA_SERVER_EXE` | `C:\Program Files\llama.cpp\llama-server.exe` |
| `LLM_MODEL_PATH` | `C:\models\model-7b-instruct-q4.gguf` |

```bash
pip install -r requirements.txt
```

Tesseract needs the Ukrainian language data (`ukr`); a fine-tuned
`ukr_custom.traineddata` is picked up automatically if present.

---

## Repository layout

```
src/                    26 modules (see docs/INVENTORY.md for the full list)
src/legacy/             2 modules kept for completeness — they import an older
                        package layout and do not run; excluded from all entry points
sample_data/            synthetic generator + generated demo inputs
docs/INVENTORY.md       per-module description: purpose, I/O, key decisions
requirements.txt        generated from actual imports
```

---

## Note on data

This project was originally built for internal inventory and document
workflows. **All sample data in this repository is synthetic** — generated by
`sample_data/make_sample_data.py` with invented item names, the placeholder
unit code `A0000`, placeholder surnames, and invented quantities and prices.
No real records, scans, personal data or organisational identifiers are
included. Absolute paths have been replaced with environment variables.

Item names in the sample catalogue are generic supply nomenclature of the kind
found in published standards; they are not extracted from any record.
