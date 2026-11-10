# Local OCR for scanned PDFs: three engines, one vote, keyword search

Reads image-only PDFs (scans, phone photos of paper) **entirely on the local
machine** and finds the ones that contain a given phrase — across hundreds
of files, even when OCR mangles the phrase.

Nothing is uploaded anywhere: the OCR engines run locally, the optional LLM
runs in [Ollama](https://ollama.com) on `127.0.0.1`, and the Chrome instance
used for OCR has its own isolated profile.

```
              ┌── Chrome PDF viewer (Screen AI), on a hidden desktop ──┐
PDF ──render──┼── Windows Photos "Scan text", window off-screen ───────┼──► align ──► vote ──► text ──► keyword filter
              └── Tesseract (ukr), psm3 + layout-aware psm6 ───────────┘   lines     2 of 3      │
                                                                                      │        └► LLM only for lines
                                                                                      │           with no majority
```

## What it does

* **Three independent OCR witnesses, in parallel** (one thread each):
  * **Chrome** — the built-in PDF viewer recognises scans with Google's
    Screen AI library. Text is read straight from the viewer's internal
    controller over the DevTools protocol — no clipboard, no screenshots.
    Chrome runs on a separate, invisible Windows desktop: no windows, no
    taskbar entries, no focus stealing.
  * **Windows Photos** — the "Scan text" feature of the stock Photos app,
    driven purely through UI Automation (no mouse, no keyboard), with its
    window moved off-screen.
  * **Tesseract** with Ukrainian data — plain `psm 3` plus a layout-aware
    `psm 6` pass that rebuilds lines from word coordinates.
* **A quorum instead of trusting one engine.** Texts are aligned line by
  line; a line two of three engines agree on is taken as is. Where they
  disagree, the default is a local **qwen2.5 7B** that may only *choose and
  combine words the engines actually produced* — any answer containing a
  word no engine saw is rejected. Without the LLM (or for rejected answers)
  the line is resolved by word-level majority (which also restores
  punctuation one engine dropped), then by aligned character voting checked
  against a Ukrainian dictionary (`pymorphy3`).
* **Keyword search over many PDFs** (`--phrase`). Only documents containing
  the phrase are saved. Matching is fuzzy, so OCR noise does not hide a hit;
  files where even fuzzy search fails get a second pass by the local LLM,
  which must quote the matching fragment *verbatim* — the quote is accepted
  only if it really is in the text and resembles the phrase.

## Accuracy

Measured on 9 synthetic pages with known ground truth, three degradation
levels (tilt, uneven light, noise; plus blur, downscaling and JPEG artefacts
for the harder ones). Character error rate / share of fully correct lines:

| | easy | medium | hard | **all** |
|---|---|---|---|---|
| Chrome alone (best single engine) | 0.28% / 92.0% | 1.07% / 79.5% | 4.05% / 34.7% | 1.80% / 68.6% |
| Windows Photos alone | 0.72% / 76.0% | 0.90% / 69.9% | 6.15% / 25.3% | 2.60% / 57.0% |
| Tesseract alone | 1.32% / 80.0% | 27.15% / 13.7% | 31.32% / 17.3% | 19.84% / 37.2% |
| quorum, no LLM | 0.04% / 98.7% | 0.41% / 90.4% | 3.56% / 41.3% | 1.34% / 76.7% |
| quorum + qwen2.5 1.5B | 0.04% / 98.7% | 0.70% / 84.9% | 3.48% / 41.3% | 1.41% / 74.9% |
| **quorum + qwen2.5 7B** (default) | **0.04% / 98.7%** | **0.25% / 93.2%** | **2.95% / 46.7%** | **1.08% / 79.4%** |

What the measurement changed:

* The first version of the vote was *worse* than Chrome alone on hard pages.
  Fixes that came out of it: re-inserting tokens the base engine dropped,
  treating an unaligned engine as abstaining rather than voting "no such
  word", and — the one that matters most for lists — penalising skipped
  lines during alignment, because item names repeat with different
  quantities and "most similar line" would otherwise pair `item - 4` with a
  later `item - 2`.
* The 1.5B model is not used automatically: it scored worse than no LLM.
* Known limit: Photos and Tesseract read the *same* rendered image and
  sometimes make the *same* mistake, outvoting a correct Chrome.

The 7B model costs about 7 seconds per disputed line on a laptop CPU;
clean pages have none.

## Quick start

```bash
pip install -r requirements.txt
python sample_data/make_sample_data.py          # 9 synthetic scanned PDFs + ground truth
python src/pdf_to_text_multithread.py sample_data/pdf --quorum --phrase "акт приймання-передачі"
```

Matching documents land in `src/output_txt/` as `<name>_text.txt`; a
per-file report with every engine's text and every decision is written to
`src/Data/quorum/`.

### Options

| flag | meaning |
|---|---|
| `--phrase TEXT` | keep only PDFs containing this phrase (fuzzy); without it every PDF is saved |
| `--quorum` | three engines + vote (recommended); without it: Tesseract, optionally with `--chrome` first |
| `--no-photos` / `--no-chrome` | drop a witness |
| `--no-llm` | resolve disagreements by voting only |
| `--llm-model NAME` | Ollama model (default: qwen2.5 7B if installed) |
| `--no-llm-recheck` | skip the LLM pass for files where fuzzy search found nothing |
| `--fuzzy-threshold N` | phrase match threshold, 0–1 (default 0.75) |
| `--preprocess` | dewarp, deskew, flat-field, CLAHE before Tesseract — for photos / crooked scans only |
| `--chrome`, `--chrome-visible`, `--chrome-workers N` | Chrome as a stage before Tesseract (without the quorum) |

## Requirements

* Python 3.10+, packages from `requirements.txt`.
* **Tesseract 5** with Ukrainian data (`ukr`, better `ukr_best`). If several
  installs exist, the one with Ukrainian data is picked automatically.
* **Google Chrome ≥ 153** (not headless — headless has no PDF viewer). Its
  OCR component (Screen AI) downloads into a fresh profile only after a
  while; if the regular Chrome profile already has it, it is copied over.
* **Windows 10/11** for the Photos witness and the hidden desktop. Elsewhere
  the quorum runs with Chrome (off-screen window) and Tesseract.
* Optional: **Ollama** with `qwen2.5:7b-instruct`.
* PDF rendering uses Poppler if installed, otherwise PyMuPDF.

## Caveats

* Photos needs the *visible* desktop: its window appears for ~40 ms before
  being moved off-screen and takes keyboard focus. On a hidden desktop it
  opens but never enables the text layer. Use `--no-photos` if that matters.
* Photos works through the system clipboard; it is shared with Chrome's
  fallback path under one lock and the user's text clipboard is restored
  afterwards (non-text clipboard content is not).
* The accuracy numbers come from 9 synthetic pages in one font — enough to
  compare variants, not a general claim.

## Layout

```
src/pdf_to_text_multithread.py   entry point: rendering, Tesseract, keyword filter, stages
src/ocr_quorum.py                alignment, voting, LLM client, three-witness orchestration
src/chrome_pdf_text.py           Chrome PDF viewer over DevTools, hidden desktop
src/рука.py                      human-paced CDP input (Chrome fallback path)
src/photos_ocr.py                Windows Photos "Scan text" via UI Automation
src/image_preprocessor.py        optional page clean-up before Tesseract
sample_data/make_sample_data.py  synthetic scanned PDFs + ground truth
```

---

### Коротко українською

Локальний OCR для сканованих PDF: текст розпізнають одночасно три рушії —
переглядач PDF у Chrome (на невидимому робочому столі), «Сканування тексту»
у застосунку «Фото» Windows і Tesseract. Рядки, де зійшлися двоє з трьох,
беруться як є; спірні вирішує локальна модель qwen2.5 7B (Ollama), якій
заборонено вигадувати слова, яких не бачив жоден рушій. З `--phrase`
програма проходить сотні PDF і зберігає лише ті, де є ключова фраза —
навіть спотворена OCR. Усе працює без інтернету й без передачі даних назовні.
