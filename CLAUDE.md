# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A small, single-purpose utility (`translate.py`) that translates the text
content of an `.xlsx` workbook while preserving images, styles, formulas, and
rich-text run formatting. It is built around a **three-step manual workflow**
that uses an LLM as the translation engine and the filesystem as the channel
between steps. There is no build/test/lint tooling — it's a script you run
manually with `python translate.py`.

## The translation pipeline (read this first)

The flow is intentionally split so that an external LLM does the translation
between the extract and replace steps. Order matters, and the file format
contract between steps is implicit and brittle (see "Data contract" below).

1. **Extract** — `translate_excel(path)` reads the source workbook with
   openpyxl and returns the *set* of unique non-empty string cell values.
   `print_translate(path)` writes them (sorted, one per record) to `1.txt`.
   This is currently commented out in `__main__` — uncomment to regenerate.
2. **Translate (external)** — an LLM translates `1.txt` → `2.txt`. The two
   files must stay record-aligned: the Nth `======`-delimited block in `1.txt`
   must correspond to the Nth block in `2.txt`. The Chinese source is the
   key; the English block is the value.
3. **Replace** — `excel_cell_replace("./1.txt", "./2.txt", path)`:
   - `loadData()` parses both files into position-indexed dicts keyed by an
     incrementing counter (NOT by content).
   - `get_translate_map()` zips them positionally:
     `combined_map[chinese_text] = english_text`. If a position exists in
     `1.txt` but not `2.txt`, the Chinese original is used unchanged.
   - openpyxl is then used to rewrite each matching cell's value.

## Data contract (the load-bearing constraint)

Positional alignment between `1.txt` and `2.txt` is the entire correctness
model. There is no fuzzy matching and no key — if blocks drift by one
(because you re-extracted, or the LLM merged/split blocks), every record
after the drift point maps to the wrong translation. When regenerating `1.txt`,
you MUST regenerate `2.txt` from the same extraction.

Records are delimited by `======` on its own line; blank/whitespace-only
segments are skipped during load.

## openpyxl limitations that shape the design

- **Images drop on save.** openpyxl does not support `wmf`; calling
  `wb.save()` on a workbook that contains wmf images silently drops them
  (`UserWarning: wmf image format is not supported so the image is being
  dropped`). This script accepts that loss — text inside dropped images
  cannot be translated by this tool.
- **Rich text is preserved** as long as you only set `cell.value` to a plain
  string per cell; openpyxl keeps the existing cell's formatting.
- Match keys are `cell.value.strip()` results, so leading/trailing
  whitespace differences between the extracted key and the live cell value
  will cause a miss (the cell keeps its original Chinese).

## Files

- `translate.py` — the whole tool. `__main__` toggles between extract and
  replace by commenting lines in/out.
- `1.txt` / `2.txt` — source-language and target-language record lists
  (gitignored).
- `excel/TRANSLATE.xlsx` — the source workbook (~150 MB; gitignored).
- `excel/TRANSLATE copy.xlsx` — backup copy.
- `*-translated.xlsx` — output, written next to the script (gitignored).

## Working in this repo

- If you change the extraction logic (e.g. sort order, dedup, filtering),
  expect `2.txt` to be stale — flag that the translation file needs
  regenerating rather than silently running replace against a mismatched map.
- The script is meant to be run with the working directory at the repo
  root (paths are relative: `./1.txt`, `./2.txt`, `./excel/translate.xlsx`).
