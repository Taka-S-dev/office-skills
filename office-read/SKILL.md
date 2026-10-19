---
name: office-read
description: Read and extract content from Excel, Word, PowerPoint, PDF and CSV files (.xlsx .xlsm .docx .pptx .pdf .csv .tsv) as Markdown or JSON using a bundled script, without any proprietary or licensed skill. Use this whenever the user mentions a spreadsheet, workbook, sheet, Word document, report, slide deck, presentation, PDF, or references a file with one of those extensions and wants it read, summarized, searched, compared, checked, or turned into something else. Also use when the user asks "what's in this file", "look at this Excel", "確認して", "読んで", "まとめて", "中身を見て" about an office file, even if they don't ask for extraction explicitly. Do not use for creating or editing office files.
---

# office-read

Office files are zipped XML; `cat` shows garbage and `Read` cannot open them. The bundled
script turns them into Markdown that fits in context, and it knows the traps that make a naive
dump misleading (hidden sheets, formulas vs cached values, speaker notes, merged cells,
tracked changes). Use the script instead of writing ad-hoc openpyxl / python-docx code each time,
so every read comes out in the same shape.

Script: `scripts/office2md.py` (relative to this skill's directory).
Dependencies: `pip install openpyxl python-docx python-pptx pypdf` (all MIT/BSD licensed).
Run it with plain `python`, not `python -I`, so user-installed packages are found. (The general
advice to use `-I` on downloaded files is about not importing code from the file's directory;
this script is the skill's own code and only reads the target as data, so plain `python` is fine.)

When formulas have no cached value (the script says so explicitly), the numbers are unknown, not
zero. Either compute them by hand from the formulas, or run office-edit's `scripts/recalc.py`
on a copy of the file to let Excel/LibreOffice fill them in.

## Workflow

0. **Check for a profile.** `profiles/` in this skill (and `.office-profiles/` in the
   project, if present) holds notes about recurring in-house files: which cells mean what, what
   must not be touched, the standard procedure for common requests. Skim the `# ` headings and
   `## 識別` lines; if one matches the file name or sheet/slide structure, read that profile
   before anything else and follow it. It encodes things nobody can see from the file itself.
   No match means no profile exists yet; proceed normally. If the user ends up explaining
   file-specific rules to you during the task, offer to save them:
   `office2md.py FILE --profile --out profiles/<name>.md` generates a skeleton with the structure
   filled in; add the rules they told you to the TODO lines. See `profiles/README.md`.

1. **Outline first, content second.** For anything bigger than a one-sheet workbook or a
   few slides, run `--outline` first. It is cheap and tells you what exists: sheet names and
   sizes, hidden sheets, heading tree, slide titles, which slides have notes.
   Then pull only what the question needs with `--sheet`, `--slide`, or `--max-rows`.

   ```bash
   python <skill>/scripts/office2md.py "path/to/file.xlsx" --outline
   python <skill>/scripts/office2md.py "path/to/file.xlsx" --sheet "売上" --max-rows 0
   python <skill>/scripts/office2md.py "path/to/deck.pptx" --slide 7
   ```

2. **Default dump is capped** at 200 rows and 50 columns per sheet. The output says when it
   truncated and prints the exact command to get the rest. Raise the cap only for the sheet
   you actually need; a 50,000-row sheet in context helps nobody.

3. **Large output goes to a file.** If a dump will clearly be long (many sheets, long report),
   pass `--out <scratch>/name.md` and then grep or read the file in parts, rather than
   flooding the terminal.

4. **Need to compute on the data, not read it?** Use `--json` and feed it to a short Python
   snippet, or load the workbook with openpyxl / pandas directly. The Markdown form is for
   reading; the JSON form is for processing.

5. **Reviewing or comparing, not just reading?** Excel output always carries row numbers and
   column letters, so cite findings as `メイン!C4 = 12`, never "the third row". For forms, sparse
   sheets, or cross-sheet comparison use `--cells`: one line per non-empty cell as
   `Sheet!C4 = value`, which is trivial to diff between sheets or files.

   ```bash
   python <skill>/scripts/office2md.py "path/to/file.xlsx" --sheet "メイン" --cells
   ```

6. **When the dump says `[uncached]` or prints a WARNING about formulas**, the workbook was
   saved by a script and never recalculated. Those cells have no value at all; treat any total
   or comparison that depends on them as unverified and say so in the report. If the value
   matters, run office-edit's `scripts/recalc.py` on a copy, or compute by hand from the
   formula shown.

7. **Text outside cells.** Text boxes and shapes are listed under "Shapes and pictures" with
   the cell they sit on; a reviewer's "修正済み" note is often there and nowhere else. Pictures
   cannot be read by any parser: when the outline reports images, or a review depends on what
   a figure shows, convert with office-edit's `scripts/pdf_tool.py topdf` and read the PDF.

## What the script surfaces and why it matters

| Format | Surfaced | Why a naive read misses it |
|---|---|---|
| xlsx | hidden / veryHidden sheets, marked `_(hidden)_` | Hidden sheets often hold the real inputs, or data the author did not mean to ship. Mention them to the user. |
| xlsx | row numbers and column letters on every grid; `--cells` for `Sheet!C4 = value` lines | A finding without an address cannot be verified in Excel. |
| xlsx | uncached formulas shown as `=B2*C2 [uncached]` plus a WARNING; `--formulas` shows `=B2*C2  → 1200` for cached ones | By default openpyxl returns None for a never-recalculated formula, which looks like an empty cell. The script always flags it. |
| xlsx | text in text boxes and shapes, with anchor cell; pictures listed | openpyxl ignores drawings entirely, so review notes written in a text box vanish from a plain dump. |
| xlsx | merged ranges, cell comments, defined names, tables | Merged header cells explain odd empty cells; comments are reviewer intent. |
| docx | heading tree in `--outline`, tables inline in document order | python-docx's `paragraphs` skips tables entirely, so a plain loop loses every table. |
| docx | header/footer text, comments, tracked-changes flag | "社外秘" lives in the header. If tracked changes exist, deletions are omitted from the dump; say so if the user is reviewing edits. |
| pptx | speaker notes, hidden slides, tables, chart series data, grouped shapes | Notes often contain the actual script; hidden slides are usually "do not present" leftovers. Chart values are embedded and recoverable. |
| pdf | page text, bookmarks, "no text layer" count | Scanned PDFs return nothing; the script tells you how many pages are empty so you can recommend OCR instead of reporting "the PDF is blank". |
| csv | encoding auto-detect (UTF-8, cp932, UTF-16), delimiter sniff | Japanese CSVs from Excel are cp932; reading them as UTF-8 raises or produces mojibake. |

## Reporting to the user

- Say which sheets / slides / pages you looked at, and name any you skipped due to size.
- Mention hidden sheets, hidden slides, tracked changes, and empty-text PDF pages when present,
  even if the user did not ask. These are the things people get burned by.
- Keep the user's cell addresses and sheet names when quoting numbers (e.g. `売上!D6`), so they
  can verify in Excel.
- Legacy `.xls` / `.doc` / `.ppt` are not supported. The script prints a LibreOffice conversion
  command; otherwise ask the user to re-save as the x-format.

## Extending

Each format is one `Handler` subclass in `scripts/office2md.py` with `outline()`, `dump()`,
`to_markdown()`. To add a format (e.g. `.odt`, `.msg`), add a subclass and register the
extension in `HANDLERS`. The CLI, truncation, JSON and `--out` handling need no change.
