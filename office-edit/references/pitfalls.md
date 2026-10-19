# What each library silently loses or gets wrong

Read the section for the format you are about to edit. Each item has caused a real
failure in use; the fix is usually one line if you know it beforehand.

## Excel (openpyxl)

**Formulas are never calculated.** Any formula you write, and any formula already in the file,
has an empty cached value after `wb.save()`. Excel recalculates on open, so a human sees nothing
wrong. But pandas, office-read, or any downstream script reads `None`. Run `recalc.py` after
saving if anything will read the values programmatically, and always run it to catch `#REF!`.

**Load-and-save drops these:**

- VBA macros unless `load_workbook(path, keep_vba=True)` and the file stays `.xlsm`.
- Images and shapes (openpyxl reads images only partially; drawings are lost).
- Charts: openpyxl keeps charts it created, but charts from Excel are dropped on load.
- Slicers, sparklines, pivot table caches (the pivot table may stay but goes stale), form controls,
  Power Query connections, data validation referencing external books.
- Conditional formatting and data validation usually survive, but check after saving.
- Comments survive; threaded comments (the newer kind) are converted to legacy notes.

When a workbook contains any of the above and the edit is small, prefer touching the XML
directly or editing via Excel COM (PowerShell `New-Object -ComObject Excel.Application`), both
of which keep everything else intact.

**Dates.** A cell shows "2026/10/17" but openpyxl gives `datetime`. Write back `datetime`
objects, not strings, or Excel will treat them as text and sorting/formulas break.

**Numbers stored as text.** Writing `"120"` (str) creates a text cell with a green triangle in
Excel. Write `120` (int). `xlsx_set.py` does this conversion for you.

**Styles are per cell, not per column.** Appending rows with `ws.append()` gives unstyled cells.
Copy `cell._style = copy(source_cell._style)` from the row above if the sheet is formatted.
`xlsx_set.py` does this automatically for unstyled target cells (and says so in its output);
pass `--no-copy-style` when the new cell is meant to look different.

**Merged cells.** Only the top-left cell holds the value. Writing to any other cell in the merge
raises or is ignored.

**Inserting rows** (`ws.insert_rows`) does NOT update formulas, merged ranges, defined names,
or conditional formatting that reference the shifted rows. If the sheet has formulas pointing
below the insert point, prefer appending at the bottom or doing the insert in Excel COM.

**read_only / write_only modes** are faster but lose everything except values. Never use them
for an edit-and-save round trip.

## Word (python-docx)

**Runs split words.** "東京支店" can be stored as "東京"+"支店" because of spell-check or an edit
history. Use `docx_replace.py` (run-spanning) rather than `run.text.replace`.

**Setting `run.text` discards tabs/breaks inside that run** (python-docx rebuilds the run's
children from the string, keeping only `\t` and `\n` as tab/break). Usually fine for
replacement; be aware if a run is a lone tab.

**`doc.paragraphs` skips tables; `doc.tables` skips nested tables; both skip text boxes,
headers, footers, footnotes, and comments.** Use the `iter_paragraphs()` pattern in
`docx_replace.py` when you need "all text".

**Tracked changes.** python-docx has no API for them. Text inside `<w:del>` is invisible to
`paragraph.text`; text inside `<w:ins>` is included. If a user wants edits *as* tracked changes,
you have to write `<w:ins>`/`<w:del>` XML by hand or ask them to turn on Track Changes and do
it in Word. Say which it is instead of silently making untracked edits in a tracked document.

**Styles must exist in the document.** `add_paragraph(style="List Bullet")` raises `KeyError`
if the template never used that style (common in Japanese templates: use "箇条書き" or check
`[s.name for s in doc.styles]`). Copy a paragraph's `pPr` from an existing one as a fallback.

**Adding a heading in a Japanese template** often gives Latin fonts. Set
`run.font.name` and also the East Asian font:
`run._element.rPr.rFonts.set(qn('w:eastAsia'), 'MS 明朝')`.

**Images:** `doc.add_picture()` only appends at the end. To insert in the middle, add a run to the
target paragraph: `paragraph.add_run().add_picture(path, width=Mm(80))`.

**Comments.** python-docx ≥ 1.2 has `doc.add_comment(runs, text, author)`. Older versions
need raw XML.

**Table cell merges.** `cell.merge(other)` works; `row.cells` then repeats the merged cell object,
so de-duplicate by `cell._tc` when iterating.

## PowerPoint (python-pptx)

**There is no "duplicate slide" API.** The usual workaround copies shapes' XML to a new slide
created from the same layout; pictures and charts need their relationships copied too, which
is fragile. For "make 5 slides like slide 3", prefer: ask the user for a template with the slide
already duplicated, or generate from a layout and fill placeholders.

**Deleting a slide:** also unofficial. `xml_slides = prs.slides._sldIdLst; xml_slides.remove(xml_slides[i])`
leaves the slide part orphaned in the zip (file still opens). Acceptable for one-offs.

**Reading order ≠ shape order.** `slide.shapes` is z-order. Sort by `(top, left)` if you need
to reason about layout.

**Placeholders differ by layout.** `slide.placeholders[1]` is the body on "Title and Content"
but may not exist on "Title Only". Check `[(p.placeholder_format.idx, p.name) for p in slide.placeholders]`.

**Charts** can be edited via `chart.replace_data(CategoryChartData(...))`; text replace does
not reach chart labels. SmartArt is not supported at all.

**Fonts and autofit.** python-pptx does not shrink text to fit. Long replacements overflow the
box. If a replacement is much longer than the original, check with a PDF render
(`soffice --headless --convert-to pdf`) or warn the user.

**Notes.** `slide.notes_slide` creates the notes slide on access if missing; harmless.

## PDF (pypdf)

Editing text in place is not realistic with pypdf. What works (all in `pdf_tool.py`): merge,
split, extract pages, rotate, bookmarks, Office → PDF. Also possible with a few lines: watermark
(overlay a page), fill AcroForm fields, metadata, encryption. For "change the wording in this
PDF", get the source document and re-export, or tell the user it is not feasible. For scanned
PDFs, OCR (tesseract + ocrmypdf) is the prerequisite for anything.

**Bookmarks (しおり) vs. a visible table of contents.** `pdf_tool.py bookmarks` writes the
navigation pane entries, not a printed TOC page. Both are often called 目次; ask if unsure.

**`PdfWriter.append()` keeps the source's bookmarks; `add_page()` drops them.** `merge` uses
append (so each input's own bookmarks survive, plus one per file); `bookmarks --set` uses
add_page on purpose so the old tree is replaced rather than duplicated.

**Office → PDF through COM** opens the real application. A file already open in Word/Excel
makes `Open` fail or prompt; close it first. Excel's `ExportAsFixedFormat` prints the active
print areas, so a sheet with no print area set exports everything, possibly across many pages.

## All formats

- Edit a copy, or let the scripts make a `.bak`. Office files cannot be partially written; a
  crash mid-save produces a zero-byte file.
- A file open in Excel/Word/PowerPoint is locked on Windows. Saving over it raises
  `PermissionError`. Ask the user to close it rather than saving under a new name silently.
- Verify after saving: run `office-read`'s `office2md.py` on the result and compare with
  the original dump, or `--outline` both and diff. A dropped sheet or table shows up there.
- Legacy `.xls` / `.doc` / `.ppt` cannot be written by these libraries. Convert first, or
  edit through COM.
