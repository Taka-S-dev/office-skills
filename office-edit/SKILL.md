---
name: office-edit
description: Edit, update, fill in, or create Excel, Word and PowerPoint files (.xlsx .xlsm .docx .pptx) safely with openpyxl / python-docx / python-pptx, and do PDF page operations (merge, split, extract, rotate, bookmarks/しおり, convert Office files to PDF) with pypdf, without any proprietary or licensed skill. Use this whenever the user wants an office file changed in any way: replace text or names across a document or deck, update cells or formulas, add rows or a sheet, fill a template, fix a report, bump a fiscal year, rename a company, insert a table or picture, append slides, or produce a new .xlsx/.docx/.pptx deliverable. Also use for Japanese phrasing like "書き換えて", "更新して", "修正して", "差し替えて", "追記して", "作って" when an office file is the target. Covers recalculating formulas and checking for #REF! errors after an edit. For reading only, use office-read instead.
---

# office-edit

Editing an office file is easy to get mostly right and easy to break in the remaining cases:
formulas that show as blank to the next script, macros that vanish, a company name that was
split across two runs and never got replaced, a tracked-changes document edited untracked.
This skill is the checklist plus four scripts that handle the recurring cases.

Scripts live in `scripts/` relative to this skill. Dependencies:
`pip install openpyxl python-docx python-pptx pypdf` (MIT/BSD licensed). Run with plain `python`,
not `python -I`: `-I` hides user-site packages, and these scripts only read the target file as
data, so the usual caution about running interpreters on downloaded files does not apply.

Replace scripts report hits per location (body / tables / header / footer, or per slide). When
the user says 「本文の」 and the dry run shows header or footer hits, pass `--scope body,tables`
rather than guessing.

## The procedure

0. **Check for a profile.** office-read's `profiles/` (and the project's `.office-profiles/`)
   may hold a note for this kind of file: protected cells, the standard procedure for the request,
   formatting rules for new rows. If a profile matches, its rules override your defaults.
   When the user teaches you a rule mid-task ("D10 は触らないで"), offer to record it there
   (`office2md.py FILE --profile --out profiles/<name>.md`, then fill the TODO lines) so the next
   session does not have to be told again.

1. **Read before writing.** Dump the file with office-read's `office2md.py --outline` (and the
   relevant sheet/slide) so you know what is there: hidden sheets, formulas, headers, notes.
   For Excel, the `--outline` output ends with an `Objects:` line that lists charts, images,
   macros, pivots and the like found in the file, because openpyxl drops some of those on save
   (see `references/pitfalls.md`, Excel section). If it says anything other than "none" and the
   edit is small, consider editing through Excel COM instead, or verify those objects survived
   after saving.

2. **Never edit the only copy.** Work on a copy or let the scripts write a `.bak`. If the user
   has the file open in Office, saving fails with `PermissionError`; ask them to close it.

3. **Prefer the bundled scripts for the common edits.** They handle the traps and print what
   they changed, which doubles as the report to the user.

   | Task | Command |
   |---|---|
   | Replace text in Word (body, tables, header/footer, run-spanning) | `python <skill>/scripts/docx_replace.py file.docx --replace "旧=新" --dry-run` then without `--dry-run` |
   | Replace text in PowerPoint (slides, tables, groups, notes) | `python <skill>/scripts/pptx_replace.py deck.pptx --replace "旧=新" [--slides 2,5-7]` |
   | Set cells or formulas in Excel | `python <skill>/scripts/xlsx_set.py book.xlsx --set "売上!D6==SUM(D2:D5)"` |
   | Recalculate and check formula errors | `python <skill>/scripts/recalc.py book.xlsx` |
   | PDF: merge / split / extract pages / rotate | `python <skill>/scripts/pdf_tool.py merge out.pdf a.pdf b.pdf`, `split a.pdf --at 4,9 --outdir parts/`, `extract a.pdf --pages 2,5-7 --out x.pdf` |
   | PDF: set bookmarks (しおり) from a text outline | `python <skill>/scripts/pdf_tool.py bookmarks a.pdf --set toc.txt --out b.pdf` (format in the script's help) |
   | PDF: build that outline from a long document | `python <skill>/scripts/pdf_headings.py a.pdf --out toc.txt` (auto: printed TOC, else built-in chapter/section patterns); `--sample` to see heading-like lines per page, then `--pattern REGEX` to apply the style you recognised to every page |

   **Bookmarks for a long PDF: never build the list by reading the pages.** A 600-page PDF does
   not fit in one reading, so a hand-built list stops partway and the job needs several rounds.
   Instead: run `pdf_headings.py a.pdf --out toc.txt` first. If its summary reports page ranges
   with no heading, or the list looks wrong, run `--sample` (one line per heading-like line,
   with page and font size, for the whole document), decide what this document's headings look
   like, and rerun with `--pattern` (a regex; named group `num` such as `1.2.3` sets nesting,
   `--level N` forces one). Matching ignores the spaces extraction scatters inside Japanese
   headings. Then run `pdf_headings.py a.pdf --check toc.txt`: it lists numbering gaps (第2章
   followed by 第4章, 3.1 then 3.3), heading-like lines that are not in the list, and long
   stretches with no entry. Decide which of those are real headings, add a `--pattern` for
   their style, rerun with every pattern so far (they combine), and check again. Stop when the
   check is empty or only lists things you decided to leave out, such as the title page or a
   level below `--max-level`. Only then pass toc.txt to `pdf_tool.py bookmarks --set`.
   | Word / Excel / PowerPoint → PDF | `python <skill>/scripts/pdf_tool.py topdf report.docx` (Office COM on Windows, else LibreOffice) |
   | Excel → PDF for visual review | default `--fit width`: each sheet one page wide, rows never split across pages; `--fit page` for small sheets; `--fit none` to keep the author's print setup when the PDF itself is the deliverable |

   PDF text itself cannot be edited here. When the user wants wording changed in a PDF, edit
   the source .docx/.pptx and run `topdf` again; if there is no source, say so.

   Always run the replace scripts with `--dry-run` first. A pattern with 0 hits means the text
   is split differently than you think, lives in a text box, or does not exist; find out before
   writing.

4. **For anything the scripts do not cover, write a short openpyxl / python-docx / python-pptx
   script** in the scratchpad. Before writing it, read the matching section of
   `references/pitfalls.md`. It is short and lists exactly the things that go wrong
   (`keep_vba`, dates as strings, missing styles in Japanese templates, no duplicate-slide API,
   insert_rows not updating formulas).

5. **Verify after saving.** Run office-read's `scripts/office_diff.py <backup> <result>`: it lists
   every changed cell, paragraph or slide, so unintended changes (a dropped text box, a shifted
   formula) show up next to the intended ones. Then check:
   - the sheets / headings / slide count match the original plus your intended change,
   - for Excel, `recalc.py` reports no `#REF!` / `#DIV/0!` / `#NAME?`,
   - nothing you did not touch disappeared (charts, images, macros).
   Then tell the user what changed, cell by cell or pattern by pattern, and what you verified.

## Creating new files

- **Excel:** openpyxl from scratch is fine. Write numbers as numbers, dates as `datetime`,
  formulas as strings starting with `=`. Set column widths and a bold header row; Excel users
  notice when a sheet is unformatted. Run `recalc.py` at the end if the sheet has formulas.
- **Word:** start from the user's template (`Document("template.docx")`) whenever one exists,
  so styles, fonts and headers match the company look. From a blank `Document()` the fonts are
  Calibri and Japanese text looks wrong; set East Asian fonts explicitly (pitfalls.md).
- **PowerPoint:** start from the user's template or an existing deck and use its layouts via
  `prs.slide_layouts`. Fill placeholders rather than drawing text boxes at coordinates.
  Building a polished deck from zero in python-pptx is slow; say so and keep the scope to
  "slides with title + bullets + one table/picture each" unless the user asks for more.

## What to say to the user

- What changed, with locations (`売上!D6`, "slide 3 title", "header").
- Any 0-hit patterns and your best guess why.
- Any feature the save may have dropped (macros, charts, pivot caches) and whether you checked.
- If the document has tracked changes and they asked for edits: whether your edits are
  tracked (they are not, unless you wrote the XML) and what they can do about it.
