#!/usr/bin/env python3
"""PDF page-level operations with pypdf: merge, split, extract, rotate, bookmarks, convert.

Text inside a PDF cannot be edited with this; fix the source document instead (see SKILL.md).

Usage:
    pdf_tool.py info a.pdf                               pages, metadata, bookmark tree
    pdf_tool.py merge out.pdf a.pdf b.pdf c.pdf          concatenate; each input becomes a top-level bookmark
    pdf_tool.py split a.pdf --every 1 --outdir parts/    one file per page (or per N pages)
    pdf_tool.py split a.pdf --at 4,9 --outdir parts/     split before pages 4 and 9 -> 1-3, 4-8, 9-end
    pdf_tool.py extract a.pdf --pages 2,5-7 --out sub.pdf
    pdf_tool.py rotate a.pdf --pages 3 --degrees 90 --out r.pdf
    pdf_tool.py bookmarks a.pdf --set toc.txt --out b.pdf   replace the bookmark tree (format below)
    pdf_tool.py bookmarks a.pdf --clear --out b.pdf
    pdf_tool.py topdf report.docx [--out report.pdf]      via Word/Excel/PowerPoint COM, else LibreOffice

Bookmark file format (toc.txt): one entry per line, indent with two spaces per level,
page number first (1-based), then the title.

    1 第1章 はじめに
      2 1.1 背景
      4 1.2 目的
    6 第2章 方法
"""

from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")  # Windows console defaults to cp932
except AttributeError:
    pass


def need_pypdf():
    try:
        import pypdf  # noqa: F401
        return pypdf
    except ImportError:
        sys.exit("[pdf_tool] missing library: pip install pypdf")


def parse_pages(spec: str, n_pages: int) -> list[int]:
    """'2,5-7' -> [1,4,5,6] (0-based). Validates range."""
    out: list[int] = []
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            a, b = int(a), (int(b) if b else n_pages)
            out.extend(range(a - 1, b))
        else:
            out.append(int(part) - 1)
    bad = [p + 1 for p in out if p < 0 or p >= n_pages]
    if bad:
        sys.exit(f"[pdf_tool] page(s) out of range (document has {n_pages}): {bad}")
    return out


# ---------------------------------------------------------------------------

def cmd_info(a):
    pypdf = need_pypdf()
    r = pypdf.PdfReader(a.file)
    if r.is_encrypted:
        if a.password is None or not r.decrypt(a.password):
            sys.exit("[pdf_tool] encrypted; pass --password")
    print(f"# {a.file.name}")
    print(f"pages: {len(r.pages)}")
    for k, v in (r.metadata or {}).items():
        print(f"{k.lstrip('/')}: {v}")
    print("bookmarks:")
    n = 0

    def walk(items, depth):
        nonlocal n
        for it in items:
            if isinstance(it, list):
                walk(it, depth + 1)
            else:
                try:
                    page = r.get_destination_page_number(it) + 1
                except Exception:
                    page = "?"
                print("  " * (depth + 1) + f"{page} {it.title}")
                n += 1
    try:
        walk(r.outline, 0)
    except Exception as e:
        print(f"  (could not read outline: {e})")
    if n == 0:
        print("  (none)")


def cmd_merge(a):
    pypdf = need_pypdf()
    w = pypdf.PdfWriter()
    for f in a.inputs:
        r = pypdf.PdfReader(f)
        start = len(w.pages)
        w.append(r)  # keeps each input's own bookmarks, nested under nothing
        if not a.no_bookmarks:
            w.add_outline_item(Path(f).stem, start)
        print(f"  + {f} ({len(r.pages)} pages)")
    w.write(a.out)
    print(f"[pdf_tool] wrote {a.out} ({len(w.pages)} pages)")


def cmd_split(a):
    pypdf = need_pypdf()
    r = pypdf.PdfReader(a.file)
    n = len(r.pages)
    a.outdir.mkdir(parents=True, exist_ok=True)
    if a.at:
        cuts = sorted({int(x) - 1 for x in a.at.split(",") if x.strip()})
        bounds = [0] + [c for c in cuts if 0 < c < n] + [n]
    else:
        every = a.every or 1
        bounds = list(range(0, n, every)) + [n]
    ranges = [(bounds[i], bounds[i + 1]) for i in range(len(bounds) - 1)]
    for i, (s, e) in enumerate(ranges, start=1):
        w = pypdf.PdfWriter()
        for p in range(s, e):
            w.add_page(r.pages[p])
        out = a.outdir / f"{a.file.stem}_{i:02d}_p{s + 1}-{e}.pdf"
        w.write(out)
        print(f"  {out.name}: pages {s + 1}-{e}")
    print(f"[pdf_tool] {len(ranges)} files in {a.outdir}")


def cmd_extract(a):
    pypdf = need_pypdf()
    r = pypdf.PdfReader(a.file)
    pages = parse_pages(a.pages, len(r.pages))
    w = pypdf.PdfWriter()
    for p in pages:
        w.add_page(r.pages[p])
    w.write(a.out)
    print(f"[pdf_tool] wrote {a.out} ({len(pages)} pages: {a.pages})")


def cmd_rotate(a):
    pypdf = need_pypdf()
    r = pypdf.PdfReader(a.file)
    targets = set(parse_pages(a.pages, len(r.pages))) if a.pages else set(range(len(r.pages)))
    w = pypdf.PdfWriter()
    for i, page in enumerate(r.pages):
        if i in targets:
            page.rotate(a.degrees)
        w.add_page(page)
    w.write(a.out)
    print(f"[pdf_tool] wrote {a.out} (rotated {len(targets)} pages by {a.degrees}°)")


def parse_toc(path: Path) -> list[tuple[int, int, str]]:
    """-> [(level, page0, title)]"""
    entries = []
    for ln in path.read_text(encoding="utf-8").splitlines():
        if not ln.strip():
            continue
        indent = len(ln) - len(ln.lstrip(" "))
        level = indent // 2
        body = ln.strip().split(None, 1)
        if len(body) != 2 or not body[0].isdigit():
            sys.exit(f"[pdf_tool] bad toc line (expected 'PAGE Title'): {ln!r}")
        entries.append((level, int(body[0]) - 1, body[1]))
    return entries


def cmd_bookmarks(a):
    pypdf = need_pypdf()
    r = pypdf.PdfReader(a.file)
    w = pypdf.PdfWriter()
    for page in r.pages:          # add_page instead of append() so old bookmarks are dropped
        w.add_page(page)
    if r.metadata:
        w.add_metadata({k: v for k, v in r.metadata.items() if isinstance(v, str)})
    if a.set:
        entries = parse_toc(a.set)
        n = len(w.pages)
        parents: list = [None] * 20
        for level, page0, title in entries:
            if page0 < 0 or page0 >= n:
                sys.exit(f"[pdf_tool] bookmark '{title}' points to page {page0 + 1}, document has {n}")
            parent = parents[level - 1] if level > 0 else None
            item = w.add_outline_item(title, page0, parent=parent)
            parents[level] = item
            print("  " * level + f"{page0 + 1} {title}")
        print(f"[pdf_tool] {len(entries)} bookmarks set")
    elif a.clear:
        print("[pdf_tool] bookmarks cleared")
    else:
        sys.exit("[pdf_tool] pass --set toc.txt or --clear")
    w.write(a.out)
    print(f"[pdf_tool] wrote {a.out}")


def find_soffice() -> str | None:
    exe = shutil.which("soffice") or shutil.which("soffice.exe")
    if exe:
        return exe
    for p in (r"C:\Program Files\LibreOffice\program\soffice.exe",
              "/Applications/LibreOffice.app/Contents/MacOS/soffice", "/usr/bin/soffice"):
        if Path(p).exists():
            return p
    return None


def cmd_topdf(a):
    src = a.file.resolve()
    out = (a.out or src.with_suffix(".pdf")).resolve()
    ext = src.suffix.lower()
    # Paths travel as environment variables, never as script text, so a file name cannot inject
    # PowerShell. AutomationSecurity=3 keeps document macros from running during the export.
    scripts = {
        ".docx": """$app = New-Object -ComObject Word.Application
try { $app.Visible = $false; $app.DisplayAlerts = 0; $app.AutomationSecurity = 3
  $doc = $app.Documents.Open($env:OFFICE_SRC, $false, $true)
  $doc.ExportAsFixedFormat($env:OFFICE_OUT, 17); $doc.Close(0) } finally { $app.Quit() }""",
        ".pptx": """$app = New-Object -ComObject PowerPoint.Application
try { $app.AutomationSecurity = 3
  $doc = $app.Presentations.Open($env:OFFICE_SRC, $true, $false, $false)
  $doc.SaveAs($env:OFFICE_OUT, 32); $doc.Close() } finally { $app.Quit() }""",
        ".xlsx": """$app = New-Object -ComObject Excel.Application
try { $app.Visible = $false; $app.DisplayAlerts = $false; $app.AutomationSecurity = 3
  $doc = $app.Workbooks.Open($env:OFFICE_SRC, 0, $true)
  if ($env:OFFICE_FIT -ne 'none') {
    # Page setup is changed in memory only; the workbook is read-only and closed without saving.
    foreach ($ws in $doc.Worksheets) {
      if ($ws.Visible -ne -1) { continue }
      try {
        $ps = $ws.PageSetup
        $ps.PrintArea = ''
        $ps.Zoom = $false
        $ps.FitToPagesWide = 1
        if ($env:OFFICE_FIT -eq 'page') { $ps.FitToPagesTall = 1 } else { $ps.FitToPagesTall = $false }
      } catch { }
    }
  }
  $doc.ExportAsFixedFormat(0, $env:OFFICE_OUT); $doc.Close($false) } finally { $app.Quit() }""",
    }
    scripts[".doc"] = scripts[".docm"] = scripts[".docx"]
    scripts[".ppt"] = scripts[".pptm"] = scripts[".pptx"]
    scripts[".xls"] = scripts[".xlsm"] = scripts[".xlsx"]
    if os.name == "nt" and ext in scripts and a.engine != "soffice":
        ps = "$ErrorActionPreference = 'Stop'\n" + scripts[ext]
        r = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", ps],
                           capture_output=True, text=True,
                           env={**os.environ, "OFFICE_SRC": str(src), "OFFICE_OUT": str(out),
                                "OFFICE_FIT": a.fit})
        if r.returncode == 0 and out.exists():
            print(f"[pdf_tool] wrote {out} (via Office)")
            return
        err = r.stderr.strip().splitlines()
        print(f"[pdf_tool] Office COM failed: {err[0] if err else 'unknown'}; trying LibreOffice")
    so = find_soffice()
    if not so:
        sys.exit("[pdf_tool] no converter: Office COM unavailable and LibreOffice not found.")
    r = subprocess.run([so, "--headless", "--convert-to", "pdf", "--outdir", str(out.parent), str(src)],
                       capture_output=True, text=True, timeout=300)
    produced = out.parent / (src.stem + ".pdf")
    if r.returncode != 0 or not produced.exists():
        sys.exit(f"[pdf_tool] soffice failed:\n{r.stdout}\n{r.stderr}")
    if produced != out:
        shutil.move(produced, out)
    print(f"[pdf_tool] wrote {out} (via LibreOffice)")


# ---------------------------------------------------------------------------

def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("info"); p.add_argument("file", type=Path); p.add_argument("--password")
    p.set_defaults(fn=cmd_info)

    p = sub.add_parser("merge"); p.add_argument("out", type=Path); p.add_argument("inputs", nargs="+")
    p.add_argument("--no-bookmarks", action="store_true", help="don't add a bookmark per input file")
    p.set_defaults(fn=cmd_merge)

    p = sub.add_parser("split"); p.add_argument("file", type=Path)
    g = p.add_mutually_exclusive_group(); g.add_argument("--every", type=int); g.add_argument("--at")
    p.add_argument("--outdir", type=Path, default=Path("."))
    p.set_defaults(fn=cmd_split)

    p = sub.add_parser("extract"); p.add_argument("file", type=Path)
    p.add_argument("--pages", required=True); p.add_argument("--out", type=Path, required=True)
    p.set_defaults(fn=cmd_extract)

    p = sub.add_parser("rotate"); p.add_argument("file", type=Path)
    p.add_argument("--pages"); p.add_argument("--degrees", type=int, default=90, choices=[90, 180, 270, -90])
    p.add_argument("--out", type=Path, required=True)
    p.set_defaults(fn=cmd_rotate)

    p = sub.add_parser("bookmarks"); p.add_argument("file", type=Path)
    p.add_argument("--set", type=Path, help="toc text file"); p.add_argument("--clear", action="store_true")
    p.add_argument("--out", type=Path, required=True)
    p.set_defaults(fn=cmd_bookmarks)

    p = sub.add_parser("topdf"); p.add_argument("file", type=Path); p.add_argument("--out", type=Path)
    p.add_argument("--engine", choices=["office", "soffice"])
    p.add_argument("--fit", choices=["width", "page", "none"], default="width",
                   help="xlsx via Office: width = each sheet one page wide, as many tall as needed (default, "
                        "keeps rows intact for review); page = whole sheet on one page; none = author's "
                        "print settings. Ignored by LibreOffice.")
    p.set_defaults(fn=cmd_topdf)

    a = ap.parse_args(argv)
    if hasattr(a, "file") and not a.file.exists():
        sys.exit(f"[pdf_tool] file not found: {a.file}")
    a.fn(a)
    return 0


if __name__ == "__main__":
    sys.exit(main())
