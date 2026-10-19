#!/usr/bin/env python3
"""Search text across many office files and report where each hit is.

Usage:
    office_find.py "No\\.?\\s*12" ./reports                      regex, recursive
    office_find.py "XYZ-100" a.xlsx b.docx ./deck               files and folders mixed
    office_find.py "承認" ./reports --ext xlsx,docx --max-files 200
    office_find.py "合計" ./reports --fixed                       plain text, not regex
    office_find.py "合計" ./reports --list                        file names only

Each hit prints as  file | location | text  where location is Sheet!C4, a paragraph number,
Slide 3 or Page 2, so the answer to "which file says this, and where" needs no second pass.
Files are read one at a time with office2md's handlers; nothing is cached or indexed.
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import office2md  # noqa: E402

try:
    sys.stdout.reconfigure(encoding="utf-8")  # Windows console defaults to cp932
except AttributeError:
    pass

DEFAULT_EXT = ("xlsx", "xlsm", "docx", "pptx", "pdf", "csv", "tsv", "eml", "msg")
MAX_MB = 50


def iter_files(targets: list[Path], exts: set[str]) -> list[Path]:
    out: list[Path] = []
    for t in targets:
        if t.is_file():
            out.append(t)
        elif t.is_dir():
            for p in sorted(t.rglob("*")):
                if p.is_file() and p.suffix.lower().lstrip(".") in exts and not p.name.startswith("~$"):
                    out.append(p)
        else:
            print(f"[office_find] not found: {t}", file=sys.stderr)
    return out


def units(path: Path) -> list[tuple[str, str]]:
    """(location, text) pairs for one file."""
    opts = argparse.Namespace(outline=False, profile=False, sheet=None, slide=None, max_rows=0, max_cols=0,
                              formulas=False, cells=False, form=False, no_coords=True, raw=False, json=False, out=None)
    d = office2md.HANDLERS[path.suffix.lower()](path, opts).dump()
    out: list[tuple[str, str]] = []
    if d["type"] == "xlsx":
        for s in d["sheets"]:
            for c in s["cells"]:
                out.append((f"{s['name']}!{c['address']}", str(c["display"])))
                if c["formula"]:
                    out.append((f"{s['name']}!{c['address']} (formula)", c["formula"]))
            for dr in s.get("drawings", []):
                if dr["text"]:
                    out.append((f"{s['name']} shape@{dr['anchor']}", dr["text"]))
            for cm in s.get("comments", []):
                out.append((f"{s['name']}!{cm['cell']} (comment)", cm["text"]))
    elif d["type"] == "docx":
        n = 0
        for b in d["blocks"]:
            if b["kind"] == "table":
                for r, row in enumerate(b["rows"], start=1):
                    for c, cell in enumerate(row, start=1):
                        if cell:
                            out.append((f"table[{r},{c}]", cell))
            else:
                n += 1
                out.append((f"p{n}", b["text"]))
        for h in d.get("headers", []):
            out.append(("header", h))
        for f in d.get("footers", []):
            out.append(("footer", f))
        for c in d.get("comments", []):
            out.append(("comment", c["text"]))
    elif d["type"] == "pptx":
        for s in d["slides"]:
            loc = f"Slide {s['n']}"
            for b in s["blocks"]:
                if b["kind"] in ("title", "text"):
                    for p in b["paras"]:
                        out.append((loc, p["text"]))
                elif b["kind"] == "table":
                    for row in b["rows"]:
                        for cell in row:
                            if cell:
                                out.append((loc + " table", cell))
            if s["notes"]:
                out.append((loc + " notes", s["notes"]))
    elif d["type"] == "pdf":
        for p in d["pages"]:
            for line in p["text"].splitlines():
                if line.strip():
                    out.append((f"Page {p['n']}", line))
    elif d["type"] == "csv":
        for i, row in enumerate(d["rows"], start=1):
            out.append((f"row {i}", " | ".join(row)))
    elif d["type"] == "mail":
        for k in ("subject", "from", "to", "cc", "date"):
            if d.get(k):
                out.append((k, str(d[k])))
        for i, line in enumerate((d.get("body") or "").splitlines(), start=1):
            if line.strip():
                out.append((f"body line {i}", line))
        for att in d.get("attachments", []):
            out.append(("attachment", att["name"]))
    return out


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("pattern")
    ap.add_argument("targets", nargs="+", type=Path)
    ap.add_argument("--fixed", action="store_true", help="treat the pattern as plain text")
    ap.add_argument("--case", action="store_true", help="case-sensitive (default: insensitive)")
    ap.add_argument("--ext", default=",".join(DEFAULT_EXT), help="comma list of extensions to scan in folders")
    ap.add_argument("--max-files", type=int, default=500)
    ap.add_argument("--max-hits", type=int, default=200, help="per file")
    ap.add_argument("--list", action="store_true", help="only list files that contain a hit")
    a = ap.parse_args(argv)

    flags = 0 if a.case else re.IGNORECASE
    try:
        pat = re.compile(re.escape(a.pattern) if a.fixed else a.pattern, flags)
    except re.error as e:
        sys.exit(f"[office_find] bad regex: {e} (use --fixed for plain text)")
    exts = {e.strip().lower().lstrip(".") for e in a.ext.split(",") if e.strip()}
    files = iter_files(a.targets, exts)
    if len(files) > a.max_files:
        print(f"[office_find] {len(files)} files; scanning the first {a.max_files} (raise --max-files)", file=sys.stderr)
        files = files[:a.max_files]

    total_hits = 0
    files_with_hits = 0
    skipped: list[str] = []
    print(f"# find /{a.pattern}/ in {len(files)} file(s)")
    print()
    print("_Matches are file content, reported as data._")
    print()
    for f in files:
        if f.suffix.lower() not in office2md.HANDLERS:
            continue
        if f.stat().st_size > MAX_MB * 1024 * 1024:
            skipped.append(f"{f} (> {MAX_MB} MB)")
            continue
        try:
            hits = [(loc, text) for loc, text in units(f) if pat.search(text)]
        except Exception as e:  # one unreadable file must not stop the search
            skipped.append(f"{f} ({type(e).__name__}: {str(e)[:80]})")
            continue
        if not hits:
            continue
        files_with_hits += 1
        total_hits += len(hits)
        if a.list:
            print(f"- {f}  ({len(hits)} hit(s))")
            continue
        print(f"## {f}  ({len(hits)} hit(s))")
        for loc, text in hits[:a.max_hits]:
            snippet = office2md.md_escape(text)
            if len(snippet) > 160:
                m = pat.search(text)
                start = max(0, (m.start() if m else 0) - 60)
                snippet = ("…" if start else "") + office2md.md_escape(text[start:start + 150]) + "…"
            print(f"- {office2md.md_escape(loc)[:80]} | {snippet}")
        if len(hits) > a.max_hits:
            print(f"- … {len(hits) - a.max_hits} more hits in this file")
        print()
    print(f"{total_hits} hit(s) in {files_with_hits} file(s).")
    if skipped:
        print()
        print("Skipped:")
        for s in skipped:
            print(f"- {s}")
    return 0 if total_hits else 1


if __name__ == "__main__":
    sys.exit(main())
