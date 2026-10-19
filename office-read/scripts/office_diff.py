#!/usr/bin/env python3
"""Show what changed between two versions of an office file, by cell, paragraph, slide or page.

Usage:
    office_diff.py old.xlsx new.xlsx [--sheet NAME] [--formulas]
    office_diff.py old.docx new.docx
    office_diff.py old.pptx new.pptx
    office_diff.py old.pdf  new.pdf
    ... [--context 2] [--out diff.md]

Comparing two full dumps by eye costs a lot of reading and misses small edits. This prints only
the differences, each with its address (Sheet!C4, paragraph number, slide number, page), so a
review can say exactly what moved. Exit code 0 = identical, 1 = differences found.
"""

from __future__ import annotations

import argparse
import difflib
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import office2md  # noqa: E402

try:
    sys.stdout.reconfigure(encoding="utf-8")  # Windows console defaults to cp932
except AttributeError:
    pass


def dump(path: Path, a) -> dict:
    ext = path.suffix.lower()
    cls = office2md.HANDLERS.get(ext)
    if cls is None:
        sys.exit(f"[office_diff] unsupported extension {ext}")
    opts = argparse.Namespace(outline=False, profile=False, sheet=a.sheet, slide=None, max_rows=50000, max_cols=1000,
                              formulas=a.formulas, cells=False, form=False, no_coords=True, raw=True,
                              json=False, out=None)   # raw: compare stored values, not display strings
    return cls(path, opts).dump()


MAX_LINES = 20000     # lines per unit compared; more than this is truncated with a note


def fmt(v) -> str:
    """Document text for the report. Pipes would break tables and a backtick run could close
    the diff fence, letting file content pose as report structure; both are neutralised."""
    if v is None:
        return "(empty)"
    s = str(v).replace("\r\n", "\n").replace("\n", " / ")
    return s.replace("|", "\\|").replace("`", "ˋ")


def bounded(lines: list[str], what: str) -> list[str]:
    if len(lines) <= MAX_LINES:
        return lines
    return lines[:MAX_LINES] + [f"(… {len(lines) - MAX_LINES} more lines of {what} not compared; "
                                "narrow with --sheet or split the file)"]


# -- xlsx ---------------------------------------------------------------------------------

def diff_xlsx(old: dict, new: dict, a) -> tuple[list[str], int]:
    out: list[str] = []
    n_changes = 0
    o_sheets = {s["name"]: s for s in old["sheets"]}
    n_sheets = {s["name"]: s for s in new["sheets"]}
    for name in o_sheets.keys() - n_sheets.keys():
        out.append(f"## Sheet removed: {fmt(name)}")
        n_changes += 1
    for name in n_sheets.keys() - o_sheets.keys():
        out.append(f"## Sheet added: {fmt(name)} ({n_sheets[name]['max_row']} rows)")
        n_changes += 1

    for name in [s["name"] for s in new["sheets"] if s["name"] in o_sheets]:
        os_, ns_ = o_sheets[name], n_sheets[name]

        def cellmap(s):
            return {c["address"]: (c["formula"], c["value"]) for c in s["cells"]}
        om, nm = cellmap(os_), cellmap(ns_)
        rows = []
        for addr in sorted(om.keys() | nm.keys(), key=lambda x: (int(''.join(ch for ch in x if ch.isdigit())), x)):
            o = om.get(addr)
            n = nm.get(addr)
            if o == n:
                continue
            if o is None:
                rows.append([addr, "added", "", fmt(n[0] or n[1])])
            elif n is None:
                rows.append([addr, "removed", fmt(o[0] or o[1]), ""])
            else:
                kind = "formula" if o[0] != n[0] else "value"
                rows.append([addr, kind, fmt(o[0] or o[1]), fmt(n[0] or n[1])])
        drows = []
        od = {(d["anchor"], d["name"]): d["text"] for d in os_.get("drawings", [])}
        nd = {(d["anchor"], d["name"]): d["text"] for d in ns_.get("drawings", [])}
        for k in sorted(od.keys() | nd.keys()):
            if od.get(k) != nd.get(k):
                drows.append([f"{k[0]} {k[1]}", fmt(od.get(k)), fmt(nd.get(k))])
        if os_["state"] != ns_["state"]:
            out.append(f"## Sheet: {fmt(name)}  (visibility {os_['state']} → {ns_['state']})")
        elif rows or drows:
            out.append(f"## Sheet: {fmt(name)}")
        if rows:
            out.append(office2md.md_table(rows, header=["cell", "change", "old", "new"]))
            n_changes += len(rows)
        if drows:
            out.append("")
            out.append("Shapes / text boxes:")
            out.append(office2md.md_table(drows, header=["shape", "old", "new"]))
            n_changes += len(drows)
        if rows or drows or os_["state"] != ns_["state"]:
            out.append("")
    return out, n_changes


# -- text-based formats -------------------------------------------------------------------

def docx_lines(d: dict) -> list[str]:
    lines = []
    t = 0
    for b in d["blocks"]:
        if b["kind"] == "table":
            t += 1
            for r, row in enumerate(b["rows"], start=1):
                for c, cell in enumerate(row, start=1):
                    if cell:
                        lines.append(f"table{t}[{r},{c}]: {fmt(cell)}")
        elif b["kind"] == "heading":
            lines.append("#" * b["level"] + " " + fmt(b["text"]))
        elif b["kind"] == "list":
            lines.append("- " + fmt(b["text"]))
        else:
            lines.append(fmt(b["text"]))
    for h in d.get("headers", []):
        lines.append(f"header: {fmt(h)}")
    for f in d.get("footers", []):
        lines.append(f"footer: {fmt(f)}")
    return lines


def pptx_units(d: dict) -> dict[str, list[str]]:
    units = {}
    for s in d["slides"]:
        lines = []
        for b in s["blocks"]:
            if b["kind"] in ("title", "text"):
                for p in b["paras"]:
                    lines.append(("[title] " if b["kind"] == "title" else "") + fmt(p["text"]))
            elif b["kind"] == "table":
                for r, row in enumerate(b["rows"], start=1):
                    for c, cell in enumerate(row, start=1):
                        if cell:
                            lines.append(f"table[{r},{c}]: {fmt(cell)}")
            elif b["kind"] == "chart":
                for se in b["series"]:
                    lines.append(f"chart {se['name']}: " + ", ".join(f"{c}={v}" for c, v in zip(se["categories"], se["values"])))
        if s["notes"]:
            lines.append("notes: " + fmt(s["notes"]))
        units[f"Slide {s['n']}"] = lines
    return units


def pdf_units(d: dict) -> dict[str, list[str]]:
    return {f"Page {p['n']}": [fmt(l) for l in p["text"].splitlines() if l.strip()] for p in d["pages"]}


def diff_lines(old: list[str], new: list[str], context: int, label_prefix: str = "") -> list[str]:
    out = []
    old, new = bounded(old, "old"), bounded(new, "new")
    sm = difflib.SequenceMatcher(a=old, b=new)
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            continue
        # context before
        for k in range(max(0, i1 - context), i1):
            out.append(f"  {label_prefix}{k + 1}: {old[k]}")
        if tag in ("replace", "delete"):
            for k in range(i1, i2):
                out.append(f"- {label_prefix}{k + 1}: {old[k]}")
        if tag in ("replace", "insert"):
            for k in range(j1, j2):
                out.append(f"+ {label_prefix}{k + 1}: {new[k]}")
        for k in range(i2, min(len(old), i2 + context)):
            out.append(f"  {label_prefix}{k + 1}: {old[k]}")
        out.append("")
    return out


def diff_units(old: dict[str, list[str]], new: dict[str, list[str]], context: int) -> tuple[list[str], int]:
    out: list[str] = []
    n = 0
    for name in old.keys() - new.keys():
        out.append(f"## {name} removed ({len(old[name])} lines)")
        n += 1
    for name in new.keys() - old.keys():
        out.append(f"## {name} added")
        out += [f"+ {l}" for l in new[name]]
        out.append("")
        n += 1
    for name in [k for k in new if k in old]:
        d = diff_lines(old[name], new[name], context)
        if d:
            out.append(f"## {name}")
            out.append("```diff")
            out += d
            out.append("```")
            out.append("")
            n += sum(1 for l in d if l.startswith(("-", "+")))
    return out, n


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("old", type=Path)
    ap.add_argument("new", type=Path)
    ap.add_argument("--sheet", help="xlsx: only this sheet")
    ap.add_argument("--formulas", action="store_true", help="xlsx: compare formulas as well as values (default: both already; this shows formula → value in output)")
    ap.add_argument("--context", type=int, default=1, help="unchanged lines shown around a text change")
    ap.add_argument("--out", type=Path)
    a = ap.parse_args(argv)
    for p in (a.old, a.new):
        if not p.exists():
            sys.exit(f"[office_diff] file not found: {p}")
    if a.old.suffix.lower() != a.new.suffix.lower():
        sys.exit("[office_diff] both files must have the same type")

    old, new = dump(a.old, a), dump(a.new, a)
    kind = old["type"]
    head = [f"# {a.old.name} → {a.new.name}", "", office2md.DATA_NOTICE, ""]
    if kind == "xlsx":
        body, n = diff_xlsx(old, new, a)
    elif kind == "docx":
        d = diff_lines(docx_lines(old), docx_lines(new), a.context, "p")
        body = (["```diff"] + d + ["```"]) if d else []
        n = sum(1 for l in d if l.startswith(("-", "+")))
        if old.get("comments") != new.get("comments"):
            body += ["", f"Comments: {len(old.get('comments', []))} → {len(new.get('comments', []))}"]
            n += 1
    elif kind == "pptx":
        body, n = diff_units(pptx_units(old), pptx_units(new), a.context)
    elif kind == "pdf":
        body, n = diff_units(pdf_units(old), pdf_units(new), a.context)
    elif kind == "csv":
        d = diff_lines([" | ".join(r) for r in old["rows"]], [" | ".join(r) for r in new["rows"]], a.context, "row")
        body = (["```diff"] + d + ["```"]) if d else []
        n = sum(1 for l in d if l.startswith(("-", "+")))
    else:
        sys.exit(f"[office_diff] no diff for type {kind}")

    summary = "No differences found." if n == 0 else f"{n} change(s) found."
    text = "\n".join(head + [summary, ""] + body).rstrip() + "\n"
    if a.out:
        a.out.write_text(text, encoding="utf-8")
        print(f"[office_diff] wrote {a.out}: {summary}")
    else:
        print(text)
    return 0 if n == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
