"""Rebuild a PDF page's paragraphs and table cells from where its text sits.

pypdf's plain text gives one line per printed line. In a table with narrow columns a cell's
sentence is printed four or five characters per line, so plain text cuts it into fragments and
interleaves them with the cells beside it. This module reads the position and size of every text
run, and

1. joins runs at the same height into a line, but splits the line where the gap is wide (a new
   column starts there);
2. joins a line to the one above when they start at the same left edge and sit at the usual line
   spacing: one paragraph, or one cell of a table column;
3. lays the blocks out top to bottom; blocks whose vertical extents overlap side by side become a
   row of a table, a block alone is a paragraph.

The result is approximate by nature (no ruling lines are read), so the caller says so.
"""

from __future__ import annotations

import re
import statistics

SENT_END = tuple("。．.!?！？")


def _width(text: str, size: float) -> float:
    """Printed width estimate: full-width characters one em, others half an em."""
    import unicodedata
    return sum(size if unicodedata.east_asian_width(ch) in ("W", "F", "A") else size * 0.5 for ch in text)


def runs_of(page) -> list[dict]:
    runs: list[dict] = []

    def visit(text, cm, tm, font_dict, font_size):
        t = text.replace("\n", "")
        if not t.strip():
            return
        scale = abs(tm[0] * cm[0]) or abs(tm[3] * cm[3]) or 1.0
        size = (font_size or 10) * scale
        x = tm[4] * cm[0] + cm[4]
        y = tm[5] * cm[3] + cm[5]
        runs.append({"x": x, "y": y, "size": size, "text": t})
    page.extract_text(visitor_text=visit)
    return runs


def segments(runs: list[dict]) -> list[dict]:
    """Runs at one height, merged left to right while the gap stays small."""
    rows: list[list[dict]] = []
    for r in sorted(runs, key=lambda r: (-r["y"], r["x"])):
        if rows and abs(rows[-1][0]["y"] - r["y"]) <= max(1.5, 0.25 * r["size"]):
            rows[-1].append(r)
        else:
            rows.append([r])
    segs: list[dict] = []
    for row in rows:
        row.sort(key=lambda r: r["x"])
        cur = None
        for r in row:
            if cur is not None and r["x"] - cur["end"] <= 0.9 * r["size"]:
                gap = r["x"] - cur["end"]
                cur["text"] += (" " if gap > 0.3 * r["size"] and cur["text"][-1:].isascii() else "") + r["text"].strip()
                cur["end"] = max(cur["end"], r["x"] + _width(r["text"], r["size"]))
            else:
                cur = {"x": r["x"], "y": r["y"], "size": r["size"], "text": r["text"].strip(),
                       "end": r["x"] + _width(r["text"], r["size"])}
                segs.append(cur)
    return segs


def blocks(segs: list[dict]) -> list[dict]:
    """Segments stacked at the same left edge and the usual spacing, joined into one block."""
    gaps = [a["y"] - b["y"] for a, b in zip(segs, segs[1:]) if 0 < a["y"] - b["y"] < 3 * a["size"]]
    spacing = statistics.median(gaps) if gaps else 14.0
    out: list[dict] = []
    for s in segs:
        best = None
        for b in out:
            if b["closed"]:
                continue
            dy = b["bottom"] - s["y"]
            # same column: the left edge matches, or this line runs on under an indented first line
            if 0.4 * spacing <= dy <= 1.6 * spacing and abs(b["x"] - s["x"]) <= 1.2 * s["size"]:
                best = b
                break
        if best is not None:
            prev = best["lines"][-1]
            joiner = " " if prev[-1:].isascii() and s["text"][:1].isascii() else ""
            best["lines"][-1] = prev + joiner + s["text"] if not prev.endswith(SENT_END) else prev
            if prev.endswith(SENT_END):
                best["lines"].append(s["text"])
            best["bottom"], best["x"] = s["y"], min(best["x"], s["x"])
            best["right"] = max(best["right"], s["end"])
        else:
            out.append({"x": s["x"], "top": s["y"], "bottom": s["y"], "right": s["end"], "size": s["size"],
                        "lines": [s["text"]], "closed": False})
        for b in out:                                  # a block far above can no longer grow
            if b is not best and b["bottom"] - s["y"] > 1.6 * spacing:
                b["closed"] = True
    return out


def layout_markdown(page, escape) -> tuple[str, bool]:
    """(Markdown of the page, whether a table was rebuilt)."""
    bl = blocks(segments(runs_of(page)))
    bl.sort(key=lambda b: (-b["top"], b["x"]))
    rows: list[list[dict]] = []
    for b in bl:
        if rows:
            band = rows[-1]
            top = max(x["top"] for x in band)
            bottom = min(x["bottom"] for x in band)
            # overlaps the band vertically and does not overlap any member horizontally
            if b["top"] >= bottom - 0.5 * b["size"] and b["bottom"] <= top + 0.5 * b["size"] and \
                    all(b["x"] >= x["right"] - 2 or b["right"] <= x["x"] + 2 for x in band):
                band.append(b)
                continue
        rows.append([b])
    out: list[str] = []
    table: list[list[str]] = []
    rebuilt = False

    def flush():
        nonlocal table, rebuilt
        if not table:
            return
        width = max(len(r) for r in table)
        if len(table) == 1 and width <= 1:
            out.append(table[0][0])
        else:
            rebuilt = True
            out.append("| " + " | ".join([""] * width) + " |")
            out.append("|" + "---|" * width)
            for r in table:
                out.append("| " + " | ".join(escape(c).replace("\n", "<br>") for c in r + [""] * (width - len(r))) + " |")
        out.append("")
        table = []
    for band in rows:
        band.sort(key=lambda b: b["x"])
        cells = ["\n".join(b["lines"]) for b in band]
        if len(band) >= 2:
            table.append(cells)
        else:
            flush()
            out += [re.sub(r"^([#>*+\-|]|\d+\. )", r"\\\1", ln) for ln in cells[0].split("\n")]
            out.append("")
    flush()
    return "\n".join(out).strip(), rebuilt
