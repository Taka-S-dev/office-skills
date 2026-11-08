#!/usr/bin/env python3
"""Convert a Word file to Markdown with office2md, then cross-check it against pandoc.

Usage:
    md_crosscheck.py 仕様書.docx                      writes 仕様書.md, 仕様書.review.md (and figures/)
    md_crosscheck.py 仕様書.docx --out-dir ./md
    md_crosscheck.py 仕様書.docx --no-markitdown      skip the optional second reference

office2md keeps what pandoc drops (automatic heading numbers, text boxes, notes, figure alt text).
pandoc keeps a few things office2md does not render the same way (inline emphasis, LaTeX for
equations, merged-cell tables). Neither is complete alone, so this script runs both, keeps the
office2md result as the Markdown to deliver, and writes a review file listing every place the
references disagree with it: sentences present only in a reference, links, emphasis, equations
and table shapes. A reader works through that list, edits the Markdown, and reruns the check
until only intentional differences remain. With pandoc absent the review says so and the
Markdown is still produced.
"""

from __future__ import annotations

import argparse
import re
import shutil
import subprocess
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import office2md  # noqa: E402
import diagram_lint  # noqa: E402

try:
    sys.stdout.reconfigure(encoding="utf-8")  # Windows console defaults to cp932
except AttributeError:
    pass

NUM_PREFIX = re.compile(r"^(第\s*[0-9０-９一二三四五六七八九十]+\s*[章節項部編]|\d+(\.\d+)*\.?|\(\d+\)|[①-⑳]|[a-zA-Z][.)])\s*")


# -- running the engines ------------------------------------------------------------------

def run_office2md(src: Path, out_md: Path, images: Path, base: Path | None = None) -> str:
    # images_base makes figure paths relative to the folder the .md lives in, with forward slashes
    opts = argparse.Namespace(outline=False, profile=False, sheet=None, slide=None, max_rows=0, max_cols=0,
                              formulas=False, cells=False, form=False, no_coords=True, raw=False, json=False,
                              out=None, extract_attachments=None, mail_engine="auto", password=None, images=images,
                              images_base=base or out_md.parent)
    h = office2md.HANDLERS[src.suffix.lower()](src, opts)
    text = h.to_markdown(h.dump())
    out_md.write_text(text, encoding="utf-8")
    return text


def run_pandoc(src: Path, out_md: Path, media: Path) -> str | None:
    exe = shutil.which("pandoc")
    if not exe:
        return None
    # pandoc's own markdown writer keeps equations as $$...$$ (the gfm writer flattens them);
    # grid/simple/multiline tables are disabled so tables come out as pipe tables or HTML
    # absolute paths: an argument that is a path can then never begin with "-"
    r = subprocess.run([exe, str(src.resolve()), "-t", "markdown-grid_tables-simple_tables-multiline_tables",
                        "--wrap=none", "--track-changes=accept",
                        f"--extract-media={media.resolve()}", "-o", str(out_md.resolve())],
                       capture_output=True, text=True, encoding="utf-8", timeout=300)
    if r.returncode != 0:
        return None
    return out_md.read_text(encoding="utf-8")


def run_markitdown(src: Path, out_md: Path) -> str | None:
    try:
        from markitdown import MarkItDown
    except ImportError:
        return None
    try:
        text = MarkItDown().convert(str(src)).text_content
    except Exception as e:
        hint = " (pip install markitdown[docx])" if "ependenc" in str(e) else ""
        print(f"[md_crosscheck] markitdown could not convert: {type(e).__name__}{hint}", file=sys.stderr)
        return None
    out_md.write_text(text, encoding="utf-8")
    return text


# -- normalising for comparison -----------------------------------------------------------

def sentences(md: str) -> dict[str, str]:
    """Normalised sentence -> original line. Markdown syntax, numbering and whitespace removed so
    that the same content written two ways compares equal."""
    out: dict[str, str] = {}
    for raw in md.splitlines():
        line = raw.strip()
        if not line or line.startswith(("|---", "```", "_Content extracted", "_(table of contents", "[figure:", ":::")):
            continue
        line = re.sub(r"^>\s*\*\*Text box:\*\*\s*", "", line)      # text box content is content
        line = re.sub(r"\[\[\d+\]\]\(#[^)]*\)", "", line)           # markitdown footnote marks
        line = re.sub(r"\$\$?[^$]{1,400}\$\$?", "", line)           # equations are compared separately
        line = re.sub(r"\s*\{[^}]*\}\s*$", "", line)                 # pandoc heading attributes {.title}
        if line.startswith("|"):
            cells = [c.strip() for c in line.strip("|").split("|")]
            for c in cells:
                for s in split_sentences(c):
                    out.setdefault(s, raw)
            continue
        line = re.sub(r"^#{1,6}\s*", "", line)
        line = re.sub(r"^(\s*[-*+]\s+|\s*\d+\.\s+|>\s*)+", "", line)
        line = re.sub(r"\[\^[^\]]+\]", "", line)
        line = re.sub(r"!\[[^\]]*\]\([^)]*\)", "", line)
        line = re.sub(r"\[([^\]]*)\]\([^)]*\)", r"\1", line)
        line = re.sub(r"<[^>]+>", "", line)
        for s in split_sentences(line):
            out.setdefault(s, raw)
    return out


def split_sentences(text: str) -> list[str]:
    text = text.replace("**", "").replace("__", "").replace("*", "").replace("`", "").replace("\\", "")
    pieces = re.split(r"(?<=[。．.!?！？])\s*", text)
    out = []
    for p in pieces:
        p = NUM_PREFIX.sub("", p.strip())
        p = re.sub(r"[\s　]+", "", p)
        p = re.sub(r"[「」『』（）()\[\]【】、，,.。:：;；・\-ー–—]", "", p)
        if len(p) >= 4:
            out.append(p)
    return out


def links(md: str) -> dict[str, str]:
    return {u: t for t, u in re.findall(r"\[([^\]]+)\]\((https?://[^)\s]+|mailto:[^)\s]+)\)", md)}


def emphasis(md: str) -> list[str]:
    return [m.strip() for m in re.findall(r"\*\*([^*\n]{2,80})\*\*", md) if m.strip()]


def equations(md: str) -> list[str]:
    return [e.strip() for e in re.findall(r"\$([^$\n]{1,200})\$", md)] + \
           [e.strip() for e in re.findall(r"\$\$([^$\n]{1,400})\$\$", md)]


def code_span(s: str) -> str:
    """Inline code whose fence is longer than any backtick run inside, so content cannot close it."""
    s = re.sub(r"[\x00-\x1f\x7f]", " ", s)[:200]
    longest = max((len(m) for m in re.findall(r"`+", s)), default=0)
    fence = "`" * (longest + 1)
    return f"{fence} {s} {fence}" if longest else f"`{s}`"


def tables(md: str) -> list[int]:
    """Cell count per table, in order."""
    counts, cur, in_table = [], 0, False
    for line in md.splitlines() + [""]:
        if line.strip().startswith("|"):
            in_table = True
            if not re.match(r"^\|?\s*:?-{2,}", line.strip()):
                cur += line.count("|") - 1
        elif in_table:
            counts.append(cur)
            cur, in_table = 0, False
    return counts


def headings(md: str) -> list[str]:
    return [re.sub(r"\s+", " ", h.strip()) for h in re.findall(r"^#{1,6}\s+(.+)$", md, re.M)]


def esc(s: str) -> str:
    return office2md.md_escape(s).replace("<br>", " / ")


# -- safe automatic fixes -----------------------------------------------------------------
# Only changes that cannot be wrong are applied: a reference's LaTeX replacing the primary's plain
# text of the same equation, a phrase bold in the reference and occurring exactly once in the
# primary, a link whose text occurs exactly once in the primary. Everything else stays a review item.

PROTECTED = re.compile(r"^(#{1,6}\s|\||```|> \*\*Text box|\[figure:|\[\^|_\(|_Content)")


def norm_eq(s: str) -> str:
    s = re.sub(r"\\(times|cdot)", "×", s)
    s = re.sub(r"\\[a-zA-Z]+", "", s)
    s = s.replace("\\", "").replace("{", "").replace("}", "")
    s = s.replace("*", "×").replace("·", "×")
    return re.sub(r"\s+", "", s)


def apply_safe(primary: str, refs: dict[str, str]) -> tuple[str, dict[str, int]]:
    lines = primary.split("\n")
    stats = {"equations": 0, "bold": 0, "links": 0}
    body_idx = [i for i, l in enumerate(lines) if l.strip() and not PROTECTED.match(l.strip())]

    def occurrences(phrase: str) -> list[int]:
        return [i for i in body_idx if phrase in lines[i]]

    for text in refs.values():
        # equations: plain-text line in the primary that is the same equation as a reference's LaTeX
        for latex in dict.fromkeys(equations(text)):
            target = norm_eq(latex)
            if len(target) < 3:
                continue
            hits = [i for i in body_idx if norm_eq(lines[i]) == target and "$" not in lines[i]]
            if len(hits) == 1:
                lines[hits[0]] = f"$${latex.strip()}$$"
                stats["equations"] += 1
        # links: the reference knows a URL for a phrase the primary has once, unlinked
        for url, phrase in links(text).items():
            phrase = phrase.strip()
            if not phrase or f"({url})" in primary or not url.lower().startswith(("http://", "https://", "mailto:")):
                continue
            hits = [i for i in occurrences(phrase) if f"[{phrase}]" not in lines[i]]
            if len(hits) == 1 and lines[hits[0]].count(phrase) == 1:
                lines[hits[0]] = lines[hits[0]].replace(phrase, f"[{phrase}]({url})")
                stats["links"] += 1
        # bold: a phrase the reference emphasises, present exactly once in the primary and not yet marked
        for phrase in dict.fromkeys(emphasis(text)):
            if len(phrase) < 2 or re.fullmatch(r"[\d\s.,%]+", phrase) or "`" in phrase or "$" in phrase:
                continue
            hits = occurrences(phrase)
            if len(hits) == 1 and lines[hits[0]].count(phrase) == 1 and f"**{phrase}**" not in lines[hits[0]] \
                    and "](" not in lines[hits[0]]:
                lines[hits[0]] = lines[hits[0]].replace(phrase, f"**{phrase}**")
                stats["bold"] += 1
    return "\n".join(lines), stats


# -- the review ---------------------------------------------------------------------------

def review(primary: str, refs: dict[str, str], src: Path, applied: dict[str, int] | None = None) -> tuple[str, int]:
    out = [f"# Cross-check of {src.name}", "", office2md.DATA_NOTICE, ""]
    if not refs:
        out.append("Status: NO REFERENCE. pandoc is not installed and markitdown is unavailable, so the "
                   "Markdown comes from office2md alone and nothing could be compared. Say so when handing it over.")
        return "\n".join(out) + "\n", 0
    status_at = len(out)          # the status block is inserted here once the counts are known
    out.append("Primary: office2md. References: " + ", ".join(refs) + ". Each item below is a place where a "
               "reference has something the primary lacks, or the two disagree. Decide per item: fix the "
               "primary Markdown, or leave it when the difference is intentional (collapsed table of contents, "
               "dropped page numbers, numbering that only the primary has).")
    out.append("")
    n_items = 0
    counts = {"missing": 0, "links": 0, "equations": 0, "bold": 0, "tables": 0, "headings": 0}
    own = {f"# {src.name}", "## Notes", "## Comments"}        # lines office2md adds itself
    # drop those lines before sentence extraction: "# 仕様書.docx" would otherwise yield the sentence
    # "仕様書" and shadow the document's real title heading
    p_sent = sentences("\n".join(l for l in primary.splitlines() if l.strip() not in own))
    p_links, p_emph, p_eq, p_tables = links(primary), emphasis(primary), equations(primary), tables(primary)
    p_heads = [h for h in headings(primary) if h not in (src.name, "Notes", "Comments")]

    for name, text in refs.items():
        r_sent = sentences(text)
        missing = [r_sent[s] for s in r_sent if s not in p_sent]
        # ignore reference lines that are just the TOC the primary collapsed on purpose
        missing = [m for m in missing if not re.search(r"\.{3,}\s*\d+\s*$|…+\s*\d+\s*$", m)]
        out.append(f"## Content in {name} not found in the primary ({len(missing)})")
        out.append("")
        if missing:
            out.append("Each line is as the reference wrote it. Search the primary for the words; if the content "
                       "is truly absent, add it where it belongs.")
            out.append("")
            for m in missing[:150]:
                out.append(f"- {esc(m.strip())[:200]}")
            if len(missing) > 150:
                out.append(f"- … {len(missing) - 150} more")
            n_items += len(missing)
            counts["missing"] += len(missing)
        else:
            out.append("None.")
        out.append("")

        extra = [p_sent[s] for s in p_sent if s not in r_sent]
        out.append(f"## Content in the primary not found in {name} ({len(extra)})")
        out.append("")
        out.append("Usually text boxes, notes, figure alt text and numbering the reference drops; listed so a "
                   "duplicated or misplaced line can be spotted.")
        out.append("")
        for m in extra[:60]:
            out.append(f"- {esc(m.strip())[:200]}")
        if len(extra) > 60:
            out.append(f"- … {len(extra) - 60} more")
        out.append("")

        r_links = links(text)
        new_links = {u: t for u, t in r_links.items() if u not in p_links}
        if new_links:
            out.append(f"## Links in {name} missing from the primary ({len(new_links)})")
            out.append("")
            out.append("Add as [text](url) at the matching text.")
            out.append("")
            for u, t in list(new_links.items())[:100]:
                out.append(f"- [{esc(t)}]({u})")
            out.append("")
            n_items += len(new_links)
            counts["links"] += len(new_links)

        r_eq = equations(text)
        p_eq_norm = {norm_eq(e) for e in p_eq}
        new_eq = list(dict.fromkeys(e for e in r_eq if norm_eq(e) not in p_eq_norm))
        if new_eq:
            out.append(f"## Equations as LaTeX in {name} ({len(new_eq)})")
            out.append("")
            out.append("The primary has each equation as plain text; replace with $...$ where the exact form matters.")
            out.append("")
            for e in new_eq[:60]:
                out.append(f"- {code_span(e)}")
            out.append("")
            n_items += len(new_eq)
            counts["equations"] += len(new_eq)

        r_emph = [e for e in emphasis(text) if e not in p_emph]
        if r_emph:
            uniq = list(dict.fromkeys(r_emph))
            out.append(f"## Bold phrases in {name} not marked in the primary ({len(uniq)})")
            out.append("")
            out.append("Mark with **...** only where the emphasis carries meaning (必須, 禁止, 注意).")
            out.append("")
            for e in uniq[:80]:
                out.append(f"- {esc(e)}")
            out.append("")
            counts["bold"] += len(uniq)

        r_tables = tables(text)
        if r_tables != p_tables:
            out.append(f"## Table shapes differ from {name}")
            out.append("")
            out.append(f"Primary: {len(p_tables)} table(s) with cell counts {p_tables[:30]}; {name}: "
                       f"{len(r_tables)} with {r_tables[:30]}. A larger count in the reference usually means "
                       "merged cells it expanded or a nested table it flattened; check those tables by eye.")
            out.append("")
            n_items += 1
            counts["tables"] += 1

        r_heads = headings(text)
        strip = lambda h: NUM_PREFIX.sub("", h).strip()  # noqa: E731
        only_ref = [h for h in r_heads if strip(h) not in {strip(x) for x in p_heads}]
        only_pri = [h for h in p_heads if strip(h) not in {strip(x) for x in r_heads}]
        if only_ref or only_pri:
            out.append(f"## Headings that differ from {name}")
            out.append("")
            for h in only_ref[:40]:
                out.append(f"- only in {name}: {esc(h)}")
            for h in only_pri[:40]:
                out.append(f"- only in primary: {esc(h)}")
            out.append("")
            n_items += len(only_ref)
            counts["headings"] += len(only_ref)

    out.append("## Done when")
    out.append("")
    out.append("Rerun this script after editing. The lists above should contain only differences you decided to keep.")

    # status block at the top: the one place to look to know whether this is finished
    lint = diagram_lint.lint_markdown(primary)
    if lint:
        out += ["## Diagram syntax problems (" + str(sum(len(e) for _, _, e in lint)) + ")", ""]
        for lang, line, errs in lint:
            out += [f"- {lang} block at line {line}: {esc(e)}" for e in errs]
        out.append("")
    blocking = counts["missing"] + counts["links"] + counts["headings"] + sum(len(e) for _, _, e in lint)
    verdict = "READY" if blocking == 0 and counts["tables"] == 0 else "NEEDS REVIEW"
    status = [f"**Status: {verdict}**", ""]
    if applied:
        status.append("Applied automatically (safe, exact matches only): "
                      + ", ".join(f"{v} {k}" for k, v in applied.items() if v) + "." if any(applied.values())
                      else "Applied automatically: nothing matched.")
    status.append(f"Remaining: {counts['missing']} sentence(s) only in a reference, {counts['links']} link(s), "
                  f"{counts['equations']} equation(s) not applied, {counts['bold']} bold phrase(s) (judgement call), "
                  f"{counts['tables']} table shape difference(s), {counts['headings']} heading difference(s).")
    status.append("READY means no reference has content, links or headings the primary lacks and the tables "
                  "match in shape; bold and equation items are optional polish.")
    status.append("")
    out[status_at:status_at] = status
    return "\n".join(out) + "\n", blocking + counts["tables"]     # the count printed = what blocks READY


# -- Excel ------------------------------------------------------------------------------------
# No converter does better than office2md on a workbook, so the references are the workbook
# itself: every text it contains must appear in the Markdown (content check, survives any
# restructuring an agent does), and each sheet is rendered to PDF so that what was reconstructed
# from shapes (sequence diagrams, flowcharts, charts, pictures) can be compared with the picture.
# Visual checks are confirmed by writing their id into <stem>.checked; the verdict is READY when
# nothing is missing and every visual check is confirmed.

XLSX_OPTS = dict(outline=False, profile=False, sheet=None, slide=None, max_rows=0, max_cols=0, formulas=False,
                 cells=False, form=False, no_coords=False, raw=False, json=False, out=None, blocks=True,
                 extract_attachments=None, mail_engine="auto", password=None)


def _norm(s: str) -> str:
    return re.sub(r"<br>|[\s\\*|/`]", "", str(s))


def xlsx_inventory(data: dict) -> list[tuple[str, str]]:
    """(where, text) for every piece of text in the workbook that the Markdown must carry."""
    items: list[tuple[str, str]] = []
    for s in data["sheets"]:
        name = s["name"]
        for c in s["cells"]:
            items.append((f"{name}!{c['address']}", str(c["display"])))
        for d in s.get("drawings", []):
            if d.get("text"):
                items.append((f"{name} {d['kind']} '{d['name']}' at {d['anchor']}", d["text"]))
        for ch in s.get("charts", []):
            if ch.get("title"):
                items.append((f"{name} chart '{ch['name']}' title", ch["title"]))
            for se in ch.get("series", []):
                if se.get("name"):
                    items.append((f"{name} chart '{ch['name']}' series", se["name"]))
        for v in s.get("validations", []):
            if v["rule"].startswith("list: "):
                for choice in v["rule"][6:].split(" (from ")[0].split(", "):
                    items.append((f"{name} input rule {v['range']}", choice))
        for c in s.get("comments", []):
            items.append((f"{name}!{c['cell']} comment", c["text"] or ""))
    keep = []
    for where, text in items:
        n = _norm(text)
        if len(n) < 2 or re.fullmatch(r"-?\d{1,2}", n):
            continue        # "1", "8": present everywhere, says nothing about completeness
        keep.append((where, text))
    return keep


def visual_items(data: dict, sheet_pdf: dict[str, str]) -> list[tuple[str, str]]:
    items = []
    for s in data["sheets"]:
        name = s["name"]
        pdf = sheet_pdf.get(name, "")
        where = f" in {pdf}" if pdf else ""
        d = s.get("diagram")
        if d and (d["sequence"] or d["edges"] or d["unexplained"]):
            what = []
            if d["sequence"]:
                what.append(f"{len(d['sequence']['messages'])} message(s) between "
                            f"{len(d['sequence']['participants'])} participant(s)")
            if d["edges"]:
                what.append(f"{len(d['edges'])} connection(s)")
            if d["unexplained"]:
                what.append(f"{len(d['unexplained'])} line(s) not interpreted")
            items.append((f"{name}#diagram", f"Compare the Diagram section of sheet {name} ({'; '.join(what)}) with the "
                          f"drawing{where}: participants, order, arrow directions, labels. Correct the Markdown, and "
                          "describe any line not interpreted."))
        for i, ch in enumerate(s.get("charts", []), start=1):
            items.append((f"{name}#chart{i}", f"Check chart '{ch['name']}' (title, series, values) against the drawn chart{where}."))
        for i, pc in enumerate(s.get("pictures", []), start=1):
            if pc.get("placeholder"):
                continue
            items.append((f"{name}#picture{i}", f"Picture '{pc['name']}' at {pc['from']}: its content is not extracted. "
                          f"Look at it{where} and describe what it shows in the Markdown, or note that it carries "
                          "no information."))
    return items


def render_sheets(src: Path, out: Path, sheet_names: list[str]) -> tuple[dict[str, str], str]:
    """Render each visible sheet to its own PDF, one page wide. Excel through COM when available,
    else the whole workbook through LibreOffice. Returns ({sheet: pdf path}, method)."""
    out.mkdir(parents=True, exist_ok=True)
    if sys.platform == "win32":
        ps = r"""
$ErrorActionPreference = 'Stop'
$app = New-Object -ComObject Excel.Application
try {
  $app.Visible = $false; $app.DisplayAlerts = $false; $app.AutomationSecurity = 3
  $wb = $app.Workbooks.Open($env:OFFICE_SRC, 0, $true)
  $i = 0
  foreach ($ws in $wb.Worksheets) {
    $i++
    if ($ws.Visible -ne -1) { continue }
    try {
      $ps = $ws.PageSetup; $ps.PrintArea = ''; $ps.Zoom = $false; $ps.FitToPagesWide = 1; $ps.FitToPagesTall = $false
    } catch { }
    try {
      $ws.ExportAsFixedFormat(0, (Join-Path $env:OFFICE_OUT ('sheet{0:D2}.pdf' -f $i)))
      Write-Output ('OK {0}' -f $i)
    } catch { Write-Output ('EMPTY {0}' -f $i) }
  }
  $wb.Close($false)
} finally { $app.Quit() }
"""
        try:
            r = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", ps],
                               capture_output=True, text=True, timeout=300,
                               env={**__import__("os").environ, "OFFICE_SRC": str(src.resolve()), "OFFICE_OUT": str(out.resolve())})
            done = {int(x.split()[1]) for x in r.stdout.split("\n") if x.startswith("OK ")}
            if done:
                return {sheet_names[i - 1]: f"{out.name}/sheet{i:02d}.pdf" for i in sorted(done)
                        if 0 < i <= len(sheet_names)}, "Excel"
        except (subprocess.TimeoutExpired, OSError):
            pass
    so = shutil.which("soffice") or (r"C:\Program Files\LibreOffice\program\soffice.exe"
                                      if Path(r"C:\Program Files\LibreOffice\program\soffice.exe").exists() else None)
    if so:
        try:
            r = subprocess.run([so, "--headless", "--convert-to", "pdf", "--outdir", str(out.resolve()), str(src.resolve())],
                               capture_output=True, text=True, timeout=300)
            pdf = out / (src.stem + ".pdf")
            if pdf.exists():
                return {name: f"{out.name}/{pdf.name} (all sheets in tab order)" for name in sheet_names}, "LibreOffice"
        except (subprocess.TimeoutExpired, OSError):
            pass
    return {}, ""


def xlsx_main(src: Path, a) -> int:
    out_dir = a.out_dir or src.parent
    out_dir.mkdir(parents=True, exist_ok=True)
    stem = src.stem
    primary_path = out_dir / f"{stem}.md"
    fresh = a.fresh or not primary_path.exists()
    opts = argparse.Namespace(**XLSX_OPTS, uml=a.uml, images=(out_dir / f"{stem}_figures") if fresh else None,
                              images_base=out_dir, drawio=out_dir / f"{stem}.drawio")
    h = office2md.HANDLERS[src.suffix.lower()](src, opts)
    data = h.dump()
    if fresh:
        primary = h.to_markdown(data)
        primary_path.write_text(primary, encoding="utf-8")
    else:
        primary = primary_path.read_text(encoding="utf-8")
        print(f"[md_crosscheck] reusing existing {primary_path.name} (pass --fresh to regenerate from the workbook)")

    sheet_pdf, method = {}, ""
    render_dir = out_dir / f"{stem}_sheets"
    if not a.no_render:
        names = [s["name"] for s in data["sheets"]]
        if render_dir.exists() and any(render_dir.glob("*.pdf")) and not a.fresh:
            sheet_pdf = {n: f"{render_dir.name}/sheet{i:02d}.pdf" for i, n in enumerate(names, start=1)
                         if (render_dir / f"sheet{i:02d}.pdf").exists()}
            method = "existing renders"
        else:
            sheet_pdf, method = render_sheets(src, render_dir, names)

    inv = xlsx_inventory(data)
    md_norm = _norm(primary)
    missing = [(w, t) for w, t in inv if _norm(t) not in md_norm]
    visuals = visual_items(data, sheet_pdf)
    checked_path = out_dir / f"{stem}.checked"
    checked_lines = checked_path.read_text(encoding="utf-8").splitlines() if checked_path.exists() else []
    done = {vid for vid, _ in visuals if any(ln.strip().startswith(vid) for ln in checked_lines)}
    lint = diagram_lint.lint_markdown(primary)
    lint_count = sum(len(e) for _, _, e in lint)
    rendered = []
    if a.render_diagrams and diagram_lint.renderers():
        rendered = diagram_lint.render(primary, render_dir / "diagrams")
    render_errors = [r for r in rendered if r[3]]
    ready = not missing and len(done) == len(visuals) and not lint_count and not render_errors

    def ident(s: str) -> str:
        return s.replace("`", "'")

    out = [f"# Cross-check of {src.name}", "", office2md.DATA_NOTICE, "",
           f"**Status: {'READY' if ready else 'NEEDS REVIEW'}**", "",
           f"Text check: {len(inv)} text item(s) in the workbook, {len(missing)} missing from the Markdown.",
           f"Diagram check: {len(diagram_lint.blocks(primary))} mermaid/PlantUML block(s), {lint_count} problem(s)"
           + (f"; rendered with the installed tool: {len(rendered) - len(render_errors)} ok, {len(render_errors)} failed."
              if rendered else "; structure checked (add --render-diagrams to also render with an installed tool)."),
           f"Visual checks: {len(done)} of {len(visuals)} confirmed (record confirmations in {checked_path.name}).",
           (f"Rendered sheets ({method}): " + "; ".join(f"{k} → {v}" for k, v in sheet_pdf.items())) if sheet_pdf
           else "Rendered sheets: none (no Excel or LibreOffice, or --no-render); visual checks need another way to see the sheets.",
           "READY means every text in the workbook appears in the Markdown and every visual check is confirmed. "
           "The Markdown may be restructured freely (rewrite a diagram, merge blocks into one table); the text check "
           "keeps that from dropping anything.", ""]
    if lint or render_errors:
        out += [f"## Diagram problems ({lint_count + len(render_errors)})", "",
                "A diagram that does not parse is drawn wrongly or not at all. Fix the block at the line given."]
        out.append("")
        for lang, line, errs in lint:
            out += [f"- {lang} block at line {line}: {esc(e)}" for e in errs]
        for lang, line, _, err in render_errors:
            out.append(f"- {lang} block at line {line}: the renderer rejected it: {esc(err)}")
        out.append("")
    if rendered and not render_errors:
        out += ["Diagrams rendered back from the Markdown (compare with the sheet renders): "
                + "; ".join(f"{Path(img).name}" for _, _, img, _ in rendered), ""]
    out += [f"## Text in the workbook missing from the Markdown ({len(missing)})", ""]
    if missing:
        out.append("Each line is where the text sits in the workbook. Put it back in the Markdown where it belongs.")
        out.append("")
        for w, t in missing[:200]:
            out.append(f"- {esc(w)}: {esc(t)[:200]}")
        if len(missing) > 200:
            out.append(f"- … {len(missing) - 200} more")
    else:
        out.append("None.")
    out += ["", f"## Visual checks ({len(done)} of {len(visuals)} confirmed)", ""]
    if visuals:
        for vid, what in sorted(visuals, key=lambda v: v[0] in done):
            out.append(f"- [{'x' if vid in done else ' '}] `{ident(vid)}` — {what}")
        out += ["", f"After checking one, add a line starting with its id to {checked_path.name}, with what you found, "
                f"for example: `{ident(visuals[0][0])}  matches the rendered sheet`."]
    else:
        out.append("None: no diagrams, charts or pictures with content.")
    notes = []
    for s in data["sheets"]:
        d = s.get("diagram")
        if d and d["unexplained"]:
            notes.append(f"- {esc(s['name'])}: {len(d['unexplained'])} line(s) the reconstruction could not attach to "
                         "shapes or lifelines; they are listed in the sheet's Diagram section.")
        if s.get("uncached_formulas"):
            notes.append(f"- {esc(s['name'])}: {len(s['uncached_formulas'])} formula(s) without a cached value; their "
                         "results are unknown until the workbook is recalculated.")
    if notes:
        out += ["", "## Notes", ""] + notes
    out += ["", "## Done when", "",
            "Rerun after editing the Markdown or the .checked file. A rerun keeps the edits (--fresh starts over)."]
    review_path = out_dir / f"{stem}.review.md"
    review_path.write_text("\n".join(out) + "\n", encoding="utf-8")
    print(f"[md_crosscheck] primary:   {primary_path}")
    if sheet_pdf:
        print(f"[md_crosscheck] rendered:  {render_dir} ({method})")
    print(f"[md_crosscheck] review:    {review_path}  "
          f"({'READY' if ready else f'{len(missing)} missing, {len(visuals) - len(done)} visual check(s) open'})")
    return 0


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("file", type=Path)
    ap.add_argument("--out-dir", type=Path, help="where to write the .md, .review.md and figures (default: next to the file)")
    ap.add_argument("--no-markitdown", action="store_true")
    ap.add_argument("--no-pandoc", action="store_true")
    ap.add_argument("--no-apply", action="store_true",
                    help="report only; do not fold safe exact-match fixes (equations, links, bold) into the primary")
    ap.add_argument("--fresh", action="store_true",
                    help="regenerate the .md from the file even if one exists (a plain rerun keeps your edits)")
    ap.add_argument("--uml", choices=["mermaid", "plantuml", "both"], default="mermaid",
                    help="xlsx: notation for reconstructed diagrams")
    ap.add_argument("--render-diagrams", action="store_true",
                    help="also render each mermaid/PlantUML block with an installed tool (PlantUML runs sandboxed)")
    ap.add_argument("--no-render", action="store_true",
                    help="xlsx: do not render sheets to PDF (visual checks then have nothing to compare against)")
    a = ap.parse_args(argv)
    src = a.file
    if not src.exists():
        sys.exit(f"[md_crosscheck] file not found: {src}")
    if src.suffix.lower() in (".xlsx", ".xlsm", ".xltx", ".xltm"):
        return xlsx_main(src, a)
    if src.suffix.lower() not in (".docx", ".docm", ".dotx"):
        sys.exit("[md_crosscheck] Word and Excel files only")
    out_dir = a.out_dir or src.parent
    out_dir.mkdir(parents=True, exist_ok=True)
    stem = src.stem
    primary_path = out_dir / f"{stem}.md"
    if primary_path.exists() and not a.fresh:
        # a rerun checks the Markdown as it is now, edits included; --fresh starts over
        primary = primary_path.read_text(encoding="utf-8")
        print(f"[md_crosscheck] reusing existing {primary_path.name} (pass --fresh to regenerate from the .docx)")
    else:
        primary = run_office2md(src, primary_path, out_dir / f"{stem}_figures", out_dir)

    refs: dict[str, str] = {}
    if not a.no_pandoc:
        t = run_pandoc(src, out_dir / f"{stem}.pandoc.md", out_dir / f"{stem}_pandoc_media")
        if t is not None:
            refs["pandoc"] = t
    if not a.no_markitdown:
        t = run_markitdown(src, out_dir / f"{stem}.markitdown.md")
        if t is not None:
            refs["markitdown"] = t

    applied = None
    if refs and not a.no_apply:
        primary, applied = apply_safe(primary, refs)
        primary_path.write_text(primary, encoding="utf-8")
    text, n = review(primary, refs, src, applied)
    review_path = out_dir / f"{stem}.review.md"
    review_path.write_text(text, encoding="utf-8")
    print(f"[md_crosscheck] primary:   {primary_path}")
    for name in refs:
        print(f"[md_crosscheck] reference: {out_dir / f'{stem}.{name}.md'}")
    if not refs:
        print("[md_crosscheck] no reference converter found (pandoc / markitdown); primary written, nothing compared")
    print(f"[md_crosscheck] review:    {review_path}  ({n} item(s) to look at)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
