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
import html
import re
import shutil
import subprocess
import sys
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import office2md  # noqa: E402
from office2md import radicals  # noqa: E402
import diagram_lint  # noqa: E402
import review_checks as rc  # noqa: E402
import office_com  # noqa: E402

try:
    sys.stdout.reconfigure(encoding="utf-8")  # Windows console defaults to cp932
except AttributeError:
    pass

# a number is a label only when a dot or a space follows it: "３類型に…" and "10件まで" keep their digits
NUM_PREFIX = re.compile(r"^(第\s*[0-9０-９一二三四五六七八九十]+\s*[章節項部編]|[0-9０-９]+([.．][0-9０-９]+)*([.．]|(?=\s))"
                        r"|[(（][0-9０-９]+[)）]|[0-9０-９]{1,3}[)）]|[①-⑳]|[a-zA-Z][.)](?=\s)"
                        r"|[ⅰ-ⅿ]{1,4}[)）]?|[ぁ-んァ-ヶ][)）.．]?(?=\s)|[(（][a-zA-Z]{1,4}[)）](?=\s))\s*")  # 1) ア ⅱ) (a)


# -- running the engines ------------------------------------------------------------------

def run_office2md(src: Path, out_md: Path, images: Path | None, base: Path | None = None,
                  write: bool = True) -> tuple[str, dict]:
    # images_base makes figure paths relative to the folder the .md lives in, with forward slashes
    opts = argparse.Namespace(outline=False, profile=False, sheet=None, slide=None, max_rows=0, max_cols=0,
                              formulas=False, cells=False, form=False, no_coords=True, raw=False, json=False,
                              out=None, extract_attachments=None, mail_engine="auto", password=None, images=images,
                              images_base=base or out_md.parent)
    h = office2md.HANDLERS[src.suffix.lower()](src, opts)
    data = h.dump()
    text = h.to_markdown(data)
    if write:
        out_md.write_text(text, encoding="utf-8")
    return text, data


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

# a line that is only an internal link whose text ends in a page number: a TOC entry as pandoc
# ("[はじめに [1](#section)](#section)") or markitdown ("[はじめに 1](#_Toc119672580)") write it
# a table-of-contents entry: [見出し 12](#anchor), also with a roman page (ii), nested, or quoted
TOC_LINE = re.compile(r"^(?:[-*>]\s{0,4}){0,3}\[[^\n]{0,300}?\s\[?(?:[0-9]+|[ivxlcIVXLC]{1,6})\]?"
                      r"(\(#[^)]{0,300}\))?\]\(#[^)]{0,300}\)\s*$")
# markitdown's form: 1.2　見出し [2](#anchor), the page as the only link at the end
TOC_TAIL = re.compile(r"^[^\[\]\n]{1,300}?\s\[(?:[0-9]+|[ivxlcIVXLC]{1,6})\]\(#[^)]{0,300}\)\s*$")
# the tags pandoc and markitdown write; anything else in angle brackets is text ("<Insert Name>")
HTML_TAG = re.compile(r"</?(?:p|td|th|tr|br|em|strong|a|sup|sub|span|div|table|tbody|thead|tfoot|ul|ol|li|img|u|s|b|i|del|ins|code|colgroup|col|caption|blockquote|h[1-6]|figure|figcaption|pre|hr|mark|section|nav|small|dl|dt|dd)\b[^<>]{0,500}>", re.I)
# characters scanned by the substring fallback per reference (a 118-page plan needs about 1e8)
SUBSTRING_BUDGET = 2_000_000_000
# Every pattern below runs on document text, so every repeat is bounded: a crafted line must not
# make the review take time growing with the square of its length.


def _strip_between(text: str, start: str, end: str) -> str:
    """text without each start…end span, scanned once (a lazy regex would rescan a long line)."""
    out, i = [], 0
    while True:
        a = text.find(start, i)
        if a < 0:
            break
        b = text.find(end, a + len(start))
        if b < 0:
            break
        out.append(text[i:a])
        i = b + len(end)
    out.append(text[i:])
    return "".join(out)


def sentence_entries(md: str, known: set[str] | None = None) -> list[tuple[str, str, str]]:
    """(normalised sentence, original line, section) for every sentence, in order. Markdown
    syntax, numbering and whitespace are removed so that the same content written two ways
    compares equal; the section is the normalised text of the nearest heading above. With known,
    only headings in that set count: a heading one converter adds (a "（１）" line promoted to a
    heading) then does not move every sentence under it to another section."""
    out: list[tuple[str, str, str]] = []
    section = ""
    for raw in md.splitlines():
        line = raw.strip()
        if not line or line.startswith(("|---", "```", "_Content extracted", "_Source:", "_(table of contents", ":::")):
            continue
        if len(line) <= 1000 and (TOC_LINE.match(line) or TOC_TAIL.match(line)):
            continue                                                 # a table-of-contents entry: [見出し 12](#anchor)
        line = _strip_between(line, "{--", "--}")                    # tracked deletions are not current text
        line = _strip_between(line, "{>>", "<<}")                    # inline comments
        line = line.replace("{++", "").replace("++}", "")
        # a word with furigana, ｜誹謗《ひぼう》: pandoc and markitdown drop the whole word, so it is left
        # out of the comparison (the primary keeps it; it is not a loss in the primary)
        line = re.sub(r"｜[^《｜]{0,20}《[^》]{0,40}》", "", line)
        # pandoc's HTML footnotes inside tables: the note number link and the back link
        # bounded and kept inside one tag: an unbounded .*? here could backtrack badly on a crafted line
        line = re.sub(r'<a [^<>]{0,300}class="footnote-(?:ref|back)"[^<>]{0,300}>'
                      r'(?:[^<]|<sup>[^<]{0,20}</sup>){0,200}</a>', "", line)
        line = re.sub(r"\[\^[^\]\n]{1,40}\]", "", line)              # footnote marks, in table cells too
        line = re.sub(r"\[[↑↩︎]{1,4}\]\(#[^)]{0,200}\)", "", line)   # markitdown's footnote back links
        line = re.sub(r"^>\s*\*\*Text box:\*\*\s*", "", line)      # text box content is content
        line = re.sub(r"\[\[\d{1,6}\]\]\(#[^)]{0,200}\)", "", line)  # markitdown footnote marks
        line = re.sub(r"\$\$?[^$]{1,400}\$\$?", "", line)           # equations are compared separately
        line = re.sub(r"\[figure:[^\]]{0,500}\]", "", line)          # office2md figure markers after a caption
        line = re.sub(r"\[([^\]]{0,3000})\]\{[^}]{0,200}\}", r"\1", line)  # pandoc spans [text]{.underline}
        line = re.sub(r"\{(?:width|height|#|\.)[^}]{0,200}\}?", "", line)  # pandoc attributes, also cut at line end
        line = re.sub(r"\s{0,20}\{[^}]{0,200}\}\s{0,20}$", "", line)  # pandoc heading attributes {.title}
        line = re.sub(r"<br\s{0,5}/?>", " ", line)
        line = re.sub(r"<((?:https?|mailto|ftp):[^<>\s]{1,500})>", r"\1", line)   # autolinks <https://...>
        line = HTML_TAG.sub(" ", line)                             # pandoc HTML tables: tags, not text
        line = html.unescape(line)                                 # &lt;Insert Name&gt; is the text <Insert Name>
        line = re.sub(r"\^([^^\s]{1,20})\^", r"\1", line)          # pandoc superscript ^注１）^
        line = re.sub(r"(?<!~)~([^~\s]{1,20})~(?!~)", r"\1", line)  # and subscript ~(2)~, not ~~struck~~
        # links and pictures read as their text, in table cells too (markitdown writes a framed
        # box as a one-cell table, links included)
        line = re.sub(r"!\[[^\]]{0,3000}\]\([^)]{0,3000}\)", "", line)
        line = re.sub(r"\[([^\]]{0,500})\]\([^)]{0,2000}\)", r"\1", line)      # long SharePoint URLs too
        if line.startswith("|"):
            cells = [HTML_TAG.sub(" ", c).strip() for c in line.strip("|").split("|")]
            for c in cells:
                for s in split_sentences(c):
                    out.append((s, raw, section))
            continue
        if re.match(r"^#{1,6}\s", line):
            head = "".join(split_sentences(re.sub(r"^#{1,6}\s*", "", line)))
            if head and (known is None or head in known):
                section = head
        line = re.sub(r"^#{1,6}\s*", "", line)
        line = re.sub(r"^(\s{0,8}[-*+]\s{1,8}|\s{0,8}\d{1,9}\.\s{1,8}|>\s{0,8}){1,10}", "", line)
        line = HTML_TAG.sub("", line)                               # a <placeholder> in the text stays
        for s in split_sentences(line):
            out.append((s, raw, section))
    return out


def sentences(md: str) -> dict[str, str]:
    """Normalised sentence -> first original line it came from."""
    out: dict[str, str] = {}
    for s, raw, _ in sentence_entries(md):
        out.setdefault(s, raw)
    return out


_CIRCLED = re.compile("[①-⑳Ⅰ-ⅿ]")


def _head_key(h: str) -> str:
    """A heading compared without its number, spacing, quote style, digit width or emphasis
    marks (as sentences are, but kept whole however short: "概要" is a heading)."""
    h = radicals(NUM_PREFIX.sub("", h.strip())).translate(_WIDE_DIGITS)
    h = re.sub(r"[“”\"‘’'*_`\\]", "", h)
    return re.sub(r"[\s　]+", "", h) or h.strip()


def _toc_entry(line: str, head_keys: set[str]) -> bool:
    """A table-of-contents line without links: a heading of the primary and its page number."""
    text = re.sub(r"[\s*]{0,5}(?:[A-Z]\.)?[0-9]{1,5}(?:[.-][0-9]{1,4}){0,3}[\s*]{0,5}$", "", line.strip()[:300])
    text = re.sub(r"^[-*>\s]{0,6}", "", text)
    return text != line.strip() and "".join(split_sentences(NUM_PREFIX.sub("", text))) in head_keys


_WIDE_DIGITS = str.maketrans("０１２３４５６７８９", "0123456789")


def split_sentences(text: str) -> list[str]:
    text = radicals(text)               # 地⽅ (radical) and 地方 are the same text to a reader
    text = text.translate(_WIDE_DIGITS)  # 図１-1 and 図1-1: converters keep or fold the width
    text = text.replace("~~", "").replace("**", "").replace("__", "").replace("*", "").replace("`", "").replace("\\", "")
    # leaders: Word's "……" comes out of pandoc as "......", which would also split the sentence
    text = re.sub(r"[…‥]+|\.{2,}", "", text)
    text = re.sub(r"[“”\"‘’']", "", text)          # pandoc turns curly quotes into straight ones
    # list numbers and bullets inside a cell: "1. by a party" in one converter, "(i) by a party" or
    # "* + item" in another; both sides drop them
    text = re.sub(r"(?:(?<=\s)|^)(?:[0-9]{1,3}[.)]|[(（][0-9a-zA-Zⅰ-ⅿ]{1,4}[)）]|[+*•-]|[ぁ-んァ-ヶ][)）.．]?)(?=\s)",
                  " ", text)                           # ア / い labels too, as NUM_PREFIX drops them at a line start
    pieces = re.split(r"(?<=[。．.!?！？])\s*", text)
    out = []
    for p in pieces:
        p = NUM_PREFIX.sub("", p.strip())
        p = re.sub(r"[\s　]+", "", p)
        p = re.sub(r"[「」『』（）()\[\]【】、，,.。:：;；・\-ー–—‐‑−－•<>]", "", p)   # • bullets in cells
        p = re.sub(r"_{2,}|[☐☑☒✓✔]", "", p)           # fill-in blanks and checkboxes ([ ] in a reference)
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
    """Rows with content per pipe table, in order. Rows, not cells: converters write merged cells
    differently (repeated, blank or dropped), but a row lost or split is a real difference; an
    all-blank row is the placeholder header markitdown adds and is not counted, and so is a row
    whose only content repeats the cell above (a label merged down a blank form section).
    Laid out the way office2md writes them, so the counts compare: a one-column table (a framed
    box, which office2md writes as a quote) is not counted, nor a first row holding one text
    across every column (a title, which office2md writes above the table), and two tables with the
    same columns and only blank lines between are one table broken across a page, as office2md
    joins them."""
    counts, cur, in_table, prev, width, first = [], 0, False, [], 0, True
    cols, last_cols, gap_text = 0, -1, True
    for line in md.splitlines() + [""]:
        s = line.strip()
        if s.startswith("|"):
            in_table = True
            if re.match(r"^\|?\s*:?-{2,}", s):
                continue
            cells = [c.strip() for c in s.strip("|").split("|")]
            cols = cols or len(cells)
            width = max(width, sum(1 for c in cells if c))
            # markitdown writes the spanning title once and leaves the merged cells blank
            title = first and len(cells) > 1 and cells[0] and (
                len(set(cells)) == 1 or (len(cells) >= 3 and not any(cells[1:])))
            first = first and not any(cells)          # markitdown's blank header comes before the title
            new = [c for i, c in enumerate(cells) if c and (i >= len(prev) or c != prev[i])]
            if new and not title:
                cur += 1
            prev = cells
        elif in_table:
            if width > 1:
                if not gap_text and cols == last_cols and counts:
                    counts[-1] += cur
                else:
                    counts.append(cur)
                last_cols = cols
            cur, in_table, prev, width, first, cols, gap_text = 0, False, [], 0, True, 0, False
        elif s:
            gap_text = True
    return counts


def html_tables(md: str) -> int:
    """HTML tables (pandoc writes merged-cell tables so) with more than one column; a one-column
    one is a framed box, which office2md writes as a quote, as tables() does not count."""
    n = 0
    for part in md.split("<table")[1:]:
        body = part.split("</table>", 1)[0]
        rows = body.split("<tr")[1:]
        if any(len(re.findall(r"<t[dh][\s>]", r)) > 1 for r in rows):
            n += 1
    return n


def headings(md: str) -> list[str]:
    # [ \t]+, not \s+: an empty heading line ("#") must not swallow the line after it
    # a heading pandoc marks .TOC-Heading belongs to the table of contents the primary collapses
    found = [re.sub(r"\s{0,20}\{[^{}]{0,200}\}\s{0,20}$", "", h[:1000])
             for h in re.findall(r"^#{1,6}[ \t]+(.+)$", md, re.M) if ".TOC-Heading" not in h]
    return [re.sub(r"\s+", " ", h.strip()) for h in found if h.strip()]


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
    head_keys = {_head_key(h) for h in headings(primary)}

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
        # bold: a phrase the reference emphasises, present exactly once in the primary and not yet marked.
        # Not a heading's text (the reference bolds a heading; the primary's one plain copy is then an
        # index line or a cross-reference), and a line wholly bold in the reference only onto a line
        # that is that phrase alone
        whole = {re.sub(r"^\W{0,6}|\W{0,6}$", "", m) for m in re.findall(r"(?m)^\s*\*\*([^*\n]{2,80})\*\*\s*$", text)}
        for phrase in dict.fromkeys(emphasis(text)):
            if len(phrase) < 2 or re.fullmatch(r"[\d\s.,%]+", phrase) or "`" in phrase or "$" in phrase:
                continue
            if _head_key(phrase) in head_keys or not NUM_PREFIX.sub("", phrase.strip()).strip():
                continue                     # a heading's text, or only a number label such as （８）
            hits = occurrences(phrase)
            if hits and phrase.strip() in whole and lines[hits[0]].strip().strip("*").strip() != phrase.strip():
                continue
            if len(hits) == 1 and lines[hits[0]].count(phrase) == 1 and f"**{phrase}**" not in lines[hits[0]] \
                    and "](" not in lines[hits[0]]:
                lines[hits[0]] = lines[hits[0]].replace(phrase, f"**{phrase}**")
                stats["bold"] += 1
    return "\n".join(lines), stats


# -- the review ---------------------------------------------------------------------------

def review(primary: str, refs: dict[str, str], src: Path, applied: dict[str, int] | None = None,
           checks: dict | None = None) -> tuple[str, int]:
    checks = checks or {}
    out = [f"# Cross-check of {src.name}", "", office2md.DATA_NOTICE, ""]
    if not refs:
        out.append("Status: NO REFERENCE. pandoc is not installed and markitdown is unavailable, so the "
                   "Markdown comes from office2md alone and nothing could be compared. Say so when handing it over.")
        out += [""] + checks.get("visual_lines", [])
        return "\n".join(out) + "\n", checks.get("total", 0) - checks.get("done", 0)
    status_at = len(out)          # the status block is inserted here once the counts are known
    out.append("Primary: office2md. References: " + ", ".join(refs) + ". Each item below is a place where a "
               "reference has something the primary lacks, or the two disagree. Decide per item: fix the "
               "primary Markdown, or leave it when the difference is intentional (collapsed table of contents, "
               "dropped page numbers, numbering that only the primary has).")
    out.append("")
    n_items = 0
    counts = {"missing": 0, "links": 0, "equations": 0, "bold": 0, "tables": 0, "headings": 0, "fewer": 0, "moved": 0}
    own = {f"# {src.name}", "## Notes", "## Comments", "## Tracked changes"}        # lines office2md adds itself
    # drop those lines before sentence extraction: "# 仕様書.docx" would otherwise yield the sentence
    # "仕様書" and shadow the document's real title heading
    p_text = "\n".join(l for l in primary.splitlines() if l.strip() not in own)
    p_sent = sentences(p_text)
    p_entries = sentence_entries(p_text)
    p_blob = "\n".join(s for s, _, _ in p_entries)
    # the same sentences run together: a reference sentence the primary splits over two lines
    # (a text box, a wrapped cell) is found across the join
    p_flat = "".join(s for s, _, _ in p_entries)
    p_flat_bare = _CIRCLED.sub("", p_flat)
    p_count = Counter(s for s, _, _ in p_entries)
    p_section = {s: sec for s, _, sec in p_entries}
    # the change list and comment table office2md appends are not tables of the document
    body = re.split(r"(?m)^## (?:Tracked changes|Comments)$", primary, maxsplit=1)[0]
    p_links, p_emph, p_eq, p_tables = links(primary), emphasis(primary), equations(primary), tables(body)
    p_heads = [h for h in headings(primary) if h not in (src.name, "Notes", "Comments")]
    p_head_keys = {"".join(split_sentences(NUM_PREFIX.sub("", h))) for h in p_heads} - {""}
    acc = rc.accepted(checks.get("checked_lines", []))
    taken: list[tuple[str, str, str]] = []        # (id, reason, what) for items accepted with a reason

    def keep(kind: str, key: str, what: str) -> tuple[bool, str]:
        """(still open, id). An item accepted in the .checked file with a reason is set aside."""
        iid = rc.item_id(kind, key)
        if iid in acc:
            taken.append((iid, acc[iid], what))
            return False, iid
        return True, iid

    printed = checks.get("print_pages") or []
    if printed:
        found = [(k, raw, page) for k, raw, page in rc.print_missing(printed, p_text)
                 if keep("print", k, f"page {page}: {raw[:100]}")[0]]
        out.append(f"## Text Word prints that is not in the primary ({len(found)})")
        out.append("")
        if found:
            out.append("Pieces of the printed page (the document rendered by Word, running headers and page "
                       "numbers removed) that the Markdown does not contain. This check owes nothing to any "
                       "converter, so text all of them drop (shapes, fields, odd layouts) shows here. Open the "
                       "page, and add the text or accept the item with a reason.")
            out.append("")
            for k, raw, page in found[:150]:
                out.append(f"- `{rc.item_id('print', k)}` page {page}: {esc(raw)[:160]}")
            if len(found) > 150:
                out.append(f"- … {len(found) - 150} more")
            counts["printed"] = len(found)
        else:
            out.append("None.")
        out.append("")

    for name, text in refs.items():
        r_sent = sentences(text)
        # found when equal to a primary sentence or part of one: a table cell pandoc writes as
        # several <p> lines is one cell, joined, in the primary
        missing, unchecked = [], 0
        budget = SUBSTRING_BUDGET
        for s in r_sent:
            if s in p_sent:
                continue
            # the substring search scans the whole primary per sentence: bounded, so a huge or
            # crafted document cannot make the review run for hours; what is left is reported
            if budget <= 0:
                unchecked += 1
                missing.append((s, r_sent[s]))
                continue
            budget -= len(p_blob) + 2 * len(p_flat)
            # a circled number starts a line in the primary (dropped there as numbering) but sits
            # mid-sentence in a reference that ran a framed box into one cell
            if s not in p_blob and s not in p_flat and _CIRCLED.sub("", s) not in p_flat_bare:
                missing.append((s, r_sent[s]))
        if unchecked:
            counts["unchecked"] = counts.get("unchecked", 0) + unchecked
        # ignore reference lines that are just the TOC the primary collapsed on purpose
        # leaders and a page ("....A.3-4"), or a heading of the primary followed by its page number
        missing = [(k, m) for k, m in missing
                   if not re.search(r"(?:\.{3,40}|…{1,40})\s{0,5}(?:[A-Z]\.)?[\d.-]{1,12}[\s*]{0,5}$", m[-200:])
                   and not _toc_entry(m, p_head_keys)]
        open_items = []
        for k, m in missing:
            still, iid = keep("missing", k, m.strip()[:120])
            if still:
                open_items.append((iid, m))
        missing = open_items
        out.append(f"## Content in {name} not found in the primary ({len(missing)})")
        out.append("")
        if missing:
            out.append("Each line is as the reference wrote it. Search the primary for the words; if the content "
                       "is truly absent, add it where it belongs. If the difference is intentional, accept it with "
                       "a reason (see Done when).")
            out.append("")
            for iid, m in missing[:150]:
                out.append(f"- `{iid}` {esc(m.strip())[:200]}")
            if len(missing) > 150:
                out.append(f"- … {len(missing) - 150} more")
            n_items += len(missing)
            counts["missing"] += len(missing)
        else:
            out.append("None.")
        out.append("")

        # the same sentence fewer times in the primary: one of the repeats was lost
        r_entries = sentence_entries(text)
        r_count = Counter(s for s, _, _ in r_entries)
        fewer = [(r_sent[s], n, p_count[s], iid) for s, n in r_count.items()
                 if p_count[s] and p_count[s] < n and len(s) >= 6
                 for still, iid in [keep("fewer", s, r_sent[s].strip()[:120])] if still]
        if fewer:
            out += [f"## Repeated content found fewer times in the primary than in {name} ({len(fewer)})", "",
                    "A specification often repeats a sentence on purpose (one per case or per section). Find where "
                    "the reference has the other occurrence and add it to the primary.", ""]
            out += [f"- `{iid}` {esc(m.strip())[:200]} ({n} times in {name}, {have} in the primary)"
                    for m, n, have, iid in fewer[:80]]
            out.append("")
            n_items += len(fewer)
            counts["fewer"] += len(fewer)
        # a sentence present once in each, under a different heading: moved, or a heading lost
        r_section = {s: sec for s, _, sec in r_entries}
        p_section = {s: sec for s, _, sec in sentence_entries(p_text, known=set(r_section.values()))}
        moved = [(r_sent[s], r_section[s], p_section[s]) for s in r_count
                 if r_count[s] == 1 and p_count[s] == 1 and r_section[s] and p_section[s]
                 and r_section[s] != p_section[s]]
        if moved:
            out += [f"## Content under a different heading than in {name} ({len(moved)})", "",
                    "Usually a heading the two converters number or word differently; check the ones where the "
                    "condition or exception now belongs to another section.", ""]
            out += [f"- {esc(m.strip())[:160]} (in {name} under '{esc(a)[:40]}', in the primary under '{esc(b)[:40]}')"
                    for m, a, b in moved[:60]]
            out.append("")
            counts["moved"] += len(moved)

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
        new_links = {u: t for u, t in r_links.items() if u not in p_links and keep("link", u, u)[0]}
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
        if html_tables(text) and len(r_tables) + html_tables(text) == len(p_tables):
            # the reference wrote tables with merged cells as HTML: rows of those cannot be compared
            counts["tables_skipped"] = counts.get("tables_skipped", 0) + html_tables(text)
        elif r_tables != p_tables and keep("tables", f"{name}:{r_tables}:{p_tables}", f"table shapes against {name}")[0]:
            out.append(f"## Table shapes differ from {name} (`{rc.item_id('tables', f'{name}:{r_tables}:{p_tables}')}`)")
            out.append("")
            out.append(f"Primary: {len(p_tables)} table(s) with row counts {p_tables[:30]}; {name}: "
                       f"{len(r_tables)} with {r_tables[:30]}. A different count means a row lost, split or "
                       "merged, or a nested table flattened; check those tables by eye.")
            out.append("")
            n_items += 1
            counts["tables"] += 1

        r_heads = headings(text)
        strip = _head_key
        only_ref = [h for h in r_heads if strip(h) not in {strip(x) for x in p_heads}
                    and keep("heading", strip(h), h)[0]]
        only_pri = [h for h in p_heads if strip(h) not in {strip(x) for x in r_heads}]
        if only_ref or only_pri:
            out.append(f"## Headings that differ from {name}")
            out.append("")
            for h in only_ref[:40]:
                out.append(f"- `{rc.item_id('heading', strip(h))}` only in {name}: {esc(h)}")
            for h in only_pri[:40]:
                out.append(f"- only in primary: {esc(h)}")
            out.append("")
            n_items += len(only_ref)
            counts["headings"] += len(only_ref)

    lint = diagram_lint.lint_markdown(primary)
    if lint:
        out += ["## Diagram syntax problems (" + str(sum(len(e) for _, _, e in lint)) + ")", ""]
        for lang, line, errs in lint:
            out += [f"- {lang} block at line {line}: {esc(e)}" for e in errs]
        out.append("")
    src_state, md_digest, now_digest, md_name = checks.get("source", ("unknown", "", "", ""))
    out += source_lines(src_state, md_digest, now_digest, md_name)
    vis_total, vis_done, vis_stale = checks.get("total", 0), checks.get("done", 0), checks.get("stale", 0)
    if checks.get("visual_lines"):
        out += checks["visual_lines"] + [""]
    if taken:
        out += [f"## Accepted with a reason ({len(taken)})", ""]
        out += [f"- `{iid}` {esc(what)} — {esc(reason)}" for iid, reason, what in taken[:200]]
        out.append("")
    out.append("## Done when")
    out.append("")
    out.append(f"Rerun this script after each round of edits. Fix what is wrong in the Markdown; for a difference "
               f"that is right as it is, add a line `accept <id> <reason>` to {checks.get('checked_name', '<stem>.checked')} "
               "(the reason is required and is listed above on the next run). Stop when the status is READY.")

    # status block at the top: the one place to look to know whether this is finished
    lint_n = sum(len(e) for _, _, e in lint)
    content_ok = counts["missing"] + counts["fewer"] + counts["links"] + counts["headings"] == 0
    blocking = (counts["missing"] + counts["fewer"] + counts["links"] + counts["headings"] + lint_n
                + counts.get("printed", 0)
                + (vis_total - vis_done) + (1 if src_state == "differs" else 0))
    verdict = "READY" if blocking == 0 and counts["tables"] == 0 else "NEEDS REVIEW"
    status = [f"**Status: {verdict}**", ""]
    if applied:
        status.append("Applied automatically (safe, exact matches only): "
                      + ", ".join(f"{v} {k}" for k, v in applied.items() if v) + "." if any(applied.values())
                      else "Applied automatically: nothing matched.")
        status.append("")
    status += ["Checks:",
               status_line(src_state != "differs",
                           {"match": "Source: the Markdown was made from this version of the document.",
                            "unknown": "Source: the Markdown names no source version (made by an older office2md).",
                            "differs": "Source: the Markdown was made from another version of the document."}[src_state]),
               status_line(content_ok, f"Content against {', '.join(refs)}: {counts['missing']} sentence(s) only in a "
                           f"reference, {counts['fewer']} repeated fewer times, {counts['links']} link(s), "
                           f"{counts['headings']} heading difference(s)"
                           + (f"; {counts['unchecked']} of those sentences were not searched as part of a longer "
                              "sentence because the document is very large, so some may be present"
                              if counts.get("unchecked") else "") + "."),
               status_line((counts.get("printed", 0) == 0) if printed else None,
                           f"Printed text: {counts.get('printed', 0)} piece(s) of the page as Word prints it not in "
                           "the primary" if printed else "Printed text: not compared (no Word or LibreOffice to "
                           "print the document, or --no-render)."),
               status_line(counts["tables"] == 0, f"Table shapes: {counts['tables']} difference(s) in row counts"
                           + (f"; {counts['tables_skipped']} table(s) a reference wrote as HTML were not compared"
                              if counts.get("tables_skipped") else "") + "."),
               status_line(vis_done == vis_total, f"Visual checks: {vis_done} of {vis_total} confirmed"
                           + (f", {vis_stale} recorded for content that has since changed" if vis_stale else "")
                           + (f" (page render: {checks['page_pdf']})" if checks.get("page_pdf") else "") + "."),
               status_line(lint_n == 0, f"Diagram syntax: {lint_n} problem(s) (built-in checker)."),
               f"- [info] {counts['moved']} sentence(s) under a different heading, {counts['equations']} equation(s) "
               f"not applied, {counts['bold']} bold phrase(s): look at them, none blocks READY.",
               "",
               "READY means the checks above passed within their scope: what the references found is in the primary, "
               "and every figure and object was looked at. It does not prove the Markdown means the same as the "
               "document: content every converter drops (text inside pictures) is covered only by the visual checks.",
               ""]
    if checks.get("rounds_path"):
        rows = rc.record_round(checks["rounds_path"], {
            "status": verdict, "missing": counts["missing"], "printed": counts.get("printed", 0),
            "fewer": counts["fewer"], "tables": counts["tables"],
            "visual open": vis_total - vis_done, "accepted": len(taken)})
        status += rc.rounds_table(rows, ["status", "missing", "printed", "fewer", "tables", "visual open",
                                         "accepted"])
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


def run_office_script(script: str, image: str, env: dict[str, str]) -> str:
    """Office through COM with a time limit; see office_com.run_script."""
    return office_com.run_script(script, image, env)


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
} finally {
  $app.Quit()
  # release every COM reference, or Excel stays running hidden after Quit
  foreach ($o in @($ps, $ws, $wb, $app)) { if ($o) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($o) } }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
"""
        try:
            stdout = run_office_script(ps, "EXCEL", {"OFFICE_SRC": str(src.resolve()), "OFFICE_OUT": str(out.resolve())})
            done = {int(x.split()[1]) for x in stdout.split("\n") if x.startswith("OK ")}
            if done:
                return {sheet_names[i - 1]: f"{out.name}/sheet{i:02d}.pdf" for i in sorted(done)
                        if 0 < i <= len(sheet_names)}, "Excel"
        except (subprocess.TimeoutExpired, OSError):
            pass
    so = shutil.which("soffice") or (r"C:\Program Files\LibreOffice\program\soffice.exe"
                                      if Path(r"C:\Program Files\LibreOffice\program\soffice.exe").exists() else None)
    if so:
        try:
            subprocess.run([so, "--headless", "--convert-to", "pdf", "--outdir", str(out.resolve()), str(src.resolve())],
                           capture_output=True, text=True, timeout=300)
            pdf = out / (src.stem + ".pdf")
            if pdf.exists():
                return {name: f"{out.name}/{pdf.name} (all sheets in tab order)" for name in sheet_names}, "LibreOffice"
        except (subprocess.TimeoutExpired, OSError):
            pass
    return {}, ""


def ident(s: str) -> str:
    return s.replace("`", "'")


def status_line(ok: bool | None, text: str) -> str:
    """One line of the status block: ok, FAIL, or not run (a check that could not run is not a pass)."""
    return f"- [{'ok' if ok else 'not run' if ok is None else 'FAIL'}] {text}"


def visual_section(visuals: list[tuple[str, str, str]], lines: list[str], checked_name: str,
                   example: str) -> tuple[list[str], int, int]:
    """The visual-check list and (confirmed, stale) counts."""
    states = {vid: rc.check_state(vid, fp, lines) for vid, fp, _ in visuals}
    done = sum(1 for v in states.values() if v == "done")
    stale = sum(1 for v in states.values() if v == "stale")
    out = [f"## Visual checks ({done} of {len(visuals)} confirmed)", ""]
    if not visuals:
        return out + ["None: nothing was reconstructed or left unextracted."], 0, 0
    for vid, fp, what in sorted(visuals, key=lambda v: states[v[0]] == "done"):
        st = states[vid]
        note = " **(recorded for different content: look again)**" if st == "stale" else ""
        # names in `what` come from the file: escaped like any other document text
        out.append(f"- [{'x' if st == 'done' else ' '}] `{ident(rc.check_key(vid, fp))}`{note} — {esc(what)}")
    first = rc.check_key(visuals[0][0], visuals[0][1])
    out += ["", f"After checking one, add a line to {checked_name} that starts with its id exactly as shown, then "
            f"what you found, for example: `{ident(first)}  {example}`. The part after @ names the content you "
            "looked at: when the file changes, a changed item gets a new id and has to be looked at again."]
    return out, done, stale


def source_lines(state: str, md_digest: str, now: str, md_name: str) -> list[str]:
    if state != "differs":
        return []
    return ["## The Markdown was made from another version of the file", "",
            f"{md_name} records sha256 {md_digest}; the file is now {now}. Its text, diagrams and the checks below "
            "describe the old version. Regenerate with --fresh and carry your edits over, then rerun.", ""]


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
    src_state, md_digest, now_digest = rc.source_state(primary, src)

    sheet_pdf, method = {}, ""
    render_dir = out_dir / f"{stem}_sheets"
    names = [s["name"] for s in data["sheets"]]
    if not a.no_render:
        if not a.fresh and render_dir.exists() and rc.renders_current(render_dir, src):
            sheet_pdf = {n: f"{render_dir.name}/sheet{i:02d}.pdf" for i, n in enumerate(names, start=1)
                         if (render_dir / f"sheet{i:02d}.pdf").exists()}
            method = "existing renders of this version"
        else:
            # renders of another version must never be compared against: remove them first
            for old in [*render_dir.glob("sheet*.pdf"), render_dir / "source.sha256"] if render_dir.exists() else []:
                try:
                    old.unlink()
                except OSError:
                    pass
            sheet_pdf, method = render_sheets(src, render_dir, names)
            if sheet_pdf:
                rc.stamp_renders(render_dir, src)

    text = rc.xlsx_text_check(data, primary)
    missing = text["missing"]
    visuals = rc.xlsx_visuals(data, sheet_pdf)
    census_visuals, census_notes = rc.xlsx_census(src, data)
    visuals += census_visuals
    checked_path = out_dir / f"{stem}.checked"
    vis_lines, done, stale = visual_section(visuals, rc.read_checked(checked_path), checked_path.name,
                                            "matches the rendered sheet")
    lint = diagram_lint.lint_markdown(primary)
    lint_count = sum(len(e) for _, _, e in lint)
    rendered = []
    if a.render_diagrams and diagram_lint.renderers():
        rendered = diagram_lint.render(primary, render_dir / "diagrams")
    render_errors = [r for r in rendered if r[3]]
    edits = rc.diagram_edits(data, primary)
    ready = (not missing and done == len(visuals) and not lint_count and not render_errors
             and src_state != "differs")

    whole = f" Not found as a heading, so checked against the whole Markdown: {', '.join(text['whole_text'])}." \
        if text["whole_text"] else ""
    out = [f"# Cross-check of {src.name}", "", office2md.DATA_NOTICE, "",
           f"**Status: {'READY' if ready else 'NEEDS REVIEW'}**", "", "Checks:",
           status_line(src_state != "differs",
                       {"match": "Source: the Markdown was made from this version of the workbook.",
                        "unknown": "Source: the Markdown names no source version (made by an older office2md).",
                        "differs": "Source: the Markdown was made from another version of the workbook."}[src_state]),
           status_line(not missing, f"Text: {text['total']} text item(s), {len(missing)} missing. Each sheet's texts are "
                       "looked for in that sheet's section, as many times as they occur; "
                       f"{text['short_checked']} short value(s) (1, 2, ○) are looked for on the line of their row label."
                       + whole),
           status_line(done == len(visuals), f"Visual checks: {done} of {len(visuals)} confirmed"
                       + (f", {stale} recorded for content that has since changed" if stale else "")
                       + f" (record them in {checked_path.name})."),
           status_line(not lint_count, f"Diagram syntax: {len(diagram_lint.blocks(primary))} mermaid/PlantUML "
                       f"block(s), {lint_count} problem(s) (built-in checker)."),
           status_line((not render_errors) if rendered else None,
                       f"Diagram rendering with the installed tool: {len(rendered) - len(render_errors)} ok, "
                       f"{len(render_errors)} failed." if rendered else
                       "Diagram rendering with a real renderer: not run (--render-diagrams and PlantUML or mermaid-cli)."),
           status_line(True if sheet_pdf else None,
                       (f"Rendered sheets ({method}): " + "; ".join(f"{k} → {v}" for k, v in sheet_pdf.items()))
                       if sheet_pdf else "Rendered sheets: none (no Excel or LibreOffice, or --no-render); visual "
                       "checks need another way to see the sheets."),
           "",
           "READY means the checks above passed within their scope. It does not prove the Markdown means the same as "
           "the workbook: meaning carried by colour or position is covered only where a visual check asks for it, and "
           "a value moved to another place in the same sheet section passes the text check. Restructure the Markdown "
           "freely (rewrite a diagram, merge blocks into one table); the text check keeps that from dropping anything.",
           ""]
    out += source_lines(src_state, md_digest, now_digest, primary_path.name)
    if lint or render_errors:
        out += [f"## Diagram problems ({lint_count + len(render_errors)})", "",
                "A diagram that does not parse is drawn wrongly or not at all. Fix the block at the line given.", ""]
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
        out += ["Each line is where the text sits in the workbook. Put it back in the sheet's section of the Markdown; "
                "keep a short value on the same line as its row label.", ""]
        for w, t, detail in missing[:200]:
            out.append(f"- {esc(w)}: {esc(t)[:200]}" + (f" ({esc(detail)})" if detail else ""))
        if len(missing) > 200:
            out.append(f"- … {len(missing) - 200} more")
    else:
        out.append("None.")
    out += [""] + vis_lines
    if edits:
        out += ["", f"## Diagram edits since the reconstruction ({len(edits)})", "",
                "The Markdown's sequence diagram differs from what was read from the shapes. Fine when you changed it "
                "after looking at the drawing; otherwise compare again.", ""]
        out += [f"- {esc(e)}" for e in edits[:60]]
    notes = []
    for s in data["sheets"]:
        d = s.get("diagram")
        if d and d["unexplained"]:
            notes.append(f"- {esc(s['name'])}: {len(d['unexplained'])} line(s) the reconstruction could not attach to "
                         "shapes or lifelines; they are listed in the sheet's Diagram section.")
        if s.get("uncached_formulas"):
            notes.append(f"- {esc(s['name'])}: {len(s['uncached_formulas'])} formula(s) without a cached value; their "
                         "results are unknown until the workbook is recalculated.")
        if s.get("struck"):
            notes.append(f"- {esc(s['name'])}: {len(s['struck'])} struck-through cell(s), written as ~~text~~; keep the "
                         "marking, it usually means deleted.")
    notes += [f"- {esc(n)}" for n in census_notes]
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
          f"({'READY' if ready else f'{len(missing)} missing, {len(visuals) - done} visual check(s) open'})")
    return 0


def render_docx(src: Path, out: Path) -> str:
    """The document as one PDF (Word through COM, else LibreOffice). Returns the PDF path or ''."""
    out.mkdir(parents=True, exist_ok=True)
    pdf = out / "document.pdf"
    if sys.platform == "win32":
        ps = """
$ErrorActionPreference = 'Stop'
$app = New-Object -ComObject Word.Application
try {
  $app.Visible = $false; $app.DisplayAlerts = 0; $app.AutomationSecurity = 3
  # read-only, no conversion prompt, no repair, no encoding dialog; a placeholder password makes a
  # protected document fail at once instead of waiting on a password prompt nobody can see.
  # Every argument is given: COM from PowerShell cannot skip one with [Type]::Missing.
  $doc = $app.Documents.Open($env:OFFICE_SRC, $false, $true, $false, '_', '_', $false, '_', '_', 0, 65001, $false, $false, 0, $true)
  $doc.ExportAsFixedFormat($env:OFFICE_OUT, 17)
  $doc.Close(0)
  Write-Output 'OK'
} finally {
  $app.Quit()
  foreach ($o in @($doc, $app)) { if ($o) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($o) } }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
"""
        try:
            run_office_script(ps, "WINWORD", {"OFFICE_SRC": str(src.resolve()), "OFFICE_OUT": str(pdf.resolve())})
            if pdf.exists():
                return str(pdf)
        except (subprocess.TimeoutExpired, OSError):
            pass
    so = shutil.which("soffice") or (r"C:\Program Files\LibreOffice\program\soffice.exe"
                                      if Path(r"C:\Program Files\LibreOffice\program\soffice.exe").exists() else None)
    if so:
        try:
            subprocess.run([so, "--headless", "--convert-to", "pdf", "--outdir", str(out.resolve()), str(src.resolve())],
                           capture_output=True, text=True, timeout=300)
            made = out / (src.stem + ".pdf")
            if made.exists():
                made.replace(pdf)
                return str(pdf)
        except (subprocess.TimeoutExpired, OSError):
            pass
    return ""


def word_main(src: Path, a) -> int:
    out_dir = a.out_dir or src.parent
    out_dir.mkdir(parents=True, exist_ok=True)
    stem = src.stem
    primary_path = out_dir / f"{stem}.md"
    fresh = a.fresh or not primary_path.exists()
    primary, data = run_office2md(src, primary_path, (out_dir / f"{stem}_figures") if fresh else None, out_dir,
                                  write=fresh)
    if not fresh:
        # a rerun checks the Markdown as it is now, edits included; --fresh starts over
        primary = primary_path.read_text(encoding="utf-8")
        print(f"[md_crosscheck] reusing existing {primary_path.name} (pass --fresh to regenerate from the .docx)")
    src_state, md_digest, now_digest = rc.source_state(primary, src)

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

    # the document as Word prints it: a reference independent of every converter, and the pages
    # an agent compares with the Markdown by eye
    page_pdf = ""
    render_dir = out_dir / f"{stem}_pages"
    if not a.no_render:
        if not a.fresh and rc.renders_current(render_dir, src) and (render_dir / "document.pdf").exists():
            page_pdf = f"{render_dir.name}/document.pdf"
        else:
            for old in (render_dir / "document.pdf", render_dir / "source.sha256"):
                if old.exists():
                    old.unlink()
            if render_docx(src, render_dir):
                rc.stamp_renders(render_dir, src)
                page_pdf = f"{render_dir.name}/document.pdf"
    visuals = rc.docx_visuals(data, page_pdf) + rc.docx_census(src, data)
    pages: list[str] = []
    if page_pdf:
        pages = rc.printed_pages(out_dir / page_pdf)
        if pages:
            spans = rc.page_spans(primary, pages)
            visuals += rc.page_checks(spans, page_pdf, primary_path.name, a.sample_pages)
    checked_path = out_dir / f"{stem}.checked"
    checked_lines = rc.read_checked(checked_path)
    vis_lines, done, stale = visual_section(visuals, checked_lines, checked_path.name, "matches the page")
    checks = {"visual_lines": vis_lines, "done": done, "total": len(visuals), "stale": stale,
              "checked_name": checked_path.name, "page_pdf": page_pdf, "checked_lines": checked_lines,
              "print_pages": pages, "rounds_path": out_dir / f"{stem}.rounds.jsonl",
              "source": (src_state, md_digest, now_digest, primary_path.name)}
    text, n = review(primary, refs, src, applied, checks)
    review_path = out_dir / f"{stem}.review.md"
    review_path.write_text(text, encoding="utf-8")
    print(f"[md_crosscheck] primary:   {primary_path}")
    for name in refs:
        print(f"[md_crosscheck] reference: {out_dir / f'{stem}.{name}.md'}")
    if not refs:
        print("[md_crosscheck] no reference converter found (pandoc / markitdown); primary written, nothing compared")
    print(f"[md_crosscheck] review:    {review_path}  ({n} item(s) to look at)")
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
    ap.add_argument("--sample-pages", type=int, default=5, metavar="N",
                    help="docx: pages spread through the document to compare by eye (plus the worst-matching ones)")
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
    return word_main(src, a)


if __name__ == "__main__":
    sys.exit(main())
