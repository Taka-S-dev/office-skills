#!/usr/bin/env python3
"""Checks shared by md_crosscheck's Word and Excel reviews.

Kept apart from md_crosscheck so each check can be read and tested on its own:

- visual checks carry a fingerprint of what was looked at; a confirmation in <stem>.checked
  holds only for that exact content, so a changed diagram has to be looked at again
- the Markdown's _Source line names the file version it was made from; a reused Markdown made
  from another version is reported
- the Excel text check works per sheet section and counts occurrences, and checks short values
  (1, 2, ○, ×) next to their row label instead of skipping them
- a census of the package lists what is in the file independently of the extraction, so an
  element the reader does not understand is reported instead of silently missing
- edits to a reconstructed sequence diagram (a reversed or relabelled arrow) are listed
"""

from __future__ import annotations

import hashlib
import json
import re
import zipfile
from collections import Counter
from pathlib import Path

import office2md

DRAWING_TAGS = ("sp", "cxnSp", "pic", "graphicFrame", "grpSp")


# -- fingerprints and the .checked file ---------------------------------------------------------

def fingerprint(obj) -> str:
    """8 hex digits naming the content of one visual check."""
    def strip(o):
        if isinstance(o, dict):
            return {k: strip(v) for k, v in o.items() if k not in ("blob", "saved", "metrics", "box")}
        if isinstance(o, (list, tuple)):
            return [strip(v) for v in o]
        return o
    raw = json.dumps(strip(obj), sort_keys=True, ensure_ascii=False, default=str)
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()[:8]


def check_key(vid: str, fp: str) -> str:
    return f"{vid}@{fp}"


def read_checked(path: Path) -> list[str]:
    try:
        return [ln.strip() for ln in path.read_text(encoding="utf-8").splitlines() if ln.strip()]
    except OSError:
        return []


def _starts_with_token(line: str, token: str) -> bool:
    """line begins with token followed by the end, a space or a colon (exact id, not a prefix)."""
    if not line.startswith(token):
        return False
    rest = line[len(token):]
    return rest == "" or rest[0].isspace() or rest[0] in ":："


def check_state(vid: str, fp: str, lines: list[str]) -> str:
    """'done' when a line records this id with this fingerprint, 'stale' when the id is recorded
    only for other content (or without a fingerprint), else 'open'."""
    key = check_key(vid, fp)
    if any(_starts_with_token(ln, key) for ln in lines):
        return "done"
    if any(_starts_with_token(ln, vid) or (ln.startswith(vid + "@") and _starts_with_token(ln, ln.split()[0]))
           for ln in lines):
        return "stale"
    return "open"


# -- which version of the file a Markdown was made from ----------------------------------------------

def source_state(md: str, src: Path) -> tuple[str, str, str]:
    """('match' | 'differs' | 'unknown', digest in the Markdown, digest of the file now)."""
    now = office2md.file_digest(src)
    for line in md.splitlines()[:12]:
        m = office2md.SOURCE_LINE.match(line.strip())
        if m:
            return ("match" if m.group(1) == now else "differs"), m.group(1), now
    return "unknown", "", now


# -- renders kept next to the Markdown ---------------------------------------------------------------

def renders_current(render_dir: Path, src: Path, stamp: str = "source.sha256") -> bool:
    """Whether what render_dir holds was made from this version of src (its stamp file)."""
    try:
        return (render_dir / stamp).read_text(encoding="utf-8").strip() == office2md.file_digest(src)
    except OSError:
        return False


def stamp_renders(render_dir: Path, src: Path, stamp: str = "source.sha256") -> None:
    try:
        render_dir.mkdir(parents=True, exist_ok=True)
        (render_dir / stamp).write_text(office2md.file_digest(src) + "\n", encoding="utf-8")
    except OSError:
        pass


# -- Excel: text per sheet, with counts and short values ---------------------------------------------

def norm(s) -> str:
    out = str(s).replace("<br>", "")
    for ch in " \t\r\n　\\*|/`~":
        out = out.replace(ch, "")
    return out


def _clean_name(name: str) -> str:
    return "".join(" " if ord(ch) < 32 or ord(ch) == 127 else ch for ch in name)


def sheet_sections(md: str, names: list[str], heading: str = "## Sheet: ") -> dict[str, str | None]:
    """Sheet name -> the Markdown under its heading, up to the next sheet heading. None when the
    heading is not found (the Markdown was restructured); the caller then checks the whole text."""
    lines = md.splitlines()
    starts = [i for i, ln in enumerate(lines) if ln.startswith(("## Sheet: ", "## Chart sheet: "))]
    out: dict[str, str | None] = {}
    for name in names:
        head = heading + office2md.md_escape(_clean_name(name))
        # exact, or followed by the state note ("  _(hidden)_"): a bare prefix match would send
        # sheet "CUI Overlay" to the section of "CUI Overlay Overview"
        at = next((i for i in starts if lines[i] == head or lines[i].startswith(head + "  _(")), None)
        if at is None:
            out[name] = None
            continue
        end = next((i for i in starts if i > at), len(lines))
        out[name] = "\n".join(lines[at:end])
    return out


def is_short(n: str) -> bool:
    return len(n) < 2 or re.fullmatch(r"-?[0-9]{1,2}", n) is not None


def _token_re(value: str):
    return re.compile(r"(?<![0-9A-Za-z.])" + re.escape(value) + r"(?![0-9A-Za-z.])")


def _addr(a: str) -> tuple[int, int] | None:
    m = re.match(r"([A-Z]+)([0-9]+)$", a)
    if not m:
        return None
    return int(m.group(2)), _col(m.group(1))


def _col(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def _letters(n: int) -> str:
    out = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        out = chr(65 + r) + out
    return out


def sheet_items(s: dict) -> list[tuple[str, str, str]]:
    """(where, text, cell address or '') for every text of one sheet that the Markdown must carry."""
    name = s["name"]
    items: list[tuple[str, str, str]] = []
    for c in s.get("cells", []):
        items.append((f"{name}!{c['address']}", str(c["display"]), c["address"]))
    for d in s.get("drawings", []):
        if d.get("text") and d["kind"] != "smartart":
            items.append((f"{name} {d['kind']} '{d['name']}' at {d['anchor']}", d["text"], ""))
    for sa in s.get("smartart", []):
        # node by node: a SmartArt rewritten as a list or a diagram still passes
        for nd in sa.get("nodes", []):
            items.append((f"{name} SmartArt '{sa['name']}' at {sa['from']}", nd["text"], ""))
    for ch in s.get("charts", []):
        if ch.get("title"):
            items.append((f"{name} chart '{ch['name']}' title", ch["title"], ""))
        for se in ch.get("series", []):
            if se.get("name"):
                items.append((f"{name} chart '{ch['name']}' series", se["name"], ""))
    for v in s.get("validations", []):
        if v["rule"].startswith("list: "):
            for choice in v["rule"][6:].split(" (from ")[0].split(", "):
                items.append((f"{name} input rule {v['range']}", choice, ""))
    for c in s.get("comments", []):
        items.append((f"{name}!{c['cell']} comment", c["text"] or "", ""))
    for hf in s.get("header_footer", []):
        items.append((f"{name} printed {hf.split(':', 1)[0]}", hf.split(":", 1)[-1].strip(), ""))
    return items


def xlsx_text_check(data: dict, md: str) -> dict:
    """Every text of each sheet must appear in that sheet's section of the Markdown, as many times
    as in the sheet. Short values (1, 2, ○) cannot be found that way; each is checked on the same
    line as its row label (the nearest text cell to its left with no gap), or with its address.

    Returns {"missing": [(where, text, detail)], "total": n, "short_checked": n,
             "whole_text": [sheet names checked against the whole Markdown]}."""
    sheets = data.get("sheets", []) + [{"name": cs["name"], "charts": cs["charts"]} for cs in data.get("chart_sheets", [])]
    names = [s["name"] for s in sheets]
    sections = sheet_sections(md, names)
    chart_sections = sheet_sections(md, [cs["name"] for cs in data.get("chart_sheets", [])], "## Chart sheet: ")
    missing: list[tuple[str, str, str]] = []
    total = short_checked = 0
    whole: list[str] = []
    for s in sheets:
        sec = sections.get(s["name"]) or chart_sections.get(s["name"])
        if sec is None:
            whole.append(s["name"])
            sec = md
        sec_norm = norm(sec)
        sec_lines = sec.splitlines()
        items = sheet_items(s)
        groups: dict[str, list[tuple[str, str]]] = {}
        shorts: list[tuple[str, str, str]] = []
        for where, text, addr in items:
            n = norm(text)
            if not n:
                continue
            total += 1
            if is_short(n):
                if addr:
                    shorts.append((where, text, addr))
                continue
            groups.setdefault(n, []).append((where, text))
        for n, hits in groups.items():
            have = sec_norm.count(n)
            if have >= len(hits):
                continue
            if have == 0:
                missing += [(w, t, "") for w, t in hits]
            else:
                where = ", ".join(w for w, _ in hits[:6]) + (" …" if len(hits) > 6 else "")
                missing.append((where, hits[0][1], f"{len(hits)} times in the sheet, {have} in its section"))
        if shorts:
            by_cell = {c["address"]: str(c["display"]) for c in s.get("cells", [])}
            for where, text, addr in shorts:
                rc = _addr(addr)
                if rc is None:
                    continue
                label = ""
                r, c = rc
                cc = c - 1
                while cc >= 1:
                    v = by_cell.get(f"{_letters(cc)}{r}")
                    if v is None:
                        break
                    vn = norm(v)
                    if not is_short(vn) and not re.fullmatch(r"[-+0-9.,%:]+", vn):
                        label = vn
                        break
                    cc -= 1
                tok = _token_re(norm(text))
                # a cell listed by address ("- B3: 1", one-column blocks) and a grid row ("| 3 | … |")
                listed = [ln.split(":", 1)[1] for ln in sec_lines
                          if re.match(r"^\s*- " + re.escape(addr) + r"( \(h\))?:", ln)]
                row_cells = [ln.split("|", 2)[2] for ln in sec_lines
                             if re.match(r"^\| " + str(r) + r"( \(h\))? \|", ln) and ln.count("|") >= 3]
                if label:
                    short_checked += 1
                    # the grid's row-number cell is not a value: "| 2 | ボルト | 3 |" does not hold a 2
                    if not (any(label in norm(ln) and tok.search(re.sub(r"^\| [0-9]+( \(h\))? \|", "|", ln))
                                for ln in sec_lines)
                            or any(tok.search(x) for x in listed)):
                        missing.append((where, text, f"not found on a line with its row label '{label[:30]}'"))
                elif listed:
                    short_checked += 1
                    if not any(tok.search(x) for x in listed):
                        missing.append((where, text, "not found after its address"))
                elif row_cells:
                    short_checked += 1
                    if not any(tok.search(x) for x in row_cells):
                        missing.append((where, text, f"not found in row {r} of the sheet's table"))
    return {"missing": missing, "total": total, "short_checked": short_checked, "whole_text": whole}


# -- Excel: visual checks --------------------------------------------------------------------------

def xlsx_visuals(data: dict, sheet_pdf: dict[str, str]) -> list[tuple[str, str, str]]:
    """(id, fingerprint, what to do) for everything reconstructed or not extracted."""
    items = []
    for s in data["sheets"]:
        name = s["name"]
        pdf = sheet_pdf.get(name, "")
        where = f" in {pdf}" if pdf else ""
        d = s.get("diagram")
        if d and (d["sequence"] or d["edges"] or d["unexplained"] or d.get("bars")):
            what = []
            if d["sequence"]:
                what.append(f"{len(d['sequence']['messages'])} message(s) between "
                            f"{len(d['sequence']['participants'])} participant(s)")
            if d["edges"]:
                what.append(f"{len(d['edges'])} connection(s)")
            if d.get("bars"):
                what.append(f"{len(d['bars'])} period bar(s) placed by position")
            if d["unexplained"]:
                what.append(f"{len(d['unexplained'])} line(s) not interpreted")
            items.append((f"{name}#diagram", fingerprint(d),
                          f"Compare the Diagram section of sheet {name} ({'; '.join(what)}) with the drawing{where}: "
                          "participants, order, arrow directions, labels. Correct the Markdown, and describe any "
                          "line not interpreted."))
        for i, ch in enumerate(s.get("charts", []), start=1):
            items.append((f"{name}#chart{i}", fingerprint(ch),
                          f"Check chart '{ch['name']}' (title, series, values) against the drawn chart{where}."))
        for i, pc in enumerate(s.get("pictures", []), start=1):
            if pc.get("placeholder"):
                continue
            items.append((f"{name}#picture{i}", fingerprint(pc),
                          f"Picture '{pc['name']}' at {pc['from']}: its content is not extracted. Look at it{where} "
                          "and describe what it shows in the Markdown, or note that it carries no information."))
        for i, sa in enumerate(s.get("smartart", []), start=1):
            items.append((f"{name}#smartart{i}", fingerprint(sa),
                          f"SmartArt '{sa['name']}' at {sa['from']}: {len(sa['nodes'])} text node(s) were read in the "
                          f"diagram's data order. Compare order and nesting with the drawing{where}; the arrows and "
                          "layout are not kept, so add what the picture says that the list does not."))
        frames = [x for x in s.get("drawings", []) if x["kind"] == "frame"]
        for i, fr in enumerate(frames, start=1):
            items.append((f"{name}#object{i}", fingerprint(fr),
                          f"Object '{fr['name']}' at {fr['anchor']} is not a chart, SmartArt or picture the reader "
                          f"understands; its content is not extracted. Look at it{where} and describe it."))
    for cs in data.get("chart_sheets", []):
        for i, ch in enumerate(cs["charts"], start=1):
            items.append((f"{cs['name']}#chart{i}", fingerprint(ch),
                          f"Check the chart on chart sheet {cs['name']} (title, series, values) in Excel."))
    return items


# -- census: what is in the package, independently of the extraction ------------------------------------

def _count_drawing_elements(root) -> int:
    """Drawing elements in a drawing part, not counting the Fallback copy of an AlternateContent."""
    n = 0
    for el in root.iter():
        if not isinstance(el.tag, str) or el.tag.rsplit("}", 1)[-1] not in DRAWING_TAGS:
            continue
        anc = el.getparent()
        fallback = False
        while anc is not None:
            if isinstance(anc.tag, str) and anc.tag.endswith("}Fallback"):
                choice = anc.getprevious()
                fallback = choice is not None and any(
                    isinstance(x.tag, str) and x.tag.rsplit("}", 1)[-1] in DRAWING_TAGS for x in choice.iter())
                break
            anc = anc.getparent()
        if not fallback:
            n += 1
    return n


def xlsx_census(src: Path, data: dict) -> tuple[list[tuple[str, str, str]], list[str]]:
    """Visual checks for parts that are present but not (fully) read, and notes for the rest."""
    import xlsx_drawing as xd
    visuals: list[tuple[str, str, str]] = []
    notes: list[str] = []
    try:
        pkg = xd.Package(src)
    except (zipfile.BadZipFile, OSError):
        return visuals, notes
    try:
        names = pkg.names
        by_sheet = {s["name"]: s for s in data["sheets"]}
        for title, part in xd.sheet_parts(pkg):
            s = by_sheet.get(title)
            if s is None:
                continue
            found = 0
            for typ, tgt, external in pkg.rels(part).values():
                if typ == "drawing" and not external and tgt in names:
                    try:
                        found += _count_drawing_elements(pkg.xml(tgt))
                    except Exception:
                        pass
            read = s.get("shape_count", 0)
            if found > read:
                visuals.append((f"{title}#unread", f"{found}-{read}",
                                f"The sheet's drawing holds {found} element(s) and {read} were read (limit or nesting). "
                                "Compare the rendered sheet with the Markdown for shapes that are missing."))
        embedded = sorted(n for n in names if n.startswith("xl/embeddings/"))
        if embedded:
            visuals.append(("(workbook)#embedded", fingerprint(embedded),
                            f"{len(embedded)} embedded object(s) ({', '.join(Path(n).name for n in embedded[:8])}"
                            f"{' …' if len(embedded) > 8 else ''}): their content is not extracted. Open them in Excel "
                            "and describe what they contain, or note that they carry no information."))
        chartex = sorted(n for n in names if re.match(r"xl/charts/chartEx[0-9]*[.]xml$", n))
        if chartex:
            visuals.append(("(workbook)#chartex", fingerprint(chartex),
                            f"{len(chartex)} chart(s) of the newer kinds (waterfall, treemap, funnel …) whose data "
                            "is not read. Look at them in the rendered sheets and write their values down."))
        smart_parts = [n for n in names if re.match(r"xl/diagrams/data[0-9]*[.]xml$", n)]
        smart_read = sum(len(s.get("smartart", [])) for s in data["sheets"])
        if len(smart_parts) > smart_read:
            notes.append(f"{len(smart_parts)} SmartArt data part(s) in the file, {smart_read} placed on a sheet and read.")
        if any(n.startswith("xl/pivotTables/") for n in names):
            notes.append("Pivot tables: their values are in the cells, but their fields and filters are not read; "
                         "say which view a value comes from when it matters.")
        if any(n.startswith("xl/ctrlProps/") for n in names):
            notes.append("Form controls (check boxes, buttons, lists): labels drawn as legacy shapes are read; "
                         "their state (checked or not) is not.")
        if any(n.startswith(("xl/slicers/", "xl/timelines/")) for n in names):
            notes.append("Slicers or timelines filter what the sheet shows; the filter state is not read.")
    finally:
        pkg.close()
    return visuals, notes


# -- Word: visual checks and census -------------------------------------------------------------------

def docx_visuals(data: dict, page_pdf: str = "", md_dir: str = "") -> list[tuple[str, str, str]]:
    """md_dir: the Markdown's folder as seen from the review file, put before figure paths."""
    where = f" in {page_pdf}" if page_pdf else " (render the document: office-edit/scripts/pdf_tool.py topdf)"
    items = []
    n: dict[str, int] = {}
    for b in data.get("blocks", []):
        if b.get("kind") != "object":
            continue
        k = b.get("object") or ""
        n[k] = n.get(k, 0) + 1
        label = b.get("name") or k
        if k == "smartart":
            what = (f"SmartArt '{label}': {len(b.get('nodes', []))} text node(s) were read in data order. Compare "
                    f"order and nesting with the page{where}, and add what the picture says that the list does not.")
        elif k == "chart":
            what = f"Chart '{label}': check title, series and values against the page{where}."
        else:
            what = (f"Embedded object '{label}': its content is not extracted. Open it in Word and describe what "
                    "it contains, or note that it carries no information.")
        items.append((f"{k}{n[k]}", fingerprint(b), what))
    drawings = [b for b in data.get("blocks", []) if b.get("kind") == "drawing"]
    for i, b in enumerate(drawings, start=1):
        # a figure drawn with shapes has no picture in the file: only Word's rendering shows it
        shown = f" (rendered as {md_dir}{b['saved']})" if b.get("saved") else ""
        items.append((f"drawing{i}", fingerprint([b.get("text")]),
                      f"Drawing {i}, {b.get('text', 'drawing')}{shown}: compare it with the page{where}. The "
                      "Markdown must show it, as the rendered picture or, where none was made, as a description of "
                      "what it says (steps, arrows and their labels, what is grouped), placed where it sits."))
    for i, im in enumerate(data.get("images", []), start=1):
        if im.get("placeholder"):
            continue
        shown = (md_dir + im["saved"]) if im.get("saved") else (im.get("file") or im.get("name"))
        items.append((f"figure{i}", fingerprint(im),
                      f"Figure {shown}{(' — ' + im['alt']) if im.get('alt') else ' (no alt text)'}: its content is not "
                      "extracted. Open it and describe in the Markdown what it shows that the text does not "
                      "(a diagram, a table drawn as a picture, values), or note that it carries no information."))
    return items


def docx_census(src: Path, data: dict) -> list[tuple[str, str, str]]:
    """Objects present in the package that the extraction did not place in the text."""
    from lxml import etree
    try:
        z = zipfile.ZipFile(src)
    except (zipfile.BadZipFile, OSError):
        return []
    with z:
        names = set(z.namelist())
        try:
            parser = etree.XMLParser(resolve_entities=False, no_network=True, huge_tree=False)
            root = etree.fromstring(z.read("word/document.xml"), parser)
        except (KeyError, etree.XMLSyntaxError):
            return []
    objects = [b for b in data.get("blocks", []) if b.get("kind") == "object"]
    have = Counter(b.get("object") for b in objects)
    out = []
    expected = {"smartart": 0, "chart": 0, "embedded": 0}
    uris = {"http://schemas.openxmlformats.org/drawingml/2006/diagram": "smartart",
            "http://schemas.openxmlformats.org/drawingml/2006/chart": "chart"}
    for el in root.iter():
        if not isinstance(el.tag, str):
            continue
        local = el.tag.rsplit("}", 1)[-1]
        kind = uris.get(el.get("uri") or "") if local == "graphicData" else ("embedded" if local == "object" else None)
        if not kind:
            continue
        anc, in_fallback = el.getparent(), False
        while anc is not None:
            if isinstance(anc.tag, str) and anc.tag.endswith("}Fallback"):
                in_fallback = True
                break
            anc = anc.getparent()
        if not in_fallback:
            expected[kind] += 1
    for kind, n in expected.items():
        if n > have.get(kind, 0):
            out.append((f"unread-{kind}", f"{n}-{have.get(kind, 0)}",
                        f"The document body holds {n} {kind} object(s) and {have.get(kind, 0)} were read (inside a "
                        "header, a table cell nested too deep, or a form the reader does not know). Find them on the "
                        "page and add their content."))
    if any(n.startswith("word/embeddings/") for n in names) and not have.get("embedded"):
        out.append(("embedded-parts", "1", "The file carries embedded files (word/embeddings/) that no object in the "
                    "text points to; open the document and check whether they matter."))
    return out


# -- sequence diagrams edited after the reconstruction --------------------------------------------------

_ARROW = re.compile(r"^\s*([^\s:+-][^:]*?)\s*(-->>|->>|-->|->|--x|-x|--\)|-\))\s*[+-]?\s*([^:]+?)\s*:\s*(.*?)\s*$")


def _messages(mermaid: str) -> list[tuple[str, str, str]]:
    out = []
    for line in mermaid.splitlines():
        m = _ARROW.match(line)
        if m:
            out.append((m.group(1).strip(), m.group(3).strip(), m.group(4).strip()))
    return out


def diagram_edits(data: dict, md: str) -> list[str]:
    """Messages of each sheet's reconstructed sequence that differ in the Markdown's mermaid
    block of that sheet: reversed, relabelled, added or removed. Listed for the reviewer; an edit
    made on purpose after looking at the drawing is fine."""
    out = []
    sections = sheet_sections(md, [s["name"] for s in data["sheets"]])
    for s in data["sheets"]:
        d = s.get("diagram")
        if not d or not d.get("mermaid_sequence"):
            continue
        sec = sections.get(s["name"])
        if sec is None:
            continue
        blocks = re.findall(r"```mermaid[ \t]*\n(.*?)```", sec, re.S)
        seq = next((b for b in blocks if b.lstrip().startswith("sequenceDiagram")), None)
        if seq is None:
            out.append(f"{s['name']}: the sequence diagram block was removed from the sheet's section.")
            continue
        orig, now = _messages(d["mermaid_sequence"]), _messages(seq)
        if orig == now:
            continue
        rev = {(b, a, lab) for a, b, lab in orig}
        for i, msg in enumerate(now, start=1):
            if msg in orig:
                continue
            if msg in rev and msg not in orig:
                out.append(f"{s['name']}: message {i} '{msg[2]}' now goes {msg[0]} → {msg[1]}; the drawing had it the other way.")
            else:
                out.append(f"{s['name']}: message {i} {msg[0]} → {msg[1]} '{msg[2]}' is not in the reconstruction.")
        for msg in orig:
            if msg not in now and (msg[1], msg[0], msg[2]) not in now:
                out.append(f"{s['name']}: the reconstruction's message {msg[0]} → {msg[1]} '{msg[2]}' is no longer in the Markdown.")
    return out


# -- the printed document as a reference, and pages to compare by eye -----------------------------

def printed_pages(pdf: Path) -> list[str]:
    """Text of each page of the document as Word printed it, wrapped lines joined, with the lines
    that repeat on many pages (running headers and footers, page numbers) removed."""
    import pypdf
    try:
        reader = pypdf.PdfReader(str(pdf))
        raw = [(p.extract_text() or "") for p in reader.pages[:2000]]
    except Exception:
        return []
    lines_per_page = [[x.strip() for x in t.splitlines() if x.strip()] for t in raw]
    seen = Counter(x for lines in lines_per_page for x in set(lines))
    repeat = max(3, int(0.3 * len(raw)))
    pages = []
    for lines in lines_per_page:
        keep = [x for x in lines if seen[x] < repeat and not re.fullmatch(r"[-‐－ー\s]*[0-9０-９]{1,4}[-‐－ー\s]*"
                                                                            r"(/\s*[0-9]{1,4})?", x)]
        pages.append("\n".join(keep))         # not reflowed: a table row must not run into the next
    return pages


def page_spans(md: str, pages: list[str], split_sentences=None) -> list[tuple[int, int | None, int | None, float]]:
    """(page number, first md line, last md line, share of the page's text found in the md) for
    every page: where in the Markdown the page's text sits, by the printed pieces it holds."""
    norm_lines = [norm_text(x) for x in md.splitlines()]
    flat = "".join(norm_lines)
    out = []
    for n, text in enumerate(pages, start=1):
        ps = [k for k, _ in pieces(text, 8)]
        if not ps:
            out.append((n, None, None, 1.0))
            continue
        found = [k for k in ps if k in flat]
        hits = []
        for k in found[:40]:
            probe = k[:16]
            at = next((i for i, line in enumerate(norm_lines) if probe in line), None)
            if at is not None:
                hits.append(at)
        share = len(found) / len(ps)
        if hits:
            hits.sort()
            # the densest stretch of hits: a page's text sits together, while a phrase it shares with
            # the table of contents or a running title matches far away
            best = max(range(len(hits)), key=lambda i: sum(1 for h in hits if 0 <= h - hits[i] <= 60))
            near = [h for h in hits if 0 <= h - hits[best] <= 60]
            out.append((n, near[0] + 1, near[-1] + 1, share))
        else:
            out.append((n, None, None, share))
    return out


def page_checks(spans, pdf_name: str, md_name: str, sample: int = 5, worst: int = 5) -> list[tuple[str, str, str]]:
    """Visual checks for a spread of pages plus the pages whose text the Markdown matches least."""
    if not spans:
        return []
    total = len(spans)
    picks = set()
    if total <= sample:
        picks = {s[0] for s in spans}
    else:
        step = (total - 1) / (sample - 1)
        picks = {spans[round(i * step)][0] for i in range(sample)}
    ranked = sorted((s for s in spans if s[3] < 0.95), key=lambda s: s[3])
    picks |= {s[0] for s in ranked[:worst]}
    items = []
    for n, first, last, share in spans:
        if n not in picks:
            continue
        where = f"lines {first}-{last} of {md_name}" if first else f"(not located in {md_name})"
        low = f"; only {share:.0%} of its sentences were found in the Markdown" if share < 0.95 else ""
        items.append((f"page{n}", fingerprint([n, first, last, round(share, 2)]),
                      f"Open {pdf_name} at page {n} and compare it with {where}{low}: paragraph order, every table "
                      "(rows, columns, merged headings), text boxes and captions, numbering. Fix the Markdown where "
                      "it differs."))
    return items


# -- accepting a review item with a reason ---------------------------------------------------------

def item_id(kind: str, text: str) -> str:
    return f"{kind}:{hashlib.sha1(text.encode('utf-8')).hexdigest()[:8]}"


def accepted(lines: list[str]) -> dict[str, str]:
    """id -> reason for lines 'accept <id> <reason>'. A line without a reason accepts nothing."""
    out = {}
    for ln in lines:
        m = re.match(r"accept\s+(\S+)\s+(.{4,})$", ln)
        if m:
            out[m.group(1)] = m.group(2).strip()[:300]
    return out


# -- the history of runs ---------------------------------------------------------------------------

def record_round(path: Path, counts: dict) -> list[dict]:
    """Append this run's counts to <stem>.rounds.jsonl and return all runs, oldest first."""
    import datetime
    rows = []
    try:
        rows = [json.loads(x) for x in path.read_text(encoding="utf-8").splitlines() if x.strip()]
    except (OSError, ValueError):
        rows = []
    row = {"round": len(rows) + 1, "at": datetime.datetime.now().isoformat(timespec="seconds"), **counts}
    rows.append(row)
    try:
        path.write_text("".join(json.dumps(r, ensure_ascii=False) + "\n" for r in rows[-200:]), encoding="utf-8")
    except OSError:
        pass
    return rows


def rounds_table(rows: list[dict], keys: list[str]) -> list[str]:
    if len(rows) < 2:
        return []
    out = ["Runs so far (newest last): each run is one pass of checking and fixing.", "",
           "| round | " + " | ".join(keys) + " |", "|---|" + "---|" * len(keys)]
    for r in rows[-8:]:
        out.append(f"| {r['round']} | " + " | ".join(str(r.get(k, "")) for k in keys) + " |")
    return out + [""]


# -- comparing with the printed text by pieces, not sentences ---------------------------------------
# A PDF gives a table row as one line with its cells separated by spaces, and joins wrapped lines,
# so sentence boundaries do not match the Markdown's. Pieces between spaces (and sentence ends) are
# compared instead, against the whole Markdown with spacing and punctuation removed.

_DROP = set(" \t\r\n　|*#>`~[](){}「」『』（）【】、，,.。:：;；・-‐－ー–—_\\/\"'“”‘’!！?？")


def norm_text(s: str) -> str:
    # as printed: a footnote mark shows its number ("UI2"), a word with furigana shows the word,
    # a radical typed for an ideograph (地⽅) reads as the ideograph the Markdown has
    s = office2md.radicals(s)
    s = re.sub(r"\[\^e?([0-9]{1,6})\]", r"\1", s)
    s = re.sub(r"｜([^《｜]{0,20})《[^》]{0,40}》", r"\1", s)
    s = re.sub(r"\[\^[^\]\n]{1,40}\]|<[^<>]{1,300}>", "", s)
    return "".join(ch for ch in s if ch not in _DROP)


def pieces(text: str, min_len: int = 6) -> list[tuple[str, str]]:
    """(normalised piece, original piece) for the parts of a printed text between spaces."""
    out = []
    for line in text.splitlines():
        for part in re.split(r"\s{1,}|(?<=[。．])", line):
            n = norm_text(part)
            if len(n) >= min_len:
                out.append((n, part.strip()))
    return out


def print_missing(pages: list[str], md: str) -> list[tuple[str, str, int]]:
    """(normalised piece, original piece, page) for printed text the Markdown does not contain."""
    flat = norm_text(md)
    out, seen = [], set()
    for n, text in enumerate(pages, start=1):
        for key, raw in pieces(text):
            if key in seen or key in flat:
                continue
            seen.add(key)
            out.append((key, raw, n))
    return out


# -- the Markdown's own pictures, links and marks against the file --------------------------------

# formats a Markdown viewer (a browser underneath) does not show
UNVIEWABLE = (".emf", ".wmf", ".tif", ".tiff")
_LINK = re.compile(r"(!?)\[(?:[^\[\]\n]|\[[^\[\]\n]{0,300}\]){0,1000}\]\((<[^<>\n]{1,1000}>|[^()\s]{1,1000})(?:\s+\"[^\"\n]{0,300}\")?\)")


def local_links(md: str) -> list[tuple[bool, str]]:
    """(is a picture, target) for every link to a file next to the Markdown, in order."""
    out = []
    for m in _LINK.finditer(md):
        target = m.group(2)
        if target.startswith("<"):
            target = target[1:-1]
        if re.match(r"[a-zA-Z][a-zA-Z0-9+.-]{1,20}:|#", target):       # https:, mailto:, data:, #anchor
            continue
        pic = m.group(1) == "!"
        # [必須](全角) in body text is no link to a file: only a target that looks like a path counts
        if pic or "/" in target or re.search(r"\.[A-Za-z0-9]{1,5}$", target):
            out.append((pic, target))
    return out


def link_check(md: str, md_dir: Path) -> dict:
    """Pictures and file links of the Markdown: which point at nothing, which at a format viewers
    do not show."""
    from urllib.parse import unquote
    broken, unviewable = [], []
    for _, target in dict.fromkeys(local_links(md)):
        path = target.split("#", 1)[0]
        if not path:
            continue
        if not (md_dir / unquote(path)).exists() and not (md_dir / path).exists():
            broken.append(target)
        elif Path(path).suffix.lower() in UNVIEWABLE:
            unviewable.append(target)
    return {"broken": broken, "unviewable": unviewable}


_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_MC_FALLBACK = "{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback"


def docx_struck_chars(src: Path) -> int:
    """Characters the document body strikes through by direct formatting, as office2md marks them:
    not in a tracked deletion (written {--...--}), a ruby, hidden text or a drawing's fallback
    copy. Headers, footers and notes are left out: office2md writes them as plain text."""
    from lxml import etree
    try:
        z = zipfile.ZipFile(src)
    except (zipfile.BadZipFile, OSError):
        return 0
    parser = etree.XMLParser(resolve_entities=False, no_network=True, huge_tree=False)
    total = 0
    with z:
        parts = [n for n in z.namelist() if n == "word/document.xml"]
        for name in parts:
            try:
                root = etree.fromstring(z.read(name), parser)
            except (KeyError, etree.XMLSyntaxError):
                continue
            for r in root.iter(_W + "r"):
                rpr = r.find(_W + "rPr")
                if rpr is None:
                    continue

                def on(tag: str) -> bool:
                    el = rpr.find(_W + tag)
                    return el is not None and (el.get(_W + "val") or "true").lower() not in ("0", "false", "off")
                if not (on("strike") or on("dstrike")) or on("vanish"):
                    continue
                anc, skip = r.getparent(), False
                while anc is not None:
                    if anc.tag in (_W + "del", _W + "ruby", _MC_FALLBACK, _W + "moveFrom"):
                        skip = True
                        break
                    anc = anc.getparent()
                if not skip:
                    total += len(re.sub(r"\s", "", "".join(t.text or "" for t in r.iter(_W + "t"))))
    return total


def md_struck_chars(md: str) -> int:
    """Characters inside ~~...~~ and <del>...</del>, spaces, escapes and inline tags left out
    (a span may run over a line break). Emphasis marks inside count, so an error leans to more."""
    spans = re.findall(r"~~(.{1,3000}?)~~", md, re.S) + re.findall(r"<del>(.{1,3000}?)</del>", md, re.S)
    return sum(len(re.sub(r"</?(?:u|em|strong|mark|sup|sub|br)\s{0,3}/?>|\\(?=\S)|\s", "", s)) for s in spans)


_FIG_NAME = re.compile(r"figure-(\d{3,4})(?:-\d+)?\.\w+$")


def figure_order(md: str, data: dict) -> list[str]:
    """The package names of the document's pictures in the order the Markdown shows them (office2md
    saves picture n as figure-00n)."""
    images = data.get("images", [])
    out = []
    for pic, target in local_links(md):
        m = _FIG_NAME.search(target)
        if pic and m and 0 < int(m[1]) <= len(images) and images[int(m[1]) - 1].get("file"):
            out.append(images[int(m[1]) - 1]["file"])
    return list(dict.fromkeys(out))


def pandoc_order(md: str) -> list[str]:
    """The package names of the pictures in pandoc's Markdown, in order. pandoc keeps
    media/imageN.ext and writes the extraction folder as given, spaces and all, so the names are
    read from the path's end rather than through the link syntax."""
    return list(dict.fromkeys(re.findall(r"[\\/]media[\\/]([\w.-]{1,100}?\.[A-Za-z0-9]{2,5})(?=[)}\s>])", md)))


def order_moves(primary: list[str], ref: list[str]) -> list[str]:
    """Pictures both sides show whose order differs: the ones outside what the two share in the
    same order, so a single moved picture is named once, not every picture after it. difflib's
    matching keeps memory linear for a document with thousands of pictures."""
    import difflib
    common = set(primary) & set(ref)
    a = [x for x in primary if x in common]
    b = [x for x in ref if x in common]
    if a == b:
        return []
    keep = set()
    for blk in difflib.SequenceMatcher(None, a, b, autojunk=False).get_matching_blocks():
        keep.update(a[blk.a:blk.a + blk.size])
    return [x for x in a if x not in keep]


def picture_count(md: str) -> int:
    """Pictures and figure markers in the Markdown: what office2md writes for each picture."""
    return (sum(1 for pic, _ in local_links(md) if pic)
            + len(re.findall(r"(?<!!)\[(?:figure|drawing): ", md)))


# what may sit between a line's Markdown prefix and its text and still be only a number label
_LABEL_ONLY = re.compile(r"\s*[(（\[]?[0-9０-９A-Za-zⅰ-ⅿⅠ-Ⅿ①-⑳一二三四五六七八九十〇ぁ-んァ-ヶｦ-ﾝ第章節項条.．\-]{1,12}"
                         r"[)）\]]?[.．)）]?\s*")


def _num_norm(s: str) -> str:
    """A line as the number check compares it: width-folded, without spaces, Markdown or tags."""
    import unicodedata
    return re.sub(r"[\s\\*_`#>|]|</?\w+[^<>]{0,50}>", "", unicodedata.normalize("NFKC", s))


def _label_of(row) -> str | None:
    """Word's label for a numbered paragraph, or None for a bullet: Word's list type says so
    (1 bullet, 6 picture bullet), or the label is a symbol or one letter in a symbol font (o, l)."""
    label = row[0].strip()
    if len(row) > 2 and row[2] in ("1", "6"):
        return None
    folded = _num_norm(label)                        # full-width １ reads as 1, half-width ｱ as ア
    if not re.search(r"[0-9A-Za-zⅰ-ⅿⅠ-Ⅿ①-⑳一二三四五六七八九十ぁ-んァ-ヶｦ-ﾝ]", folded) or re.fullmatch(r"[A-Za-z]", folded):
        return None
    return label


def apply_numbers(numbers: list[tuple], md: str) -> tuple[str, int]:
    """Word's own number on each Markdown line whose number differs from it: where Word is at
    hand it is the authority, and office2md's numbering is the fallback. Only a line found once,
    holding the paragraph's text after nothing but its Markdown prefix and a number label, is
    changed, so a table cell or a quote that repeats the words is left alone."""
    n = _num_norm
    lines = md.split("\n")
    norm = [n(l) for l in lines]
    changed = 0
    for row in numbers:
        label = _label_of(row)
        text = row[1].strip() if len(row) > 1 else ""
        # the label comes from the document's numbering definition: only something shaped like a
        # number goes into the Markdown, never brackets with a target, tags or other markup
        if not label or len(label) > 40 or not _LABEL_ONLY.fullmatch(label) or len(n(text)) < 4:
            continue
        key = n(text)[:24]
        hits = [i for i, nl in enumerate(norm) if key in nl]
        if len(hits) != 1:
            continue
        i = hits[0]
        lab = n(label).rstrip(".")
        if lab in norm[i][:max(0, norm[i].find(key))]:
            continue                                   # the number is already there
        raw = lines[i]
        start = raw.find(text[:6])
        if start < 0:
            continue
        prefix = re.match(r"\s*(?:#{1,9}\s+|[-*+]\s+|>\s*)*(?:\*\*|<strong>)?", raw[:start])
        between = raw[prefix.end():start]
        if between.strip() and not _LABEL_ONLY.fullmatch(between):
            continue
        lines[i] = raw[:prefix.end()] + label + " " + raw[start:]
        norm[i] = n(lines[i])
        changed += 1
    return "\n".join(lines), changed


def numbering_misses(numbers: list[tuple], md: str) -> list[tuple[str, str, str]]:
    """Paragraphs Word numbers (its ListString) whose number the Markdown line lacks. Returns
    (label, text, the Markdown line). A paragraph whose text is not found is the text check's
    business; a text found on several lines is fine when one of them carries the number."""
    n = _num_norm
    lines = [(n(l), l) for l in md.splitlines() if l.strip()]
    out = []
    for row in numbers:
        label, text = _label_of(row), row[1]
        if not label:
            continue                                      # a bullet: Word's list type, a symbol, a symbol-font letter
        lab = n(label).rstrip(".")
        key = n(text)[:24]
        if len(key) < 4:
            continue
        hits = [(nl, raw) for nl, raw in lines if key in nl]
        if not hits:
            continue
        if not any(lab in nl[:max(0, nl.find(key))] for nl, _ in hits):
            out.append((label, text, hits[0][1]))
    return out
