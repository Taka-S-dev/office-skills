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

def renders_current(render_dir: Path, src: Path) -> bool:
    stamp = render_dir / "source.sha256"
    try:
        return stamp.read_text(encoding="utf-8").strip() == office2md.file_digest(src)
    except OSError:
        return False


def stamp_renders(render_dir: Path, src: Path) -> None:
    try:
        render_dir.mkdir(parents=True, exist_ok=True)
        (render_dir / "source.sha256").write_text(office2md.file_digest(src) + "\n", encoding="utf-8")
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

def docx_visuals(data: dict, page_pdf: str = "") -> list[tuple[str, str, str]]:
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
    for i, im in enumerate(data.get("images", []), start=1):
        if im.get("placeholder"):
            continue
        shown = im.get("saved") or im.get("file") or im.get("name")
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
