#!/usr/bin/env python3
"""Read a Visio drawing (.vsdx / .vsdm) as Markdown: each page's flow as mermaid plus a list.

Usage:
    visio_reader.py 業務フロー.vsdx
    visio_reader.py 業務フロー.vsdx --json
    office2md.py 業務フロー.vsdx                 the same, through the common entry point

Business flows are usually drawn in Visio with lanes (one per department), tasks, events,
gateways and flows glued between them. Visio stores which shape each connector is glued to
(the page's Connect elements), so the flow is read from the glue, not guessed from positions:
each connector becomes an edge from the shape glued at its begin to the shape glued at its end,
with its own text as the label. A shape is placed in the lane whose area holds its centre.
Shapes that carry text but are not part of the flow (notes, titles, legends) are listed so
nothing written on the page is lost.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import office2md  # noqa: E402

try:
    sys.stdout.reconfigure(encoding="utf-8")  # Windows console defaults to cp932
except AttributeError:
    pass

V = "{http://schemas.microsoft.com/office/visio/2012/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
RELS = "{http://schemas.openxmlformats.org/package/2006/relationships}"
MAX_SHAPES = 20000
CONNECTOR_MASTERS = re.compile(r"connector|sequence flow|message flow|association|コネクタ|線", re.I)
LANE_MASTERS = re.compile(r"swimlane|pool|lane|cff container|スイムレーン|レーン", re.I)
DECOR_MASTERS = re.compile(r"title block|revision|parts list|projection|ruled column|section header|phase list|"
                           r"separator|swimlane list|block - side label", re.I)
NODE_KIND = [(re.compile(r"gateway|decision|判断|分岐", re.I), "gateway"),
             (re.compile(r"start event|開始", re.I), "start"), (re.compile(r"end event|終了", re.I), "end"),
             (re.compile(r"intermediate event|event|イベント", re.I), "event"),
             (re.compile(r"data object|document|データ|帳票", re.I), "data"),
             (re.compile(r"data store|database|データストア", re.I), "store"),
             (re.compile(r"annotation|注釈", re.I), "note")]


class VisioError(office2md.ReadError):
    pass


def _parser():
    from lxml import etree
    return etree.XMLParser(resolve_entities=False, no_network=True, huge_tree=False)


def _xml(z: zipfile.ZipFile, name: str):
    from lxml import etree
    return etree.fromstring(z.read(name), _parser())


def _rels(z: zipfile.ZipFile, part: str) -> dict[str, str]:
    d, base = part.rsplit("/", 1)
    rp = f"{d}/_rels/{base}.rels"
    if rp not in z.namelist():
        return {}
    out = {}
    for r in _xml(z, rp).iter(RELS + "Relationship"):
        if r.get("TargetMode") == "External":
            continue
        segs = d.split("/")
        for s in (r.get("Target") or "").split("/"):
            if s == "..":
                segs = segs[:-1]
            elif s and s != ".":
                segs.append(s)
        out[r.get("Id")] = "/".join(segs)
    return out


def _cells(shape) -> dict[str, str]:
    return {c.get("N"): c.get("V") or "" for c in shape.findall(V + "Cell")}


def _num(v, default=0.0) -> float:
    try:
        return float(v)
    except (TypeError, ValueError):
        return default


def _text(shape) -> str:
    t = shape.find(V + "Text")
    if t is None:
        return ""
    s = "".join(t.itertext())
    return re.sub(r"\s+", " ", re.sub(r"[\x00-\x1f\x7f]", " ", s)).strip()[:500]


def _masters(z: zipfile.ZipFile) -> tuple[dict[str, str], dict[tuple[str, str], dict]]:
    """master id -> name, and (master id, shape id) -> cells of the master's shapes (inherited)."""
    names: dict[str, str] = {}
    cells: dict[tuple[str, str], dict] = {}
    if "visio/masters/masters.xml" not in z.namelist():
        return names, cells
    rels = _rels(z, "visio/masters/masters.xml")
    for m in _xml(z, "visio/masters/masters.xml").iter(V + "Master"):
        mid = m.get("ID") or ""
        names[mid] = m.get("NameU") or m.get("Name") or ""
        rel = m.find(V + "Rel")
        part = rels.get(rel.get(R + "id")) if rel is not None else None
        if part and part in z.namelist():
            try:
                for s in _xml(z, part).iter(V + "Shape"):
                    cells[(mid, s.get("ID") or "")] = _cells(s)
            except Exception:
                pass
    return names, cells


def _pages(z: zipfile.ZipFile) -> list[tuple[str, str]]:
    if "visio/pages/pages.xml" not in z.namelist():
        return []
    rels = _rels(z, "visio/pages/pages.xml")
    out = []
    for p in _xml(z, "visio/pages/pages.xml").iter(V + "Page"):
        if p.get("Background") == "1":
            continue
        rel = p.find(V + "Rel")
        part = rels.get(rel.get(R + "id")) if rel is not None else None
        if part and part in z.namelist():
            out.append((office2md.md_escape(re.sub(r"[\x00-\x1f\x7f]", " ", p.get("NameU") or p.get("Name") or "page")), part))
    return out


def read_page(z: zipfile.ZipFile, part: str, masters: dict[str, str], mcells: dict) -> dict:
    root = _xml(z, part)
    shapes: dict[str, dict] = {}
    top = root.find(V + "Shapes")

    def walk(container, ox: float, oy: float, master: str, parent: str | None, depth: int):
        if container is None or depth > 12 or len(shapes) >= MAX_SHAPES:
            return
        for s in container.findall(V + "Shape"):
            sid = s.get("ID") or ""
            m = s.get("Master") or master
            own = _cells(s)
            inherited = mcells.get((m, s.get("MasterShape") or ""), {}) if m else {}
            if not inherited and m and s.get("Master"):
                inherited = next((c for (mm, _), c in mcells.items() if mm == m), {})
            c = {**inherited, **own}
            w, h = _num(c.get("Width")), _num(c.get("Height"))
            px, py = ox + _num(c.get("PinX")), oy + _num(c.get("PinY"))
            lx, ly = _num(c.get("LocPinX"), w / 2), _num(c.get("LocPinY"), h / 2)
            x0, y0 = px - lx, py - ly
            name = masters.get(m, "") if m else ""
            one_d = "BeginX" in c and "EndX" in c
            shapes[sid] = {"id": sid, "master": name, "text": _text(s), "parent": parent,
                           "box": (x0, y0, x0 + w, y0 + h), "one_d": one_d,
                           "begin": (ox + _num(c.get("BeginX")), oy + _num(c.get("BeginY"))) if one_d else None,
                           "end": (ox + _num(c.get("EndX")), oy + _num(c.get("EndY"))) if one_d else None,
                           "dashed": _num(c.get("LinePattern"), 1) not in (0, 1)}
            walk(s.find(V + "Shapes"), x0, y0, m, sid, depth + 1)
    walk(top, 0.0, 0.0, "", None, 0)

    def top_of(sid: str) -> str:
        """The shape a glue point stands for: climb to the outermost group below the page."""
        seen = 0
        while sid in shapes and shapes[sid]["parent"] and seen < 20:
            sid = shapes[sid]["parent"]
            seen += 1
        return sid

    def label_of(sid: str) -> str:
        s = shapes.get(sid)
        if not s:
            return ""
        if s["text"]:
            return s["text"]
        kids = [x["text"] for x in shapes.values() if x["parent"] == sid and x["text"]]
        return " / ".join(kids)[:300]

    ends: dict[str, dict[str, str]] = {}
    for cn in root.iter(V + "Connect"):
        src, cell, dst = cn.get("FromSheet") or "", cn.get("FromCell") or "", cn.get("ToSheet") or ""
        if cell in ("BeginX", "EndX") and src in shapes and dst in shapes:
            ends.setdefault(src, {})["begin" if cell == "BeginX" else "end"] = top_of(dst)
    lanes = [s for s in shapes.values() if LANE_MASTERS.search(s["master"]) and label_of(s["id"])
             and not re.search(r"swimlane list", s["master"], re.I)]
    lane_ids = {ln["id"] for ln in lanes}

    def is_node(s: dict) -> bool:
        # a named shape, or a BPMN element that carries no text (events, gateways): never a bare
        # background rectangle, which would swallow every loose line end near it
        named = bool(label_of(s["id"])) or any(rx.search(s["master"] or "") for rx, _ in NODE_KIND)             or re.search(r"task|process|sub-process|タスク|処理", s["master"] or "", re.I) is not None
        return (not s["one_d"] and not s["parent"] and s["id"] not in lane_ids
                and not DECOR_MASTERS.search(s["master"] or "") and named)

    candidates = [s for s in shapes.values() if is_node(s)]

    def at_point(pt) -> tuple[str, str]:
        """(shape id, how) for a loose connector end: the shape under or next to it, else the lane."""
        x, y = pt
        best, dist = None, 0.35                  # inches: a hair beside a box still counts as touching it
        for s in candidates:
            b = s["box"]
            dx, dy = max(b[0] - x, 0, x - b[2]), max(b[1] - y, 0, y - b[3])
            d = (dx * dx + dy * dy) ** 0.5
            if d < dist:
                best, dist = s["id"], d
        if best:
            return best, "position"
        for ln in sorted(lanes, key=lambda l_: (l_["box"][2] - l_["box"][0]) * (l_["box"][3] - l_["box"][1])):
            b = ln["box"]
            if b[0] <= x <= b[2] and b[1] <= y <= b[3]:
                return ln["id"], "lane"
        return "", ""

    edges, used, notes = [], set(), []
    for cid, c in shapes.items():
        if not c["one_d"] or DECOR_MASTERS.search(c["master"] or ""):
            continue
        e = ends.get(cid, {})
        how = "glued"
        b_id, e_id = e.get("begin"), e.get("end")
        if not b_id and c["begin"]:
            b_id, how = at_point(c["begin"])[0], "position"
        if not e_id and c["end"]:
            e_id, how = at_point(c["end"])[0], "position"
        if not b_id or not e_id or b_id == e_id:
            continue
        touches_data = any(re.search(r"data|document|データ|帳票", shapes[x]["master"] or "", re.I)
                           for x in (b_id, e_id))
        # a line from a plain text box (no master) to a shape is an annotation of that shape
        plain = [x for x in (b_id, e_id) if not shapes[x]["master"] and x not in lane_ids and shapes[x]["text"]]
        if plain and len(plain) == 1:
            note, target = plain[0], (e_id if plain[0] == b_id else b_id)
            notes.append({"on": target, "text": shapes[note]["text"], "how": how})
            used |= {note, cid}
            continue
        kind = ("association" if re.search(r"association", c["master"], re.I) or (c["dashed"] and touches_data) else
                "message" if re.search(r"message", c["master"], re.I) or (how == "position" and c["dashed"]) else "flow")
        edges.append({"from": b_id, "to": e_id, "label": label_of(cid), "kind": kind, "how": how, "id": cid})
        used |= {b_id, e_id, cid}

    # a line glued to another line: a data object or note riding on a message flow
    by_line = {e["id"]: e for e in edges}
    kept = []
    for e in edges:
        line_end = next((x for x in (e["to"], e["from"]) if x in by_line and x != e["id"]), None)
        if line_end:
            other = e["from"] if line_end == e["to"] else e["to"]
            by_line[line_end].setdefault("carries", []).append(other)
            continue
        kept.append(e)
    edges = kept

    # a flow's label is often a separate text box beside the line ("却下の場合", "進達書類の送付")
    free = [s for s in shapes.values() if s["text"] and not s["one_d"] and not s["parent"] and s["id"] not in used
            and s["id"] not in lane_ids and not s["master"] and not DECOR_MASTERS.search(s["master"] or "")]
    taken: set[str] = set()
    for e in edges:
        if e["label"] or e["kind"] == "association":
            continue
        c = shapes[e["id"]]
        pts = [p for p in (c["begin"], c["end"]) if p] + ([((c["begin"][0] + c["end"][0]) / 2,
                                                            (c["begin"][1] + c["end"][1]) / 2)] if c["begin"] else [])
        best, dist = None, 0.6
        for s in free:
            if s["id"] in taken or len(s["text"]) > 40:
                continue
            b = s["box"]
            cx, cy = (b[0] + b[2]) / 2, (b[1] + b[3]) / 2
            d = min(((cx - x) ** 2 + (cy - y) ** 2) ** 0.5 for x, y in pts) if pts else 9
            if d < dist:
                best, dist = s, d
        if best:
            e["label"], e["label_by_position"] = best["text"], True
            taken.add(best["id"])
            used.add(best["id"])

    # reading order: from the start events along the flow, then whatever is left by position
    order: list[dict] = []
    out_of: dict[str, list[dict]] = {}
    for e in edges:
        out_of.setdefault(e["from"], []).append(e)
    starts = [e["from"] for e in edges if not any(x["to"] == e["from"] for x in edges)]
    starts.sort(key=lambda i: (-shapes[i]["box"][3], shapes[i]["box"][0]))
    seen_e: set[int] = set()
    stack = list(reversed(starts))
    while stack:
        n = stack.pop()
        for e in sorted(out_of.get(n, []), key=lambda e: (-shapes[e["to"]]["box"][3], shapes[e["to"]]["box"][0]),
                        reverse=True):
            if id(e) not in seen_e:
                seen_e.add(id(e))
                order.append(e)
                stack.append(e["to"])
    order += [e for e in edges if id(e) not in seen_e]
    edges = order

    def lane_of(sid: str) -> str:
        b = shapes[sid]["box"]
        cx, cy = (b[0] + b[2]) / 2, (b[1] + b[3]) / 2
        hits = [ln for ln in lanes if ln["id"] != sid and ln["box"][0] <= cx <= ln["box"][2] and ln["box"][1] <= cy <= ln["box"][3]]
        if not hits:
            return ""
        best = min(hits, key=lambda ln: (ln["box"][2] - ln["box"][0]) * (ln["box"][3] - ln["box"][1]))
        return label_of(best["id"])

    def kind_of(sid: str) -> str:
        if sid in lane_ids:
            return "lane"
        m = shapes[sid]["master"]
        for rx, k in NODE_KIND:
            if rx.search(m):
                return k
        return "task"
    node_ids = sorted({x for e in edges for x in (e["from"], e["to"])},
                      key=lambda i: (-shapes[i]["box"][3], shapes[i]["box"][0]))
    nodes = [{"id": i, "label": label_of(i), "kind": kind_of(i), "lane": "" if i in lane_ids else lane_of(i)}
             for i in node_ids]
    other = []
    for s in shapes.values():
        if not s["text"] or s["id"] in used or s["id"] in lane_ids or s["one_d"]:
            continue
        anc, inside = s["parent"], False
        while anc:
            if anc in used:
                inside = True
                break
            anc = shapes[anc]["parent"] if anc in shapes else None
        if inside or DECOR_MASTERS.search(s["master"] or ""):
            continue
        other.append({"text": s["text"], "lane": lane_of(s["id"]), "y": -s["box"][3], "x": s["box"][0]})
    other.sort(key=lambda o: (o["y"], o["x"]))
    decor = [s["text"] for s in shapes.values() if s["text"] and DECOR_MASTERS.search(s["master"] or "")]
    names = [label_of(ln["id"]) for ln in sorted(lanes, key=lambda ln: (ln["box"][0], -ln["box"][3]))]
    return {"lanes": list(dict.fromkeys(names)),
            "nodes": nodes,
            "edges": [{**e, "carries": [label_of(x) or kind_of(x) for x in e.get("carries", [])]} for e in edges],
            "notes": [{"on": label_of(n["on"]) or kind_of(n["on"]), "text": n["text"], "how": n["how"]} for n in notes], "other": [{"text": o["text"], "lane": o["lane"]} for o in other],
            "decor": decor, "shapes": len(shapes)}


def read(path: Path) -> dict:
    try:
        z = zipfile.ZipFile(path)
    except (zipfile.BadZipFile, OSError) as e:
        raise VisioError(f"cannot open {path.name}: {e}") from None
    with z:
        masters, mcells = _masters(z)
        pages = []
        for name, part in _pages(z):
            try:
                pages.append({"name": name, **read_page(z, part, masters, mcells)})
            except Exception as e:                    # one broken page must not hide the others
                pages.append({"name": name, "error": f"{type(e).__name__}: {e}"[:200]})
    return {"file": str(path), "type": "vsdx", "pages": pages}


def _mermaid(page: dict) -> list[str]:
    def safe(s: str) -> str:
        return re.sub(r'["`\[\]{}()<>|#;]', " ", s).strip()[:80] or " "
    shape = {"gateway": ('{"', '"}'), "start": ('(("', '"))'), "end": ('((("', '")))'), "event": ('(("', '"))'),
             "data": ('[/"', '"/]'), "store": ('[("', '")]'), "note": ('>"', '"]'), "task": ('["', '"]'),
             "lane": ('(["', '"])')}
    ids = {n["id"]: f"N{i}" for i, n in enumerate(page["nodes"], start=1)}
    out = ["```mermaid", "flowchart LR"]
    by_lane: dict[str, list[dict]] = {}
    for n in page["nodes"]:
        by_lane.setdefault(n["lane"], []).append(n)
    for k, (lane, ns) in enumerate(by_lane.items(), start=1):
        if lane:
            out.append(f'    subgraph L{k}["{safe(lane)}"]')
        for n in ns:
            a, b = shape[n["kind"]]
            out.append(f"    {'    ' if lane else ''}{ids[n['id']]}{a}{safe(n['label'])}{b}")
        if lane:
            out.append("    end")
    arrow = {"flow": "-->", "message": "-.->", "association": "-.-"}
    for e in page["edges"]:
        lab = f'|"{safe(e["label"])}"|' if e["label"] else ""
        out.append(f"    {ids[e['from']]} {arrow[e['kind']]}{lab} {ids[e['to']]}")
    return out + ["```"]


def to_markdown(data: dict, path: Path) -> str:
    esc = office2md.md_escape
    out = office2md.md_title(str(path))
    for p in data["pages"]:
        out += [f"## Page: {p['name']}", ""]
        if p.get("error"):
            out += [f"_(page could not be read: {esc(p['error'])})_", ""]
            continue
        if p["lanes"]:
            out += ["Lanes: " + " · ".join(esc(x) for x in p["lanes"]), ""]
        if p["edges"]:
            out += ["**Flow** (from the shapes the connectors are glued to; verify against the drawing)", ""]
            out += _mermaid(p) + [""]
            names = {n["id"]: (f"[{esc(n['lane'])}] " if n["lane"] else "") + esc(n["label"] or f"({n['kind']})")
                     for n in p["nodes"]}
            for e in p["edges"]:
                kind = "" if e["kind"] == "flow" else f" ({e['kind']})"
                if e.get("how") == "position":
                    kind += " (end found by position, not glued)"
                lab = f" [{esc(e['label'])}]" if e["label"] else ""
                if e.get("label_by_position"):
                    lab += " (label: nearest text)"
                if e.get("carries"):
                    lab += " carrying " + ", ".join(esc(c) for c in e["carries"])
                out.append(f"- {names[e['from']]} → {names[e['to']]}{lab}{kind}")
            out.append("")
        else:
            out += ["_(no connectors glued between shapes on this page)_", ""]
        if p.get("notes"):
            out += ["**Annotations** (text boxes linked to a shape)", ""]
            out += [f"- {esc(n['on'])}: {esc(n['text'])}" + (" (linked by position)" if n["how"] == "position" else "")
                    for n in p["notes"]]
            out.append("")
        if p["other"]:
            out += ["**Other text on the page**", ""]
            out += [f"- {('[' + esc(o['lane']) + '] ') if o['lane'] else ''}{esc(o['text'])}" for o in p["other"][:300]]
            out.append("")
        if p["decor"]:
            out += ["Title and revision blocks: " + " / ".join(esc(t) for t in p["decor"][:40]), ""]
    if not data["pages"]:
        out.append("_(no pages)_")
    return "\n".join(out) + "\n"


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("file", type=Path)
    ap.add_argument("--json", action="store_true")
    ap.add_argument("--out", type=Path)
    a = ap.parse_args(argv)
    if not a.file.exists():
        sys.exit(f"[visio_reader] file not found: {a.file}")
    try:
        data = read(a.file)
    except VisioError as e:
        sys.exit(f"[visio_reader] {e}")
    text = json.dumps(data, ensure_ascii=False, indent=2) if a.json else to_markdown(data, a.file)
    if a.out:
        a.out.write_text(text, encoding="utf-8")
        print(f"[visio_reader] wrote {a.out} ({len(text)} chars)")
    else:
        print(text, end="")
    return 0


if __name__ == "__main__":
    sys.exit(main())
