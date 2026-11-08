"""Shapes, connectors, charts, pictures and input rules in an .xlsx, placed on the sheet's grid.

openpyxl reads none of this. Sequence diagrams, flowcharts and record layouts in Excel are drawn
with boxes, text boxes and arrows, so their meaning is in where the shapes sit and which way the
arrows point. This module reads the drawing parts from the package, puts every shape in one
coordinate system (EMU, derived from column widths and row heights so that shapes and cells can
be compared), and reconstructs what the drawing says:

- connectors: start and end point, arrowheads, and the shapes they are glued to
- sequence diagrams: vertical lines are lifelines, the boxes above them are participants,
  horizontal arrows between two lifelines are messages, labelled by the nearest text box or
  cell text above them, ordered top to bottom
- flowcharts: connectors between shapes, glued or ending near a shape
- charts: type, title, series names, categories and values (cached, or read from the cells)
- input rules (data validation), including the newer kind openpyxl drops

Everything returned is plain data. office2md renders it; md_crosscheck uses it to list what to
confirm visually. Reconstruction is heuristic, so every result keeps the cell positions it came
from and anything not explained is reported, not dropped.
"""

from __future__ import annotations

import re
import struct
import zipfile
from bisect import bisect_right
from pathlib import Path

NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "x14": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
    "xm": "http://schemas.microsoft.com/office/excel/2006/main",
}
EMU_PER_PX = 9525
EMU_PER_PT = 12700
MAX_COL = 16384
MAX_ROW = 1048576
MAX_SHAPES = 5000                    # per sheet; a crafted file cannot make this loop forever
LINE_PRST = {"line", "straightConnector1"} | {f"bentConnector{i}" for i in range(2, 6)} \
    | {f"curvedConnector{i}" for i in range(2, 6)}
ARROW_PRST = {"rightArrow": (0, 0.5, 1, 0.5), "leftArrow": (1, 0.5, 0, 0.5), "upArrow": (0.5, 1, 0.5, 0),
              "downArrow": (0.5, 0, 0.5, 1), "leftRightArrow": (0, 0.5, 1, 0.5), "upDownArrow": (0.5, 0, 0.5, 1),
              "notchedRightArrow": (0, 0.5, 1, 0.5), "stripedRightArrow": (0, 0.5, 1, 0.5),
              "rightArrowCallout": (0, 0.5, 1, 0.5), "leftArrowCallout": (1, 0.5, 0, 0.5),
              "chevron": (0, 0.5, 1, 0.5), "homePlate": (0, 0.5, 1, 0.5)}
ARROW_HEADS = {"triangle", "stealth", "arrow"}
IMAGE_EXT = {"png", "jpg", "jpeg", "gif", "bmp", "tif", "tiff", "emf", "wmf", "svg", "webp"}


def _q(prefix: str, name: str) -> str:
    return "{%s}%s" % (NS[prefix], name)


def _local(tag) -> str:
    return tag.rsplit("}", 1)[-1] if isinstance(tag, str) else ""


def _int(v, default=0, lo=-(2 ** 40), hi=2 ** 40) -> int:
    try:
        return max(lo, min(hi, int(v)))
    except (TypeError, ValueError):
        return default


def col_letter(idx: int) -> str:
    s = ""
    while idx > 0:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s


def clean_text(s: str, limit: int = 400) -> str:
    return re.sub(r"\s+", " ", re.sub(r"[\x00-\x1f\x7f]", " ", s or "")).strip()[:limit]


# -- sheet geometry -----------------------------------------------------------------------

class Metrics:
    """Cell grid in EMU. Column widths and row heights follow Excel's defaults closely enough that
    shapes anchored to cells and cell text can be compared; exact pixels are not needed."""

    def __init__(self, ws):
        fmt = ws.sheet_format
        base = fmt.baseColWidth or 8
        self.def_w = float(fmt.defaultColWidth) if fmt.defaultColWidth else (8.43 if base == 8 else base + 0.71)
        self.def_h = float(fmt.defaultRowHeight) if fmt.defaultRowHeight else 15.0
        self.col_w: dict[int, float] = {}
        for letter, dim in ws.column_dimensions.items():
            try:
                lo = dim.min or _col_index(letter)
                hi = dim.max or lo
            except Exception:
                continue
            w = 0.0 if dim.hidden else (dim.width if dim.width else self.def_w)
            for c in range(max(1, lo), min(MAX_COL, hi, lo + 512) + 1):
                self.col_w[c] = w
        self.row_h: dict[int, float] = {}
        for r, dim in ws.row_dimensions.items():
            if dim.hidden:
                self.row_h[r] = 0.0
            elif dim.ht:
                self.row_h[r] = float(dim.ht)
        self._xs = [0]          # _xs[i] = left edge of column i+1
        self._ys = [0]

    def col_emu(self, c1: int) -> int:
        w = self.col_w.get(c1, self.def_w)
        return 0 if w <= 0 else int(w * 7 + 5) * EMU_PER_PX

    def row_emu(self, r1: int) -> int:
        return int(self.row_h.get(r1, self.def_h) * EMU_PER_PT)

    @property
    def default_col(self) -> int:
        return int(self.def_w * 7 + 5) * EMU_PER_PX

    @property
    def default_row(self) -> int:
        return int(self.def_h * EMU_PER_PT)

    def _grow_x(self, n: int):
        n = min(n, MAX_COL)
        while len(self._xs) <= n:
            self._xs.append(self._xs[-1] + self.col_emu(len(self._xs)))

    def _grow_y(self, n: int):
        n = min(n, MAX_ROW)
        while len(self._ys) <= n:
            self._ys.append(self._ys[-1] + self.row_emu(len(self._ys)))

    def x(self, col0: int, off: int = 0) -> int:
        col0 = max(0, min(MAX_COL - 1, col0))
        self._grow_x(col0)
        return self._xs[col0] + off

    def y(self, row0: int, off: int = 0) -> int:
        row0 = max(0, min(MAX_ROW - 1, row0))
        self._grow_y(row0)
        return self._ys[row0] + off

    def cell_box(self, row1: int, col1: int) -> tuple[int, int, int, int]:
        return self.x(col1 - 1), self.y(row1 - 1), self.x(col1), self.y(row1)

    def cell_at(self, x: float, y: float) -> str:
        col = 1
        self._grow_x(64)
        while self._xs[-1] <= x and len(self._xs) <= MAX_COL:
            self._grow_x(len(self._xs) + 64)
        col = max(1, min(MAX_COL, bisect_right(self._xs, x)))
        self._grow_y(64)
        while self._ys[-1] <= y and len(self._ys) <= min(MAX_ROW, 200000):
            self._grow_y(len(self._ys) + 256)
        row = max(1, min(MAX_ROW, bisect_right(self._ys, y)))
        return f"{col_letter(col)}{row}"


def _col_index(letters: str) -> int:
    n = 0
    for ch in letters.upper():
        if "A" <= ch <= "Z":
            n = n * 26 + (ord(ch) - 64)
    return n


# -- package access -----------------------------------------------------------------------

class Package:
    def __init__(self, path: Path):
        from lxml import etree
        self._etree = etree
        self.parser = etree.XMLParser(resolve_entities=False, no_network=True, huge_tree=False)
        self.z = zipfile.ZipFile(path)
        self.names = set(self.z.namelist())

    def xml(self, name: str):
        return self._etree.fromstring(self.z.read(name), self.parser)

    def rels(self, part: str) -> dict[str, tuple[str, str, bool]]:
        """rId -> (type suffix, absolute target, external?)"""
        d, base = part.rsplit("/", 1) if "/" in part else ("", part)
        rp = f"{d}/_rels/{base}.rels" if d else f"_rels/{base}.rels"
        if rp not in self.names:
            return {}
        out = {}
        for r in self.xml(rp).findall(_q("rel", "Relationship")):
            tgt = r.get("Target") or ""
            external = r.get("TargetMode") == "External"
            if not external:
                if tgt.startswith("/"):
                    tgt = tgt.lstrip("/")
                else:
                    segs = d.split("/") if d else []
                    for s in tgt.split("/"):
                        if s == "..":
                            if segs:
                                segs.pop()
                        elif s and s != ".":
                            segs.append(s)
                    tgt = "/".join(segs)
            out[r.get("Id")] = ((r.get("Type") or "").rsplit("/", 1)[-1], tgt, external)
        return out

    def close(self):
        self.z.close()


def sheet_parts(pkg: Package) -> list[tuple[str, str]]:
    """[(sheet title, part path)] in tab order."""
    if "xl/workbook.xml" not in pkg.names:
        return []
    root = pkg.xml("xl/workbook.xml")
    rels = pkg.rels("xl/workbook.xml")
    out = []
    for sh in root.findall(_q("m", "sheets") + "/" + _q("m", "sheet")):
        rid = sh.get(_q("r", "id"))
        if rid in rels and rels[rid][1] in pkg.names:
            out.append((sh.get("name"), rels[rid][1]))
    return out


# -- shapes ---------------------------------------------------------------------------------

def _anchor_box(anchor, m: Metrics):
    tag = _local(anchor.tag)
    X = lambda n: _q("xdr", n)  # noqa: E731

    def point(el):
        return (_int(el.findtext(X("col")), 0, 0, MAX_COL - 1), _int(el.findtext(X("colOff"))),
                _int(el.findtext(X("row")), 0, 0, MAX_ROW - 1), _int(el.findtext(X("rowOff"))))

    if tag == "twoCellAnchor":
        f, t = anchor.find(X("from")), anchor.find(X("to"))
        if f is None or t is None:
            return None
        c1, o1, r1, p1 = point(f)
        c2, o2, r2, p2 = point(t)
        return (m.x(c1, o1), m.y(r1, p1), m.x(c2, o2), m.y(r2, p2))
    if tag == "oneCellAnchor":
        f, ext = anchor.find(X("from")), anchor.find(X("ext"))
        if f is None:
            return None
        c1, o1, r1, p1 = point(f)
        x1, y1 = m.x(c1, o1), m.y(r1, p1)
        cx = _int(ext.get("cx"), 0, 0) if ext is not None else 0
        cy = _int(ext.get("cy"), 0, 0) if ext is not None else 0
        return (x1, y1, x1 + cx, y1 + cy)
    if tag == "absoluteAnchor":
        pos, ext = anchor.find(X("pos")), anchor.find(X("ext"))
        if pos is None:
            return None
        x1, y1 = _int(pos.get("x")), _int(pos.get("y"))
        return (x1, y1, x1 + _int(ext.get("cx") if ext is not None else 0), y1 + _int(ext.get("cy") if ext is not None else 0))
    return None


def _xfrm(el):
    """(off x, off y, ext cx, ext cy, chOff x, chOff y, chExt cx, chExt cy, flipH, flipV, rot deg) or None."""
    for path in ("xdr:spPr/a:xfrm", "xdr:grpSpPr/a:xfrm", "xdr:xfrm"):
        x = el.find("/".join(_q(*p.split(":")) for p in path.split("/")))
        if x is not None:
            off, ext = x.find(_q("a", "off")), x.find(_q("a", "ext"))
            cho, che = x.find(_q("a", "chOff")), x.find(_q("a", "chExt"))
            g = lambda e, k: _int(e.get(k)) if e is not None else 0  # noqa: E731
            return (g(off, "x"), g(off, "y"), g(ext, "cx"), g(ext, "cy"), g(cho, "x"), g(cho, "y"),
                    g(che, "cx"), g(che, "cy"), x.get("flipH") in ("1", "true"), x.get("flipV") in ("1", "true"),
                    _int(x.get("rot"), 0) / 60000.0)
    return None


def _text(el) -> str:
    paras = []
    body = el.find(_q("xdr", "txBody"))
    if body is None:
        return ""
    for p in body.findall(_q("a", "p")):
        t = "".join(x.text or "" for x in p.iter(_q("a", "t")))
        if t.strip():
            paras.append(t.strip())
    return "\n".join(paras)


def _endpoints(box, flip_h: bool, flip_v: bool, rot: float, frac=(0, 0, 1, 1)):
    x1, y1, x2, y2 = box
    w, h = x2 - x1, y2 - y1
    fx1, fy1, fx2, fy2 = frac
    if flip_h:
        fx1, fx2 = 1 - fx1, 1 - fx2
    if flip_v:
        fy1, fy2 = 1 - fy1, 1 - fy2
    sx, sy, ex, ey = x1 + fx1 * w, y1 + fy1 * h, x1 + fx2 * w, y1 + fy2 * h
    if rot % 360:
        import math
        cx, cy = (x1 + x2) / 2, (y1 + y2) / 2
        t = math.radians(rot)
        rotp = lambda px, py: (cx + (px - cx) * math.cos(t) - (py - cy) * math.sin(t),  # noqa: E731
                               cy + (px - cx) * math.sin(t) + (py - cy) * math.cos(t))
        (sx, sy), (ex, ey) = rotp(sx, sy), rotp(ex, ey)
    return (sx, sy), (ex, ey)


def _arrows(el) -> tuple[bool, bool, bool]:
    """(arrow at start, arrow at end, dashed)"""
    ln = el.find(_q("xdr", "spPr") + "/" + _q("a", "ln"))
    if ln is None:
        return False, False, False
    he, te, dash = ln.find(_q("a", "headEnd")), ln.find(_q("a", "tailEnd")), ln.find(_q("a", "prstDash"))
    return ((he is not None and (he.get("type") or "none") in ARROW_HEADS),
            (te is not None and (te.get("type") or "none") in ARROW_HEADS),
            dash is not None and (dash.get("val") or "solid") != "solid")


def _walk(el, box, m: Metrics, out: list[dict], depth: int, group_id: str | None, rels: dict, pkg: Package):
    if len(out) >= MAX_SHAPES or depth > 8:
        return
    tag = _local(el.tag)
    X = lambda n: _q("xdr", n)  # noqa: E731
    xf = _xfrm(el)
    flip_h, flip_v, rot = (xf[8], xf[9], xf[10]) if xf else (False, False, 0.0)

    def base(nv_path):
        nv = el.find(nv_path)
        return {"id": (nv.get("id") or "") if nv is not None else "", "name": clean_text(nv.get("name") if nv is not None else "", 120),
                "alt": clean_text((nv.get("descr") or nv.get("title") or "") if nv is not None else "", 300),
                "box": box, "group": group_id, "from": m.cell_at(box[0] + 1, box[1] + 1),
                "to": m.cell_at(max(box[0], box[2] - 1), max(box[1], box[3] - 1))}

    if tag == "sp":
        s = base(X("nvSpPr") + "/" + X("cNvPr"))
        geo = el.find(X("spPr") + "/" + _q("a", "prstGeom"))
        prst = geo.get("prst") if geo is not None else ("custom" if el.find(X("spPr") + "/" + _q("a", "custGeom")) is not None else "")
        s["prst"] = prst or ""
        s["text"] = clean_text(_text(el).replace("\n", " / "), 1000)
        cnv = el.find(X("nvSpPr") + "/" + X("cNvSpPr"))
        s["textbox"] = cnv is not None and cnv.get("txBox") in ("1", "true")
        if prst in LINE_PRST:
            s["kind"] = "connector"
            s["start"], s["end"] = _endpoints(box, flip_h, flip_v, rot)
            s["head"], s["tail"], s["dash"] = _arrows(el)
            s["st"] = s["en"] = None
        elif prst in ARROW_PRST:
            s["kind"] = "arrowshape"
            s["start"], s["end"] = _endpoints(box, flip_h, flip_v, rot, ARROW_PRST[prst])
            both = prst in ("leftRightArrow", "upDownArrow")
            s["head"], s["tail"], s["dash"] = both, True, False
            s["st"] = s["en"] = None
        else:
            s["kind"] = "shape"
        out.append(s)
    elif tag == "cxnSp":
        s = base(X("nvCxnSpPr") + "/" + X("cNvPr"))
        cnv = el.find(X("nvCxnSpPr") + "/" + X("cNvCxnSpPr"))
        st = cnv.find(_q("a", "stCxn")) if cnv is not None else None
        en = cnv.find(_q("a", "endCxn")) if cnv is not None else None
        geo = el.find(X("spPr") + "/" + _q("a", "prstGeom"))
        s.update({"kind": "connector", "prst": geo.get("prst") if geo is not None else "", "text": "", "textbox": False,
                  "st": st.get("id") if st is not None else None, "en": en.get("id") if en is not None else None})
        s["start"], s["end"] = _endpoints(box, flip_h, flip_v, rot)
        s["head"], s["tail"], s["dash"] = _arrows(el)
        out.append(s)
    elif tag == "pic":
        s = base(X("nvPicPr") + "/" + X("cNvPr"))
        s.update({"kind": "picture", "text": "", "prst": "", "textbox": False})
        blip = el.find(".//" + _q("a", "blip"))
        rid = blip.get(_q("r", "embed")) if blip is not None else None
        if rid and rid in rels and not rels[rid][2] and rels[rid][1] in pkg.names:
            media = rels[rid][1]
            ext = media.rsplit(".", 1)[-1].lower()
            if ext in IMAGE_EXT:
                blob = pkg.z.read(media)
                s.update({"media_ext": "jpg" if ext == "jpeg" else ext, "blob": blob, "dims": image_dims(blob)})
        out.append(s)
    elif tag == "graphicFrame":
        s = base(X("nvGraphicFramePr") + "/" + X("cNvPr"))
        s.update({"kind": "frame", "text": "", "prst": "", "textbox": False})
        ch = el.find(".//" + _q("c", "chart"))
        if ch is not None:
            rid = ch.get(_q("r", "id"))
            if rid in rels and rels[rid][1] in pkg.names:
                s["kind"] = "chart"
                s["chart_part"] = rels[rid][1]
        out.append(s)
    elif tag == "grpSp":
        g = base(X("nvGrpSpPr") + "/" + X("cNvPr"))
        g.update({"kind": "group", "text": "", "prst": "", "textbox": False})
        out.append(g)
        children = [c for c in el if _local(c.tag) in ("sp", "cxnSp", "pic", "graphicFrame", "grpSp")]
        for c in children:
            cx = _xfrm(c)
            if xf and cx and xf[6] and xf[7]:
                sx = (box[2] - box[0]) / xf[6]
                sy = (box[3] - box[1]) / xf[7]
                cb = (box[0] + (cx[0] - xf[4]) * sx, box[1] + (cx[1] - xf[5]) * sy,
                      box[0] + (cx[0] - xf[4] + cx[2]) * sx, box[1] + (cx[1] - xf[5] + cx[3]) * sy)
            else:
                cb = box
            _walk(c, cb, m, out, depth + 1, g["id"], rels, pkg)


def image_dims(blob: bytes) -> tuple[int, int] | None:
    try:
        if blob[:8] == b"\x89PNG\r\n\x1a\n":
            return struct.unpack(">II", blob[16:24])
        if blob[:6] in (b"GIF87a", b"GIF89a"):
            return struct.unpack("<HH", blob[6:10])
        if blob[:2] == b"BM":
            w, h = struct.unpack("<ii", blob[18:26])
            return w, abs(h)
        if blob[:2] == b"\xff\xd8":
            i = 2
            while i + 9 < len(blob):
                if blob[i] != 0xFF:
                    return None
                marker = blob[i + 1]
                seg = struct.unpack(">H", blob[i + 2:i + 4])[0]
                if marker in (0xC0, 0xC1, 0xC2):
                    h, w = struct.unpack(">HH", blob[i + 5:i + 9])
                    return w, h
                i += 2 + seg
    except (struct.error, IndexError):
        return None
    return None


# -- charts -----------------------------------------------------------------------------------

def read_chart(pkg: Package, part: str, resolve) -> dict:
    root = pkg.xml(part)
    C = lambda n: _q("c", n)  # noqa: E731
    plot = root.find(".//" + C("plotArea"))
    types = [_local(ch.tag) for ch in plot if _local(ch.tag).endswith("Chart")] if plot is not None else []
    title_el = root.find(C("chart") + "/" + C("title"))
    title = clean_text(" ".join(t.text or "" for t in title_el.iter(_q("a", "t")))) if title_el is not None else ""

    def values(el) -> list:
        if el is None:
            return []
        pts = [(int(p.get("idx") or 0), p.findtext(C("v")) or "") for p in el.iter(C("pt"))]
        if pts:
            return [v for _, v in sorted(pts)]
        f = el.find(".//" + C("f"))
        return [str(v) if v is not None else "" for v in resolve(f.text)] if f is not None and f.text else []

    series = []
    if plot is not None:
        for ser in plot.iter(C("ser")):
            tx = ser.find(C("tx"))
            name = ""
            if tx is not None:
                v = tx.find(".//" + C("v"))
                if v is not None and v.text:
                    name = v.text
                else:
                    f = tx.find(".//" + C("f"))
                    if f is not None and f.text:
                        got = resolve(f.text)
                        name = str(got[0]) if got else ""
            cats = values(ser.find(C("cat"))) or values(ser.find(C("xVal")))
            vals = values(ser.find(C("val"))) or values(ser.find(C("yVal")))
            series.append({"name": clean_text(name, 120), "categories": [clean_text(str(c), 80) for c in cats[:200]],
                           "values": [clean_text(str(v), 40) for v in vals[:200]]})
    return {"type": ", ".join(t.replace("Chart", "") for t in types) or "chart", "title": title, "series": series}


# -- input rules ------------------------------------------------------------------------------

def read_validations(pkg: Package, sheet_part: str) -> list[dict]:
    """Both <dataValidations> and the x14 extension list (cross-sheet lists in newer Excel)."""
    root = pkg.xml(sheet_part)
    out = []
    dvs = root.find(_q("m", "dataValidations"))
    if dvs is not None:
        for dv in dvs.findall(_q("m", "dataValidation")):
            out.append({"sqref": dv.get("sqref") or "", "type": dv.get("type") or "any",
                        "op": dv.get("operator") or "", "f1": dv.findtext(_q("m", "formula1")) or "",
                        "f2": dv.findtext(_q("m", "formula2")) or "", "prompt": clean_text(dv.get("prompt") or "", 200)})
    for dv in root.iter(_q("x14", "dataValidation")):
        f1 = dv.find(_q("x14", "formula1"))
        f2 = dv.find(_q("x14", "formula2"))
        out.append({"sqref": dv.findtext(_q("xm", "sqref")) or "", "type": dv.get("type") or "any",
                    "op": dv.get("operator") or "", "f1": f1.findtext(_q("xm", "f")) if f1 is not None else "",
                    "f2": f2.findtext(_q("xm", "f")) if f2 is not None else "", "prompt": clean_text(dv.get("prompt") or "", 200)})
    return out[:500]


# -- reading a workbook -----------------------------------------------------------------------

def read_workbook(path: Path, values_wb, resolve) -> dict[str, dict]:
    """sheet title -> {"shapes": [...], "charts": [...], "validations": [...]}"""
    result: dict[str, dict] = {}
    try:
        pkg = Package(path)
    except (zipfile.BadZipFile, OSError):
        return result
    try:
        for title, part in sheet_parts(pkg):
            entry = {"shapes": [], "charts": [], "validations": []}
            try:
                entry["validations"] = read_validations(pkg, part)
            except Exception:
                pass
            if title not in values_wb.sheetnames:
                result[title] = entry
                continue
            m = Metrics(values_wb[title])
            entry["metrics"] = m
            for typ, tgt, external in pkg.rels(part).values():
                if typ != "drawing" or external or tgt not in pkg.names:
                    continue
                drels = pkg.rels(tgt)
                root = pkg.xml(tgt)
                for anchor in root:
                    box = _anchor_box(anchor, m)
                    if box is None:
                        continue
                    box = (min(box[0], box[2]), min(box[1], box[3]), max(box[0], box[2]), max(box[1], box[3]))
                    for el in anchor:
                        if _local(el.tag) in ("sp", "cxnSp", "pic", "graphicFrame", "grpSp"):
                            _walk(el, box, m, entry["shapes"], 0, None, drels, pkg)
            for s in entry["shapes"]:
                if s["kind"] == "chart":
                    try:
                        ch = read_chart(pkg, s["chart_part"], resolve)
                    except Exception:
                        ch = {"type": "chart", "title": "", "series": []}
                    ch.update({"name": s["name"], "from": s["from"], "to": s["to"], "alt": s["alt"]})
                    entry["charts"].append(ch)
            result[title] = entry
    finally:
        pkg.close()
    return result


# -- reading the drawing --------------------------------------------------------------------

def _center(b):
    return (b[0] + b[2]) / 2, (b[1] + b[3]) / 2


def _dist_to_box(px, py, b) -> float:
    dx = max(b[0] - px, 0, px - b[2])
    dy = max(b[1] - py, 0, py - b[3])
    return (dx * dx + dy * dy) ** 0.5


def analyse(shapes: list[dict], cells: list[dict], m: Metrics) -> dict:
    """Sequence diagram and connector graph from the shapes; cells supply labels written in cells.

    cells: [{"address": "D5", "text": "..."}] for non-empty cells.
    """
    lines = [s for s in shapes if s["kind"] in ("connector", "arrowshape")]
    texts = [s for s in shapes if s["kind"] == "shape" and s.get("text")]
    col, row = m.default_col, m.default_row
    tol = max(col * 0.6, 1)
    result: dict = {"sequence": None, "edges": [], "unexplained": []}
    used: set[int] = set()

    def dxdy(l):
        return abs(l["end"][0] - l["start"][0]), abs(l["end"][1] - l["start"][1])

    # cell text with its box, for labels and participant names written in cells
    cell_items = []
    for c in cells:
        mm = re.match(r"([A-Z]+)(\d+)$", c["address"])
        if not mm or not str(c.get("text", "")).strip():
            continue
        r1, c1 = int(mm.group(2)), _col_index(mm.group(1))
        if r1 > 5000 or c1 > 500:
            continue
        cell_items.append({"address": c["address"], "text": clean_text(str(c["text"]), 200), "box": m.cell_box(r1, c1)})

    # lifelines are plain vertical lines: no arrowheads and not glued to shapes at both ends
    # (a vertical arrow between two glued boxes is a flowchart step, not a lifeline)
    vert = [l for l in lines if l["kind"] == "connector" and dxdy(l)[1] > row and dxdy(l)[0] <= 0.15 * dxdy(l)[1]
            and not l["head"] and not l["tail"] and not (l.get("st") and l.get("en"))]
    horiz = [l for l in lines if dxdy(l)[0] > 0 and dxdy(l)[1] <= 0.2 * dxdy(l)[0]]

    if len(vert) >= 2 and horiz:
        lifelines: list[dict] = []
        for l in sorted(vert, key=lambda l: (l["start"][0] + l["end"][0]) / 2):
            x = (l["start"][0] + l["end"][0]) / 2
            top, bottom = min(l["start"][1], l["end"][1]), max(l["start"][1], l["end"][1])
            if lifelines and abs(lifelines[-1]["x"] - x) <= tol:
                ll = lifelines[-1]
                ll["top"], ll["bottom"] = min(ll["top"], top), max(ll["bottom"], bottom)
                ll["ids"].append(l["id"])
            else:
                lifelines.append({"x": x, "top": top, "bottom": bottom, "ids": [l["id"]]})
        if len(lifelines) >= 2:
            gaps = [b["x"] - a["x"] for a, b in zip(lifelines, lifelines[1:])]
            tol2 = max(tol, 0.25 * min(gaps)) if gaps else tol
            taken: set[str] = set()
            for i, ll in enumerate(lifelines, start=1):
                cand = [t for t in texts if t["box"][0] - tol <= ll["x"] <= t["box"][2] + tol
                        and _center(t["box"])[1] <= ll["top"] + row]
                if cand:
                    best = max(cand, key=lambda t: t["box"][3])
                    ll["name"], ll["source"] = best["text"], f"shape '{best['name']}' at {best['from']}"
                    taken.add(best["id"])
                else:
                    cc = [c for c in cell_items if c["box"][0] - tol <= ll["x"] <= c["box"][2] + tol
                          and c["box"][3] <= ll["top"] + row and ll["top"] - c["box"][3] <= 4 * row]
                    if cc:
                        best = max(cc, key=lambda c: c["box"][3])
                        ll["name"], ll["source"] = best["text"], f"cell {best['address']}"
                    else:
                        ll["name"], ll["source"] = f"L{i}", "no label found above the lifeline"
            messages = []
            for l in horiz:
                xa, xb = l["start"][0], l["end"][0]
                y = (l["start"][1] + l["end"][1]) / 2

                def nearest(x):
                    best = min(lifelines, key=lambda ll: abs(ll["x"] - x))
                    return best if abs(best["x"] - x) <= tol2 and best["top"] - row <= y <= best["bottom"] + row else None

                a, b = nearest(xa), nearest(xb)
                if not a or not b or a is b:
                    continue
                if l["tail"] and not l["head"]:
                    frm, to, sure = a, b, True
                elif l["head"] and not l["tail"]:
                    frm, to, sure = b, a, True
                else:
                    frm, to, sure = a, b, False
                label, source = (l["text"], f"text inside arrow '{l['name']}'") if l.get("text") else ("", "")
                if not label:
                    lo, hi = min(xa, xb), max(xa, xb)
                    best, best_score = None, None
                    for t in texts:
                        if t["id"] in taken:
                            continue
                        cx, cy = _center(t["box"])
                        if lo - tol <= cx <= hi + tol and y - 3 * row <= cy <= y + row:
                            score = abs(cy - y) + (row if cy > y else 0)
                            if best_score is None or score < best_score:
                                best, best_score = (t["text"], f"shape '{t['name']}' at {t['from']}"), score
                    for c in cell_items:
                        cx, cy = _center(c["box"])
                        if lo <= cx <= hi and y - 2.5 * row <= cy <= y + 0.5 * row:
                            score = abs(cy - y) + (row if cy > y else 0) + 1
                            if best_score is None or score < best_score:
                                best, best_score = (c["text"], f"cell {c['address']}"), score
                    if best:
                        label, source = best
                messages.append({"y": y, "from": frm["name"], "to": to["name"], "label": label, "label_source": source,
                                 "direction_known": sure, "at": m.cell_at(min(xa, xb) + 1, y), "line": l["name"]})
                used.add(id(l))
            if messages:
                for ll in lifelines:
                    for l in vert:
                        if l["id"] in ll["ids"]:
                            used.add(id(l))
                messages.sort(key=lambda mm: mm["y"])
                result["sequence"] = {"participants": [{"name": ll["name"], "source": ll["source"],
                                                        "at": m.cell_at(ll["x"], ll["top"] + 1)} for ll in lifelines],
                                      "messages": messages}

    # connector graph: glued ends first, otherwise the shape an end touches
    by_id = {s["id"]: s for s in shapes if s.get("id")}
    nodes = [s for s in shapes if s["kind"] == "shape"]
    for l in lines:
        if id(l) in used:
            continue

        def end_shape(glued_id, point):
            if glued_id and glued_id in by_id and by_id[glued_id]["kind"] == "shape":
                return by_id[glued_id], "glued"
            best = min(nodes, key=lambda s: _dist_to_box(point[0], point[1], s["box"]), default=None)
            if best is not None and _dist_to_box(point[0], point[1], best["box"]) <= tol * 0.5:
                return best, "touching"
            return None, ""

        a, ka = end_shape(l.get("st"), l["start"])
        b, kb = end_shape(l.get("en"), l["end"])
        if a and b and a is not b:
            if l["head"] and not l["tail"]:
                a, b, ka, kb = b, a, kb, ka
            result["edges"].append({"from": a["text"] or a["name"], "to": b["text"] or b["name"],
                                    "both": l["head"] and l["tail"], "none": not l["head"] and not l["tail"],
                                    "how": ka if ka == kb else f"{ka}/{kb}", "line": l["name"],
                                    "label": l.get("text", "")})
        else:
            result["unexplained"].append({"line": l["name"], "from": m.cell_at(*l["start"]), "to": m.cell_at(*l["end"]),
                                          "arrow": "both" if l["head"] and l["tail"] else "start" if l["head"]
                                          else "end" if l["tail"] else "none"})
    return result


def mermaid_sequence(seq: dict) -> str:
    def safe(s: str) -> str:
        return re.sub(r"[;:#`\n\r]", " ", s or "").strip()[:80] or "?"
    ids = {}
    lines = ["sequenceDiagram"]
    for i, p in enumerate(seq["participants"], start=1):
        ids[p["name"]] = f"P{i}"
        lines.append(f"    participant P{i} as {safe(p['name'])}")
    for msg in seq["messages"]:
        lines.append(f"    {ids.get(msg['from'], 'P?')}->>{ids.get(msg['to'], 'P?')}: {safe(msg['label']) if msg['label'] else '(no label)'}")
    return "\n".join(lines)


def mermaid_flow(edges: list[dict]) -> str:
    def safe(s: str) -> str:
        return re.sub(r'["`\[\]\n\r]', " ", s or "").strip()[:60] or "?"
    ids: dict[str, str] = {}
    lines = ["flowchart TD"]
    for e in edges:
        for k in ("from", "to"):
            if e[k] not in ids:
                ids[e[k]] = f"N{len(ids) + 1}"
                lines.append(f'    {ids[e[k]]}["{safe(e[k])}"]')
    for e in edges:
        arrow = "<-->" if e["both"] else "---" if e["none"] else "-->"
        label = f"|{safe(e['label'])}|" if e.get("label") else ""
        lines.append(f"    {ids[e['from']]} {arrow}{label} {ids[e['to']]}")
    return "\n".join(lines)
