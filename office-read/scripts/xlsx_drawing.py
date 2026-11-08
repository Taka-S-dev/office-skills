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
              "rightArrowCallout": (0, 0.5, 1, 0.5), "leftArrowCallout": (1, 0.5, 0, 0.5)}
# chevron / homePlate are left out on purpose: in specifications they are step labels and the
# tag of a UML frame ("alt"), not arrows
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


def _paint(el) -> tuple[str, str]:
    """(fill, line) of a shape: "none", "#RRGGBB", or "default" (theme colour, drawn as the
    editor's default). A shape with no fill and no theme style is transparent in Excel, and a
    transcription must keep it so, or a frame drawn over a diagram hides what is inside it."""
    sp = el.find(_q("xdr", "spPr"))
    style = el.find(_q("xdr", "style"))

    def color(node):
        if node is None:
            return None
        if node.find(_q("a", "noFill")) is not None:
            return "none"
        sf = node.find(_q("a", "solidFill"))
        if sf is not None:
            c = sf.find(_q("a", "srgbClr"))
            v = c.get("val") if c is not None else ""
            return f"#{v.upper()}" if re.fullmatch(r"[0-9A-Fa-f]{6}", v or "") else "default"
        return None

    def themed(ref_name):
        ref = style.find(_q("a", ref_name)) if style is not None else None
        return ref is not None and (ref.get("idx") or "0") != "0"

    fill = color(sp)
    if fill is None:
        fill = "default" if themed("fillRef") else "none"
    ln = sp.find(_q("a", "ln")) if sp is not None else None
    line = color(ln)
    if line is None:
        line = "default" if (ln is not None or themed("lnRef")) else "none"
    return fill, line


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
        s["fill"], s["line"] = _paint(el)
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


FRAME_KW = re.compile(r"^\s*\[?\s*(alt|opt|loop|par|break|critical|neg|ref|seq|strict|assert|ignore|consider|group)\b"
                      r"\s*\]?\s*(.*)$", re.I)


def _join_segments(lines: list[dict], eps: float, keep_out: set[int]) -> list[dict]:
    """Elbow arrows drawn as two or three straight lines become one path. Only unambiguous joints
    (exactly two endpoints meeting) are joined; glued connectors and lifelines are left alone."""
    segs = [l for l in lines if l["kind"] == "connector" and not l.get("st") and not l.get("en")
            and id(l) not in keep_out]
    if len(segs) < 2 or len(segs) > 600:
        return lines
    buckets: dict[tuple[int, int], list[tuple[int, int]]] = {}
    for i, l in enumerate(segs):
        for w, (x, y) in enumerate((l["start"], l["end"])):
            buckets.setdefault((int(x // eps), int(y // eps)), []).append((i, w))

    def near(i, w):
        x, y = segs[i]["start"] if w == 0 else segs[i]["end"]
        bx, by = int(x // eps), int(y // eps)
        found = []
        for dx in (-1, 0, 1):
            for dy in (-1, 0, 1):
                for j, wj in buckets.get((bx + dx, by + dy), []):
                    if j == i:
                        continue
                    xj, yj = segs[j]["start"] if wj == 0 else segs[j]["end"]
                    if abs(xj - x) <= eps and abs(yj - y) <= eps:
                        found.append((j, wj))
        return found

    joint = {}
    for i in range(len(segs)):
        for w in (0, 1):
            f = near(i, w)
            if len(f) == 1 and len(near(*f[0])) == 1:
                joint[(i, w)] = f[0]
    if not joint:
        return lines
    seen: set[int] = set()
    merged: list[dict] = []
    for i in range(len(segs)):
        if i in seen or ((i, 0) in joint and (i, 1) in joint):
            continue                          # start walking from a free end only
        free = 0 if (i, 0) not in joint else 1
        path, cur, enter_free = [], i, free
        while cur is not None and cur not in seen:
            seen.add(cur)
            exit_w = 1 - enter_free
            path.append((cur, enter_free))
            nxt = joint.get((cur, exit_w))
            if nxt is None:
                break
            cur, enter_free = nxt[0], nxt[1]
        if len(path) < 2:
            continue
        (fi, fw), (li, lw) = path[0], path[-1]
        first, last = segs[fi], segs[li]
        start = first["start"] if fw == 0 else first["end"]
        last_free = 1 - lw
        end = last["end"] if last_free == 1 else last["start"]
        head = first["head"] if fw == 0 else first["tail"]
        tail = last["tail"] if last_free == 1 else last["head"]
        xs = [c for k, _ in path for c in (segs[k]["box"][0], segs[k]["box"][2])]
        ys = [c for k, _ in path for c in (segs[k]["box"][1], segs[k]["box"][3])]
        merged.append({"kind": "connector", "id": first["id"], "name": " + ".join(segs[k]["name"] for k, _ in path),
                       "start": start, "end": end, "head": head, "tail": tail, "dash": first.get("dash", False),
                       "st": None, "en": None, "text": "", "box": (min(xs), min(ys), max(xs), max(ys)),
                       "from": first["from"], "to": last["to"], "parts": [segs[k]["id"] for k, _ in path]})
        for k, _ in path:
            segs[k]["_joined"] = True
    if not merged:
        return lines
    return [l for l in lines if not l.get("_joined")] + merged


def analyse(shapes: list[dict], cells: list[dict], m: Metrics) -> dict:
    """Sequence diagram and connector graph from the shapes; cells supply labels written in cells.

    cells: [{"address": "D5", "text": "..."}] for non-empty cells.
    Recognised in a sequence: lifelines, participants, messages (solid; dashed = reply), self
    messages, activation bars, frames (alt / opt / loop / par / critical / break ... with else
    separators and [conditions]) and notes. Each is recognised only from an unambiguous cue;
    anything else is reported as not interpreted.
    """
    col, row = m.default_col, m.default_row
    tol = max(col * 0.6, 1)
    lines = [s for s in shapes if s["kind"] in ("connector", "arrowshape")]
    texts = [s for s in shapes if s["kind"] == "shape" and s.get("text")]
    result: dict = {"sequence": None, "edges": [], "unexplained": []}
    used: set[int] = set()

    def dxdy(l):
        return abs(l["end"][0] - l["start"][0]), abs(l["end"][1] - l["start"][1])

    cell_items = []
    for c in cells:
        mm = re.match(r"([A-Z]+)(\d+)$", c["address"])
        if not mm or not str(c.get("text", "")).strip():
            continue
        r1, c1 = int(mm.group(2)), _col_index(mm.group(1))
        if r1 > 5000 or c1 > 500:
            continue
        cell_items.append({"address": c["address"], "text": clean_text(str(c["text"]), 200), "box": m.cell_box(r1, c1)})

    # lifelines: long plain vertical lines, no arrowheads, not glued at both ends
    vert = [l for l in lines if l["kind"] == "connector" and dxdy(l)[1] > 2 * row and dxdy(l)[0] <= 0.15 * dxdy(l)[1]
            and not l["head"] and not l["tail"] and not (l.get("st") and l.get("en"))]
    lines = _join_segments(lines, max(col * 0.12, 1), {id(l) for l in vert})

    seq = _sequence(shapes, lines, vert, texts, cell_items, m, col, row, tol, used)
    if seq:
        result["sequence"] = seq

    by_id = {s["id"]: s for s in shapes if s.get("id")}
    nodes = [s for s in shapes if s["kind"] == "shape" and id(s) not in used]
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


def _sequence(shapes, lines, vert, texts, cell_items, m, col, row, tol, used) -> dict | None:
    if len(vert) < 2:
        return None
    lifelines: list[dict] = []
    for l in sorted(vert, key=lambda l: (l["start"][0] + l["end"][0]) / 2):
        x = (l["start"][0] + l["end"][0]) / 2
        top, bottom = min(l["start"][1], l["end"][1]), max(l["start"][1], l["end"][1])
        if lifelines and abs(lifelines[-1]["x"] - x) <= tol:
            ll = lifelines[-1]
            ll["top"], ll["bottom"] = min(ll["top"], top), max(ll["bottom"], bottom)
            ll["lines"].append(l)
        else:
            lifelines.append({"x": x, "top": top, "bottom": bottom, "lines": [l]})
    if len(lifelines) < 2:
        return None
    gaps = [b["x"] - a["x"] for a, b in zip(lifelines, lifelines[1:])]
    tol2 = max(tol, 0.25 * min(gaps))
    taken: set[str] = set()

    for i, ll in enumerate(lifelines, start=1):
        ll["key"] = f"P{i}"
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

    def lifeline_at(x, y):
        best = min(lifelines, key=lambda ll: abs(ll["x"] - x))
        return best if abs(best["x"] - x) <= tol2 and best["top"] - row <= y <= best["bottom"] + row else None

    # activation bars: narrow tall boxes centred on a lifeline
    acts = []
    for s in shapes:
        if s["kind"] != "shape" or len(s.get("text") or "") > 3:
            continue
        w, h = s["box"][2] - s["box"][0], s["box"][3] - s["box"][1]
        cx, _ = _center(s["box"])
        if 0 < w <= 0.8 * col and h >= 0.8 * row:
            ll = min(lifelines, key=lambda ll: abs(ll["x"] - cx))
            if abs(ll["x"] - cx) <= tol and ll["top"] - row <= s["box"][1] <= ll["bottom"]:
                acts.append({"ll": ll, "y1": s["box"][1], "y2": s["box"][3], "at": s["from"]})
                taken.add(s["id"])
                used.add(id(s))

    # frames: a box with an interaction keyword in it or in a tag at its top-left corner
    frames = []
    for s in shapes:
        if s["kind"] != "shape" or s["id"] in taken:
            continue
        b = s["box"]
        w, h = b[2] - b[0], b[3] - b[1]
        if w < 1.5 * col or h < 1.5 * row:
            continue
        kw, cond, tags = None, "", []
        mm = FRAME_KW.match(s.get("text") or "")
        if mm:
            kw, cond = mm.group(1).lower(), mm.group(2)
        else:
            for t in texts:
                if t["id"] in taken or t is s:
                    continue
                cx, cy = _center(t["box"])
                if b[0] - 0.2 * col <= t["box"][0] <= b[0] + 0.5 * w and b[1] - 0.5 * row <= cy <= b[1] + 2 * row:
                    mt = FRAME_KW.match(t["text"])
                    if mt:
                        kw, cond = mt.group(1).lower(), mt.group(2)
                        tags.append(t["id"])
                        break
        if not kw:
            continue
        cond = cond.strip()
        if not cond:
            for t in texts:
                if t["id"] in taken or t["id"] in tags or t is s:
                    continue
                if t["text"].startswith("[") and b[0] <= t["box"][0] <= b[2] and b[1] - 0.5 * row <= t["box"][1] <= b[1] + 2.5 * row:
                    cond = t["text"]
                    tags.append(t["id"])
                    break
        frames.append({"kw": kw, "cond": cond.strip("[] "), "box": b, "elses": [], "at": s["from"]})
        taken.add(s["id"])
        taken.update(tags)
        used.add(id(s))
    # frames that overlap without one containing the other cannot be written as nested blocks
    frames.sort(key=lambda f: (f["box"][1], -(f["box"][3] - f["box"][1])))
    keep = []
    for f in frames:
        ok = True
        for g in keep:
            fb, gb = f["box"], g["box"]
            overlap = not (fb[3] <= gb[1] or gb[3] <= fb[1])
            contained = gb[1] <= fb[1] and fb[3] <= gb[3]
            if overlap and not contained:
                ok = False
        if ok:
            keep.append(f)
    dropped_frames = [f for f in frames if f not in keep]
    frames = keep

    # else separators: dashed lines without arrowheads across most of a frame
    for l in lines:
        if not l.get("dash") or l["head"] or l["tail"]:
            continue
        dx, dy = abs(l["end"][0] - l["start"][0]), abs(l["end"][1] - l["start"][1])
        if dx == 0 or dy > 0.2 * dx:
            continue
        y = (l["start"][1] + l["end"][1]) / 2
        inner = [f for f in frames if f["box"][1] < y < f["box"][3] and dx >= 0.6 * (f["box"][2] - f["box"][0])]
        if not inner:
            continue
        f = min(inner, key=lambda f: f["box"][3] - f["box"][1])
        cond = ""
        for t in texts:
            if t["id"] in taken:
                continue
            if t["text"].startswith("[") and f["box"][0] <= t["box"][0] <= (f["box"][0] + f["box"][2]) / 2 \
                    and y - row <= t["box"][1] <= y + 1.5 * row:
                cond = t["text"]
                taken.add(t["id"])
                break
        f["elses"].append({"y": y, "cond": cond.strip("[] ")})
        used.add(id(l))

    # messages
    msgs = []
    label_ids: set[str] = set()

    def find_label(lo, hi, y_lo, y_hi, y_ref):
        best, best_score = None, None
        for t in texts:
            if t["id"] in taken or t["id"] in label_ids:
                continue
            cx, cy = _center(t["box"])
            if lo - tol <= cx <= hi + tol and y_lo <= cy <= y_hi:
                score = abs(cy - y_ref) + (row if cy > y_ref else 0)
                if best_score is None or score < best_score:
                    best, best_score = (t["text"], f"shape '{t['name']}' at {t['from']}", t["id"]), score
        for c in cell_items:
            cx, cy = _center(c["box"])
            if lo <= cx <= hi and y_lo + 0.5 * row <= cy <= y_hi - 0.5 * row:
                score = abs(cy - y_ref) + (row if cy > y_ref else 0) + 1
                if best_score is None or score < best_score:
                    best, best_score = (c["text"], f"cell {c['address']}", None), score
        return best

    for l in lines:
        if id(l) in used or any(l is v for ll in lifelines for v in ll["lines"]):
            continue
        if l.get("dash") and not l["head"] and not l["tail"]:
            continue
        xa, ya = l["start"]
        xb, yb = l["end"]
        dx, dy = abs(xb - xa), abs(yb - ya)
        a, b = lifeline_at(xa, ya), lifeline_at(xb, yb)
        if a is None or b is None:
            continue
        self_msg = a is b
        if self_msg:
            if not (l["head"] or l["tail"]) or not (0.2 * row <= dy <= 4 * row):
                continue
            y = min(ya, yb)
            lo, hi, y_lo, y_hi = a["x"], a["x"] + 3 * col, y - row, max(ya, yb) + row
        else:
            if dx == 0 or dy > 0.2 * dx:
                continue
            y = (ya + yb) / 2
            lo, hi, y_lo, y_hi = min(xa, xb), max(xa, xb), y - 3 * row, y + row
        if l["tail"] and not l["head"]:
            frm, to, sure = a, b, True
        elif l["head"] and not l["tail"]:
            frm, to, sure = b, a, True
        else:
            frm, to, sure = a, b, False
        label, source = (l["text"], f"text inside arrow '{l['name']}'") if l.get("text") else ("", "")
        if not label:
            got = find_label(lo, hi, y_lo, y_hi, y)
            if got:
                label, source, tid = got
                if tid:
                    label_ids.add(tid)
        msgs.append({"y": y, "from": frm["name"], "to": to["name"], "from_key": frm["key"], "to_key": to["key"],
                     "label": label, "label_source": source, "direction_known": sure, "reply": bool(l.get("dash")),
                     "self": self_msg, "at": m.cell_at(min(xa, xb) + 1, y), "line": l["name"]})
        used.add(id(l))
    if not msgs:
        return None
    for ll in lifelines:
        for l in ll["lines"]:
            used.add(id(l))

    # notes: remaining text inside the diagram area
    notes = []
    xmin = lifelines[0]["x"] - 2 * col
    xmax = lifelines[-1]["x"] + 2 * col
    ytop = min(ll["top"] for ll in lifelines)
    ybot = max(ll["bottom"] for ll in lifelines)
    for t in texts:
        if t["id"] in taken or t["id"] in label_ids:
            continue
        cx, cy = _center(t["box"])
        if not (xmin <= cx <= xmax and ytop <= cy <= ybot):
            continue
        over = [ll for ll in lifelines if t["box"][0] <= ll["x"] <= t["box"][2]]
        if over:
            where, keys = "over", [over[0], over[-1]] if len(over) > 1 else [over[0]]
        else:
            ll = min(lifelines, key=lambda ll: abs(ll["x"] - cx))
            where, keys = ("right of" if cx > ll["x"] else "left of"), [ll]
        notes.append({"y": t["box"][1], "where": where, "keys": [k["key"] for k in keys],
                      "names": [k["name"] for k in keys], "text": t["text"], "at": t["from"]})
        taken.add(t["id"])

    events = []
    for f in frames:
        events.append((f["box"][1], 0, "start", f))
        for e in f["elses"]:
            events.append((e["y"], 1, "else", {"frame": f, **e}))
        events.append((f["box"][3], 4, "end", f))
    for mm in msgs:
        events.append((mm["y"], 2, "msg", mm))
    for n in notes:
        events.append((n["y"], 2, "note", n))
    for a in acts:
        events.append((a["y1"], 3, "act", {"key": a["ll"]["key"], "name": a["ll"]["name"], "at": a["at"]}))
        events.append((a["y2"], 3.5, "deact", {"key": a["ll"]["key"], "name": a["ll"]["name"], "at": a["at"]}))
    events.sort(key=lambda e: (e[0], e[1]))
    return {"participants": [{"key": ll["key"], "name": ll["name"], "source": ll["source"],
                              "at": m.cell_at(ll["x"], ll["top"] + 1)} for ll in lifelines],
            "messages": [mm for mm in msgs],
            "events": [{"type": e[2], **{k: v for k, v in e[3].items() if k not in ("box", "elses", "frame")},
                        **({"kw": e[3]["frame"]["kw"]} if e[2] == "else" else {}),
                        **({"kw": e[3]["kw"], "cond": e[3]["cond"]} if e[2] in ("start", "end") else {})}
                       for e in events],
            "dropped_frames": [{"kw": f["kw"], "at": f["at"]} for f in dropped_frames]}


# -- writing it out ---------------------------------------------------------------------------

MERMAID_BLOCKS = {"alt": ("alt", "else"), "opt": ("opt", None), "loop": ("loop", None), "par": ("par", "and"),
                  "critical": ("critical", "option"), "break": ("break", None)}
PLANTUML_BLOCKS = {"alt": ("alt", "else"), "opt": ("opt", None), "loop": ("loop", None), "par": ("par", "else"),
                   "critical": ("critical", "else"), "break": ("break", None), "group": ("group", "else")}


def _safe_mermaid(s: str, limit: int = 80) -> str:
    return re.sub(r"[;#`\n\r]", " ", s or "").replace(":", "：").strip()[:limit]


def mermaid_sequence(seq: dict) -> str:
    out = ["sequenceDiagram"]
    for p in seq["participants"]:
        out.append(f"    participant {p['key']} as {_safe_mermaid(p['name']) or p['key']}")
    depth = 1
    for e in seq.get("events") or [{"type": "msg", **m} for m in seq["messages"]]:
        pad = "    " * depth
        t = e["type"]
        if t == "msg":
            arrow = "-->>" if e.get("reply") else "->>"
            out.append(f"{pad}{e['from_key']}{arrow}{e['to_key']}: {_safe_mermaid(e['label']) or '(no label)'}")
        elif t == "note":
            out.append(f"{pad}Note {e['where']} {','.join(e['keys'])}: {_safe_mermaid(e['text'])}")
        elif t == "act":
            out.append(f"{pad}activate {e['key']}")
        elif t == "deact":
            out.append(f"{pad}deactivate {e['key']}")
        elif t == "start":
            kw, _ = MERMAID_BLOCKS.get(e["kw"], ("rect", None))
            if kw == "rect":
                out.append(f"{pad}rect rgba(128,128,128,0.1)")
                out.append(f"{pad}    Note over {seq['participants'][0]['key']},{seq['participants'][-1]['key']}: "
                           f"{e['kw']} {_safe_mermaid(e['cond'])}".rstrip())
            else:
                out.append(f"{pad}{kw} {_safe_mermaid(e['cond'])}".rstrip())
            depth += 1
        elif t == "else":
            sep = MERMAID_BLOCKS.get(e["kw"], (None, None))[1]
            if sep:
                out.append(f"{'    ' * (depth - 1)}{sep} {_safe_mermaid(e['cond'])}".rstrip())
        elif t == "end":
            depth = max(1, depth - 1)
            out.append(f"{'    ' * depth}end")
    return "\n".join(out)


def plantuml_sequence(seq: dict) -> str:
    def safe(s: str, limit: int = 80) -> str:
        # text from the file must stay text: % starts a preprocessor builtin (%getenv ...),
        # <img:...> and <&...> are creole that loads resources, a leading ! is a directive
        s = re.sub(r"[\n\r]", " ", s or "").replace('"', "'").replace("%", "％").replace("<", "＜").replace(">", "＞")
        return s.lstrip("!").strip()[:limit]
    out = ["@startuml"]
    for p in seq["participants"]:
        out.append(f'participant "{safe(p["name"]) or p["key"]}" as {p["key"]}')
    depth = 0
    for e in seq.get("events") or [{"type": "msg", **m} for m in seq["messages"]]:
        pad = "  " * depth
        t = e["type"]
        if t == "msg":
            arrow = "-->" if e.get("reply") else "->"
            out.append(f"{pad}{e['from_key']} {arrow} {e['to_key']} : {safe(e['label']) or '(no label)'}")
        elif t == "note":
            out.append(f"{pad}note {e['where']} {', '.join(e['keys'])} : {safe(e['text'])}")
        elif t == "act":
            out.append(f"{pad}activate {e['key']}")
        elif t == "deact":
            out.append(f"{pad}deactivate {e['key']}")
        elif t == "start":
            kw = PLANTUML_BLOCKS.get(e["kw"], ("group", None))[0]
            label = f"{e['kw']} {safe(e['cond'])}".strip() if kw == "group" and e["kw"] != "group" else safe(e["cond"])
            out.append(f"{pad}{kw} {label}".rstrip())
            depth += 1
        elif t == "else":
            sep = PLANTUML_BLOCKS.get(e["kw"], ("group", "else"))[1]
            if sep:
                out.append(f"{'  ' * (depth - 1)}{sep} {safe(e['cond'])}".rstrip())
        elif t == "end":
            depth = max(0, depth - 1)
            out.append(f"{'  ' * depth}end")
    out.append("@enduml")
    return "\n".join(out)


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


def plantuml_flow(edges: list[dict]) -> str:
    def safe(s: str) -> str:
        s = re.sub(r'["\n\r]', " ", s or "").replace("%", "％").replace("<", "＜").replace(">", "＞")
        return s.lstrip("!").strip()[:60] or "?"
    ids: dict[str, str] = {}
    out = ["@startuml"]
    for e in edges:
        for k in ("from", "to"):
            if e[k] not in ids:
                ids[e[k]] = f"N{len(ids) + 1}"
                out.append(f'rectangle "{safe(e[k])}" as {ids[e[k]]}')
    for e in edges:
        arrow = "<-->" if e["both"] else "--" if e["none"] else "-->"
        label = f" : {safe(e['label'])}" if e.get("label") else ""
        out.append(f"{ids[e['from']]} {arrow} {ids[e['to']]}{label}")
    out.append("@enduml")
    return "\n".join(out)


# -- draw.io ----------------------------------------------------------------------------------
# A transcription, not an interpretation: every shape, line and picture keeps its position, size,
# outline, arrowheads and glue, so the drawing can be edited in draw.io exactly as it was drawn.
# Text written in cells inside the drawing area comes along as text, since diagrams in workbooks
# often label arrows in cells. Labels are plain text (html=0): nothing from the file is read as HTML.

DRAWIO_SHAPE = {"rect": "", "roundRect": "rounded=1;", "ellipse": "ellipse;", "diamond": "rhombus;",
                "flowChartDecision": "rhombus;", "flowChartProcess": "", "flowChartTerminator": "rounded=1;arcSize=50;",
                "parallelogram": "shape=parallelogram;perimeter=parallelogramPerimeter;",
                "flowChartInputOutput": "shape=parallelogram;perimeter=parallelogramPerimeter;",
                "can": "shape=cylinder3;", "flowChartMagneticDisk": "shape=cylinder3;",
                "triangle": "triangle;", "hexagon": "shape=hexagon;perimeter=hexagonPerimeter2;",
                "foldedCorner": "shape=note;", "cloud": "ellipse;shape=cloud;",
                "homePlate": "shape=step;perimeter=stepPerimeter;fixedSize=1;",
                "chevron": "shape=step;perimeter=stepPerimeter;fixedSize=1;"}


def _x(s) -> str:
    from xml.sax.saxutils import escape
    return escape(str(s), {'"': "&quot;", "\n": "&#10;"})


def _paint_style(s: dict) -> str:
    out = ""
    fill, line = s.get("fill"), s.get("line")
    if fill == "none":
        out += "fillColor=none;"
    elif fill and fill.startswith("#"):
        out += f"fillColor={fill};"
    if line == "none":
        out += "strokeColor=none;"
    elif line and line.startswith("#"):
        out += f"strokeColor={line};"
    return out


def drawio_page(name: str, shapes: list[dict], cells: list[dict], m: Metrics, page_id: str,
                charts: list[dict] | None = None) -> str:
    px = lambda v: round(v / EMU_PER_PX, 1)  # noqa: E731
    out = []
    ids = {}
    n = 0

    def nid():
        nonlocal n
        n += 1
        return f"{page_id}-{n}"

    drawn = [s for s in shapes if s["kind"] != "group"]
    if not drawn:
        return ""
    x1 = min(s["box"][0] for s in drawn) - m.default_col
    y1 = min(s["box"][1] for s in drawn) - m.default_row
    x2 = max(s["box"][2] for s in drawn) + m.default_col
    y2 = max(s["box"][3] for s in drawn) + m.default_row
    for s in drawn:
        b = s["box"]
        geo = f'<mxGeometry x="{px(b[0])}" y="{px(b[1])}" width="{max(px(b[2] - b[0]), 1)}" height="{max(px(b[3] - b[1]), 1)}" as="geometry"/>'
        if s["kind"] == "shape":
            cid = nid()
            ids[s.get("id")] = cid
            if s.get("textbox"):
                style = "text;html=0;align=left;verticalAlign=top;strokeColor=none;fillColor=none;"
            else:
                style = DRAWIO_SHAPE.get(s.get("prst") or "rect", "") + "html=0;" + _paint_style(s)
            out.append(f'<mxCell id="{cid}" value="{_x(s.get("text", ""))}" style="{style}" vertex="1" parent="1">{geo}</mxCell>')
        elif s["kind"] in ("connector", "arrowshape"):
            cid = nid()
            style = "html=0;rounded=0;" + ("dashed=1;" if s.get("dash") else "")
            style += f"startArrow={'block' if s.get('head') else 'none'};endArrow={'block' if s.get('tail') else 'none'};"
            if s["kind"] == "arrowshape":
                style += "strokeWidth=3;"
            if (s.get("prst") or "").startswith("bentConnector"):
                style += "edgeStyle=orthogonalEdgeStyle;"
            src = f' source="{ids[s["st"]]}"' if s.get("st") in ids else ""
            tgt = f' target="{ids[s["en"]]}"' if s.get("en") in ids else ""
            (sx, sy), (ex, ey) = s["start"], s["end"]
            out.append(f'<mxCell id="{cid}" value="{_x(s.get("text", ""))}" style="{style}" edge="1" parent="1"{src}{tgt}>'
                       f'<mxGeometry relative="1" as="geometry"><mxPoint x="{px(sx)}" y="{px(sy)}" as="sourcePoint"/>'
                       f'<mxPoint x="{px(ex)}" y="{px(ey)}" as="targetPoint"/></mxGeometry></mxCell>')
        elif s["kind"] == "picture":
            cid = nid()
            blob = s.get("blob")
            if blob and len(blob) <= 1_000_000 and s.get("media_ext") in ("png", "jpg", "gif", "bmp", "svg", "webp"):
                import base64
                mime = {"jpg": "jpeg", "svg": "svg+xml"}.get(s["media_ext"], s["media_ext"])
                style = f"shape=image;aspect=fixed;image=data:image/{mime},{base64.b64encode(blob).decode()};"
                out.append(f'<mxCell id="{cid}" value="" style="{style}" vertex="1" parent="1">{geo}</mxCell>')
            else:
                out.append(f'<mxCell id="{cid}" value="{_x("[picture: " + s.get("name", "") + "]")}" '
                           f'style="rounded=0;html=0;dashed=1;" vertex="1" parent="1">{geo}</mxCell>')
        elif s["kind"] in ("chart", "frame"):
            cid = nid()
            label = f"[{s['kind']}: {s.get('name', '')}]"
            ch = next((c for c in (charts or []) if c.get("name") == s.get("name") and c.get("from") == s.get("from")), None)
            if ch:
                # draw.io cannot draw an Excel chart, so its placeholder carries the data instead
                lines = [f"{label} {ch.get('title') or ''}".strip(), f"({ch.get('type', 'chart')})"]
                for se in ch.get("series", [])[:8]:
                    pairs = ", ".join(f"{c}={v}" for c, v in zip(se["categories"], se["values"])) if se["categories"] \
                        else ", ".join(se["values"])
                    lines.append(f"{se['name'] or 'series'}: {pairs}"[:300])
                label = "\n".join(lines)
            out.append(f'<mxCell id="{cid}" value="{_x(label)}" '
                       f'style="rounded=0;html=0;dashed=1;" vertex="1" parent="1">{geo}</mxCell>')
    for c in cells:
        mm = re.match(r"([A-Z]+)(\d+)$", c["address"])
        if not mm or not str(c.get("text", "")).strip():
            continue
        r1, c1 = int(mm.group(2)), _col_index(mm.group(1))
        if r1 > 5000 or c1 > 500:
            continue
        b = m.cell_box(r1, c1)
        if b[0] >= x1 and b[2] <= x2 and b[1] >= y1 and b[3] <= y2:
            cid = nid()
            out.append(f'<mxCell id="{cid}" value="{_x(c["text"])}" style="text;html=0;align=left;verticalAlign=middle;'
                       f'strokeColor=none;fillColor=none;whiteSpace=nowrap;" vertex="1" parent="1">'
                       f'<mxGeometry x="{px(b[0])}" y="{px(b[1])}" width="{px(b[2] - b[0])}" height="{px(b[3] - b[1])}" as="geometry"/></mxCell>')
    return (f'<diagram id="{page_id}" name="{_x(name)}"><mxGraphModel dx="1200" dy="800" grid="1" gridSize="10" guides="1" '
            f'page="1" pageScale="1" math="0" shadow="0"><root><mxCell id="0"/><mxCell id="1" parent="0"/>'
            + "".join(out) + "</root></mxGraphModel></diagram>")


def drawio_file(pages: list[str]) -> str:
    return '<mxfile host="office2md">' + "".join(p for p in pages if p) + "</mxfile>\n"
