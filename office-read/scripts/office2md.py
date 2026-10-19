#!/usr/bin/env python3
"""Dump Office documents (xlsx / docx / pptx / pdf / csv) as Markdown.

Usage:
    office2md.py FILE                      full content as Markdown
    office2md.py FILE --outline            structure only (sheets / headings / slide titles)
    office2md.py FILE --sheet "Sales"      one sheet (xlsx)         --slide 3   one slide (pptx)
    office2md.py FILE --max-rows 0         no row limit (default 200 rows per sheet)
    office2md.py FILE --formulas           show formulas instead of cached values (xlsx)
    office2md.py FILE --cells              one line per non-empty cell: Sheet!C4 = value (xlsx)
    office2md.py FILE --json               structured JSON instead of Markdown
    office2md.py FILE --out result.md      write to a file instead of stdout

Adding a format: subclass Handler, implement outline() and dump(), and register
the extensions in HANDLERS at the bottom. Nothing else needs to change.
"""

from __future__ import annotations

import argparse
import csv
import io
import json
import re
import sys
import zipfile
from pathlib import Path
from typing import Any

try:
    sys.stdout.reconfigure(encoding="utf-8")  # Windows console defaults to cp932
except AttributeError:
    pass

# ----------------------------------------------------------------------------
# Shared helpers
# ----------------------------------------------------------------------------

DEFAULT_MAX_ROWS = 200
DEFAULT_MAX_COLS = 50


def md_escape(value: Any) -> str:
    """Make a cell safe for a Markdown table."""
    if value is None:
        return ""
    text = str(value)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = text.replace("\n", "<br>")
    text = text.replace("|", "\\|")
    return text


def md_table(rows: list[list[Any]], header: list[Any] | None = None) -> str:
    """Render rows as a Markdown table. Pads ragged rows."""
    if not rows and not header:
        return "_(empty)_"
    width = max(len(r) for r in ([header] if header else []) + rows)
    if header is None:
        header, rows = rows[0], rows[1:]
    def line(r: list[Any]) -> str:
        cells = [md_escape(c) for c in r] + [""] * (width - len(r))
        return "| " + " | ".join(cells) + " |"
    out = [line(header), "|" + "---|" * width]
    out += [line(r) for r in rows]
    return "\n".join(out)


def etree_local(tag) -> str:
    """'{ns}name' -> 'name' for lxml tags."""
    return tag.rsplit("}", 1)[-1] if isinstance(tag, str) else ""


def need(module: str, pip_name: str):
    try:
        return __import__(module)
    except ImportError:
        sys.exit(f"[office2md] missing library '{module}'. Install with: pip install {pip_name}")


class Handler:
    """Base class. Subclasses fill `outline()` and `dump()`."""

    name = "generic"

    def __init__(self, path: Path, opts: argparse.Namespace):
        self.path = path
        self.opts = opts

    def outline(self) -> dict:
        raise NotImplementedError

    def dump(self) -> dict:
        raise NotImplementedError

    # Markdown rendering lives here so JSON and Markdown stay in sync.
    def to_markdown(self, data: dict) -> str:
        raise NotImplementedError


# ----------------------------------------------------------------------------
# xlsx / xlsm / xltx
# ----------------------------------------------------------------------------

class XlsxHandler(Handler):
    name = "xlsx"

    def _open(self, data_only: bool):
        openpyxl = need("openpyxl", "openpyxl")
        return openpyxl.load_workbook(self.path, data_only=data_only, read_only=False)

    def _sheet_info(self, ws, fws=None) -> dict:
        fws = fws if fws is not None else ws
        return {
            "name": ws.title,
            "state": ws.sheet_state,            # visible / hidden / veryHidden
            "dimensions": ws.dimensions,
            "max_row": ws.max_row,
            "max_col": ws.max_column,
            "merged_ranges": [str(r) for r in ws.merged_cells.ranges],
            "freeze_panes": str(ws.freeze_panes) if ws.freeze_panes else None,
            "tables": list(ws.tables.keys()) if hasattr(ws, "tables") else [],
            "formula_cells": sum(1 for row in fws.iter_rows() for c in row
                                 if isinstance(c.value, str) and c.value.startswith("=")),
            "objects": {
                "data_validations": len(ws.data_validations.dataValidation) if ws.data_validations else 0,
                "conditional_formats": len(ws.conditional_formatting),
            },
        }

    def outline(self) -> dict:
        wb = self._open(data_only=False)  # formulas visible, so they can be counted
        info = {
            "file": str(self.path),
            "type": "xlsx",
            "active_sheet": wb.active.title if wb.active else None,
            "sheets": [self._sheet_info(ws) for ws in wb.worksheets],
            "defined_names": sorted(wb.defined_names.keys()) if hasattr(wb.defined_names, "keys") else [],
            "package_objects": self._package_objects(),
        }
        return info

    def _package_objects(self) -> list[str]:
        """Parts in the zip that openpyxl does not round-trip. Names only; cheap and reliable."""
        markers = (
            ("xl/charts/", "charts"), ("xl/media/", "images"), ("xl/pivotTables/", "pivot tables"),
            ("xl/slicers/", "slicers"), ("xl/vbaProject.bin", "VBA macros"),
            ("xl/drawings/", "drawings/shapes"), ("xl/ctrlProps/", "form controls"),
            ("xl/queryTables/", "external data"), ("xl/threadedComments/", "threaded comments"),
        )
        hits = set()
        try:
            with zipfile.ZipFile(self.path) as z:
                for n in z.namelist():
                    for prefix, label in markers:
                        if n.startswith(prefix):
                            hits.add(label)
        except zipfile.BadZipFile:
            pass
        return sorted(hits)

    # -- drawings (text boxes, shapes, pictures) -------------------------------------------
    # openpyxl ignores drawings, so a text box saying "修正済み" is invisible to it. The text is
    # plain XML in xl/drawings/, so read it from the package directly and anchor it to a cell.

    _NS = {
        "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
        "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
        "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    }

    def _sheet_drawings(self) -> dict[str, list[dict]]:
        """sheet title -> [{kind, name, anchor, text}] for every shape/picture/chart on the sheet."""
        from lxml import etree
        parser = etree.XMLParser(resolve_entities=False, no_network=True, huge_tree=False)
        ns = self._NS
        result: dict[str, list[dict]] = {}
        try:
            z = zipfile.ZipFile(self.path)
        except zipfile.BadZipFile:
            return result
        with z:
            names = set(z.namelist())

            def rels(part: str) -> dict[str, tuple[str, str]]:
                d = part.rsplit("/", 1)[0]
                rp = f"{d}/_rels/{part.rsplit('/', 1)[1]}.rels"
                if rp not in names:
                    return {}
                root = etree.fromstring(z.read(rp), parser)
                out = {}
                for r in root.findall("rel:Relationship", ns):
                    tgt = r.get("Target") or ""
                    if tgt.startswith("/"):
                        tgt = tgt.lstrip("/")
                    else:
                        base = d.split("/")
                        for seg in tgt.split("/"):
                            if seg == "..":
                                base.pop()
                            else:
                                base.append(seg)
                        tgt = "/".join(base)
                    out[r.get("Id")] = (r.get("Type", "").rsplit("/", 1)[-1], tgt)
                return out

            if "xl/workbook.xml" not in names:
                return result
            wb_root = etree.fromstring(z.read("xl/workbook.xml"), parser)
            wb_rels = rels("xl/workbook.xml")
            for sh in wb_root.findall("m:sheets/m:sheet", ns):
                title = sh.get("name")
                rid = sh.get("{%s}id" % ns["r"])
                sheet_part = wb_rels.get(rid, ("", ""))[1]
                if not sheet_part:
                    continue
                for typ, tgt in rels(sheet_part).values():
                    if typ != "drawing" or tgt not in names:
                        continue
                    droot = etree.fromstring(z.read(tgt), parser)
                    items: list[dict] = []
                    for anchor in droot:
                        frm = anchor.find("xdr:from", ns)
                        if frm is not None:
                            col = int(frm.findtext("xdr:col", "0", ns)) + 1
                            row = int(frm.findtext("xdr:row", "0", ns)) + 1
                            at = f"{_col_letter(col)}{row}"
                        else:
                            at = "(absolute)"
                        self._collect_shapes(anchor, at, items, ns)
                    if items:
                        result.setdefault(title, []).extend(items)
        return result

    def _collect_shapes(self, node, at: str, items: list[dict], ns: dict) -> None:
        for child in node:
            tag = etree_local(child.tag)
            if tag == "sp" or tag == "cxnSp":
                name = child.find("xdr:nvSpPr/xdr:cNvPr", ns) if tag == "sp" else child.find("xdr:nvCxnSpPr/xdr:cNvPr", ns)
                paras = []
                for p in child.findall("xdr:txBody/a:p", ns):
                    t = "".join(x.text or "" for x in p.iter("{%s}t" % ns["a"]))
                    if t.strip():
                        paras.append(t.strip())
                items.append({"kind": "shape", "name": name.get("name") if name is not None else "",
                              "anchor": at, "text": "\n".join(paras)})
            elif tag == "pic":
                name = child.find("xdr:nvPicPr/xdr:cNvPr", ns)
                items.append({"kind": "picture", "name": name.get("name") if name is not None else "",
                              "anchor": at, "text": ""})
            elif tag == "graphicFrame":
                name = child.find("xdr:nvGraphicFramePr/xdr:cNvPr", ns)
                items.append({"kind": "chart/frame", "name": name.get("name") if name is not None else "",
                              "anchor": at, "text": ""})
            elif tag == "grpSp":
                self._collect_shapes(child, at, items, ns)

    # -- content ------------------------------------------------------------------------

    def dump(self) -> dict:
        values_wb = self._open(data_only=True)
        formula_wb = self._open(data_only=False)   # always: uncached formulas must never look blank
        max_rows = self.opts.max_rows
        max_cols = self.opts.max_cols
        wanted = self.opts.sheet
        show_formulas = self.opts.formulas
        drawings = self._sheet_drawings()

        sheets = []
        for ws in values_wb.worksheets:
            if wanted and ws.title != wanted:
                continue
            fws = formula_wb[ws.title]
            info = self._sheet_info(ws, fws)

            rows: list[list[Any]] = []
            cells: list[dict] = []          # for --cells: every non-empty cell with its address
            truncated_rows = False
            truncated_cols = ws.max_column > max_cols if max_cols else False
            comments: list[dict] = []
            formulas_seen = 0
            formulas_cached = 0
            uncached_addrs: list[str] = []

            for r_idx, row in enumerate(ws.iter_rows(), start=1):
                if max_rows and r_idx > max_rows:
                    truncated_rows = True
                    break
                out_row: list[Any] = []
                for c_idx, cell in enumerate(row, start=1):
                    if max_cols and c_idx > max_cols:
                        break
                    val = cell.value
                    fval = fws.cell(row=cell.row, column=cell.column).value
                    is_formula = isinstance(fval, str) and fval.startswith("=")
                    if is_formula:
                        formulas_seen += 1
                        if val is None:
                            uncached_addrs.append(cell.coordinate)
                            val = f"{fval} [uncached]"
                        else:
                            formulas_cached += 1
                            if show_formulas:
                                val = f"{fval}  → {val}"
                    if cell.comment:
                        comments.append({"cell": cell.coordinate, "author": cell.comment.author,
                                         "text": cell.comment.text})
                    if val is not None and val != "":
                        cells.append({"address": cell.coordinate, "value": val,
                                      "formula": fval if is_formula else None})
                    out_row.append(val)
                rows.append(out_row)

            # Drop trailing empty rows/columns so tables stay narrow.
            while rows and all(c is None or c == "" for c in rows[-1]):
                rows.pop()
            last_col = 0
            for r in rows:
                for i, c in enumerate(r, start=1):
                    if c is not None and c != "":
                        last_col = max(last_col, i)
            rows = [r[:last_col] for r in rows]

            info.update({
                "rows": rows,
                "cells": cells,
                "truncated_rows": truncated_rows,
                "truncated_cols": truncated_cols,
                "comments": comments,
                "formulas_shown": formulas_seen,
                "formulas_cached": formulas_cached,
                "uncached_formulas": uncached_addrs,
                "drawings": drawings.get(ws.title, []),
            })
            sheets.append(info)

        if wanted and not sheets:
            sys.exit(f"[office2md] sheet '{wanted}' not found. Sheets: {values_wb.sheetnames}")
        return {"file": str(self.path), "type": "xlsx", "sheets": sheets}

    def _grid(self, rows: list[list[Any]]) -> str:
        """Markdown table with Excel's column letters as header and row numbers as first column,
        so every value can be cited as Sheet!C4."""
        if not rows:
            return "_(empty)_"
        width = max(len(r) for r in rows)
        header = [""] + [_col_letter(i) for i in range(1, width + 1)]
        body = [[i] + r for i, r in enumerate(rows, start=1)]
        return md_table(body, header=header)

    def to_markdown(self, data: dict) -> str:
        out = [f"# {Path(data['file']).name}", ""]
        if "rows" not in (data["sheets"][0] if data["sheets"] else {}):
            # outline
            out.append(f"Sheets: {len(data['sheets'])}  (active: {data.get('active_sheet')})")
            out.append("")
            out.append(md_table(
                [[s["name"], s["state"], s["dimensions"], s["max_row"], s["max_col"],
                  len(s["merged_ranges"]), s["formula_cells"], ", ".join(s["tables"])] for s in data["sheets"]],
                header=["sheet", "state", "range", "rows", "cols", "merged", "formulas", "tables"]))
            if any(s["formula_cells"] for s in data["sheets"]):
                out += ["", "Sheets with formulas: pass --formulas to see them; cached values may be empty "
                        "if the file was never recalculated (office-edit/scripts/recalc.py can compute them)."]
            if data.get("defined_names"):
                out += ["", "Defined names: " + ", ".join(data["defined_names"])]
            objs = data.get("package_objects") or []
            dv = sum(s["objects"]["data_validations"] for s in data["sheets"])
            cf = sum(s["objects"]["conditional_formats"] for s in data["sheets"])
            extra = ([f"{dv} data validation(s)"] if dv else []) + ([f"{cf} conditional format(s)"] if cf else [])
            if objs or extra:
                out += ["", "Objects: " + ", ".join(objs + extra) + "."]
                if "images" in objs:
                    out.append("> Images may contain text no parser can read. For a visual check, convert with "
                               "office-edit/scripts/pdf_tool.py topdf and read the PDF.")
                if objs:
                    out.append("> Saving with openpyxl may drop: " + ", ".join(objs)
                               + ". For small edits prefer Excel COM, or verify after saving (see office-edit pitfalls).")
            else:
                out += ["", "Objects: none (no charts, images, macros, pivots). Safe to edit with openpyxl."]
            return "\n".join(out)

        for s in data["sheets"]:
            flag = "" if s["state"] == "visible" else f"  _({s['state']})_"
            out.append(f"## Sheet: {s['name']}{flag}")
            out.append(f"Range {s['dimensions']} · {s['max_row']} rows × {s['max_col']} cols")
            if s["merged_ranges"]:
                out.append("Merged: " + ", ".join(s["merged_ranges"][:20])
                           + (" …" if len(s["merged_ranges"]) > 20 else ""))
            unc = s["uncached_formulas"]
            if unc:
                shown = ", ".join(unc[:10]) + (" …" if len(unc) > 10 else "")
                out.append(f"> WARNING: {len(unc)} formula cell(s) have no cached value ({shown}). "
                           "The file was never recalculated, so those values are unknown, not blank. "
                           "Do not judge totals or comparisons from them; run office-edit/scripts/recalc.py "
                           "on a copy, or compute by hand from the formulas shown.")
            out.append("")
            if self.opts.cells:
                for c in s["cells"]:
                    v = md_escape(c["value"]).replace("<br>", " / ")
                    out.append(f"- {s['name']}!{c['address']} = {v}")
                if not s["cells"]:
                    out.append("_(empty)_")
            elif self.opts.no_coords:
                out.append(md_table(s["rows"]))
            else:
                out.append(self._grid(s["rows"]))
            notes = []
            if s["truncated_rows"]:
                notes.append(f"rows truncated at {self.opts.max_rows}; rerun with --sheet \"{s['name']}\" --max-rows 0")
            if s["truncated_cols"]:
                notes.append(f"columns truncated at {self.opts.max_cols}; use --max-cols 0")
            if self.opts.formulas and s["formulas_cached"]:
                notes.append(f"{s['formulas_cached']} formula cells shown as `=formula  → cached value`")
            for n in notes:
                out.append(f"\n> {n}")
            if s["comments"]:
                out += ["", "**Cell comments**", md_table(
                    [[c["cell"], c["author"], c["text"]] for c in s["comments"]],
                    header=["cell", "author", "comment"])]
            if s["drawings"]:
                out += ["", "**Shapes and pictures** (text boxes are not cells; this is the only place their text appears)",
                        md_table([[d["anchor"], d["kind"], d["name"], d["text"]] for d in s["drawings"]],
                                 header=["anchor", "kind", "name", "text"])]
                if any(d["kind"] == "picture" for d in s["drawings"]):
                    out.append("> Pictures may contain text; convert with office-edit/scripts/pdf_tool.py topdf "
                               "and read the PDF to see them.")
            out.append("")
        return "\n".join(out)


# ----------------------------------------------------------------------------
# csv / tsv
# ----------------------------------------------------------------------------

class CsvHandler(Handler):
    name = "csv"

    def _read(self) -> list[list[str]]:
        raw = self.path.read_bytes()
        for enc in ("utf-8-sig", "cp932", "utf-16", "latin-1"):
            try:
                text = raw.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        delim = "\t" if self.path.suffix.lower() == ".tsv" else None
        if delim is None:
            try:
                delim = csv.Sniffer().sniff(text[:4096], delimiters=",;\t|").delimiter
            except csv.Error:
                delim = ","
        return list(csv.reader(io.StringIO(text), delimiter=delim))

    def outline(self) -> dict:
        rows = self._read()
        return {"file": str(self.path), "type": "csv", "rows": len(rows),
                "cols": max((len(r) for r in rows), default=0), "header": rows[0] if rows else []}

    def dump(self) -> dict:
        rows = self._read()
        n = self.opts.max_rows
        return {"file": str(self.path), "type": "csv", "total_rows": len(rows),
                "rows": rows[:n] if n else rows, "truncated_rows": bool(n and len(rows) > n)}

    def to_markdown(self, data: dict) -> str:
        out = [f"# {Path(data['file']).name}", ""]
        if "header" in data:
            out.append(f"{data['rows']} rows × {data['cols']} cols")
            out.append("Header: " + ", ".join(data["header"]))
            return "\n".join(out)
        out.append(md_table(data["rows"]))
        if data["truncated_rows"]:
            out.append(f"\n> rows truncated at {self.opts.max_rows} of {data['total_rows']}; use --max-rows 0")
        return "\n".join(out)


# ----------------------------------------------------------------------------
# docx
# ----------------------------------------------------------------------------

class DocxHandler(Handler):
    name = "docx"

    def _open(self):
        docx = need("docx", "python-docx")
        return docx.Document(self.path)

    @staticmethod
    def _heading_level(paragraph) -> int:
        name = (paragraph.style.name or "") if paragraph.style else ""
        m = re.match(r"(?:Heading|見出し)\s*(\d)", name)
        if m:
            return int(m[1])
        if name == "Title":
            return 1
        return 0

    @staticmethod
    def _is_list(paragraph) -> bool:
        name = (paragraph.style.name or "") if paragraph.style else ""
        if "List" in name or "箇条書き" in name or "段落番号" in name:
            return True
        return paragraph._p.pPr is not None and paragraph._p.pPr.numPr is not None

    @staticmethod
    def _para_text(paragraph) -> str:
        """Text including hyperlinks and tracked insertions; deletions are skipped."""
        from docx.oxml.ns import qn
        parts = []
        for node in paragraph._p.iter():
            if node.tag == qn("w:t"):
                parts.append(node.text or "")
            elif node.tag == qn("w:tab"):
                parts.append("\t")
            elif node.tag == qn("w:br"):
                parts.append("\n")
        return "".join(parts)

    def _iter_body(self, doc):
        """Yield ('p', paragraph) / ('t', table) in document order."""
        from docx.table import Table
        from docx.text.paragraph import Paragraph
        from docx.oxml.ns import qn
        for child in doc.element.body.iterchildren():
            if child.tag == qn("w:p"):
                yield "p", Paragraph(child, doc)
            elif child.tag == qn("w:tbl"):
                yield "t", Table(child, doc)

    def _table_rows(self, table) -> list[list[str]]:
        rows = []
        for row in table.rows:
            cells = []
            prev = None
            for cell in row.cells:
                # python-docx repeats merged cells; collapse the repeats.
                if cell._tc is prev:
                    cells.append("")
                else:
                    cells.append("\n".join(p.text for p in cell.paragraphs).strip())
                prev = cell._tc
            rows.append(cells)
        return rows

    def _comments(self) -> list[dict]:
        """Comments via raw XML so it works on any python-docx version."""
        try:
            with zipfile.ZipFile(self.path) as z:
                if "word/comments.xml" not in z.namelist():
                    return []
                xml = z.read("word/comments.xml")
        except zipfile.BadZipFile:
            return []
        # lxml ships with python-docx; disable entity resolution so a crafted file can't XXE us.
        from lxml import etree
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        parser = etree.XMLParser(resolve_entities=False, no_network=True, huge_tree=False)
        root = etree.fromstring(xml, parser)
        out = []
        for c in root.findall("w:comment", ns):
            text = "".join(t.text or "" for t in c.iter("{%s}t" % ns["w"]))
            out.append({"author": c.get("{%s}author" % ns["w"]), "date": c.get("{%s}date" % ns["w"]),
                        "text": text})
        return out

    def _has_tracked_changes(self) -> bool:
        try:
            with zipfile.ZipFile(self.path) as z:
                xml = z.read("word/document.xml")
        except (zipfile.BadZipFile, KeyError):
            return False
        return b"<w:ins " in xml or b"<w:del " in xml

    def outline(self) -> dict:
        doc = self._open()
        headings = []
        n_par = n_tbl = 0
        for kind, obj in self._iter_body(doc):
            if kind == "t":
                n_tbl += 1
                continue
            n_par += 1
            lvl = self._heading_level(obj)
            if lvl:
                headings.append({"level": lvl, "text": obj.text.strip()})
        cp = doc.core_properties
        return {"file": str(self.path), "type": "docx",
                "title": cp.title, "author": cp.author, "modified": str(cp.modified) if cp.modified else None,
                "paragraphs": n_par, "tables": n_tbl, "sections": len(doc.sections),
                "headings": headings, "comments": len(self._comments()),
                "tracked_changes": self._has_tracked_changes()}

    def dump(self) -> dict:
        doc = self._open()
        blocks = []
        for kind, obj in self._iter_body(doc):
            if kind == "t":
                blocks.append({"kind": "table", "rows": self._table_rows(obj)})
                continue
            text = self._para_text(obj).strip()
            if not text:
                continue
            lvl = self._heading_level(obj)
            if lvl:
                blocks.append({"kind": "heading", "level": lvl, "text": text})
            elif self._is_list(obj):
                blocks.append({"kind": "list", "text": text})
            else:
                blocks.append({"kind": "para", "text": text})
        headers = [p.text for s in doc.sections for p in s.header.paragraphs if p.text.strip()]
        footers = [p.text for s in doc.sections for p in s.footer.paragraphs if p.text.strip()]
        return {"file": str(self.path), "type": "docx", "blocks": blocks,
                "headers": sorted(set(headers)), "footers": sorted(set(footers)),
                "comments": self._comments(), "tracked_changes": self._has_tracked_changes()}

    def to_markdown(self, data: dict) -> str:
        out = [f"# {Path(data['file']).name}", ""]
        if "headings" in data:
            out.append(f"{data['paragraphs']} paragraphs · {data['tables']} tables · "
                       f"{data['comments']} comments · tracked changes: {data['tracked_changes']}")
            if data.get("title") or data.get("author"):
                out.append(f"Title: {data.get('title') or '-'} · Author: {data.get('author') or '-'} · Modified: {data.get('modified') or '-'}")
            out.append("")
            for h in data["headings"]:
                out.append("  " * (h["level"] - 1) + f"- {h['text']}")
            if not data["headings"]:
                out.append("_(no heading styles used)_")
            return "\n".join(out)

        if data["tracked_changes"]:
            out.append("> This document has tracked changes. Insertions are shown; deletions are omitted.\n")
        if data["headers"]:
            out.append("> Header: " + " / ".join(data["headers"]))
        if data["footers"]:
            out.append("> Footer: " + " / ".join(data["footers"]))
        if data["headers"] or data["footers"]:
            out.append("")
        blocks = data["blocks"]
        for i, b in enumerate(blocks):
            if b["kind"] == "heading":
                out.append("#" * (b["level"] + 1) + " " + b["text"])
            elif b["kind"] == "list":
                out.append("- " + b["text"])
                # keep consecutive list items together
                if i + 1 < len(blocks) and blocks[i + 1]["kind"] == "list":
                    continue
            elif b["kind"] == "table":
                out.append(md_table(b["rows"]))
            else:
                out.append(b["text"])
            out.append("")
        if data["comments"]:
            out += ["## Comments", md_table(
                [[c["author"], c["date"], c["text"]] for c in data["comments"]],
                header=["author", "date", "comment"])]
        return "\n".join(out)


# ----------------------------------------------------------------------------
# pptx
# ----------------------------------------------------------------------------

class PptxHandler(Handler):
    name = "pptx"

    def _open(self):
        pptx = need("pptx", "python-pptx")
        return pptx.Presentation(self.path)

    def _shapes(self, shapes, depth=0):
        """Flatten groups; yield shapes in z-order (top-to-bottom reading is approximate)."""
        from pptx.enum.shapes import MSO_SHAPE_TYPE
        for sh in shapes:
            if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
                yield from self._shapes(sh.shapes, depth + 1)
            else:
                yield sh

    def _slide_blocks(self, slide) -> list[dict]:
        blocks = []
        shapes = list(self._shapes(slide.shapes))
        # Reading order: sort by top, then left. z-order is rarely reading order.
        shapes.sort(key=lambda s: ((s.top or 0), (s.left or 0)))
        for sh in shapes:
            if sh.has_text_frame and sh.text_frame.text.strip():
                is_title = sh == slide.shapes.title if slide.shapes.title is not None else False
                paras = []
                for p in sh.text_frame.paragraphs:
                    t = "".join(r.text for r in p.runs).strip()
                    if t:
                        paras.append({"level": p.level, "text": t})
                blocks.append({"kind": "title" if is_title else "text", "name": sh.name, "paras": paras})
            elif getattr(sh, "has_table", False) and sh.has_table:
                rows = [[c.text for c in r.cells] for r in sh.table.rows]
                blocks.append({"kind": "table", "name": sh.name, "rows": rows})
            elif getattr(sh, "has_chart", False) and sh.has_chart:
                ch = sh.chart
                series = []
                for plot in ch.plots:
                    cats = list(plot.categories) if plot.categories else []
                    for s in plot.series:
                        series.append({"name": s.name, "categories": [str(c) for c in cats],
                                       "values": list(s.values)})
                blocks.append({"kind": "chart", "name": sh.name,
                               "title": ch.chart_title.text_frame.text if ch.has_title else None,
                               "series": series})
            elif sh.shape_type is not None and "PICTURE" in str(sh.shape_type):
                blocks.append({"kind": "picture", "name": sh.name})
        return blocks

    def outline(self) -> dict:
        prs = self._open()
        slides = []
        for i, slide in enumerate(prs.slides, start=1):
            title = slide.shapes.title.text.strip() if slide.shapes.title is not None and slide.shapes.title.has_text_frame else ""
            hidden = slide._element.get("show") == "0"
            slides.append({"n": i, "title": title, "layout": slide.slide_layout.name, "hidden": hidden,
                           "has_notes": slide.has_notes_slide and bool(slide.notes_slide.notes_text_frame.text.strip())})
        return {"file": str(self.path), "type": "pptx", "slide_count": len(slides), "slides": slides}

    def dump(self) -> dict:
        prs = self._open()
        slides = []
        for i, slide in enumerate(prs.slides, start=1):
            if self.opts.slide and i != self.opts.slide:
                continue
            notes = slide.notes_slide.notes_text_frame.text.strip() if slide.has_notes_slide else ""
            slides.append({"n": i, "layout": slide.slide_layout.name,
                           "hidden": slide._element.get("show") == "0",
                           "blocks": self._slide_blocks(slide), "notes": notes})
        if self.opts.slide and not slides:
            sys.exit(f"[office2md] slide {self.opts.slide} not found (deck has {len(prs.slides)} slides)")
        return {"file": str(self.path), "type": "pptx", "slides": slides}

    def to_markdown(self, data: dict) -> str:
        out = [f"# {Path(data['file']).name}", ""]
        if "slide_count" in data:
            out.append(f"{data['slide_count']} slides")
            out.append("")
            out.append(md_table(
                [[s["n"], s["title"], s["layout"], "hidden" if s["hidden"] else "", "yes" if s["has_notes"] else ""]
                 for s in data["slides"]],
                header=["#", "title", "layout", "hidden", "notes"]))
            return "\n".join(out)
        for s in data["slides"]:
            flag = "  _(hidden)_" if s["hidden"] else ""
            out.append(f"## Slide {s['n']}{flag}")
            for b in s["blocks"]:
                if b["kind"] == "title":
                    out.append("**" + " ".join(p["text"] for p in b["paras"]) + "**")
                elif b["kind"] == "text":
                    for p in b["paras"]:
                        out.append("  " * p["level"] + "- " + p["text"])
                elif b["kind"] == "table":
                    out.append(md_table(b["rows"]))
                elif b["kind"] == "chart":
                    out.append(f"_Chart{': ' + b['title'] if b['title'] else ''}_")
                    for se in b["series"]:
                        pairs = ", ".join(f"{c}={v}" for c, v in zip(se["categories"], se["values"]))
                        out.append(f"  - {se['name']}: {pairs}")
                elif b["kind"] == "picture":
                    out.append(f"_[picture: {b['name']}]_")
                out.append("")
            if s["notes"]:
                out.append("> **Notes:** " + s["notes"].replace("\n", "\n> "))
                out.append("")
        return "\n".join(out)


# ----------------------------------------------------------------------------
# pdf (text layer only; scanned PDFs need OCR, out of scope here)
# ----------------------------------------------------------------------------

class PdfHandler(Handler):
    name = "pdf"

    def _open(self):
        pypdf = need("pypdf", "pypdf")
        return pypdf.PdfReader(self.path)

    def outline(self) -> dict:
        r = self._open()
        meta = r.metadata or {}
        bookmarks = []
        def walk(items, depth=0):
            for it in items:
                if isinstance(it, list):
                    walk(it, depth + 1)
                else:
                    bookmarks.append({"level": depth, "title": it.title})
        try:
            walk(r.outline)
        except Exception:
            pass
        return {"file": str(self.path), "type": "pdf", "pages": len(r.pages),
                "title": meta.get("/Title"), "author": meta.get("/Author"),
                "encrypted": r.is_encrypted, "bookmarks": bookmarks}

    def dump(self) -> dict:
        r = self._open()
        pages = []
        for i, p in enumerate(r.pages, start=1):
            if self.opts.slide and i != self.opts.slide:
                continue
            pages.append({"n": i, "text": (p.extract_text() or "").strip()})
        return {"file": str(self.path), "type": "pdf", "pages": pages}

    def to_markdown(self, data: dict) -> str:
        out = [f"# {Path(data['file']).name}", ""]
        if "bookmarks" in data:
            out.append(f"{data['pages']} pages · Title: {data.get('title') or '-'} · Author: {data.get('author') or '-'}")
            for b in data["bookmarks"]:
                out.append("  " * b["level"] + "- " + b["title"])
            return "\n".join(out)
        empty = 0
        for p in data["pages"]:
            out.append(f"## Page {p['n']}")
            if p["text"]:
                out.append(p["text"])
            else:
                empty += 1
                out.append("_(no text layer)_")
            out.append("")
        if empty:
            out.append(f"> {empty} page(s) had no text layer; likely scanned. OCR is needed for those.")
        return "\n".join(out)


# ----------------------------------------------------------------------------
# Registry + CLI
# ----------------------------------------------------------------------------

HANDLERS: dict[str, type[Handler]] = {
    ".xlsx": XlsxHandler, ".xlsm": XlsxHandler, ".xltx": XlsxHandler, ".xltm": XlsxHandler,
    ".csv": CsvHandler, ".tsv": CsvHandler,
    ".docx": DocxHandler, ".dotx": DocxHandler, ".docm": DocxHandler,
    ".pptx": PptxHandler, ".potx": PptxHandler, ".pptm": PptxHandler,
    ".pdf": PdfHandler,
}


# ----------------------------------------------------------------------------
# Profile skeleton: structure facts filled in, meaning left as TODO for a human
# ----------------------------------------------------------------------------

def _col_letter(idx: int) -> str:
    s = ""
    while idx:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s


def _formula_ranges(ws) -> list[str]:
    """Group formula cells per column into 'H3:H120'-style ranges."""
    cols: dict[int, list[int]] = {}
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("="):
                cols.setdefault(c.column, []).append(c.row)
    out = []
    for col, rows in sorted(cols.items()):
        rows.sort()
        start = prev = rows[0]
        for r in rows[1:] + [None]:
            if r is None or r != prev + 1:
                L = _col_letter(col)
                out.append(f"{L}{start}" if start == prev else f"{L}{start}:{L}{prev}")
                if r is not None:
                    start = r
            if r is not None:
                prev = r
    return out


def make_profile(path: Path, ext: str) -> str:
    name = path.name
    stem = re.sub(r"[\d_\-]+$", "", path.stem) or path.stem
    head = [f"# {stem}（{name}）", "",
            "<!-- Generated skeleton. Lines marked TODO are where human knowledge goes. "
            "Delete this comment and anything that is not true for this file family. -->", "",
            "## 識別", f"- ファイル名: `{name}` のような名前  <!-- TODO: 命名規則 (例: 売上報告_YYYYMM.xlsx) -->"]
    body: list[str] = []
    if ext in (".xlsx", ".xlsm", ".xltx", ".xltm"):
        openpyxl = need("openpyxl", "openpyxl")
        wb = openpyxl.load_workbook(path, data_only=False)
        head.append("- シート構成: " + ", ".join(f"「{ws.title}」" + ("(非表示)" if ws.sheet_state != "visible" else "")
                                              for ws in wb.worksheets))
        body += ["", "## シート"]
        for ws in wb.worksheets:
            body.append(f"### {ws.title}" + ("  (非表示)" if ws.sheet_state != "visible" else ""))
            body.append("- 役割: TODO  <!-- 何のシートか。例: 入力用 / 集計用 / マスタ / 印刷用 -->")
            body.append(f"- 範囲: {ws.dimensions}")
            header = None
            for row in ws.iter_rows(min_row=1, max_row=10):
                vals = [c.value for c in row]
                if sum(1 for v in vals if v not in (None, "")) >= 2:
                    header = (row[0].row, [str(v).replace("\n", " ") for v in vals if v not in (None, "")])
                    break
            if header:
                body.append(f"- ヘッダ行: {header[0]} 行目 → {', '.join(header[1][:15])}"
                            + (" …" if len(header[1]) > 15 else ""))
                body.append(f"- データ開始行: {header[0] + 1}  <!-- TODO: 確認 -->")
            fr = _formula_ranges(ws)
            if fr:
                body.append("- 数式セル: " + ", ".join(fr[:12]) + (" …" if len(fr) > 12 else "")
                            + "  <!-- TODO: 触ってよいか -->")
            if ws.merged_cells.ranges:
                body.append("- 結合: " + ", ".join(str(r) for r in list(ws.merged_cells.ranges)[:10]))
            body.append("- 重要セル: TODO  <!-- 例: B3 = 対象年月(日付型), D10 = 合計(数式、変更禁止) -->")
            body.append("")
        if hasattr(wb.defined_names, "keys") and wb.defined_names.keys():
            body.append("- 定義名: " + ", ".join(sorted(wb.defined_names.keys())))
    elif ext in (".docx", ".dotx", ".docm"):
        h = DocxHandler(path, argparse.Namespace())
        o = h.outline()
        head.append(f"- 構成: 見出し {len(o['headings'])} 個、表 {o['tables']} 個"
                    + ("、変更履歴あり" if o["tracked_changes"] else ""))
        body += ["", "## 構成"]
        for hd in o["headings"]:
            body.append("  " * (hd["level"] - 1) + f"- {hd['text']}  <!-- TODO: この節に何を書くか -->")
        if o["tables"]:
            body.append(f"- 表 {o['tables']} 個: TODO  <!-- 何の表か、どの列を埋めるか -->")
    elif ext in (".pptx", ".potx", ".pptm"):
        h = PptxHandler(path, argparse.Namespace())
        o = h.outline()
        head.append(f"- スライド数: {o['slide_count']}")
        body += ["", "## スライド"]
        for s in o["slides"]:
            body.append(f"- {s['n']}. {s['title'] or '(無題)'} [{s['layout']}]"
                        + ("(非表示)" if s["hidden"] else "") + "  <!-- TODO: 更新対象か固定か -->")
    else:
        body += ["", "## 構成", "- TODO"]

    tail = ["", "## 運用ルール", "- TODO  <!-- 例: 追加は最終行の下に、書式は上の行をコピー -->",
            "- TODO  <!-- 例: 「マスタ」シートは編集禁止、社外に出す前に削除 -->",
            "", "## よくある依頼と手順", "- 「TODO」 = TODO  <!-- 例: 「先月分を作って」= コピーして B3 を更新し明細をクリア -->",
            "", "## 注意", "- TODO  <!-- 過去に壊れた点、見落としやすい点 -->"]
    return "\n".join(head + body + tail) + "\n"


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("file", type=Path)
    ap.add_argument("--outline", action="store_true", help="structure only, no content")
    ap.add_argument("--profile", action="store_true",
                    help="emit a profile skeleton (structure filled, meaning as TODO) for profiles/")
    ap.add_argument("--sheet", help="xlsx: only this sheet")
    ap.add_argument("--slide", type=int, help="pptx: only this slide number / pdf: only this page")
    ap.add_argument("--max-rows", type=int, default=DEFAULT_MAX_ROWS, help="rows per sheet, 0 = unlimited")
    ap.add_argument("--max-cols", type=int, default=DEFAULT_MAX_COLS, help="columns per sheet, 0 = unlimited")
    ap.add_argument("--formulas", action="store_true", help="xlsx: show formulas next to cached values")
    ap.add_argument("--cells", action="store_true",
                    help="xlsx: list non-empty cells as Sheet!A1 = value instead of a grid (good for forms and comparisons)")
    ap.add_argument("--no-coords", action="store_true", help="xlsx: plain table without row numbers / column letters")
    ap.add_argument("--json", action="store_true", help="emit JSON instead of Markdown")
    ap.add_argument("--out", type=Path, help="write output to this file")
    opts = ap.parse_args(argv)

    if not opts.file.exists():
        sys.exit(f"[office2md] file not found: {opts.file}")
    ext = opts.file.suffix.lower()
    if ext in (".xls", ".doc", ".ppt"):
        sys.exit(f"[office2md] legacy binary format {ext} is not supported. "
                 "Open it in Office and 'Save as' the x-format, or convert with LibreOffice: "
                 f"soffice --headless --convert-to {ext}x \"{opts.file}\"")
    cls = HANDLERS.get(ext)
    if cls is None:
        sys.exit(f"[office2md] unsupported extension {ext}. Supported: {', '.join(sorted(HANDLERS))}")

    if opts.profile:
        text = make_profile(opts.file, ext)
    else:
        h = cls(opts.file, opts)
        data = h.outline() if opts.outline else h.dump()
        text = json.dumps(data, ensure_ascii=False, indent=2, default=str) if opts.json else h.to_markdown(data)

    if opts.out:
        opts.out.write_text(text, encoding="utf-8")
        print(f"[office2md] wrote {opts.out} ({len(text)} chars)")
    else:
        print(text)
    return 0


if __name__ == "__main__":
    sys.exit(main())
