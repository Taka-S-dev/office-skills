"""Generate sample office files for testing office2md.py."""
import sys
from pathlib import Path

out = Path(sys.argv[1])
out.mkdir(parents=True, exist_ok=True)

# --- xlsx ---
import openpyxl
from openpyxl.comments import Comment
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "売上"
ws.append(["商品", "単価", "数量", "小計"])
for i, (n, p, q) in enumerate([("りんご", 120, 10), ("みかん", 80, 25), ("ぶどう|高級", 600, 3)], start=2):
    ws.append([n, p, q, f"=B{i}*C{i}"])
ws["A6"] = "合計"; ws["D6"] = "=SUM(D2:D4)"
ws["A1"].comment = Comment("商品名は正式名称で", "reviewer")
ws.merge_cells("F1:G1"); ws["F1"] = "備考\n改行あり"
ws2 = wb.create_sheet("hidden_memo"); ws2.sheet_state = "hidden"; ws2["A1"] = "秘密のメモ"
wsf = wb.create_sheet("申請書")          # form-like layout: label cells with values to the right / below
wsf["B2"] = "申請日："; wsf["C2"] = "2026/10/01"
wsf["B3"] = "対象機種"; wsf.merge_cells("B3:C3"); wsf["D3"] = "XYZ-100"
wsf["B5"] = "氏名"; wsf["C5"] = "所属"          # sibling labels in one row
wsf["B6"] = "山田"; wsf["C6"] = "品質保証部"
wsf["B8"] = "備考"; wsf["B9"] = "特になし"       # value below the label
wsf["E2"] = "合計"; wsf["F2"] = "=SUM(F3:F4)"; wsf["F3"] = 1; wsf["F4"] = 2
# display formats and hidden cells: what the screen shows vs what is stored
import datetime
wsf["C2"] = datetime.datetime(2026, 10, 1); wsf["C2"].number_format = "yyyy/mm/dd"
wsf["E5"] = "達成率"; wsf["F5"] = 0.15; wsf["F5"].number_format = "0%"
wsf["E6"] = "金額"; wsf["F6"] = 1234.5; wsf["F6"].number_format = "#,##0.0"
wsf["E7"] = "時刻"; wsf["F7"] = datetime.datetime(2026, 10, 1, 9, 5); wsf["F7"].number_format = "h:mm"
wsf["E9"] = "内部メモ"; wsf["F9"] = "非表示の値"
wsf.row_dimensions[9].hidden = True
wsf["H2"] = "隠し列"; wsf.column_dimensions["H"].hidden = True
ws3 = wb.create_sheet("big")
for r in range(300):
    ws3.append([r, r * 2])
wb.save(out / "sample.xlsx")


def inject_textbox(xlsx: Path, text: str, name: str = "注記") -> None:
    """Add a text box to sheet1 by writing the drawing parts directly into the package.

    openpyxl cannot create shapes, but the reader must handle them, so the fixture builds the
    same XML Excel would. Element order inside <worksheet> matters: <drawing> goes before
    <legacyDrawing> (which openpyxl adds for cell comments)."""
    import re
    import shutil
    import zipfile
    DRAWING = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
<xdr:twoCellAnchor><xdr:from><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
<xdr:to><xdr:col>7</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="2" name="{name}"/><xdr:cNvSpPr txBox="1"/></xdr:nvSpPr>
<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>
<xdr:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="ja-JP"/><a:t>{text}</a:t></a:r></a:p></xdr:txBody></xdr:sp>
<xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>"""
    REL = ('<Relationship Id="rIdDrw1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" '
           'Target="../drawings/drawing1.xml"/>')
    tmp = xlsx.with_suffix(".tmp")
    with zipfile.ZipFile(xlsx) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        names = zin.namelist()
        for n in names:
            data = zin.read(n)
            if n == "[Content_Types].xml":
                data = data.replace(b"</Types>", b'<Override PartName="/xl/drawings/drawing1.xml" '
                                    b'ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>')
            elif n == "xl/worksheets/sheet1.xml":
                tag = b'<drawing xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rIdDrw1"/>'
                if b"<legacyDrawing" in data:
                    data = data.replace(b"<legacyDrawing", tag + b"<legacyDrawing", 1)
                else:
                    data = re.sub(rb"</worksheet>", tag + b"</worksheet>", data, count=1)
            elif n == "xl/worksheets/_rels/sheet1.xml.rels":
                data = data.replace(b"</Relationships>", REL.encode() + b"</Relationships>")
            zout.writestr(n, data)
        if "xl/worksheets/_rels/sheet1.xml.rels" not in names:
            zout.writestr("xl/worksheets/_rels/sheet1.xml.rels",
                          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                          + REL + "</Relationships>")
        zout.writestr("xl/drawings/drawing1.xml", DRAWING)
    shutil.move(tmp, xlsx)


inject_textbox(out / "sample.xlsx", "修正済み: 単価は税抜")


def inject_parts(pkg: Path, parts: dict[str, str], rels: dict[str, str], content_types: list[tuple[str, str]]) -> None:
    """Add parts, relationship entries and content-type overrides to an OOXML package."""
    import shutil
    import zipfile
    tmp = pkg.with_suffix(".tmp")
    with zipfile.ZipFile(pkg) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        names = zin.namelist()
        for n in names:
            data = zin.read(n)
            if n == "[Content_Types].xml":
                adds = "".join(f'<Override PartName="/{p}" ContentType="{c}"/>' for p, c in content_types)
                data = data.replace(b"</Types>", adds.encode() + b"</Types>")
            elif n in rels:
                data = data.replace(b"</Relationships>", rels[n].encode() + b"</Relationships>")
            zout.writestr(n, data)
        for n, body in rels.items():
            if n not in names:
                zout.writestr(n, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                              '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                              + body + "</Relationships>")
        for n, body in parts.items():
            zout.writestr(n, body)
    shutil.move(tmp, pkg)


# --- review fixtures: tracked changes and comments (kept separate so Excel/Word never open them) ---
import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from pptx import Presentation
rd = docx.Document()
rd.add_paragraph("レビュー対象の文書")
p = rd.add_paragraph("変更前の文 ")
ins = OxmlElement("w:ins"); ins.set(qn("w:id"), "1"); ins.set(qn("w:author"), "reviewer"); ins.set(qn("w:date"), "2026-10-18T09:00:00Z")
r = OxmlElement("w:r"); t = OxmlElement("w:t"); t.text = "追加された文"; r.append(t); ins.append(r); p._p.append(ins)
dl = OxmlElement("w:del"); dl.set(qn("w:id"), "2"); dl.set(qn("w:author"), "reviewer"); dl.set(qn("w:date"), "2026-10-18T09:05:00Z")
r2 = OxmlElement("w:r"); dt = OxmlElement("w:delText"); dt.text = "削除された文"; r2.append(dt); dl.append(r2); p._p.append(dl)
p3 = rd.add_paragraph("この数値は要確認です。")
rd.add_comment(runs=p3.runs, text="出典を追記してください", author="reviewer", initials="rv")
rd.save(out / "review.docx")

TCNS = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"
rw = openpyxl.Workbook(); rws = rw.active; rws.title = "メイン"; rws["B2"] = 120; rws["C2"] = "メモ"
rws["C2"].comment = Comment("旧式のメモ", "taro")
rw.save(out / "review.xlsx")
inject_parts(out / "review.xlsx",
    parts={
        "xl/persons/person.xml": f'<personList xmlns="{TCNS}"><person displayName="reviewer" id="{{P1}}" userId="reviewer" providerId="None"/></personList>',
        "xl/threadedComments/threadedComment1.xml":
            f'<ThreadedComments xmlns="{TCNS}">'
            '<threadedComment ref="B2" dT="2026-10-18T09:00:00.00" personId="{P1}" id="{C1}"><text>単価を確認</text></threadedComment>'
            '<threadedComment ref="B2" dT="2026-10-18T10:00:00.00" personId="{P1}" id="{C2}" parentId="{C1}"><text>確認済み</text></threadedComment>'
            '</ThreadedComments>',
    },
    rels={
        "xl/worksheets/_rels/sheet1.xml.rels": '<Relationship Id="rIdTc1" Type="http://schemas.microsoft.com/office/2017/10/relationships/threadedComment" Target="../threadedComments/threadedComment1.xml"/>',
        "xl/_rels/workbook.xml.rels": '<Relationship Id="rIdPer1" Type="http://schemas.microsoft.com/office/2017/10/relationships/person" Target="persons/person.xml"/>',
    },
    content_types=[("xl/persons/person.xml", "application/vnd.ms-excel.person+xml"),
                   ("xl/threadedComments/threadedComment1.xml", "application/vnd.ms-excel.threadedcomments+xml")])

PNS = "http://schemas.openxmlformats.org/presentationml/2006/main"
rp = Presentation(); rs = rp.slides.add_slide(rp.slide_layouts[5]); rs.shapes.title.text = "レビュー対象"
rp.save(out / "review.pptx")
inject_parts(out / "review.pptx",
    parts={
        "ppt/commentAuthors.xml": f'<p:cmAuthorLst xmlns:p="{PNS}"><p:cmAuthor id="0" name="reviewer" initials="rv" lastIdx="1" clrIdx="0"/></p:cmAuthorLst>',
        "ppt/comments/comment1.xml": f'<p:cmLst xmlns:p="{PNS}"><p:cm authorId="0" dt="2026-10-18T09:00:00" idx="1"><p:pos x="10" y="10"/><p:text>タイトルを見直す</p:text></p:cm></p:cmLst>',
    },
    rels={
        "ppt/slides/_rels/slide1.xml.rels": '<Relationship Id="rIdCm1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments/comment1.xml"/>',
        "ppt/_rels/presentation.xml.rels": '<Relationship Id="rIdCa1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors" Target="commentAuthors.xml"/>',
    },
    content_types=[("ppt/commentAuthors.xml", "application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml"),
                   ("ppt/comments/comment1.xml", "application/vnd.openxmlformats-officedocument.presentationml.comments+xml")])

# --- docx ---
import docx
d = docx.Document()
d.add_heading("報告書タイトル", 0)
d.add_heading("1. 概要", 1)
d.add_paragraph("これは概要の段落です。")
d.add_heading("1.1 詳細", 2)
d.add_paragraph("箇条書き その1", style="List Bullet")
d.add_paragraph("箇条書き その2", style="List Bullet")
t = d.add_table(rows=2, cols=2)
t.cell(0, 0).text = "項目"; t.cell(0, 1).text = "値"
t.cell(1, 0).text = "A"; t.cell(1, 1).text = "100"
d.sections[0].header.paragraphs[0].text = "社外秘"
d.save(out / "sample.docx")

# --- spec.docx: what a real specification contains and a plain walk loses ---
# auto-numbered headings (numbering.xml), nested lists, a footnote, a figure with alt text,
# a floating text box, an equation, and a table-of-contents field.
import struct as _struct
import zlib as _zlib
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import Part

sd = docx.Document()
numbering = sd.part.numbering_part.element
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
abstract = parse_xml(
    f'<w:abstractNum {nsdecls("w")} w:abstractNumId="900">'
    '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="第%1章"/></w:lvl>'
    '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2"/></w:lvl>'
    '<w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2.%3"/></w:lvl>'
    '</w:abstractNum>')
num = parse_xml(f'<w:num {nsdecls("w")} w:numId="900"><w:abstractNumId w:val="900"/></w:num>')
# abstractNum elements must precede num elements in numbering.xml
first_num = numbering.find(qn("w:num"))
(first_num.addprevious(abstract) if first_num is not None else numbering.append(abstract))
numbering.append(num)


def numbered(paragraph, ilvl: int, num_id: int = 900):
    pPr = paragraph._p.get_or_add_pPr()
    np_ = OxmlElement("w:numPr")
    il = OxmlElement("w:ilvl"); il.set(qn("w:val"), str(ilvl))
    ni = OxmlElement("w:numId"); ni.set(qn("w:val"), str(num_id))
    np_.append(il); np_.append(ni); pPr.append(np_)


sd.add_heading("注文処理仕様書", 0)
toc = parse_xml(
    f'<w:sdt {nsdecls("w")}><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/></w:docPartObj></w:sdtPr>'
    '<w:sdtContent><w:p><w:r><w:t>第1章 概要 ....... 1</w:t></w:r></w:p><w:p><w:r><w:t>1.1 目的 ....... 1</w:t></w:r></w:p>'
    '<w:p><w:r><w:t>第2章 データ形式 ....... 2</w:t></w:r></w:p></w:sdtContent></w:sdt>')
sd.paragraphs[-1]._p.addnext(toc)          # right after the title, where Word puts it
h = sd.add_heading("概要", 1); numbered(h, 0)
h = sd.add_heading("目的", 2); numbered(h, 1)
p = sd.add_paragraph("本書は注文データの形式を定める。")
# footnote reference in the paragraph, note body in a footnotes part
r = p.add_run(); fr = OxmlElement("w:footnoteReference"); fr.set(qn("w:id"), "1"); r._r.append(fr)
footnotes_xml = (
    f'<w:footnotes {nsdecls("w")}>'
    '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>'
    '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>'
    '<w:footnote w:id="1"><w:p><w:r><w:t>旧版（第2版）の形式は対象外。</w:t></w:r></w:p></w:footnote></w:footnotes>')
fn_part = Part(PackURI("/word/footnotes.xml"),
               "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
               footnotes_xml.encode("utf-8"), sd.part.package)
sd.part.relate_to(fn_part, RT.FOOTNOTES)
p = sd.add_paragraph("関連文書: ")
rid = sd.part.relate_to("https://example.com/spec/order-api-v3", RT.HYPERLINK, is_external=True)
hl = OxmlElement("w:hyperlink"); hl.set(qn("r:id"), rid)
hr = OxmlElement("w:r"); ht = OxmlElement("w:t"); ht.text = "注文 API 仕様 v3"; hr.append(ht); hl.append(hr)
p._p.append(hl)
h = sd.add_heading("用語", 2); numbered(h, 1)
sd.add_paragraph("必須項目", style="List Bullet")
sd.add_paragraph("任意項目", style="List Bullet")
q = sd.add_paragraph("省略可能な拡張項目", style="List Bullet 2")
h = sd.add_heading("データ形式", 1); numbered(h, 0)
h = sd.add_heading("見出し部", 2); numbered(h, 1)
h = sd.add_heading("件数項目", 3); numbered(h, 2)
sd.add_paragraph("合計金額は次式で求める。")
eq = parse_xml('<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
               '<m:oMath><m:r><m:t>T = P + N × 4</m:t></m:r></m:oMath></m:oMathPara>')
sd.add_paragraph()._p.append(eq)
pb = sd.add_paragraph("在庫が無い場合は"); rb = pb.add_run("30日以内に入荷"); rb.bold = True; pb.add_run("を通知する。")
sd.add_paragraph("手順は以下の順に行う。")
sd.add_paragraph("注文を受け付ける", style="List Number")
sd.add_paragraph("在庫を確認する", style="List Number")
sd.add_paragraph("出荷を指示する", style="List Number")
h = sd.add_heading("明細部", 2); numbered(h, 1)
# a 1x1 PNG, inline figure with alt text, plus a caption
png = (b"\x89PNG\r\n\x1a\n" + _struct.pack(">I", 13) + b"IHDR" + _struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0))
png = png[:8] + png[8:] + _struct.pack(">I", _zlib.crc32(png[12:]) & 0xffffffff)
raw = _zlib.compress(b"\x00\xff\x00\x00")
png += _struct.pack(">I", len(raw)) + b"IDAT" + raw + _struct.pack(">I", _zlib.crc32(b"IDAT" + raw) & 0xffffffff)
png += _struct.pack(">I", 0) + b"IEND" + _struct.pack(">I", _zlib.crc32(b"IEND") & 0xffffffff)
(out / "fig.png").write_bytes(png)
sd.add_picture(str(out / "fig.png"))
docpr = sd.element.body.findall(".//" + qn("wp:docPr"))[-1]
docpr.set("descr", "データの全体構成図")
sd.add_paragraph("図 1 データの構成", style="Caption")
tb = parse_xml(
    f'<w:p {nsdecls("w")} xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:v="urn:schemas-microsoft-com:vml">'
    '<w:r><w:t>明細部は見出し部の直後に続く。</w:t></w:r>'
    '<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor><wp:docPr id="7" name="注記"/>'
    '<wps:wsp><wps:txbx><w:txbxContent><w:p><w:r><w:t>注記: 改訂時は見出し部も見直すこと</w:t></w:r></w:p></w:txbxContent></wps:txbx></wps:wsp>'
    '</wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:pict><v:shape><v:textbox><w:txbxContent><w:p><w:r><w:t>注記: 改訂時は見出し部も見直すこと</w:t></w:r></w:p>'
    '</w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p>')
sd.paragraphs[-1]._p.addnext(tb)           # keeps sectPr last, as Word requires
sd.save(out / "spec.docx")
(out / "fig.png").unlink()

# --- diagram.xlsx: what Excel "drawing paper" specifications contain ---------------------
# a sequence diagram drawn with boxes, dashed lifelines and arrows (labels in text boxes and in a
# cell), a flowchart with glued connectors, a record layout with merged cells and input rules,
# and a chart with a picture.
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart import BarChart, Reference as ChartRef

dw = openpyxl.Workbook()
seq_ws = dw.active; seq_ws.title = "シーケンス"
seq_ws["G9"] = "在庫照会"                                   # a message label written in a cell
flow_ws = dw.create_sheet("フロー")
pk = dw.create_sheet("レイアウト")
pk["A2"] = "桁"
for i in range(8):
    pk.cell(row=2, column=2 + i, value=i)
pk["A3"] = 0; pk["B3"] = "区分"; pk.merge_cells("B3:E3"); pk["F3"] = "店舗"; pk.merge_cells("F3:I3")
pk["A4"] = 1; pk["B4"] = "日付"; pk.merge_cells("B4:I4")
pk["A5"] = "2-3"; pk["B5"] = "金額"; pk.merge_cells("B5:I5")
pk["K2"] = "要否"; pk["L2"] = "版"
for v, r in (("必須", 3), ("任意", 4), ("必須", 5)):
    pk.cell(row=r, column=11, value=v)
pk["N2"] = 1; pk["N3"] = 2; pk["N4"] = 3
dv1 = DataValidation(type="list", formula1='"必須,任意"', allow_blank=True); dv1.add("K3:K5"); pk.add_data_validation(dv1)
dv2 = DataValidation(type="list", formula1="=$N$2:$N$4"); dv2.add("L3:L5"); pk.add_data_validation(dv2)
gr = dw.create_sheet("グラフ")
for row in (("月", "売上"), ("1月", 120), ("2月", 80), ("3月", 150)):
    gr.append(row)
bc = BarChart(); bc.title = "月別売上"
bc.add_data(ChartRef(gr, min_col=2, min_row=1, max_row=4), titles_from_data=True)
bc.set_categories(ChartRef(gr, min_col=1, min_row=2, max_row=4))
gr.add_chart(bc, "D2")
try:
    from PIL import Image as _PIL
    from openpyxl.drawing.image import Image as XLImage
    _PIL.new("RGB", (40, 20), (200, 30, 30)).save(out / "_pic.png")
    pic = XLImage(str(out / "_pic.png")); gr.add_image(pic, "D18")
except ImportError:
    pass
dw.save(out / "diagram.xlsx")
if (out / "_pic.png").exists():
    (out / "_pic.png").unlink()

XDR = ('xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
       'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"')


def _anchor(c1, r1, c2, r2, body):
    return (f"<xdr:twoCellAnchor><xdr:from><xdr:col>{c1}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{r1}</xdr:row>"
            f"<xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>{c2}</xdr:col><xdr:colOff>0</xdr:colOff>"
            f"<xdr:row>{r2}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>{body}<xdr:clientData/></xdr:twoCellAnchor>")


def _box(sid, name, text, prst="rect", txbox=False):
    tb = ' txBox="1"' if txbox else ""
    return (f'<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="{sid}" name="{name}"/><xdr:cNvSpPr{tb}/></xdr:nvSpPr>'
            f'<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="{prst}"><a:avLst/></a:prstGeom></xdr:spPr>'
            f'<xdr:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="ja-JP"/><a:t>{text}</a:t></a:r></a:p></xdr:txBody></xdr:sp>')


def _line(sid, name, flip_h=False, head=False, tail=False, dash=False, st=None, en=None):
    flip = ' flipH="1"' if flip_h else ""
    glue = (f'<a:stCxn id="{st}" idx="2"/>' if st else "") + (f'<a:endCxn id="{en}" idx="0"/>' if en else "")
    ln = ('<a:solidFill><a:srgbClr val="000000"/></a:solidFill>' + ('<a:prstDash val="dash"/>' if dash else "")
          + ('<a:headEnd type="triangle"/>' if head else "") + ('<a:tailEnd type="triangle"/>' if tail else ""))
    return (f'<xdr:cxnSp macro=""><xdr:nvCxnSpPr><xdr:cNvPr id="{sid}" name="{name}"/><xdr:cNvCxnSpPr>{glue}</xdr:cNvCxnSpPr></xdr:nvCxnSpPr>'
            f'<xdr:spPr><a:xfrm{flip}><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>'
            f'<a:ln w="9525">{ln}</a:ln></xdr:spPr></xdr:cxnSp>')


seq_xml = "".join([
    _anchor(1, 1, 3, 3, _box(2, "ブラウザ", "ブラウザ")), _anchor(5, 1, 7, 3, _box(3, "APIサーバ", "APIサーバ")),
    _anchor(9, 1, 11, 3, _box(4, "DB", "DB")),
    _anchor(2, 3, 2, 18, _line(5, "生存線 ブラウザ", dash=True)), _anchor(6, 3, 6, 18, _line(6, "生存線 APIサーバ", dash=True)),
    _anchor(10, 3, 10, 18, _line(7, "生存線 DB", dash=True)),
    _anchor(2, 6, 6, 6, _line(8, "矢印 1", tail=True)), _anchor(3, 4, 5, 6, _box(9, "ラベル 1", "注文登録", txbox=True)),
    _anchor(6, 9, 10, 9, _line(10, "矢印 2", tail=True)),
    _anchor(6, 12, 10, 12, _line(11, "矢印 3", head=True)), _anchor(7, 10, 9, 11, _box(12, "ラベル 3", "在庫数", txbox=True)),
    _anchor(2, 15, 6, 15, _line(13, "矢印 4", flip_h=True, tail=True)), _anchor(3, 13, 5, 14, _box(14, "ラベル 4", "登録完了", txbox=True)),
])
flow_xml = "".join([
    _anchor(1, 1, 4, 3, _box(2, "開始", "開始")), _anchor(1, 5, 4, 8, _box(3, "判定", "判定", prst="diamond")),
    _anchor(1, 10, 4, 12, _box(4, "終了", "終了")), _anchor(5, 5, 8, 8, _box(7, "差戻し", "差戻し")),
    _anchor(2, 3, 2, 5, _line(5, "コネクタ 開始-判定", tail=True, st=2, en=3)),
    _anchor(2, 8, 2, 10, _line(6, "コネクタ 判定-終了", tail=True, st=3, en=4)),
    _anchor(4, 6, 5, 6, _line(8, "矢印 判定-差戻し", tail=True)),
])


def inject_drawing(xlsx: Path, sheet_xml: str, name: str, body: str) -> None:
    """Add a drawing part with the given anchors to one worksheet of an existing package."""
    import shutil
    import zipfile
    rels_name = sheet_xml.replace("worksheets/", "worksheets/_rels/") + ".rels"
    rel = (f'<Relationship Id="rIdInj{name}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" '
           f'Target="../drawings/{name}.xml"/>')
    drawing = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr {XDR}>{body}</xdr:wsDr>'
    tmp = xlsx.with_suffix(".tmp")
    with zipfile.ZipFile(xlsx) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        names = zin.namelist()
        for n in names:
            data = zin.read(n)
            if n == "[Content_Types].xml":
                data = data.replace(b"</Types>", f'<Override PartName="/xl/drawings/{name}.xml" '
                                    'ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>'.encode())
            elif n == sheet_xml:
                tag = f'<drawing xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rIdInj{name}"/>'.encode()
                data = data.replace(b"</worksheet>", tag + b"</worksheet>", 1)
            elif n == rels_name:
                data = data.replace(b"</Relationships>", rel.encode() + b"</Relationships>")
            zout.writestr(n, data)
        if rels_name not in names:
            zout.writestr(rels_name, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                          + rel + "</Relationships>")
        zout.writestr(f"xl/drawings/{name}.xml", drawing)
    shutil.move(tmp, xlsx)


inject_drawing(out / "diagram.xlsx", "xl/worksheets/sheet1.xml", "drawing91", seq_xml)
inject_drawing(out / "diagram.xlsx", "xl/worksheets/sheet2.xml", "drawing92", flow_xml)

# --- pptx ---
from pptx import Presentation
from pptx.util import Inches
prs = Presentation()
s = prs.slides.add_slide(prs.slide_layouts[1])
s.shapes.title.text = "第1章 はじめに"
s.placeholders[1].text_frame.text = "ポイント1"
p = s.placeholders[1].text_frame.add_paragraph(); p.text = "サブポイント"; p.level = 1
s.notes_slide.notes_text_frame.text = "ここは発表者ノート"
s2 = prs.slides.add_slide(prs.slide_layouts[5])
s2.shapes.title.text = "表のスライド"
tbl = s2.shapes.add_table(2, 2, Inches(1), Inches(2), Inches(4), Inches(1)).table
tbl.cell(0, 0).text = "x"; tbl.cell(0, 1).text = "y"; tbl.cell(1, 0).text = "1"; tbl.cell(1, 1).text = "2"
prs.save(out / "sample.pptx")

# --- eml (mail kept in a folder, with an attachment) ---
from email.message import EmailMessage
for n, (subj, date, body) in enumerate([
        ("月次報告の確認依頼", "Mon, 06 Oct 2026 09:00:00 +0900", "添付の月次報告を確認してください。\n合計が前月と合いません。"),
        ("Re: 月次報告の確認依頼", "Tue, 07 Oct 2026 18:30:00 +0900", "評価シートの No.11 はメインの No.12 と不一致です。\n修正版を金曜までに送ります。"),
        ("Re: Re: 月次報告の確認依頼", "Fri, 10 Oct 2026 11:15:00 +0900", "修正版を添付します。D6 の数式を D2:D5 に広げました。")], start=1):
    em = EmailMessage()
    em["Subject"] = subj; em["From"] = "山田 <yamada@example.com>"; em["To"] = "reviewer@example.com"
    em["Cc"] = "team@example.com"; em["Date"] = date
    em.set_content(body)
    if n == 3:
        em.add_attachment((out / "sample.xlsx").read_bytes(), maintype="application",
                          subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename="月次報告_修正版.xlsx")
    (out / f"mail{n}.eml").write_bytes(bytes(em))

# --- archives: a zip with Japanese names and a nested 7z, attached to a 4th mail ---
import zipfile
import py7zr
with py7zr.SevenZipFile(out / "inner.7z", "w") as z7:
    z7.writestr("7z の中のメモ".encode("utf-8"), "内側/メモ.txt")
with zipfile.ZipFile(out / "添付資料.zip", "w", zipfile.ZIP_DEFLATED) as zf:
    zf.write(out / "sample.docx", "資料/報告書.docx")
    zf.write(out / "inner.7z", "inner.7z")
(out / "inner.7z").unlink()
em = EmailMessage()
em["Subject"] = "Re: Re: Re: 月次報告の確認依頼"; em["From"] = "山田 <yamada@example.com>"; em["To"] = "reviewer@example.com"
em["Date"] = "Sat, 11 Oct 2026 08:00:00 +0900"
em.set_content("関連資料を圧縮して添付します。")
em.add_attachment((out / "添付資料.zip").read_bytes(), maintype="application", subtype="zip", filename="添付資料.zip")
(out / "mail4.eml").write_bytes(bytes(em))

# --- csv ---
(out / "sample.csv").write_text("名前,年齢\n太郎,30\n花子,25\n", encoding="cp932")
print("samples written to", out)
