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
# a table with a header spanning two columns and groups spanning rows, as Word stores them
md_ = docx.Document()
mt = md_.add_table(rows=4, cols=4)
for (r, c), v in {(0, 0): "区分", (0, 1): "項目", (0, 3): "備考", (1, 1): "名称", (1, 2): "桁数",
                  (2, 0): "基本", (2, 1): "受注番号", (2, 2): "10", (2, 3): "必須",
                  (3, 1): "受注日", (3, 2): "8", (3, 3): "任意"}.items():
    mt.cell(r, c).text = v
mt.cell(0, 1).merge(mt.cell(0, 2))
mt.cell(0, 0).merge(mt.cell(1, 0))
mt.cell(2, 0).merge(mt.cell(3, 0))
md_.save(out / "merged.docx")

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


# --- seq_rich.xlsx: the rest of UML sequence notation ------------------------------------
# activation bar, self message drawn as three separate lines, alt frame with a tag box, an else
# separator and [conditions], dashed replies, and a note. Shapes have outlines so the rendered
# sheet shows them.
def _rect(sid, name, text="", fill=None, a_from=None):
    fill_xml = '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>' if fill == "white" else "<a:noFill/>"
    body = (f'<xdr:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="ja-JP" sz="900"/><a:t>{text}</a:t></a:r></a:p></xdr:txBody>'
            if text else "")
    return (f'<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="{sid}" name="{name}"/><xdr:cNvSpPr/></xdr:nvSpPr>'
            f'<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
            f'{fill_xml}<a:ln w="9525"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></xdr:spPr>{body}</xdr:sp>')


def _anchor_off(c1, o1, r1, c2, o2, r2, body):
    return (f"<xdr:twoCellAnchor><xdr:from><xdr:col>{c1}</xdr:col><xdr:colOff>{o1}</xdr:colOff><xdr:row>{r1}</xdr:row>"
            f"<xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>{c2}</xdr:col><xdr:colOff>{o2}</xdr:colOff>"
            f"<xdr:row>{r2}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>{body}<xdr:clientData/></xdr:twoCellAnchor>")


COL = 609600                     # default column width in EMU (64 px)
rw = openpyxl.Workbook(); rw.active.title = "シーケンス2"; rw.save(out / "seq_rich.xlsx")
rich_xml = "".join([
    _anchor(1, 1, 3, 3, _rect(2, "ブラウザ", "ブラウザ")), _anchor(5, 1, 7, 3, _rect(3, "APIサーバ", "APIサーバ")),
    _anchor(9, 1, 11, 3, _rect(4, "DB", "DB")),
    _anchor(2, 3, 2, 22, _line(5, "生存線 1", dash=True)), _anchor(6, 3, 6, 22, _line(6, "生存線 2", dash=True)),
    _anchor(10, 3, 10, 22, _line(7, "生存線 3", dash=True)),
    _anchor_off(5, int(COL * 0.85), 5, 6, int(COL * 0.15), 20, _rect(8, "活性区間", fill="white")),
    _anchor(2, 5, 6, 5, _line(9, "矢印 注文登録", tail=True)), _anchor(3, 3, 5, 5, _box(10, "ラベル 注文登録", "注文登録", txbox=True)),
    _anchor(6, 7, 7, 7, _line(11, "自己 1")), _anchor(7, 7, 7, 8, _line(12, "自己 2")),
    _anchor(6, 8, 7, 8, _line(13, "自己 3", flip_h=True, tail=True)), _anchor(7, 6, 9, 7, _box(14, "ラベル 検証", "検証", txbox=True)),
    _anchor(1, 9, 11, 16, _rect(15, "枠")), _anchor(1, 9, 3, 10, _box(16, "枠タグ", "alt [在庫あり]", txbox=True)),
    _anchor(6, 11, 10, 11, _line(17, "矢印 引当", tail=True)), _anchor(7, 10, 9, 11, _box(18, "ラベル 引当", "引当", txbox=True)),
    _anchor(1, 13, 11, 13, _line(19, "else 区切り", dash=True)), _anchor(1, 13, 3, 14, _box(20, "else 条件", "[在庫なし]", txbox=True)),
    _anchor(2, 15, 6, 15, _line(21, "戻り 在庫切れ", flip_h=True, tail=True, dash=True)),
    _anchor(3, 14, 5, 15, _box(22, "ラベル 在庫切れ", "在庫切れ通知", txbox=True)),
    _anchor(9, 17, 12, 18, _box(23, "ノート", "読み取り専用", txbox=True)),
    _anchor(2, 20, 6, 20, _line(24, "戻り 注文番号", flip_h=True, tail=True, dash=True)),
    _anchor(3, 18, 5, 19, _box(25, "ラベル 注文番号", "注文番号", txbox=True)),
])
inject_drawing(out / "seq_rich.xlsx", "xl/worksheets/sheet1.xml", "drawing93", rich_xml)

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

# --- objects.*: content a plain reader drops (SmartArt, legacy shapes, chart sheets) --------
# SmartArt (same data part in all three formats), legacy VML text box, chart sheet, printed
# header, struck-through cells and runs, repeated values and sentences, an embedded object.
DGM_NS = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
DGM_DATA_TYPE = "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml"
DGM_DATA_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData"


def _dgm_pt(mid, text="", typ=None):
    t = f' type="{typ}"' if typ else ""
    body = f"<a:p><a:r><a:t>{text}</a:t></a:r></a:p>" if text else "<a:p/>"
    return f'<dgm:pt modelId="{mid}"{t}><dgm:prSet/><dgm:spPr/><dgm:t><a:bodyPr/><a:lstStyle/>{body}</dgm:t></dgm:pt>'


def _dgm_cxn(mid, src, dst, order, typ=None):
    t = f' type="{typ}"' if typ else ""
    return f'<dgm:cxn modelId="{mid}"{t} srcId="{src}" destId="{dst}" srcOrd="{order}" destOrd="0"/>'


# document node 0; 受付, 審査 (with child 書類確認), 承認; a presentation point that must not show up
SMARTART_DATA = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><dgm:dataModel xmlns:dgm="{DGM_NS}" xmlns:a="{A_NS}">'
                 "<dgm:ptLst>" + _dgm_pt(0, typ="doc") + _dgm_pt(4, "承認") + _dgm_pt(1, "受付") + _dgm_pt(2, "審査")
                 + _dgm_pt(3, "書類確認") + _dgm_pt(9, "表示用", typ="pres") + "</dgm:ptLst><dgm:cxnLst>"
                 + _dgm_cxn(10, 0, 4, 2) + _dgm_cxn(11, 0, 1, 0) + _dgm_cxn(12, 0, 2, 1) + _dgm_cxn(13, 2, 3, 0)
                 + _dgm_cxn(14, 1, 9, 0, typ="presOf") + "</dgm:cxnLst><dgm:bg/><dgm:whole/></dgm:dataModel>")


def _smartart_frame(prefix, rid, sid=40, name="図表 1"):
    return (f'<{prefix}:graphicFrame macro=""><{prefix}:nvGraphicFramePr><{prefix}:cNvPr id="{sid}" name="{name}"/>'
            f'<{prefix}:cNvGraphicFramePr/></{prefix}:nvGraphicFramePr><{prefix}:xfrm><a:off x="0" y="0"/>'
            f'<a:ext cx="3000000" cy="1500000"/></{prefix}:xfrm><a:graphic><a:graphicData uri="{DGM_NS}">'
            f'<dgm:relIds xmlns:dgm="{DGM_NS}" xmlns:r="{R_NS}" r:dm="{rid}" r:lo="" r:qs="" r:cs=""/>'
            f'</a:graphicData></a:graphic></{prefix}:graphicFrame>')


from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Font
ow = openpyxl.Workbook()
ws = ow.active
ws.title = "部品"
for row in [["品名", "数量", "状態"], ["ボルト", 2, "○"], ["ナット", 5, "未定"], ["ワッシャ", "旧品番", "未定"]]:
    ws.append(row)
ws["B4"].font = Font(strike=True)
ws.oddHeader.center.text = "文書番号 DOC-001 第2版"
cs = ow.create_chartsheet("グラフ専用")
bar = BarChart()
bar.title = "数量の推移"
bar.add_data(Reference(ws, min_col=2, min_row=1, max_row=3), titles_from_data=True)
cs.add_chart(bar)
ow.save(out / "objects.xlsx")
inject_drawing(out / "objects.xlsx", "xl/worksheets/sheet1.xml", "drawing94",
               _anchor(5, 1, 9, 8, _smartart_frame("xdr", "rIdDgm")))
VML = ('<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" '
       'xmlns:x="urn:schemas-microsoft-com:office:excel"><v:shape id="_x0000_s1025" type="#_x0000_t202">'
       '<v:textbox><div><font>旧形式の注記: 単位は mm</font></div></v:textbox>'
       '<x:ClientData ObjectType="Text"><x:Anchor>1, 0, 6, 0, 3, 0, 8, 0</x:Anchor></x:ClientData></v:shape></xml>')
inject_parts(out / "objects.xlsx",
             {"xl/diagrams/data1.xml": SMARTART_DATA, "xl/drawings/vmlDrawing9.vml": VML},
             {"xl/drawings/_rels/drawing94.xml.rels": f'<Relationship Id="rIdDgm" Type="{DGM_DATA_REL}" Target="../diagrams/data1.xml"/>',
              "xl/worksheets/_rels/sheet1.xml.rels": '<Relationship Id="rIdVml" Type="http://schemas.openxmlformats.org/'
              'officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing9.vml"/>'},
             [("xl/diagrams/data1.xml", DGM_DATA_TYPE),
              ("xl/drawings/vmlDrawing9.vml", "application/vnd.openxmlformats-officedocument.vmlDrawing")])

od = docx.Document()
od.add_heading("通常時の動作", level=1)
od.add_paragraph("応答がない場合は三回まで再送する。")
od.add_paragraph("送信の前に状態を確認する。")
p = od.add_paragraph("改訂: ")
struck_run = p.add_run("旧仕様では再送しない。")
struck_run.font.strike = True
od.add_heading("異常時の動作", level=1)
od.add_paragraph("応答がない場合は三回まで再送する。")
od.add_paragraph("終了要求の後は再送しない。")
from docx.opc.part import Part as DocxPart
from docx.opc.packuri import PackURI as DocxURI
from docx.oxml import parse_xml
dpart = DocxPart(DocxURI("/word/diagrams/data1.xml"), DGM_DATA_TYPE, SMARTART_DATA.encode("utf-8"), od.part.package)
drid = od.part.relate_to(dpart, DGM_DATA_REL)
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
sp = od.add_paragraph()
sp._p.append(parse_xml(
    f'<w:r xmlns:w="{W_NS}" xmlns:wp="{WP_NS}" xmlns:a="{A_NS}"><w:drawing><wp:inline><wp:extent cx="3000000" cy="1500000"/>'
    f'<wp:docPr id="50" name="図表 1"/><a:graphic><a:graphicData uri="{DGM_NS}"><dgm:relIds xmlns:dgm="{DGM_NS}" '
    f'xmlns:r="{R_NS}" r:dm="{drid}" r:lo="" r:qs="" r:cs=""/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'))
op = od.add_paragraph()
op._p.append(parse_xml(
    f'<w:r xmlns:w="{W_NS}" xmlns:o="urn:schemas-microsoft-com:office:office"><w:object>'
    '<o:OLEObject Type="Embed" ProgID="Excel.Sheet.12" ShapeID="_x0000_i1025" DrawAspect="Icon" ObjectID="_1"/>'
    "</w:object></w:r>"))
od.save(out / "objects.docx")

from pptx.opc.package import Part as PptxPart
from pptx.opc.packuri import PackURI as PptxURI
from pptx.oxml import parse_xml as pptx_parse
op_prs = Presentation()
sl = op_prs.slides.add_slide(op_prs.slide_layouts[5])
sl.shapes.title.text = "承認の流れ"
ppart = PptxPart(PptxURI("/ppt/diagrams/data1.xml"), DGM_DATA_TYPE, op_prs.part.package, SMARTART_DATA.encode("utf-8"))
prid = sl.part.relate_to(ppart, DGM_DATA_REL)
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
sl.shapes._spTree.append(pptx_parse(_smartart_frame("p", prid).replace(
    "<p:graphicFrame", f'<p:graphicFrame xmlns:p="{P_NS}" xmlns:a="{A_NS}"', 1).replace(' macro=""', "", 1)))
(out / "fig2.png").write_bytes(png)
pic = sl.shapes.add_picture(str(out / "fig2.png"), Inches(5), Inches(4))
pic._element.nvPicPr.cNvPr.set("descr", "承認フローの図")
(out / "fig2.png").unlink()
op_prs.save(out / "objects.pptx")

# --- plan.xml: Microsoft Project XML (what Project and MPXJ write for a .mpp) ---------------
def _ptask(uid, name, level, outline, start, finish, dur_h, pct=0, summary=0, milestone=0, critical=0,
           preds=(), base_finish=None, deadline=None, notes=None, custom=None):
    links = "".join(f"<PredecessorLink><PredecessorUID>{u}</PredecessorUID><Type>{t}</Type><LinkLag>{lag}</LinkLag>"
                    "<LagFormat>7</LagFormat></PredecessorLink>" for u, t, lag in preds)
    base = f"<Baseline><Number>0</Number><Start>{start}</Start><Finish>{base_finish}</Finish></Baseline>" if base_finish else ""
    extra = (f"<Deadline>{deadline}</Deadline>" if deadline else "") + (f"<Notes>{notes}</Notes>" if notes else "")
    ext = (f"<ExtendedAttribute><FieldID>188743731</FieldID><Value>{custom}</Value></ExtendedAttribute>" if custom else "")
    return (f"<Task><UID>{uid}</UID><ID>{uid}</ID><Name>{name}</Name><OutlineNumber>{outline}</OutlineNumber>"
            f"<OutlineLevel>{level}</OutlineLevel><Start>{start}</Start><Finish>{finish}</Finish>"
            f"<Duration>PT{dur_h}H0M0S</Duration><PercentComplete>{pct}</PercentComplete><Summary>{summary}</Summary>"
            f"<Milestone>{milestone}</Milestone><Critical>{critical}</Critical>{extra}{links}{base}{ext}</Task>")


plan = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Project xmlns="http://schemas.microsoft.com/project">'
        "<Name>plan.xml</Name><Title>受注システム更新</Title><Manager>山田</Manager><StartDate>2026-10-01T08:00:00</StartDate>"
        "<FinishDate>2026-11-30T17:00:00</FinishDate><StatusDate>2026-11-06T17:00:00</StatusDate><MinutesPerDay>480</MinutesPerDay>"
        "<ExtendedAttributes><ExtendedAttribute><FieldID>188743731</FieldID><FieldName>Text1</FieldName><Alias>担当部署</Alias>"
        "</ExtendedAttribute></ExtendedAttributes><Tasks>"
        "<Task><UID>0</UID><ID>0</ID><Name>受注システム更新</Name><OutlineLevel>0</OutlineLevel><Summary>1</Summary>"
        "<Work>PT400H0M0S</Work><ActualWork>PT150H0M0S</ActualWork><RemainingWork>PT250H0M0S</RemainingWork>"
        "<BCWS>800</BCWS><BCWP>600</BCWP><ACWP>700</ACWP></Task>"
        + _ptask(1, "設計", 1, "1", "2026-10-01T08:00:00", "2026-10-20T17:00:00", 112, pct=100, summary=1)
        + _ptask(2, "基本設計", 2, "1.1", "2026-10-01T08:00:00", "2026-10-09T17:00:00", 56, pct=100, custom="開発部")
        + _ptask(3, "詳細設計", 2, "1.2", "2026-10-12T08:00:00", "2026-10-20T17:00:00", 56, pct=100, preds=[(2, 1, 0)])
        + _ptask(4, "実装", 1, "2", "2026-10-21T08:00:00", "2026-11-30T17:00:00", 232, pct=30, summary=1)
        + _ptask(5, "画面の実装", 2, "2.1", "2026-10-21T08:00:00", "2026-11-05T17:00:00", 96, pct=60, critical=1,
                 preds=[(3, 1, 0)], base_finish="2026-10-30T17:00:00", deadline="2026-11-02T17:00:00",
                 notes="外部の画面部品の納品待ち。", custom="開発部")
        + _ptask(6, "帳票の実装", 2, "2.2", "2026-10-23T08:00:00", "2026-11-20T17:00:00", 160, pct=10, preds=[(5, 3, 9600)])
        + _ptask(7, "受入判定", 2, "2.3", "2026-11-30T17:00:00", "2026-11-30T17:00:00", 0, milestone=1, critical=1,
                 preds=[(6, 1, 0)])
        # a task outside the logic: no links, a typed-in start date, two days of slack, not started
        + "<Task><UID>8</UID><ID>8</ID><Name>外部調整</Name><OutlineNumber>2.4</OutlineNumber><OutlineLevel>2</OutlineLevel>"
        "<Start>2026-10-26T08:00:00</Start><Finish>2026-10-28T17:00:00</Finish><Duration>PT24H0M0S</Duration>"
        "<PercentComplete>0</PercentComplete><Summary>0</Summary><Milestone>0</Milestone><Critical>0</Critical>"
        "<ConstraintType>2</ConstraintType><ConstraintDate>2026-10-26T08:00:00</ConstraintDate>"
        "<TotalSlack>9600</TotalSlack></Task>"
        + "</Tasks><Resources><Resource><UID>0</UID><Name></Name></Resource><Resource><UID>1</UID><Name>佐藤</Name>"
        "<Type>1</Type><OverAllocated>1</OverAllocated><RemainingWork>PT40H0M0S</RemainingWork></Resource><Resource><UID>2</UID><Name>鈴木</Name><Type>1</Type></Resource></Resources>"
        "<Assignments><Assignment><TaskUID>5</TaskUID><ResourceUID>1</ResourceUID><Units>1</Units></Assignment>"
        "<Assignment><TaskUID>6</TaskUID><ResourceUID>2</ResourceUID><Units>0.5</Units></Assignment></Assignments></Project>")
(out / "plan.xml").write_text(plan, encoding="utf-8")
(out / "plan.mpp").write_bytes(b"not a real project file")

# --- history/: four saved copies of one plan, file names carry the date ---------------------
def _plan(status, tasks, finish):
    st = f"<StatusDate>{status}T17:00:00</StatusDate>" if status else ""
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Project xmlns="http://schemas.microsoft.com/project">'
            f"<Title>受注システム更新</Title><StartDate>2026-10-01T08:00:00</StartDate><FinishDate>{finish}T17:00:00</FinishDate>"
            f"{st}<MinutesPerDay>480</MinutesPerDay><Tasks>" + "".join(tasks) + "</Tasks></Project>")


def _hist(status, design_name, p3, p5, f5, f6, p6, milestone, extra=False):
    tasks = [_ptask(1, "設計", 1, "1", "2026-10-01T08:00:00", "2026-10-20T17:00:00", 112, summary=1),
             _ptask(2, design_name, 2, "1.1", "2026-10-01T08:00:00", "2026-10-09T17:00:00", 56, pct=100),
             _ptask(3, "詳細設計", 2, "1.2", "2026-10-12T08:00:00", "2026-10-20T17:00:00", 56, pct=p3, preds=[(2, 1, 0)]),
             _ptask(4, "実装", 1, "2", "2026-10-21T08:00:00", f"{milestone}T17:00:00", 200, summary=1),
             _ptask(5, "画面の実装", 2, "2.1", "2026-10-21T08:00:00", f"{f5}T17:00:00", 80, pct=p5, preds=[(3, 1, 0)],
                    custom="開発部")]
    if f6:
        tasks.append(_ptask(6, "帳票の実装", 2, "2.2", "2026-10-23T08:00:00", f"{f6}T17:00:00", 120, pct=p6))
    if extra:
        tasks.append(_ptask(8, "移行リハーサル", 2, "2.3", "2026-11-16T08:00:00", "2026-11-25T17:00:00", 64))
    tasks.append(_ptask(7, "受入判定", 2, "2.9", f"{milestone}T17:00:00", f"{milestone}T17:00:00", 0, milestone=1))
    return _plan(status, tasks, milestone).replace("<Tasks>", "<ExtendedAttributes><ExtendedAttribute><FieldID>188743731"
                                                  "</FieldID><Alias>担当部署</Alias></ExtendedAttribute></ExtendedAttributes><Tasks>", 1)


(out / "history").mkdir(exist_ok=True)
(out / "history" / "計画_20261009.xml").write_text(
    _hist("2026-10-09", "基本設計", 20, 0, "2026-10-30", None, 0, "2026-11-20"), encoding="utf-8")
# status date left at the previous report: should be flagged
(out / "history" / "計画_20261023.xml").write_text(
    _hist("2026-10-09", "基本設計", 100, 30, "2026-11-02", "2026-11-18", 0, "2026-11-27"), encoding="utf-8")
(out / "history" / "計画_20261106.xml").write_text(
    _hist("2026-11-06", "基本設計（改）", 100, 60, "2026-11-05", "2026-11-20", 10, "2026-11-30", extra=True), encoding="utf-8")
# no status date: ordered by the file name; nothing moved, so progress stalled
(out / "history" / "計画_20261113.xml").write_text(
    _hist(None, "基本設計（改）", 100, 60, "2026-11-05", "2026-11-20", 10, "2026-11-30", extra=True), encoding="utf-8")


# --- flow.vsdx: a Visio business flow with two lanes -------------------------------------------
# glued sequence flows, a dashed message line whose ends are only near its shapes, a decision
# label written as a separate text box, a data object riding on the message, an annotation.
VNS = "http://schemas.microsoft.com/office/visio/2012/main"
RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _vcells(**kv):
    return "".join(f'<Cell N="{k}" V="{v}"/>' for k, v in kv.items())


def _vshape(sid, text="", master=None, **cells):
    m = f' Master="{master}"' if master else ""
    t = f"<Text>{text}</Text>" if text else ""
    return f'<Shape ID="{sid}" Type="Shape"{m}>{_vcells(**cells)}{t}</Shape>'


def _vline(sid, master, x1, y1, x2, y2, text="", dashed=False):
    extra = {"LinePattern": 2} if dashed else {}
    return _vshape(sid, text, master, BeginX=x1, BeginY=y1, EndX=x2, EndY=y2, PinX=(x1 + x2) / 2,
                   PinY=(y1 + y2) / 2, Width=abs(x2 - x1) or 0.1, Height=abs(y2 - y1) or 0.1, **extra)


page = "".join([
    _vshape(1, "窓口", 1, PinX=2, PinY=4, Width=3, Height=8),            # lane: left column
    _vshape(2, "審査課", 1, PinX=6, PinY=4, Width=3, Height=8),          # lane: right column
    _vshape(10, "", 2, PinX=2, PinY=7, Width=0.5, Height=0.5),            # start event
    _vshape(11, "申請受付", 3, PinX=2, PinY=5.5, Width=1.5, Height=0.8),
    _vshape(12, "", 4, PinX=2, PinY=3.5, Width=0.8, Height=0.8),         # gateway
    _vshape(13, "受理通知", 3, PinX=2, PinY=1.5, Width=1.5, Height=0.8),
    _vshape(14, "内容審査", 3, PinX=6, PinY=3.5, Width=1.5, Height=0.8),
    _vshape(15, "申請書", 6, PinX=4, PinY=6.2, Width=0.5, Height=0.6),   # data object
    _vshape(16, "受理の場合", None, PinX=2.4, PinY=2.6, Width=0.9, Height=0.3),
    _vshape(17, "書類の不備は電話で確認する", None, PinX=0.6, PinY=5.5, Width=1.0, Height=0.5),
    _vline(20, 5, 2, 6.75, 2, 5.9),                        # start -> 申請受付 (glued)
    _vline(21, 5, 2, 5.1, 2, 3.9),                         # 申請受付 -> gateway (glued)
    _vline(22, 5, 2, 3.1, 2, 1.9),                         # gateway -> 受理通知 (glued, label nearby)
    _vline(23, 7, 2.8, 5.4, 5.2, 3.6, dashed=True),        # 申請受付 -> 内容審査, not glued
    _vline(24, 8, 4, 5.9, 4, 4.5, dashed=True),            # 申請書 rides on line 23
    _vline(25, 8, 1.12, 5.5, 1.25, 5.5, dashed=True),      # annotation -> 申請受付, not glued
])
connects = "".join(f'<Connect FromSheet="{a}" FromCell="{c}" ToSheet="{b}"/>' for a, c, b in [
    (20, "BeginX", 10), (20, "EndX", 11), (21, "BeginX", 11), (21, "EndX", 12), (22, "BeginX", 12),
    (22, "EndX", 13), (24, "BeginX", 15), (24, "EndX", 23)])
masters = "".join(f'<Master ID="{i}" NameU="{n}"/>' for i, n in [
    (1, "Pool / Lane"), (2, "Start Event"), (3, "Task"), (4, "Gateway"), (5, "Sequence Flow"),
    (6, "Data Object"), (7, "Message Flow"), (8, "Association")])
with zipfile.ZipFile(out / "flow.vsdx", "w", zipfile.ZIP_DEFLATED) as vz:
    vz.writestr("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8"?><Types '
                'xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>')
    vz.writestr("visio/masters/masters.xml", f'<Masters xmlns="{VNS}">{masters}</Masters>')
    vz.writestr("visio/pages/pages.xml", f'<Pages xmlns="{VNS}" xmlns:r="{RNS}"><Page ID="0" NameU="申請"><Rel '
                'r:id="rId1"/></Page></Pages>')
    vz.writestr("visio/pages/_rels/pages.xml.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/'
                'package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.microsoft.com/visio/'
                '2010/relationships/page" Target="page1.xml"/></Relationships>')
    vz.writestr("visio/pages/page1.xml", f'<PageContents xmlns="{VNS}"><Shapes>{page}</Shapes>'
                f"<Connects>{connects}</Connects></PageContents>")

# --- nested.docx: a table inside a table cell --------------------------------------------------
nd = docx.Document()
outer = nd.add_table(rows=1, cols=2)
outer.cell(0, 0).text = "申請者"
inner = outer.cell(0, 1).add_table(rows=2, cols=2)
inner.cell(0, 0).text = "氏名"; inner.cell(0, 1).text = "山田"
inner.cell(1, 0).text = "対象者との続柄"; inner.cell(1, 1).text = "本人"
nd.save(out / "nested.docx")

# --- csv ---
(out / "sample.csv").write_text("名前,年齢\n太郎,30\n花子,25\n", encoding="cp932")
print("samples written to", out)
