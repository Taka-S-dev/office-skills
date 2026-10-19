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

# --- csv ---
(out / "sample.csv").write_text("名前,年齢\n太郎,30\n花子,25\n", encoding="cp932")
print("samples written to", out)
