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

# --- csv ---
(out / "sample.csv").write_text("名前,年齢\n太郎,30\n花子,25\n", encoding="cp932")
print("samples written to", out)
