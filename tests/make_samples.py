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
ws3 = wb.create_sheet("big")
for r in range(300):
    ws3.append([r, r * 2])
wb.save(out / "sample.xlsx")

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
