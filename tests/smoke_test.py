#!/usr/bin/env python3
"""End-to-end smoke test for office-read and office-edit.

Generates sample files, runs every script in its main modes, and checks the output
contains what it should. Run after changing anything:

    python tests/smoke_test.py

Exit code 0 = all good. Needs: openpyxl python-docx python-pptx pypdf.
Excel COM / LibreOffice are optional; recalc and topdf are skipped when neither exists.
"""

from __future__ import annotations

import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
READ = ROOT / "office-read" / "scripts" / "office2md.py"
EDIT = ROOT / "office-edit" / "scripts"
PY = sys.executable

fails: list[str] = []


def run(*args, expect_ok=True) -> str:
    r = subprocess.run([PY, *map(str, args)], capture_output=True, text=True, encoding="utf-8")
    out = (r.stdout or "") + (r.stderr or "")
    if expect_ok and r.returncode != 0:
        fails.append(f"exit {r.returncode}: {' '.join(map(str, args))}\n{out[-800:]}")
    return out


def check(name: str, cond: bool, detail: str = ""):
    status = "ok  " if cond else "FAIL"
    print(f"  [{status}] {name}")
    if not cond:
        fails.append(f"{name}: {detail[-500:]}")


def main() -> int:
    with tempfile.TemporaryDirectory() as td:
        d = Path(td)
        run(ROOT / "tests" / "make_samples.py", d)
        xlsx, docx, pptx, csv = d / "sample.xlsx", d / "sample.docx", d / "sample.pptx", d / "sample.csv"

        print("office-read")
        o = run(READ, xlsx, "--outline")
        check("xlsx outline lists hidden sheet", "hidden_memo | hidden" in o, o)
        o = run(READ, xlsx, "--sheet", "売上", "--formulas")
        check("xlsx formulas shown", "=B2*C2" in o, o)
        check("xlsx cell comment shown", "商品名は正式名称で" in o, o)
        o = run(READ, xlsx, "--sheet", "big")
        check("xlsx truncation notice", "rows truncated at 200" in o, o)
        o = run(READ, xlsx, "--sheet", "売上")
        check("xlsx grid has column letters and row numbers", "|  | A | B |" in o and "| 6 | 合計 |" in o, o)
        check("xlsx uncached formula is marked, not blank", "=SUM(D2:D4) [uncached]" in o and "WARNING: 4 formula" in o, o)
        check("xlsx text box text extracted with anchor", "| F2 | shape | 注記 | 修正済み: 単価は税抜 |" in o, o)
        o = run(READ, xlsx, "--sheet", "売上", "--cells")
        check("xlsx --cells lists Sheet!A1 = value", "- 売上!B2 = 120" in o, o)
        o = run(ROOT / "office-read" / "scripts" / "xlsx_deps.py", xlsx, "--cell", "売上!D6")
        check("xlsx_deps precedents tree", "- 売上!D6 = =SUM(D2:D4)" in o and "- 売上!B2 = 120" in o, o)
        o = run(ROOT / "office-read" / "scripts" / "xlsx_deps.py", xlsx, "--cell", "売上!B2", "--dependents")
        check("xlsx_deps dependents tree", "  - 売上!D2 = =B2*C2" in o and "    - 売上!D6 = =SUM(D2:D4)" in o, o)
        rules = d / "checks.txt"
        rules.write_text("required 申請書!C2 申請書!D3\nequal 申請書!C2 申請書!C2\ncompare 売上!B2 < 売上!B4\n"
                         "regex 申請書!D3 ^[A-Z]{3}-\\d{3}$\nrequired 売上!E1\ncompare 申請書!F2 > 0\nno_errors\n",
                         encoding="utf-8")
        o = run(ROOT / "office-read" / "scripts" / "xlsx_check.py", xlsx, "--rules", rules, expect_ok=False)
        check("xlsx_check PASS / FAIL / UNKNOWN per rule",
              "| PASS | required 申請書!C2 申請書!D3 |" in o and "| FAIL | required 売上!E1 | empty: 売上!E1 |" in o
              and "| UNKNOWN | compare 申請書!F2 > 0 |" in o and "PASS: 5, FAIL: 1, UNKNOWN: 1" in o, o)
        rv = ROOT / "office-read" / "scripts" / "office_review.py"
        o = run(rv, d / "review.docx")
        check("office_review docx tracked changes + comment",
              "| 2 | insert | reviewer |" in o and "| 2 | delete | reviewer |" in o and "出典を追記してください" in o, o)
        o = run(READ, d / "review.docx")
        check("docx tracked changes and comments marked in place (CriticMarkup) with a change list",
              all(x in o for x in ["{++追加された文++}", "{--削除された文--}", "{>>reviewer: 出典を追記してください<<}",
                                   "## Tracked changes", "| inserted | 追加された文 |"]), o)
        o = run(READ, d / "review.docx", "--accept-changes")
        check("docx --accept-changes prints the accepted text without marks",
              "追加された文" in o and "削除された文" not in o and "{++" not in o and "{>>" not in o, o)
        o = run(READ, d / "merged.docx")
        check("docx header cell spanning two columns names both; groups spanning rows repeat",
              "| 区分 | 項目 | 項目 | 備考 |" in o and "| 基本 | 受注日 | 8 | 任意 |" in o, o)
        code = "\n".join([
            "import md_crosscheck as m",
            "same = lambda a, b: set(m.sentences(a)) == set(m.sentences(b)) != set()",
            "assert same('RFI【あ】......情報提供依頼書のこと。', 'RFI【あ】……情報提供依頼書のこと。')",
            "assert m.sentences('[はじめに [1](#section)](#section)') == {}",
            "assert m.sentences('[第１章 概要 2](#_Toc11)') == {}",
            "assert '3類型に分類する' in m.sentences('1. ３類型に分類する。')",  # the list number goes, the ３ stays (width folded)
            "assert same('| 全ての団体で、<br>必須機能である |', '全ての団体で、必須機能である')",
            "assert same('[RFI【あ】]{.underline}の説明文です。', 'RFI【あ】の説明文です。')",
            "assert m.tables('|  |  |\\n| --- | --- |\\n| a | b |\\n| c | d |\\n') == [2]",
            "assert m.headings('# 目次 {#目次 .TOC-Heading}\\n#\\n## はじめに\\n') == ['はじめに']",
            "assert same('前の{--古い--}{++新しい++}文です。{>>a: x<<}', '前の新しい文です。')",
            "assert m.tables('| 区分 | a |\\n|---|---|\\n| 区分 | |\\n| 区分 | b |\\n') == [2]",
            "assert m.sentences('&#10;</blockquote></td>') == {}",
            "assert same('[figure: x.png 2x2]様式第三号の説明文', '様式第三号の説明文')",
            "print('NORMALISATION OK')"])
        o = subprocess.run([PY, "-X", "utf8", "-c", code], capture_output=True, text=True, encoding="utf-8",
                           cwd=ROOT / "office-read" / "scripts")
        check("cross-check normalisation: leaders, TOC links, full-width digits, cell breaks, spans, rows, marks",
              "NORMALISATION OK" in o.stdout, o.stdout + o.stderr)
        o = run(rv, d / "review.xlsx")
        check("office_review xlsx threaded + legacy", "| メイン!B2 | reviewer |" in o and "reply to reviewer" in o and "| メイン!C2 | taro |" in o, o)
        o = run(rv, d / "review.pptx")
        check("office_review pptx comment", "| 1 | reviewer |" in o and "タイトルを見直す" in o, o)
        o = run(READ, xlsx, "--sheet", "申請書", "--cells")
        check("xlsx display formats (date, percent, thousands, time) with raw shown",
              all(x in o for x in ["申請書!C2 = 2026/10/01", "申請書!F5 = 15%  (raw 0.15)", "申請書!F6 = 1,234.5  (raw 1234.5)",
                                   "申請書!F7 = 9:05"]), o)
        check("xlsx hidden row and column flagged", "申請書!F9 = 非表示の値  [hidden]" in o and "申請書!H2 = 隠し列  [hidden]" in o
              and "hidden rows 9; hidden columns H" in o, o)
        o = run(READ, xlsx, "--sheet", "申請書")
        check("xlsx grid marks hidden row/col", "| 9 (h) |" in o and "| H (h) |" in o, o)
        o = run(READ, xlsx, "--outline")
        check("xlsx outline lists hidden rows/cols", "申請書: rows 9; columns H" in o, o)
        pdir = d / "profs"
        pdir.mkdir()
        (pdir / "sales.md").write_text("# 売上\n\n## 識別\n- match: sheets=売上,申請書\n", encoding="utf-8")
        (pdir / "other.md").write_text("# other\n\n## 識別\n- match: filename=^nomatch\n", encoding="utf-8")
        o = run(READ, xlsx, "--outline", "--profiles-dir", pdir)
        check("profile auto-match by sheets", "Matched profile: " in o and "sales.md (by sheets, given directory)" in o
              and "confirm with the user" in o and "other.md" not in o, o)
        o = run(READ, xlsx, "--profile")
        check("profile skeleton emits match lines", "- match: filename=^sample\\.xlsx$" in o and "- match: sheets=売上,hidden_memo,申請書,big" in o, o)
        o = run(ROOT / "office-read" / "scripts" / "office_find.py", "XYZ-100|発表者", d, "--ext", "xlsx,pptx")
        check("office_find reports file and address", "- 申請書!D3 | XYZ-100" in o and "- Slide 1 notes | ここは発表者ノート" in o, o)
        att = d / "att"
        o = run(READ, d / "mail3.eml", "--extract-attachments", att)
        check("eml headers, body and attachment extraction",
              "| from | 山田 <yamada@example.com> |" in o and "D6 の数式を D2:D5 に広げました" in o
              and (att / "月次報告_修正版.xlsx").exists(), o)
        o = run(ROOT / "office-read" / "scripts" / "mail_digest.py", d, "--threads", "--bodies", "40")
        check("mail_digest chronological thread table", "| 1 | 2026-10-06 09:00 |" in o and "| 3 | 2026-10-10 11:15 |" in o
              and "月次報告_修正版.xlsx |" in o and "## 月次報告の確認依頼  (4 mail(s))" in o, o)
        o = run(ROOT / "office-read" / "scripts" / "office_find.py", r"No\.1[12]", d, "--ext", "eml")
        check("office_find searches mail bodies", "- body line 1 | 評価シートの No.11" in o, o)
        o = run(READ, d / "添付資料.zip")
        check("archive listing", "資料/報告書.docx" in o and "inner.7z" in o and "2 file(s) inside" in o, o)
        ex = d / "ex"
        o = run(ROOT / "office-read" / "scripts" / "unpack.py", d / "添付資料.zip", "--out", ex)
        check("unpack zip with nested 7z", (ex / "資料" / "報告書.docx").exists() and (ex / "inner" / "内側" / "メモ.txt").exists(), o)
        import py7zr as _py7zr
        enc = d / "enc.7z"
        with _py7zr.SevenZipFile(enc, "w", password="pw") as z7:
            z7.writestr("秘密の内容", "機密.txt")
        o = run(ROOT / "office-read" / "scripts" / "unpack.py", enc, "--out", d / "enc_out",
                "--password", "wrong", "--password", "pw", expect_ok=True)
        check("unpack tries password candidates in order", "opened with candidate 2 of 2" in o and (d / "enc_out" / "機密.txt").exists(), o)
        o = run(ROOT / "office-read" / "scripts" / "unpack.py", enc, "--out", d / "enc_out2", "--password", "a", "--password", "b", expect_ok=False)
        check("unpack reports when no candidate works", "none of the 2 password candidate" in o, o)
        o = run(ROOT / "office-read" / "scripts" / "mail_digest.py", d, "--extract", d / "att2", "--deep")
        check("mail_digest --deep summarises attachments and archive contents",
              "| 3 | 月次報告_修正版.xlsx | 003_月次報告_修正版.xlsx | sheets: 売上 6x7 4f" in o
              and "資料/報告書.docx |" in o and "headings: 報告書タイトル / 1. 概要" in o, o)
        msg = ROOT / "tests" / "fixtures" / "sample.msg"
        if msg.exists():
            o = run(READ, msg, "--mail-engine", "olefile")
            check("msg via olefile: subject, body, attachment", "月次報告の修正依頼" in o and "No.11" in o and "見積条件.txt" in o, o)
        else:
            print("  [skip] msg fixture not present")
        dg = d / "diagram.xlsx"
        o = run(READ, dg, "--sheet", "シーケンス")
        check("xlsx sequence diagram reconstructed from shapes (order, direction, labels from text box and cell)",
              all(x in o for x in ["P1->>P2: 注文登録", "P2->>P3: 在庫照会", "P3->>P2: 在庫数", "P2->>P1: 登録完了",
                                   "label from cell G9", "participant P3 as DB"]), o)
        o = run(READ, dg, "--sheet", "フロー")
        check("xlsx flowchart connections, glued and touching",
              "- 開始 → 判定  (glued" in o and "- 判定 → 終了  (glued" in o and "- 判定 → 差戻し  (touching" in o, o)
        o = run(READ, dg, "--sheet", "レイアウト", "--blocks")
        check("xlsx record layout: merged spans and input rules (literal and range)",
              "区分 [B3:E3, 4 cols]" in o and "日付 [B4:I4, 8 cols]" in o and "Block A2:I5:" in o
              and "| K3:K5 | list: 必須, 任意 |" in o and "| L3:L5 | list: 1, 2, 3 (from $N$2:$N$4) |" in o, o)
        o = run(READ, dg, "--sheet", "グラフ")
        check("xlsx chart series and picture size", "売上: 1月=120, 2月=80, 3月=150" in o and "40x20" in o, o)
        rich = d / "seq_rich.xlsx"
        o = run(READ, rich, "--uml", "both", "--drawio", d / "rich.drawio", "--out", d / "rich.md")
        rm = (d / "rich.md").read_text(encoding="utf-8")
        check("sequence: activation, self message from 3 lines, alt/else with conditions, dashed replies, note (mermaid)",
              all(x in rm for x in ["P1->>P2: 注文登録\n    activate P2", "P2->>P2: 検証", "alt 在庫あり\n        P2->>P3: 引当",
                                    "else 在庫なし\n        P2-->>P1: 在庫切れ通知", "Note over P3: 読み取り専用",
                                    "P2-->>P1: 注文番号\n    deactivate P2"]), rm)
        check("sequence: same in PlantUML", all(x in rm for x in ["@startuml", "alt 在庫あり", "P2 --> P1 : 注文番号", "@enduml"]), rm)
        from lxml import etree as _et
        dio = _et.parse(str(d / "rich.drawio"), _et.XMLParser(resolve_entities=False)).getroot()
        dcells = dio.findall(".//mxCell")
        check("draw.io copy keeps lines, arrowheads, dashes and boxes",
              # 3 lifelines + 4 messages + 3 self-message segments + 1 else separator; 6 dashed (lifelines, separator, replies)
              sum(1 for c in dcells if c.get("edge")) == 11 and sum(1 for c in dcells if "dashed=1" in (c.get("style") or "")) == 6
              and any("在庫切れ通知" == c.get("value") for c in dcells) and "Editable copy" in rm, str(len(dcells)))
        frame = next((c for c in dcells if c.get("vertex") and not c.get("value") and "fillColor=none" in (c.get("style") or "")), None)
        check("draw.io copy keeps a transparent frame transparent", frame is not None, str([c.get("style") for c in dcells][:6]))
        lint = ROOT / "office-read" / "scripts" / "diagram_lint.py"
        o = run(lint, d / "rich.md")
        check("diagram_lint: generated diagrams are clean", "2 diagram block(s), 0 with problems" in o, o)
        (d / "broken.md").write_text(rm.replace("        P2->>P3: 引当", "        P2->>P4: 引当")
                                     .replace("    end\n    Note over P3", "    Note over P3", 1), encoding="utf-8")
        o = run(lint, d / "broken.md", expect_ok=False)
        check("diagram_lint: catches undeclared participant and unclosed alt",
              "undeclared participant 'P4'" in o and "'alt' block is never closed" in o, o)
        xc = ROOT / "office-read" / "scripts" / "md_crosscheck.py"
        run(xc, rich, "--out-dir", d / "xr", "--no-render")
        bm = (d / "xr" / "seq_rich.md").read_text(encoding="utf-8")
        (d / "xr" / "seq_rich.md").write_text(bm.replace("    end\n    Note over P3", "    Note over P3", 1), encoding="utf-8")
        (d / "xr" / "seq_rich.checked").write_text("シーケンス2#diagram ok\n", encoding="utf-8")
        run(xc, rich, "--out-dir", d / "xr", "--no-render")
        rv = (d / "xr" / "seq_rich.review.md").read_text(encoding="utf-8")
        check("cross-check refuses READY while a diagram block is broken",
              "NEEDS REVIEW" in rv and "## Diagram problems" in rv and "never closed" in rv and (d / "xr" / "seq_rich.drawio").exists(), rv)
        o = run(xc, dg, "--out-dir", d / "xx", "--no-render")
        rv = (d / "xx" / "diagram.review.md").read_text(encoding="utf-8")
        check("xlsx cross-check: nothing missing, 4 visual checks open",
              "[ok] Text: " in rv and "Visual checks: 0 of 4 confirmed" in rv and "NEEDS REVIEW" in rv, rv)
        ids = re.findall(r"`([^`]+@[0-9a-f]{8})`", rv.split("## Visual checks", 1)[1])
        (d / "xx" / "diagram.checked").write_text("シーケンス#diagram ok\nグラフ#chart1@00000000 ok\n", encoding="utf-8")
        run(xc, dg, "--out-dir", d / "xx", "--no-render")
        rv = (d / "xx" / "diagram.review.md").read_text(encoding="utf-8")
        check("visual check ids match exactly, with the content fingerprint (no prefix, no old id)",
              "Visual checks: 0 of 4 confirmed, 2 recorded for content that has since changed" in rv, rv[:1500])
        dio_text = (d / "xx" / "diagram.drawio").read_text(encoding="utf-8")
        check("draw.io chart placeholder carries the series values", "売上: 1月=120, 2月=80, 3月=150" in dio_text, dio_text[:300])
        md = (d / "xx" / "diagram.md").read_text(encoding="utf-8")
        (d / "xx" / "diagram.md").write_text(md.replace("日付 [B4:I4, 8 cols]", ""), encoding="utf-8")
        run(xc, dg, "--out-dir", d / "xx", "--no-render")
        rv = (d / "xx" / "diagram.review.md").read_text(encoding="utf-8")
        check("xlsx cross-check catches text dropped from the Markdown", "- レイアウト!B4: 日付" in rv, rv)
        (d / "xx" / "diagram.md").write_text(md, encoding="utf-8")
        (d / "xx" / "diagram.checked").write_text("".join(f"{i}  looked at the rendered sheet\n" for i in ids),
                                                    encoding="utf-8")
        run(xc, dg, "--out-dir", d / "xx", "--no-render")
        rv = (d / "xx" / "diagram.review.md").read_text(encoding="utf-8")
        check("xlsx cross-check READY once text is complete and visual checks are recorded", "**Status: READY**" in rv, rv)
        # an edited mermaid arrow is listed against the reconstruction
        m = re.search(r"^(\s*)(\w+)->>(\w+): (.+)$", md, re.M)
        (d / "xx" / "diagram.md").write_text(
            md.replace(m.group(0), f"{m.group(1)}{m.group(3)}->>{m.group(2)}: {m.group(4)}", 1), encoding="utf-8")
        run(xc, dg, "--out-dir", d / "xx", "--no-render")
        rv = (d / "xx" / "diagram.review.md").read_text(encoding="utf-8")
        check("a reversed arrow in the Markdown is listed as an edit of the reconstruction",
              "## Diagram edits since the reconstruction (1)" in rv and "the drawing had it the other way" in rv, rv)
        (d / "xx" / "diagram.md").write_text(md, encoding="utf-8")
        # the workbook changes: the Markdown is reported stale, only the changed diagram reopens
        (d / "v2").mkdir()
        with zipfile.ZipFile(dg) as zin, zipfile.ZipFile(d / "v2" / "diagram.xlsx", "w", zipfile.ZIP_DEFLATED) as zout:
            for n in zin.namelist():
                b = zin.read(n)
                zout.writestr(n, b.replace("在庫数".encode(), "在庫件数".encode()) if n == "xl/drawings/drawing91.xml" else b)
        run(xc, d / "v2" / "diagram.xlsx", "--out-dir", d / "xx", "--no-render")
        rv = (d / "xx" / "diagram.review.md").read_text(encoding="utf-8")
        check("a changed workbook makes the Markdown stale and reopens only the changed visual check",
              "[FAIL] Source: the Markdown was made from another version" in rv
              and "Visual checks: 3 of 4 confirmed, 1 recorded" in rv and "- [ ] `シーケンス#diagram@" in rv
              and "NEEDS REVIEW" in rv, rv[:2500])
        # objects.xlsx: SmartArt, legacy text box, chart sheet, printed header, strikethrough
        ob = d / "objects.xlsx"
        o = run(READ, ob)
        check("xlsx reads SmartArt in order with nesting, a legacy text box, a chart sheet, the printed header, ~~struck~~",
              all(x in o for x in ["  - 受付\n  - 審査\n    - 書類確認\n  - 承認", "旧形式の注記: 単位は mm",
                                   "## Chart sheet: グラフ専用", "数量の推移", "Printed header: 文書番号 DOC-001 第2版",
                                   "~~旧品番~~"]) and "表示用" not in o, o)
        check("xlsx shows numbers as formatted (a negative in brackets, a unit in the format), one struck word, page fields",
              all(x in o for x in ["| 割引 | (1,200) | ~~旧単価~~ 新単価 |", "| 件数 | 5件 |",
                                   "Printed footer: [page] / [pages] ページ", "(2 cell(s): B4, C5)"]), o)
        code = "\n".join([
            "from office2md import format_number as f",
            "assert f(1200, '\"\\\\\"#,##0') == '¥1,200' and f(0, '#,##0;-#,##0;\"-\"') == '-'",
            "assert f(1234567, '#,##0,') == '1,235' and f(1200, 'General') == 1200 and f(-5, '0') == '-5'",
            # an empty section hides, a weekday code keeps the serial, conditions pick the section,
            # half rounds up, leading zeros stay, a negative General stays a number
            "assert f(-5, '0;;') == '' and f(45000, 'aaa') == 45000 and f(-1.5, 'General') == -1.5",
            "assert f(500, '[>=100]\"big\";0') == 'big' and f(5, '[>=100]\"big\";0') == '5'",
            "assert f(1234.5, '#,##0') == '1,235' and f(12, '00000') == '00012'",
            "assert f(0, '_(\"$\"* #,##0_);_(\"$\"* (#,##0);_(\"$\"* \"-\"??_);_(@_)') == '$-'",
            "print('FORMAT OK')"])
        o = subprocess.run([PY, "-X", "utf8", "-c", code], capture_output=True, text=True, encoding="utf-8",
                           cwd=ROOT / "office-read" / "scripts")
        check("xlsx number formats: yen written with a backslash, a zero section, thousands scaling, General kept",
              "FORMAT OK" in o.stdout, o.stdout + o.stderr)
        run(xc, ob, "--out-dir", d / "xo", "--no-render")
        rv = (d / "xo" / "objects.review.md").read_text(encoding="utf-8")
        check("SmartArt and the chart sheet become visual checks", "`部品#smartart1@" in rv and "`グラフ専用#chart1@" in rv
              and "[ok] Text: " in rv, rv[:2000])
        om = (d / "xo" / "objects.md").read_text(encoding="utf-8")
        (d / "xo" / "objects.md").write_text(om.replace("| 未定 |", "|  |", 1).replace("| ボルト | 2 |", "| ボルト | 3 |"),
                                             encoding="utf-8")
        run(xc, ob, "--out-dir", d / "xo", "--no-render")
        rv = (d / "xo" / "objects.review.md").read_text(encoding="utf-8")
        check("text check catches one of two equal values removed and a short number changed",
              "未定 (2 times in the sheet, 1 in its section)" in rv
              and "- 部品!B2: 2 (not found on a line with its row label 'ボルト')" in rv, rv[:2500])
        o = run(READ, xlsx, "--sheet", "申請書", "--form")
        check("xlsx --form pairs label right / merged / below / header row",
              all(x in o for x in ["- 申請日 = 2026/10/01  (申請書!B2 → C2)", "- 対象機種 = XYZ-100  (申請書!B3 → D3)",
                                   "- 氏名 = 山田  (申請書!B5 → B6)", "- 所属 = 品質保証部  (申請書!C5 → C6)",
                                   "- 備考 = 特になし  (申請書!B8 → B9)"]), o)
        o = run(READ, docx)
        check("docx header + table + list", "社外秘" in o and "| 項目 | 値 |" in o and "- 箇条書き その1" in o, o)
        spec = d / "spec.docx"
        o = run(READ, spec, "--outline")
        check("spec.docx outline carries automatic heading numbers",
              "- 第1章 概要" in o and "  - 1.1 目的" in o and "    - 2.1.1 件数項目" in o, o)
        o = run(READ, spec, "--images", d / "img")
        check("spec.docx: nested bullets, ordered list, footnote, figure+alt, caption, text box, TOC, equation",
              all(x in o for x in ["- 任意項目\n  - 省略可能な拡張項目", "1. 注文を受け付ける\n2. 在庫を確認する",
                                   "[^1]", "[^1]: 旧版（第2版）の形式は対象外。",
                                   "![データの全体構成図](", "figure-001.png) _(1x1, placeholder image, no visual content)_",
                                   "*図 1 データの構成*", "> **Text box:** 注記: 改訂時は見出し部も見直すこと",
                                   "(table of contents field, 3 entries)", "T = P + N × 4"])
              and (d / "img" / "figure-001.png").exists(), o)
        check("spec.docx keeps hyperlink URLs", "[注文 API 仕様 v3](https://example.com/spec/order-api-v3)" in o, o)
        if shutil.which("pandoc"):
            o = run(ROOT / "office-read" / "scripts" / "md_crosscheck.py", spec, "--out-dir", d / "xc", "--no-markitdown")
            rv = (d / "xc" / "spec.review.md").read_text(encoding="utf-8")
            pm = (d / "xc" / "spec.md").read_text(encoding="utf-8")
            check("md_crosscheck: nothing in pandoc missing from the primary",
                  "## Content in pandoc not found in the primary (0)" in rv and (d / "xc" / "spec.md").exists(), o + rv)
            check("md_crosscheck: safe fixes applied (equation to LaTeX, bold) and status block present",
                  "$$" in pm and "**30日以内に入荷**" in pm and "**Status: " in rv and "Applied automatically" in rv, rv + pm)
            check("md_crosscheck: READY when only a reference's dropped title differs",
                  "**Status: READY**" in rv, rv)
            check("md_crosscheck: figure as an image link relative to the .md with forward slashes, size shown",
                  "![データの全体構成図](spec_figures/figure-001.png) _(1x1, placeholder image, no visual content)_" in pm, pm)
            (d / "xc" / "spec.md").write_text(pm + "\n手で追記した行。\n", encoding="utf-8")
            o = run(ROOT / "office-read" / "scripts" / "md_crosscheck.py", spec, "--out-dir", d / "xc", "--no-markitdown")
            check("md_crosscheck: a rerun keeps manual edits",
                  "reusing existing spec.md" in o and "手で追記した行。" in (d / "xc" / "spec.md").read_text(encoding="utf-8"), o)
            od = d / "objects.docx"
            run(xc, od, "--out-dir", d / "xd", "--no-markitdown", "--no-render")
            rv = (d / "xd" / "objects.review.md").read_text(encoding="utf-8")
            check("Word SmartArt and embedded object are visual checks; content check passes",
                  "`smartart1@" in rv and "`embedded1@" in rv and "[ok] Content against pandoc" in rv
                  and "NEEDS REVIEW" in rv, rv[:2500])
            wm = (d / "xd" / "objects.md").read_text(encoding="utf-8")
            rep = "応答がない場合は三回まで再送する。"
            cut = wm.rindex(rep)
            wm2 = wm[:cut] + wm[cut + len(rep):]
            wm2 = wm2.replace("終了要求の後は再送しない。\n", "").replace(
                "送信の前に状態を確認する。\n", "送信の前に状態を確認する。\n\n終了要求の後は再送しない。\n")
            (d / "xd" / "objects.md").write_text(wm2, encoding="utf-8")
            run(xc, od, "--out-dir", d / "xd", "--no-markitdown", "--no-render")
            rv = (d / "xd" / "objects.review.md").read_text(encoding="utf-8")
            check("Word cross-check catches a repeated sentence lost and lists one moved to another heading",
                  "1 repeated fewer times" in rv and "(2 times in pandoc, 1 in the primary)" in rv
                  and "## Content under a different heading than in pandoc (1)" in rv, rv[:3000])
        else:
            print("  [skip] md_crosscheck (pandoc not installed)")
        o = run(READ, d / "objects.docx")
        check("docx strikethrough, SmartArt with nesting, embedded object",
              all(x in o for x in ["~~旧仕様では再送しない。~~", "> - 審査\n>   - 書類確認", "**Embedded object** Excel.Sheet.12"]), o)
        plan = d / "plan.xml"
        o = run(READ, plan, "--gantt")
        check("Project XML: tasks with hierarchy, link types and lag, resources, custom field, notes",
              all(x in o for x in ["Read with: Project XML file.", "| 6 | 2.2 | 　帳票の実装 |", "5SS+2d", "鈴木 50%",
                                   "| 2 | 基本設計 | 開発部 |", "外部の画面部品の納品待ち。", "**設計**"]), o)
        check("Project XML: late, behind baseline, past deadline, critical; gantt",
              all(x in o for x in ["- 5 画面の実装 — finish 2026-11-05, 60% done",
                                   "baseline 2026-10-30 (+6 days)", "deadline 2026-11-02",
                                   "Critical and not complete: 2", "受入判定 :milestone, crit, t7"]), o)
        check("Project: behind expected progress before the finish passes, not started, SPI/CPI",
              all(x in o for x in ["- 6 帳票の実装 — 10% done, 51% expected by now", "Should have started and have not: 1",
                                   "SPI 0.75", "CPI 0.86"]), o)
        check("Project: schedule health (no links, fixed date), finish-driving chain, near-critical, overallocation, work",
              all(x in o for x in ["No predecessor: nothing drives its start", "- 8 外部調整 — must start on 2026-10-26",
                                   "3. 5 画面の実装 — 2026-10-21 → 2026-11-05, 60% (critical)",
                                   "- 8 外部調整 — 2 days of slack", "Over-allocated (Project's own flag",
                                   "Work: 400 h planned, 150 h done, 250 h remaining (38% of the work)."]), o)
        o2 = run(ROOT / "office-read" / "scripts" / "mpp_reader.py", plan, "--gantt")
        check("mpp_reader runs on its own with the same output", o2.split("_Source:")[1].strip() == o.split("_Source:")[1].strip(), o2[:300])
        hist = ROOT / "office-read" / "scripts" / "mpp_history.py"
        h = d / "history"
        o = run(hist, h / "計画_20261009.xml", h / "計画_20261106.xml")
        check("mpp_history diff: added, renamed, finish moves, progress; a renamed predecessor is not a link change",
              all(x in o for x in ["Plan finish: 2026-11-20 → 2026-11-30 (+10 days)", "- 8 移行リハーサル under '実装'",
                                   "- 2 '基本設計' → '基本設計（改）'", "5 画面の実装: finish 2026-10-30 → 2026-11-05 (+6 days)",
                                   "3 詳細設計: 20% → 100%"]) and "Dependencies changed" not in o, o)
        o = run(hist, h, "--cache", d / "mppcache")
        check("mpp_history timeline: ordered by status date, stale and missing status dates flagged, worse periods",
              all(x in o for x in ["| 2026-10-23 | 計画_20261023.xml |", "the status date was probably not updated",
                                   "no status date; placed and judged for lateness at its file name",
                                   "2026-10-23 → 2026-11-06: late tasks 0 → 1", "(stalled)", "- 2026-11-06: 移行リハーサル",
                                   "- 受入判定: 2 time(s), 2026-11-20 → 2026-11-30 (+10 days)", "xychart-beta",
                                   "| expected | behind expected |", "## Slips by 担当部署", "| 開発部 | 1 | 2 | 6 |"]), o)
        o2 = run(hist, h, "--cache", d / "mppcache")
        check("mpp_history --cache: a rerun reuses the conversions and gives the same result",
              len(list((d / "mppcache").glob("*.xml"))) == 4 and "cached conversion" in o2
              and o2.split("## Snapshots")[1] == o.split("## Snapshots")[1],
              o2[:400])
        o = run(READ, d / "plan.mpp", expect_ok=False)
        check(".mpp without Project or MPXJ says what to install instead of a traceback",
              "cannot read plan.mpp" in o and "Save As > XML" in o and "Traceback" not in o, o)
        o = run(READ, pptx)
        check("pptx notes + table", "発表者ノート" in o and "| x | y |" in o, o)
        o = run(READ, d / "objects.pptx", "--images", d / "pimg")
        check("pptx SmartArt text and picture alt text, picture extracted",
              "> - 審査\n>   - 書類確認" in o and "— 承認フローの図" in o and any((d / "pimg").glob("figure-*.png")), o)
        o = run(READ, d / "flow.vsdx")
        check("vsdx: lanes, glued flow in order, loose dashed message by position carrying data, decision label, annotation",
              all(x in o for x in ["Lanes: 窓口 · 審査課", "- [窓口] (start) → [窓口] 申請受付",
                                   "- [窓口] 申請受付 → [審査課] 内容審査 carrying 申請書 (message) (end found by position",
                                   "[受理の場合] (label: nearest text)", "- 申請受付: 書類の不備は電話で確認する",
                                   'subgraph L2["審査課"]', '-->|"受理の場合"|']), o)
        o = run(READ, d / "nested.docx")
        check("docx: a table nested in a cell is read, one line per row",
              "氏名 \\| 山田<br>対象者との続柄 \\| 本人" in o, o)
        o = run(READ, d / "ruby.docx")
        check("docx: furigana written ｜base《reading》, the word stays whole",
              "SNS等を通じた｜誹謗《ひぼう》中傷への対策を進める。" in o, o)
        o = run(READ, d / "shapes.docx")
        check("docx: （１） in body style is a subsection heading, a line break stays a break",
              "### 対象業務\n\n#### （１）申請の受付" in o and "一行目の文  \n二行目の文" in o, o)
        check("docx: （n） siblings share a level; a run of them with nothing between stays paragraphs",
              "\n#### （２）審査の流れ" in o and "\n（１）概要について\n\n（２）詳細について" in o, o)
        check("docx: a bold ④ after the ③ heading is its sibling; a borderless layout table has no header row",
              "\n##### ③ 第三の施策" in o and "\n##### ④ 第四の施策" in o and "|  |  |\n|---|---|\n| ①開放 | ③安全 |" in o, o)
        check("docx: a whole-bold short line stays bold; a one-column table under its caption stays a table",
              "\n**ア 小項目の見出し**\n" in o and "表1-1　要件の一覧\n\n|  |\n|---|\n| 一つ目の要件 |\n| 二つ目の要件 |" in o, o)
        check("docx: a cell and its text inside content controls are read; the table's alt-text title shown",
              "Table: 回答の表" in o and "| 回答欄 | ここに入力してください |" in o, o)
        check("docx: a table broken over a page is joined, the repeated header dropped",
              "| 期限 | 月末 |\n| 提出先 | 窓口 |" in o, o)
        check("docx: a one-column box is a quoted block, a spanning title row sits above its table, spacer column dropped",
              "> 留意事項  \n> 箱の中の二行目" in o and "\n申請日\n\n|  |  |\n|---|---|\n| 氏名 | 山田 |" in o, o)
        check("docx: text box in a cell marked, radical form read as the ideograph, notes numbered as displayed",
              "[text box] 枠内の注記" in o and "地方の制度[^1][^2]" in o and "[^2]: 二番目の注" in o, o)
        check("docx: a heading below Markdown's six levels is a bold line", "\n**第七段の見出し**\n" in o and "#######" not in o, o)
        o = run(READ, d / "drawings.docx")
        check("docx: figures drawn with shapes are marked where they sit, a lone text box is not one",
              "[drawing: canvas, 2 shapes, 1 line or connector]\n\n> **Text box:** 受付" in o
              and "図2 印だけの図\n\n[drawing: 1 shape]\n\n> **Text box:** 単独の注記" in o
              and "図3 古い形式の線\n\n[drawing: 1 line or connector]" in o
              and "3 drawing(s) made of shapes" in o, o)
        o = run(READ, d / "drawings.docx", "--outline")
        check("docx outline counts the drawings", "· 3 drawings ·" in o, o)
        if sys.platform == "win32":
            o = run(READ, d / "drawings.docx", "--images", d / "drawn", "--out", d / "drawn.md")
            dm = (d / "drawn.md").read_text(encoding="utf-8") if (d / "drawn.md").exists() else o
            check("docx --images: each drawing rendered through Word as a PNG and linked, or the reason said",
                  ("![drawing: canvas, 2 shapes, 1 line or connector](drawn/drawing-001.png)" in dm
                   and (d / "drawn" / "drawing-003.png").exists()) or "without a picture (Word" in dm, dm)
            o = run(READ, d / "emf.docx", "--images", d / "emf_img", "--out", d / "emf.md")
            em = (d / "emf.md").read_text(encoding="utf-8")
            check("docx --images: an EMF picture saved as PNG and linked as an image",
                  "![Picture 1](emf_img/figure-001.png)" in em and (d / "emf_img" / "figure-001.png").exists(), em)
        code = "\n".join([
            "import review_checks as rc, tempfile, os",
            "from pathlib import Path",
            "pages = ['見出し\\n本文の一文目です 第二の項目名\\n1', '見出し\\n脚注つきの文です2\\n2', '見出し\\nもう一つの行です\\n3']",
            "md = '本文の一文目です\\n\\n第二の項目名\\n\\n脚注つきの文です[^2]\\n'",
            "miss = rc.print_missing(pages, md)",
            "assert [m[1] for m in miss] == ['もう一つの行です'] and miss[0][2] == 3, miss",
            "spans = rc.page_spans(md, pages)",
            "assert spans[0][:2] == (1, 1) and spans[1][:3] == (2, 5, 5) and spans[2][1] is None, spans",
            "iid = rc.item_id('missing', 'もう一つの行です')",
            "acc = rc.accepted([f'accept {iid} printed only in the page footer', 'accept missing:00000000'])",
            "assert acc == {iid: 'printed only in the page footer'}, acc",
            "p = Path(tempfile.mkdtemp()) / 'x.rounds.jsonl'",
            "rc.record_round(p, {'missing': 2}); rows = rc.record_round(p, {'missing': 0})",
            "assert [r['round'] for r in rows] == [1, 2] and rc.rounds_table(rows, ['missing'])[-2] == '| 2 | 0 |'",
            "secs = rc.sheet_sections('## Sheet: CUI Overlay Overview\\na\\n## Sheet: CUI Overlay\\nb', ['CUI Overlay'])",
            "assert secs['CUI Overlay'].endswith('\\nb'), secs",
            "import md_crosscheck as m",
            "assert set(m.sentences('<p>&lt;Insert CSP Name&gt;</p>')) == set(m.sentences('<Insert CSP Name>'))",
            "assert set(m.sentences('図 １‑1　適用対象の範囲')) == set(m.sentences('図 1-1　適用対象の範囲'))",
            "assert m.TOC_LINE.match('> [改定履歴 [ii](#a)](#a)')",
            # a level-7 heading pandoc writes with seven # is office2md's bold line; a drawing's image
            # link leaves no text behind
            # bold lines count only as a match for the other side's heading; **a** and **b** is no line of bold
            "assert m.headings('####### 第七段\\n**18.1.1 第七段**\\n') == ['第七段']",
            "assert m.headings('####### 第七段\\n**18.1.1 第七段**\\n**a** and **b**\\n', bold=True) == ['第七段', '18.1.1 第七段']",
            "assert set(m.sentences('####### 第七段の見出し')) == set(m.sentences('**18.1.1.1.1.1.1. 第七段の見出し**'))",
            "assert m.sentences('![drawing: 2 shapes](x/drawing-001.png)') == {}",
            "assert m._head_key('II. an assessment') == m._head_key('an assessment')",
            "assert m._head_key('until 31^st^ March') == m._head_key('until 31st March')",
            "fixed, _ = m.apply_safe('#### （１）管理項目について\\n\\n索引の行 （１）管理項目について\\n\\n大切な語を含む文\\n',",
            "                        {'r': '**（１）管理項目について**\\n\\n**大切な語**を含む文\\n'})",
            "assert '**（１）' not in fixed and '**大切な語**を含む文' in fixed, fixed",  # a heading's bold stays off its index copy
            "print('AGENT LOOP OK')"])
        o = subprocess.run([PY, "-X", "utf8", "-c", code], capture_output=True, text=True, encoding="utf-8",
                           cwd=ROOT / "office-read" / "scripts")
        check("review loop: printed text missing by page, page spans, accept needs a reason, rounds recorded",
              "AGENT LOOP OK" in o.stdout, o.stdout + o.stderr)
        code = "\n".join([
            "import xlsx_drawing as xd",
            "def sh(name, box, text='', prst='rect', **k): return {'kind': 'shape', 'name': name, 'box': box, 'text': text, 'prst': prst, **k}",
            "a = sh('a', (120, 55, 230, 91), '申請書提出'); b = sh('b', (270, 165, 380, 201), '受付'); c = sh('c', (120, 275, 230, 311), '差し戻し')",
            "bands = [sh('l1', (20, 20, 580, 130)), sh('n1', (20, 20, 90, 130), '申請者'),",
            "         sh('l2', (20, 130, 580, 240)), sh('n2', (20, 130, 90, 240), '窓口')]",
            "lab = sh('t', (240, 120, 280, 136), 'あり', prst='', textbox=True)",
            "l1 = {'name': 'c1', 'start': (230, 73), 'end': (270, 183), 'head': True, 'tail': False}",
            "l2 = {'name': 'c2', 'start': (270, 201), 'end': (230, 293), 'head': True, 'tail': False}",
            "e = lambda f, t, x, y, l: {'from': f, 'to': t, 'label': '', 'both': False, 'none': False, '_a': x, '_b': y, '_l': l}",
            "r = {'edges': [e('申請書提出', '受付', a, b, l1), e('受付', '差し戻し', b, c, l2)]}",
            "xd._flow_labels_and_lanes(r, [a, b, c, lab] + bands, 64.0)",
            "assert [x['label'] for x in r['edges']] == ['あり', ''] and r['edges'][0]['label_by_position'], r",
            "assert r['lanes'] == {'申請書提出': '申請者', '受付': '窓口'} and '_a' not in r['edges'][0], r",
            "m = xd.mermaid_flow(r['edges'], r['lanes'])",
            "assert 'subgraph L1[\"申請者\"]' in m and 'subgraph L2[\"窓口\"]' in m, m",
            "print('FLOW OK')"])
        o = subprocess.run([PY, "-X", "utf8", "-c", code], capture_output=True, text=True, encoding="utf-8",
                           cwd=ROOT / "office-read" / "scripts")
        check("xlsx flow: a branch label beside the line, lanes from labelled bands, subgraph per lane",
              "FLOW OK" in o.stdout, o.stdout + o.stderr)
        code = "\n".join([
            "import xlsx_drawing as xd",
            "class M:",
            "    def cell_at(self, x, y): return f'{xd._col_letter(int(x // 10) + 1)}{int(y // 10) + 1}'",
            "cells = [{'address': 'B14', 'text': '作業'}, {'address': 'D14', 'text': '４～６月'}, {'address': 'F14', 'text': '７～９月'},",
            "         {'address': 'B16', 'text': '設計'}]",
            "items = [{'address': c['address'], 'text': c['text']} for c in cells]",
            "bar = {'start': (31, 155), 'end': (60, 155), 'head': True, 'tail': True}",
            "r = {'unexplained': [{'line': '直線 1', 'from': 'D16', 'to': 'G16', 'arrow': 'both', '_l': bar}], 'edges': []}",
            "xd._period_bars(r, items, M(), 10.0, 10.0)",
            "b = r['bars'][0]",
            "assert (b['label'], b['from'], b['to'], b['from_head'], b['to_head']) == ('設計', 'D16', 'F16', '４～６月', '７～９月'), b",
            "assert r['unexplained'] == [], r",
            "from lxml import etree",
            "A = 'http://schemas.openxmlformats.org/drawingml/2006/main'",
            "geo = etree.fromstring(f'<a:prstGeom xmlns:a=\"{A}\" prst=\"borderCallout1\"><a:avLst>'",
            "                      '<a:gd name=\"adj1\" fmla=\"val 50000\"/><a:gd name=\"adj2\" fmla=\"val 0\"/>'",
            "                      '<a:gd name=\"adj3\" fmla=\"val 150000\"/><a:gd name=\"adj4\" fmla=\"val -50000\"/></a:avLst></a:prstGeom>')",
            "assert xd._callout_tip('borderCallout1', geo, (0, 0, 100, 100)) == (-50.0, 150.0)",
            "assert xd._callout_tip('wedgeRectCallout', None, (0, 0, 100, 100)) == (50 - 20.833, 50 + 62.5)",
            "assert xd._callout_tip('rect', None, (0, 0, 100, 100)) is None",
            "print('BARS OK')"])
        o = subprocess.run([PY, "-X", "utf8", "-c", code], capture_output=True, text=True, encoding="utf-8",
                           cwd=ROOT / "office-read" / "scripts")
        check("xlsx schedule bars with row label and month headings; a callout's pointer tip from its adjust values",
              "BARS OK" in o.stdout, o.stdout + o.stderr)
        code = "\n".join([
            "import pdf_blocks as b",
            "runs = [{'x': 110, 'y': 300 - 14 * i, 'size': 10, 'text': t} for i, t in enumerate(['操作権限', '設定・管', '理に関す', 'ること'])]",
            "runs += [{'x': 176, 'y': 300 - 14 * i, 'size': 10, 'text': t} for i, t in enumerate(['操作権限設定・', '管理は、すべて', 'の基幹業務シス', 'テムで必要。'])]",
            "bl = b.blocks(b.segments(runs))",
            "texts = sorted(x['lines'][0] for x in bl)",
            "assert texts == ['操作権限設定・管理に関すること', '操作権限設定・管理は、すべての基幹業務システムで必要。'], texts",
            "print('COLUMNS OK')"])
        o = subprocess.run([PY, "-X", "utf8", "-c", code], capture_output=True, text=True, encoding="utf-8",
                           cwd=ROOT / "office-read" / "scripts")
        check("pdf --columns: narrow cells printed a few characters a line are joined per column",
              "COLUMNS OK" in o.stdout, o.stdout + o.stderr)
        code = "\n".join([
            "import md_crosscheck as m",
            "same = lambda a, b: set(m.sentences(a)) == set(m.sentences(b)) != set()",
            "assert same('SNS等を通じた｜誹謗《ひぼう》中傷の対策。', 'SNS等を通じた中傷の対策。')",
            "assert same('| 運営する[^97]という役割。 |', '運営するという役割。')",
            "assert same('平成14年法律第151号 [↑](#footnote-ref-11)', '平成14年法律第151号')",
            "assert same('推進すること（“DX with Cybersecurity”）が重要。', '推進すること（\"DX with Cybersecurity\"）が重要。')",
            "from pathlib import Path",
            "m.SUBSTRING_BUDGET = 0",
            "text, _ = m.review('# a\\n\\n一つめの文です。\\n', {'pandoc': '一つめの文です。別の長い文がここに。\\n'}, Path('a.docx'))",
            "assert 'were not searched as part of a longer sentence' in text, text",
            "print('REFS OK')"])
        o = subprocess.run([PY, "-X", "utf8", "-c", code], capture_output=True, text=True, encoding="utf-8",
                           cwd=ROOT / "office-read" / "scripts")
        check("cross-check normalisation: furigana, notes in table cells, back links, curly quotes",
              "REFS OK" in o.stdout, o.stdout + o.stderr)
        code = "\n".join([
            "import office2md as m",
            "wrapped = 'あ' * 30 + '\\n' + 'い' * 30 + '\\n' + 'うえお。' + '\\n' + '短い行' + '\\n' + 'か' * 30 + '\\n' + '・項目'",
            "out = m.reflow(wrapped)",
            "assert out.splitlines()[0] == 'あ' * 30 + 'い' * 30 + 'うえお。', out",
            "assert '短い行' in out.splitlines() and '・項目' in out.splitlines(), out",
            "assert m.tabular('\\n'.join('行 %d 120 3,400 5.5 ▲7' % i for i in range(6)))",
            "assert not m.tabular('本文です。\\n' * 6)",
            "print('PDF OK')"])
        o = subprocess.run([PY, "-X", "utf8", "-c", code], capture_output=True, text=True, encoding="utf-8",
                           cwd=ROOT / "office-read" / "scripts")
        check("pdf: wrapped lines joined into their paragraph, short lines and list items kept, tables detected",
              "PDF OK" in o.stdout, o.stdout + o.stderr)
        if sys.platform == "win32" and shutil.which("powershell"):
            sys.path.insert(0, str(ROOT / "office-read" / "scripts"))
            import office_com
            legacy = d / "legacy.doc"
            office_com.run_script(
                "$a = New-Object -ComObject Word.Application; try { $a.DisplayAlerts = 0; "
                "$x = $a.Documents.Open($env:OFFICE_SRC, $false, $true); $x.SaveAs2($env:OFFICE_OUT, 0); $x.Close(0) } "
                "finally { $a.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($a) }",
                "WINWORD", {"OFFICE_SRC": str(d / "sample.docx"), "OFFICE_OUT": str(legacy)}, limit=120)
            if legacy.exists():
                o = run(READ, legacy)
                check("legacy .doc is converted by Word and read", "is the old .doc format; converted to .docx by "
                      "Microsoft Office" in o and "箇条書き その1" in o, o)
            else:
                print("  [skip] legacy .doc (Word not available)")
        o = run(READ, csv)
        check("csv cp932 decoded", "太郎" in o, o)
        o = run(READ, xlsx, "--json")
        check("json mode", '"type": "xlsx"' in o, o)
        o = run(READ, xlsx, "--profile")
        check("profile skeleton", "## 運用ルール" in o and "数式セル: D2:D4, D6" in o, o)

        print("office-edit")
        o = run(EDIT / "docx_replace.py", docx, "--replace", "概要=要約", "--replace", "社外秘=部外秘", "--no-backup")
        o2 = run(READ, docx)
        check("docx replace body + header", "要約" in o2 and "部外秘" in o2, o + o2)
        o = run(EDIT / "pptx_replace.py", pptx, "--replace", "発表者=スピーカー", "--no-backup")
        o2 = run(READ, pptx)
        check("pptx replace notes", "スピーカーノート" in o2, o + o2)
        before = d / "before.xlsx"
        shutil.copyfile(xlsx, before)
        o = run(EDIT / "xlsx_set.py", xlsx, "--set", "売上!A7=追加", "--set", "売上!B7=42", "--no-backup")
        o2 = run(READ, xlsx, "--sheet", "売上", "--no-coords")
        check("xlsx_set writes number and text", "| 追加 | 42 |" in o2, o + o2)
        o = run(ROOT / "office-read" / "scripts" / "office_diff.py", before, xlsx, "--sheet", "売上", expect_ok=False)
        check("office_diff reports added cells by address", "| A7 | added |  | 追加 |" in o and "| B7 | added |  | 42 |" in o, o)
        o = run(ROOT / "office-read" / "scripts" / "office_diff.py", docx, docx)
        check("office_diff identical docx", "No differences found." in o, o)

        have_excel = os.name == "nt" and subprocess.run(
            ["powershell", "-NoProfile", "-Command", "try{(New-Object -ComObject Excel.Application).Quit();'y'}catch{'n'}"],
            capture_output=True, text=True).stdout.strip() == "y"
        have_soffice = shutil.which("soffice") is not None or Path(r"C:\Program Files\LibreOffice\program\soffice.exe").exists()
        if have_excel or have_soffice:
            run(EDIT / "xlsx_set.py", xlsx, "--set", "売上!D7==1/0", "--no-backup")
            o = run(EDIT / "recalc.py", xlsx, expect_ok=False)
            check("recalc detects #DIV/0!", "#DIV/0!" in o, o)
            o = run(EDIT / "pdf_tool.py", "topdf", docx)
            pdf = docx.with_suffix(".pdf")
            check("topdf produced pdf", pdf.exists(), o)
            if pdf.exists():
                o = run(EDIT / "pdf_headings.py", pdf, "--out", d / "toc_auto.txt")
                check("pdf_headings finds numbered headings and writes toc.txt",
                      "heading(s)" in o and (d / "toc_auto.txt").exists()
                      and "1 1. 要約" in (d / "toc_auto.txt").read_text(encoding="utf-8"), o)
                o = run(EDIT / "pdf_headings.py", pdf, "--sample")
                check("pdf_headings --sample lists page | font | text", "| 報告書タイトル" in o, o)
                (d / "toc_short.txt").write_text("1 1.1 詳細\n", encoding="utf-8")
                o = run(EDIT / "pdf_headings.py", pdf, "--check", d / "toc_short.txt", expect_ok=False)
                check("pdf_headings --check reports unlisted heading-like lines", "not in the list" in o and "1. 要約" in o, o)
                merged = d / "m.pdf"
                run(EDIT / "pdf_tool.py", "merge", merged, pdf, pdf)
                toc = d / "toc.txt"
                toc.write_text("1 第1章\n  2 1.1 節\n", encoding="utf-8")
                run(EDIT / "pdf_tool.py", "bookmarks", merged, "--set", toc, "--out", d / "mb.pdf")
                o = run(EDIT / "pdf_tool.py", "info", d / "mb.pdf")
                check("pdf merge + bookmarks", "pages: 2" in o and "1.1 節" in o, o)
                run(EDIT / "pdf_tool.py", "split", d / "mb.pdf", "--every", "1", "--outdir", d / "parts")
                check("pdf split", len(list((d / "parts").glob("*.pdf"))) == 2)
        else:
            print("  [skip] recalc / topdf (no Excel COM or LibreOffice)")

    print()
    if fails:
        print(f"{len(fails)} failure(s):")
        for f in fails:
            print("-", f)
        return 1
    print("all checks passed")
    return 0


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except AttributeError:
        pass
    sys.exit(main())
