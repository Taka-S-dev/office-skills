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
import shutil
import subprocess
import sys
import tempfile
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
        o = run(READ, xlsx, "--sheet", "申請書", "--form")
        check("xlsx --form pairs label right / merged / below / header row",
              all(x in o for x in ["- 申請日 = 2026/10/01  (申請書!B2 → C2)", "- 対象機種 = XYZ-100  (申請書!B3 → D3)",
                                   "- 氏名 = 山田  (申請書!B5 → B6)", "- 所属 = 品質保証部  (申請書!C5 → C6)",
                                   "- 備考 = 特になし  (申請書!B8 → B9)"]), o)
        o = run(READ, docx)
        check("docx header + table + list", "社外秘" in o and "| 項目 | 値 |" in o and "- 箇条書き その1" in o, o)
        o = run(READ, pptx)
        check("pptx notes + table", "発表者ノート" in o and "| x | y |" in o, o)
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
        o = run(EDIT / "xlsx_set.py", xlsx, "--set", "売上!A7=追加", "--set", "売上!B7=42", "--no-backup")
        o2 = run(READ, xlsx, "--sheet", "売上", "--no-coords")
        check("xlsx_set writes number and text", "| 追加 | 42 |" in o2, o + o2)

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
