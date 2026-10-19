#!/usr/bin/env python3
"""Run declarative checks against a workbook: required cells, cross-sheet equality, date order.

Usage:
    xlsx_check.py book.xlsx --rules checks.txt
    xlsx_check.py book.xlsx --profile profiles/monthly-report.md      rules from its ```checks block
    xlsx_check.py book.xlsx --rule "equal メイン!C4 評価シート!C4" --rule "required メイン!B3"

Rule lines (one per line, '#' starts a comment, blank lines ignored):

    required  メイン!B3 メイン!B5 申請書!D3          every listed cell is non-empty
    equal     メイン!C4 評価シート!C4                all listed cells hold the same value
    compare   メイン!B10 <= メイン!B11               <, <=, ==, !=, >=, >  (dates, numbers, text)
    compare   明細!H2 > 0                            right side may be a literal
    regex     メイン!B7 ^[A-Z]{3}-\\d{3}$             text matches the pattern
    in        メイン!C9 承認,却下,保留               value is one of the choices
    count     明細!A3:A200 >= 1                      number of non-empty cells in the range
    no_errors                                       no #REF!, #DIV/0!, #NAME? ... anywhere
    recalculated                                    no formula without a cached value

A rule about a formula cell that was never recalculated cannot be judged; it is reported as
UNKNOWN rather than passed. Exit code 0 = all pass, 1 = any FAIL or UNKNOWN, 2 = bad rule.
"""

from __future__ import annotations

import argparse
import datetime as dt
import re
import sys
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")  # Windows console defaults to cp932
except AttributeError:
    pass

ERRORS = ("#REF!", "#DIV/0!", "#NAME?", "#VALUE!", "#N/A", "#NUM!", "#NULL!")
OPS = {"<": lambda a, b: a < b, "<=": lambda a, b: a <= b, "==": lambda a, b: a == b,
       "!=": lambda a, b: a != b, ">=": lambda a, b: a >= b, ">": lambda a, b: a > b}
REF_RE = re.compile(r"^(?:'([^']+)'|([^'!]+))!(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)$")


class Book:
    def __init__(self, path: Path):
        import openpyxl
        self.vals = openpyxl.load_workbook(path, data_only=True)
        self.forms = openpyxl.load_workbook(path, data_only=False)

    def cell(self, ref: str):
        """-> (sheet, addr, value, state) where state is 'ok' | 'uncached' | 'missing'."""
        m = REF_RE.match(ref.strip())
        if not m:
            raise ValueError(f"not a cell reference: {ref!r} (write Sheet!A1)")
        sheet = m[1] or m[2]
        addr = m[3].replace("$", "")
        if sheet not in self.vals.sheetnames:
            raise ValueError(f"sheet not found: {sheet!r}; sheets: {self.vals.sheetnames}")
        if ":" in addr:
            raise ValueError(f"single cell expected, got range {ref!r}")
        v = self.vals[sheet][addr].value
        f = self.forms[sheet][addr].value
        if isinstance(f, str) and f.startswith("=") and v is None:
            return sheet, addr, None, "uncached"
        return sheet, addr, v, "ok"

    def range_cells(self, ref: str):
        m = REF_RE.match(ref.strip())
        if not m or ":" not in m[3]:
            raise ValueError(f"range expected: {ref!r} (write Sheet!A1:B9)")
        sheet = m[1] or m[2]
        if sheet not in self.vals.sheetnames:
            raise ValueError(f"sheet not found: {sheet!r}")
        return sheet, list(self.vals[sheet][m[3].replace("$", "")])


def show(v) -> str:
    """Cell value for the report: one line, no table or markdown structure characters."""
    if v is None:
        return "(empty)"
    if isinstance(v, dt.datetime):
        return v.strftime("%Y-%m-%d") if v.time() == dt.time() else v.isoformat(sep=" ")
    s = str(v).replace("\r\n", " / ").replace("\r", " / ").replace("\n", " / ")
    s = s.replace("|", "\\|").replace("`", "ˋ")
    return "\\" + s if s[:1] in ("#", ">", "-", "*", "+") else s


def literal(text: str):
    """A right-hand literal: number, ISO date, or text."""
    t = text.strip()
    if re.fullmatch(r"-?\d+", t):
        return int(t)
    if re.fullmatch(r"-?\d+\.\d*", t):
        return float(t)
    for fmt in ("%Y-%m-%d", "%Y/%m/%d"):
        try:
            return dt.datetime.strptime(t, fmt)
        except ValueError:
            pass
    return t.strip('"')


def coerce(a, b):
    """Make two values comparable: numbers with numeric strings, dates with date strings."""
    if isinstance(a, dt.datetime) and isinstance(b, str):
        b = literal(b)
    if isinstance(b, dt.datetime) and isinstance(a, str):
        a = literal(a)
    if isinstance(a, (int, float)) and isinstance(b, str):
        try:
            b = float(b)
        except ValueError:
            pass
    if isinstance(b, (int, float)) and isinstance(a, str):
        try:
            a = float(a)
        except ValueError:
            pass
    if isinstance(a, dt.date) and not isinstance(a, dt.datetime):
        a = dt.datetime.combine(a, dt.time())
    if isinstance(b, dt.date) and not isinstance(b, dt.datetime):
        b = dt.datetime.combine(b, dt.time())
    return a, b


def run_rule(book: Book, line: str) -> tuple[str, str, str]:
    """-> (status, rule, detail) with status PASS | FAIL | UNKNOWN."""
    rule = line.split("#", 1)[0].strip()
    parts = rule.split()
    kind, args = parts[0], parts[1:]

    if kind == "required":
        if not args:
            raise ValueError("required needs at least one REF")
        bad, unk = [], []
        for ref in args:
            sh, addr, v, st = book.cell(ref)
            if st == "uncached":
                unk.append(f"{sh}!{addr}")
            elif v is None or (isinstance(v, str) and not v.strip()):
                bad.append(f"{sh}!{addr}")
        if bad:
            return "FAIL", rule, "empty: " + ", ".join(bad)
        if unk:
            return "UNKNOWN", rule, "uncached formula: " + ", ".join(unk)
        return "PASS", rule, f"{len(args)} cell(s) filled"

    if kind == "equal":
        if len(args) < 2:
            raise ValueError("equal needs at least two REFs")
        got = [book.cell(r) for r in args]
        if any(st == "uncached" for *_, st in got):
            return "UNKNOWN", rule, "uncached: " + ", ".join(f"{s}!{a}" for s, a, _, st in got if st == "uncached")
        vals = [coerce(v, got[0][2])[0] for *_, v, _ in got]
        detail = ", ".join(f"{s}!{a}={show(v)}" for s, a, v, _ in got)
        if all(v is None or (isinstance(v, str) and not v.strip()) for v in vals):
            return "FAIL", rule, detail + " (all empty; equality of nothing is not a match)"
        return ("PASS" if all(v == vals[0] for v in vals) else "FAIL"), rule, detail

    if kind == "compare":
        if len(args) != 3 or args[1] not in OPS:
            raise ValueError("compare needs: REF OP REF_OR_LITERAL")
        s1, a1, v1, st1 = book.cell(args[0])
        if REF_RE.match(args[2]):
            # looks like Sheet!A1: a bad sheet or a range is an error in the rule, not a literal
            s2, a2, v2, st2 = book.cell(args[2])
            rhs = f"{s2}!{a2}={show(v2)}"
        else:
            v2, st2, rhs = literal(args[2]), "ok", show(literal(args[2]))
        if st1 == "uncached" or st2 == "uncached":
            return "UNKNOWN", rule, "uncached formula on one side"
        if v1 is None or v2 is None:
            return "FAIL", rule, f"{s1}!{a1}={show(v1)} {args[1]} {rhs}: empty side"
        x, y = coerce(v1, v2)
        try:
            ok = OPS[args[1]](x, y)
        except TypeError:
            return "FAIL", rule, f"{s1}!{a1}={show(v1)} ({type(v1).__name__}) vs {rhs}: not comparable"
        return ("PASS" if ok else "FAIL"), rule, f"{s1}!{a1}={show(v1)} {args[1]} {rhs}"

    if kind == "regex":
        if len(args) < 2:
            raise ValueError("regex needs: REF PATTERN")
        sh, addr, v, st = book.cell(args[0])
        pat = rule.split(None, 2)[2]
        if st == "uncached":
            return "UNKNOWN", rule, "uncached formula"
        ok = v is not None and re.fullmatch(pat, str(v)) is not None
        return ("PASS" if ok else "FAIL"), rule, f"{sh}!{addr}={show(v)}"

    if kind == "in":
        if len(args) < 2:
            raise ValueError("in needs: REF A,B,C")
        sh, addr, v, st = book.cell(args[0])
        choices = [c.strip() for c in rule.split(None, 2)[2].split(",")]
        if st == "uncached":
            return "UNKNOWN", rule, "uncached formula"
        ok = show(v) in choices
        return ("PASS" if ok else "FAIL"), rule, f"{sh}!{addr}={show(v)}; allowed: {', '.join(choices)}"

    if kind == "count":
        if len(args) != 3 or args[1] not in OPS:
            raise ValueError("count needs: RANGE OP NUMBER")
        sh, cells = book.range_cells(args[0])
        n = sum(1 for row in cells for c in (row if isinstance(row, tuple) else (row,))
                if c.value is not None and str(c.value).strip() != "")
        ok = OPS[args[1]](n, int(args[2]))
        return ("PASS" if ok else "FAIL"), rule, f"{n} non-empty cell(s) in {args[0]}"

    if kind == "no_errors":
        hits = [f"{ws.title}!{c.coordinate}={c.value}" for ws in book.vals.worksheets
                for row in ws.iter_rows() for c in row if isinstance(c.value, str) and c.value in ERRORS]
        return ("FAIL" if hits else "PASS"), rule, (", ".join(hits[:10]) + (" …" if len(hits) > 10 else "")) or "no error values"

    if kind == "recalculated":
        unc = [f"{ws.title}!{c.coordinate}" for ws in book.forms.worksheets for row in ws.iter_rows() for c in row
               if isinstance(c.value, str) and c.value.startswith("=")
               and book.vals[ws.title][c.coordinate].value is None]
        return ("FAIL" if unc else "PASS"), rule, (f"{len(unc)} uncached: " + ", ".join(unc[:10])) if unc else "all formulas have cached values"

    raise ValueError(f"unknown rule kind {kind!r}")


def rules_from_profile(path: Path) -> list[str]:
    text = path.read_text(encoding="utf-8")
    m = re.search(r"```checks\s*\n(.*?)```", text, re.S)
    if not m:
        sys.exit(f"[xlsx_check] no ```checks block in {path}")
    return m[1].splitlines()


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("file", type=Path)
    ap.add_argument("--rules", type=Path, help="text file with one rule per line")
    ap.add_argument("--profile", type=Path, help="profile .md with a ```checks block")
    ap.add_argument("--rule", action="append", default=[], help="inline rule (repeatable)")
    a = ap.parse_args(argv)
    if not a.file.exists():
        sys.exit(f"[xlsx_check] file not found: {a.file}")
    lines: list[str] = list(a.rule)
    if a.rules:
        lines += a.rules.read_text(encoding="utf-8").splitlines()
    if a.profile:
        lines += rules_from_profile(a.profile)
    lines = [l for l in lines if l.split("#", 1)[0].strip()]
    if not lines:
        sys.exit("[xlsx_check] no rules; pass --rules, --profile or --rule")

    book = Book(a.file)
    results = []
    bad_rule = False
    for line in lines:
        try:
            results.append(run_rule(book, line))
        except ValueError as e:
            results.append(("ERROR", line.strip(), str(e)))
            bad_rule = True

    print(f"# {a.file.name}: {len(results)} check(s)")
    print()
    print("| result | rule | detail |")
    print("|---|---|---|")
    for st, rule, detail in results:
        print(f"| {st} | {rule.replace('|', '\\|')} | {detail.replace('|', '\\|')} |")
    counts = {k: sum(1 for st, _, _ in results if st == k) for k in ("PASS", "FAIL", "UNKNOWN", "ERROR")}
    print()
    print(", ".join(f"{k}: {v}" for k, v in counts.items() if v))
    if counts["UNKNOWN"]:
        print("> UNKNOWN = the cell holds a formula that was never recalculated; run office-edit/scripts/recalc.py and check again.")
    if bad_rule:
        return 2
    return 1 if counts["FAIL"] or counts["UNKNOWN"] else 0


if __name__ == "__main__":
    sys.exit(main())
