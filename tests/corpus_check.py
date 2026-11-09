#!/usr/bin/env python3
"""Run md_crosscheck on every Word and Excel file under a folder and compare with a previous run.

Usage:
    corpus_check.py CORPUS_DIR --save run.json                     record a run
    corpus_check.py CORPUS_DIR --save new.json --baseline run.json   and list what got worse

The documents stay where they are and are never copied into the repository; the conversions go
to --out (a temporary folder by default). For each file the review's checks are recorded (ok,
FAIL or not run) with the counts that matter: sentences missing against the references, printed
text missing, numbers that differ from Word's, workbook text missing. Against a baseline a file
is worse when a check went from ok to FAIL, one of those counts grew, or the run stopped
producing a review. The exit status is 1 when anything is worse, so the run can gate a commit.
"""

from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
import tempfile
import time
from pathlib import Path

CROSSCHECK = Path(__file__).resolve().parent.parent / "office-read" / "scripts" / "md_crosscheck.py"
KINDS = (".docx", ".docm", ".xlsx", ".xlsm")
# counts read off the review's check lines; a larger number is worse
COUNTS = {
    "missing": r"Content against [^:]*: (\d+) sentence",
    "fewer": r"(\d+) repeated fewer times",
    "printed": r"Printed text: (\d+) piece",
    "numbers": r"(\d+) with another number",
    "text missing": r"Text: \d+ text item\(s\), (\d+) missing",
    "links": r"Pictures and links: (\d+) broken",
}


def review_of(path: Path) -> dict:
    text = path.read_text(encoding="utf-8")
    checks = {}
    for m in re.finditer(r"^- \[(ok|FAIL|not run)\] ([^:]+):", text, re.M):
        checks[m[2].strip()] = m[1]
    counts = {k: int(m[1]) for k, p in COUNTS.items() if (m := re.search(p, text))}
    verdict = re.search(r"\*\*Status: ([A-Z ]+)\*\*|Status: (NO REFERENCE)", text)
    return {"status": (verdict[1] or verdict[2]) if verdict else "", "checks": checks, "counts": counts}


def run(corpus: Path, out: Path, render: bool, timeout: int) -> dict:
    files = sorted(p for p in corpus.rglob("*") if p.suffix.lower() in KINDS and not p.name.startswith("~$"))
    results = {}
    for i, f in enumerate(files, 1):
        rel = f.relative_to(corpus).as_posix()
        work = out / f"{i:04d}"
        cmd = [sys.executable, "-X", "utf8", str(CROSSCHECK), str(f), "--out-dir", str(work), "--fresh",
               "--no-markitdown"] + ([] if render else ["--no-render"])
        t = time.time()
        try:
            r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace",
                               timeout=timeout)
            reviews = list(work.glob("*_review/review.md"))
            entry = review_of(reviews[0]) if reviews else {"error": (r.stderr or r.stdout).strip()[-300:]}
        except subprocess.TimeoutExpired:
            entry = {"error": f"timed out after {timeout} s"}
        entry["secs"] = round(time.time() - t, 1)
        results[rel] = entry
        state = entry.get("status") or "ERROR"
        print(f"[{i}/{len(files)}] {entry['secs']:6.1f}s {state:13} {rel}", flush=True)
    return results


def compare(old: dict, new: dict) -> tuple[list[str], list[str]]:
    worse, better = [], []
    for f, n in new.items():
        o = old.get(f)
        if o is None:
            continue
        if "error" in n and "error" not in o:
            worse.append(f"{f}: no review any more ({n['error'][:120]})")
            continue
        for check, state in n.get("checks", {}).items():
            before = o.get("checks", {}).get(check)
            if before == "ok" and state == "FAIL":
                worse.append(f"{f}: {check} went from ok to FAIL")
            elif before == "FAIL" and state == "ok":
                better.append(f"{f}: {check} now ok")
        for k, v in n.get("counts", {}).items():
            was = o.get("counts", {}).get(k)
            if was is not None and v > was:
                worse.append(f"{f}: {k} {was} -> {v}")
            elif was is not None and v < was:
                better.append(f"{f}: {k} {was} -> {v}")
    return worse, better


def totals(results: dict) -> dict:
    out: dict[str, int] = {}
    for r in results.values():
        for k, v in r.get("counts", {}).items():
            out[k] = out.get(k, 0) + v
    out["errors"] = sum(1 for r in results.values() if "error" in r)
    return out


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("corpus", type=Path)
    ap.add_argument("--save", type=Path, required=True, help="where to write this run's results (JSON)")
    ap.add_argument("--baseline", type=Path, help="a previous run's JSON to compare with")
    ap.add_argument("--out", type=Path, help="where the conversions go (default: a temporary folder)")
    ap.add_argument("--no-render", action="store_true", help="skip Word/Excel printing (faster, fewer checks)")
    ap.add_argument("--timeout", type=int, default=900, help="seconds per file (default 900)")
    a = ap.parse_args(argv)
    if not a.corpus.is_dir():
        sys.exit(f"[corpus_check] not a folder: {a.corpus}")
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmp:
        results = run(a.corpus, a.out or Path(tmp), not a.no_render, a.timeout)
    a.save.write_text(json.dumps(results, ensure_ascii=False, indent=1), encoding="utf-8")
    now = totals(results)
    print(f"[corpus_check] {len(results)} file(s); totals {now}; saved {a.save}")
    if not a.baseline:
        return 0
    old = json.loads(a.baseline.read_text(encoding="utf-8"))
    worse, better = compare(old, results)
    print(f"[corpus_check] against {a.baseline}: totals {totals(old)} -> {now}")
    print(f"[corpus_check] better ({len(better)}):")
    for line in better[:60]:
        print("  " + line)
    print(f"[corpus_check] worse ({len(worse)}):")
    for line in worse:
        print("  " + line)
    return 1 if worse else 0


if __name__ == "__main__":
    sys.exit(main())
