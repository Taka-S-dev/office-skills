#!/usr/bin/env python3
"""Compare Microsoft Project plans: two versions, or a year of saved copies as a timeline.

Usage:
    mpp_history.py 計画_20261002.mpp 計画_20261106.mpp     what changed between two versions
    mpp_history.py ./計画の履歴/                            every .mpp / Project XML in the folder, as a timeline
    mpp_history.py a.mpp b.mpp c.mpp                        three or more files: a timeline
    mpp_history.py ./計画の履歴/ --cache ./mpp_cache        keep converted XML; a rerun converts only new files
    mpp_history.py old.mpp new.mpp --json

Each file is read with mpp_reader (Microsoft Project, else MPXJ; Project XML directly).
Snapshots are ordered by the plan's status date (状況報告日); a plan without one falls back to
a date in its file name (20261106, 2026-11-06, 2026_11_06), then to the file's modified time.
A status date far from the file-name date is reported: the status date was probably not
updated, and "late" for that snapshot is judged at the old date.

Tasks are matched by Project's unique id (UID), which survives inserting and reordering rows.
When most shared UIDs carry different names (the plan was rebuilt or pasted into a new file),
tasks are matched by their outline path of names instead, and the output says so.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import office2md  # noqa: E402
import mpp_reader as mr  # noqa: E402

try:
    sys.stdout.reconfigure(encoding="utf-8")  # Windows console defaults to cp932
except AttributeError:
    pass

PLAN_EXT = (".mpp", ".mpt", ".xml")
STALE_DAYS = 7           # status date this far from the file-name date: probably not updated
STALL_POINTS = 1.0       # less overall progress than this between snapshots counts as stalled
esc = office2md.md_escape


# -- loading snapshots ------------------------------------------------------------------------

def name_date(path: Path) -> datetime | None:
    for m in re.finditer(r"(20[0-9]{2})[-_.]?([01][0-9])[-_.]?([0-3][0-9])(?![0-9])", path.stem):
        try:
            return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            continue
    return None


def load(path: Path, engine: str, cache: Path | None) -> dict:
    """Plan data plus when it describes ('date', 'date_from') and any warning about that date."""
    xml = None
    key = ""
    if cache:
        key = hashlib.sha256(path.read_bytes()).hexdigest()[:24]
        hit = cache / f"{key}.xml"
        if hit.exists():
            xml, label = hit.read_bytes(), "cached conversion"
    if xml is None:
        xml, label = mr.to_mspdi(path, engine)
        if cache:
            cache.mkdir(parents=True, exist_ok=True)
            (cache / f"{key}.xml").write_bytes(xml)
    data = mr.parse(xml)
    data.update({"engine": label, "file": str(path), "name": path.name})
    status = mr._date(data["project"].get("StatusDate"))
    fdate = name_date(path)
    if status and fdate and abs((status.date() - fdate.date()).days) > STALE_DAYS:
        # placed by its file name, but lateness stays judged at the status date the plan states
        data["date"], data["date_from"] = fdate, "file name (status date not updated)"
        data["date_warning"] = (f"status date {status.date()} is {abs((status.date() - fdate.date()).days)} days "
                                f"from the date in the file name ({fdate.date()}); the status date was probably not "
                                "updated. The snapshot is placed by its file name; lateness in it is judged at the "
                                "old status date")
    elif status:
        data["date"], data["date_from"] = status, "status date"
    else:
        data["date"], data["date_from"] = ((fdate, "file name") if fdate else
                                           (datetime.fromtimestamp(path.stat().st_mtime), "file modified time"))
        # no status date: lateness is judged at the snapshot's date instead of not at all
        data["project"]["StatusDate"] = data["date"].isoformat()
    return data


def plan_files(targets: list[Path]) -> list[Path]:
    files: list[Path] = []
    for t in targets:
        if t.is_dir():
            files += sorted(p for p in t.iterdir() if p.is_file() and p.suffix.lower() in PLAN_EXT
                            and not p.name.startswith("~$") and (p.suffix.lower() != ".xml" or mr.is_mspdi(p)))
        elif t.is_file():
            files.append(t)
        else:
            sys.exit(f"[mpp_history] not found: {t}")
    return files


# -- matching tasks across versions -----------------------------------------------------------

def with_paths(data: dict) -> None:
    """Give each task its parent and outline path of names (for matching and for display)."""
    stack: list[dict] = []
    for t in data["tasks"]:
        while stack and stack[-1]["level"] >= t["level"]:
            stack.pop()
        t["parent"] = stack[-1]["name"] if stack else ""
        t["path"] = " / ".join([s["name"] for s in stack] + [t["name"]])
        stack.append(t)


def match_mode(a: dict, b: dict) -> str:
    ua = {t["uid"]: t["name"] for t in a["tasks"] if t["uid"]}
    ub = {t["uid"]: t["name"] for t in b["tasks"] if t["uid"]}
    common = set(ua) & set(ub)
    if not common:
        return "path"
    same = sum(ua[u] == ub[u] for u in common)
    return "uid" if same / len(common) >= 0.6 else "path"


def keyed(data: dict, mode: str) -> dict[str, dict]:
    out = {}
    for t in data["tasks"]:
        k = t["uid"] if mode == "uid" else t["path"]
        if k and k not in out:
            out[k] = t
    return out


# -- what one plan says -----------------------------------------------------------------------

def leaves(data: dict) -> list[dict]:
    return [t for t in data["tasks"] if not t["summary"] and t["active"]]


def overall_percent(data: dict) -> float:
    """Project's own % complete when the plan carries it, else leaf tasks weighted by duration."""
    if data["project"].get("percent"):
        return float(data["project"]["percent"])
    ls = leaves(data)
    total = sum(t["duration_min"] for t in ls)
    if total > 0:
        return sum(t["duration_min"] * t["percent"] for t in ls) / total
    return sum(t["percent"] for t in ls) / len(ls) if ls else 0.0


def plan_finish(data: dict) -> datetime | None:
    d = mr._date(data["project"].get("FinishDate"))
    if d:
        return d
    fins = [mr._date(t["finish"]) for t in data["tasks"]]
    fins = [f for f in fins if f]
    return max(fins) if fins else None


def metrics(data: dict) -> dict:
    att = mr.attention(data)
    pc = mr.progress_check(data)
    fin = plan_finish(data)
    return {"expected": pc["expected"], "behind_expected": len(pc["behind"]), "spi": pc["spi"],"date": data["date"].date().isoformat(), "date_from": data["date_from"], "file": data["name"],
            "tasks": len(leaves(data)), "percent": round(overall_percent(data), 1), "late": len(att["late"]),
            "behind_baseline": len(att["slipped"]), "critical_open": len(att["critical_open"]),
            "finish": fin.date().isoformat() if fin else ""}


# -- the difference between two plans ---------------------------------------------------------

def _days_between(a: str, b: str) -> int | None:
    da, db = mr._date(a), mr._date(b)
    return (db.date() - da.date()).days if da and db else None


def diff(a: dict, b: dict) -> dict:
    with_paths(a)
    with_paths(b)
    mode = match_mode(a, b)
    ka, kb = keyed(a, mode), keyed(b, mode)
    a_by_uid = {t["uid"]: t for t in a["tasks"]}
    b_by_uid = {t["uid"]: t for t in b["tasks"]}

    def pred_set(t, by_uid):
        # by uid when the ids carry over (a renamed predecessor is the same link), else by name path
        return sorted(f"{p['uid'] if mode == 'uid' else by_uid.get(p['uid'], {}).get('path', p['uid'])}"
                      f"|{p['type']}|{p['lag_min']:g}" for p in t["predecessors"])
    out = {"mode": mode, "added": [kb[k] for k in kb if k not in ka], "removed": [ka[k] for k in ka if k not in kb],
           "renamed": [], "moved": [], "dates": [], "progress": [], "links": [], "resources": [], "rebaselined": [],
           "other": []}
    for k in kb:
        if k not in ka:
            continue
        old, new = ka[k], kb[k]
        if mode == "uid" and old["name"] != new["name"]:
            out["renamed"].append((old, new))
        if old["parent"] != new["parent"]:
            out["moved"].append((old, new))
        ds, df = _days_between(old["start"], new["start"]), _days_between(old["finish"], new["finish"])
        if (ds or df) or old["duration"] != new["duration"]:
            out["dates"].append((old, new, ds or 0, df or 0))
        if old["percent"] != new["percent"]:
            out["progress"].append((old, new))
        if pred_set(old, a_by_uid) != pred_set(new, b_by_uid):
            out["links"].append((old, new))
        if sorted(old["resources"]) != sorted(new["resources"]):
            out["resources"].append((old, new))
        if old["baseline_finish"] != new["baseline_finish"] and old["baseline_finish"] and new["baseline_finish"]:
            out["rebaselined"].append((old, new))
        for field, label in (("deadline", "deadline"), ("notes", "notes"), ("custom", "custom fields"),
                             ("active", "active"), ("milestone", "milestone"), ("constraint", "constraint")):
            if old[field] != new[field]:
                out["other"].append((old, new, label))
    return out


def _d(s: str) -> str:
    return s[:10] if s else "-"


def _task(t: dict) -> str:
    return f"{t['id']} {esc(t['name'])}"


def _signed(n: int) -> str:
    return f"+{n}" if n > 0 else str(n)


def diff_markdown(a: dict, b: dict, d: dict, heading: str = "##") -> list[str]:
    ma, mb = metrics(a), metrics(b)
    out = []
    fa, fb = plan_finish(a), plan_finish(b)
    move = (fb.date() - fa.date()).days if fa and fb else None
    out.append(f"Plan finish: {ma['finish'] or '-'} → {mb['finish'] or '-'}"
               + (f" ({_signed(move)} days)" if move else "") + f" · progress {ma['percent']}% → {mb['percent']}% · "
               f"late {ma['late']} → {mb['late']} · tasks {ma['tasks']} → {mb['tasks']}")
    if d["mode"] == "path":
        out.append("> Tasks matched by their outline path of names: the unique ids differ between the files "
                   "(the plan was rebuilt or pasted). A renamed task then shows as removed plus added.")
    out.append("")

    def section(title: str, rows: list[str]):
        if rows:
            out.extend([f"{heading} {title} ({len(rows)})", ""] + rows[:200]
                       + ([f"- … {len(rows) - 200} more"] if len(rows) > 200 else []) + [""])
    section("Added tasks", [f"- {_task(t)} under '{esc(t['parent']) or '(top)'}' — {_d(t['start'])} → {_d(t['finish'])}"
                            for t in d["added"]])
    section("Removed tasks", [f"- {_task(t)} under '{esc(t['parent']) or '(top)'}' — was {_d(t['start'])} → {_d(t['finish'])}, "
                              f"{t['percent']}% done" for t in d["removed"]])
    section("Renamed", [f"- {o['id']} '{esc(o['name'])}' → '{esc(n['name'])}'" for o, n in d["renamed"]])
    section("Moved to another summary task", [f"- {_task(n)}: '{esc(o['parent']) or '(top)'}' → '{esc(n['parent']) or '(top)'}'"
                                              for o, n in d["moved"]])
    dates = sorted(d["dates"], key=lambda x: -x[3])
    section("Dates changed (latest finish moves first)",
            [f"- {_task(n)}: finish {_d(o['finish'])} → {_d(n['finish'])} ({_signed(df)} days)"
             + (f", start {_signed(ds)} days" if ds else "") + (f", duration {o['duration']} → {n['duration']} days"
                                                               if o["duration"] != n["duration"] else "")
             for o, n, ds, df in dates])
    section("Progress", [f"- {_task(n)}: {o['percent']}% → {n['percent']}%" for o, n in
                         sorted(d["progress"], key=lambda x: x[1]["percent"] - x[0]["percent"])])
    section("Dependencies changed", [f"- {_task(n)}: '{o['predecessor_text'] or '-'}' → '{n['predecessor_text'] or '-'}'"
                                     for o, n in d["links"]])
    section("Resources changed", [f"- {_task(n)}: {esc(', '.join(o['resources']) or '-')} → {esc(', '.join(n['resources']) or '-')}"
                                  for o, n in d["resources"]])
    section("Baseline reset (the plan's reference dates were changed)",
            [f"- {_task(n)}: baseline finish {_d(o['baseline_finish'])} → {_d(n['baseline_finish'])}" for o, n in d["rebaselined"]])
    section("Other changes", [f"- {_task(n)}: {label} changed" for o, n, label in d["other"]])
    if not any(d[k] for k in ("added", "removed", "renamed", "moved", "dates", "progress", "links", "resources",
                              "rebaselined", "other")):
        out += ["No task differs.", ""]
    return out


# -- a timeline of many plans -----------------------------------------------------------------

def timeline(snaps: list[dict]) -> dict:
    rows, periods = [], []
    history: dict[str, list[tuple[str, str]]] = {}      # task key -> [(snapshot date, finish)]
    names: dict[str, str] = {}
    custom_of: dict[str, dict] = {}
    first_seen: dict[str, str] = {}
    milestones: dict[str, list[tuple[str, str]]] = {}
    for i, s in enumerate(snaps):
        with_paths(s)
        m = metrics(s)
        rows.append(m)
        mode = match_mode(snaps[i - 1], s) if i else "uid"
        for t in s["tasks"]:
            k = t["uid"] if mode == "uid" and t["uid"] else t["path"]
            names[k] = t["name"]
            custom_of[k] = t["custom"]
            first_seen.setdefault(k, m["date"])
            if not t["summary"]:
                history.setdefault(k, []).append((m["date"], t["finish"]))
            if t["milestone"]:
                milestones.setdefault(k, []).append((m["date"], t["finish"]))
        if i:
            prev = snaps[i - 1]
            d = diff(prev, s)
            pm = rows[i - 1]
            reasons = []
            if m["late"] > pm["late"]:
                reasons.append(f"late tasks {pm['late']} → {m['late']}")
            if m["behind_expected"] > pm["behind_expected"]:
                reasons.append(f"tasks behind their expected progress {pm['behind_expected']} → {m['behind_expected']}")
            fa, fb = plan_finish(prev), plan_finish(s)
            if fa and fb and fb > fa:
                reasons.append(f"plan finish moved {(fb.date() - fa.date()).days} days later")
            gain = m["percent"] - pm["percent"]
            if gain < STALL_POINTS and pm["percent"] < 100:
                reasons.append(f"progress {pm['percent']}% → {m['percent']}% (stalled)")
            later = [x for x in d["dates"] if x[3] > 0]
            if later:
                reasons.append(f"{len(later)} task(s) finish later")
            if d["added"]:
                reasons.append(f"{len(d['added'])} task(s) added")
            if d["removed"]:
                reasons.append(f"{len(d['removed'])} task(s) removed")
            periods.append({"from": pm["date"], "to": m["date"], "added": d["added"], "removed": d["removed"],
                            "later": later, "rebaselined": d["rebaselined"], "reasons": reasons, "mode": d["mode"]})
    slips = []
    for k, seq in history.items():
        fins = [(day, mr._date(f)) for day, f in seq if mr._date(f)]
        moves = [(fins[j][0], (fins[j][1].date() - fins[j - 1][1].date()).days)
                 for j in range(1, len(fins)) if fins[j][1].date() > fins[j - 1][1].date()]
        if moves:
            total = (fins[-1][1].date() - fins[0][1].date()).days
            slips.append({"name": names[k], "custom": custom_of.get(k, {}), "moves": len(moves), "total_days": total, "first": fins[0][1].date().isoformat(),
                          "last": fins[-1][1].date().isoformat(), "when": [w for w, _ in moves]})
    slips.sort(key=lambda x: (-x["moves"], -x["total_days"]))
    added_later = [(day, names[k]) for k, day in first_seen.items() if day != rows[0]["date"]]
    return {"rows": rows, "periods": periods, "slips": slips, "added_later": sorted(added_later),
            "milestones": {names[k]: v for k, v in milestones.items()}}


def _chart(title: str, labels: list[str], values: list[float], unit: str, second: list[float] | None = None) -> list[str]:
    def safe(s: str) -> str:
        return re.sub(r'["`\n]', " ", s)
    hi = max(values + (second or []) + [1])
    lines = ["```mermaid", "xychart-beta", f'    title "{safe(title)}"',
             "    x-axis [" + ", ".join(f'"{safe(x[5:])}"' for x in labels) + "]",
             f'    y-axis "{safe(unit)}" 0 --> {int(hi * 1.2) + 1}',
             "    line [" + ", ".join(f"{v:g}" for v in values) + "]"]
    if second:
        lines.append("    line [" + ", ".join(f"{v:g}" for v in second) + "]")
    return lines + ["```"]


def timeline_markdown(snaps: list[dict], tl: dict) -> list[str]:
    rows = tl["rows"]
    out = [f"{len(snaps)} snapshots, {rows[0]['date']} → {rows[-1]['date']}. Each is one saved copy of the plan; "
           "dates are status dates unless noted.", ""]
    warn = [f"- {esc(s['name'])}: {s['date_warning']}" for s in snaps if s.get("date_warning")]
    fallback = [f"- {esc(s['name'])}: no status date; placed and judged for lateness at its {s['date_from']} "
                f"({s['date'].date()})" for s in snaps if s["date_from"] in ("file name", "file modified time")]
    if warn or fallback:
        out += ["## Dates to check", ""] + warn + fallback + [""]
    out += ["## Snapshots", "",
            office2md.md_table([[r["date"], esc(r["file"]), r["tasks"], f"{r['percent']}%",
                                 "" if r["expected"] is None else f"{r['expected']}%", r["behind_expected"], r["late"],
                                 r["behind_baseline"], r["critical_open"], "" if r["spi"] is None else r["spi"], r["finish"]]
                                for r in rows],
                               header=["date", "file", "tasks", "done", "expected", "behind expected", "late",
                                       "behind baseline", "critical open", "SPI", "plan finish"]), ""]
    bad = [p for p in tl["periods"] if p["reasons"]]
    out += [f"## Periods that got worse ({len(bad)} of {len(tl['periods'])})", "",
            "A period is the time between two snapshots. Listed when late tasks grew, the plan finish moved later, "
            f"progress grew less than {STALL_POINTS:g} point(s), tasks finish later, or tasks were added or removed.", ""]
    for p in bad:
        out.append(f"- {p['from']} → {p['to']}: " + "; ".join(p["reasons"]))
        for x in sorted(p["later"], key=lambda x: -x[3])[:5]:
            out.append(f"  - {_task(x[1])} finish {_d(x[0]['finish'])} → {_d(x[1]['finish'])} ({_signed(x[3])} days)")
        for t in p["added"][:5]:
            out.append(f"  - added: {_task(t)} under '{esc(t['parent']) or '(top)'}'")
        for t in p["removed"][:5]:
            out.append(f"  - removed: {_task(t)} ({t['percent']}% done when removed)")
        if p["rebaselined"]:
            out.append(f"  - baseline reset on {len(p['rebaselined'])} task(s): comparisons with the baseline restart here")
        if p["mode"] == "path":
            out.append("  - tasks matched by name path for this period (unique ids changed)")
    if not bad:
        out.append("None.")
    out += ["", f"## Tasks added after the first snapshot ({len(tl['added_later'])})", ""]
    by_day: dict[str, list[str]] = {}
    for day, name in tl["added_later"]:
        by_day.setdefault(day, []).append(name)
    out += [f"- {day}: " + ", ".join(esc(n) for n in ns[:15]) + (f" … {len(ns) - 15} more" if len(ns) > 15 else "")
            for day, ns in by_day.items()] or ["None."]
    out += ["", f"## Tasks whose finish moved later most often ({len(tl['slips'])})", ""]
    out += [f"- {esc(s['name'])}: {s['moves']} time(s), {s['first']} → {s['last']} ({_signed(s['total_days'])} days); "
            f"moved in snapshots {', '.join(s['when'][:8])}" for s in tl["slips"][:30]] or ["None."]
    groups: dict[str, dict[str, list]] = {}
    for sl in tl["slips"]:
        for field, value in sl["custom"].items():
            groups.setdefault(field, {}).setdefault(value, []).append(sl)
    for field, by_value in groups.items():
        if len(by_value) > 20:
            continue                          # a free-text field, not a grouping
        out += ["", f"## Slips by {esc(field)}", "",
                office2md.md_table([[esc(v), len(sls), sum(x["moves"] for x in sls), sum(x["total_days"] for x in sls)]
                                    for v, sls in sorted(by_value.items(), key=lambda kv: -sum(x["total_days"] for x in kv[1]))],
                                   header=[esc(field), "tasks that slipped", "slips", "days slipped in total"])]
    if tl["milestones"]:
        out += ["", "## Milestones over time", "",
                "Each milestone's date as each snapshot saw it; a line that keeps moving right is a milestone that "
                "keeps slipping.", ""]
        for name, seq in tl["milestones"].items():
            changes = [seq[0]] + [seq[j] for j in range(1, len(seq)) if _d(seq[j][1]) != _d(seq[j - 1][1])]
            out.append(f"- {esc(name)}: " + " → ".join(f"{_d(f)} (as of {day})" for day, f in changes))
    labels = [r["date"] for r in rows]
    if len(rows) >= 3:
        first = mr._date(rows[0]["finish"] or "")
        drift = [((mr._date(r["finish"]).date() - first.date()).days if first and mr._date(r["finish"]) else 0)
                 for r in rows]
        out += ["", "## Trends", ""]
        out += _chart("Late tasks", labels, [r["late"] for r in rows], "tasks") + [""]
        out += _chart("Plan finish, days later than the first snapshot", labels, drift, "days") + [""]
        out += _chart("Overall progress (done, then expected by the dates)", labels, [r["percent"] for r in rows], "%",
                      [r["expected"] or 0 for r in rows])
    return out


# -- entry point ------------------------------------------------------------------------------

def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("targets", nargs="+", type=Path, help="two plans to compare, or a folder / three or more plans")
    ap.add_argument("--timeline", action="store_true", help="treat two files as a timeline instead of a diff")
    ap.add_argument("--engine", choices=["auto", "project", "mpxj"], default="auto")
    ap.add_argument("--cache", type=Path, help="keep converted XML here (keyed by file content) for later runs")
    ap.add_argument("--json", action="store_true")
    ap.add_argument("--out", type=Path)
    a = ap.parse_args(argv)
    files = plan_files(a.targets)
    if len(files) < 2:
        sys.exit("[mpp_history] need at least two plans")
    snaps, skipped = [], []
    for f in files:
        try:
            snaps.append(load(f, a.engine, a.cache))
        except mr.MppError as e:
            skipped.append(f"{f.name}: {e}")
        print(f"[mpp_history] read {len(snaps)}/{len(files)}", file=sys.stderr)
    if len(snaps) < 2:
        sys.exit("[mpp_history] fewer than two plans could be read: " + "; ".join(skipped))
    snaps.sort(key=lambda s: (s["date"], s["name"]))
    as_diff = len(snaps) == 2 and not a.timeline and not any(t.is_dir() for t in a.targets)
    if as_diff:
        result = diff(snaps[0], snaps[1])
        if a.json:
            text = json.dumps({"old": metrics(snaps[0]), "new": metrics(snaps[1]), **result}, ensure_ascii=False,
                              indent=2, default=str)
        else:
            old, new = snaps
            out = [f"# {esc(old['name'])} → {esc(new['name'])}", "", office2md.DATA_NOTICE, "",
                   f"Older: {old['date'].date()} ({old['date_from']}), read with {old['engine']}. "
                   f"Newer: {new['date'].date()} ({new['date_from']}), read with {new['engine']}.", ""]
            out += [f"> {esc(s['name'])}: {s['date_warning']}" for s in snaps if s.get("date_warning")]
            out += diff_markdown(old, new, result)
            text = "\n".join(out) + "\n"
    else:
        tl = timeline(snaps)
        if a.json:
            text = json.dumps(tl, ensure_ascii=False, indent=2, default=str)
        else:
            out = [f"# Plan history: {esc(snaps[0]['name'])} … {esc(snaps[-1]['name'])}", "", office2md.DATA_NOTICE, ""]
            engines = sorted({s["engine"] for s in snaps})
            out += [f"Read with: {', '.join(engines)}.", ""]
            out += timeline_markdown(snaps, tl)
            if skipped:
                out += ["", "## Not read", ""] + [f"- {esc(x)}" for x in skipped]
            text = "\n".join(out) + "\n"
    if a.out:
        a.out.write_text(text, encoding="utf-8")
        print(f"[mpp_history] wrote {a.out} ({len(text)} chars)", file=sys.stderr)
    else:
        print(text, end="")
    return 0


if __name__ == "__main__":
    sys.exit(main())
