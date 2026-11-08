#!/usr/bin/env python3
"""Read a Microsoft Project plan (.mpp / .mpt, or Project XML) as Markdown or JSON.

Usage:
    mpp_reader.py 計画.mpp                        tasks, dependencies, what is late
    mpp_reader.py 計画.mpp --gantt                also a mermaid gantt chart
    mpp_reader.py 計画.mpp --engine mpxj          choose the converter (auto | project | mpxj)
    mpp_reader.py 計画.xml                        Project XML (Save As > XML): nothing else needed
    mpp_reader.py 計画.mpp --save-xml 計画.xml    keep the converted XML for a later run
    office2md.py 計画.mpp                         the same, through the common entry point
                                              (office2md hands the file to this script)

.mpp is an undocumented binary format, so a plan is first converted to Project's own XML
(MSPDI) and one reader handles everything after that:
  1. Microsoft Project, when installed (driven through COM): Project writes the XML itself,
     the most faithful conversion.
  2. MPXJ (pip install mpxj; needs Java): reads .mpp without Project. Close to Project's own
     export; the output says when this engine was used.
  3. Neither available: the error says what to install, or to save the plan as XML in Project.
The engine used is named in the output, so a reader knows how far to trust it.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import office2md  # noqa: E402

try:
    sys.stdout.reconfigure(encoding="utf-8")  # Windows console defaults to cp932
except AttributeError:
    pass

NS = "http://schemas.microsoft.com/project"
LINK_TYPES = {"0": "FF", "1": "FS", "2": "SF", "3": "SS"}
CONSTRAINTS = {"0": "", "1": "as late as possible", "2": "must start on", "3": "must finish on",
               "4": "start no earlier than", "5": "start no later than", "6": "finish no earlier than",
               "7": "finish no later than"}
PROJECT_LIMIT_S = 240
MAX_TASKS = 5000


class MppError(office2md.ReadError):
    """Raised instead of exiting, so a folder search can skip one unreadable plan and go on."""


# -- converting a .mpp to Project XML ---------------------------------------------------------

def is_mspdi(path: Path) -> bool:
    try:
        with open(path, "rb") as f:
            head = f.read(4096)
    except OSError:
        return False
    return b"schemas.microsoft.com/project" in head and b"<Project" in head


def project_available() -> bool:
    if sys.platform != "win32":
        return False
    try:
        import winreg
        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"MSProject.Application"))
        return True
    except OSError:
        return False


def mpxj_available() -> bool:
    import importlib.util
    has_java = bool(shutil.which("java") or os.environ.get("JAVA_HOME"))
    return has_java and all(importlib.util.find_spec(m) is not None for m in ("mpxj", "jpype"))


_PROJECT_PS = r"""
$ErrorActionPreference = 'Stop'
$app = New-Object -ComObject MSProject.Application
try {
  $app.Visible = $false
  try { $app.DisplayAlerts = $false } catch { }
  # macros in the plan must never run: refuse to open unless they are disabled for this session
  $app.AutomationSecurity = 3
  if ($app.AutomationSecurity -ne 3) { throw 'could not disable macros in Project; not opening the plan' }
  # named arguments through IDispatch: PowerShell cannot skip positional COM arguments
  $bf = [Reflection.BindingFlags]::InvokeMethod
  [void]$app.GetType().InvokeMember('FileOpenEx', $bf, $null, $app, @($env:OFFICE_SRC, $true), $null, $null, @('Name', 'ReadOnly'))
  [void]$app.GetType().InvokeMember('FileSaveAs', $bf, $null, $app, @($env:OFFICE_OUT, 'MSProject.XML'), $null, $null, @('Name', 'FormatID'))
  [void]$app.GetType().InvokeMember('FileCloseEx', $bf, $null, $app, @(0), $null, $null, @('Save'))
  Write-Output 'OK'
} finally {
  try { $app.Quit(0) } catch { }
  [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
"""

_GUARD_PS = """
$t0 = Get-Date
$job = Start-Job -ScriptBlock ([scriptblock]::Create($env:OFFICE_SCRIPT))
if (Wait-Job $job -Timeout LIMIT) { Receive-Job $job 2>&1 }
else {
  Stop-Job $job
  # only hidden instances started by this call: a Project window the user opened is left alone
  Get-Process WINPROJ -ErrorAction SilentlyContinue |
    Where-Object { $_.StartTime -ge $t0 -and $_.MainWindowHandle -eq 0 } | Stop-Process -Force
  Write-Output 'TIMEOUT'
}
Remove-Job $job -Force
"""


def convert_with_project(src: Path, out_xml: Path) -> str:
    """'' on success, else the reason. Paths and the script travel in environment variables."""
    try:
        r = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command",
                            _GUARD_PS.replace("LIMIT", str(PROJECT_LIMIT_S))],
                           capture_output=True, text=True, timeout=PROJECT_LIMIT_S + 60,
                           env={**os.environ, "OFFICE_SCRIPT": _PROJECT_PS,
                                "OFFICE_SRC": str(src.resolve()), "OFFICE_OUT": str(out_xml.resolve())})
    except (subprocess.TimeoutExpired, OSError) as e:
        return f"Project did not run ({type(e).__name__})"
    if "TIMEOUT" in r.stdout:
        return f"Project did not answer within {PROJECT_LIMIT_S} s (a hidden dialog?); it was stopped"
    if out_xml.exists() and out_xml.stat().st_size > 0:
        return ""
    last = [ln for ln in (r.stdout + r.stderr).splitlines() if ln.strip()][-1:] or ["no output"]
    return "Project could not export the plan: " + last[0][:300]


def convert_with_mpxj(src: Path, out_xml: Path) -> str:
    """MPXJ runs in a child process: the JVM cannot be restarted, and a crash stays contained."""
    try:
        r = subprocess.run([sys.executable, str(Path(__file__).resolve()), "--mpxj-convert",
                            str(src.resolve()), str(out_xml.resolve())],
                           capture_output=True, text=True, timeout=300)
    except (subprocess.TimeoutExpired, OSError) as e:
        return f"MPXJ did not finish ({type(e).__name__})"
    if r.returncode == 0 and out_xml.exists() and out_xml.stat().st_size > 0:
        return ""
    last = [ln for ln in (r.stdout + r.stderr).splitlines() if ln.strip()][-1:] or ["no output"]
    return "MPXJ could not read the plan: " + last[0][:300]


def _mpxj_child(src: str, out: str) -> int:
    import jpype
    import jpype.imports  # noqa: F401  (enables "from org.mpxj ... import")
    __import__("mpxj")                         # puts the MPXJ jars on the class path
    jpype.startJVM()
    try:
        try:
            from org.mpxj.reader import UniversalProjectReader
            from org.mpxj.mspdi import MSPDIWriter
        except ImportError:                       # MPXJ before version 13
            from net.sf.mpxj.reader import UniversalProjectReader
            from net.sf.mpxj.mspdi import MSPDIWriter
        project = UniversalProjectReader().read(src)
        if project is None:
            print("not a project file MPXJ recognises")
            return 2
        MSPDIWriter().write(project, out)
        return 0
    finally:
        jpype.shutdownJVM()


def to_mspdi(src: Path, engine: str = "auto", save_xml: Path | None = None) -> tuple[bytes, str]:
    """(Project XML bytes, engine label). Raises SystemExit with what to do when nothing works."""
    if src.suffix.lower() == ".xml" or is_mspdi(src):
        return src.read_bytes(), "Project XML file"
    order = {"auto": ["project", "mpxj"], "project": ["project"], "mpxj": ["mpxj"]}[engine]
    reasons = []
    tmp = Path(tempfile.mkdtemp(prefix="mpp_"))
    try:
        out = tmp / "plan.xml"
        for name in order:
            if name == "project":
                if not project_available():
                    reasons.append("Microsoft Project is not installed")
                    continue
                why = convert_with_project(src, out)
                label = "Microsoft Project (its own XML export)"
            else:
                if not mpxj_available():
                    reasons.append("MPXJ is not installed (pip install mpxj; needs Java)")
                    continue
                why = convert_with_mpxj(src, out)
                label = "MPXJ (reads .mpp without Project; close to Project's own export)"
            if not why:
                data = out.read_bytes()
                if save_xml:
                    save_xml.write_bytes(data)
                return data, label
            reasons.append(why)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    raise MppError("cannot read " + src.name + ": " + "; ".join(reasons) + ". Install one of them, or open "
                   "the plan in Project and Save As > XML, then read the .xml.")


# -- reading Project XML ----------------------------------------------------------------------

def _q(name: str) -> str:
    return "{%s}%s" % (NS, name)


def _t(el, name: str) -> str:
    v = el.findtext(_q(name)) if el is not None else None
    return (v or "").strip()


def _date(s: str) -> datetime | None:
    try:
        return datetime.fromisoformat(s[:19]) if s else None
    except ValueError:
        return None


def _minutes(iso: str) -> float | None:
    """ISO 8601 duration as Project writes it (PT16H0M0S, P2DT4H...) in minutes."""
    m = re.fullmatch(r"(-)?P(?:([0-9.]+)D)?(?:T(?:([0-9.]+)H)?(?:([0-9.]+)M)?(?:([0-9.]+)S)?)?", iso or "")
    if not m or iso in ("P", "PT"):
        return None
    d, h, mi, s = (float(x) if x else 0.0 for x in m.groups()[1:])
    total = d * 24 * 60 + h * 60 + mi + s / 60
    return -total if m.group(1) else total


def _days(minutes: float | None, per_day: float) -> str:
    if minutes is None:
        return ""
    v = minutes / per_day
    return f"{v:.0f}" if abs(v - round(v)) < 0.05 else f"{v:.1f}"


def _clean(s: str, limit: int = 300, one_line: bool = True) -> str:
    """Text from the plan with control characters removed; single-line fields (names, aliases)
    also lose line breaks, so a crafted name cannot start a heading or a diagram line."""
    s = re.sub(r"[\x00-\x08\x0b-\x1f\x7f]", " ", s or "")
    if one_line:
        s = re.sub(r"[\r\n  ]+", " ", s)
    return s.strip()[:limit]


def _int(s: str, limit: int = 10) -> str:
    """A numeric id from the plan, or '' when it is anything else."""
    return s if re.fullmatch(r"[0-9]{1,%d}" % limit, s or "") else ""


def _iso(s: str) -> str:
    """A date as the normalised ISO text, or '' when it does not parse."""
    d = _date(s)
    return d.isoformat() if d else ""


def _num(s: str) -> float | None:
    try:
        return float(s) if s else None
    except ValueError:
        return None


def _work(el) -> dict:
    """Work, actual and remaining work in hours (Project writes them as durations)."""
    out = {}
    for key, name in (("work_h", "Work"), ("actual_work_h", "ActualWork"), ("remaining_work_h", "RemainingWork")):
        m = _minutes(_t(el, name))
        out[key] = round(m / 60, 1) if m is not None else None
    return out


def _money(el) -> dict:
    """Cost and the earned-value figures Project keeps per task (planned, earned, actual cost)."""
    return {key: _num(_t(el, name)) for key, name in (("cost", "Cost"), ("actual_cost", "ActualCost"),
                                                      ("bcws", "BCWS"), ("bcwp", "BCWP"), ("acwp", "ACWP"))}


def parse(xml: bytes) -> dict:
    from lxml import etree
    parser = etree.XMLParser(resolve_entities=False, no_network=True, huge_tree=False)
    root = etree.fromstring(xml, parser)
    if root.tag != _q("Project"):
        raise MppError("not Microsoft Project XML (root element is not Project)")
    per_day = float(_t(root, "MinutesPerDay") or 480) or 480.0
    info = {k: _clean(_t(root, k)) for k in ("Name", "Title", "Subject", "Author", "Manager", "Company")}
    info.update({k: _iso(_t(root, k)) for k in ("StartDate", "FinishDate", "StatusDate", "CurrentDate",
                                                "CreationDate", "LastSaved")})
    fields = {}
    for ea in root.iter(_q("ExtendedAttribute")):
        if ea.getparent() is not None and ea.getparent().tag == _q("ExtendedAttributes"):
            fid = _t(ea, "FieldID")
            if fid:
                fields[fid] = _clean(_t(ea, "Alias") or _t(ea, "FieldName"), 80)
    resources = {}
    for r in root.iter(_q("Resource")):
        uid, name = _int(_t(r, "UID")), _clean(_t(r, "Name"), 120)
        if uid and name and _t(r, "IsNull") != "1":
            resources[uid] = {"uid": uid, "name": name, "type": {"0": "material", "1": "work", "2": "cost"}.get(_t(r, "Type"), ""),
                              "overallocated": _t(r, "OverAllocated") == "1", **_work(r)}
    assigned: dict[str, list[str]] = {}
    for a in root.iter(_q("Assignment")):
        ru, tu = _int(_t(a, "ResourceUID")), _int(_t(a, "TaskUID"))
        if ru in resources and tu:
            resources[ru]["tasks"] = resources[ru].get("tasks", 0) + 1
            units = _t(a, "Units")
            share = f" {float(units) * 100:.0f}%" if units and abs(float(units) - 1.0) > 0.001 else ""
            assigned.setdefault(tu, []).append(resources[ru]["name"] + share)
    tasks = []
    for el in root.iter(_q("Task")):
        if _t(el, "OutlineLevel") == "0" and _t(el, "IsNull") != "1":
            # the project summary task: Project's own overall % complete
            info["percent"] = _int(_t(el, "PercentComplete").split(".")[0], 3)
            info["totals"] = {**_work(el), **_money(el)}
        if _t(el, "IsNull") == "1" or _t(el, "OutlineLevel") in ("", "0"):
            continue                                   # the project summary task, empty rows
        base = next((b for b in el.findall(_q("Baseline")) if _t(b, "Number") in ("0", "")), None)
        preds = []
        for pl in el.findall(_q("PredecessorLink")):
            lag = _t(pl, "LinkLag")
            preds.append({"uid": _int(_t(pl, "PredecessorUID")), "type": LINK_TYPES.get(_t(pl, "Type"), "FS"),
                          "lag_min": float(lag) / 10 if lag else 0.0})   # LinkLag is in tenths of a minute
        custom = {}
        for ea in el.findall(_q("ExtendedAttribute")):
            fid, val = _t(ea, "FieldID"), _clean(_t(ea, "Value"), 200)
            if val:
                custom[fields.get(fid) or f"field {fid}"] = val
        slack = _t(el, "TotalSlack")
        tasks.append({
            "uid": _int(_t(el, "UID")), "id": _int(_t(el, "ID")), "name": _clean(_t(el, "Name"), 300),
            "wbs": _clean(_t(el, "WBS"), 40), "outline": _clean(_t(el, "OutlineNumber"), 40),
            "level": max(1, min(int(_int(_t(el, "OutlineLevel"), 3) or 1), 32)),   # Project allows 9 levels
            "start": _iso(_t(el, "Start")), "finish": _iso(_t(el, "Finish")),
            "duration": _days(_minutes(_t(el, "Duration")), per_day),
            "duration_min": _minutes(_t(el, "Duration")) or 0.0,
            "constraint_code": _int(_t(el, "ConstraintType"), 2), **_work(el), **_money(el),
            "percent": int(float(_t(el, "PercentComplete") or 0)),
            "milestone": _t(el, "Milestone") == "1", "summary": _t(el, "Summary") == "1",
            "critical": _t(el, "Critical") == "1", "active": _t(el, "Active") != "0",
            "deadline": _iso(_t(el, "Deadline")), "constraint": CONSTRAINTS.get(_t(el, "ConstraintType"), ""),
            "constraint_date": _iso(_t(el, "ConstraintDate")),
            "actual_start": _iso(_t(el, "ActualStart")), "actual_finish": _iso(_t(el, "ActualFinish")),
            "baseline_start": _iso(_t(base, "Start")), "baseline_finish": _iso(_t(base, "Finish")),
            "total_slack_days": _days(float(slack) / 10, per_day) if slack else "",
            "notes": _clean(_t(el, "Notes"), 2000, one_line=False), "resources": assigned.get(_t(el, "UID"), []),
            "predecessors": preds, "custom": custom,
        })
        if len(tasks) >= MAX_TASKS:
            break
    by_uid = {t["uid"]: t for t in tasks}
    for t in tasks:
        parts = []
        for p in t["predecessors"]:
            ref = by_uid.get(p["uid"], {}).get("id") or f"uid {p['uid']}"
            lag = ""
            if p["lag_min"]:
                d = _days(abs(p["lag_min"]), per_day)
                lag = f"{'+' if p['lag_min'] > 0 else '-'}{d}d"
            parts.append(f"{ref}{'' if p['type'] == 'FS' and not lag else p['type']}{lag}")
        t["predecessor_text"] = ", ".join(parts)
    for t in tasks:
        t.setdefault("successors", [])
        for p in t["predecessors"]:
            if p["uid"] in by_uid:
                by_uid[p["uid"]].setdefault("successors", []).append(t["uid"])
    return {"project": info, "minutes_per_day": per_day, "tasks": tasks, "resources": list(resources.values()),
            "truncated": len(tasks) >= MAX_TASKS}


def attention(data: dict) -> dict:
    """What a reader asked about a plan wants first: late, behind the baseline, deadlines, critical."""
    info = data["project"]
    status = _date(info.get("StatusDate")) or _date(info.get("CurrentDate"))
    late, slipped, missed, critical = [], [], [], []
    for t in data["tasks"]:
        if t["summary"] or not t["active"]:
            continue
        fin = _date(t["finish"])
        if status and fin and fin < status and t["percent"] < 100:
            late.append(t)
        bf = _date(t["baseline_finish"])
        if bf and fin and (fin.date() - bf.date()).days > 0:
            slipped.append((t, (fin.date() - bf.date()).days))
        dl = _date(t["deadline"])
        if dl and fin and fin > dl:
            missed.append(t)
        if t["critical"] and t["percent"] < 100:
            critical.append(t)
    return {"status_date": status.date().isoformat() if status else "", "late": late, "slipped": slipped,
            "deadline_missed": missed, "critical_open": critical}


BEHIND_POINTS = 20      # expected minus actual % complete at or above this: behind schedule
LONG_DAYS = 44          # working days; longer tasks hide progress (a common schedule-review limit)
NEAR_DAYS = 2           # total slack at or below this, above zero: near-critical
HARD = {"2": "must start on", "3": "must finish on", "5": "start no later than", "7": "finish no later than"}


def status_of(data: dict) -> datetime | None:
    info = data["project"]
    return _date(info.get("StatusDate")) or _date(info.get("CurrentDate"))


def expected_percent(t: dict, status: datetime) -> float | None:
    """How far the task should be at the status date, by the baseline dates (else the current ones),
    assuming even progress over its span."""
    s, f = _date(t["baseline_start"] or t["start"]), _date(t["baseline_finish"] or t["finish"])
    if not s or not f:
        return None
    if status <= s:
        return 0.0
    if status >= f or f <= s:
        return 100.0
    return 100.0 * (status - s).total_seconds() / (f - s).total_seconds()


def progress_check(data: dict) -> dict:
    """Tasks behind where they should be, tasks that should have started, and the plan as a whole."""
    status = status_of(data)
    if not status:
        return {"behind": [], "not_started": [], "expected": None, "actual": None, "spi": None, "cpi": None}
    behind, not_started = [], []
    weight = exp_sum = act_sum = 0.0
    for t in data["tasks"]:
        if t["summary"] or not t["active"] or t["milestone"]:
            continue
        e = expected_percent(t, status)
        if e is None:
            continue
        w = t["duration_min"] or 1.0
        weight += w
        exp_sum += w * e
        act_sum += w * t["percent"]
        if t["percent"] < 100 and e - t["percent"] >= BEHIND_POINTS:
            behind.append((t, round(e), t["percent"]))
        s = _date(t["start"])
        if t["percent"] == 0 and s and s < status:
            not_started.append(t)
    tot = data["project"].get("totals") or {}
    bcws = tot.get("bcws") or sum(t["bcws"] or 0 for t in data["tasks"] if not t["summary"])
    bcwp = tot.get("bcwp") or sum(t["bcwp"] or 0 for t in data["tasks"] if not t["summary"])
    acwp = tot.get("acwp") or sum(t["acwp"] or 0 for t in data["tasks"] if not t["summary"])
    return {"behind": sorted(behind, key=lambda x: x[2] - x[1]), "not_started": not_started,
            "expected": round(exp_sum / weight, 1) if weight else None,
            "actual": round(act_sum / weight, 1) if weight else None,
            "spi": round(bcwp / bcws, 2) if bcws else None, "cpi": round(bcwp / acwp, 2) if acwp and bcwp else None}


def health(data: dict) -> dict:
    """Ways of building a plan that make its dates untrustworthy, each with the tasks affected."""
    status = status_of(data)
    tasks = [t for t in data["tasks"] if t["active"]]
    leaves = [t for t in tasks if not t["summary"]]
    open_leaves = [t for t in leaves if t["percent"] < 100]
    last = max((t for t in leaves if _date(t["finish"])), key=lambda t: _date(t["finish"]), default=None)
    first = min((t for t in leaves if _date(t["start"])), key=lambda t: _date(t["start"]), default=None)
    per_day = data["minutes_per_day"]
    out = {
        "no_predecessor": [t for t in open_leaves if not t["predecessors"] and t is not first],
        "no_successor": [t for t in open_leaves if not t.get("successors") and t is not last],
        "hard_constraint": [(t, HARD[t["constraint_code"]]) for t in open_leaves if t["constraint_code"] in HARD],
        "negative_slack": [t for t in open_leaves if t["total_slack_days"] and float(t["total_slack_days"]) < 0],
        "lead": [(t, p) for t in leaves for p in t["predecessors"] if p["lag_min"] < 0],
        "long": [t for t in open_leaves if not t["milestone"] and t["duration_min"] / per_day > LONG_DAYS],
        "summary_links": [t for t in tasks if t["summary"] and t["predecessors"]],
        "future_actuals": [],
    }
    if status:
        for t in leaves:
            for label, key in (("actual start", "actual_start"), ("actual finish", "actual_finish")):
                d = _date(t[key])
                if d and d > status:
                    out["future_actuals"].append((t, label, t[key]))
    return out


def driving_path(data: dict, limit: int = 60) -> list[dict]:
    """From the task that finishes last, back through the predecessor that drives each date
    (the one whose linked date plus lag is latest). The chain that sets the plan's finish."""
    by_uid = {t["uid"]: t for t in data["tasks"]}
    leaves = [t for t in data["tasks"] if not t["summary"] and t["active"] and _date(t["finish"])]
    if not leaves:
        return []
    cur = max(leaves, key=lambda t: (_date(t["finish"]), t["critical"]))
    chain, seen = [cur], {cur["uid"]}
    while len(chain) < limit:
        best, best_date = None, None
        for p in cur["predecessors"]:
            pt = by_uid.get(p["uid"])
            if not pt or pt["uid"] in seen:
                continue
            d = _date(pt["start"] if p["type"] in ("SS", "SF") else pt["finish"])
            if d and (best_date is None or d > best_date):
                best, best_date = pt, d
        if best is None:
            break
        if best["summary"]:
            break
        chain.append(best)
        seen.add(best["uid"])
        cur = best
    return list(reversed(chain))


def near_critical(data: dict) -> list[dict]:
    out = []
    for t in data["tasks"]:
        if t["summary"] or t["critical"] or t["percent"] >= 100 or not t["total_slack_days"]:
            continue
        slack = float(t["total_slack_days"])
        if 0 < slack <= NEAR_DAYS:
            out.append(t)
    return out


# -- Markdown ---------------------------------------------------------------------------------

def _d(s: str) -> str:
    return s[:10] if s else ""


def _label(t: dict) -> str:
    return f"{t['id']} {office2md.md_escape(t['name'])}"   # id is digits only (parse)


def gantt(data: dict, limit: int = 300) -> list[str]:
    def safe(s: str) -> str:
        return re.sub(r"[:#;`\n]", " ", s).strip() or "task"
    title = safe(data["project"].get("Title") or data["project"].get("Name") or "plan")
    out = ["```mermaid", "gantt", f"    title {title}", "    dateFormat YYYY-MM-DD"]
    n = 0
    for t in data["tasks"]:
        if n >= limit:
            break
        if t["summary"] and t["level"] == 1:
            out.append(f"    section {safe(t['name'])}")
            continue
        if t["summary"] or not t["start"] or not t["uid"]:
            continue
        tags = (["milestone"] if t["milestone"] else []) + (["crit"] if t["critical"] else []) \
            + (["done"] if t["percent"] >= 100 else ["active"] if t["percent"] > 0 else [])
        end = "0d" if t["milestone"] else _d(t["finish"])
        out.append(f"    {safe(t['name'])} :{', '.join(tags + ['t' + t['uid']])}, {_d(t['start'])}, {end}")
        n += 1
    out.append("```")
    if n >= limit:
        out.append(f"_(gantt shows the first {limit} tasks; the table has all of them)_")
    return out


def progress_markdown(data: dict) -> list[str]:
    pc = progress_check(data)
    if pc["expected"] is None:
        return []
    out = [f"Progress against the schedule: {pc['actual']}% done where the dates call for {pc['expected']}% "
           f"(baseline dates where saved, assuming even progress)."]
    if pc["spi"] is not None:
        out[-1] += (f" Earned value: SPI {pc['spi']} (below 1 = behind schedule)"
                    + (f", CPI {pc['cpi']} (below 1 = over cost)" if pc["cpi"] is not None else "") + ".")
    out.append(f"Behind their expected progress by {BEHIND_POINTS} points or more (finish not yet passed is "
               f"included, so these are caught early): {len(pc['behind'])}")
    out += [f"- {_label(t)} — {act}% done, {exp}% expected by now, finish {_d(t['finish'])}" for t, exp, act in pc["behind"][:50]]
    out.append(f"Should have started and have not: {len(pc['not_started'])}")
    out += [f"- {_label(t)} — planned start {_d(t['start'])}" for t in pc["not_started"][:50]]
    return out


def health_markdown(data: dict) -> list[str]:
    h = health(data)
    rows = [
        ("no_predecessor", "No predecessor: nothing drives its start, so a delay upstream does not move it",
         lambda t: _label(t)),
        ("no_successor", "No successor: its delay moves nothing, so the plan finish cannot show it", lambda t: _label(t)),
        ("hard_constraint", "Fixed date constraint: the date is typed in, not calculated from the work before it",
         lambda x: f"{_label(x[0])} — {x[1]} {_d(x[0]['constraint_date'])}"),
        ("negative_slack", "Negative slack: already cannot meet a date it is tied to",
         lambda t: f"{_label(t)} — {t['total_slack_days']} days"),
        ("lead", "Negative lag (starts before its predecessor allows)",
         lambda x: f"{_label(x[0])} — {x[0]['predecessor_text']}"),
        ("long", f"Longer than {LONG_DAYS} working days: progress inside it is hard to see",
         lambda t: f"{_label(t)} — {t['duration']} days"),
        ("summary_links", "Summary task with its own predecessors: links belong on the work tasks",
         lambda t: _label(t)),
        ("future_actuals", "Actual date after the status date: progress recorded in the future",
         lambda x: f"{_label(x[0])} — {x[1]} {_d(x[2])}"),
    ]
    found = [(title, h[key], fmt) for key, title, fmt in rows if h[key]]
    out = ["## Schedule health", ""]
    if not found:
        return out + ["No problems found in how the plan is built (links, constraints, slack, durations, actuals).", ""]
    out += ["How the plan is built, not how the work is going: each item makes the dates less trustworthy.", ""]
    for title, items, fmt in found:
        out.append(f"{title}: {len(items)}")
        out += [f"- {fmt(x)}" for x in items[:30]] + ([f"- … {len(items) - 30} more"] if len(items) > 30 else [])
    return out + [""]


def path_markdown(data: dict) -> list[str]:
    chain = driving_path(data)
    near = near_critical(data)
    out = []
    if len(chain) > 1:
        out += ["## What sets the finish date", "",
                "The chain of tasks behind the last finish, each driven by the one before it. Shortening or "
                "starting any of these earlier moves the finish; work off this chain does not.", ""]
        out += [f"{i}. {_label(t)} — {_d(t['start'])} → {_d(t['finish'])}, {t['percent']}%"
                + (" (critical)" if t["critical"] else "") for i, t in enumerate(chain, start=1)]
        out.append("")
    if near:
        out += [f"## Near-critical (total slack {NEAR_DAYS} days or less)", "",
                "Not critical yet; a small slip makes them so.", ""]
        out += [f"- {_label(t)} — {t['total_slack_days']} days of slack, finish {_d(t['finish'])}" for t in near[:50]]
        out.append("")
    return out


def work_markdown(data: dict) -> list[str]:
    tot = data["project"].get("totals") or {}
    work = tot.get("work_h")
    if not work:
        return []
    done = tot.get("actual_work_h") or 0
    out = ["## Work and cost", "",
           f"Work: {work:g} h planned, {done:g} h done, {tot.get('remaining_work_h') or 0:g} h remaining "
           f"({100 * done / work:.0f}% of the work)."]
    if tot.get("cost"):
        out.append(f"Cost: {tot['cost']:g} planned, {tot.get('actual_cost') or 0:g} spent.")
    rows = [[_label(t), f"{t['work_h'] or 0:g}", f"{t['actual_work_h'] or 0:g}", f"{t['remaining_work_h'] or 0:g}",
             f"{t['cost']:g}" if t["cost"] else ""] for t in data["tasks"]
            if t["summary"] and t["level"] == 1 and t["work_h"]]
    if rows:
        out += ["", office2md.md_table(rows, header=["summary task", "work h", "done h", "remaining h", "cost"])]
    return out + [""]


def to_markdown(data: dict, path: Path, engine: str, with_gantt: bool = False) -> str:
    info = data["project"]
    att = attention(data)
    tasks = data["tasks"]
    out = office2md.md_title(str(path))
    out.append(f"Read with: {engine}.")
    head = " · ".join(x for x in [
        office2md.md_escape(info.get("Title") or info.get("Name") or ""),
        f"{_d(info.get('StartDate'))} → {_d(info.get('FinishDate'))}" if info.get("StartDate") else "",
        f"status date {att['status_date']}" if att["status_date"] else "no status date",
        f"{len(tasks)} tasks ({sum(t['milestone'] for t in tasks)} milestones)",
        f"{len(data['resources'])} resources",
        f"manager {office2md.md_escape(info['Manager'])}" if info.get("Manager") else ""] if x)
    out += [head, ""]
    if data.get("truncated"):
        out += [f"> Only the first {MAX_TASKS} tasks are read.", ""]

    out += ["## Needs attention", ""]
    if att["status_date"]:
        out.append(f"Late (should have finished by {att['status_date']}, not complete): {len(att['late'])}")
        out += [f"- {_label(t)} — finish {_d(t['finish'])}, {t['percent']}% done" for t in att["late"][:50]]
    else:
        out.append("Late: not judged, the plan has no status date.")
    out.append(f"Behind the baseline (finish later than planned): {len(att['slipped'])}")
    out += [f"- {_label(t)} — finish {_d(t['finish'])}, baseline {_d(t['baseline_finish'])} (+{days} days)"
            for t, days in sorted(att["slipped"], key=lambda x: -x[1])[:50]]
    if not any(t["baseline_finish"] for t in tasks):
        out.append("- (no baseline saved in the plan)")
    out.append(f"Past their deadline: {len(att['deadline_missed'])}")
    out += [f"- {_label(t)} — finish {_d(t['finish'])}, deadline {_d(t['deadline'])}" for t in att["deadline_missed"][:50]]
    out.append(f"Critical and not complete: {len(att['critical_open'])}")
    out += [f"- {_label(t)} — {_d(t['start'])} → {_d(t['finish'])}" for t in att["critical_open"][:50]]
    out.append("")

    out += progress_markdown(data)
    out.append("")
    out += health_markdown(data) + path_markdown(data) + work_markdown(data)
    out += ["## Tasks", "",
            "Predecessors are task IDs with the link type when not finish-to-start (SS, FF, SF) and the lag in "
            "working days. Summary tasks are bold; the outline number gives the hierarchy.", ""]
    rows = []
    for t in tasks:
        name = office2md.md_escape(t["name"])
        name = ("　" * (t["level"] - 1)) + (f"**{name}**" if t["summary"] else name)
        flags = []
        if t["milestone"]:
            flags.append("milestone")
        if t["critical"]:
            flags.append("critical")
        if not t["active"]:
            flags.append("inactive")
        if t["deadline"]:
            flags.append(f"deadline {_d(t['deadline'])}")
        if t["constraint"]:
            flags.append(f"{t['constraint']} {_d(t['constraint_date'])}".strip())
        if t["baseline_finish"] and _d(t["baseline_finish"]) != _d(t["finish"]):
            flags.append(f"baseline finish {_d(t['baseline_finish'])}")
        rows.append([t["id"], office2md.md_escape(t["outline"] or t["wbs"]), name, _d(t["start"]), _d(t["finish"]), t["duration"],
                     f"{t['percent']}%", t["predecessor_text"], office2md.md_escape(", ".join(t["resources"])),
                     ", ".join(flags)])
    out.append(office2md.md_table(rows, header=["ID", "Outline", "Task", "Start", "Finish", "Days", "Done",
                                               "Predecessors", "Resources", "Notes on dates"]) if rows else "_(no tasks)_")
    custom = [t for t in tasks if t["custom"]]
    if custom:
        keys = sorted({k for t in custom for k in t["custom"]})
        if len(keys) > 50:
            out += ["", f"> {len(keys)} custom fields; the table shows the first 50 by name (--json has all)."]
            keys = keys[:50]
        out += ["", "## Custom fields", "",
                office2md.md_table([[t["id"], office2md.md_escape(t["name"])] + [office2md.md_escape(t["custom"].get(k, ""))
                                                                                 for k in keys] for t in custom],
                                   header=["ID", "Task"] + [office2md.md_escape(k) for k in keys])]
    noted = [t for t in tasks if t["notes"]]
    if noted:
        out += ["", "## Task notes", ""]
        for t in noted:
            out.append(f"- {_label(t)}: " + office2md.md_para(t["notes"]).replace("\n", " / "))
    if data["resources"]:
        over = [r for r in data["resources"] if r["overallocated"]]
        out += ["", "## Resources", ""]
        if over:
            out += ["Over-allocated (Project's own flag: more work assigned than the resource has time for): "
                    + ", ".join(office2md.md_escape(r["name"]) for r in over), ""]
        out.append(office2md.md_table(
            [[office2md.md_escape(r["name"]), r["type"], r.get("tasks", 0),
              "" if r.get("remaining_work_h") is None else f"{r['remaining_work_h']:g}",
              "yes" if r["overallocated"] else ""] for r in data["resources"]],
            header=["resource", "type", "tasks", "remaining work h", "over-allocated"]))
    if with_gantt:
        out += ["", "## Gantt", "", "Drawn from the dates above; check against the plan before relying on it.", ""]
        out += gantt(data)
    return "\n".join(out) + "\n"


def outline_markdown(data: dict, path: Path, engine: str) -> str:
    out = office2md.md_title(str(path)) + [f"Read with: {engine}.", ""]
    for t in data["tasks"]:
        if t["level"] <= 2:
            out.append("  " * (t["level"] - 1) + f"- {office2md.md_escape(t['outline'])} {office2md.md_escape(t['name'])} "
                       f"({_d(t['start'])} → {_d(t['finish'])}, {t['percent']}%)")
    return "\n".join(out) + "\n"


def read(path: Path, engine: str = "auto", save_xml: Path | None = None) -> dict:
    xml, label = to_mspdi(path, engine, save_xml)
    data = parse(xml)
    data["engine"] = label
    data["file"] = str(path)
    data["type"] = "mpp"
    return data


def main(argv=None) -> int:
    argv = sys.argv[1:] if argv is None else argv
    if argv[:1] == ["--mpxj-convert"] and len(argv) == 3:
        return _mpxj_child(argv[1], argv[2])
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("file", type=Path)
    ap.add_argument("--engine", choices=["auto", "project", "mpxj"], default="auto",
                    help="how to convert a .mpp: auto tries Microsoft Project, then MPXJ")
    ap.add_argument("--gantt", action="store_true", help="add a mermaid gantt chart")
    ap.add_argument("--outline", action="store_true", help="top two levels only")
    ap.add_argument("--json", action="store_true")
    ap.add_argument("--out", type=Path, help="write to this file")
    ap.add_argument("--save-xml", type=Path, help="keep the converted Project XML here")
    a = ap.parse_args(argv)
    if not a.file.exists():
        sys.exit(f"[mpp_reader] file not found: {a.file}")
    try:
        data = read(a.file, a.engine, a.save_xml)
    except MppError as e:
        sys.exit(f"[mpp_reader] {e}")
    if a.json:
        text = json.dumps(data, ensure_ascii=False, indent=2, default=str)
    elif a.outline:
        text = outline_markdown(data, a.file, data["engine"])
    else:
        text = to_markdown(data, a.file, data["engine"], a.gantt)
    if a.out:
        a.out.write_text(text, encoding="utf-8")
        print(f"[mpp_reader] wrote {a.out} ({len(text)} chars)")
    else:
        print(text, end="")
    return 0


if __name__ == "__main__":
    sys.exit(main())
