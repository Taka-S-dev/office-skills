#!/usr/bin/env python3
"""Check the mermaid and PlantUML diagrams in a Markdown file before anyone renders them.

Usage:
    diagram_lint.py file.md                 report problems, exit 1 if any
    diagram_lint.py file.md --render DIR    also render with PlantUML / mermaid-cli when installed

A diagram reconstructed from a drawing, then corrected by hand, is easy to break: a participant
renamed in one line but not another, an `alt` without its `end`, a `deactivate` with nothing
active. Renderers either refuse such a diagram or draw something different from what was meant.
This checks the structure of sequence diagrams and flowcharts without any renderer installed.
When PlantUML (PLANTUML_JAR or `plantuml` on PATH) or mermaid-cli (`mmdc`) is available, --render
also runs the real parser and saves an image next to the Markdown for a side-by-side look.
"""

from __future__ import annotations

import argparse
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")  # Windows console defaults to cp932
except AttributeError:
    pass

FENCE = re.compile(r"^```(mermaid|plantuml|puml)\s*\n(.*?)^```", re.M | re.S)
MER_ARROW = re.compile(r"^(?P<a>[^\s:;\-<>]+?)\s*(?P<arrow><<->>|<<-->>|-->>|->>|-->|->|--x|-x|--\)|-\))\s*[+-]?"
                       r"(?P<b>[^\s:;\-<>]+?)\s*:\s*(?P<text>.*)$")
MER_BLOCK_OPEN = re.compile(r"^(loop|alt|opt|par|critical|break|rect|box)\b")
MER_SEPARATORS = {"else": {"alt"}, "and": {"par"}, "option": {"critical"}}
PUML_ARROW = re.compile(r"^(?P<a>\"[^\"]+\"|[^\s\"]+)\s*(?P<arrow>[<]?[-.]+[>]{1,2}[xo]?|[xo]?[<]{1,2}[-.]+)\s*"
                        r"(?P<b>\"[^\"]+\"|[^\s\":]+)\s*(:\s*(?P<text>.*))?$")
PUML_BLOCK_OPEN = re.compile(r"^(alt|opt|loop|par|break|critical|group|box)\b")


def blocks(md: str) -> list[tuple[str, int, str]]:
    """(language, line number of the fence, code) for every diagram block."""
    out = []
    for m in FENCE.finditer(md):
        lang = "plantuml" if m.group(1) in ("plantuml", "puml") else "mermaid"
        out.append((lang, md.count("\n", 0, m.start()) + 1, m.group(2)))
    return out


def lint_mermaid(code: str) -> list[str]:
    lines = [ln.strip() for ln in code.strip("\n").split("\n")]
    lines = [ln for ln in lines if ln and not ln.startswith("%%")]
    if not lines:
        return ["empty diagram"]
    kind = lines[0].split()[0]
    if kind == "sequenceDiagram":
        return _lint_mermaid_sequence(lines[1:])
    if kind in ("flowchart", "graph"):
        return _lint_mermaid_flow(lines[1:])
    return []          # other diagram kinds: not checked


def _lint_mermaid_sequence(lines: list[str]) -> list[str]:
    errs = []
    declared: set[str] = set()
    stack: list[str] = []
    active: dict[str, int] = {}
    for i, ln in enumerate(lines, start=2):
        m = re.match(r"^(participant|actor)\s+(\S+?)(\s+as\s+(.+))?$", ln)
        if m:
            declared.add(m.group(2))
            continue
        if ";" in ln:
            errs.append(f"line {i}: ';' ends a statement in mermaid and cuts the text: {ln[:60]}")
        if ln == "end":
            if not stack:
                errs.append(f"line {i}: 'end' without an open block")
            else:
                stack.pop()
            continue
        mo = MER_BLOCK_OPEN.match(ln)
        if mo:
            stack.append(mo.group(1))
            continue
        word = ln.split()[0]
        if word in MER_SEPARATORS:
            if not stack or stack[-1] not in MER_SEPARATORS[word]:
                errs.append(f"line {i}: '{word}' outside {'/'.join(MER_SEPARATORS[word])}")
            continue
        m = re.match(r"^(activate|deactivate)\s+(\S+)$", ln)
        if m:
            p = m.group(2)
            if declared and p not in declared:
                errs.append(f"line {i}: {m.group(1)} of undeclared participant '{p}'")
            active[p] = active.get(p, 0) + (1 if m.group(1) == "activate" else -1)
            if active[p] < 0:
                errs.append(f"line {i}: deactivate {p} while it is not active")
                active[p] = 0
            continue
        m = re.match(r"^[Nn]ote\s+(over|left of|right of)\s+([^:]+):", ln)
        if m:
            for p in [x.strip() for x in m.group(2).split(",")]:
                if declared and p not in declared:
                    errs.append(f"line {i}: note on undeclared participant '{p}'")
            continue
        m = MER_ARROW.match(ln)
        if m:
            for p in (m.group("a"), m.group("b")):
                if declared and p not in declared:
                    errs.append(f"line {i}: message uses undeclared participant '{p}' (typo, or missing 'participant' line)")
            continue
        if re.match(r"^(title|autonumber|link|links|properties|details|create|destroy)\b", ln):
            continue
        errs.append(f"line {i}: not a mermaid sequence statement: {ln[:60]}")
    for b in reversed(stack):
        errs.append(f"'{b}' block is never closed with 'end'")
    return errs


def _lint_mermaid_flow(lines: list[str]) -> list[str]:
    errs = []
    for i, ln in enumerate(lines, start=2):
        for o, c in (("[", "]"), ("(", ")"), ("{", "}")):
            if ln.count(o) != ln.count(c):
                errs.append(f"line {i}: unbalanced '{o}{c}': {ln[:60]}")
        if ln.count('"') % 2:
            errs.append(f"line {i}: unbalanced quote: {ln[:60]}")
    return errs


def lint_plantuml(code: str) -> list[str]:
    lines = [ln.strip() for ln in code.strip("\n").split("\n")]
    lines = [ln for ln in lines if ln and not ln.startswith("'")]
    errs = []
    if not lines or lines[0] != "@startuml" or lines[-1] != "@enduml":
        errs.append("must start with @startuml and end with @enduml")
    body = [ln for ln in lines if ln not in ("@startuml", "@enduml")]
    declared: set[str] = set()
    stack: list[str] = []
    is_sequence = any(PUML_ARROW.match(ln) and not ln.startswith(("rectangle", "node", "component")) for ln in body) \
        and not any(ln.startswith(("rectangle", "node", "component", "usecase", "class ")) for ln in body)
    for i, ln in enumerate(body, start=2):
        m = re.match(r'^(participant|actor|boundary|control|entity|database|collections|queue|rectangle|node|component)\s+'
                     r'(?:"[^"]*"\s+as\s+(\S+)|(\S+))', ln)
        if m:
            declared.add(m.group(2) or m.group(3))
            continue
        if ln == "end" or ln.startswith("end "):
            if not stack:
                errs.append(f"line {i}: 'end' without an open block")
            else:
                stack.pop()
            continue
        mo = PUML_BLOCK_OPEN.match(ln)
        if mo:
            stack.append(mo.group(1))
            continue
        if ln.split()[0] == "else":
            if not stack or stack[-1] not in ("alt", "par", "critical", "group"):
                errs.append(f"line {i}: 'else' outside alt/par")
            continue
        m = re.match(r"^(activate|deactivate)\s+(\S+)", ln)
        if m:
            if declared and m.group(2) not in declared:
                errs.append(f"line {i}: {m.group(1)} of undeclared participant '{m.group(2)}'")
            continue
        if re.match(r"^note\b", ln, re.I):
            continue
        m = PUML_ARROW.match(ln)
        if m:
            if is_sequence and declared:
                for p in (m.group("a"), m.group("b")):
                    if not p.startswith('"') and p not in declared:
                        errs.append(f"line {i}: message uses undeclared participant '{p}'")
            continue
        if re.match(r"^(title|autonumber|skinparam|hide|show|left to right|top to bottom|==|\.\.\.|\|\|\||header|footer)", ln):
            continue
        errs.append(f"line {i}: not a recognised PlantUML statement: {ln[:60]}")
    for b in reversed(stack):
        errs.append(f"'{b}' block is never closed with 'end'")
    return errs


def lint_markdown(md: str) -> list[tuple[str, int, list[str]]]:
    """[(language, fence line, errors)] for blocks with errors."""
    out = []
    for lang, line, code in blocks(md):
        errs = lint_mermaid(code) if lang == "mermaid" else lint_plantuml(code)
        if errs:
            out.append((lang, line, errs))
    return out


def _which(name: str) -> str | None:
    """shutil.which over PATH only: never the current directory, where a file being processed
    could have planted a same-named executable."""
    cwd = Path.cwd().resolve()
    dirs = []
    for p in os.environ.get("PATH", "").split(os.pathsep):
        if not p or p == ".":
            continue
        try:
            if Path(p).resolve() == cwd:
                continue
        except OSError:
            continue
        dirs.append(p)
    return shutil.which(name, path=os.pathsep.join(dirs)) if dirs else None


def renderers() -> dict[str, list[str]]:
    found = {}
    jar = os.environ.get("PLANTUML_JAR")
    java = _which("java")
    if jar and Path(jar).is_file() and java:
        # SANDBOX: no file inclusion, no URL fetch, no environment access from diagram text
        found["plantuml"] = [java, "-DPLANTUML_SECURITY_PROFILE=SANDBOX", "-jar", str(Path(jar).resolve())]
    elif _which("plantuml"):
        found["plantuml"] = [_which("plantuml")]
    if _which("mmdc"):
        found["mermaid"] = [_which("mmdc")]
    return found


def render(md: str, out_dir: Path) -> list[tuple[str, int, str, str]]:
    """Render each block with the real tool when installed: [(language, line, image path or '', error)].

    Each block is rendered in a fresh temporary directory under a fixed name, with that directory
    as the working directory, so no path derived from the document reaches the tool's command line
    (mmdc is a .cmd wrapper on Windows, where batch argument parsing is unsafe)."""
    import tempfile
    tools = renderers()
    results = []
    out_dir.mkdir(parents=True, exist_ok=True)
    env = {**os.environ, "PLANTUML_SECURITY_PROFILE": "SANDBOX"}
    for n, (lang, line, code) in enumerate(blocks(md), start=1):
        if lang not in tools:
            continue
        ext = "puml" if lang == "plantuml" else "mmd"
        with tempfile.TemporaryDirectory() as td:
            (Path(td) / f"diagram.{ext}").write_text(code, encoding="utf-8")
            cmd = tools[lang] + ([f"diagram.{ext}", "-tpng", "-charset", "UTF-8"] if lang == "plantuml"
                                 else ["-i", f"diagram.{ext}", "-o", "diagram.png"])
            try:
                r = subprocess.run(cmd, capture_output=True, text=True, timeout=120, cwd=td, env=env)
                made = Path(td) / "diagram.png"
                err = "" if r.returncode == 0 and made.exists() else (r.stderr or r.stdout).strip()[:300] or "render failed"
            except (subprocess.TimeoutExpired, OSError) as e:
                made, err = Path(td) / "none", str(e)[:200]
            img = out_dir / f"diagram{n:02d}.png"
            if made.exists():
                shutil.copyfile(made, img)
        results.append((lang, line, str(img) if img.exists() and not err else "", err))
    return results


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("file", type=Path)
    ap.add_argument("--render", type=Path, metavar="DIR", help="also render with installed PlantUML / mermaid-cli")
    a = ap.parse_args(argv)
    md = a.file.read_text(encoding="utf-8")
    found = blocks(md)
    problems = lint_markdown(md)
    print(f"{len(found)} diagram block(s), {len(problems)} with problems")
    for lang, line, errs in problems:
        for e in errs:
            print(f"- {lang} block at line {line}: {e}")
    if a.render:
        tools = renderers()
        if not tools:
            print("no renderer installed (PLANTUML_JAR / plantuml / mmdc); structure check only")
        for lang, line, img, err in render(md, a.render):
            print(f"- {lang} block at line {line}: " + (f"render error: {err}" if err else f"rendered {img}"))
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
