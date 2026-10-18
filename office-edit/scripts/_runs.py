"""Run-spanning text replacement shared by docx_replace.py and pptx_replace.py.

Word and PowerPoint split a paragraph into "runs" at every formatting change, spell-check
boundary, or edit point. "東京支店" is often stored as "東京" + "支店", so a per-run replace
silently finds nothing. This works on the paragraph's joined text and then writes the result
back into the runs, keeping the first run's formatting for the replacement.
"""

from __future__ import annotations

import re
from typing import Iterable


def replace_in_runs(runs: list, old: str, new: str, regex: bool = False) -> int:
    """Replace every occurrence of `old` across `runs` (objects with a .text str property).

    Returns the number of replacements. Formatting of the run where a match starts is kept;
    text of later runs covered by the match is removed.
    """
    if not runs:
        return 0
    pattern = re.compile(old if regex else re.escape(old))
    count = 0
    while True:
        texts = [r.text or "" for r in runs]
        full = "".join(texts)
        m = pattern.search(full)
        if not m:
            return count
        start, end = m.start(), m.end()
        if start == end:  # zero-width match guard
            return count
        replacement = m.expand(new) if regex else new

        # map char offsets to run indices
        pos = 0
        first = last = None
        first_off = last_off = 0
        for i, t in enumerate(texts):
            run_start, run_end = pos, pos + len(t)
            if first is None and run_start <= start < run_end:
                first, first_off = i, start - run_start
            if run_start < end <= run_end:
                last, last_off = i, end - run_start
                break
            pos = run_end
        if first is None or last is None:
            return count

        if first == last:
            t = texts[first]
            runs[first].text = t[:first_off] + replacement + t[last_off:]
        else:
            runs[first].text = texts[first][:first_off] + replacement
            for i in range(first + 1, last):
                runs[i].text = ""
            runs[last].text = texts[last][last_off:]
        count += 1


def find_in_runs(runs: Iterable, old: str, regex: bool = False) -> int:
    full = "".join((r.text or "") for r in runs)
    pattern = re.compile(old if regex else re.escape(old))
    return len(pattern.findall(full))
