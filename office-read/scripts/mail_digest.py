#!/usr/bin/env python3
"""Timeline of the mails in a folder: who wrote what to whom, when, with which attachments.

Usage:
    mail_digest.py ./案件フォルダ                      chronological table of every .eml / .msg
    mail_digest.py ./案件フォルダ --bodies 300          add the first 300 characters of each body
    mail_digest.py ./案件フォルダ --threads             group by subject (Re:/Fw: stripped)
    mail_digest.py ./案件フォルダ --extract ./attachments   save all attachments, numbered by mail

Mails kept next to a document explain why it looks the way it does. Reading them one by one
loses the order; this prints the sequence first so the story can be followed, then each mail
can be opened with office2md when the detail matters.
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import office2md  # noqa: E402

try:
    sys.stdout.reconfigure(encoding="utf-8")  # Windows console defaults to cp932
except AttributeError:
    pass

PREFIX = re.compile(r"^\s*((re|fw|fwd|返信|転送)\s*[:：]\s*|\[[^\]]{1,30}\]\s*|【[^】]{1,30}】\s*)+", re.I)


def thread_key(subject: str | None) -> str:
    s = PREFIX.sub("", subject or "")
    return re.sub(r"\s+", " ", s).strip().lower() or "(no subject)"


def one_line(s, n: int) -> str:
    s = office2md.md_escape(s).replace("<br>", " / ")
    return s if len(s) <= n else s[: n - 1] + "…"


def summarize(path: Path) -> str:
    """One line about a readable file: what is inside, from the handler's outline."""
    ext = path.suffix.lower()
    cls = office2md.HANDLERS.get(ext)
    if cls is None:
        return "(not readable with office2md)"
    opts = argparse.Namespace(outline=True, profile=False, sheet=None, slide=None, max_rows=200, max_cols=50,
                              formulas=False, cells=False, form=False, no_coords=True, raw=False, json=False,
                              out=None, extract_attachments=None, mail_engine="auto", password=None)
    try:
        o = cls(path, opts).outline()
    except Exception as e:
        return f"(could not read: {type(e).__name__})"
    t = o.get("type")
    if t == "xlsx":
        sheets = [f"{s['name']}{' (hidden)' if s['state'] != 'visible' else ''} {s['max_row']}x{s['max_col']}"
                  + (f" {s['formula_cells']}f" if s.get("formula_cells") else "") for s in o["sheets"]]
        return "sheets: " + "; ".join(sheets) + ("; objects: " + ", ".join(o["package_objects"]) if o.get("package_objects") else "")
    if t == "docx":
        heads = [h["text"] for h in o.get("headings", [])][:6]
        return f"{o['paragraphs']} paragraphs, {o['tables']} tables" + (f"; headings: {' / '.join(heads)}" if heads else "") \
            + ("; tracked changes" if o.get("tracked_changes") else "")
    if t == "pptx":
        titles = [s["title"] for s in o["slides"] if s["title"]][:6]
        return f"{o['slide_count']} slides" + (f": {' / '.join(titles)}" if titles else "")
    if t == "pdf":
        return f"{o['pages']} pages" + (f"; bookmarks: {len(o['bookmarks'])}" if o.get("bookmarks") else "")
    if t == "csv":
        return f"{o['rows']} rows x {o['cols']} cols; header: {', '.join(o['header'][:8])}"
    if t == "mail":
        return f"mail: {o.get('subject')} ({o.get('date')})"
    if t == "archive":
        return f"archive with {len(o['entries'])} file(s)"
    return t or "?"


def deep_outline(mail_no: int, att_name: str, saved: Path, password: str | None) -> list[list]:
    """Rows [mail#, attachment, file, summary] for the attachment and, for archives, its contents."""
    sys.path.insert(0, str(Path(__file__).parent))
    import unpack
    rows: list[list] = []
    if saved.suffix.lower() in unpack.ARCHIVE_EXTS:
        out = saved.with_suffix("")
        try:
            files = unpack.unpack(saved, out, password, depth=1)
        except PermissionError:
            return [[mail_no, att_name, saved.name, "(encrypted archive; rerun with --password)"]]
        except Exception as e:
            return [[mail_no, att_name, saved.name, f"(could not unpack: {type(e).__name__})"]]
        rows.append([mail_no, att_name, "", f"archive, {len(files)} file(s) unpacked to {out}"])
        for inner, target in files:
            if target.suffix.lower() in unpack.ARCHIVE_EXTS:
                continue
            rows.append([mail_no, att_name, inner, summarize(target)])
    else:
        rows.append([mail_no, att_name, saved.name, summarize(saved)])
    return rows


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("folder", type=Path)
    ap.add_argument("--bodies", type=int, default=0, metavar="N", help="include the first N characters of each body")
    ap.add_argument("--threads", action="store_true", help="group by normalised subject")
    ap.add_argument("--extract", type=Path, metavar="DIR", help="save every attachment into DIR")
    ap.add_argument("--deep", action="store_true",
                    help="with --extract: unpack zip/7z attachments and outline every readable file (sheets, headings, slides)")
    ap.add_argument("--password", help="password for zip/7z attachments (used with --deep)")
    ap.add_argument("--mail-engine", choices=["auto", "outlook", "olefile"], default="auto")
    ap.add_argument("--max-mails", type=int, default=500)
    a = ap.parse_args(argv)
    if not a.folder.is_dir():
        sys.exit(f"[mail_digest] not a folder: {a.folder}")

    files = sorted(p for p in a.folder.rglob("*") if p.suffix.lower() in (".eml", ".msg") and not p.name.startswith("~$"))
    if len(files) > a.max_mails:
        print(f"[mail_digest] {len(files)} mails; reading the first {a.max_mails} (raise --max-mails)", file=sys.stderr)
        files = files[: a.max_mails]
    mails, failed = [], []
    for f in files:
        opts = argparse.Namespace(extract_attachments=None, mail_engine=a.mail_engine)
        try:
            m = office2md.HANDLERS[f.suffix.lower()](f, opts)._message()
        except Exception as e:
            failed.append(f"{f.name}: {type(e).__name__}: {str(e)[:80]}")
            continue
        m["_file"] = f
        mails.append(m)
    mails.sort(key=lambda m: (m.get("date") or "9999", m["_file"].name))

    deep_rows: list[list] = []
    if a.extract:
        a.extract.mkdir(parents=True, exist_ok=True)
        for i, m in enumerate(mails, start=1):
            for att in m["attachments"]:
                if att.get("data") is None:
                    continue
                target = a.extract / f"{i:03d}_{office2md._safe_name(att['name'])}"
                target.write_bytes(att["data"])
                att["saved"] = str(target)
                if a.deep:
                    deep_rows += deep_outline(i, att["name"], target, a.password)
    elif a.deep:
        sys.exit("[mail_digest] --deep needs --extract DIR")

    print(f"# {a.folder.name}: {len(mails)} mail(s)")
    print()
    print(office2md.DATA_NOTICE)
    print()
    if not mails:
        print("No .eml or .msg files found.")
    header = ["#", "date", "from", "to", "subject", "attachments"]

    def row(i, m):
        atts = ", ".join(one_line(x["name"], 40) for x in m["attachments"])
        return [i, (m.get("date") or "?")[:16], one_line(m.get("from"), 30), one_line(m.get("to"), 30),
                one_line(m.get("subject"), 60), atts]

    if a.threads:
        groups: dict[str, list] = {}
        order: list[tuple[int, dict]] = list(enumerate(mails, start=1))
        for i, m in order:
            groups.setdefault(thread_key(m.get("subject")), []).append((i, m))
        for key, items in groups.items():
            print(f"## {office2md.md_escape(key)}  ({len(items)} mail(s))")
            print(office2md.md_table([row(i, m) for i, m in items], header=header))
            print()
    else:
        print(office2md.md_table([row(i, m) for i, m in enumerate(mails, start=1)], header=header))
        print()
    if a.bodies:
        for i, m in enumerate(mails, start=1):
            print(f"### {i}. {office2md.md_escape(m.get('subject') or '(no subject)')}  ({(m.get('date') or '?')[:16]}, {one_line(m.get('from'), 40)})")
            print()
            print(office2md.md_para((m.get("body") or "").strip()[: a.bodies]))
            print()
    if deep_rows:
        print("## Attachments in depth")
        print()
        print(office2md.md_table([[r[0], one_line(r[1], 40), one_line(r[2], 50), one_line(r[3], 120)] for r in deep_rows],
                                 header=["mail", "attachment", "file", "what is inside"]))
        print()
    if a.extract:
        print(f"Attachments saved under {a.extract} (prefix = mail number above). Read them with office2md.")
    if failed:
        print()
        print("Could not read:")
        for x in failed:
            print(f"- {x}")
    print()
    print("Open a mail in full: office2md.py <file> ; extract its attachments: --extract-attachments DIR")
    return 0


if __name__ == "__main__":
    sys.exit(main())
