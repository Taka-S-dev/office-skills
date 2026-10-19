#!/usr/bin/env python3
"""Unpack zip and 7z archives safely, with Japanese file names intact.

Usage:
    unpack.py 添付.zip --out ./extracted                extract (nested archives too, one level)
    unpack.py 添付.7z --out ./extracted --password xxxx
    unpack.py 添付.zip --list                           names and sizes only, nothing written

Archives sent by mail arrive with three problems: file names from Windows are cp932 and come
out as mojibake, an entry can be named ../../something and land outside the target folder,
and a small file can expand to gigabytes. This handles all three and then prints what landed
where, so the files can be read with office2md.

Exit code 0 = extracted, 1 = nothing extracted or an error, 2 = wrong or missing password.
"""

from __future__ import annotations

import argparse
import sys
import zipfile
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")  # Windows console defaults to cp932
except AttributeError:
    pass

ARCHIVE_EXTS = (".zip", ".7z")
MAX_TOTAL = 2 * 1024 ** 3        # 2 GB uncompressed per archive
MAX_RATIO = 200                  # uncompressed / compressed; above this it is a bomb, not a document
MAX_FILES = 5000


def fix_zip_name(info: zipfile.ZipInfo) -> str:
    """zipfile decodes non-UTF-8 names as cp437; Windows zips are cp932 (Shift_JIS)."""
    name = info.filename
    if info.flag_bits & 0x800:      # UTF-8 flag set: already right
        return name
    try:
        return name.encode("cp437").decode("cp932")
    except (UnicodeEncodeError, UnicodeDecodeError):
        return name


def safe_target(out: Path, name: str) -> Path | None:
    """Target path inside out, or None when the entry would escape it."""
    name = name.replace("\\", "/")
    parts = [p for p in name.split("/") if p not in ("", ".")]
    if not parts or any(p == ".." for p in parts) or Path(name).is_absolute() or (len(parts[0]) > 1 and parts[0][1] == ":"):
        return None
    target = out.joinpath(*parts)
    try:
        target.resolve().relative_to(out.resolve())
    except ValueError:
        return None
    return target


def list_zip(path: Path, password: str | None):
    with zipfile.ZipFile(path) as z:
        if password:
            z.setpassword(password.encode("utf-8"))
        for info in z.infolist():
            yield fix_zip_name(info), info.file_size, info.compress_size, info.is_dir()


def extract_zip(path: Path, out: Path, password: str | None) -> list[tuple[str, Path]]:
    done = []
    with zipfile.ZipFile(path) as z:
        if password:
            z.setpassword(password.encode("utf-8"))
        for info in z.infolist():
            name = fix_zip_name(info)
            if info.is_dir():
                continue
            target = safe_target(out, name)
            if target is None:
                print(f"  skipped (unsafe path): {name}", file=sys.stderr)
                continue
            target.parent.mkdir(parents=True, exist_ok=True)
            try:
                with z.open(info) as src, open(target, "wb") as dst:
                    remaining = MAX_TOTAL
                    while True:
                        chunk = src.read(1024 * 1024)
                        if not chunk:
                            break
                        remaining -= len(chunk)
                        if remaining < 0:
                            raise RuntimeError("archive expands beyond the size limit")
                        dst.write(chunk)
            except RuntimeError as e:
                if "password" in str(e).lower() or "Bad password" in str(e):
                    raise PermissionError(str(e))
                raise
            done.append((name, target))
    return done


def list_7z(path: Path, password: str | None):
    import py7zr
    with py7zr.SevenZipFile(path, mode="r", password=password) as z:
        for f in z.list():
            yield f.filename, f.uncompressed or 0, f.compressed or 0, f.is_directory


def extract_7z(path: Path, out: Path, password: str | None) -> list[tuple[str, Path]]:
    import py7zr
    done = []
    with py7zr.SevenZipFile(path, mode="r", password=password) as z:
        encrypted = z.needs_password()
        if encrypted and not password:
            raise PermissionError("archive is encrypted; pass --password")
        names = [f.filename for f in z.list() if not f.is_directory]
        safe = [n for n in names if safe_target(out, n) is not None]
        for n in names:
            if n not in safe:
                print(f"  skipped (unsafe path): {n}", file=sys.stderr)
        if safe:
            try:
                z.extract(path=out, targets=safe)
            except Exception as e:
                # a wrong password surfaces as corrupt LZMA data, not as a password error
                if encrypted:
                    raise PermissionError(f"wrong password ({type(e).__name__})")
                raise
        for n in safe:
            t = safe_target(out, n)
            if t and t.exists():
                done.append((n, t))
    return done


def check_bomb(path: Path, entries) -> None:
    total = comp = 0
    n = 0
    for _, size, csize, is_dir in entries:
        if is_dir:
            continue
        n += 1
        total += size
        comp += csize
    if n > MAX_FILES:
        raise RuntimeError(f"{n} entries; more than {MAX_FILES} is not a document archive")
    if total > MAX_TOTAL:
        raise RuntimeError(f"would expand to {total / 1024 ** 3:.1f} GB; limit is {MAX_TOTAL / 1024 ** 3:.0f} GB")
    if comp and total / comp > MAX_RATIO and total > 50 * 1024 ** 2:
        raise RuntimeError(f"compression ratio {total / comp:.0f}:1 looks like an archive bomb")


def unpack(path: Path, out: Path, password: str | None, depth: int = 1) -> list[tuple[str, Path]]:
    ext = path.suffix.lower()
    try:
        if ext == ".zip":
            check_bomb(path, list_zip(path, password))
            done = extract_zip(path, out, password)
        elif ext == ".7z":
            check_bomb(path, list_7z(path, password))
            done = extract_7z(path, out, password)
        else:
            raise RuntimeError(f"unsupported archive type {ext}")
    except (RuntimeError, zipfile.BadZipFile) as e:
        msg = str(e)
        if "password" in msg.lower():
            raise PermissionError(msg)
        raise
    except Exception as e:  # py7zr raises its own classes for bad passwords
        if "password" in type(e).__name__.lower() or "password" in str(e).lower():
            raise PermissionError(str(e))
        raise
    if depth > 0:
        for name, target in list(done):
            if target.suffix.lower() in ARCHIVE_EXTS:
                sub = target.with_suffix("")
                try:
                    inner = unpack(target, sub, password, depth - 1)
                    done.extend((f"{name}/{n}", t) for n, t in inner)
                except Exception as e:
                    print(f"  nested archive not unpacked: {name} ({e})", file=sys.stderr)
    return done


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("archive", type=Path)
    ap.add_argument("--out", type=Path, help="target folder (default: next to the archive, same name)")
    ap.add_argument("--password")
    ap.add_argument("--list", action="store_true")
    ap.add_argument("--no-nested", action="store_true", help="leave archives inside the archive packed")
    a = ap.parse_args(argv)
    if not a.archive.exists():
        sys.exit(f"[unpack] file not found: {a.archive}")
    ext = a.archive.suffix.lower()
    if ext not in ARCHIVE_EXTS:
        sys.exit(f"[unpack] unsupported: {ext} (zip and 7z only)")
    try:
        if a.list:
            entries = list(list_zip(a.archive, a.password) if ext == ".zip" else list_7z(a.archive, a.password))
            print(f"# {a.archive.name}: {sum(1 for e in entries if not e[3])} file(s)")
            for name, size, _, is_dir in entries:
                if not is_dir:
                    print(f"- {name}  ({size} bytes)")
            return 0
        out = a.out or a.archive.with_suffix("")
        out.mkdir(parents=True, exist_ok=True)
        done = unpack(a.archive, out, a.password, depth=0 if a.no_nested else 1)
    except PermissionError as e:
        print(f"[unpack] password required or wrong: {e}", file=sys.stderr)
        return 2
    except Exception as e:
        print(f"[unpack] {type(e).__name__}: {e}", file=sys.stderr)
        return 1
    print(f"# {a.archive.name} → {out}: {len(done)} file(s)")
    for name, target in done:
        print(f"- {name}  →  {target}")
    if not done:
        return 1
    print()
    print("Read a file: office2md.py <path>   Search all: office_find.py PATTERN <folder>")
    return 0


if __name__ == "__main__":
    sys.exit(main())
