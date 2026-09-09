#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
extract_archive.py - extracts archives found in a tender documents folder
(TenderDataExtractor project), so the LLM can then scan/read the extracted
files with the docx/pdf/xlsx skills (Шаг 3-4 in CLAUDE.md).

Handles the mechanical work that used to require the LLM to reason about
archive contents:
  - supports .zip, .tar, .tar.gz/.tgz, .tar.bz2/.tbz2, .gz, .7z, .rar;
  - extracts each archive into a sibling folder "<archive_name>_extracted"
    next to the archive (or a folder you specify with --dest);
  - guards against zip-slip / path traversal (skips any entry that would
    write outside the destination folder);
  - RECURSIVE BY DEFAULT: if an extracted archive itself contains further
    archives (zip-in-zip, rar-in-zip, etc.), those are extracted too,
    all the way down, without needing to be told to do it again. Depth is
    capped (--max-depth, default 5) and extraction aborts with an error if
    it blows past sane file-count/size limits, as protection against
    decompression-bomb style archives. Pass --no-recursive to only extract
    the top-level archive(s).

Only call this script AFTER the user has confirmed they want the archive(s)
unpacked (see CLAUDE.md Шаг 3 - archives are never auto-extracted without
asking first).

Usage:
    python extract_archive.py --archive "/path/to/file.zip"
    python extract_archive.py --archive "/path/to/a.zip" --archive "/path/to/b.7z"
    python extract_archive.py --scan-dir "/path/to/tender/folder"      # finds & extracts all archives in a folder
    python extract_archive.py --archive "/path/to/file.rar" --password "secret"
    python extract_archive.py --archive "/path/to/file.zip" --no-recursive
    python extract_archive.py --archive "/path/to/file.zip" --max-depth 2

Prints a JSON result to stdout, one entry per top-level archive:
    [
      {"archive": "...", "status": "ok", "dest": "...", "extracted_files": ["...", ...],
       "nested_archives": [{"archive": "...", "status": "ok", "dest": "..."}, ...]},
      {"archive": "...", "status": "error", "error": "..."}
    ]
"""

import argparse
import json
import os
import shutil
import subprocess
import sys
import tarfile
import zipfile

ARCHIVE_EXTENSIONS = (".zip", ".rar", ".7z", ".tar", ".gz", ".tgz", ".bz2", ".tbz2")

TOOLS_DIR = os.path.dirname(os.path.abspath(__file__))

DEFAULT_MAX_DEPTH = 5
DEFAULT_MAX_FILES = 20000
DEFAULT_MAX_TOTAL_BYTES = 2 * 1024 ** 3  # 2 GB


class LimitExceeded(RuntimeError):
    """Raised when recursive extraction blows past the decompression-bomb guard."""


def _ensure_pip_module(module_name, pip_name=None):
    """Import module_name, auto-installing pip_name (or module_name) if missing.

    The sandbox environment is not guaranteed to keep pip packages between
    sessions, so scripts that depend on py7zr/rarfile install themselves on
    first use instead of requiring a manual `pip install` step every time.
    """
    try:
        return __import__(module_name)
    except ImportError:
        subprocess.check_call([
            sys.executable, "-m", "pip", "install", "--break-system-packages", "-q",
            pip_name or module_name,
        ])
        return __import__(module_name)


def find_unrar_tool():
    """Locate a working unrar binary.

    Prefers the copy bundled in tools/bin/unrar (committed to the project,
    so it survives across sandbox sessions) over whatever may or may not be
    present in the current sandbox image.

    Each candidate is test-executed (not just checked for the executable
    bit) before being accepted: tools/bin/unrar is a Linux x86-64 binary
    committed to the repo, so on a different OS/CPU (e.g. macOS, including
    Apple Silicon) it exists and is technically "executable" but fails to
    actually run. In that case we fall through to a system-installed unrar
    instead of hard-failing (on macOS: `brew install unrar`).
    """
    candidates = [
        os.path.join(TOOLS_DIR, "bin", "unrar"),
        os.path.join(TOOLS_DIR, "bin", "unrar.exe"),
        shutil.which("unrar"),
        shutil.which("unrar-nonfree"),
        "/tmp/extracted_pkg/usr/bin/unrar-nonfree",
    ]
    for c in candidates:
        if not c or not os.path.isfile(c) or not os.access(c, os.X_OK):
            continue
        try:
            subprocess.run([c], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=5)
        except (OSError, subprocess.SubprocessError):
            continue  # not runnable on this platform/arch (e.g. bundled Linux binary on macOS)
        return c
    return None


def find_archives(scan_dir):
    found = []
    for root, _dirs, files in os.walk(scan_dir):
        for fname in files:
            if fname.lower().endswith(ARCHIVE_EXTENSIONS):
                found.append(os.path.join(root, fname))
    return found


def safe_dest_path(dest_dir, member_name):
    """Resolve a member path inside dest_dir, rejecting path traversal (zip-slip)."""
    target = os.path.normpath(os.path.join(dest_dir, member_name))
    dest_abs = os.path.abspath(dest_dir)
    target_abs = os.path.abspath(target)
    if not (target_abs == dest_abs or target_abs.startswith(dest_abs + os.sep)):
        return None
    return target_abs


def _fix_zip_name_encoding(name, flag_bits):
    """Python's zipfile decodes non-UTF8-flagged entry names as cp437.
    Many Russian DOS/Windows zip tools actually wrote names in cp866 (or
    cp1251), so the cp437 decode produces mojibake. Undo the cp437 decode
    to get the raw bytes back, then try the encodings actually used by
    Russian tools before falling back to the original (already-decoded)
    name.
    """
    if flag_bits & 0x800:
        return name  # UTF-8 flag set - name is already correct
    try:
        raw = name.encode("cp437")
    except UnicodeEncodeError:
        return name
    for enc in ("cp866", "cp1251"):
        try:
            candidate = raw.decode(enc)
        except UnicodeDecodeError:
            continue
        if candidate and all(ord(c) >= 0x20 for c in candidate):
            return candidate
    return name


def extract_zip(archive_path, dest_dir):
    extracted = []
    with zipfile.ZipFile(archive_path) as zf:
        for info in zf.infolist():
            fixed_name = _fix_zip_name_encoding(info.filename, info.flag_bits)
            target = safe_dest_path(dest_dir, fixed_name)
            if target is None:
                continue  # skip unsafe entry
            if info.is_dir():
                os.makedirs(target, exist_ok=True)
                continue
            os.makedirs(os.path.dirname(target), exist_ok=True)
            with zf.open(info) as src, open(target, "wb") as out:
                out.write(src.read())
            extracted.append(target)
    return extracted


def extract_tar(archive_path, dest_dir):
    extracted = []
    with tarfile.open(archive_path) as tf:
        for member in tf.getmembers():
            target = safe_dest_path(dest_dir, member.name)
            if target is None:
                continue
            if member.isdir():
                os.makedirs(target, exist_ok=True)
                continue
            if not (member.isfile() or member.islnk() or member.issym()):
                continue
            os.makedirs(os.path.dirname(target), exist_ok=True)
            src = tf.extractfile(member)
            if src is None:
                continue
            with open(target, "wb") as out:
                out.write(src.read())
            extracted.append(target)
    return extracted


def extract_gz_single(archive_path, dest_dir):
    import gzip
    base = os.path.basename(archive_path)
    out_name = base[:-3] if base.lower().endswith(".gz") else base + ".out"
    target = os.path.join(dest_dir, out_name)
    with gzip.open(archive_path, "rb") as src, open(target, "wb") as out:
        out.write(src.read())
    return [target]


def extract_7z(archive_path, dest_dir, password=None):
    py7zr = _ensure_pip_module("py7zr")
    extracted = []
    with py7zr.SevenZipFile(archive_path, mode="r", password=password) as zf:
        names = zf.getnames()
        zf.extractall(path=dest_dir)
    for name in names:
        target = safe_dest_path(dest_dir, name)
        if target and os.path.isfile(target):
            extracted.append(target)
    return extracted


def extract_rar(archive_path, dest_dir, password=None):
    rarfile = _ensure_pip_module("rarfile")

    unrar_tool = find_unrar_tool()
    if not unrar_tool:
        raise RuntimeError(
            "Не найден бинарник unrar (ни в tools/bin/unrar, ни в системе). "
            "Положите рабочий unrar в tools/bin/unrar рядом со скриптом."
        )
    rarfile.UNRAR_TOOL = unrar_tool

    extracted = []
    try:
        with rarfile.RarFile(archive_path) as rf:
            for info in rf.infolist():
                target = safe_dest_path(dest_dir, info.filename)
                if target is None or info.is_dir():
                    continue
                os.makedirs(os.path.dirname(target), exist_ok=True)
                with rf.open(info, pwd=password) as src, open(target, "wb") as out:
                    out.write(src.read())
                extracted.append(target)
    except rarfile.RarCannotExec as e:
        raise RuntimeError(
            "Для распаковки .rar на этой машине не найден системный unrar/bsdtar. "
            "Установите unrar (например, apt install unrar) или распакуйте архив вручную."
        ) from e
    return extracted


def extract_one(archive_path, dest_root, password=None):
    lower = archive_path.lower()
    stem = os.path.basename(archive_path)
    for ext in (".tar.gz", ".tar.bz2"):
        if lower.endswith(ext):
            stem = stem[: -len(ext)]
            break
    else:
        stem = os.path.splitext(stem)[0]

    dest_dir = os.path.join(dest_root, stem + "_extracted")
    os.makedirs(dest_dir, exist_ok=True)

    if lower.endswith(".zip"):
        extracted = extract_zip(archive_path, dest_dir)
    elif lower.endswith((".tar", ".tar.gz", ".tgz", ".tar.bz2", ".tbz2")):
        extracted = extract_tar(archive_path, dest_dir)
    elif lower.endswith(".gz"):
        extracted = extract_gz_single(archive_path, dest_dir)
    elif lower.endswith(".7z"):
        extracted = extract_7z(archive_path, dest_dir, password=password)
    elif lower.endswith(".rar"):
        extracted = extract_rar(archive_path, dest_dir, password=password)
    else:
        raise RuntimeError(f"Неподдерживаемый формат архива: {archive_path}")

    return dest_dir, extracted


def extract_recursive(archive_path, dest_root, password, max_depth, stats):
    """Extract archive_path, then recurse into any archives found inside it.

    `stats` accumulates {'files', 'bytes', 'max_files', 'max_bytes'} across
    the whole run (all top-level archives share one budget) so a chain of
    nested archives can't be used as a decompression bomb.

    Returns (dest_dir, all_extracted_files, nested_results) where
    nested_results is a flat list describing every nested archive that was
    (or failed to be) extracted, for reporting purposes.
    """
    dest_dir, extracted = extract_one(archive_path, dest_root, password=password)

    stats["files"] += len(extracted)
    stats["bytes"] += sum(os.path.getsize(p) for p in extracted if os.path.isfile(p))
    if stats["files"] > stats["max_files"] or stats["bytes"] > stats["max_bytes"]:
        raise LimitExceeded(
            f"Превышен лимит распаковки (файлов: {stats['files']}, байт: {stats['bytes']}) — "
            "похоже на архивную бомбу (архив с аномально большим/глубоко вложенным содержимым), "
            "дальнейшая распаковка остановлена."
        )

    all_extracted = list(extracted)
    nested_results = []

    if max_depth > 0:
        for nested_archive in [f for f in extracted if f.lower().endswith(ARCHIVE_EXTENSIONS)]:
            try:
                n_dest, n_extracted, n_nested = extract_recursive(
                    nested_archive, os.path.dirname(nested_archive), password, max_depth - 1, stats
                )
                all_extracted.extend(n_extracted)
                nested_results.append({"archive": nested_archive, "status": "ok", "dest": n_dest})
                nested_results.extend(n_nested)
            except LimitExceeded:
                raise
            except Exception as ne:
                nested_results.append({"archive": nested_archive, "status": "error", "error": str(ne)})
    elif any(f.lower().endswith(ARCHIVE_EXTENSIONS) for f in extracted):
        nested_results.append({
            "archive": None,
            "status": "skipped",
            "error": f"Достигнута максимальная глубина вложенности архивов в {archive_path} — часть архивов внутри не распакована.",
        })

    return dest_dir, all_extracted, nested_results


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--archive", action="append", default=[], help="Путь к архиву (можно указывать несколько раз)")
    parser.add_argument("--scan-dir", help="Папка, в которой нужно найти и распаковать все архивы")
    parser.add_argument("--dest", help="Куда распаковывать (по умолчанию - рядом с архивом, папка <имя>_extracted)")
    parser.add_argument("--password", help="Пароль архива, если требуется")
    parser.add_argument("--no-recursive", action="store_true", help="Не распаковывать архивы, найденные внутри других архивов (по умолчанию распаковка рекурсивная)")
    parser.add_argument("--max-depth", type=int, default=DEFAULT_MAX_DEPTH, help=f"Максимальная глубина вложенных архивов (по умолчанию {DEFAULT_MAX_DEPTH})")
    parser.add_argument("--max-files", type=int, default=DEFAULT_MAX_FILES, help="Лимит суммарного числа извлечённых файлов (защита от архивных бомб)")
    parser.add_argument("--max-total-mb", type=int, default=DEFAULT_MAX_TOTAL_BYTES // (1024 * 1024), help="Лимит суммарного объёма извлечённых данных, МБ")
    args = parser.parse_args()

    archives = list(args.archive)
    if args.scan_dir:
        archives.extend(find_archives(args.scan_dir))
    archives = sorted(set(archives))

    if not archives:
        print(json.dumps({"error": "Архивы не найдены/не указаны"}, ensure_ascii=False))
        sys.exit(1)

    max_depth = 0 if args.no_recursive else args.max_depth
    stats = {"files": 0, "bytes": 0, "max_files": args.max_files, "max_bytes": args.max_total_mb * 1024 * 1024}

    results = []
    for archive_path in archives:
        dest_root = args.dest or os.path.dirname(archive_path)
        try:
            dest_dir, extracted, nested_results = extract_recursive(
                archive_path, dest_root, args.password, max_depth, stats
            )
            results.append({
                "archive": archive_path,
                "status": "ok",
                "dest": dest_dir,
                "extracted_files": extracted,
                "nested_archives": nested_results,
            })
        except Exception as e:
            results.append({"archive": archive_path, "status": "error", "error": str(e)})

    print(json.dumps(results, ensure_ascii=False, indent=2))
    if any(r["status"] == "error" for r in results):
        sys.exit(1)


if __name__ == "__main__":
    main()
