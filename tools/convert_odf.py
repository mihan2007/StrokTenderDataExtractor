#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
convert_odf.py - converts LibreOffice / OpenDocument files (.odt, .ods, .odp)
found in a tender documents folder (TenderDataExtractor project) into their
Microsoft Office equivalents (.docx, .xlsx, .pptx), so the LLM can then read
them with the existing docx/xlsx/pptx skills (Шаг 3-4 in CLAUDE.md) instead
of needing dedicated ODF-reading skills.

Uses the LibreOffice headless CLI (soffice --headless --convert-to ...),
which is already available in the sandbox - no new dependency to install.

Handles the mechanical work that used to require the LLM to reason about it:
  - supports .odt -> .docx, .ods -> .xlsx, .odp -> .pptx
    (also handles the flat-XML variants .fodt/.fods/.fodp and the
    template variants .ott/.ots/.otp the same way);
  - converts each file next to the original by default (same folder), or
    into a folder you specify with --dest;
  - runs each conversion with its own throwaway LibreOffice user profile
    (-env:UserInstallation=...), so parallel/repeated calls in the same
    sandbox session don't collide on a lock file or "soffice is already
    running" error;
  - times out (default 120s per file) instead of hanging forever if
    LibreOffice gets stuck on a malformed file.

Usage:
    python convert_odf.py --file "/path/to/file.odt"
    python convert_odf.py --file "/path/to/a.ods" --file "/path/to/b.odp"
    python convert_odf.py --scan-dir "/path/to/tender/folder"   # finds & converts all ODF files in a folder (recursive)
    python convert_odf.py --file "/path/to/file.odt" --dest "/path/to/AIExtract"
    python convert_odf.py --file "/path/to/file.odt" --timeout 60

Prints a JSON result to stdout, one entry per file:
    [
      {"source": "...", "status": "ok", "converted": "...", "target_format": "docx"},
      {"source": "...", "status": "error", "error": "..."}
    ]
"""

import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile
import uuid

ODF_TARGET_FORMAT = {
    ".odt": "docx",
    ".ott": "docx",
    ".fodt": "docx",
    ".ods": "xlsx",
    ".ots": "xlsx",
    ".fods": "xlsx",
    ".odp": "pptx",
    ".otp": "pptx",
    ".fodp": "pptx",
}

ODF_EXTENSIONS = tuple(ODF_TARGET_FORMAT.keys())

DEFAULT_TIMEOUT_SECONDS = 120


def find_odf_files(scan_dir):
    found = []
    for root, _dirs, files in os.walk(scan_dir):
        for fname in files:
            if fname.lower().endswith(ODF_EXTENSIONS):
                found.append(os.path.join(root, fname))
    return found


def find_soffice():
    candidates = [
        shutil.which("soffice"),
        shutil.which("libreoffice"),
        "/usr/bin/soffice",  # common Linux path
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",  # macOS default install path (not on PATH by default)
    ]
    for c in candidates:
        if c and os.path.isfile(c) and os.access(c, os.X_OK):
            return c
    return None


def convert_one(file_path, dest_dir, soffice_bin, timeout):
    ext = os.path.splitext(file_path)[1].lower()
    target_format = ODF_TARGET_FORMAT.get(ext)
    if not target_format:
        raise RuntimeError(f"Неподдерживаемый формат ODF-документа: {file_path}")

    if not os.path.isfile(file_path):
        raise RuntimeError(f"Файл не найден: {file_path}")

    os.makedirs(dest_dir, exist_ok=True)

    # Throwaway user profile per call: avoids "soffice is already running"
    # lock conflicts when converting several files back to back or in
    # parallel within the same sandbox session.
    profile_dir = os.path.join(tempfile.gettempdir(), f"lo_profile_{uuid.uuid4().hex}")
    profile_uri = "file://" + profile_dir

    cmd = [
        soffice_bin,
        "--headless",
        "--norestore",
        f"-env:UserInstallation={profile_uri}",
        "--convert-to", target_format,
        "--outdir", dest_dir,
        file_path,
    ]

    try:
        proc = subprocess.run(
            cmd, capture_output=True, text=True, timeout=timeout
        )
    except subprocess.TimeoutExpired as e:
        raise RuntimeError(
            f"LibreOffice не ответил за {timeout} сек. при конвертации {file_path} "
            "(возможно, файл повреждён)."
        ) from e
    finally:
        shutil.rmtree(profile_dir, ignore_errors=True)

    stem = os.path.splitext(os.path.basename(file_path))[0]
    expected_out = os.path.join(dest_dir, f"{stem}.{target_format}")

    if proc.returncode != 0 or not os.path.isfile(expected_out):
        detail = (proc.stderr or proc.stdout or "").strip()
        raise RuntimeError(
            f"Не удалось сконвертировать {file_path} в .{target_format}"
            + (f": {detail}" if detail else "")
        )

    return expected_out, target_format


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--file", action="append", default=[], help="Путь к ODF-файлу (можно указывать несколько раз)")
    parser.add_argument("--scan-dir", help="Папка, в которой нужно найти и сконвертировать все ODF-файлы (рекурсивно)")
    parser.add_argument("--dest", help="Куда сохранять результат (по умолчанию - рядом с исходным файлом, в той же папке)")
    parser.add_argument("--timeout", type=int, default=DEFAULT_TIMEOUT_SECONDS, help=f"Таймаут на файл, сек (по умолчанию {DEFAULT_TIMEOUT_SECONDS})")
    args = parser.parse_args()

    files = list(args.file)
    if args.scan_dir:
        files.extend(find_odf_files(args.scan_dir))
    files = sorted(set(files))

    if not files:
        print(json.dumps({"error": "ODF-файлы (.odt/.ods/.odp и т.п.) не найдены/не указаны"}, ensure_ascii=False))
        sys.exit(1)

    soffice_bin = find_soffice()
    if not soffice_bin:
        print(json.dumps({"error": "LibreOffice (soffice) не найден в системе"}, ensure_ascii=False))
        sys.exit(1)

    results = []
    for file_path in files:
        dest_dir = args.dest or os.path.dirname(file_path)
        try:
            converted_path, target_format = convert_one(file_path, dest_dir, soffice_bin, args.timeout)
            results.append({
                "source": file_path,
                "status": "ok",
                "converted": converted_path,
                "target_format": target_format,
            })
        except Exception as e:
            results.append({"source": file_path, "status": "error", "error": str(e)})

    print(json.dumps(results, ensure_ascii=False, indent=2))
    if any(r["status"] == "error" for r in results):
        sys.exit(1)


if __name__ == "__main__":
    main()
