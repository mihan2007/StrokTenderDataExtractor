#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
env_check.py - one-time environment diagnostic (Шаг 0 in CLAUDE.md).

Checks, independent of OS (Windows/macOS/Linux), whether the local tools this
project depends on (unrar, LibreOffice, a few Python packages) are present
and actually usable - not just "a file with that name exists somewhere".
Prints one JSON report to stdout with a status + suggested fix per check,
the fix command tailored to the detected OS.

Meant to run once per machine, right after the project folder is copied
over (see Шаг 0 in CLAUDE.md) - not on every tender. Safe to re-run any
time by hand:

    python3 tools/env_check.py
"""
import importlib
import json
import os
import platform
import sys

TOOLS_DIR = os.path.dirname(os.path.abspath(__file__))
if TOOLS_DIR not in sys.path:
    sys.path.insert(0, TOOLS_DIR)


def _os_name():
    return {"Windows": "windows", "Darwin": "macos", "Linux": "linux"}.get(platform.system(), "other")


def check_python():
    ok = sys.version_info >= (3, 8)
    return {
        "check": "python3",
        "status": "ok" if ok else "missing",
        "detail": sys.version.split()[0],
        "fix": None if ok else "Нужен Python 3.8+: https://www.python.org/downloads/",
    }


def check_unrar():
    try:
        from extract_archive import find_unrar_tool
    except Exception as e:
        return {"check": "unrar", "status": "unknown", "detail": "не удалось проверить: %s" % e, "fix": None}
    tool = find_unrar_tool()
    if tool:
        return {"check": "unrar", "status": "ok", "detail": tool, "fix": None}
    fixes = {
        "macos": "brew install unrar",
        "windows": "choco install unrar  (или скачать unrar.exe с rarlab.com и положить в tools/bin/)",
        "linux": "sudo apt install unrar",
    }
    return {
        "check": "unrar",
        "status": "missing",
        "detail": "рабочий unrar не найден - .rar архивы распаковать не получится (остальные форматы .zip/.7z/.tar не затронуты)",
        "fix": fixes.get(_os_name(), "установите unrar через пакетный менеджер вашей ОС"),
    }


def check_soffice():
    try:
        from convert_odf import find_soffice
    except Exception as e:
        return {"check": "libreoffice", "status": "unknown", "detail": "не удалось проверить: %s" % e, "fix": None}
    tool = find_soffice()
    if tool:
        return {"check": "libreoffice", "status": "ok", "detail": tool, "fix": None}
    fixes = {
        "macos": "brew install --cask libreoffice",
        "windows": "скачать и установить с libreoffice.org/download",
        "linux": "sudo apt install libreoffice",
    }
    return {
        "check": "libreoffice",
        "status": "missing",
        "detail": "LibreOffice (soffice) не найден - конвертация .odt/.ods/.odp работать не будет (остальные форматы не затронуты)",
        "fix": fixes.get(_os_name(), "установите LibreOffice для вашей ОС"),
    }


def check_module(module_name, pip_name=None, note=""):
    try:
        importlib.import_module(module_name)
        return {"check": module_name, "status": "ok", "detail": note, "fix": None}
    except ImportError:
        return {
            "check": module_name,
            "status": "missing",
            "detail": note,
            "fix": "%s -m pip install --break-system-packages %s" % (sys.executable, pip_name or module_name),
        }


def main():
    checks = [
        check_python(),
        check_unrar(),
        check_soffice(),
        check_module("openpyxl", note="используется при работе с Позиции_и_цены.xlsx"),
        check_module("rarfile", note="доустанавливается автоматически при первой распаковке .rar, если отсутствует - можно игнорировать"),
        check_module("py7zr", note="доустанавливается автоматически при первой распаковке .7z, если отсутствует - можно игнорировать"),
    ]
    result = {
        "os": platform.system(),
        "python_version": sys.version.split()[0],
        "checks": checks,
        "all_ok": all(c["status"] == "ok" for c in checks if c["check"] not in ("rarfile", "py7zr")),
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
