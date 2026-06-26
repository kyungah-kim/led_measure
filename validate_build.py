"""
Cross-platform build validation script for LED Panel Analyzer.

Usage:
    python validate_build.py           # full check
    python validate_build.py --fix     # auto-fix where possible (not yet implemented)

Checks:
  1. led_measure.spec  — datas & hiddenimports completeness
  2. build_windows.bat — pip install list vs requirements.txt
  3. Desktop asset files referenced in code actually exist
  4. Web vs Desktop feature gap summary
"""
from __future__ import annotations

import re
import os
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))

PASS  = "\033[32m✔\033[0m"
FAIL  = "\033[31m✘\033[0m"
WARN  = "\033[33m⚠\033[0m"
INFO  = "\033[34mℹ\033[0m"

issues: list[str] = []

def _ok(msg: str)   -> None: print(f"  {PASS}  {msg}")
def _fail(msg: str) -> None: issues.append(msg); print(f"  {FAIL}  {msg}")
def _warn(msg: str) -> None: print(f"  {WARN}  {msg}")
def _info(msg: str) -> None: print(f"  {INFO}  {msg}")


# ── 1. Asset files referenced in desktop code ──────────────────────────────

def check_desktop_assets():
    print("\n[1] Desktop asset files")
    # SVG arrows
    for svg in ("arr_down.svg", "arr_up.svg"):
        p = os.path.join(ROOT, "desktop", svg)
        if os.path.isfile(p):
            _ok(f"desktop/{svg} exists")
        else:
            _fail(f"desktop/{svg} MISSING — referenced in stylesheet")

    # x.png
    p = os.path.join(ROOT, "x.png")
    if os.path.isfile(p):
        _ok("x.png exists")
    else:
        _fail("x.png MISSING — CenterAlignPanel reference image")


# ── 2. led_measure.spec datas completeness ─────────────────────────────────

def check_spec_datas():
    print("\n[2] led_measure.spec — datas")
    spec = open(os.path.join(ROOT, "led_measure.spec"), encoding="utf-8").read()

    required_datas = {
        "x.png":            "'x.png'",
        "arr_down.svg":     "arr_down.svg",
        "arr_up.svg":       "arr_up.svg",
        "web/templates":    "web.*templates",
    }
    for label, pattern in required_datas.items():
        if re.search(pattern, spec):
            _ok(f"spec datas includes {label}")
        else:
            _fail(f"spec datas MISSING: {label}")


# ── 3. led_measure.spec hiddenimports completeness ─────────────────────────

def check_spec_hiddenimports():
    print("\n[3] led_measure.spec — hiddenimports")
    spec = open(os.path.join(ROOT, "led_measure.spec"), encoding="utf-8").read()

    required_imports = [
        ("pptx",                         "python-pptx PPT export"),
        ("lxml",                          "lxml (pptx dependency)"),
        ("serial",                        "pyserial"),
        ("openpyxl",                      "openpyxl Excel export"),
        ("matplotlib",                    "matplotlib chart rendering"),
        ("matplotlib.backends.backend_agg", "matplotlib Agg backend"),
        ("PySide6.QtCharts",              "PySide6 Charts"),
    ]
    for pkg, reason in required_imports:
        if f"'{pkg}'" in spec or f'"{pkg}"' in spec:
            _ok(f"hiddenimports: {pkg}")
        else:
            _fail(f"hiddenimports MISSING: {pkg}  ({reason})")


# ── 4. build_windows.bat pip packages ─────────────────────────────────────

def check_bat_packages():
    print("\n[4] build_windows.bat — pip install packages")
    bat_path = os.path.join(ROOT, "build_windows.bat")
    if not os.path.isfile(bat_path):
        _warn("build_windows.bat not found — skipping")
        return
    bat = open(bat_path, encoding="utf-8", errors="ignore").read()

    required_pkgs = [
        "PySide6", "openpyxl", "pyserial", "python-pptx",
        "matplotlib", "lxml", "pyinstaller",
    ]
    for pkg in required_pkgs:
        if pkg.lower() in bat.lower():
            _ok(f"bat includes: {pkg}")
        else:
            _fail(f"bat MISSING pip install: {pkg}")


# ── 5. requirements.txt vs bat ─────────────────────────────────────────────

def check_requirements():
    print("\n[5] requirements.txt completeness")
    req_path = os.path.join(ROOT, "requirements.txt")
    if not os.path.isfile(req_path):
        _warn("requirements.txt not found — skipping")
        return
    req = open(req_path, encoding="utf-8").read().lower()

    essential = ["pyside6", "openpyxl", "pyserial", "python-pptx", "matplotlib", "lxml"]
    for pkg in essential:
        if pkg in req:
            _ok(f"requirements.txt: {pkg}")
        else:
            _warn(f"requirements.txt may be missing: {pkg}")


# ── 6. Web vs Desktop feature gap ─────────────────────────────────────────

def check_web_gap():
    print("\n[6] Web vs Desktop feature gap (informational)")
    routes_path = os.path.join(ROOT, "web", "routes.py")
    if not os.path.isfile(routes_path):
        _warn("web/routes.py not found — skipping")
        return
    routes = open(routes_path, encoding="utf-8").read()

    # Desktop panels that should have corresponding web endpoints
    web_checks = {
        "LumSwing  (/run/<seq_name>)": r'route.*"/run/',
        "AutoAll   (/auto_all/run)":   r'route.*"/auto_all',
        "PPT export(/export/ppt)":     r'route.*"/export/ppt',
        "Progress  (/progress)":       r'route.*"/progress',
    }
    for label, pattern in web_checks.items():
        if re.search(pattern, routes):
            _ok(f"Web route exists: {label}")
        else:
            _fail(f"Web route MISSING: {label}")

    # Known gaps — informational only
    _info("Known Web gaps (not critical — desktop-only features):")
    _info("  • LgTvPanel (serial terminal) — no web equivalent")
    _info("  • ModulePanel / GammaSubPanel / ColorSubPanel — no web equivalent")
    _info("  • Center align x.png preview — static image, not API-dependent")


# ── 7. Frozen path correctness ─────────────────────────────────────────────

def check_frozen_paths():
    print("\n[7] PyInstaller frozen path handling")
    mw = open(os.path.join(ROOT, "desktop", "main_window.py"), encoding="utf-8").read()
    rd = open(os.path.join(ROOT, "run_desktop.py"), encoding="utf-8").read()

    if "_MEIPASS" in rd and "MATPLOTLIBDATA" in rd:
        _ok("run_desktop.py sets MATPLOTLIBDATA for frozen build")
    else:
        _fail("run_desktop.py missing frozen MATPLOTLIBDATA setup")

    if "_MEIPASS" in mw and "x.png" in mw:
        _ok("CenterAlignPanel uses _MEIPASS fallback for x.png")
    else:
        _fail("CenterAlignPanel missing _MEIPASS fallback for x.png")

    if "_dark_style()" in mw:
        _ok("Stylesheet uses _dark_style() with runtime path for SVG arrows")
    else:
        _fail("Stylesheet not using _dark_style() — SVG arrows may not load in frozen build")


# ── Main ───────────────────────────────────────────────────────────────────

def main():
    print("=" * 60)
    print("  LED Panel Analyzer — Cross-Platform Build Validator")
    print("=" * 60)

    check_desktop_assets()
    check_spec_datas()
    check_spec_hiddenimports()
    check_bat_packages()
    check_requirements()
    check_web_gap()
    check_frozen_paths()

    print("\n" + "=" * 60)
    if issues:
        print(f"  {FAIL}  {len(issues)} issue(s) found:\n")
        for i, msg in enumerate(issues, 1):
            print(f"     {i}. {msg}")
        sys.exit(1)
    else:
        print(f"  {PASS}  All checks passed — build should be consistent across platforms")
    print("=" * 60)


if __name__ == "__main__":
    main()
