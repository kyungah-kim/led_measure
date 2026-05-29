"""
Mock 장비로 전체 측정 시퀀스를 실행하고 Excel 출력을 검증하는 스크립트.

실행:
    python test_output.py

출력 파일은 현재 디렉터리의 output/ 폴더에 저장됩니다.
"""
from __future__ import annotations

import os
import sys
import time

sys.path.insert(0, os.path.dirname(__file__))

from core.engine import MeasurementEngine
from core.equipment.mock import MockGenerator, MockMeter
from core.export import ExcelExporter

# ── 출력 폴더 ─────────────────────────────────────────────────────────────────
OUT_DIR = os.path.join(os.path.dirname(__file__), "output")
os.makedirs(OUT_DIR, exist_ok=True)

BRAND = "ex)LG"
MODEL = "ex)XXXXXXX"


# ── 색상 출력 헬퍼 ────────────────────────────────────────────────────────────
def _cyan(s: str)    -> str: return f"\033[96m{s}\033[0m"
def _green(s: str)   -> str: return f"\033[92m{s}\033[0m"
def _yellow(s: str)  -> str: return f"\033[93m{s}\033[0m"
def _bold(s: str)    -> str: return f"\033[1m{s}\033[0m"

def _print_result(name: str, path: str) -> None:
    size_kb = os.path.getsize(path) / 1024
    print(f"  {_green('✔')} {name:<20} → {os.path.basename(path)}  ({size_kb:.1f} KB)")


# ── Mock 엔진 초기화 ──────────────────────────────────────────────────────────
def _make_engine() -> MeasurementEngine:
    engine = MeasurementEngine(brand=BRAND, model_name=MODEL)
    m = MockMeter();    m.connect("MOCK")
    g = MockGenerator(); g.connect("MOCK")
    engine.meter     = m
    engine.generator = g
    return engine


# ── 시퀀스별 실행 함수 ────────────────────────────────────────────────────────

def run_lum_swing(engine: MeasurementEngine, exporter: ExcelExporter) -> str:
    """케이스 A(SDR) 30샘플 + 케이스 B(HDR) 30샘플."""
    results_by_case: dict = {}
    for case, is_hdr in [("A", False), ("B", True)]:
        collected = []
        def _cb(r, _c=case, _col=collected):
            _col.append(r)
            pct = len(_col) / 30 * 100
            bar = "█" * (len(_col) * 20 // 30)
            print(f"\r    Case {_c} [{bar:<20}] {pct:5.1f}%", end="", flush=True)

        engine.run_sequence("lum_swing", case=case, is_hdr=is_hdr,
                            callback=_cb, sample_count=30)
        print()  # newline
        results_by_case[case] = collected

    case_str = "-".join(f"Case{k}" for k in sorted(results_by_case.keys()))
    path = exporter.export_lum_swing(
        results_by_case, BRAND, MODEL,
        file_path=os.path.join(OUT_DIR, f"{MODEL}_LumSwing_{case_str}.xlsx")
    )
    return path


def run_lum_loading(engine: MeasurementEngine, exporter: ExcelExporter) -> str:
    """10단계 버전, 케이스 A(SDR) + B(HDR)."""
    results_by_case: dict = {}
    for case, is_hdr in [("A", False), ("B", True)]:
        step_results: dict = {}
        total_steps = 10

        def _cb(step_idx, apl, step_res, _case=case, _sr=step_results):
            _sr[int(apl)] = step_res
            bar = "█" * ((step_idx + 1) * 20 // total_steps)
            print(f"\r    Case {_case} [{bar:<20}] APL {apl:3.0f}%  "
                  f"Lv={step_res[0].Lv:7.3f} cd/m²", end="", flush=True)

        engine.run_sequence("lum_loading", version="10", case=case,
                            is_hdr=is_hdr, use_avg=True, cooling_enabled=False,
                            callback=_cb)
        print()
        results_by_case[case] = step_results

    path = exporter.export_lum_loading(
        results_by_case, BRAND, MODEL, use_avg=True,
        file_path=os.path.join(OUT_DIR, f"{MODEL}_LumLoading.xlsx")
    )
    return path


def run_gamut(engine: MeasurementEngine, exporter: ExcelExporter) -> str:
    """R→G→B→W→BK 색재현율 측정."""
    results: dict = {}
    color_icons = {"red": "🔴", "green": "🟢", "blue": "🔵",
                   "white": "⬜", "black": "⬛"}

    def _cb(color, r):
        results[color] = r
        icon = color_icons.get(color, "  ")
        print(f"    {icon} {color:<8}  "
              f"Lv={r.Lv:7.3f}  x={r.x:.4f}  y={r.y:.4f}  "
              f"u'={r.u_prime:.4f}  v'={r.v_prime:.4f}")

    engine.run_sequence("gamut", callback=_cb)

    path = exporter.export_gamut(
        results, BRAND, MODEL,
        file_path=os.path.join(OUT_DIR, f"{MODEL}_Gamut.xlsx")
    )
    return path


def run_contrast(engine: MeasurementEngine, exporter: ExcelExporter) -> str:
    """White Raster + Black Window 명암비 측정."""
    results: dict = {}

    def _cb(win_size, r):
        results[win_size] = r
        print(f"    Window {win_size:5.1f}%  Lv={r.Lv:.4f} cd/m²")

    engine.run_sequence("contrast", callback=_cb)

    # 명암비 계산 출력
    lv_white = results.get(0.0)
    lv_100   = results.get(100.0)
    if lv_white and lv_100 and lv_100.Lv > 0:
        cr = lv_white.Lv / lv_100.Lv
        print(f"    {_yellow(f'Contrast Ratio (0% vs 100% window): {cr:,.1f}:1')}")

    path = exporter.export_contrast(
        results, BRAND, MODEL,
        file_path=os.path.join(OUT_DIR, f"{MODEL}_ContrastRatio.xlsx")
    )
    return path


# ── 메인 ──────────────────────────────────────────────────────────────────────

def main() -> None:
    print(_bold(f"\n{'='*60}"))
    print(_bold(f"  LED Measure — Mock Output 검증"))
    print(_bold(f"  Brand: {BRAND}  /  Model: {MODEL}"))
    print(_bold(f"{'='*60}\n"))

    engine   = _make_engine()
    exporter = ExcelExporter()
    paths    = []

    # 1. 휘도 스윙
    print(_cyan("▶ [1/4] 휘도 스윙 (Luminance Swing) — Case A(SDR) · B(HDR)"))
    t0 = time.time()
    paths.append(("Luminance Swing", run_lum_swing(engine, exporter)))
    print(f"    완료 ({time.time()-t0:.1f}s)\n")

    # 2. APL 로딩
    print(_cyan("▶ [2/4] APL 로딩 (Luminance Loading) — 10단계 · Case A·B"))
    t0 = time.time()
    paths.append(("Luminance Loading", run_lum_loading(engine, exporter)))
    print(f"    완료 ({time.time()-t0:.1f}s)\n")

    # 3. 색재현율
    print(_cyan("▶ [3/4] 색재현율 (Gamut) — R·G·B·W·BK"))
    t0 = time.time()
    paths.append(("Gamut", run_gamut(engine, exporter)))
    print(f"    완료 ({time.time()-t0:.1f}s)\n")

    # 4. 명암비
    print(_cyan("▶ [4/4] 명암비 (Contrast Ratio) — 100·50·20·14.1·0%"))
    t0 = time.time()
    paths.append(("Contrast Ratio", run_contrast(engine, exporter)))
    print(f"    완료 ({time.time()-t0:.1f}s)\n")

    # 결과 요약
    print(_bold(f"{'─'*60}"))
    print(_bold("  생성된 Excel 파일:"))
    for name, path in paths:
        _print_result(name, path)
    print(_bold(f"{'─'*60}"))
    print(f"\n  저장 폴더: {_yellow(OUT_DIR)}\n")


if __name__ == "__main__":
    main()
