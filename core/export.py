"""Excel export for all measurement sequences (openpyxl).

NOTE: openpyxl writes .xlsx natively.  True .xlsm (VBA macros) requires an
existing macro-enabled template opened with keep_vba=True.  For new files we
use .xlsx — Excel opens these identically for pure data + charts.
"""
from __future__ import annotations

import os
import statistics
from datetime import datetime
from typing import Any, Dict, List

import openpyxl
from openpyxl.chart import LineChart, Reference, ScatterChart
from openpyxl.styles import (
    Alignment, Border, Font, PatternFill, Side
)
from openpyxl.utils import get_column_letter

from .equipment.base import MeasureResult
from .gamut_utils import calc_gamut_stats, DCI_P3_UV, BT2020_UV

# ── 브랜드 색상 ───────────────────────────────────────────────────────────────
_BRAND_HEX = {"samsung": "0070C0", "lg": "FF0000", "sony": "00B050", "philips": "FF8800"}


def _brand_hex(brand: str) -> str:
    return _BRAND_HEX.get(brand.lower().strip(), "FF8800")


# ── 스타일 상수 ──────────────────────────────────────────────────────────────
_BLUE_FILL   = PatternFill("solid", fgColor="1F4E79")
_GRAY_FILL   = PatternFill("solid", fgColor="D6DCE4")
_GREEN_FILL  = PatternFill("solid", fgColor="E2EFDA")
_YELLOW_FILL = PatternFill("solid", fgColor="FFF2CC")
_RED_FILL    = PatternFill("solid", fgColor="FCE4D6")

_WHITE_FONT  = Font(color="FFFFFF", bold=True, name="Calibri", size=10)
_BOLD_FONT   = Font(bold=True, name="Calibri", size=10)
_BASE_FONT   = Font(name="Calibri", size=10)

_THIN = Side(style="thin", color="BFBFBF")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_CENTER = Alignment(horizontal="center", vertical="center")
_RIGHT  = Alignment(horizontal="right",  vertical="center")

# 숫자 형식
_FMT_LV    = '0.000'
_FMT_CHROM = '0.0000'
_FMT_XYZ   = '0.000'
_FMT_INT   = '0'
_FMT_RATIO = '0.0'

# 측정값 컬럼 정의: (헤더, 속성경로, 숫자포맷, 열너비)
_MEAS_COLS: list[tuple[str, str, str, int]] = [
    ("Time (s)",        "timestamp_ms",           _FMT_INT,   10),
    ("Lv (cd/m²)",      "Lv",                     _FMT_LV,    12),
    ("x",               "x",                      _FMT_CHROM, 10),
    ("y",               "y",                      _FMT_CHROM, 10),
    ("u'",              "u_prime",                _FMT_CHROM, 10),
    ("v'",              "v_prime",                _FMT_CHROM, 10),
    ("X",               "X",                      _FMT_XYZ,   10),
    ("Y",               "Y",                      _FMT_XYZ,   10),
    ("Z",               "Z",                      _FMT_XYZ,   10),
    ("CCT (K)",         "cct",                    _FMT_INT,   10),
    ("Duv",             "duv",                    "0.00000",  10),
    ("Pattern",         "pattern_info.type",      "@",         12),
    ("APL (%)",         "pattern_info.apl_pct",   _FMT_RATIO, 10),
    ("W (%)",           "pattern_info.width_pct", _FMT_RATIO, 10),
    ("H (%)",           "pattern_info.height_pct",_FMT_RATIO, 10),
    ("Color",           "pattern_info.color",     "@",         12),
    ("SDR/HDR",         "pattern_info.is_hdr",    "@",         10),
]


# ── 헬퍼 ─────────────────────────────────────────────────────────────────────

def _get_attr(obj: object, path: str) -> object:
    for part in path.split("."):
        obj = getattr(obj, part)
    if isinstance(obj, bool):
        return "HDR" if obj else "SDR"
    return obj


def _write_header_row(ws, row: int, labels: list[str],
                      fill=_BLUE_FILL, font=_WHITE_FONT) -> None:
    for ci, label in enumerate(labels, 1):
        c = ws.cell(row=row, column=ci, value=label)
        c.fill = fill
        c.font = font
        c.alignment = _CENTER
        c.border = _BORDER


def _write_meas_row(ws, row: int, r: MeasureResult,
                    col_offset: int = 0, seq: int | None = None) -> None:
    for ci, (_, path, fmt, _) in enumerate(_MEAS_COLS, 1 + col_offset):
        val = seq if (seq is not None and path == "timestamp_ms") else _get_attr(r, path)
        c = ws.cell(row=row, column=ci, value=val)
        c.number_format = fmt
        c.alignment = _CENTER
        c.border = _BORDER
        c.font = _BASE_FONT


def _set_col_widths(ws, widths: list[int], col_offset: int = 0) -> None:
    for ci, w in enumerate(widths, 1 + col_offset):
        ws.column_dimensions[get_column_letter(ci)].width = w


def _freeze(ws, cell: str = "B2") -> None:
    ws.freeze_panes = cell


def _file_prefix(brand: str, model: str) -> str:
    safe = lambda s: "".join(c for c in s if c.isalnum() or c in "._- ").strip()
    return f"{safe(brand)}_{safe(model)}"


def _default_path(brand: str, model: str, suffix: str) -> str:
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    fname = f"{_file_prefix(brand, model)}_{suffix}_{ts}.xlsx"
    return os.path.join(os.getcwd(), fname)


def _add_info_sheet(wb, brand: str, model: str, sequence: str) -> None:
    # openpyxl.Workbook() always creates a default empty "Sheet" — remove it
    for default_name in ("Sheet", "Sheet1"):
        if default_name in wb.sheetnames:
            wb.remove(wb[default_name])
    ws = wb.create_sheet("Info", 0)
    rows = [
        ("Brand",     brand),
        ("Model",     model),
        ("Sequence",  sequence),
        ("Export",    datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        ("Software",  "LED Measure v0.1"),
    ]
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 30
    for ri, (k, v) in enumerate(rows, 1):
        ca = ws.cell(row=ri, column=1, value=k)
        cb = ws.cell(row=ri, column=2, value=v)
        ca.font = _BOLD_FONT
        cb.font = _BASE_FONT
        ca.border = _BORDER
        cb.border = _BORDER


# ── ExcelExporter ─────────────────────────────────────────────────────────────

class ExcelExporter:
    """Creates formatted Excel workbooks from MeasureResult data."""

    # ── 1. Luminance Swing ────────────────────────────────────────────────────

    def export_lum_swing(
        self,
        results_by_case: Dict[str, List[MeasureResult]],
        brand: str,
        model: str,
        file_path: str | None = None,
    ) -> str:
        """단일 시트 피벗: 행=측정#, 열=모드(SDR_Vivid 등), 값=Lv."""
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        _add_info_sheet(wb, brand, model, "Luminance Swing")

        ws = wb.create_sheet("LumSwing")
        keys = sorted(results_by_case.keys())
        max_n = max((len(v) for v in results_by_case.values()), default=0)

        # 헤더: # | SDR_Vivid | SDR_Standard | ...
        _write_header_row(ws, 1, ["#"] + keys)
        ws.column_dimensions["A"].width = 6
        for ci in range(2, len(keys) + 2):
            ws.column_dimensions[get_column_letter(ci)].width = 14
        _freeze(ws, "B2")

        # 데이터 행
        for i in range(max_n):
            ri = i + 2
            idx_cell = ws.cell(row=ri, column=1, value=i + 1)
            idx_cell.alignment = _CENTER
            idx_cell.border = _BORDER
            idx_cell.font = _BASE_FONT
            for ci, key in enumerate(keys, 2):
                row_list = results_by_case.get(key, [])
                val = row_list[i].Lv if i < len(row_list) else None
                c = ws.cell(row=ri, column=ci, value=val)
                c.number_format = _FMT_LV
                c.alignment = _CENTER
                c.border = _BORDER
                c.font = _BASE_FONT

        # 통계 블록
        if max_n > 0:
            stat_row = max_n + 3
            ws.cell(row=stat_row, column=1, value="Statistics").font = _BOLD_FONT
            stat_row += 1
            for label in ["Count", "Mean Lv", "Max Lv", "Min Lv", "Std Dev"]:
                ws.cell(row=stat_row, column=1, value=label).font = _BOLD_FONT
                for ci, key in enumerate(keys, 2):
                    lvs = [r.Lv for r in results_by_case.get(key, [])]
                    if not lvs:
                        continue
                    if   label == "Count":   val = len(lvs)
                    elif label == "Mean Lv": val = round(statistics.mean(lvs), 3)
                    elif label == "Max Lv":  val = round(max(lvs), 3)
                    elif label == "Min Lv":  val = round(min(lvs), 3)
                    else:                    val = round(statistics.stdev(lvs), 4) if len(lvs) > 1 else 0
                    c = ws.cell(row=stat_row, column=ci, value=val)
                    c.number_format = _FMT_LV
                    c.alignment = _CENTER
                    c.border = _BORDER
                stat_row += 1

        if max_n >= 2:
            self._add_swing_pivot_chart(ws, keys, max_n)

        path = file_path or _default_path(brand, model, "LumSwing")
        wb.save(path)
        return path

    def _add_swing_pivot_chart(self, ws, keys: list, data_rows: int) -> None:
        if data_rows < 2:
            return
        chart = LineChart()
        chart.title = "Luminance Swing"
        chart.style = 10
        chart.y_axis.title = "Lv (cd/m²)"
        chart.x_axis.title = "측정 #"
        chart.width  = 26
        chart.height = 14

        _COLORS = ["1F4E79", "C55A11", "375623", "7030A0", "FF0000", "00B0F0"]
        for i, key in enumerate(keys):
            ref = Reference(ws, min_col=i + 2, min_row=1, max_row=data_rows + 1)
            chart.add_data(ref, titles_from_data=True)
            chart.series[i].graphicalProperties.line.solidFill = _COLORS[i % len(_COLORS)]
            chart.series[i].graphicalProperties.line.width = 15000

        cat_ref = Reference(ws, min_col=1, min_row=2, max_row=data_rows + 1)
        chart.set_categories(cat_ref)
        ws.add_chart(chart, f"A{data_rows + 6}")

    # ── 2. Luminance Loading ──────────────────────────────────────────────────

    def export_lum_loading(
        self,
        results_by_case: Dict[str, Dict[int, List[MeasureResult]]],
        brand: str,
        model: str,
        use_avg: bool = True,   # kept for API compatibility; Summary shows both Avg & Max
        file_path: str | None = None,
        brand_name: str = "",
    ) -> str:
        """Summary 시트(APL × 케이스) + 케이스별 Raw 시트."""
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        _add_info_sheet(wb, brand, model, "Luminance Loading")

        cases   = sorted(results_by_case.keys())
        all_apls: list[int] = sorted(
            {apl for cd in results_by_case.values() for apl in cd}
        )

        # ── Summary 시트 ──────────────────────────────────────────────────────
        ws_sum = wb.create_sheet("Summary")

        # 헤더: APL | Case A Avg | Case A Max | Case A Min | Case B ...
        hdr = ["APL (%)"]
        for case in cases:
            hdr += [f"Case {case} Avg Lv", f"Case {case} Max Lv",
                    f"Case {case} Min Lv", f"Case {case} StdDev"]
        _write_header_row(ws_sum, 1, hdr)
        ws_sum.column_dimensions["A"].width = 10
        for ci in range(2, len(hdr) + 1):
            ws_sum.column_dimensions[get_column_letter(ci)].width = 14

        for ri, apl in enumerate(all_apls, 2):
            ws_sum.cell(row=ri, column=1, value=apl).alignment = _CENTER
            ci = 2
            for case in cases:
                step_results = results_by_case.get(case, {}).get(apl, [])
                lvs = [r.Lv for r in step_results]
                avg = round(statistics.mean(lvs), 3)   if lvs else None
                mx  = round(max(lvs), 3)               if lvs else None
                mn  = round(min(lvs), 3)               if lvs else None
                std = round(statistics.stdev(lvs), 4)  if len(lvs) > 1 else (0 if lvs else None)
                for val, fmt in [(avg, _FMT_LV), (mx, _FMT_LV), (mn, _FMT_LV), (std, _FMT_LV)]:
                    c = ws_sum.cell(row=ri, column=ci, value=val)
                    c.number_format = fmt
                    c.alignment = _CENTER
                    c.border = _BORDER
                    ci += 1

        _freeze(ws_sum, "B2")

        # Summary 라인 차트 (Avg Lv per case)
        self._add_apl_chart(ws_sum, all_apls, cases, len(all_apls), brand=brand_name)

        # ── Raw 시트 ─────────────────────────────────────────────────────────
        meas_headers = [h for h, *_ in _MEAS_COLS]
        meas_widths  = [w for *_, w in _MEAS_COLS]

        for case, case_data in sorted(results_by_case.items()):
            ws_raw = wb.create_sheet(title=f"Raw_{case}")
            _write_header_row(ws_raw, 1, ["APL (%)"] + meas_headers,
                              fill=PatternFill("solid", fgColor="375623"),
                              font=_WHITE_FONT)
            ws_raw.column_dimensions["A"].width = 10
            _set_col_widths(ws_raw, meas_widths, col_offset=1)
            _freeze(ws_raw, "B2")

            ri = 2
            for apl in sorted(case_data.keys()):
                fill = _GRAY_FILL if apl % 20 < 10 else None
                for r in case_data[apl]:
                    c = ws_raw.cell(row=ri, column=1, value=apl)
                    c.alignment = _CENTER
                    c.border = _BORDER
                    if fill:
                        c.fill = fill
                    _write_meas_row(ws_raw, ri, r, col_offset=1, seq=ri - 1)
                    if fill:
                        for ci2 in range(2, len(_MEAS_COLS) + 2):
                            ws_raw.cell(row=ri, column=ci2).fill = fill
                    ri += 1

        path = file_path or _default_path(brand, model, "LumLoading")
        wb.save(path)
        return path

    def _add_apl_chart(self, ws, _all_apls: list, cases: list, n_apl: int, brand: str = "") -> None:
        chart = LineChart()
        chart.title = "APL vs Lv (Average)"
        chart.style = 10
        chart.y_axis.title = "Lv (cd/m²)"
        chart.x_axis.title = "APL (%)"
        chart.width  = 26
        chart.height = 14

        fallback_colors = ["1F4E79", "C55A11", "375623"]
        for i, case in enumerate(cases):
            col = 2 + i * 4
            ref = Reference(ws, min_col=col, min_row=1, max_row=n_apl + 1)
            chart.add_data(ref, titles_from_data=True)
            color = _brand_hex(brand) if i == 0 and brand else (
                fallback_colors[i] if i < len(fallback_colors) else "888888"
            )
            chart.series[i].graphicalProperties.line.solidFill = color

        apl_ref = Reference(ws, min_col=1, min_row=2, max_row=n_apl + 1)
        chart.set_categories(apl_ref)
        ws.add_chart(chart, f"A{n_apl + 4}")

    # ── 3. Gamut ──────────────────────────────────────────────────────────────

    def export_gamut(
        self,
        results: Dict[str, MeasureResult],
        brand: str,
        model: str,
        file_path: str | None = None,
    ) -> str:
        """컬러별 행 + 배경색, 전체 색채측정값."""
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        _add_info_sheet(wb, brand, model, "Gamut")

        ws = wb.create_sheet("Gamut")

        color_order = ["red", "green", "blue", "white", "black"]
        color_fills = {
            "red":   PatternFill("solid", fgColor="FCE4D6"),
            "green": PatternFill("solid", fgColor="E2EFDA"),
            "blue":  PatternFill("solid", fgColor="DAE8FC"),
            "white": PatternFill("solid", fgColor="F2F2F2"),
            "black": PatternFill("solid", fgColor="D9D9D9"),
        }
        color_label = {
            "red": "Red", "green": "Green", "blue": "Blue",
            "white": "White", "black": "Black",
        }

        headers = ["Color"] + [h for h, *_ in _MEAS_COLS]
        widths  = [8] + [w for *_, w in _MEAS_COLS]
        _write_header_row(ws, 1, headers)
        _set_col_widths(ws, widths)
        _freeze(ws, "B2")

        for ri, color in enumerate(color_order, 2):
            r = results.get(color)
            fill = color_fills.get(color)
            label_cell = ws.cell(row=ri, column=1, value=color_label.get(color, color))
            label_cell.font = _BOLD_FONT
            label_cell.alignment = _CENTER
            label_cell.border = _BORDER
            if fill:
                label_cell.fill = fill

            if r:
                _write_meas_row(ws, ri, r, col_offset=1)
                if fill:
                    for ci in range(2, len(_MEAS_COLS) + 2):
                        ws.cell(row=ri, column=ci).fill = fill

        # White/Black 명암비 계산
        w_r = results.get("white")
        bk_r = results.get("black")
        if w_r and bk_r and bk_r.Lv > 0:
            cr = round(w_r.Lv / bk_r.Lv, 1)
            stat_row = len(color_order) + 3
            ws.cell(row=stat_row, column=1, value="Contrast Ratio").font = _BOLD_FONT
            c = ws.cell(row=stat_row, column=2, value=cr)
            c.number_format = _FMT_RATIO
            c.font = Font(bold=True, color="C55A11", size=11)

        path = file_path or _default_path(brand, model, "Gamut")
        wb.save(path)
        return path

    # ── 4. Contrast Ratio ─────────────────────────────────────────────────────

    def export_contrast(
        self,
        results: Dict[float, MeasureResult],
        brand: str,
        model: str,
        file_path: str | None = None,
    ) -> str:
        """윈도우 크기별 Lv + 명암비(CR) 자동 계산."""
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        _add_info_sheet(wb, brand, model, "Contrast Ratio")

        ws = wb.create_sheet("ContrastRatio")

        meas_headers = [h for h, *_ in _MEAS_COLS]
        meas_widths  = [w for *_, w in _MEAS_COLS]

        # 0% window = solid white (peak Lv) / 100% window = full black window (min Lv)
        # CR = Lv_white(0% win) / Lv_black(100% win)
        ref_lv = results.get(0.0, None)
        ref_lv_val = ref_lv.Lv if ref_lv else None

        headers = ["Window (%)", "Lv (cd/m²)", "CR (White/Lv)"] + meas_headers
        _write_header_row(ws, 1, headers)
        ws.column_dimensions["A"].width = 12
        ws.column_dimensions["B"].width = 14
        ws.column_dimensions["C"].width = 16
        _set_col_widths(ws, meas_widths, col_offset=3)
        _freeze(ws, "B2")

        sorted_items = sorted(results.items(), reverse=True)  # Full White first, then 100%→14.1%

        for ri, (win_size, r) in enumerate(sorted_items, 2):
            # CR: full white / black window Lv (full white row itself shows "—")
            if ref_lv_val and r.Lv > 0 and win_size > 0.0:
                cr_val = round(ref_lv_val / r.Lv, 1)
            else:
                cr_val = None

            fill = _GREEN_FILL if win_size == 0.0 else (
                   _YELLOW_FILL if win_size <= 20.0 else None)

            win_label = "Full White" if win_size == 0.0 else win_size
            win_fmt = "@" if win_size == 0.0 else _FMT_RATIO
            for ci, (val, fmt) in enumerate(
                [(win_label, win_fmt), (r.Lv, _FMT_LV), (cr_val, _FMT_RATIO)], 1
            ):
                c = ws.cell(row=ri, column=ci, value=val)
                c.number_format = fmt
                c.alignment = _CENTER
                c.border = _BORDER
                c.font = _BOLD_FONT if ci == 3 else _BASE_FONT
                if fill:
                    c.fill = fill

            _write_meas_row(ws, ri, r, col_offset=3)
            if fill:
                for ci2 in range(4, len(_MEAS_COLS) + 4):
                    ws.cell(row=ri, column=ci2).fill = fill

        path = file_path or _default_path(brand, model, "ContrastRatio")
        wb.save(path)
        return path

    # ── 5. All-session combined export ───────────────────────────────────────

    def export_all_session(
        self,
        brand: str,
        model: str,
        session_swing:    Dict[str, Any],   # "SDR_Vivid" → [MeasureResult]
        session_loading:  Dict[str, Any],   # "SDR_Vivid" → {apl → [MeasureResult]}
        session_gamut:    Dict[str, Any],   # "SDR"/"HDR" → {color → MeasureResult}
        session_contrast: Dict[str, Any],   # "SDR"/"HDR" → {side → MeasureResult}
        file_path: str = "",
    ) -> str:
        """{brand}_{model}_all.xlsx — 모든 측정 결과를 탭으로 나눠 통합 저장."""
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        _add_info_sheet(wb, brand, model, "All Sessions")

        meas_headers = [h for h, *_ in _MEAS_COLS]
        meas_widths  = [w for *_, w in _MEAS_COLS]

        # ── 휘도 스윙 (피벗: 행=#, 열=모드, 값=Lv) ──────────────────────────
        if session_swing:
            swing_keys = sorted(k for k, v in session_swing.items() if v)
            if swing_keys:
                max_n = max(len(session_swing[k]) for k in swing_keys)
                ws_sw = wb.create_sheet("LumSwing")
                _write_header_row(ws_sw, 1, ["#"] + swing_keys)
                ws_sw.column_dimensions["A"].width = 6
                for ci in range(2, len(swing_keys) + 2):
                    ws_sw.column_dimensions[get_column_letter(ci)].width = 14
                _freeze(ws_sw, "B2")
                for i in range(max_n):
                    ri = i + 2
                    idx_c = ws_sw.cell(row=ri, column=1, value=i + 1)
                    idx_c.alignment = _CENTER
                    idx_c.border = _BORDER
                    idx_c.font = _BASE_FONT
                    for ci, key in enumerate(swing_keys, 2):
                        row_list = session_swing.get(key, [])
                        val = row_list[i].Lv if i < len(row_list) else None
                        c = ws_sw.cell(row=ri, column=ci, value=val)
                        c.number_format = _FMT_LV
                        c.alignment = _CENTER
                        c.border = _BORDER
                        c.font = _BASE_FONT
                # 통계
                stat_row = max_n + 3
                ws_sw.cell(row=stat_row, column=1, value="Statistics").font = _BOLD_FONT
                stat_row += 1
                for label in ["Count", "Mean Lv", "Max Lv", "Min Lv"]:
                    ws_sw.cell(row=stat_row, column=1, value=label).font = _BOLD_FONT
                    for ci, key in enumerate(swing_keys, 2):
                        lvs = [r.Lv for r in session_swing.get(key, [])]
                        if not lvs:
                            continue
                        if   label == "Count":   val = len(lvs)
                        elif label == "Mean Lv": val = round(statistics.mean(lvs), 3)
                        elif label == "Max Lv":  val = round(max(lvs), 3)
                        else:                    val = round(min(lvs), 3)
                        c = ws_sw.cell(row=stat_row, column=ci, value=val)
                        c.number_format = _FMT_LV
                        c.alignment = _CENTER
                        c.border = _BORDER
                    stat_row += 1
                if max_n >= 2:
                    self._add_swing_pivot_chart(ws_sw, swing_keys, max_n)

        # ── APL 로딩 Summary ──────────────────────────────────────────────────
        if session_loading:
            cases    = sorted(session_loading.keys())
            all_apls = sorted({apl for cd in session_loading.values() for apl in cd})
            ws_sum = wb.create_sheet("Loading_Summary")
            hdr = ["APL (%)"]
            for c in cases:
                hdr += [f"{c} Avg", f"{c} Max", f"{c} Min"]
            _write_header_row(ws_sum, 1, hdr)
            ws_sum.column_dimensions["A"].width = 10
            for ci in range(2, len(hdr) + 1):
                ws_sum.column_dimensions[get_column_letter(ci)].width = 14
            for ri, apl in enumerate(all_apls, 2):
                ws_sum.cell(row=ri, column=1, value=apl).alignment = _CENTER
                ci = 2
                for case in cases:
                    lvs = [r.Lv for r in session_loading.get(case, {}).get(apl, [])]
                    avg = round(statistics.mean(lvs), 3) if lvs else None
                    mx  = round(max(lvs), 3)             if lvs else None
                    mn  = round(min(lvs), 3)             if lvs else None
                    for val in [avg, mx, mn]:
                        c2 = ws_sum.cell(row=ri, column=ci, value=val)
                        c2.number_format = _FMT_LV
                        c2.alignment = _CENTER
                        c2.border = _BORDER
                        ci += 1
            _freeze(ws_sum, "B2")
            # 케이스별 Raw 시트
            for case, apl_dict in sorted(session_loading.items()):
                ws_raw = wb.create_sheet(f"Loading_{case}"[:31])
                headers_raw = ["APL (%)", "#"] + meas_headers
                _write_header_row(ws_raw, 1, headers_raw)
                _set_col_widths(ws_raw, [10, 6] + meas_widths)
                _freeze(ws_raw, "C2")
                ri = 2
                for apl in sorted(apl_dict):
                    for idx, r in enumerate(apl_dict[apl], 1):
                        ws_raw.cell(row=ri, column=1, value=apl).alignment = _CENTER
                        ws_raw.cell(row=ri, column=2, value=idx).alignment = _CENTER
                        _write_meas_row(ws_raw, ri, r, col_offset=2)
                        ri += 1

        # ── 색재현율 ──────────────────────────────────────────────────────────
        color_order = ["red", "green", "blue", "white", "black"]
        for mode_key, gamut_results in sorted(session_gamut.items()):
            if not gamut_results:
                continue
            ws = wb.create_sheet(f"Gamut_{mode_key}"[:31])
            _write_header_row(ws, 1, ["Color"] + meas_headers)
            _set_col_widths(ws, [8] + meas_widths)
            _freeze(ws, "B2")
            for ri, color in enumerate(color_order, 2):
                r = gamut_results.get(color)
                ws.cell(row=ri, column=1, value=color.capitalize()).font = _BOLD_FONT
                ws.cell(row=ri, column=1).alignment = _CENTER
                ws.cell(row=ri, column=1).border = _BORDER
                if r:
                    _write_meas_row(ws, ri, r, col_offset=1)

        # ── 명암비 ────────────────────────────────────────────────────────────
        for mode_key, contrast_results in sorted(session_contrast.items()):
            if not contrast_results:
                continue
            ws = wb.create_sheet(f"Contrast_{mode_key}"[:31])
            _write_header_row(ws, 1, ["Black H/V (%)", "Lv (cd/m²)", "CR (White/Lv)"] + meas_headers)
            _set_col_widths(ws, [14, 14, 14] + meas_widths)
            _freeze(ws, "B2")
            # 기준: Full White (0.0)
            ref_lv = None
            if 0.0 in contrast_results:
                ref_lv = contrast_results[0.0].Lv
            for ri, side in enumerate(sorted(contrast_results, reverse=True), 2):
                r = contrast_results[side]
                cr = round(ref_lv / r.Lv, 1) if (ref_lv and r.Lv > 0 and side > 0.0) else None
                side_label = "Full White" if side == 0.0 else side
                side_fmt = "@" if side == 0.0 else _FMT_RATIO
                for ci, (val, fmt) in enumerate(
                    [(side_label, side_fmt), (r.Lv, _FMT_LV), (cr, _FMT_RATIO)], 1
                ):
                    c2 = ws.cell(row=ri, column=ci, value=val)
                    c2.number_format = fmt
                    c2.alignment = _CENTER
                    c2.border = _BORDER
                _write_meas_row(ws, ri, r, col_offset=3)

        path = file_path or _default_path(brand, model, "all")
        wb.save(path)
        return path

    # ── 6. Report Template ────────────────────────────────────────────────────

    def export_report_template(
        self,
        brand: str,
        model: str,
        gamut_results: Dict[str, "MeasureResult"] | None = None,
        lum_loading_results: Dict[str, Dict[int, List["MeasureResult"]]] | None = None,
        contrast_results: Dict[float, "MeasureResult"] | None = None,
        file_path: str | None = None,
    ) -> str:
        """모든 측정 결과를 하나의 통합 리포트 워크북으로 저장.

        포함 시트:
          - Info      : 기본 정보
          - Summary   : 측정 항목별 핵심 지표 요약
          - Gamut     : 색재현율 측정값 + 통계
          - LumLoading: APL vs Lv 요약 + 차트
          - Contrast  : 윈도우별 명암비
        """
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        _add_info_sheet(wb, brand, model, "Report")

        # ── Summary 시트 ──────────────────────────────────────────────────────
        ws_sum = wb.create_sheet("Summary")
        ws_sum.column_dimensions["A"].width = 26
        ws_sum.column_dimensions["B"].width = 20
        ws_sum.column_dimensions["C"].width = 20

        _write_header_row(ws_sum, 1, ["항목", "측정값", "비고"])

        rows_data: list[tuple[str, object, str]] = []

        # Gamut 통계
        if gamut_results:
            r_r = gamut_results.get("red")
            r_g = gamut_results.get("green")
            r_b = gamut_results.get("blue")
            if r_r and r_g and r_b:
                stats = calc_gamut_stats(
                    (r_r.u_prime, r_r.v_prime),
                    (r_g.u_prime, r_g.v_prime),
                    (r_b.u_prime, r_b.v_prime),
                )
                rows_data.append(("DCI-P3 Coverage (%)",   stats["dci_overlap"],   ""))
                rows_data.append(("BT.2020 Coverage (%)",  stats["bt2020_overlap"],""))
            if gamut_results.get("white") and gamut_results.get("black"):
                w_lv  = gamut_results["white"].Lv
                bk_lv = gamut_results["black"].Lv
                if bk_lv > 0:
                    rows_data.append(("Gamut CR (White/Black)", round(w_lv / bk_lv, 1), ""))

        # LumLoading 통계 — 케이스별 Peak Lv (APL 10%)
        if lum_loading_results:
            for case, case_data in sorted(lum_loading_results.items()):
                apl10 = case_data.get(10, [])
                if apl10:
                    peak = max(r.Lv for r in apl10)
                    rows_data.append((f"LumLoading Case {case} Peak Lv (APL10)", round(peak, 3), "cd/m²"))

        # Contrast
        if contrast_results:
            ref = contrast_results.get(0.0)
            bk  = contrast_results.get(100.0)
            if ref and bk and bk.Lv > 0:
                cr = round(ref.Lv / bk.Lv, 1)
                rows_data.append(("Contrast Ratio (0%/100%)", cr, ""))

        for ri, (label, val, note) in enumerate(rows_data, 2):
            fill = _GREEN_FILL if ri % 2 == 0 else None
            for ci, v in enumerate([label, val, note], 1):
                c = ws_sum.cell(row=ri, column=ci, value=v)
                c.alignment = _CENTER
                c.border = _BORDER
                c.font = _BASE_FONT
                if fill:
                    c.fill = fill

        # ── Gamut 시트 ────────────────────────────────────────────────────────
        if gamut_results:
            ws_g = wb.create_sheet("Gamut")
            color_order = ["red", "green", "blue", "white", "black"]
            color_fills = {
                "red":   PatternFill("solid", fgColor="FCE4D6"),
                "green": PatternFill("solid", fgColor="E2EFDA"),
                "blue":  PatternFill("solid", fgColor="DAE8FC"),
                "white": PatternFill("solid", fgColor="F2F2F2"),
                "black": PatternFill("solid", fgColor="D9D9D9"),
            }
            headers = ["Color"] + [h for h, *_ in _MEAS_COLS]
            widths  = [8] + [w for *_, w in _MEAS_COLS]
            _write_header_row(ws_g, 1, headers)
            _set_col_widths(ws_g, widths)
            _freeze(ws_g, "B2")
            for ri, color in enumerate(color_order, 2):
                r = gamut_results.get(color)
                fill = color_fills.get(color)
                lc = ws_g.cell(row=ri, column=1, value=color.capitalize())
                lc.font = _BOLD_FONT
                lc.alignment = _CENTER
                lc.border = _BORDER
                if fill:
                    lc.fill = fill
                if r:
                    _write_meas_row(ws_g, ri, r, col_offset=1)
                    if fill:
                        for ci in range(2, len(_MEAS_COLS) + 2):
                            ws_g.cell(row=ri, column=ci).fill = fill

            # Gamut 통계 블록
            r_r = gamut_results.get("red")
            r_g = gamut_results.get("green")
            r_b = gamut_results.get("blue")
            if r_r and r_g and r_b:
                stats = calc_gamut_stats(
                    (r_r.u_prime, r_r.v_prime),
                    (r_g.u_prime, r_g.v_prime),
                    (r_b.u_prime, r_b.v_prime),
                )
                stat_row = len(color_order) + 3
                for label, val in [
                    ("DCI-P3 Coverage (%)",  stats["dci_overlap"]),
                    ("BT.2020 Coverage (%)", stats["bt2020_overlap"]),
                ]:
                    c_label = ws_g.cell(row=stat_row, column=1, value=label)
                    c_label.font = _BOLD_FONT
                    c_val   = ws_g.cell(row=stat_row, column=2, value=val)
                    c_val.number_format = "0.00"
                    stat_row += 1

        # ── LumLoading 시트 ───────────────────────────────────────────────────
        if lum_loading_results:
            import statistics as _stats
            ws_ll = wb.create_sheet("LumLoading")
            cases    = sorted(lum_loading_results.keys())
            all_apls = sorted({apl for cd in lum_loading_results.values() for apl in cd})
            hdr = ["APL (%)"]
            for case in cases:
                hdr += [f"Case {case} Avg Lv", f"Case {case} Max Lv",
                        f"Case {case} Min Lv", f"Case {case} StdDev"]
            _write_header_row(ws_ll, 1, hdr)
            ws_ll.column_dimensions["A"].width = 10
            for ci in range(2, len(hdr) + 1):
                ws_ll.column_dimensions[get_column_letter(ci)].width = 14
            for ri, apl in enumerate(all_apls, 2):
                ws_ll.cell(row=ri, column=1, value=apl).alignment = _CENTER
                ci = 2
                for case in cases:
                    lvs = [r.Lv for r in lum_loading_results.get(case, {}).get(apl, [])]
                    avg = round(_stats.mean(lvs), 3)   if lvs else None
                    mx  = round(max(lvs), 3)            if lvs else None
                    mn  = round(min(lvs), 3)            if lvs else None
                    std = round(_stats.stdev(lvs), 4)   if len(lvs) > 1 else (0 if lvs else None)
                    for val, fmt in [(avg, _FMT_LV), (mx, _FMT_LV), (mn, _FMT_LV), (std, _FMT_LV)]:
                        c = ws_ll.cell(row=ri, column=ci, value=val)
                        c.number_format = fmt
                        c.alignment = _CENTER
                        c.border = _BORDER
                        ci += 1
            _freeze(ws_ll, "B2")
            self._add_apl_chart(ws_ll, all_apls, cases, len(all_apls), brand=brand)

        # ── Contrast 시트 ─────────────────────────────────────────────────────
        if contrast_results:
            ws_cr = wb.create_sheet("Contrast")
            ref_lv_val = contrast_results.get(0.0)
            ref_lv_val = ref_lv_val.Lv if ref_lv_val else None
            meas_headers = [h for h, *_ in _MEAS_COLS]
            meas_widths  = [w for *_, w in _MEAS_COLS]
            headers = ["Window (%)", "Lv (cd/m²)", "CR (White/Lv)"] + meas_headers
            _write_header_row(ws_cr, 1, headers)
            ws_cr.column_dimensions["A"].width = 12
            ws_cr.column_dimensions["B"].width = 14
            ws_cr.column_dimensions["C"].width = 16
            _set_col_widths(ws_cr, meas_widths, col_offset=3)
            _freeze(ws_cr, "B2")
            for ri, (win_size, r) in enumerate(sorted(contrast_results.items(), reverse=True), 2):
                cr_val = round(ref_lv_val / r.Lv, 1) if (ref_lv_val and r.Lv > 0 and win_size > 0.0) else None
                fill = _GREEN_FILL if win_size == 0.0 else (
                       _YELLOW_FILL if win_size <= 20.0 else None)
                win_label = "Full White" if win_size == 0.0 else win_size
                win_fmt = "@" if win_size == 0.0 else _FMT_RATIO
                for ci, (val, fmt) in enumerate(
                    [(win_label, win_fmt), (r.Lv, _FMT_LV), (cr_val, _FMT_RATIO)], 1
                ):
                    c = ws_cr.cell(row=ri, column=ci, value=val)
                    c.number_format = fmt
                    c.alignment = _CENTER
                    c.border = _BORDER
                    c.font = _BOLD_FONT if ci == 3 else _BASE_FONT
                    if fill:
                        c.fill = fill
                _write_meas_row(ws_cr, ri, r, col_offset=3)
                if fill:
                    for ci2 in range(4, len(_MEAS_COLS) + 4):
                        ws_cr.cell(row=ri, column=ci2).fill = fill

        path = file_path or _default_path(brand, model, "Report")
        wb.save(path)
        return path
