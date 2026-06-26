"""ModulePanel — LED module measurement panel.

Layout:
  Controls bar (top)
  ┌─────────────┬─────────────────────────────────────────────────┐
  │  Left       │  Chart grid                                     │
  │  sidebar    │  [Gamma_W][Gamma_R][Gamma_G][Gamma_B]           │
  │  (info,     │  [RevG_W ][RevG_R ][RevG_G ][RevG_B ]          │
  │   spec      │  [duv    ][CIE 1976  (×2)  ][ColorAccuracy]    │
  │   tables)   │  [White Tracking (×2)       ][              ]   │
  │             ├─────────────────────────────────────────────────┤
  │             │  ΔE2000 stats table (full width)                │
  └─────────────┴─────────────────────────────────────────────────┘
"""
from __future__ import annotations

import math
from typing import Any, Dict, List, Optional, Tuple

from PySide6.QtCharts import QChart, QChartView, QLineSeries, QScatterSeries, QValueAxis
from PySide6.QtCore import QMargins, Qt, Signal, Slot
from PySide6.QtGui import QColor, QPainter, QPen
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from core.engine import MeasurementEngine
from core.equipment.base import MeasureResult
from core.gamut_utils import DCI_P3_UV
from core.sequences.module_measure import BT709_REF_UV, DEFAULT_GAMMA_STEPS
from core.sequences.calman_sweep import COLOR_ORDER as CALMAN_COLOR_ORDER
from .worker import MeasurementWorker, wire_worker_cleanup

# ── Color constants ───────────────────────────────────────────────────────────
_CH_COLORS = {"W": "#666677", "R": "#e74c3c", "G": "#27ae60", "B": "#2980e8"}
_PATCH_COLORS = {
    "R": "#e74c3c", "G": "#27ae60", "B": "#2980e8",
    "C": "#1abc9c", "M": "#9b59b6", "Y": "#c8a000", "W": "#666677",
}
_PATCH_ORDER = ["R", "G", "B", "C", "M", "Y", "W"]

# BT.709 xy for color accuracy reference chart
def _uv_to_xy(u: float, v: float) -> Tuple[float, float]:
    d = 6.0 * u - 16.0 * v + 12.0
    if abs(d) < 1e-12:
        return 0.0, 0.0
    return 9.0 * u / d, 4.0 * v / d

_BT709_REF_XY: Dict[str, Tuple[float, float]] = {
    name: _uv_to_xy(u, v) for name, (u, v) in BT709_REF_UV.items()
}

# BT.709 gamut triangle in xy (R,G,B primaries)
_BT709_GAMUT_XY = [(0.6400, 0.3300), (0.3000, 0.6000), (0.1500, 0.0600)]
# DCI-P3 gamut triangle in xy
_DCI_P3_GAMUT_XY = [(0.6800, 0.3200), (0.2650, 0.6900), (0.1500, 0.0600)]

# ── Color science helpers ─────────────────────────────────────────────────────

def _meas_to_xyz(r: MeasureResult) -> Tuple[float, float, float]:
    yc = r.y if r.y > 1e-6 else 1e-6
    return r.x * r.Lv / yc, r.Lv, (1.0 - r.x - r.y) * r.Lv / yc


def _xyz_to_lab(X: float, Y: float, Z: float) -> Tuple[float, float, float]:
    Xn, Yn, Zn = 95.047, 100.0, 108.883

    def f(t: float) -> float:
        return t ** (1.0 / 3.0) if t > 0.008856 else 7.787 * t + 16.0 / 116.0

    return (
        116.0 * f(Y / Yn) - 16.0,
        500.0 * (f(X / Xn) - f(Y / Yn)),
        200.0 * (f(Y / Yn) - f(Z / Zn)),
    )


def _delta_e2000(L1: float, a1: float, b1: float,
                 L2: float, a2: float, b2: float) -> float:
    C1 = math.sqrt(a1 * a1 + b1 * b1)
    C2 = math.sqrt(a2 * a2 + b2 * b2)
    Cab7 = ((C1 + C2) / 2.0) ** 7
    G = 0.5 * (1.0 - math.sqrt(Cab7 / (Cab7 + 25.0 ** 7)))
    a1p, a2p = a1 * (1.0 + G), a2 * (1.0 + G)
    C1p = math.sqrt(a1p ** 2 + b1 ** 2)
    C2p = math.sqrt(a2p ** 2 + b2 ** 2)
    h1p = math.degrees(math.atan2(b1, a1p)) % 360.0
    h2p = math.degrees(math.atan2(b2, a2p)) % 360.0
    dLp = L2 - L1
    dCp = C2p - C1p
    if C1p * C2p == 0:
        dhp = 0.0
    elif abs(h2p - h1p) <= 180.0:
        dhp = h2p - h1p
    elif h2p > h1p:
        dhp = h2p - h1p - 360.0
    else:
        dhp = h2p - h1p + 360.0
    dHp = 2.0 * math.sqrt(C1p * C2p) * math.sin(math.radians(dhp / 2.0))
    Lp_avg = (L1 + L2) / 2.0
    Cp_avg = (C1p + C2p) / 2.0
    if C1p * C2p == 0:
        hp_avg = h1p + h2p
    elif abs(h1p - h2p) <= 180.0:
        hp_avg = (h1p + h2p) / 2.0
    elif h1p + h2p < 360.0:
        hp_avg = (h1p + h2p + 360.0) / 2.0
    else:
        hp_avg = (h1p + h2p - 360.0) / 2.0
    T = (1.0
         - 0.17 * math.cos(math.radians(hp_avg - 30.0))
         + 0.24 * math.cos(math.radians(2.0 * hp_avg))
         + 0.32 * math.cos(math.radians(3.0 * hp_avg + 6.0))
         - 0.20 * math.cos(math.radians(4.0 * hp_avg - 63.0)))
    SL = 1.0 + 0.015 * (Lp_avg - 50.0) ** 2 / math.sqrt(20.0 + (Lp_avg - 50.0) ** 2)
    SC = 1.0 + 0.045 * Cp_avg
    SH = 1.0 + 0.015 * Cp_avg * T
    Cp7 = Cp_avg ** 7
    RC = 2.0 * math.sqrt(Cp7 / (Cp7 + 25.0 ** 7))
    d_theta = 30.0 * math.exp(-((hp_avg - 275.0) / 25.0) ** 2)
    RT = -math.sin(math.radians(2.0 * d_theta)) * RC
    return math.sqrt(
        (dLp / SL) ** 2 + (dCp / SC) ** 2 + (dHp / SH) ** 2
        + RT * (dCp / SC) * (dHp / SH)
    )


def _calc_de2000(r: MeasureResult, ref_x: float, ref_y: float) -> float:
    """ΔE2000 vs reference chromaticity at same luminance."""
    X_m, Y_m, Z_m = _meas_to_xyz(r)
    L_m, a_m, b_m = _xyz_to_lab(X_m, Y_m, Z_m)
    ryc = ref_y if ref_y > 1e-6 else 1e-6
    X_r = ref_x * Y_m / ryc
    Z_r = (1.0 - ref_x - ref_y) * Y_m / ryc
    L_r, a_r, b_r = _xyz_to_lab(X_r, Y_m, Z_r)
    return _delta_e2000(L_r, a_r, b_r, L_m, a_m, b_m)


# ── Chart factory ─────────────────────────────────────────────────────────────

_CHART_BG      = QColor("#111828")
_CHART_GRID    = QColor("#1e2a48")
_CHART_AX_TEXT = QColor("#506090")
_CHART_TITLE   = QColor("#6878b0")


def _make_chart(title: str, x_label: str, x_min: float, x_max: float,
                y_label: str, y_min: float, y_max: float,
                tick_x: int = 5, tick_y: int = 5) -> Tuple[QChart, QValueAxis, QValueAxis]:
    chart = QChart()
    chart.setTitle(title)
    chart.setBackgroundBrush(_CHART_BG)
    chart.setTitleBrush(_CHART_TITLE)
    chart.setMargins(QMargins(2, 4, 2, 2))
    tf = chart.titleFont(); tf.setPointSize(8); tf.setBold(True); chart.setTitleFont(tf)
    chart.legend().hide()

    def _ax(label: str, lo: float, hi: float, ticks: int) -> QValueAxis:
        ax = QValueAxis()
        ax.setTitleText(label); ax.setRange(lo, hi); ax.setTickCount(ticks)
        ax.setLabelsBrush(_CHART_AX_TEXT); ax.setTitleBrush(_CHART_AX_TEXT)
        ax.setGridLineColor(_CHART_GRID)
        lf = ax.labelsFont(); lf.setPointSize(7); ax.setLabelsFont(lf)
        tf2 = ax.titleFont(); tf2.setPointSize(7); ax.setTitleFont(tf2)
        ax.setLabelFormat("%.2g")
        return ax

    ax_x = _ax(x_label, x_min, x_max, tick_x)
    ax_y = _ax(y_label, y_min, y_max, tick_y)
    chart.addAxis(ax_x, Qt.AlignmentFlag.AlignBottom)
    chart.addAxis(ax_y, Qt.AlignmentFlag.AlignLeft)
    return chart, ax_x, ax_y


def _chart_view(chart: QChart, min_h: int = 220) -> QChartView:
    cv = QChartView(chart)
    cv.setRenderHint(QPainter.RenderHint.Antialiasing)
    cv.setStyleSheet(
        "background:#111828;"
        "border:1px solid #1e2a48;"
        "border-radius:8px;"
    )
    cv.setMinimumHeight(min_h)
    cv.setSizePolicy(
        QSizePolicy.Policy.Expanding,
        QSizePolicy.Policy.Fixed,
    )
    return cv


# ── Left sidebar ──────────────────────────────────────────────────────────────

def _section_label(text: str) -> QLabel:
    lbl = QLabel(text.upper())
    lbl.setStyleSheet(
        "font-size:9px;font-weight:700;color:#4e8df8;"
        "background:transparent;padding:6px 0 2px 0;"
        "letter-spacing:0.1em;"
    )
    return lbl


def _make_table(rows: int, cols: int, headers: List[str],
                row_height: int = 18) -> QTableWidget:
    t = QTableWidget(rows, cols)
    t.setHorizontalHeaderLabels(headers)
    t.verticalHeader().hide()
    t.verticalHeader().setDefaultSectionSize(row_height)
    t.horizontalHeader().setDefaultSectionSize(60)
    t.setAlternatingRowColors(True)
    t.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
    t.setSelectionMode(QTableWidget.SelectionMode.NoSelection)
    t.setShowGrid(True)
    t.setFixedHeight(rows * row_height + 24)
    t.setStyleSheet("font-size:10px;")
    return t


def _titem(text: str, align=Qt.AlignmentFlag.AlignCenter) -> QTableWidgetItem:
    it = QTableWidgetItem(text)
    it.setTextAlignment(align)
    return it


class _LeftSidebar(QWidget):
    """Left info + spec sidebar."""

    # Row indices for each section table
    _MODULE_ROWS = ["White", "Black", "CR", "Color x,y", "CCT", "duv"]
    _GAMMA_ROWS  = ["White", "Red", "Green", "Blue"]
    _GAMUT_ROWS  = ["DCI Overlap", "Size"]
    _DE_ROWS     = ["AVG.", "Max."]

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setFixedWidth(175)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(2, 2, 2, 2)
        layout.setSpacing(2)

        # ── Model Info ────────────────────────────────────────────────
        layout.addWidget(_section_label("Model Info."))
        _info_fields = [
            "IP Target Year", "Model Name", "Event",
            "Panel Maker", "BLU Type", "Dev. Grade",
            "SoC", "SW Version", "Serial No.", "Date",
        ]
        self._info_table = QTableWidget(len(_info_fields), 2)
        self._info_table.setHorizontalHeaderLabels(["Field", "Value"])
        self._info_table.verticalHeader().hide()
        self._info_table.verticalHeader().setDefaultSectionSize(17)
        self._info_table.setAlternatingRowColors(True)
        self._info_table.setShowGrid(True)
        self._info_table.setStyleSheet("font-size:10px;")
        self._info_table.setColumnWidth(0, 90)
        self._info_table.horizontalHeader().setStretchLastSection(True)
        self._info_table.setFixedHeight(len(_info_fields) * 17 + 24)
        for ri, field in enumerate(_info_fields):
            it = _titem(field)
            it.setFlags(it.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self._info_table.setItem(ri, 0, it)
            self._info_table.setItem(ri, 1, QTableWidgetItem(""))
        layout.addWidget(self._info_table)

        # ── Module specs ──────────────────────────────────────────────
        layout.addWidget(_section_label("Module"))
        self._mod_table = _make_table(len(self._MODULE_ROWS), 3,
                                      ["Pattern", "Spec.", "Result"])
        _mod_specs = ["400", "0.0001", "400000×", "-", "10000K", "±0.006"]
        for ri, (row, spec) in enumerate(zip(self._MODULE_ROWS, _mod_specs)):
            self._mod_table.setItem(ri, 0, _titem(row))
            self._mod_table.setItem(ri, 1, _titem(spec))
            self._mod_table.setItem(ri, 2, _titem("—"))
        layout.addWidget(self._mod_table)

        # ── Gamma specs ───────────────────────────────────────────────
        layout.addWidget(_section_label("Gamma"))
        self._gamma_table = _make_table(len(self._GAMMA_ROWS), 3,
                                        ["Pattern", "Spec.", "Result"])
        for ri, row in enumerate(self._GAMMA_ROWS):
            self._gamma_table.setItem(ri, 0, _titem(row))
            self._gamma_table.setItem(ri, 1, _titem("2.2 ± 0.2"))
            self._gamma_table.setItem(ri, 2, _titem("—"))
        layout.addWidget(self._gamma_table)

        # ── Gamut ─────────────────────────────────────────────────────
        layout.addWidget(_section_label("Gamut"))
        self._gamut_table = _make_table(len(self._GAMUT_ROWS), 3,
                                        ["Standard", "Spec.", "Result"])
        _gamut_specs = ["100%", "-"]
        for ri, (row, spec) in enumerate(zip(self._GAMUT_ROWS, _gamut_specs)):
            self._gamut_table.setItem(ri, 0, _titem(row))
            self._gamut_table.setItem(ri, 1, _titem(spec))
            self._gamut_table.setItem(ri, 2, _titem("—"))
        layout.addWidget(self._gamut_table)

        # ── Color Accuracy ────────────────────────────────────────────
        layout.addWidget(_section_label("Module Color Accuracy"))
        self._de_table = _make_table(len(self._DE_ROWS), 3,
                                     ["Standard", "Target", "Result"])
        _de_targets = ["< 3.0", "< 5.0"]
        for ri, (row, tgt) in enumerate(zip(self._DE_ROWS, _de_targets)):
            self._de_table.setItem(ri, 0, _titem(row))
            self._de_table.setItem(ri, 1, _titem(tgt))
            self._de_table.setItem(ri, 2, _titem("—"))
        self._de_table.setStyleSheet("font-size:10px;background:#ffffd0;")
        layout.addWidget(self._de_table)

        layout.addStretch()

    # ── Public update methods ─────────────────────────────────────────

    def _set_result(self, table: QTableWidget, row: int, value: str,
                    ok: Optional[bool] = None) -> None:
        it = _titem(value)
        if ok is True:
            it.setForeground(QColor("#1a7040"))
        elif ok is False:
            it.setForeground(QColor("#c02020"))
        table.setItem(row, 2, it)

    def update_module_white(self, lv: float, cct: float, duv: float,
                            xy: Tuple[float, float]) -> None:
        self._set_result(self._mod_table, 0, f"{lv:.1f}")
        self._set_result(self._mod_table, 3, f"({xy[0]:.4f}, {xy[1]:.4f})")
        self._set_result(self._mod_table, 4, f"{cct:.0f} K")
        ok_duv = abs(duv) <= 0.006
        self._set_result(self._mod_table, 5, f"{duv:.5f}", ok_duv)

    def update_gamma(self, ch: str, avg_gamma: Optional[float]) -> None:
        ch_map = {"W": 0, "R": 1, "G": 2, "B": 3}
        ri = ch_map.get(ch)
        if ri is None:
            return
        if avg_gamma is None:
            self._set_result(self._gamma_table, ri, "—")
        else:
            ok = 2.0 <= avg_gamma <= 2.4
            self._set_result(self._gamma_table, ri, f"{avg_gamma:.3f}", ok)

    def update_gamut(self, dci_overlap: float) -> None:
        ok = dci_overlap >= 100.0
        self._set_result(self._gamut_table, 0, f"{dci_overlap:.1f}%", ok)

    def update_de2000(self, avg: float, max_: float) -> None:
        self._set_result(self._de_table, 0, f"{avg:.2f}", avg < 3.0)
        self._set_result(self._de_table, 1, f"{max_:.2f}", max_ < 5.0)

    def clear(self) -> None:
        for table, count in [(self._mod_table, len(self._MODULE_ROWS)),
                              (self._gamma_table, len(self._GAMMA_ROWS)),
                              (self._gamut_table, len(self._GAMUT_ROWS)),
                              (self._de_table, len(self._DE_ROWS))]:
            for ri in range(count):
                table.setItem(ri, 2, _titem("—"))


# ── Chart area ────────────────────────────────────────────────────────────────

class _ChartArea(QWidget):
    """Right-side scrollable chart grid + bottom stats table."""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(4)

        # ── Scrollable chart grid ─────────────────────────────────────
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        grid_widget = QWidget()
        self._grid = QGridLayout(grid_widget)
        self._grid.setSpacing(3)
        self._grid.setContentsMargins(2, 2, 2, 2)
        for col in range(4):
            self._grid.setColumnStretch(col, 1)

        # ── Row 0: Gamma charts ───────────────────────────────────────
        self._gamma_series: Dict[str, QLineSeries] = {}
        self._gamma_charts: Dict[str, QChart] = {}
        for ci, ch in enumerate(["W", "R", "G", "B"]):
            chart, ax_x, ax_y = _make_chart(
                f"Gamma_{ch}", "Input level", 0, 255,
                "Normalized Lum.", 0.0, 1.05, tick_x=6, tick_y=6,
            )
            # Reference γ2.2
            ref = QLineSeries()
            ref_pen = QPen(QColor("#aaaaaa")); ref_pen.setWidth(1)
            ref_pen.setStyle(Qt.PenStyle.DashLine); ref.setPen(ref_pen)
            for i in range(11):
                x = i / 10 * 255
                y = (i / 10) ** 2.2
                ref.append(x, y)
            chart.addSeries(ref); ref.attachAxis(ax_x); ref.attachAxis(ax_y)
            # Measured
            s = QLineSeries()
            pen = QPen(QColor(_CH_COLORS[ch])); pen.setWidth(2); s.setPen(pen)
            chart.addSeries(s); s.attachAxis(ax_x); s.attachAxis(ax_y)
            self._gamma_series[ch] = s
            self._gamma_charts[ch] = chart
            self._grid.addWidget(_chart_view(chart), 0, ci)

        # ── Row 1: Reverse Gamma charts ───────────────────────────────
        self._rev_series: Dict[str, QLineSeries] = {}
        for ci, ch in enumerate(["W", "R", "G", "B"]):
            chart, ax_x, ax_y = _make_chart(
                f"Reverse Gamma_{ch}", "Input level", 0, 255,
                "#Lv_Nor.", 0.0, 1.05, tick_x=6, tick_y=6,
            )
            s = QLineSeries()
            pen = QPen(QColor(_CH_COLORS[ch])); pen.setWidth(2); s.setPen(pen)
            chart.addSeries(s); s.attachAxis(ax_x); s.attachAxis(ax_y)
            self._rev_series[ch] = s
            self._grid.addWidget(_chart_view(chart), 1, ci)

        # ── Row 2: duv | CIE 1976 (span2) | Color Accuracy ───────────

        # duv chart
        self._duv_chart, ax_dx, ax_dy = _make_chart(
            "duv", "Input level", 0, 255, "duv", -0.012, 0.012, tick_x=6, tick_y=5,
        )
        self._duv_series = QLineSeries()
        pen = QPen(QColor("#2980e8")); pen.setWidth(2); self._duv_series.setPen(pen)
        self._duv_chart.addSeries(self._duv_series)
        self._duv_series.attachAxis(ax_dx); self._duv_series.attachAxis(ax_dy)
        # Zero line
        zero = QLineSeries()
        zpen = QPen(QColor("#aaaaaa")); zpen.setWidth(1); zpen.setStyle(Qt.PenStyle.DashLine); zero.setPen(zpen)
        zero.append(0, 0); zero.append(255, 0)
        self._duv_chart.addSeries(zero); zero.attachAxis(ax_dx); zero.attachAxis(ax_dy)
        self._grid.addWidget(_chart_view(self._duv_chart), 2, 0)

        # CIE 1976 u'v' chart
        self._cie_chart, ax_cu, ax_cv = _make_chart(
            "CIE 1976  Coordinates", "u'", 0.0, 0.70, "v'", 0.0, 0.65, tick_x=8, tick_y=7,
        )
        self._cie_chart.legend().show()
        lf = self._cie_chart.legend().font(); lf.setPointSize(7)
        self._cie_chart.legend().setFont(lf)
        # BT.709 reference triangle
        self._cie_bt709: Optional[QLineSeries] = None
        self._cie_dci: Optional[QLineSeries] = None
        self._cie_meas_tri: Optional[QLineSeries] = None
        self._cie_scatter: Optional[QScatterSeries] = None
        self._cie_axes = (ax_cu, ax_cv)
        self._add_cie_references()
        self._grid.addWidget(_chart_view(self._cie_chart), 2, 1, 1, 2)

        # Color Accuracy xy chart
        self._ca_chart, ax_cx, ax_cy = _make_chart(
            "Module Color Accuracy", "x", 0.0, 0.80, "y", 0.0, 0.90, tick_x=9, tick_y=10,
        )
        self._ca_chart.legend().show()
        lf2 = self._ca_chart.legend().font(); lf2.setPointSize(7)
        self._ca_chart.legend().setFont(lf2)
        self._ca_meas_scatter: Optional[QScatterSeries] = None
        self._ca_target_scatter: Optional[QScatterSeries] = None
        self._ca_axes = (ax_cx, ax_cy)
        self._add_ca_references()
        self._grid.addWidget(_chart_view(self._ca_chart), 2, 3)

        # ── Row 3: White Tracking ─────────────────────────────────────
        self._wt_chart, ax_wx, ax_wy = _make_chart(
            "White Tracking", "Input level", 0, 255,
            "Color Temp. [K]", 0, 15000, tick_x=6, tick_y=7,
        )
        self._wt_series = QLineSeries()
        wpen = QPen(QColor("#2980e8")); wpen.setWidth(2); self._wt_series.setPen(wpen)
        self._wt_chart.addSeries(self._wt_series)
        self._wt_series.attachAxis(ax_wx); self._wt_series.attachAxis(ax_wy)
        self._grid.addWidget(_chart_view(self._wt_chart), 3, 0, 1, 2)
        # placeholder
        ph = QWidget(); ph.setStyleSheet("background:#f4f6fc;border:1px solid #dde0ec;")
        self._grid.addWidget(ph, 3, 2, 1, 2)

        scroll.setWidget(grid_widget)
        outer.addWidget(scroll, stretch=1)

        # ── Bottom: ΔE2000 stats table ────────────────────────────────
        self._de_stats = QTableWidget(1, 5)
        self._de_stats.setHorizontalHeaderLabels(
            ["Parameter", "Target", "Avg.", "Max.", "Min."]
        )
        self._de_stats.verticalHeader().hide()
        self._de_stats.verticalHeader().setDefaultSectionSize(20)
        self._de_stats.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self._de_stats.setSelectionMode(QTableWidget.SelectionMode.NoSelection)
        self._de_stats.setAlternatingRowColors(True)
        self._de_stats.setStyleSheet("font-size:11px;")
        self._de_stats.setFixedHeight(44)
        self._de_stats.horizontalHeader().setStretchLastSection(True)
        for ci, val in enumerate(["ΔE2000", "< 3", "—", "—", "—"]):
            self._de_stats.setItem(0, ci, _titem(val))
        outer.addWidget(self._de_stats)

    # ── Reference gamuts ──────────────────────────────────────────────

    def _add_cie_references(self) -> None:
        ax_u, ax_v = self._cie_axes
        # BT.709 dashed
        s = QLineSeries(); s.setName("BT.709")
        pen = QPen(QColor("#222222")); pen.setWidth(1); pen.setStyle(Qt.PenStyle.DashLine); s.setPen(pen)
        bt709_pts = [BT709_REF_UV[ch] for ch in ["R", "G", "B"]] + [BT709_REF_UV["R"]]
        for u, v in bt709_pts:
            s.append(u, v)
        self._cie_chart.addSeries(s); s.attachAxis(ax_u); s.attachAxis(ax_v)
        # DCI solid
        s2 = QLineSeries(); s2.setName("DCI")
        pen2 = QPen(QColor("#222222")); pen2.setWidth(1); s2.setPen(pen2)
        pts = DCI_P3_UV + [DCI_P3_UV[0]]
        for u, v in pts:
            s2.append(u, v)
        self._cie_chart.addSeries(s2); s2.attachAxis(ax_u); s2.attachAxis(ax_v)

    def _add_ca_references(self) -> None:
        ax_x, ax_y = self._ca_axes
        # BT.709 dashed
        s = QLineSeries(); s.setName("BT.709")
        pen = QPen(QColor("#222222")); pen.setWidth(1); pen.setStyle(Qt.PenStyle.DashLine); s.setPen(pen)
        pts = _BT709_GAMUT_XY + [_BT709_GAMUT_XY[0]]
        for x, y in pts:
            s.append(x, y)
        self._ca_chart.addSeries(s); s.attachAxis(ax_x); s.attachAxis(ax_y)
        # DCI solid
        s2 = QLineSeries(); s2.setName("DCI")
        pen2 = QPen(QColor("#111111")); pen2.setWidth(1); s2.setPen(pen2)
        pts2 = _DCI_P3_GAMUT_XY + [_DCI_P3_GAMUT_XY[0]]
        for x, y in pts2:
            s2.append(x, y)
        self._ca_chart.addSeries(s2); s2.attachAxis(ax_x); s2.attachAxis(ax_y)

    # ── Gamma update ──────────────────────────────────────────────────

    def update_gamma_channel(self, ch: str, points: List[dict]) -> None:
        s = self._gamma_series.get(ch)
        if not s or not points:
            return
        lv_max = max((p["Lv"] for p in points), default=0.0)
        s.clear()
        for p in points:
            y = (p["Lv"] / lv_max) if lv_max > 0 else 0.0
            s.append(p["level"], y)

        # Reverse gamma: delta_Lv normalized
        rev = self._rev_series.get(ch)
        if rev:
            rev.clear()
            lvs = [p["Lv"] for p in points]
            deltas = [max(0.0, lvs[i] - lvs[i - 1]) for i in range(1, len(lvs))]
            d_max = max(deltas) if deltas else 1.0
            levels = [p["level"] for p in points]
            for i, d in enumerate(deltas, 1):
                rev.append(levels[i], d / d_max if d_max > 0 else 0.0)

        # White tracking data (CCT / Duv) from W channel
        if ch == "W":
            self._wt_series.clear()
            self._duv_series.clear()
            for p in points:
                r: Optional[MeasureResult] = p.get("result")
                if r is not None and r.cct and r.cct > 0:
                    self._wt_series.append(p["level"], r.cct)
                    self._duv_series.append(p["level"], r.duv)

    # ── Color patch update ────────────────────────────────────────────

    def update_colors(self, color_results: Dict[str, MeasureResult]) -> None:
        ax_u, ax_v = self._cie_axes
        ax_x, ax_y = self._ca_axes

        # Remove old measured series from their respective charts
        for series in [self._cie_meas_tri, self._cie_scatter]:
            if series is not None:
                try:
                    self._cie_chart.removeSeries(series)
                except Exception:
                    pass
        for series in [self._ca_meas_scatter, self._ca_target_scatter]:
            if series is not None:
                try:
                    self._ca_chart.removeSeries(series)
                except Exception:
                    pass
        self._cie_meas_tri = self._cie_scatter = None
        self._ca_meas_scatter = self._ca_target_scatter = None

        # CIE 1976: measured RGB triangle
        r_r = color_results.get("R"); r_g = color_results.get("G"); r_b = color_results.get("B")
        if r_r and r_g and r_b:
            tri = QLineSeries()
            pen = QPen(QColor("#f7c94f")); pen.setWidth(2); tri.setPen(pen)
            for mres in [r_r, r_g, r_b, r_r]:
                tri.append(mres.u_prime, mres.v_prime)
            self._cie_chart.addSeries(tri); tri.attachAxis(ax_u); tri.attachAxis(ax_v)
            self._cie_meas_tri = tri

        # CIE 1976: scatter for all patches
        cie_sc = QScatterSeries()
        cie_sc.setMarkerSize(7.0)
        for name in _PATCH_ORDER:
            mres = color_results.get(name)
            if mres:
                cie_sc.append(mres.u_prime, mres.v_prime)
                cie_sc.setColor(QColor(_PATCH_COLORS.get(name, "#888888")))
        if cie_sc.count() > 0:
            self._cie_chart.addSeries(cie_sc); cie_sc.attachAxis(ax_u); cie_sc.attachAxis(ax_v)
            self._cie_scatter = cie_sc

        # Color Accuracy xy: measured scatter
        meas_sc = QScatterSeries(); meas_sc.setName("Measured")
        meas_sc.setMarkerSize(8.0); meas_sc.setBorderColor(QColor("#ffffff"))
        for name in _PATCH_ORDER:
            mres = color_results.get(name)
            if mres:
                meas_sc.append(mres.x, mres.y)
                meas_sc.setColor(QColor(_PATCH_COLORS.get(name, "#888888")))
        if meas_sc.count() > 0:
            self._ca_chart.addSeries(meas_sc); meas_sc.attachAxis(ax_x); meas_sc.attachAxis(ax_y)
            self._ca_meas_scatter = meas_sc

        # Color Accuracy xy: target scatter (squares)
        tgt_sc = QScatterSeries(); tgt_sc.setName("Target")
        tgt_sc.setMarkerShape(QScatterSeries.MarkerShape.MarkerShapeRectangle)
        tgt_sc.setMarkerSize(8.0); tgt_sc.setColor(QColor("#000000"))
        tgt_sc.setBorderColor(QColor("#000000"))
        for name in _PATCH_ORDER:
            ref = _BT709_REF_XY.get(name)
            if ref:
                tgt_sc.append(ref[0], ref[1])
        if tgt_sc.count() > 0:
            self._ca_chart.addSeries(tgt_sc); tgt_sc.attachAxis(ax_x); tgt_sc.attachAxis(ax_y)
            self._ca_target_scatter = tgt_sc

    # ── ΔE2000 stats ──────────────────────────────────────────────────

    def update_de_stats(self, de_values: Dict[str, float]) -> None:
        if not de_values:
            return
        vals = list(de_values.values())
        avg = sum(vals) / len(vals)
        mx  = max(vals)
        mn  = min(vals)
        for ci, (val, ok_thr) in enumerate([(avg, 3.0), (mx, 5.0), (mn, None)], 2):
            it = _titem(f"{val:.2f}")
            if ok_thr is not None:
                it.setForeground(QColor("#1a7040" if val < ok_thr else "#c02020"))
            self._de_stats.setItem(0, ci, it)

    def clear(self) -> None:
        for s in self._gamma_series.values():
            s.clear()
        for s in self._rev_series.values():
            s.clear()
        self._wt_series.clear()
        self._duv_series.clear()
        for series in [self._cie_meas_tri, self._cie_scatter]:
            if series is not None:
                try:
                    self._cie_chart.removeSeries(series)
                except Exception:
                    pass
        for series in [self._ca_meas_scatter, self._ca_target_scatter]:
            if series is not None:
                try:
                    self._ca_chart.removeSeries(series)
                except Exception:
                    pass
        self._cie_meas_tri = self._cie_scatter = None
        self._ca_meas_scatter = self._ca_target_scatter = None
        for ci in range(2, 5):
            self._de_stats.setItem(0, ci, _titem("—"))


# ── Gamma sub-panel ───────────────────────────────────────────────────────────

class GammaSubPanel(QWidget):
    """Gamma-only module measurement: W/R/G/B gamma + reverse gamma charts."""

    def __init__(self, engine: MeasurementEngine,
                 parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._engine = engine
        self._worker: Optional[MeasurementWorker] = None

        root = QVBoxLayout(self)
        root.setSpacing(4); root.setContentsMargins(0, 0, 0, 0)

        # Controls
        ctrl = QHBoxLayout(); ctrl.setSpacing(6)
        title = QLabel("Gamma Measurement")
        title.setStyleSheet("font-size:13px;font-weight:bold;")
        ctrl.addWidget(title)
        self._hdr_check = QCheckBox("HDR"); ctrl.addWidget(self._hdr_check)
        ctrl.addSpacing(8)
        ctrl.addWidget(QLabel("Channel:"))
        self._ch_w = QCheckBox("W"); self._ch_w.setChecked(True)
        self._ch_r = QCheckBox("R"); self._ch_r.setChecked(True)
        self._ch_g = QCheckBox("G"); self._ch_g.setChecked(True)
        self._ch_b = QCheckBox("B"); self._ch_b.setChecked(True)
        for cb in (self._ch_w, self._ch_r, self._ch_g, self._ch_b):
            ctrl.addWidget(cb)
        ctrl.addSpacing(8)
        ctrl.addWidget(QLabel("Step:"))
        self._step_combo = QComboBox()
        self._step_combo.addItems(["10-step (default)", "17-step", "5-step"])
        self._step_combo.setFixedWidth(130)
        ctrl.addWidget(self._step_combo)
        ctrl.addStretch()
        self._btn_run = QPushButton("▶  Start")
        self._btn_run.setObjectName("primary")
        self._btn_run.clicked.connect(self._run)
        self._btn_stop = QPushButton("■  Stop")
        self._btn_stop.setObjectName("danger")
        self._btn_stop.setEnabled(False)
        self._btn_stop.clicked.connect(self._stop)
        self._btn_clear = QPushButton("Clear")
        self._btn_clear.clicked.connect(self._clear)
        for btn in (self._btn_run, self._btn_stop, self._btn_clear):
            ctrl.addWidget(btn)
        root.addLayout(ctrl)

        prog_row = QHBoxLayout(); prog_row.setSpacing(6)
        self._progress = QProgressBar(); self._progress.setFixedHeight(5)
        prog_row.addWidget(self._progress, stretch=1)
        self._status_lbl = QLabel("Idle")
        self._status_lbl.setObjectName("muted")
        self._status_lbl.setFixedWidth(240)
        prog_row.addWidget(self._status_lbl)
        root.addLayout(prog_row)

        # 4 gamma + 4 reverse gamma charts
        grid_w = QWidget()
        grid = QGridLayout(grid_w)
        grid.setSpacing(2); grid.setContentsMargins(0, 0, 0, 0)
        self._gamma_series: Dict[str, QLineSeries] = {}
        self._rev_series: Dict[str, QLineSeries] = {}
        for ci, ch in enumerate(["W", "R", "G", "B"]):
            gc, gx, gy = _make_chart(f"Gamma {ch}", "Level", 0, 255, "Norm Lv", 0, 1.1)
            ref = QLineSeries()
            rp = QPen(QColor("#bbbbbb")); rp.setWidth(1)
            rp.setStyle(Qt.PenStyle.DashLine); ref.setPen(rp)
            for lv in range(0, 256, 5):
                ref.append(lv, (lv / 255) ** 2.2)
            gc.addSeries(ref); ref.attachAxis(gx); ref.attachAxis(gy)
            s = QLineSeries()
            pen = QPen(QColor(_CH_COLORS[ch])); pen.setWidth(2); s.setPen(pen)
            gc.addSeries(s); s.attachAxis(gx); s.attachAxis(gy)
            self._gamma_series[ch] = s
            grid.addWidget(_chart_view(gc), 0, ci)

            rc, rx, ry = _make_chart(f"Rev.Gamma {ch}", "Level", 0, 255, "Δ Norm", 0, 1.05)
            rs = QLineSeries()
            rpen = QPen(QColor(_CH_COLORS[ch])); rpen.setWidth(2); rs.setPen(rpen)
            rc.addSeries(rs); rs.attachAxis(rx); rs.attachAxis(ry)
            self._rev_series[ch] = rs
            grid.addWidget(_chart_view(rc), 1, ci)

        root.addWidget(grid_w, stretch=1)

    def _get_channels(self) -> List[str]:
        return [ch for cb, ch in [(self._ch_w, "W"), (self._ch_r, "R"),
                                  (self._ch_g, "G"), (self._ch_b, "B")]
                if cb.isChecked()]

    def _get_steps(self) -> List[int]:
        txt = self._step_combo.currentText()
        if "17" in txt:
            s = list(range(0, 256, 16))
            if s[-1] != 255: s.append(255)
            return s
        if "5" in txt:
            return [0, 64, 128, 192, 255]
        return list(DEFAULT_GAMMA_STEPS)

    def _run(self) -> None:
        self._clear()
        channels = self._get_channels()
        if not channels:
            QMessageBox.warning(self, "Settings Error", "Select at least one channel.")
            return
        self._btn_run.setEnabled(False); self._btn_stop.setEnabled(True)
        self._status_lbl.setText("Measuring…")
        self._worker = MeasurementWorker(
            self._engine, "module_measure",
            is_hdr=self._hdr_check.isChecked(),
            gamma_channels=channels,
            gamma_steps=self._get_steps(),
            ref_uv={},
            run_gamma=True,
            run_colors=False,
        )
        self._worker.progress.connect(self._on_progress)
        self._worker.succeeded.connect(self._on_finished)
        self._worker.error.connect(lambda m: (
            QMessageBox.critical(self, "Error", m),
            self._btn_run.setEnabled(True),
            self._btn_stop.setEnabled(False),
        ))
        wire_worker_cleanup(self._worker, self, "_worker")
        self._worker.start()

    def _stop(self) -> None:
        self._engine.stop_sequence()
        if self._worker:
            self._worker.requestInterruption()

    def _clear(self) -> None:
        for s in self._gamma_series.values(): s.clear()
        for s in self._rev_series.values(): s.clear()
        self._progress.setValue(0)
        self._status_lbl.setText("Cleared")

    @Slot(str, float, object)
    def _on_progress(self, step: str, pct: float, data: Any) -> None:
        self._progress.setValue(int(pct * 100))
        if step == "module_gamma" and isinstance(data, dict):
            ch = data.get("channel", ""); level = data.get("level", 0)
            r = data.get("result")
            if r:
                s = self._gamma_series.get(ch)
                if s:
                    lv_max = r.Lv
                    s.append(level, min(lv_max / max(lv_max, 1e-9), 1.5))
                self._status_lbl.setText(
                    f"Gamma {ch} L={level}  Lv={r.Lv:.2f}  ({int(pct*100)}%)"
                )

    @Slot(object)
    def _on_finished(self, result: Any) -> None:
        self._btn_run.setEnabled(True); self._btn_stop.setEnabled(False)
        self._progress.setValue(100)
        if not isinstance(result, dict):
            return
        for ch, points in result.get("gamma", {}).items():
            gs = self._gamma_series.get(ch)
            rs = self._rev_series.get(ch)
            if not gs or not points:
                continue
            gs.clear()
            lvs = [p["Lv"] for p in points]
            lv_max = max(lvs) if lvs else 1.0
            for p in points:
                gs.append(p["level"], p["Lv"] / lv_max if lv_max > 0 else 0)
            if rs:
                rs.clear()
                norm = [lv / lv_max if lv_max > 0 else 0 for lv in lvs]
                deltas = [norm[i] - norm[i - 1] for i in range(1, len(norm))]
                d_max = max(deltas) if deltas else 1.0
                for i, d in enumerate(deltas):
                    rs.append(points[i + 1]["level"], d / d_max if d_max > 0 else 0)
        n = sum(len(p) for p in result.get("gamma", {}).values())
        self._status_lbl.setText(f"Done — {n} measurements")


# ── Color sub-panel ────────────────────────────────────────────────────────────

class ColorSubPanel(QWidget):
    """Color-patch-only module measurement: CIE u'v', color accuracy xy, ΔE2000."""

    def __init__(self, engine: MeasurementEngine,
                 parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._engine = engine
        self._worker: Optional[MeasurementWorker] = None
        self._de_values: Dict[str, float] = {}

        root = QVBoxLayout(self)
        root.setSpacing(4); root.setContentsMargins(0, 0, 0, 0)

        # Controls
        ctrl = QHBoxLayout(); ctrl.setSpacing(6)
        title = QLabel("Chromaticity / Color Accuracy")
        title.setStyleSheet("font-size:13px;font-weight:bold;")
        ctrl.addWidget(title)
        self._hdr_check = QCheckBox("HDR"); ctrl.addWidget(self._hdr_check)
        ctrl.addStretch()
        self._btn_run = QPushButton("▶  Start")
        self._btn_run.setObjectName("primary")
        self._btn_run.clicked.connect(self._run)
        self._btn_stop = QPushButton("■  Stop")
        self._btn_stop.setObjectName("danger")
        self._btn_stop.setEnabled(False)
        self._btn_stop.clicked.connect(self._stop)
        self._btn_clear = QPushButton("Clear")
        self._btn_clear.clicked.connect(self._clear)
        for btn in (self._btn_run, self._btn_stop, self._btn_clear):
            ctrl.addWidget(btn)
        root.addLayout(ctrl)

        prog_row = QHBoxLayout(); prog_row.setSpacing(6)
        self._progress = QProgressBar(); self._progress.setFixedHeight(5)
        prog_row.addWidget(self._progress, stretch=1)
        self._status_lbl = QLabel("Idle")
        self._status_lbl.setObjectName("muted")
        self._status_lbl.setFixedWidth(240)
        prog_row.addWidget(self._status_lbl)
        root.addLayout(prog_row)

        # CIE u'v' + Color Accuracy xy charts side by side
        chart_row = QHBoxLayout(); chart_row.setSpacing(2)

        self._cie_chart, self._cie_xax, self._cie_yax = _make_chart(
            "CIE 1976 u'v'", "u'", 0.0, 0.7, "v'", 0.0, 0.65, tick_x=8, tick_y=7)
        self._cie_scatter: Optional[QScatterSeries] = None
        chart_row.addWidget(_chart_view(self._cie_chart, 200), stretch=1)

        self._ca_chart, self._ca_xax, self._ca_yax = _make_chart(
            "Color Accuracy xy", "x", 0.0, 0.8, "y", 0.0, 0.9, tick_x=9, tick_y=10)
        self._ca_scatter: Optional[QScatterSeries] = None
        self._ca_target: Optional[QScatterSeries] = None
        chart_row.addWidget(_chart_view(self._ca_chart, 200), stretch=1)

        root.addLayout(chart_row, stretch=1)

        # ΔE2000 table
        self._de_table = QTableWidget(0, 3)
        self._de_table.setHorizontalHeaderLabels(["Color", "ΔE2000", "Pass(<3)"])
        self._de_table.verticalHeader().hide()
        self._de_table.verticalHeader().setDefaultSectionSize(18)
        self._de_table.setAlternatingRowColors(True)
        self._de_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self._de_table.setSelectionMode(QTableWidget.SelectionMode.NoSelection)
        self._de_table.horizontalHeader().setStretchLastSection(True)
        self._de_table.setFixedHeight(160)
        root.addWidget(self._de_table)

    def _run(self) -> None:
        self._clear()
        self._btn_run.setEnabled(False); self._btn_stop.setEnabled(True)
        self._status_lbl.setText("Measuring…")
        from core.sequences.module_measure import BT709_REF_UV
        self._worker = MeasurementWorker(
            self._engine, "module_measure",
            is_hdr=self._hdr_check.isChecked(),
            gamma_channels=["W"],
            gamma_steps=[255],
            ref_uv=dict(BT709_REF_UV),
            run_gamma=False,
            run_colors=True,
        )
        self._worker.progress.connect(self._on_progress)
        self._worker.succeeded.connect(self._on_finished)
        self._worker.error.connect(lambda m: (
            QMessageBox.critical(self, "Error", m),
            self._btn_run.setEnabled(True),
            self._btn_stop.setEnabled(False),
        ))
        wire_worker_cleanup(self._worker, self, "_worker")
        self._worker.start()

    def _stop(self) -> None:
        self._engine.stop_sequence()
        if self._worker:
            self._worker.requestInterruption()

    def _clear(self) -> None:
        for s in [self._cie_scatter, self._ca_scatter, self._ca_target]:
            if s is not None:
                try:
                    self._cie_chart.removeSeries(s)
                    self._ca_chart.removeSeries(s)
                except Exception:
                    pass
        self._cie_scatter = self._ca_scatter = self._ca_target = None
        self._de_values.clear()
        self._de_table.setRowCount(0)
        self._progress.setValue(0)
        self._status_lbl.setText("Cleared")

    @Slot(str, float, object)
    def _on_progress(self, step: str, pct: float, data: Any) -> None:
        self._progress.setValue(int(pct * 100))
        if step == "module_color" and isinstance(data, dict):
            self._status_lbl.setText(f"Color {data.get('name','')} measuring ({int(pct*100)}%)")

    @Slot(object)
    def _on_finished(self, result: Any) -> None:
        self._btn_run.setEnabled(True); self._btn_stop.setEnabled(False)
        self._progress.setValue(100)
        if not isinstance(result, dict):
            return
        color_results: Dict[str, MeasureResult] = result.get("colors", {})

        # CIE u'v' scatter
        sc = QScatterSeries()
        sc.setMarkerSize(9.0)
        sc.setMarkerShape(QScatterSeries.MarkerShape.MarkerShapeCircle)
        for name in _PATCH_ORDER:
            r = color_results.get(name)
            if r:
                sc.append(r.u_prime, r.v_prime)
        self._cie_chart.addSeries(sc)
        sc.attachAxis(self._cie_xax); sc.attachAxis(self._cie_yax)
        self._cie_scatter = sc

        # CA xy scatter (measured)
        msc = QScatterSeries(); msc.setMarkerSize(9.0)
        msc.setMarkerShape(QScatterSeries.MarkerShape.MarkerShapeCircle)
        tsc = QScatterSeries(); tsc.setMarkerSize(7.0)
        tsc.setMarkerShape(QScatterSeries.MarkerShape.MarkerShapeRectangle)
        tsc.setColor(QColor("#000000")); tsc.setBorderColor(QColor("#000000"))
        for name in _PATCH_ORDER:
            r = color_results.get(name)
            if r:
                msc.append(r.x, r.y)
                msc.setColor(QColor(_PATCH_COLORS.get(name, "#888888")))
            ref = _BT709_REF_XY.get(name)
            if ref:
                tsc.append(ref[0], ref[1])
        self._ca_chart.addSeries(msc); msc.attachAxis(self._ca_xax); msc.attachAxis(self._ca_yax)
        self._ca_chart.addSeries(tsc); tsc.attachAxis(self._ca_xax); tsc.attachAxis(self._ca_yax)
        self._ca_scatter = msc; self._ca_target = tsc

        # ΔE2000 table
        self._de_table.setRowCount(0)
        self._de_values = {}
        for name in _PATCH_ORDER:
            r = color_results.get(name)
            ref = _BT709_REF_XY.get(name)
            if r and ref:
                try:
                    de = _calc_de2000(r, ref[0], ref[1])
                    self._de_values[name] = de
                except Exception:
                    de = 0.0
                ok = de <= 3.0
                row = self._de_table.rowCount()
                self._de_table.insertRow(row)
                it_c = _titem(name)
                it_c.setForeground(QColor(_PATCH_COLORS.get(name, "#333")))
                it_de = _titem(f"{de:.2f}")
                it_de.setForeground(QColor("#1a7040" if ok else "#c02020"))
                it_p = _titem("PASS" if ok else "FAIL")
                it_p.setForeground(QColor("#1a7040" if ok else "#c02020"))
                self._de_table.setItem(row, 0, it_c)
                self._de_table.setItem(row, 1, it_de)
                self._de_table.setItem(row, 2, it_p)

        if self._de_values:
            vals = list(self._de_values.values())
            self._status_lbl.setText(
                f"Done — ΔE Avg={sum(vals)/len(vals):.2f}  Max={max(vals):.2f}"
            )
        else:
            self._status_lbl.setText("Done")


# ── Calman Sweep panel ───────────────────────────────────────────────────────

_CALMAN_PATCH_COLORS = {
    "R": "#e74c3c", "G": "#27ae60", "B": "#2980e8",
    "C": "#1abc9c", "M": "#9b59b6", "Y": "#c8a000",
}
_DE_PASS_THRESHOLD = 3.0


class _CalmanSweepPanel(QWidget):
    """Calman saturation sweep controls + results table.

    30 patches: R/G/B/C/M/Y × 20/40/60/80/100 %
    """

    def __init__(self, engine: MeasurementEngine,
                 parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._engine = engine
        self._worker: Optional[MeasurementWorker] = None
        self._results: Dict[str, Any] = {}

        self._gamut_data: Dict[str, MeasureResult] = {}

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 2, 0, 2)
        outer.setSpacing(3)

        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet("color:#263058;")
        outer.addWidget(sep)

        # ── Controls row ──────────────────────────────────────────────
        ctrl = QHBoxLayout(); ctrl.setSpacing(6)
        hdr = QLabel("Calman Color Sweep")
        hdr.setStyleSheet("font-size:12px;font-weight:bold;")
        ctrl.addWidget(hdr)
        ctrl.addSpacing(8)

        self._gamut_status = QLabel("⚠ Run Module All measurement first.")
        self._gamut_status.setStyleSheet("font-size:10px;color:#f0b040;")
        ctrl.addWidget(self._gamut_status)
        ctrl.addStretch()

        # Per-colour ΔE76 avg summary labels
        self._summary_lbls: Dict[str, QLabel] = {}
        for color in CALMAN_COLOR_ORDER:
            lbl2 = QLabel(f"{color}: —")
            lbl2.setStyleSheet(
                f"color:{_CALMAN_PATCH_COLORS[color]};"
                "font-size:10px;font-weight:bold;"
            )
            self._summary_lbls[color] = lbl2
            ctrl.addWidget(lbl2)

        ctrl.addSpacing(8)
        self._btn_run = QPushButton("▶ Calman Sweep")
        self._btn_run.setObjectName("primary")
        self._btn_run.setEnabled(False)
        self._btn_run.clicked.connect(self._run)
        self._btn_stop = QPushButton("■")
        self._btn_stop.setObjectName("danger")
        self._btn_stop.setEnabled(False)
        self._btn_stop.setFixedWidth(32)
        self._btn_stop.clicked.connect(self._stop)
        ctrl.addWidget(self._btn_run)
        ctrl.addWidget(self._btn_stop)
        outer.addLayout(ctrl)

        # ── Progress ──────────────────────────────────────────────────
        prog_row = QHBoxLayout(); prog_row.setSpacing(6)
        self._progress = QProgressBar(); self._progress.setFixedHeight(5)
        prog_row.addWidget(self._progress, stretch=1)
        self._status_lbl = QLabel("Idle")
        self._status_lbl.setObjectName("muted")
        self._status_lbl.setFixedWidth(260)
        prog_row.addWidget(self._status_lbl)
        outer.addLayout(prog_row)

        # ── Results table ─────────────────────────────────────────────
        # Columns: Color | Sat% | Lv | x | y | L* ref | a* ref | b* ref | ΔE76 | Pass
        self._table = QTableWidget(0, 10)
        self._table.setHorizontalHeaderLabels([
            "Color", "Sat%", "Lv", "x", "y",
            "L* ref", "a* ref", "b* ref", "ΔE76", "Pass(<3)",
        ])
        self._table.verticalHeader().hide()
        self._table.verticalHeader().setDefaultSectionSize(18)
        self._table.setAlternatingRowColors(True)
        self._table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self._table.setSelectionMode(QTableWidget.SelectionMode.NoSelection)
        self._table.setStyleSheet("font-size:10px;")
        self._table.horizontalHeader().setStretchLastSection(True)
        self._table.setFixedHeight(200)
        outer.addWidget(self._table)

    # ── Gamut data injection (from ModulePanel) ───────────────────────

    def set_gamut_data(self, colors: Dict[str, MeasureResult]) -> None:
        """Receive measured R/G/B/W primaries from ModulePanel._on_finished."""
        self._gamut_data = colors
        has_rgbw = all(c in colors for c in ("R", "G", "B", "W"))
        if has_rgbw:
            w = colors["W"]
            self._gamut_status.setText(
                f"✓ Measured gamut loaded  (W: x={w.x:.4f}, y={w.y:.4f})"
            )
            self._gamut_status.setStyleSheet("font-size:10px;color:#1dd9a0;")
            self._btn_run.setEnabled(True)
        else:
            self._gamut_status.setText("⚠ Run Module All measurement first.")
            self._gamut_status.setStyleSheet("font-size:10px;color:#f0b040;")
            self._btn_run.setEnabled(False)

    # ── Run / Stop ────────────────────────────────────────────────────

    def _run(self) -> None:
        if not all(c in self._gamut_data for c in ("R", "G", "B", "W")):
            QMessageBox.warning(self, "No Data",
                                "Run Module All measurement first.")
            return
        self._table.setRowCount(0)
        for lbl in self._summary_lbls.values():
            lbl.setText(lbl.text().split(":")[0] + ": —")
        self._progress.setValue(0)
        self._btn_run.setEnabled(False)
        self._btn_stop.setEnabled(True)
        self._status_lbl.setText("Measuring…")

        self._worker = MeasurementWorker(
            self._engine, "calman_sweep",
            measured_colors=dict(self._gamut_data),
        )
        self._worker.progress.connect(self._on_progress)
        self._worker.succeeded.connect(self._on_finished)
        self._worker.error.connect(lambda m: (
            QMessageBox.critical(self, "Error", m),
            self._btn_run.setEnabled(True),
            self._btn_stop.setEnabled(False),
        ))
        wire_worker_cleanup(self._worker, self, "_worker")
        self._worker.start()

    def _stop(self) -> None:
        self._engine.stop_sequence()
        if self._worker:
            self._worker.requestInterruption()

    # ── Progress ──────────────────────────────────────────────────────

    @Slot(str, float, object)
    def _on_progress(self, step: str, pct: float, data: Any) -> None:
        self._progress.setValue(int(pct * 100))
        if step == "calman_sweep" and isinstance(data, dict):
            color = data.get("color", "")
            sat   = data.get("sat", 0)
            r     = data.get("result")
            de    = data.get("de76", 0.0)
            if r is not None:
                self._add_row(color, sat, r, de)
                self._status_lbl.setText(
                    f"{color} {sat}%  Lv={r.Lv:.2f}  ΔE76={de:.2f}  ({int(pct*100)}%)"
                )

    def _add_row(self, color: str, sat: int,
                 r: MeasureResult, de: float) -> None:
        ok = de <= _DE_PASS_THRESHOLD
        row = self._table.rowCount()
        self._table.insertRow(row)

        # Compute ref Lab for display using linear XYZ mixing (Calman method)
        ref_L = ref_a = ref_b = 0.0
        try:
            from core.sequences.calman_sweep import (
                calc_target_xyz, xyz_to_lab, meas_to_xyz,
            )
            w_meas = self._gamut_data.get("W")
            if w_meas is not None:
                white_xyz = meas_to_xyz(w_meas)
                ref_xyz = calc_target_xyz(color, sat, self._gamut_data)
                if ref_xyz is not None:
                    ref_L, ref_a, ref_b = xyz_to_lab(*ref_xyz, *white_xyz)
        except Exception:
            pass

        vals = [color, f"{sat}%",
                f"{r.Lv:.3f}", f"{r.x:.4f}", f"{r.y:.4f}",
                f"{ref_L:.1f}", f"{ref_a:.1f}", f"{ref_b:.1f}",
                f"{de:.2f}", "PASS" if ok else "FAIL"]
        for ci, val in enumerate(vals):
            it = _titem(val)
            if ci == 0:
                it.setForeground(QColor(_CALMAN_PATCH_COLORS.get(color, "#aaa")))
            if ci == 8:
                it.setForeground(QColor("#1dd9a0" if ok else "#f05050"))
            if ci == 9:
                it.setForeground(QColor("#1dd9a0" if ok else "#f05050"))
            self._table.setItem(row, ci, it)
        self._table.scrollToBottom()

    # ── Finished ──────────────────────────────────────────────────────

    @Slot(object)
    def _on_finished(self, result: Any) -> None:
        self._results = result or {}
        self._btn_run.setEnabled(True)
        self._btn_stop.setEnabled(False)
        self._progress.setValue(100)
        if not isinstance(result, dict):
            return
        for color in CALMAN_COLOR_ORDER:
            color_data = result.get(color, {})
            des = [v["de76"] for v in color_data.values() if "de76" in v]
            lbl = self._summary_lbls.get(color)
            if lbl and des:
                avg = sum(des) / len(des)
                lbl.setText(f"{color}: {avg:.2f}")
                ok = avg <= _DE_PASS_THRESHOLD
                lbl.setStyleSheet(
                    f"color:{'#1dd9a0' if ok else '#f05050'};"
                    "font-size:10px;font-weight:bold;"
                )
        total_des = [
            v["de76"]
            for cd in result.values() if isinstance(cd, dict)
            for v in cd.values() if isinstance(v, dict) and "de76" in v
        ]
        if total_des:
            avg_all = sum(total_des) / len(total_des)
            self._status_lbl.setText(
                f"Done — 30 patches  ΔE76 Avg={avg_all:.2f}  Max={max(total_des):.2f}"
            )


# ── Main panel ────────────────────────────────────────────────────────────────

class ModulePanel(QWidget):
    """LED module measurement panel."""

    gamut_data_ready = Signal(dict)   # emitted with color_results after full measurement

    def __init__(self, engine: MeasurementEngine, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._engine = engine
        self._worker: Optional[MeasurementWorker] = None
        self._result: Dict[str, Any] = {}
        self._de_values: Dict[str, float] = {}
        self._live_gamma: Dict[str, List[dict]] = {}
        self._live_colors: Dict[str, Any] = {}

        root = QVBoxLayout(self)
        root.setSpacing(4)
        root.setContentsMargins(6, 4, 6, 4)

        # ── Controls bar ──────────────────────────────────────────────
        ctrl = QHBoxLayout(); ctrl.setSpacing(6)
        title = QLabel("Module Measurement")
        title.setStyleSheet("font-size:13px;font-weight:bold;")
        ctrl.addWidget(title)
        self._hdr_check = QCheckBox("HDR")
        ctrl.addWidget(self._hdr_check)
        ctrl.addSpacing(8)
        ctrl.addWidget(QLabel("Gamma Channel:"))
        self._ch_w = QCheckBox("W"); self._ch_w.setChecked(True)
        self._ch_r = QCheckBox("R"); self._ch_r.setChecked(True)
        self._ch_g = QCheckBox("G"); self._ch_g.setChecked(True)
        self._ch_b = QCheckBox("B"); self._ch_b.setChecked(True)
        for cb in (self._ch_w, self._ch_r, self._ch_g, self._ch_b):
            ctrl.addWidget(cb)
        ctrl.addSpacing(8)
        ctrl.addWidget(QLabel("Step:"))
        self._step_combo = QComboBox()
        self._step_combo.addItems(["10-step (default)", "17-step", "5-step"])
        self._step_combo.setFixedWidth(130)
        ctrl.addWidget(self._step_combo)
        ctrl.addStretch()
        self._btn_run = QPushButton("▶  Start")
        self._btn_run.setObjectName("primary")
        self._btn_run.clicked.connect(self._run)
        self._btn_stop = QPushButton("■  Stop")
        self._btn_stop.setObjectName("danger")
        self._btn_stop.setEnabled(False)
        self._btn_stop.clicked.connect(self._stop)
        self._btn_clear = QPushButton("Clear")
        self._btn_clear.clicked.connect(self._clear)
        self._btn_export = QPushButton("Excel")
        self._btn_export.clicked.connect(self._export)
        for btn in (self._btn_run, self._btn_stop, self._btn_clear, self._btn_export):
            ctrl.addWidget(btn)
        root.addLayout(ctrl)

        # Progress
        prog_row = QHBoxLayout(); prog_row.setSpacing(6)
        self._progress = QProgressBar(); self._progress.setFixedHeight(6)
        prog_row.addWidget(self._progress, stretch=1)
        self._status_lbl = QLabel("Idle")
        self._status_lbl.setObjectName("muted")
        self._status_lbl.setFixedWidth(240)
        prog_row.addWidget(self._status_lbl)
        root.addLayout(prog_row)

        # ── Main content: sidebar + charts ────────────────────────────
        splitter = QSplitter(Qt.Orientation.Horizontal)
        self._sidebar = _LeftSidebar()
        self._charts = _ChartArea()
        splitter.addWidget(self._sidebar)
        splitter.addWidget(self._charts)
        splitter.setSizes([175, 900])
        splitter.setCollapsible(0, False)
        root.addWidget(splitter, stretch=1)

        # ── Calman saturation sweep ───────────────────────────────────
        self._calman_panel = _CalmanSweepPanel(engine, self)
        root.addWidget(self._calman_panel)

    # ── Helpers ───────────────────────────────────────────────────────

    def _get_gamma_steps(self) -> List[int]:
        text = self._step_combo.currentText()
        if "17" in text:
            steps = list(range(0, 256, 16))
            if steps[-1] != 255:
                steps.append(255)
            return steps
        if "5" in text:
            return [0, 64, 128, 192, 255]
        return list(DEFAULT_GAMMA_STEPS)

    def _get_gamma_channels(self) -> List[str]:
        return [ch for cb, ch in [(self._ch_w, "W"), (self._ch_r, "R"),
                                  (self._ch_g, "G"), (self._ch_b, "B")]
                if cb.isChecked()]

    # ── Run / Stop / Clear ────────────────────────────────────────────

    def _run(self) -> None:
        self._clear()
        channels = self._get_gamma_channels()
        if not channels:
            QMessageBox.warning(self, "Settings Error", "Select at least one gamma channel.")
            return
        self._btn_run.setEnabled(False)
        self._btn_stop.setEnabled(True)
        self._status_lbl.setText("Measuring…")
        self._worker = MeasurementWorker(
            self._engine, "module_measure",
            is_hdr=self._hdr_check.isChecked(),
            gamma_channels=channels,
            gamma_steps=self._get_gamma_steps(),
            ref_uv=dict(BT709_REF_UV),
        )
        self._worker.progress.connect(self._on_progress)
        self._worker.succeeded.connect(self._on_finished)
        self._worker.error.connect(lambda m: (
            QMessageBox.critical(self, "Error", m),
            self._btn_run.setEnabled(True),
            self._btn_stop.setEnabled(False),
        ))
        wire_worker_cleanup(self._worker, self, "_worker")
        self._worker.start()

    def _stop(self) -> None:
        self._engine.stop_sequence()
        if self._worker:
            self._worker.requestInterruption()

    def _clear(self) -> None:
        self._result.clear()
        self._de_values.clear()
        self._live_gamma.clear()
        self._live_colors.clear()
        self._sidebar.clear()
        self._charts.clear()
        self._progress.setValue(0)
        self._status_lbl.setText("Cleared")

    # ── Progress ──────────────────────────────────────────────────────

    @Slot(str, float, object)
    def _on_progress(self, step: str, pct: float, data: Any) -> None:
        self._progress.setValue(int(pct * 100))
        if step == "module_gamma" and isinstance(data, dict):
            ch = data.get("channel", "")
            level = data.get("level", 0)
            result = data.get("result")
            if result and ch:
                self._status_lbl.setText(
                    f"Gamma {ch} L={level}  Lv={result.Lv:.2f}  ({int(pct*100)}%)"
                )
                pts = self._live_gamma.setdefault(ch, [])
                pts.append({"level": level, "Lv": result.Lv, "result": result})
                self._charts.update_gamma_channel(ch, pts)
        elif step == "module_color" and isinstance(data, dict):
            name = data.get("name", "")
            result = data.get("result")
            self._status_lbl.setText(f"Color {name} measuring  ({int(pct*100)}%)")
            if result and name:
                self._live_colors[name] = result
                self._charts.update_colors(self._live_colors)

    # ── Finished ──────────────────────────────────────────────────────

    @Slot(object)
    def _on_finished(self, result: Any) -> None:
        self._result = result or {}
        self._btn_run.setEnabled(True)
        self._btn_stop.setEnabled(False)
        self._progress.setValue(100)
        if not isinstance(result, dict):
            return

        gamma_data: Dict[str, List[dict]] = result.get("gamma", {})
        color_results: Dict[str, MeasureResult] = result.get("colors", {})
        gamut_stats: Dict[str, float] = result.get("gamut_stats", {})

        # Gamma charts + sidebar
        for ch, points in gamma_data.items():
            self._charts.update_gamma_channel(ch, points)
            gammas = [p["gamma"] for p in points if p.get("gamma") is not None]
            avg_g = sum(gammas) / len(gammas) if gammas else None
            self._sidebar.update_gamma(ch, avg_g)

        # Color patch charts
        self._charts.update_colors(color_results)

        # ΔE2000 per patch
        self._de_values = {}
        for name in _PATCH_ORDER:
            mres = color_results.get(name)
            ref = _BT709_REF_XY.get(name)
            if mres and ref:
                try:
                    de = _calc_de2000(mres, ref[0], ref[1])
                    self._de_values[name] = de
                except Exception:
                    pass

        # Update sidebar
        w = color_results.get("W")
        if w:
            self._sidebar.update_module_white(w.Lv, w.cct or 0.0, w.duv or 0.0,
                                              (w.x, w.y))
        dci = gamut_stats.get("dci_overlap", 0.0)
        self._sidebar.update_gamut(dci)
        if self._de_values:
            vals = list(self._de_values.values())
            self._sidebar.update_de2000(sum(vals) / len(vals), max(vals))
            self._charts.update_de_stats(self._de_values)

        n = sum(len(p) for p in gamma_data.values()) + len(color_results)
        self._status_lbl.setText(f"Done — {n} measurements")

        # Emit for CalmanSweepPanel to receive measured primaries
        if color_results:
            self.gamut_data_ready.emit(color_results)

    # ── Export ────────────────────────────────────────────────────────

    def _export(self) -> None:
        if not self._result:
            QMessageBox.information(self, "Notice", "No data to save.")
            return
        brand = self._engine.brand or "brand"
        model = self._engine.model_name or "model"
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Excel", f"module_measure_{brand}_{model}.xlsx",
            "Excel (*.xlsx)")
        if not path:
            return
        try:
            import openpyxl
            wb = openpyxl.Workbook()
            ws_g = wb.active
            assert ws_g is not None
            ws_g.title = "Gamma"
            ws_g.append(["Channel", "Level", "Lv (cd/m²)", "Norm Lv", "γ", "u'", "v'",
                          "CCT (K)", "Duv"])
            for ch, points in self._result.get("gamma", {}).items():
                lv_max = max((p["Lv"] for p in points), default=0.0)
                for p in points:
                    r: MeasureResult = p["result"]
                    norm = p["Lv"] / lv_max if lv_max > 0 else 0.0
                    ws_g.append([ch, p["level"], round(p["Lv"], 4), round(norm, 4),
                                 p.get("gamma"), round(r.u_prime, 4), round(r.v_prime, 4),
                                 round(r.cct, 0) if r.cct else None,
                                 round(r.duv, 5) if r.cct else None])
            ws_c = wb.create_sheet("Color")
            ws_c.append(["Color", "Lv", "x", "y", "u'", "v'", "CCT (K)", "Duv", "ΔE2000"])
            color_results = self._result.get("colors", {})
            for name in _PATCH_ORDER:
                r = color_results.get(name)
                if r:
                    de = self._de_values.get(name)
                    ws_c.append([name, round(r.Lv, 3), round(r.x, 4), round(r.y, 4),
                                 round(r.u_prime, 4), round(r.v_prime, 4),
                                 round(r.cct, 0) if r.cct else None,
                                 round(r.duv, 5) if r.cct else None,
                                 round(de, 3) if de is not None else None])
            if self._de_values:
                vals = list(self._de_values.values())
                ws_c.append(["AVG", "", "", "", "", "", "", "",
                              round(sum(vals) / len(vals), 3)])
                ws_c.append(["MAX", "", "", "", "", "", "", "", round(max(vals), 3)])
            wb.save(path)
            QMessageBox.information(self, "Saved", f"Saved:\n{path}")
        except Exception:
            import traceback
            QMessageBox.critical(self, "Save Error", traceback.format_exc())
