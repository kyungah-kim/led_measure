from __future__ import annotations

import os
import statistics
from datetime import datetime
from typing import Any, Dict, List, Optional

import openpyxl

from PySide6.QtCharts import QChart, QChartView, QLineSeries, QScatterSeries, QValueAxis
from PySide6.QtCore import QMargins, QRectF, Qt, Slot
from PySide6.QtGui import QColor, QPainter
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QColorDialog,
    QComboBox,
    QFileDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSplitter,
    QSpinBox,
    QStackedWidget,
    QStatusBar,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
    QLineEdit,
)

from core.engine import MeasurementEngine
from core.equipment.base import MeasureResult
from core.export import ExcelExporter
from core.gamut_utils import DCI_P3_UV, BT2020_UV, calc_gamut_stats
from .worker import ConnectWorker, MeasurementWorker, wire_worker_cleanup


def _save_all_session(engine: MeasurementEngine) -> str:
    """모든 세션 데이터를 {brand}_{model}_all.xlsx 한 파일에 자동 저장."""
    if not engine.auto_save_dir:
        return ""
    brand = engine.brand or "brand"
    model = engine.model_name or "model"
    path = os.path.join(engine.auto_save_dir, f"{brand}_{model}_all.xlsx")
    try:
        ExcelExporter().export_all_session(
            brand=brand, model=model,
            session_swing=engine.session_swing,
            session_loading=engine.session_loading,
            session_gamut=engine.session_gamut,
            session_contrast=engine.session_contrast,
            file_path=path,
        )
    except Exception as e:
        print(f"[_save_all_session] 오류: {e}")
    return path

_DARK_STYLE = """
QMainWindow, QWidget { background: #f5f6fa; color: #1a1d2e; font-size: 13px; }
QGroupBox { border: 1px solid #d0d3e0; border-radius: 6px; margin-top: 8px; padding: 8px;
            background: #ffffff; }
QGroupBox::title { color: #6b7080; font-size: 11px; text-transform: uppercase; }

QPushButton { background: #ffffff; border: 1px solid #c8ccd8; border-radius: 5px;
              padding: 6px 14px; color: #1a1d2e; font-weight: 500; }
QPushButton:hover { background: #eef0f8; border-color: #a8b0cc; }
QPushButton:pressed { background: #dde2f0; }
QPushButton#primary { background: #4f8ef7; border-color: #4f8ef7; color: white; font-weight: bold; }
QPushButton#primary:hover { background: #3a7ae8; border-color: #3a7ae8; }
QPushButton#primary:pressed { background: #2d6fd6; }
QPushButton#danger  { background: #e74c3c; border-color: #e74c3c; color: white; font-weight: bold; }
QPushButton#danger:hover  { background: #d44030; }
QPushButton#warning { background: #e67e22; border-color: #e67e22; color: white; font-weight: bold; }
QPushButton#warning:hover { background: #cf6d1a; }
QPushButton#success { background: #27ae60; border-color: #27ae60; color: white; font-weight: bold; }
QPushButton:disabled { color: #aab0c0; border-color: #dde0ea; background: #f5f6fa; }

QComboBox { background: #ffffff; border: 1px solid #c8ccd8; border-radius: 4px;
            padding: 5px 8px; color: #1a1d2e; }
QComboBox:hover { border-color: #4f8ef7; }
QComboBox::drop-down { border: none; width: 20px; }
QComboBox QAbstractItemView { background: #ffffff; border: 1px solid #c8ccd8;
                               selection-background-color: #eef0f8; }
QLineEdit { background: #ffffff; border: 1px solid #c8ccd8; border-radius: 4px;
            padding: 5px 8px; color: #1a1d2e; }
QLineEdit:focus { border-color: #4f8ef7; }

QListWidget { background: #ffffff; border: 1px solid #d0d3e0; border-radius: 4px; }
QListWidget::item { padding: 6px 12px; }
QListWidget::item:hover { background: #f0f2f8; }
QListWidget::item:selected { background: rgba(79,142,247,0.12); color: #2d6fd6;
                              border-left: 3px solid #4f8ef7; }

QProgressBar { background: #e8eaf0; border: 1px solid #d0d3e0; border-radius: 3px; height: 8px; }
QProgressBar::chunk { background: #4f8ef7; border-radius: 3px; }

QTableWidget { background: #ffffff; gridline-color: #eaecf4; border: 1px solid #d0d3e0;
               alternate-background-color: #f8f9fc; }
QTableWidget::item:hover { background: #eef0f8; }
QTableWidget::item:selected { background: rgba(79,142,247,0.15); color: #1a1d2e; }
QHeaderView::section { background: #f0f2f8; color: #6b7080; border: none;
                       border-bottom: 2px solid #d0d3e0; padding: 6px 8px;
                       font-size: 11px; font-weight: bold; letter-spacing: 0.03em; }
QHeaderView::section:hover { background: #e8eaf4; }

QCheckBox { spacing: 6px; }
QCheckBox::indicator { width: 15px; height: 15px; border-radius: 3px;
                       border: 1px solid #c8ccd8; background: #ffffff; }
QCheckBox::indicator:checked { background: #4f8ef7; border-color: #4f8ef7; }

QSpinBox { background: #ffffff; border: 1px solid #c8ccd8; border-radius: 4px;
           padding: 4px 6px; color: #1a1d2e; }
QSpinBox:focus { border-color: #4f8ef7; }

QSplitter::handle { background: #d0d3e0; }
QSplitter::handle:horizontal { width: 4px; }
QSplitter::handle:vertical   { height: 4px; }
QSplitter::handle:hover { background: #4f8ef7; }

QScrollBar:vertical { background: #f5f6fa; width: 8px; border-radius: 4px; }
QScrollBar::handle:vertical { background: #c8ccd8; border-radius: 4px; min-height: 20px; }
QScrollBar::handle:vertical:hover { background: #a0a8c0; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0px; }
QScrollBar:horizontal { background: #f5f6fa; height: 8px; border-radius: 4px; }
QScrollBar::handle:horizontal { background: #c8ccd8; border-radius: 4px; min-width: 20px; }
QScrollBar::handle:horizontal:hover { background: #a0a8c0; }
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal { width: 0px; }

QLabel#status_ok  { color: #27ae60; font-weight: bold; }
QLabel#status_err { color: #e74c3c; font-weight: bold; }
QLabel#muted      { color: #8890a8; font-size: 12px; }
"""


class ConnectionPanel(QGroupBox):
    def __init__(self, engine: MeasurementEngine, parent: QWidget | None = None) -> None:
        super().__init__("장비 연결", parent)
        self._engine = engine
        self._connect_worker: Optional[ConnectWorker] = None
        self._reset_worker: Optional[ConnectWorker] = None
        self._build_ui()

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setSpacing(4)
        root.setContentsMargins(6, 6, 6, 4)

        # ── 상단 행: 기본 정보 / CA 색채휘도계 / 패턴 제너레이터 ──────────────
        top_row = QHBoxLayout()
        top_row.setSpacing(6)

        # Brand / Model
        info_box = QGroupBox("기본 정보")
        info_form = QFormLayout(info_box)
        info_form.setContentsMargins(8, 4, 8, 4)
        info_form.setSpacing(4)
        self._brand_edit = QLineEdit("입력")
        self._model_edit = QLineEdit("입력")
        self._brand_edit.setMinimumWidth(140)
        self._model_edit.setMinimumWidth(140)
        self._brand_edit.textChanged.connect(self._sync_info)
        self._model_edit.textChanged.connect(self._sync_info)
        info_form.addRow("Brand:", self._brand_edit)
        info_form.addRow("Model:", self._model_edit)
        top_row.addWidget(info_box, stretch=1)

        # 포트 스캔 공통 버튼
        btn_scan = QPushButton("포트 스캔")
        btn_scan.setToolTip("연결된 시리얼 포트를 다시 스캔합니다")
        btn_scan.clicked.connect(self._scan_ports)

        # Meter
        meter_box = QGroupBox("CA 색채휘도계")
        meter_layout = QHBoxLayout(meter_box)
        meter_layout.setContentsMargins(8, 4, 8, 4)
        self._meter_port = QComboBox()
        self._meter_port.setEditable(True)
        self._meter_model = QComboBox()
        self._meter_model.addItems(["CA-410", "CA-310"])
        self._btn_meter = QPushButton("연결")
        self._btn_meter.setObjectName("primary")
        self._btn_meter.clicked.connect(self._connect_meter)
        self._btn_meter_dis = QPushButton("연결 해제")
        self._btn_meter_dis.setObjectName("danger")
        self._btn_meter_dis.setEnabled(False)
        self._btn_meter_dis.clicked.connect(self._disconnect_meter)
        self._meter_status = QLabel("미연결")
        self._meter_status.setObjectName("status_err")
        meter_layout.addWidget(QLabel("포트:"))
        meter_layout.addWidget(self._meter_port)
        meter_layout.addWidget(self._meter_model)
        meter_layout.addWidget(self._btn_meter)
        meter_layout.addWidget(self._btn_meter_dis)
        meter_layout.addWidget(self._meter_status)
        top_row.addWidget(meter_box, stretch=2)

        # Generator
        gen_box = QGroupBox("패턴 제너레이터")
        gen_layout = QHBoxLayout(gen_box)
        gen_layout.setContentsMargins(8, 4, 8, 4)
        self._gen_port = QComboBox()
        self._gen_port.setEditable(True)
        self._gen_model = QComboBox()
        self._gen_model.addItems(["VG-879", "VG-876"])
        self._btn_gen = QPushButton("연결")
        self._btn_gen.setObjectName("primary")
        self._btn_gen.clicked.connect(self._connect_generator)
        self._btn_gen_dis = QPushButton("연결 해제")
        self._btn_gen_dis.setObjectName("danger")
        self._btn_gen_dis.setEnabled(False)
        self._btn_gen_dis.clicked.connect(self._disconnect_generator)
        self._btn_gen_reset = QPushButton("장비 리셋")
        self._btn_gen_reset.setObjectName("warning")
        self._btn_gen_reset.setEnabled(False)
        self._btn_gen_reset.setToolTip("장비가 멈추었을 때 ENQ 재진입 + 컬러바 재로드로 복구합니다")
        self._btn_gen_reset.clicked.connect(self._reset_generator)
        self._gen_status = QLabel("미연결")
        self._gen_status.setObjectName("status_err")
        gen_layout.addWidget(QLabel("포트:"))
        gen_layout.addWidget(self._gen_port)
        gen_layout.addWidget(self._gen_model)
        gen_layout.addWidget(self._btn_gen)
        gen_layout.addWidget(self._btn_gen_dis)
        gen_layout.addWidget(self._btn_gen_reset)
        gen_layout.addWidget(self._gen_status)
        top_row.addWidget(gen_box, stretch=2)

        root.addLayout(top_row)

        # ── 하단 행: 포트 스캔 / 자동 저장 폴더 / 전체 해제 ──────────────
        bottom_row = QHBoxLayout()
        bottom_row.setSpacing(4)
        self._btn_dis_all = QPushButton("전체 해제")
        self._btn_dis_all.setObjectName("danger")
        self._btn_dis_all.clicked.connect(self._disconnect_all)
        bottom_row.addWidget(btn_scan)
        bottom_row.addSpacing(12)
        bottom_row.addWidget(QLabel("자동 저장 폴더:"))
        self._save_dir_edit = QLineEdit()
        self._save_dir_edit.setPlaceholderText("측정 완료 시 자동 저장할 폴더")
        self._save_dir_edit.setReadOnly(True)
        self._save_dir_edit.setMinimumWidth(220)
        bottom_row.addWidget(self._save_dir_edit, stretch=1)
        btn_folder = QPushButton("📁")
        btn_folder.setFixedWidth(32)
        btn_folder.clicked.connect(self._pick_save_dir)
        bottom_row.addWidget(btn_folder)
        bottom_row.addSpacing(12)
        bottom_row.addWidget(self._btn_dis_all)
        root.addLayout(bottom_row)

        # 시작 시 포트 목록 채우기 + 자동 저장 폴더 초기값 = 실행 폴더
        self._scan_ports()
        default_dir = os.path.dirname(os.path.abspath(__file__))
        self._engine.auto_save_dir = default_dir
        self._save_dir_edit.setText(default_dir)

    def _scan_ports(self) -> None:
        """시리얼 포트를 스캔해 두 콤보박스를 갱신한다."""
        import serial.tools.list_ports
        ports = sorted(p.device for p in serial.tools.list_ports.comports())

        for combo in (self._meter_port, self._gen_port):
            current = combo.currentText()
            combo.clear()
            combo.addItems(ports)
            idx = combo.findText(current)
            if idx >= 0:
                combo.setCurrentIndex(idx)

    def _pick_save_dir(self) -> None:
        folder = QFileDialog.getExistingDirectory(
            self, "자동 저장 폴더 선택",
            self._engine.auto_save_dir or os.path.expanduser("~"),
        )
        if folder:
            self._engine.auto_save_dir = folder
            self._save_dir_edit.setText(folder)

    def _sync_info(self) -> None:
        self._engine.brand = self._brand_edit.text().strip()
        self._engine.model_name = self._model_edit.text().strip()

    def _connect_meter(self) -> None:
        port = self._meter_port.currentText()
        model = self._meter_model.currentText()
        self._meter_status.setText("연결 중...")
        self._btn_meter.setEnabled(False)
        self._connect_worker = ConnectWorker(lambda: self._engine.connect_meter(port, model))
        self._connect_worker.succeeded.connect(self._on_meter_connected)
        self._connect_worker.error.connect(lambda msg: self._on_connect_error("CA 연결 오류", msg, self._btn_meter))
        wire_worker_cleanup(self._connect_worker, self, '_connect_worker')
        self._connect_worker.start()

    def _on_meter_connected(self) -> None:
        self._sync_info()
        ident = getattr(self._engine.meter, "ident", None)
        label = f"연결됨  {ident}" if ident else "연결됨"
        self._meter_status.setText(label)
        self._meter_status.setStyleSheet("color:#1a9e50;font-weight:bold;")
        self._btn_meter.setEnabled(False)
        self._btn_meter_dis.setEnabled(True)

    def _disconnect_meter(self) -> None:
        try:
            if self._engine.meter and self._engine.meter.is_connected:
                self._engine.meter.disconnect()
            self._engine.meter = None
        except Exception:
            pass
        self._meter_status.setText("미연결")
        self._meter_status.setStyleSheet("color:#e74c3c;font-weight:bold;")
        self._btn_meter.setEnabled(True)
        self._btn_meter_dis.setEnabled(False)

    def _connect_generator(self) -> None:
        port = self._gen_port.currentText()
        model = self._gen_model.currentText()
        self._gen_status.setText("연결 중...")
        self._btn_gen.setEnabled(False)
        self._connect_worker = ConnectWorker(lambda: self._engine.connect_generator(port, model))
        self._connect_worker.succeeded.connect(self._on_gen_connected)
        self._connect_worker.error.connect(lambda msg: self._on_connect_error("VG 연결 오류", msg, self._btn_gen))
        wire_worker_cleanup(self._connect_worker, self, '_connect_worker')
        self._connect_worker.start()

    def _on_gen_connected(self) -> None:
        self._gen_status.setText("연결됨")
        self._gen_status.setStyleSheet("color:#1a9e50;font-weight:bold;")
        self._btn_gen.setEnabled(False)
        self._btn_gen_dis.setEnabled(True)
        self._btn_gen_reset.setEnabled(True)

    def _on_connect_error(self, title: str, msg: str, retry_btn: "QPushButton") -> None:
        retry_btn.setEnabled(True)
        QMessageBox.critical(self, title, msg)

    def _reset_generator(self) -> None:
        """장비 freeze 시 ENQ 재진입 + EXPDN4(2286,0) 컬러바 복귀."""
        gen = self._engine.generator
        if gen is None or not gen.is_connected:
            return
        self._btn_gen_reset.setEnabled(False)
        self._gen_status.setText("리셋 중...")
        worker = ConnectWorker(lambda: gen.reset())
        worker.succeeded.connect(self._on_gen_reset_done)
        worker.error.connect(lambda msg: (
            QMessageBox.critical(self, "VG 리셋 오류", msg),
            self._btn_gen_reset.setEnabled(True),
            self._gen_status.setText("연결됨"),
            self._gen_status.setStyleSheet("color:#1a9e50;font-weight:bold;"),
        ))
        wire_worker_cleanup(worker, self, '_reset_worker')
        worker.start()
        self._reset_worker = worker

    def _on_gen_reset_done(self) -> None:
        self._gen_status.setText("연결됨")
        self._gen_status.setStyleSheet("color:#1a9e50;font-weight:bold;")
        self._btn_gen_reset.setEnabled(True)

    def _disconnect_generator(self) -> None:
        try:
            if self._engine.generator and self._engine.generator.is_connected:
                self._engine.generator.disconnect()
            self._engine.generator = None
        except Exception:
            pass
        self._gen_status.setText("미연결")
        self._gen_status.setStyleSheet("color:#e74c3c;font-weight:bold;")
        self._btn_gen.setEnabled(True)
        self._btn_gen_dis.setEnabled(False)
        self._btn_gen_reset.setEnabled(False)

    def _disconnect_all(self) -> None:
        self._disconnect_meter()
        self._disconnect_generator()


# ---------------------------------------------------------------------------
# Center Alignment Panel
# ---------------------------------------------------------------------------

class CenterAlignPanel(QWidget):
    def __init__(self, engine: MeasurementEngine, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._engine = engine
        self._worker: Optional[MeasurementWorker] = None
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignmentFlag.AlignTop)

        title = QLabel("🎯 센터 맞추기")
        title.setStyleSheet("font-size:15px;font-weight:bold;")
        layout.addWidget(title)
        desc = QLabel("패턴 제너레이터에서 ABC 센터 정렬 패턴을 출력합니다.\n"
                      "측정기 렌즈가 화면 정중앙을 향하도록 조정 후 [OK]를 누르세요.")
        desc.setObjectName("muted")
        desc.setWordWrap(True)
        layout.addWidget(desc)

        notice = QLabel(
            "⚠  패턴 출력 전 TV 설정 확인\n"
            "   •  Aspect Ratio :  Original\n"
            "   •  Just Scan      :  On"
        )
        notice.setStyleSheet(
            "background:#fff8e1; border:1px solid #f0c040; border-radius:5px;"
            "padding:8px 12px; color:#7a5800; font-size:12px;"
        )
        layout.addWidget(notice)

        btn_row = QHBoxLayout()
        self._btn_start = QPushButton("▶  패턴 출력")
        self._btn_start.setObjectName("primary")
        self._btn_start.clicked.connect(self._start)
        self._btn_ok = QPushButton("✔  OK — 센터 확인")
        self._btn_ok.setObjectName("success")
        self._btn_ok.setEnabled(False)
        self._btn_ok.clicked.connect(self._confirm)
        btn_row.addWidget(self._btn_start)
        btn_row.addWidget(self._btn_ok)
        btn_row.addStretch()
        layout.addLayout(btn_row)

        self._status = QLabel("대기 중")
        self._status.setObjectName("muted")
        layout.addWidget(self._status)

    def _start(self) -> None:
        self._btn_start.setEnabled(False)
        self._status.setText("패턴 출력 중...")
        self._worker = MeasurementWorker(self._engine, "center_align")
        self._worker.succeeded.connect(self._on_ready)
        self._worker.error.connect(self._on_error)
        wire_worker_cleanup(self._worker, self, '_worker')
        self._worker.start()

    def _on_ready(self, _result: Any) -> None:
        self._status.setText("패턴 출력 중 — 센터 확인 후 OK 클릭")
        self._btn_ok.setEnabled(True)
        self._btn_start.setEnabled(True)

    def _on_error(self, msg: str) -> None:
        self._status.setText(f"오류: {msg}")
        self._btn_start.setEnabled(True)

    def _confirm(self) -> None:
        self._status.setText("✔  센터 확인 완료. 다음 단계로 진행하세요.")
        self._btn_ok.setEnabled(False)


# ---------------------------------------------------------------------------
# Luminance Swing Panel
# ---------------------------------------------------------------------------

_SWING_CASE_COLORS = {"Vivid": "#e74c3c", "Standard": "#4f8ef7", "Cinema": "#27ae60"}


_SWING_CASE_ORDER = ["Vivid", "Standard", "Cinema"]


class LumSwingPanel(QWidget):
    def __init__(self, engine: MeasurementEngine, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._engine = engine
        self._worker: Optional[MeasurementWorker] = None
        # "SDR_Vivid", "HDR_Standard" 등 키로 결과 누적
        self._all_data: Dict[str, List[MeasureResult]] = {}
        # 현재 측정 중인 키
        self._current_key: str = ""
        # 범례 순서 고정: 초기화 시 Vivid→Standard→Cinema 순으로 미리 생성
        # {case: QLineSeries}
        self._sdr_series: Dict[str, QLineSeries] = {}
        self._hdr_series: Dict[str, QLineSeries] = {}

        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignmentFlag.AlignTop)

        # ── 컨트롤 한 줄 ────────────────────────────────────────────────
        ctrl = QHBoxLayout()
        title = QLabel("📈 휘도 스윙")
        title.setStyleSheet("font-size:15px;font-weight:bold;")
        ctrl.addWidget(title)
        ctrl.addWidget(QLabel("PSM:"))
        self._case_combo = QComboBox()
        self._case_combo.addItems(["Vivid", "Standard", "Cinema"])
        ctrl.addWidget(self._case_combo)
        self._hdr_check = QCheckBox("HDR")
        self._hdr_check.setChecked(False)
        self._hdr_check.stateChanged.connect(self._on_hdr_toggled)
        ctrl.addWidget(self._hdr_check)
        self._btn_run = QPushButton("▶  시작")
        self._btn_run.setObjectName("primary")
        self._btn_run.clicked.connect(self._run)
        ctrl.addWidget(self._btn_run)
        self._btn_stop = QPushButton("■  중지")
        self._btn_stop.setObjectName("danger")
        self._btn_stop.setEnabled(False)
        self._btn_stop.clicked.connect(self._stop)
        ctrl.addWidget(self._btn_stop)
        self._btn_clear = QPushButton("🗑 초기화")
        self._btn_clear.clicked.connect(self._clear)
        ctrl.addWidget(self._btn_clear)
        self._btn_export = QPushButton("💾  Excel 저장")
        self._btn_export.clicked.connect(self._export)
        ctrl.addWidget(self._btn_export)
        ctrl.addStretch()
        layout.addLayout(ctrl)

        # ── 프로그레스 + 상태 (다음 줄) ─────────────────────────────────
        prog_row = QHBoxLayout()
        self._progress = QProgressBar()
        self._progress.setFixedHeight(10)
        prog_row.addWidget(self._progress)
        self._status_label = QLabel("대기 중")
        self._status_label.setObjectName("muted")
        prog_row.addWidget(self._status_label)
        prog_row.addStretch()
        layout.addLayout(prog_row)

        # ── SDR 차트 | HDR 차트 (좌우 분할) ─────────────────────────────
        def _make_chart(title: str):
            chart = QChart()
            chart.setTitle(title)
            chart.setBackgroundBrush(QColor("#ffffff"))

            # 타이틀 폰트 작게
            title_font = chart.titleFont()
            title_font.setPointSize(10)
            chart.setTitleFont(title_font)
            chart.setTitleBrush(QColor("#1a1d2e"))

            # 범례 폰트 작게
            legend = chart.legend()
            legend.setVisible(True)
            legend.setLabelColor(QColor("#1a1d2e"))
            legend_font = legend.font()
            legend_font.setPointSize(9)
            legend.setFont(legend_font)

            # 여백 최소화
            chart.setMargins(QMargins(2, 2, 2, 2))

            # X축
            ax = QValueAxis()
            ax.setTitleText("측정 #")
            ax.setLabelFormat("%d")
            ax.setLabelsBrush(QColor("#6b7080"))
            ax.setTitleBrush(QColor("#6b7080"))
            ax_font = ax.labelsFont()
            ax_font.setPointSize(8)
            ax.setLabelsFont(ax_font)
            ax_title_font = ax.titleFont()
            ax_title_font.setPointSize(8)
            ax.setTitleFont(ax_title_font)

            # Y축
            ay = QValueAxis()
            ay.setTitleText("Lv (cd/m²)")
            ay.setLabelFormat("%d")
            ay.setLabelsBrush(QColor("#6b7080"))
            ay.setTitleBrush(QColor("#6b7080"))
            ay.setLabelsFont(ax_font)
            ay.setTitleFont(ax_title_font)

            chart.addAxis(ax, Qt.AlignBottom)
            chart.addAxis(ay, Qt.AlignLeft)
            view = QChartView(chart)
            view.setRenderHint(QPainter.Antialiasing)
            view.setMinimumHeight(320)
            view.setStyleSheet("background: #ffffff;")
            return chart, ax, ay, view

        (self._chart_sdr, self._ax_x_sdr, self._ax_y_sdr, view_sdr) = _make_chart("SDR")
        (self._chart_hdr, self._ax_x_hdr, self._ax_y_hdr, view_hdr) = _make_chart("HDR")

        # ── 시리즈를 Vivid→Standard→Cinema 순서로 미리 생성 (범례 순서 고정) ──
        def _make_series(case: str, chart: QChart, ax_x: QValueAxis, ax_y: QValueAxis) -> QLineSeries:
            s = QLineSeries()
            s.setName(case)
            pen = s.pen()
            pen.setColor(QColor(_SWING_CASE_COLORS[case]))
            pen.setWidth(2)
            s.setPen(pen)
            chart.addSeries(s)
            s.attachAxis(ax_x)
            s.attachAxis(ax_y)
            return s

        for _case in _SWING_CASE_ORDER:
            self._sdr_series[_case] = _make_series(
                _case, self._chart_sdr, self._ax_x_sdr, self._ax_y_sdr)
            self._hdr_series[_case] = _make_series(
                _case, self._chart_hdr, self._ax_x_hdr, self._ax_y_hdr)

        chart_split = QSplitter(Qt.Horizontal)
        chart_split.addWidget(view_sdr)
        chart_split.addWidget(view_hdr)
        chart_split.setStretchFactor(0, 1)
        chart_split.setStretchFactor(1, 1)
        chart_split.setMaximumHeight(520)
        layout.addWidget(chart_split)

        # ── Lv 피벗 테이블: 행=#, 열=모드(SDR_Vivid 등) ──────────────────
        self._table = QTableWidget(0, 1)
        self._table.setAlternatingRowColors(True)
        self._table.setHorizontalHeaderLabels(["#"])
        self._table.setFixedHeight(300)  # 테이블 크기150 → 300으로 조정
        layout.addWidget(self._table)
        layout.addStretch()

    # ── 차트 헬퍼 ────────────────────────────────────────────────────────

    def _get_axes(self, is_hdr: bool):
        if is_hdr:
            return self._ax_x_hdr, self._ax_y_hdr
        return self._ax_x_sdr, self._ax_y_sdr

    def _get_series(self, case: str, is_hdr: bool) -> QLineSeries:
        return (self._hdr_series if is_hdr else self._sdr_series)[case]

    def _clear(self) -> None:
        self._all_data.clear()
        self._current_key = ""
        # 시리즈 데이터만 삭제 — 차트에서 제거하지 않아 범례 순서 유지
        for s in list(self._sdr_series.values()) + list(self._hdr_series.values()):
            s.clear()
        # 축 범위 초기화
        for ax in (self._ax_x_sdr, self._ax_x_hdr):
            ax.setRange(0, 10)
        for ay in (self._ax_y_sdr, self._ax_y_hdr):
            ay.setRange(0, 100)
        self._table.setRowCount(0)
        self._table.setColumnCount(1)
        self._table.setHorizontalHeaderLabels(["#"])
        self._status_label.setText("초기화됨")

    # ── 실행 / 중지 ──────────────────────────────────────────────────────

    def _run(self) -> None:
        case = self._case_combo.currentText()
        is_hdr = self._hdr_check.isChecked()
        mode = "HDR" if is_hdr else "SDR"
        self._current_key = f"{mode}_{case}"

        # 재측정: 해당 시리즈 데이터만 초기화 (범례 순서 그대로 유지)
        self._get_series(case, is_hdr).clear()
        self._all_data[self._current_key] = []

        self._worker = MeasurementWorker(self._engine, "lum_swing",
                                          case=case, is_hdr=is_hdr)
        self._worker.progress.connect(self._on_progress)
        self._worker.succeeded.connect(self._on_finished)
        self._worker.error.connect(self._on_error)
        wire_worker_cleanup(self._worker, self, '_worker')
        self._worker.start()
        self._btn_run.setEnabled(False)
        self._btn_stop.setEnabled(True)
        self._status_label.setText(f"{mode} {case} 측정 중...")

    def _stop(self) -> None:
        self._engine.stop_sequence()
        if self._worker:
            self._worker.requestInterruption()

    @Slot(int)
    def _on_hdr_toggled(self, state: int) -> None:
        gen = self._engine.generator
        if gen is None or not gen.is_connected:
            self._hdr_check.blockSignals(True)
            self._hdr_check.setChecked(not bool(state))
            self._hdr_check.blockSignals(False)
            self._status_label.setText("제너레이터가 연결되지 않았습니다.")
            return
        if getattr(self, '_hdr_worker', None) is not None:
            return
        enabled = bool(state)
        self._hdr_check.setEnabled(False)
        self._status_label.setText("HDR 전환 중..." if enabled else "SDR 전환 중...")

        def _done():
            self._hdr_check.setEnabled(True)
            self._status_label.setText("HDR 모드" if enabled else "SDR 모드")
            setattr(self, '_hdr_worker', None)

        worker = ConnectWorker(lambda: gen.set_hdr(enabled))
        worker.error.connect(lambda msg: (
            QMessageBox.critical(self, "HDR 전환 오류", msg),
            self._hdr_check.blockSignals(True),
            self._hdr_check.setChecked(not enabled),
            self._hdr_check.blockSignals(False),
        ))
        wire_worker_cleanup(worker, self, '_hdr_worker', extra_cb=_done)
        worker.start()
        self._hdr_worker = worker

    def _auto_save(self) -> None:
        if not self._engine.auto_save_dir or not self._all_data:
            return
        brand = self._engine.brand or "brand"
        model = self._engine.model_name or "model"
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"lum_swing_{brand}_{model}_{ts}.xlsx"
        path = os.path.join(self._engine.auto_save_dir, filename)
        try:
            ExcelExporter().export_lum_swing(
                self._all_data, self._engine.brand, self._engine.model_name,
                file_path=path,
            )
            self._status_label.setText(f"완료  |  저장: {path}")
        except Exception as e:
            QMessageBox.warning(self, "자동 저장 실패", str(e))

    def _export(self) -> None:
        if not self._all_data:
            QMessageBox.information(self, "알림", "저장할 데이터가 없습니다.")
            return
        brand = self._engine.brand or "brand"
        model = self._engine.model_name or "model"
        path, _ = QFileDialog.getSaveFileName(self, "Excel 저장",
                                               f"lum_swing_{brand}_{model}.xlsx",
                                               "Excel (*.xlsx)")
        if path:
            ExcelExporter().export_lum_swing(
                self._all_data, self._engine.brand, self._engine.model_name,
                file_path=path,
            )
            QMessageBox.information(self, "저장 완료", f"저장됨: {path}")

    @Slot(str, float, object)
    def _on_progress(self, _step: str, pct: float, data: Any) -> None:
        self._progress.setValue(int(pct * 100))
        if not isinstance(data, MeasureResult):
            return
        key = self._current_key
        self._all_data.setdefault(key, []).append(data)
        rows = self._all_data[key]
        n = len(rows)

        is_hdr = key.startswith("HDR")
        case = key.split("_", 1)[1]
        series = self._get_series(case, is_hdr)
        ax_x, ax_y = self._get_axes(is_hdr)
        series.append(float(n), data.Lv)
        ax_x.setMax(max(ax_x.max(), float(n) + 2))
        ax_y.setMax(max(ax_y.max(), data.Lv * 1.15))

        self._update_table(key, data, n)
        self._status_label.setText(f"{key} #{n}  Lv={data.Lv:.3f} cd/m²")

    @Slot(object)
    def _on_finished(self, _result: Any) -> None:
        self._progress.setValue(100)
        key = self._current_key
        n = len(self._all_data.get(key, []))
        self._status_label.setText(f"{key} 완료 — {n}건")
        self._btn_run.setEnabled(True)
        self._btn_stop.setEnabled(False)
        self._engine.session_swing[key] = list(self._all_data.get(key, []))
        path = _save_all_session(self._engine)
        if path:
            self._status_label.setText(f"{key} 완료 — {n}건  |  저장: {path}")

    @Slot(str)
    def _on_error(self, msg: str) -> None:
        QMessageBox.critical(self, "오류", msg)
        self._btn_run.setEnabled(True)
        self._btn_stop.setEnabled(False)

    def _get_col_for_key(self, key: str) -> int:
        for c in range(1, self._table.columnCount()):
            h = self._table.horizontalHeaderItem(c)
            if h and h.text() == key:
                return c
        col = self._table.columnCount()
        self._table.setColumnCount(col + 1)
        self._table.setHorizontalHeaderItem(col, QTableWidgetItem(key))
        return col

    def _update_table(self, key: str, r: MeasureResult, n: int) -> None:
        col = self._get_col_for_key(key)
        # 행이 부족하면 추가 (# 셀 포함)
        while n > self._table.rowCount():
            row_idx = self._table.rowCount()
            self._table.insertRow(row_idx)
            idx_item = QTableWidgetItem(str(row_idx + 1))
            idx_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self._table.setItem(row_idx, 0, idx_item)
        item = QTableWidgetItem(f"{r.Lv:.3f}")
        item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        self._table.setItem(n - 1, col, item)
        self._table.scrollToBottom()


# ---------------------------------------------------------------------------
# Luminance Loading Panel
# ---------------------------------------------------------------------------

_BRAND_COLOR_QT = {"samsung": "#0070C0", "lg": "#FF0000", "sony": "#00B050"}


def _qt_brand_color(brand: str) -> str:
    return _BRAND_COLOR_QT.get(brand.lower().strip(), "#FF8800")


_CASE_COLORS = {"Vivid": "#e74c3c", "Standard": "#4f8ef7", "Cinema": "#27ae60"}


class LumLoadingPanel(QWidget):
    def __init__(self, engine: MeasurementEngine, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._engine = engine
        self._results: Dict[str, Any] = {}
        self._raw_data: Dict[int, List[MeasureResult]] = {}
        # mode("SDR"/"HDR") → case → apl → results
        self._all_data: Dict[str, Dict[str, Dict[int, List[MeasureResult]]]] = {
            "SDR": {}, "HDR": {}
        }
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignmentFlag.AlignTop)

        title = QLabel("📊 APL 로딩 (Luminance Loading)")
        title.setStyleSheet("font-size:15px;font-weight:bold;")
        layout.addWidget(title)

        # ── 설정 한 줄 ────────────────────────────────────────────────
        cfg_row = QHBoxLayout()
        cfg_row.addWidget(QLabel("버전:"))
        self._version_combo = QComboBox()
        self._version_combo.addItems(["10단계", "37단계", "2단계"])
        self._version_combo.setFixedWidth(80)
        cfg_row.addWidget(self._version_combo)
        cfg_row.addSpacing(12)
        cfg_row.addWidget(QLabel("케이스:"))
        self._case_combo = QComboBox()
        self._case_combo.addItems(["Vivid", "Standard", "Cinema"])
        self._case_combo.setFixedWidth(90)
        cfg_row.addWidget(self._case_combo)
        cfg_row.addSpacing(12)
        cfg_row.addWidget(QLabel("측정 횟수:"))
        self._meas_count = QSpinBox()
        self._meas_count.setRange(1, 10)
        self._meas_count.setValue(1)
        self._meas_count.setSuffix(" 회")
        self._meas_count.setFixedWidth(100)
        cfg_row.addWidget(self._meas_count)
        cfg_row.addSpacing(12)
        self._hdr_check = QCheckBox("HDR")
        self._hdr_check.stateChanged.connect(self._on_hdr_toggled)
        cfg_row.addWidget(self._hdr_check)
        cfg_row.addStretch()
        layout.addLayout(cfg_row)

        # ── 쿨링 조건 한 줄 ──────────────────────────────────────────
        cool_row = QHBoxLayout()
        self._cooling_check = QCheckBox("쿨링")
        cool_row.addWidget(self._cooling_check)
        cool_row.addWidget(QLabel("APL <"))
        self._cool_apl_spin = QSpinBox()
        self._cool_apl_spin.setRange(1, 100)
        self._cool_apl_spin.setValue(10)
        self._cool_apl_spin.setSuffix(" %")
        self._cool_apl_spin.setFixedWidth(100)
        cool_row.addWidget(self._cool_apl_spin)
        cool_row.addWidget(QLabel("일 때"))
        self._cool_sec_spin = QSpinBox()
        self._cool_sec_spin.setRange(1, 60)
        self._cool_sec_spin.setValue(5)
        self._cool_sec_spin.setSuffix(" 초")
        self._cool_sec_spin.setFixedWidth(100)
        cool_row.addWidget(self._cool_sec_spin)
        cool_row.addWidget(QLabel("Black 출력"))
        cool_row.addStretch()
        layout.addLayout(cool_row)

        # ── 버튼 행 ──────────────────────────────────────────────────
        btn_row = QHBoxLayout()
        self._btn_run = QPushButton("▶  측정 시작")
        self._btn_run.setObjectName("primary")
        self._btn_run.clicked.connect(self._run)
        btn_row.addWidget(self._btn_run)
        self._btn_stop = QPushButton("■  중지")
        self._btn_stop.setObjectName("danger")
        self._btn_stop.setEnabled(False)
        self._btn_stop.clicked.connect(self._stop)
        btn_row.addWidget(self._btn_stop)
        self._btn_export = QPushButton("💾  Excel 저장")
        self._btn_export.clicked.connect(self._export)
        btn_row.addWidget(self._btn_export)
        self._btn_clear = QPushButton("🗑  초기화")
        self._btn_clear.clicked.connect(self._clear_chart)
        btn_row.addWidget(self._btn_clear)
        btn_row.addStretch()
        layout.addLayout(btn_row)

        self._progress = QProgressBar()
        layout.addWidget(self._progress)
        self._status_label = QLabel("대기 중")
        self._status_label.setObjectName("muted")
        layout.addWidget(self._status_label)

        # ── SDR / HDR 차트 좌우 분할 ─────────────────────────────────
        def _make_apl_chart(title: str):
            chart = QChart()
            chart.setTitle(title)
            chart.setBackgroundBrush(QColor("#ffffff"))
            chart.setMargins(QMargins(2, 2, 2, 2))
            chart.setContentsMargins(0, 0, 0, 0)

            title_font = chart.titleFont()
            title_font.setPointSize(9)
            chart.setTitleFont(title_font)
            chart.setTitleBrush(QColor("#1a1d2e"))

            legend = chart.legend()
            legend.setLabelColor(QColor("#1a1d2e"))
            legend.show()
            leg_font = legend.font()
            leg_font.setPointSize(8)
            legend.setFont(leg_font)

            ax = QValueAxis()
            ax.setTitleText("APL (%)")
            ax.setRange(0, 100)
            ax.setTickCount(11)
            ax.setLabelsBrush(QColor("#6b7080"))
            ax.setTitleBrush(QColor("#6b7080"))
            ax.setLabelFormat("%d")
            ax_font = ax.labelsFont()
            ax_font.setPointSize(8)
            ax.setLabelsFont(ax_font)
            ax_title_font = ax.titleFont()
            ax_title_font.setPointSize(8)
            ax.setTitleFont(ax_title_font)

            ay = QValueAxis()
            ay.setTitleText("Lv (cd/m²)")
            ay.setLabelsBrush(QColor("#6b7080"))
            ay.setTitleBrush(QColor("#6b7080"))
            ay.setLabelFormat("%d")
            ay.setLabelsFont(ax_font)
            ay.setTitleFont(ax_title_font)

            chart.addAxis(ax, Qt.AlignBottom)
            chart.addAxis(ay, Qt.AlignLeft)
            view = QChartView(chart)
            view.setRenderHint(QPainter.Antialiasing)
            view.setMinimumHeight(60)
            view.setStyleSheet("background: #ffffff;")
            return chart, ax, ay, view

        (self._chart_sdr, self._ax_x_sdr,
         self._ax_y_sdr, view_sdr) = _make_apl_chart("SDR")
        (self._chart_hdr, self._ax_x_hdr,
         self._ax_y_hdr, view_hdr) = _make_apl_chart("HDR")

        # ── 시리즈를 Vivid→Standard→Cinema 순서로 미리 생성 (범례 순서 고정) ──
        _CASE_ORDER = ["Vivid", "Standard", "Cinema"]
        self._sdr_apl_series: Dict[str, QLineSeries] = {}
        self._hdr_apl_series: Dict[str, QLineSeries] = {}

        def _make_apl_series(case: str, chart: QChart,
                             ax_x: QValueAxis, ax_y: QValueAxis) -> QLineSeries:
            s = QLineSeries()
            s.setName(case)
            pen = s.pen()
            pen.setColor(QColor(_CASE_COLORS.get(case, "#888899")))
            pen.setWidth(2)
            s.setPen(pen)
            chart.addSeries(s)
            s.attachAxis(ax_x)
            s.attachAxis(ax_y)
            return s

        for _case in _CASE_ORDER:
            self._sdr_apl_series[_case] = _make_apl_series(
                _case, self._chart_sdr, self._ax_x_sdr, self._ax_y_sdr)
            self._hdr_apl_series[_case] = _make_apl_series(
                _case, self._chart_hdr, self._ax_x_hdr, self._ax_y_hdr)

        # SDR / HDR 차트 좌우 분할
        chart_split = QSplitter(Qt.Horizontal)
        chart_split.addWidget(view_sdr)
        chart_split.addWidget(view_hdr)
        chart_split.setStretchFactor(0, 1)
        chart_split.setStretchFactor(1, 1)

        self._table = QTableWidget(0, 9)
        self._table.setAlternatingRowColors(True)
        self._table.setHorizontalHeaderLabels(["APL%", "#", "Lv (cd/m²)", "x", "y", "u'", "v'", "CCT (K)", "Duv"])
        self._table.setMinimumHeight(80)

        # 차트-테이블 상하 분할 (드래그로 비율 조정 가능)
        v_split = QSplitter(Qt.Vertical)
        v_split.addWidget(chart_split)
        v_split.addWidget(self._table)
        v_split.setSizes([110, 140])   # ← 초기 비율 (픽셀), 여기를 바꾸면 됨
        layout.addWidget(v_split)

    @Slot(int)
    def _on_hdr_toggled(self, state: int) -> None:
        gen = self._engine.generator
        if gen is None or not gen.is_connected:
            self._hdr_check.blockSignals(True)
            self._hdr_check.setChecked(not bool(state))
            self._hdr_check.blockSignals(False)
            self._status_label.setText("제너레이터가 연결되지 않았습니다.")
            return
        if getattr(self, '_hdr_worker', None) is not None:
            return
        enabled = bool(state)
        self._hdr_check.setEnabled(False)
        self._status_label.setText("HDR 전환 중..." if enabled else "SDR 전환 중...")

        def _done():
            self._hdr_check.setEnabled(True)
            self._status_label.setText("HDR 모드" if enabled else "SDR 모드")
            setattr(self, '_hdr_worker', None)

        worker = ConnectWorker(lambda: gen.set_hdr(enabled))
        worker.error.connect(lambda msg: (
            QMessageBox.critical(self, "HDR 전환 오류", msg),
            self._hdr_check.blockSignals(True),
            self._hdr_check.setChecked(not enabled),
            self._hdr_check.blockSignals(False),
        ))
        wire_worker_cleanup(worker, self, '_hdr_worker', extra_cb=_done)
        worker.start()
        self._hdr_worker = worker

    def _run(self) -> None:
        version_map = {"37단계": "37", "10단계": "10", "2단계": "2"}
        version = version_map[self._version_combo.currentText()]
        self._raw_data.clear()
        self._table.setRowCount(0)
        self._live_series: Optional["QLineSeries"] = None   # 현재 측정 중인 시리즈
        self._btn_run.setEnabled(False)
        self._worker = MeasurementWorker(
            self._engine, "lum_loading",
            version=version,
            case=self._case_combo.currentText(),
            is_hdr=self._hdr_check.isChecked(),
            cooling_enabled=self._cooling_check.isChecked(),
            cooling_apl_threshold=self._cool_apl_spin.value(),
            cooling_duration_sec=float(self._cool_sec_spin.value()),
            measurements_per_step=self._meas_count.value(),
        )
        self._worker.progress.connect(self._on_progress)
        self._worker.succeeded.connect(self._on_finished)
        self._worker.error.connect(lambda m: (
            QMessageBox.critical(self, "오류", m),
            self._btn_run.setEnabled(True),
            self._btn_stop.setEnabled(False),
        ))
        wire_worker_cleanup(self._worker, self, '_worker')
        self._worker.start()
        self._btn_run.setEnabled(False)
        self._btn_stop.setEnabled(True)

    def _stop(self) -> None:
        self._engine.stop_sequence()
        if self._worker:
            self._worker.requestInterruption()

    def _all_data_flat(self) -> Dict[str, Dict[int, List[MeasureResult]]]:
        """_all_data 전체를 {"SDR_Vivid": {...}, "HDR_Standard": {...}, ...} 로 평탄화.

        export_lum_loading 에 넘길 results_by_case 형태.
        """
        flat: Dict[str, Dict[int, List[MeasureResult]]] = {}
        for mode, cases in self._all_data.items():
            for case, apl_dict in cases.items():
                if apl_dict:
                    flat[f"{mode}_{case}"] = apl_dict
        return flat

    def _auto_save(self) -> None:
        flat = self._all_data_flat()
        if not self._engine.auto_save_dir or not flat:
            return
        brand = self._engine.brand or "brand"
        model = self._engine.model_name or "model"
        # 고정 파일명 — 케이스가 추가될 때마다 덮어씌워 항상 최신 전체 데이터 유지
        filename = f"lum_loading_{brand}_{model}.xlsx"
        path = os.path.join(self._engine.auto_save_dir, filename)
        try:
            ExcelExporter().export_lum_loading(
                flat, self._engine.brand, self._engine.model_name,
                file_path=path,
            )
            cases_str = ", ".join(flat.keys())
            self._status_label.setText(
                f"완료 — {len(self._raw_data)}개 APL  |  자동 저장 ({cases_str}): {path}"
            )
        except Exception as e:
            QMessageBox.warning(self, "자동 저장 실패", str(e))

    def _on_finished(self, result: Any) -> None:
        self._results = result or {}
        self._live_series = None  # 라이브 시리즈 참조 해제
        mode = "HDR" if self._hdr_check.isChecked() else "SDR"
        case = self._case_combo.currentText()
        self._all_data[mode][case] = dict(self._raw_data)
        # 세션 업데이트 후 통합 파일 자동 저장
        self._engine.session_loading[f"{mode}_{case}"] = dict(self._raw_data)
        self._status_label.setText(f"완료 — {len(self._raw_data)}개 APL 측정")
        self._btn_run.setEnabled(True)
        self._btn_stop.setEnabled(False)
        path = _save_all_session(self._engine)
        if path:
            self._status_label.setText(
                f"완료 — {len(self._raw_data)}개 APL  |  저장: {path}"
            )

    @Slot(str, float, object)
    def _on_progress(self, _step: str, pct: float, data: Any) -> None:
        self._progress.setValue(int(pct * 100))
        if not (isinstance(data, dict) and "apl" in data):
            return
        apl = int(data["apl"])
        results: List[MeasureResult] = data.get("results", [])
        self._raw_data[apl] = results

        # ── 테이블: 새 APL 스텝 행만 추가 (전체 재구성 없음) ─────────────
        for idx, r in enumerate(results, start=1):
            self._add_table_row(f"{apl}%", str(idx), r)
        self._table.scrollToBottom()

        lv_avg = sum(r.Lv for r in results) / len(results) if results else 0
        self._status_label.setText(f"APL {apl}% — Lv={lv_avg:.3f} cd/m²  ({int(pct*100)}%)")

        # ── 차트: 현재 케이스 시리즈에 포인트 추가만 (전체 재생성 없음) ──
        mode = "HDR" if self._hdr_check.isChecked() else "SDR"
        case = self._case_combo.currentText()
        self._all_data[mode].setdefault(case, {})[apl] = results

        ax_y  = self._ax_y_sdr  if mode == "SDR" else self._ax_y_hdr

        if self._live_series is None:
            # 새 측정 시작 — 미리 만든 시리즈 데이터만 초기화 후 사용
            series_map = self._sdr_apl_series if mode == "SDR" else self._hdr_apl_series
            series = series_map[case]
            series.clear()
            self._live_series = series

        self._live_series.append(float(apl), lv_avg)

        # Y축 범위 동적 확장
        all_lv = [
            sum(r.Lv for r in rs) / len(rs)
            for rs in self._raw_data.values() if rs
        ]
        if all_lv:
            ax_y.setRange(0, max(all_lv) * 1.15)

    def _refresh_table(self) -> None:
        self._table.setRowCount(0)
        for apl in sorted(self._raw_data):
            results = self._raw_data[apl]
            for idx, r in enumerate(results, start=1):
                self._add_table_row(f"{apl}%", str(idx), r)

    def _refresh_apl_chart(self) -> None:
        for s in list(self._sdr_apl_series.values()) + list(self._hdr_apl_series.values()):
            s.clear()

        for mode, series_map, ax_y in [
            ("SDR", self._sdr_apl_series, self._ax_y_sdr),
            ("HDR", self._hdr_apl_series, self._ax_y_hdr),
        ]:
            all_lv: List[float] = []
            for case, apl_dict in self._all_data[mode].items():
                series = series_map.get(case)
                if not series or not apl_dict:
                    continue
                for apl in sorted(apl_dict):
                    results = apl_dict[apl]
                    if not results:
                        continue
                    lv = sum(r.Lv for r in results) / len(results)
                    series.append(float(apl), lv)
                    all_lv.append(lv)
            if all_lv:
                ax_y.setRange(0, max(all_lv) * 1.15)

    def _clear_chart(self) -> None:
        self._all_data = {"SDR": {}, "HDR": {}}
        self._raw_data.clear()
        self._live_series = None
        for s in list(self._sdr_apl_series.values()) + list(self._hdr_apl_series.values()):
            s.clear()
        for ay in (self._ax_y_sdr, self._ax_y_hdr):
            ay.setRange(0, 100)
        self._table.setRowCount(0)
        self._status_label.setText("그래프 초기화됨")

    def _add_table_row(self, apl_label: str, idx_label: str, r: MeasureResult) -> None:
        row = self._table.rowCount()
        self._table.insertRow(row)
        cct_str = f"{r.cct:.0f}" if r.cct else "—"
        duv_str = f"{r.duv:.5f}" if r.cct else "—"
        for ci, val in enumerate([apl_label, idx_label,
                                   f"{r.Lv:.3f}", f"{r.x:.4f}", f"{r.y:.4f}",
                                   f"{r.u_prime:.4f}", f"{r.v_prime:.4f}",
                                   cct_str, duv_str]):
            item = QTableWidgetItem(str(val))
            item.setTextAlignment(Qt.AlignCenter)
            self._table.setItem(row, ci, item)

    def _export(self) -> None:
        if not self._raw_data:
            QMessageBox.information(self, "알림", "저장할 데이터가 없습니다.")
            return
        case = self._case_combo.currentText()
        mode = "HDR" if self._hdr_check.isChecked() else "SDR"
        brand = self._engine.brand or "brand"
        model = self._engine.model_name or "model"
        default_name = f"lum_loading_{mode}_{case.lower()}_{brand}_{model}.xlsx"
        path, _ = QFileDialog.getSaveFileName(self, "Excel 저장", default_name, "Excel (*.xlsx)")
        if path:
            try:
                ExcelExporter().export_lum_loading(
                    {case: self._raw_data},
                    self._engine.brand, self._engine.model_name,
                    file_path=path,
                )
                QMessageBox.information(self, "저장 완료", f"저장됨: {path}")
            except Exception:
                import traceback
                QMessageBox.critical(self, "저장 오류", traceback.format_exc())


# ---------------------------------------------------------------------------
# Gamut Panel
# ---------------------------------------------------------------------------

class GamutPanel(QWidget):
    # u'v' 색도 좌표 색상 매핑
    _UV_COLOR = {
        "red":   QColor("#e74c3c"),
        "green": QColor("#2ecc71"),
        "blue":  QColor("#4f8ef7"),
        "white": QColor("#888899"),
        "black": QColor("#aab0c0"),
    }

    def __init__(self, engine: MeasurementEngine, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._engine = engine
        self._results: Dict[str, MeasureResult] = {}
        self._gamut_stats: Dict[str, float] = {}
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignmentFlag.AlignTop)

        # ── 타이틀 + 버튼 + 프로그레스 + 통계 (한 줄) ───────────────────
        top_row = QHBoxLayout()
        top_row.setSpacing(6)
        title = QLabel("🎨 색재현율")
        title.setStyleSheet("font-size:15px;font-weight:bold;")
        top_row.addWidget(title)
        self._hdr_check = QCheckBox("HDR")
        self._hdr_check.stateChanged.connect(self._on_hdr_toggled)
        top_row.addWidget(self._hdr_check)
        self._btn_run = QPushButton("▶  측정 시작")
        self._btn_run.setObjectName("primary")
        self._btn_run.clicked.connect(self._run)
        top_row.addWidget(self._btn_run)
        self._btn_stop = QPushButton("■  중지")
        self._btn_stop.setObjectName("danger")
        self._btn_stop.setEnabled(False)
        self._btn_stop.clicked.connect(self._stop)
        top_row.addWidget(self._btn_stop)
        self._btn_export = QPushButton("💾  Excel 저장")
        self._btn_export.clicked.connect(self._export)
        top_row.addWidget(self._btn_export)
        top_row.addStretch()
        self._lbl_dci    = QLabel("DCI-P3: —")
        self._lbl_bt2020 = QLabel("BT.2020: —")
        for lbl in (self._lbl_dci, self._lbl_bt2020):
            lbl.setStyleSheet("font-weight:bold; font-size:13px; padding:2px 8px;")
            top_row.addWidget(lbl)
        layout.addLayout(top_row)

        # ── 프로그레스 + 상태 (다음 줄) ─────────────────────────────────
        prog_row = QHBoxLayout()
        self._progress = QProgressBar()
        self._progress.setFixedHeight(10)
        prog_row.addWidget(self._progress)
        self._status_label = QLabel("대기 중")
        self._status_label.setObjectName("muted")
        prog_row.addWidget(self._status_label)
        prog_row.addStretch()
        layout.addLayout(prog_row)

        # ── 저장 경로 (별도 줄, 평소에는 숨김) ────────────────────────────
        self._path_label = QLabel()
        self._path_label.setObjectName("muted")
        self._path_label.setStyleSheet("font-size:11px; padding:0px 2px;")
        self._path_label.setVisible(False)
        layout.addWidget(self._path_label)

        # ↓ 버튼 행~저장경로 라벨까지와 차트+표 사이의 간격
        #   숫자를 늘리면 차트+표가 아래로 내려갑니다 (단위: px)
        layout.addSpacing(60)  # ← 이 값을 조절하세요 (예: 16, 24, 40 ...)

        # u'v' 차트 + 테이블 분할
        splitter = QSplitter(Qt.Horizontal)

        # u'v' ScatterChart
        self._chart = QChart()
        self._chart.setTitle("u'v' 색도 다이어그램")
        self._chart.setBackgroundBrush(QColor("#ffffff"))
        self._chart.setTitleBrush(QColor("#1a1d2e"))
        self._chart.setMargins(QMargins(4, 4, 4, 4))
        self._axis_u = QValueAxis()
        self._axis_u.setTitleText("u'")
        self._axis_u.setRange(0.0, 0.65)
        self._axis_u.setLabelsBrush(QColor("#6b7080"))
        self._axis_u.setTitleBrush(QColor("#6b7080"))
        self._axis_v = QValueAxis()
        self._axis_v.setTitleText("v'")
        self._axis_v.setRange(0.0, 0.65)
        self._axis_v.setLabelsBrush(QColor("#6b7080"))
        self._axis_v.setTitleBrush(QColor("#6b7080"))
        self._chart.addAxis(self._axis_u, Qt.AlignBottom)
        self._chart.addAxis(self._axis_v, Qt.AlignLeft)
        self._chart.legend().setLabelColor(QColor("#1a1d2e"))

        # DCI-P3 / BT.2020 기준 삼각형
        self._add_reference_gamut()

        chart_view = QChartView(self._chart)
        chart_view.setRenderHint(QPainter.Antialiasing)
        chart_view.setMinimumWidth(420)   # ← 차트 최소 너비 (px)
        chart_view.setMinimumHeight(420)  # ← 차트 최소 높이 (px)
        chart_view.setStyleSheet("background: #ffffff;")
        splitter.addWidget(chart_view)

        self._table = QTableWidget(0, 9)
        self._table.setAlternatingRowColors(True)
        self._table.setHorizontalHeaderLabels(["컬러", "Lv", "x", "y", "u'", "v'", "X", "Y", "Z"])
        splitter.addWidget(self._table)
        splitter.setStretchFactor(0, 3)  # ← 차트 : 표 = 3 : 2 비율 (숫자로 조절)
        splitter.setStretchFactor(1, 2)
        layout.addWidget(splitter)

    def _add_reference_gamut(self) -> None:
        """DCI-P3(점선 흰색), BT.2020(점선 회색) 기준 삼각형 표시."""
        for ref_pts, color, name in [
            (DCI_P3_UV, QColor("#4f8ef7"), "DCI-P3"),
            (BT2020_UV, QColor("#aab0c0"), "BT.2020"),
        ]:
            series = QLineSeries()
            series.setName(name)
            pen = series.pen()
            pen.setColor(color)
            pen.setWidth(1)
            series.setPen(pen)
            for u, v in ref_pts:
                series.append(u, v)
            # 닫힌 삼각형
            series.append(ref_pts[0][0], ref_pts[0][1])
            self._chart.addSeries(series)
            series.attachAxis(self._axis_u)
            series.attachAxis(self._axis_v)

    @Slot(int)
    def _on_hdr_toggled(self, state: int) -> None:
        gen = self._engine.generator
        if gen is None or not gen.is_connected:
            self._hdr_check.blockSignals(True)
            self._hdr_check.setChecked(not bool(state))
            self._hdr_check.blockSignals(False)
            self._status_label.setText("제너레이터가 연결되지 않았습니다.")
            return
        if getattr(self, '_hdr_worker', None) is not None:
            return
        enabled = bool(state)
        self._hdr_check.setEnabled(False)
        self._status_label.setText("HDR 전환 중..." if enabled else "SDR 전환 중...")

        def _done():
            self._hdr_check.setEnabled(True)
            self._status_label.setText("HDR 모드" if enabled else "SDR 모드")
            setattr(self, '_hdr_worker', None)

        worker = ConnectWorker(lambda: gen.set_hdr(enabled))
        worker.error.connect(lambda msg: (
            QMessageBox.critical(self, "HDR 전환 오류", msg),
            self._hdr_check.blockSignals(True),
            self._hdr_check.setChecked(not enabled),
            self._hdr_check.blockSignals(False),
        ))
        wire_worker_cleanup(worker, self, '_hdr_worker', extra_cb=_done)
        worker.start()
        self._hdr_worker = worker

    def _run(self) -> None:
        self._table.setRowCount(0)
        self._results.clear()
        self._set_path("")
        # 기존 측정점 시리즈 제거 후 기준 gamut 재설정
        self._chart.removeAllSeries()
        self._add_reference_gamut()
        self._lbl_dci.setText("DCI-P3: —")
        self._lbl_bt2020.setText("BT.2020: —")
        self._btn_run.setEnabled(False)
        self._btn_stop.setEnabled(True)
        self._worker = MeasurementWorker(self._engine, "gamut", is_hdr=self._hdr_check.isChecked())
        self._worker.progress.connect(self._on_progress)
        self._worker.succeeded.connect(self._on_finished)
        self._worker.error.connect(lambda m: (
            QMessageBox.critical(self, "오류", m),
            self._btn_run.setEnabled(True),
            self._btn_stop.setEnabled(False),
        ))
        wire_worker_cleanup(self._worker, self, '_worker')
        self._worker.start()

    def _stop(self) -> None:
        self._engine.stop_sequence()
        if self._worker:
            self._worker.requestInterruption()

    def _set_path(self, path: str) -> None:
        if path:
            self._path_label.setText(f"저장: {path}")
            self._path_label.setVisible(True)
        else:
            self._path_label.setVisible(False)

    def _auto_save(self) -> None:
        if not self._engine.auto_save_dir or not self._results:
            return
        mode = "HDR" if self._hdr_check.isChecked() else "SDR"
        brand = self._engine.brand or "brand"
        model = self._engine.model_name or "model"
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"gamut_{mode}_{brand}_{model}_{ts}.xlsx"
        path = os.path.join(self._engine.auto_save_dir, filename)
        try:
            ExcelExporter().export_gamut(self._results, self._engine.brand,
                                         self._engine.model_name, file_path=path)
            self._status_label.setText("완료  |  자동 저장")
            self._set_path(path)
        except Exception as e:
            QMessageBox.warning(self, "자동 저장 실패", str(e))

    @Slot(object)
    def _on_finished(self, result: Any) -> None:
        self._results = result or {}
        self._status_label.setText("완료")
        self._btn_run.setEnabled(True)
        self._btn_stop.setEnabled(False)
        self._update_gamut_stats()
        self._draw_meas_triangle()
        # 세션 업데이트 후 통합 파일 자동 저장
        mode = "HDR" if self._hdr_check.isChecked() else "SDR"
        self._engine.session_gamut[mode] = dict(self._results)
        path = _save_all_session(self._engine)
        self._set_path(path)

    def _update_gamut_stats(self) -> None:
        r_r = self._results.get("red")
        r_g = self._results.get("green")
        r_b = self._results.get("blue")
        if r_r and r_g and r_b:
            self._gamut_stats = calc_gamut_stats(
                (r_r.u_prime, r_r.v_prime),
                (r_g.u_prime, r_g.v_prime),
                (r_b.u_prime, r_b.v_prime),
            )
            self._lbl_dci.setText(f"DCI-P3: {self._gamut_stats['dci_overlap']:.1f}%")
            self._lbl_bt2020.setText(f"BT.2020: {self._gamut_stats['bt2020_overlap']:.1f}%")

    def _draw_meas_triangle(self) -> None:
        r_r = self._results.get("red")
        r_g = self._results.get("green")
        r_b = self._results.get("blue")
        if not (r_r and r_g and r_b):
            return
        pts = [
            (r_r.u_prime, r_r.v_prime),
            (r_g.u_prime, r_g.v_prime),
            (r_b.u_prime, r_b.v_prime),
        ]
        tri = QLineSeries()
        tri.setName("Measured")
        pen = tri.pen()
        pen.setColor(QColor("#f7c94f"))
        pen.setWidth(2)
        tri.setPen(pen)
        for u, v in pts:
            tri.append(u, v)
        tri.append(pts[0][0], pts[0][1])
        self._chart.addSeries(tri)
        tri.attachAxis(self._axis_u)
        tri.attachAxis(self._axis_v)

    @Slot(str, float, object)
    def _on_progress(self, _step: str, pct: float, data: Any) -> None:
        self._progress.setValue(int(pct * 100))
        if isinstance(data, dict) and "color" in data:
            color = data["color"]
            if color is None:
                self._status_label.setText("초기화 중...")
                return
            r = data.get("result")
            if r:
                # 테이블 행 추가
                row = self._table.rowCount()
                self._table.insertRow(row)
                for ci, val in enumerate([color, f"{r.Lv:.3f}", f"{r.x:.4f}", f"{r.y:.4f}",
                                           f"{r.u_prime:.4f}", f"{r.v_prime:.4f}",
                                           f"{r.X:.3f}", f"{r.Y:.3f}", f"{r.Z:.3f}"]):
                    item = QTableWidgetItem(str(val))
                    item.setTextAlignment(Qt.AlignCenter)
                    self._table.setItem(row, ci, item)
                # u'v' 차트에 점 추가
                dot = QScatterSeries()
                dot.setName(color)
                dot.setMarkerSize(12.0)
                dot_color = self._UV_COLOR.get(color, QColor("#888899"))
                dot.setColor(dot_color)
                dot.setBorderColor(dot_color)
                dot.append(r.u_prime, r.v_prime)
                self._chart.addSeries(dot)
                dot.attachAxis(self._axis_u)
                dot.attachAxis(self._axis_v)
            self._status_label.setText(f"{color} 측정 완료 — {int(pct*100)}%")

    def _export(self) -> None:
        if not self._results:
            QMessageBox.information(self, "알림", "저장할 데이터가 없습니다.")
            return
        brand = self._engine.brand or "brand"
        model = self._engine.model_name or "model"
        default_name = f"gamut_{brand}_{model}.xlsx"
        path, _ = QFileDialog.getSaveFileName(self, "Excel 저장", default_name, "Excel (*.xlsx)")
        if path:
            try:
                ExcelExporter().export_gamut(self._results, self._engine.brand,
                                             self._engine.model_name, file_path=path)
                QMessageBox.information(self, "저장 완료", f"저장됨: {path}")
            except Exception as e:
                QMessageBox.critical(self, "저장 오류", str(e))


# ---------------------------------------------------------------------------
# Contrast Panel
# ---------------------------------------------------------------------------

class ContrastPanel(QWidget):
    def __init__(self, engine: MeasurementEngine, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._engine = engine
        self._results: Dict[float, MeasureResult] = {}
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignmentFlag.AlignTop)

        title = QLabel("⬛ 명암비 (Contrast Ratio)")
        title.setStyleSheet("font-size:15px;font-weight:bold;")
        layout.addWidget(title)
        desc = QLabel("White Raster + Black Window — 창 H/V 100% → 50% → 20% → 14.1% 순으로 측정합니다.")
        desc.setObjectName("muted")
        layout.addWidget(desc)

        btn_row = QHBoxLayout()
        self._hdr_check = QCheckBox("HDR")
        self._hdr_check.stateChanged.connect(self._on_hdr_toggled)
        btn_row.addWidget(self._hdr_check)
        self._btn_run = QPushButton("▶  측정 시작")
        self._btn_run.setObjectName("primary")
        self._btn_run.clicked.connect(self._run)
        btn_row.addWidget(self._btn_run)
        self._btn_stop = QPushButton("■  중지")
        self._btn_stop.setObjectName("danger")
        self._btn_stop.setEnabled(False)
        self._btn_stop.clicked.connect(self._stop)
        btn_row.addWidget(self._btn_stop)
        self._btn_export = QPushButton("💾  Excel 저장")
        self._btn_export.clicked.connect(self._export)
        btn_row.addWidget(self._btn_export)
        btn_row.addStretch()
        layout.addLayout(btn_row)

        self._progress = QProgressBar()
        layout.addWidget(self._progress)
        self._status_label = QLabel("대기 중")
        self._status_label.setObjectName("muted")
        layout.addWidget(self._status_label)

        self._lv_ref: Optional[float] = None  # 창 100% 기준 흑휘도
        self._table = QTableWidget(0, 5)
        self._table.setAlternatingRowColors(True)
        self._table.setHorizontalHeaderLabels(["Black H/V (%)", "Lv (cd/m²)", "x", "y", "CR (White/Lv)"])
        layout.addWidget(self._table)

    @Slot(int)
    def _on_hdr_toggled(self, state: int) -> None:
        gen = self._engine.generator
        if gen is None or not gen.is_connected:
            self._hdr_check.blockSignals(True)
            self._hdr_check.setChecked(not bool(state))
            self._hdr_check.blockSignals(False)
            self._status_label.setText("제너레이터가 연결되지 않았습니다.")
            return
        if getattr(self, '_hdr_worker', None) is not None:
            return
        enabled = bool(state)
        self._hdr_check.setEnabled(False)
        self._status_label.setText("HDR 전환 중..." if enabled else "SDR 전환 중...")

        def _done():
            self._hdr_check.setEnabled(True)
            self._status_label.setText("HDR 모드" if enabled else "SDR 모드")
            setattr(self, '_hdr_worker', None)

        worker = ConnectWorker(lambda: gen.set_hdr(enabled))
        worker.error.connect(lambda msg: (
            QMessageBox.critical(self, "HDR 전환 오류", msg),
            self._hdr_check.blockSignals(True),
            self._hdr_check.setChecked(not enabled),
            self._hdr_check.blockSignals(False),
        ))
        wire_worker_cleanup(worker, self, '_hdr_worker', extra_cb=_done)
        worker.start()
        self._hdr_worker = worker

    def _auto_save(self) -> None:
        if not self._engine.auto_save_dir or not self._results:
            return
        mode = "HDR" if self._hdr_check.isChecked() else "SDR"
        brand = self._engine.brand or "brand"
        model = self._engine.model_name or "model"
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"contrast_{mode}_{brand}_{model}_{ts}.xlsx"
        path = os.path.join(self._engine.auto_save_dir, filename)
        try:
            ExcelExporter().export_contrast(self._results, self._engine.brand,
                                            self._engine.model_name, file_path=path)
            self._status_label.setText(f"완료  |  자동 저장: {path}")
        except Exception as e:
            QMessageBox.warning(self, "자동 저장 실패", str(e))

    def _on_finished(self, result: Any) -> None:
        self._results = result or {}
        self._status_label.setText("완료")
        self._btn_run.setEnabled(True)
        self._btn_stop.setEnabled(False)
        # 세션 업데이트 후 통합 파일 자동 저장
        mode = "HDR" if self._hdr_check.isChecked() else "SDR"
        self._engine.session_contrast[mode] = dict(self._results)
        path = _save_all_session(self._engine)
        if path:
            self._status_label.setText(f"완료  |  저장: {path}")

    def _run(self) -> None:
        self._table.setRowCount(0)
        self._lv_ref = None
        self._btn_run.setEnabled(False)
        self._btn_stop.setEnabled(True)
        self._worker = MeasurementWorker(self._engine, "contrast", is_hdr=self._hdr_check.isChecked())
        self._worker.progress.connect(self._on_progress)
        self._worker.succeeded.connect(self._on_finished)
        self._worker.error.connect(lambda m: (
            QMessageBox.critical(self, "오류", m),
            self._btn_run.setEnabled(True),
            self._btn_stop.setEnabled(False),
        ))
        wire_worker_cleanup(self._worker, self, '_worker')
        self._worker.start()

    def _stop(self) -> None:
        self._engine.stop_sequence()
        if self._worker:
            self._worker.requestInterruption()

    @Slot(str, float, object)
    def _on_progress(self, _step: str, pct: float, data: Any) -> None:
        self._progress.setValue(int(pct * 100))
        if not (isinstance(data, dict) and "win_size" in data):
            return
        win_size = data["win_size"]   # 0.0 = full white ref, else black window H/V side %
        r = data.get("result")
        if not r:
            return

        # win_size 0.0 = Full White → 기준 Lv 저장
        if win_size == 0.0:
            self._lv_ref = r.Lv

        hv_label = "Full White" if win_size == 0.0 else f"{win_size:.1f} × {win_size:.1f}"

        # CR = Full White Lv / Black Lv (black window 행에만, Lv > 0 가드)
        if win_size > 0.0 and self._lv_ref and self._lv_ref > 0 and r.Lv > 0:
            cr_str = f"{self._lv_ref / r.Lv:.1f} : 1"
        else:
            cr_str = "—"

        row = self._table.rowCount()
        self._table.insertRow(row)
        for ci, val in enumerate([hv_label, f"{r.Lv:.4f}",
                                   f"{r.x:.4f}", f"{r.y:.4f}", cr_str]):
            item = QTableWidgetItem(str(val))
            item.setTextAlignment(Qt.AlignCenter)
            self._table.setItem(row, ci, item)

        step_label = "Full White" if win_size == 0.0 else f"창 H/V {win_size:.1f}%"
        self._status_label.setText(f"{step_label} 측정 — {int(pct*100)}%")

    def _export(self) -> None:
        if not self._results:
            QMessageBox.information(self, "알림", "저장할 데이터가 없습니다.")
            return
        brand = self._engine.brand or "brand"
        model = self._engine.model_name or "model"
        default_name = f"contrast_{brand}_{model}.xlsx"
        path, _ = QFileDialog.getSaveFileName(self, "Excel 저장", default_name, "Excel (*.xlsx)")
        if path:
            ExcelExporter().export_contrast(self._results, self._engine.brand, self._engine.model_name,
                                            file_path=path)
            QMessageBox.information(self, "저장 완료", f"저장됨: {path}")


# ---------------------------------------------------------------------------
# Report Panel
# ---------------------------------------------------------------------------

_DEFAULT_MODEL_COLORS = [
    "#4f8ef7", "#e74c3c", "#27ae60", "#f39c12",
    "#9b59b6", "#1abc9c", "#e67e22", "#34495e",
]


class ReportPanel(QWidget):
    # ── 경쟁사 비교 장표 ──────────────────────────────────────────────────────
    _ROW_LABELS = [
        ("White 휘도", "HDR 10%"),
        ("White 휘도", "HDR 100%"),
        ("White 휘도", "SDR 10%"),
        ("White 휘도", "SDR 100%"),
        ("White 휘도", "Contrast Ratio"),
        ("White 휘도", "Black (cd/m²)"),
        ("Color Gamut", "DCI-P3 (%)"),
        ("Color Gamut", "BT.2020 (%)"),
    ]
    _COMP_KEYS = [
        "hdr_10", "hdr_100", "sdr_10", "sdr_100",
        "contrast_ratio", "black_lv", "dci_overlap", "bt2020_overlap",
    ]

    # ── 광학 측정 데이터 ──────────────────────────────────────────────────────
    _OPTICAL_ROW_LABELS = [
        ("휘도", "Vivid SDR 10% / 100%"),
        ("휘도", "Standard SDR 10% / 100%"),
        ("휘도", "Vivid HDR 10% / 100%"),
        ("휘도", "Standard HDR 10% / 100%"),
        ("휘도", "Cinema HDR 10% / 100%"),
        ("Contrast", "Black (Ratio)"),
        ("Color Gamut", "DCI-P3 (%)"),
        ("Color Gamut", "BT.2020 (%)"),
    ]
    # (key_10, key_100) — key_100=None for single-value rows
    _OPTICAL_KEYS: List[tuple] = [
        ("sdr_vivid_10",    "sdr_vivid_100"),
        ("sdr_standard_10", "sdr_standard_100"),
        ("hdr_vivid_10",    "hdr_vivid_100"),
        ("hdr_standard_10", "hdr_standard_100"),
        ("hdr_cinema_10",   "hdr_cinema_100"),
        ("contrast_ratio",  None),
        ("dci_overlap",     None),
        ("bt2020_overlap",  None),
    ]

    def __init__(
        self,
        engine: MeasurementEngine,
        gamut_panel: "GamutPanel",
        lum_panel: "LumLoadingPanel",
        contrast_panel: "ContrastPanel",
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self._engine = engine
        self._gamut_panel = gamut_panel
        self._lum_panel = lum_panel
        self._contrast_panel = contrast_panel
        self._models: List[Dict] = []
        self._model_colors: Dict[str, str] = {}

        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignmentFlag.AlignTop)

        title = QLabel("📋 보고서 템플릿")
        title.setStyleSheet("font-size:15px;font-weight:bold;")
        layout.addWidget(title)

        top_row = QHBoxLayout()
        self._btn_load = QPushButton("📂 파일 불러오기")
        self._btn_load.setObjectName("primary")
        self._btn_load.clicked.connect(self._load_files)
        top_row.addWidget(self._btn_load)

        top_row.addWidget(QLabel("보고서 형식:"))
        self._format_combo = QComboBox()
        self._format_combo.addItems(["경쟁사 비교 장표", "광학 측정 데이터"])
        self._format_combo.setFixedWidth(150)
        self._format_combo.currentIndexChanged.connect(self._on_format_changed)
        top_row.addWidget(self._format_combo)

        top_row.addWidget(QLabel("휘도 집계:"))
        self._agg_combo = QComboBox()
        self._agg_combo.addItems(["최대값", "중간값", "최소값"])
        self._agg_combo.setFixedWidth(90)
        top_row.addWidget(self._agg_combo)

        self._btn_color = QPushButton("🎨 색상 변경")
        self._btn_color.clicked.connect(self._change_model_color)
        top_row.addWidget(self._btn_color)
        self._btn_del = QPushButton("✕ 선택 삭제")
        self._btn_del.clicked.connect(self._delete_selected)
        top_row.addWidget(self._btn_del)
        self._btn_excel = QPushButton("💾 Excel 저장")
        self._btn_excel.clicked.connect(self._export_excel)
        top_row.addWidget(self._btn_excel)
        self._btn_copy = QPushButton("📋 클립보드 복사")
        self._btn_copy.clicked.connect(self._copy_clipboard)
        top_row.addWidget(self._btn_copy)
        top_row.addStretch()
        layout.addLayout(top_row)

        self._model_list = QListWidget()
        self._model_list.setMaximumHeight(60)
        layout.addWidget(self._model_list)

        self._report_table = QTableWidget(len(self._ROW_LABELS), 2)
        self._report_table.setHorizontalHeaderLabels(["구분", "항목"])
        self._report_table.setMinimumHeight(50)
        self._report_table.verticalHeader().setDefaultSectionSize(26)  # ← 행 높이 (px)
        self._report_table.setColumnWidth(0, 80)   # ← 첫 번째 열(구분) 너비 (px)
        self._report_table.setColumnWidth(1, 80)   # ← 두 번째 열(항목) 너비 (px)

        # ── APL 차트 생성 헬퍼 ────────────────────────────────────────
        def _make_apl_chart(title: str):
            chart = QChart()
            chart.setTitle(title)
            chart.setBackgroundBrush(QColor("#ffffff"))
            chart.setMargins(QMargins(2, 2, 2, 2))

            tf = chart.titleFont(); tf.setPointSize(8); chart.setTitleFont(tf)
            chart.setTitleBrush(QColor("#1a1d2e"))

            sf = tf  # 8pt 공용
            ax = QValueAxis()
            ax.setTitleText("APL (%)"); ax.setRange(0, 100); ax.setTickCount(6)
            ax.setLabelFormat("%d")
            ax.setLabelsBrush(QColor("#6b7080")); ax.setTitleBrush(QColor("#6b7080"))
            ax.setLabelsFont(sf); ax.setTitleFont(sf)

            ay = QValueAxis()
            ay.setTitleText("Lv (cd/m²)"); ay.setLabelFormat("%d")
            ay.setLabelsBrush(QColor("#6b7080")); ay.setTitleBrush(QColor("#6b7080"))
            ay.setLabelsFont(sf); ay.setTitleFont(sf)

            chart.addAxis(ax, Qt.AlignBottom)
            chart.addAxis(ay, Qt.AlignLeft)

            # 범례 차트 내부 우상단 고정
            legend = chart.legend()
            lf = legend.font(); lf.setPointSize(7); legend.setFont(lf)
            legend.setLabelColor(QColor("#1a1d2e"))
            legend.detachFromChart()
            legend.setBackgroundVisible(True)
            legend.setBrush(QColor(255, 255, 255, 210))
            chart.plotAreaChanged.connect(
                lambda rect, c=chart: c.legend().setGeometry(
                    QRectF(rect.right() - 106, rect.top() + 4, 102, 54)
                )
            )

            view = QChartView(chart)
            view.setRenderHint(QPainter.Antialiasing)
            view.setMinimumHeight(100)
            view.setStyleSheet("background: #ffffff;")
            return chart, ax, ay, view

        # ── SDR Vivid APL 차트 (위) ────────────────────────────────────
        (self._apl_chart_sdr,
         self._apl_axis_x_sdr,
         self._apl_axis_y_sdr,
         apl_view_sdr) = _make_apl_chart("SDR Vivid")

        # ── HDR Vivid APL 차트 (아래) ──────────────────────────────────
        (self._apl_chart_hdr,
         self._apl_axis_x_hdr,
         self._apl_axis_y_hdr,
         apl_view_hdr) = _make_apl_chart("HDR Vivid")

        apl_vsplit = QSplitter(Qt.Vertical)
        apl_vsplit.addWidget(apl_view_sdr)
        apl_vsplit.addWidget(apl_view_hdr)
        apl_vsplit.setStretchFactor(0, 1)
        apl_vsplit.setStretchFactor(1, 1)

        # ── Gamut u'v' 차트 ────────────────────────────────────────────
        self._gamut_chart = QChart()
        self._gamut_chart.setTitle("u'v' 색도")
        self._gamut_chart.setBackgroundBrush(QColor("#ffffff"))
        self._gamut_chart.setMargins(QMargins(2, 2, 2, 2))
        self._gamut_chart.setTitleBrush(QColor("#1a1d2e"))

        # 범례 차트 내부 우상단 고정
        _gl = self._gamut_chart.legend()
        _gl.setLabelColor(QColor("#1a1d2e"))
        _glf = _gl.font(); _glf.setPointSize(7); _gl.setFont(_glf)
        _gl.detachFromChart()
        _gl.setBackgroundVisible(True)
        _gl.setBrush(QColor(255, 255, 255, 210))
        self._gamut_chart.plotAreaChanged.connect(
            lambda rect, c=self._gamut_chart: c.legend().setGeometry(
                QRectF(rect.right() - 86, rect.top() + 4, 82, 54)
            )
        )
        self._gamut_axis_u = QValueAxis()
        self._gamut_axis_u.setTitleText("u'")
        self._gamut_axis_u.setRange(0.0, 0.65)
        self._gamut_axis_u.setLabelsBrush(QColor("#6b7080"))
        self._gamut_axis_u.setTitleBrush(QColor("#6b7080"))
        self._gamut_axis_v = QValueAxis()
        self._gamut_axis_v.setTitleText("v'")
        self._gamut_axis_v.setRange(0.0, 0.65)
        self._gamut_axis_v.setLabelsBrush(QColor("#6b7080"))
        self._gamut_axis_v.setTitleBrush(QColor("#6b7080"))
        self._gamut_chart.addAxis(self._gamut_axis_u, Qt.AlignBottom)
        self._gamut_chart.addAxis(self._gamut_axis_v, Qt.AlignLeft)
        self._add_ref_gamuts()
        gamut_chart_view = QChartView(self._gamut_chart)
        gamut_chart_view.setRenderHint(QPainter.Antialiasing)
        gamut_chart_view.setMinimumHeight(200)
        gamut_chart_view.setMinimumWidth(200)
        gamut_chart_view.setStyleSheet("background: #ffffff;")

        # ── 데이터 테이블 ────────────────────────────────────────────
        self._report_table.setMinimumHeight(45)        # ← 테이블 최소 높이 (px)
        self._report_table.setMaximumHeight(16777215)  # 제한 없음 (줄이면 최대 높이 고정)
        # addWidget 두 번째 인수 = stretch 비율
        # 테이블 : 차트영역 = 3 : 2  →  테이블을 키우려면 3을 4·5로, 줄이려면 1·2로
        layout.addWidget(self._report_table, 1)        # ← 이 숫자가 테이블 높이 비율

        # ── 좌우: APL 차트 | u'v' 차트 ────────────────────────────────
        chart_splitter = QSplitter(Qt.Horizontal)
        chart_splitter.addWidget(apl_vsplit)
        chart_splitter.addWidget(gamut_chart_view)
        chart_splitter.setStretchFactor(0, 3)   # ← APL 차트 너비 비율  (1:1 → 같은 값 / 3:2 → 3,2)
        chart_splitter.setStretchFactor(1, 2)   # ← u'v' 차트 너비 비율
        layout.addWidget(chart_splitter, 2)     # ← 이 숫자가 차트 영역 높이 비율

        self._refresh_report_table()

    def _change_model_color(self) -> None:
        row = self._model_list.currentRow()
        if row < 0 or row >= len(self._models):
            QMessageBox.information(self, "알림", "색상을 변경할 모델을 선택하세요.")
            return
        entry = self._models[row]
        key = f"{entry['brand']}_{entry['model']}"
        current = QColor(self._model_colors.get(key, _DEFAULT_MODEL_COLORS[0]))
        color = QColorDialog.getColor(current, self, "모델 색상 선택")
        if color.isValid():
            self._model_colors[key] = color.name()
            self._refresh_apl_chart()
            self._refresh_gamut_chart()

    def _delete_selected(self) -> None:
        row = self._model_list.currentRow()
        if row < 0 or row >= len(self._models):
            return
        self._models.pop(row)
        self._model_list.takeItem(row)
        self._refresh_report_table()

    def _is_optical_format(self) -> bool:
        return self._format_combo.currentIndex() == 1

    def _on_format_changed(self) -> None:
        self._refresh_report_table()

    def _current_row_labels(self):
        return self._OPTICAL_ROW_LABELS if self._is_optical_format() else self._ROW_LABELS

    def _cell_value(self, entry: Dict, ri: int) -> str:
        if self._is_optical_format():
            k10, k100 = self._OPTICAL_KEYS[ri]
            v10  = entry.get(k10)
            v100 = entry.get(k100) if k100 else None
            if v10 is None and v100 is None:
                return "—"
            if k100 is None:
                return str(v10) if v10 is not None else "—"
            s10  = f"{v10}"  if v10  is not None else "—"
            s100 = f"{v100}" if v100 is not None else "—"
            return f"{s10} / {s100}"
        else:
            raw = entry.get(self._COMP_KEYS[ri])
            return f"{raw}" if raw is not None else "—"

    def _refresh_report_table(self) -> None:
        row_labels = self._current_row_labels()
        self._report_table.setRowCount(len(row_labels))
        self._report_table.setColumnCount(2 + len(self._models))
        headers = ["구분", "항목"] + [f"{e['brand']}_{e['model']}" for e in self._models]
        self._report_table.setHorizontalHeaderLabels(headers)

        for ri, (section, item) in enumerate(row_labels):
            for ci, val in enumerate([section, item]):
                c = QTableWidgetItem(val)
                c.setTextAlignment(Qt.AlignCenter)
                self._report_table.setItem(ri, ci, c)
            for mi, entry in enumerate(self._models):
                c = QTableWidgetItem(self._cell_value(entry, ri))
                c.setTextAlignment(Qt.AlignCenter)
                self._report_table.setItem(ri, 2 + mi, c)

    def _table_to_text(self) -> str:
        row_labels = self._current_row_labels()
        lines = []
        headers = ["구분", "항목"] + [f"{e['brand']}_{e['model']}" for e in self._models]
        lines.append("\t".join(headers))
        for ri, (section, item) in enumerate(row_labels):
            row = [section, item]
            for entry in self._models:
                row.append(self._cell_value(entry, ri))
            lines.append("\t".join(row))
        return "\n".join(lines)

    def _copy_clipboard(self) -> None:
        QApplication.clipboard().setText(self._table_to_text())
        QMessageBox.information(self, "복사 완료", "클립보드에 복사되었습니다.")

    def _export_excel(self) -> None:
        if not self._models:
            QMessageBox.information(self, "알림", "추가된 모델이 없습니다.")
            return
        brand = self._engine.brand or "brand"
        model = self._engine.model_name or "model"
        path, _ = QFileDialog.getSaveFileName(
            self, "Excel 저장", f"report_{brand}_{model}.xlsx", "Excel (*.xlsx)"
        )
        if path:
            ExcelExporter().export_report_template(brand, model, file_path=path)
            QMessageBox.information(self, "저장 완료", f"저장됨: {path}")

    # ── 파일 불러오기 ────────────────────────────────────────────────────

    def _load_files(self) -> None:
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Excel 파일 선택", "", "Excel (*.xlsx)"
        )
        if not paths:
            return
        errors: List[str] = []
        for path in paths:
            try:
                self._parse_xlsx(path)
            except Exception as exc:
                errors.append(f"{os.path.basename(path)}: {exc}")
        if errors:
            QMessageBox.warning(self, "파일 오류", "\n".join(errors))
        self._refresh_report_table()
        self._refresh_apl_chart()
        self._refresh_gamut_chart()

    def _find_or_create_entry(self, brand: str, model: str) -> Dict:
        for entry in self._models:
            if entry["brand"] == brand and entry["model"] == model:
                return entry
        key = f"{brand}_{model}"
        if key not in self._model_colors:
            idx = len(self._models)
            self._model_colors[key] = _DEFAULT_MODEL_COLORS[idx % len(_DEFAULT_MODEL_COLORS)]
        entry: Dict = {
            "brand": brand, "model": model,
            "hdr_10": None, "hdr_100": None,
            "sdr_10": None, "sdr_100": None,
            "contrast_ratio": None, "black_lv": None,
            "dci_overlap": None, "bt2020_overlap": None,
            "apl_hdr": {}, "apl_sdr": {}, "gamut_uv": {},
            # per-mode (광학 측정 데이터 format)
            "sdr_vivid_10": None,    "sdr_vivid_100": None,
            "sdr_standard_10": None, "sdr_standard_100": None,
            "hdr_vivid_10": None,    "hdr_vivid_100": None,
            "hdr_standard_10": None, "hdr_standard_100": None,
            "hdr_cinema_10": None,   "hdr_cinema_100": None,
        }
        self._models.append(entry)
        color = self._model_colors[key]
        self._model_list.addItem(f"  {brand}  {model}")
        item = self._model_list.item(self._model_list.count() - 1)
        item.setForeground(QColor(color))
        return entry

    def _parse_xlsx(self, path: str) -> None:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        try:
            sheet_names = wb.sheetnames
            brand = ""
            model = ""
            sequence = ""

            if "Info" in sheet_names:
                info_ws = wb["Info"]
                for row in info_ws.iter_rows(min_row=1, max_row=5, values_only=True):
                    if not row or row[0] is None:
                        continue
                    key = str(row[0]).strip()
                    val = str(row[1] or "").strip() if len(row) > 1 else ""
                    if key == "Brand":
                        brand = val
                    elif key == "Model":
                        model = val
                    elif key == "Sequence":
                        sequence = val

            if not sequence:
                if "Summary" in sheet_names and any(s.startswith("Raw_") for s in sheet_names):
                    sequence = "Luminance Loading"
                elif "Gamut" in sheet_names:
                    sequence = "Gamut"
                else:
                    raise ValueError(f"파일 형식을 인식할 수 없습니다 (시트: {sheet_names})")

            if not brand or not model:
                basename = os.path.splitext(os.path.basename(path))[0]
                parts = basename.split("_")
                lower = basename.lower()
                if lower.startswith("lum_loading_") and len(parts) >= 5:
                    brand = brand or parts[3]
                    model = model or "_".join(parts[4:])
                elif lower.startswith("gamut_") and len(parts) >= 3:
                    brand = brand or parts[1]
                    model = model or "_".join(parts[2:])
                else:
                    brand = brand or "Unknown"
                    model = model or basename

            # 시퀀스 자동 감지 보완 — all 파일은 여러 시트를 포함
            if not sequence:
                if any(s.startswith("Loading_") for s in sheet_names):
                    sequence = "All Sessions"
                elif "Summary" in sheet_names and any(s.startswith("Raw_") for s in sheet_names):
                    sequence = "Luminance Loading"
                elif any(s.startswith("Gamut") for s in sheet_names):
                    sequence = "Gamut"
                else:
                    raise ValueError(f"파일 형식을 인식할 수 없습니다 (시트: {sheet_names})")

            entry = self._find_or_create_entry(brand, model)

            if "All Sessions" in sequence:
                self._parse_all_sessions_wb(wb, entry)
            elif "Luminance Loading" in sequence:
                basename_lower = os.path.basename(path).lower()
                is_hdr = "hdr" in basename_lower
                mode = ("cinema" if "cinema" in basename_lower
                        else "standard" if ("standard" in basename_lower or "_std_" in basename_lower)
                        else "vivid")
                self._parse_lum_loading_wb(wb, entry, is_hdr, mode=mode)
            elif "Gamut" in sequence:
                self._parse_gamut_wb(wb, entry)
            else:
                raise ValueError(f"알 수 없는 Sequence: {sequence!r}")
        finally:
            wb.close()

    def _parse_lum_loading_wb(self, wb: Any, entry: Dict, is_hdr: bool,
                              mode: str = "") -> None:
        sheet_names = wb.sheetnames
        raw_sheets = [s for s in sheet_names if s.startswith("Raw_")]
        agg = self._agg_combo.currentText()  # "최대값" / "중간값" / "최소값"

        apl_dict: Dict[int, float] = {}

        if raw_sheets:
            apl_lv: Dict[int, List[float]] = {}
            for sheet_name in raw_sheets:
                ws = wb[sheet_name]
                for row in ws.iter_rows(min_row=2, values_only=True):
                    if not row or row[0] is None:
                        continue
                    try:
                        apl = int(round(float(row[0])))
                        lv = float(row[2])  # col3 = Lv (cd/m²)
                    except (TypeError, ValueError, IndexError):
                        continue
                    apl_lv.setdefault(apl, []).append(lv)

            for apl, lvs in apl_lv.items():
                if agg == "최대값":
                    apl_dict[apl] = round(max(lvs), 3)
                elif agg == "최소값":
                    apl_dict[apl] = round(min(lvs), 3)
                else:
                    apl_dict[apl] = round(statistics.median(lvs), 3)

        elif "Summary" in sheet_names:
            col_idx = {"최대값": 2, "최소값": 3}.get(agg, 1)  # 1=Avg, 2=Max, 3=Min
            ws = wb["Summary"]
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row or row[0] is None:
                    continue
                try:
                    apl = int(round(float(row[0])))
                    lv = float(row[col_idx])
                except (TypeError, ValueError, IndexError):
                    continue
                apl_dict[apl] = round(lv, 3)
        else:
            raise ValueError("Raw_ 또는 Summary 시트를 찾을 수 없습니다.")

        if is_hdr:
            entry["apl_hdr"].update(apl_dict)
            if 10  in apl_dict: entry["hdr_10"]  = apl_dict[10]
            if 100 in apl_dict: entry["hdr_100"] = apl_dict[100]
        else:
            entry["apl_sdr"].update(apl_dict)
            if 10  in apl_dict: entry["sdr_10"]  = apl_dict[10]
            if 100 in apl_dict: entry["sdr_100"] = apl_dict[100]

        # per-mode keys (광학 측정 데이터 format)
        m = mode.lower()
        if is_hdr:
            if "cinema" in m:
                pfx = "hdr_cinema"
            elif "standard" in m or "std" in m:
                pfx = "hdr_standard"
            else:
                pfx = "hdr_vivid"
        else:
            if "standard" in m or "std" in m:
                pfx = "sdr_standard"
            else:
                pfx = "sdr_vivid"
        if 10  in apl_dict: entry[f"{pfx}_10"]  = apl_dict[10]
        if 100 in apl_dict: entry[f"{pfx}_100"] = apl_dict[100]

    def _parse_gamut_wb(self, wb: Any, entry: Dict) -> None:
        try:
            gamut_ws = wb["Gamut"]
        except KeyError:
            raise ValueError("Gamut 시트를 찾을 수 없습니다.")

        label_map = {
            "Red": "red", "Green": "green", "Blue": "blue",
            "White": "white", "Black": "black",
        }
        rows = list(gamut_ws.iter_rows(min_row=2, max_row=6, values_only=True))
        for row in rows:
            if row[0] is None:
                continue
            color_label = label_map.get(str(row[0]).strip())
            if color_label is None:
                continue
            try:
                u_prime = float(row[5])
                v_prime = float(row[6])
            except (TypeError, ValueError, IndexError):
                continue
            entry["gamut_uv"][color_label] = (u_prime, v_prime)

        uv = entry["gamut_uv"]
        if "red" in uv and "green" in uv and "blue" in uv:
            stats = calc_gamut_stats(uv["red"], uv["green"], uv["blue"])
            entry["dci_overlap"] = stats.get("dci_overlap")
            entry["bt2020_overlap"] = stats.get("bt2020_overlap")

    def _parse_all_sessions_wb(self, wb: Any, entry: Dict) -> None:
        """_all.xlsx 파싱 — Loading_*, Gamut_* 시트를 자동 감지해 데이터 추출.

        시트별 컬럼 구조 (export_all_session 기준):
          Loading_SDR_Vivid : APL(0) | #(1) | Time(2) | Lv(3) | x(4) | y(5) | u'(6) | ...
          Gamut_SDR/HDR     : Color(0) | Time(1) | Lv(2) | x(3) | y(4) | u'(5) | v'(6) | ...
        """
        sheet_names = wb.sheetnames
        agg = self._agg_combo.currentText()
        label_map = {"Red": "red", "Green": "green", "Blue": "blue",
                     "White": "white", "Black": "black"}

        # ── Loading_* 시트 ────────────────────────────────────────────────
        loading_sheets = [s for s in sheet_names if s.startswith("Loading_Raw_")
                          or (s.startswith("Loading_") and not s == "Loading_Summary")]
        for sheet_name in loading_sheets:
            # 시트명에서 모드 판단: Loading_HDR_Vivid → HDR, Loading_SDR_Vivid → SDR
            upper = sheet_name.upper()
            is_hdr    = "_HDR_" in upper or upper.endswith("_HDR")
            is_cinema = "CINEMA"   in upper
            is_std    = "STANDARD" in upper or "_STD_" in upper

            apl_lv: Dict[int, List[float]] = {}
            ws = wb[sheet_name]
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row or row[0] is None:
                    continue
                try:
                    apl = round(int(float(row[0])))
                    lv  = float(row[3])   # APL(0) | #(1) | Time(2) | Lv(3)
                except (TypeError, ValueError, IndexError):
                    continue
                apl_lv.setdefault(apl, []).append(lv)

            apl_dict: Dict[int, float] = {}
            for apl, lvs in apl_lv.items():
                if agg == "최대값":
                    apl_dict[apl] = round(max(lvs), 3)
                elif agg == "최소값":
                    apl_dict[apl] = round(min(lvs), 3)
                else:
                    apl_dict[apl] = round(statistics.median(lvs), 3)

            if is_hdr:
                entry["apl_hdr"].update(apl_dict)
                if 10  in apl_dict: entry["hdr_10"]  = apl_dict[10]
                if 100 in apl_dict: entry["hdr_100"] = apl_dict[100]
                if is_cinema:
                    pfx = "hdr_cinema"
                elif is_std:
                    pfx = "hdr_standard"
                else:
                    pfx = "hdr_vivid"
            else:
                entry["apl_sdr"].update(apl_dict)
                if 10  in apl_dict: entry["sdr_10"]  = apl_dict[10]
                if 100 in apl_dict: entry["sdr_100"] = apl_dict[100]
                pfx = "sdr_standard" if is_std else "sdr_vivid"
            if 10  in apl_dict: entry[f"{pfx}_10"]  = apl_dict[10]
            if 100 in apl_dict: entry[f"{pfx}_100"] = apl_dict[100]

        # ── Gamut_* 시트 ──────────────────────────────────────────────────
        gamut_sheets = [s for s in sheet_names if s.startswith("Gamut_")]
        for sheet_name in gamut_sheets:
            ws = wb[sheet_name]
            for row in ws.iter_rows(min_row=2, max_row=7, values_only=True):
                if not row or row[0] is None:
                    continue
                color_label = label_map.get(str(row[0]).strip().capitalize())
                if color_label is None:
                    continue
                try:
                    u_prime = float(row[5])   # Color(0)|Time(1)|Lv(2)|x(3)|y(4)|u'(5)|v'(6)
                    v_prime = float(row[6])
                except (TypeError, ValueError, IndexError):
                    continue
                entry["gamut_uv"][color_label] = (u_prime, v_prime)

        uv = entry["gamut_uv"]
        if "red" in uv and "green" in uv and "blue" in uv:
            stats = calc_gamut_stats(uv["red"], uv["green"], uv["blue"])
            entry["dci_overlap"]   = stats.get("dci_overlap")
            entry["bt2020_overlap"] = stats.get("bt2020_overlap")

        # ── Contrast_* 시트 ───────────────────────────────────────────────
        # CR 기준: Full White Lv / Black 100% window Lv (가장 높은 CR 수치)
        # 컬럼 구조 (export_all_session 기준):
        #   Black H/V %(0) | Lv(1) | CR(2) | Time(3) | Lv_col(4) | ...
        contrast_sheets = [s for s in sheet_names if s.startswith("Contrast_")]
        for sheet_name in contrast_sheets:
            ws = wb[sheet_name]
            white_lv: float | None = None
            black100_lv: float | None = None
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row or row[0] is None:
                    continue
                try:
                    label = str(row[0]).strip()
                    lv = float(row[1])
                except (TypeError, ValueError, IndexError):
                    continue
                if label == "Full White":
                    white_lv = lv
                else:
                    try:
                        side = float(label)
                        if abs(side - 100.0) < 0.1:
                            black100_lv = lv
                    except ValueError:
                        pass
            if white_lv and black100_lv and black100_lv > 0:
                entry["contrast_ratio"] = round(white_lv / black100_lv, 1)
                entry["black_lv"] = round(black100_lv, 4)

    # ── 차트 헬퍼 ────────────────────────────────────────────────────────

    def _add_ref_gamuts(self) -> None:
        for ref_pts, color, name, dash in [
            (DCI_P3_UV, QColor("#aab0c0"), "DCI-P3",  True),
            (BT2020_UV, QColor("#aab0c0"), "BT.2020", False),
        ]:
            series = QLineSeries()
            series.setName(name)
            pen = series.pen()
            pen.setColor(color)
            pen.setWidth(1)
            if dash:
                pen.setStyle(Qt.DashLine)
            series.setPen(pen)
            for u, v in ref_pts:
                series.append(u, v)
            series.append(ref_pts[0][0], ref_pts[0][1])
            self._gamut_chart.addSeries(series)
            series.attachAxis(self._gamut_axis_u)
            series.attachAxis(self._gamut_axis_v)

    def _refresh_apl_chart(self) -> None:
        self._apl_chart_sdr.removeAllSeries()
        self._apl_chart_hdr.removeAllSeries()
        lv_sdr: List[float] = []
        lv_hdr: List[float] = []

        for entry in self._models:
            model_key = f"{entry['brand']}_{entry['model']}"
            color_hex = self._model_colors.get(model_key, _DEFAULT_MODEL_COLORS[0])
            label = f"{entry['brand']}_{entry['model']}"

            for data_key, chart, axis_x, axis_y, lv_list in [
                ("apl_sdr", self._apl_chart_sdr,
                 self._apl_axis_x_sdr, self._apl_axis_y_sdr, lv_sdr),
                ("apl_hdr", self._apl_chart_hdr,
                 self._apl_axis_x_hdr, self._apl_axis_y_hdr, lv_hdr),
            ]:
                apl_dict: Dict[int, float] = entry.get(data_key, {})
                if not apl_dict or len(apl_dict) <= 2:
                    continue
                series = QLineSeries()
                series.setName(label)
                pen = series.pen()
                pen.setColor(QColor(color_hex))
                pen.setWidth(2)
                series.setPen(pen)
                for apl in sorted(apl_dict):
                    lv = apl_dict[apl]
                    series.append(float(apl), lv)
                    lv_list.append(lv)
                chart.addSeries(series)
                series.attachAxis(axis_x)
                series.attachAxis(axis_y)

        if lv_sdr:
            self._apl_axis_y_sdr.setRange(0, max(lv_sdr) * 1.15)
        if lv_hdr:
            self._apl_axis_y_hdr.setRange(0, max(lv_hdr) * 1.15)

    def _refresh_gamut_chart(self) -> None:
        self._gamut_chart.removeAllSeries()
        self._add_ref_gamuts()

        _SCATTER_COLOR = {
            "red":   QColor("#e74c3c"),
            "green": QColor("#27ae60"),
            "blue":  QColor("#4f8ef7"),
        }

        for entry in self._models:
            uv: Dict[str, tuple] = entry.get("gamut_uv", {})
            if not uv:
                continue
            model_key = f"{entry['brand']}_{entry['model']}"
            color_hex = self._model_colors.get(model_key, _DEFAULT_MODEL_COLORS[0])

            # R→G→B→R 삼각형 라인 (범례에 brand_model로 표시)
            if "red" in uv and "green" in uv and "blue" in uv:
                tri = QLineSeries()
                tri.setName(f"{entry['brand']}_{entry['model']}")
                pen = tri.pen()
                pen.setColor(QColor(color_hex))
                pen.setWidth(2)
                tri.setPen(pen)
                for key in ("red", "green", "blue", "red"):
                    u, v = uv[key]
                    tri.append(u, v)
                self._gamut_chart.addSeries(tri)
                tri.attachAxis(self._gamut_axis_u)
                tri.attachAxis(self._gamut_axis_v)

            # R/G/B 포인트만 표시 (white/black 제거), 범례에서는 숨김
            for color_key in ("red", "green", "blue"):
                if color_key not in uv:
                    continue
                u, v = uv[color_key]
                dot = QScatterSeries()
                dot.setName(color_key)
                dot.setMarkerSize(10.0)
                dot_color = _SCATTER_COLOR[color_key]
                dot.setColor(dot_color)
                dot.setBorderColor(dot_color)
                dot.append(u, v)
                self._gamut_chart.addSeries(dot)
                dot.attachAxis(self._gamut_axis_u)
                dot.attachAxis(self._gamut_axis_v)
                # 범례에서 R/G/B 개별 점 항목 숨김
                for marker in self._gamut_chart.legend().markers(dot):
                    marker.setVisible(False)


# ---------------------------------------------------------------------------
# Main Window
# ---------------------------------------------------------------------------

class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("패널 측정 프로그램 v0.1")
        self.setMinimumSize(1024, 1000) # 프로그램 전체 크기 조절
        self.setStyleSheet(_DARK_STYLE)

        self._engine = MeasurementEngine()

        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        self._conn_panel = ConnectionPanel(self._engine)
        self._conn_panel.setMaximumHeight(172)
        root.addWidget(self._conn_panel)

        splitter = QSplitter(Qt.Horizontal)
        splitter.setHandleWidth(1)

        self._sidebar = QListWidget()
        self._sidebar.setMaximumWidth(210)
        self._sidebar.addItems([
            "🎯  1. 센터 맞추기",
            "📈  2. 휘도 스윙",
            "📊  3. APL 로딩",
            "🎨  4. 색재현율",
            "⬛  5. 명암비",
            "📋  6. 보고서 템플릿",
        ])
        splitter.addWidget(self._sidebar)

        self._stack = QStackedWidget()
        self._lum_panel = LumLoadingPanel(self._engine)
        self._gamut_panel = GamutPanel(self._engine)
        self._contrast_panel = ContrastPanel(self._engine)
        self._report_panel = ReportPanel(
            self._engine, self._gamut_panel, self._lum_panel, self._contrast_panel
        )
        for panel in (
            CenterAlignPanel(self._engine),
            LumSwingPanel(self._engine),
            self._lum_panel,
            self._gamut_panel,
            self._contrast_panel,
            self._report_panel,
        ):
            panel.setContentsMargins(20, 16, 20, 16)
            self._stack.addWidget(panel)
        splitter.addWidget(self._stack)
        splitter.setStretchFactor(1, 1)

        root.addWidget(splitter)

        self._sidebar.currentRowChanged.connect(self._stack.setCurrentIndex)
        self._sidebar.setCurrentRow(0)

        self.setStatusBar(QStatusBar())
        self.statusBar().showMessage("준비 — 장비를 연결하세요")

    def closeEvent(self, event: Any) -> None:
        self._engine.disconnect_all()
        super().closeEvent(event)
