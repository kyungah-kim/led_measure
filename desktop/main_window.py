from __future__ import annotations

import os
import statistics
from typing import Any, Dict, List, Optional

import openpyxl

from PySide6.QtCharts import QChart, QChartView, QLineSeries, QScatterSeries, QValueAxis
from PySide6.QtCore import QMargins, Qt, Slot
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

_DARK_STYLE = """
QMainWindow, QWidget { background: #f5f6fa; color: #1a1d2e; font-size: 13px; }
QGroupBox { border: 1px solid #d0d3e0; border-radius: 6px; margin-top: 8px; padding: 8px;
            background: #ffffff; }
QGroupBox::title { color: #6b7080; font-size: 11px; text-transform: uppercase; }
QPushButton { background: #ffffff; border: 1px solid #c8ccd8; border-radius: 5px;
              padding: 6px 14px; color: #1a1d2e; }
QPushButton:hover { background: #eef0f8; }
QPushButton#primary { background: #4f8ef7; border-color: #4f8ef7; color: white; font-weight: bold; }
QPushButton#primary:hover { background: #3a7ae8; }
QPushButton#danger { background: #e74c3c; border-color: #e74c3c; color: white; }
QPushButton#warning { background: #e67e22; border-color: #e67e22; color: white; }
QPushButton#success { background: #27ae60; border-color: #27ae60; color: white; font-weight: bold; }
QPushButton:disabled { color: #aab0c0; border-color: #dde0ea; }
QComboBox, QLineEdit { background: #ffffff; border: 1px solid #c8ccd8;
                       border-radius: 4px; padding: 5px 8px; color: #1a1d2e; }
QListWidget { background: #ffffff; border: 1px solid #d0d3e0; border-radius: 4px; }
QListWidget::item { padding: 8px 12px; }
QListWidget::item:selected { background: rgba(79,142,247,0.12); color: #2d6fd6; border-left: 3px solid #4f8ef7; }
QProgressBar { background: #e8eaf0; border: 1px solid #c8ccd8; border-radius: 3px; height: 8px; }
QProgressBar::chunk { background: #4f8ef7; border-radius: 3px; }
QTableWidget { background: #ffffff; gridline-color: #dde0ea; border: 1px solid #d0d3e0; }
QHeaderView::section { background: #f0f2f8; color: #6b7080; border: none;
                        border-bottom: 1px solid #d0d3e0; padding: 5px 8px;
                        font-size: 11px; font-weight: bold; }
QCheckBox { spacing: 6px; }
QLabel#status_ok { color: #27ae60; font-weight: bold; }
QLabel#status_err { color: #e74c3c; font-weight: bold; }
QLabel#muted { color: #8890a8; font-size: 12px; }
QSplitter::handle { background: #d0d3e0; }
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
        self._brand_edit = QLineEdit("Samsung")
        self._model_edit = QLineEdit("QN65S95D")
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

        # ── 하단 행: Mock 버튼 (추후 제거 예정, 초록/기본/빨강으로 구분) ────────
        mock_row = QHBoxLayout()
        mock_row.setSpacing(4)
        self._btn_mock_meter = QPushButton("Mock 미터")
        self._btn_mock_meter.setObjectName("success")
        self._btn_mock_meter.clicked.connect(self._connect_mock_meter)
        self._btn_mock = QPushButton("Mock 전체")
        self._btn_mock.clicked.connect(self._connect_mock)
        self._btn_dis_all = QPushButton("전체 해제")
        self._btn_dis_all.setObjectName("danger")
        self._btn_dis_all.clicked.connect(self._disconnect_all)
        mock_row.addWidget(btn_scan)
        mock_row.addStretch()
        mock_row.addWidget(self._btn_mock_meter)
        mock_row.addWidget(self._btn_mock)
        mock_row.addWidget(self._btn_dis_all)
        root.addLayout(mock_row)

        # 시작 시 포트 목록 채우기
        self._scan_ports()

    def _scan_ports(self) -> None:
        """시리얼 포트를 스캔해 두 콤보박스를 갱신한다."""
        import serial.tools.list_ports
        ports = sorted(p.device for p in serial.tools.list_ports.comports())

        for combo in (self._meter_port, self._gen_port):
            current = combo.currentText()
            combo.clear()
            combo.addItems(ports)
            # 이전 선택 포트가 아직 존재하면 유지
            idx = combo.findText(current)
            if idx >= 0:
                combo.setCurrentIndex(idx)

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

    def _connect_mock_meter(self) -> None:
        from core.equipment.mock import MockMeter
        m = MockMeter()
        m.connect("MOCK")
        self._engine.meter = m
        self._engine.brand = self._brand_edit.text().strip() or "Samsung"
        self._engine.model_name = self._model_edit.text().strip() or "QN65S95D"
        self._meter_status.setText("Mock 연결됨")
        self._meter_status.setStyleSheet("color:#d4820a;font-weight:bold;")
        self._btn_meter.setEnabled(False)
        self._btn_meter_dis.setEnabled(True)

    def _connect_mock(self) -> None:
        from core.equipment.mock import MockMeter, MockGenerator
        m = MockMeter()
        m.connect("MOCK")
        g = MockGenerator()
        g.connect("MOCK")
        self._engine.meter = m
        self._engine.generator = g
        self._engine.brand = self._brand_edit.text().strip() or "Samsung"
        self._engine.model_name = self._model_edit.text().strip() or "QN65S95D"
        self._meter_status.setText("Mock 연결됨")
        self._meter_status.setStyleSheet("color:#d4820a;font-weight:bold;")
        self._gen_status.setText("Mock 연결됨")
        self._gen_status.setStyleSheet("color:#d4820a;font-weight:bold;")
        self._btn_meter.setEnabled(False)
        self._btn_meter_dis.setEnabled(True)
        self._btn_gen.setEnabled(False)
        self._btn_gen_dis.setEnabled(True)

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
        layout.setAlignment(Qt.AlignTop)

        title = QLabel("🎯 센터 맞추기")
        title.setStyleSheet("font-size:16px;font-weight:bold;margin-bottom:4px;")
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

class LumSwingPanel(QWidget):
    def __init__(self, engine: MeasurementEngine, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._engine = engine
        self._worker: Optional[MeasurementWorker] = None
        self._series: Dict[str, QLineSeries] = {}
        self._rows: List[MeasureResult] = []

        layout = QVBoxLayout(self)

        title = QLabel("📈 휘도 스윙 (Luminance Swing)")
        title.setStyleSheet("font-size:16px;font-weight:bold;margin-bottom:4px;")
        layout.addWidget(title)

        ctrl = QHBoxLayout()
        ctrl.addWidget(QLabel("PSM:"))
        self._case_combo = QComboBox()
        self._case_combo.addItems(["Vivid", "Standard", "Cinema"])
        ctrl.addWidget(self._case_combo)
        self._hdr_check = QCheckBox("HDR 자동 전환")
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
        self._btn_export = QPushButton("💾  Excel 저장")
        self._btn_export.clicked.connect(self._export)
        ctrl.addWidget(self._btn_export)
        ctrl.addStretch()
        layout.addLayout(ctrl)

        self._progress = QProgressBar()
        layout.addWidget(self._progress)

        # Chart
        self._chart = QChart()
        self._chart.setTitle("BPL Swing")
        self._chart.setBackgroundBrush(QColor("#ffffff"))
        self._chart.setTitleBrush(QColor("#1a1d2e"))
        self._axis_x = QValueAxis()
        self._axis_x.setTitleText("측정 #")
        self._axis_x.setLabelsBrush(QColor("#6b7080"))
        self._axis_x.setTitleBrush(QColor("#6b7080"))
        self._axis_y = QValueAxis()
        self._axis_y.setTitleText("Lv (cd/m²)")
        self._axis_y.setLabelsBrush(QColor("#6b7080"))
        self._axis_y.setTitleBrush(QColor("#6b7080"))
        self._chart.addAxis(self._axis_x, Qt.AlignBottom)
        self._chart.addAxis(self._axis_y, Qt.AlignLeft)
        self._chart.legend().hide()
        chart_view = QChartView(self._chart)
        chart_view.setRenderHint(QPainter.Antialiasing)
        chart_view.setMinimumHeight(280)
        chart_view.setStyleSheet("background: #ffffff;")
        layout.addWidget(chart_view)

        self._status_label = QLabel("")
        self._status_label.setObjectName("muted")
        layout.addWidget(self._status_label)

        # Table (last 10 rows)
        self._table = QTableWidget(0, 6)
        self._table.setHorizontalHeaderLabels(["#", "Lv (cd/m²)", "x", "y", "u'", "v'"])
        self._table.setMaximumHeight(180)
        layout.addWidget(self._table)

    def _run(self) -> None:
        case_text = self._case_combo.currentText()
        case = case_text[0]
        is_hdr = self._hdr_check.isChecked()
        self._rows.clear()
        # fresh series
        series = QLineSeries()
        series.setName(f"Case {case}")
        pen = series.pen()
        pen.setColor(QColor("#4f8ef7"))
        pen.setWidth(2)
        series.setPen(pen)
        self._chart.removeAllSeries()
        self._chart.addSeries(series)
        series.attachAxis(self._axis_x)
        series.attachAxis(self._axis_y)
        self._series = {case: series}
        self._axis_x.setRange(0, 30)
        self._axis_y.setRange(0, 600)

        self._worker = MeasurementWorker(self._engine, "lum_swing",
                                          case=case, is_hdr=is_hdr)
        self._worker.progress.connect(self._on_progress)
        self._worker.succeeded.connect(self._on_finished)
        self._worker.error.connect(self._on_error)
        wire_worker_cleanup(self._worker, self, '_worker')
        self._worker.start()
        self._btn_run.setEnabled(False)
        self._btn_stop.setEnabled(True)

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

    def _export(self) -> None:
        if not self._rows:
            QMessageBox.information(self, "알림", "저장할 데이터가 없습니다.")
            return
        case = self._case_combo.currentText().split()[0].lower()
        mode = "HDR" if self._hdr_check.isChecked() else "SDR"
        brand = self._engine.brand or "brand"
        model = self._engine.model_name or "model"
        default_name = f"lum_swing_{mode}_{case}_{brand}_{model}.xlsx"
        path, _ = QFileDialog.getSaveFileName(self, "Excel 저장", default_name,
                                               "Excel (*.xlsx)")
        if path:
            ExcelExporter().export_lum_swing(
                {case: self._rows},
                self._engine.brand, self._engine.model_name,
                file_path=path,
            )
            QMessageBox.information(self, "저장 완료", f"저장됨: {path}")

    @Slot(str, float, object)
    def _on_progress(self, _step: str, pct: float, data: Any) -> None:
        self._progress.setValue(int(pct * 100))
        if isinstance(data, MeasureResult):
            self._rows.append(data)
            case = self._case_combo.currentText()[0]
            series = self._series.get(case)
            if series:
                n = len(self._rows)
                series.append(float(n), data.Lv)
                self._axis_x.setMax(max(self._axis_x.max(), float(n) + 2))
                self._axis_y.setMax(max(self._axis_y.max(), data.Lv * 1.15))
            self._update_table()
            self._status_label.setText(
                f"측정 중 — #{len(self._rows)}  Lv={data.Lv:.3f}  x={data.x:.4f}  y={data.y:.4f}"
            )

    @Slot(object)
    def _on_finished(self, _result: Any) -> None:
        self._progress.setValue(100)
        self._status_label.setText(f"완료 — {len(self._rows)}건 측정")
        self._btn_run.setEnabled(True)
        self._btn_stop.setEnabled(False)

    @Slot(str)
    def _on_error(self, msg: str) -> None:
        QMessageBox.critical(self, "오류", msg)
        self._btn_run.setEnabled(True)
        self._btn_stop.setEnabled(False)

    def _update_table(self) -> None:
        rows = self._rows[-8:]
        self._table.setRowCount(len(rows))
        for ri, r in enumerate(rows):
            for ci, val in enumerate([len(self._rows) - len(rows) + ri + 1,
                                       f"{r.Lv:.3f}", f"{r.x:.4f}", f"{r.y:.4f}",
                                       f"{r.u_prime:.4f}", f"{r.v_prime:.4f}"]):
                item = QTableWidgetItem(str(val))
                item.setTextAlignment(Qt.AlignCenter)
                self._table.setItem(ri, ci, item)


# ---------------------------------------------------------------------------
# Luminance Loading Panel
# ---------------------------------------------------------------------------

_BRAND_COLOR_QT = {"samsung": "#0070C0", "lg": "#FF0000", "sony": "#00B050"}


def _qt_brand_color(brand: str) -> str:
    return _BRAND_COLOR_QT.get(brand.lower().strip(), "#FF8800")


class LumLoadingPanel(QWidget):
    def __init__(self, engine: MeasurementEngine, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._engine = engine
        self._results: Dict[str, Any] = {}
        self._raw_data: Dict[int, List[MeasureResult]] = {}
        self._last_hdr_raw: Dict[int, List[MeasureResult]] = {}
        self._last_sdr_raw: Dict[int, List[MeasureResult]] = {}
        layout = QVBoxLayout(self)

        title = QLabel("📊 APL 로딩 (Luminance Loading)")
        title.setStyleSheet("font-size:16px;font-weight:bold;margin-bottom:4px;")
        layout.addWidget(title)

        ctrl = QFormLayout()
        self._version_combo = QComboBox()
        self._version_combo.addItems(["37단계", "10단계", "2단계"])
        ctrl.addRow("버전:", self._version_combo)
        self._case_combo = QComboBox()
        self._case_combo.addItems(["Vivid", "Standard", "Cinema"])
        ctrl.addRow("케이스:", self._case_combo)
        self._hdr_check = QCheckBox("HDR")
        self._hdr_check.stateChanged.connect(self._on_hdr_toggled)
        ctrl.addRow("", self._hdr_check)
        self._cooling_check = QCheckBox("쿨링 (APL≤10 측정 전 Black 출력)")
        ctrl.addRow("", self._cooling_check)
        self._meas_count = QSpinBox()
        self._meas_count.setRange(1, 20)
        self._meas_count.setValue(1)
        self._meas_count.setSuffix(" 회")
        ctrl.addRow("패턴당 측정 횟수:", self._meas_count)
        layout.addLayout(ctrl)

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
        self._btn_clear = QPushButton("🗑  그래프 초기화")
        self._btn_clear.clicked.connect(self._clear_chart)
        btn_row.addWidget(self._btn_clear)
        btn_row.addStretch()
        layout.addLayout(btn_row)

        self._progress = QProgressBar()
        layout.addWidget(self._progress)
        self._status_label = QLabel("대기 중")
        self._status_label.setObjectName("muted")
        layout.addWidget(self._status_label)

        # APL vs Lv 인라인 차트 (HDR + SDR 두 시리즈)
        self._apl_chart = QChart()
        self._apl_chart.setTitle("APL vs Lv")
        self._apl_chart.setBackgroundBrush(QColor("#ffffff"))
        self._apl_chart.setTitleBrush(QColor("#1a1d2e"))
        self._apl_chart.legend().setLabelColor(QColor("#1a1d2e"))
        self._apl_chart.legend().show()

        self._apl_series_hdr = QLineSeries()
        self._apl_series_hdr.setName("HDR")
        self._apl_series_sdr = QLineSeries()
        self._apl_series_sdr.setName("SDR")

        self._apl_axis_x = QValueAxis()
        self._apl_axis_x.setTitleText("APL (%)")
        self._apl_axis_x.setRange(0, 100)
        self._apl_axis_x.setTickCount(11)
        self._apl_axis_x.setLabelsBrush(QColor("#6b7080"))
        self._apl_axis_x.setTitleBrush(QColor("#6b7080"))
        self._apl_axis_y = QValueAxis()
        self._apl_axis_y.setTitleText("Lv (cd/m²)")
        self._apl_axis_y.setLabelsBrush(QColor("#6b7080"))
        self._apl_axis_y.setTitleBrush(QColor("#6b7080"))
        self._apl_chart.addAxis(self._apl_axis_x, Qt.AlignBottom)
        self._apl_chart.addAxis(self._apl_axis_y, Qt.AlignLeft)

        for series, color, dash in [
            (self._apl_series_hdr, QColor("#e74c3c"), False),
            (self._apl_series_sdr, QColor("#4f8ef7"), True),
        ]:
            self._apl_chart.addSeries(series)
            pen = series.pen()
            pen.setColor(color)
            pen.setWidth(2)
            if dash:
                pen.setStyle(Qt.DashLine)
            series.setPen(pen)
            series.attachAxis(self._apl_axis_x)
            series.attachAxis(self._apl_axis_y)

        apl_chart_view = QChartView(self._apl_chart)
        apl_chart_view.setRenderHint(QPainter.Antialiasing)
        apl_chart_view.setMinimumHeight(220)
        apl_chart_view.setStyleSheet("background: #ffffff;")
        layout.addWidget(apl_chart_view)

        self._table = QTableWidget(0, 9)
        self._table.setHorizontalHeaderLabels(["APL%", "#", "Lv (cd/m²)", "x", "y", "u'", "v'", "CCT (K)", "Duv"])
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

    def _run(self) -> None:
        version_map = {"37단계": "37", "10단계": "10", "2단계": "2"}
        version = version_map[self._version_combo.currentText()]
        self._raw_data.clear()
        self._table.setRowCount(0)
        self._btn_run.setEnabled(False)
        self._worker = MeasurementWorker(
            self._engine, "lum_loading",
            version=version,
            case=self._case_combo.currentText(),
            is_hdr=self._hdr_check.isChecked(),
            cooling_enabled=self._cooling_check.isChecked(),
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

    @Slot(object)
    def _on_finished(self, result: Any) -> None:
        self._results = result or {}
        if self._hdr_check.isChecked():
            self._last_hdr_raw = dict(self._raw_data)
        else:
            self._last_sdr_raw = dict(self._raw_data)
        self._status_label.setText(f"완료 — {len(self._raw_data)}개 APL 측정")
        self._btn_run.setEnabled(True)
        self._btn_stop.setEnabled(False)
        self._refresh_apl_chart()

    @Slot(str, float, object)
    def _on_progress(self, _step: str, pct: float, data: Any) -> None:
        self._progress.setValue(int(pct * 100))
        if isinstance(data, dict) and "apl" in data:
            apl = int(data["apl"])
            results: List[MeasureResult] = data.get("results", [])
            self._raw_data[apl] = results
            self._refresh_table()
            self._table.scrollToBottom()
            lv_avg = sum(r.Lv for r in results) / len(results) if results else 0
            self._status_label.setText(f"APL {apl}% — Lv={lv_avg:.3f} cd/m²  ({int(pct*100)}%)")

    def _refresh_table(self) -> None:
        self._table.setRowCount(0)
        for apl in sorted(self._raw_data):
            results = self._raw_data[apl]
            for idx, r in enumerate(results, start=1):
                self._add_table_row(f"{apl}%", str(idx), r)
        self._refresh_apl_chart()

    def _refresh_apl_chart(self) -> None:
        self._apl_series_hdr.clear()
        self._apl_series_sdr.clear()

        all_lv: List[float] = []

        # 측정 중이면 현재 raw_data를 해당 모드 시리즈에 실시간 반영
        is_hdr_mode = self._hdr_check.isChecked()
        live_series = self._apl_series_hdr if is_hdr_mode else self._apl_series_sdr

        for apl in sorted(self._raw_data):
            results = self._raw_data[apl]
            if not results:
                continue
            lv = sum(r.Lv for r in results) / len(results)
            live_series.append(float(apl), lv)
            all_lv.append(lv)

        # 완료된 반대 모드 데이터도 함께 표시
        other_raw = self._last_sdr_raw if is_hdr_mode else self._last_hdr_raw
        other_series = self._apl_series_sdr if is_hdr_mode else self._apl_series_hdr
        for apl in sorted(other_raw):
            results = other_raw[apl]
            if not results:
                continue
            lv = sum(r.Lv for r in results) / len(results)
            other_series.append(float(apl), lv)
            all_lv.append(lv)

        if all_lv:
            self._apl_axis_y.setRange(0, max(all_lv) * 1.15)

    def _clear_chart(self) -> None:
        self._last_hdr_raw.clear()
        self._last_sdr_raw.clear()
        self._raw_data.clear()
        self._apl_series_hdr.clear()
        self._apl_series_sdr.clear()
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
            except Exception as exc:
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

        title = QLabel("🎨 색재현율 (Gamut)")
        title.setStyleSheet("font-size:16px;font-weight:bold;margin-bottom:4px;")
        layout.addWidget(title)
        desc = QLabel("Full Pattern (100%) — R → G → B → W → BK 순으로 자동 측정합니다.")
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

        # 통계 레이블 (DCI-P3 / BT.2020 Coverage)
        stats_row = QHBoxLayout()
        self._lbl_dci   = QLabel("DCI-P3: —")
        self._lbl_bt2020 = QLabel("BT.2020: —")
        for lbl in (self._lbl_dci, self._lbl_bt2020):
            lbl.setStyleSheet("font-weight:bold; font-size:13px; padding:2px 10px;")
            stats_row.addWidget(lbl)
        stats_row.addStretch()
        layout.addLayout(stats_row)

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
        chart_view.setMinimumWidth(420)
        chart_view.setMinimumHeight(420)
        chart_view.setStyleSheet("background: #ffffff;")
        splitter.addWidget(chart_view)

        self._table = QTableWidget(0, 9)
        self._table.setHorizontalHeaderLabels(["컬러", "Lv", "x", "y", "u'", "v'", "X", "Y", "Z"])
        splitter.addWidget(self._table)
        splitter.setStretchFactor(0, 3)
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

    @Slot(object)
    def _on_finished(self, result: Any) -> None:
        self._results = result or {}
        self._status_label.setText("완료")
        self._btn_run.setEnabled(True)
        self._btn_stop.setEnabled(False)
        # 색재현율 통계 계산
        self._update_gamut_stats()
        # 측정 삼각형(R-G-B) 라인 추가
        self._draw_meas_triangle()

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
            ExcelExporter().export_gamut(self._results, self._engine.brand, self._engine.model_name,
                                         file_path=path)
            QMessageBox.information(self, "저장 완료", f"저장됨: {path}")


# ---------------------------------------------------------------------------
# Contrast Panel
# ---------------------------------------------------------------------------

class ContrastPanel(QWidget):
    def __init__(self, engine: MeasurementEngine, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._engine = engine
        self._results: Dict[float, MeasureResult] = {}
        layout = QVBoxLayout(self)

        title = QLabel("⬛ 명암비 (Contrast Ratio)")
        title.setStyleSheet("font-size:16px;font-weight:bold;margin-bottom:4px;")
        layout.addWidget(title)
        desc = QLabel("White Raster + Black Window — 100% / 50% / 20% / 14.1% / 0% 순으로 측정합니다.")
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

        self._table = QTableWidget(0, 5)
        self._table.setHorizontalHeaderLabels(["윈도우 크기", "Lv (cd/m²)", "x", "y", "비고"])
        layout.addWidget(self._table)

    @Slot(int)
    def _on_hdr_toggled(self, state: int) -> None:
        gen = self._engine.generator
        if gen is None or not gen.is_connected:
            return
        fn = lambda: gen.set_hdr(bool(state))
        worker = ConnectWorker(fn)
        worker.error.connect(lambda msg: QMessageBox.critical(self, "HDR 전환 오류", msg))
        wire_worker_cleanup(worker, self, '_hdr_worker')
        worker.start()
        self._hdr_worker = worker

    def _run(self) -> None:
        self._table.setRowCount(0)
        self._btn_run.setEnabled(False)
        self._btn_stop.setEnabled(True)
        self._worker = MeasurementWorker(self._engine, "contrast", is_hdr=self._hdr_check.isChecked())
        self._worker.progress.connect(self._on_progress)
        self._worker.succeeded.connect(lambda r: (
            setattr(self, '_results', r or {}),
            self._status_label.setText("완료"),
            self._btn_run.setEnabled(True),
            self._btn_stop.setEnabled(False),
        ))
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
        if isinstance(data, dict) and "win_size" in data:
            win_size = data["win_size"]
            r = data.get("result")
            if r:
                row = self._table.rowCount()
                self._table.insertRow(row)
                for ci, val in enumerate([f"{win_size}%", f"{r.Lv:.4f}",
                                           f"{r.x:.4f}", f"{r.y:.4f}", ""]):
                    item = QTableWidgetItem(str(val))
                    item.setTextAlignment(Qt.AlignCenter)
                    self._table.setItem(row, ci, item)
            self._status_label.setText(f"윈도우 {win_size}% 측정 — {int(pct*100)}%")

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

        title = QLabel("📋 보고서 템플릿")
        title.setStyleSheet("font-size:16px;font-weight:bold;margin-bottom:4px;")
        layout.addWidget(title)

        top_row = QHBoxLayout()
        self._btn_load = QPushButton("📂 파일 불러오기")
        self._btn_load.setObjectName("primary")
        self._btn_load.clicked.connect(self._load_files)
        top_row.addWidget(self._btn_load)

        top_row.addWidget(QLabel("White 휘도 집계:"))
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
        self._model_list.setMaximumHeight(80)
        layout.addWidget(self._model_list)

        n_rows = len(self._ROW_LABELS)
        self._report_table = QTableWidget(n_rows, 2)
        self._report_table.setHorizontalHeaderLabels(["구분", "항목"])
        for ri, (section, item) in enumerate(self._ROW_LABELS):
            for ci, val in enumerate([section, item]):
                c = QTableWidgetItem(val)
                c.setTextAlignment(Qt.AlignCenter)
                self._report_table.setItem(ri, ci, c)
        layout.addWidget(self._report_table)

        # ── APL 차트 생성 헬퍼 ────────────────────────────────────────
        def _make_apl_chart(title: str):
            chart = QChart()
            chart.setTitle(title)
            chart.setBackgroundBrush(QColor("#ffffff"))
            chart.setTitleBrush(QColor("#1a1d2e"))
            chart.legend().setLabelColor(QColor("#1a1d2e"))
            ax = QValueAxis()
            ax.setTitleText("APL (%)")
            ax.setRange(0, 100)
            ax.setTickCount(11)
            ax.setLabelsBrush(QColor("#6b7080"))
            ax.setTitleBrush(QColor("#6b7080"))
            ay = QValueAxis()
            ay.setTitleText("Lv (cd/m²)")
            ay.setLabelsBrush(QColor("#6b7080"))
            ay.setTitleBrush(QColor("#6b7080"))
            chart.addAxis(ax, Qt.AlignBottom)
            chart.addAxis(ay, Qt.AlignLeft)
            view = QChartView(chart)
            view.setRenderHint(QPainter.Antialiasing)
            view.setStyleSheet("background: #ffffff;")
            return chart, ax, ay, view

        # ── SDR Vivid APL 차트 (위) ────────────────────────────────────
        (self._apl_chart_sdr,
         self._apl_axis_x_sdr,
         self._apl_axis_y_sdr,
         apl_view_sdr) = _make_apl_chart("APL vs Lv  [SDR Vivid]")

        # ── HDR Vivid APL 차트 (아래) ──────────────────────────────────
        (self._apl_chart_hdr,
         self._apl_axis_x_hdr,
         self._apl_axis_y_hdr,
         apl_view_hdr) = _make_apl_chart("APL vs Lv  [HDR Vivid]")

        apl_vsplit = QSplitter(Qt.Vertical)
        apl_vsplit.addWidget(apl_view_sdr)
        apl_vsplit.addWidget(apl_view_hdr)
        apl_vsplit.setStretchFactor(0, 1)
        apl_vsplit.setStretchFactor(1, 1)

        # ── Gamut u'v' 차트 ────────────────────────────────────────────
        self._gamut_chart = QChart()
        self._gamut_chart.setTitle("u'v' 색도")
        self._gamut_chart.setBackgroundBrush(QColor("#ffffff"))
        self._gamut_chart.setTitleBrush(QColor("#1a1d2e"))
        self._gamut_chart.legend().setLabelColor(QColor("#1a1d2e"))
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
        gamut_chart_view.setMinimumHeight(700)
        gamut_chart_view.setMinimumWidth(400)
        gamut_chart_view.setStyleSheet("background: #ffffff;")

        chart_splitter = QSplitter(Qt.Horizontal)
        chart_splitter.addWidget(apl_vsplit)
        chart_splitter.addWidget(gamut_chart_view)
        chart_splitter.setStretchFactor(0, 2)
        chart_splitter.setStretchFactor(1, 3)
        layout.addWidget(chart_splitter)

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

    def _refresh_report_table(self) -> None:
        n_models = len(self._models)
        self._report_table.setColumnCount(2 + n_models)
        headers = ["구분", "항목"] + [f"{e['brand']}_{e['model']}" for e in self._models]
        self._report_table.setHorizontalHeaderLabels(headers)

        keys = ["hdr_10", "hdr_100", "sdr_10", "sdr_100",
                "contrast_ratio", "black_lv", "dci_overlap", "bt2020_overlap"]

        for ri, (section, item) in enumerate(self._ROW_LABELS):
            for ci, val in enumerate([section, item]):
                c = QTableWidgetItem(val)
                c.setTextAlignment(Qt.AlignCenter)
                self._report_table.setItem(ri, ci, c)
            for mi, entry in enumerate(self._models):
                raw = entry.get(keys[ri])
                text = f"{raw}" if raw is not None else "—"
                c = QTableWidgetItem(text)
                c.setTextAlignment(Qt.AlignCenter)
                self._report_table.setItem(ri, 2 + mi, c)

    def _table_to_text(self) -> str:
        lines = []
        headers = ["구분", "항목"] + [f"{e['brand']}_{e['model']}" for e in self._models]
        lines.append("\t".join(headers))
        keys = ["hdr_10", "hdr_100", "sdr_10", "sdr_100",
                "contrast_ratio", "black_lv", "dci_overlap", "bt2020_overlap"]
        for ri, (section, item) in enumerate(self._ROW_LABELS):
            row = [section, item]
            for entry in self._models:
                raw = entry.get(keys[ri])
                row.append(str(raw) if raw is not None else "—")
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
            ExcelExporter().export_report_template(self._models, file_path=path)
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

            entry = self._find_or_create_entry(brand, model)

            if "Luminance Loading" in sequence:
                is_hdr = "hdr" in os.path.basename(path).lower()
                self._parse_lum_loading_wb(wb, entry, is_hdr)
            elif "Gamut" in sequence:
                self._parse_gamut_wb(wb, entry)
            else:
                raise ValueError(f"알 수 없는 Sequence: {sequence!r}")
        finally:
            wb.close()

    def _parse_lum_loading_wb(self, wb: Any, entry: Dict, is_hdr: bool) -> None:
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
            if 10 in apl_dict:
                entry["hdr_10"] = apl_dict[10]
            if 100 in apl_dict:
                entry["hdr_100"] = apl_dict[100]
        else:
            entry["apl_sdr"].update(apl_dict)
            if 10 in apl_dict:
                entry["sdr_10"] = apl_dict[10]
            if 100 in apl_dict:
                entry["sdr_100"] = apl_dict[100]

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

    # ── 차트 헬퍼 ────────────────────────────────────────────────────────

    def _add_ref_gamuts(self) -> None:
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
            "white": QColor("#888899"),
            "black": QColor("#aab0c0"),
        }

        for entry in self._models:
            uv: Dict[str, tuple] = entry.get("gamut_uv", {})
            if not uv:
                continue
            model_key = f"{entry['brand']}_{entry['model']}"
            color_hex = self._model_colors.get(model_key, _DEFAULT_MODEL_COLORS[0])

            # R→G→B→R 삼각형 라인
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

            # 각 색 포인트 ScatterSeries
            for color_key, (u, v) in uv.items():
                dot = QScatterSeries()
                dot.setName(color_key)
                dot.setMarkerSize(10.0)
                dot_color = _SCATTER_COLOR.get(color_key, QColor("#888899"))
                dot.setColor(dot_color)
                dot.setBorderColor(dot_color)
                dot.append(u, v)
                self._gamut_chart.addSeries(dot)
                dot.attachAxis(self._gamut_axis_u)
                dot.attachAxis(self._gamut_axis_v)


# ---------------------------------------------------------------------------
# Main Window
# ---------------------------------------------------------------------------

class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("패널 측정 프로그램 v0.1")
        self.resize(1280, 1500)  # 가로 x 세로
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
        self.statusBar().showMessage("준비 — Mock 장비 연결 후 시작하세요")

    def closeEvent(self, event: Any) -> None:
        self._engine.disconnect_all()
        super().closeEvent(event)
