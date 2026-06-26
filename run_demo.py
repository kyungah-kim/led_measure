"""Demo mode entry point — no real hardware required.

Launches the full LED Panel Analyzer UI with simulated (mock) devices so
the program can be demonstrated without a CA meter, VG generator, or LG TV.

Usage:
    python run_demo.py

What the demo does
------------------
- MockMeter, MockGenerator, MockLgSerial injected → all devices appear connected.
- Measurement delays are reduced (lum_swing 30 × 0.1 s, lum_loading instant).
- Basic Info pre-filled: Brand = "DEMO", Model = "LED-2024".
- Auto All brand set to "LG" with mock serial → PSM commands auto-confirm.
- Temporary folder used for auto-save.

All patches are local to this file.  The core package is NOT modified.
"""
from __future__ import annotations

import os
import sys
import tempfile
import time

sys.path.insert(0, os.path.dirname(__file__))

os.environ.setdefault("QT_ENABLE_HIGHDPI_SCALING", "1")
os.environ.setdefault("QT_SCALE_FACTOR_ROUNDING_POLICY", "PassThrough")

# ── 1. Reduce mock timing BEFORE importing anything from core ─────────────────
import core.equipment.mock as _mock_mod

_orig_mock_measure = _mock_mod.MockMeter.measure

def _fast_measure(self):  # type: ignore[override]
    time.sleep(0.01)      # 0.05 → 0.01 s per sample
    return _orig_mock_measure(self)

_mock_mod.MockMeter.measure = _fast_measure  # type: ignore[method-assign]

# lum_swing: interval 1.0 s → 0.1 s  (30 samples = 3 s total)
import core.sequences.lum_swing as _ls_mod
_ls_mod.LumSwingSequence.DEFAULT_INTERVAL_SEC = 0.1

# lum_loading: inter-measurement sleep 0.3 s → 0.0 s
import core.sequences.lum_loading as _ll_mod
_ll_mod._INTER_MEAS_SLEEP = 0.0

# ── 2. Mock LG TV serial port ────────────────────────────────────────────────
class MockLgSerial:
    """Simulates a pyserial Serial object connected to an LG TV debug shell."""

    def __init__(self) -> None:
        self.is_open = True
        self.in_waiting = 0   # never has data → LgTvReadWorker just sleeps

    def read(self, _n: int = 1) -> bytes:
        return b""

    def write(self, data: bytes) -> int:
        return len(data)

    def flush(self) -> None:
        pass

    def close(self) -> None:
        self.is_open = False


# ── 4. Patch MeasurementEngine.__init__ to inject mocks automatically ─────────
from core.engine import MeasurementEngine
from core.equipment.mock import MockMeter, MockGenerator

_orig_engine_init = MeasurementEngine.__init__

def _demo_engine_init(self, *args, **kwargs):  # type: ignore[override]
    _orig_engine_init(self, *args, **kwargs)
    self.meter     = MockMeter()
    self.generator = MockGenerator()
    self.brand      = "DEMO"
    self.model_name = "LED-2024"
    self.auto_save_dir = _DEMO_SAVE_DIR

MeasurementEngine.__init__ = _demo_engine_init  # type: ignore[method-assign]

# ── 5. Create a temp dir for auto-save (cleaned up at process exit) ───────────
_DEMO_SAVE_DIR = tempfile.mkdtemp(prefix="led_analyzer_demo_")

# ── 6. Build the Qt application and window ───────────────────────────────────
from PySide6.QtWidgets import QApplication, QSplashScreen
from PySide6.QtCore import Qt, QRect
from PySide6.QtGui import (QPixmap, QPainter, QColor, QFont,
                            QLinearGradient, QPen)


def _make_demo_splash() -> QSplashScreen:
    W, H = 540, 320
    pix = QPixmap(W, H)
    pix.fill(QColor("#0a0d12"))

    p = QPainter(pix)
    p.setRenderHint(QPainter.RenderHint.Antialiasing)

    grad = QLinearGradient(0, 0, W, 0)
    grad.setColorAt(0.0, QColor("#305090"))
    grad.setColorAt(0.5, QColor("#1878d0"))
    grad.setColorAt(1.0, QColor("#305090"))
    p.fillRect(0, 0, W, 4, grad)

    p.setPen(QPen(QColor("#1a2535"), 1))
    p.drawRect(0, 0, W - 1, H - 1)
    p.fillRect(40, 30, W - 80, H - 60, QColor("#0e121a"))

    cx, cy = W // 2, 90
    p.setPen(QPen(QColor("#1878d0"), 2))
    p.setBrush(QColor("#0e1a2e"))
    p.drawEllipse(cx - 28, cy - 28, 56, 56)
    p.setPen(QPen(QColor("#5aadff"), 2))
    p.drawLine(cx, cy - 16, cx, cy - 6)
    p.drawLine(cx + 14, cy - 8, cx + 10, cy - 2)
    p.drawLine(cx - 14, cy - 8, cx - 10, cy - 2)
    pen_needle = QPen(QColor("#1dd9a0"), 2)
    pen_needle.setCapStyle(Qt.PenCapStyle.RoundCap)
    p.setPen(pen_needle)
    p.drawLine(cx, cy, cx + 12, cy - 16)

    p.setPen(QColor("#e0eaf8"))
    f_title = QFont("Segoe UI", 22, QFont.Weight.Bold)
    f_title.setLetterSpacing(QFont.SpacingType.AbsoluteSpacing, 1.0)
    p.setFont(f_title)
    p.drawText(QRect(0, 124, W, 40), Qt.AlignmentFlag.AlignCenter, "LED Panel Analyzer")

    p.setPen(QColor("#1dd9a0"))
    f_demo = QFont("Segoe UI", 11, QFont.Weight.Bold)
    f_demo.setLetterSpacing(QFont.SpacingType.AbsoluteSpacing, 3.0)
    p.setFont(f_demo)
    p.drawText(QRect(0, 168, W, 26), Qt.AlignmentFlag.AlignCenter, "DEMO  MODE")

    p.setPen(QPen(QColor("#1a2535"), 1))
    p.drawLine(80, 208, W - 80, 208)

    p.setPen(QColor("#2a6040"))
    f_info = QFont("Segoe UI", 9)
    p.setFont(f_info)
    p.drawText(QRect(0, 220, W, 20), Qt.AlignmentFlag.AlignCenter,
               "Mock devices  ·  Simulated measurements  ·  No hardware required")

    p.setPen(QColor("#253545"))
    f_ver = QFont("Segoe UI", 8)
    p.setFont(f_ver)
    p.drawText(QRect(0, H - 26, W - 16, 18),
               Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter,
               "DEMO")

    p.fillRect(0, H - 3, W, 3, grad)
    p.end()

    splash = QSplashScreen(pix)
    splash.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint)
    splash.setWindowFlag(Qt.WindowType.FramelessWindowHint)
    return splash


def _apply_demo_ui(window) -> None:
    """Post-init patches: set status dots, spinboxes, and status bar."""
    from desktop.worker import LgTvReadWorker, wire_worker_cleanup

    cp  = window._conn_panel      # ConnectionPanel — meter / generator
    lgp = window._setting_panel  # SettingPanel — LG TV serial proxy
    engine = window._engine

    # ── CA meter / VG generator status dots ──────────────────────────────
    cp._meter_status.set_state("connected", tooltip="DEMO — MockMeter")
    cp._gen_status.set_state("connected", tooltip="DEMO — MockGenerator")

    # ── LG TV serial mock ─────────────────────────────────────────────────
    mock_ser = MockLgSerial()
    engine.lg_tv_serial = mock_ser
    engine.lg_serial_number = "DEMO-SN-001"
    engine.lg_sw_version    = "demo-os-1.0"
    engine.lg_sw_codename   = "DEMO"
    engine.lg_log_tx = lambda msg: lgp._lg_terminal.appendPlainText(f"[TX] {msg}")

    lgp._lg_status.set_state("connected", tooltip="DEMO — MockLgSerial")
    lgp._btn_lg.setEnabled(False)
    lgp._btn_lg_dis.setEnabled(True)
    lgp._lg_terminal.appendPlainText("[DEMO] LG TV serial connected (simulated)")
    lgp._lg_terminal.appendPlainText("[DEMO] Model: DEMO-LED2024  S/N: DEMO-SN-001")
    lgp._lg_terminal.appendPlainText("── Init complete, ready for luna commands ──")

    # Start the read worker so LgTvReadWorker doesn't crash
    # (in_waiting is always 0 → it just sleeps harmlessly)
    lgp._lg_read_worker = LgTvReadWorker(mock_ser)
    lgp._lg_read_worker.data_received.connect(lgp._on_lg_data)
    wire_worker_cleanup(lgp._lg_read_worker, lgp, '_lg_read_worker')
    lgp._lg_read_worker.start()

    # ── Basic Info ────────────────────────────────────────────────────────
    cp._brand_edit.setText("DEMO")
    cp._model_edit.setText("LED-2024")
    cp._save_dir_edit.setText(_DEMO_SAVE_DIR)
    cp._sync_info()

    # ── Auto All: LG brand + faster swing ────────────────────────────────
    ap = window._auto_panel
    ap._brand_combo.setCurrentText("LG")
    ap._swing_total_sec.setValue(30)
    ap._lg_wait_spin.setValue(0)       # PSM 전환 대기 0 s → 즉시 자동 확인
    ap._hdr_detect_spin.setValue(0)    # HDR 감지 대기 0 s

    # Status bar
    window.statusBar().showMessage(
        "DEMO MODE  —  Mock devices connected  ·  Measurements are simulated"
    )
    window._update_header_badges()


def main() -> None:
    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
    )
    app = QApplication(sys.argv)
    app.setApplicationName("LED Panel Analyzer [DEMO]")

    splash = _make_demo_splash()
    splash.show()
    app.processEvents()

    from desktop.main_window import MainWindow
    window = MainWindow()

    # Apply demo-specific UI patches after the window is fully constructed
    _apply_demo_ui(window)

    window.showMaximized()
    splash.finish(window)

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
