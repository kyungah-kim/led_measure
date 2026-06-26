"""Desktop app entry point: python run_desktop.py"""
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))

os.environ.setdefault("QT_ENABLE_HIGHDPI_SCALING", "1")
os.environ.setdefault("QT_SCALE_FACTOR_ROUNDING_POLICY", "PassThrough")

# PyInstaller 번들 실행 시 matplotlib 데이터 경로 설정
# (빌드 시 mpl-data 를 번들에 포함했으므로 _MEIPASS 아래 경로를 지정)
if getattr(sys, "frozen", False):
    _mpl_data = os.path.join(sys._MEIPASS, "matplotlib", "mpl-data")  # type: ignore[attr-defined]
    if os.path.isdir(_mpl_data):
        os.environ.setdefault("MATPLOTLIBDATA", _mpl_data)

from PySide6.QtWidgets import QApplication, QSplashScreen
from PySide6.QtCore import Qt, QRect
from PySide6.QtGui import (QPixmap, QPainter, QColor, QFont,
                            QLinearGradient, QPen)


_APP_VERSION = "v 1.0"


def _make_splash() -> QSplashScreen:
    W, H = 540, 320
    pix = QPixmap(W, H)
    pix.fill(QColor("#0a0d12"))

    p = QPainter(pix)
    p.setRenderHint(QPainter.RenderHint.Antialiasing)

    # 상단 그라데이션 accent bar
    grad = QLinearGradient(0, 0, W, 0)
    grad.setColorAt(0.0, QColor("#0a5090"))
    grad.setColorAt(0.5, QColor("#1878d0"))
    grad.setColorAt(1.0, QColor("#0a5090"))
    p.fillRect(0, 0, W, 4, grad)

    # 미세 외곽 테두리
    p.setPen(QPen(QColor("#1a2535"), 1))
    p.drawRect(0, 0, W - 1, H - 1)

    # 중앙 배경 패널 (살짝 밝은 영역)
    p.fillRect(40, 30, W - 80, H - 60, QColor("#0e121a"))

    # 아이콘 — 간단한 측정기 심볼
    cx, cy = W // 2, 95
    p.setPen(QPen(QColor("#1878d0"), 2))
    p.setBrush(QColor("#0e1a2e"))
    p.drawEllipse(cx - 28, cy - 28, 56, 56)
    p.setPen(QPen(QColor("#5aadff"), 2))
    p.drawLine(cx, cy - 16, cx, cy - 6)      # 12시 눈금
    p.drawLine(cx + 14, cy - 8, cx + 10, cy - 2)   # 2시 눈금
    p.drawLine(cx - 14, cy - 8, cx - 10, cy - 2)   # 10시 눈금
    pen_needle = QPen(QColor("#ff6060"), 2)
    pen_needle.setCapStyle(Qt.PenCapStyle.RoundCap)
    p.setPen(pen_needle)
    p.drawLine(cx, cy, cx + 12, cy - 16)     # 바늘 (약 70% 위치)

    # 타이틀
    p.setPen(QColor("#e0eaf8"))
    f_title = QFont("Segoe UI", 22, QFont.Weight.Bold)
    f_title.setLetterSpacing(QFont.SpacingType.AbsoluteSpacing, 1.0)
    p.setFont(f_title)
    p.drawText(QRect(0, 128, W, 40), Qt.AlignmentFlag.AlignCenter,
               "LED Panel Analyzer")

    # 서브타이틀
    p.setPen(QColor("#3d6080"))
    f_sub = QFont("Segoe UI", 10)
    f_sub.setLetterSpacing(QFont.SpacingType.AbsoluteSpacing, 2.5)
    p.setFont(f_sub)
    p.drawText(QRect(0, 174, W, 24), Qt.AlignmentFlag.AlignCenter,
               "PROFESSIONAL  DISPLAY  MEASUREMENT  SUITE")

    # 구분선
    pen_line = QPen(QColor("#1a2535"), 1)
    p.setPen(pen_line)
    p.drawLine(80, 210, W - 80, 210)

    # 장비 지원 정보
    p.setPen(QColor("#2a4060"))
    f_info = QFont("Segoe UI", 9)
    p.setFont(f_info)
    p.drawText(QRect(0, 222, W, 20), Qt.AlignmentFlag.AlignCenter,
               "CA-310  /  CA-410  ·  VG-876  /  VG-879")

    # 버전
    p.setPen(QColor("#253545"))
    f_ver = QFont("Segoe UI", 8)
    p.setFont(f_ver)
    p.drawText(QRect(0, H - 26, W - 16, 18),
               Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter,
               _APP_VERSION)

    # 하단 accent bar
    p.fillRect(0, H - 3, W, 3, grad)

    p.end()

    splash = QSplashScreen(pix)
    splash.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint)
    splash.setWindowFlag(Qt.WindowType.FramelessWindowHint)
    return splash


def main() -> None:
    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
    )
    app = QApplication(sys.argv)
    app.setApplicationName("LED Panel Analyzer")

    splash = _make_splash()
    splash.show()
    app.processEvents()

    from desktop.main_window import MainWindow
    window = MainWindow()
    window.showMaximized()
    splash.finish(window)

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
