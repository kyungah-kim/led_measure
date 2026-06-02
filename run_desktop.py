"""Desktop app entry point: python run_desktop.py"""
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))

# Windows High DPI 지원 — QApplication 생성 전에 설정해야 함
os.environ.setdefault("QT_ENABLE_HIGHDPI_SCALING", "1")
os.environ.setdefault("QT_SCALE_FACTOR_ROUNDING_POLICY", "PassThrough")

from PySide6.QtWidgets import QApplication
from PySide6.QtCore import Qt
from desktop.main_window import MainWindow


def main() -> None:
    app = QApplication(sys.argv)
    app.setApplicationName("패널 측정 프로그램")
    # High DPI 정책 (PySide6 6.x)
    app.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
    )
    window = MainWindow()
    window.showMaximized()   # 최대화로 시작 — resize(1280,1500) 대신
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
