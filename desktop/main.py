"""PySide6 desktop application entry point."""
from __future__ import annotations

import sys
import os

# Allow `from core.xxx import ...` when running from project root
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from PySide6.QtWidgets import QApplication

from .main_window import MainWindow


def main() -> None:
    app = QApplication(sys.argv)
    app.setApplicationName("Panel Measurement System")
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
