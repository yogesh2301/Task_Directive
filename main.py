import os
import sys

from PyQt6.QtWidgets import QApplication
from PyQt6.QtGui import QFont

from ui.main_window import OfflineApp
from ui.styles import get_light_theme


def configure_runtime_environment():
    if sys.platform.startswith("linux"):
        os.environ.setdefault("QT_QPA_PLATFORMTHEME", "fusion")


def main():
    configure_runtime_environment()
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setFont(QFont("Inter", 10))
    app.setStyleSheet(get_light_theme())

    window = OfflineApp()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
