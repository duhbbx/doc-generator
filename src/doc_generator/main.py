"""Main entry point for Doc Generator."""

import sys

from PyQt6.QtWidgets import QApplication
from PyQt6.QtCore import Qt

from .gui.main_window import MainWindow


def main():
    """Main entry point."""
    # Enable high DPI scaling
    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
    )

    app = QApplication(sys.argv)
    app.setApplicationName("Doc Generator")
    app.setApplicationVersion("1.0.0")
    app.setOrganizationName("DocGenerator")

    # Set style
    app.setStyle("Fusion")

    window = MainWindow()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
