from interface import ConverterUI
import sys
from PySide6.QtWidgets import QApplication, QSplashScreen
from PySide6.QtGui import QPixmap
from PySide6.QtCore import Qt
from interface import ConverterUI, resource_path

if __name__ == "__main__":
    app = QApplication(sys.argv)

    splash_pixmap = QPixmap(resource_path("images/logo.png"))
    splash_pixmap = splash_pixmap.scaled(
        500,
        500,
        Qt.KeepAspectRatio,
        Qt.SmoothTransformation
    )

    splash = QSplashScreen(splash_pixmap)
    splash.setFixedSize(splash.size())
    splash.showMessage(
        "Loading...",
        Qt.AlignBottom | Qt.AlignCenter
    )
    splash.show()

    app.processEvents()

    window = ConverterUI()
    window.resize(350, 550)
    splash_screen = splash.screen()
    screen_geometry = splash_screen.availableGeometry()
    window.move(
        screen_geometry.center().x() - window.width() // 2,
        screen_geometry.center().y() - window.height() // 2,
    )
    window.show()
    splash.finish(window)

    sys.exit(app.exec())