import sys
import traceback
from pathlib import Path

APP_ICON_PATH = Path(__file__).resolve().parent / "resources" / "app_icon.ico"


def run():
    from PySide6.QtGui import QIcon
    from PySide6.QtWidgets import QApplication
    from ui.main_window import MainWindow

    app = QApplication(sys.argv)
    if APP_ICON_PATH.exists():
        app.setWindowIcon(QIcon(str(APP_ICON_PATH)))
    win = MainWindow()
    win.show()
    return app.exec()


if __name__ == "__main__":
    try:
        raise SystemExit(run())
    except Exception:
        traceback.print_exc()
        raise SystemExit(1)
