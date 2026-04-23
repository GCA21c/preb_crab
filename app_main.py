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
    print("[INFO] app_main.py start")
    try:
        code = run()
        print(f"[INFO] app exit code = {code}")
    except Exception:
        print("\n[FATAL ERROR]")
        traceback.print_exc()
    finally:
        input("\n엔터 누르기 전까지 안 닫힘 > ")
