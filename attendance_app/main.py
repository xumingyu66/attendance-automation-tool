"""Entry point for the attendance automation desktop application."""

import asyncio
import signal
import sys

import qasync
from PySide6.QtWidgets import QApplication

from ui.main_window import MainWindow


def main() -> int:
    app = QApplication(sys.argv)
    loop = qasync.QEventLoop(app)
    asyncio.set_event_loop(loop)
    signal.signal(signal.SIGINT, signal.SIG_DFL)

    with loop:
        window = MainWindow()
        window.show()
        loop.run_forever()

    return 0


if __name__ == "__main__":
    sys.exit(main())
