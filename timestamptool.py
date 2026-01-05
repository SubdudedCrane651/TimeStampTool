import sys
import os
import time
import pywintypes
import win32file
import win32con

from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QLabel,
    QVBoxLayout, QFileDialog, QMessageBox
)


def get_file_times(path):
    """Return (created, accessed, modified) as Python timestamps."""
    handle = win32file.CreateFile(
        path,
        win32con.GENERIC_READ,
        win32con.FILE_SHARE_READ | win32con.FILE_SHARE_WRITE | win32con.FILE_SHARE_DELETE,
        None,
        win32con.OPEN_EXISTING,
        win32con.FILE_ATTRIBUTE_NORMAL,
        None
    )

    created, accessed, modified = win32file.GetFileTime(handle)
    handle.close()

    # Convert to Python timestamps
    return (
        created.timestamp(),
        accessed.timestamp(),
        modified.timestamp()
    )


def set_all_dates_to_created(path):
    """Set modified + accessed + creation time to the creation time."""
    handle = win32file.CreateFile(
        path,
        win32con.GENERIC_WRITE,
        win32con.FILE_SHARE_READ | win32con.FILE_SHARE_WRITE | win32con.FILE_SHARE_DELETE,
        None,
        win32con.OPEN_EXISTING,
        win32con.FILE_ATTRIBUTE_NORMAL,
        None
    )

    created, accessed, modified = win32file.GetFileTime(handle)

    # Set all timestamps to the creation timestamp
    win32file.SetFileTime(handle, created, created, created)
    handle.close()


class TimestampFixer(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("File Timestamp Normalizer")
        self.resize(400, 200)

        self.file_path = None

        self.label = QLabel("No file selected")
        self.btn_select = QPushButton("Select File")
        self.btn_fix = QPushButton("Set All Dates to Created Date")
        self.btn_fix.setEnabled(False)

        layout = QVBoxLayout()
        layout.addWidget(self.label)
        layout.addWidget(self.btn_select)
        layout.addWidget(self.btn_fix)
        self.setLayout(layout)

        self.btn_select.clicked.connect(self.select_file)
        self.btn_fix.clicked.connect(self.fix_timestamps)

    def select_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Choose a file")
        if not path:
            return

        self.file_path = path
        created, accessed, modified = get_file_times(path)

        self.label.setText(
            f"Selected:\n{path}\n\n"
            f"Created : {time.ctime(created)}\n"
            f"Accessed: {time.ctime(accessed)}\n"
            f"Modified: {time.ctime(modified)}"
        )

        self.btn_fix.setEnabled(True)

    def fix_timestamps(self):
        if not self.file_path:
            return

        try:
            set_all_dates_to_created(self.file_path)
            created, accessed, modified = get_file_times(self.file_path)

            QMessageBox.information(
                self,
                "Success",
                f"Timestamps updated!\n\n"
                f"Created : {time.ctime(created)}\n"
                f"Accessed: {time.ctime(accessed)}\n"
                f"Modified: {time.ctime(modified)}"
            )

        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = TimestampFixer()
    window.show()
    sys.exit(app.exec_())