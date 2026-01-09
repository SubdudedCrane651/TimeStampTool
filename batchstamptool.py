import sys
import os
import time
import pywintypes
import win32file
import win32con

from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QLabel,
    QVBoxLayout, QFileDialog, QMessageBox, QListWidget, QCheckBox
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


class BatchTimestampFixer(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Batch File Timestamp Normalizer")
        self.resize(550, 500)

        self.folder_path = None

        self.label = QLabel("No folder selected")
        self.btn_select = QPushButton("Select Folder")
        self.btn_fix = QPushButton("Fix All Files")
        self.btn_fix.setEnabled(False)

        # NEW: Recursive option
        self.chk_recursive = QCheckBox("Include subfolders")

        self.file_list = QListWidget()

        layout = QVBoxLayout()
        layout.addWidget(self.label)
        layout.addWidget(self.btn_select)
        layout.addWidget(self.chk_recursive)
        layout.addWidget(self.file_list)
        layout.addWidget(self.btn_fix)
        self.setLayout(layout)

        self.btn_select.clicked.connect(self.select_folder)
        self.btn_fix.clicked.connect(self.fix_all_files)

    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Choose a folder")
        if not folder:
            return

        self.folder_path = folder
        self.label.setText(f"Selected folder:\n{folder}")

        self.file_list.clear()

        recursive = self.chk_recursive.isChecked()

        if recursive:
            # Walk all subdirectories
            for root, dirs, files in os.walk(folder):
                for filename in files:
                    full_path = os.path.join(root, filename)
                    self.file_list.addItem(full_path)
        else:
            # Only top-level files
            for filename in os.listdir(folder):
                full_path = os.path.join(folder, filename)
                if os.path.isfile(full_path):
                    self.file_list.addItem(full_path)

        if self.file_list.count() > 0:
            self.btn_fix.setEnabled(True)
        else:
            self.btn_fix.setEnabled(False)
            QMessageBox.information(self, "No Files", "This folder contains no files.")

    def fix_all_files(self):
        if not self.folder_path:
            return

        total = self.file_list.count()
        processed = 0
        skipped = 0

        for i in range(total):
            path = self.file_list.item(i).text()
            try:
                set_all_dates_to_created(path)
                processed += 1
            except Exception:
                skipped += 1
                continue  # silently skip

        QMessageBox.information(
            self,
            "Done",
            f"Processed {processed} of {total} files.\n"
            f"Skipped {skipped} files that could not be modified."
        )


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = BatchTimestampFixer()
    window.show()
    sys.exit(app.exec_())