import sys
import os
import win32file
import win32con

from PyQt5.QtCore import QObject, QThread, pyqtSignal
from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QLabel,
    QVBoxLayout, QFileDialog, QMessageBox, QListWidget,
    QCheckBox, QProgressBar
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
    win32file.SetFileTime(handle, created, created, created)
    handle.close()


# ---------------------------------------------------------
# WORKER THREAD
# ---------------------------------------------------------
class TimestampWorker(QObject):
    progress = pyqtSignal(int)       # emits % progress
    finished = pyqtSignal(int, int)  # emits processed, skipped

    def __init__(self, files):
        super().__init__()
        self.files = files

    def run(self):
        total = len(self.files)
        processed = 0
        skipped = 0

        for i, path in enumerate(self.files):
            try:
                set_all_dates_to_created(path)
                processed += 1
            except Exception:
                skipped += 1

            pct = int(((i + 1) / total) * 100)
            self.progress.emit(pct)

        self.finished.emit(processed, skipped)


# ---------------------------------------------------------
# MAIN WINDOW
# ---------------------------------------------------------
class BatchTimestampFixer(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Batch File Timestamp Normalizer (Multithreaded)")
        self.resize(550, 550)

        self.folder_path = None

        self.label = QLabel("No folder selected")
        self.btn_select = QPushButton("Select Folder")
        self.btn_fix = QPushButton("Fix All Files")
        self.btn_fix.setEnabled(False)

        self.chk_recursive = QCheckBox("Include subfolders")

        self.file_list = QListWidget()
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)

        layout = QVBoxLayout()
        layout.addWidget(self.label)
        layout.addWidget(self.btn_select)
        layout.addWidget(self.chk_recursive)
        layout.addWidget(self.file_list)
        layout.addWidget(self.progress_bar)
        layout.addWidget(self.btn_fix)
        self.setLayout(layout)

        self.btn_select.clicked.connect(self.select_folder)
        self.btn_fix.clicked.connect(self.start_thread)

    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Choose a folder")
        if not folder:
            return

        self.folder_path = folder
        self.label.setText(f"Selected folder:\n{folder}")

        self.file_list.clear()
        recursive = self.chk_recursive.isChecked()

        if recursive:
            for root, dirs, files in os.walk(folder):
                for filename in files:
                    full_path = os.path.join(root, filename)
                    self.file_list.addItem(full_path)
        else:
            for filename in os.listdir(folder):
                full_path = os.path.join(folder, filename)
                if os.path.isfile(full_path):
                    self.file_list.addItem(full_path)

        self.btn_fix.setEnabled(self.file_list.count() > 0)

    # ---------------------------------------------------------
    # START THREAD
    # ---------------------------------------------------------
    def start_thread(self):
        files = [self.file_list.item(i).text() for i in range(self.file_list.count())]

        self.thread = QThread()
        self.worker = TimestampWorker(files)
        self.worker.moveToThread(self.thread)

        self.thread.started.connect(self.worker.run)
        self.worker.progress.connect(self.update_progress)
        self.worker.finished.connect(self.thread_finished)

        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)

        self.btn_fix.setEnabled(False)
        self.progress_bar.setValue(0)

        self.thread.start()

    def update_progress(self, pct):
        self.progress_bar.setValue(pct)

    def thread_finished(self, processed, skipped):
        self.btn_fix.setEnabled(True)
        QMessageBox.information(
            self,
            "Done",
            f"Processed: {processed}\nSkipped: {skipped}"
        )


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = BatchTimestampFixer()
    window.show()
    sys.exit(app.exec_())