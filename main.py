import sys
from pathlib import Path

from PySide6.QtCore import QStandardPaths, QThread, QUrl, Signal, Qt
from PySide6.QtGui import QColor, QDesktopServices, QFont
from PySide6.QtWidgets import (
    QApplication, QFileDialog, QHBoxLayout, QLabel, QLineEdit, QMainWindow,
    QMessageBox, QProgressBar, QPushButton, QStyle, QTableWidget,
    QTableWidgetItem, QVBoxLayout, QWidget,
)

from excel_merge import HEADERS, build_summary, merge_lis, save_summary


PROJECT = Path(__file__).resolve().parent


def default_directories():
    if getattr(sys, "frozen", False):
        documents = QStandardPaths.writableLocation(QStandardPaths.StandardLocation.DocumentsLocation)
        folder = Path(documents) if documents else Path.home()
        return folder, folder
    inputs = PROJECT / "execl"
    return (inputs if inputs.is_dir() else PROJECT), PROJECT


class Worker(QThread):
    completed = Signal(object)
    failed = Signal(str)

    def __init__(self, task, parent):
        super().__init__(parent)
        self.task = task

    def run(self):
        try:
            self.completed.emit(self.task())
        except Exception as error:
            self.failed.emit(str(error))


class SummaryWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.summary = None
        self.output_path = None
        self.source_paths = set()
        self.worker = None
        self.busy = False
        self.last_error = None
        self.input_directory, self.output_directory = default_directories()
        self.setWindowTitle("Excel 药敏汇总")
        self.resize(1180, 740)
        self.setMinimumSize(780, 520)

        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setContentsMargins(26, 22, 26, 20)
        layout.setSpacing(16)
        title = QLabel("药敏结果汇总")
        title.setObjectName("title")
        layout.addWidget(title)

        self.vitek_path = QLineEdit()
        self.lis_path = QLineEdit()
        self.vitek_button = QPushButton("1. 选择 vitek.xlsx")
        self.lis_button = QPushButton("2. 选择 lis.xlsx")
        for field, button in ((self.vitek_path, self.vitek_button), (self.lis_path, self.lis_button)):
            field.setReadOnly(True)
            field.setMinimumHeight(36)
            button.setMinimumSize(190, 36)
            button.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DialogOpenButton))
            row = QHBoxLayout()
            row.addWidget(button)
            row.addWidget(field, 1)
            layout.addLayout(row)
        self.vitek_button.clicked.connect(self.choose_vitek)
        self.lis_button.clicked.connect(self.choose_lis)
        self.lis_button.setEnabled(False)

        self.status = QLabel("等待 VITEK 文件")
        self.status.setWordWrap(True)
        self.status.setTextFormat(Qt.TextFormat.PlainText)
        layout.addWidget(self.status)
        self.progress = QProgressBar()
        self.progress.setRange(0, 0)
        self.progress.setFixedHeight(4)
        self.progress.setTextVisible(False)
        self.progress.hide()
        layout.addWidget(self.progress)

        self.preview_label = QLabel("汇总预览")
        layout.addWidget(self.preview_label)
        self.table = QTableWidget(0, len(HEADERS))
        self.table.setHorizontalHeaderLabels(HEADERS)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setAlternatingRowColors(True)
        self.table.setWordWrap(False)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(self.table, 1)

        footer = QHBoxLayout()
        self.output_label = QLabel("尚未保存")
        self.output_label.setWordWrap(True)
        self.output_label.setTextFormat(Qt.TextFormat.PlainText)
        footer.addWidget(self.output_label, 1)
        self.open_button = QPushButton("打开结果")
        self.open_button.setMinimumHeight(36)
        self.open_button.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_FileIcon))
        self.open_button.setEnabled(False)
        self.open_button.clicked.connect(self.open_output)
        footer.addWidget(self.open_button)
        layout.addLayout(footer)
        self.setCentralWidget(container)

        self.setStyleSheet("""
            QMainWindow, QWidget { background: #f5f7f6; color: #23322c; }
            QLabel#title { font-size: 24px; font-weight: 600; }
            QLineEdit { background: white; border: 1px solid #cbd4ce; padding: 0 10px; }
            QPushButton { background: #176b5b; color: white; border: none;
                          border-radius: 4px; padding: 6px 12px; }
            QPushButton:hover { background: #115446; }
            QPushButton:disabled { background: #dde4df; color: #68766d; }
            QTableWidget { background: white; alternate-background-color: #f0f5f2;
                           border: 1px solid #cbd4ce; gridline-color: #e2e8e4; }
            QHeaderView::section { background: #e3ebe6; padding: 8px;
                                   border: none; border-right: 1px solid #cbd4ce; }
            QProgressBar { border: none; background: #dde4df; }
            QProgressBar::chunk { background: #176b5b; }
        """)

    def set_busy(self, busy):
        self.busy = busy
        self.vitek_button.setEnabled(not busy)
        self.lis_button.setEnabled(not busy and self.summary is not None)
        self.open_button.setEnabled(not busy and self.output_path is not None)
        self.progress.setVisible(busy)

    def start_task(self, task, callback):
        self.last_error = None
        self.set_busy(True)
        self.worker = Worker(task, self)
        self.worker.completed.connect(callback)
        self.worker.failed.connect(self.show_error)
        self.worker.finished.connect(self.task_finished)
        self.worker.start()

    def task_finished(self):
        self.set_busy(False)
        self.worker.deleteLater()
        self.worker = None

    def show_error(self, message):
        self.last_error = message
        self.status.setText("处理失败，已保留上次成功结果")
        QMessageBox.critical(self, "无法完成汇总", message)

    def choose_vitek(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择 vitek.xlsx", str(self.input_directory), "Excel (*.xlsx)")
        if not path:
            return
        output, _ = QFileDialog.getSaveFileName(
            self, "保存汇总结果", str(self.output_directory / "summary_result.xlsx"), "Excel (*.xlsx)",
        )
        if output:
            self.load_vitek(path, output)

    def load_vitek(self, path, output):
        source = Path(path).resolve()
        protected = {source, source.parent / "lis.xlsx", source.parent / "summary.xlsx"}
        protected.update((PROJECT / "execl" / name).resolve() for name in ("vitek.xlsx", "lis.xlsx", "summary.xlsx"))
        self.status.setText("正在汇总 VITEK…")

        def task():
            book = build_summary(source)
            destination = save_summary(book, output, protected)
            return book, destination, protected, source

        self.start_task(task, self.vitek_completed)

    def vitek_completed(self, result):
        self.summary, self.output_path, self.source_paths, source = result
        self.input_directory = source.parent
        self.output_directory = self.output_path.parent
        self.vitek_path.setText(str(source))
        self.lis_path.clear()
        self.status.setText(f"VITEK 已完成：{self.summary.active.max_row - 1} 行。等待 LIS 文件。")
        self.update_preview(self.summary)

    def choose_lis(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择 lis.xlsx", str(self.input_directory), "Excel (*.xlsx)")
        if path:
            self.load_lis(path)

    def load_lis(self, path):
        if self.summary is None:
            return
        source = Path(path).resolve()
        protected = self.source_paths | {source}
        self.status.setText("正在匹配 LIS 并标记差异…")

        def task():
            book, stats = merge_lis(self.summary, source)
            save_summary(book, self.output_path, protected)
            return book, stats, source

        self.start_task(task, self.lis_completed)

    def lis_completed(self, result):
        book, stats, source = result
        self.lis_path.setText(str(source))
        self.status.setText(
            f"汇总完成：{stats.output_rows} 行，LIS 匹配 {stats.matched_rows} 行，"
            f"未匹配 {stats.missing_rows} 行；MIC 差异 {stats.mic_differences} 行，"
            f"解释差异 {stats.interpretation_differences} 行，日期差异 {stats.date_differences} 行。\n"
            f"重复匹配键 {stats.duplicate_keys} 组（全部保留）；"
            f"LIS 新菌株追加 {stats.appended_lis_rows} 行，"
            f"菌株已存在但抗菌剂未匹配 {stats.unmatched_lis_rows} 行，详见“匹配明细”。"
        )
        self.update_preview(book)

    def update_preview(self, book):
        sheet = book.active
        count = min(sheet.max_row - 1, 200)
        self.table.setRowCount(count)
        for row in range(count):
            for column in range(11):
                cell = sheet.cell(row + 2, column + 1)
                item = QTableWidgetItem("" if cell.value is None else str(cell.value))
                if cell.fill.patternType == "solid" and cell.fill.fgColor.type == "rgb":
                    item.setBackground(QColor("#" + cell.fill.fgColor.rgb[-6:]))
                if cell.font.color is not None and cell.font.color.type == "rgb":
                    item.setForeground(QColor("#" + cell.font.color.rgb[-6:]))
                self.table.setItem(row, column, item)
        self.table.resizeColumnsToContents()
        for column in range(11):
            self.table.setColumnWidth(column, min(280, max(140, self.table.columnWidth(column))))
        self.preview_label.setText(f"汇总预览 · {count} / {sheet.max_row - 1} 行")
        self.output_label.setText(str(self.output_path))

    def open_output(self):
        if self.output_path and not QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.output_path))):
            QMessageBox.warning(self, "无法打开文件", f"请手动打开：{self.output_path}")

    def closeEvent(self, event):
        if self.busy:
            QMessageBox.information(self, "正在处理", "请等待当前文件处理完成后再关闭窗口。")
            event.ignore()
        else:
            event.accept()


def main():
    app = QApplication(sys.argv)
    if sys.platform == "darwin":
        app.setFont(QFont("PingFang SC", 12))
    elif sys.platform == "win32":
        app.setFont(QFont("Microsoft YaHei", 10))
    window = SummaryWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())