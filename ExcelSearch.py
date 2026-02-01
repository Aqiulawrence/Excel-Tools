import sys
import os
import json
import time
import traceback
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Tuple, Optional
import warnings

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

import pandas as pd
from openpyxl import load_workbook
from xlrd import open_workbook

# PyQt6 imports
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QTextEdit, QFileDialog,
    QProgressBar, QTableWidget, QTableWidgetItem, QHeaderView,
    QGroupBox, QCheckBox, QSpinBox, QMessageBox, QSplitter,
    QMenu, QStyleFactory
)
from PyQt6.QtCore import (
    Qt, QThread, pyqtSignal, QSettings, QTimer,
    QSize, QPoint, QEvent
)
from PyQt6.QtGui import (
    QFont, QPalette, QColor, QIcon, QAction,
    QTextCursor, QGuiApplication, QCursor
)

# 版本信息
VERSION = "1.0"
APP_NAME = "Excel搜索工具"
COMPANY_NAME = "Sam's Tools"


class SearchResult:
    """搜索结果数据类"""

    def __init__(self):
        self.file_path: str = ""
        self.sheet_name: str = ""
        self.cell_address: str = ""
        self.cell_value: str = ""
        self.part_number: str = ""
        self.description: str = ""
        self.rmb_price: str = ""
        self.usd_price: str = ""
        self.row_data: Dict[str, str] = {}
        self.timestamp: datetime = datetime.now()

    def to_dict(self) -> Dict:
        """转换为字典"""
        return {
            'file_path': self.file_path,
            'sheet_name': self.sheet_name,
            'cell_address': self.cell_address,
            'cell_value': self.cell_value,
            'part_number': self.part_number,
            'description': self.description,
            'rmb_price': self.rmb_price,
            'usd_price': self.usd_price,
            'timestamp': self.timestamp.isoformat()
        }


class ExcelSearchWorker(QThread):
    """Excel搜索工作线程"""

    # 信号定义
    progress_signal = pyqtSignal(int, int)  # 当前进度, 总数
    result_signal = pyqtSignal(SearchResult)  # 单个搜索结果
    error_signal = pyqtSignal(str, str)  # 错误文件路径, 错误信息
    finished_signal = pyqtSignal(int, int)  # 成功数, 失败数
    status_signal = pyqtSignal(str)  # 状态信息

    def __init__(self):
        super().__init__()
        self.search_term: str = ""
        self.search_paths: List[str] = []
        self.max_workers: int = 4
        self.is_running: bool = False

    def setup(self, search_term: str, search_paths: List[str], max_workers: int = 4):
        """设置搜索参数"""
        self.search_term = search_term.lower().strip()
        self.search_paths = search_paths
        self.max_workers = max_workers

    def run(self):
        """执行搜索"""
        self.is_running = True

        # 收集所有文件
        all_files = []
        for path in self.search_paths:
            if os.path.isfile(path):
                all_files.append(path)
            elif os.path.isdir(path):
                for root, _, files in os.walk(path):
                    for file in files:
                        if file.endswith(('.xlsx', '.xls')) and not file.startswith(('~$', '$')):
                            all_files.append(os.path.join(root, file))

        total_files = len(all_files)
        success_count = 0
        fail_count = 0

        self.status_signal.emit(f"发现 {total_files} 个文件，开始搜索...")

        for i, file_path in enumerate(all_files, 1):
            if not self.is_running:
                break

            self.progress_signal.emit(i, total_files)

            try:
                if file_path.endswith('.xls'):
                    self._search_xls_file(file_path)
                else:
                    self._search_xlsx_file(file_path)
                success_count += 1
            except Exception as e:
                fail_count += 1
                self.error_signal.emit(file_path, str(e))

        self.finished_signal.emit(success_count, fail_count)
        self.is_running = False

    def _search_xls_file(self, file_path: str):
        """搜索.xls文件"""
        try:
            wb = open_workbook(file_path)

            for sheet in wb.sheets():
                for row in range(sheet.nrows):
                    for col in range(sheet.ncols):
                        cell_value = str(sheet.cell_value(row, col)).lower().strip()
                        if self.search_term in cell_value:
                            result = SearchResult()
                            result.file_path = file_path
                            result.sheet_name = sheet.name
                            result.cell_address = f"{chr(65 + col)}{row + 1}"
                            result.cell_value = str(sheet.cell_value(row, col))

                            # 获取附近数据（简化版本）
                            self._extract_nearby_data_xls(sheet, row, col, result)

                            self.result_signal.emit(result)
        except Exception as e:
            raise Exception(f"搜索.xls文件失败: {str(e)}")

    def _search_xlsx_file(self, file_path: str):
        """搜索.xlsx文件"""
        try:
            workbook = load_workbook(file_path, read_only=True, data_only=True)

            for sheet_name in workbook.sheetnames:
                worksheet = workbook[sheet_name]

                # 预扫描列标题
                column_headers = self._scan_headers(worksheet)

                for row in worksheet.iter_rows():
                    for cell in row:
                        if cell.value is None:
                            continue

                        cell_value = str(cell.value).lower().strip()
                        if self.search_term in cell_value:
                            result = SearchResult()
                            result.file_path = file_path
                            result.sheet_name = sheet_name
                            result.cell_address = cell.coordinate
                            result.cell_value = str(cell.value)

                            # 提取整行数据
                            self._extract_row_data(worksheet, row, column_headers, result)

                            self.result_signal.emit(result)

            workbook.close()
        except Exception as e:
            raise Exception(f"搜索.xlsx文件失败: {str(e)}")

    def _scan_headers(self, worksheet) -> Dict[str, str]:
        """扫描列标题"""
        headers = {}
        header_row = None

        # 尝试查找标题行（前5行）
        for i, row in enumerate(worksheet.iter_rows(min_row=1, max_row=5, values_only=True), 1):
            for j, cell in enumerate(row):
                if cell and isinstance(cell, str):
                    cell_lower = cell.lower()
                    if '件号' in cell_lower or 'part' in cell_lower:
                        headers['part'] = chr(65 + j)
                    elif '品名' in cell_lower or 'description' in cell_lower:
                        headers['description'] = chr(65 + j)
                    elif '单价' in cell_lower or '价格' in cell_lower or 'rmb' in cell_lower:
                        headers['rmb'] = chr(65 + j)
                    elif 'usd' in cell_lower or '$' in cell:
                        headers['usd'] = chr(65 + j)

            if headers:
                header_row = i
                break

        return headers

    def _extract_row_data(self, worksheet, row, headers: Dict[str, str], result: SearchResult):
        """提取整行数据"""
        row_num = row[0].row

        # 提取件号
        if 'part' in headers:
            cell = worksheet[f"{headers['part']}{row_num}"]
            if cell.value:
                result.part_number = str(cell.value)
                result.row_data['part_number'] = result.part_number

        # 提取品名
        if 'description' in headers:
            cell = worksheet[f"{headers['description']}{row_num}"]
            if cell.value:
                result.description = str(cell.value)
                result.row_data['description'] = result.description

        # 提取RMB价格
        if 'rmb' in headers:
            cell = worksheet[f"{headers['rmb']}{row_num}"]
            if cell.value is not None:
                result.rmb_price = str(cell.value)
                result.row_data['rmb_price'] = result.rmb_price

        # 提取USD价格
        if 'usd' in headers:
            cell = worksheet[f"{headers['usd']}{row_num}"]
            if cell.value is not None:
                result.usd_price = str(cell.value)
                result.row_data['usd_price'] = result.usd_price

        # 提取其他列的数据
        for cell in row:
            col_letter = cell.column_letter
            if col_letter not in headers.values() and cell.value is not None:
                header_name = f"列{col_letter}"
                result.row_data[header_name] = str(cell.value)

    def _extract_nearby_data_xls(self, sheet, row: int, col: int, result: SearchResult):
        """提取.xls文件附近的数椐（简化版本）"""
        # 尝试获取件号（同一行的其他列）
        for c in range(max(0, col - 5), min(sheet.ncols, col + 5)):
            cell_value = str(sheet.cell_value(row, c))
            if any(keyword in cell_value.lower() for keyword in ['件号', 'part']):
                result.part_number = cell_value
                result.row_data['part_number'] = result.part_number

        # 尝试获取品名
        for c in range(max(0, col - 5), min(sheet.ncols, col + 5)):
            cell_value = str(sheet.cell_value(row, c))
            if any(keyword in cell_value.lower() for keyword in ['品名', 'description']):
                result.description = cell_value
                result.row_data['description'] = result.description

        # 获取价格
        for c in range(max(0, col - 5), min(sheet.ncols, col + 5)):
            cell_value = str(sheet.cell_value(row, c))
            if any(keyword in cell_value.lower() for keyword in ['单价', '价格', 'rmb']):
                result.rmb_price = cell_value
                result.row_data['rmb_price'] = result.rmb_price
            elif any(keyword in cell_value.lower() for keyword in ['usd', '$']):
                result.usd_price = cell_value
                result.row_data['usd_price'] = result.usd_price

    def stop(self):
        """停止搜索"""
        self.is_running = False


class ResultTableWidget(QTableWidget):
    """搜索结果表格控件"""

    def __init__(self):
        super().__init__()
        self.setup_ui()
        self.results: List[SearchResult] = []

    def setup_ui(self):
        """设置表格UI"""
        self.setColumnCount(7)
        self.setHorizontalHeaderLabels([
            "文件", "工作表", "单元格", "件号", "品名",
            "RMB价格", "USD价格"
        ])

        # 设置列宽
        self.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        self.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)
        self.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeMode.ResizeToContents)
        self.horizontalHeader().setSectionResizeMode(6, QHeaderView.ResizeMode.ResizeToContents)

        # 设置表格属性
        self.setAlternatingRowColors(True)
        self.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)

        # 添加上下文菜单
        self.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.customContextMenuRequested.connect(self.show_context_menu)

    def add_result(self, result: SearchResult):
        """添加搜索结果到表格"""
        self.results.append(result)
        row = self.rowCount()
        self.insertRow(row)

        # 文件（只显示文件名）
        file_item = QTableWidgetItem(os.path.basename(result.file_path))
        file_item.setToolTip(result.file_path)
        file_item.setData(Qt.ItemDataRole.UserRole, result)
        self.setItem(row, 0, file_item)

        # 工作表
        self.setItem(row, 1, QTableWidgetItem(result.sheet_name))

        # 单元格
        self.setItem(row, 2, QTableWidgetItem(result.cell_address))

        # 件号
        self.setItem(row, 3, QTableWidgetItem(result.part_number))

        # 品名
        desc_item = QTableWidgetItem(result.description)
        desc_item.setToolTip(result.description if len(result.description) > 50 else "")
        self.setItem(row, 4, desc_item)

        # RMB价格
        rmb_item = QTableWidgetItem(result.rmb_price)
        if result.rmb_price and any(char.isdigit() for char in result.rmb_price):
            try:
                rmb_value = float(''.join(filter(str.isdigit, result.rmb_price)))
                if rmb_value > 0:
                    rmb_item.setForeground(QColor("#c00000"))  # 红色
            except:
                pass
        self.setItem(row, 5, rmb_item)

        # USD价格
        usd_item = QTableWidgetItem(result.usd_price)
        if result.usd_price and any(char.isdigit() for char in result.usd_price):
            try:
                usd_value = float(''.join(filter(str.isdigit, result.usd_price)))
                if usd_value > 0:
                    usd_item.setForeground(QColor("#0070c0"))  # 蓝色
            except:
                pass
        self.setItem(row, 6, usd_item)

    def clear_results(self):
        """清空结果"""
        self.results.clear()
        self.setRowCount(0)

    def show_context_menu(self, pos: QPoint):
        """显示上下文菜单"""
        menu = QMenu(self)

        # 复制动作
        copy_action = QAction("复制选中内容", self)
        copy_action.triggered.connect(self.copy_selected)
        menu.addAction(copy_action)

        # 打开文件动作
        open_action = QAction("打开文件所在位置", self)
        open_action.triggered.connect(self.open_file_location)
        menu.addAction(open_action)

        # 导出结果动作
        export_action = QAction("导出所有结果为Excel", self)
        export_action.triggered.connect(self.export_results)
        menu.addAction(export_action)

        menu.addSeparator()

        # 清空结果动作
        clear_action = QAction("清空结果", self)
        clear_action.triggered.connect(self.clear_results)
        menu.addAction(clear_action)

        menu.exec(self.mapToGlobal(pos))

    def copy_selected(self):
        """复制选中内容"""
        selected = self.selectedItems()
        if not selected:
            return

        text = ""
        current_row = -1

        for item in selected:
            if item.row() != current_row:
                if text:
                    text += "\n"
                current_row = item.row()
                text += f"文件: {self.item(current_row, 0).text()}\n"
                text += f"工作表: {self.item(current_row, 1).text()}\n"
                text += f"单元格: {self.item(current_row, 2).text()}\n"
                text += f"件号: {self.item(current_row, 3).text()}\n"
                text += f"品名: {self.item(current_row, 4).text()}\n"
                text += f"RMB价格: {self.item(current_row, 5).text()}\n"
                text += f"USD价格: {self.item(current_row, 6).text()}\n"

        QApplication.clipboard().setText(text)

    def open_file_location(self):
        """打开文件所在位置"""
        selected_rows = set(item.row() for item in self.selectedItems())
        if not selected_rows:
            return

        for row in selected_rows:
            file_item = self.item(row, 0)
            if file_item:
                result = file_item.data(Qt.ItemDataRole.UserRole)
                if result and os.path.exists(result.file_path):
                    os.startfile(os.path.dirname(result.file_path))

    def export_results(self):
        """导出结果为Excel"""
        if not self.results:
            QMessageBox.warning(self, "提示", "没有可导出的结果")
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self, "导出结果",
            f"搜索结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            "Excel文件 (*.xlsx)"
        )

        if not file_path:
            return

        try:
            # 准备数据
            data = []
            for result in self.results:
                row_data = {
                    '文件路径': result.file_path,
                    '文件名': os.path.basename(result.file_path),
                    '工作表': result.sheet_name,
                    '单元格': result.cell_address,
                    '单元格内容': result.cell_value,
                    '件号': result.part_number,
                    '品名': result.description,
                    'RMB价格': result.rmb_price,
                    'USD价格': result.usd_price,
                    '搜索时间': result.timestamp.strftime('%Y-%m-%d %H:%M:%S')
                }
                # 添加额外数据
                row_data.update(result.row_data)
                data.append(row_data)

            # 创建DataFrame并保存
            df = pd.DataFrame(data)
            df.to_excel(file_path, index=False)

            QMessageBox.information(self, "成功", f"结果已导出到: {file_path}")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"导出失败: {str(e)}")


class LogTextEdit(QTextEdit):
    """日志文本编辑框"""

    def __init__(self):
        super().__init__()
        self.setup_ui()

    def setup_ui(self):
        """设置UI"""
        self.setReadOnly(True)
        self.setMaximumHeight(150)
        self.setFont(QFont("Consolas", 9))

        # 设置颜色
        palette = self.palette()
        palette.setColor(QPalette.ColorRole.Base, QColor("#f8f8f8"))
        self.setPalette(palette)

    def add_log(self, message: str, log_type: str = "info"):
        """添加日志"""
        timestamp = datetime.now().strftime("%H:%M:%S")

        if log_type == "error":
            formatted_message = f"[{timestamp}] <span style='color:#c00000'>{message}</span>"
        elif log_type == "warning":
            formatted_message = f"[{timestamp}] <span style='color:#ff9900'>{message}</span>"
        elif log_type == "success":
            formatted_message = f"[{timestamp}] <span style='color:#00b050'>{message}</span>"
        else:
            formatted_message = f"[{timestamp}] {message}"

        self.append(formatted_message)

        # 滚动到底部
        cursor = self.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.End)
        self.setTextCursor(cursor)

    def clear_logs(self):
        """清空日志"""
        self.clear()


class ExcelSearchTool(QMainWindow):
    """Excel搜索工具主窗口"""

    def __init__(self):
        super().__init__()

        # 初始化变量
        self.search_worker = None
        self.settings = QSettings(COMPANY_NAME, APP_NAME)
        self.search_history = []

        # 设置窗口
        self.setWindowTitle(f"{APP_NAME} v{VERSION}")
        self.setGeometry(100, 100, 1200, 800)

        # 设置图标（如果有）
        self._setup_icon()

        # 创建UI
        self._create_ui()

        # 加载设置
        self._load_settings()

        # 初始化拖放
        self._setup_drag_drop()

    def _setup_icon(self):
        """设置窗口图标"""
        try:
            # 尝试加载图标文件
            icon_path = "icon.ico"
            if os.path.exists(icon_path):
                self.setWindowIcon(QIcon(icon_path))
        except:
            pass

    def _create_ui(self):
        """创建主UI"""
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(10)

        # 创建顶部搜索区域
        self._create_search_area(main_layout)

        # 创建进度区域
        self._create_progress_area(main_layout)

        # 创建结果区域
        self._create_result_area(main_layout)

        # 创建日志区域
        self._create_log_area(main_layout)

        # 创建状态栏
        self.statusBar().showMessage("就绪")

    def _create_search_area(self, parent_layout):
        """创建搜索区域"""
        search_group = QGroupBox("搜索设置")
        search_layout = QVBoxLayout(search_group)

        # 搜索关键词
        keyword_layout = QHBoxLayout()
        keyword_layout.addWidget(QLabel("搜索关键词:"))

        self.keyword_input = QLineEdit()
        self.keyword_input.setPlaceholderText("请输入要搜索的内容（支持中文、英文、数字）")
        self.keyword_input.setClearButtonEnabled(True)
        self.keyword_input.returnPressed.connect(self.start_search)
        keyword_layout.addWidget(self.keyword_input)

        # 历史按钮
        history_btn = QPushButton("历史")
        history_btn.setMaximumWidth(60)
        history_btn.clicked.connect(self.show_search_history)
        keyword_layout.addWidget(history_btn)

        search_layout.addLayout(keyword_layout)

        # 文件选择区域
        file_layout = QHBoxLayout()
        file_layout.addWidget(QLabel("搜索目标:"))

        self.file_input = QLineEdit()
        self.file_input.setPlaceholderText("选择文件或文件夹，或直接拖拽到此处")
        self.file_input.setReadOnly(True)
        file_layout.addWidget(self.file_input)

        # 文件选择按钮
        file_btn = QPushButton("选择文件")
        file_btn.clicked.connect(lambda: self.select_target("file"))
        file_layout.addWidget(file_btn)

        folder_btn = QPushButton("选择文件夹")
        folder_btn.clicked.connect(lambda: self.select_target("folder"))
        file_layout.addWidget(folder_btn)

        search_layout.addLayout(file_layout)

        # 设置区域
        settings_layout = QHBoxLayout()

        # 最大线程数
        settings_layout.addWidget(QLabel("最大并发:"))
        self.thread_spin = QSpinBox()
        self.thread_spin.setRange(1, 16)
        self.thread_spin.setValue(4)
        settings_layout.addWidget(self.thread_spin)

        settings_layout.addStretch()

        # 搜索按钮
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
        settings_layout.addWidget(self.search_btn)

        # 停止按钮
        self.stop_btn = QPushButton("停止搜索")
        self.stop_btn.clicked.connect(self.stop_search)
        self.stop_btn.setEnabled(False)
        settings_layout.addWidget(self.stop_btn)

        search_layout.addLayout(settings_layout)

        parent_layout.addWidget(search_group)

    def _create_progress_area(self, parent_layout):
        """创建进度区域"""
        progress_layout = QHBoxLayout()

        self.progress_bar = QProgressBar()
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setFormat("准备搜索...")
        progress_layout.addWidget(self.progress_bar)

        self.count_label = QLabel("结果: 0")
        self.count_label.setMinimumWidth(100)
        progress_layout.addWidget(self.count_label)

        parent_layout.addLayout(progress_layout)

    def _create_result_area(self, parent_layout):
        """创建结果区域"""
        result_group = QGroupBox("搜索结果")
        result_layout = QVBoxLayout(result_group)

        # 创建表格
        self.result_table = ResultTableWidget()
        result_layout.addWidget(self.result_table)

        # 操作按钮
        button_layout = QHBoxLayout()
        button_layout.addStretch()

        clear_btn = QPushButton("清空结果")
        clear_btn.clicked.connect(self.clear_results)
        button_layout.addWidget(clear_btn)

        export_btn = QPushButton("导出Excel")
        export_btn.clicked.connect(self.result_table.export_results)
        button_layout.addWidget(export_btn)

        copy_btn = QPushButton("复制选中")
        copy_btn.clicked.connect(self.result_table.copy_selected)
        button_layout.addWidget(copy_btn)

        result_layout.addLayout(button_layout)

        parent_layout.addWidget(result_group)

    def _create_log_area(self, parent_layout):
        """创建日志区域"""
        log_group = QGroupBox("操作日志")
        log_layout = QVBoxLayout(log_group)

        self.log_text = LogTextEdit()
        log_layout.addWidget(self.log_text)

        # 日志操作按钮
        log_btn_layout = QHBoxLayout()
        log_btn_layout.addStretch()

        clear_log_btn = QPushButton("清空日志")
        clear_log_btn.clicked.connect(self.log_text.clear_logs)
        log_btn_layout.addWidget(clear_log_btn)

        save_log_btn = QPushButton("保存日志")
        save_log_btn.clicked.connect(self.save_logs)
        log_btn_layout.addWidget(save_log_btn)

        log_layout.addLayout(log_btn_layout)

        parent_layout.addWidget(log_group)

    def _setup_drag_drop(self):
        """设置拖放功能"""
        self.setAcceptDrops(True)
        self.file_input.setAcceptDrops(True)

    def _load_settings(self):
        """加载设置"""
        # 加载上次的搜索路径
        last_path = self.settings.value("last_search_path", "")
        if last_path and os.path.exists(last_path):
            self.file_input.setText(last_path)

        # 加载线程数
        thread_count = self.settings.value("thread_count", 4, type=int)
        self.thread_spin.setValue(thread_count)

        # 加载搜索历史
        history = self.settings.value("search_history", [])
        if history:
            self.search_history = history

    def _save_settings(self):
        """保存设置"""
        self.settings.setValue("thread_count", self.thread_spin.value())
        self.settings.setValue("search_history", self.search_history[-20:])  # 保存最近20条

    def select_target(self, target_type: str):
        """选择文件或文件夹"""
        if target_type == "file":
            file_path, _ = QFileDialog.getOpenFileName(
                self, "选择Excel文件", "",
                "Excel文件 (*.xlsx *.xls);;所有文件 (*.*)"
            )
            if file_path:
                self.file_input.setText(file_path)
                self.settings.setValue("last_search_path", file_path)
        else:
            folder_path = QFileDialog.getExistingDirectory(
                self, "选择文件夹", ""
            )
            if folder_path:
                self.file_input.setText(folder_path)
                self.settings.setValue("last_search_path", folder_path)

    def dragEnterEvent(self, event):
        """拖拽进入事件"""
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        """拖放事件"""
        urls = event.mimeData().urls()
        if urls:
            path = urls[0].toLocalFile()
            if os.path.exists(path):
                self.file_input.setText(path)
                self.settings.setValue("last_search_path", path)

    def start_search(self):
        """开始搜索"""
        # 验证输入
        keyword = self.keyword_input.text().strip()
        if not keyword:
            QMessageBox.warning(self, "提示", "请输入搜索关键词")
            self.keyword_input.setFocus()
            return

        target_path = self.file_input.text().strip()
        if not target_path:
            QMessageBox.warning(self, "提示", "请选择要搜索的文件或文件夹")
            return

        if not os.path.exists(target_path):
            QMessageBox.warning(self, "错误", "指定的路径不存在")
            return

        # 添加到搜索历史
        if keyword not in self.search_history:
            self.search_history.append(keyword)

        # 清空之前的结果
        self.result_table.clear_results()
        self.log_text.add_log(f"开始搜索: {keyword}，目标: {target_path}")

        # 更新UI状态
        self.search_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("准备搜索...")
        self.count_label.setText("结果: 0")

        # 创建并启动工作线程
        self.search_worker = ExcelSearchWorker()
        self.search_worker.setup(keyword, [target_path], self.thread_spin.value())

        # 连接信号
        self.search_worker.progress_signal.connect(self.update_progress)
        self.search_worker.result_signal.connect(self.add_search_result)
        self.search_worker.error_signal.connect(self.add_error_log)
        self.search_worker.finished_signal.connect(self.search_finished)
        self.search_worker.status_signal.connect(self.update_status)

        # 启动线程
        self.search_worker.start()

        self.statusBar().showMessage("搜索进行中...")

    def stop_search(self):
        """停止搜索"""
        if self.search_worker and self.search_worker.isRunning():
            self.search_worker.stop()
            self.log_text.add_log("搜索已停止", "warning")
            self.statusBar().showMessage("搜索已停止")

    def update_progress(self, current: int, total: int):
        """更新进度"""
        if total > 0:
            percent = int((current / total) * 100)
            self.progress_bar.setValue(percent)
            self.progress_bar.setFormat(f"正在搜索: {current}/{total} ({percent}%)")

    def add_search_result(self, result: SearchResult):
        """添加搜索结果"""
        self.result_table.add_result(result)
        self.count_label.setText(f"结果: {self.result_table.rowCount()}")

    def add_error_log(self, file_path: str, error_msg: str):
        """添加错误日志"""
        file_name = os.path.basename(file_path)
        self.log_text.add_log(f"处理失败 [{file_name}]: {error_msg}", "error")

    def search_finished(self, success: int, fail: int):
        """搜索完成"""
        # 更新UI状态
        self.search_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.progress_bar.setValue(100)
        self.progress_bar.setFormat(f"搜索完成 - 成功: {success}, 失败: {fail}")

        # 显示完成消息
        total_results = self.result_table.rowCount()
        message = f"搜索完成！共找到 {total_results} 个结果"
        if fail > 0:
            message += f"，其中 {fail} 个文件处理失败"

        self.log_text.add_log(message, "success")
        self.statusBar().showMessage(message)

        # 保存设置
        self._save_settings()

        # 播放完成音效（可选）
        if total_results > 0:
            QApplication.beep()

    def update_status(self, message: str):
        """更新状态"""
        self.statusBar().showMessage(message)
        self.log_text.add_log(message)

    def clear_results(self):
        """清空结果"""
        self.result_table.clear_results()
        self.count_label.setText("结果: 0")
        self.log_text.add_log("已清空搜索结果")

    def save_logs(self):
        """保存日志"""
        file_path, _ = QFileDialog.getSaveFileName(
            self, "保存日志",
            f"搜索日志_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
            "文本文件 (*.txt)"
        )

        if file_path:
            try:
                log_text = self.log_text.toPlainText()
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(log_text)
                self.log_text.add_log(f"日志已保存到: {file_path}", "success")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"保存日志失败: {str(e)}")

    def show_search_history(self):
        """显示搜索历史"""
        if not self.search_history:
            QMessageBox.information(self, "搜索历史", "暂无搜索历史")
            return

        menu = QMenu(self)

        for keyword in reversed(self.search_history[-10:]):  # 显示最近10条
            action = QAction(keyword, self)
            action.triggered.connect(lambda checked, k=keyword: self.set_keyword(k))
            menu.addAction(action)

        menu.addSeparator()
        clear_action = QAction("清空历史", self)
        clear_action.triggered.connect(self.clear_search_history)
        menu.addAction(clear_action)

        menu.exec(QCursor.pos())

    def set_keyword(self, keyword: str):
        """设置关键词"""
        self.keyword_input.setText(keyword)
        self.keyword_input.setFocus()

    def clear_search_history(self):
        """清空搜索历史"""
        self.search_history.clear()
        self._save_settings()
        QMessageBox.information(self, "提示", "搜索历史已清空")

    def closeEvent(self, event):
        """关闭事件"""
        # 停止正在运行的搜索
        if self.search_worker and self.search_worker.isRunning():
            self.search_worker.stop()
            self.search_worker.wait(2000)  # 等待2秒

        # 保存设置
        self._save_settings()

        event.accept()


def main():
    """主函数"""
    # 创建应用
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setOrganizationName(COMPANY_NAME)

    # 设置样式
    app.setStyle(QStyleFactory.create("Fusion"))

    # 设置调色板
    palette = QPalette()
    palette.setColor(QPalette.ColorRole.Window, QColor(240, 240, 240))
    palette.setColor(QPalette.ColorRole.WindowText, QColor(0, 0, 0))
    palette.setColor(QPalette.ColorRole.Base, QColor(255, 255, 255))
    palette.setColor(QPalette.ColorRole.AlternateBase, QColor(245, 245, 245))
    palette.setColor(QPalette.ColorRole.ToolTipBase, QColor(255, 255, 220))
    palette.setColor(QPalette.ColorRole.ToolTipText, QColor(0, 0, 0))
    palette.setColor(QPalette.ColorRole.Text, QColor(0, 0, 0))
    palette.setColor(QPalette.ColorRole.Button, QColor(240, 240, 240))
    palette.setColor(QPalette.ColorRole.ButtonText, QColor(0, 0, 0))
    palette.setColor(QPalette.ColorRole.BrightText, QColor(255, 255, 255))
    palette.setColor(QPalette.ColorRole.Link, QColor(42, 130, 218))
    palette.setColor(QPalette.ColorRole.Highlight, QColor(42, 130, 218))
    palette.setColor(QPalette.ColorRole.HighlightedText, QColor(255, 255, 255))
    app.setPalette(palette)

    # 创建并显示主窗口
    window = ExcelSearchTool()
    window.show()

    # 运行应用
    sys.exit(app.exec())


if __name__ == "__main__":
    main()