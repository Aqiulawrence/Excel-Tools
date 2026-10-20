"""Excel workbook text search tool."""

import concurrent.futures
import json
import multiprocessing
import os
import sys
import threading
import warnings
import winreg
from bisect import bisect_left, bisect_right
from concurrent.futures.process import BrokenProcessPool
from dataclasses import dataclass
from decimal import Decimal, DecimalException, InvalidOperation
from itertools import chain, islice
from time import monotonic
from typing import Mapping, Sequence

from openpyxl import load_workbook
from openpyxl.styles.numbers import is_timedelta_format
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import MAC_EPOCH, WINDOWS_EPOCH, from_excel
from PyQt6.QtCore import QSettings, QThread, Qt, pyqtSignal
from PyQt6.QtGui import QColor
from PyQt6.QtWidgets import (
    QApplication, QDialog, QDialogButtonBox, QFileDialog, QFormLayout,
    QGroupBox, QHBoxLayout, QHeaderView, QLabel, QLineEdit, QMainWindow,
    QMenu, QMessageBox, QProgressBar, QPushButton, QSpinBox,
    QStyleFactory, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget,
)
from xlrd import (
    XL_CELL_BOOLEAN, XL_CELL_DATE, XL_CELL_ERROR, error_text_from_code, open_workbook,
)
from UsageReporter import UsageReporter

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

VERSION = "2.4"
APP_NAME = "Excel价格搜索工具 by Sam"
COMPANY_NAME = "Sam"
HEADER_SCAN_ROWS = 30
CANCEL_CHECK_INTERVAL_ROWS = 128
CANCEL_CHECK_INTERVAL_SECONDS = 0.5
DEFAULT_MAX_WORKERS = min(8, os.cpu_count() or 4)
PRICE_MATCH_ABSOLUTE_TOLERANCE = Decimal("1e-9")
PRICE_MATCH_RELATIVE_TOLERANCE = Decimal("1e-12")
EXCEL_EXTENSIONS = (".xls", ".xlsx", ".xlsm")
DEFAULT_KEYWORDS = {
    "description": ["品名", "description", "desc"],
    "rmb": ["RMB", "人民币", "￥", "¥"],
    "usd": ["USD", "$", "美元"],
}
RESULT_FIELDS = (
    ("description", "description"),
    ("rmb", "rmb_price"),
    ("usd", "usd_price"),
)


@dataclass
class SearchResult:
    file_path: str
    sheet_name: str
    cell_address: str
    cell_value: str
    description: str = ""
    rmb_price: str = ""
    usd_price: str = ""


def open_excel_at_cell(result):
    """Open or reuse a workbook in desktop Excel and reveal the matching cell."""
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        if getattr(sys, "frozen", False):
            message = (
                "打包程序的 Excel 自动定位依赖缺失或加载失败。"
                "请使用修复后的 pack.py 重新打包，并保留完整的 _internal 目录。"
            )
        else:
            message = "自动定位需要 Microsoft Excel 和 pywin32，请安装：pip install pywin32"
        raise RuntimeError(
            f"{message}\n\n详细信息：{exc}"
        ) from exc

    file_path = os.path.abspath(result.file_path)
    pythoncom.CoInitialize()
    try:
        try:
            excel = win32com.client.GetActiveObject("Excel.Application")
        except pythoncom.com_error:
            excel = win32com.client.Dispatch("Excel.Application")

        if excel.WindowState == -4140:  # xlMinimized
            excel.WindowState = -4143  # xlNormal
        excel.Visible = True
        # Show Excel in front before the potentially slow workbook load.
        try:
            bring_excel_to_front(int(excel.Hwnd))
        except Exception:
            # A newly started window may not accept focus yet. The caller
            # retries after navigation and reports a persistent focus failure.
            pass
        workbook = None
        normalized_path = os.path.normcase(file_path)
        for candidate in excel.Workbooks:
            if os.path.normcase(os.path.abspath(candidate.FullName)) == normalized_path:
                workbook = candidate
                break
        if workbook is None:
            workbook = excel.Workbooks.Open(file_path)

        worksheet = workbook.Worksheets(result.sheet_name)
        workbook.Activate()
        worksheet.Activate()
        target = worksheet.Range(result.cell_address)
        excel.Goto(target, True)
        center_excel_cell(excel.ActiveWindow, worksheet, target)
        # Keep Excel and the workbook open for the user; do not save or quit.
        return int(excel.Hwnd)
    finally:
        pythoncom.CoUninitialize()


def centered_scroll_index(target_index, minimum, viewport_size, span_size):
    """Find the scroll boundary closest to centering the target, in points."""
    half_cell = span_size(target_index) / 2
    half_viewport = viewport_size / 2
    low, high = minimum, target_index
    # Range dimensions account for unequal sizes and hidden rows/columns.
    while low < high:
        middle = (low + high) // 2
        if span_size(middle) - half_cell > half_viewport:
            low = middle + 1
        else:
            high = middle
    candidates = {low, max(minimum, low - 1)}
    return min(
        candidates,
        key=lambda index: (abs(span_size(index) - half_cell - half_viewport), index),
    )


def center_excel_cell(window, worksheet, target):
    """Keep the target visible with maximum left context and vertical centering."""
    frozen_rows = int(window.SplitRow) if window.FreezePanes else 0
    frozen_columns = int(window.SplitColumn) if window.FreezePanes else 0
    # The final pane is the scrollable bottom/right area for frozen headers.
    pane = window.Panes(window.Panes.Count)
    visible = pane.VisibleRange
    row, column = int(target.Row), int(target.Column)
    if column > frozen_columns:
        pane.ScrollColumn = leftmost_visible_scroll_index(
            column, frozen_columns + 1, float(visible.Width),
            lambda start: float(worksheet.Range(
                worksheet.Cells(row, start), worksheet.Cells(row, column)
            ).Width),
        )
    if row > frozen_rows:
        pane.ScrollRow = centered_scroll_index(
            row, frozen_rows + 1, float(visible.Height),
            lambda start: float(worksheet.Range(
                worksheet.Cells(start, column), worksheet.Cells(row, column)
            ).Height),
        )


def leftmost_visible_scroll_index(target_index, minimum, viewport_size, span_size):
    """Show as much left context as fits while keeping the target visible."""
    low, high = minimum, target_index
    while low < high:
        middle = (low + high) // 2
        if span_size(middle) > viewport_size:
            low = middle + 1
        else:
            high = middle
    return low


def bring_excel_to_front(hwnd):
    """Transfer foreground focus to Excel from the search UI."""
    import win32api
    import win32con
    import win32gui
    import win32process

    if win32gui.IsIconic(hwnd):
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)

    current_thread = win32api.GetCurrentThreadId()
    foreground = win32gui.GetForegroundWindow()
    foreground_thread = (
        win32process.GetWindowThreadProcessId(foreground)[0] if foreground else 0
    )
    attached = False
    try:
        # Excel may have a different UI thread from the window currently focused.
        if foreground_thread and foreground_thread != current_thread:
            win32process.AttachThreadInput(current_thread, foreground_thread, True)
            attached = True
        win32gui.BringWindowToTop(hwnd)
        win32gui.SetForegroundWindow(hwnd)
    finally:
        if attached:
            win32process.AttachThreadInput(current_thread, foreground_thread, False)


def clean_keywords(keywords: Mapping[str, Sequence[str]]) -> dict[str, list[str]]:
    """Normalize settings and ensure every supported field is present."""
    if not isinstance(keywords, Mapping):
        return {name: list(words) for name, words in DEFAULT_KEYWORDS.items()}
    cleaned = {}
    for name in DEFAULT_KEYWORDS:
        values = []
        words = keywords.get(name, [])
        if isinstance(words, str):
            words = words.split(",")
        elif not isinstance(words, Sequence):
            words = []
        for word in words:
            if word is None:
                continue
            value = str(word).strip()
            if value:
                values.append(value)
        cleaned[name] = values
    return cleaned


def find_keyword_columns(rows, keywords):
    """Locate the first matching header column for each output field."""
    columns = dict.fromkeys(DEFAULT_KEYWORDS)
    normalized = {
        name: [word.lower() for word in words]
        for name, words in keywords.items()
    }

    for row in rows:
        for index, value in enumerate(row):
            if value is None:
                continue

            text = str(value).lower()
            for name, words in normalized.items():
                if columns[name] is None and any(word in text for word in words):
                    columns[name] = index
    return columns


def numeric_value(value):
    """Read a finite numeric cell, including numbers stored as plain text."""
    if value is None or isinstance(value, bool):
        return None
    try:
        number = Decimal(str(value).strip())
    except InvalidOperation:
        return None
    return number if number.is_finite() else None


def find_unit_price_columns(row, columns):
    """Infer unit prices from three distinct cells: quantity * price = amount."""
    prices = {}
    for name in ("rmb", "usd"):
        column = columns[name]
        if column is not None and column < len(row):
            price = numeric_value(row[column])
            if price is not None and price > 0:
                prices[column] = price
    if not prices:
        return set()

    # Index values once, excluding the first column from the scan.
    value_columns = {}
    quantities = []
    for column in range(1, len(row)):
        number = numeric_value(row[column])
        if number is None or number <= 0:
            continue
        value_columns.setdefault(number, set()).add(column)
        if number == number.to_integral_value():
            quantities.append((column, number))

    sorted_values = sorted(value_columns)
    unit_prices = set()
    for price_column, price in prices.items():
        for quantity_column, quantity in quantities:
            if quantity_column == price_column:
                continue
            try:
                product = quantity * price
                # Excel's binary formula cache may store 35.70 as
                # 35.699999999999996. Allow numerical noise, not cent rounding.
                tolerance = max(
                    PRICE_MATCH_ABSOLUTE_TOLERANCE,
                    abs(product) * PRICE_MATCH_RELATIVE_TOLERANCE,
                )
                start = bisect_left(sorted_values, product - tolerance)
                end = bisect_right(sorted_values, product + tolerance)
            except DecimalException:
                # Price annotation must not prevent searching ordinary text
                # that happens to look like an out-of-range numeric value.
                continue
            if any(
                column != price_column and column != quantity_column
                for amount in sorted_values[start:end]
                for column in value_columns[amount]
            ):
                unit_prices.add(price_column)
                break
    return unit_prices


def build_result(
    file_path, sheet_name, row_number, column, value, row, columns,
    unit_price_columns=None,
):
    """Create one result and copy values from the detected columns."""
    if unit_price_columns is None:
        unit_price_columns = find_unit_price_columns(row, columns)
    details = {}
    for keyword, field in RESULT_FIELDS:
        index = columns[keyword]
        if index is not None and index < len(row) and row[index] is not None:
            details[field] = str(row[index])
            if keyword in ("rmb", "usd") and index in unit_price_columns:
                details[field] += "/件"
    address = f"{get_column_letter(column + 1)}{row_number}"
    return SearchResult(file_path, sheet_name, address, str(value), **details)


def search_rows(file_path, sheet_name, rows, search_term, keywords, cancel_event):
    """Search all rows in one worksheet and return matching cells."""
    iterator = iter(rows)
    if cancel_event.is_set():
        return None

    # Headers are normally near the top. A bounded scan avoids treating a
    # matching value deep in the sheet as a column header.
    headers = list(islice(iterator, HEADER_SCAN_ROWS))
    columns = find_keyword_columns(headers, keywords)
    matches = []
    next_cancel_check = monotonic() + CANCEL_CHECK_INTERVAL_SECONDS

    for row_number, row in enumerate(chain(headers, iterator), 1):
        # Read the clock in batches; only check the shared event twice a second.
        if (row_number - 1) % CANCEL_CHECK_INTERVAL_ROWS == 0:
            now = monotonic()
            if now >= next_cancel_check:
                if cancel_event.is_set():
                    return None
                next_cancel_check = now + CANCEL_CHECK_INTERVAL_SECONDS

        unit_price_columns = None
        for column, value in enumerate(row):
            if value is not None and search_term in str(value).lower():
                if unit_price_columns is None:
                    unit_price_columns = find_unit_price_columns(row, columns)
                matches.append(
                    build_result(
                        file_path,
                        sheet_name,
                        row_number,
                        column,
                        value,
                        row,
                        columns,
                        unit_price_columns,
                    )
                )

    return None if cancel_event.is_set() else matches


def xls_rows(sheet, workbook):
    """Normalize typed BIFF values to the same kinds of values as XLSX readers."""
    epoch = MAC_EPOCH if workbook.datemode else WINDOWS_EPOCH
    for row_number in range(sheet.nrows):
        values = sheet.row_values(row_number)
        for column, cell_type in enumerate(sheet.row_types(row_number)):
            if cell_type == XL_CELL_DATE:
                xf = workbook.xf_list[sheet.cell_xf_index(row_number, column)]
                number_format = workbook.format_map[xf.format_key].format_str
                values[column] = from_excel(
                    values[column], epoch=epoch,
                    timedelta=is_timedelta_format(number_format or ""),
                )
            elif cell_type == XL_CELL_BOOLEAN:
                values[column] = bool(values[column])
            elif cell_type == XL_CELL_ERROR:
                values[column] = error_text_from_code.get(values[column], str(values[column]))
        yield values


def search_file(file_path, search_term, keywords, cancel_event):
    """Return matches and an error message, or None for a cancelled workbook."""
    try:
        if cancel_event.is_set():
            return None, ""
        if file_path.lower().endswith(".xls"):
            # xlrd is used for the legacy binary format.
            workbook = open_workbook(file_path, formatting_info=True)

            worksheets = (
                (sheet.name, xls_rows(sheet, workbook))
                for sheet in workbook.sheets()
            )
            close = workbook.release_resources
        else:
            # openpyxl handles modern workbooks in read-only mode.
            workbook = load_workbook(file_path, read_only=True, data_only=True, keep_links=False)

            # Some producers write stale/incorrect worksheet dimensions. Trust the
            # actual XML rows and cells, otherwise read-only iteration truncates data.
            for sheet in workbook.worksheets:
                sheet.reset_dimensions()

            worksheets = (
                (sheet.title, sheet.iter_rows(values_only=True))
                for sheet in workbook.worksheets
            )
            close = workbook.close

        try:
            results = []
            for sheet_name, rows in worksheets:
                matches = search_rows(
                    file_path,
                    sheet_name,
                    rows,
                    search_term,
                    keywords,
                    cancel_event,
                )
                if matches is None:
                    return None, ""
                results.extend(matches)

            return results, ""
        finally:
            if close is not None:
                close()
    except Exception as exc:
        return [], f"{type(exc).__name__}: {exc}"


_process_cancel_event = None


def initialize_search_process(cancel_event):
    global _process_cancel_event
    _process_cancel_event = cancel_event


def search_file_in_process(file_path, search_term, keywords):
    return search_file(file_path, search_term, keywords, _process_cancel_event)


class SearchProcessPool:
    """Keep workers ready between searches; only one search uses the pool at a time."""

    def __init__(self):
        self.executor = None
        self.cancel_event = None
        self.max_workers = 0

    def acquire(self, max_workers):
        if self.executor is None or self.max_workers != max_workers:
            self.shutdown()
            context = multiprocessing.get_context("spawn")
            self.cancel_event = context.Event()
            self.executor = concurrent.futures.ProcessPoolExecutor(
                max_workers=max_workers, mp_context=context,
                initializer=initialize_search_process, initargs=(self.cancel_event,),
            )
            self.max_workers = max_workers
        self.cancel_event.clear()
        return self.executor, self.cancel_event

    def shutdown(self):
        if self.executor is not None:
            self.cancel_event.set()
            self.executor.shutdown(wait=True, cancel_futures=True)
        self.executor = None
        self.cancel_event = None
        self.max_workers = 0


class KeywordSettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("关键词设置")

        layout = QVBoxLayout(self)
        form = QFormLayout()
        self.inputs = {name: QLineEdit() for name in DEFAULT_KEYWORDS}

        fields = (
            ("description", "品名关键词"),
            ("rmb", "RMB关键词"),
            ("usd", "USD关键词"),
        )
        for name, label in fields:
            form.addRow(label, self.inputs[name])
        layout.addLayout(form)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok
            | QDialogButtonBox.StandardButton.Cancel
            | QDialogButtonBox.StandardButton.Reset
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        buttons.button(QDialogButtonBox.StandardButton.Reset).clicked.connect(self.reset_default)
        layout.addWidget(buttons)

    def reset_default(self):
        self.set_keywords(DEFAULT_KEYWORDS)

    def get_keywords(self):
        values = {
            name: input_.text().split(",")
            for name, input_ in self.inputs.items()
        }
        return clean_keywords(values)

    def set_keywords(self, keywords):
        for name, input_ in self.inputs.items():
            input_.setText(",".join(keywords.get(name, [])))


class ExcelSearchWorker(QThread):
    progress_signal = pyqtSignal(int, int, int)
    batch_result_signal = pyqtSignal(list)

    def __init__(self, process_pool=None):
        super().__init__()
        self.search_term = ""
        self.search_paths: list[str] = []
        self.max_workers = DEFAULT_MAX_WORKERS
        self.keywords = clean_keywords(DEFAULT_KEYWORDS)
        self._stop_requested = threading.Event()
        self.counts = (0, 0, 0)
        self.error = ""
        self.file_errors = []
        self.scan_errors = []
        self.duration_seconds = 0.0
        self.process_pool = process_pool if process_pool is not None else SearchProcessPool()
        self._owns_process_pool = process_pool is None

    def setup(self, search_term, search_paths, max_workers, keywords):
        self.search_term = search_term.lower().strip()
        self.search_paths = list(search_paths)
        self.max_workers = max_workers
        self.keywords = clean_keywords(keywords)

    def run(self):
        started_at = monotonic()
        completed = succeeded = failed = found = 0
        try:
            files = list(self.collect_files())
            total_files = len(files)
            if self.cancelled or not files:
                return

            worker_count = min(self.max_workers, total_files)
            executor, cancel_event = self.process_pool.acquire(worker_count)
            pending = iter(files)
            futures = {}
            try:
                while not self.cancelled:
                    while not self.cancelled and len(futures) < worker_count * 2:
                        file_path = next(pending, None)
                        if file_path is None:
                            break
                        try:
                            future = executor.submit(
                                search_file_in_process, file_path, self.search_term, self.keywords,
                            )
                        except BrokenProcessPool:
                            # A resident worker may die between searches. Retry only
                            # before any file has been submitted in this search.
                            if completed or futures:
                                raise
                            self.process_pool.shutdown()
                            executor, cancel_event = self.process_pool.acquire(worker_count)
                            future = executor.submit(
                                search_file_in_process, file_path, self.search_term, self.keywords,
                            )
                        futures[future] = file_path
                    if not futures or self.cancelled:
                        break
                    done, _ = concurrent.futures.wait(
                        futures, timeout=0.1,
                        return_when=concurrent.futures.FIRST_COMPLETED,
                    )
                    if self.cancelled:
                        break
                    ready_results = []
                    for future in done:
                        file_path = futures.pop(future)
                        completed += 1
                        try:
                            results, error = future.result()
                        except Exception as exc:
                            results, error = [], f"{type(exc).__name__}: {exc}"
                        if error:
                            failed += 1
                            self.file_errors.append(f"{file_path}: {error}")
                        elif results is not None:
                            succeeded += 1
                            found += len(results)
                            ready_results.extend(results)
                    if ready_results:
                        self.batch_result_signal.emit(ready_results)
                    self.progress_signal.emit(completed, total_files, found)
            finally:
                cancel_event.set()
                for future in futures:
                    future.cancel()
                # Drain the old search before a later search clears the event.
                if futures:
                    concurrent.futures.wait(futures)
        except Exception as exc:
            self.error = f"{type(exc).__name__}: {exc}"
            self.process_pool.shutdown()
        finally:
            if self._owns_process_pool:
                self.process_pool.shutdown()
            self.counts = (succeeded, failed, found)
            self.duration_seconds = round(monotonic() - started_at, 3)

    def collect_files(self):
        def report_scan_error(error):
            self.scan_errors.append(f"{error.filename or '未知目录'}: {error}")

        for path in self.search_paths:
            if self.cancelled:
                return
            if os.path.isdir(path):
                for root, _, filenames in os.walk(path, onerror=report_scan_error):
                    if self.cancelled:
                        return
                    for filename in filenames:
                        if self.cancelled:
                            return
                        is_excel = filename.lower().endswith(EXCEL_EXTENSIONS)
                        is_temporary = filename.startswith("~$")
                        if is_excel and not is_temporary:
                            yield os.path.join(root, filename)
            else:
                self.scan_errors.append(f"{path}: 目录不存在或无法访问")

    def stop(self):
        self._stop_requested.set()

    @property
    def cancelled(self):
        return self._stop_requested.is_set()


class ResultTableWidget(QTableWidget):
    HEADERS = ("文件", "单元格", "单元格内容", "品名", "RMB 价格", "USD 价格")
    PRICE_COLORS = {4: "#0070c0", 5: "#c00000"}

    def __init__(self):
        super().__init__()
        self.results: list[SearchResult] = []
        self.session_events = []

        self.setColumnCount(len(self.HEADERS))
        self.setHorizontalHeaderLabels(self.HEADERS)
        self.configure_columns()
        self.setAlternatingRowColors(True)
        self.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)

        self.doubleClicked.connect(self.open_excel_file)
        self.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.customContextMenuRequested.connect(self.show_context_menu)

    def configure_columns(self):
        header = self.horizontalHeader()

        for column in (0, 2, 3):
            header.setSectionResizeMode(column, QHeaderView.ResizeMode.Stretch)

        for column, width in ((1, 60), (4, 80), (5, 80)):
            header.setSectionResizeMode(column, QHeaderView.ResizeMode.Fixed)
            self.setColumnWidth(column, width)

    @staticmethod
    def text_item(value):
        item = QTableWidgetItem(value)
        item.setToolTip(value)
        return item

    def add_batch_results(self, results):
        start_row = self.rowCount()
        self.results.extend(results)
        self.setRowCount(start_row + len(results))

        for offset, result in enumerate(results):
            row = start_row + offset
            file_item = self.text_item(os.path.basename(result.file_path))
            file_item.setToolTip(result.file_path)
            self.setItem(row, 0, file_item)
            values = (result.cell_address, result.cell_value, result.description)
            for column, value in enumerate(values, 1):
                self.setItem(row, column, self.text_item(value))

            self.set_price_item(row, 4, result.rmb_price, self.PRICE_COLORS[4])
            self.set_price_item(row, 5, result.usd_price, self.PRICE_COLORS[5])

    def set_price_item(self, row, column, price, color):
        item = self.text_item(price)
        if price and any(character.isdigit() for character in price):
            item.setForeground(QColor(color))
        self.setItem(row, column, item)

    def clear_results(self):
        self.results.clear()
        self.setRowCount(0)

    def selected_result(self):
        row = self.currentRow()
        return self.results[row] if 0 <= row < len(self.results) else None

    def record_action(self, action, result, **details):
        self.session_events.append({"action": action, "file_path": result.file_path,
                                    **details})

    def open_excel_file(self):
        if result := self.selected_result():
            self.record_action("open_file", result)
            if not os.path.isfile(result.file_path):
                QMessageBox.warning(self, "打开失败", f"文件不存在：\n{result.file_path}")
                return
            try:
                excel_hwnd = open_excel_at_cell(result)
            except Exception as exc:
                try:
                    os.startfile(result.file_path)
                except OSError as open_error:
                    QMessageBox.warning(
                        self, "打开失败", f"{exc}\n\n打开文件也失败：{open_error}"
                    )
                    return
                QMessageBox.warning(
                    self, "未能自动定位",
                    f"已尝试用默认程序打开文件，但未能跳转到 "
                    f"{result.sheet_name}!{result.cell_address}。\n\n{exc}",
                )
            else:
                try:
                    bring_excel_to_front(excel_hwnd)
                except Exception as exc:
                    QMessageBox.warning(
                        self, "未能切换窗口",
                        f"Excel 已定位到 {result.sheet_name}!{result.cell_address}，"
                        f"但未能切到前台，请从任务栏切换到 Excel。\n\n{exc}",
                    )

    def show_context_menu(self, position):
        menu = QMenu(self)
        copy_action = menu.addAction("复制")
        copy_action.triggered.connect(self.copy_cell_content)
        menu.addSeparator()
        open_file_action = menu.addAction("打开文件并定位单元格")
        open_file_action.triggered.connect(self.open_excel_file)
        location_action = menu.addAction("打开文件所在目录")
        location_action.triggered.connect(self.open_file_location)
        menu.exec(self.mapToGlobal(position))

    def copy_cell_content(self):
        if item := self.currentItem():
            if result := self.selected_result():
                self.record_action("copy", result, text=item.text())
            QApplication.clipboard().setText(item.text())

    def open_file_location(self):
        for row in {item.row() for item in self.selectedItems()}:
            if row < len(self.results):
                self.record_action("open_folder", self.results[row])
                os.startfile(os.path.dirname(self.results[row].file_path))


class ExcelSearchTool(QMainWindow):
    def __init__(self):
        super().__init__()
        self.search_worker: ExcelSearchWorker | None = None
        self.process_pool = SearchProcessPool()
        self._close_pending = False
        self.settings = QSettings(COMPANY_NAME, APP_NAME)
        self.search_history: list[str] = []
        self.session_searches: list[dict[str, str | float]] = []
        self.custom_keywords = clean_keywords(DEFAULT_KEYWORDS)
        self.setWindowTitle(f"{APP_NAME} v{VERSION}")
        self.resize(700, 550)
        self.init_ui()
        self.load_settings()
        self.center_window()

    def init_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        layout.addWidget(self.create_search_group())
        self.progress_bar = QProgressBar()
        self.progress_bar.setFormat("就绪")
        layout.addWidget(self.progress_bar)

        group = QGroupBox("搜索结果")
        result_layout = QVBoxLayout(group)
        self.result_table = ResultTableWidget()
        result_layout.addWidget(self.result_table)
        layout.addWidget(group)

    def create_search_group(self):
        group = QGroupBox("搜索设置")
        layout = QVBoxLayout(group)

        layout.addLayout(self.create_keyword_row())
        layout.addLayout(self.create_folder_row())
        layout.addLayout(self.create_action_row())
        return group

    def create_keyword_row(self):
        row = QHBoxLayout()
        row.addWidget(QLabel("搜索关键词"))

        self.keyword_input = QLineEdit()
        self.keyword_input.returnPressed.connect(self.start_search)
        row.addWidget(self.keyword_input)

        self.keyword_buttons = []
        for label, callback in (
            ("历史", self.show_history),
            ("关键词设置", self.show_keyword_settings),
        ):
            button = QPushButton(label)
            button.clicked.connect(callback)
            row.addWidget(button)
            self.keyword_buttons.append(button)

        return row

    def create_folder_row(self):
        row = QHBoxLayout()
        row.addWidget(QLabel("搜索文件夹"))

        self.folder_input = QLineEdit()
        self.folder_input.setReadOnly(True)
        row.addWidget(self.folder_input)

        folder_button = QPushButton("选择文件夹")
        folder_button.clicked.connect(self.select_folder)
        row.addWidget(folder_button)

        return row

    def create_action_row(self):
        row = QHBoxLayout()
        row.addWidget(QLabel("并发文件数"))

        self.thread_spin = QSpinBox()
        self.thread_spin.setRange(1, 16)
        self.thread_spin.setValue(DEFAULT_MAX_WORKERS)
        row.addWidget(self.thread_spin)
        row.addStretch()

        self.search_btn = QPushButton("开始搜索")
        self.search_btn.clicked.connect(self.start_search)
        row.addWidget(self.search_btn)

        self.stop_btn = QPushButton("停止搜索")
        self.stop_btn.clicked.connect(self.stop_search)
        self.stop_btn.setEnabled(False)
        row.addWidget(self.stop_btn)

        return row

    def center_window(self):
        screen = QApplication.primaryScreen().availableGeometry()
        self.move(screen.center() - self.rect().center())

    def load_settings(self):
        if path := self.settings.value("last_search_path", ""):
            if os.path.isdir(path):
                self.folder_input.setText(path)
        thread_count = self.settings.value(
            "thread_count",
            DEFAULT_MAX_WORKERS,
            type=int,
        )
        self.thread_spin.setValue(thread_count)
        self.search_history = list(self.settings.value("search_history", []))

        try:
            saved_keywords = json.loads(
                self.settings.value("custom_keywords", "")
            )
            self.custom_keywords = clean_keywords(saved_keywords)
        except (TypeError, json.JSONDecodeError):
            pass

    def save_settings(self):
        self.settings.setValue("thread_count", self.thread_spin.value())
        self.settings.setValue("search_history", self.search_history[-20:])
        self.settings.setValue(
            "custom_keywords",
            json.dumps(self.custom_keywords, ensure_ascii=False),
        )

    def select_folder(self):
        if folder := QFileDialog.getExistingDirectory(self, "选择文件夹"):
            self.folder_input.setText(folder)
            self.settings.setValue("last_search_path", folder)

    def show_keyword_settings(self):
        dialog = KeywordSettingsDialog(self)
        dialog.set_keywords(self.custom_keywords)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.custom_keywords = dialog.get_keywords()
            self.save_settings()

    def start_search(self):
        if self.search_worker is not None or self._close_pending:
            return
        keyword = self.keyword_input.text().strip()
        folder = self.folder_input.text().strip()
        if not keyword:
            QMessageBox.warning(self, "提示", "请输入搜索关键词")
            return
        if not os.path.isdir(folder):
            QMessageBox.warning(self, "提示", "请选择有效的文件夹")
            return
        if keyword not in self.search_history:
            self.search_history.append(keyword)
        self.result_table.clear_results()
        self.search_btn.setEnabled(False)
        self.keyword_input.setEnabled(False)
        for button in self.keyword_buttons:
            button.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("准备搜索…")
        self.progress_bar.setToolTip("")
        search_record = None
        try:
            self.search_worker = ExcelSearchWorker(self.process_pool)
            self.search_worker.setup(
                keyword,
                [folder],
                self.thread_spin.value(),
                self.custom_keywords,
            )
            self.search_worker.progress_signal.connect(self.update_progress)
            self.search_worker.batch_result_signal.connect(self.result_table.add_batch_results)
            self.search_worker.finished.connect(self.search_finished)
            search_record = {
                "keyword": keyword, "folder": folder, "duration": 0.0,
                "overview": "准备搜索"}
            self.session_searches.append(search_record)
            self.search_worker.start()
        except Exception as exc:
            if search_record is not None:
                search_record["overview"] = "搜索失败"
            if self.search_worker is not None:
                self.search_worker.deleteLater()
                self.search_worker = None
            self.restore_search_controls()
            self.progress_bar.setFormat("搜索失败")
            QMessageBox.warning(self, "搜索失败", f"{type(exc).__name__}: {exc}")

    def stop_search(self):
        if self.search_worker is not None:
            self.search_worker.stop()
            self.stop_btn.setEnabled(False)
            self.progress_bar.setFormat("正在停止…")

    def update_progress(self, current, total, found):
        if self.search_worker is None or self.search_worker.cancelled:
            return
        percentage = int(current / total * 100) if total else 0
        self.progress_bar.setValue(percentage)
        self.progress_bar.setFormat(
            f"正在搜索：{current}/{total} ({percentage}%) — 已找到 {found} 条"
        )

    def restore_search_controls(self):
        self.search_btn.setEnabled(True)
        self.keyword_input.setEnabled(True)
        for button in self.keyword_buttons:
            button.setEnabled(True)
        self.stop_btn.setEnabled(False)

    def search_finished(self):
        worker = self.search_worker
        if worker is None:
            return
        self.session_searches[-1]["duration"] = worker.duration_seconds
        success, failed, found = worker.counts
        error = worker.error
        details = "\n".join(([error] if error else []) + worker.file_errors + worker.scan_errors)
        if error:
            status = "搜索失败"
        elif worker.cancelled:
            status = "已停止"
        elif failed:
            status = "搜索完成（部分文件读取失败）"
        else:
            status = "搜索完成"
        if not error and not worker.cancelled:
            self.progress_bar.setValue(100)
        self.progress_bar.setFormat(f"{status}：找到 {found} 条（成功：{success}，失败：{failed}）")
        self.progress_bar.setToolTip(details)
        self.session_searches[-1]["overview"] = self.progress_bar.format()
        self.search_worker = None
        worker.deleteLater()
        self.restore_search_controls()
        # if found and not error and not worker.cancelled and not self._close_pending:
        QApplication.beep()
        self.save_settings()
        if self._close_pending:
            self.close()
        elif error or worker.file_errors:
            message = QMessageBox(self)
            message.setWindowTitle(status)
            message.setIcon(QMessageBox.Icon.Warning)
            message.setText(error or (
                f"有 {failed} 个文件读取失败，{len(worker.scan_errors)} 个目录未能扫描。"
            ))
            message.setDetailedText(details)
            message.exec()

    def show_history(self):
        if not self.search_history:
            return
        menu = QMenu(self)
        for keyword in reversed(self.search_history[-10:]):
            action = menu.addAction(keyword)
            action.triggered.connect(lambda _, value=keyword: self.set_and_search(value))
        menu.addSeparator()
        clear_action = menu.addAction("清空历史")
        clear_action.triggered.connect(self.clear_history)
        menu.exec(self.cursor().pos())

    def set_and_search(self, keyword):
        if self.search_worker is not None or self._close_pending:
            return
        self.keyword_input.setText(keyword)
        self.start_search()

    def clear_history(self):
        self.search_history.clear()
        self.save_settings()

    def closeEvent(self, event):
        if self.search_worker is not None:
            self._close_pending = True
            self.stop_search()
            event.ignore()
            return
        self.save_settings()
        self.process_pool.shutdown()
        event.accept()


def register_updater_startup():
    """Keep the original updater auto-start behavior when the updater is present."""
    updater_path = os.path.join(os.path.dirname(sys.argv[0]), "Updater.exe")
    if not os.path.isfile(updater_path):
        return

    run_key = r"Software\Microsoft\Windows\CurrentVersion\Run"
    with winreg.OpenKey(
        winreg.HKEY_CURRENT_USER,
        run_key,
        0,
        winreg.KEY_SET_VALUE,
    ) as key:
        winreg.SetValueEx(key, "Updater", 0, winreg.REG_SZ, f'"{updater_path}"')


def main():
    multiprocessing.freeze_support()
    usage = UsageReporter("ExcelSearch")
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setOrganizationName(COMPANY_NAME)
    app.setStyle(QStyleFactory.create("Fusion"))
    window = ExcelSearchTool()
    window.show()
    register_updater_startup()
    exit_code = app.exec()
    usage.send(searches=window.session_searches,
               max_workers=window.thread_spin.value(),
               events=window.result_table.session_events)
    sys.exit(exit_code)


if __name__ == "__main__":
    main()
