"""Excel workbook text search tool."""

from __future__ import annotations

import concurrent.futures
import contextlib
import ctypes
import hashlib
import json
import multiprocessing
import os
import re
import sqlite3
import struct
import time
import sys
import threading
import warnings
import winreg
from bisect import bisect_left, bisect_right
from concurrent.futures.process import BrokenProcessPool
from collections import deque
from ctypes import wintypes
from dataclasses import dataclass, field
from decimal import Decimal, DecimalException, InvalidOperation
from itertools import chain, islice
from pathlib import Path
from time import monotonic
from typing import Mapping, Sequence

from openpyxl import load_workbook
from openpyxl.styles.numbers import is_timedelta_format
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import MAC_EPOCH, WINDOWS_EPOCH, from_excel
from PyQt6.QtCore import QPropertyAnimation, QObject, QSettings, QThread, QTimer, Qt, pyqtSignal
from PyQt6.QtGui import QColor
from PyQt6.QtWidgets import (
    QApplication, QComboBox, QDialog, QDialogButtonBox, QFileDialog, QFormLayout,
    QGroupBox, QHBoxLayout, QHeaderView, QLabel, QLineEdit, QMainWindow,
    QMenu, QMessageBox, QProgressBar, QPushButton,
    QStyleFactory, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget,
)
from xlrd import (
    XL_CELL_BOOLEAN, XL_CELL_DATE, XL_CELL_ERROR, error_text_from_code, open_workbook,
)
from UsageReporter import UsageReporter

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

VERSION = "3.0"
# Developer switch: log progress text changes to the console immediately.
DEBUG_PROGRESS_TEXT = False
APP_NAME = "Excel价格搜索工具 by Sam"
COMPANY_NAME = "Sam"
HEADER_SCAN_ROWS = 30
CANCEL_CHECK_INTERVAL_ROWS = 128
CANCEL_CHECK_INTERVAL_SECONDS = 0.5
DEFAULT_MAX_WORKERS = 12
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


# Persistent index engine and native directory watcher.

BASE = Path(__file__).resolve().parent

DEFAULT_DATABASE = BASE / "ExcelSearch-index.sqlite3"

EXTENSIONS = (".xlsx", ".xlsm", ".xls")

SYSTEM_DIRS = {"$recycle.bin", "system volume information"}

INSERT_BATCH = 2000

class Cancelled(Exception):
    pass

class SourceChanged(Exception):
    pass

def check_cancel(cancel=None):
    if cancel is not None and cancel.is_set():
        raise Cancelled("操作已停止，已提交的索引保留。")

def path_key(path):
    return os.path.normcase(os.path.abspath(os.path.normpath(path)))

def database_path(database=None):
    path = Path(database or DEFAULT_DATABASE).resolve()
    if not path.is_relative_to(BASE) or path == BASE:
        raise ValueError("索引数据库必须位于程序文件夹内。")
    return path

def signature(path):
    stat = os.stat(path)
    return stat.st_size, stat.st_mtime_ns

@dataclass
class Source:
    path: str
    size: int | None = None
    mtime_ns: int | None = None
    error: str = ""

    @property
    def key(self):
        return path_key(self.path)

@dataclass
class Inventory:
    root: str
    sources: list[Source] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    inaccessible: list[str] = field(default_factory=list)

def inventory(folder, cancel=None, log=None):
    if not str(folder).strip():
        raise ValueError("请选择 Excel 文件夹。")
    root = os.path.abspath(folder)
    if not os.path.isdir(root):
        raise ValueError(f"目录不存在或无法访问：{root}")
    result = Inventory(root)
    last_progress = time.monotonic()
    directory_count = 0
    completed_weight = 0.0

    def report_scan(force=False):
        nonlocal last_progress
        if log is not None and (force or time.monotonic() - last_progress >= 0.1):
            # Each discovered subtree receives a share of its parent's budget.
            # This is an estimate of traversal coverage, not elapsed time.
            percentage = 100 if force and not pending else min(99, int(completed_weight * 100))
            log(f"目录扫描中：{directory_count} 个目录，已发现 {len(result.sources)} 个 Excel 文件；估算进度 {percentage}%")
            last_progress = time.monotonic()

    def walk_error(error):
        result.warnings.append(f"跳过目录：{error.filename}；{error}")
        if error.filename:
            result.inaccessible.append(path_key(error.filename))

    pending = [(root, 1.0)]
    while pending:
        check_cancel(cancel)
        current, weight = pending.pop()
        directories, files = [], []
        try:
            with os.scandir(current) as entries:
                for entry in entries:
                    check_cancel(cancel)
                    try:
                        is_directory = entry.is_dir()
                    except OSError:
                        # Match os.walk: an unclassifiable entry is a file.
                        is_directory = False
                    if is_directory:
                        if entry.name.casefold() not in SYSTEM_DIRS:
                            directories.append(entry)
                    elif (not entry.name.startswith("~$")
                          and entry.name.lower().endswith(EXTENSIONS)):
                        files.append(entry)
        except OSError as exc:
            walk_error(exc)
            completed_weight += weight
            report_scan()
            continue
        directory_count += 1
        # Keep sorted depth-first traversal without following directory symlinks.
        children = [entry.path for entry in sorted(
            directories, key=lambda entry: entry.name, reverse=True
        ) if not os.path.islink(entry.path)]
        share = weight / (len(children) + 1)
        completed_weight += share
        pending.extend((path, share) for path in children)
        for entry in files:
            check_cancel(cancel)
            source = Source(entry.path)
            try:
                # Windows directory enumeration already supplies this metadata.
                # Keep DirEntry instead of issuing another path-based os.stat.
                stat = entry.stat()
                source.size, source.mtime_ns = stat.st_size, stat.st_mtime_ns
            except OSError as exc:
                source.error = f"{type(exc).__name__}: {exc}"
            result.sources.append(source)
        report_scan()
    report_scan(force=True)
    result.sources.sort(key=lambda source: source.key)
    return result

@contextlib.contextmanager
def workbook_sheets(path):
    # XLS support is optional; missing xlrd cannot prevent XLSX indexing.
    if path.lower().endswith(".xls"):
        try:
            from xlrd import open_workbook
        except ImportError as exc:
            raise ValueError("读取 .xls 需要 xlrd，请安装依赖。") from exc
        book = open_workbook(path, formatting_info=True)
        try:
            yield ((sheet.name, xls_rows(sheet, book)) for sheet in book.sheets())
        finally:
            book.release_resources()
        return
    try:
        from openpyxl import load_workbook
    except ImportError as exc:
        raise ValueError("读取 .xlsx/.xlsm 需要 openpyxl，请安装依赖。") from exc
    book = load_workbook(path, read_only=True, data_only=True, keep_links=False)
    try:
        for sheet in book.worksheets:
            sheet.reset_dimensions()
        yield ((sheet.title, sheet.iter_rows(values_only=True)) for sheet in book.worksheets)
    finally:
        book.close()

SCHEMA = """
CREATE TABLE IF NOT EXISTS es_roots (
    root_key TEXT PRIMARY KEY, path TEXT NOT NULL, updated_ns INTEGER NOT NULL DEFAULT 0
);
CREATE TABLE IF NOT EXISTS es_files (
    id INTEGER PRIMARY KEY, root_key TEXT NOT NULL, path_key TEXT NOT NULL,
    path TEXT NOT NULL, size INTEGER, mtime_ns INTEGER, status TEXT NOT NULL,
    error TEXT NOT NULL DEFAULT '', sheets INTEGER NOT NULL DEFAULT 0,
    rows INTEGER NOT NULL DEFAULT 0, cells INTEGER NOT NULL DEFAULT 0,
    UNIQUE(root_key, path_key),
    FOREIGN KEY(root_key) REFERENCES es_roots(root_key)
);
CREATE TABLE IF NOT EXISTS es_cells (
    id INTEGER PRIMARY KEY, file_id INTEGER NOT NULL,
    sheet_order INTEGER NOT NULL, sheet_name TEXT NOT NULL,
    row_number INTEGER NOT NULL, column_number INTEGER NOT NULL,
    value TEXT NOT NULL, value_lower TEXT NOT NULL,
    FOREIGN KEY(file_id) REFERENCES es_files(id) ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS es_cells_file ON es_cells(file_id);
"""

FTS_SCHEMA = """
CREATE VIRTUAL TABLE IF NOT EXISTS es_fts USING fts5(
    value_lower, content='es_cells', content_rowid='id',
    tokenize='trigram case_sensitive 1'
);
CREATE TRIGGER IF NOT EXISTS es_cell_insert AFTER INSERT ON es_cells BEGIN
    INSERT INTO es_fts(rowid, value_lower) VALUES (new.id, new.value_lower);
END;
CREATE TRIGGER IF NOT EXISTS es_cell_delete AFTER DELETE ON es_cells BEGIN
    INSERT INTO es_fts(es_fts, rowid, value_lower) VALUES ('delete', old.id, old.value_lower);
END;
"""

@dataclass
class Matches:
    count: int = 0
    digest: object = field(default_factory=hashlib.sha256)
    preview: list = field(default_factory=list)
    limit: int | None = 500

    def add(self, hit):
        self.count += 1
        encoded = json.dumps(hit, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
        self.digest.update(len(encoded).to_bytes(8, "little"))
        self.digest.update(encoded)
        if self.limit is None or len(self.preview) < self.limit:
            self.preview.append(hit)

class SearchEngine:
    def __init__(self, database=None, log=print, cancel=None):
        self.database = database_path(database)
        self.log = log
        self.cancel = cancel
        self.connection = None
        self.fts = False

    def __enter__(self):
        check_cancel(self.cancel)
        self.database.parent.mkdir(parents=True, exist_ok=True)
        try:
            self.connection = sqlite3.connect(str(self.database), timeout=2)
            self.connection.row_factory = sqlite3.Row
            self.connection.execute("PRAGMA foreign_keys=ON")
            self.connection.execute("PRAGMA journal_mode=WAL")
            self.connection.execute("PRAGMA synchronous=NORMAL")
            self.connection.execute("PRAGMA cache_size=-16384")
            self.connection.executescript(SCHEMA)
            had_fts = self.connection.execute(
                "SELECT 1 FROM sqlite_master WHERE name='es_fts'"
            ).fetchone()
            try:
                self.connection.executescript(FTS_SCHEMA)
                if not had_fts:
                    self.connection.execute("INSERT INTO es_fts(es_fts) VALUES ('rebuild')")
                self.fts = True
            except sqlite3.OperationalError as exc:
                if "no such module" not in str(exc) and "tokenizer" not in str(exc):
                    raise
                self.log("当前 SQLite 不支持 trigram，使用 SQLite 子串查询。")
            self.connection.commit()
            return self
        except Exception:
            self.__exit__(None, None, None)
            raise

    def __exit__(self, *args):
        if self.connection is not None:
            self.connection.close()
            self.connection = None

    def _file_record(self, root_key, source, status, error=""):
        self.connection.execute(
            "INSERT INTO es_files(root_key,path_key,path,size,mtime_ns,status,error) "
            "VALUES (?,?,?,?,?,?,?) ON CONFLICT(root_key,path_key) DO UPDATE SET "
            "path=excluded.path,size=excluded.size,mtime_ns=excluded.mtime_ns,"
            "status=excluded.status,error=excluded.error,sheets=0,rows=0,cells=0",
            (root_key, source.key, source.path, source.size, source.mtime_ns, status, error),
        )
        return self.connection.execute(
            "SELECT id FROM es_files WHERE root_key=? AND path_key=?", (root_key, source.key)
        ).fetchone()[0]

    def _index_file(self, root_key, source):
        check_cancel(self.cancel)
        if source.error:
            raise OSError(source.error)
        with self.connection:
            file_id = self._file_record(root_key, source, "ok")
            self.connection.execute("DELETE FROM es_cells WHERE file_id=?", (file_id,))
            sheet_count = row_count = cell_count = 0
            batch = []
            insert_sql = ("INSERT INTO es_cells(file_id,sheet_order,sheet_name,row_number,"
                          "column_number,value,value_lower) VALUES (?,?,?,?,?,?,?)")
            with workbook_sheets(source.path) as sheets:
                for sheet_order, (name, rows) in enumerate(sheets):
                    sheet_count += 1
                    for row_number, row in enumerate(rows, 1):
                        check_cancel(self.cancel)
                        row_count += 1
                        for column, value in enumerate(row, 1):
                            if value is None or value == "":
                                continue
                            value = str(value)
                            batch.append((file_id, sheet_order, name, row_number, column, value, value.lower()))
                            cell_count += 1
                            if len(batch) >= INSERT_BATCH:
                                check_cancel(self.cancel)
                                self.connection.executemany(insert_sql, batch)
                                batch.clear()
                if batch:
                    self.connection.executemany(insert_sql, batch)
            if signature(source.path) != (source.size, source.mtime_ns):
                raise SourceChanged("文件在读取期间发生变化，将在下次更新时重试。")
            self.connection.execute("UPDATE es_files SET sheets=?,rows=?,cells=? WHERE id=?",
                                    (sheet_count, row_count, cell_count, file_id))

    def prepare(self, folder, retry_errors=False):
        """One inventory, per-file transactions, no global all-files-ready gate."""
        started = time.perf_counter()
        self.log("扫描文件夹并检查增量变化…")
        snapshot = inventory(folder, self.cancel, self.log)
        root_key = path_key(snapshot.root)
        for warning in snapshot.warnings[:30]:
            self.log(warning)
        if len(snapshot.warnings) > 30:
            self.log(f"另有 {len(snapshot.warnings) - 30} 个目录访问失败。")
        with self.connection:
            self.connection.execute("INSERT INTO es_roots(root_key,path) VALUES (?,?) "
                                    "ON CONFLICT(root_key) DO UPDATE SET path=excluded.path",
                                    (root_key, snapshot.root))
        previous = {row["path_key"]: row for row in self.connection.execute(
            "SELECT * FROM es_files WHERE root_key=?", (root_key,))}
        updated = reused = 0
        self.log(f"发现 {len(snapshot.sources)} 个 Excel 文件。")
        last_progress = time.monotonic()
        for position, source in enumerate(snapshot.sources, 1):
            check_cancel(self.cancel)
            old = previous.get(source.key)
            same = old is not None and (old["size"], old["mtime_ns"]) == (source.size, source.mtime_ns)
            if (same and not source.error and old["status"] != "inaccessible"
                    and (old["status"] == "ok" or not retry_errors)):
                reused += 1
            else:
                try:
                    self._index_file(root_key, source)
                    updated += 1
                except Cancelled:
                    raise
                except sqlite3.Error:
                    # Database failures must not be misclassified as bad Excel files.
                    raise
                except Exception as exc:
                    error = f"{type(exc).__name__}: {exc}"
                    with self.connection:
                        file_id = self._file_record(root_key, source, "error", error)
                        self.connection.execute("DELETE FROM es_cells WHERE file_id=?", (file_id,))
                    self.log(f"跳过文件：{source.path}；{error}")
            if position == 1 or time.monotonic() - last_progress >= 0.1 or position == len(snapshot.sources):
                self.log(f"处理进度：{position}/{len(snapshot.sources)}，已用 {time.perf_counter() - started:.1f} 秒")
                last_progress = time.monotonic()
        current_keys = {source.key for source in snapshot.sources}
        inaccessible = {
            key for key in previous
            if any(key == folder or key.startswith(os.path.join(folder, ""))
                   for folder in snapshot.inaccessible)
        }
        removed = set(previous) - current_keys - inaccessible
        with self.connection:
            for key in inaccessible:
                check_cancel(self.cancel)
                self.connection.execute(
                    "UPDATE es_files SET status='inaccessible',error=? WHERE root_key=? AND path_key=?",
                    ("所在目录暂时无法访问；保留缓存，恢复后重新核对。", root_key, key),
                )
            for key in removed:
                check_cancel(self.cancel)
                self.connection.execute("DELETE FROM es_files WHERE root_key=? AND path_key=?", (root_key, key))
            self.connection.execute("UPDATE es_roots SET updated_ns=? WHERE root_key=?", (time.time_ns(), root_key))
        counts = dict(self.connection.execute(
            "SELECT count(*) AS files,coalesce(sum(sheets),0) AS sheets,"
            "coalesce(sum(rows),0) AS rows,coalesce(sum(cells),0) AS cells "
            "FROM es_files WHERE root_key=? AND status='ok'", (root_key,)).fetchone())
        failed = self.connection.execute("SELECT count(*) FROM es_files WHERE root_key=? AND status!='ok'",
                                         (root_key,)).fetchone()[0]
        elapsed = time.perf_counter() - started
        self.log(f"索引可用：成功 {counts['files']} 个，跳过 {failed} 个；更新 {updated} 个，"
                 f"复用 {reused} 个，移除 {len(removed)} 个；准备耗时 {elapsed:.3f} 秒。")
        if not counts["files"]:
            self.log("当前范围没有可搜索的 Excel 数据；可检查上面的目录或文件错误。")
        self.log(f"索引：{self.database}")
        size = sum(path.stat().st_size for path in (self.database, Path(str(self.database) + "-wal")) if path.exists())
        return {"seconds": elapsed, "ready": True, "counts": counts, "skipped": failed,
                "updated": updated, "reused": reused, "removed": len(removed),
                "directory_errors": len(snapshot.warnings), "database_mib": size / 1024**2}

    def search(self, folder, term, limit=500):
        lowered = term.lower()
        use_fts = self.fts and len(lowered) >= 3 and "\x00" not in lowered
        select = ("SELECT f.path,c.sheet_name,c.row_number,c.column_number,c.value "
                  "FROM es_cells c JOIN es_files f ON f.id=c.file_id ")
        if use_fts:
            select += "JOIN es_fts ON es_fts.rowid=c.id "
            where = "es_fts MATCH ? AND "
            parameters = ['"' + lowered.replace('"', '""') + '"']
        else:
            where, parameters = "", []
        select += ("WHERE " + where + "f.root_key=? AND f.status='ok' AND instr(c.value_lower,?)>0 "
                   "ORDER BY f.path_key,c.sheet_order,c.row_number,c.column_number")
        parameters += [path_key(folder), lowered]
        matches = Matches(limit=limit)
        self.connection.set_progress_handler(lambda: int(self.cancel is not None and self.cancel.is_set()), 2000)
        try:
            for row in self.connection.execute(select, parameters):
                check_cancel(self.cancel)
                matches.add(tuple(row))
        except sqlite3.OperationalError:
            check_cancel(self.cancel)
            raise
        finally:
            self.connection.set_progress_handler(None, 0)
        return matches, "FTS5 trigram" if use_fts else "SQLite 子串扫描"

class DirectoryWatcher:
    def __init__(self, folder, changed, invalidated):
        self.folder = os.path.abspath(folder)
        self.changed, self.invalidated = changed, invalidated
        self.stopping = threading.Event()
        self.ready = threading.Event()
        self.condition = threading.Condition()
        self.requested = self.completed = 0
        self.error = None
        self.thread = None

    def start(self):
        if os.name != "nt":
            raise RuntimeError("实时监听需要 Windows。")
        self.thread = threading.Thread(target=self._run, name="ExcelDirectoryWatcher", daemon=False)
        self.thread.start()
        if not self.ready.wait(5):
            self.stop()
            raise RuntimeError("监听启动超时。")
        if self.error:
            self.stop()
            raise RuntimeError(f"监听启动失败：{self.error}")

    def flush(self, cancel=None):
        with self.condition:
            target = self.requested = self.requested + 1
            deadline = time.monotonic() + 5
            while self.completed < target:
                check_cancel(cancel)
                if self.error or self.stopping.is_set():
                    raise RuntimeError(f"监听已中断：{self.error or '已停止'}")
                if time.monotonic() > deadline:
                    raise RuntimeError("等待监听同步超时。")
                self.condition.wait(0.02)

    def stop(self):
        self.stopping.set()
        if self.thread and self.thread is not threading.current_thread():
            self.thread.join(5)
            if self.thread.is_alive():
                raise RuntimeError("目录监听未能退出。")

    def _run(self):
        api = None
        handle = event = None
        armed = False
        overlapped = None
        try:
            api = ctypes.WinDLL("kernel32", use_last_error=True)

            class OVERLAPPED(ctypes.Structure):
                _fields_ = [("Internal", ctypes.c_size_t), ("InternalHigh", ctypes.c_size_t),
                            ("Offset", wintypes.DWORD), ("OffsetHigh", wintypes.DWORD),
                            ("hEvent", wintypes.HANDLE)]

            api.CreateFileW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD,
                                       ctypes.c_void_p, wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE]
            api.CreateFileW.restype = wintypes.HANDLE
            api.CreateEventW.argtypes = [ctypes.c_void_p, wintypes.BOOL, wintypes.BOOL, wintypes.LPCWSTR]
            api.CreateEventW.restype = wintypes.HANDLE
            api.ReadDirectoryChangesW.argtypes = [wintypes.HANDLE, ctypes.c_void_p, wintypes.DWORD,
                wintypes.BOOL, wintypes.DWORD, ctypes.c_void_p, ctypes.POINTER(OVERLAPPED), ctypes.c_void_p]
            api.ReadDirectoryChangesW.restype = wintypes.BOOL
            api.GetOverlappedResult.argtypes = [wintypes.HANDLE, ctypes.POINTER(OVERLAPPED),
                                               ctypes.POINTER(wintypes.DWORD), wintypes.BOOL]
            api.GetOverlappedResult.restype = wintypes.BOOL
            api.WaitForSingleObject.argtypes = [wintypes.HANDLE, wintypes.DWORD]
            api.WaitForSingleObject.restype = wintypes.DWORD
            api.ResetEvent.argtypes = [wintypes.HANDLE]
            api.CancelIoEx.argtypes = [wintypes.HANDLE, ctypes.POINTER(OVERLAPPED)]
            api.CloseHandle.argtypes = [wintypes.HANDLE]
            handle = api.CreateFileW(self.folder, 1, 7, None, 3, 0x42000000, None)
            if handle == ctypes.c_void_p(-1).value:
                handle = None
                raise ctypes.WinError(ctypes.get_last_error())
            event = api.CreateEventW(None, True, False, None)
            if not event:
                raise ctypes.WinError(ctypes.get_last_error())
            buffer = ctypes.create_string_buffer(64 * 1024)
            overlapped = OVERLAPPED()
            overlapped.hEvent = event

            def arm():
                api.ResetEvent(event)
                if not api.ReadDirectoryChangesW(handle, buffer, len(buffer), True, 0x5B,
                                                 None, ctypes.byref(overlapped), None):
                    if ctypes.get_last_error() != 997:
                        raise ctypes.WinError(ctypes.get_last_error())

            arm()
            armed = True
            self.ready.set()
            while not self.stopping.is_set():
                result = api.WaitForSingleObject(event, 10)
                if result == 0:
                    count = wintypes.DWORD()
                    ok = api.GetOverlappedResult(handle, ctypes.byref(overlapped), ctypes.byref(count), False)
                    armed = False
                    if not ok and ctypes.get_last_error() != 1022:
                        raise ctypes.WinError(ctypes.get_last_error())
                    data = buffer.raw[:count.value] if ok else b""
                    arm()  # Re-arm before parsing/callbacks; OS buffers new changes.
                    armed = True
                    if not data:
                        self.invalidated("监听通知溢出，后台补扫。")
                    offset = 0
                    while offset + 12 <= len(data):
                        next_offset, action, length = struct.unpack_from("<III", data, offset)
                        if offset + 12 + length > len(data):
                            self.invalidated("监听通知不完整，后台补扫。")
                            break
                        name = data[offset + 12:offset + 12 + length].decode("utf-16-le")
                        self.changed(os.path.join(self.folder, name))
                        if not next_offset:
                            break
                        if next_offset < 12:
                            raise RuntimeError("无效目录通知。")
                        offset += next_offset
                elif result != 258:
                    raise ctypes.WinError(ctypes.get_last_error())
                if api.WaitForSingleObject(event, 0) == 258:
                    with self.condition:
                        self.completed = self.requested
                        self.condition.notify_all()
        except Exception as exc:
            self.error = str(exc)
            if not self.stopping.is_set():
                self.invalidated(f"监听中断：{exc}；暂停搜索，重连后补扫。")
        finally:
            self.ready.set()
            if api and handle:
                if armed:
                    api.CancelIoEx(handle, ctypes.byref(overlapped))
                    count = wintypes.DWORD()
                    api.GetOverlappedResult(handle, ctypes.byref(overlapped), ctypes.byref(count), True)
                api.CloseHandle(handle)
            if api and event:
                api.CloseHandle(event)
            with self.condition:
                self.condition.notify_all()

class StopToken(threading.Event):
    def __init__(self, shutdown):
        super().__init__()
        self.shutdown = shutdown

    def is_set(self):
        return self.shutdown.is_set() or super().is_set()

class LiveIndexService:
    def __init__(self, folder, database, emit, watcher_factory=DirectoryWatcher):
        self.folder = os.path.abspath(folder)
        self.database = database_path(database)
        self.emit, self.watcher_factory = emit, watcher_factory
        self.condition = threading.Condition()
        self.shutdown = threading.Event()
        self.operation_cancel = StopToken(self.shutdown)
        self.pending = set()
        self.rescan = True
        self.requests = deque()
        self.retries = {}
        self.watcher = self.engine = None
        self.last_report = self.last_error = None
        self.thread = threading.Thread(target=self._run, name="ExcelLiveIndex", daemon=False)

    def start(self):
        self.thread.start()

    def log(self, text):
        self.emit("log", text)

    def state(self, text):
        self.emit("live", text)

    def changed(self, path):
        parts = Path(os.path.relpath(path, self.folder)).parts
        if any(part.casefold() in SYSTEM_DIRS for part in parts) or os.path.basename(path).startswith("~$"):
            return
        if os.path.isfile(path) and not path.lower().endswith(EXTENSIONS):
            return
        with self.condition:
            self.pending.add(path_key(path))
            if len(self.pending) > 10000:
                self.pending.clear()
                self.rescan = True
            self.condition.notify_all()

    def invalidated(self, reason):
        self.log(reason)
        self.state("监听待恢复 / 正在补扫；当前搜索结果可能变化。")
        with self.condition:
            self.rescan = True
            self.condition.notify_all()

    def request(self, action, **options):
        with self.condition:
            self.requests.append((action, options))
            self.condition.notify_all()

    def _emit_done(self, options, outcome):
        # Optional request identity lets clients ignore a cancelled request's
        # late completion while another search is already queued.
        if "search_cancel" in options:
            outcome = {"search_cancel": options["search_cancel"], "outcome": outcome}
        self.emit("done", outcome)

    def cancel(self):
        self.operation_cancel.set()
        with self.condition:
            queued = bool(self.requests)
            self.requests.clear()
            self.condition.notify_all()
        if queued:
            self.emit("done", "操作已停止；下次搜索前会继续同步。")

    def stop(self):
        self.shutdown.set()
        self.cancel()

    def join(self, timeout=None):
        self.thread.join(timeout)

    def _connect_watcher(self):
        if not os.path.isdir(self.folder):
            raise ValueError(f"目录不存在或无法访问：{self.folder}")
        if self.watcher is not None and not self.watcher.error and not self.watcher.stopping.is_set():
            return
        if self.watcher:
            self.watcher.stop()
        self.watcher = self.watcher_factory(self.folder, self.changed, self.invalidated)
        self.watcher.start()
        self.rescan = True
        self.log("目录监听已开启（无需管理员权限）；开始后台核对索引。")

    def _counts(self, elapsed=0):
        key = path_key(self.folder)
        counts = dict(self.engine.connection.execute(
            "SELECT count(*) AS files,coalesce(sum(sheets),0) AS sheets,"
            "coalesce(sum(rows),0) AS rows,coalesce(sum(cells),0) AS cells "
            "FROM es_files WHERE root_key=? AND status='ok'", (key,)).fetchone())
        failed = self.engine.connection.execute(
            "SELECT count(*) FROM es_files WHERE root_key=? AND status!='ok'", (key,)).fetchone()[0]
        self.last_report = {"seconds": elapsed, "counts": counts, "skipped": failed, "ready": True}
        return self.last_report

    def _set_retry(self, path, attempt=0):
        path = path_key(path)
        if attempt < 3:
            self.retries[path] = (time.monotonic() + (1, 3, 10)[attempt], attempt + 1)
        else:
            self.retries.pop(path, None)

    def _update_file(self, path, attempt=0):
        check_cancel(self.operation_cancel)
        path = path_key(path)
        key = path_key(self.folder)
        connection = self.engine.connection
        source = Source(path)
        try:
            source.size, source.mtime_ns = signature(path)
        except FileNotFoundError:
            with connection:
                connection.execute("DELETE FROM es_files WHERE root_key=? AND path_key=?", (key, path_key(path)))
            self.retries.pop(path, None)
            return
        except OSError as exc:
            source.error = str(exc)
        try:
            self.engine._index_file(key, source)
            self.retries.pop(path, None)
        except (Cancelled, sqlite3.Error):
            raise
        except Exception as exc:
            with connection:
                file_id = self.engine._file_record(key, source, "error", f"{type(exc).__name__}: {exc}")
                connection.execute("DELETE FROM es_cells WHERE file_id=?", (file_id,))
            self.log(f"文件暂不可读：{path}；{exc}")
            self._set_retry(path, attempt)

    def _apply_paths(self, paths):
        if not paths:
            return
        previous = [row[0] for row in self.engine.connection.execute(
            "SELECT path FROM es_files WHERE root_key=?", (path_key(self.folder),))]
        files = set()
        inaccessible = set()
        for path in paths:
            check_cancel(self.operation_cancel)
            prefix = path_key(path) + os.sep
            descendants = [old for old in previous if path_key(old).startswith(prefix)]
            if os.path.isdir(path) or descendants:
                files.update(descendants)
                if os.path.isdir(path):
                    snapshot = inventory(path, self.operation_cancel)
                    files.update(source.path for source in snapshot.sources)
                    inaccessible.update(snapshot.inaccessible)
                    for warning in snapshot.warnings[:10]:
                        self.log(warning)
            elif path.lower().endswith(EXTENSIONS):
                files.add(path)
        for path in sorted(files):
            normalized = path_key(path)
            if any(normalized == folder or normalized.startswith(os.path.join(folder, ""))
                   for folder in inaccessible):
                check_cancel(self.operation_cancel)
                with self.engine.connection:
                    self.engine.connection.execute(
                        "UPDATE es_files SET status='inaccessible',error=? WHERE root_key=? AND path_key=?",
                        ("所在目录暂时无法访问；保留缓存，恢复后重新核对。", path_key(self.folder), normalized),
                    )
                self.retries.pop(normalized, None)
            else:
                self._update_file(normalized)
        if files:
            self.log(f"后台增量更新：处理 {len(files)} 个 Excel 路径。")

    def _synchronize(self, force=False):
        self._connect_watcher()
        if force:
            self.rescan = True
        self.state("正在后台更新索引，当前搜索结果为暂定结果，可能变化…")
        started = time.perf_counter()
        deadline = time.monotonic() + 30
        while True:
            check_cancel(self.operation_cancel)
            self.watcher.flush(self.operation_cancel)
            with self.condition:
                full_scan, paths = self.rescan, self.pending
                self.rescan, self.pending = False, set()
            try:
                if full_scan:
                    self.last_report = self.engine.prepare(self.folder, retry_errors=True)
                    inaccessible = self.engine.connection.execute(
                        "SELECT path FROM es_files WHERE root_key=? AND status='inaccessible'",
                        (path_key(self.folder),),
                    )
                    for row in inaccessible:
                        self.retries.pop(path_key(row[0]), None)
                    failed = self.engine.connection.execute(
                        "SELECT path FROM es_files WHERE root_key=? AND status='error'", (path_key(self.folder),))
                    for row in failed:
                        if path_key(row[0]) not in self.retries:
                            self._set_retry(row[0])
                else:
                    self._apply_paths(paths)
                self.watcher.flush(self.operation_cancel)
                with self.condition:
                    has_changes = self.rescan or bool(self.pending)
                if not has_changes:
                    break
                if time.monotonic() >= deadline:
                    raise RuntimeError("文件持续变化，暂未完成同步；请等待后台更新后再搜索。")
            except Exception:
                with self.condition:
                    self.pending.update(paths)
                    self.rescan = self.rescan or full_scan
                raise
        self.last_error = None
        report = self._counts(time.perf_counter() - started)
        suffix = f"；{report['skipped']} 个文件不可读，未纳入搜索" if report["skipped"] else ""
        self.state(f"持续监听中 · 已处理当前收到的变化 · {time.strftime('%H:%M:%S')}"
                   f" · {report['counts']['files']} 个成功文件{suffix}")
        return report

    def _execute(self, action, options):
        """Run a request in the service thread; supplied by the UI adapter."""
        raise NotImplementedError

    def _run(self):
        try:
            while not self.shutdown.is_set():
                action = None
                try:
                    self.operation_cancel.clear()
                    with self.condition:
                        if self.requests:
                            action, options = self.requests.popleft()
                    self._connect_watcher()
                    if self.engine is None:
                        self.engine = SearchEngine(self.database, self.log, self.operation_cancel)
                        self.engine.__enter__()
                    with self.condition:
                        needs_sync = self.rescan or bool(self.pending)
                    if action:
                        self._emit_done(options, self._execute(action, options))
                    elif needs_sync:
                        if self.shutdown.wait(0.15):
                            break
                        self._synchronize()
                    else:
                        due = [(path, attempt) for path, (when, attempt) in self.retries.items()
                               if when <= time.monotonic()]
                        for path, attempt in due:
                            self._update_file(path, attempt)
                        if due:
                            self._synchronize()
                        with self.condition:
                            if not self.requests and not self.pending and not self.rescan:
                                self.condition.wait(0.1)
                except Cancelled as exc:
                    self.state("更新已停止；下次搜索前会继续同步，当前不标记为最新。")
                    self.rescan = True
                    if action:
                        self._emit_done(options, str(exc))
                    with self.condition:
                        while not self.shutdown.is_set() and not self.requests:
                            self.condition.wait(0.1)
                except Exception as exc:
                    self.last_error = str(exc)
                    self.state(f"索引同步未完成：{exc}；将自动重试。")
                    self.log(f"后台更新失败：{type(exc).__name__}: {exc}")
                    self.rescan = True
                    if self.engine:
                        self.engine.__exit__(None, None, None)
                        self.engine = None
                    if action:
                        self._emit_done(options, f"搜索/更新未完成：{exc}")
                    deadline = time.monotonic() + 3
                    with self.condition:
                        while not self.shutdown.is_set() and not self.requests and time.monotonic() < deadline:
                            self.condition.wait(0.1)
        finally:
            if self.engine:
                self.engine.__exit__(None, None, None)
            if self.watcher:
                self.watcher.stop()
            self.emit("stopped", None)

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
    columns = None
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
                if columns is None:
                    columns = find_keyword_columns(headers, keywords)
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
            positions = {path: position for position, path in enumerate(files)}
            ready_files = {}
            next_position = 0
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
                        ready_files[positions[file_path]] = results or []
                    # Workers may finish out of order; publish in the same
                    # file/sheet/row/column order as the index query.
                    while next_position in ready_files:
                        ready_results.extend(ready_files.pop(next_position))
                        next_position += 1
                    if ready_results:
                        found += len(ready_results)
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
        seen = set()
        for path in self.search_paths:
            if self.cancelled:
                return
            try:
                snapshot = inventory(path, self._stop_requested)
            except Cancelled:
                return
            self.scan_errors.extend(snapshot.warnings)
            for source in snapshot.sources:
                if self.cancelled:
                    return
                if source.key not in seen:
                    seen.add(source.key)
                    yield source.path

    def stop(self):
        self._stop_requested.set()

    @property
    def cancelled(self):
        return self._stop_requested.is_set()


class ExcelLiveIndexService(LiveIndexService):
    """Adapt the local index service to progress and result signals for the UI."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._search_request = None
        self._active_search = None
        self._syncing = True
        self.search_thread = threading.Thread(
            target=self._run_searches, name="ExcelIndexReader", daemon=False)

    def start(self):
        super().start()
        self.search_thread.start()

    def stop(self):
        super().stop()
        with self.condition:
            self.condition.notify_all()

    def join(self, timeout=None):
        super().join(timeout)
        self.search_thread.join(timeout)

    def _run(self):
        try:
            super()._run()
        finally:
            self.shutdown.set()
            with self.condition:
                self.condition.notify_all()
            if self.search_thread.ident is not None:
                self.search_thread.join()

    def request(self, action, **options):
        if action != "search":
            return super().request(action, **options)
        options.setdefault("search_cancel", StopToken(self.shutdown))
        with self.condition:
            for previous in (self._search_request, self._active_search):
                if previous is not None and previous.get("search_cancel") is not options.get("search_cancel"):
                    previous["search_cancel"].set()
            self._search_request = options
            self.condition.notify_all()

    def results_provisional(self):
        with self.condition:
            return (self._syncing or self.rescan or bool(self.pending) or bool(self.retries)
                    or self.last_report is None or self.last_error is not None)

    def changed(self, path):
        super().changed(path)
        if self.results_provisional():
            self.emit("index_updating", True)

    def _run_searches(self):
        engine = None
        version = None
        initial = False
        options = None
        last_outcome = None
        try:
            while not self.shutdown.is_set():
                with self.condition:
                    if self._search_request is not None:
                        options = self._search_request
                        self._search_request = None
                        self._active_search = options
                        initial = True
                        version = None
                        last_outcome = None
                if options is None or options["search_cancel"].is_set():
                    if options is not None and initial:
                        self._emit_done(options, "本次搜索已停止。")
                        initial = False
                    with self.condition:
                        self.condition.wait(0.1)
                    continue
                try:
                    if engine is None:
                        # Do not run schema/FTS creation on the reader connection.
                        # The writer initializes new databases; old ones are usable immediately.
                        connection = sqlite3.connect(self.database.as_uri() + "?mode=ro", uri=True, timeout=0.1)
                        try:
                            connection.row_factory = sqlite3.Row
                            connection.execute("SELECT 1 FROM es_files LIMIT 1")
                            engine = SearchEngine(self.database, cancel=options["search_cancel"])
                            engine.connection = connection
                            engine.fts = bool(connection.execute(
                                "SELECT 1 FROM sqlite_master WHERE name='es_fts'").fetchone())
                        except Exception:
                            connection.close()
                            raise
                    current_version = engine.connection.execute("PRAGMA data_version").fetchone()[0]
                    provisional = self.results_provisional()
                    if (not initial and current_version == version and not provisional
                            and last_outcome is not None):
                        # Only the lifecycle state changed: the displayed snapshot
                        # is already current, so do not query/format every hit again.
                        final = dict(last_outcome, provisional=False)
                        self.emit("search_final", (options["search_cancel"], final))
                        options = None
                        continue
                    # Lifecycle changes alone do not change the result snapshot.
                    if initial or current_version != version:
                        engine.cancel = options["search_cancel"]
                        query = dict(options, _refresh=not initial, _results=[])
                        outcome = self._execute_search("search", query, engine.cancel, engine=engine)
                        outcome["provisional"] = (
                            self.results_provisional()
                            or engine.connection.execute("PRAGMA data_version").fetchone()[0] != current_version
                        )
                        if not engine.cancel.is_set():
                            if initial:
                                self._emit_done(options, outcome)
                            else:
                                self.emit("live_results", (options["search_cancel"], query["_results"], outcome))
                        last_outcome = outcome
                        version, initial = current_version, False
                        # Keep refreshing only until this search has a final snapshot.
                        # The writer and directory watcher continue independently.
                        if not outcome["provisional"]:
                            options = None
                except Cancelled:
                    if initial:
                        self._emit_done(options, "本次搜索已停止。")
                    initial = False
                except sqlite3.OperationalError as exc:
                    # A brand-new database may still be waiting for schema creation.
                    if engine is not None:
                        if initial:
                            self._emit_done(options, f"搜索失败：{exc}")
                        else:
                            self.emit("live_search_error", (options["search_cancel"], str(exc)))
                        options = None
                except Exception as exc:
                    if initial:
                        self._emit_done(options, f"搜索失败：{exc}")
                    else:
                        self.emit("live_search_error", (options["search_cancel"], str(exc)))
                    options = None
                with self.condition:
                    if self._search_request is None:
                        self.condition.wait(0.1)
        finally:
            if engine is not None:
                engine.__exit__(None, None, None)

    def log(self, text):
        # Expose user-facing progress from internal index lifecycle messages.
        if text.startswith("扫描文件夹并检查增量变化"):
            self._scan_started = monotonic()
            self.emit("progress", ("正在扫描文件夹…", 0, 0))
        elif match := re.match(r"目录扫描中：(\d+) 个目录，已发现 (\d+) 个 Excel 文件；估算进度 (\d+)%", text):
            percentage = int(match[3])
            self.emit("progress", (
                f"扫描进度估算 {percentage}%：已扫描 {match[1]} 个目录，发现 {match[2]} 个 Excel 文件",
                percentage, 100))
        elif match := re.match(r"发现 (\d+) 个 Excel 文件", text):
            self._index_started = monotonic()
            self._scan_seconds = self._index_started - self._scan_started
            total = int(match[1])
            # The count is already known; use the same phase label as the next update.
            self.emit("progress", (f"正在更新索引：0/{total} 个文件（0%）", 0, total))
        elif match := re.match(r"处理进度：(\d+)/(\d+)，已用 ([\d.]+) 秒", text):
            current, total = int(match[1]), int(match[2])
            percentage = int(current / total * 100) if total else 0
            self.emit("progress", (
                f"正在更新索引：{current}/{total} 个文件（{percentage}%），已用 {match[3]} 秒", current, total))
        elif match := re.match(r"后台增量更新：处理 (\d+) 个 Excel 路径", text):
            total = int(match[1])
            self.emit("progress", (f"已更新 {total} 个文件", total, total))
        elif text.startswith("索引可用："):
            self.emit("index_time", f"{self._scan_seconds:.2f}+{monotonic() - self._index_started:.2f}")

    def state(self, text):
        super().state(text)
        self.emit("index_updating", not text.startswith("持续监听中") or self.results_provisional())
        if text.startswith("持续监听中") and self.last_report is not None:
            report = self.last_report
            total = report["counts"]["files"]
            failed = report["skipped"]
            suffix = f"，{failed} 个文件暂时无法读取" if failed else ""
            self.emit("progress", (f"就绪：可搜索 {total} 个文件{suffix}", total, total))

    def _synchronize(self, force=False):
        self._syncing = True
        self.emit("index_updating", True)
        try:
            self.engine.connection.execute(
                "CREATE INDEX IF NOT EXISTS es_cells_location "
                "ON es_cells(file_id,sheet_name,row_number,column_number)"
            )
            self.engine.connection.commit()
            return super()._synchronize(force)
        finally:
            self._syncing = False
            self.emit("index_updating", self.results_provisional())

    def _execute(self, action, options):
        cancel = options.get("search_cancel", self.operation_cancel)
        original_cancel = self.engine.cancel
        try:
            check_cancel(cancel)
            return self._execute_search(action, options, cancel)
        except Cancelled:
            # Only shutdown cancels the index lifecycle. A cancelled search
            # returns normally so the inherited loop keeps applying changes.
            if self.shutdown.is_set():
                raise
            return "本次搜索已停止。"
        finally:
            self.engine.cancel = original_cancel

    def _execute_search(self, action, options, cancel, engine=None):
        if action == "build":
            self._synchronize(force=True)
        engine = engine or self.engine
        check_cancel(cancel)
        engine.cancel = cancel
        connection = engine.connection
        def progress(value):
            if not options.get("_refresh"):
                self.emit("search_progress", value)
        connection.execute("BEGIN")
        try:
            files = connection.execute(
                "SELECT count(*) FROM es_files WHERE root_key=? AND status='ok'",
                (path_key(self.folder),)).fetchone()[0]
            file_errors = [f"{row[0]}: {row[1]}" for row in connection.execute(
                "SELECT path,error FROM es_files WHERE root_key=? AND status!='ok' ORDER BY path_key",
                (path_key(self.folder),),
            )]
            found = 0
            backend = ""
            if action != "build":
                progress(("正在搜索…", 0, 0))
                matches, backend = engine.search(self.folder, options["term"], limit=None)
                progress((f"已找到 {matches.count} 条，正在整理结果…", 0, matches.count))
                previous_sheet = previous_row = None
                columns = row = unit_prices = None
                batch = []
                connection.set_progress_handler(lambda: int(cancel.is_set()), 2000)
                try:
                    for path, sheet, number, column, value in matches.preview:
                        check_cancel(cancel)
                        if previous_sheet != (path, sheet):
                            header_rows = {}
                            for cell in connection.execute(
                                "SELECT c.row_number,c.column_number,c.value FROM es_cells c "
                                "JOIN es_files f ON f.id=c.file_id "
                                "WHERE f.root_key=? AND f.path_key=? AND c.sheet_name=? AND c.row_number<=? "
                                "ORDER BY c.row_number,c.column_number",
                                (path_key(self.folder), path_key(path), sheet, HEADER_SCAN_ROWS),
                            ):
                                values = header_rows.setdefault(cell[0], [])
                                values.extend([None] * (cell[1] - len(values)))
                                values[cell[1] - 1] = cell[2]
                            columns = find_keyword_columns(header_rows.values(), options["keywords"])
                            previous_sheet = (path, sheet)
                        if previous_row != (path, sheet, number):
                            row = []
                            for cell in connection.execute(
                                "SELECT c.column_number,c.value FROM es_cells c "
                                "JOIN es_files f ON f.id=c.file_id "
                                "WHERE f.root_key=? AND f.path_key=? AND c.sheet_name=? AND c.row_number=? "
                                "ORDER BY c.column_number",
                                (path_key(self.folder), path_key(path), sheet, number),
                            ):
                                row.extend([None] * (cell[0] - len(row)))
                                row[cell[0] - 1] = cell[1]
                            unit_prices = find_unit_price_columns(row, columns)
                            previous_row = (path, sheet, number)
                        batch.append(build_result(path, sheet, number, column - 1, value, row, columns, unit_prices))
                        if len(batch) >= 500:
                            found += len(batch)
                            self.emit_search_batch(options, batch)
                            progress((
                                f"正在显示结果：{found}/{matches.count} 条", found, matches.count))
                            batch = []
                    check_cancel(cancel)
                    if batch:
                        found += len(batch)
                        self.emit_search_batch(options, batch)
                except Exception:
                    check_cancel(cancel)
                    raise
                finally:
                    connection.set_progress_handler(None, 0)
            with self.condition:
                changing = bool(self.pending) or self.rescan
            if changing:
                self.log("查询期间又收到文件变化；本次结果对应查询开始时的索引快照，后台将继续更新。")
            return {"counts": (files, len(file_errors), found),
                    "file_errors": file_errors, "backend": backend}
        finally:
            connection.rollback()

    def emit_search_batch(self, options, batch):
        if options.get("_refresh"):
            options["_results"].extend(batch)
            return
        if "search_cancel" in options:
            self.emit("search_batch", (options["search_cancel"], batch))
        else:
            self.emit("batch", batch)


class ExcelIndexSearchTask(QObject):
    """UI-side request; SQLite readers and the watcher run off the GUI thread."""

    progress_signal = pyqtSignal(int, int, int)
    batch_result_signal = pyqtSignal(list)
    finished = pyqtSignal()

    def __init__(self, enqueue, action="search"):
        super().__init__()
        self.enqueue = enqueue
        self.action = action
        self.service = None
        self._stop_requested = False
        self.search_cancel = StopToken(threading.Event())
        self._finished = False
        self.counts = (0, 0, 0)
        self.error = ""
        self.file_errors = []
        self.scan_errors = []
        self.duration_seconds = 0.0

    def setup(self, search_term, search_paths, max_workers, keywords):
        self.search_term = search_term.lower().strip()
        self.search_paths = list(search_paths)
        self.keywords = clean_keywords(keywords)

    def start(self):
        self.started_at = monotonic()
        self.enqueue(self)

    def dispatch(self, service):
        if self.cancelled or self._finished:
            return
        self.service = service
        self.search_cancel.shutdown = service.shutdown
        service.request(self.action, term=self.search_term, keywords=self.keywords,
                        search_cancel=self.search_cancel)

    def complete(self, outcome):
        if self._finished:
            return
        self._finished = True
        self.duration_seconds = round(monotonic() - self.started_at, 3)
        if isinstance(outcome, dict):
            self.counts = outcome["counts"]
            self.file_errors = outcome["file_errors"]
        elif not self.cancelled:
            self.error = str(outcome)
        self.finished.emit()

    def add_batch(self, results):
        if not self.cancelled and not self._finished:
            success, failed, found = self.counts
            self.counts = (success, failed, found + len(results))
            self.batch_result_signal.emit(results)

    def stop(self):
        if self._finished:
            return
        self._stop_requested = True
        self.search_cancel.set()
        self.complete(None)

    @property
    def cancelled(self):
        return self._stop_requested


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

        self.doubleClicked.connect(lambda: self.open_excel_file(action="double_click"))
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

    def replace_results(self, results):
        if self.results == results:
            return
        selected = self.selected_result()
        selected_key = ((selected.file_path, selected.sheet_name, selected.cell_address)
                        if selected else None)
        scroll = self.verticalScrollBar().value()
        column = max(0, self.currentColumn())
        self.setUpdatesEnabled(False)
        try:
            self.clear_results()
            self.add_batch_results(results)
            for row, result in enumerate(results):
                if (result.file_path, result.sheet_name, result.cell_address) == selected_key:
                    self.setCurrentCell(row, column)
                    break
            self.verticalScrollBar().setValue(scroll)
        finally:
            self.setUpdatesEnabled(True)

    def clear_results(self):
        self.results.clear()
        self.setRowCount(0)

    def selected_result(self):
        row = self.currentRow()
        return self.results[row] if 0 <= row < len(self.results) else None

    def record_action(self, action, result, **details):
        self.session_events.append({"action": action, "file_path": result.file_path,
                                    **details})

    def open_excel_file(self, *, action="open_file"):
        if result := self.selected_result():
            self.record_action(action, result)
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


class DebugProgressBar(QProgressBar):
    """Trace displayed progress text across all search modes and error paths."""

    def setFormat(self, text):
        changed = self.format() != text
        super().setFormat(text)
        if DEBUG_PROGRESS_TEXT and changed and text and sys.stdout is not None:
            print(f"[{time.strftime('%H:%M:%S')}] [progress] {text}", flush=True)


class ExcelSearchTool(QMainWindow):
    index_event_signal = pyqtSignal(int, str, object)

    def __init__(self):
        super().__init__()
        self.search_worker: ExcelSearchWorker | ExcelIndexSearchTask | None = None
        self.process_pool = SearchProcessPool()
        self._close_pending = False
        self.index_service = None
        self.index_generation = 0
        self.index_database = DEFAULT_DATABASE
        self._live_search_cancel = None
        self._index_is_updating = True
        self._index_search_summary = None
        self._index_results_final = False
        self._search_beeped = False
        self._index_query_progress = ("正在搜索…", 0, 0)
        self.index_state = "正在扫描文件夹…"
        self.index_progress = (self.index_state, 0, 0)
        self._keep_search_summary = False
        self._queued_index_task = None
        self._index_restart_timer = QTimer(self)
        self._index_restart_timer.setSingleShot(True)
        self._index_restart_timer.timeout.connect(self.restart_index_service)
        self._index_close_timer = QTimer(self)
        self._index_close_timer.setInterval(80)
        self._index_close_timer.timeout.connect(self.finish_pending_close)
        self.index_event_signal.connect(self.handle_index_event, Qt.ConnectionType.QueuedConnection)
        self.settings = QSettings(COMPANY_NAME, APP_NAME)
        self.search_history: list[str] = []
        self.session_searches: list[dict[str, str | float]] = []
        self.index_time = ""
        self.custom_keywords = clean_keywords(DEFAULT_KEYWORDS)
        self.setWindowTitle(f"{APP_NAME} v{VERSION}")
        self.resize(700, 550)
        self.init_ui()
        self.load_settings()
        self.search_mode_combo.currentIndexChanged.connect(self.search_mode_changed)
        self.search_mode_changed()
        self.center_window()
        self.folder_input.textChanged.connect(self.schedule_index_service)
        self.schedule_index_service(delay=0)

    def init_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        layout.addWidget(self.create_search_group())
        self.progress_bar = DebugProgressBar()
        self.progress_bar.setFormat("")
        layout.addWidget(self.progress_bar)
        self._progress_animation = QPropertyAnimation(self.progress_bar, b"value", self)
        self._progress_animation.setDuration(160)
        self._progress_target = None
        self._progress_refresh_timer = QTimer(self)
        self._progress_refresh_timer.setSingleShot(True)
        self._progress_refresh_timer.setInterval(150)
        self._progress_refresh_timer.timeout.connect(self.render_index_progress)

        group = QGroupBox("搜索结果")
        result_layout = QVBoxLayout(group)
        self.result_table = ResultTableWidget()
        result_layout.addWidget(self.result_table)
        layout.addWidget(group)

    def create_search_group(self):
        group = QGroupBox("搜索设置")
        layout = QVBoxLayout(group)
        self._search_label_width = max(
            QLabel(text).sizeHint().width()
            for text in ("搜索关键词", "搜索文件夹", "搜索方式")
        )

        layout.addLayout(self.create_keyword_row())
        layout.addLayout(self.create_folder_row())
        layout.addLayout(self.create_action_row())
        return group

    def create_keyword_row(self):
        row = QHBoxLayout()
        label = QLabel("搜索关键词")
        label.setFixedWidth(self._search_label_width)
        row.addWidget(label)

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
        label = QLabel("搜索文件夹")
        label.setFixedWidth(self._search_label_width)
        row.addWidget(label)

        self.folder_input = QLineEdit()
        self.folder_input.setReadOnly(True)
        row.addWidget(self.folder_input)

        self.folder_button = QPushButton("选择文件夹")
        self.folder_button.clicked.connect(self.select_folder)
        row.addWidget(self.folder_button)

        return row

    def create_action_row(self):
        row = QHBoxLayout()
        label = QLabel("搜索方式")
        label.setFixedWidth(self._search_label_width)
        row.addWidget(label)

        self.search_mode_combo = QComboBox()
        self.search_mode_combo.addItem("索引搜索", "index")
        self.search_mode_combo.addItem("常规搜索", "regular")
        self.search_mode_combo.setToolTip("推荐使用索引搜索")
        row.addWidget(self.search_mode_combo)
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
        mode = self.search_mode_combo.findData(self.settings.value("search_mode", "index"))
        self.search_mode_combo.setCurrentIndex(max(0, mode))
        # Preserve the old concurrency preference for the unchanged scan engine.
        self.regular_max_workers = max(1, min(16, self.settings.value(
            "thread_count", DEFAULT_MAX_WORKERS, type=int,
        )))
        self.search_history = list(self.settings.value("search_history", []))

        try:
            saved_keywords = json.loads(
                self.settings.value("custom_keywords", "")
            )
            self.custom_keywords = clean_keywords(saved_keywords)
        except (TypeError, json.JSONDecodeError):
            pass

    def save_settings(self):
        self.settings.setValue("search_mode", self.search_mode_combo.currentData())
        self.settings.setValue("search_history", self.search_history[-20:])
        self.settings.setValue(
            "custom_keywords",
            json.dumps(self.custom_keywords, ensure_ascii=False),
        )

    def schedule_index_service(self, *args, delay=400):
        if self._close_pending:
            return
        if self.search_worker is None:
            self.result_table.clear_results()
        self._keep_search_summary = False
        self._live_search_cancel = None
        self._index_search_summary = None
        self._index_results_final = False
        self._search_beeped = False
        self.index_progress = ("正在扫描文件夹…", 0, 0)
        self.update_result_status(True)
        self.index_generation += 1
        self._index_restart_timer.stop()
        if self.index_service is not None:
            self.index_service.stop()
        self.set_index_state("正在扫描文件夹…")
        self._index_restart_timer.start(delay)

    def restart_index_service(self):
        if self._close_pending:
            return
        if self.index_service is not None and self.index_service.thread.is_alive():
            self._index_restart_timer.start(40)
            return
        folder = self.folder_input.text().strip()
        if not os.path.isdir(folder):
            self.set_index_state("请选择搜索文件夹")
            return
        generation = self.index_generation
        try:
            self.index_service = ExcelLiveIndexService(
                folder, self.index_database,
                lambda kind, value: self.index_event_signal.emit(generation, kind, value),
            )
            self.index_service.start()
        except Exception as exc:
            error = f"{type(exc).__name__}: {exc}"
            self.set_index_state("搜索准备失败，请重新选择文件夹")
            if self._queued_index_task is not None:
                task, self._queued_index_task = self._queued_index_task, None
                task.complete(error)
            return
        if self._queued_index_task is not None:
            task, self._queued_index_task = self._queued_index_task, None
            task.dispatch(self.index_service)

    def enqueue_index_search(self, task):
        service = self.index_service
        folder = self.folder_input.text().strip()
        if (service is not None and service.thread.is_alive() and not service.shutdown.is_set()
                and not self._index_restart_timer.isActive()
                and path_key(service.folder) == path_key(folder)):
            task.dispatch(service)
        else:
            self._queued_index_task = task
            self._index_restart_timer.start(0)

    def handle_index_event(self, generation, kind, value):
        if kind == "index_time":
            if not self.index_time:
                self.index_time = value
            return
        if generation != self.index_generation:
            return
        task = self.search_worker if isinstance(self.search_worker, ExcelIndexSearchTask) else None
        if kind == "index_updating":
            self.update_result_status(bool(value))
        elif kind == "live_results":
            cancel, results, outcome = value
            if (cancel is self._live_search_cancel and not cancel.is_set()
                    and self.search_worker is None and self.search_mode_combo.currentData() == "index"):
                self.result_table.replace_results(results)
                self._index_results_final = not outcome["provisional"]
                success, failed, found = outcome["counts"]
                status = "暂定结果" if outcome["provisional"] else "搜索完成"
                self._index_search_summary = f"{status}：找到 {found} 条（成功：{success}，失败：{failed}）"
                self.progress_bar.setToolTip("\n".join(outcome["file_errors"]))
                self.update_result_status(outcome["provisional"])
                if self._index_results_final:
                    self.restore_search_controls()
                    self.beep_search_complete()
        elif kind == "search_final":
            cancel, outcome = value
            if (cancel is self._live_search_cancel and not cancel.is_set()
                    and self.search_worker is None and self.search_mode_combo.currentData() == "index"):
                self._index_results_final = True
                success, failed, found = outcome["counts"]
                self._index_search_summary = f"搜索完成：找到 {found} 条（成功：{success}，失败：{failed}）"
                self.restore_search_controls()
                self.refresh_index_progress()
                self.beep_search_complete()
        elif kind == "live_search_error":
            cancel, error = value
            if cancel is self._live_search_cancel and not cancel.is_set():
                cancel.set()
                self.restore_search_controls()
                self._index_search_summary = f"结果刷新失败，当前结果可能过时：{error}"
                self.progress_bar.setToolTip(error)
                self.refresh_index_progress()
        elif kind == "progress":
            self.index_progress = value
            self.refresh_index_progress()
        elif kind == "search_progress":
            self._index_query_progress = value
            self.refresh_index_progress()
        elif kind == "live":
            text = str(value)
            if text.startswith("持续监听中"):
                # The same service callback immediately sends the detailed ready
                # progress. Do not render a separate bare "ready" frame here.
                self.index_state = "就绪"
            elif "已停止" in text:
                self.set_index_state("已暂停，下次搜索时继续准备")
            elif "未完成" in text or "待恢复" in text:
                self.set_index_state("正在重新准备搜索，请稍候…")
            else:
                # Normal lifecycle notifications must not replace real phase progress.
                self.refresh_index_progress()
        elif kind == "search_batch" and task is not None:
            cancel, results = value
            if cancel is task.search_cancel:
                task.add_batch(results)
        elif kind == "done" and task is not None:
            if isinstance(value, dict) and value.get("search_cancel") is task.search_cancel:
                outcome = value["outcome"]
                if task.action == "search" and isinstance(outcome, dict):
                    self._index_results_final = not outcome.get("provisional", True)
                task.complete(outcome)
        elif kind == "stopped":
            if self._live_search_cancel is not None:
                self._live_search_cancel.set()
            self.restore_search_controls()
            if task is not None:
                task.complete("索引服务已停止。")
            if self._close_pending:
                self.finish_pending_close()

    def update_result_status(self, updating):
        self._index_is_updating = updating
        self.refresh_index_progress()

    def refresh_index_progress(self):
        # Coalesce writer, reader and watcher events into one stable UI frame.
        if not self._progress_refresh_timer.isActive():
            self._progress_refresh_timer.start()

    def render_index_progress(self):
        if (self.search_mode_combo.currentData() != "index"
                or self.search_worker is not None and not isinstance(self.search_worker, ExcelIndexSearchTask)):
            return
        if self._index_results_final and self._index_search_summary:
            text, current, total = self._index_search_summary, 1, 1
        elif self._index_is_updating:
            text, current, total = self.index_progress
        elif isinstance(self.search_worker, ExcelIndexSearchTask) and not self.search_worker.cancelled:
            text, current, total = self._index_query_progress
        elif self._index_search_summary:
            text, current, total = self._index_search_summary, 1, 1
        else:
            text, current, total = self.index_progress
        self.progress_bar.setRange(0, 100)
        target = int(current / total * 100) if total else 0
        if target != self._progress_target:
            self._progress_animation.stop()
            self._progress_target = target
            if target < self.progress_bar.value():
                # A new phase starts at zero rather than animating backwards.
                self.progress_bar.setValue(target)
            else:
                self._progress_animation.setStartValue(max(0, self.progress_bar.value()))
                self._progress_animation.setEndValue(target)
                self._progress_animation.start()
        if self.progress_bar.format() != text:
            self.progress_bar.setFormat(text)

    def set_index_state(self, text):
        self.index_state = text
        if self.index_progress[2] == 0 or text != "就绪" and not text.startswith("正在准备"):
            self.index_progress = (text, 0, 0)
        if self.search_mode_combo.currentData() == "index":
            self.refresh_index_progress()
            return
        task = self.search_worker
        if isinstance(task, ExcelIndexSearchTask) and not task.cancelled:
            self.progress_bar.setFormat("正在搜索…" if text == "就绪" else text)
        elif task is None and self.search_worker is None and not self._keep_search_summary:
            self.progress_bar.setValue(0)
            self.progress_bar.setFormat(
                text if self.search_mode_combo.currentData() == "index" else "就绪"
            )

    def search_mode_changed(self):
        self.update_result_status(self._index_is_updating)
        if self.search_mode_combo.currentData() != "index" and self._live_search_cancel is not None:
            self._live_search_cancel.set()
            self._live_search_cancel = None
        if self.search_worker is not None or self._keep_search_summary:
            return
        if self.search_mode_combo.currentData() == "regular":
            self._progress_animation.stop()
            self._progress_target = None
            self.progress_bar.setValue(0)
            self.progress_bar.setFormat("就绪")
        else:
            self.refresh_index_progress()

    def finish_pending_close(self):
        if (self._close_pending and self.search_worker is None
                and (self.index_service is None or not self.index_service.thread.is_alive())):
            self._index_close_timer.stop()
            self.close()

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

    def start_search(self, action="search"):
        if isinstance(action, bool):
            action = "search"
        if self.search_worker is not None or self._close_pending or self.index_search_pending():
            return
        keyword = self.keyword_input.text().strip()
        folder = self.folder_input.text().strip()
        if action == "search" and not keyword:
            QMessageBox.warning(self, "提示", "请输入搜索关键词")
            return
        if not os.path.isdir(folder):
            QMessageBox.warning(self, "提示", "请选择有效的文件夹")
            return
        if action == "search" and keyword not in self.search_history:
            self.search_history.append(keyword)
        if self._live_search_cancel is not None:
            self._live_search_cancel.set()
            self._live_search_cancel = None
        self.result_table.clear_results()
        self._index_search_summary = None
        self._index_results_final = False
        self._search_beeped = False
        self._index_query_progress = ("正在搜索…", 0, 0)
        self._keep_search_summary = False
        self.search_btn.setEnabled(False)
        self.search_mode_combo.setEnabled(False)
        self.folder_button.setEnabled(False)
        self.folder_input.setEnabled(False)
        self.keyword_input.setEnabled(False)
        for button in self.keyword_buttons:
            button.setEnabled(False)
        self.stop_btn.setEnabled(True)
        if action != "build" and self.search_mode_combo.currentData() != "index":
            self._progress_animation.stop()
            self.progress_bar.setValue(0)
            self.progress_bar.setFormat("准备搜索…")
        self.progress_bar.setToolTip("")
        search_record = None
        try:
            if action == "build" or self.search_mode_combo.currentData() == "index":
                self.search_worker = ExcelIndexSearchTask(self.enqueue_index_search, action)
                if action == "search":
                    self._live_search_cancel = self.search_worker.search_cancel
            else:
                self.search_worker = ExcelSearchWorker(self.process_pool)
            self.search_worker.setup(
                keyword,
                [folder],
                self.regular_max_workers,
                self.custom_keywords,
            )
            self.search_worker.progress_signal.connect(self.update_progress)
            self.search_worker.batch_result_signal.connect(self.result_table.add_batch_results)
            self.search_worker.finished.connect(self.search_finished)
            search_record = {
                "keyword": keyword, "folder": folder, "duration": 0.0,
                "overview": "准备搜索" if action == "search" else "手动更新索引"}
            self.session_searches.append(search_record)
            self.search_worker.start()
            if isinstance(self.search_worker, ExcelIndexSearchTask):
                self.refresh_index_progress()
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
            self.stop_btn.setEnabled(False)
            self.progress_bar.setFormat("正在停止…")
            self.search_worker.stop()
        elif self.index_search_pending():
            self._live_search_cancel.set()
            self._index_search_summary = "搜索已停止；后台继续更新索引"
            self.restore_search_controls()
            self.refresh_index_progress()

    def index_search_pending(self):
        return (self._live_search_cancel is not None
                and not self._live_search_cancel.is_set() and not self._index_results_final)

    def update_progress(self, current, total, found):
        if self.search_worker is None or self.search_worker.cancelled:
            return
        percentage = int(current / total * 100) if total else 0
        self.progress_bar.setValue(percentage)
        self.progress_bar.setFormat(
            f"正在搜索：{current}/{total} ({percentage}%) — 已找到 {found} 条"
        )

    def restore_search_controls(self):
        busy = self.search_worker is not None or self.index_search_pending()
        self.search_btn.setEnabled(not busy)
        self.search_mode_combo.setEnabled(not busy)
        self.folder_button.setEnabled(not busy)
        self.folder_input.setEnabled(not busy)
        self.keyword_input.setEnabled(not busy)
        for button in self.keyword_buttons:
            button.setEnabled(not busy)
        self.stop_btn.setEnabled(busy)

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
        if (isinstance(worker, ExcelIndexSearchTask) and worker.action == "search"
                and not error and not worker.cancelled):
            status = "搜索完成" if self._index_results_final else "暂定结果"
        if isinstance(worker, ExcelIndexSearchTask) and worker.action == "build":
            status = "索引更新失败" if error else "索引更新已停止" if worker.cancelled else "索引更新完成"
        summary = f"{status}：找到 {found} 条（成功：{success}，失败：{failed}）"
        if not isinstance(worker, ExcelIndexSearchTask):
            if not error and not worker.cancelled:
                self.progress_bar.setValue(100)
            self.progress_bar.setFormat(summary)
        self.progress_bar.setToolTip(details)
        self.session_searches[-1]["overview"] = summary
        self._keep_search_summary = True
        self.search_worker = None
        self._queued_index_task = None
        if (error or worker.cancelled) and isinstance(worker, ExcelIndexSearchTask):
            worker.search_cancel.set()
        worker.deleteLater()
        self.restore_search_controls()
        if isinstance(worker, ExcelIndexSearchTask):
            self._index_search_summary = summary
            self.refresh_index_progress()
        if (not error and not worker.cancelled
                and (not isinstance(worker, ExcelIndexSearchTask)
                     or worker.action == "search" and self._index_results_final)):
            self.beep_search_complete()
        self.save_settings()
        if self._close_pending:
            self.close()
        elif error and not worker.cancelled:
            message = QMessageBox(self)
            message.setWindowTitle(status)
            message.setIcon(QMessageBox.Icon.Warning)
            message.setText(error)
            message.setDetailedText(details)
            message.exec()

    def beep_search_complete(self):
        if not self._search_beeped and not self._close_pending:
            self._search_beeped = True
            QApplication.beep()

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
        self._close_pending = True
        self._index_restart_timer.stop()
        if self.index_service is not None:
            self.index_service.stop()
        if self.search_worker is not None:
            self.stop_search()
            event.ignore()
            self._index_close_timer.start()
            return
        if self.index_service is not None and self.index_service.thread.is_alive():
            event.ignore()
            self._index_close_timer.start()
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
               index_time=window.index_time,
               events=window.result_table.session_events)
    sys.exit(exit_code)


if __name__ == "__main__":
    main()
