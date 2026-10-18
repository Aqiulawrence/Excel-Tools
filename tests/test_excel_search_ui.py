"""Exercise queued QThread delivery and window close with an offscreen Qt app."""

import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from types import SimpleNamespace

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from openpyxl import Workbook
from PyQt6.QtCore import QEventLoop, QTimer
from PyQt6.QtWidgets import QApplication

import ExcelSearch as search


class SearchWindowTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])
        cls.app.setQuitOnLastWindowClosed(False)

    def setUp(self):
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        self.folder = Path(folder.name)
        for index in range(24):
            book = Workbook()
            sheet = book.active
            sheet.append(["description", "RMB", "USD", "quantity", "amount"])
            sheet.append([f"Alpha {index}", 12.5, 2, 4, 50])
            sheet.append([f"Beta {index}", 10, 1.5, 3, 30])
            book.save(self.folder / f"{index}.xlsx")
            book.close()
        # Do not touch the user's registry settings or trigger usage reporting.
        for method in ("load_settings", "save_settings"):
            mock = patch.object(search.ExcelSearchTool, method)
            mock.start()
            self.addCleanup(mock.stop)
        self.window = search.ExcelSearchTool()
        self.addCleanup(self.window.close)
        self.addCleanup(self.window.process_pool.shutdown)
        self.window.folder_input.setText(str(self.folder))
        self.window.thread_spin.setValue(2)
        self.window.show()

    def wait_finished(self):
        loop = QEventLoop()
        timer = QTimer()
        timer.setInterval(10)
        timer.timeout.connect(lambda: loop.quit() if self.window.search_worker is None else None)
        timeout = QTimer()
        timeout.setSingleShot(True)
        timeout.timeout.connect(loop.quit)
        timer.start()
        timeout.start(15000)
        loop.exec()
        timer.stop()
        timeout.stop()
        self.assertIsNone(self.window.search_worker, "search did not complete within 15 seconds")
        self.app.processEvents()

    def start(self, term):
        self.window.keyword_input.setText(term)
        self.window.start_search()
        self.assertIsNotNone(self.window.search_worker)

    def test_queued_results_progress_and_new_keyword(self):
        for term in ("Alpha", "Beta"):
            self.start(term)
            self.wait_finished()
            table = self.window.result_table
            self.assertEqual(table.rowCount(), 24)
            self.assertEqual(len(table.results), 24)
            self.assertTrue(all(term in item.cell_value for item in table.results))
            for index, result in enumerate(table.results):
                self.assertEqual(table.item(index, 2).text(), result.cell_value)
                self.assertEqual(table.item(index, 4).text(), result.rmb_price)
            self.assertEqual(self.window.progress_bar.value(), 100)
            self.assertEqual(self.window.progress_bar.maximum(), 100)
            self.assertTrue(self.window.search_btn.isEnabled())

    def test_stop_then_restart_does_not_mix_queued_results(self):
        self.start("Alpha")
        self.window.search_worker.progress_signal.connect(lambda *_: self.window.stop_search())
        self.wait_finished()
        self.start("Beta")
        self.wait_finished()
        self.assertEqual(len(self.window.result_table.results), 24)
        self.assertTrue(all("Beta" in r.cell_value for r in self.window.result_table.results))

    def test_close_during_search_drains_and_shuts_down_pool(self):
        self.start("Alpha")
        self.window.search_worker.progress_signal.connect(lambda *_: self.window.close())
        self.wait_finished()
        self.assertFalse(self.window.isVisible())
        self.assertIsNone(self.window.process_pool.executor)

    def test_scan_errors_remain_visible_without_completion_popup(self):
        def denied():
            self.window.search_worker.scan_errors.append("restricted: Access denied")
            return iter(())

        with patch.object(search.ExcelSearchWorker, "collect_files", side_effect=denied), \
                patch.object(search.QMessageBox, "exec", return_value=0) as warning:
            self.start("Alpha")
            self.wait_finished()
            self.assertTrue(self.window.progress_bar.format().startswith("搜索完成："))
            self.assertNotIn("部分路径失败", self.window.progress_bar.format())
            self.assertIn("restricted", self.window.progress_bar.toolTip())
            warning.assert_not_called()

    def test_completion_status_popup_and_record_matrix(self):
        cases = [
            ("normal", False, "", [], [], "搜索完成：", False),
            ("directory-only", False, "", [], ["denied"], "搜索完成：", False),
            ("stopped", True, "", [], ["denied"], "已停止：", False),
            ("file-error", False, "", ["broken.xlsx"], [], "搜索完成（部分文件读取失败）：", True),
            ("mixed", False, "", ["broken.xlsx"], ["denied"], "搜索完成（部分文件读取失败）：", True),
            ("fatal", False, "pool failed", [], ["denied"], "搜索失败：", True),
        ]
        for name, cancelled, error, file_errors, scan_errors, prefix, popup in cases:
            with self.subTest(name=name), patch.object(search.QMessageBox, "exec", return_value=0) as warning:
                self.window.session_searches.append({})
                self.window.search_worker = SimpleNamespace(
                    duration_seconds=1.0, counts=(3, len(file_errors), 2),
                    error=error, file_errors=file_errors, scan_errors=scan_errors,
                    cancelled=cancelled, deleteLater=lambda: None,
                )
                self.window.search_finished()
                status = self.window.progress_bar.format()
                self.assertTrue(status.startswith(prefix), status)
                self.assertEqual(self.window.session_searches[-1]["overview"], status)
                self.assertEqual(warning.call_count, int(popup))
                self.assertTrue(self.window.search_btn.isEnabled())
                self.assertFalse(self.window.stop_btn.isEnabled())


if __name__ == "__main__":
    unittest.main()
