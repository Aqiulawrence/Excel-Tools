import sys
import os
from PyQt6.QtWidgets import *
from PyQt6.QtCore import *
from PyQt6.QtGui import *
import core_logic


class DragDropLabel(QLabel):
    file_dropped = pyqtSignal(str)

    def __init__(self, text="拖拽Excel文件到这里"):
        super().__init__(text)
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setFixedSize(280, 70)
        self.setStyleSheet(
            "border: 3px dashed #aaa; border-radius: 10px; font-size: 14px; color: #666; background: #f9f9f9; padding: 5px;")
        self.setAcceptDrops(True)
        self.selected = False

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls() and event.mimeData().urls()[0].toLocalFile().lower().endswith('.xlsx'):
            event.accept()
            self.setStyleSheet(
                "border: 3px dashed #0078d7; border-radius: 10px; font-size: 14px; color: #0078d7; background: #e6f2ff; padding: 5px;")

    def dragLeaveEvent(self, event):
        self.setStyleSheet(
            "border: 3px solid #4caf50; border-radius: 10px; font-size: 14px; color: #666; background: #f9f9f9; padding: 5px;" if self.selected else "border: 3px dashed #aaa; border-radius: 10px; font-size: 14px; color: #666; background: #f9f9f9; padding: 5px;")

    def dropEvent(self, event):
        file = event.mimeData().urls()[0].toLocalFile()
        self.file_dropped.emit(file)
        self.setText(f"已选择:\n{file.split('/')[-1]}")
        self.selected = True
        self.setStyleSheet(
            "border: 3px solid #4caf50; border-radius: 10px; font-size: 14px; color: #666; background: #f9f9f9; padding: 5px;")

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            file, _ = QFileDialog.getOpenFileName(self, "选择Excel文件", "", "Excel文件 (*.xlsx *.xls)")
            if file:
                self.file_dropped.emit(file)
                self.setText(f"已选择:\n{file.split('/')[-1]}")
                self.selected = True
                self.setStyleSheet(
                    "border: 3px solid #4caf50; border-radius: 10px; font-size: 14px; color: #666; background: #f9f9f9; padding: 5px;")


class App(QWidget):
    def __init__(self):
        super().__init__()
        self.result_texts = []
        self.comparator = core_logic.ExcelComparator()
        self.load_keywords()
        self.file1 = ""
        self.file2 = ""
        self.init_ui()

    def load_keywords(self):
        if not os.path.exists('./keywords.txt'):
            with open('./keywords.txt', 'w', encoding='utf-8') as f:
                f.write("part,件号\nqty,quantit,pcs in ctn,数量,每箱数")
        with open('./keywords.txt', 'r', encoding='utf-8') as f:
            lines = f.readlines()
            self.comparator.PART_KEYWORDS = lines[0].strip().split(',')
            self.comparator.QTY_KEYWORDS = lines[1].strip().split(',')

    def init_ui(self):
        self.setWindowTitle("Excel件号数量对比工具 by Sam")
        self.setFixedSize(700, 600)

        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)

        title = QLabel("Excel件号数量对比工具 by Sam")
        title.setStyleSheet("font-size: 24px; font-weight: bold; color: #2c3e50;")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(title)

        drop_layout = QHBoxLayout()
        drop_layout.setSpacing(20)

        self.drop1 = DragDropLabel("点击或拖拽\n第一个Excel文件")
        self.drop1.file_dropped.connect(lambda f: self.set_file(1, f))
        drop_layout.addWidget(self.drop1)

        vs_label = QLabel("VS")
        vs_label.setStyleSheet("font-size: 24px; font-weight: bold; color: #666;")
        vs_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        drop_layout.addWidget(vs_label)

        self.drop2 = DragDropLabel("点击或拖拽\n第二个Excel文件")
        self.drop2.file_dropped.connect(lambda f: self.set_file(2, f))
        drop_layout.addWidget(self.drop2)

        main_layout.addLayout(drop_layout)

        self.compare_btn = QPushButton("开始对比")
        self.compare_btn.setFixedHeight(45)
        self.compare_btn.clicked.connect(self.compare_files)
        self.compare_btn.setEnabled(False)
        self.compare_btn.setStyleSheet(
            "background: #e0e0e0; color: #666; border: none; border-radius: 8px; font-size: 16px; font-weight: bold; padding: 10px;")
        main_layout.addWidget(self.compare_btn)

        result_layout = QHBoxLayout()
        result_layout.setSpacing(15)

        for title in ["前者缺失", "后者缺失", "数量差异"]:
            group = QGroupBox(title)
            group.setStyleSheet(
                "font-size: 16px; font-weight: bold; color: #2c3e50; border: 2px solid #ddd; border-radius: 8px; margin-top: 10px; padding-top: 10px;")
            layout = QVBoxLayout()
            text_edit = QTextEdit()
            text_edit.setReadOnly(True)
            text_edit.setStyleSheet("font-size: 14px;")
            layout.addWidget(text_edit)
            group.setLayout(layout)
            result_layout.addWidget(group)
            self.result_texts.append(text_edit)

        main_layout.addLayout(result_layout)
        self.setLayout(main_layout)
        self.center()

    def center(self):
        screen = QApplication.primaryScreen().geometry()
        self.move((screen.width() - self.width()) // 2, (screen.height() - self.height()) // 2)

    def set_file(self, num, file_path):
        if num == 1:
            self.file1 = file_path
        else:
            self.file2 = file_path
        enabled = bool(self.file1 and self.file2)
        self.compare_btn.setEnabled(enabled)
        self.compare_btn.setStyleSheet(
            f"background: {'#0078d7' if enabled else '#e0e0e0'}; color: {'white' if enabled else '#666'}; border: none; border-radius: 8px; font-size: 16px; font-weight: bold; padding: 10px;")

    def compare_files(self):
        try:
            result = self.comparator.compare(self.file1, self.file2)
            for text_edit in self.result_texts:
                text_edit.clear()
            for i, (desc, dic) in enumerate(result.items()):
                if dic:
                    output = [f"{part}: {value}" for part, value in dic.items()]
                    total = sum(abs(v) for v in dic.values())
                    output.append(f"\n总计: {total}")
                    self.result_texts[i].setText("\n".join(output))
                else:
                    self.result_texts[i].setText(f"无{desc[2:4]}")
            if not any(result.values()):
                QMessageBox.information(self, "提示", "✅ 件号所对数量完全一致！")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"{str(e)}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = App()
    window.show()
    sys.exit(app.exec())
