"""
Cross-platform bid document manager (macOS/Windows)
---------------------------------------------------

This PyQt5 application is an optimized version of the original elevator
bid-file manager. It keeps the original feature set (open/rename/move/
delete/refresh, copy to directory/clipboard/text, PDF-to-image, and the RMB
amount converter) while adding quality-of-life improvements:

* Configurable base path stored via ``QSettings`` with a toolbar action to
  change it at runtime.
* Automatic company list discovery from folders named ``【公司名称】`` under the
  base path, with a manual fallback list when nothing is found.
* Safer file operations with early existence checks and clearer error
  messages.
* A simple text filter to quickly locate companies in large lists.
* Window/layout settings (splitter sizes, zoom level, geometry) persisted
  across launches.

Run ``python tools/bid_manager.py`` to launch the UI. Packaging instructions
for macOS/Windows live in ``tools/bid_manager_build.md``.
"""

from __future__ import annotations

import atexit
import os
import re
import shutil
import struct
import subprocess
import sys
import tempfile
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import List, Optional

from PyQt5.QtCore import (
    QPoint,
    Qt,
    QSettings,
    QFileInfo,
    QMimeData,
    QThread,
    pyqtSignal,
    QUrl,
)
from PyQt5.QtGui import QFont, QFontMetrics, QIcon, QImage
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFileIconProvider,
    QFormLayout,
    QFrame,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMenu,
    QMessageBox,
    QProgressDialog,
    QPushButton,
    QSizePolicy,
    QSlider,
    QSplitter,
    QStatusBar,
    QToolBar,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
    QAction,
    QHeaderView,
)

import fitz

try:
    import docx  # type: ignore
except Exception:
    docx = None


_RMB_DIGITS = "零壹贰叁肆伍陆柒捌玖"
_RMB_UNITS1 = ["", "拾", "佰", "仟"]
_RMB_UNITS2 = ["", "万", "亿", "兆"]


def amount_to_rmb_upper(value_str: str) -> str:
    """Convert an amount string to uppercase RMB without the RMB prefix."""

    if value_str is None:
        raise ValueError("金额为空")
    s = value_str.strip()
    if not s:
        raise ValueError("金额为空")

    s = s.replace(",", "").replace("￥", "").replace("¥", "")
    s = re.sub(r"(?i)\bRMB\b", "", s)

    neg = s.startswith("-")
    if neg:
        s = s[1:]

    s = re.sub(r"[^0-9.]", "", s)
    if s.count(".") > 1 or not re.search(r"\d", s):
        raise ValueError("金额格式无效")

    d = Decimal(s).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    if d == Decimal("0.00"):
        return "零元整"

    total_cents = int((d * 100).to_integral_value(rounding=ROUND_HALF_UP))
    int_part = total_cents // 100
    jiao = (total_cents // 10) % 10
    fen = total_cents % 10

    int_str = _int_to_rmb_upper(int_part)
    prefix = "负" if neg and d != 0 else ""
    if jiao == 0 and fen == 0:
        return prefix + int_str + "元整"

    dec = ""
    if jiao != 0:
        dec += _RMB_DIGITS[jiao] + "角"
    if fen != 0:
        dec += _RMB_DIGITS[fen] + "分"
    if int_part == 0:
        return prefix + "零元" + dec
    return prefix + int_str + "元" + dec


def _int_to_rmb_upper(n: int) -> str:
    if n == 0:
        return "零"

    parts: List[str] = []
    unit_section_idx = 0
    need_zero = False

    while n > 0:
        section = n % 10000
        if section == 0:
            if not need_zero and parts:
                need_zero = True
        else:
            sec_str = _four_digits_to_cn(section)
            if need_zero:
                parts.insert(0, "零")
                need_zero = False
            parts.insert(0, sec_str + _RMB_UNITS2[unit_section_idx])
        n //= 10000
        unit_section_idx += 1

    res = "".join(parts)
    res = re.sub(r"零+$", "", res)
    return res or "零"


def _four_digits_to_cn(num: int) -> str:
    assert 0 <= num <= 9999
    if num == 0:
        return ""
    res: List[str] = []
    unit_pos = 0
    zero_flag = False
    while num > 0:
        digit = num % 10
        if digit == 0:
            if not zero_flag and res:
                res.insert(0, "零")
                zero_flag = True
        else:
            res.insert(0, _RMB_DIGITS[digit] + _RMB_UNITS1[unit_pos])
            zero_flag = False
        unit_pos += 1
        num //= 10
    s = "".join(res)
    s = re.sub(r"零+$", "", s)
    s = re.sub(r"零+", "零", s)
    return s


class AmountConvertDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("数字大小写转化工具（金额转人民币大写）")
        self.setModal(True)
        self.setMinimumWidth(520)

        self.input_edit = QLineEdit()
        self.input_edit.setPlaceholderText("请输入金额，例如：1234.56 或 ￥12,345.67")
        self.result_edit = QLineEdit()
        self.result_edit.setReadOnly(True)
        self.result_edit.setStyleSheet(
            "QLineEdit { background: #f8f9fa; font-weight: bold; }"
        )

        self.tips_label = QLabel("说明：自动四舍五入到分；支持负号、逗号、¥/￥/RMB 前缀。")
        self.tips_label.setStyleSheet("color:#6c757d;")

        self.btn_copy = QPushButton("复制大写金额")
        self.btn_close = QPushButton("关闭")

        form = QFormLayout()
        form.addRow("小写金额：", self.input_edit)
        form.addRow("大写金额：", self.result_edit)

        btns = QHBoxLayout()
        btns.addStretch(1)
        btns.addWidget(self.btn_copy)
        btns.addWidget(self.btn_close)

        layout = QVBoxLayout(self)
        layout.addLayout(form)
        layout.addWidget(self.tips_label)
        layout.addLayout(btns)

        self.input_edit.textChanged.connect(self.on_text_changed)
        self.btn_copy.clicked.connect(self.copy_result)
        self.btn_close.clicked.connect(self.close)

        clip = QApplication.clipboard().text()
        if any(x in clip for x in list("0123456789")):
            self.input_edit.setText(clip)

    def on_text_changed(self, txt: str):
        try:
            upper = amount_to_rmb_upper(txt)
            self.result_edit.setText(upper)
        except Exception as e:  # noqa: BLE001
            self.result_edit.setText(f"（输入无效）{e}")

    def copy_result(self):
        text = self.result_edit.text().strip()
        if not text or text.startswith("（输入无效）"):
            QMessageBox.warning(self, "提示", "当前结果无效，无法复制。")
            return
        QApplication.clipboard().setText(text)
        QMessageBox.information(self, "已复制", f"已复制到剪贴板：\n{text}")


class FileOperationThread(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal(bool, str)

    def __init__(self, operation_type: str, source_path: str, target_path: str | None = None):
        super().__init__()
        self.operation_type = operation_type
        self.source_path = source_path
        self.target_path = target_path
        self.total_files = 0
        self.processed_files = 0

    def run(self):
        try:
            if self.operation_type == "delete":
                self.delete_files()
            elif self.operation_type == "move":
                self.move_files()
            elif self.operation_type == "copy":
                self.copy_files()
            self.finished.emit(True, f"{self.operation_type.capitalize()} operation completed successfully")
        except Exception as e:  # noqa: BLE001
            self.finished.emit(False, f"Error during {self.operation_type} operation: {str(e)}")

    def count_files(self, path: str) -> int:
        if os.path.isfile(path):
            return 1
        count = 0
        for _, _, files in os.walk(path):
            count += len(files)
        return count

    def delete_files(self):
        if os.path.isfile(self.source_path):
            os.remove(self.source_path)
            self.progress.emit(100)
            return

        self.total_files = self.count_files(self.source_path) or 1
        self._delete_folder(self.source_path)
        self.progress.emit(100)

    def _delete_folder(self, folder_path: str):
        if not os.path.exists(folder_path):
            return
        for item in os.listdir(folder_path):
            item_path = os.path.join(folder_path, item)
            if os.path.isfile(item_path):
                os.remove(item_path)
                self.processed_files += 1
                self.progress.emit(int(self.processed_files / self.total_files * 100))
            else:
                self._delete_folder(item_path)
        os.rmdir(folder_path)

    def move_files(self):
        if os.path.isfile(self.source_path):
            shutil.move(self.source_path, self.target_path)
            self.progress.emit(100)
            return

        self.total_files = self.count_files(self.source_path) or 1
        self._move_folder(self.source_path, self.target_path)
        self.progress.emit(100)

    def _move_folder(self, source: str, target: str):
        if not os.path.exists(source):
            return
        if not os.path.exists(target):
            os.makedirs(target)
        for item in os.listdir(source):
            s_item = os.path.join(source, item)
            t_item = os.path.join(target, item)
            if os.path.isfile(s_item):
                shutil.move(s_item, t_item)
                self.processed_files += 1
                self.progress.emit(int(self.processed_files / self.total_files * 100))
            else:
                self._move_folder(s_item, t_item)
        if os.path.exists(source) and not os.listdir(source):
            os.rmdir(source)

    def copy_files(self):
        if os.path.isfile(self.source_path):
            os.makedirs(os.path.dirname(self.target_path), exist_ok=True)
            shutil.copy2(self.source_path, self.target_path)
            self.progress.emit(100)
            return

        self.total_files = self.count_files(self.source_path) or 1
        self._copy_folder(self.source_path, self.target_path)
        self.progress.emit(100)

    def _copy_folder(self, source: str, target: str):
        if not os.path.exists(source):
            return
        if not os.path.exists(target):
            os.makedirs(target)
        for item in os.listdir(source):
            s_item = os.path.join(source, item)
            t_item = os.path.join(target, item)
            if os.path.isfile(s_item):
                shutil.copy2(s_item, t_item)
                self.processed_files += 1
                self.progress.emit(int(self.processed_files / self.total_files * 100))
            else:
                self._copy_folder(s_item, t_item)


class CompanyManager(QMainWindow):
    DEFAULT_COMPANIES = [
        "上海三菱电梯有限公司",
        "上海三菱电梯有限公司江苏分公司",
        "江苏省元菱电梯有限公司",
        "南京骏菱电梯工程有限公司",
        "江苏海菱机电设备工程有限公司南京分公司",
        "卧龙溪岸投标项目",
        "双砚居投标项目",
        "烟草局",
        "红山学院",
        "南京泉峰科技有限公司2025年度电梯维保项目",
        "南通四建电梯工程有限公司南京分公司",
    ]

    def __init__(self):
        super().__init__()
        self.settings = QSettings("BidManager", "CompanyManager")
        self.base_path = Path(
            self.settings.value("base_path", "~/BidLibrary").replace("\\", "/")
        ).expanduser()
        self.font_scale_factor = float(self.settings.value("font_scale", 1.0))
        self.base_font_size = 10
        self.temp_dirs: List[str] = []
        self.file_thread: Optional[FileOperationThread] = None
        self.progress_dialog: Optional[QProgressDialog] = None

        atexit.register(self.cleanup_temp_dirs)

        self.initUI()
        self.load_companies()

    def initUI(self):  # noqa: N802
        self.setWindowTitle("电梯维保资质库管理系统")
        self.resize(1200, 800)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        main_layout = QHBoxLayout(central_widget)
        main_layout.setContentsMargins(5, 5, 5, 5)
        main_layout.setSpacing(5)

        left_container = QFrame()
        left_container.setFrameStyle(QFrame.StyledPanel)
        left_layout = QVBoxLayout(left_container)
        left_layout.setContentsMargins(0, 0, 0, 0)

        header_layout = QHBoxLayout()
        company_label = QLabel("公司列表")
        company_label.setAlignment(Qt.AlignCenter)
        company_label.setStyleSheet(
            """
            QLabel {
                background-color: #007bff;
                color: white;
                padding: 5px;
                font-weight: bold;
                border-radius: 4px;
            }
            """
        )

        self.filter_edit = QLineEdit()
        self.filter_edit.setPlaceholderText("筛选公司...")
        self.filter_edit.textChanged.connect(self.filter_companies)

        header_layout.addWidget(company_label)
        header_layout.addWidget(self.filter_edit)
        left_layout.addLayout(header_layout)

        self.company_list = QListWidget()
        self.company_list.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.company_list.currentItemChanged.connect(self.on_company_selected)
        self.company_list.setContextMenuPolicy(Qt.CustomContextMenu)
        self.company_list.customContextMenuRequested.connect(
            self.on_company_list_context_menu
        )
        left_layout.addWidget(self.company_list)

        self.file_tree = QTreeWidget()
        self.file_tree.setHeaderLabels(["资质文件", "类型", "修改日期"])
        self.file_tree.setColumnWidth(0, 300)
        self.file_tree.header().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.file_tree.itemDoubleClicked.connect(self.on_file_double_clicked)
        self.file_tree.setSelectionMode(QAbstractItemView.ExtendedSelection)

        self.splitter = QSplitter(Qt.Horizontal)
        self.splitter.addWidget(left_container)
        self.splitter.addWidget(self.file_tree)
        self.splitter.setSizes([300, 900])
        self.splitter.setChildrenCollapsible(False)
        main_layout.addWidget(self.splitter)

        toolbar = QToolBar()
        self.addToolBar(toolbar)

        convert_action = QAction("数字转大写", self)
        convert_action.setToolTip("金额转人民币大写（投标书用）")
        convert_action.triggered.connect(self.open_amount_converter)
        toolbar.addAction(convert_action)
        toolbar.addSeparator()

        set_base_action = QAction("设置资质库路径", self)
        set_base_action.setToolTip("选择公司资质库根目录")
        set_base_action.triggered.connect(self.set_base_path)
        toolbar.addAction(set_base_action)
        toolbar.addSeparator()

        open_action = QAction("打开文件", self)
        open_action.triggered.connect(self.open_file)
        toolbar.addAction(open_action)

        open_folder_action = QAction("打开所在文件夹", self)
        open_folder_action.triggered.connect(self.open_file_folder)
        toolbar.addAction(open_folder_action)

        refresh_action = QAction("刷新", self)
        refresh_action.triggered.connect(self.refresh_files)
        toolbar.addAction(refresh_action)

        rename_action = QAction("重命名", self)
        rename_action.triggered.connect(self.rename_file)
        toolbar.addAction(rename_action)

        move_action = QAction("移动文件", self)
        move_action.triggered.connect(self.move_file)
        toolbar.addAction(move_action)

        copy_action = QAction("复制到…", self)
        copy_action.triggered.connect(self.copy_file)
        toolbar.addAction(copy_action)

        direct_copy_action = QAction("直接复制", self)
        direct_copy_action.setShortcut("Ctrl+C")
        direct_copy_action.triggered.connect(self.copy_to_clipboard_files)
        toolbar.addAction(direct_copy_action)

        copy_text_action = QAction("复制文字", self)
        copy_text_action.triggered.connect(self.copy_text_from_file)
        toolbar.addAction(copy_text_action)

        copy_path_action = QAction("复制路径", self)
        copy_path_action.triggered.connect(self.copy_path_to_clipboard)
        toolbar.addAction(copy_path_action)

        delete_action = QAction("删除文件", self)
        delete_action.triggered.connect(self.delete_file)
        toolbar.addAction(delete_action)

        pdf_to_img_action = QAction("PDF转图片", self)
        pdf_to_img_action.triggered.connect(self.pdf_to_images_and_copy)
        toolbar.addAction(pdf_to_img_action)

        delete_temp_action = QAction("删除临时文件", self)
        delete_temp_action.triggered.connect(self.delete_temp_files)
        toolbar.addAction(delete_temp_action)

        toolbar.addSeparator()

        zoom_out_action = QAction("缩小字体", self)
        zoom_out_action.triggered.connect(self.zoom_out)
        toolbar.addAction(zoom_out_action)

        self.zoom_slider = QSlider(Qt.Horizontal)
        self.zoom_slider.setMinimum(50)
        self.zoom_slider.setMaximum(200)
        self.zoom_slider.setValue(int(self.font_scale_factor * 100))
        self.zoom_slider.setTickPosition(QSlider.TicksBelow)
        self.zoom_slider.setTickInterval(10)
        self.zoom_slider.setFixedWidth(100)
        self.zoom_slider.valueChanged.connect(self.on_zoom_changed)
        toolbar.addWidget(self.zoom_slider)

        zoom_in_action = QAction("放大字体", self)
        zoom_in_action.triggered.connect(self.zoom_in)
        toolbar.addAction(zoom_in_action)

        reset_zoom_action = QAction("重置字体", self)
        reset_zoom_action.triggered.connect(self.reset_zoom)
        toolbar.addAction(reset_zoom_action)

        self.statusBar = QStatusBar()
        self.setStatusBar(self.statusBar)

        self.apply_styles()
        self.restore_window_state()

    def apply_styles(self):
        font = QFont("Microsoft YaHei", self.base_font_size)
        self.setFont(font)
        self.setStyleSheet(
            """
            QMainWindow { background-color: #f5f5f5; }
            QListWidget, QTreeWidget {
                background-color: white; border: 1px solid #ddd; border-radius: 4px;
            }
            QTreeWidget { alternate-background-color: #f9f9f9; }
            QToolBar { background-color: #f8f9fa; border-bottom: 1px solid #e9ecef; spacing: 5px; padding: 5px; }
            QToolButton { background-color: #007bff; color: white; border: none; border-radius: 4px; padding: 5px 10px; }
            QToolButton:hover { background-color: #0056b3; }
            QStatusBar { background-color: #e9ecef; color: #6c757d; }
            QListWidget::item:selected, QTreeWidget::item:selected { background-color: #007bff; color: white; }
            QSlider::groove:horizontal { border: 1px solid #999; height: 8px; background: #c4c4c4; margin: 2px 0; }
            QSlider::handle:horizontal { background: #b4b4b4; border: 1px solid #5c5c5c; width: 18px; margin: -2px 0; border-radius: 3px; }
            QSplitter::handle { background-color: #ccc; width: 3px; }
            QSplitter::handle:hover { background-color: #007bff; }
        """
        )
        self.apply_font_zoom()

    def restore_window_state(self):
        geometry = self.settings.value("geometry")
        if geometry:
            self.restoreGeometry(geometry)
        state = self.settings.value("window_state")
        if state:
            self.restoreState(state)
        splitter_sizes = self.settings.value("splitter_sizes")
        if splitter_sizes:
            try:
                self.splitter.setSizes([int(x) for x in splitter_sizes])
            except Exception:  # noqa: BLE001
                pass

    def closeEvent(self, event):  # noqa: N802
        self.settings.setValue("geometry", self.saveGeometry())
        self.settings.setValue("window_state", self.saveState())
        self.settings.setValue("splitter_sizes", self.splitter.sizes())
        self.settings.setValue("font_scale", self.font_scale_factor)
        super().closeEvent(event)

    def showEvent(self, event):  # noqa: N802
        super().showEvent(event)
        QApplication.processEvents()
        self.adjust_company_list_width()
        self.adjust_file_tree_columns()

    def adjust_company_list_width(self):
        max_width = 0
        fm = QFontMetrics(self.company_list.font())
        for i in range(self.company_list.count()):
            text = self.company_list.item(i).text()
            w = fm.horizontalAdvance(text)
            max_width = max(max_width, w)
        self.company_list.setMinimumWidth(min(max_width + 40, 400))

    def adjust_file_tree_columns(self):
        self.file_tree.resizeColumnToContents(0)
        if self.file_tree.columnWidth(0) < 300:
            self.file_tree.setColumnWidth(0, 300)
        for col in range(1, self.file_tree.columnCount()):
            self.file_tree.resizeColumnToContents(col)

    def zoom_in(self):
        v = self.zoom_slider.value()
        if v < self.zoom_slider.maximum():
            self.zoom_slider.setValue(v + 10)

    def zoom_out(self):
        v = self.zoom_slider.value()
        if v > self.zoom_slider.minimum():
            self.zoom_slider.setValue(v - 10)

    def reset_zoom(self):
        self.zoom_slider.setValue(100)

    def on_zoom_changed(self, value):
        self.font_scale_factor = value / 100.0
        self.apply_font_zoom()
        self.statusBar.showMessage(f"字体大小: {value}%", 3000)

    def apply_font_zoom(self):
        new_font_size = int(self.base_font_size * self.font_scale_factor)
        font = QFont("Microsoft YaHei", new_font_size)
        self.setFont(font)
        self.company_list.setFont(font)
        self.file_tree.setFont(font)
        self.statusBar.setFont(font)
        toolbar = self.findChild(QToolBar)
        if toolbar:
            for action in toolbar.actions():
                widget = toolbar.widgetForAction(action)
                if widget:
                    widget.setFont(font)

    def resolve_company_dirs(self) -> List[str]:
        if not self.base_path.exists():
            return []
        pattern = re.compile(r"^【(.+?)】$")
        resolved = []
        for entry in self.base_path.iterdir():
            if entry.is_dir():
                match = pattern.match(entry.name)
                if match:
                    resolved.append(match.group(1))
        resolved.sort(key=str.lower)
        return resolved

    def load_companies(self):
        companies = self.resolve_company_dirs() or self.DEFAULT_COMPANIES
        self.company_list.clear()
        max_width = 0
        fm = QFontMetrics(self.company_list.font())
        for company in companies:
            item = QListWidgetItem(company)
            item.setData(Qt.UserRole, company)
            item.setIcon(QIcon.fromTheme("folder"))
            item.setSizeHint(QSize(100, 40))
            self.company_list.addItem(item)
            max_width = max(max_width, fm.horizontalAdvance(company))
        self.company_list.setMinimumWidth(min(max_width + 40, 400))
        if self.company_list.count() > 0:
            self.company_list.setCurrentRow(0)

    def filter_companies(self, text: str):
        text = text.strip().lower()
        for i in range(self.company_list.count()):
            item = self.company_list.item(i)
            item.setHidden(text not in item.text().lower())

    def on_company_selected(self, current: QListWidgetItem, previous: QListWidgetItem | None):
        if current is None:
            return
        company_name = current.data(Qt.UserRole)
        company_path = self.base_path / f"【{company_name}】"
        if not company_path.exists():
            QMessageBox.warning(self, "路径不存在", f"找不到公司目录: {company_path}")
            return
        self.statusBar.showMessage(f"正在加载: {company_name}")
        self.file_tree.clear()
        self.load_company_structure(company_path, self.file_tree.invisibleRootItem())
        for i in range(self.file_tree.topLevelItemCount()):
            self.file_tree.topLevelItem(i).setExpanded(True)
        self.adjust_file_tree_columns()
        self.statusBar.showMessage(f"已加载: {company_name}")

    def load_company_structure(self, path: Path, parent_item: QTreeWidgetItem):
        try:
            entries = sorted(
                [p for p in path.iterdir() if not p.name.startswith("._") and not p.name.startswith(".")],
                key=lambda p: (not p.is_dir(), p.name.lower()),
            )
            for item in entries:
                if item.is_dir():
                    dir_item = QTreeWidgetItem(parent_item)
                    dir_item.setText(0, item.name)
                    dir_item.setData(0, Qt.UserRole, str(item))
                    dir_item.setIcon(0, QIcon.fromTheme("folder"))
                    self.load_company_structure(item, dir_item)
                else:
                    file_item = QTreeWidgetItem(parent_item)
                    file_item.setText(0, item.name)
                    file_item.setData(0, Qt.UserRole, str(item))
                    fi = QFileInfo(str(item))
                    suffix = fi.suffix().upper()
                    file_item.setText(1, f"{suffix}文件" if suffix else "文件")
                    file_item.setText(2, fi.lastModified().toString("yyyy-MM-dd hh:mm"))
                    icon_provider = QFileIconProvider()
                    file_item.setIcon(0, icon_provider.icon(fi))
        except PermissionError:
            QMessageBox.warning(self, "权限错误", f"无法访问: {path}")

    def on_file_double_clicked(self, item: QTreeWidgetItem, column: int):
        path = item.data(0, Qt.UserRole)
        if path and os.path.exists(path):
            self._open_path(path)

    def _open_path(self, path: str):
        try:
            if sys.platform == "win32":
                os.startfile(path)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", path])
            else:
                subprocess.Popen(["xdg-open", path])
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "错误", f"无法打开: {exc}")

    def open_file(self):
        current_item = self.file_tree.currentItem()
        if current_item and current_item.data(0, Qt.UserRole):
            path = current_item.data(0, Qt.UserRole)
            if os.path.exists(path):
                self._open_path(path)
                self.statusBar.showMessage(f"已打开: {os.path.basename(path)}", 3000)

    def open_file_folder(self):
        current_item = self.file_tree.currentItem()
        if not current_item or not current_item.data(0, Qt.UserRole):
            return
        file_path = current_item.data(0, Qt.UserRole)
        if not os.path.exists(file_path):
            QMessageBox.warning(self, "警告", "文件不存在")
            return
        try:
            if sys.platform == "win32":
                if os.path.isfile(file_path):
                    subprocess.Popen(f'explorer /select,"{file_path}"')
                else:
                    os.startfile(file_path)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", "-R", file_path])
            else:
                subprocess.Popen(["xdg-open", os.path.dirname(file_path)])
            self.statusBar.showMessage("已打开文件所在位置", 3000)
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "错误", f"无法打开文件夹: {str(e)}")

    def refresh_files(self):
        current_company = self.company_list.currentItem()
        if current_company:
            self.on_company_selected(current_company, None)
            self.statusBar.showMessage("已刷新文件列表", 3000)

    def on_company_list_context_menu(self, pos: QPoint):
        item = self.company_list.itemAt(pos)
        if not item:
            return
        menu = QMenu(self)
        act_copy_name = menu.addAction("复制公司名称")

        def _do():
            QApplication.clipboard().setText(item.text())
            self.statusBar.showMessage(f"已复制公司名称：{item.text()}", 3000)

        act_copy_name.triggered.connect(_do)
        menu.exec_(self.company_list.mapToGlobal(pos))

    def contextMenuEvent(self, event):  # noqa: N802
        current_item = self.file_tree.currentItem()
        if not current_item or not current_item.data(0, Qt.UserRole):
            return
        path = current_item.data(0, Qt.UserRole)
        menu = QMenu(self)

        act_open = menu.addAction("打开")
        act_open.triggered.connect(self.open_file)

        act_open_folder = menu.addAction("打开所在文件夹")
        act_open_folder.triggered.connect(self.open_file_folder)

        menu.addSeparator()
        act_rename = menu.addAction("重命名")
        act_rename.triggered.connect(self.rename_file)

        act_direct_copy = menu.addAction("直接复制(到剪贴板)")
        act_direct_copy.triggered.connect(self.copy_to_clipboard_files)

        act_copy = menu.addAction("复制到…")
        act_copy.triggered.connect(self.copy_file)

        act_move = menu.addAction("移动到…")
        act_move.triggered.connect(self.move_file)

        menu.addSeparator()
        if os.path.isfile(path):
            act_copy_text = menu.addAction("复制文字到剪贴板")
            act_copy_text.triggered.connect(self.copy_text_from_file)
        act_copy_path = menu.addAction("复制完整路径")
        act_copy_path.triggered.connect(self.copy_path_to_clipboard)
        act_copy_name = menu.addAction("复制文件名")
        act_copy_name.triggered.connect(self.copy_name_to_clipboard)

        menu.addSeparator()
        if os.path.isfile(path) and path.lower().endswith(".pdf"):
            act_pdf = menu.addAction("PDF转图片并复制（第一页）")
            act_pdf.triggered.connect(self.pdf_to_images_and_copy)

        menu.addSeparator()
        act_delete = menu.addAction("删除")
        act_delete.triggered.connect(self.delete_file)

        menu.exec_(event.globalPos())

    def open_amount_converter(self):
        dlg = AmountConvertDialog(self)
        dlg.exec_()

    def rename_file(self):
        current_item = self.file_tree.currentItem()
        if not current_item or not current_item.data(0, Qt.UserRole):
            return
        file_path = current_item.data(0, Qt.UserRole)
        if not os.path.exists(file_path):
            QMessageBox.warning(self, "警告", "请选择一个文件或文件夹进行重命名")
            return
        dir_path = os.path.dirname(file_path)
        old_name = os.path.basename(file_path)
        new_name, ok = QInputDialog.getText(self, "重命名", "请输入新的名称:", QLineEdit.Normal, old_name)
        if ok and new_name and new_name != old_name:
            try:
                new_path = os.path.join(dir_path, new_name)
                if os.path.exists(new_path):
                    QMessageBox.warning(self, "错误", "已存在同名文件或文件夹")
                    return
                os.rename(file_path, new_path)
                self.refresh_files()
                self.statusBar.showMessage(f"已重命名: {old_name} -> {new_name}", 3000)
            except Exception as e:  # noqa: BLE001
                QMessageBox.critical(self, "错误", f"重命名失败: {str(e)}")

    def move_file(self):
        current_item = self.file_tree.currentItem()
        if not current_item or not current_item.data(0, Qt.UserRole):
            return
        source_path = current_item.data(0, Qt.UserRole)
        if not os.path.exists(source_path):
            QMessageBox.warning(self, "警告", "请选择一个文件或文件夹进行移动")
            return
        target_dir = QFileDialog.getExistingDirectory(self, "选择目标目录", str(self.base_path))
        if not target_dir:
            return
        name = os.path.basename(source_path)
        target_path = os.path.join(target_dir, name)
        if os.path.exists(target_path):
            reply = QMessageBox.question(
                self,
                "确认覆盖",
                f"目标目录已存在同名项：\n{name}\n是否覆盖？",
                QMessageBox.Yes | QMessageBox.No,
            )
            if reply == QMessageBox.No:
                return
        self.progress_dialog = QProgressDialog(f"正在移动 {name}...", "取消", 0, 100, self)
        self.progress_dialog.setWindowTitle("移动文件")
        self.progress_dialog.setWindowModality(Qt.WindowModal)
        self.progress_dialog.show()

        self.file_thread = FileOperationThread("move", source_path, target_path)
        self.file_thread.progress.connect(self.progress_dialog.setValue)
        self.file_thread.finished.connect(self.on_file_operation_finished)
        self.file_thread.start()

    def copy_file(self):
        current_item = self.file_tree.currentItem()
        if not current_item or not current_item.data(0, Qt.UserRole):
            return
        source_path = current_item.data(0, Qt.UserRole)
        if not os.path.exists(source_path):
            QMessageBox.warning(self, "警告", "请选择一个文件或文件夹进行复制")
            return
        target_dir = QFileDialog.getExistingDirectory(self, "选择目标目录", str(self.base_path))
        if not target_dir:
            return
        name = os.path.basename(source_path)
        target_path = os.path.join(target_dir, name)
        if os.path.exists(target_path):
            reply = QMessageBox.question(
                self,
                "确认覆盖",
                f"目标目录已存在同名项：\n{name}\n是否覆盖？",
                QMessageBox.Yes | QMessageBox.No,
            )
            if reply == QMessageBox.No:
                return
        self.progress_dialog = QProgressDialog(f"正在复制 {name}...", "取消", 0, 100, self)
        self.progress_dialog.setWindowTitle("复制文件")
        self.progress_dialog.setWindowModality(Qt.WindowModal)
        self.progress_dialog.show()

        self.file_thread = FileOperationThread("copy", source_path, target_path)
        self.file_thread.progress.connect(self.progress_dialog.setValue)
        self.file_thread.finished.connect(self.on_file_operation_finished)
        self.file_thread.start()

    def delete_file(self):
        current_item = self.file_tree.currentItem()
        if not current_item or not current_item.data(0, Qt.UserRole):
            return
        file_path = current_item.data(0, Qt.UserRole)
        if not os.path.exists(file_path):
            QMessageBox.warning(self, "警告", "文件或文件夹不存在")
            return
        name = os.path.basename(file_path)
        reply = QMessageBox.question(
            self,
            "确认删除",
            f"确定要删除 {name} 吗？此操作无法撤销。",
            QMessageBox.Yes | QMessageBox.No,
        )
        if reply == QMessageBox.Yes:
            self.progress_dialog = QProgressDialog(f"正在删除 {name}...", "取消", 0, 100, self)
            self.progress_dialog.setWindowTitle("删除文件")
            self.progress_dialog.setWindowModality(Qt.WindowModal)
            self.progress_dialog.show()

            self.file_thread = FileOperationThread("delete", file_path)
            self.file_thread.progress.connect(self.progress_dialog.setValue)
            self.file_thread.finished.connect(self.on_file_operation_finished)
            self.file_thread.start()

    def on_file_operation_finished(self, success: bool, message: str):
        if self.progress_dialog:
            self.progress_dialog.close()
            self.progress_dialog = None
        if success:
            self.refresh_files()
            self.statusBar.showMessage(message, 5000)
        else:
            QMessageBox.critical(self, "错误", message)

    def copy_to_clipboard_files(self):
        items = self.file_tree.selectedItems()
        if not items:
            QMessageBox.information(self, "提示", "请先选择要复制的文件或文件夹")
            return

        urls: List[QUrl] = []
        for it in items:
            p = it.data(0, Qt.UserRole)
            if p and os.path.exists(p):
                urls.append(QUrl.fromLocalFile(p))

        if not urls:
            QMessageBox.warning(self, "提示", "没有可复制的有效路径")
            return

        mime = QMimeData()
        mime.setUrls(urls)
        try:
            mime.setData('application/x-qt-windows-mime;value="Preferred DropEffect"', struct.pack("<I", 1))
        except Exception:
            pass

        QApplication.clipboard().setMimeData(mime)
        self.statusBar.showMessage(f"已复制 {len(urls)} 项到剪贴板，可在资源管理器中 Ctrl+V 粘贴", 4000)

    def copy_text_from_file(self):
        current_item = self.file_tree.currentItem()
        if not current_item or not current_item.data(0, Qt.UserRole):
            return
        file_path = current_item.data(0, Qt.UserRole)
        if not os.path.isfile(file_path):
            QMessageBox.information(self, "提示", "请选择具体文件再复制文字")
            return

        ext = os.path.splitext(file_path)[1].lower()
        text = ""
        ok = True
        err: Optional[str] = None

        try:
            if ext in (".txt", ".csv", ".py", ".json", ".md", ".log"):
                with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                    text = f.read()

            elif ext == ".pdf":
                pdf = fitz.open(file_path)
                total = len(pdf)
                prog = QProgressDialog("正在从 PDF 提取文字…", "取消", 0, total, self)
                prog.setWindowModality(Qt.WindowModal)
                prog.setWindowTitle("提取文字")
                prog.show()
                chunks: List[str] = []
                for i, page in enumerate(pdf):
                    chunks.append(page.get_text("text"))
                    prog.setValue(i + 1)
                    QApplication.processEvents()
                    if prog.wasCanceled():
                        ok = False
                        break
                pdf.close()
                if ok:
                    text = "\n".join(chunks)

            elif ext == ".docx":
                if docx is None:
                    ok = False
                    err = "未安装 python-docx，无法读取 DOCX。请执行：pip install python-docx"
                else:
                    d = docx.Document(file_path)
                    text = "\n".join(p.text for p in d.paragraphs)

            else:
                ok = False
                err = f"暂不支持从此类型文件提取文字：{ext}"

        except Exception as e:  # noqa: BLE001
            ok = False
            err = str(e)

        if ok:
            QApplication.clipboard().setText(text or "")
            msg = f"已复制文字到剪贴板（{len(text)} 字符）"
            self.statusBar.showMessage(msg, 4000)
            if not text.strip():
                QMessageBox.information(self, "完成", "已复制，但未提取到有效文字。")
        else:
            QMessageBox.warning(self, "无法复制文字", err or "未知错误")

    def copy_path_to_clipboard(self):
        current_item = self.file_tree.currentItem()
        if not current_item or not current_item.data(0, Qt.UserRole):
            return
        path = current_item.data(0, Qt.UserRole)
        QApplication.clipboard().setText(path)
        self.statusBar.showMessage("已复制完整路径到剪贴板", 3000)

    def copy_name_to_clipboard(self):
        current_item = self.file_tree.currentItem()
        if not current_item or not current_item.data(0, Qt.UserRole):
            return
        name = os.path.basename(current_item.data(0, Qt.UserRole))
        QApplication.clipboard().setText(name)
        self.statusBar.showMessage("已复制文件名到剪贴板", 3000)

    def delete_temp_files(self):
        if not self.temp_dirs:
            QMessageBox.information(self, "提示", "没有可删除的临时文件")
            return
        reply = QMessageBox.question(
            self,
            "确认删除",
            f"确定要删除所有临时文件吗？\n共 {len(self.temp_dirs)} 个临时文件夹将被删除。",
            QMessageBox.Yes | QMessageBox.No,
        )
        if reply != QMessageBox.Yes:
            return
        deleted, failed = 0, 0
        for temp_dir in self.temp_dirs[:]:
            try:
                if os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir)
                    deleted += 1
                self.temp_dirs.remove(temp_dir)
            except Exception:
                failed += 1
        if failed == 0:
            QMessageBox.information(self, "完成", f"已成功删除 {deleted} 个临时文件夹")
        else:
            QMessageBox.warning(self, "部分完成", f"成功 {deleted} 个，失败 {failed} 个")
        self.statusBar.showMessage(f"已删除 {deleted} 个临时文件夹", 3000)

    def cleanup_temp_dirs(self):
        for temp_dir in self.temp_dirs:
            try:
                if os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir)
            except Exception as e:  # noqa: BLE001
                print(f"清理临时文件夹失败: {temp_dir}, 错误: {str(e)}")

    def pdf_to_images_and_copy(self):
        current_item = self.file_tree.currentItem()
        if not current_item or not current_item.data(0, Qt.UserRole):
            return
        file_path = current_item.data(0, Qt.UserRole)
        if not os.path.isfile(file_path) or not file_path.lower().endswith(".pdf"):
            QMessageBox.warning(self, "警告", "请选择一个 PDF 文件")
            return
        try:
            pdf_document = fitz.open(file_path)
            page_count = len(pdf_document)
            mat = fitz.Matrix(2.0, 2.0)
            temp_dir = tempfile.mkdtemp(prefix="pdf_images_")
            self.temp_dirs.append(temp_dir)

            image_paths: List[str] = []
            for page_num in range(page_count):
                page = pdf_document.load_page(page_num)
                pix = page.get_pixmap(matrix=mat)
                image_path = os.path.join(temp_dir, f"page_{page_num + 1}.png")
                pix.save(image_path)
                image_paths.append(image_path)
            pdf_document.close()

            if image_paths:
                image = QImage(image_paths[0])
                QApplication.clipboard().setImage(image)
                self.statusBar.showMessage(
                    f"PDF已转换为{page_count}张图片，已复制第一页到剪贴板。临时文件：{temp_dir}",
                    5000,
                )
                reply = QMessageBox.question(
                    self,
                    "转换成功",
                    f"PDF已转换为{page_count}张图片，已复制第一页到剪贴板。\n"
                    f"临时文件保存在: {temp_dir}\n\n是否打开临时文件夹查看所有图片?",
                    QMessageBox.Yes | QMessageBox.No,
                )
                if reply == QMessageBox.Yes:
                    self._open_path(temp_dir)
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "错误", f"PDF转换失败: {str(e)}")

    def set_base_path(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("设置资质库路径")
        layout = QFormLayout(dialog)
        path_edit = QLineEdit(str(self.base_path))
        layout.addRow("资质库路径:", path_edit)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addRow(buttons)
        if dialog.exec_() == QDialog.Accepted:
            self.base_path = Path(path_edit.text()).expanduser()
            self.settings.setValue("base_path", str(self.base_path))
            self.load_companies()


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    manager = CompanyManager()
    manager.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
