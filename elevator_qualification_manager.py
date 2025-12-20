# -*- coding: utf-8 -*-
"""
电梯维保资质库管理系统（PyQt6）
- 支持：打开 / 重命名 / 移动 / 删除 / 刷新
- 复制：
  1) 复制到目录（线程+进度）
  2) 直接复制到剪贴板（支持多选，资源管理器 Ctrl+V 粘贴）
  3) 复制文字（TXT / PDF / DOCX）
  4) 复制路径 / 复制文件名 / 复制公司名称
- PDF 转图片（第一页自动复制到剪贴板，可打开临时目录）
- 新增：数字大小写转化工具（金额转人民币大写），一键复制

可选依赖：
  pip install python-docx PyMuPDF
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

import fitz  # PyMuPDF
from PyQt6.QtCore import (
    QMimeData,
    QPoint,
    QThread,
    QUrl,
    QSize,
    Qt,
    QFileInfo,
    pyqtSignal,
)
from PyQt6.QtGui import QFont, QFontMetrics, QIcon, QImage
from PyQt6.QtWidgets import (
    QApplication,
    QAbstractItemView,
    QAction,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFileIconProvider,
    QFormLayout,
    QFrame,
    QHeaderView,
    QHBoxLayout,
    QLabel,
    QInputDialog,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMenu,
    QMessageBox,
    QProgressDialog,
    QPushButton,
    QSlider,
    QSizePolicy,
    QSplitter,
    QStatusBar,
    QToolBar,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

# ===== 可选依赖：python-docx（DOCX 文本提取） =====
try:
    import docx  # python-docx
except Exception:
    docx = None


# ===================== 金额转大写（人民币）工具 =====================
_RMB_DIGITS = "零壹贰叁肆伍陆柒捌玖"
_RMB_UNITS1 = ["", "拾", "佰", "仟"]  # 个、十、百、千
_RMB_UNITS2 = ["", "万", "亿", "兆"]  # 万、亿、兆（按4位分节）


def amount_to_rmb_upper(value_str: str) -> str:
    """
    将金额字符串转换为人民币大写（不带“人民币”前缀），规则：
    - 四舍五入到分
    - 整数部分使用“元”，小数部分用“角”“分”；若无小数，末尾加“整”
    - 0 -> 零元整
    - 负数在最前面加“负”
    """
    if value_str is None:
        raise ValueError("金额为空")
    s = value_str.strip()
    if not s:
        raise ValueError("金额为空")

    # 允许常见前缀与分隔符
    s = s.replace(",", "")
    s = s.replace("￥", "").replace("¥", "")
    s = re.sub(r"(?i)\bRMB\b", "", s)

    neg = False
    if s.startswith("-"):
        neg = True
        s = s[1:]

    # 保留数字和一个小数点
    s = re.sub(r"[^0-9.]", "", s)
    if s.count(".") > 1 or not re.search(r"\d", s):
        raise ValueError("金额格式无效")

    # 使用 Decimal 精确到分
    d = Decimal(s).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    # 防止 -0.00
    if d == Decimal("0.00"):
        return "零元整"

    # 拆分整数与小数
    total_cents = int((d * 100).to_integral_value(rounding=ROUND_HALF_UP))
    int_part = total_cents // 100
    jiao = (total_cents // 10) % 10
    fen = total_cents % 10

    int_str = _int_to_rmb_upper(int_part)

    # 组装
    prefix = "负" if neg and d != 0 else ""
    if jiao == 0 and fen == 0:
        return prefix + int_str + "元整"
    dec = ""
    if jiao != 0:
        dec += _RMB_DIGITS[jiao] + "角"
    if fen != 0:
        dec += _RMB_DIGITS[fen] + "分"
    # 如果整数为0但有角分，整数部分仍显示“零元”
    if int_part == 0:
        return prefix + "零元" + dec
    return prefix + int_str + "元" + dec


def _int_to_rmb_upper(n: int) -> str:
    """整数部分（>=0）转中文大写（不含“元”）"""
    if n == 0:
        return "零"

    parts: list[str] = []
    unit_section_idx = 0
    need_zero = False  # 跨节需要补零

    while n > 0:
        section = n % 10000  # 4位一节
        if section == 0:
            if not need_zero and parts:  # 仅当更高节已有内容时，标记需要零
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
    # 清理可能的尾部“零”
    res = re.sub(r"零+$", "", res)
    # 处理“壹拾X”是否简写为“拾X”——人民币大写一般保留“壹拾”
    return res or "零"


def _four_digits_to_cn(num: int) -> str:
    """将 0-9999 转中文，内部用于节内转换"""
    assert 0 <= num <= 9999
    if num == 0:
        return ""
    res = []
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
    s = re.sub(r"零+$", "", s)  # 去掉节末尾零
    s = re.sub(r"零+", "零", s)  # 合并多零
    return s


class AmountConvertDialog(QDialog):
    """金额大小写转换弹窗"""

    def __init__(self, parent: QWidget | None = None):
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

        # 若剪贴板有可能的金额，自动填充
        clip = QApplication.clipboard().text()
        if any(x in clip for x in list("0123456789")):
            self.input_edit.setText(clip)

    def on_text_changed(self, txt: str) -> None:
        try:
            upper = amount_to_rmb_upper(txt)
            self.result_edit.setText(upper)
        except Exception as exc:  # noqa: BLE001
            self.result_edit.setText(f"（输入无效）{exc}")

    def copy_result(self) -> None:
        text = self.result_edit.text().strip()
        if not text or text.startswith("（输入无效）"):
            QMessageBox.warning(self, "提示", "当前结果无效，无法复制。")
            return
        QApplication.clipboard().setText(text)
        QMessageBox.information(self, "已复制", f"已复制到剪贴板：\n{text}")


# ===================== 文件操作线程 =====================
class FileOperationThread(QThread):
    """文件操作线程：复制 / 移动 / 删除（递归），支持进度信号"""

    progress = pyqtSignal(int)
    finished = pyqtSignal(bool, str)

    def __init__(self, operation_type: str, source_path: str, target_path: str | None = None):
        super().__init__()
        self.operation_type = operation_type  # 'copy', 'move', 'delete'
        self.source_path = source_path
        self.target_path = target_path
        self.total_files = 0
        self.processed_files = 0

    def run(self) -> None:  # type: ignore[override]
        try:
            if self.operation_type == "delete":
                self.delete_files()
            elif self.operation_type == "move":
                self.move_files()
            elif self.operation_type == "copy":
                self.copy_files()
            self.finished.emit(True, f"{self.operation_type.capitalize()} operation completed successfully")
        except Exception as exc:  # noqa: BLE001
            self.finished.emit(False, f"Error during {self.operation_type} operation: {str(exc)}")

    def count_files(self, path: str) -> int:
        """递归计算文件总数"""
        if os.path.isfile(path):
            return 1
        count = 0
        for _, _, files in os.walk(path):
            count += len(files)
        return count

    # ==== 删除 ====
    def delete_files(self) -> None:
        if os.path.isfile(self.source_path):
            os.remove(self.source_path)
            self.progress.emit(100)
            return
        self.total_files = self.count_files(self.source_path) or 1
        self._delete_folder(self.source_path)
        self.progress.emit(100)

    def _delete_folder(self, folder_path: str) -> None:
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

    # ==== 移动 ====
    def move_files(self) -> None:
        if os.path.isfile(self.source_path):
            shutil.move(self.source_path, self.target_path)
            self.progress.emit(100)
            return
        self.total_files = self.count_files(self.source_path) or 1
        self._move_folder(self.source_path, self.target_path)
        self.progress.emit(100)

    def _move_folder(self, source: str, target: str) -> None:
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

    # ==== 复制 ====
    def copy_files(self) -> None:
        if os.path.isfile(self.source_path):
            os.makedirs(os.path.dirname(self.target_path), exist_ok=True)
            shutil.copy2(self.source_path, self.target_path)
            self.progress.emit(100)
            return
        self.total_files = self.count_files(self.source_path) or 1
        self._copy_folder(self.source_path, self.target_path)
        self.progress.emit(100)

    def _copy_folder(self, source: str, target: str) -> None:
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
    def __init__(self):
        super().__init__()
        self.base_path = "/Volumes/移动U盘/电梯维保投标核心资质库"
        self.companies = [
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

        self.font_scale_factor = 1.0
        self.base_font_size = 10

        self.temp_dirs: list[str] = []
        self.file_thread: FileOperationThread | None = None
        self.progress_dialog: QProgressDialog | None = None

        atexit.register(self.cleanup_temp_dirs)

        self.initUI()
        self.load_companies()

    # ---------- UI ----------
    def initUI(self) -> None:
        self.setWindowTitle("电梯维保资质库管理系统")
        self.setGeometry(100, 100, 1200, 800)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        main_layout = QHBoxLayout(central_widget)
        main_layout.setContentsMargins(5, 5, 5, 5)
        main_layout.setSpacing(5)

        # 左侧：公司列表
        left_container = QFrame()
        left_container.setFrameStyle(QFrame.Shape.StyledPanel)
        left_layout = QVBoxLayout(left_container)
        left_layout.setContentsMargins(0, 0, 0, 0)

        company_label = QLabel("公司列表")
        company_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
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
        left_layout.addWidget(company_label)

        self.company_list = QListWidget()
        self.company_list.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.company_list.currentItemChanged.connect(self.on_company_selected)
        # 右键复制公司名
        self.company_list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.company_list.customContextMenuRequested.connect(self.on_company_list_context_menu)
        left_layout.addWidget(self.company_list)

        # 右侧：文件树
        self.file_tree = QTreeWidget()
        self.file_tree.setHeaderLabels(["资质文件", "类型", "修改日期"])
        self.file_tree.setColumnWidth(0, 300)
        self.file_tree.header().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.file_tree.itemDoubleClicked.connect(self.on_file_double_clicked)
        # 支持多选（直接复制到剪贴板时很有用）
        self.file_tree.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)

        # 分割
        self.splitter = QSplitter(Qt.Orientation.Horizontal)
        self.splitter.addWidget(left_container)
        self.splitter.addWidget(self.file_tree)
        self.splitter.setSizes([300, 900])
        self.splitter.setChildrenCollapsible(False)
        main_layout.addWidget(self.splitter)

        # 工具栏
        toolbar = QToolBar()
        self.addToolBar(toolbar)

        # === 新增：数字大小写转化工具（放在第一个） ===
        convert_action = QAction("数字转大写", self)
        convert_action.setToolTip("金额转人民币大写（投标书用）")
        convert_action.triggered.connect(self.open_amount_converter)
        toolbar.addAction(convert_action)
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

        # 复制到目录
        copy_action = QAction("复制到…", self)
        copy_action.triggered.connect(self.copy_file)
        toolbar.addAction(copy_action)

        # 直接复制到剪贴板（多选）
        direct_copy_action = QAction("直接复制", self)
        direct_copy_action.setShortcut("Ctrl+C")
        direct_copy_action.triggered.connect(self.copy_to_clipboard_files)
        toolbar.addAction(direct_copy_action)

        # 复制文字
        copy_text_action = QAction("复制文字", self)
        copy_text_action.triggered.connect(self.copy_text_from_file)
        toolbar.addAction(copy_text_action)

        # 复制路径
        copy_path_action = QAction("复制路径", self)
        copy_path_action.triggered.connect(self.copy_path_to_clipboard)
        toolbar.addAction(copy_path_action)

        # 删除
        delete_action = QAction("删除文件", self)
        delete_action.triggered.connect(self.delete_file)
        toolbar.addAction(delete_action)

        # PDF转图片
        pdf_to_img_action = QAction("PDF转图片", self)
        pdf_to_img_action.triggered.connect(self.pdf_to_images_and_copy)
        toolbar.addAction(pdf_to_img_action)

        # 删除临时文件
        delete_temp_action = QAction("删除临时文件", self)
        delete_temp_action.triggered.connect(self.delete_temp_files)
        toolbar.addAction(delete_temp_action)

        toolbar.addSeparator()

        zoom_out_action = QAction("缩小字体", self)
        zoom_out_action.triggered.connect(self.zoom_out)
        toolbar.addAction(zoom_out_action)

        self.zoom_slider = QSlider(Qt.Orientation.Horizontal)
        self.zoom_slider.setMinimum(50)
        self.zoom_slider.setMaximum(200)
        self.zoom_slider.setValue(100)
        self.zoom_slider.setTickPosition(QSlider.TickPosition.TicksBelow)
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

        # 状态栏
        self.statusBar = QStatusBar()
        self.setStatusBar(self.statusBar)

        self.apply_styles()

    def apply_styles(self) -> None:
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

    def showEvent(self, event) -> None:  # type: ignore[override]
        super().showEvent(event)
        QApplication.processEvents()
        self.adjust_company_list_width()
        self.adjust_file_tree_columns()

    def adjust_company_list_width(self) -> None:
        max_width = 0
        fm = QFontMetrics(self.company_list.font())
        for i in range(self.company_list.count()):
            text = self.company_list.item(i).text()
            w = fm.horizontalAdvance(text)
            max_width = max(max_width, w)
        self.company_list.setMinimumWidth(min(max_width + 40, 400))

    def adjust_file_tree_columns(self) -> None:
        self.file_tree.resizeColumnToContents(0)
        if self.file_tree.columnWidth(0) < 300:
            self.file_tree.setColumnWidth(0, 300)
        for col in range(1, self.file_tree.columnCount()):
            self.file_tree.resizeColumnToContents(col)

    # ---------- 缩放 ----------
    def zoom_in(self) -> None:
        value = self.zoom_slider.value()
        if value < self.zoom_slider.maximum():
            self.zoom_slider.setValue(value + 10)

    def zoom_out(self) -> None:
        value = self.zoom_slider.value()
        if value > self.zoom_slider.minimum():
            self.zoom_slider.setValue(value - 10)

    def reset_zoom(self) -> None:
        self.zoom_slider.setValue(100)

    def on_zoom_changed(self, value: int) -> None:
        self.font_scale_factor = value / 100.0
        self.apply_font_zoom()
        self.statusBar.showMessage(f"字体大小: {value}%", 3000)

    def apply_font_zoom(self) -> None:
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

    # ---------- 数据加载 ----------
    def load_companies(self) -> None:
        self.company_list.clear()
        max_width = 0
        fm = QFontMetrics(self.company_list.font())
        for company in self.companies:
            item = QListWidgetItem(company)
            item.setData(Qt.ItemDataRole.UserRole, company)
            item.setIcon(QIcon.fromTheme("folder"))
            item.setSizeHint(QSize(100, 40))
            self.company_list.addItem(item)
            max_width = max(max_width, fm.horizontalAdvance(company))
        self.company_list.setMinimumWidth(min(max_width + 40, 400))
        if self.company_list.count() > 0:
            self.company_list.setCurrentRow(0)

    def on_company_selected(self, current: QListWidgetItem, previous) -> None:  # type: ignore[override]
        if current is None:
            return
        company_name = current.data(Qt.ItemDataRole.UserRole)
        company_path = Path(self.base_path) / f"【{company_name}】"
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

    def load_company_structure(self, path: Path, parent_item: QTreeWidgetItem) -> None:
        try:
            for item in sorted(
                [p for p in path.iterdir() if not p.name.startswith("._") and not p.name.startswith(".")],
                key=lambda p: (not p.is_dir(), p.name.lower()),
            ):
                if item.is_dir():
                    dir_item = QTreeWidgetItem(parent_item)
                    dir_item.setText(0, item.name)
                    dir_item.setData(0, Qt.ItemDataRole.UserRole, str(item))
                    dir_item.setIcon(0, QIcon.fromTheme("folder"))
                    self.load_company_structure(item, dir_item)
                else:
                    file_item = QTreeWidgetItem(parent_item)
                    file_item.setText(0, item.name)
                    file_item.setData(0, Qt.ItemDataRole.UserRole, str(item))
                    fi = QFileInfo(str(item))
                    file_item.setText(1, fi.suffix().upper() + "文件")
                    file_item.setText(2, fi.lastModified().toString("yyyy-MM-dd hh:mm"))
                    icon_provider = QFileIconProvider()
                    file_item.setIcon(0, icon_provider.icon(fi))
        except PermissionError:
            QMessageBox.warning(self, "权限错误", f"无法访问: {path}")

    # ---------- 交互 ----------
    def on_file_double_clicked(self, item: QTreeWidgetItem, column: int) -> None:  # type: ignore[override]
        path = item.data(0, Qt.ItemDataRole.UserRole)
        if path and os.path.exists(path):
            os.startfile(path)

    def open_file(self) -> None:
        current_item = self.file_tree.currentItem()
        if current_item and current_item.data(0, Qt.ItemDataRole.UserRole):
            path = current_item.data(0, Qt.ItemDataRole.UserRole)
            if os.path.exists(path):
                try:
                    os.startfile(path)
                    self.statusBar.showMessage(f"已打开: {os.path.basename(path)}", 3000)
                except Exception as exc:  # noqa: BLE001
                    QMessageBox.critical(self, "错误", f"无法打开: {str(exc)}")

    def open_file_folder(self) -> None:
        current_item = self.file_tree.currentItem()
        if current_item and current_item.data(0, Qt.ItemDataRole.UserRole):
            file_path = current_item.data(0, Qt.ItemDataRole.UserRole)
            if os.path.exists(file_path):
                try:
                    if sys.platform == "win32":
                        if os.path.isfile(file_path):
                            subprocess.Popen(f'explorer /select,"{file_path}"')
                        else:
                            os.startfile(file_path)
                    elif sys.platform == "darwin":
                        subprocess.Popen(["open", "-R", file_path])
                    elif sys.platform == "linux":
                        subprocess.Popen(["xdg-open", os.path.dirname(file_path)])
                    self.statusBar.showMessage("已打开文件所在位置", 3000)
                except Exception as exc:  # noqa: BLE001
                    QMessageBox.critical(self, "错误", f"无法打开文件夹: {str(exc)}")

    def refresh_files(self) -> None:
        current_company = self.company_list.currentItem()
        if current_company:
            self.on_company_selected(current_company, None)
            self.statusBar.showMessage("已刷新文件列表", 3000)

    # ---- 公司列表右键：复制公司名称 ----
    def on_company_list_context_menu(self, pos: QPoint) -> None:
        item = self.company_list.itemAt(pos)
        if not item:
            return
        menu = QMenu(self)
        act_copy_name = menu.addAction("复制公司名称")

        def _do() -> None:
            QApplication.clipboard().setText(item.text())
            self.statusBar.showMessage(f"已复制公司名称：{item.text()}", 3000)

        act_copy_name.triggered.connect(_do)
        menu.exec(self.company_list.mapToGlobal(pos))

    # ---- 文件树右键菜单 ----
    def contextMenuEvent(self, event) -> None:  # type: ignore[override]
        current_item = self.file_tree.currentItem()
        if not current_item or not current_item.data(0, Qt.ItemDataRole.UserRole):
            return
        path = current_item.data(0, Qt.ItemDataRole.UserRole)
        menu = QMenu(self)

        # 通用
        act_open = menu.addAction("打开")
        act_open.triggered.connect(self.open_file)

        act_open_folder = menu.addAction("打开所在文件夹")
        act_open_folder.triggered.connect(self.open_file_folder)

        menu.addSeparator()
        act_rename = menu.addAction("重命名")
        act_rename.triggered.connect(self.rename_file)

        # 直接复制到剪贴板（多选）
        act_direct_copy = menu.addAction("直接复制(到剪贴板)")
        act_direct_copy.triggered.connect(self.copy_to_clipboard_files)

        # 复制到目录
        act_copy = menu.addAction("复制到…")
        act_copy.triggered.connect(self.copy_file)

        # 移动
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

        menu.exec(event.globalPos())

    # ---------- 新增：打开金额转换器 ----------
    def open_amount_converter(self) -> None:
        dlg = AmountConvertDialog(self)
        dlg.exec()

    # ---------- 重命名 ----------
    def rename_file(self) -> None:
        current_item = self.file_tree.currentItem()
        if not current_item or not current_item.data(0, Qt.ItemDataRole.UserRole):
            return
        file_path = current_item.data(0, Qt.ItemDataRole.UserRole)
        if not os.path.exists(file_path):
            QMessageBox.warning(self, "警告", "请选择一个文件或文件夹进行重命名")
            return
        dir_path = os.path.dirname(file_path)
        old_name = os.path.basename(file_path)
        new_name, ok = QInputDialog.getText(self, "重命名", "请输入新的名称:", QLineEdit.EchoMode.Normal, old_name)
        if ok and new_name and new_name != old_name:
            try:
                new_path = os.path.join(dir_path, new_name)
                if os.path.exists(new_path):
                    QMessageBox.warning(self, "错误", "已存在同名文件或文件夹")
                    return
                os.rename(file_path, new_path)
                self.refresh_files()
                self.statusBar.showMessage(f"已重命名: {old_name} -> {new_name}", 3000)
            except Exception as exc:  # noqa: BLE001
                QMessageBox.critical(self, "错误", f"重命名失败: {str(exc)}")

    # ---------- 移动 ----------
    def move_file(self) -> None:
        current_item = self.file_tree.currentItem()
        if not current_item or not current_item.data(0, Qt.ItemDataRole.UserRole):
            return
        source_path = current_item.data(0, Qt.ItemDataRole.UserRole)
        if not os.path.exists(source_path):
            QMessageBox.warning(self, "警告", "请选择一个文件或文件夹进行移动")
            return
        target_dir = QFileDialog.getExistingDirectory(self, "选择目标目录", self.base_path)
        if not target_dir:
            return
        name = os.path.basename(source_path)
        target_path = os.path.join(target_dir, name)
        if os.path.exists(target_path):
            reply = QMessageBox.question(
                self,
                "确认覆盖",
                f"目标目录已存在同名项：\n{name}\n是否覆盖？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if reply == QMessageBox.StandardButton.No:
                return
        self.progress_dialog = QProgressDialog(f"正在移动 {name}...", "取消", 0, 100, self)
        self.progress_dialog.setWindowTitle("移动文件")
        self.progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
        self.progress_dialog.show()

        self.file_thread = FileOperationThread("move", source_path, target_path)
        self.file_thread.progress.connect(self.progress_dialog.setValue)
        self.file_thread.finished.connect(self.on_file_operation_finished)
        self.file_thread.start()

    # ---------- 复制到目录 ----------
    def copy_file(self) -> None:
        current_item = self.file_tree.currentItem()
        if not current_item or not current_item.data(0, Qt.ItemDataRole.UserRole):
            return
        source_path = current_item.data(0, Qt.ItemDataRole.UserRole)
        if not os.path.exists(source_path):
            QMessageBox.warning(self, "警告", "请选择一个文件或文件夹进行复制")
            return
        target_dir = QFileDialog.getExistingDirectory(self, "选择目标目录", self.base_path)
        if not target_dir:
            return
        name = os.path.basename(source_path)
        target_path = os.path.join(target_dir, name)
        if os.path.exists(target_path):
            reply = QMessageBox.question(
                self,
                "确认覆盖",
                f"目标目录已存在同名项：\n{name}\n是否覆盖？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if reply == QMessageBox.StandardButton.No:
                return
        self.progress_dialog = QProgressDialog(f"正在复制 {name}...", "取消", 0, 100, self)
        self.progress_dialog.setWindowTitle("复制文件")
        self.progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
        self.progress_dialog.show()

        self.file_thread = FileOperationThread("copy", source_path, target_path)
        self.file_thread.progress.connect(self.progress_dialog.setValue)
        self.file_thread.finished.connect(self.on_file_operation_finished)
        self.file_thread.start()

    # ---------- 删除 ----------
    def delete_file(self) -> None:
        current_item = self.file_tree.currentItem()
        if not current_item or not current_item.data(0, Qt.ItemDataRole.UserRole):
            return
        file_path = current_item.data(0, Qt.ItemDataRole.UserRole)
        if not os.path.exists(file_path):
            QMessageBox.warning(self, "警告", "文件或文件夹不存在")
            return
        name = os.path.basename(file_path)
        reply = QMessageBox.question(
            self,
            "确认删除",
            f"确定要删除 {name} 吗？此操作无法撤销。",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if reply == QMessageBox.StandardButton.Yes:
            self.progress_dialog = QProgressDialog(f"正在删除 {name}...", "取消", 0, 100, self)
            self.progress_dialog.setWindowTitle("删除文件")
            self.progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
            self.progress_dialog.show()

            self.file_thread = FileOperationThread("delete", file_path)
            self.file_thread.progress.connect(self.progress_dialog.setValue)
            self.file_thread.finished.connect(self.on_file_operation_finished)
            self.file_thread.start()

    def on_file_operation_finished(self, success: bool, message: str) -> None:
        if self.progress_dialog:
            self.progress_dialog.close()
            self.progress_dialog = None
        if success:
            self.refresh_files()
            self.statusBar.showMessage(message, 5000)
        else:
            QMessageBox.critical(self, "错误", message)

    # ---------- 直接复制到剪贴板（多选） ----------
    def copy_to_clipboard_files(self) -> None:
        """将选中的文件/文件夹放入系统剪贴板；资源管理器中 Ctrl+V 即可粘贴"""
        items = self.file_tree.selectedItems()
        if not items:
            QMessageBox.information(self, "提示", "请先选择要复制的文件或文件夹")
            return

        urls = []
        for it in items:
            path = it.data(0, Qt.ItemDataRole.UserRole)
            if path and os.path.exists(path):
                urls.append(QUrl.fromLocalFile(path))

        if not urls:
            QMessageBox.warning(self, "提示", "没有可复制的有效路径")
            return

        mime = QMimeData()
        mime.setUrls(urls)
        # 明确告知 Windows：复制（1 = CF_PREFERRED_DROPEFFECT: COPY）
        try:
            mime.setData('application/x-qt-windows-mime;value="Preferred DropEffect"', struct.pack('<I', 1))
        except Exception:  # noqa: BLE001
            pass

        QApplication.clipboard().setMimeData(mime)
        self.statusBar.showMessage(f"已复制 {len(urls)} 项到剪贴板，可在资源管理器中 Ctrl+V 粘贴", 4000)

    # ---------- 复制文字（TXT / PDF / DOCX） ----------
    def copy_text_from_file(self) -> None:
        """从选中文件提取文本并复制到剪贴板；支持 txt / pdf / docx"""
        current_item = self.file_tree.currentItem()
        if not current_item or not current_item.data(0, Qt.ItemDataRole.UserRole):
            return
        file_path = current_item.data(0, Qt.ItemDataRole.UserRole)
        if not os.path.isfile(file_path):
            QMessageBox.information(self, "提示", "请选择具体文件再复制文字")
            return

        ext = os.path.splitext(file_path)[1].lower()
        text = ""
        ok = True
        err = None

        try:
            if ext in (".txt", ".csv", ".py", ".json", ".md", ".log"):
                with open(file_path, "r", encoding="utf-8", errors="ignore") as file:
                    text = file.read()

            elif ext == ".pdf":
                pdf = fitz.open(file_path)
                total = len(pdf)
                prog = QProgressDialog("正在从 PDF 提取文字…", "取消", 0, total, self)
                prog.setWindowModality(Qt.WindowModality.WindowModal)
                prog.setWindowTitle("提取文字")
                prog.show()
                chunks = []
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
                    document = docx.Document(file_path)
                    text = "\n".join(paragraph.text for paragraph in document.paragraphs)

            else:
                ok = False
                err = f"暂不支持从此类型文件提取文字：{ext}"

        except Exception as exc:  # noqa: BLE001
            ok = False
            err = str(exc)

        if ok:
            QApplication.clipboard().setText(text or "")
            msg = f"已复制文字到剪贴板（{len(text)} 字符）"
            self.statusBar.showMessage(msg, 4000)
            if not text.strip():
                QMessageBox.information(self, "完成", "已复制，但未提取到有效文字。")
        else:
            QMessageBox.warning(self, "无法复制文字", err or "未知错误")

    # ---------- 复制路径 / 名称 ----------
    def copy_path_to_clipboard(self) -> None:
        current_item = self.file_tree.currentItem()
        if not current_item or not current_item.data(0, Qt.ItemDataRole.UserRole):
            return
        path = current_item.data(0, Qt.ItemDataRole.UserRole)
        QApplication.clipboard().setText(path)
        self.statusBar.showMessage("已复制完整路径到剪贴板", 3000)

    def copy_name_to_clipboard(self) -> None:
        current_item = self.file_tree.currentItem()
        if not current_item or not current_item.data(0, Qt.ItemDataRole.UserRole):
            return
        name = os.path.basename(current_item.data(0, Qt.ItemDataRole.UserRole))
        QApplication.clipboard().setText(name)
        self.statusBar.showMessage("已复制文件名到剪贴板", 3000)

    # ---------- 临时文件 / PDF转图片 ----------
    def delete_temp_files(self) -> None:
        if not self.temp_dirs:
            QMessageBox.information(self, "提示", "没有可删除的临时文件")
            return
        reply = QMessageBox.question(
            self,
            "确认删除",
            f"确定要删除所有临时文件吗？\n共 {len(self.temp_dirs)} 个临时文件夹将被删除。",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        deleted, failed = 0, 0
        for temp_dir in self.temp_dirs[:]:
            try:
                if os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir)
                    deleted += 1
                self.temp_dirs.remove(temp_dir)
            except Exception:  # noqa: BLE001
                failed += 1
        if failed == 0:
            QMessageBox.information(self, "完成", f"已成功删除 {deleted} 个临时文件夹")
        else:
            QMessageBox.warning(self, "部分完成", f"成功 {deleted} 个，失败 {failed} 个")
        self.statusBar.showMessage(f"已删除 {deleted} 个临时文件夹", 3000)

    def cleanup_temp_dirs(self) -> None:
        for temp_dir in self.temp_dirs:
            try:
                if os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir)
            except Exception as exc:  # noqa: BLE001
                print(f"清理临时文件夹失败: {temp_dir}, 错误: {str(exc)}")

    def pdf_to_images_and_copy(self) -> None:
        current_item = self.file_tree.currentItem()
        if not current_item or not current_item.data(0, Qt.ItemDataRole.UserRole):
            return
        file_path = current_item.data(0, Qt.ItemDataRole.UserRole)
        if not os.path.isfile(file_path) or not file_path.lower().endswith(".pdf"):
            QMessageBox.warning(self, "警告", "请选择一个 PDF 文件")
            return
        try:
            pdf_document = fitz.open(file_path)
            page_count = len(pdf_document)
            mat = fitz.Matrix(2.0, 2.0)
            temp_dir = tempfile.mkdtemp(prefix="pdf_images_")
            self.temp_dirs.append(temp_dir)

            image_paths = []
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
                    f"PDF已转换为{page_count}张图片，已复制第一页到剪贴板。临时文件：{temp_dir}", 5000,
                )
                reply = QMessageBox.question(
                    self,
                    "转换成功",
                    f"PDF已转换为{page_count}张图片，已复制第一页到剪贴板。\n"
                    f"临时文件保存在: {temp_dir}\n\n是否打开临时文件夹查看所有图片?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                )
                if reply == QMessageBox.StandardButton.Yes:
                    if sys.platform == "win32":
                        os.startfile(temp_dir)
                    elif sys.platform == "darwin":
                        subprocess.Popen(["open", temp_dir])
                    elif sys.platform == "linux":
                        subprocess.Popen(["xdg-open", temp_dir])
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "错误", f"PDF转换失败: {str(exc)}")

    # ---------- 其他 ----------
    def set_base_path(self) -> None:
        dialog = QDialog(self)
        dialog.setWindowTitle("设置资质库路径")
        layout = QFormLayout(dialog)
        path_edit = QLineEdit(self.base_path)
        layout.addRow("资质库路径:", path_edit)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addRow(buttons)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.base_path = path_edit.text()
            self.load_companies()


def main() -> None:
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    manager = CompanyManager()
    manager.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
