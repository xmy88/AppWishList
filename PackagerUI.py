from __future__ import annotations

import sys
import subprocess
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional

from PyQt6.QtCore import QThread, pyqtSignal
from PyQt6.QtGui import QIcon
from PyQt6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QPlainTextEdit,
    QVBoxLayout,
    QWidget,
)


@dataclass
class BuildOptions:
    script_path: Path
    output_dir: Path
    app_name: str
    icon_path: Optional[Path]
    target_platform: str
    one_file: bool
    windowed: bool
    clean: bool


class PackThread(QThread):
    log_line = pyqtSignal(str)
    finished = pyqtSignal(int)

    def __init__(self, command: List[str]) -> None:
        super().__init__()
        self.command = command

    def run(self) -> None:  # pragma: no cover - side effects only
        process = subprocess.Popen(
            self.command,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
        )
        if process.stdout:
            for line in process.stdout:
                self.log_line.emit(line.rstrip())
        process.wait()
        self.finished.emit(process.returncode or 0)


class PackagerWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("跨平台 PyQt6 打包工具")
        self.setMinimumSize(760, 520)

        if sys.platform.startswith("darwin"):
            self.setWindowIcon(QIcon.fromTheme("folder"))

        self.script_input = QLineEdit()
        self.output_input = QLineEdit(str(Path.cwd() / "dist"))
        self.app_name_input = QLineEdit("AppWishList")
        self.icon_input = QLineEdit()
        self.platform_select = QComboBox()
        self.platform_select.addItems(["Windows (.exe)", "macOS (.app)"])

        self.one_file_check = QCheckBox("单文件 (--onefile)")
        self.one_file_check.setChecked(True)
        self.windowed_check = QCheckBox("隐藏控制台 (--windowed)")
        self.windowed_check.setChecked(True)
        self.clean_check = QCheckBox("清理上次构建 (--clean)")

        self.log_panel = QPlainTextEdit()
        self.log_panel.setReadOnly(True)
        self.status_label = QLabel("准备就绪")

        form_layout = QFormLayout()
        form_layout.addRow("入口脚本:", self._with_button(self.script_input, self.select_script))
        form_layout.addRow("输出目录:", self._with_button(self.output_input, self.select_output_dir))
        form_layout.addRow("应用名称:", self.app_name_input)
        form_layout.addRow("图标路径:", self._with_button(self.icon_input, self.select_icon))
        form_layout.addRow("目标平台:", self.platform_select)

        option_row = QHBoxLayout()
        option_row.addWidget(self.one_file_check)
        option_row.addWidget(self.windowed_check)
        option_row.addWidget(self.clean_check)

        self.build_button = QPushButton("开始打包")
        self.build_button.clicked.connect(self.start_packaging)

        layout = QVBoxLayout()
        layout.addLayout(form_layout)
        layout.addLayout(option_row)
        layout.addWidget(self.build_button)
        layout.addWidget(QLabel("构建日志:"))
        layout.addWidget(self.log_panel)
        layout.addWidget(self.status_label)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

        self.pack_thread: Optional[PackThread] = None

    def _with_button(self, line_edit: QLineEdit, handler) -> QWidget:
        button = QPushButton("选择")
        button.setFixedWidth(80)
        button.clicked.connect(handler)
        wrapper = QWidget()
        layout = QHBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(line_edit)
        layout.addWidget(button)
        wrapper.setLayout(layout)
        return wrapper

    def select_script(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "选择入口脚本", str(Path.cwd()), "Python Files (*.py)")
        if path:
            self.script_input.setText(path)
            if not self.app_name_input.text():
                self.app_name_input.setText(Path(path).stem)

    def select_output_dir(self) -> None:
        path = QFileDialog.getExistingDirectory(self, "选择输出目录", str(Path.cwd()))
        if path:
            self.output_input.setText(path)

    def select_icon(self) -> None:
        filters = "Icon Files (*.ico *.icns *.png)"
        path, _ = QFileDialog.getOpenFileName(self, "选择图标", str(Path.cwd()), filters)
        if path:
            self.icon_input.setText(path)

    def start_packaging(self) -> None:
        options = self.collect_options()
        if not options:
            return

        command = self.build_command(options)
        self.log_panel.clear()
        self.log_panel.appendPlainText("执行命令: " + " ".join(command))

        self.build_button.setEnabled(False)
        self.status_label.setText("正在打包……")

        self.pack_thread = PackThread(command)
        self.pack_thread.log_line.connect(self.log_panel.appendPlainText)
        self.pack_thread.finished.connect(self.handle_finished)
        self.pack_thread.start()

    def collect_options(self) -> Optional[BuildOptions]:
        script_path = Path(self.script_input.text()).expanduser()
        output_dir = Path(self.output_input.text()).expanduser()
        icon_path = Path(self.icon_input.text()).expanduser() if self.icon_input.text() else None
        app_name = self.app_name_input.text().strip() or script_path.stem

        if not script_path.exists():
            self.show_warning("入口脚本不存在，请重新选择。")
            return None

        if not output_dir.exists():
            output_dir.mkdir(parents=True, exist_ok=True)

        target = "windows" if self.platform_select.currentIndex() == 0 else "macos"

        if target == "macos" and sys.platform.startswith("win"):
            self.show_warning("在 Windows 上创建 macOS 包需要在 macOS 环境执行 PyInstaller。")
        if target == "windows" and sys.platform.startswith("darwin"):
            self.show_warning("在 macOS 上创建 Windows 包建议在 Windows 环境运行 PyInstaller。")

        if icon_path and not icon_path.exists():
            self.show_warning("图标路径无效，将忽略图标设置。")
            icon_path = None

        return BuildOptions(
            script_path=script_path,
            output_dir=output_dir,
            app_name=app_name,
            icon_path=icon_path,
            target_platform=target,
            one_file=self.one_file_check.isChecked(),
            windowed=self.windowed_check.isChecked(),
            clean=self.clean_check.isChecked(),
        )

    def build_command(self, options: BuildOptions) -> List[str]:
        command = [
            sys.executable,
            "-m",
            "PyInstaller",
            f"--name={options.app_name}",
            f"--distpath={options.output_dir}",
            f"--workpath={options.output_dir / 'build'}",
        ]

        if options.icon_path:
            command.append(f"--icon={options.icon_path}")
        if options.one_file:
            command.append("--onefile")
        if options.windowed:
            command.append("--windowed")
        if options.clean:
            command.append("--clean")

        if options.target_platform == "macos" and sys.platform.startswith("darwin"):
            command.append("--target-arch=universal2")

        command.append(str(options.script_path))
        return command

    def handle_finished(self, exit_code: int) -> None:
        success = exit_code == 0
        self.status_label.setText("打包完成" if success else "打包失败，查看日志")
        self.build_button.setEnabled(True)
        if success:
            self.log_panel.appendPlainText("构建完成，文件位于: " + self.output_input.text())
        else:
            self.show_warning("打包失败，请查看日志输出。")

    def show_warning(self, message: str) -> None:
        QMessageBox.warning(self, "提示", message)


def main() -> None:  # pragma: no cover - GUI start
    app = QApplication(sys.argv)
    window = PackagerWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
