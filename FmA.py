import sys
import os
import subprocess
import importlib
import time
import base64
from pathlib import Path
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QLabel, QLineEdit, QPushButton, QFileDialog, QComboBox,
                             QCheckBox, QProgressBar, QTextEdit, QMessageBox, QFrame,
                             QGroupBox, QDialog)
from PyQt5.QtGui import QFont, QIcon, QPalette, QColor, QPixmap
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QSize


class DependencyChecker(QThread):
    """依赖检查与安装线程"""
    progress = pyqtSignal(str)
    finished = pyqtSignal(bool)

    REQUIRED_PACKAGES = [
        'pandas',
        'openpyxl',
        'python-docx',
        'chardet',
        'psutil',
        'xlrd'
    ]

    def run(self):
        """检查并安装所需依赖"""
        self.progress.emit("正在检查依赖...")
        missing_packages = []

        # 检查所有必需包
        for package in self.REQUIRED_PACKAGES:
            if not self.is_package_installed(package):
                missing_packages.append(package)

        if not missing_packages:
            self.progress.emit("所有依赖已安装")
            self.finished.emit(True)
            return

        self.progress.emit(f"缺少依赖: {', '.join(missing_packages)}")
        self.progress.emit("正在尝试安装...")

        # 尝试安装缺失包
        success = self.install_packages(missing_packages)

        if success:
            self.progress.emit("依赖安装成功!")
            self.finished.emit(True)
        else:
            self.progress.emit("依赖安装失败，请手动安装")
            self.finished.emit(False)

    def is_package_installed(self, package_name):
        """检查包是否已安装"""
        try:
            importlib.import_module(package_name)
            return True
        except ImportError:
            return False

    def install_packages(self, packages):
        """安装指定的包"""
        try:
            # 使用pip安装包
            for package in packages:
                self.progress.emit(f"安装 {package}...")
                subprocess.check_call([sys.executable, "-m", "pip", "install", package])
            return True
        except subprocess.CalledProcessError:
            return False
        except Exception as e:
            self.progress.emit(f"安装错误: {str(e)}")
            return False


class DependencyDialog(QDialog):
    """依赖检查对话框"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("依赖检查")
        self.setWindowIcon(parent.windowIcon())
        self.setFixedSize(400, 200)

        layout = QVBoxLayout()
        self.setLayout(layout)

        # 标题
        title_label = QLabel("依赖检查中...")
        title_label.setStyleSheet("font-size: 16px; font-weight: bold;")
        title_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(title_label)

        # 进度标签
        self.progress_label = QLabel("正在初始化...")
        self.progress_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.progress_label)

        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 0)  # 不确定模式
        layout.addWidget(self.progress_bar)

        # 按钮区域
        button_layout = QHBoxLayout()
        self.retry_button = QPushButton("重试")
        self.retry_button.setVisible(False)
        self.retry_button.clicked.connect(self.retry_check)

        self.cancel_button = QPushButton("退出")
        self.cancel_button.clicked.connect(self.reject)

        button_layout.addStretch()
        button_layout.addWidget(self.retry_button)
        button_layout.addWidget(self.cancel_button)
        layout.addLayout(button_layout)

        # 启动依赖检查线程
        self.checker = DependencyChecker()
        self.checker.progress.connect(self.update_progress)
        self.checker.finished.connect(self.on_check_finished)
        self.checker.start()

    def update_progress(self, message):
        """更新进度消息"""
        self.progress_label.setText(message)

    def on_check_finished(self, success):
        """依赖检查完成"""
        if success:
            self.accept()
        else:
            self.progress_bar.setRange(0, 100)
            self.progress_bar.setValue(100)
            self.retry_button.setVisible(True)
            self.progress_label.setText("依赖安装失败，请手动安装或重试")

    def retry_check(self):
        """重试依赖检查"""
        self.progress_bar.setRange(0, 0)
        self.retry_button.setVisible(False)
        self.progress_label.setText("正在重试依赖检查...")

        self.checker = DependencyChecker()
        self.checker.progress.connect(self.update_progress)
        self.checker.finished.connect(self.on_check_finished)
        self.checker.start()


class FmAUI(QMainWindow):
    """FmA 文件合并助手 UI - 依赖自动检查版"""

    def __init__(self):
        super().__init__()
        self.init_ui()
        self.setWindowTitle("FmA 文件合并助手")
        self.setGeometry(100, 100, 800, 700)
        self.setMinimumSize(700, 600)
        self.setWindowIcon(QIcon(self.create_icon()))

    def create_icon(self):
        """创建应用图标（简约风格）"""
        # 使用base64编码一个简约文件合并图标
        icon_data = base64.b64decode("""
        AAABAAEAEBAAAAEAIABoBAAAFgAAACgAAAAQAAAAIAAAAAEAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
        AAAAAADg4OD//////////////////////7u7u0dHRwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
        AAAAAAAAAAAAAAAAALu7u0dHR0dHRzs7OwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEdHR0dH
        Rzs7Ozs7Ozs7OwAAAAAAAAAAAAAAAAAAAAAAAAAAAEdHR0dHRzs7Ozs7Ozs7Ozs7OwAAAAAAAAAA
        AAAAAAAAAAAAR0dHR0dHOzs7Ozs7Ozs7Ozs7Ozs7OwAAAAAAAAAAAAAAR0dHR0dHRzs7Ozs7Ozs7
        Ozs7Ozs7Ozs7OwAAAAAAAAAAR0dHR0dHRzs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7OwAAAAA7R0dHR0dH
        Rzs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7OwAAAEdHR0dHR0dHOzs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7
        Ozs7Ozs7OztHR0dHR0dHOzs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7R0dHR0dHRzs7Ozs7
        Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7O0dHR0dHR0dHOzs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7
        Ozs7Ozs7O0dHR0dHR0dHR0dHOzs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7R0dHR0dHR0dH
        R0dHOzs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7R0dHR0dHR0dHR0dHRzs7Ozs7Ozs7Ozs7
        Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7
        Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7
        Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7
        Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7
        Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7
        Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7
        Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7
        Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7
        Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7
        Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7
        Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7
        Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7
        Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7
        Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7
        Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7
        O/==
        """)
        pixmap = QPixmap()
        pixmap.loadFromData(icon_data)
        return pixmap

    def init_ui(self):
        """初始化UI界面"""
        # 设置主窗口样式
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f8f9fa;
                font-family: 'Segoe UI', 'Microsoft YaHei', sans-serif;
            }
        """)

        # 主窗口部件
        main_widget = QWidget()
        self.setCentralWidget(main_widget)

        # 主布局
        main_layout = QVBoxLayout(main_widget)
        main_layout.setContentsMargins(20, 15, 20, 15)
        main_layout.setSpacing(15)

        # 标题区域
        title_layout = self.create_title_layout()
        main_layout.addLayout(title_layout)

        # 输入/输出区域
        input_output_group = self.create_input_output_group()
        main_layout.addWidget(input_output_group)

        # 选项区域
        options_group = self.create_options_group()
        main_layout.addWidget(options_group)

        # 进度区域
        progress_group = self.create_progress_group()
        main_layout.addWidget(progress_group)

        # 日志区域
        log_group = self.create_log_group()
        main_layout.addWidget(log_group, 2)  # 给日志区域更多空间

        # 按钮区域
        button_layout = self.create_button_layout()
        main_layout.addLayout(button_layout)

        # 初始状态禁用按钮，直到依赖检查完成
        self.merge_btn.setEnabled(False)

    def create_title_layout(self):
        """创建标题区域"""
        title_layout = QHBoxLayout()
        title_layout.setContentsMargins(0, 0, 0, 10)

        # 标题
        title_label = QLabel("文件合并助手")
        title_label.setStyleSheet("""
            QLabel {
                font-size: 24px;
                font-weight: 500;
                color: #2c3e50;
            }
        """)

        # 版本信息
        version_label = QLabel("v1.0")
        version_label.setStyleSheet("""
            QLabel {
                font-size: 14px;
                color: #7f8c8d;
                padding-top: 8px;
            }
        """)

        title_layout.addWidget(title_label)
        title_layout.addStretch()
        title_layout.addWidget(version_label)

        return title_layout

    def create_input_output_group(self):
        """创建输入输出分组"""
        group = QGroupBox("文件路径")
        group.setStyleSheet("""
            QGroupBox {
                font-size: 14px;
                font-weight: 500;
                color: #34495e;
                border: 1px solid #e0e0e0;
                border-radius: 6px;
                padding-top: 20px;
                margin-top: 5px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
            }
        """)

        layout = QVBoxLayout()
        layout.setSpacing(12)
        layout.setContentsMargins(15, 15, 15, 15)

        # 输入路径
        input_layout = QHBoxLayout()
        input_layout.setSpacing(10)

        self.input_path = QLineEdit()
        self.input_path.setPlaceholderText("选择输入文件或文件夹...")
        self.input_path.setStyleSheet("""
            QLineEdit {
                padding: 8px;
                border: 1px solid #dcdee2;
                border-radius: 4px;
                font-size: 14px;
            }
            QLineEdit:focus {
                border-color: #3498db;
            }
        """)

        input_btn = QPushButton("选择")
        input_btn.setStyleSheet("""
            QPushButton {
                background-color: #ecf0f1;
                color: #34495e;
                border: none;
                border-radius: 4px;
                padding: 8px 15px;
                font-size: 14px;
                min-width: 70px;
            }
            QPushButton:hover {
                background-color: #d0d3d4;
            }
        """)
        input_btn.clicked.connect(self.select_input)

        input_layout.addWidget(self.input_path)
        input_layout.addWidget(input_btn)

        # 输出路径
        output_layout = QHBoxLayout()
        output_layout.setSpacing(10)

        self.output_path = QLineEdit()
        self.output_path.setPlaceholderText("设置输出文件路径...")
        self.output_path.setStyleSheet(self.input_path.styleSheet())

        output_btn = QPushButton("选择")
        output_btn.setStyleSheet(input_btn.styleSheet())
        output_btn.clicked.connect(self.select_output)

        output_layout.addWidget(self.output_path)
        output_layout.addWidget(output_btn)

        layout.addLayout(input_layout)
        layout.addLayout(output_layout)
        group.setLayout(layout)

        return group

    def create_options_group(self):
        """创建选项分组"""
        group = QGroupBox("合并选项")
        group.setStyleSheet(self.create_input_output_group().styleSheet())
        group.setMaximumHeight(120)

        layout = QVBoxLayout()
        layout.setSpacing(15)
        layout.setContentsMargins(15, 20, 15, 15)

        # 第一行选项
        options_row1 = QHBoxLayout()
        options_row1.setSpacing(20)

        # 输出格式
        format_layout = QVBoxLayout()
        format_label = QLabel("输出格式")
        format_label.setStyleSheet("font-size: 13px; color: #7f8c8d;")
        self.format_combo = QComboBox()
        self.format_combo.addItems(["Excel (.xlsx)", "Word (.docx)", "JSON (.json)", "Text (.txt)"])
        self.format_combo.setStyleSheet("""
            QComboBox {
                padding: 6px;
                border: 1px solid #dcdee2;
                border-radius: 4px;
                font-size: 14px;
            }
        """)

        format_layout.addWidget(format_label)
        format_layout.addWidget(self.format_combo)

        # 添加来源信息
        self.add_source_cb = QCheckBox("添加来源信息")
        self.add_source_cb.setChecked(True)
        self.add_source_cb.setStyleSheet("""
            QCheckBox {
                font-size: 14px;
                color: #34495e;
                padding: 4px;
            }
        """)

        # 包含子文件夹
        self.recursive_cb = QCheckBox("包含子文件夹")
        self.recursive_cb.setChecked(True)
        self.recursive_cb.setStyleSheet(self.add_source_cb.styleSheet())

        options_row1.addLayout(format_layout)
        options_row1.addWidget(self.add_source_cb)
        options_row1.addWidget(self.recursive_cb)

        layout.addLayout(options_row1)
        group.setLayout(layout)

        return group

    def create_progress_group(self):
        """创建进度分组"""
        group = QGroupBox()
        group.setStyleSheet(self.create_input_output_group().styleSheet())

        layout = QVBoxLayout()
        layout.setContentsMargins(15, 15, 15, 15)

        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 1px solid #dcdee2;
                border-radius: 4px;
                height: 24px;
                text-align: center;
                background: white;
            }
            QProgressBar::chunk {
                background-color: #3498db;
                border-radius: 4px;
            }
        """)

        # 进度状态
        self.progress_label = QLabel("准备就绪")
        self.progress_label.setStyleSheet("font-size: 13px; color: #7f8c8d; padding-top: 8px;")
        self.progress_label.setAlignment(Qt.AlignCenter)

        layout.addWidget(self.progress_bar)
        layout.addWidget(self.progress_label)
        group.setLayout(layout)

        return group

    def create_log_group(self):
        """创建日志分组（增加高度）"""
        group = QGroupBox("操作日志")
        group.setStyleSheet(self.create_input_output_group().styleSheet())

        layout = QVBoxLayout()
        layout.setContentsMargins(15, 15, 15, 15)

        # 日志文本框 - 增加高度
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setStyleSheet("""
            QTextEdit {
                border: 1px solid #dcdee2;
                border-radius: 4px;
                background-color: white;
                font-family: 'Consolas', 'Courier New', monospace;
                font-size: 13px;
                min-height: 150px;
            }
        """)

        # 清除日志按钮
        clear_btn = QPushButton("清除日志")
        clear_btn.setStyleSheet("""
            QPushButton {
                background-color: #ecf0f1;
                color: #34495e;
                border: none;
                border-radius: 4px;
                padding: 6px 12px;
                font-size: 13px;
                min-width: 80px;
            }
            QPushButton:hover {
                background-color: #d0d3d4;
            }
        """)
        clear_btn.clicked.connect(self.log_text.clear)

        layout.addWidget(self.log_text)
        layout.addWidget(clear_btn, 0, Qt.AlignRight)
        group.setLayout(layout)

        return group

    def create_button_layout(self):
        """创建按钮布局"""
        button_layout = QHBoxLayout()
        button_layout.setContentsMargins(10, 10, 10, 5)

        # 合并按钮
        self.merge_btn = QPushButton("开始合并")
        self.merge_btn.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                font-weight: 500;
                font-size: 15px;
                border: none;
                border-radius: 6px;
                padding: 12px 25px;
                min-width: 120px;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
            QPushButton:disabled {
                background-color: #bdc3c7;
            }
        """)
        self.merge_btn.setCursor(Qt.PointingHandCursor)
        self.merge_btn.clicked.connect(self.start_merge)

        # 退出按钮
        exit_btn = QPushButton("退出程序")
        exit_btn.setStyleSheet("""
            QPushButton {
                background-color: #ecf0f1;
                color: #34495e;
                font-weight: 500;
                font-size: 15px;
                border: none;
                border-radius: 6px;
                padding: 12px 25px;
                min-width: 120px;
            }
            QPushButton:hover {
                background-color: #d0d3d4;
            }
        """)
        exit_btn.setCursor(Qt.PointingHandCursor)
        exit_btn.clicked.connect(self.close)

        button_layout.addStretch()
        button_layout.addWidget(self.merge_btn)
        button_layout.addSpacing(15)
        button_layout.addWidget(exit_btn)
        button_layout.addStretch()

        return button_layout

    def select_input(self):
        """选择输入文件或文件夹"""
        path, _ = QFileDialog.getOpenFileName(
            self, "选择输入文件", "",
            "所有文件 (*);;Excel文件 (*.xlsx *.xls);;Word文件 (*.docx);;JSON文件 (*.json);;文本文件 (*.txt *.csv)"
        )

        if path:
            self.input_path.setText(path)

            # 自动设置输出路径
            if not self.output_path.text():
                output_dir = os.path.dirname(path)
                output_file = os.path.join(output_dir, "合并结果.xlsx")
                self.output_path.setText(output_file)

                # 更新进度状态
                self.progress_label.setText("已选择输入文件")

    def select_output(self):
        """选择输出文件路径"""
        path, _ = QFileDialog.getSaveFileName(
            self, "选择输出文件", "",
            "Excel文件 (*.xlsx);;Word文件 (*.docx);;JSON文件 (*.json);;文本文件 (*.txt)"
        )

        if path:
            self.output_path.setText(path)
            self.progress_label.setText("已设置输出路径")

    def start_merge(self):
        """开始合并操作"""
        # 确保依赖已安装
        if not self.check_dependencies():
            QMessageBox.warning(self, "依赖缺失", "请确保所有依赖已安装")
            return

        input_path = self.input_path.text()
        output_file = self.output_path.text()

        if not input_path or not output_file:
            QMessageBox.warning(self, "输入错误", "请选择输入和输出路径")
            return

        # 获取设置
        settings = {
            'add_source': self.add_source_cb.isChecked(),
            'recursive': self.recursive_cb.isChecked(),
            'combine_sheets': True,  # 默认只显示一个选项
            'output_format': self.get_output_format()
        }

        # 更新UI状态
        self.progress_bar.setValue(0)
        self.log_text.clear()
        self.progress_label.setText("开始处理文件...")
        self.merge_btn.setEnabled(False)

        # 记录开始时间
        self.log_text.append(f"[{datetime.now().strftime('%H:%M:%S')}] 开始合并操作")
        self.log_text.append(f"输入: {input_path}")
        self.log_text.append(f"输出: {output_file}")
        self.log_text.append("-" * 40)

        # 创建并启动合并线程
        self.merge_thread = MergeThread(
            self.fm, input_path, output_file, settings
        )
        self.merge_thread.log.connect(self.log_text.append)
        self.merge_thread.finished.connect(self.merge_finished)
        self.merge_thread.file_processed.connect(self.update_progress)
        self.merge_thread.start()

    def get_output_format(self):
        """获取选择的输出格式"""
        format_text = self.format_combo.currentText()
        if "Excel" in format_text:
            return "excel"
        elif "Word" in format_text:
            return "word"
        elif "JSON" in format_text:
            return "json"
        else:
            return "text"

    def update_progress(self, filename, count):
        """更新进度显示"""
        # 进度值是一个0-100的整数
        progress = min(count * 10, 100)  # 简单的进度计算，每处理一个文件增加10%
        self.progress_bar.setValue(progress)
        self.progress_label.setText(f"正在处理: {filename}")

    def merge_finished(self, success, stats):
        """合并完成处理"""
        # 恢复UI状态
        self.merge_btn.setEnabled(True)
        self.progress_label.setText("操作完成" if success else "操作失败")

        if success:
            self.log_text.append(f"[{datetime.now().strftime('%H:%M:%S')}] 合并成功!")
            self.log_text.append(f"处理时间: {stats['time']:.2f}秒")
            self.log_text.append(f"处理文件数: {stats['files']}")
            self.log_text.append(f"成功合并: {stats['success']}")
            self.progress_bar.setValue(100)

            # 显示成功消息
            QMessageBox.information(
                self, "合并成功",
                f"文件合并成功!\n\n"
                f"输出文件: {stats['output']}\n"
                f"处理时间: {stats['time']:.2f}秒\n"
                f"文件数量: {stats['files']}"
            )
        else:
            self.log_text.append(f"[{datetime.now().strftime('%H:%M:%S')}] 合并失败")
            self.progress_bar.setValue(0)

            # 显示错误消息
            error_msg = stats.get('error', '未知错误')
            QMessageBox.critical(
                self, "合并失败",
                f"文件合并失败!\n\n错误信息: {error_msg}"
            )

    def check_dependencies(self):
        """检查所有依赖是否已安装"""
        required = ['pandas', 'openpyxl', 'python-docx', 'chardet', 'psutil']
        missing = []

        for package in required:
            try:
                importlib.import_module(package)
            except ImportError:
                missing.append(package)

        if not missing:
            return True

        # 显示依赖安装对话框
        dialog = DependencyDialog(self)
        result = dialog.exec_()

        return result == QDialog.Accepted


class MergeThread(QThread):
    """后台合并线程"""
    progress = pyqtSignal(int)
    finished = pyqtSignal(bool, dict)
    log = pyqtSignal(str)
    file_processed = pyqtSignal(str, int)

    def __init__(self, fm, input_path, output_file, settings):
        super().__init__()
        self.fm = fm
        self.input_path = input_path
        self.output_file = output_file
        self.settings = settings

    def run(self):
        try:
            self.log.emit("开始文件合并...")
            success, stats = self.fm.merge_files(self.input_path, self.output_file, self.settings)

            if success:
                self.log.emit(f"合并成功! 输出文件: {self.output_file}")
                self.log.emit(f"处理时间: {stats['time']:.2f}秒 | 文件数: {stats['files']} | 成功: {stats['success']}")
            else:
                self.log.emit(f"合并失败: {stats}")

            self.finished.emit(success, stats)
        except Exception as e:
            self.log.emit(f"错误: {str(e)}")
            self.finished.emit(False, {'error': str(e)})


def main():
    """应用入口"""
    app = QApplication(sys.argv)

    # 设置应用样式
    app.setStyle("Fusion")

    # 创建主窗口
    window = FmAUI()
    window.show()

    # 启动依赖检查
    if not window.check_dependencies():
        QMessageBox.critical(None, "依赖缺失", "无法运行，请确保所有依赖已安装")
        return

    # 应用执行
    sys.exit(app.exec_())


if __name__ == "__main__":
    # 确保正确初始化
    import datetime
    import glob
    import json
    import importlib

    # 尝试导入核心模块
    try:
        import pandas as pd
        from docx import Document
        import psutil
    except ImportError:
        # 如果导入失败，将在依赖检查中处理
        pass

    main()
