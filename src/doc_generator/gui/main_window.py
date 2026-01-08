"""Main application window."""

import os
from pathlib import Path
from typing import Optional

from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGroupBox,
    QPushButton, QLineEdit, QLabel, QFileDialog, QProgressBar,
    QMessageBox, QComboBox, QSpinBox, QStatusBar, QMenuBar,
    QMenu, QSplitter, QTextEdit, QApplication
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QAction

from ..core.excel_reader import ExcelReader
from ..core.word_renderer import WordRenderer
from ..core.mapping import MappingConfig
from ..utils.config import get_app_config
from .mapping_widget import MappingWidget


class GeneratorWorker(QThread):
    """Worker thread for document generation."""

    progress = pyqtSignal(int, int, str)  # current, total, message
    finished = pyqtSignal(bool, str)  # success, message
    error = pyqtSignal(str)

    def __init__(
        self,
        excel_path: str,
        template_path: str,
        output_dir: str,
        mapping_config: MappingConfig,
        filename_pattern: str,
        sheet_name: str,
        header_row: int,
        start_row: int,
    ):
        super().__init__()
        self.excel_path = excel_path
        self.template_path = template_path
        self.output_dir = output_dir
        self.mapping_config = mapping_config
        self.filename_pattern = filename_pattern
        self.sheet_name = sheet_name
        self.header_row = header_row
        self.start_row = start_row
        self._cancelled = False

    def cancel(self) -> None:
        """Cancel the generation."""
        self._cancelled = True

    def run(self) -> None:
        """Run the document generation."""
        try:
            renderer = WordRenderer(self.template_path)
            mappings = self.mapping_config.get_mappings_dict()

            with ExcelReader(self.excel_path) as reader:
                sheet = self.sheet_name if self.sheet_name else None
                rows = reader.get_all_rows(sheet, self.header_row, self.start_row)
                total = len(rows)

                if total == 0:
                    self.finished.emit(False, "Excel文件中没有数据行")
                    return

                generated = 0
                for i, row_data in enumerate(rows):
                    if self._cancelled:
                        self.finished.emit(False, f"已取消，已生成 {generated} 个文件")
                        return

                    # Add row index to data
                    row_data["_index"] = i + 1
                    row_data["_row"] = i + self.start_row

                    # Generate filename
                    filename = renderer.generate_filename(self.filename_pattern, row_data, i)
                    output_path = Path(self.output_dir) / filename

                    # Render document
                    self.progress.emit(i + 1, total, f"正在生成: {filename}")
                    renderer.render(row_data, mappings, output_path)
                    generated += 1

                self.finished.emit(True, f"成功生成 {generated} 个文件")

        except Exception as e:
            self.error.emit(str(e))


class MainWindow(QMainWindow):
    """Main application window."""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Doc Generator - Excel 批量生成 Word")
        self.setMinimumSize(900, 700)

        self._excel_columns: list[str] = []
        self._word_placeholders: list[str] = []
        self._worker: Optional[GeneratorWorker] = None

        self._setup_ui()
        self._setup_menu()
        self._load_config()

    def _setup_ui(self) -> None:
        """Set up the user interface."""
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        # File selection area
        files_group = QGroupBox("文件选择")
        files_layout = QVBoxLayout(files_group)

        # Excel file
        excel_row = QHBoxLayout()
        excel_row.addWidget(QLabel("Excel文件:"))
        self.excel_path_edit = QLineEdit()
        self.excel_path_edit.setReadOnly(True)
        excel_row.addWidget(self.excel_path_edit)
        self.excel_browse_btn = QPushButton("浏览...")
        self.excel_browse_btn.clicked.connect(self._browse_excel)
        excel_row.addWidget(self.excel_browse_btn)
        files_layout.addLayout(excel_row)

        # Excel options
        excel_opts = QHBoxLayout()
        excel_opts.addWidget(QLabel("工作表:"))
        self.sheet_combo = QComboBox()
        self.sheet_combo.setMinimumWidth(150)
        self.sheet_combo.currentTextChanged.connect(self._on_sheet_changed)
        excel_opts.addWidget(self.sheet_combo)
        excel_opts.addSpacing(20)
        excel_opts.addWidget(QLabel("表头行:"))
        self.header_row_spin = QSpinBox()
        self.header_row_spin.setRange(1, 100)
        self.header_row_spin.setValue(1)
        self.header_row_spin.valueChanged.connect(self._reload_excel)
        excel_opts.addWidget(self.header_row_spin)
        excel_opts.addSpacing(20)
        excel_opts.addWidget(QLabel("数据起始行:"))
        self.start_row_spin = QSpinBox()
        self.start_row_spin.setRange(1, 1000)
        self.start_row_spin.setValue(2)
        excel_opts.addWidget(self.start_row_spin)
        excel_opts.addStretch()
        files_layout.addLayout(excel_opts)

        # Word template
        word_row = QHBoxLayout()
        word_row.addWidget(QLabel("Word模板:"))
        self.template_path_edit = QLineEdit()
        self.template_path_edit.setReadOnly(True)
        word_row.addWidget(self.template_path_edit)
        self.template_browse_btn = QPushButton("浏览...")
        self.template_browse_btn.clicked.connect(self._browse_template)
        word_row.addWidget(self.template_browse_btn)
        files_layout.addLayout(word_row)

        # Output directory
        output_row = QHBoxLayout()
        output_row.addWidget(QLabel("输出目录:"))
        self.output_path_edit = QLineEdit()
        output_row.addWidget(self.output_path_edit)
        self.output_browse_btn = QPushButton("浏览...")
        self.output_browse_btn.clicked.connect(self._browse_output)
        output_row.addWidget(self.output_browse_btn)
        files_layout.addLayout(output_row)

        # Filename pattern
        filename_row = QHBoxLayout()
        filename_row.addWidget(QLabel("文件名模式:"))
        self.filename_pattern_edit = QLineEdit()
        self.filename_pattern_edit.setText("{{_index}}_output.docx")
        self.filename_pattern_edit.setPlaceholderText("可使用 {{列名}} 作为占位符，如: {{姓名}}_合同.docx")
        filename_row.addWidget(self.filename_pattern_edit)
        files_layout.addLayout(filename_row)

        layout.addWidget(files_group)

        # Splitter for mapping and preview
        splitter = QSplitter(Qt.Orientation.Vertical)

        # Mapping configuration
        mapping_group = QGroupBox("映射配置")
        mapping_layout = QVBoxLayout(mapping_group)
        self.mapping_widget = MappingWidget()
        mapping_layout.addWidget(self.mapping_widget)
        splitter.addWidget(mapping_group)

        # Info area
        info_group = QGroupBox("信息")
        info_layout = QVBoxLayout(info_group)
        self.info_text = QTextEdit()
        self.info_text.setReadOnly(True)
        self.info_text.setMaximumHeight(100)
        info_layout.addWidget(self.info_text)
        splitter.addWidget(info_group)

        splitter.setSizes([400, 100])
        layout.addWidget(splitter)

        # Progress and actions
        action_layout = QHBoxLayout()

        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        action_layout.addWidget(self.progress_bar)

        self.generate_btn = QPushButton("生成文档")
        self.generate_btn.setMinimumWidth(120)
        self.generate_btn.setStyleSheet("QPushButton { padding: 8px 16px; font-weight: bold; }")
        self.generate_btn.clicked.connect(self._generate)
        action_layout.addWidget(self.generate_btn)

        self.cancel_btn = QPushButton("取消")
        self.cancel_btn.setVisible(False)
        self.cancel_btn.clicked.connect(self._cancel_generation)
        action_layout.addWidget(self.cancel_btn)

        layout.addLayout(action_layout)

        # Status bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("就绪")

    def _setup_menu(self) -> None:
        """Set up the menu bar."""
        menubar = self.menuBar()

        # File menu
        file_menu = menubar.addMenu("文件(&F)")

        save_config_action = QAction("保存配置(&S)", self)
        save_config_action.triggered.connect(self._save_config)
        file_menu.addAction(save_config_action)

        load_config_action = QAction("加载配置(&L)", self)
        load_config_action.triggered.connect(self._load_config_file)
        file_menu.addAction(load_config_action)

        file_menu.addSeparator()

        exit_action = QAction("退出(&X)", self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        # Help menu
        help_menu = menubar.addMenu("帮助(&H)")

        about_action = QAction("关于(&A)", self)
        about_action.triggered.connect(self._show_about)
        help_menu.addAction(about_action)

    def _load_config(self) -> None:
        """Load application config."""
        config = get_app_config()
        # Could restore recent files, window geometry, etc.

    def _browse_excel(self) -> None:
        """Browse for Excel file."""
        path, _ = QFileDialog.getOpenFileName(
            self, "选择Excel文件", "",
            "Excel文件 (*.xlsx *.xls);;所有文件 (*)"
        )
        if path:
            self.excel_path_edit.setText(path)
            self._load_excel(path)

    def _load_excel(self, path: str) -> None:
        """Load Excel file and update UI."""
        try:
            with ExcelReader(path) as reader:
                # Update sheet combo
                self.sheet_combo.clear()
                self.sheet_combo.addItems(reader.sheet_names)

                # Load columns from first sheet
                self._excel_columns = reader.get_headers(
                    header_row=self.header_row_spin.value()
                )

                row_count = reader.get_row_count(start_row=self.start_row_spin.value())

            self._update_mapping()
            self._update_info()
            self.status_bar.showMessage(f"已加载Excel: {len(self._excel_columns)}列, {row_count}行数据")

            # Save to recent
            config = get_app_config()
            config.add_recent_file("excel", path)
            config.save()

        except Exception as e:
            QMessageBox.critical(self, "错误", f"无法加载Excel文件:\n{e}")

    def _on_sheet_changed(self, sheet_name: str) -> None:
        """Handle sheet selection change."""
        if sheet_name and self.excel_path_edit.text():
            self._reload_excel()

    def _reload_excel(self) -> None:
        """Reload Excel with current settings."""
        path = self.excel_path_edit.text()
        if path:
            try:
                with ExcelReader(path) as reader:
                    sheet = self.sheet_combo.currentText() or None
                    self._excel_columns = reader.get_headers(
                        sheet_name=sheet,
                        header_row=self.header_row_spin.value()
                    )
                self._update_mapping()
                self._update_info()
            except Exception as e:
                self.info_text.setText(f"重新加载Excel出错: {e}")

    def _browse_template(self) -> None:
        """Browse for Word template."""
        path, _ = QFileDialog.getOpenFileName(
            self, "选择Word模板", "",
            "Word文档 (*.docx);;所有文件 (*)"
        )
        if path:
            self.template_path_edit.setText(path)
            self._load_template(path)

    def _load_template(self, path: str) -> None:
        """Load Word template and extract placeholders."""
        try:
            renderer = WordRenderer(path)
            self._word_placeholders = renderer.get_placeholders()
            self._update_mapping()
            self._update_info()
            self.status_bar.showMessage(f"已加载模板: {len(self._word_placeholders)}个占位符")

            # Save to recent
            config = get_app_config()
            config.add_recent_file("template", path)
            config.save()

        except Exception as e:
            QMessageBox.critical(self, "错误", f"无法加载Word模板:\n{e}")

    def _browse_output(self) -> None:
        """Browse for output directory."""
        path = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if path:
            self.output_path_edit.setText(path)

    def _update_mapping(self) -> None:
        """Update the mapping widget with current columns and placeholders."""
        self.mapping_widget.set_data(self._excel_columns, self._word_placeholders)

    def _update_info(self) -> None:
        """Update the info text area."""
        info_lines = []

        if self._excel_columns:
            info_lines.append(f"Excel列 ({len(self._excel_columns)}): {', '.join(self._excel_columns[:10])}")
            if len(self._excel_columns) > 10:
                info_lines[-1] += f" ... 等{len(self._excel_columns)}列"

        if self._word_placeholders:
            info_lines.append(f"Word占位符 ({len(self._word_placeholders)}): {', '.join(self._word_placeholders[:10])}")
            if len(self._word_placeholders) > 10:
                info_lines[-1] += f" ... 等{len(self._word_placeholders)}个"

        self.info_text.setText("\n".join(info_lines))

    def _save_config(self) -> None:
        """Save current configuration to file."""
        path, _ = QFileDialog.getSaveFileName(
            self, "保存配置", "",
            "配置文件 (*.json);;所有文件 (*)"
        )
        if path:
            try:
                config = self.mapping_widget.get_mapping_config()
                config.excel_file = self.excel_path_edit.text()
                config.template_file = self.template_path_edit.text()
                config.output_directory = self.output_path_edit.text()
                config.output_filename_pattern = self.filename_pattern_edit.text()
                config.sheet_name = self.sheet_combo.currentText()
                config.header_row = self.header_row_spin.value()
                config.start_row = self.start_row_spin.value()
                config.save(path)
                self.status_bar.showMessage(f"配置已保存: {path}")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"保存配置失败:\n{e}")

    def _load_config_file(self) -> None:
        """Load configuration from file."""
        path, _ = QFileDialog.getOpenFileName(
            self, "加载配置", "",
            "配置文件 (*.json);;所有文件 (*)"
        )
        if path:
            try:
                config = MappingConfig.load(path)

                # Load files
                if config.excel_file and os.path.exists(config.excel_file):
                    self.excel_path_edit.setText(config.excel_file)
                    self._load_excel(config.excel_file)

                if config.template_file and os.path.exists(config.template_file):
                    self.template_path_edit.setText(config.template_file)
                    self._load_template(config.template_file)

                self.output_path_edit.setText(config.output_directory)
                self.filename_pattern_edit.setText(config.output_filename_pattern)

                if config.sheet_name:
                    idx = self.sheet_combo.findText(config.sheet_name)
                    if idx >= 0:
                        self.sheet_combo.setCurrentIndex(idx)

                self.header_row_spin.setValue(config.header_row)
                self.start_row_spin.setValue(config.start_row)

                # Load mappings
                self.mapping_widget.load_mapping_config(config)

                self.status_bar.showMessage(f"配置已加载: {path}")

            except Exception as e:
                QMessageBox.critical(self, "错误", f"加载配置失败:\n{e}")

    def _validate(self) -> bool:
        """Validate inputs before generation."""
        if not self.excel_path_edit.text():
            QMessageBox.warning(self, "提示", "请选择Excel文件")
            return False

        if not self.template_path_edit.text():
            QMessageBox.warning(self, "提示", "请选择Word模板")
            return False

        if not self.output_path_edit.text():
            QMessageBox.warning(self, "提示", "请选择输出目录")
            return False

        if not self._word_placeholders:
            QMessageBox.warning(self, "提示", "Word模板中没有找到占位符")
            return False

        return True

    def _generate(self) -> None:
        """Start document generation."""
        if not self._validate():
            return

        mapping_config = self.mapping_widget.get_mapping_config()

        self._worker = GeneratorWorker(
            excel_path=self.excel_path_edit.text(),
            template_path=self.template_path_edit.text(),
            output_dir=self.output_path_edit.text(),
            mapping_config=mapping_config,
            filename_pattern=self.filename_pattern_edit.text(),
            sheet_name=self.sheet_combo.currentText(),
            header_row=self.header_row_spin.value(),
            start_row=self.start_row_spin.value(),
        )

        self._worker.progress.connect(self._on_progress)
        self._worker.finished.connect(self._on_finished)
        self._worker.error.connect(self._on_error)

        self.generate_btn.setEnabled(False)
        self.cancel_btn.setVisible(True)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)

        self._worker.start()

    def _cancel_generation(self) -> None:
        """Cancel the current generation."""
        if self._worker:
            self._worker.cancel()

    def _on_progress(self, current: int, total: int, message: str) -> None:
        """Handle progress updates."""
        self.progress_bar.setMaximum(total)
        self.progress_bar.setValue(current)
        self.status_bar.showMessage(message)

    def _on_finished(self, success: bool, message: str) -> None:
        """Handle generation finished."""
        self.generate_btn.setEnabled(True)
        self.cancel_btn.setVisible(False)
        self.progress_bar.setVisible(False)

        if success:
            self.status_bar.showMessage(message)
            QMessageBox.information(self, "完成", message)
        else:
            self.status_bar.showMessage(message)
            QMessageBox.warning(self, "提示", message)

        self._worker = None

    def _on_error(self, error: str) -> None:
        """Handle generation error."""
        self.generate_btn.setEnabled(True)
        self.cancel_btn.setVisible(False)
        self.progress_bar.setVisible(False)

        self.status_bar.showMessage(f"错误: {error}")
        QMessageBox.critical(self, "错误", f"生成文档时出错:\n{error}")

        self._worker = None

    def _show_about(self) -> None:
        """Show about dialog."""
        QMessageBox.about(
            self, "关于 Doc Generator",
            "Doc Generator v1.0.0\n\n"
            "批量将Excel数据填充到Word模板\n\n"
            "支持:\n"
            "- Excel列到Word占位符映射\n"
            "- 表达式计算\n"
            "- 批量生成文档"
        )

    def closeEvent(self, event) -> None:
        """Handle window close."""
        if self._worker and self._worker.isRunning():
            reply = QMessageBox.question(
                self, "确认",
                "正在生成文档，确定要退出吗？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.No:
                event.ignore()
                return

            self._worker.cancel()
            self._worker.wait()

        # Save window state
        config = get_app_config()
        config.save()

        event.accept()
