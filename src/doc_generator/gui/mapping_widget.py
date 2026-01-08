"""Mapping configuration widget."""

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QTableWidget, QTableWidgetItem,
    QPushButton, QComboBox, QLineEdit, QLabel, QHeaderView, QMessageBox,
    QDialog, QDialogButtonBox, QTextEdit, QGroupBox
)
from PyQt6.QtCore import Qt, pyqtSignal

from ..core.mapping import MappingRule, MappingType, MappingConfig


class ExpressionEditorDialog(QDialog):
    """Dialog for editing expressions."""

    def __init__(self, expression: str, excel_columns: list[str], parent=None):
        super().__init__(parent)
        self.setWindowTitle("编辑表达式")
        self.setMinimumSize(500, 400)

        layout = QVBoxLayout(self)

        # Help text
        help_group = QGroupBox("表达式帮助")
        help_layout = QVBoxLayout(help_group)
        help_text = QLabel(
            "使用 {{列名}} 引用 Excel 列的值\n\n"
            "支持的运算符: +, -, *, /, ==, !=, <, >, <=, >=\n\n"
            "支持的函数:\n"
            "  concat(a, b, ...) - 连接字符串\n"
            "  sum(a, b, ...) - 求和\n"
            "  avg(a, b, ...) - 平均值\n"
            "  round(x, n) - 四舍五入到n位小数\n"
            "  if(条件, 真值, 假值) - 条件判断\n"
            "  ifempty(值, 默认值) - 空值替换\n"
            "  upper(s), lower(s) - 大小写转换\n"
            "  left(s, n), right(s, n) - 截取字符串\n\n"
            "示例: {{单价}} * {{数量}}\n"
            "示例: concat({{姓}}, {{名}})\n"
            "示例: round({{金额}} * 1.1, 2)"
        )
        help_text.setWordWrap(True)
        help_layout.addWidget(help_text)
        layout.addWidget(help_group)

        # Column buttons
        if excel_columns:
            col_group = QGroupBox("快速插入列")
            col_layout = QHBoxLayout(col_group)
            col_layout.setSpacing(5)

            for col in excel_columns[:10]:  # Show max 10 columns
                btn = QPushButton(col)
                btn.setMaximumWidth(100)
                btn.clicked.connect(lambda checked, c=col: self._insert_column(c))
                col_layout.addWidget(btn)

            col_layout.addStretch()
            layout.addWidget(col_group)

        # Expression editor
        expr_label = QLabel("表达式:")
        layout.addWidget(expr_label)

        self.expression_edit = QTextEdit()
        self.expression_edit.setPlainText(expression)
        self.expression_edit.setMinimumHeight(80)
        layout.addWidget(self.expression_edit)

        # Buttons
        button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

    def _insert_column(self, column: str) -> None:
        """Insert a column placeholder at cursor."""
        cursor = self.expression_edit.textCursor()
        cursor.insertText(f"{{{{{column}}}}}")
        self.expression_edit.setFocus()

    def get_expression(self) -> str:
        """Get the edited expression."""
        return self.expression_edit.toPlainText().strip()


class MappingWidget(QWidget):
    """Widget for configuring placeholder mappings."""

    mappings_changed = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self._excel_columns: list[str] = []
        self._word_placeholders: list[str] = []
        self._setup_ui()

    def _setup_ui(self) -> None:
        """Set up the user interface."""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        # Toolbar
        toolbar = QHBoxLayout()

        self.auto_map_btn = QPushButton("自动映射")
        self.auto_map_btn.setToolTip("自动匹配同名的列和占位符")
        self.auto_map_btn.clicked.connect(self._auto_map)
        toolbar.addWidget(self.auto_map_btn)

        self.clear_btn = QPushButton("清除全部")
        self.clear_btn.clicked.connect(self._clear_all)
        toolbar.addWidget(self.clear_btn)

        toolbar.addStretch()
        layout.addLayout(toolbar)

        # Mapping table
        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["Word占位符", "映射类型", "来源/表达式", "操作"])

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)

        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setAlternatingRowColors(True)

        layout.addWidget(self.table)

    def set_data(self, excel_columns: list[str], word_placeholders: list[str]) -> None:
        """Set the available columns and placeholders.

        Args:
            excel_columns: List of column names from Excel.
            word_placeholders: List of placeholder names from Word template.
        """
        self._excel_columns = excel_columns
        self._word_placeholders = word_placeholders
        self._rebuild_table()

    def _rebuild_table(self) -> None:
        """Rebuild the mapping table."""
        self.table.setRowCount(len(self._word_placeholders))

        for row, placeholder in enumerate(self._word_placeholders):
            # Placeholder name (read-only)
            item = QTableWidgetItem(placeholder)
            item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.table.setItem(row, 0, item)

            # Mapping type combo
            type_combo = QComboBox()
            type_combo.addItems(["直接映射", "表达式"])
            type_combo.currentIndexChanged.connect(
                lambda idx, r=row: self._on_type_changed(r, idx)
            )
            self.table.setCellWidget(row, 1, type_combo)

            # Source combo / expression field
            source_combo = QComboBox()
            source_combo.setEditable(False)
            source_combo.addItem("(未映射)")
            source_combo.addItems(self._excel_columns)

            # Auto-select matching column
            if placeholder in self._excel_columns:
                idx = self._excel_columns.index(placeholder) + 1
                source_combo.setCurrentIndex(idx)

            source_combo.currentIndexChanged.connect(lambda: self.mappings_changed.emit())
            self.table.setCellWidget(row, 2, source_combo)

            # Edit button (for expressions)
            edit_btn = QPushButton("编辑...")
            edit_btn.setVisible(False)
            edit_btn.clicked.connect(lambda checked, r=row: self._edit_expression(r))
            self.table.setCellWidget(row, 3, edit_btn)

    def _on_type_changed(self, row: int, type_index: int) -> None:
        """Handle mapping type change."""
        source_widget = self.table.cellWidget(row, 2)
        edit_btn = self.table.cellWidget(row, 3)

        if type_index == 0:  # Direct mapping
            # Replace with combo box
            if isinstance(source_widget, QLineEdit):
                combo = QComboBox()
                combo.addItem("(未映射)")
                combo.addItems(self._excel_columns)
                combo.currentIndexChanged.connect(lambda: self.mappings_changed.emit())
                self.table.setCellWidget(row, 2, combo)
            edit_btn.setVisible(False)
        else:  # Expression
            # Replace with line edit
            if isinstance(source_widget, QComboBox):
                edit = QLineEdit()
                edit.setReadOnly(True)
                edit.setPlaceholderText("点击编辑按钮设置表达式...")
                edit.textChanged.connect(lambda: self.mappings_changed.emit())
                self.table.setCellWidget(row, 2, edit)
            edit_btn.setVisible(True)

        self.mappings_changed.emit()

    def _edit_expression(self, row: int) -> None:
        """Open expression editor dialog."""
        source_widget = self.table.cellWidget(row, 2)
        current_expr = ""
        if isinstance(source_widget, QLineEdit):
            current_expr = source_widget.text()

        dialog = ExpressionEditorDialog(current_expr, self._excel_columns, self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            source_widget.setText(dialog.get_expression())

    def _auto_map(self) -> None:
        """Automatically map matching column names to placeholders."""
        for row in range(self.table.rowCount()):
            placeholder = self.table.item(row, 0).text()
            source_widget = self.table.cellWidget(row, 2)

            if isinstance(source_widget, QComboBox):
                if placeholder in self._excel_columns:
                    idx = self._excel_columns.index(placeholder) + 1
                    source_widget.setCurrentIndex(idx)

        self.mappings_changed.emit()

    def _clear_all(self) -> None:
        """Clear all mappings."""
        for row in range(self.table.rowCount()):
            # Reset to direct mapping
            type_combo = self.table.cellWidget(row, 1)
            if isinstance(type_combo, QComboBox):
                type_combo.setCurrentIndex(0)

            source_widget = self.table.cellWidget(row, 2)
            if isinstance(source_widget, QComboBox):
                source_widget.setCurrentIndex(0)
            elif isinstance(source_widget, QLineEdit):
                source_widget.clear()

        self.mappings_changed.emit()

    def get_mapping_config(self) -> MappingConfig:
        """Get the current mapping configuration.

        Returns:
            MappingConfig with all the rules.
        """
        config = MappingConfig()

        for row in range(self.table.rowCount()):
            placeholder = self.table.item(row, 0).text()
            type_combo = self.table.cellWidget(row, 1)
            source_widget = self.table.cellWidget(row, 2)

            if type_combo.currentIndex() == 0:  # Direct mapping
                if isinstance(source_widget, QComboBox) and source_widget.currentIndex() > 0:
                    source = source_widget.currentText()
                    config.add_rule(MappingRule(
                        placeholder=placeholder,
                        mapping_type=MappingType.DIRECT,
                        source=source,
                    ))
            else:  # Expression
                if isinstance(source_widget, QLineEdit) and source_widget.text().strip():
                    config.add_rule(MappingRule(
                        placeholder=placeholder,
                        mapping_type=MappingType.EXPRESSION,
                        expression=source_widget.text().strip(),
                    ))

        return config

    def load_mapping_config(self, config: MappingConfig) -> None:
        """Load mappings from a configuration.

        Args:
            config: MappingConfig to load.
        """
        for row in range(self.table.rowCount()):
            placeholder = self.table.item(row, 0).text()
            rule = config.get_rule(placeholder)

            if rule is None:
                continue

            type_combo = self.table.cellWidget(row, 1)

            if rule.mapping_type == MappingType.DIRECT:
                type_combo.setCurrentIndex(0)
                # Wait for widget to update
                source_widget = self.table.cellWidget(row, 2)
                if isinstance(source_widget, QComboBox) and rule.source in self._excel_columns:
                    idx = self._excel_columns.index(rule.source) + 1
                    source_widget.setCurrentIndex(idx)
            else:
                type_combo.setCurrentIndex(1)
                # Trigger type change
                self._on_type_changed(row, 1)
                source_widget = self.table.cellWidget(row, 2)
                if isinstance(source_widget, QLineEdit):
                    source_widget.setText(rule.expression)

        self.mappings_changed.emit()
