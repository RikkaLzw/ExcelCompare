# -*- coding: utf-8 -*-
"""
比较配置面板

提供比较模式、区域选择、智能匹配等配置。
"""
from typing import List, Optional
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QComboBox, QCheckBox, QFrame, QListWidget, QListWidgetItem,
    QGroupBox, QLineEdit, QScrollArea
)
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QCursor

from src.services.compare_service import CompareMode, CompareOptions


class ConfigPanel(QFrame):
    """配置面板"""
    
    compare_clicked = pyqtSignal()
    smart_compare_clicked = pyqtSignal()
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self._setup_ui()
        self._apply_styles()
    
    def _setup_ui(self):
        """设置 UI"""
        # 外层布局
        outer_layout = QVBoxLayout(self)
        outer_layout.setContentsMargins(0, 0, 0, 0)
        outer_layout.setSpacing(0)
        
        # 滚动区域
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        
        # 内容容器
        content = QWidget()
        layout = QVBoxLayout(content)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)
        
        # 标题
        title = QLabel("比较配置")
        title.setObjectName("panelTitle")
        layout.addWidget(title)
        
        # 比较模式
        mode_group = QGroupBox("比较模式")
        mode_layout = QVBoxLayout(mode_group)
        
        self.mode_combo = QComboBox()
        self.mode_combo.addItem("精确匹配", CompareMode.EXACT)
        self.mode_combo.addItem("数值比较", CompareMode.NUMERIC)
        self.mode_combo.addItem("结构比较", CompareMode.STRUCTURE)
        self.mode_combo.addItem("公式比较", CompareMode.FORMULA)
        self.mode_combo.addItem("智能匹配", "SMART")
        self.mode_combo.currentIndexChanged.connect(self._on_mode_changed)
        mode_layout.addWidget(self.mode_combo)
        
        layout.addWidget(mode_group)
        
        # 智能匹配选项（默认隐藏）
        self.smart_group = QWidget()
        self.smart_group.setObjectName("smartWidget")
        self.smart_group.setMinimumHeight(180)  # 设置最小高度
        smart_layout = QVBoxLayout(self.smart_group)
        smart_layout.setContentsMargins(10, 10, 10, 10)
        smart_layout.setSpacing(8)
        
        # 智能匹配标题
        smart_title = QLabel("-- 智能匹配设置 --")
        smart_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        smart_layout.addWidget(smart_title)
        
        # 区域选择
        range_lbl = QLabel("比较区域 (如 A1:D10):")
        smart_layout.addWidget(range_lbl)
        
        self.range_input = QLineEdit()
        self.range_input.setPlaceholderText("留空比较全表")
        smart_layout.addWidget(self.range_input)
        
        # 标题行
        self.use_header_check = QCheckBox("首行作为列标题")
        self.use_header_check.setChecked(True)
        smart_layout.addWidget(self.use_header_check)
        
        # 主键列
        self.use_key_check = QCheckBox("使用主键列匹配行")
        self.use_key_check.stateChanged.connect(self._on_key_check_changed)
        smart_layout.addWidget(self.use_key_check)
        
        # 主键列输入
        key_widget = QWidget()
        key_layout = QHBoxLayout(key_widget)
        key_layout.setContentsMargins(20, 0, 0, 0)
        key_layout.addWidget(QLabel("主键列:"))
        self.key_col_input = QLineEdit()
        self.key_col_input.setPlaceholderText("A")
        self.key_col_input.setMaximumWidth(50)
        self.key_col_input.setEnabled(False)
        key_layout.addWidget(self.key_col_input)
        key_layout.addStretch()
        smart_layout.addWidget(key_widget)
        
        self.smart_group.hide()
        layout.addWidget(self.smart_group)
        
        # 工作表选择
        sheet_group = QGroupBox("工作表")
        sheet_layout = QVBoxLayout(sheet_group)
        
        self.all_sheets_check = QCheckBox("比较全部工作表")
        self.all_sheets_check.setChecked(True)
        self.all_sheets_check.stateChanged.connect(self._on_all_sheets_changed)
        sheet_layout.addWidget(self.all_sheets_check)
        
        self.sheet_list = QListWidget()
        self.sheet_list.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        self.sheet_list.setMaximumHeight(80)
        self.sheet_list.setEnabled(False)
        sheet_layout.addWidget(self.sheet_list)
        
        layout.addWidget(sheet_group)
        
        # 忽略选项
        ignore_group = QGroupBox("忽略选项")
        ignore_layout = QVBoxLayout(ignore_group)
        
        self.ignore_format_check = QCheckBox("忽略格式差异")
        self.ignore_format_check.setChecked(True)
        ignore_layout.addWidget(self.ignore_format_check)
        
        self.ignore_case_check = QCheckBox("忽略大小写")
        ignore_layout.addWidget(self.ignore_case_check)
        
        self.ignore_whitespace_check = QCheckBox("忽略前后空格")
        ignore_layout.addWidget(self.ignore_whitespace_check)
        
        self.ignore_empty_rows_check = QCheckBox("忽略空白行")
        ignore_layout.addWidget(self.ignore_empty_rows_check)
        
        layout.addWidget(ignore_group)
        
        # 匹配方式选项（适用于所有比较模式）
        match_group = QGroupBox()
        match_main_layout = QVBoxLayout(match_group)

        # 标题行（包含问号提示）
        title_layout = QHBoxLayout()
        title_layout.setContentsMargins(0, 0, 0, 5)
        title_label = QLabel("匹配方式")
        title_label.setStyleSheet("font-weight: bold; font-size: 13px;")
        title_layout.addWidget(title_label)

        help_label = QLabel("?")
        help_label.setStyleSheet("""
            QLabel {
                color: #666;
                background-color: #e8e8e8;
                border: 1px solid #ccc;
                border-radius: 8px;
                font-size: 11px;
                font-weight: bold;
                padding: 0px;
                min-width: 16px;
                max-width: 16px;
                min-height: 16px;
                max-height: 16px;
            }
            QLabel:hover {
                background-color: #d0d0d0;
                color: #333;
            }
        """)
        help_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        help_label.setCursor(QCursor(Qt.CursorShape.WhatsThisCursor))
        help_label.setToolTip(
            "<b>匹配方式说明：</b><br><br>"

            "<b>📌 使用主键列匹配行</b><br>"
            "适用场景：两个文件的数据行顺序不一致<br>"
            "工作原理：根据指定列的值来匹配对应的行进行比较<br>"
            "使用示例：<br>"
            "• A文件第3行的ID是'001'，B文件第5行的ID也是'001'<br>"
            "• 系统会自动将这两行匹配起来进行比较<br>"
            "• 支持设置两个主键列进行组合匹配（如：姓名+日期）<br><br>"
            "<b>填写说明：</b><br>"
            "• 列顺序相同时：只需填写A文件的主键列，B文件留空即可<br>"
            "• 列顺序不同时：需要分别指定A文件和B文件的主键列<br><br>"

            "<b>📌 根据标题行匹配列</b><br>"
            "适用场景：两个文件的列顺序不一致<br>"
            "工作原理：根据标题行的列名来匹配对应的列进行比较<br>"
            "使用示例：<br>"
            "• A文件的'姓名'列在第2列（B列）<br>"
            "• B文件的'姓名'列在第4列（D列）<br>"
            "• 系统会自动将这两列匹配起来进行比较<br>"
            "• 默认使用第1行作为标题行，可自定义<br><br>"

            "<b>💡 使用技巧：</b><br>"
            "• 两种匹配方式可以同时启用<br>"
            "• 同时启用时可处理行列都乱序的情况<br>"
            "• 如果不启用，则按位置逐行逐列比较<br>"
            "• 主键列必须包含唯一值，否则可能匹配错误"
        )
        title_layout.addWidget(help_label)
        title_layout.addStretch()
        match_main_layout.addLayout(title_layout)

        match_layout = QVBoxLayout()
        
        # 主键列匹配行
        self.use_key_match_check = QCheckBox("使用主键列匹配行")
        self.use_key_match_check.setToolTip("勾选后根据指定列的值匹配行，处理行顺序不同的情况")
        self.use_key_match_check.stateChanged.connect(self._on_key_match_changed)
        match_layout.addWidget(self.use_key_match_check)

        key_input_widget = QWidget()
        key_input_layout = QVBoxLayout(key_input_widget)
        key_input_layout.setContentsMargins(20, 0, 0, 0)

        # A文件主键列
        key_a_layout = QHBoxLayout()
        key_a_layout.addWidget(QLabel("A文件:"))
        self.global_key_col_input = QLineEdit()
        self.global_key_col_input.setPlaceholderText("如 B")
        self.global_key_col_input.setMaximumWidth(50)
        self.global_key_col_input.setEnabled(False)
        key_a_layout.addWidget(self.global_key_col_input)

        key_a_layout.addWidget(QLabel("+"))
        self.global_key_col2_input = QLineEdit()
        self.global_key_col2_input.setPlaceholderText("如 C")
        self.global_key_col2_input.setMaximumWidth(50)
        self.global_key_col2_input.setEnabled(False)
        self.global_key_col2_input.setToolTip("第二主键列（可选）")
        key_a_layout.addWidget(self.global_key_col2_input)
        key_a_layout.addStretch()
        key_input_layout.addLayout(key_a_layout)

        # B文件主键列
        key_b_layout = QHBoxLayout()
        key_b_layout.addWidget(QLabel("B文件:"))
        self.global_key_col_input_b = QLineEdit()
        self.global_key_col_input_b.setPlaceholderText("如 B")
        self.global_key_col_input_b.setMaximumWidth(50)
        self.global_key_col_input_b.setEnabled(False)
        key_b_layout.addWidget(self.global_key_col_input_b)

        key_b_layout.addWidget(QLabel("+"))
        self.global_key_col2_input_b = QLineEdit()
        self.global_key_col2_input_b.setPlaceholderText("如 C")
        self.global_key_col2_input_b.setMaximumWidth(50)
        self.global_key_col2_input_b.setEnabled(False)
        self.global_key_col2_input_b.setToolTip("第二主键列（可选）")
        key_b_layout.addWidget(self.global_key_col2_input_b)
        key_b_layout.addStretch()
        key_input_layout.addLayout(key_b_layout)

        match_layout.addWidget(key_input_widget)
        
        # 根据标题行匹配列
        self.use_header_match_check = QCheckBox("根据标题行匹配列")
        self.use_header_match_check.setToolTip("根据标题行的列名匹配列，处理两个文件列顺序不同的情况")
        self.use_header_match_check.stateChanged.connect(self._on_header_match_changed)
        match_layout.addWidget(self.use_header_match_check)
        
        header_input_widget = QWidget()
        header_input_layout = QHBoxLayout(header_input_widget)
        header_input_layout.setContentsMargins(20, 0, 0, 0)
        header_input_layout.addWidget(QLabel("标题行:"))
        self.global_header_row_input = QLineEdit()
        self.global_header_row_input.setPlaceholderText("如 1")
        self.global_header_row_input.setMaximumWidth(50)
        self.global_header_row_input.setEnabled(False)
        self.global_header_row_input.setText("1")  # 默认第1行
        header_input_layout.addWidget(self.global_header_row_input)
        header_input_layout.addStretch()
        match_layout.addWidget(header_input_widget)

        match_main_layout.addLayout(match_layout)
        layout.addWidget(match_group)
        
        # 开始比较按钮
        self.compare_btn = QPushButton("开始比较")
        self.compare_btn.setObjectName("compareBtn")
        self.compare_btn.clicked.connect(self._on_compare_clicked)
        layout.addWidget(self.compare_btn)
        
        layout.addStretch()
        
        # 设置滚动区域
        scroll.setWidget(content)
        outer_layout.addWidget(scroll)
    
    def _apply_styles(self):
        """应用样式"""
        self.setStyleSheet("""
            ConfigPanel {
                background-color: #ffffff;
                border: 1px solid #e0e0e0;
                border-radius: 8px;
            }
            #panelTitle {
                font-size: 14px;
                font-weight: bold;
                color: #333333;
            }
            #smartWidget {
                background-color: #f5f5f5;
                border: 1px solid #ddd;
                border-radius: 4px;
            }
            QGroupBox {
                font-weight: bold;
                border: 1px solid #e0e0e0;
                border-radius: 4px;
                margin-top: 12px;
                padding-top: 12px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                left: 10px;
                padding: 0 5px;
            }
            QComboBox, QLineEdit {
                padding: 6px;
                border: 1px solid #e0e0e0;
                border-radius: 4px;
            }
            QListWidget {
                border: 1px solid #e0e0e0;
                border-radius: 4px;
            }
            #compareBtn {
                background-color: #4caf50;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 12px;
                font-size: 14px;
                font-weight: bold;
            }
            #compareBtn:hover {
                background-color: #43a047;
            }
        """)
    
    def _on_mode_changed(self, index: int):
        """比较模式变化"""
        mode = self.mode_combo.currentData()
        if mode == "SMART":
            self.smart_group.show()
            self.compare_btn.setText("智能比较")
        else:
            self.smart_group.hide()
            self.compare_btn.setText("开始比较")
    
    def _on_key_check_changed(self, state: int):
        """主键列复选框变化（智能匹配）"""
        self.key_col_input.setEnabled(state == Qt.CheckState.Checked.value)
    
    def _on_key_match_changed(self, state: int):
        """全局主键列复选框变化"""
        enabled = state == Qt.CheckState.Checked.value
        self.global_key_col_input.setEnabled(enabled)
        self.global_key_col2_input.setEnabled(enabled)
        self.global_key_col_input_b.setEnabled(enabled)
        self.global_key_col2_input_b.setEnabled(enabled)
    
    def _on_header_match_changed(self, state: int):
        """首行匹配列复选框变化"""
        self.global_header_row_input.setEnabled(state == Qt.CheckState.Checked.value)
    
    def _on_all_sheets_changed(self, state: int):
        """全部工作表复选框变化"""
        self.sheet_list.setEnabled(state != Qt.CheckState.Checked.value)
    
    def _on_compare_clicked(self):
        """比较按钮点击"""
        mode = self.mode_combo.currentData()
        if mode == "SMART":
            self.smart_compare_clicked.emit()
        else:
            self.compare_clicked.emit()
    
    def set_sheet_list(self, sheets: List[str]):
        """设置工作表列表"""
        self.sheet_list.clear()
        for sheet in sheets:
            item = QListWidgetItem(sheet)
            item.setSelected(True)
            self.sheet_list.addItem(item)
    
    def get_compare_mode(self) -> CompareMode:
        """获取比较模式"""
        mode = self.mode_combo.currentData()
        if mode == "SMART":
            return CompareMode.EXACT
        return mode
    
    def is_smart_mode(self) -> bool:
        """是否为智能匹配模式"""
        return self.mode_combo.currentData() == "SMART"
    
    def get_compare_options(self) -> CompareOptions:
        """获取比较选项"""
        options = CompareOptions()
        options.ignore_format = self.ignore_format_check.isChecked()
        options.ignore_case = self.ignore_case_check.isChecked()
        options.ignore_whitespace = self.ignore_whitespace_check.isChecked()
        options.ignore_empty_rows = self.ignore_empty_rows_check.isChecked()
        return options
    
    def get_smart_compare_settings(self) -> dict:
        """获取智能比较设置"""
        return {
            'range_str': self.range_input.text().strip(),
            'use_header': self.use_header_check.isChecked(),
            'use_key': self.use_key_check.isChecked(),
            'key_column': self.key_col_input.text().strip(),
        }
    
    def get_selected_sheets(self) -> Optional[List[str]]:
        """获取选中的工作表"""
        if self.all_sheets_check.isChecked():
            return None
        return [item.text() for item in self.sheet_list.selectedItems()]
    
    def get_key_column_config(self) -> dict:
        """
        获取全局主键列配置（用于精确匹配等模式）
        返回: {'a': (主键列1索引, 主键列2索引), 'b': (主键列1索引, 主键列2索引)}
              0-indexed，None 表示未指定
        """
        if not self.use_key_match_check.isChecked():
            return {'a': (None, None), 'b': (None, None)}

        def parse_col(text):
            key_str = text.strip().upper()
            if not key_str:
                return None
            if key_str.isdigit():
                return int(key_str) - 1
            else:
                col_idx = 0
                for char in key_str:
                    if 'A' <= char <= 'Z':
                        col_idx = col_idx * 26 + (ord(char) - ord('A') + 1)
                return col_idx - 1 if col_idx > 0 else None

        # A文件主键列
        key_col1_a = parse_col(self.global_key_col_input.text())
        key_col2_a = parse_col(self.global_key_col2_input.text())

        # B文件主键列（如果未填写，使用A文件的配置）
        key_col1_b_text = self.global_key_col_input_b.text().strip()
        key_col2_b_text = self.global_key_col2_input_b.text().strip()

        if key_col1_b_text:
            key_col1_b = parse_col(key_col1_b_text)
        else:
            key_col1_b = key_col1_a  # 默认使用A文件的配置

        if key_col2_b_text:
            key_col2_b = parse_col(key_col2_b_text)
        else:
            key_col2_b = key_col2_a  # 默认使用A文件的配置

        return {
            'a': (key_col1_a, key_col2_a),
            'b': (key_col1_b, key_col2_b)
        }
    
    def get_header_row_config(self) -> Optional[int]:
        """
        获取首行匹配列配置（用于处理列顺序不同的情况）
        返回: 标题行索引（0-indexed），如果未启用返回 None
        """
        if not self.use_header_match_check.isChecked():
            return None
        
        row_str = self.global_header_row_input.text().strip()
        if not row_str or not row_str.isdigit():
            return 0  # 默认第一行
        
        return int(row_str) - 1  # 用户输入是1-indexed
