"""
差异列表面板

以表格形式展示所有差异。
"""
from typing import List
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QTableWidget, QTableWidgetItem,
    QHeaderView, QFrame, QAbstractItemView
)
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QColor, QBrush

from src.models.diff_model import DiffResult, DiffType


class DiffListPanel(QFrame):
    """差异列表面板"""
    
    diff_selected = pyqtSignal(int)  # 差异选中信号（索引）
    
    # 差异类型背景色
    TYPE_COLORS = {
        DiffType.MODIFIED: QColor("#fff9c4"),
        DiffType.ADDED: QColor("#c8e6c9"),
        DiffType.DELETED: QColor("#ffcdd2"),
        DiffType.FORMAT_CHANGED: QColor("#ffe0b2"),
    }
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self._diffs: List[DiffResult] = []
        self._setup_ui()
        self._apply_styles()
    
    def _setup_ui(self):
        """设置 UI"""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(4)
        
        # 标题
        title = QLabel("差异列表")
        title.setObjectName("panelTitle")
        layout.addWidget(title)
        
        # 表格
        self.table = QTableWidget()
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels([
            "序号", "工作表", "位置", "类型", "原值", "新值"
        ])
        
        # 设置列宽
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Fixed)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Fixed)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(5, QHeaderView.ResizeMode.Stretch)
        
        self.table.setColumnWidth(0, 50)
        self.table.setColumnWidth(2, 60)
        self.table.setColumnWidth(3, 80)
        
        # 设置行为
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        
        # 连接信号
        self.table.itemSelectionChanged.connect(self._on_selection_changed)
        
        layout.addWidget(self.table)
    
    def _apply_styles(self):
        """应用样式"""
        self.setStyleSheet("""
            DiffListPanel {
                background-color: #ffffff;
                border: 1px solid #e0e0e0;
                border-radius: 8px;
            }
            #panelTitle {
                font-size: 14px;
                font-weight: bold;
                color: #333333;
            }
            QTableWidget {
                border: 1px solid #e0e0e0;
                border-radius: 4px;
                gridline-color: #e0e0e0;
            }
            QHeaderView::section {
                background-color: #f5f5f5;
                padding: 6px;
                border: 1px solid #e0e0e0;
                font-weight: bold;
            }
        """)
    
    def set_diffs(self, diffs: List[DiffResult]):
        """设置差异列表"""
        self._diffs = diffs
        self.table.setRowCount(len(diffs))
        
        for i, diff in enumerate(diffs):
            # 序号
            item_idx = QTableWidgetItem(str(i + 1))
            item_idx.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table.setItem(i, 0, item_idx)
            
            # 工作表
            item_sheet = QTableWidgetItem(diff.sheet)
            self.table.setItem(i, 1, item_sheet)
            
            # 位置
            item_pos = QTableWidgetItem(diff.position)
            item_pos.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table.setItem(i, 2, item_pos)
            
            # 类型
            item_type = QTableWidgetItem(diff.type_display)
            item_type.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            item_type.setBackground(QBrush(self.TYPE_COLORS.get(diff.diff_type, QColor("#ffffff"))))
            self.table.setItem(i, 3, item_type)
            
            # 原值
            old_val = str(diff.old_value) if diff.old_value is not None else ""
            item_old = QTableWidgetItem(old_val[:100])  # 截断过长内容
            self.table.setItem(i, 4, item_old)
            
            # 新值
            new_val = str(diff.new_value) if diff.new_value is not None else ""
            item_new = QTableWidgetItem(new_val[:100])
            self.table.setItem(i, 5, item_new)
    
    def _on_selection_changed(self):
        """选中变化"""
        selected = self.table.selectedItems()
        if selected:
            row = selected[0].row()
            self.diff_selected.emit(row)
    
    def select_diff(self, index: int):
        """选中指定差异"""
        if 0 <= index < self.table.rowCount():
            self.table.selectRow(index)
            self.table.scrollToItem(self.table.item(index, 0))
