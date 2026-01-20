# -*- coding: utf-8 -*-
"""
差异视图

双栏表格显示，支持鼠标拖拽选择区域。
"""
from typing import Optional, List, Dict, Tuple
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QTabWidget, QTableView,
    QHeaderView, QLabel, QAbstractItemView, QPushButton, QFrame,
    QCheckBox, QLineEdit
)
from PyQt6.QtCore import Qt, QAbstractTableModel, QModelIndex, pyqtSignal
from PyQt6.QtGui import QBrush, QColor

from src.models.excel_model import WorkbookData, SheetData
from src.models.diff_model import DiffResult, DiffType


class SheetTableModel(QAbstractTableModel):
    """工作表表格数据模型"""
    
    DIFF_COLORS = {
        DiffType.MODIFIED: QColor("#fff9c4"),
        DiffType.ADDED: QColor("#c8e6c9"),
        DiffType.DELETED: QColor("#ffcdd2"),
        DiffType.FORMAT_CHANGED: QColor("#ffe0b2"),
    }
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self._sheet: Optional[SheetData] = None
        self._diff_map: Dict[Tuple[int, int], DiffType] = {}
    
    def set_data(self, sheet: SheetData, diff_map: Dict[Tuple[int, int], DiffType]):
        self.beginResetModel()
        self._sheet = sheet
        self._diff_map = diff_map
        self.endResetModel()
    
    def rowCount(self, parent=QModelIndex()) -> int:
        return self._sheet.row_count if self._sheet else 0
    
    def columnCount(self, parent=QModelIndex()) -> int:
        return self._sheet.col_count if self._sheet else 0
    
    def data(self, index: QModelIndex, role: int = Qt.ItemDataRole.DisplayRole):
        if not index.isValid() or self._sheet is None:
            return None
        
        row, col = index.row(), index.column()
        
        if role == Qt.ItemDataRole.DisplayRole:
            cell = self._sheet.get_cell(row, col)
            return cell.display_value if cell else ""
        
        elif role == Qt.ItemDataRole.BackgroundRole:
            diff_type = self._diff_map.get((row, col))
            if diff_type:
                return QBrush(self.DIFF_COLORS.get(diff_type, QColor("#ffffff")))
        
        elif role == Qt.ItemDataRole.ToolTipRole:
            cell = self._sheet.get_cell(row, col)
            if cell and cell.value is not None:
                tip = f"值: {cell.value}"
                if cell.formula:
                    tip += f"\n公式: {cell.formula}"
                return tip
        
        return None
    
    def headerData(self, section: int, orientation: Qt.Orientation, role: int = Qt.ItemDataRole.DisplayRole):
        if role != Qt.ItemDataRole.DisplayRole:
            return None
        
        if orientation == Qt.Orientation.Horizontal:
            return self._col_to_letter(section)
        else:
            return str(section + 1)
    
    @staticmethod
    def _col_to_letter(col: int) -> str:
        result = ""
        while col >= 0:
            result = chr(col % 26 + ord('A')) + result
            col = col // 26 - 1
        return result


class SelectableTableView(QTableView):
    """支持区域选择的表格视图"""

    selection_changed = pyqtSignal(str)  # 选区变化信号，发送区域字符串如 "A1:D10"
    cell_clicked = pyqtSignal(int, int)  # 单元格点击信号，发送 (row, col)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setSelectionMode(QAbstractItemView.SelectionMode.ContiguousSelection)
        self.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectItems)

        # 强制设置选中颜色，确保即使表格失去焦点也能看到明显的蓝色高亮
        self.setStyleSheet("""
            QTableView {
                selection-background-color: #0078d7;
                selection-color: white;
            }
            QTableView:!active {
                selection-background-color: #0078d7;
                selection-color: white;
            }
        """)

        self.selectionModel()

    def mousePressEvent(self, event):
        """鼠标点击事件"""
        super().mousePressEvent(event)
        index = self.indexAt(event.pos())
        if index.isValid():
            self.cell_clicked.emit(index.row(), index.column())
    
    def selectionChanged(self, selected, deselected):
        super().selectionChanged(selected, deselected)
        # 获取选中区域
        indexes = self.selectionModel().selectedIndexes()
        if indexes:
            rows = [idx.row() for idx in indexes]
            cols = [idx.column() for idx in indexes]
            min_row, max_row = min(rows), max(rows)
            min_col, max_col = min(cols), max(cols)
            
            # 转换为 Excel 格式
            range_str = f"{self._col_to_letter(min_col)}{min_row + 1}:{self._col_to_letter(max_col)}{max_row + 1}"
            self.selection_changed.emit(range_str)
        else:
            self.selection_changed.emit("")
    
    @staticmethod
    def _col_to_letter(col: int) -> str:
        result = ""
        while col >= 0:
            result = chr(col % 26 + ord('A')) + result
            col = col // 26 - 1
        return result
    
    def get_selection_range(self) -> Optional[Tuple[int, int, int, int]]:
        """获取选中区域 (min_row, min_col, max_row, max_col)，0-indexed"""
        indexes = self.selectionModel().selectedIndexes()
        if not indexes:
            return None
        rows = [idx.row() for idx in indexes]
        cols = [idx.column() for idx in indexes]
        return (min(rows), min(cols), max(rows), max(cols))


class DiffView(QWidget):
    """差异视图"""
    
    compare_selection_clicked = pyqtSignal()  # 比较选区按钮点击信号
    
    def __init__(self, parent=None):
        super().__init__(parent)
        
        self._workbook_a: Optional[WorkbookData] = None
        self._workbook_b: Optional[WorkbookData] = None
        self._diffs: List[DiffResult] = []
        self._current_tables: Dict[str, Tuple[SelectableTableView, SelectableTableView]] = {}
        
        self._setup_ui()
        self._apply_styles()
    
    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(4)
        
        # 工作表标签页
        self.tab_widget = QTabWidget()
        layout.addWidget(self.tab_widget, 1)
        
        # 选区信息栏
        selection_bar = QFrame()
        selection_bar.setObjectName("selectionBar")
        selection_layout = QHBoxLayout(selection_bar)
        selection_layout.setContentsMargins(8, 6, 8, 6)
        selection_layout.setSpacing(10)
        
        # 文件A选区
        selection_layout.addWidget(QLabel("文件A:"))
        self.range_a_label = QLabel("未选择")
        self.range_a_label.setObjectName("rangeLabel")
        self.range_a_label.setMinimumWidth(70)
        selection_layout.addWidget(self.range_a_label)
        
        # 文件B选区
        selection_layout.addWidget(QLabel("文件B:"))
        self.range_b_label = QLabel("未选择")
        self.range_b_label.setObjectName("rangeLabel")
        self.range_b_label.setMinimumWidth(70)
        selection_layout.addWidget(self.range_b_label)
        
        # 分隔线
        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.VLine)
        sep.setStyleSheet("color: #ccc;")
        selection_layout.addWidget(sep)
        
        # 同步滚动开关
        self.sync_scroll_check = QCheckBox("同步滚动")
        self.sync_scroll_check.setChecked(True)
        self.sync_scroll_check.setToolTip("勾选后，滚动一个表格时另一个表格也会同步滚动")
        selection_layout.addWidget(self.sync_scroll_check)
        
        selection_layout.addStretch()
        
        # 比较选区按钮
        self.compare_selection_btn = QPushButton("比较选中区域")
        self.compare_selection_btn.setObjectName("compareSelectionBtn")
        self.compare_selection_btn.clicked.connect(self.compare_selection_clicked.emit)
        selection_layout.addWidget(self.compare_selection_btn)
        
        layout.addWidget(selection_bar)
    
    def _apply_styles(self):
        self.setStyleSheet("""
            QTabWidget::pane {
                border: 1px solid #e0e0e0;
                background-color: #ffffff;
            }
            QTabBar::tab {
                padding: 8px 16px;
                margin-right: 2px;
                background-color: #f5f5f5;
                border: 1px solid #e0e0e0;
                border-bottom: none;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
            }
            QTabBar::tab:selected {
                background-color: #ffffff;
            }
            QTableView {
                gridline-color: #e0e0e0;
                selection-background-color: #bbdefb;
            }
            QHeaderView::section {
                background-color: #f5f5f5;
                padding: 4px;
                border: 1px solid #e0e0e0;
                font-weight: bold;
            }
            #selectionBar {
                background-color: #e8f5e9;
                border: 1px solid #c8e6c9;
                border-radius: 4px;
            }
            #rangeLabel {
                font-weight: bold;
                color: #1976d2;
            }
            #compareSelectionBtn {
                background-color: #ff9800;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 8px 16px;
                font-weight: bold;
            }
            #compareSelectionBtn:hover {
                background-color: #f57c00;
            }
        """)
    
    def set_data(
        self, 
        workbook_a: Optional[WorkbookData], 
        workbook_b: Optional[WorkbookData],
        diffs: List[DiffResult]
    ):
        """设置数据"""
        self._workbook_a = workbook_a
        self._workbook_b = workbook_b
        self._diffs = diffs
        self._current_tables.clear()
        
        if not workbook_a and not workbook_b:
            return
        
        # 构建差异映射（文件A和文件B可能有不同的位置）
        # 新增(ADDED)只在B高亮，删除(DELETED)只在A高亮，修改(MODIFIED)两边都高亮
        diff_map_a: Dict[str, Dict[Tuple[int, int], DiffType]] = {}
        diff_map_b: Dict[str, Dict[Tuple[int, int], DiffType]] = {}
        
        for diff in diffs:
            if diff.sheet not in diff_map_a:
                diff_map_a[diff.sheet] = {}
                diff_map_b[diff.sheet] = {}
            
            # 根据差异类型决定高亮位置
            if diff.diff_type == DiffType.ADDED:
                # 新增：只在文件B中高亮
                row_b = diff.row_b if diff.row_b is not None else diff.row
                col_b = diff.col_b if diff.col_b is not None else diff.col
                pos_b = (row_b, col_b)
                diff_map_b[diff.sheet][pos_b] = diff.diff_type
            elif diff.diff_type == DiffType.DELETED:
                # 删除：只在文件A中高亮
                pos_a = (diff.row, diff.col)
                diff_map_a[diff.sheet][pos_a] = diff.diff_type
            else:
                # 修改或格式变化：两边都高亮
                pos_a = (diff.row, diff.col)
                diff_map_a[diff.sheet][pos_a] = diff.diff_type
                
                row_b = diff.row_b if diff.row_b is not None else diff.row
                col_b = diff.col_b if diff.col_b is not None else diff.col
                pos_b = (row_b, col_b)
                diff_map_b[diff.sheet][pos_b] = diff.diff_type
        
        self.tab_widget.clear()
        
        sheets_a = workbook_a.sheet_names if workbook_a else []
        sheets_b = workbook_b.sheet_names if workbook_b else []
        all_sheets = set(sheets_a) | set(sheets_b)
        
        for sheet_name in sorted(all_sheets):
            sheet_a = workbook_a.get_sheet(sheet_name) if workbook_a else None
            sheet_b = workbook_b.get_sheet(sheet_name) if workbook_b else None
            
            sheet_widget, table_a, table_b = self._create_sheet_widget(
                sheet_name,
                sheet_a,
                sheet_b,
                diff_map_a.get(sheet_name, {}),
                diff_map_b.get(sheet_name, {})
            )
            
            self._current_tables[sheet_name] = (table_a, table_b)
            
            diff_count = len(diff_map_a.get(sheet_name, {}))
            tab_text = f"{sheet_name} ({diff_count})" if diff_count > 0 else sheet_name
            self.tab_widget.addTab(sheet_widget, tab_text)
        
        # 重置选区显示
        self.range_a_label.setText("未选择")
        self.range_b_label.setText("未选择")
    
    def _create_sheet_widget(
        self,
        sheet_name: str,
        sheet_a: Optional[SheetData],
        sheet_b: Optional[SheetData],
        diff_map_a: Dict[Tuple[int, int], DiffType],
        diff_map_b: Dict[Tuple[int, int], DiffType]
    ) -> Tuple[QWidget, SelectableTableView, SelectableTableView]:
        """创建工作表视图，返回 (widget, table_a, table_b)"""
        widget = QWidget()
        layout = QHBoxLayout(widget)
        layout.setContentsMargins(4, 4, 4, 4)
        layout.setSpacing(4)
        
        # 左侧表格（文件 A）
        container_a, table_a = self._create_table_view(sheet_a, diff_map_a, "文件 A", 'a')
        layout.addWidget(container_a, 1)
        
        # 右侧表格（文件 B）
        container_b, table_b = self._create_table_view(sheet_b, diff_map_b, "文件 B", 'b')
        layout.addWidget(container_b, 1)
        
        # 同步滚动
        self._sync_scroll(table_a, table_b)
        
        return widget, table_a, table_b
    
    def _create_table_view(
        self,
        sheet: Optional[SheetData],
        diff_map: Dict[Tuple[int, int], DiffType],
        title: str,
        which: str
    ) -> Tuple[QWidget, SelectableTableView]:
        """创建表格视图，返回 (container, table)"""
        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(4)
        
        title_label = QLabel(title)
        title_label.setStyleSheet("font-weight: bold; font-size: 13px; color: #333;")
        layout.addWidget(title_label)
        
        table = SelectableTableView()
        table.setAlternatingRowColors(True)
        table.horizontalHeader().setDefaultSectionSize(80)
        table.verticalHeader().setDefaultSectionSize(24)
        
        # 连接选区变化信号
        if which == 'a':
            table.selection_changed.connect(self._on_selection_a_changed)
            table.cell_clicked.connect(lambda r, c: self._on_cell_clicked_a(r, c))
        else:
            table.selection_changed.connect(self._on_selection_b_changed)
            table.cell_clicked.connect(lambda r, c: self._on_cell_clicked_b(r, c))
        
        model = SheetTableModel()
        if sheet:
            model.set_data(sheet, diff_map)
        table.setModel(model)
        
        layout.addWidget(table, 1)
        return container, table
    
    def _on_selection_a_changed(self, range_str: str):
        self.range_a_label.setText(range_str if range_str else "未选择")

    def _on_selection_b_changed(self, range_str: str):
        self.range_b_label.setText(range_str if range_str else "未选择")

    def _on_cell_clicked_a(self, row: int, col: int):
        """文件A单元格点击，定位到文件B对应位置"""
        current_idx = self.tab_widget.currentIndex()
        if current_idx < 0:
            return

        tab_text = self.tab_widget.tabText(current_idx)
        sheet_name = tab_text.split(" (")[0]

        if sheet_name not in self._current_tables:
            return

        # 查找对应的差异
        for diff in self._diffs:
            if diff.sheet == sheet_name and diff.row == row and diff.col == col:
                # 只处理修改类型的差异（有对应关系）
                if diff.diff_type == DiffType.MODIFIED:
                    row_b = diff.row_b if diff.row_b is not None else diff.row
                    col_b = diff.col_b if diff.col_b is not None else diff.col
                    self._locate_cell_in_table_b(sheet_name, row_b, col_b)
                break

    def _on_cell_clicked_b(self, row: int, col: int):
        """文件B单元格点击，定位到文件A对应位置"""
        current_idx = self.tab_widget.currentIndex()
        if current_idx < 0:
            return

        tab_text = self.tab_widget.tabText(current_idx)
        sheet_name = tab_text.split(" (")[0]

        if sheet_name not in self._current_tables:
            return

        # 查找对应的差异
        for diff in self._diffs:
            if diff.sheet == sheet_name:
                row_b = diff.row_b if diff.row_b is not None else diff.row
                col_b = diff.col_b if diff.col_b is not None else diff.col
                if row_b == row and col_b == col:
                    # 只处理修改类型的差异（有对应关系）
                    if diff.diff_type == DiffType.MODIFIED:
                        self._locate_cell_in_table_a(sheet_name, diff.row, diff.col)
                    break

    def _locate_cell_in_table_a(self, sheet_name: str, row: int, col: int):
        """在文件A表格中定位单元格"""
        if sheet_name not in self._current_tables:
            return

        table_a, _ = self._current_tables[sheet_name]
        model = table_a.model()
        if model:
            index = model.index(row, col)
            table_a.scrollTo(index, QAbstractItemView.ScrollHint.PositionAtCenter)
            table_a.setCurrentIndex(index)

    def _locate_cell_in_table_b(self, sheet_name: str, row: int, col: int):
        """在文件B表格中定位单元格"""
        if sheet_name not in self._current_tables:
            return

        _, table_b = self._current_tables[sheet_name]
        model = table_b.model()
        if model:
            index = model.index(row, col)
            table_b.scrollTo(index, QAbstractItemView.ScrollHint.PositionAtCenter)
            table_b.setCurrentIndex(index)
    
    def _sync_scroll(self, table_a: SelectableTableView, table_b: SelectableTableView):
        """同步两个表格的滚动"""
        # 使用 lambda 捕获表格引用，并检查同步开关
        def sync_h(source, target):
            if self.sync_scroll_check.isChecked():
                target.horizontalScrollBar().setValue(source.horizontalScrollBar().value())
        
        def sync_v(source, target):
            if self.sync_scroll_check.isChecked():
                target.verticalScrollBar().setValue(source.verticalScrollBar().value())
        
        # 连接信号
        table_a.horizontalScrollBar().valueChanged.connect(lambda: sync_h(table_a, table_b))
        table_b.horizontalScrollBar().valueChanged.connect(lambda: sync_h(table_b, table_a))
        table_a.verticalScrollBar().valueChanged.connect(lambda: sync_v(table_a, table_b))
        table_b.verticalScrollBar().valueChanged.connect(lambda: sync_v(table_b, table_a))
    
    def get_current_selections(self) -> Tuple[Optional[str], Optional[Tuple], Optional[str], Optional[Tuple]]:
        """
        获取当前选中的区域
        返回: (sheet_name, range_a, sheet_name, range_b)
        range 格式: (min_row, min_col, max_row, max_col) 0-indexed
        """
        current_idx = self.tab_widget.currentIndex()
        if current_idx < 0:
            return None, None, None, None
        
        tab_text = self.tab_widget.tabText(current_idx)
        # 移除可能的差异数量后缀
        sheet_name = tab_text.split(" (")[0]
        
        if sheet_name not in self._current_tables:
            return None, None, None, None
        
        table_a, table_b = self._current_tables[sheet_name]
        range_a = table_a.get_selection_range()
        range_b = table_b.get_selection_range()
        
        return sheet_name, range_a, sheet_name, range_b

    def scroll_to_diff(self, diff: DiffResult):
        """滚动到指定差异，根据差异类型高亮对应的表格"""
        from src.models.diff_model import DiffType
        from PyQt6.QtWidgets import QAbstractItemView
        
        for i in range(self.tab_widget.count()):
            tab_text = self.tab_widget.tabText(i)
            if tab_text.startswith(diff.sheet):
                self.tab_widget.setCurrentIndex(i)
                
                if diff.sheet in self._current_tables:
                    table_a, table_b = self._current_tables[diff.sheet]
                    
                    # 获取差异位置
                    row_a = diff.row
                    col_a = diff.col
                    row_b = diff.row_b if diff.row_b is not None else diff.row
                    col_b = diff.col_b if diff.col_b is not None else diff.col
                    
                    # 如果行号差异过大，自动关闭同步滚动
                    if abs(row_a - row_b) > 20:
                        self.sync_scroll_check.setChecked(False)
                    
                    # 根据差异类型决定高亮哪边
                    if diff.diff_type == DiffType.ADDED:
                        # 新增：只高亮文件B
                        model_b = table_b.model()
                        if model_b:
                            index_b = model_b.index(row_b, col_b)
                            table_b.scrollTo(index_b, QAbstractItemView.ScrollHint.PositionAtCenter)
                            table_b.setCurrentIndex(index_b)
                        # 清除A的选择
                        table_a.clearSelection()
                    elif diff.diff_type == DiffType.DELETED:
                        # 删除：只高亮文件A
                        model_a = table_a.model()
                        if model_a:
                            index_a = model_a.index(row_a, col_a)
                            table_a.scrollTo(index_a, QAbstractItemView.ScrollHint.PositionAtCenter)
                            table_a.setCurrentIndex(index_a)
                        # 清除B的选择
                        table_b.clearSelection()
                    else:
                        # 修改：两边都高亮
                        model_a = table_a.model()
                        if model_a:
                            index_a = model_a.index(row_a, col_a)
                            table_a.scrollTo(index_a, QAbstractItemView.ScrollHint.PositionAtCenter)
                            table_a.setCurrentIndex(index_a)
                        
                        model_b = table_b.model()
                        if model_b:
                            index_b = model_b.index(row_b, col_b)
                            table_b.scrollTo(index_b, QAbstractItemView.ScrollHint.PositionAtCenter)
                            table_b.setCurrentIndex(index_b)
                break
