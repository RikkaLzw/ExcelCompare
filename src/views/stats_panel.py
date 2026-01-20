"""
统计面板

显示差异统计信息和图表。
"""
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QFrame, QGridLayout
)
from PyQt6.QtCore import Qt

from src.models.diff_model import DiffSummary


class StatsPanel(QFrame):
    """统计面板"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self._setup_ui()
        self._apply_styles()
    
    def _setup_ui(self):
        """设置 UI"""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)
        
        # 标题
        title = QLabel("差异统计")
        title.setObjectName("panelTitle")
        layout.addWidget(title)
        
        # 统计网格
        stats_grid = QGridLayout()
        stats_grid.setSpacing(8)
        
        # 总计
        self.total_label = QLabel("0")
        self.total_label.setObjectName("statValue")
        stats_grid.addWidget(QLabel("总计:"), 0, 0)
        stats_grid.addWidget(self.total_label, 0, 1)
        
        # 修改
        self.modified_label = QLabel("0")
        self.modified_label.setObjectName("modifiedValue")
        stats_grid.addWidget(QLabel("修改:"), 1, 0)
        stats_grid.addWidget(self.modified_label, 1, 1)
        
        # 新增
        self.added_label = QLabel("0")
        self.added_label.setObjectName("addedValue")
        stats_grid.addWidget(QLabel("新增:"), 2, 0)
        stats_grid.addWidget(self.added_label, 2, 1)
        
        # 删除
        self.deleted_label = QLabel("0")
        self.deleted_label.setObjectName("deletedValue")
        stats_grid.addWidget(QLabel("删除:"), 3, 0)
        stats_grid.addWidget(self.deleted_label, 3, 1)
        
        # 格式变化
        self.format_label = QLabel("0")
        self.format_label.setObjectName("formatValue")
        stats_grid.addWidget(QLabel("格式:"), 4, 0)
        stats_grid.addWidget(self.format_label, 4, 1)
        
        layout.addLayout(stats_grid)
    
    def _apply_styles(self):
        """应用样式"""
        self.setStyleSheet("""
            StatsPanel {
                background-color: #ffffff;
                border: 1px solid #e0e0e0;
                border-radius: 8px;
            }
            #panelTitle {
                font-size: 14px;
                font-weight: bold;
                color: #333333;
            }
            #statValue {
                font-size: 18px;
                font-weight: bold;
                color: #2196f3;
            }
            #modifiedValue {
                font-size: 16px;
                font-weight: bold;
                color: #ffc107;
            }
            #addedValue {
                font-size: 16px;
                font-weight: bold;
                color: #4caf50;
            }
            #deletedValue {
                font-size: 16px;
                font-weight: bold;
                color: #f44336;
            }
            #formatValue {
                font-size: 16px;
                font-weight: bold;
                color: #ff9800;
            }
        """)
    
    def set_summary(self, summary: DiffSummary):
        """设置统计摘要"""
        self.total_label.setText(str(summary.total))
        self.modified_label.setText(str(summary.modified))
        self.added_label.setText(str(summary.added))
        self.deleted_label.setText(str(summary.deleted))
        self.format_label.setText(str(summary.format_changed))
