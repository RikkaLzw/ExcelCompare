"""
文件选择面板

支持拖拽上传和点击选择文件。
"""
from pathlib import Path
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QFrame
)
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QDragEnterEvent, QDropEvent

from src.models.excel_model import WorkbookData
from src.services.excel_service import ExcelService


class FilePanel(QFrame):
    """文件选择面板"""
    
    file_dropped = pyqtSignal(str)  # 文件拖入信号
    
    def __init__(self, title: str = "文件", parent=None):
        super().__init__(parent)
        self.title = title
        self._file_path: str = ""
        
        self.setAcceptDrops(True)
        self._setup_ui()
        self._apply_styles()
    
    def _setup_ui(self):
        """设置 UI"""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)
        
        # 标题
        self.title_label = QLabel(self.title)
        self.title_label.setObjectName("titleLabel")
        layout.addWidget(self.title_label)
        
        # 拖拽区域
        self.drop_area = QLabel("拖拽 Excel 文件到此处\n或点击下方按钮选择")
        self.drop_area.setObjectName("dropArea")
        self.drop_area.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.drop_area.setMinimumHeight(80)
        layout.addWidget(self.drop_area)
        
        # 文件信息
        self.file_info = QLabel("")
        self.file_info.setObjectName("fileInfo")
        self.file_info.setWordWrap(True)
        self.file_info.hide()
        layout.addWidget(self.file_info)
        
        # 按钮区
        button_layout = QHBoxLayout()
        self.select_btn = QPushButton("选择文件")
        self.select_btn.setObjectName("selectBtn")
        self.select_btn.clicked.connect(self._on_select_clicked)
        button_layout.addWidget(self.select_btn)
        
        self.clear_btn = QPushButton("清除")
        self.clear_btn.setObjectName("clearBtn")
        self.clear_btn.clicked.connect(self._clear_file)
        self.clear_btn.hide()
        button_layout.addWidget(self.clear_btn)
        
        layout.addLayout(button_layout)
    
    def _apply_styles(self):
        """应用样式"""
        self.setStyleSheet("""
            FilePanel {
                background-color: #ffffff;
                border: 1px solid #e0e0e0;
                border-radius: 8px;
            }
            #titleLabel {
                font-size: 14px;
                font-weight: bold;
                color: #333333;
            }
            #dropArea {
                background-color: #fafafa;
                border: 2px dashed #cccccc;
                border-radius: 6px;
                color: #888888;
                font-size: 12px;
                padding: 16px;
            }
            #dropArea[dragOver="true"] {
                background-color: #e3f2fd;
                border-color: #2196f3;
            }
            #fileInfo {
                color: #666666;
                font-size: 12px;
                padding: 8px;
                background-color: #f5f5f5;
                border-radius: 4px;
            }
            #selectBtn {
                background-color: #2196f3;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 8px 16px;
                font-size: 13px;
            }
            #selectBtn:hover {
                background-color: #1976d2;
            }
            #clearBtn {
                background-color: #f5f5f5;
                color: #666666;
                border: 1px solid #e0e0e0;
                border-radius: 4px;
                padding: 8px 16px;
                font-size: 13px;
            }
            #clearBtn:hover {
                background-color: #eeeeee;
            }
        """)
    
    def dragEnterEvent(self, event: QDragEnterEvent):
        """拖入事件"""
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if urls and self._is_excel_file(urls[0].toLocalFile()):
                event.acceptProposedAction()
                self.drop_area.setProperty("dragOver", True)
                self.drop_area.style().unpolish(self.drop_area)
                self.drop_area.style().polish(self.drop_area)
                return
        event.ignore()
    
    def dragLeaveEvent(self, event):
        """拖离事件"""
        self.drop_area.setProperty("dragOver", False)
        self.drop_area.style().unpolish(self.drop_area)
        self.drop_area.style().polish(self.drop_area)
    
    def dropEvent(self, event: QDropEvent):
        """放下事件"""
        self.drop_area.setProperty("dragOver", False)
        self.drop_area.style().unpolish(self.drop_area)
        self.drop_area.style().polish(self.drop_area)
        
        urls = event.mimeData().urls()
        if urls:
            file_path = urls[0].toLocalFile()
            if self._is_excel_file(file_path):
                self._file_path = file_path
                self.file_dropped.emit(file_path)
    
    def _is_excel_file(self, path: str) -> bool:
        """检查是否为 Excel 文件"""
        ext = Path(path).suffix.lower()
        return ext in {'.xlsx', '.xls'}
    
    def _on_select_clicked(self):
        """选择按钮点击"""
        from PyQt6.QtWidgets import QFileDialog
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择 Excel 文件",
            "",
            "Excel 文件 (*.xlsx *.xls);;所有文件 (*.*)"
        )
        if file_path:
            self._file_path = file_path
            self.file_dropped.emit(file_path)
    
    def _clear_file(self):
        """清除文件"""
        self._file_path = ""
        self.drop_area.show()
        self.file_info.hide()
        self.clear_btn.hide()
    
    def set_file_info(self, workbook: WorkbookData):
        """设置文件信息"""
        self._file_path = workbook.file_path
        
        size_str = ExcelService.format_file_size(workbook.file_size)
        sheets_str = ", ".join(workbook.sheet_names[:3])
        if len(workbook.sheet_names) > 3:
            sheets_str += f" ... (+{len(workbook.sheet_names) - 3})"
        
        info_text = f"""<b>{workbook.file_name}</b><br>
        大小: {size_str}<br>
        修改时间: {workbook.modified_time}<br>
        工作表: {sheets_str}"""
        
        self.file_info.setText(info_text)
        self.drop_area.hide()
        self.file_info.show()
        self.clear_btn.show()
    
    @property
    def file_path(self) -> str:
        return self._file_path
