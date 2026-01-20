"""
比较工作线程

在后台执行 Excel 文件比较，避免阻塞 UI。
"""
from typing import Optional, List
from PyQt6.QtCore import QThread, pyqtSignal

from src.models.excel_model import WorkbookData
from src.models.diff_model import CompareResult
from src.services.excel_service import ExcelService
from src.services.compare_service import CompareService, CompareMode, CompareOptions


class CompareWorker(QThread):
    """比较工作线程"""
    
    # 信号定义
    progress_updated = pyqtSignal(int, str)         # 进度更新 (百分比, 消息)
    file_loaded = pyqtSignal(str, object)           # 文件加载完成 (路径, WorkbookData)
    compare_finished = pyqtSignal(object)           # 比较完成 (CompareResult)
    error_occurred = pyqtSignal(str)                # 发生错误 (错误消息)
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.file_a_path: Optional[str] = None
        self.file_b_path: Optional[str] = None
        self.mode: CompareMode = CompareMode.EXACT
        self.options: Optional[CompareOptions] = None
        self.selected_sheets: Optional[List[str]] = None
        
        self._workbook_a: Optional[WorkbookData] = None
        self._workbook_b: Optional[WorkbookData] = None
    
    def set_files(self, file_a: str, file_b: str):
        """设置要比较的文件"""
        self.file_a_path = file_a
        self.file_b_path = file_b
    
    def set_compare_options(
        self, 
        mode: CompareMode = CompareMode.EXACT,
        options: Optional[CompareOptions] = None,
        selected_sheets: Optional[List[str]] = None
    ):
        """设置比较选项"""
        self.mode = mode
        self.options = options or CompareOptions()
        self.selected_sheets = selected_sheets
    
    def run(self):
        """执行比较任务"""
        try:
            # 1. 加载文件 A
            self.progress_updated.emit(10, "正在加载文件 A...")
            self._workbook_a = ExcelService.load_file(self.file_a_path)
            self.file_loaded.emit(self.file_a_path, self._workbook_a)
            
            # 2. 加载文件 B
            self.progress_updated.emit(30, "正在加载文件 B...")
            self._workbook_b = ExcelService.load_file(self.file_b_path)
            self.file_loaded.emit(self.file_b_path, self._workbook_b)
            
            # 3. 执行比较
            self.progress_updated.emit(50, "正在比较文件...")
            result = CompareService.compare(
                self._workbook_a,
                self._workbook_b,
                mode=self.mode,
                options=self.options,
                selected_sheets=self.selected_sheets
            )
            
            # 4. 完成
            self.progress_updated.emit(100, "比较完成")
            self.compare_finished.emit(result)
            
        except FileNotFoundError as e:
            self.error_occurred.emit(f"文件不存在: {str(e)}")
        except ValueError as e:
            self.error_occurred.emit(f"文件错误: {str(e)}")
        except Exception as e:
            self.error_occurred.emit(f"发生错误: {str(e)}")
    
    @property
    def workbook_a(self) -> Optional[WorkbookData]:
        return self._workbook_a
    
    @property
    def workbook_b(self) -> Optional[WorkbookData]:
        return self._workbook_b


class FileLoadWorker(QThread):
    """单文件加载线程"""
    
    loaded = pyqtSignal(str, object)    # 加载完成 (路径, WorkbookData)
    error = pyqtSignal(str, str)        # 发生错误 (路径, 错误消息)
    
    def __init__(self, file_path: str, parent=None):
        super().__init__(parent)
        self.file_path = file_path
    
    def run(self):
        try:
            workbook = ExcelService.load_file(self.file_path)
            self.loaded.emit(self.file_path, workbook)
        except Exception as e:
            self.error.emit(self.file_path, str(e))
