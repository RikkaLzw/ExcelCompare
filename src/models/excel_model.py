"""
Excel 数据模型定义

定义 Excel 文件解析后的统一数据结构。
"""
from dataclasses import dataclass, field
from typing import Any, Optional, List, Dict
from enum import Enum


class CellType(Enum):
    """单元格数据类型"""
    EMPTY = "empty"
    STRING = "string"
    NUMBER = "number"
    BOOLEAN = "boolean"
    DATE = "date"
    FORMULA = "formula"
    ERROR = "error"


@dataclass
class CellStyle:
    """单元格样式"""
    font_name: Optional[str] = None
    font_size: Optional[float] = None
    font_bold: bool = False
    font_italic: bool = False
    font_color: Optional[str] = None
    bg_color: Optional[str] = None
    border: Optional[str] = None
    alignment: Optional[str] = None
    number_format: Optional[str] = None


@dataclass
class CellData:
    """单元格数据"""
    value: Any = None
    formula: Optional[str] = None
    cell_type: CellType = CellType.EMPTY
    style: Optional[CellStyle] = None
    comment: Optional[str] = None
    
    @property
    def display_value(self) -> str:
        """获取显示值"""
        if self.value is None:
            return ""
        return str(self.value)
    
    def is_empty(self) -> bool:
        """判断是否为空单元格"""
        return self.value is None or self.value == ""


@dataclass
class SheetData:
    """工作表数据"""
    name: str
    rows: List[List[CellData]] = field(default_factory=list)
    row_count: int = 0
    col_count: int = 0
    
    def get_cell(self, row: int, col: int) -> Optional[CellData]:
        """获取指定位置的单元格"""
        if 0 <= row < len(self.rows) and 0 <= col < len(self.rows[row]):
            return self.rows[row][col]
        return None


@dataclass
class WorkbookData:
    """工作簿数据"""
    file_path: str
    file_name: str
    file_size: int = 0
    modified_time: str = ""
    sheets: List[SheetData] = field(default_factory=list)
    sheet_names: List[str] = field(default_factory=list)
    
    def get_sheet(self, name: str) -> Optional[SheetData]:
        """根据名称获取工作表"""
        for sheet in self.sheets:
            if sheet.name == name:
                return sheet
        return None
