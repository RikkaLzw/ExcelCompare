"""
差异数据模型定义

定义比较结果的数据结构。
"""
from dataclasses import dataclass
from typing import Any, Optional, List
from enum import Enum


class DiffType(Enum):
    """差异类型"""
    MODIFIED = "modified"       # 修改
    ADDED = "added"             # 新增
    DELETED = "deleted"         # 删除
    FORMAT_CHANGED = "format"   # 格式变化


@dataclass
class DiffResult:
    """单个差异结果"""
    sheet: str              # 工作表名称
    row: int                # 文件A行号（0-indexed）
    col: int                # 文件A列号（0-indexed）
    diff_type: DiffType     # 差异类型
    old_value: Any = None   # 原值（文件A）
    new_value: Any = None   # 新值（文件B）
    old_formula: Optional[str] = None
    new_formula: Optional[str] = None
    row_b: Optional[int] = None  # 文件B行号（主键匹配时可能不同）
    col_b: Optional[int] = None  # 文件B列号（主键匹配时可能不同）
    
    @property
    def position(self) -> str:
        """获取单元格位置字符串（如 A1, B2）"""
        col_letter = self._col_to_letter(self.col)
        return f"{col_letter}{self.row + 1}"
    
    @staticmethod
    def _col_to_letter(col: int) -> str:
        """将列索引转换为字母（0=A, 1=B, ...）"""
        result = ""
        while col >= 0:
            result = chr(col % 26 + ord('A')) + result
            col = col // 26 - 1
        return result
    
    @property
    def type_display(self) -> str:
        """差异类型的中文显示"""
        type_map = {
            DiffType.MODIFIED: "修改",
            DiffType.ADDED: "新增",
            DiffType.DELETED: "删除",
            DiffType.FORMAT_CHANGED: "格式变化"
        }
        return type_map.get(self.diff_type, "未知")


@dataclass
class DiffSummary:
    """差异统计摘要"""
    total: int = 0
    modified: int = 0
    added: int = 0
    deleted: int = 0
    format_changed: int = 0
    
    def add_diff(self, diff_type: DiffType):
        """添加一个差异计数"""
        self.total += 1
        if diff_type == DiffType.MODIFIED:
            self.modified += 1
        elif diff_type == DiffType.ADDED:
            self.added += 1
        elif diff_type == DiffType.DELETED:
            self.deleted += 1
        elif diff_type == DiffType.FORMAT_CHANGED:
            self.format_changed += 1


@dataclass
class CompareResult:
    """完整比较结果"""
    file_a: str
    file_b: str
    diffs: List[DiffResult]
    summary: DiffSummary
    
    # 按工作表分组的差异
    diffs_by_sheet: dict = None
    
    # 比较配置（用于报告记录）
    compare_config: dict = None
    
    def __post_init__(self):
        if self.diffs_by_sheet is None:
            self.diffs_by_sheet = {}
            for diff in self.diffs:
                if diff.sheet not in self.diffs_by_sheet:
                    self.diffs_by_sheet[diff.sheet] = []
                self.diffs_by_sheet[diff.sheet].append(diff)
        
        if self.compare_config is None:
            self.compare_config = {}

