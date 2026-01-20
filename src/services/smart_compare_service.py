"""
智能比较服务

支持基于行/列标题的智能匹配，以及自定义区域比较。
"""
import re
from dataclasses import dataclass
from typing import List, Optional, Dict, Any, Tuple

from src.models.excel_model import WorkbookData, SheetData, CellData
from src.models.diff_model import DiffResult, DiffSummary, DiffType, CompareResult


@dataclass
class CellRange:
    """单元格区域（如 A1:D10）"""
    start_row: int  # 0-indexed
    start_col: int  # 0-indexed
    end_row: int
    end_col: int
    
    @classmethod
    def from_string(cls, range_str: str) -> 'CellRange':
        """
        从字符串解析区域，如 "A1:D10" 或 "B2:F20"
        """
        range_str = range_str.strip().upper()
        
        # 匹配 A1:B2 格式
        match = re.match(r'^([A-Z]+)(\d+):([A-Z]+)(\d+)$', range_str)
        if not match:
            raise ValueError(f"无效的区域格式: {range_str}，请使用如 A1:D10 的格式")
        
        start_col = cls._col_to_index(match.group(1))
        start_row = int(match.group(2)) - 1
        end_col = cls._col_to_index(match.group(3))
        end_row = int(match.group(4)) - 1
        
        if start_row > end_row or start_col > end_col:
            raise ValueError(f"无效的区域: 起始位置必须小于结束位置")
        
        return cls(start_row, start_col, end_row, end_col)
    
    @staticmethod
    def _col_to_index(col_str: str) -> int:
        """列字母转索引（A=0, B=1, ..., Z=25, AA=26）"""
        result = 0
        for char in col_str:
            result = result * 26 + (ord(char) - ord('A') + 1)
        return result - 1
    
    @staticmethod
    def _index_to_col(index: int) -> str:
        """索引转列字母"""
        result = ""
        while index >= 0:
            result = chr(index % 26 + ord('A')) + result
            index = index // 26 - 1
        return result
    
    @property
    def row_count(self) -> int:
        return self.end_row - self.start_row + 1
    
    @property
    def col_count(self) -> int:
        return self.end_col - self.start_col + 1
    
    def __str__(self) -> str:
        return f"{self._index_to_col(self.start_col)}{self.start_row + 1}:{self._index_to_col(self.end_col)}{self.end_row + 1}"


@dataclass 
class SmartCompareOptions:
    """智能比较选项"""
    # 区域设置
    range_a: Optional[CellRange] = None  # 文件A的比较区域
    range_b: Optional[CellRange] = None  # 文件B的比较区域（必须与A大小一致）
    
    # 标题设置
    use_header_row: bool = False         # 是否使用首行作为列标题
    header_row_index: int = 0            # 标题行索引（相对于区域起始）
    use_key_column: bool = False         # 是否使用某列作为行键（主键）
    key_column_index: int = 0            # 键列索引（相对于区域起始）
    
    # 比较选项
    ignore_case: bool = False
    ignore_whitespace: bool = False
    ignore_empty_rows: bool = False


class SmartCompareService:
    """智能比较服务"""
    
    @classmethod
    def compare_with_range(
        cls,
        workbook_a: WorkbookData,
        workbook_b: WorkbookData,
        sheet_name: str,
        options: SmartCompareOptions
    ) -> CompareResult:
        """
        使用指定区域和智能匹配进行比较
        
        Args:
            workbook_a: 工作簿 A
            workbook_b: 工作簿 B
            sheet_name: 要比较的工作表名称
            options: 智能比较选项
            
        Returns:
            CompareResult 对象
        """
        sheet_a = workbook_a.get_sheet(sheet_name)
        sheet_b = workbook_b.get_sheet(sheet_name)
        
        if not sheet_a or not sheet_b:
            raise ValueError(f"工作表 '{sheet_name}' 在两个文件中必须都存在")
        
        diffs: List[DiffResult] = []
        summary = DiffSummary()
        
        # 计算区域偏移量（用于将相对索引转换为绝对行号列号）
        row_offset_a = options.range_a.start_row if options.range_a else 0
        col_offset_a = options.range_a.start_col if options.range_a else 0
        row_offset_b = options.range_b.start_row if options.range_b else 0
        col_offset_b = options.range_b.start_col if options.range_b else 0
        
        # 提取区域数据
        data_a = cls._extract_range_data(sheet_a, options.range_a)
        data_b = cls._extract_range_data(sheet_b, options.range_b)
        
        # 根据选项选择比较方式
        if options.use_key_column:
            # 基于主键列的智能匹配
            diffs = cls._compare_by_key(
                sheet_name, data_a, data_b, options,
                row_offset_a, col_offset_a
            )
        elif options.use_header_row:
            # 基于列标题的匹配
            diffs = cls._compare_by_header(
                sheet_name, data_a, data_b, options,
                row_offset_a, col_offset_a
            )
        else:
            # 位置对位置比较
            diffs = cls._compare_by_position(
                sheet_name, data_a, data_b, options,
                row_offset_a, col_offset_a
            )
        
        # 更新统计
        for diff in diffs:
            summary.add_diff(diff.diff_type)
        
        return CompareResult(
            file_a=workbook_a.file_name,
            file_b=workbook_b.file_name,
            diffs=diffs,
            summary=summary
        )
    
    @classmethod
    def _extract_range_data(
        cls,
        sheet: SheetData,
        cell_range: Optional[CellRange]
    ) -> List[List[Any]]:
        """
        提取指定区域的数据
        
        Returns:
            二维列表，包含区域内所有单元格的值
        """
        if cell_range is None:
            # 如果没有指定区域，返回整个工作表
            return [[cell.value if cell else None for cell in row] for row in sheet.rows]
        
        data = []
        for row_idx in range(cell_range.start_row, cell_range.end_row + 1):
            row_data = []
            for col_idx in range(cell_range.start_col, cell_range.end_col + 1):
                cell = sheet.get_cell(row_idx, col_idx)
                value = cell.value if cell else None
                row_data.append(value)
            data.append(row_data)
        
        return data
    
    @classmethod
    def _compare_by_key(
        cls,
        sheet_name: str,
        data_a: List[List[Any]],
        data_b: List[List[Any]],
        options: SmartCompareOptions,
        row_offset: int = 0,
        col_offset: int = 0
    ) -> List[DiffResult]:
        """
        基于主键列进行比较
        
        将主键列的值作为行标识，匹配两个表中相同键的行进行比较。
        row_offset, col_offset: 原始区域的起始偏移，用于转换为绝对坐标
        """
        diffs = []
        key_col = options.key_column_index
        header_row = options.header_row_index if options.use_header_row else -1
        
        # 获取列标题（用于报告）
        headers = data_a[header_row] if header_row >= 0 and header_row < len(data_a) else None
        
        # 构建键到行的映射
        def build_key_map(data: List[List[Any]], skip_header: bool) -> Dict[Any, Tuple[int, List[Any]]]:
            key_map = {}
            start_row = (header_row + 1) if skip_header else 0
            for i, row in enumerate(data[start_row:], start=start_row):
                if len(row) > key_col:
                    key = row[key_col]
                    if key is not None and key != "":
                        # 标准化键值
                        if options.ignore_case and isinstance(key, str):
                            key = key.lower()
                        if options.ignore_whitespace and isinstance(key, str):
                            key = key.strip()
                        key_map[key] = (i, row)
            return key_map
        
        map_a = build_key_map(data_a, options.use_header_row)
        map_b = build_key_map(data_b, options.use_header_row)
        
        all_keys = set(map_a.keys()) | set(map_b.keys())
        
        for key in sorted(all_keys, key=lambda x: str(x)):
            row_a = map_a.get(key)
            row_b = map_b.get(key)
            
            if row_a is None:
                # 文件B中新增的行
                row_idx, row_data = row_b
                for col_idx, value in enumerate(row_data):
                    if value is not None and value != "":
                        diffs.append(DiffResult(
                            sheet=sheet_name,
                            row=row_idx + row_offset,  # 加上偏移量
                            col=col_idx + col_offset,
                            diff_type=DiffType.ADDED,
                            new_value=value
                        ))
            elif row_b is None:
                # 文件A中删除的行
                row_idx, row_data = row_a
                for col_idx, value in enumerate(row_data):
                    if value is not None and value != "":
                        diffs.append(DiffResult(
                            sheet=sheet_name,
                            row=row_idx + row_offset,  # 加上偏移量
                            col=col_idx + col_offset,
                            diff_type=DiffType.DELETED,
                            old_value=value
                        ))
            else:
                # 两边都有，比较每个单元格
                row_idx_a, row_data_a = row_a
                row_idx_b, row_data_b = row_b
                
                max_cols = max(len(row_data_a), len(row_data_b))
                for col_idx in range(max_cols):
                    val_a = row_data_a[col_idx] if col_idx < len(row_data_a) else None
                    val_b = row_data_b[col_idx] if col_idx < len(row_data_b) else None
                    
                    # 标准化比较
                    cmp_a, cmp_b = val_a, val_b
                    if options.ignore_case:
                        if isinstance(cmp_a, str): cmp_a = cmp_a.lower()
                        if isinstance(cmp_b, str): cmp_b = cmp_b.lower()
                    if options.ignore_whitespace:
                        if isinstance(cmp_a, str): cmp_a = cmp_a.strip()
                        if isinstance(cmp_b, str): cmp_b = cmp_b.strip()
                    
                    if cmp_a != cmp_b:
                        diff_type = DiffType.MODIFIED
                        if (val_a is None or val_a == "") and val_b is not None:
                            diff_type = DiffType.ADDED
                        elif val_a is not None and (val_b is None or val_b == ""):
                            diff_type = DiffType.DELETED
                        
                        diffs.append(DiffResult(
                            sheet=sheet_name,
                            row=row_idx_a + row_offset,  # 加上偏移量
                            col=col_idx + col_offset,
                            diff_type=diff_type,
                            old_value=val_a,
                            new_value=val_b
                        ))
        
        return diffs
    
    @classmethod
    def _compare_by_header(
        cls,
        sheet_name: str,
        data_a: List[List[Any]],
        data_b: List[List[Any]],
        options: SmartCompareOptions,
        row_offset: int = 0,
        col_offset: int = 0
    ) -> List[DiffResult]:
        """
        基于列标题进行比较
        
        匹配两个表中相同标题的列进行比较。
        row_offset, col_offset: 原始区域的起始偏移，用于转换为绝对坐标
        """
        diffs = []
        header_row = options.header_row_index
        
        if header_row >= len(data_a) or header_row >= len(data_b):
            raise ValueError("标题行索引超出数据范围")
        
        headers_a = data_a[header_row]
        headers_b = data_b[header_row]
        
        # 构建标题到列索引的映射
        def build_header_map(headers: List[Any]) -> Dict[str, int]:
            header_map = {}
            for idx, h in enumerate(headers):
                if h is not None and h != "":
                    key = str(h)
                    if options.ignore_case:
                        key = key.lower()
                    if options.ignore_whitespace:
                        key = key.strip()
                    header_map[key] = idx
            return header_map
        
        map_a = build_header_map(headers_a)
        map_b = build_header_map(headers_b)
        
        # 找出共同的标题
        common_headers = set(map_a.keys()) & set(map_b.keys())
        only_in_a = set(map_a.keys()) - set(map_b.keys())
        only_in_b = set(map_b.keys()) - set(map_a.keys())
        
        # 比较共同标题列的数据
        data_start = header_row + 1
        max_rows = max(len(data_a), len(data_b)) - data_start
        
        for header in common_headers:
            col_a = map_a[header]
            col_b = map_b[header]
            
            for row_offset_local in range(max_rows):
                row_idx = data_start + row_offset_local
                
                val_a = data_a[row_idx][col_a] if row_idx < len(data_a) and col_a < len(data_a[row_idx]) else None
                val_b = data_b[row_idx][col_b] if row_idx < len(data_b) and col_b < len(data_b[row_idx]) else None
                
                # 忽略空行
                if options.ignore_empty_rows and (val_a is None or val_a == "") and (val_b is None or val_b == ""):
                    continue
                
                # 标准化比较
                cmp_a, cmp_b = val_a, val_b
                if options.ignore_case:
                    if isinstance(cmp_a, str): cmp_a = cmp_a.lower()
                    if isinstance(cmp_b, str): cmp_b = cmp_b.lower()
                if options.ignore_whitespace:
                    if isinstance(cmp_a, str): cmp_a = cmp_a.strip()
                    if isinstance(cmp_b, str): cmp_b = cmp_b.strip()
                
                if cmp_a != cmp_b:
                    diff_type = DiffType.MODIFIED
                    if (val_a is None or val_a == "") and val_b is not None:
                        diff_type = DiffType.ADDED
                    elif val_a is not None and (val_b is None or val_b == ""):
                        diff_type = DiffType.DELETED
                    
                    diffs.append(DiffResult(
                        sheet=sheet_name,
                        row=row_idx + row_offset,  # 加上偏移量
                        col=col_a + col_offset,
                        diff_type=diff_type,
                        old_value=val_a,
                        new_value=val_b
                    ))
        
        return diffs
    
    @classmethod
    def _compare_by_position(
        cls,
        sheet_name: str,
        data_a: List[List[Any]],
        data_b: List[List[Any]],
        options: SmartCompareOptions,
        row_offset: int = 0,
        col_offset: int = 0
    ) -> List[DiffResult]:
        """
        基于位置进行比较（传统方式）
        row_offset, col_offset: 原始区域的起始偏移，用于转换为绝对坐标
        """
        diffs = []
        
        max_rows = max(len(data_a), len(data_b))
        max_cols = max(
            max(len(row) for row in data_a) if data_a else 0,
            max(len(row) for row in data_b) if data_b else 0
        )
        
        for row_idx in range(max_rows):
            row_a = data_a[row_idx] if row_idx < len(data_a) else []
            row_b = data_b[row_idx] if row_idx < len(data_b) else []
            
            for col_idx in range(max_cols):
                val_a = row_a[col_idx] if col_idx < len(row_a) else None
                val_b = row_b[col_idx] if col_idx < len(row_b) else None
                
                # 忽略空值
                if options.ignore_empty_rows and (val_a is None or val_a == "") and (val_b is None or val_b == ""):
                    continue
                
                # 标准化比较
                cmp_a, cmp_b = val_a, val_b
                if options.ignore_case:
                    if isinstance(cmp_a, str): cmp_a = cmp_a.lower()
                    if isinstance(cmp_b, str): cmp_b = cmp_b.lower()
                if options.ignore_whitespace:
                    if isinstance(cmp_a, str): cmp_a = cmp_a.strip()
                    if isinstance(cmp_b, str): cmp_b = cmp_b.strip()
                
                if cmp_a != cmp_b:
                    diff_type = DiffType.MODIFIED
                    if (val_a is None or val_a == "") and val_b is not None:
                        diff_type = DiffType.ADDED
                    elif val_a is not None and (val_b is None or val_b == ""):
                        diff_type = DiffType.DELETED
                    
                    diffs.append(DiffResult(
                        sheet=sheet_name,
                        row=row_idx + row_offset,  # 加上偏移量
                        col=col_idx + col_offset,
                        diff_type=diff_type,
                        old_value=val_a,
                        new_value=val_b
                    ))
        
        return diffs
