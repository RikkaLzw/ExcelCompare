"""
Excel 比较服务

提供多种比较模式来比较两个 Excel 文件的差异。
"""
from enum import Enum
from typing import List, Optional, Set, Tuple

from src.models.excel_model import WorkbookData, SheetData, CellData, CellType
from src.models.diff_model import DiffResult, DiffSummary, DiffType, CompareResult


class CompareMode(Enum):
    """比较模式"""
    EXACT = "exact"         # 精确匹配：逐单元格完全匹配
    NUMERIC = "numeric"     # 数值比较：只比较数值，忽略文本格式
    STRUCTURE = "structure" # 结构比较：比较行列增删变化
    FORMULA = "formula"     # 公式比较：比较单元格公式


class CompareOptions:
    """比较选项"""
    def __init__(self):
        self.ignore_format = True           # 忽略格式
        self.ignore_empty_rows = False      # 忽略空白行
        self.ignore_empty_cols = False      # 忽略空白列
        self.ignore_case = False            # 忽略大小写
        self.ignore_whitespace = False      # 忽略前后空格
        self.ignore_hidden = False          # 忽略隐藏行列
        self.ignore_comments = True         # 忽略批注


class CompareService:
    """比较服务"""
    
    @classmethod
    def compare(
        cls,
        workbook_a: WorkbookData,
        workbook_b: WorkbookData,
        mode: CompareMode = CompareMode.EXACT,
        options: Optional[CompareOptions] = None,
        selected_sheets: Optional[List[str]] = None
    ) -> CompareResult:
        """
        比较两个工作簿
        
        Args:
            workbook_a: 第一个工作簿
            workbook_b: 第二个工作簿
            mode: 比较模式
            options: 比较选项
            selected_sheets: 要比较的工作表列表，None 表示全部
            
        Returns:
            CompareResult 对象
        """
        if options is None:
            options = CompareOptions()
        
        diffs: List[DiffResult] = []
        summary = DiffSummary()
        
        # 确定要比较的工作表
        sheets_a = set(workbook_a.sheet_names)
        sheets_b = set(workbook_b.sheet_names)
        
        if selected_sheets:
            sheets_to_compare = set(selected_sheets) & (sheets_a | sheets_b)
        else:
            sheets_to_compare = sheets_a | sheets_b
        
        # 比较每个工作表
        for sheet_name in sheets_to_compare:
            sheet_a = workbook_a.get_sheet(sheet_name)
            sheet_b = workbook_b.get_sheet(sheet_name)
            
            if sheet_a is None and sheet_b is not None:
                # 工作表在 B 中新增
                sheet_diffs = cls._mark_all_cells(sheet_b, DiffType.ADDED, is_new=True)
            elif sheet_a is not None and sheet_b is None:
                # 工作表在 B 中删除
                sheet_diffs = cls._mark_all_cells(sheet_a, DiffType.DELETED, is_new=False)
            else:
                # 两个工作表都存在，进行比较
                if mode == CompareMode.EXACT:
                    sheet_diffs = cls._compare_exact(sheet_a, sheet_b, options)
                elif mode == CompareMode.NUMERIC:
                    sheet_diffs = cls._compare_numeric(sheet_a, sheet_b, options)
                elif mode == CompareMode.FORMULA:
                    sheet_diffs = cls._compare_formula(sheet_a, sheet_b, options)
                else:  # STRUCTURE
                    sheet_diffs = cls._compare_structure(sheet_a, sheet_b, options)
            
            # 更新统计
            for diff in sheet_diffs:
                summary.add_diff(diff.diff_type)
            diffs.extend(sheet_diffs)
        
        return CompareResult(
            file_a=workbook_a.file_name,
            file_b=workbook_b.file_name,
            diffs=diffs,
            summary=summary
        )
    
    @classmethod
    def _compare_exact(
        cls, 
        sheet_a: SheetData, 
        sheet_b: SheetData, 
        options: CompareOptions
    ) -> List[DiffResult]:
        """精确匹配比较"""
        diffs = []
        
        max_rows = max(sheet_a.row_count, sheet_b.row_count)
        max_cols = max(sheet_a.col_count, sheet_b.col_count)
        
        for row in range(max_rows):
            # 检查是否需要跳过空行
            if options.ignore_empty_rows:
                row_a_empty = cls._is_row_empty(sheet_a, row)
                row_b_empty = cls._is_row_empty(sheet_b, row)
                if row_a_empty and row_b_empty:
                    continue
            
            for col in range(max_cols):
                cell_a = sheet_a.get_cell(row, col)
                cell_b = sheet_b.get_cell(row, col)
                
                diff = cls._compare_cells(
                    sheet_a.name, row, col, cell_a, cell_b, options
                )
                if diff:
                    diffs.append(diff)
        
        return diffs
    
    @classmethod
    def _compare_numeric(
        cls,
        sheet_a: SheetData,
        sheet_b: SheetData,
        options: CompareOptions
    ) -> List[DiffResult]:
        """数值比较（只比较数值类型）"""
        diffs = []
        
        max_rows = max(sheet_a.row_count, sheet_b.row_count)
        max_cols = max(sheet_a.col_count, sheet_b.col_count)
        
        for row in range(max_rows):
            for col in range(max_cols):
                cell_a = sheet_a.get_cell(row, col)
                cell_b = sheet_b.get_cell(row, col)
                
                # 尝试转换为数值比较
                val_a = cls._to_numeric(cell_a)
                val_b = cls._to_numeric(cell_b)
                
                if val_a != val_b:
                    diff_type = DiffType.MODIFIED
                    if val_a is None and val_b is not None:
                        diff_type = DiffType.ADDED
                    elif val_a is not None and val_b is None:
                        diff_type = DiffType.DELETED
                    
                    diffs.append(DiffResult(
                        sheet=sheet_a.name,
                        row=row,
                        col=col,
                        diff_type=diff_type,
                        old_value=val_a,
                        new_value=val_b
                    ))
        
        return diffs
    
    @classmethod
    def _compare_formula(
        cls,
        sheet_a: SheetData,
        sheet_b: SheetData,
        options: CompareOptions
    ) -> List[DiffResult]:
        """公式比较"""
        diffs = []
        
        max_rows = max(sheet_a.row_count, sheet_b.row_count)
        max_cols = max(sheet_a.col_count, sheet_b.col_count)
        
        for row in range(max_rows):
            for col in range(max_cols):
                cell_a = sheet_a.get_cell(row, col)
                cell_b = sheet_b.get_cell(row, col)
                
                formula_a = cell_a.formula if cell_a else None
                formula_b = cell_b.formula if cell_b else None
                
                # 如果没有公式，比较值
                if formula_a is None and formula_b is None:
                    diff = cls._compare_cells(
                        sheet_a.name, row, col, cell_a, cell_b, options
                    )
                    if diff:
                        diffs.append(diff)
                elif formula_a != formula_b:
                    diff_type = DiffType.MODIFIED
                    if formula_a is None and formula_b is not None:
                        diff_type = DiffType.ADDED
                    elif formula_a is not None and formula_b is None:
                        diff_type = DiffType.DELETED
                    
                    diffs.append(DiffResult(
                        sheet=sheet_a.name,
                        row=row,
                        col=col,
                        diff_type=diff_type,
                        old_value=cell_a.value if cell_a else None,
                        new_value=cell_b.value if cell_b else None,
                        old_formula=formula_a,
                        new_formula=formula_b
                    ))
        
        return diffs
    
    @classmethod
    def _compare_structure(
        cls,
        sheet_a: SheetData,
        sheet_b: SheetData,
        options: CompareOptions
    ) -> List[DiffResult]:
        """结构比较（检测行列增删）"""
        # 简化实现：直接比较行数和列数的差异
        # 完整实现需要使用 LCS 算法
        diffs = []
        
        # 检测行变化
        if sheet_a.row_count != sheet_b.row_count:
            if sheet_a.row_count < sheet_b.row_count:
                # 新增行
                for row in range(sheet_a.row_count, sheet_b.row_count):
                    for col in range(sheet_b.col_count):
                        cell_b = sheet_b.get_cell(row, col)
                        if cell_b and not cell_b.is_empty():
                            diffs.append(DiffResult(
                                sheet=sheet_a.name,
                                row=row,
                                col=col,
                                diff_type=DiffType.ADDED,
                                new_value=cell_b.value
                            ))
            else:
                # 删除行
                for row in range(sheet_b.row_count, sheet_a.row_count):
                    for col in range(sheet_a.col_count):
                        cell_a = sheet_a.get_cell(row, col)
                        if cell_a and not cell_a.is_empty():
                            diffs.append(DiffResult(
                                sheet=sheet_a.name,
                                row=row,
                                col=col,
                                diff_type=DiffType.DELETED,
                                old_value=cell_a.value
                            ))
        
        # 对共同行进行精确比较
        common_rows = min(sheet_a.row_count, sheet_b.row_count)
        common_cols = min(sheet_a.col_count, sheet_b.col_count)
        
        for row in range(common_rows):
            for col in range(common_cols):
                cell_a = sheet_a.get_cell(row, col)
                cell_b = sheet_b.get_cell(row, col)
                
                diff = cls._compare_cells(
                    sheet_a.name, row, col, cell_a, cell_b, options
                )
                if diff:
                    diffs.append(diff)
        
        return diffs
    
    @classmethod
    def _compare_cells(
        cls,
        sheet_name: str,
        row: int,
        col: int,
        cell_a: Optional[CellData],
        cell_b: Optional[CellData],
        options: CompareOptions
    ) -> Optional[DiffResult]:
        """比较两个单元格"""
        val_a = cell_a.value if cell_a else None
        val_b = cell_b.value if cell_b else None
        
        # 应用忽略选项
        if options.ignore_case and isinstance(val_a, str) and isinstance(val_b, str):
            val_a = val_a.lower()
            val_b = val_b.lower()
        
        if options.ignore_whitespace:
            if isinstance(val_a, str):
                val_a = val_a.strip()
            if isinstance(val_b, str):
                val_b = val_b.strip()
        
        # 判断是否相等
        is_empty_a = val_a is None or val_a == ""
        is_empty_b = val_b is None or val_b == ""
        
        if is_empty_a and is_empty_b:
            return None
        
        if val_a == val_b:
            # 值相同，检查格式差异
            if not options.ignore_format:
                style_diff = cls._compare_styles(cell_a, cell_b)
                if style_diff:
                    return DiffResult(
                        sheet=sheet_name,
                        row=row,
                        col=col,
                        diff_type=DiffType.FORMAT_CHANGED,
                        old_value=val_a,
                        new_value=val_b
                    )
            return None
        
        # 确定差异类型
        if is_empty_a and not is_empty_b:
            diff_type = DiffType.ADDED
        elif not is_empty_a and is_empty_b:
            diff_type = DiffType.DELETED
        else:
            diff_type = DiffType.MODIFIED
        
        return DiffResult(
            sheet=sheet_name,
            row=row,
            col=col,
            diff_type=diff_type,
            old_value=cell_a.value if cell_a else None,
            new_value=cell_b.value if cell_b else None
        )
    
    @classmethod
    def _compare_styles(
        cls,
        cell_a: Optional[CellData],
        cell_b: Optional[CellData]
    ) -> bool:
        """比较样式是否不同"""
        style_a = cell_a.style if cell_a else None
        style_b = cell_b.style if cell_b else None
        
        if style_a is None and style_b is None:
            return False
        if style_a is None or style_b is None:
            return True
        
        return (
            style_a.font_name != style_b.font_name or
            style_a.font_size != style_b.font_size or
            style_a.font_bold != style_b.font_bold or
            style_a.bg_color != style_b.bg_color
        )
    
    @classmethod
    def _mark_all_cells(
        cls,
        sheet: SheetData,
        diff_type: DiffType,
        is_new: bool
    ) -> List[DiffResult]:
        """标记工作表中所有非空单元格"""
        diffs = []
        for row in range(sheet.row_count):
            for col in range(sheet.col_count):
                cell = sheet.get_cell(row, col)
                if cell and not cell.is_empty():
                    diffs.append(DiffResult(
                        sheet=sheet.name,
                        row=row,
                        col=col,
                        diff_type=diff_type,
                        old_value=None if is_new else cell.value,
                        new_value=cell.value if is_new else None
                    ))
        return diffs
    
    @classmethod
    def _is_row_empty(cls, sheet: SheetData, row: int) -> bool:
        """检查行是否为空"""
        if row >= len(sheet.rows):
            return True
        return all(cell.is_empty() for cell in sheet.rows[row])
    
    @classmethod
    def _to_numeric(cls, cell: Optional[CellData]) -> Optional[float]:
        """尝试将单元格值转换为数值"""
        if cell is None:
            return None
        if cell.cell_type == CellType.NUMBER:
            return float(cell.value)
        if cell.cell_type == CellType.STRING:
            try:
                return float(cell.value)
            except (ValueError, TypeError):
                return None
        return None
