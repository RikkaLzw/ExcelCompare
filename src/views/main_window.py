"""
主窗口

应用程序的主界面，包含菜单栏、工具栏、文件面板、表格视图、差异列表等。
"""
import os
from typing import Optional
from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QSplitter, QStatusBar, QToolBar, QMenuBar, QMenu,
    QFileDialog, QMessageBox, QProgressDialog, QLabel
)
from PyQt6.QtCore import Qt, QSize
from PyQt6.QtGui import QAction, QIcon, QKeySequence

from src.views.file_panel import FilePanel
from src.views.config_panel import ConfigPanel
from src.views.diff_view import DiffView
from src.views.diff_list import DiffListPanel
from src.views.stats_panel import StatsPanel
from src.models.excel_model import WorkbookData
from src.models.diff_model import CompareResult
from src.services.compare_service import CompareMode, CompareOptions
from src.workers.compare_worker import CompareWorker


class MainWindow(QMainWindow):
    """主窗口"""
    
    def __init__(self):
        super().__init__()
        
        self.setWindowTitle("Excel 文件比较工具")
        self.setMinimumSize(1280, 720)
        self.resize(1600, 900)
        
        # 设置窗口图标
        icon_path = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 
                                  "resources", "icon.svg")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
        
        # 数据
        self._workbook_a: Optional[WorkbookData] = None
        self._workbook_b: Optional[WorkbookData] = None
        self._compare_result: Optional[CompareResult] = None
        self._compare_worker: Optional[CompareWorker] = None
        self._current_diff_index: int = -1
        
        # 初始化 UI
        self._setup_ui()
        self._setup_menu()
        self._setup_toolbar()
        self._setup_statusbar()
        self._connect_signals()
        
        # 应用样式
        self._apply_styles()
    
    def _setup_ui(self):
        """设置 UI 布局"""
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        main_layout = QHBoxLayout(central_widget)
        main_layout.setContentsMargins(8, 8, 8, 8)
        main_layout.setSpacing(8)
        
        # 左侧面板（配置 + 统计）
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(8)
        
        # 配置面板
        self.config_panel = ConfigPanel()
        left_layout.addWidget(self.config_panel, 1)  # 配置面板可伸缩
        
        # 弹性空间
        left_layout.addStretch()
        
        # 统计面板（放在底部）
        self.stats_panel = StatsPanel()
        left_layout.addWidget(self.stats_panel)
        
        left_panel.setMaximumWidth(280)
        left_panel.setMinimumWidth(220)
        
        # 右侧主区域
        right_splitter = QSplitter(Qt.Orientation.Vertical)
        
        # 上部：文件面板 + 表格视图
        top_widget = QWidget()
        top_layout = QVBoxLayout(top_widget)
        top_layout.setContentsMargins(0, 0, 0, 0)
        top_layout.setSpacing(8)
        
        # 文件选择面板
        file_panel_widget = QWidget()
        file_panel_layout = QHBoxLayout(file_panel_widget)
        file_panel_layout.setContentsMargins(0, 0, 0, 0)
        file_panel_layout.setSpacing(8)
        
        self.file_panel_a = FilePanel("文件 A")
        self.file_panel_b = FilePanel("文件 B")
        file_panel_layout.addWidget(self.file_panel_a)
        file_panel_layout.addWidget(self.file_panel_b)
        
        top_layout.addWidget(file_panel_widget)
        
        # 差异视图（双栏表格）
        self.diff_view = DiffView()
        top_layout.addWidget(self.diff_view, 1)
        
        right_splitter.addWidget(top_widget)
        
        # 下部：差异列表
        self.diff_list_panel = DiffListPanel()
        right_splitter.addWidget(self.diff_list_panel)
        
        # 设置分割比例
        right_splitter.setSizes([600, 200])
        
        # 组装布局
        main_layout.addWidget(left_panel)
        main_layout.addWidget(right_splitter, 1)
    
    def _setup_menu(self):
        """设置菜单栏"""
        menubar = self.menuBar()
        
        # 文件菜单
        file_menu = menubar.addMenu("文件(&F)")
        
        open_a_action = QAction("打开文件 A...", self)
        open_a_action.setShortcut(QKeySequence("Ctrl+O"))
        open_a_action.triggered.connect(lambda: self._open_file('a'))
        file_menu.addAction(open_a_action)
        
        open_b_action = QAction("打开文件 B...", self)
        open_b_action.setShortcut(QKeySequence("Ctrl+Shift+O"))
        open_b_action.triggered.connect(lambda: self._open_file('b'))
        file_menu.addAction(open_b_action)
        
        file_menu.addSeparator()
        
        export_action = QAction("导出报告...", self)
        export_action.setShortcut(QKeySequence("Ctrl+E"))
        export_action.triggered.connect(self._export_report)
        file_menu.addAction(export_action)
        
        file_menu.addSeparator()
        
        exit_action = QAction("退出(&X)", self)
        exit_action.setShortcut(QKeySequence("Alt+F4"))
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)
        
        # 编辑菜单
        edit_menu = menubar.addMenu("编辑(&E)")
        
        compare_action = QAction("开始比较", self)
        compare_action.setShortcut(QKeySequence("F5"))
        compare_action.triggered.connect(self._start_compare)
        edit_menu.addAction(compare_action)
        
        # 查看菜单
        view_menu = menubar.addMenu("查看(&V)")
        
        prev_diff_action = QAction("上一个差异", self)
        prev_diff_action.setShortcut(QKeySequence("Shift+F3"))
        prev_diff_action.triggered.connect(self._prev_diff)
        view_menu.addAction(prev_diff_action)
        
        next_diff_action = QAction("下一个差异", self)
        next_diff_action.setShortcut(QKeySequence("F3"))
        next_diff_action.triggered.connect(self._next_diff)
        view_menu.addAction(next_diff_action)
        
        # 帮助菜单
        help_menu = menubar.addMenu("帮助(&H)")
        
        about_action = QAction("关于", self)
        about_action.triggered.connect(self._show_about)
        help_menu.addAction(about_action)
    
    def _setup_toolbar(self):
        """设置工具栏"""
        toolbar = QToolBar("主工具栏")
        toolbar.setIconSize(QSize(24, 24))
        toolbar.setMovable(False)
        self.addToolBar(toolbar)
        
        # 打开文件 A
        open_a_action = QAction("打开 A", self)
        open_a_action.setToolTip("打开文件 A (Ctrl+O)")
        open_a_action.triggered.connect(lambda: self._open_file('a'))
        toolbar.addAction(open_a_action)
        
        # 打开文件 B
        open_b_action = QAction("打开 B", self)
        open_b_action.setToolTip("打开文件 B (Ctrl+Shift+O)")
        open_b_action.triggered.connect(lambda: self._open_file('b'))
        toolbar.addAction(open_b_action)
        
        toolbar.addSeparator()
        
        # 开始比较
        self.compare_action = QAction("开始比较", self)
        self.compare_action.setToolTip("开始比较 (F5)")
        self.compare_action.triggered.connect(self._start_compare)
        toolbar.addAction(self.compare_action)
        
        toolbar.addSeparator()
        
        # 差异导航
        self.prev_action = QAction("◀ 上一个", self)
        self.prev_action.setToolTip("上一个差异 (Shift+F3)")
        self.prev_action.triggered.connect(self._prev_diff)
        self.prev_action.setEnabled(False)
        toolbar.addAction(self.prev_action)
        
        self.next_action = QAction("下一个 ▶", self)
        self.next_action.setToolTip("下一个差异 (F3)")
        self.next_action.triggered.connect(self._next_diff)
        self.next_action.setEnabled(False)
        toolbar.addAction(self.next_action)
        
        # 差异位置标签
        self.diff_position_label = QLabel(" 0/0 ")
        toolbar.addWidget(self.diff_position_label)
        
        toolbar.addSeparator()
        
        # 导出报告
        export_action = QAction("导出报告", self)
        export_action.setToolTip("导出比较报告 (Ctrl+E)")
        export_action.triggered.connect(self._export_report)
        toolbar.addAction(export_action)
    
    def _setup_statusbar(self):
        """设置状态栏"""
        self.statusbar = QStatusBar()
        self.setStatusBar(self.statusbar)
        self.statusbar.showMessage("就绪")
    
    def _connect_signals(self):
        """连接信号"""
        # 文件面板信号
        self.file_panel_a.file_dropped.connect(lambda p: self._load_file(p, 'a'))
        self.file_panel_b.file_dropped.connect(lambda p: self._load_file(p, 'b'))
        
        # 配置面板信号
        self.config_panel.compare_clicked.connect(self._start_compare)
        self.config_panel.smart_compare_clicked.connect(self._start_smart_compare)
        
        # 差异列表信号
        self.diff_list_panel.diff_selected.connect(self._on_diff_selected)
        
        # 选区比较信号
        self.diff_view.compare_selection_clicked.connect(self._compare_selection)
    
    def _apply_styles(self):
        """应用样式"""
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f5f5f5;
            }
            QToolBar {
                background-color: #ffffff;
                border-bottom: 1px solid #e0e0e0;
                padding: 4px;
                spacing: 4px;
            }
            QToolBar QToolButton {
                padding: 6px 12px;
                border-radius: 4px;
                font-size: 13px;
            }
            QToolBar QToolButton:hover {
                background-color: #e3f2fd;
            }
            QToolBar QToolButton:pressed {
                background-color: #bbdefb;
            }
            QStatusBar {
                background-color: #ffffff;
                border-top: 1px solid #e0e0e0;
            }
            QMenuBar {
                background-color: #ffffff;
                border-bottom: 1px solid #e0e0e0;
            }
            QMenuBar::item:selected {
                background-color: #e3f2fd;
            }
            QMenu {
                background-color: #ffffff;
                border: 1px solid #e0e0e0;
            }
            QMenu::item:selected {
                background-color: #e3f2fd;
            }
        """)
    
    def _open_file(self, which: str):
        """打开文件对话框"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            f"选择文件 {which.upper()}",
            "",
            "Excel 文件 (*.xlsx *.xls);;所有文件 (*.*)"
        )
        if file_path:
            self._load_file(file_path, which)
    
    def _load_file(self, file_path: str, which: str):
        """加载文件"""
        try:
            from src.services.excel_service import ExcelService
            workbook = ExcelService.load_file(file_path)
            
            if which == 'a':
                self._workbook_a = workbook
                self.file_panel_a.set_file_info(workbook)
                self.statusbar.showMessage(f"已加载文件 A: {workbook.file_name}")
            else:
                self._workbook_b = workbook
                self.file_panel_b.set_file_info(workbook)
                self.statusbar.showMessage(f"已加载文件 B: {workbook.file_name}")
            
            # 更新配置面板的 sheet 列表
            self._update_sheet_list()
            
            # 预览文件内容
            self._preview_files()
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"无法加载文件:\n{str(e)}")
    
    def _preview_files(self):
        """预览文件内容（加载后立即显示）"""
        if self._workbook_a or self._workbook_b:
            # 使用空差异列表显示预览
            self.diff_view.set_data(
                self._workbook_a,
                self._workbook_b,
                []  # 空差异列表，仅显示内容预览
            )
            if self._workbook_a and self._workbook_b:
                self.statusbar.showMessage("已加载两个文件，可查看预览或开始比较")
    
    def _update_sheet_list(self):
        """更新 sheet 列表"""
        sheets_a = self._workbook_a.sheet_names if self._workbook_a else []
        sheets_b = self._workbook_b.sheet_names if self._workbook_b else []
        all_sheets = list(set(sheets_a + sheets_b))
        self.config_panel.set_sheet_list(all_sheets)
    
    def _start_compare(self):
        """开始比较"""
        if not self._workbook_a or not self._workbook_b:
            QMessageBox.warning(self, "提示", "请先加载两个要比较的 Excel 文件")
            return
        
        # 获取配置
        mode = self.config_panel.get_compare_mode()
        options = self.config_panel.get_compare_options()
        selected_sheets = self.config_panel.get_selected_sheets()
        key_config = self.config_panel.get_key_column_config()  # 主键列配置
        key_col1_a, key_col2_a = key_config['a']  # A文件主键列
        key_col1_b, key_col2_b = key_config['b']  # B文件主键列
        header_row = self.config_panel.get_header_row_config()  # 首行匹配列

        # 显示进度对话框
        progress = QProgressDialog("正在比较...", "取消", 0, 100, self)
        progress.setWindowTitle("比较中")
        progress.setWindowModality(Qt.WindowModality.WindowModal)
        progress.setMinimumDuration(0)
        progress.setValue(0)

        try:
            progress.setValue(30)
            progress.setLabelText("正在分析差异...")

            # 构建匹配模式描述
            mode_parts = []
            if key_col1_a is not None:
                if key_col2_a is not None:
                    mode_parts.append(f"A文件主键:{key_col1_a + 1}+{key_col2_a + 1}")
                else:
                    mode_parts.append(f"A文件主键:{key_col1_a + 1}")
            if key_col1_b is not None:
                if key_col2_b is not None:
                    mode_parts.append(f"B文件主键:{key_col1_b + 1}+{key_col2_b + 1}")
                else:
                    mode_parts.append(f"B文件主键:{key_col1_b + 1}")
            if header_row is not None:
                mode_parts.append(f"标题行:{header_row + 1}")

            if key_col1_a is not None or header_row is not None:
                # 使用智能匹配方式比较
                result = self._compare_with_smart_match(key_col1_a, key_col2_a, key_col1_b, key_col2_b, header_row, options, selected_sheets)
                mode_desc = "智能匹配 (" + ", ".join(mode_parts) + ")"
            else:
                # 使用标准比较服务
                from src.services.compare_service import CompareService
                result = CompareService.compare(
                    self._workbook_a,
                    self._workbook_b,
                    mode=mode,
                    options=options,
                    selected_sheets=selected_sheets if selected_sheets else None
                )
                mode_desc = "按位置"
            
            # 设置比较配置（用于报告记录）
            result.compare_config = {
                'mode': mode_desc,
                'key_column': key_col1_a,
                'key_column2': key_col2_a,
                'header_row': header_row,
                'ignore_case': options.ignore_case,
                'ignore_whitespace': options.ignore_whitespace,
                'ignore_format': options.ignore_format,
                'ignore_empty_rows': options.ignore_empty_rows,
            }
            
            progress.setValue(100)
            self._compare_result = result
            self._current_diff_index = 0 if result.diffs else -1
            
            # 更新 UI
            self._update_compare_result()
            self.statusbar.showMessage(f"比较完成 ({mode_desc})，共发现 {result.summary.total} 处差异")
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"比较过程中发生错误:\n{str(e)}")
        finally:
            progress.close()

    def _compare_with_smart_match(self, key_col1_a: Optional[int], key_col2_a: Optional[int], key_col1_b: Optional[int], key_col2_b: Optional[int], header_row: Optional[int], options, selected_sheets):
        """使用智能匹配方式比较（支持A文件和B文件分别指定主键列）"""
        from src.models.diff_model import DiffResult, DiffSummary, DiffType, CompareResult

        print("\n" + "="*80)
        print("【调试】开始智能匹配比较")
        print(f"【调试】A文件主键列: key_col1_a={key_col1_a}, key_col2_a={key_col2_a}")
        print(f"【调试】B文件主键列: key_col1_b={key_col1_b}, key_col2_b={key_col2_b}")
        print(f"【调试】标题行: header_row={header_row}")
        print("="*80 + "\n")

        all_diffs = []

        # 获取要比较的工作表
        sheets_a = self._workbook_a.sheet_names
        sheets_b = self._workbook_b.sheet_names
        
        if selected_sheets:
            sheets_to_compare = [s for s in selected_sheets if s in sheets_a and s in sheets_b]
        else:
            sheets_to_compare = [s for s in sheets_a if s in sheets_b]
        
        for sheet_name in sheets_to_compare:
            sheet_a = self._workbook_a.get_sheet(sheet_name)
            sheet_b = self._workbook_b.get_sheet(sheet_name)
            
            if not sheet_a or not sheet_b:
                continue
            
            # 构建列映射（首行匹配列）
            col_map_b_to_a = {}  # B的列索引 -> A的列索引
            col_map_a_to_b = {}  # A的列索引 -> B的列索引

            if header_row is not None:
                print(f"【调试】工作表 '{sheet_name}': 启用标题行匹配，标题行索引={header_row}")
                # 获取标题行
                headers_a = {}
                headers_b = {}
                for col_idx in range(sheet_a.col_count):
                    cell = sheet_a.get_cell(header_row, col_idx)
                    val = cell.value if cell else None
                    if val is not None and str(val).strip() != "":
                        key = str(val).strip().lower()  # 标题匹配始终忽略大小写
                        headers_a[key] = col_idx

                for col_idx in range(sheet_b.col_count):
                    cell = sheet_b.get_cell(header_row, col_idx)
                    val = cell.value if cell else None
                    if val is not None and str(val).strip() != "":
                        key = str(val).strip().lower()  # 标题匹配始终忽略大小写
                        headers_b[key] = col_idx

                print(f"【调试】A文件标题: {headers_a}")
                print(f"【调试】B文件标题: {headers_b}")

                # 建立列映射
                for header, col_a in headers_a.items():
                    if header in headers_b:
                        col_b = headers_b[header]
                        col_map_a_to_b[col_a] = col_b
                        col_map_b_to_a[col_b] = col_a

                print(f"【调试】列映射 A->B: {col_map_a_to_b}")
                print(f"【调试】列映射 B->A: {col_map_b_to_a}")
            else:
                print(f"【调试】工作表 '{sheet_name}': 未启用标题行匹配")

            # 提取行数据
            def extract_row_data(sheet, row_idx, use_col_map=None):
                """提取一行数据，可按列映射重排"""
                row_data = []
                for col_idx in range(sheet.col_count):
                    cell = sheet.get_cell(row_idx, col_idx)
                    row_data.append(cell.value if cell else None)
                return row_data

            if key_col1_a is not None:
                print(f"【调试】使用主键列匹配行")
                # 使用主键列匹配行（A文件和B文件分别指定主键列）
                rows_a = {}
                rows_b = {}

                def make_key(row_data, col1, col2, col_map=None):
                    """生成复合主键"""
                    actual_col1 = col_map.get(col1, col1) if col_map else col1
                    val1 = row_data[actual_col1] if actual_col1 < len(row_data) else None
                    if val1 is None or str(val1).strip() == "":
                        return None
                    key = str(val1).strip()
                    if col2 is not None:
                        actual_col2 = col_map.get(col2, col2) if col_map else col2
                        val2 = row_data[actual_col2] if actual_col2 < len(row_data) else None
                        if val2 is not None and str(val2).strip() != "":
                            key += "|" + str(val2).strip()
                    if options.ignore_case:
                        key = key.lower()
                    return key

                print(f"【调试】开始提取A文件主键（从列索引 {key_col1_a}, {key_col2_a}）")
                for row_idx in range(sheet_a.row_count):
                    if header_row is not None and row_idx == header_row:
                        continue  # 跳过标题行
                    row_data = extract_row_data(sheet_a, row_idx)
                    norm_key = make_key(row_data, key_col1_a, key_col2_a)
                    if norm_key:
                        if norm_key not in rows_a:
                            rows_a[norm_key] = []
                        rows_a[norm_key].append((row_idx, row_data))

                print(f"【调试】A文件提取到 {len(rows_a)} 个唯一主键，前5个: {list(rows_a.keys())[:5]}")

                print(f"【调试】开始提取B文件主键（从列索引 {key_col1_b}, {key_col2_b}）")
                for row_idx in range(sheet_b.row_count):
                    if header_row is not None and row_idx == header_row:
                        continue
                    row_data = extract_row_data(sheet_b, row_idx)
                    # 主键列不使用列映射，始终从指定的列索引提取
                    norm_key = make_key(row_data, key_col1_b, key_col2_b)
                    if norm_key:
                        if norm_key not in rows_b:
                            rows_b[norm_key] = []
                        rows_b[norm_key].append((row_idx, row_data))

                print(f"【调试】B文件提取到 {len(rows_b)} 个唯一主键，前5个: {list(rows_b.keys())[:5]}")

                all_keys = set(rows_a.keys()) | set(rows_b.keys())
                print(f"【调试】总共 {len(all_keys)} 个唯一主键需要比较")

                for key in all_keys:
                    list_a = rows_a.get(key, [])
                    list_b = rows_b.get(key, [])

                    # 贪婪匹配：按位置距离最小的优先匹配
                    matches = []
                    used_a = set()
                    used_b = set()

                    # 计算所有可能的配对及其距离
                    pairs = []
                    for i, (row_idx_a, row_data_a) in enumerate(list_a):
                        for j, (row_idx_b, row_data_b) in enumerate(list_b):
                            dist = abs(row_idx_a - row_idx_b)
                            pairs.append((dist, i, j))

                    # 按距离排序
                    pairs.sort(key=lambda x: x[0])

                    # 贪婪选择
                    for _, i, j in pairs:
                        if i not in used_a and j not in used_b:
                            matches.append((list_a[i], list_b[j]))
                            used_a.add(i)
                            used_b.add(j)

                    # 处理匹配的行
                    match_count = 0
                    for (row_idx_a, row_data_a), (row_idx_b, row_data_b) in matches:
                        match_count += 1
                        if match_count == 1:  # 只打印第一个匹配的详细信息
                            print(f"【调试】第一个匹配: A行{row_idx_a} <-> B行{row_idx_b}")

                        # 按列映射或按位置比较
                        if col_map_a_to_b:
                            if match_count == 1:
                                print(f"【调试】使用列映射比较，共 {len(col_map_a_to_b)} 个列映射")
                            for col_a, col_b in col_map_a_to_b.items():
                                # 跳过主键列（A文件和B文件的主键列都要跳过）
                                skip_reason = None
                                if col_a == key_col1_a:
                                    skip_reason = f"col_a({col_a})==key_col1_a({key_col1_a})"
                                elif col_a == key_col2_a:
                                    skip_reason = f"col_a({col_a})==key_col2_a({key_col2_a})"
                                elif col_b == key_col1_b:
                                    skip_reason = f"col_b({col_b})==key_col1_b({key_col1_b})"
                                elif col_b == key_col2_b:
                                    skip_reason = f"col_b({col_b})==key_col2_b({key_col2_b})"

                                if skip_reason:
                                    if match_count == 1:
                                        print(f"【调试】  跳过列映射 ({col_a}->{col_b}): {skip_reason}")
                                    continue

                                if match_count == 1:
                                    print(f"【调试】  比较列映射 ({col_a}->{col_b})")

                                val_a = row_data_a[col_a] if col_a < len(row_data_a) else None
                                val_b = row_data_b[col_b] if col_b < len(row_data_b) else None
                                if self._values_differ(val_a, val_b, options):
                                    diff_type = self._get_diff_type(val_a, val_b)
                                    if match_count == 1 or col_a == 0:  # 打印第一个匹配或主键列的差异
                                        print(f"【调试】  发现差异: A行{row_idx_a}列{col_a} vs B行{row_idx_b}列{col_b}, 值: {val_a} -> {val_b}")
                                    all_diffs.append(DiffResult(
                                        sheet=sheet_name, row=row_idx_a, col=col_a,
                                        diff_type=diff_type, old_value=val_a, new_value=val_b,
                                        row_b=row_idx_b, col_b=col_b
                                    ))
                        else:
                            if match_count == 1:
                                print(f"【调试】使用按位置比较")
                            max_cols = max(len(row_data_a), len(row_data_b))
                            for col_idx in range(max_cols):
                                # 跳过主键列
                                skip_reason = None
                                if col_idx == key_col1_a:
                                    skip_reason = f"col_idx({col_idx})==key_col1_a({key_col1_a})"
                                elif col_idx == key_col2_a:
                                    skip_reason = f"col_idx({col_idx})==key_col2_a({key_col2_a})"
                                elif col_idx == key_col1_b:
                                    skip_reason = f"col_idx({col_idx})==key_col1_b({key_col1_b})"
                                elif col_idx == key_col2_b:
                                    skip_reason = f"col_idx({col_idx})==key_col2_b({key_col2_b})"

                                if skip_reason:
                                    if match_count == 1:
                                        print(f"【调试】  跳过列索引 {col_idx}: {skip_reason}")
                                    continue

                                val_a = row_data_a[col_idx] if col_idx < len(row_data_a) else None
                                val_b = row_data_b[col_idx] if col_idx < len(row_data_b) else None
                                if self._values_differ(val_a, val_b, options):
                                    diff_type = self._get_diff_type(val_a, val_b)
                                    all_diffs.append(DiffResult(
                                        sheet=sheet_name, row=row_idx_a, col=col_idx,
                                        diff_type=diff_type, old_value=val_a, new_value=val_b,
                                        row_b=row_idx_b, col_b=col_idx
                                    ))

                    # 处理未匹配的A（删除整行）
                    for i, (row_idx_a, row_data_a) in enumerate(list_a):
                        if i not in used_a:
                            print(f"【调试】未匹配的A行（删除整行）: 行{row_idx_a}, 主键={key}")
                            for col_idx, val in enumerate(row_data_a):
                                if val is not None and str(val).strip() != "":
                                    all_diffs.append(DiffResult(
                                        sheet=sheet_name, row=row_idx_a, col=col_idx,
                                        diff_type=DiffType.DELETED, old_value=val
                                    ))

                    # 处理未匹配的B（新增整行）
                    for j, (row_idx_b, row_data_b) in enumerate(list_b):
                        if j not in used_b:
                            print(f"【调试】未匹配的B行（新增整行）: 行{row_idx_b}, 主键={key}")
                            for col_idx, val in enumerate(row_data_b):
                                if val is not None and str(val).strip() != "":
                                    all_diffs.append(DiffResult(
                                        sheet=sheet_name, row=row_idx_b, col=col_idx,
                                        diff_type=DiffType.ADDED, new_value=val,
                                        row_b=row_idx_b, col_b=col_idx
                                    ))
            else:
                # 只使用首行匹配列，按位置匹配行
                max_rows = max(sheet_a.row_count, sheet_b.row_count)
                for row_idx in range(max_rows):
                    if header_row is not None and row_idx == header_row:
                        continue
                    
                    row_data_a = extract_row_data(sheet_a, row_idx) if row_idx < sheet_a.row_count else []
                    row_data_b = extract_row_data(sheet_b, row_idx) if row_idx < sheet_b.row_count else []
                    
                    if col_map_a_to_b:
                        for col_a, col_b in col_map_a_to_b.items():
                            val_a = row_data_a[col_a] if col_a < len(row_data_a) else None
                            val_b = row_data_b[col_b] if col_b < len(row_data_b) else None
                            if self._values_differ(val_a, val_b, options):
                                diff_type = self._get_diff_type(val_a, val_b)
                                all_diffs.append(DiffResult(
                                    sheet=sheet_name, row=row_idx, col=col_a,
                                    diff_type=diff_type, old_value=val_a, new_value=val_b,
                                    row_b=row_idx, col_b=col_b
                                ))
        
        # 创建结果
        summary = DiffSummary()
        for diff in all_diffs:
            summary.add_diff(diff.diff_type)

        print("\n" + "="*80)
        print(f"【调试】比较完成，共发现 {len(all_diffs)} 处差异")
        print(f"【调试】差异统计: 新增={summary.added}, 删除={summary.deleted}, 修改={summary.modified}")
        print("="*80 + "\n")

        return CompareResult(
            file_a=self._workbook_a.file_name,
            file_b=self._workbook_b.file_name,
            diffs=all_diffs,
            summary=summary
        )
    
    def _update_compare_result(self):
        """更新比较结果显示"""
        if not self._compare_result:
            return
        
        result = self._compare_result
        
        # 更新统计面板
        self.stats_panel.set_summary(result.summary)
        
        # 更新差异列表
        self.diff_list_panel.set_diffs(result.diffs)
        
        # 更新差异视图
        self.diff_view.set_data(self._workbook_a, self._workbook_b, result.diffs)
        
        # 更新导航按钮
        has_diffs = len(result.diffs) > 0
        self.prev_action.setEnabled(has_diffs)
        self.next_action.setEnabled(has_diffs)
        self._update_diff_position()
    
    def _update_diff_position(self):
        """更新差异位置显示"""
        if not self._compare_result or not self._compare_result.diffs:
            self.diff_position_label.setText(" 0/0 ")
            return
        
        total = len(self._compare_result.diffs)
        current = self._current_diff_index + 1 if self._current_diff_index >= 0 else 0
        self.diff_position_label.setText(f" {current}/{total} ")
    
    def _prev_diff(self):
        """跳转到上一个差异"""
        if not self._compare_result or not self._compare_result.diffs:
            return
        
        self._current_diff_index -= 1
        if self._current_diff_index < 0:
            self._current_diff_index = len(self._compare_result.diffs) - 1
        
        self._navigate_to_current_diff()
    
    def _next_diff(self):
        """跳转到下一个差异"""
        if not self._compare_result or not self._compare_result.diffs:
            return
        
        self._current_diff_index += 1
        if self._current_diff_index >= len(self._compare_result.diffs):
            self._current_diff_index = 0
        
        self._navigate_to_current_diff()
    
    def _navigate_to_current_diff(self):
        """导航到当前差异"""
        if self._current_diff_index < 0:
            return
        
        diff = self._compare_result.diffs[self._current_diff_index]
        self.diff_view.scroll_to_diff(diff)
        self.diff_list_panel.select_diff(self._current_diff_index)
        self._update_diff_position()
    
    def _on_diff_selected(self, index: int):
        """差异列表选中事件"""
        self._current_diff_index = index
        self._navigate_to_current_diff()
    
    def _export_report(self):
        """导出报告"""
        if not self._compare_result:
            QMessageBox.warning(self, "提示", "请先执行比较操作")
            return
        
        file_path, selected_filter = QFileDialog.getSaveFileName(
            self,
            "导出报告",
            "compare_report",
            "Excel 文件 (*.xlsx);;HTML 文件 (*.html)"
        )
        
        if not file_path:
            return
        
        try:
            from src.services.report_service import ReportService
            
            if file_path.endswith('.html'):
                ReportService.export_html(
                    self._compare_result, 
                    self._workbook_a, 
                    self._workbook_b, 
                    file_path
                )
            else:
                if not file_path.endswith('.xlsx'):
                    file_path += '.xlsx'
                ReportService.export_excel(
                    self._compare_result,
                    self._workbook_a,
                    self._workbook_b,
                    file_path
                )
            
            QMessageBox.information(self, "成功", f"报告已导出到:\n{file_path}")
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"导出失败:\n{str(e)}")
    
    def _start_smart_compare(self):
        """开始智能比较"""
        if not self._workbook_a or not self._workbook_b:
            QMessageBox.warning(self, "提示", "请先加载两个要比较的 Excel 文件")
            return
        
        # 获取智能比较设置
        settings = self.config_panel.get_smart_compare_settings()
        options = self.config_panel.get_compare_options()
        selected_sheets = self.config_panel.get_selected_sheets()
        
        # 确定要比较的工作表
        if selected_sheets:
            sheet_name = selected_sheets[0]  # 智能比较每次只比较一个工作表
        else:
            # 使用第一个共同的工作表
            common_sheets = set(self._workbook_a.sheet_names) & set(self._workbook_b.sheet_names)
            if not common_sheets:
                QMessageBox.warning(self, "提示", "两个文件没有相同名称的工作表")
                return
            sheet_name = sorted(common_sheets)[0]
        
        # 显示进度
        progress = QProgressDialog("正在进行智能比较...", "取消", 0, 100, self)
        progress.setWindowTitle("智能比较")
        progress.setWindowModality(Qt.WindowModality.WindowModal)
        progress.setMinimumDuration(0)
        progress.setValue(0)
        
        try:
            from src.services.smart_compare_service import SmartCompareService, SmartCompareOptions, CellRange
            
            progress.setValue(30)
            progress.setLabelText("解析设置...")
            
            # 构建智能比较选项
            smart_options = SmartCompareOptions()
            
            # 解析区域
            if settings['range_str']:
                try:
                    cell_range = CellRange.from_string(settings['range_str'])
                    smart_options.range_a = cell_range
                    smart_options.range_b = cell_range  # 两个文件使用相同区域
                except ValueError as e:
                    QMessageBox.warning(self, "区域格式错误", str(e))
                    return
            
            # 设置标题和主键
            smart_options.use_header_row = settings['use_header']
            smart_options.use_key_column = settings['use_key']
            
            # 解析主键列
            if settings['use_key'] and settings['key_column']:
                key_col = settings['key_column'].strip().upper()
                if key_col.isdigit():
                    smart_options.key_column_index = int(key_col) - 1  # 转为0-indexed
                else:
                    # 字母列名
                    smart_options.key_column_index = 0
                    for char in key_col:
                        smart_options.key_column_index = smart_options.key_column_index * 26 + (ord(char) - ord('A') + 1)
                    smart_options.key_column_index -= 1
            
            # 传递忽略选项
            smart_options.ignore_case = options.ignore_case
            smart_options.ignore_whitespace = options.ignore_whitespace
            smart_options.ignore_empty_rows = options.ignore_empty_rows
            
            progress.setValue(50)
            progress.setLabelText("正在智能匹配...")
            
            # 执行智能比较
            result = SmartCompareService.compare_with_range(
                self._workbook_a,
                self._workbook_b,
                sheet_name,
                smart_options
            )
            
            progress.setValue(100)
            self._compare_result = result
            self._current_diff_index = 0 if result.diffs else -1
            
            # 更新 UI
            self._update_compare_result()
            
            mode_desc = "基于主键列" if settings['use_key'] else ("基于列标题" if settings['use_header'] else "基于位置")
            self.statusbar.showMessage(f"智能比较完成 ({mode_desc})，共发现 {result.summary.total} 处差异")
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"智能比较时发生错误:\n{str(e)}")
        finally:
            progress.close()
    
    def _compare_selection(self):
        """比较选中的区域（支持主键列匹配）"""
        if not self._workbook_a or not self._workbook_b:
            QMessageBox.warning(self, "提示", "请先加载两个要比较的 Excel 文件")
            return

        # 获取选区
        sheet_name, range_a, _, range_b = self.diff_view.get_current_selections()

        if not range_a or not range_b:
            QMessageBox.warning(self, "提示", "请在两个表格中分别用鼠标拖拽选择要比较的区域")
            return

        # range 格式: (min_row, min_col, max_row, max_col)
        rows_a = range_a[2] - range_a[0] + 1
        cols_a = range_a[3] - range_a[1] + 1
        rows_b = range_b[2] - range_b[0] + 1
        cols_b = range_b[3] - range_b[1] + 1

        # 列数必须相同
        if cols_a != cols_b:
            QMessageBox.warning(
                self, "列数不匹配",
                f"文件A选区: {cols_a}列\n文件B选区: {cols_b}列\n\n两个选区的列数必须相同"
            )
            return

        # 从配置面板获取主键列和标题行设置
        key_config = self.config_panel.get_key_column_config()
        key_col1_a_abs, key_col2_a_abs = key_config['a']  # A文件绝对列索引
        key_col1_b_abs, key_col2_b_abs = key_config['b']  # B文件绝对列索引
        header_row_abs = self.config_panel.get_header_row_config()  # 绝对行索引

        # 转换为相对于选区的偏移
        key_col1 = None if key_col1_a_abs is None else (key_col1_a_abs - range_a[1])
        key_col2 = None if key_col2_a_abs is None else (key_col2_a_abs - range_a[1])

        # 标题行特殊处理：如果标题行不在选区内，从工作表读取
        header_row = None
        use_external_header = False
        if header_row_abs is not None:
            if header_row_abs < range_a[0] or header_row_abs > range_a[2]:
                # 标题行不在选区内，使用外部标题行
                use_external_header = True
                header_row = None  # 选区内不使用标题行
            else:
                # 标题行在选区内，转换为相对偏移
                header_row = header_row_abs - range_a[0]

        # 验证主键列是否在选区内
        if key_col1 is not None and (key_col1 < 0 or key_col1 >= cols_a):
            QMessageBox.warning(
                self, "主键列不在选区内",
                f"配置的主键列（第{key_col1_a_abs + 1}列）不在选区范围内\n"
                f"选区列范围: 第{range_a[1] + 1}列 到 第{range_a[3] + 1}列"
            )
            return

        # 如果没有指定主键列，行数也必须相同
        if key_col1 is None and rows_a != rows_b:
            QMessageBox.warning(
                self, "区域大小不匹配",
                f"文件A选区: {rows_a}行\n文件B选区: {rows_b}行\n\n"
                "按位置比较时行数必须相同。\n如果行顺序不同，请在配置面板中指定主键列作为匹配依据。"
            )
            return

        self.statusbar.showMessage("正在比较选中区域...")

        try:
            from src.models.diff_model import DiffResult, DiffSummary, DiffType, CompareResult

            sheet_a = self._workbook_a.get_sheet(sheet_name)
            sheet_b = self._workbook_b.get_sheet(sheet_name)

            if not sheet_a or not sheet_b:
                QMessageBox.warning(self, "错误", f"工作表 '{sheet_name}' 不存在")
                return

            diffs = []
            options = self.config_panel.get_compare_options()

            # 构建模式描述
            mode_parts = []
            if key_col1 is not None:
                if key_col2 is not None:
                    mode_parts.append(f"复合主键:{key_col1_a_abs + 1}+{key_col2_a_abs + 1}")
                else:
                    mode_parts.append(f"主键列:{key_col1_a_abs + 1}")
            if header_row_abs is not None:
                mode_parts.append(f"标题行:{header_row_abs + 1}")

            if key_col1 is not None or header_row is not None or use_external_header:
                # 使用增强的选区比较
                diffs = self._compare_selection_smart(
                    sheet_name, sheet_a, sheet_b,
                    range_a, range_b, key_col1, key_col2, header_row, options,
                    external_header_row=header_row_abs if use_external_header else None
                )
                mode_desc = "智能匹配 (" + ", ".join(mode_parts) + ")"
            else:
                # 按位置比较
                diffs = self._compare_by_position(
                    sheet_name, sheet_a, sheet_b,
                    range_a, range_b, options
                )
                mode_desc = "按位置"

            # 创建结果
            summary = DiffSummary()
            for diff in diffs:
                summary.add_diff(diff.diff_type)

            result = CompareResult(
                file_a=self._workbook_a.file_name,
                file_b=self._workbook_b.file_name,
                diffs=diffs,
                summary=summary
            )

            # 设置比较配置（用于报告记录）
            result.compare_config = {
                'mode': mode_desc,
                'key_column': key_col1_a_abs,
                'key_column2': key_col2_a_abs,
                'header_row': header_row_abs,
                'ignore_case': options.ignore_case,
                'ignore_whitespace': options.ignore_whitespace,
                'selection_a': self._format_range(range_a),
                'selection_b': self._format_range(range_b),
            }

            self._compare_result = result
            self._current_diff_index = 0 if diffs else -1

            # 更新 UI
            self._update_compare_result()

            self.statusbar.showMessage(
                f"选区比较完成 ({mode_desc}): 文件A [{self._format_range(range_a)}] vs 文件B [{self._format_range(range_b)}]，"
                f"共发现 {summary.total} 处差异"
            )

        except Exception as e:
            QMessageBox.critical(self, "错误", f"选区比较时发生错误:\n{str(e)}")
    
    def _compare_selection_smart(self, sheet_name, sheet_a, sheet_b, range_a, range_b, key_col1, key_col2, header_row, options, external_header_row=None):
        """选区智能比较（支持复合主键列匹配行 + 标题行匹配列）

        Args:
            external_header_row: 外部标题行的绝对行索引，用于标题行不在选区内的情况
        """
        from src.models.diff_model import DiffResult, DiffType

        diffs = []

        # 提取选区数据
        def extract_range_data(sheet, rng):
            data = []
            for row_idx in range(rng[0], rng[2] + 1):
                row_data = []
                for col_idx in range(rng[1], rng[3] + 1):
                    cell = sheet.get_cell(row_idx, col_idx)
                    row_data.append(cell.value if cell else None)
                data.append((row_idx, row_data))
            return data

        data_a = extract_range_data(sheet_a, range_a)
        data_b = extract_range_data(sheet_b, range_b)

        # 构建列映射（标题行匹配列）
        col_map_a_to_b = {}  # A的相对列索引 -> B的相对列索引
        data_start_offset = 0  # 数据开始行偏移（跳过标题行）

        # 处理外部标题行
        if external_header_row is not None:
            headers_a = {}
            headers_b = {}

            # 从工作表读取标题行
            for col_offset in range(range_a[3] - range_a[1] + 1):
                col_idx_a = range_a[1] + col_offset
                cell_a = sheet_a.get_cell(external_header_row, col_idx_a)
                val_a = cell_a.value if cell_a else None
                if val_a is not None and str(val_a).strip() != "":
                    key = str(val_a).strip().lower()
                    headers_a[key] = col_offset

                col_idx_b = range_b[1] + col_offset
                cell_b = sheet_b.get_cell(external_header_row, col_idx_b)
                val_b = cell_b.value if cell_b else None
                if val_b is not None and str(val_b).strip() != "":
                    key = str(val_b).strip().lower()
                    headers_b[key] = col_offset

            for header, col_a in headers_a.items():
                if header in headers_b:
                    col_map_a_to_b[col_a] = headers_b[header]

            data_start_offset = 0  # 外部标题行不占用选区行
        elif header_row is not None and header_row < len(data_a) and header_row < len(data_b):
            headers_a = {}
            headers_b = {}
            _, header_row_a = data_a[header_row]
            _, header_row_b = data_b[header_row]
            
            for col_offset, val in enumerate(header_row_a):
                if val is not None and str(val).strip() != "":
                    key = str(val).strip().lower()  # 标题匹配始终忽略大小写
                    headers_a[key] = col_offset

            for col_offset, val in enumerate(header_row_b):
                if val is not None and str(val).strip() != "":
                    key = str(val).strip().lower()  # 标题匹配始终忽略大小写
                    headers_b[key] = col_offset
            
            for header, col_a in headers_a.items():
                if header in headers_b:
                    col_map_a_to_b[col_a] = headers_b[header]
            
            data_start_offset = header_row + 1  # 跳过标题行及之前的行
        
        if key_col1 is not None:
            # 主键列匹配行（支持复合主键）
            rows_a = {}
            rows_b = {}
            
            def make_key(row_data, col1, col2, col_map=None):
                """生成复合主键"""
                actual_col1 = col_map.get(col1, col1) if col_map else col1
                val1 = row_data[actual_col1] if actual_col1 < len(row_data) else None
                if val1 is None or str(val1).strip() == "":
                    return None
                key = str(val1).strip()
                if col2 is not None:
                    actual_col2 = col_map.get(col2, col2) if col_map else col2
                    val2 = row_data[actual_col2] if actual_col2 < len(row_data) else None
                    if val2 is not None and str(val2).strip() != "":
                        key += "|" + str(val2).strip()
                if options.ignore_case:
                    key = key.lower()
                return key
            
            for row_idx, row_data in data_a[data_start_offset:]:
                norm_key = make_key(row_data, key_col1, key_col2)
                if norm_key:
                    if norm_key not in rows_a:
                        rows_a[norm_key] = []
                    rows_a[norm_key].append((row_idx, row_data))
            
            for row_idx, row_data in data_b[data_start_offset:]:
                # 主键列不使用列映射，始终从指定的列索引提取
                norm_key = make_key(row_data, key_col1, key_col2)
                if norm_key:
                    if norm_key not in rows_b:
                        rows_b[norm_key] = []
                    rows_b[norm_key].append((row_idx, row_data))

            
            all_keys = set(rows_a.keys()) | set(rows_b.keys())
            
            for key in all_keys:
                list_a = rows_a.get(key, [])
                list_b = rows_b.get(key, [])
                
                # 贪婪匹配：按位置距离最小的优先匹配
                matches = []
                used_a = set()
                used_b = set()
                
                # 计算所有可能的配对及其距离
                pairs = []
                for i, (row_idx_a, row_data_a) in enumerate(list_a):
                    for j, (row_idx_b, row_data_b) in enumerate(list_b):
                        dist = abs(row_idx_a - row_idx_b)
                        pairs.append((dist, i, j))
                
                # 按距离排序
                pairs.sort(key=lambda x: x[0])
                
                # 贪婪选择
                for _, i, j in pairs:
                    if i not in used_a and j not in used_b:
                        matches.append((list_a[i], list_b[j]))
                        used_a.add(i)
                        used_b.add(j)
                
                # 处理匹配的行
                for (row_idx_a, row_data_a), (row_idx_b, row_data_b) in matches:
                    if col_map_a_to_b:
                        for col_a, col_b in col_map_a_to_b.items():
                            # 跳过主键列
                            if col_a == key_col1 or col_a == key_col2:
                                continue
                            val_a = row_data_a[col_a] if col_a < len(row_data_a) else None
                            val_b = row_data_b[col_b] if col_b < len(row_data_b) else None
                            if self._values_differ(val_a, val_b, options):
                                diff_type = self._get_diff_type(val_a, val_b)
                                diffs.append(DiffResult(
                                    sheet=sheet_name, row=row_idx_a, col=range_a[1] + col_a,
                                    diff_type=diff_type, old_value=val_a, new_value=val_b,
                                    row_b=row_idx_b, col_b=range_b[1] + col_b
                                ))
                    else:
                        for col_offset in range(max(len(row_data_a), len(row_data_b))):
                            # 跳过主键列
                            if col_offset == key_col1 or col_offset == key_col2:
                                continue
                            val_a = row_data_a[col_offset] if col_offset < len(row_data_a) else None
                            val_b = row_data_b[col_offset] if col_offset < len(row_data_b) else None
                            if self._values_differ(val_a, val_b, options):
                                diff_type = self._get_diff_type(val_a, val_b)
                                diffs.append(DiffResult(
                                    sheet=sheet_name, row=row_idx_a, col=range_a[1] + col_offset,
                                    diff_type=diff_type, old_value=val_a, new_value=val_b,
                                    row_b=row_idx_b, col_b=range_b[1] + col_offset
                                ))
                
                # 处理未匹配的A（删除）
                for i, (row_idx_a, row_data_a) in enumerate(list_a):
                    if i not in used_a:
                        for col_offset, val in enumerate(row_data_a):
                            if val is not None and str(val).strip() != "":
                                diffs.append(DiffResult(
                                    sheet=sheet_name, row=row_idx_a, col=range_a[1] + col_offset,
                                    diff_type=DiffType.DELETED, old_value=val
                                ))

                # 处理未匹配的B（新增）
                for j, (row_idx_b, row_data_b) in enumerate(list_b):
                    if j not in used_b:
                        for col_offset, val in enumerate(row_data_b):
                            if val is not None and str(val).strip() != "":
                                diffs.append(DiffResult(
                                    sheet=sheet_name, row=row_idx_b, col=range_b[1] + col_offset,
                                    diff_type=DiffType.ADDED, new_value=val,
                                    row_b=row_idx_b, col_b=range_b[1] + col_offset
                                ))
        else:
            # 只使用首行匹配列，按位置匹配行
            for i in range(data_start_offset, max(len(data_a), len(data_b))):
                row_idx_a, row_data_a = data_a[i] if i < len(data_a) else (range_a[0] + i, [])
                row_idx_b, row_data_b = data_b[i] if i < len(data_b) else (range_b[0] + i, [])
                
                if col_map_a_to_b:
                    for col_a, col_b in col_map_a_to_b.items():
                        val_a = row_data_a[col_a] if col_a < len(row_data_a) else None
                        val_b = row_data_b[col_b] if col_b < len(row_data_b) else None
                        if self._values_differ(val_a, val_b, options):
                            diff_type = self._get_diff_type(val_a, val_b)
                            diffs.append(DiffResult(
                                sheet=sheet_name, row=row_idx_a, col=range_a[1] + col_a,
                                diff_type=diff_type, old_value=val_a, new_value=val_b,
                                row_b=row_idx_b, col_b=range_b[1] + col_b
                            ))
        
        return diffs
    
    def _compare_by_position(self, sheet_name, sheet_a, sheet_b, range_a, range_b, options):
        """按位置比较选区"""
        from src.models.diff_model import DiffResult, DiffType
        
        diffs = []
        rows = range_a[2] - range_a[0] + 1
        cols = range_a[3] - range_a[1] + 1
        
        for row_offset in range(rows):
            for col_offset in range(cols):
                row_a = range_a[0] + row_offset
                col_a = range_a[1] + col_offset
                row_b = range_b[0] + row_offset
                col_b = range_b[1] + col_offset
                
                cell_a = sheet_a.get_cell(row_a, col_a)
                cell_b = sheet_b.get_cell(row_b, col_b)
                
                val_a = cell_a.value if cell_a else None
                val_b = cell_b.value if cell_b else None
                
                if self._values_differ(val_a, val_b, options):
                    diff_type = self._get_diff_type(val_a, val_b)
                    diffs.append(DiffResult(
                        sheet=sheet_name, row=row_a, col=col_a,
                        diff_type=diff_type, old_value=val_a, new_value=val_b,
                        row_b=row_b, col_b=col_b  # 记录B的位置
                    ))
        return diffs
    
    def _compare_by_key_column(self, sheet_name, sheet_a, sheet_b, range_a, range_b, key_col, options):
        """基于主键列匹配行进行比较"""
        from src.models.diff_model import DiffResult, DiffType
        
        diffs = []
        cols = range_a[3] - range_a[1] + 1
        
        # 提取选区数据并建立主键映射
        def extract_rows(sheet, rng):
            rows = {}
            for row_idx in range(rng[0], rng[2] + 1):
                row_data = []
                for col_idx in range(rng[1], rng[3] + 1):
                    cell = sheet.get_cell(row_idx, col_idx)
                    row_data.append(cell.value if cell else None)
                
                # 主键值（相对于选区起始列的偏移）
                if key_col < len(row_data):
                    key = row_data[key_col]
                    if key is not None and str(key).strip() != "":
                        # 标准化键值
                        if options.ignore_case and isinstance(key, str):
                            key = key.lower()
                        if options.ignore_whitespace and isinstance(key, str):
                            key = key.strip()
                        rows[key] = (row_idx, row_data)
            return rows
        
        rows_a = extract_rows(sheet_a, range_a)
        rows_b = extract_rows(sheet_b, range_b)
        
        all_keys = set(rows_a.keys()) | set(rows_b.keys())
        
        for key in all_keys:
            data_a = rows_a.get(key)
            data_b = rows_b.get(key)
            
            if data_a is None:
                # 文件B中新增的行
                row_idx_b, row_data = data_b
                for col_offset, val in enumerate(row_data):
                    if val is not None and str(val).strip() != "":
                        diffs.append(DiffResult(
                            sheet=sheet_name, row=row_idx_b, col=range_b[1] + col_offset,
                            diff_type=DiffType.ADDED, new_value=val,
                            row_b=row_idx_b, col_b=range_b[1] + col_offset
                        ))
            elif data_b is None:
                # 文件A中删除的行
                row_idx_a, row_data = data_a
                for col_offset, val in enumerate(row_data):
                    if val is not None and str(val).strip() != "":
                        diffs.append(DiffResult(
                            sheet=sheet_name, row=row_idx_a, col=range_a[1] + col_offset,
                            diff_type=DiffType.DELETED, old_value=val
                        ))
            else:
                # 两边都有，逐列比较
                row_idx_a, row_data_a = data_a
                row_idx_b, row_data_b = data_b

                for col_offset in range(min(len(row_data_a), len(row_data_b))):
                    # 跳过主键列
                    if col_offset == key_col:
                        continue
                    val_a = row_data_a[col_offset]
                    val_b = row_data_b[col_offset]

                    if self._values_differ(val_a, val_b, options):
                        diff_type = self._get_diff_type(val_a, val_b)
                        diffs.append(DiffResult(
                            sheet=sheet_name, row=row_idx_a, col=range_a[1] + col_offset,
                            diff_type=diff_type, old_value=val_a, new_value=val_b,
                            row_b=row_idx_b, col_b=range_b[1] + col_offset  # 记录B的位置
                        ))
        
        return diffs
    
    def _values_differ(self, val_a, val_b, options):
        """判断两个值是否不同（应用忽略选项）"""
        cmp_a, cmp_b = val_a, val_b
        if options.ignore_case:
            if isinstance(cmp_a, str): cmp_a = cmp_a.lower()
            if isinstance(cmp_b, str): cmp_b = cmp_b.lower()
        if options.ignore_whitespace:
            if isinstance(cmp_a, str): cmp_a = cmp_a.strip()
            if isinstance(cmp_b, str): cmp_b = cmp_b.strip()
        if options.ignore_empty_rows:
            if (cmp_a is None or cmp_a == "") and (cmp_b is None or cmp_b == ""):
                return False
        return cmp_a != cmp_b
    
    def _get_diff_type(self, val_a, val_b):
        """获取差异类型"""
        from src.models.diff_model import DiffType
        if (val_a is None or val_a == "") and val_b:
            return DiffType.ADDED
        elif val_a and (val_b is None or val_b == ""):
            return DiffType.DELETED
        return DiffType.MODIFIED
    
    def _format_range(self, r: tuple) -> str:
        """格式化区域元组为 Excel 格式"""
        def col_to_letter(col: int) -> str:
            result = ""
            while col >= 0:
                result = chr(col % 26 + ord('A')) + result
                col = col // 26 - 1
            return result
        return f"{col_to_letter(r[1])}{r[0]+1}:{col_to_letter(r[3])}{r[2]+1}"
    
    def _show_about(self):
        """显示关于对话框"""
        QMessageBox.about(
            self,
            "关于 Excel 文件比较工具",
            """<h3>Excel 文件比较工具</h3>
            <p>版本: 1.0.0</p>
            <p>功能: 比较两个 Excel 文件的内容差异，提供可视化差异展示和详细比较报告。</p>
            <p>支持格式: .xlsx, .xls</p>
            <hr>
            <p>💝 看着怼怼庆每天被表格差异折磨得焦头烂额，作为男友的我决定用代码来拯救她的加班时光！</p>
            <p>愿这个小工具能让你少熬点夜，多睡点觉~ 😊</p>
            """
        )
