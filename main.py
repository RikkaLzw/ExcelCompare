"""
Excel 文件比较工具 - 程序入口

功能：比较两个 Excel 文件的内容差异，提供可视化差异展示和详细比较报告。
"""
import sys
from PyQt6.QtWidgets import QApplication
from PyQt6.QtCore import Qt
from src.views.main_window import MainWindow


def main():
    # 启用高DPI缩放
    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
    )
    
    app = QApplication(sys.argv)
    app.setApplicationName("Excel Compare")
    app.setApplicationDisplayName("Excel 文件比较工具")
    
    # 创建并显示主窗口
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
