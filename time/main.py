import sys
from PyQt6.QtWidgets import QApplication, QMainWindow, QTabWidget, QWidget
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont

from models import AppData
from tab_action_list import ActionListTab
from tab_mode_editor import ModeEditorTab


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("节拍动作设计工具")
        self.resize(1200, 700)

        self.app_data = AppData()

        self.tabs = QTabWidget()
        self.tabs.setStyleSheet(
            "QTabBar::tab{padding:8px 20px;font-size:13px;}"
            "QTabBar::tab:selected{font-weight:bold;color:#1565C0;}"
        )

        self.action_list_tab = ActionListTab(self.app_data)
        self.mode_editor_tab = ModeEditorTab(self.app_data)

        self.tabs.addTab(self.action_list_tab, "一级动作列表")
        self.tabs.addTab(self.mode_editor_tab, "模式编辑")

        self.action_list_tab.data_changed.connect(self.mode_editor_tab.on_data_changed)

        self.setCentralWidget(self.tabs)


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    font = QFont("Microsoft YaHei", 10)
    app.setFont(font)

    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
