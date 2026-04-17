# measurement_table.py
# 自定义表格控件：支持单格选中、拖拽列范围选择、右键菜单

from PyQt5.QtWidgets import (
    QTableWidget, QTableWidgetItem, QHeaderView, QMenu, QAction, QAbstractItemView
)
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QColor, QBrush, QFont

TESTROWS  = 10
STAT_ROWS = 4                          # 统计行数（跟在数据行后面）
STAT_LABELS = ["均值 (uL)", "CV (%)", "理论值 (uL)", "补偿值 (uL)"]

COLOR_SELECTED  = QColor(0,   120, 215)   # 深蓝：单格选中
COLOR_DRAG      = QColor(144, 202, 249)   # 浅蓝：拖拽范围
COLOR_NORMAL    = QColor(255, 255, 255)   # 白色：普通
COLOR_INDEX     = QColor(245, 245, 245)   # 浅灰：序号列
COLOR_STAT_LBL  = QColor(220, 230, 245)   # 淡蓝：统计标签列
COLOR_STAT_VAL  = QColor(240, 248, 255)   # 极淡蓝：统计数值列


class MeasurementTable(QTableWidget):
    """
    4列表格：序号 | 加注前(g) | 加注后(g) | 加注量(g)
    数据行 10 行 + 统计行 4 行（均值/CV/理论值/偏移量）
    支持：
      - 左键单击选中单元格
      - 左键拖拽（列1/2/3）选择连续行范围（仅数据行）
      - 右键菜单信号（由父窗口处理具体菜单）
    """

    # 右键信号：row=-1 表示点击在列头
    context_menu_requested = pyqtSignal(int, int)   # (row, col)
    # 列头右键信号
    header_context_menu_requested = pyqtSignal(int)  # col
    # 用户手动编辑单元格后触发
    cell_value_changed = pyqtSignal(int, int, str)   # (row, col, new_text)
    # 快捷键信号
    copy_shortcut   = pyqtSignal()
    paste_shortcut  = pyqtSignal()
    delete_shortcut = pyqtSignal()
    undo_shortcut   = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(TESTROWS + STAT_ROWS, 4, parent)

        self._updating = False   # 防止 set_cell_value 触发 itemChanged 循环

        # 拖拽选择状态
        self._dragging    = False
        self._has_drag_sel = False
        self._drag_start  = -1
        self._drag_end    = -1
        self._drag_col    = -1

        # 单格选中状态
        self._sel_row = -1
        self._sel_col = -1

        self._setup_ui()

    # ------------------------------------------------------------------
    # 初始化
    # ------------------------------------------------------------------
    def _setup_ui(self):
        headers = ["序号", "加注前(g)", "加注后(g)", "加注量(g)"]
        self.setHorizontalHeaderLabels(headers)

        # 列宽
        self.setColumnWidth(0, 40)
        self.setColumnWidth(1, 85)
        self.setColumnWidth(2, 85)
        self.setColumnWidth(3, 85)
        self.horizontalHeader().setSectionResizeMode(0, QHeaderView.Fixed)
        self.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)

        # 行高自动拉伸填满表格
        for row in range(TESTROWS + STAT_ROWS):
            self.verticalHeader().setSectionResizeMode(row, QHeaderView.Stretch)

        # 关闭默认选择行为（我们自己管理选中高亮）
        self.setSelectionMode(QAbstractItemView.NoSelection)
        # 列1/2 允许双击或直接键入编辑
        self.setEditTriggers(
            QAbstractItemView.DoubleClicked | QAbstractItemView.AnyKeyPressed
        )
        self.setMouseTracking(True)

        # 默认字体
        cell_font = QFont()
        cell_font.setPointSize(11)
        stat_font = QFont()
        stat_font.setPointSize(11)
        stat_font.setBold(True)

        # ── 数据行 0‥TESTROWS-1 ──
        for row in range(TESTROWS):
            idx_item = QTableWidgetItem(str(row + 1))
            idx_item.setTextAlignment(Qt.AlignCenter)
            idx_item.setBackground(QBrush(COLOR_INDEX))
            idx_item.setFlags(Qt.ItemIsEnabled)
            idx_item.setFont(cell_font)
            self.setItem(row, 0, idx_item)
            for col in (1, 2):
                item = QTableWidgetItem("")
                item.setTextAlignment(Qt.AlignCenter)
                item.setBackground(QBrush(COLOR_NORMAL))
                item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsEditable | Qt.ItemIsSelectable)
                item.setFont(cell_font)
                self.setItem(row, col, item)
            item3 = QTableWidgetItem("")
            item3.setTextAlignment(Qt.AlignCenter)
            item3.setBackground(QBrush(COLOR_NORMAL))
            item3.setFlags(Qt.ItemIsEnabled)
            item3.setFont(cell_font)
            self.setItem(row, 3, item3)

        # ── 统计行 TESTROWS‥TESTROWS+STAT_ROWS-1 ──
        # col 0-2 合并为一个标签单元格，col 3 显示数值
        for i, lbl_text in enumerate(STAT_LABELS):
            row = TESTROWS + i
            # 合并 col 0/1/2 → 标签
            self.setSpan(row, 0, 1, 3)
            lbl_item = QTableWidgetItem(lbl_text)
            lbl_item.setTextAlignment(Qt.AlignCenter)
            lbl_item.setBackground(QBrush(COLOR_STAT_LBL))
            lbl_item.setFont(stat_font)
            lbl_item.setFlags(Qt.ItemIsEnabled)
            self.setItem(row, 0, lbl_item)
            # col 3: 数值（初始空）
            val_item = QTableWidgetItem("")
            val_item.setTextAlignment(Qt.AlignCenter)
            val_item.setBackground(QBrush(COLOR_STAT_VAL))
            val_item.setFont(stat_font)
            val_item.setFlags(Qt.ItemIsEnabled)
            self.setItem(row, 3, val_item)

        # 监听用户编辑
        self.itemChanged.connect(self._on_item_changed)

        # 列头右键
        self.horizontalHeader().setContextMenuPolicy(Qt.CustomContextMenu)
        self.horizontalHeader().customContextMenuRequested.connect(
            self._on_header_right_click
        )

        self.verticalHeader().setVisible(False)
        self.setAlternatingRowColors(False)

    # ------------------------------------------------------------------
    # 公开 API
    # ------------------------------------------------------------------
    def set_stat_value(self, stat_idx: int, text: str, alert: bool = False):
        """设置统计行（0=均值,1=CV,2=理论值,3=补偿值）的数值列；alert=True 时文字标红"""
        row = TESTROWS + stat_idx
        self._updating = True
        try:
            item = self.item(row, 3)
            if item:
                item.setText(text)
                if alert:
                    item.setForeground(QBrush(QColor(200, 0, 0)))
                else:
                    item.setForeground(QBrush(QColor(0, 0, 128)))
                item.setBackground(QBrush(COLOR_STAT_VAL))
        finally:
            self._updating = False

    def get_drag_selection(self):
        """返回 (start_row, end_row, col) 或 None"""
        if not self._has_drag_sel:
            return None
        return (
            min(self._drag_start, self._drag_end),
            max(self._drag_start, self._drag_end),
            self._drag_col,
        )

    def clear_drag_selection(self):
        self._has_drag_sel = False
        self._dragging     = False
        self._drag_start   = -1
        self._drag_end     = -1
        self._drag_col     = -1
        self._refresh_all_colors()

    def get_cell_value(self, row, col):
        """返回单元格文字，不存在时返回空串"""
        item = self.item(row, col)
        return item.text() if item else ""

    def set_cell_value(self, row, col, text):
        self._updating = True
        try:
            item = self.item(row, col)
            if item is None:
                item = QTableWidgetItem()
                item.setTextAlignment(Qt.AlignCenter)
                self.setItem(row, col, item)
            item.setText(text)
        finally:
            self._updating = False
        self._refresh_cell_color(row, col)

    def _on_item_changed(self, item):
        """用户手动编辑单元格时触发，通知外部更新数据模型"""
        if self._updating:
            return
        row, col = item.row(), item.column()
        if row < TESTROWS and col in (1, 2):
            self.cell_value_changed.emit(row, col, item.text())

    def clear_row_data(self, row):
        for col in range(1, 4):
            self.set_cell_value(row, col, "")

    def clear_all_data(self):
        self.clear_drag_selection()
        for row in range(TESTROWS):
            self.clear_row_data(row)

    # ------------------------------------------------------------------
    # 鼠标事件：拖拽选择（仅数据行响应）
    # ------------------------------------------------------------------
    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            index = self.indexAt(event.pos())
            if index.isValid():
                row, col = index.row(), index.column()
                if row >= TESTROWS:          # 统计行不参与交互
                    super().mousePressEvent(event)
                    return
                self._sel_row = row
                self._sel_col = col
                if col in (1, 2, 3):
                    self._dragging     = True
                    self._has_drag_sel = False
                    self._drag_start   = row
                    self._drag_end     = row
                    self._drag_col     = col
                else:
                    self._dragging     = False
                    self._has_drag_sel = False
                self._refresh_all_colors()
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self._dragging and (event.buttons() & Qt.LeftButton):
            index = self.indexAt(event.pos())
            if index.isValid() and index.column() == self._drag_col:
                new_row = min(index.row(), TESTROWS - 1)   # 不超出数据行
                if new_row != self._drag_end:
                    self._drag_end = new_row
                    self._refresh_all_colors()
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton and self._dragging:
            self._dragging     = False
            self._has_drag_sel = (self._drag_start != self._drag_end)
            self._refresh_all_colors()
        super().mouseReleaseEvent(event)

    def keyPressEvent(self, event):
        from PyQt5.QtGui import QKeySequence
        if event.matches(QKeySequence.Copy):
            self.copy_shortcut.emit()
            return
        if event.matches(QKeySequence.Paste):
            self.paste_shortcut.emit()
            return
        if event.matches(QKeySequence.Undo):
            self.undo_shortcut.emit()
            return
        if event.key() == Qt.Key_Delete:
            self.delete_shortcut.emit()
            return
        super().keyPressEvent(event)

    def contextMenuEvent(self, event):
        index = self.indexAt(event.pos())
        if index.isValid() and index.row() < TESTROWS:
            self.context_menu_requested.emit(index.row(), index.column())

    # ------------------------------------------------------------------
    # 列头右键
    # ------------------------------------------------------------------
    def _on_header_right_click(self, pos):
        col = self.horizontalHeader().logicalIndexAt(pos)
        if col in (1, 2, 3):
            self.header_context_menu_requested.emit(col)

    # ------------------------------------------------------------------
    # 颜色刷新（仅数据行）
    # ------------------------------------------------------------------
    def _refresh_all_colors(self):
        drag_sel = self.get_drag_selection()  # None 或 (s,e,c)
        for row in range(TESTROWS):
            for col in range(1, 4):
                self._refresh_cell_color(row, col, drag_sel)

    def _refresh_cell_color(self, row, col, drag_sel=None):
        if row >= TESTROWS:
            return
        item = self.item(row, col)
        if item is None:
            return
        if row == self._sel_row and col == self._sel_col:
            item.setBackground(QBrush(COLOR_SELECTED))
            item.setForeground(QBrush(QColor(255, 255, 255)))
        elif drag_sel and col == drag_sel[2] and drag_sel[0] <= row <= drag_sel[1]:
            item.setBackground(QBrush(COLOR_DRAG))
            item.setForeground(QBrush(QColor(0, 0, 0)))
        else:
            item.setBackground(QBrush(COLOR_NORMAL))
            item.setForeground(QBrush(QColor(0, 0, 0)))
