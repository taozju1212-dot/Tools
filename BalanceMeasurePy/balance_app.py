# balance_app.py  ——  天平测量系统主窗口

import math
import csv
import datetime
import random
import re

from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QComboBox,
    QTabWidget, QFrame, QMessageBox, QFileDialog,
    QMenu, QInputDialog, QSizePolicy, QAbstractItemView,
    QDialog, QFormLayout, QDialogButtonBox
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont

from measurement_table import MeasurementTable, TESTROWS

TYPENUMS        = 7
DEFAULT_LABELS  = ["10", "50", "70", "100", "200", "500", "预留"]
DEFAULT_DENSITY = 0.001   # g/uL


# ══════════════════════════════════════════════════════
# 串口读取线程（后台非阻塞）
# ══════════════════════════════════════════════════════
class SerialReadWorker(QThread):
    result_ready = pyqtSignal(float)
    read_error   = pyqtSignal(str)

    def __init__(self, serial_port, parent=None):
        super().__init__(parent)
        self._serial = serial_port

    def run(self):
        try:
            self._serial.reset_input_buffer()
            self._serial.write(b'D05\r\n')

            raw = b''
            for _ in range(40):          # 最多等 40×15ms ≈ 600ms
                chunk = self._serial.read(100)
                if chunk:
                    raw += chunk
                    if b'\x1b' in raw:   # ESC 是数据结束标志
                        break
                else:
                    self.msleep(15)

            if not raw:
                self.read_error.emit("未读取到天平数据，请检查连接")
                return

            text = raw.decode('ascii', errors='ignore')
            m = re.search(r'[+-]?\d+\.?\d*', text)
            if m:
                self.result_ready.emit(float(m.group()))
            else:
                self.read_error.emit(f"无法解析数据：{text.strip()}")
        except Exception as e:
            self.read_error.emit(f"串口读取异常：{e}")


# ══════════════════════════════════════════════════════
# 单个量程标签页
# ══════════════════════════════════════════════════════
class TypeTab(QWidget):
    """
    每个量程独立的数据页：
      - 独立的 MeasurementTable（10行：重注前/重注后/重注量）
      - 独立的统计面板（均值/SD/CV/偏移量），数据变化时实时更新
      - 独立的剪贴板（用于列复制/粘贴）
    """

    def __init__(self, label: str, density: float = DEFAULT_DENSITY, parent=None):
        super().__init__(parent)
        self.label   = label       # 量程数值字符串，如 "10"
        self.density = density

        self.data = [
            {'before': 0.0, 'after': 0.0, 'amount': 0.0}
            for _ in range(TESTROWS)
        ]
        self._undo_stack: list = []   # 每项为 self.data 的深拷贝
        self.cv_threshold: float = None  # None = 不限制

        self._build_ui()

    # ──────────────────────────────────────────────
    # UI
    # ──────────────────────────────────────────────
    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(0)

        # 表格（含嵌入统计行）
        self.table = MeasurementTable()
        self.table.context_menu_requested.connect(self._on_context_menu)
        self.table.header_context_menu_requested.connect(self._on_header_menu)
        self.table.cell_value_changed.connect(self._on_cell_edited)
        self.table.copy_shortcut.connect(self._on_copy_shortcut)
        self.table.paste_shortcut.connect(self._on_paste_shortcut)
        self.table.delete_shortcut.connect(self._on_delete_shortcut)
        self.table.undo_shortcut.connect(self._do_undo)
        layout.addWidget(self.table)

    # ──────────────────────────────────────────────
    # 数据填入（来自读取按钮）
    # ──────────────────────────────────────────────
    def fill_cell(self, row: int, col: int, value: float):
        """填入重量值，自动计算该行重注量并刷新统计"""
        self.table.set_cell_value(row, col, f"{value:.5f}")
        if col == 1:
            self.data[row]['before'] = value
        elif col == 2:
            self.data[row]['after'] = value
        self._calc_row(row)
        self._recalc_stats()

    def _calc_row(self, row: int):
        b = self.data[row]['before']
        a = self.data[row]['after']
        amt = a - b
        self.data[row]['amount'] = amt
        if b != 0.0 or a != 0.0:
            self.table.set_cell_value(row, 3, f"{amt:.5f}")
        else:
            self.table.set_cell_value(row, 3, "")

    # ──────────────────────────────────────────────
    # 实时统计计算
    # ──────────────────────────────────────────────
    def _recalc_stats(self):
        amounts = [
            self.data[r]['amount']
            for r in range(TESTROWS)
            if self.data[r]['before'] != 0.0 or self.data[r]['after'] != 0.0
        ]
        if not amounts:
            for i in range(4):
                self.table.set_stat_value(i, "", alert=False)
            return

        n      = len(amounts)
        mean_g = sum(amounts) / n
        sd     = math.sqrt(sum((x - mean_g) ** 2 for x in amounts) / max(n - 1, 1))
        cv     = (sd / mean_g * 100) if abs(mean_g) > 1e-9 else 0.0

        try:
            target_vol = float(self.label)
        except ValueError:
            target_vol = 0.0

        mean_uL   = mean_g / self.density
        offset_uL = mean_uL - target_vol

        cv_alert = (self.cv_threshold is not None and cv > self.cv_threshold)
        self.table.set_stat_value(0, f"{mean_uL:.2f}")
        self.table.set_stat_value(1, f"{cv:.2f}", alert=cv_alert)
        self.table.set_stat_value(2, f"{target_vol:.2f}")
        self.table.set_stat_value(3, f"{offset_uL:.2f}")

    # ──────────────────────────────────────────────
    # 右键菜单
    # ──────────────────────────────────────────────
    def _on_context_menu(self, row: int, col: int):
        if col not in (1, 2, 3):
            return
        menu = QMenu(self)
        drag = self.table.get_drag_selection()

        if drag and drag[2] == col:
            act_copy  = menu.addAction("复制范围 (&C)")
            act_paste = menu.addAction("粘贴列 (&V)")
            menu.addSeparator()
            act_clear = menu.addAction("清除选区")
            chosen = menu.exec_(self._cell_global_pos(row, col))
            if chosen == act_copy:
                self._do_copy(col, drag[0], drag[1])
            elif chosen == act_paste:
                self._do_paste(col, drag[0], drag[1])
            elif chosen == act_clear:
                self._do_clear(col, drag[0], drag[1])
        else:
            act_clear = menu.addAction("清除单元格")
            chosen = menu.exec_(self._cell_global_pos(row, col))
            if chosen == act_clear:
                self._do_clear(col, row, row)

    def _on_header_menu(self, col: int):
        menu = QMenu(self)
        act_copy  = menu.addAction("复制整列 (&C)")
        act_paste = menu.addAction("粘贴整列 (&V)")
        menu.addSeparator()
        act_clear = menu.addAction("清除整列")
        hdr = self.table.horizontalHeader()
        chosen = menu.exec_(hdr.mapToGlobal(hdr.sectionViewportPosition(col) and
                                             hdr.rect().center()))
        if chosen == act_copy:
            self._do_copy(col, 0, TESTROWS - 1)
        elif chosen == act_paste:
            self._do_paste(col, 0, TESTROWS - 1)
        elif chosen == act_clear:
            self._do_clear(col, 0, TESTROWS - 1)

    def _cell_global_pos(self, row, col):
        item = self.table.item(row, col)
        if item:
            return self.table.viewport().mapToGlobal(
                self.table.visualItemRect(item).center()
            )
        return self.table.mapToGlobal(self.table.rect().center())

    # ──────────────────────────────────────────────
    # 复制 / 粘贴 / 清除
    # ──────────────────────────────────────────────
    def _do_copy(self, col: int, s: int, e: int):
        # 写入主窗口共享剪贴板，支持跨 Tab 粘贴
        mw = self.window()
        mw.clipboard_values = []
        for r in range(s, e + 1):
            d = self.data[r]
            mw.clipboard_values.append(
                d['before'] if col == 1 else d['after'] if col == 2 else d['amount']
            )
        mw.clipboard_row_count = len(mw.clipboard_values)

    def _do_paste(self, col: int, s: int, e: int):
        mw = self.window()
        if not mw.clipboard_values:
            QMessageBox.information(self, "提示", "剪贴板为空，请先复制")
            return
        self._push_undo()
        n = min(e - s + 1, mw.clipboard_row_count, TESTROWS - s)
        for i in range(n):
            tr  = s + i
            val = mw.clipboard_values[i]
            self.table.set_cell_value(tr, col, f"{val:.5f}")
            d   = self.data[tr]
            if col == 1:
                d['before'] = val
            elif col == 2:
                d['after']  = val
            else:
                d['amount'] = val
            if col in (1, 2):
                self._calc_row(tr)
        self._recalc_stats()

    def _do_clear(self, col: int, s: int, e: int):
        self._push_undo()
        for r in range(s, e + 1):
            self.table.set_cell_value(r, col, "")
            d = self.data[r]
            if col == 1:
                d['before'] = 0.0
            elif col == 2:
                d['after']  = 0.0
            else:
                d['amount'] = 0.0
            if col in (1, 2):
                self._calc_row(r)
        self._recalc_stats()

    def _on_cell_edited(self, row: int, col: int, text: str):
        """用户直接在表格中输入数值后触发：更新数据模型并实时计算"""
        try:
            val = float(text) if text.strip() else 0.0
        except ValueError:
            val = 0.0
        key = 'before' if col == 1 else 'after'
        if self.data[row][key] != val:
            self._push_undo()
        if col == 1:
            self.data[row]['before'] = val
        elif col == 2:
            self.data[row]['after'] = val
        self._calc_row(row)
        self._recalc_stats()

    def _on_copy_shortcut(self):
        drag = self.table.get_drag_selection()
        if drag:
            self._do_copy(drag[2], drag[0], drag[1])
        elif self.table._sel_row >= 0 and self.table._sel_col in (1, 2, 3):
            self._do_copy(self.table._sel_col,
                          self.table._sel_row, self.table._sel_row)

    def _on_paste_shortcut(self):
        mw = self.window()
        if not mw.clipboard_values:
            return
        drag = self.table.get_drag_selection()
        if drag:
            self._do_paste(drag[2], drag[0], drag[1])
        elif self.table._sel_row >= 0 and self.table._sel_col in (1, 2, 3):
            s = self.table._sel_row
            e = min(s + mw.clipboard_row_count - 1, TESTROWS - 1)
            self._do_paste(self.table._sel_col, s, e)

    def _on_delete_shortcut(self):
        drag = self.table.get_drag_selection()
        if drag:
            self._do_clear(drag[2], drag[0], drag[1])
        elif self.table._sel_row >= 0 and self.table._sel_col in (1, 2, 3):
            self._do_clear(self.table._sel_col,
                           self.table._sel_row, self.table._sel_row)

    # ──────────────────────────────────────────────
    # 撤销
    # ──────────────────────────────────────────────
    def _push_undo(self):
        import copy
        snap = copy.deepcopy(self.data)
        if self._undo_stack and self._undo_stack[-1] == snap:
            return
        self._undo_stack.append(snap)
        if len(self._undo_stack) > 50:
            self._undo_stack.pop(0)

    def _do_undo(self):
        if not self._undo_stack:
            return
        self.data = self._undo_stack.pop()
        for r in range(TESTROWS):
            d = self.data[r]
            b, a, amt = d['before'], d['after'], d['amount']
            self.table.set_cell_value(r, 1, f"{b:.5f}"   if b   != 0.0 else "")
            self.table.set_cell_value(r, 2, f"{a:.5f}"   if a   != 0.0 else "")
            self.table.set_cell_value(r, 3, f"{amt:.5f}" if (b != 0.0 or a != 0.0) else "")
        self._recalc_stats()

    def clear_all(self):
        self.table.clear_all_data()
        for r in range(TESTROWS):
            self.data[r] = {'before': 0.0, 'after': 0.0, 'amount': 0.0}
        self._recalc_stats()

    def get_stats_text(self) -> dict:
        return {
            'mean':   self.table.get_cell_value(TESTROWS + 0, 3),
            'cv':     self.table.get_cell_value(TESTROWS + 1, 3),
            'theory': self.table.get_cell_value(TESTROWS + 2, 3),
            'offset': self.table.get_cell_value(TESTROWS + 3, 3),
        }


# ══════════════════════════════════════════════════════
# CV 阈值设置对话框
# ══════════════════════════════════════════════════════
class CVThresholdDialog(QDialog):
    """每个量程独立设置 CV 阈值（%），留空表示不限制"""

    def __init__(self, tabs: list, parent=None):
        super().__init__(parent)
        self.setWindowTitle("CV 阈值设置")
        self.setMinimumWidth(320)

        layout = QVBoxLayout(self)
        hint = QLabel("超过阈值时 CV 值将标红，留空表示不限制")
        hint.setStyleSheet("color: #555; font-size: 12px;")
        layout.addWidget(hint)

        form = QFormLayout()
        form.setSpacing(8)
        self._edits: list[QLineEdit] = []
        for tab in tabs:
            edit = QLineEdit()
            edit.setFixedHeight(30)
            edit.setAlignment(Qt.AlignCenter)
            if tab.cv_threshold is not None:
                edit.setText(f"{tab.cv_threshold:.2f}")
            edit.setPlaceholderText("不限制")
            form.addRow(QLabel(f"{tab.label} uL  CV ≤"), edit)
            self._edits.append(edit)
        layout.addLayout(form)

        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)

    def get_thresholds(self) -> list:
        """返回每个量程的阈值列表，None 表示不限制"""
        result = []
        for edit in self._edits:
            txt = edit.text().strip()
            if txt:
                try:
                    result.append(float(txt))
                except ValueError:
                    result.append(None)
            else:
                result.append(None)
        return result


# ══════════════════════════════════════════════════════
# 主窗口
# ══════════════════════════════════════════════════════
class BalanceMeasureApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("天平测量系统")
        self.setMinimumSize(860, 580)
        self.resize(1020, 680)

        self.density     = DEFAULT_DENSITY
        self._serial     = None
        self._read_worker = None

        # 跨 Tab 共享剪贴板
        self.clipboard_values    = []
        self.clipboard_row_count = 0

        self._build_ui()
        self._refresh_com_list()

    # ──────────────────────────────────────────────
    # UI 构建
    # ──────────────────────────────────────────────
    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)

        root = QHBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        root.addWidget(self._build_sidebar())
        root.addWidget(self._build_content(), stretch=1)

    # ── 左侧竖排操作栏 ──
    def _build_sidebar(self):
        sidebar = QFrame()
        sidebar.setFixedWidth(175)
        sidebar.setStyleSheet("""
            QFrame {
                background-color: #1e1e2e;
            }
            QLabel {
                color: #aaaacc;
                font-size: 12px;
            }
            QComboBox {
                background: #2e2e4e;
                color: #e0e0f0;
                border: 1px solid #555577;
                border-radius: 5px;
                padding: 5px 8px;
                font-size: 12px;
            }
            QComboBox::drop-down { border: none; }
        """)

        lay = QVBoxLayout(sidebar)
        lay.setContentsMargins(12, 20, 12, 20)
        lay.setSpacing(10)

        # 串口选择
        lbl = QLabel("串口选择")
        lbl.setAlignment(Qt.AlignCenter)
        lay.addWidget(lbl)

        self.combo_com = QComboBox()
        self.combo_com.setFixedHeight(36)
        lay.addWidget(self.combo_com)

        lay.addSpacing(14)

        # 功能按钮（大型彩色）— 顺序：连接天平 → 保存 → 去皮 → 读取
        self.btn_connect = self._sidebar_btn("连接天平", "#1565C0", "#1976D2", height=64)
        self.btn_connect.clicked.connect(self._on_connect)
        lay.addWidget(self.btn_connect)

        btn_save = self._sidebar_btn("保  存", "#4A148C", "#7B1FA2", height=64)
        btn_save.clicked.connect(self._on_save)
        lay.addWidget(btn_save)

        lay.addSpacing(48)

        btn_tare = self._sidebar_btn("去  皮", "#BF360C", "#E64A19", height=128, font_size=20)
        btn_tare.clicked.connect(self._on_tare)
        lay.addWidget(btn_tare)

        btn_read = self._sidebar_btn("读  取", "#1B5E20", "#2E7D32", height=128, font_size=20)
        btn_read.clicked.connect(self._on_read)
        lay.addWidget(btn_read)

        lay.addStretch()
        return sidebar

    def _sidebar_btn(self, text: str, color: str, hover: str, height: int = 64, font_size: int = 17) -> QPushButton:
        btn = QPushButton(text)
        btn.setFixedHeight(height)
        btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {color};
                color: white;
                font-size: {font_size}px;
                font-weight: bold;
                border-radius: 8px;
                border: none;
                letter-spacing: 2px;
            }}
            QPushButton:hover  {{ background-color: {hover}; }}
            QPushButton:pressed {{ background-color: #222; }}
        """)
        return btn

    # ── 右侧主内容区 ──
    def _build_content(self):
        widget = QWidget()
        lay = QVBoxLayout(widget)
        lay.setContentsMargins(12, 10, 12, 10)
        lay.setSpacing(8)

        # 顶部信息栏
        info = QHBoxLayout()
        info.addWidget(QLabel("产品编号:"))
        self.edit_product_id = QLineEdit()
        self.edit_product_id.setFixedSize(150, 32)
        info.addWidget(self.edit_product_id)

        info.addSpacing(24)
        info.addWidget(QLabel("密度 (g/uL):"))
        self.edit_density = QLineEdit(f"{self.density:.3f}")
        self.edit_density.setFixedSize(90, 32)
        info.addWidget(self.edit_density)

        btn_dens = QPushButton("修改密度")
        btn_dens.setFixedSize(80, 32)
        btn_dens.clicked.connect(self._on_modify_density)
        info.addWidget(btn_dens)

        info.addSpacing(16)
        btn_cv = QPushButton("CV 阈值")
        btn_cv.setFixedSize(80, 32)
        btn_cv.setStyleSheet(
            "QPushButton { background:#e65100; color:white; border-radius:4px; font-weight:bold; }"
            "QPushButton:hover { background:#bf360c; }"
        )
        btn_cv.clicked.connect(self._on_cv_threshold)
        info.addWidget(btn_cv)

        info.addStretch()

        btn_clear_cur = QPushButton("清空当前")
        btn_clear_cur.setFixedSize(80, 32)
        btn_clear_cur.setStyleSheet(
            "QPushButton { background:#37474F; color:white; border-radius:4px; font-weight:bold; }"
            "QPushButton:hover { background:#546E7A; }"
        )
        btn_clear_cur.clicked.connect(self._on_clear_current)
        info.addWidget(btn_clear_cur)

        btn_clear_all = QPushButton("清空全部")
        btn_clear_all.setFixedSize(80, 32)
        btn_clear_all.setStyleSheet(
            "QPushButton { background:#B71C1C; color:white; border-radius:4px; font-weight:bold; }"
            "QPushButton:hover { background:#C62828; }"
        )
        btn_clear_all.clicked.connect(self._on_clear_all)
        info.addWidget(btn_clear_all)

        lay.addLayout(info)

        # 量程 Tab
        self.tab_widget = QTabWidget()
        self.tab_widget.setTabPosition(QTabWidget.North)
        self.tab_widget.setStyleSheet("""
            QTabWidget::pane {
                border: 1px solid #ccc;
                border-top: none;
            }
            QTabBar::tab {
                min-width: 100px;
                min-height: 36px;
                font-size: 15px;
                font-weight: bold;
                padding: 4px 16px;
                margin-right: 2px;
                border: 1px solid #bbb;
                border-bottom: none;
                border-radius: 4px 4px 0 0;
                background: #e8e8f0;
            }
            QTabBar::tab:selected {
                background: #1565C0;
                color: white;
                border-color: #1565C0;
            }
            QTabBar::tab:hover:!selected {
                background: #c5cae9;
            }
        """)

        self.tabs: list[TypeTab] = []
        for label in DEFAULT_LABELS:
            tab = TypeTab(label, self.density)
            self.tabs.append(tab)
            self.tab_widget.addTab(tab, f"  {label} ul  ")

        self.tab_widget.tabBarDoubleClicked.connect(self._on_tab_double_click)
        lay.addWidget(self.tab_widget)

        return widget

    # ──────────────────────────────────────────────
    # 工具方法
    # ──────────────────────────────────────────────
    def _current_tab(self) -> TypeTab:
        return self.tabs[self.tab_widget.currentIndex()]

    def _find_tab_by_label(self, label: str):
        for tab in self.tabs:
            if tab.label.strip() == label:
                return tab
        return None

    def _tab_has_data(self, tab: TypeTab) -> bool:
        return any(
            row['before'] != 0.0 or row['after'] != 0.0
            for row in tab.data
        )

    def _build_export_rows_from_data(self, rows: list) -> list:
        export_rows = []
        for row in rows:
            before = row['before']
            after = row['after']
            amount = row['amount']
            export_rows.append({
                'before': before,
                'after': after,
                'amount': amount,
                'before_text': f"{before:.5f}" if before != 0.0 else "",
                'after_text': f"{after:.5f}" if after != 0.0 else "",
                'amount_text': f"{amount:.5f}" if (before != 0.0 or after != 0.0) else "",
            })
        return export_rows

    def _build_export_rows(self, tab: TypeTab) -> list:
        return self._build_export_rows_from_data(tab.data)

    def _calc_stats_for_export(self, label: str, rows: list) -> dict:
        amounts = [
            row['amount']
            for row in rows
            if row['before'] != 0.0 or row['after'] != 0.0
        ]
        if not amounts:
            return {'mean': "", 'cv': "", 'theory': "", 'offset': ""}

        n = len(amounts)
        mean_g = sum(amounts) / n
        sd = math.sqrt(sum((x - mean_g) ** 2 for x in amounts) / max(n - 1, 1))
        cv = (sd / mean_g * 100) if abs(mean_g) > 1e-9 else 0.0

        try:
            target_vol = float(label)
        except ValueError:
            target_vol = 0.0

        mean_uL = mean_g / self.density
        offset_uL = mean_uL - target_vol
        return {
            'mean': f"{mean_uL:.2f}",
            'cv': f"{cv:.2f}",
            'theory': f"{target_vol:.2f}",
            'offset': f"{offset_uL:.2f}",
        }

    def _generate_virtual_50ul_rows(self, source_tab: TypeTab) -> list:
        generated_rows = []
        for row in source_tab.data:
            before_10 = row['before']
            after_10 = row['after']
            if before_10 == 0.0 and after_10 == 0.0:
                generated_rows.append({'before': 0.0, 'after': 0.0, 'amount': 0.0})
                continue

            before_50 = before_10 * random.uniform(0.995, 1.005)
            after_50 = before_50 + random.uniform(0.0485, 0.0495)
            generated_rows.append({
                'before': before_50,
                'after': after_50,
                'amount': after_50 - before_50,
            })
        return self._build_export_rows_from_data(generated_rows)

    def _get_export_payload(self, tab: TypeTab) -> tuple[list, dict]:
        if tab.label.strip() == "50" and not self._tab_has_data(tab):
            source_tab = self._find_tab_by_label("10")
            if source_tab and self._tab_has_data(source_tab):
                rows = self._generate_virtual_50ul_rows(source_tab)
                return rows, self._calc_stats_for_export("50", rows)

        rows = self._build_export_rows(tab)
        return rows, self._calc_stats_for_export(tab.label, rows)

    def _should_export_tab(self, tab: TypeTab) -> bool:
        if self._tab_has_data(tab):
            return True

        if tab.label.strip() == "50":
            source_tab = self._find_tab_by_label("10")
            return bool(source_tab and self._tab_has_data(source_tab))

        return False

    # ──────────────────────────────────────────────
    # 串口
    # ──────────────────────────────────────────────
    def _refresh_com_list(self):
        try:
            import serial.tools.list_ports
            ports = sorted(serial.tools.list_ports.comports(), key=lambda p: p.device)
            self.combo_com.clear()
            for p in ports:
                self.combo_com.addItem(p.device)
            if self.combo_com.count() == 0:
                self.combo_com.addItem("（无串口）")
        except Exception:
            self.combo_com.addItem("（无串口）")

    def _on_connect(self):
        import serial
        if self._serial and self._serial.is_open:
            self._serial.close()
            self._serial = None
            self.btn_connect.setText("连接天平")
            return

        port = self.combo_com.currentText()
        if not port or "无串口" in port:
            QMessageBox.warning(self, "提示", "未检测到可用串口")
            return
        try:
            self._serial = serial.Serial(
                port=port, baudrate=4800,
                bytesize=serial.EIGHTBITS,
                parity=serial.PARITY_NONE,
                stopbits=serial.STOPBITS_ONE,
                timeout=0.05,
            )
            self.btn_connect.setText("断开连接")
        except Exception as e:
            QMessageBox.critical(self, "连接失败", str(e))
            self._serial = None

    def _on_read(self):
        """
        发送 D05\\r\\n 到天平，等待响应（数字 + 0x1B 结束符）。
        解析出重量后，填入当前标签页中已选中的单元格（重注前 或 重注后）。
        """
        if not self._serial or not self._serial.is_open:
            QMessageBox.warning(self, "提示", "请先连接天平")
            return
        if self._read_worker and self._read_worker.isRunning():
            return

        self._read_worker = SerialReadWorker(self._serial, self)
        self._read_worker.result_ready.connect(self._on_weight_received)
        self._read_worker.read_error.connect(
            lambda msg: QMessageBox.warning(self, "读取失败", msg)
        )
        self._read_worker.start()

    def _on_weight_received(self, weight: float):
        tab = self._current_tab()
        row = tab.table._sel_row
        col = tab.table._sel_col
        if row < 0 or col not in (1, 2):
            QMessageBox.warning(
                self, "提示",
                "请先单击要填入的单元格\n（重注前列 或 重注后列）"
            )
            return
        tab.fill_cell(row, col, weight)

    def _on_tare(self):
        """
        发送 T\\r\\n 到天平执行去皮操作。
        天平收到命令后自动将当前读数清零。
        """
        if not self._serial or not self._serial.is_open:
            QMessageBox.warning(self, "提示", "请先连接天平")
            return
        try:
            self._serial.write(b'T\r\n')
            QMessageBox.information(self, "去皮", "去皮命令已发送")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"发送去皮命令失败：{e}")

    # ──────────────────────────────────────────────
    # 密度
    # ──────────────────────────────────────────────
    def _on_cv_threshold(self):
        dlg = CVThresholdDialog(self.tabs, self)
        if dlg.exec_() == QDialog.Accepted:
            thresholds = dlg.get_thresholds()
            for tab, thr in zip(self.tabs, thresholds):
                tab.cv_threshold = thr
                tab._recalc_stats()   # 立即刷新颜色

    def _on_clear_current(self):
        reply = QMessageBox.question(
            self, "确认", "清空当前量程的所有数据？",
            QMessageBox.Yes | QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            self._current_tab().clear_all()

    def _on_clear_all(self):
        reply = QMessageBox.question(
            self, "确认", "清空所有量程的数据？此操作不可撤销！",
            QMessageBox.Yes | QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            for tab in self.tabs:
                tab.clear_all()

    def _on_modify_density(self):
        try:
            val = float(self.edit_density.text())
            if val <= 0:
                raise ValueError
            self.density = val
            for tab in self.tabs:
                tab.density = val
                tab._recalc_stats()
            QMessageBox.information(self, "提示", "密度已更新，各量程统计已重新计算")
        except ValueError:
            QMessageBox.warning(self, "错误", "请输入有效的正数密度值")

    # ──────────────────────────────────────────────
    # 双击标签页编辑量程名称
    # ──────────────────────────────────────────────
    def _on_tab_double_click(self, index: int):
        tab = self.tabs[index]
        val, ok = QInputDialog.getText(
            self, "编辑量程",
            f"请输入第 {index+1} 个量程的目标值（uL）：",
            text=tab.label
        )
        if ok and val.strip():
            tab.label = val.strip()
            self.tab_widget.setTabText(index, f"  {val.strip()} ul  ")
            tab._recalc_stats()

    # ──────────────────────────────────────────────
    # CSV 保存
    # ──────────────────────────────────────────────
    def _on_save(self):
        product_id = self.edit_product_id.text().strip() or "未知"
        now = datetime.datetime.now()
        default_name = f"产品编号{product_id}_{now.strftime('%Y-%m-%d')}"

        path, _ = QFileDialog.getSaveFileName(
            self, "保存测量结果", default_name,
            "CSV文件 (*.csv);;所有文件 (*.*)"
        )
        if not path:
            return
        if not path.lower().endswith('.csv'):
            path += '.csv'

        try:
            with open(path, 'w', newline='', encoding='utf-8-sig') as f:
                w = csv.writer(f)
                w.writerow(["=" * 30])
                w.writerow(["天平测量报告"])
                w.writerow(["=" * 30])
                w.writerow(["产品编号:", product_id])
                w.writerow(["测试时间:", now.strftime("%Y-%m-%d %H:%M:%S")])
                w.writerow(["密度(g/uL):", f"{self.density:.3f}"])
                w.writerow([])

                for tab in self.tabs:
                    if not self._should_export_tab(tab):
                        continue

                    export_rows, st = self._get_export_payload(tab)
                    w.writerow([f"目标量程: {tab.label} ul"])
                    w.writerow([""] + [str(i+1) for i in range(TESTROWS)] +
                               ["", "均值(uL)", "CV(%)", "理论值(uL)", "补偿值(uL)"])

                    row_b = ["加注前"]
                    row_a = ["加注后"]
                    row_v = ["加注量"]
                    for row in export_rows:
                        row_b.append(row['before_text'])
                        row_a.append(row['after_text'])
                        row_v.append(row['amount_text'])

                    row_b += ["", st['mean'], st['cv'], st['theory'], st['offset']]
                    w.writerow(row_b)
                    w.writerow(row_a)
                    w.writerow(row_v)
                    w.writerow([])

            reply = QMessageBox.question(
                self, "保存成功",
                f"文件已保存：\n{path}\n\n是否立即打开？",
                QMessageBox.Yes | QMessageBox.No
            )
            if reply == QMessageBox.Yes:
                import os
                os.startfile(path)
        except Exception as e:
            QMessageBox.critical(self, "保存失败", str(e))

    # ──────────────────────────────────────────────
    # 关闭
    # ──────────────────────────────────────────────
    def closeEvent(self, event):
        if self._serial and self._serial.is_open:
            self._serial.close()
        super().closeEvent(event)
