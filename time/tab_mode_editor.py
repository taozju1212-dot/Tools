from __future__ import annotations
import json
from typing import Optional

from PyQt6.QtCore import Qt, pyqtSignal, QSize, QEvent, QPoint
from PyQt6.QtGui import QColor, QBrush
from PyQt6.QtWidgets import (
    QWidget, QHBoxLayout, QVBoxLayout, QLabel, QPushButton,
    QComboBox, QDoubleSpinBox, QFrame, QScrollArea, QFileDialog,
    QMessageBox, QSpinBox, QTableWidget, QTableWidgetItem,
    QAbstractItemView, QHeaderView, QSizePolicy,
)

from models import AppData, ModeConfig, TimelineRow

# ── color palette ──────────────────────────────────────────────────────────────
_PALETTE = [
    "#4FC3F7", "#81C784", "#FFB74D", "#F06292",
    "#BA68C8", "#4DD0E1", "#FF8A65", "#AED581",
    "#90A4AE", "#FFD54F", "#80CBC4", "#EF9A9A",
]

def module_color(module_id: str) -> QColor:
    idx = (ord(module_id) - ord("A")) % len(_PALETTE)
    return QColor(_PALETTE[idx])

COL_MODULE = 0
COL_ACTION = 1
COL_NAME   = 2
FIXED_COLS = 3

ROLE_ROW_TYPE = Qt.ItemDataRole.UserRole
ROW_MAIN   = "main"
ROW_DETAIL = "detail"

_COMBO_STYLE = (
    "QComboBox{border:1px solid #ddd;border-radius:3px;padding:1px 4px;background:#fff;}"
    "QComboBox:focus{border-color:#2196F3;}"
)
_BTN_BLUE = (
    "QPushButton{background:#2196F3;color:white;border-radius:12px;"
    "padding:5px 18px;font-weight:bold;}"
    "QPushButton:hover{background:#1976D2;}"
)
_BTN_LIGHT = (
    "QPushButton{background:#E3F2FD;color:#1565C0;border:1px solid #90CAF9;"
    "border-radius:12px;padding:5px 14px;font-weight:bold;}"
    "QPushButton:hover{background:#BBDEFB;}"
)


# ── DetailWidget ───────────────────────────────────────────────────────────────

class DetailWidget(QWidget):
    changed = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setStyleSheet("background:#F8F9FA;")
        layout = QHBoxLayout(self)
        layout.setContentsMargins(8, 4, 8, 4)
        layout.setSpacing(16)

        self.lbl_start = QLabel("起始时间")
        self.combo_start = QComboBox()
        self.combo_start.setMinimumWidth(90)
        self.combo_start.currentIndexChanged.connect(self.changed)
        layout.addWidget(self.lbl_start)
        layout.addWidget(self.combo_start)

        self.lbl_prev = QLabel("前级动作")
        self.combo_prev = QComboBox()
        self.combo_prev.setMinimumWidth(150)
        self.combo_prev.setEditable(True)
        self.combo_prev.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.combo_prev.lineEdit().setPlaceholderText("——  输入筛选…")
        # MatchContains filter so typing "B" shows all B-module actions
        self.combo_prev.completer().setFilterMode(Qt.MatchFlag.MatchContains)
        self.combo_prev.completer().setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
        self.combo_prev.currentIndexChanged.connect(self.changed)
        layout.addWidget(self.lbl_prev)
        layout.addWidget(self.combo_prev)

        self.lbl_dur = QLabel("动作时间")
        self.spin_dur = QDoubleSpinBox()
        self.spin_dur.setRange(0.1, 200.0)
        self.spin_dur.setSingleStep(0.1)
        self.spin_dur.setDecimals(1)
        self.spin_dur.setValue(1.0)
        self.spin_dur.valueChanged.connect(self.changed)
        layout.addWidget(self.lbl_dur)
        layout.addWidget(self.spin_dur)
        layout.addWidget(QLabel("S"))
        layout.addStretch()

    def populate_start_times(self, beat_time: int, step: float):
        self.combo_start.blockSignals(True)
        cur = self.get_start_time()
        self.combo_start.clear()
        self.combo_start.addItem("——", None)
        t = 0.0
        while t <= beat_time:
            self.combo_start.addItem(f"{t:.1f}", t)
            t = round(t + step, 1)
        if cur is not None:
            idx = self.combo_start.findData(cur)
            if idx >= 0:
                self.combo_start.setCurrentIndex(idx)
        self.combo_start.blockSignals(False)

    def populate_prev_actions(self, options: list[tuple[str, str]]):
        self.combo_prev.blockSignals(True)
        cur = self.get_prev_action_key()
        self.combo_prev.clear()
        self.combo_prev.addItem("——", None)
        for key, text in options:
            self.combo_prev.addItem(text, key)
        if cur:
            idx = self.combo_prev.findData(cur)
            if idx >= 0:
                self.combo_prev.setCurrentIndex(idx)
        self.combo_prev.blockSignals(False)

    def get_start_time(self) -> Optional[float]:
        return self.combo_start.currentData()

    def get_prev_action_key(self) -> Optional[str]:
        data = self.combo_prev.currentData()
        if data is not None:
            return data
        # editable combo: user may have typed; try to match by displayed text
        text = self.combo_prev.currentText().strip()
        if text and text not in ("——", ""):
            idx = self.combo_prev.findText(text, Qt.MatchFlag.MatchContains)
            if idx >= 0:
                return self.combo_prev.itemData(idx)
        return None

    def get_duration(self) -> float:
        return self.spin_dur.value()

    def set_values(self, start_time, prev_key, duration):
        self.combo_start.blockSignals(True)
        idx = self.combo_start.findData(start_time)
        self.combo_start.setCurrentIndex(idx if idx >= 0 else 0)
        self.combo_start.blockSignals(False)

        self.combo_prev.blockSignals(True)
        idx2 = self.combo_prev.findData(prev_key)
        self.combo_prev.setCurrentIndex(idx2 if idx2 >= 0 else 0)
        self.combo_prev.blockSignals(False)

        self.spin_dur.blockSignals(True)
        self.spin_dur.setValue(duration)
        self.spin_dur.blockSignals(False)


# ── TimelineTable ──────────────────────────────────────────────────────────────

class TimelineTable(QTableWidget):
    rows_changed = pyqtSignal()

    def __init__(self, app_data: AppData):
        super().__init__()
        self.app_data = app_data
        self._beat_time: int = 30
        self._step: float = 0.5
        self._time_col_width: int = 34
        self._time_cols: list[float] = []

        self.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.setAlternatingRowColors(False)
        self.setShowGrid(True)
        self.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)

        # vertical header as drag handle
        vh = self.verticalHeader()
        vh.setVisible(True)
        vh.setFixedWidth(20)
        vh.setDefaultSectionSize(30)
        vh.setDefaultAlignment(Qt.AlignmentFlag.AlignCenter)

        # event-filter drag on the vertical-header viewport
        self._vh_drag_src: Optional[int] = None   # source main-row index
        self._vh_drag_line: Optional[int] = None  # current drop target main-row index
        vh.viewport().installEventFilter(self)
        vh.viewport().setMouseTracking(True)

        self._setup_fixed_columns()

    # ── wheel = zoom ───────────────────────────────────────────────────────────

    def wheelEvent(self, event):
        dy = event.angleDelta().y()
        if dy != 0 and self._time_cols:
            step = 3 if dy > 0 else -3
            new_w = max(16, min(120, self._time_col_width + step))
            if new_w != self._time_col_width:
                self._time_col_width = new_w
                for ci in range(FIXED_COLS, self.columnCount()):
                    self.setColumnWidth(ci, new_w)
            event.accept()
            return
        super().wheelEvent(event)

    # ── drag-drop row reordering via vertical-header event filter ────────────────

    def eventFilter(self, obj, event):
        vh_vp = self.verticalHeader().viewport()
        if obj is not vh_vp:
            return super().eventFilter(obj, event)

        et = event.type()

        if et == QEvent.Type.MouseButtonPress:
            row = self.verticalHeader().logicalIndexAt(event.pos().y())
            if self._row_type(row) == ROW_DETAIL:
                row -= 1
            main_rows = self._main_rows()
            if row >= 0 and row in main_rows:
                self._vh_drag_src = main_rows.index(row)
                vh_vp.setCursor(Qt.CursorShape.ClosedHandCursor)
            return False

        if et == QEvent.Type.MouseMove:
            if self._vh_drag_src is not None:
                row = self.verticalHeader().logicalIndexAt(event.pos().y())
                main_rows = self._main_rows()
                if row < 0:
                    self._vh_drag_line = len(main_rows)
                else:
                    if self._row_type(row) == ROW_DETAIL:
                        row += 1
                    self._vh_drag_line = (
                        main_rows.index(row) if row in main_rows else len(main_rows)
                    )
                vh_vp.update()  # repaint drop indicator
            return False

        if et == QEvent.Type.MouseButtonRelease:
            if self._vh_drag_src is not None:
                src = self._vh_drag_src
                tgt = self._vh_drag_line if self._vh_drag_line is not None else src
                self._vh_drag_src = None
                self._vh_drag_line = None
                vh_vp.unsetCursor()
                vh_vp.update()
                self._move_row_pair(src, tgt)
            return False

        if et == QEvent.Type.Paint and self._vh_drag_src is not None and self._vh_drag_line is not None:
            # let default paint happen first, then draw indicator line on top
            super().eventFilter(obj, event)
            from PyQt6.QtGui import QPainter, QPen
            main_rows = self._main_rows()
            tgt_i = self._vh_drag_line
            if tgt_i < len(main_rows):
                indicator_row = main_rows[tgt_i]
            else:
                # after last row
                indicator_row = self.rowCount()
            y = self.verticalHeader().sectionViewportPosition(indicator_row) if indicator_row < self.rowCount() else (
                self.verticalHeader().sectionViewportPosition(self.rowCount() - 1)
                + self.verticalHeader().sectionSize(self.rowCount() - 1)
            )
            painter = QPainter(vh_vp)
            pen = QPen(QColor("#2196F3"), 2)
            painter.setPen(pen)
            painter.drawLine(0, y, vh_vp.width(), y)
            painter.end()
            return True  # we handled the paint

        return False

    def _move_row_pair(self, source_i: int, target_i: int):
        """Rebuild table with the source main-row moved to target_i position."""
        main_rows = self._main_rows()
        if source_i < 0 or source_i >= len(main_rows):
            return
        # no-op when source is already at target position
        if target_i in (source_i, source_i + 1):
            return

        rows_data = self._collect_timeline_rows()
        expanded  = [not self.isRowHidden(mi + 1) for mi in main_rows]

        tr  = rows_data.pop(source_i)
        exp = expanded.pop(source_i)

        insert_i = target_i if target_i <= source_i else target_i - 1
        insert_i = max(0, min(insert_i, len(rows_data)))
        rows_data.insert(insert_i, tr)
        expanded.insert(insert_i, exp)

        # rebuild all rows
        self.setRowCount(0)
        for tr_item in rows_data:
            self.add_row(tr_item)

        # restore detail-row expanded state
        new_main = self._main_rows()
        for i, exp_state in enumerate(expanded):
            if i < len(new_main) and exp_state:
                self.showRow(new_main[i] + 1)

        self.updateGeometry()

    # ── size hint: table grows to fit all rows (no internal v-scroll) ──────────

    def sizeHint(self) -> QSize:
        h = self.horizontalHeader().height()
        for r in range(self.rowCount()):
            if not self.isRowHidden(r):
                h += self.rowHeight(r)
        # always reserve horizontal scrollbar height to avoid clipping last row
        h += self.horizontalScrollBar().sizeHint().height()
        return QSize(super().sizeHint().width(), h + 4)

    def minimumSizeHint(self) -> QSize:
        return self.sizeHint()

    # ── columns ────────────────────────────────────────────────────────────────

    def _setup_fixed_columns(self):
        self.setColumnCount(FIXED_COLS)
        self.setHorizontalHeaderLabels(["模块", "动作编号", "动作名称"])
        for col, w in [(COL_MODULE, 130), (COL_ACTION, 110), (COL_NAME, 110)]:
            self.horizontalHeader().setSectionResizeMode(col, QHeaderView.ResizeMode.Fixed)
            self.setColumnWidth(col, w)

    def rebuild_time_axis(self, beat_time: int, step: float):
        self._beat_time = beat_time
        self._step = step
        times: list[float] = []
        t = step
        while t <= beat_time:
            times.append(round(t, 1))
            t = round(t + step, 1)
        self._time_cols = times

        total_cols = FIXED_COLS + len(times)
        self.setColumnCount(total_cols)
        headers = ["模块", "动作编号", "动作名称"] + [
            str(int(tc)) if tc == int(tc) else str(tc) for tc in times
        ]
        self.setHorizontalHeaderLabels(headers)
        for i in range(FIXED_COLS, total_cols):
            self.setColumnWidth(i, self._time_col_width)
            self.horizontalHeader().setSectionResizeMode(i, QHeaderView.ResizeMode.Fixed)

        for r in range(self.rowCount()):
            if self._row_type(r) == ROW_DETAIL:
                self.setSpan(r, 0, 1, total_cols)
                dw = self.cellWidget(r, 0)
                if isinstance(dw, DetailWidget):
                    dw.populate_start_times(beat_time, step)

        self._refresh_all_colors()

    # ── row management ─────────────────────────────────────────────────────────

    def _row_type(self, row: int) -> str:
        item = self.item(row, COL_MODULE)
        return (item.data(ROLE_ROW_TYPE) or "") if item else ""

    def _main_rows(self) -> list[int]:
        return [r for r in range(self.rowCount()) if self._row_type(r) == ROW_MAIN]

    def add_row(self, timeline_row: Optional[TimelineRow] = None):
        tr = timeline_row or TimelineRow()
        main_idx = self.rowCount()
        self.insertRow(main_idx)
        self.insertRow(main_idx + 1)
        self.setRowHeight(main_idx, 30)
        self.setRowHeight(main_idx + 1, 60)
        self.hideRow(main_idx + 1)

        item0 = QTableWidgetItem()
        item0.setData(ROLE_ROW_TYPE, ROW_MAIN)
        self.setItem(main_idx, COL_MODULE, item0)
        item1 = QTableWidgetItem()
        item1.setData(ROLE_ROW_TYPE, ROW_MAIN)
        self.setItem(main_idx, COL_ACTION, item1)

        # vertical header: show drag-hint icon for main rows, blank for detail
        self.setVerticalHeaderItem(main_idx, QTableWidgetItem("⠿"))
        self.setVerticalHeaderItem(main_idx + 1, QTableWidgetItem(""))

        mod_combo = QComboBox()
        mod_combo.setStyleSheet(_COMBO_STYLE)
        for m in self.app_data.modules:
            mod_combo.addItem(f"{m.id} {m.name}", m.id)
        if tr.module_id:
            idx = mod_combo.findData(tr.module_id)
            if idx >= 0:
                mod_combo.setCurrentIndex(idx)
        self.setCellWidget(main_idx, COL_MODULE, mod_combo)

        act_combo = QComboBox()
        act_combo.setStyleSheet(_COMBO_STYLE)
        self.setCellWidget(main_idx, COL_ACTION, act_combo)

        name_lbl = QLabel()
        name_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        name_lbl.setStyleSheet(
            "QLabel{padding:2px 6px;border:1px solid #ddd;border-radius:3px;"
            "background:#f0f7ff;}"
            "QLabel:hover{background:#ddeeff;border-color:#90CAF9;}"
        )
        name_lbl.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setCellWidget(main_idx, COL_NAME, name_lbl)

        def _on_mod_change(_, mi=main_idx):
            self._refresh_action_combo(mi)
            self._on_row_edited(mi)
        mod_combo.currentIndexChanged.connect(_on_mod_change)

        def _on_act_change(_, mi=main_idx):
            self._refresh_name_label(mi)
            self._on_row_edited(mi)
        act_combo.currentIndexChanged.connect(_on_act_change)

        detail_idx = main_idx + 1
        d_item = QTableWidgetItem()
        d_item.setData(ROLE_ROW_TYPE, ROW_DETAIL)
        self.setItem(detail_idx, COL_MODULE, d_item)
        self.setSpan(detail_idx, 0, 1, max(self.columnCount(), FIXED_COLS))

        dw = DetailWidget()
        dw.populate_start_times(self._beat_time, self._step)
        self.setCellWidget(detail_idx, 0, dw)
        dw.set_values(tr.start_time, tr.prev_action_key, tr.duration)
        dw.changed.connect(lambda mi=main_idx: self._on_row_edited(mi))

        name_lbl.mousePressEvent = lambda e, mi=main_idx: self._toggle_detail(mi)

        self._refresh_action_combo(main_idx)
        if tr.action_no:
            key = f"{tr.module_id}{tr.action_no}"
            act_combo.setCurrentIndex(act_combo.findData(key))

        self._refresh_name_label(main_idx)
        self._refresh_prev_action_options()
        self._refresh_all_colors()
        self.updateGeometry()

    def _refresh_action_combo(self, main_idx: int):
        mod_combo = self.cellWidget(main_idx, COL_MODULE)
        act_combo = self.cellWidget(main_idx, COL_ACTION)
        if not isinstance(mod_combo, QComboBox) or not isinstance(act_combo, QComboBox):
            return
        mid = mod_combo.currentData()
        act_combo.blockSignals(True)
        act_combo.clear()
        if mid:
            for a in self.app_data.get_actions_for_module(mid):
                act_combo.addItem(a.key, a.key)
        act_combo.blockSignals(False)
        self._refresh_name_label(main_idx)

    def _refresh_name_label(self, main_idx: int):
        act_combo = self.cellWidget(main_idx, COL_ACTION)
        name_lbl = self.cellWidget(main_idx, COL_NAME)
        if not isinstance(act_combo, QComboBox) or not isinstance(name_lbl, QLabel):
            return
        key = act_combo.currentData()
        action = self.app_data.get_action(key) if key else None
        name_lbl.setText(action.name if action else "")

    def _toggle_detail(self, main_idx: int):
        detail_idx = main_idx + 1
        if self.isRowHidden(detail_idx):
            self.showRow(detail_idx)
        else:
            self.hideRow(detail_idx)
        self.updateGeometry()

    def _on_row_edited(self, main_idx: int):
        self._refresh_all_colors()
        self._refresh_prev_action_options()
        self.rows_changed.emit()

    def _refresh_prev_action_options(self):
        # All actions from the global action library (not just current-mode rows)
        all_actions = [
            (a.key, f"{a.key} {a.name}")
            for a in self.app_data.actions
        ]

        for mi in self._main_rows():
            dw = self.cellWidget(mi + 1, 0)
            act_combo = self.cellWidget(mi, COL_ACTION)
            if not isinstance(dw, DetailWidget) or not isinstance(act_combo, QComboBox):
                continue
            own_key = act_combo.currentData()
            dw.populate_prev_actions([(k, t) for k, t in all_actions if k != own_key])

    # ── color rendering ────────────────────────────────────────────────────────

    def _refresh_all_colors(self):
        rows_data = self._collect_timeline_rows()
        for i, mi in enumerate(self._main_rows()):
            tr = rows_data[i] if i < len(rows_data) else None
            mod_combo = self.cellWidget(mi, COL_MODULE)
            mid = mod_combo.currentData() if isinstance(mod_combo, QComboBox) else None
            color = module_color(mid) if mid else QColor("#eeeeee")
            color.setAlpha(200)
            bg = QBrush(color)
            eff_start = tr.effective_start(rows_data) if tr else None

            for ci in range(FIXED_COLS, self.columnCount()):
                item = self.item(mi, ci)
                if item is None:
                    item = QTableWidgetItem()
                    self.setItem(mi, ci, item)
                t = self._time_cols[ci - FIXED_COLS]
                filled = (
                    eff_start is not None
                    and tr is not None
                    and eff_start < t <= round(eff_start + tr.duration, 1)
                )
                item.setBackground(bg if filled else QBrush(QColor("#ffffff")))

    def _collect_timeline_rows(self) -> list[TimelineRow]:
        rows: list[TimelineRow] = []
        for mi in self._main_rows():
            mod_combo = self.cellWidget(mi, COL_MODULE)
            act_combo = self.cellWidget(mi, COL_ACTION)
            dw = self.cellWidget(mi + 1, 0)
            mid = mod_combo.currentData() or "" if isinstance(mod_combo, QComboBox) else ""
            key = act_combo.currentData() if isinstance(act_combo, QComboBox) else ""
            no = key[1:] if key and len(key) >= 3 else ""
            if isinstance(dw, DetailWidget):
                tr = TimelineRow(mid, no, dw.get_start_time(), dw.get_prev_action_key(), dw.get_duration())
            else:
                tr = TimelineRow(module_id=mid, action_no=no)
            rows.append(tr)
        return rows

    # ── public ─────────────────────────────────────────────────────────────────

    def get_mode_config(self, mode_id: str, beat_time: int, step: float) -> ModeConfig:
        cfg = ModeConfig(mode_id=mode_id, beat_time=beat_time, step=step)
        cfg.rows = self._collect_timeline_rows()
        return cfg

    def get_color_grid(self) -> list[list[Optional[str]]]:
        """Return hex color per (row_idx, time_col_idx) for Excel export."""
        rows_data = self._collect_timeline_rows()
        grid = []
        for i, mi in enumerate(self._main_rows()):
            tr = rows_data[i] if i < len(rows_data) else None
            mod_combo = self.cellWidget(mi, COL_MODULE)
            mid = mod_combo.currentData() if isinstance(mod_combo, QComboBox) else None
            hex_color = _PALETTE[(ord(mid) - ord("A")) % len(_PALETTE)] if mid else None
            eff_start = tr.effective_start(rows_data) if tr else None
            row_colors: list[Optional[str]] = []
            for t in self._time_cols:
                filled = (
                    eff_start is not None and tr is not None
                    and eff_start < t <= round(eff_start + tr.duration, 1)
                )
                row_colors.append(hex_color if filled else None)
            grid.append(row_colors)
        return grid

    def load_mode_config(self, cfg: ModeConfig):
        self.setRowCount(0)
        self.rebuild_time_axis(cfg.beat_time, cfg.step)
        for tr in cfg.rows:
            self.add_row(tr)
        self.updateGeometry()

    def on_data_changed(self):
        for mi in self._main_rows():
            self._refresh_action_combo(mi)
            self._refresh_name_label(mi)
        self._refresh_prev_action_options()
        self._refresh_all_colors()


# ── ModeSection ────────────────────────────────────────────────────────────────

class ModeSection(QFrame):
    remove_requested = pyqtSignal(object)
    changed = pyqtSignal()

    def __init__(self, app_data: AppData):
        super().__init__()
        self.app_data = app_data
        self.setFrameShape(QFrame.Shape.Box)
        self.setStyleSheet(
            "ModeSection{border:1px solid #b0c4de;border-radius:6px;background:#fafcff;}"
        )

        root = QVBoxLayout(self)
        root.setContentsMargins(10, 8, 10, 8)
        root.setSpacing(6)

        # ── header bar ────────────────────────────────────────────────────────
        header = QHBoxLayout()
        header.setSpacing(8)

        header.addWidget(QLabel("模式:"))
        self.combo_mode = QComboBox()
        self.combo_mode.setMinimumWidth(130)
        self.combo_mode.setStyleSheet(_COMBO_STYLE)
        header.addWidget(self.combo_mode)

        header.addSpacing(8)
        header.addWidget(QLabel("节拍时间"))
        self.spin_beat = QSpinBox()
        self.spin_beat.setRange(1, 200)
        self.spin_beat.setValue(30)
        self.spin_beat.setSuffix(" S")
        self.spin_beat.setMinimumWidth(80)
        header.addWidget(self.spin_beat)

        header.addWidget(QLabel("步长"))
        self.combo_step = QComboBox()
        self.combo_step.setStyleSheet(_COMBO_STYLE)
        for label, val in [("0.5 S", 0.5), ("0.1 S", 0.1), ("1 S", 1.0)]:
            self.combo_step.addItem(label, val)
        header.addWidget(self.combo_step)

        btn_ok = QPushButton("OK")
        btn_ok.setStyleSheet(_BTN_BLUE)
        btn_ok.clicked.connect(self._on_ok)
        header.addWidget(btn_ok)

        header.addStretch()

        btn_remove = QPushButton("× 删除此模式")
        btn_remove.setStyleSheet(
            "QPushButton{background:#FFEBEE;color:#c62828;border:1px solid #FFCDD2;"
            "border-radius:10px;padding:4px 12px;}"
            "QPushButton:hover{background:#FFCDD2;}"
        )
        btn_remove.clicked.connect(lambda: self.remove_requested.emit(self))
        header.addWidget(btn_remove)

        root.addLayout(header)

        # ── table ─────────────────────────────────────────────────────────────
        self.table = TimelineTable(app_data)
        root.addWidget(self.table)

        # ── add row ────────────────────────────────────────────────────────────
        footer = QHBoxLayout()
        btn_add = QPushButton("+ 添加动作行")
        btn_add.setStyleSheet(
            "QPushButton{background:#E8F5E9;color:#2E7D32;border:1px solid #A5D6A7;"
            "border-radius:10px;padding:4px 14px;}"
            "QPushButton:hover{background:#C8E6C9;}"
        )
        btn_add.clicked.connect(lambda: self.table.add_row())
        footer.addWidget(btn_add)
        footer.addStretch()
        root.addLayout(footer)

        self._refresh_mode_combo()

    def _on_ok(self):
        self.table.rebuild_time_axis(self.spin_beat.value(), self.combo_step.currentData())

    def _refresh_mode_combo(self):
        cur = self.combo_mode.currentData()
        self.combo_mode.blockSignals(True)
        self.combo_mode.clear()
        for m in self.app_data.modes:
            self.combo_mode.addItem(f"{m.id} {m.name}", m.id)
        if cur:
            idx = self.combo_mode.findData(cur)
            if idx >= 0:
                self.combo_mode.setCurrentIndex(idx)
        self.combo_mode.blockSignals(False)

    def on_data_changed(self):
        self._refresh_mode_combo()
        self.table.on_data_changed()

    def get_config(self) -> ModeConfig:
        mode_id = self.combo_mode.currentData() or ""
        return self.table.get_mode_config(mode_id, self.spin_beat.value(), self.combo_step.currentData())

    def get_color_grid(self) -> list[list[Optional[str]]]:
        return self.table.get_color_grid()

    def get_time_cols(self) -> list[float]:
        return list(self.table._time_cols)

    def load_config(self, cfg: ModeConfig):
        idx = self.combo_mode.findData(cfg.mode_id)
        if idx >= 0:
            self.combo_mode.setCurrentIndex(idx)
        self.spin_beat.setValue(cfg.beat_time)
        step_idx = self.combo_step.findData(cfg.step)
        if step_idx >= 0:
            self.combo_step.setCurrentIndex(step_idx)
        self.table.load_mode_config(cfg)


# ── scroll container that emits wheel delta for global zoom ────────────────────

class _ScrollInner(QWidget):
    zoom_delta = pyqtSignal(int)

    def wheelEvent(self, event):
        dy = event.angleDelta().y()
        if dy != 0:
            self.zoom_delta.emit(dy)
            event.accept()
            return
        super().wheelEvent(event)


# ── ModeEditorTab ──────────────────────────────────────────────────────────────

class ModeEditorTab(QWidget):
    def __init__(self, app_data: AppData):
        super().__init__()
        self.app_data = app_data
        self._sections: list[ModeSection] = []
        self._build_ui()

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(12, 10, 12, 10)
        root.setSpacing(8)

        # ── global toolbar ────────────────────────────────────────────────────
        bar = QHBoxLayout()
        bar.setSpacing(8)

        for label, slot in [
            ("Load JSON", self._on_load),
            ("Save JSON", self._on_save_json),
        ]:
            b = QPushButton(label)
            b.setStyleSheet(_BTN_LIGHT)
            b.clicked.connect(slot)
            bar.addWidget(b)

        btn_load_excel = QPushButton("Load Excel")
        btn_load_excel.setStyleSheet(
            "QPushButton{background:#FFF8E1;color:#E65100;border:1px solid #FFCC80;"
            "border-radius:12px;padding:5px 14px;font-weight:bold;}"
            "QPushButton:hover{background:#FFE0B2;}"
        )
        btn_load_excel.clicked.connect(self._on_load_excel)
        bar.addWidget(btn_load_excel)

        btn_excel = QPushButton("Save as Excel")
        btn_excel.setStyleSheet(
            "QPushButton{background:#E8F5E9;color:#1B5E20;border:1px solid #A5D6A7;"
            "border-radius:12px;padding:5px 14px;font-weight:bold;}"
            "QPushButton:hover{background:#C8E6C9;}"
        )
        btn_excel.clicked.connect(self._on_save_excel)
        bar.addWidget(btn_excel)

        bar.addStretch()

        btn_add_mode = QPushButton("+ 新增模式")
        btn_add_mode.setStyleSheet(_BTN_BLUE)
        btn_add_mode.clicked.connect(self._add_section)
        bar.addWidget(btn_add_mode)

        root.addLayout(bar)

        # ── scroll area ───────────────────────────────────────────────────────
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QFrame.Shape.NoFrame)

        self.inner = _ScrollInner()
        self.inner.zoom_delta.connect(self._zoom_all_tables)
        self.sections_layout = QVBoxLayout(self.inner)
        self.sections_layout.setContentsMargins(0, 0, 0, 0)
        self.sections_layout.setSpacing(12)
        self.sections_layout.addStretch(1)

        self.scroll.setWidget(self.inner)
        root.addWidget(self.scroll, 1)

    # ── sections ──────────────────────────────────────────────────────────────

    def _add_section(self, cfg: Optional[ModeConfig] = None) -> ModeSection:
        sec = ModeSection(self.app_data)
        sec.remove_requested.connect(self._remove_section)
        if cfg:
            sec.load_config(cfg)
        self._sections.append(sec)
        self.sections_layout.insertWidget(self.sections_layout.count() - 1, sec)
        return sec

    def _remove_section(self, sec: ModeSection):
        reply = QMessageBox.question(
            self, "确认删除", "确认删除此模式的编辑内容？",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if reply == QMessageBox.StandardButton.Yes:
            self._sections.remove(sec)
            sec.setParent(None)

    def _zoom_all_tables(self, delta: int):
        """Zoom every table's time columns by the same step."""
        if not self._sections:
            return
        ref_w = self._sections[0].table._time_col_width
        new_w = max(16, min(120, ref_w + (3 if delta > 0 else -3)))
        for sec in self._sections:
            sec.table._time_col_width = new_w
            for ci in range(FIXED_COLS, sec.table.columnCount()):
                sec.table.setColumnWidth(ci, new_w)

    # ── notifications from Tab 1 ──────────────────────────────────────────────

    def on_data_changed(self):
        for sec in self._sections:
            sec.on_data_changed()

    # ── file operations ───────────────────────────────────────────────────────

    def _on_load(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "导入模式编辑文件", "", "JSON 文件 (*.json);;所有文件 (*)"
        )
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                d = json.load(f)
            loaded = AppData.from_dict(d)
            # Clear existing sections
            for sec in list(self._sections):
                sec.setParent(None)
            self._sections.clear()
            # Also merge modes/modules/actions into app_data
            self.app_data.modes = loaded.modes or self.app_data.modes
            self.app_data.modules = loaded.modules or self.app_data.modules
            self.app_data.actions = loaded.actions or self.app_data.actions
            for cfg in loaded.mode_configs:
                self._add_section(cfg)
        except Exception as e:
            QMessageBox.critical(self, "导入失败", str(e))

    def _collect_all_configs(self):
        return [sec.get_config() for sec in self._sections]

    def _on_save_json(self):
        self.app_data.mode_configs = self._collect_all_configs()
        path, _ = QFileDialog.getSaveFileName(
            self, "保存模式文件", "mode_config.json", "JSON 文件 (*.json)"
        )
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(self.app_data.to_dict(), f, ensure_ascii=False, indent=2)
            QMessageBox.information(self, "保存成功", f"已保存至 {path}")
        except Exception as e:
            QMessageBox.critical(self, "保存失败", str(e))

    def _on_load_excel(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "导入 Excel 文件", "", "Excel 文件 (*.xlsx);;所有文件 (*)"
        )
        if not path:
            return
        try:
            from excel_export import import_from_excel
            configs = import_from_excel(path)
            if not configs:
                QMessageBox.warning(self, "导入提示", "未在文件中找到有效的模式数据。")
                return
            for sec in list(self._sections):
                sec.setParent(None)
            self._sections.clear()
            for cfg in configs:
                self._add_section(cfg)
            QMessageBox.information(self, "导入成功", f"已导入 {len(configs)} 个模式。")
        except Exception as e:
            QMessageBox.critical(self, "导入失败", str(e))

    def _on_save_excel(self):
        if not self._sections:
            QMessageBox.warning(self, "无数据", "请先添加并编辑至少一个模式。")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "导出 Excel", "timeline.xlsx", "Excel 文件 (*.xlsx)"
        )
        if not path:
            return
        try:
            from excel_export import export_to_excel
            sections_data = [
                (sec.get_config(), sec.get_color_grid(), sec.get_time_cols())
                for sec in self._sections
            ]
            export_to_excel(sections_data, self.app_data, path)
            QMessageBox.information(self, "导出成功", f"已导出至 {path}")
        except Exception as e:
            QMessageBox.critical(self, "导出失败", str(e))
