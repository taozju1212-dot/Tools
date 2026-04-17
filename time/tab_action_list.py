from __future__ import annotations
import json
from typing import Optional

from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QFont
from PyQt6.QtWidgets import (
    QWidget, QHBoxLayout, QVBoxLayout, QLabel, QLineEdit, QPushButton,
    QScrollArea, QFrame, QFileDialog, QMessageBox, QComboBox, QSizePolicy,
)

from models import AppData, Mode, Module, Action


# ─── small reusable row widget ─────────────────────────────────────────────────

class ItemRow(QWidget):
    removed = pyqtSignal(object)  # emits self
    name_changed = pyqtSignal()

    def __init__(self, id_text: str, name: str = ""):
        super().__init__()
        self.id_text = id_text
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 2, 0, 2)
        layout.setSpacing(6)

        self.lbl_id = QLabel(id_text)
        self.lbl_id.setFixedWidth(24)
        self.lbl_id.setAlignment(Qt.AlignmentFlag.AlignCenter)
        font = QFont()
        font.setBold(True)
        self.lbl_id.setFont(font)

        self.edit_name = QLineEdit(name)
        self.edit_name.setPlaceholderText("名称…")
        self.edit_name.textChanged.connect(self.name_changed)

        self.btn_del = QPushButton("✕")
        self.btn_del.setFixedSize(24, 24)
        self.btn_del.setStyleSheet(
            "QPushButton{color:#f44336;border:none;font-weight:bold;}"
            "QPushButton:hover{color:#d32f2f;}"
        )
        self.btn_del.clicked.connect(lambda: self.removed.emit(self))

        layout.addWidget(self.lbl_id)
        layout.addWidget(self.edit_name, 1)
        layout.addWidget(self.btn_del)

    def get_name(self) -> str:
        return self.edit_name.text().strip()


# ─── panel for modes / modules ─────────────────────────────────────────────────

class LetterItemPanel(QFrame):
    """Panel for 模式编号 or 模块编号 (A-Z letters)."""
    changed = pyqtSignal()

    def __init__(self, title: str):
        super().__init__()
        self.setFrameShape(QFrame.Shape.Box)
        self.setStyleSheet("QFrame{border:1px solid #ddd;border-radius:6px;background:#fff;}")
        self._rows: list[ItemRow] = []

        root = QVBoxLayout(self)
        root.setContentsMargins(12, 10, 12, 10)
        root.setSpacing(8)

        lbl = QLabel(title)
        lbl.setStyleSheet("font-weight:bold;font-size:13px;border:none;")
        root.addWidget(lbl)

        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setFrameShape(QFrame.Shape.NoFrame)
        self.scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

        self.inner = QWidget()
        self.inner_layout = QVBoxLayout(self.inner)
        self.inner_layout.setContentsMargins(0, 0, 0, 0)
        self.inner_layout.setSpacing(2)
        self.inner_layout.addStretch(1)

        self.scroll_area.setWidget(self.inner)
        root.addWidget(self.scroll_area, 1)

        self.btn_add = QPushButton("+ Add")
        self.btn_add.setStyleSheet(
            "QPushButton{background:#2196F3;color:white;border-radius:14px;"
            "padding:6px 20px;font-weight:bold;}"
            "QPushButton:hover{background:#1976D2;}"
        )
        self.btn_add.clicked.connect(self._on_add)
        root.addWidget(self.btn_add, 0, Qt.AlignmentFlag.AlignHCenter)

    def _next_id(self) -> Optional[str]:
        used = {r.id_text for r in self._rows}
        for c in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
            if c not in used:
                return c
        return None

    def _on_add(self):
        nxt = self._next_id()
        if nxt is None:
            return
        self._add_row(nxt, "")
        self.changed.emit()

    def _add_row(self, id_text: str, name: str):
        row = ItemRow(id_text, name)
        row.removed.connect(self._on_remove)
        row.name_changed.connect(self.changed)
        self._rows.append(row)
        self.inner_layout.insertWidget(self.inner_layout.count() - 1, row)

    def _on_remove(self, row: ItemRow):
        reply = QMessageBox.question(
            self, "确认删除", f"确认删除编号 {row.id_text}？",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if reply == QMessageBox.StandardButton.Yes:
            self._rows.remove(row)
            row.setParent(None)
            self.changed.emit()

    def get_items(self) -> list[tuple[str, str]]:
        return [(r.id_text, r.get_name()) for r in self._rows]

    def load_items(self, items: list[tuple[str, str]]):
        for row in list(self._rows):
            row.setParent(None)
        self._rows.clear()
        for id_text, name in items:
            self._add_row(id_text, name)


# ─── panel for actions (00-99, per module) ─────────────────────────────────────

class ActionPanel(QFrame):
    changed = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.setFrameShape(QFrame.Shape.Box)
        self.setStyleSheet("QFrame{border:1px solid #ddd;border-radius:6px;background:#fff;}")
        self._rows: list[ItemRow] = []

        root = QVBoxLayout(self)
        root.setContentsMargins(12, 10, 12, 10)
        root.setSpacing(8)

        lbl = QLabel("动作编号   (00-99)")
        lbl.setStyleSheet("font-weight:bold;font-size:13px;border:none;")
        root.addWidget(lbl)

        # module selector row
        module_row = QHBoxLayout()
        self.module_combo = QComboBox()
        self.module_combo.setMinimumWidth(80)
        self.module_combo.currentIndexChanged.connect(self._on_module_changed)
        module_row.addWidget(QLabel("模块:"))
        module_row.addWidget(self.module_combo, 1)
        module_row.addStretch()
        root.addLayout(module_row)

        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setFrameShape(QFrame.Shape.NoFrame)
        self.scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

        self.inner = QWidget()
        self.inner_layout = QVBoxLayout(self.inner)
        self.inner_layout.setContentsMargins(0, 0, 0, 0)
        self.inner_layout.setSpacing(2)
        self.inner_layout.addStretch(1)

        self.scroll_area.setWidget(self.inner)
        root.addWidget(self.scroll_area, 1)

        self.btn_add = QPushButton("+ Add")
        self.btn_add.setStyleSheet(
            "QPushButton{background:#2196F3;color:white;border-radius:14px;"
            "padding:6px 20px;font-weight:bold;}"
            "QPushButton:hover{background:#1976D2;}"
        )
        self.btn_add.clicked.connect(self._on_add)
        root.addWidget(self.btn_add, 0, Qt.AlignmentFlag.AlignHCenter)

        self._module_actions: dict[str, list[ItemRow]] = {}  # module_id -> rows

    def _current_module_id(self) -> Optional[str]:
        idx = self.module_combo.currentIndex()
        if idx < 0:
            return None
        return self.module_combo.itemData(idx)

    def update_modules(self, modules: list[tuple[str, str]]):
        """Called when module list changes. Preserve existing action rows."""
        current = self._current_module_id()
        self.module_combo.blockSignals(True)
        self.module_combo.clear()
        for mid, mname in modules:
            self.module_combo.addItem(f"{mid} {mname}", mid)
            if mid not in self._module_actions:
                self._module_actions[mid] = []
        self.module_combo.blockSignals(False)

        # remove stale module keys
        valid_ids = {m[0] for m in modules}
        for stale in [k for k in self._module_actions if k not in valid_ids]:
            for r in self._module_actions[stale]:
                r.setParent(None)
            del self._module_actions[stale]

        # restore selection
        if current:
            idx = self.module_combo.findData(current)
            if idx >= 0:
                self.module_combo.setCurrentIndex(idx)
        self._refresh_list()

    def _on_module_changed(self):
        self._refresh_list()

    def _refresh_list(self):
        # Hide all rows then show current module's rows
        for rows in self._module_actions.values():
            for r in rows:
                r.setVisible(False)
        mid = self._current_module_id()
        if mid and mid in self._module_actions:
            for r in self._module_actions[mid]:
                r.setVisible(True)
                if r.parent() != self.inner:
                    self.inner_layout.insertWidget(self.inner_layout.count() - 1, r)

    def _next_no(self, module_id: str) -> Optional[str]:
        # id_text is like "A00"; extract the 2-digit number part
        used = {r.id_text[len(module_id):] for r in self._module_actions.get(module_id, [])}
        for i in range(100):
            no = f"{i:02d}"
            if no not in used:
                return no
        return None

    def _on_add(self):
        mid = self._current_module_id()
        if mid is None:
            return
        nxt = self._next_no(mid)
        if nxt is None:
            return
        label = f"{mid}{nxt}"
        row = ItemRow(label, "")
        row.removed.connect(lambda r, m=mid: self._on_remove(m, r))
        row.name_changed.connect(self.changed)
        self._module_actions[mid].append(row)
        self.inner_layout.insertWidget(self.inner_layout.count() - 1, row)
        self.changed.emit()

    def _on_remove(self, module_id: str, row: ItemRow):
        reply = QMessageBox.question(
            self, "确认删除", f"确认删除动作 {row.id_text}？",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if reply == QMessageBox.StandardButton.Yes:
            self._module_actions[module_id].remove(row)
            row.setParent(None)
            self.changed.emit()

    def get_all_actions(self) -> list[tuple[str, str, str]]:
        """Returns list of (module_id, no, name)."""
        result = []
        for mid, rows in self._module_actions.items():
            for r in rows:
                id_text = r.id_text  # e.g. "A00"
                no = id_text[1:]     # "00"
                result.append((mid, no, r.get_name()))
        return result

    def load_actions(self, actions: list[tuple[str, str, str]]):
        """Load (module_id, no, name) tuples."""
        for rows in self._module_actions.values():
            for r in rows:
                r.setParent(None)
        self._module_actions = {mid: [] for mid in self._module_actions}

        for mid, no, name in actions:
            if mid not in self._module_actions:
                self._module_actions[mid] = []
            label = f"{mid}{no}"
            row = ItemRow(label, name)
            row.removed.connect(lambda r, m=mid: self._on_remove(m, r))
            row.name_changed.connect(self.changed)
            self._module_actions[mid].append(row)
            # Don't add to layout yet; _refresh_list will do it
        self._refresh_list()


# ─── Tab 1 ─────────────────────────────────────────────────────────────────────

class ActionListTab(QWidget):
    data_changed = pyqtSignal()

    def __init__(self, app_data: AppData):
        super().__init__()
        self.app_data = app_data
        self._build_ui()

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(16, 16, 16, 16)
        root.setSpacing(12)

        # Three panels
        panels_row = QHBoxLayout()
        panels_row.setSpacing(12)

        self.modes_panel = LetterItemPanel("模式编号   (A-Z)")
        self.modules_panel = LetterItemPanel("模块编号   (A-Z)")
        self.actions_panel = ActionPanel()

        self.modes_panel.changed.connect(self._on_changed)
        self.modules_panel.changed.connect(self._on_modules_changed)
        self.actions_panel.changed.connect(self._on_changed)

        for p in (self.modes_panel, self.modules_panel, self.actions_panel):
            p.setMinimumWidth(200)
            panels_row.addWidget(p, 1)

        root.addLayout(panels_row, 1)

        # Bottom buttons
        btn_row = QHBoxLayout()
        btn_row.setSpacing(12)
        btn_row.addStretch()
        for label, slot in [
            ("Load", self._on_load),
            ("Save", self._on_save),
            ("Save as", self._on_save_as),
            ("Clear", self._on_clear),
        ]:
            btn = QPushButton(label)
            btn.setMinimumWidth(90)
            btn.setStyleSheet(
                "QPushButton{background:#E3F2FD;color:#1565C0;border:1px solid #90CAF9;"
                "border-radius:14px;padding:6px 18px;font-weight:bold;}"
                "QPushButton:hover{background:#BBDEFB;}"
            )
            btn.clicked.connect(slot)
            btn_row.addWidget(btn)
        btn_row.addStretch()
        root.addLayout(btn_row)

    # ── data sync ──────────────────────────────────────────────────────────────

    def _sync_to_model(self):
        self.app_data.modes = [Mode(id_, name) for id_, name in self.modes_panel.get_items()]
        self.app_data.modules = [Module(id_, name) for id_, name in self.modules_panel.get_items()]
        self.app_data.actions = [
            Action(mid, no, name)
            for mid, no, name in self.actions_panel.get_all_actions()
        ]

    def _sync_from_model(self):
        self.modes_panel.load_items([(m.id, m.name) for m in self.app_data.modes])
        self.modules_panel.load_items([(m.id, m.name) for m in self.app_data.modules])
        self.actions_panel.update_modules([(m.id, m.name) for m in self.app_data.modules])
        self.actions_panel.load_actions(
            [(a.module_id, a.no, a.name) for a in self.app_data.actions]
        )

    def _on_changed(self):
        self._sync_to_model()
        self.data_changed.emit()

    def _on_modules_changed(self):
        self._sync_to_model()
        # propagate module list to action panel
        self.actions_panel.update_modules([(m.id, m.name) for m in self.app_data.modules])
        self.data_changed.emit()

    # ── buttons ────────────────────────────────────────────────────────────────

    def _on_load(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "导入一级动作列表", "", "JSON 文件 (*.json);;所有文件 (*)"
        )
        if not path:
            return
        try:
            import json
            with open(path, "r", encoding="utf-8") as f:
                d = json.load(f)
            new_data = AppData.from_dict(d)
            self.app_data.modes = new_data.modes
            self.app_data.modules = new_data.modules
            self.app_data.actions = new_data.actions
            self._sync_from_model()
            self.data_changed.emit()
        except Exception as e:
            QMessageBox.critical(self, "导入失败", str(e))

    def _on_save(self):
        self._sync_to_model()
        QMessageBox.information(self, "已保存", "数据已保存至当前软件（未写入文件）。")

    def _on_save_as(self):
        self._sync_to_model()
        path, _ = QFileDialog.getSaveFileName(
            self, "保存一级动作列表", "action_list.json", "JSON 文件 (*.json)"
        )
        if not path:
            return
        try:
            import json
            with open(path, "w", encoding="utf-8") as f:
                json.dump(self.app_data.to_dict(), f, ensure_ascii=False, indent=2)
            QMessageBox.information(self, "保存成功", f"已保存至 {path}")
        except Exception as e:
            QMessageBox.critical(self, "保存失败", str(e))

    def _on_clear(self):
        reply = QMessageBox.question(
            self, "确认清空",
            "确认清空所有模式、模块和动作数据？此操作不可撤销。",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if reply == QMessageBox.StandardButton.Yes:
            self.app_data.modes.clear()
            self.app_data.modules.clear()
            self.app_data.actions.clear()
            self._sync_from_model()
            self.data_changed.emit()
