from __future__ import annotations
from dataclasses import dataclass, field
from typing import Optional, List


@dataclass
class Mode:
    id: str
    name: str


@dataclass
class Module:
    id: str
    name: str


@dataclass
class Action:
    module_id: str
    no: str  # "00"-"99"
    name: str

    @property
    def key(self) -> str:
        return f"{self.module_id}{self.no}"

    @property
    def display(self) -> str:
        return f"{self.module_id}{self.no} {self.name}"


@dataclass
class TimelineRow:
    module_id: str = ""
    action_no: str = ""
    start_time: Optional[float] = None
    prev_action_key: Optional[str] = None
    duration: float = 1.0

    @property
    def action_key(self) -> str:
        return f"{self.module_id}{self.action_no}"

    def effective_start(self, rows: List[TimelineRow], visited: Optional[set] = None) -> Optional[float]:
        if visited is None:
            visited = set()
        if self.action_key in visited:
            return None  # cycle detected
        visited.add(self.action_key)

        if self.prev_action_key:
            for r in rows:
                if r is not self and r.action_key == self.prev_action_key:
                    prev_start = r.effective_start(rows, visited)
                    if prev_start is not None:
                        return round(prev_start + r.duration, 1)
            return None
        return self.start_time


@dataclass
class ModeConfig:
    mode_id: str = ""
    beat_time: int = 30
    step: float = 0.5
    rows: List[TimelineRow] = field(default_factory=list)


@dataclass
class AppData:
    modes: List[Mode] = field(default_factory=list)
    modules: List[Module] = field(default_factory=list)
    actions: List[Action] = field(default_factory=list)
    mode_configs: List[ModeConfig] = field(default_factory=list)

    # ── serialization ──────────────────────────────────────────────────────────

    def to_dict(self) -> dict:
        return {
            "version": "1.0",
            "modes": [{"id": m.id, "name": m.name} for m in self.modes],
            "modules": [{"id": m.id, "name": m.name} for m in self.modules],
            "actions": [{"moduleId": a.module_id, "no": a.no, "name": a.name}
                        for a in self.actions],
            "modeConfigs": [
                {
                    "modeId": mc.mode_id,
                    "beatTime": mc.beat_time,
                    "step": mc.step,
                    "rows": [
                        {
                            "moduleId": r.module_id,
                            "actionNo": r.action_no,
                            "startTime": r.start_time,
                            "prevActionKey": r.prev_action_key,
                            "duration": r.duration,
                        }
                        for r in mc.rows
                    ],
                }
                for mc in self.mode_configs
            ],
        }

    @classmethod
    def from_dict(cls, d: dict) -> AppData:
        data = cls()
        data.modes = [Mode(m["id"], m["name"]) for m in d.get("modes", [])]
        data.modules = [Module(m["id"], m["name"]) for m in d.get("modules", [])]
        data.actions = [
            Action(a["moduleId"], a["no"], a["name"]) for a in d.get("actions", [])
        ]
        for mc in d.get("modeConfigs", []):
            cfg = ModeConfig(
                mode_id=mc["modeId"],
                beat_time=mc.get("beatTime", 30),
                step=mc.get("step", 0.5),
            )
            for r in mc.get("rows", []):
                cfg.rows.append(
                    TimelineRow(
                        module_id=r["moduleId"],
                        action_no=r["actionNo"],
                        start_time=r.get("startTime"),
                        prev_action_key=r.get("prevActionKey"),
                        duration=r.get("duration", 1.0),
                    )
                )
            data.mode_configs.append(cfg)
        return data

    # ── helpers ────────────────────────────────────────────────────────────────

    def get_actions_for_module(self, module_id: str) -> List[Action]:
        return [a for a in self.actions if a.module_id == module_id]

    def get_action(self, key: str) -> Optional[Action]:
        for a in self.actions:
            if a.key == key:
                return a
        return None

    def get_mode_config(self, mode_id: str) -> Optional[ModeConfig]:
        for mc in self.mode_configs:
            if mc.mode_id == mode_id:
                return mc
        return None

    def next_mode_id(self) -> Optional[str]:
        used = {m.id for m in self.modes}
        for c in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
            if c not in used:
                return c
        return None

    def next_module_id(self) -> Optional[str]:
        used = {m.id for m in self.modules}
        for c in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
            if c not in used:
                return c
        return None

    def next_action_no(self, module_id: str) -> Optional[str]:
        used = {a.no for a in self.actions if a.module_id == module_id}
        for i in range(100):
            no = f"{i:02d}"
            if no not in used:
                return no
        return None
