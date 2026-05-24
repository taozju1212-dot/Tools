from __future__ import annotations

from dataclasses import asdict, dataclass, field
from typing import Literal
from uuid import uuid4


MechanismType = Literal["linear_belt", "rotary", "leadscrew", "belt", "custom"]
Strategy = Literal["balanced", "eight_low_impact", "eight_stall_guard", "low_peak_speed", "low_impact", "symmetric"]


def new_id() -> str:
    return uuid4().hex[:10]


@dataclass
class Axis:
    id: str = field(default_factory=new_id)
    name: str = "Axis 1"
    mechanism_type: MechanismType = "linear_belt"
    motor_step_angle: float = 1.8
    microstep: int = 256
    gear_ratio: float = 1.0
    pulley_teeth: int = 20
    pulley_pitch_mm: float = 2.0
    lead_mm_per_rev: float = 6.35
    custom_microsteps_per_mm: float = 1280.0

    def normalized_type(self) -> str:
        if self.mechanism_type == "belt":
            return "linear_belt"
        return self.mechanism_type

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict) -> "Axis":
        values = {k: v for k, v in data.items() if k in cls.__dataclass_fields__}
        if "id" not in values or not values["id"]:
            values["id"] = new_id()
        if values.get("mechanism_type") == "belt":
            values["mechanism_type"] = "linear_belt"
        return cls(**values)


@dataclass
class MotionParams:
    vmax: float = 0.0
    tvmax_ms: float = 0.0
    a1: float = 0.0
    a2: float = 0.0
    amax: float = 0.0
    dmax: float = 0.0
    d2: float = 0.0
    d1: float = 0.0
    vstart: float = 0.0
    vstop: float = 0.0
    v1: float = 0.0
    v2: float = 0.0

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict) -> "MotionParams":
        return cls(**{k: v for k, v in data.items() if k in cls.__dataclass_fields__})


@dataclass
class DistanceCase:
    id: str = field(default_factory=new_id)
    distance: float = 500.0
    note: str = ""

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict) -> "DistanceCase":
        values = {k: v for k, v in data.items() if k in cls.__dataclass_fields__}
        if "id" not in values or not values["id"]:
            values["id"] = new_id()
        return cls(**values)


@dataclass
class MotionAction:
    id: str = field(default_factory=new_id)
    name: str = "Action 1"
    axis_id: str = ""
    params: MotionParams = field(default_factory=MotionParams)
    distances: list[DistanceCase] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "id": self.id,
            "name": self.name,
            "axis_id": self.axis_id,
            "params": self.params.to_dict(),
            "distances": [distance.to_dict() for distance in self.distances],
        }

    @classmethod
    def from_dict(cls, data: dict) -> "MotionAction":
        return cls(
            id=data.get("id") or new_id(),
            name=data.get("name", "Action 1"),
            axis_id=data.get("axis_id", ""),
            params=MotionParams.from_dict(data.get("params", {})),
            distances=[DistanceCase.from_dict(item) for item in data.get("distances", [])],
        )


@dataclass
class CompositeStep:
    id: str = field(default_factory=new_id)
    name: str = "STEP"
    note: str = ""
    delay_ms: float = 0.0
    action1_distance_id: str = ""
    action2_distance_id: str = ""
    action3_distance_id: str = ""

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict) -> "CompositeStep":
        values = {k: v for k, v in data.items() if k in cls.__dataclass_fields__}
        legacy_map = {
            "x_distance_id": "action1_distance_id",
            "y_distance_id": "action2_distance_id",
            "z_distance_id": "action3_distance_id",
        }
        for old_key, new_key in legacy_map.items():
            if old_key in data and new_key not in values:
                values[new_key] = data[old_key]
        if "id" not in values or not values["id"]:
            values["id"] = new_id()
        return cls(**values)


@dataclass
class CompositeAction:
    id: str = field(default_factory=new_id)
    number: str = ""
    name: str = "组合动作1"
    steps: list[CompositeStep] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "id": self.id,
            "number": self.number,
            "name": self.name,
            "steps": [step.to_dict() for step in self.steps],
        }

    @classmethod
    def from_dict(cls, data: dict) -> "CompositeAction":
        return cls(
            id=data.get("id") or new_id(),
            number=str(data.get("number", "")),
            name=data.get("name", "组合动作1"),
            steps=[CompositeStep.from_dict(item) for item in data.get("steps", [])],
        )


@dataclass
class AppState:
    project_name: str = "Untitled Project"
    fclk_hz: float = 16_000_000.0
    axes: list[Axis] = field(default_factory=list)
    actions: list[MotionAction] = field(default_factory=list)
    composite_actions: list[CompositeAction] = field(default_factory=list)
    current_project_path: str = ""

    def to_dict(self) -> dict:
        return {
            "schemaVersion": 2,
            "project_name": self.project_name,
            "fclk_hz": self.fclk_hz,
            "axes": [axis.to_dict() for axis in self.axes],
            "actions": [action.to_dict() for action in self.actions],
            "composite_actions": [action.to_dict() for action in self.composite_actions],
        }

    @classmethod
    def from_dict(cls, data: dict) -> "AppState":
        if "axis" in data:
            axis = Axis.from_dict(data.get("axis", {}))
            cases = data.get("cases", [])
            action = MotionAction(axis_id=axis.id)
            action.distances = [DistanceCase(distance=item.get("distance_mm", 500.0), note=item.get("note", "")) for item in cases] or [DistanceCase()]
            if cases:
                first = cases[0]
                action.params = MotionParams(
                    vmax=first.get("vmax_mm_s", 300.0),
                    a1=first.get("acc_mm_s2", 1000.0),
                    a2=first.get("acc_mm_s2", 1000.0),
                    amax=first.get("acc_mm_s2", 1000.0),
                    dmax=first.get("dec_mm_s2", 1000.0),
                    d2=first.get("dec_mm_s2", 1000.0),
                    d1=first.get("dec_mm_s2", 1000.0),
                    vstart=first.get("start_velocity_mm_s", 0.0),
                    vstop=first.get("end_velocity_mm_s", 0.0),
                )
            return cls(axes=[axis], actions=[action])
        composite_actions = [CompositeAction.from_dict(item) for item in data.get("composite_actions", [])]
        if not composite_actions and data.get("composite_steps"):
            composite_actions = [
                CompositeAction(
                    name="组合动作1",
                    steps=[CompositeStep.from_dict(item) for item in data.get("composite_steps", [])],
                )
            ]
        return cls(
            project_name=data.get("project_name", "Untitled Project"),
            fclk_hz=data.get("fclk_hz", 16_000_000.0),
            axes=[Axis.from_dict(item) for item in data.get("axes", [])],
            actions=[MotionAction.from_dict(item) for item in data.get("actions", [])],
            composite_actions=composite_actions,
        )


# Legacy structures kept for calculation tests and older project imports.
@dataclass
class MotionCase:
    distance_mm: float = 500.0
    vmax_mm_s: float = 300.0
    acc_mm_s2: float = 1000.0
    dec_mm_s2: float = 1000.0
    start_velocity_mm_s: float = 0.0
    end_velocity_mm_s: float = 0.0
    note: str = ""


@dataclass
class TProfileCase:
    acc_distance_microsteps: float = 4000.0
    acc_time_s: float = 0.05
    const_distance_microsteps: float = 591680.0
    dec_distance_microsteps: float = 4000.0
    dec_time_s: float = 0.05
    vstart_register: int = 0
    vstop_register: int = 10
