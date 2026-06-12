from __future__ import annotations

from dataclasses import asdict, dataclass, field
from typing import Literal
from uuid import uuid4


MechanismType = Literal["linear_belt", "rotary", "leadscrew", "belt", "custom", "timer"]
Strategy = Literal["balanced", "eight_low_impact", "eight_stall_guard", "low_peak_speed", "low_impact", "symmetric"]


def new_id() -> str:
    return uuid4().hex[:10]


@dataclass
class AxisPoint:
    id: str = field(default_factory=new_id)
    name: str = "零位"
    position_mm: float = 0.0

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict) -> "AxisPoint":
        values = {k: v for k, v in data.items() if k in cls.__dataclass_fields__}
        if "id" not in values or not values["id"]:
            values["id"] = new_id()
        return cls(**values)


@dataclass
class Axis:
    id: str = field(default_factory=new_id)
    number: str = ""
    name: str = "Axis 1"
    mechanism_type: MechanismType = "linear_belt"
    motor_step_angle: float = 1.8
    microstep: int = 256
    gear_ratio: float = 1.0
    pulley_teeth: int = 20
    pulley_pitch_mm: float = 2.0
    lead_mm_per_rev: float = 6.35
    custom_microsteps_per_mm: float = 1280.0
    points: list[AxisPoint] = field(default_factory=list)

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
        values["points"] = [AxisPoint.from_dict(item) for item in data.get("points", [])]
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
    number: str = ""
    distance: float = 0.0
    duration_s: float = 0.0
    note: str = ""
    start_point_id: str = ""
    end_point_id: str = ""

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict) -> "DistanceCase":
        values = {k: v for k, v in data.items() if k in cls.__dataclass_fields__}
        if "id" not in values or not values["id"]:
            values["id"] = new_id()
        return cls(**values)


@dataclass
class CompositeStepItem:
    id: str = field(default_factory=new_id)
    kind: Literal["motion", "delay"] = "motion"
    axis_id: str = ""
    distance_id: str = ""
    delay_ms: float = 0.0

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict) -> "CompositeStepItem":
        values = {k: v for k, v in data.items() if k in cls.__dataclass_fields__}
        if "id" not in values or not values["id"]:
            values["id"] = new_id()
        if values.get("kind") not in {"motion", "delay"}:
            values["kind"] = "motion"
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
    items: list[CompositeStepItem] = field(default_factory=list)

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict) -> "CompositeStep":
        values = {k: v for k, v in data.items() if k in cls.__dataclass_fields__}
        values["items"] = [CompositeStepItem.from_dict(item) for item in data.get("items", [])]
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
        if not values["items"]:
            for key in ("action1_distance_id", "action2_distance_id", "action3_distance_id"):
                if values.get(key):
                    values["items"].append(CompositeStepItem(kind="motion", distance_id=values[key]))
            if float(values.get("delay_ms") or 0) > 0:
                values["items"].append(CompositeStepItem(kind="delay", delay_ms=float(values.get("delay_ms") or 0)))
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
            action.distances = [DistanceCase(distance=item.get("distance_mm", 0.0), note=item.get("note", "")) for item in cases]
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
            state = cls(axes=[axis], actions=[action])
            state.migrate_coordinate_points()
            return state
        composite_actions = [CompositeAction.from_dict(item) for item in data.get("composite_actions", [])]
        if not composite_actions and data.get("composite_steps"):
            composite_actions = [
                CompositeAction(
                    name="组合动作1",
                    steps=[CompositeStep.from_dict(item) for item in data.get("composite_steps", [])],
                )
            ]
        state = cls(
            project_name=data.get("project_name", "Untitled Project"),
            fclk_hz=data.get("fclk_hz", 16_000_000.0),
            axes=[Axis.from_dict(item) for item in data.get("axes", [])],
            actions=[MotionAction.from_dict(item) for item in data.get("actions", [])],
            composite_actions=composite_actions,
        )
        state.migrate_coordinate_points()
        return state

    def migrate_coordinate_points(self) -> None:
        axis_by_id = {axis.id: axis for axis in self.axes}
        for axis in self.axes:
            if axis.normalized_type() == "timer":
                axis.points = []
                continue
            if not axis.points:
                axis.points.append(AxisPoint(name="零位", position_mm=0.0))
            zero = next((point for point in axis.points if abs(point.position_mm) < 1e-12), None)
            if not zero:
                zero = AxisPoint(name="零位", position_mm=0.0)
                axis.points.insert(0, zero)
            elif not zero.name:
                zero.name = "零位"

        for action in self.actions:
            axis = axis_by_id.get(action.axis_id)
            if not axis:
                continue
            if axis.normalized_type() == "timer":
                continue
            zero = next((point for point in axis.points if abs(point.position_mm) < 1e-12), axis.points[0])
            point_by_id = {point.id: point for point in axis.points}
            for idx, distance in enumerate(action.distances, start=1):
                if distance.start_point_id in point_by_id and distance.end_point_id in point_by_id:
                    distance.distance = point_by_id[distance.end_point_id].position_mm - point_by_id[distance.start_point_id].position_mm
                    continue
                point_name = distance.note.strip() or f"坐标{idx}"
                point = AxisPoint(name=point_name, position_mm=distance.distance)
                axis.points.append(point)
                distance.start_point_id = zero.id
                distance.end_point_id = point.id

        distance_axis: dict[str, str] = {}
        for action in self.actions:
            for distance in action.distances:
                distance_axis[distance.id] = action.axis_id
        for composite in self.composite_actions:
            for step in composite.steps:
                if not step.items:
                    for key in ("action1_distance_id", "action2_distance_id", "action3_distance_id"):
                        distance_id = getattr(step, key)
                        if distance_id:
                            step.items.append(CompositeStepItem(kind="motion", axis_id=distance_axis.get(distance_id, ""), distance_id=distance_id))
                    if step.delay_ms > 0:
                        step.items.append(CompositeStepItem(kind="delay", delay_ms=step.delay_ms))
                for item in step.items:
                    if item.kind == "motion" and item.distance_id and not item.axis_id:
                        item.axis_id = distance_axis.get(item.distance_id, "")


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
