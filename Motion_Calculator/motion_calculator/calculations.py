from __future__ import annotations

from dataclasses import dataclass
from math import sqrt

from .models import Axis, CompositeStep, DistanceCase, MotionAction, MotionCase, MotionParams, Strategy, TProfileCase


VMAX_MAX = 8_388_607
AMAX_MAX = 262_143
DMAX_MAX = 262_143
TVMAX_MAX = 65_535
XTARGET_MIN = -(2**31)
XTARGET_MAX = 2**31 - 1


@dataclass
class AxisDerived:
    count: float
    unit_per_motor_rev: float
    microsteps_per_unit: float
    unit_name: str
    velocity_unit: str
    acceleration_unit: str


@dataclass
class RegisterValue:
    raw: float
    value: int
    hex_value: str
    limit: tuple[int, int]


@dataclass
class Segment:
    phase: str
    start_time: float
    end_time: float
    start_velocity: float
    end_velocity: float
    acceleration: float
    distance: float


@dataclass
class MotionResult:
    profile_type: str
    reached_vmax: bool
    vpeak: float
    acc_time: float
    const_time: float
    dec_time: float
    total_time: float
    acc_distance: float
    const_distance: float
    dec_distance: float
    segments: list[Segment]
    vmax_register: RegisterValue
    tvmax_register: RegisterValue
    amax_register: RegisterValue
    dmax_register: RegisterValue
    xtarget_register: RegisterValue

    @property
    def vpeak_mm_s(self) -> float:
        return self.vpeak

    @property
    def acc_time_s(self) -> float:
        return self.acc_time

    @property
    def const_time_s(self) -> float:
        return self.const_time

    @property
    def dec_time_s(self) -> float:
        return self.dec_time

    @property
    def total_time_s(self) -> float:
        return self.total_time

    @property
    def acc_distance_mm(self) -> float:
        return self.acc_distance

    @property
    def const_distance_mm(self) -> float:
        return self.const_distance

    @property
    def dec_distance_mm(self) -> float:
        return self.dec_distance


ForwardResult = MotionResult


@dataclass
class TargetTimeResult:
    strategy: str
    vmax_mm_s: float
    acc_mm_s2: float
    dec_mm_s2: float
    forward: MotionResult
    params: MotionParams
    time_error_s: float = 0.0


@dataclass
class CompositeMoveResult:
    axis_name: str
    distance_note: str
    distance: float
    total_time: float


@dataclass
class CompositeStepResult:
    step: CompositeStep
    move_results: list[CompositeMoveResult]
    motion_time: float
    delay_time: float
    total_time: float


@dataclass
class CompositeMotionResult:
    steps: list[CompositeStepResult]
    total_time: float


@dataclass
class TProfileResult:
    acc_rps2: float
    vmax_rps: float
    dec_rps2: float
    const_time_s: float
    total_time_s: float
    amax_register: RegisterValue
    vmax_register: RegisterValue
    dmax_register: RegisterValue
    xtarget_register: RegisterValue


def validate_axis(axis: Axis) -> list[str]:
    messages: list[str] = []
    if axis.motor_step_angle <= 0:
        messages.append("电机步距角必须大于 0")
    if axis.microstep not in {1, 2, 4, 8, 16, 32, 64, 128, 256}:
        messages.append("微步细分必须为 1/2/4/8/16/32/64/128/256")
    if axis.gear_ratio <= 0:
        messages.append("减速比必须大于 0")
    mechanism = axis.normalized_type()
    if mechanism == "linear_belt" and (axis.pulley_teeth <= 0 or axis.pulley_pitch_mm <= 0):
        messages.append("同步轮齿数和齿距必须大于 0")
    if mechanism == "leadscrew" and axis.lead_mm_per_rev <= 0:
        messages.append("丝杆导程必须大于 0")
    if mechanism == "custom" and axis.custom_microsteps_per_mm <= 0:
        messages.append("自定义 microsteps/mm 必须大于 0")
    return messages


def derive_axis(axis: Axis) -> AxisDerived:
    messages = validate_axis(axis)
    if messages:
        raise ValueError("; ".join(messages))
    count = 360.0 / axis.motor_step_angle * axis.microstep
    mechanism = axis.normalized_type()
    if mechanism == "linear_belt":
        unit_per_motor_rev = axis.pulley_teeth * axis.pulley_pitch_mm / axis.gear_ratio
        unit = "mm"
    elif mechanism == "leadscrew":
        unit_per_motor_rev = axis.lead_mm_per_rev / axis.gear_ratio
        unit = "mm"
    elif mechanism == "rotary":
        unit_per_motor_rev = 360.0 / axis.gear_ratio
        unit = "deg"
    else:
        unit_per_motor_rev = count / axis.custom_microsteps_per_mm
        unit = "mm"
    return AxisDerived(
        count=count,
        unit_per_motor_rev=unit_per_motor_rev,
        microsteps_per_unit=count / unit_per_motor_rev,
        unit_name=unit,
        velocity_unit=f"{unit}/s",
        acceleration_unit=f"{unit}/s²",
    )


def _reg(raw: float, limit: tuple[int, int]) -> RegisterValue:
    value = int(round(raw))
    width = max(1, len(f"{limit[1]:X}"))
    return RegisterValue(raw=raw, value=value, hex_value=f"0x{value & ((1 << 32) - 1):0{width}X}", limit=limit)


def velocity_to_register(axis: Axis, fclk_hz: float, velocity: float) -> RegisterValue:
    derived = derive_axis(axis)
    rps = velocity / derived.unit_per_motor_rev
    return _reg(rps * derived.count * 2**24 / fclk_hz, (0, VMAX_MAX))


def register_to_velocity(axis: Axis, fclk_hz: float, register: float) -> float:
    derived = derive_axis(axis)
    rps = register * fclk_hz / 2**24 / derived.count
    return rps * derived.unit_per_motor_rev


def motor_rpm_to_velocity(axis: Axis, rpm: float) -> float:
    derived = derive_axis(axis)
    return rpm / 60.0 * derived.unit_per_motor_rev


def velocity_to_motor_rpm(axis: Axis, velocity: float) -> float:
    derived = derive_axis(axis)
    return velocity / derived.unit_per_motor_rev * 60.0


def register_to_motor_rpm(axis: Axis, fclk_hz: float, register: float) -> float:
    derived = derive_axis(axis)
    rps = register * fclk_hz / 2**24 / derived.count
    return rps * 60.0


def acceleration_to_register(axis: Axis, fclk_hz: float, acceleration: float, maximum: int = AMAX_MAX) -> RegisterValue:
    derived = derive_axis(axis)
    rps2 = acceleration / derived.unit_per_motor_rev
    return _reg(rps2 * derived.count * 2**41 / fclk_hz**2, (0, maximum))


def register_to_acceleration(axis: Axis, fclk_hz: float, register: float) -> float:
    derived = derive_axis(axis)
    rps2 = register * fclk_hz**2 / 2**41 / derived.count
    return rps2 * derived.unit_per_motor_rev


def motor_rpm_s_to_acceleration(axis: Axis, rpm_s: float) -> float:
    derived = derive_axis(axis)
    return rpm_s / 60.0 * derived.unit_per_motor_rev


def acceleration_to_motor_rpm_s(axis: Axis, acceleration: float) -> float:
    derived = derive_axis(axis)
    return acceleration / derived.unit_per_motor_rev * 60.0


def register_to_motor_rpm_s(axis: Axis, fclk_hz: float, register: float) -> float:
    derived = derive_axis(axis)
    rps2 = register * fclk_hz**2 / 2**41 / derived.count
    return rps2 * 60.0


def tvmax_to_register(fclk_hz: float, tvmax_ms: float) -> RegisterValue:
    return _reg(max(0.0, tvmax_ms) / 1000.0 * fclk_hz / 512.0, (0, TVMAX_MAX))


def register_to_tvmax_ms(fclk_hz: float, register: float) -> float:
    return max(0.0, register) * 512.0 / fclk_hz * 1000.0


def distance_to_register(axis: Axis, distance: float) -> RegisterValue:
    derived = derive_axis(axis)
    return _reg(distance * derived.microsteps_per_unit, (XTARGET_MIN, XTARGET_MAX))


def register_to_distance(axis: Axis, register: float) -> float:
    derived = derive_axis(axis)
    return register / derived.microsteps_per_unit


def _segments_for_accel(start_v: float, end_v: float, stages: list[tuple[float, float]], start_time: float, phase: str) -> list[Segment]:
    segments: list[Segment] = []
    t = start_time
    current = start_v
    for target, accel in stages:
        if accel <= 0:
            raise ValueError("加速度和减速度必须大于 0")
        if end_v >= start_v:
            next_v = min(target, end_v)
            if next_v <= current:
                continue
            dt = (next_v - current) / accel
            dist = (current + next_v) * 0.5 * dt
            segments.append(Segment(phase, t, t + dt, current, next_v, accel, dist))
            current = next_v
            t += dt
        else:
            next_v = max(target, end_v)
            if next_v >= current:
                continue
            dt = (current - next_v) / accel
            dist = (current + next_v) * 0.5 * dt
            segments.append(Segment(phase, t, t + dt, current, next_v, -accel, dist))
            current = next_v
            t += dt
    return segments


def _accel_segments(params: MotionParams, vpeak: float, start_time: float = 0.0) -> list[Segment]:
    stages = [
        (min(params.v1, vpeak), params.a1),
        (min(params.v2, vpeak), params.a2),
        (vpeak, params.amax),
    ]
    return _segments_for_accel(params.vstart, vpeak, stages, start_time, "加速")


def _decel_segments(params: MotionParams, vpeak: float, start_time: float) -> list[Segment]:
    stages = [
        (max(params.v2, params.vstop), params.dmax),
        (max(params.v1, params.vstop), params.d2),
        (params.vstop, params.d1),
    ]
    return _segments_for_accel(vpeak, params.vstop, stages, start_time, "减速")


def _sum_distance(segments: list[Segment]) -> float:
    return sum(segment.distance for segment in segments)


def _sum_time(segments: list[Segment]) -> float:
    return sum(segment.end_time - segment.start_time for segment in segments)


def _profile_distance_with_tvmax(params: MotionParams, vpeak: float, tvmax_s: float) -> float:
    accel_distance = _sum_distance(_accel_segments(params, vpeak))
    decel_distance = _sum_distance(_decel_segments(params, vpeak, 0))
    return accel_distance + decel_distance + vpeak * tvmax_s


def calculate_motion(axis: Axis, fclk_hz: float, params: MotionParams, distance: DistanceCase) -> MotionResult:
    if distance.distance < 0:
        raise ValueError("距离不能为负")
    if params.vmax <= 0:
        raise ValueError("Vmax 必须大于 0")
    if max(params.vstart, params.vstop, params.v1, params.v2) > params.vmax:
        raise ValueError("VSTART/VSTOP/V1/V2 不能大于 Vmax")

    tvmax_s = max(0.0, params.tvmax_ms) / 1000.0
    accel_at_vmax = _accel_segments(params, params.vmax)
    decel_at_vmax = _decel_segments(params, params.vmax, _sum_time(accel_at_vmax))
    accel_distance = _sum_distance(accel_at_vmax)
    decel_distance = _sum_distance(decel_at_vmax)

    if accel_distance + decel_distance + params.vmax * tvmax_s <= distance.distance:
        vpeak = params.vmax
        reached = True
        profile = "T型"
    else:
        reached = False
        profile = "三角"
        low = max(params.vstart, params.vstop)
        high = params.vmax
        for _ in range(80):
            mid = (low + high) / 2
            candidate = _profile_distance_with_tvmax(params, mid, tvmax_s)
            if candidate > distance.distance:
                high = mid
            else:
                low = mid
        vpeak = low

    accel_segments = _accel_segments(params, vpeak)
    acc_time = _sum_time(accel_segments)
    acc_distance = _sum_distance(accel_segments)
    decel_segments = _decel_segments(params, vpeak, acc_time)
    dec_time = _sum_time(decel_segments)
    dec_distance = _sum_distance(decel_segments)
    const_distance = max(0.0, distance.distance - acc_distance - dec_distance)
    const_time = const_distance / vpeak if vpeak > 0 else 0.0
    segments = list(accel_segments)
    if const_time > 0:
        start = acc_time
        segments.append(Segment("匀速", start, start + const_time, vpeak, vpeak, 0.0, const_distance))
    dec_start = acc_time + const_time
    segments.extend(_decel_segments(params, vpeak, dec_start))
    total = acc_time + const_time + dec_time

    return MotionResult(
        profile_type=profile,
        reached_vmax=reached,
        vpeak=vpeak,
        acc_time=acc_time,
        const_time=const_time,
        dec_time=dec_time,
        total_time=total,
        acc_distance=acc_distance,
        const_distance=const_distance,
        dec_distance=dec_distance,
        segments=segments,
        vmax_register=velocity_to_register(axis, fclk_hz, params.vmax),
        tvmax_register=tvmax_to_register(fclk_hz, params.tvmax_ms),
        amax_register=acceleration_to_register(axis, fclk_hz, params.amax, AMAX_MAX),
        dmax_register=acceleration_to_register(axis, fclk_hz, params.dmax, DMAX_MAX),
        xtarget_register=distance_to_register(axis, distance.distance),
    )


def calculate_forward(axis: Axis, case: MotionCase, fclk_hz: float = 16_000_000.0) -> MotionResult:
    params = MotionParams(
        vmax=case.vmax_mm_s,
        a1=case.acc_mm_s2,
        a2=case.acc_mm_s2,
        amax=case.acc_mm_s2,
        dmax=case.dec_mm_s2,
        d2=case.dec_mm_s2,
        d1=case.dec_mm_s2,
        vstart=case.start_velocity_mm_s,
        vstop=case.end_velocity_mm_s,
        v1=min(case.vmax_mm_s, case.vmax_mm_s / 3),
        v2=min(case.vmax_mm_s, case.vmax_mm_s * 2 / 3),
    )
    return calculate_motion(axis, fclk_hz, params, DistanceCase(distance=case.distance_mm, note=case.note))


def calculate_composite_motion(
    axes: list[Axis],
    actions: list[MotionAction],
    steps: list[CompositeStep],
    fclk_hz: float = 16_000_000.0,
) -> CompositeMotionResult:
    axis_by_id = {axis.id: axis for axis in axes}
    distance_lookup: dict[str, tuple[Axis, MotionAction, DistanceCase]] = {}
    for action in actions:
        axis = axis_by_id.get(action.axis_id)
        if not axis:
            continue
        for distance in action.distances:
            distance_lookup[distance.id] = (axis, action, distance)

    step_results: list[CompositeStepResult] = []
    for step in steps:
        move_results: list[CompositeMoveResult] = []
        for distance_id in (step.action1_distance_id, step.action2_distance_id, step.action3_distance_id):
            if not distance_id:
                continue
            axis, action, distance = distance_lookup[distance_id]
            result = calculate_motion(axis, fclk_hz, action.params, distance)
            move_results.append(
                CompositeMoveResult(
                    axis_name=axis.name,
                    distance_note=distance.note,
                    distance=distance.distance,
                    total_time=result.total_time,
                )
            )
        motion_time = max((item.total_time for item in move_results), default=0.0)
        delay_time = max(0.0, step.delay_ms) / 1000.0
        step_results.append(
            CompositeStepResult(
                step=step,
                move_results=move_results,
                motion_time=motion_time,
                delay_time=delay_time,
                total_time=motion_time + delay_time,
            )
        )

    return CompositeMotionResult(
        steps=step_results,
        total_time=sum(item.total_time for item in step_results),
    )


def _strategy_defaults(strategy: Strategy) -> tuple[float, float, tuple[float, float, float], tuple[float, float, float]]:
    if strategy == "eight_low_impact":
        return 0.30, 0.30, (0.45, 1.00, 0.45), (0.45, 1.00, 0.45)
    if strategy == "eight_stall_guard":
        return 0.35, 0.35, (1.00, 0.65, 0.35), (0.35, 0.65, 1.00)
    if strategy == "low_peak_speed":
        return 0.45, 0.45, (1.00, 1.00, 1.00), (1.00, 1.00, 1.00)
    return 0.25, 0.25, (1.00, 1.00, 1.00), (1.00, 1.00, 1.00)


def _segment_time_shares(delta_fractions: tuple[float, float, float], ratios: tuple[float, float, float]) -> tuple[float, float, float]:
    weighted = tuple(delta / ratio for delta, ratio in zip(delta_fractions, ratios, strict=True))
    total = sum(weighted)
    return tuple(item / total for item in weighted)


def _accel_values(vpeak: float, total_time: float, ratios: tuple[float, float, float]) -> tuple[float, float, float]:
    deltas = (0.30, 0.40, 0.30)
    shares = _segment_time_shares(deltas, ratios)
    return tuple((delta * vpeak) / (share * total_time) for delta, share in zip(deltas, shares, strict=True))


def _ramp_distance_coeff(total_time: float, ratios: tuple[float, float, float], velocity_points: tuple[float, float, float, float]) -> float:
    deltas = tuple(abs(velocity_points[i + 1] - velocity_points[i]) for i in range(3))
    shares = _segment_time_shares(deltas, ratios)
    return sum((velocity_points[i] + velocity_points[i + 1]) * 0.5 * shares[i] * total_time for i in range(3))


def recommend_for_target_time(
    axis: Axis,
    distance_mm: float,
    target_time_s: float,
    strategy: Strategy,
    vmax_limit: float | None = None,
    acc_limit: float | None = None,
    dec_limit: float | None = None,
    fclk_hz: float = 16_000_000.0,
    accel_time_fraction: float | None = None,
    decel_time_fraction: float | None = None,
) -> TargetTimeResult:
    if distance_mm <= 0:
        raise ValueError("运动距离必须大于 0")
    if target_time_s <= 0:
        raise ValueError("目标时间必须大于 0")
    default_accel, default_decel, accel_ratios, decel_ratios = _strategy_defaults(strategy)
    accel_fraction = default_accel if accel_time_fraction is None else accel_time_fraction
    decel_fraction = default_decel if decel_time_fraction is None else decel_time_fraction
    if accel_fraction <= 0 or decel_fraction <= 0 or accel_fraction + decel_fraction >= 0.95:
        raise ValueError("加速/减速时间占比不合理")
    ta = target_time_s * accel_fraction
    td = target_time_s * decel_fraction
    const_time = max(0.0, target_time_s - ta - td)
    v_points = (0.0, 0.30, 0.70, 1.0)
    dec_points = (1.0, 0.70, 0.30, 0.0)
    denominator = _ramp_distance_coeff(ta, accel_ratios, v_points) + const_time + _ramp_distance_coeff(td, decel_ratios, dec_points)
    if denominator <= 0:
        raise ValueError("目标时间过短，无法反算参数")
    vmax = distance_mm / denominator
    if vmax_limit and vmax > vmax_limit:
        raise ValueError("当前 VMAX 上限下无法满足目标时间")
    a1, a2, amax = _accel_values(vmax, ta, accel_ratios)
    dmax, d2, d1 = _accel_values(vmax, td, decel_ratios)
    if acc_limit and max(a1, a2, amax) > acc_limit:
        raise ValueError("当前 A/D 上限下无法满足目标时间")
    if dec_limit and max(dmax, d2, d1) > dec_limit:
        raise ValueError("当前 A/D 上限下无法满足目标时间")
    params = MotionParams(
        vmax=vmax,
        tvmax_ms=0.0,
        a1=a1,
        a2=a2,
        amax=amax,
        dmax=dmax,
        d2=d2,
        d1=d1,
        v1=vmax * 0.30,
        v2=vmax * 0.70,
    )
    result = calculate_motion(axis, fclk_hz, params, DistanceCase(distance=distance_mm))
    tolerance = max(0.05, target_time_s * 0.05)
    if abs(result.total_time - target_time_s) > tolerance:
        raise ValueError("当前上限条件下无法接近目标时间")
    return TargetTimeResult(
        strategy=strategy,
        vmax_mm_s=vmax,
        acc_mm_s2=max(a1, a2, amax),
        dec_mm_s2=max(dmax, d2, d1),
        forward=result,
        params=params,
        time_error_s=result.total_time - target_time_s,
    )


def calculate_t_profile(axis: Axis, case: TProfileCase, fclk_hz: float = 16_000_000.0) -> TProfileResult:
    derived = derive_axis(axis)
    if case.acc_time_s <= 0 or case.dec_time_s <= 0:
        raise ValueError("加速时间和减速时间必须大于 0")
    vstart_rps = case.vstart_register * fclk_hz / 2**24 / derived.count
    vstop_rps = case.vstop_register * fclk_hz / 2**24 / derived.count
    s_acc_rev = case.acc_distance_microsteps / derived.count
    s_dec_rev = case.dec_distance_microsteps / derived.count
    s_const_rev = case.const_distance_microsteps / derived.count
    acc_rps2 = 2 * (s_acc_rev - vstart_rps * case.acc_time_s) / case.acc_time_s**2
    vmax_rps = vstart_rps + acc_rps2 * case.acc_time_s
    dec_rps2 = (vmax_rps - vstop_rps) / case.dec_time_s
    const_time = s_const_rev / vmax_rps if vmax_rps > 0 else 0
    return TProfileResult(
        acc_rps2=acc_rps2,
        vmax_rps=vmax_rps,
        dec_rps2=dec_rps2,
        const_time_s=const_time,
        total_time_s=case.acc_time_s + const_time + case.dec_time_s,
        amax_register=_reg(acc_rps2 * derived.count * 2**41 / fclk_hz**2, (0, AMAX_MAX)),
        vmax_register=_reg(vmax_rps * derived.count * 2**24 / fclk_hz, (0, VMAX_MAX)),
        dmax_register=_reg(dec_rps2 * derived.count * 2**41 / fclk_hz**2, (0, DMAX_MAX)),
        xtarget_register=_reg(case.acc_distance_microsteps + case.const_distance_microsteps + case.dec_distance_microsteps, (XTARGET_MIN, XTARGET_MAX)),
    )
