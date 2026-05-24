from __future__ import annotations

import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

from .calculations import (
    MotionResult,
    acceleration_to_motor_rpm_s,
    acceleration_to_register,
    calculate_composite_motion,
    calculate_motion,
    derive_axis,
    distance_to_register,
    motor_rpm_s_to_acceleration,
    motor_rpm_to_velocity,
    register_to_motor_rpm,
    register_to_motor_rpm_s,
    register_to_acceleration,
    register_to_distance,
    register_to_tvmax_ms,
    register_to_velocity,
    tvmax_to_register,
    velocity_to_motor_rpm,
    recommend_for_target_time,
    velocity_to_register,
)
from .models import AppState, Axis, CompositeAction, CompositeStep, DistanceCase, MotionAction, MotionParams, new_id
from .io_excel import export_excel
from .project_io import load_project, save_project


VELOCITY_FIELDS = {"vmax", "vstart", "vstop", "v1", "v2"}
ACCEL_FIELDS = {"a1", "a2", "amax", "dmax", "d2", "d1"}
TIME_FIELDS = {"tvmax_ms"}
MECHANISM_LABELS = {
    "步进直线": "linear_belt",
    "步进旋转": "rotary",
    "步进丝杆": "leadscrew",
}
MECHANISM_NAMES = {value: key for key, value in MECHANISM_LABELS.items()}
FIELD_LABELS = {
    "vmax": "VMAX",
    "tvmax_ms": "TVMAX",
    "vstart": "VSTART",
    "vstop": "VSTOP",
    "v1": "V1",
    "v2": "V2",
    "a1": "A1",
    "a2": "A2",
    "amax": "AMAX",
    "dmax": "DMAX",
    "d2": "D2",
    "d1": "D1",
}
STRATEGY_LABELS = {
    "平衡策略": "balanced",
    "低冲击策略": "low_impact",
    "低峰值速度策略": "low_peak_speed",
    "对称加减速策略": "symmetric",
}
STRATEGY_LABELS = {
    "平衡（梯形）": "balanced",
    "8点低冲击": "eight_low_impact",
    "8点防失步": "eight_stall_guard",
    "低峰值（三角）": "low_peak_speed",
}

STRATEGY_DEFAULT_RATIOS = {
    "balanced": (25, 25),
    "eight_low_impact": (30, 30),
    "eight_stall_guard": (35, 35),
    "low_peak_speed": (45, 45),
}


def _float(value: str, default: float = 0.0) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def _int(value: str, default: int = 0) -> int:
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return default


def _fmt(value: float, digits: int = 3) -> str:
    return f"{value:.{digits}f}"


class ProjectDialog(simpledialog.Dialog):
    def body(self, master: tk.Widget) -> tk.Widget:
        self.title("新建项目")
        ttk.Label(master, text="项目名称").grid(row=0, column=0, sticky="w", padx=6, pady=6)
        self.name_var = tk.StringVar(value="Motion Project")
        ttk.Entry(master, textvariable=self.name_var, width=28).grid(row=0, column=1, padx=6, pady=6)
        ttk.Label(master, text="主控器主频").grid(row=1, column=0, sticky="w", padx=6, pady=6)
        self.fclk_var = tk.StringVar(value="16 MHz")
        ttk.Combobox(master, textvariable=self.fclk_var, values=["12.5 MHz", "16 MHz"], state="readonly", width=25).grid(row=1, column=1, padx=6, pady=6)
        return master

    def apply(self) -> None:
        self.result = {
            "name": self.name_var.get().strip() or "Motion Project",
            "fclk_hz": 12_500_000.0 if self.fclk_var.get().startswith("12.5") else 16_000_000.0,
        }


class AxisDialog(simpledialog.Dialog):
    def __init__(self, parent: tk.Widget, axis: Axis | None = None) -> None:
        self.axis = axis
        super().__init__(parent, "运动轴")

    def body(self, master: tk.Widget) -> tk.Widget:
        axis = self.axis or Axis(name=f"Axis {new_id()[:3]}")
        self.vars: dict[str, tk.StringVar] = {}
        self.rows: dict[str, tuple[ttk.Label, tk.Widget]] = {}
        fields = [
            ("name", "名称", axis.name),
            ("mechanism_type", "结构类型", MECHANISM_NAMES.get(axis.normalized_type(), "步进直线")),
            ("motor_step_angle", "步距角", axis.motor_step_angle),
            ("microstep", "微步", axis.microstep),
            ("gear_ratio", "减速比", axis.gear_ratio),
            ("pulley_teeth", "同步轮齿数", axis.pulley_teeth),
            ("pulley_pitch_mm", "同步轮齿距 mm", axis.pulley_pitch_mm),
            ("lead_mm_per_rev", "丝杆导程 mm", axis.lead_mm_per_rev),
        ]
        for row, (key, label, value) in enumerate(fields):
            label_widget = ttk.Label(master, text=label)
            label_widget.grid(row=row, column=0, sticky="w", padx=6, pady=4)
            var = tk.StringVar(value=str(value))
            self.vars[key] = var
            if key == "mechanism_type":
                widget = ttk.Combobox(master, textvariable=var, values=list(MECHANISM_LABELS), state="readonly", width=24)
                widget.bind("<<ComboboxSelected>>", lambda _e: self.update_visible_fields())
            elif key == "motor_step_angle":
                widget = ttk.Combobox(master, textvariable=var, values=["1.8", "0.9"], width=24)
            elif key == "microstep":
                widget = ttk.Combobox(master, textvariable=var, values=["256", "128", "64", "32", "16", "8", "4", "2", "1"], state="readonly", width=24)
            else:
                widget = ttk.Entry(master, textvariable=var, width=26)
            widget.grid(row=row, column=1, sticky="ew", padx=6, pady=4)
            self.rows[key] = (label_widget, widget)
        self.help_label = ttk.Label(master, text="", foreground="#666")
        self.help_label.grid(row=len(fields), column=0, columnspan=2, sticky="w", padx=6, pady=4)
        self.update_visible_fields()
        return master

    def update_visible_fields(self) -> None:
        mechanism = MECHANISM_LABELS.get(self.vars["mechanism_type"].get(), "linear_belt")
        visible = {
            "linear_belt": {"name", "mechanism_type", "motor_step_angle", "microstep", "gear_ratio", "pulley_teeth", "pulley_pitch_mm"},
            "rotary": {"name", "mechanism_type", "motor_step_angle", "microstep", "gear_ratio"},
            "leadscrew": {"name", "mechanism_type", "motor_step_angle", "microstep", "lead_mm_per_rev"},
        }[mechanism]
        help_text = {
            "linear_belt": "步进直线：步距角、微步、减速比、同步轮齿数、同步轮齿距。",
            "rotary": "步进旋转：步距角、减速比、微步。",
            "leadscrew": "步进丝杆：步距角、丝杆导程、微步。",
        }[mechanism]
        for key, (label, widget) in self.rows.items():
            if key in visible:
                label.grid()
                widget.grid()
            else:
                label.grid_remove()
                widget.grid_remove()
        self.help_label.configure(text=help_text)

    def apply(self) -> None:
        self.result = Axis(
            id=self.axis.id if self.axis else new_id(),
            name=self.vars["name"].get().strip() or "Axis",
            mechanism_type=MECHANISM_LABELS.get(self.vars["mechanism_type"].get(), "linear_belt"),
            motor_step_angle=_float(self.vars["motor_step_angle"].get(), 1.8),
            microstep=_int(self.vars["microstep"].get(), 256),
            gear_ratio=_float(self.vars["gear_ratio"].get(), 1.0),
            pulley_teeth=_int(self.vars["pulley_teeth"].get(), 20),
            pulley_pitch_mm=_float(self.vars["pulley_pitch_mm"].get(), 2.0),
            lead_mm_per_rev=_float(self.vars["lead_mm_per_rev"].get(), 6.35),
        )


class MotionCalculatorApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("HM Motion Calculator")
        self.geometry("1420x820")
        self.minsize(980, 620)
        self.state_data = AppState()
        self.current_axis_id = ""
        self.current_action_id = ""
        self.current_distance_id = ""
        self.project_path = ""
        self.param_phys_vars: dict[str, tk.StringVar] = {}
        self.param_motor_vars: dict[str, tk.StringVar] = {}
        self.param_reg_vars: dict[str, tk.StringVar] = {}
        self.param_reg_hint_vars: dict[str, tk.StringVar] = {}
        self.param_reg_entries: dict[str, ttk.Entry] = {}
        self.param_reg_hint_labels: dict[str, tk.Label] = {}
        self.target_vars: dict[str, tk.StringVar] = {}
        self.target_result_vars: dict[str, tk.StringVar] = {}
        self.target_apply_button: ttk.Button | None = None
        self._loading_action_params = False
        self.current_composite_action_id = ""
        self.composite_number_var = tk.StringVar(value="")
        self.composite_name_var = tk.StringVar(value="")
        self._refreshing_composite_action_list = False
        self.composite_editor: ttk.Frame | None = None
        self.composite_total_var = tk.StringVar(value="")
        self.composite_status_var = tk.StringVar(value="")
        self.composite_vars: dict[str, dict[str, tk.StringVar]] = {}
        self._composite_option_by_id: dict[str, str] = {}
        self._composite_id_by_option: dict[str, str] = {}
        self._composite_total_by_id: dict[str, float | None] = {}
        self._last_target_params: MotionParams | None = None
        self._last_target_result: MotionResult | None = None
        self._target_has_calculated = False
        self.detail_vars: dict[str, tk.StringVar] = {}
        self.detail_labels: dict[str, ttk.Label] = {}
        self._chart_result: MotionResult | None = None
        self._chart_w = 400
        self._chart_h = 200
        self._chart_pad = 46
        self._chart_total = 1.0
        self._chart_vmax = 1.0
        self._chart_unit = "mm"
        self._marker_x = 46.0
        self._show_start_page()

    def _clear(self) -> None:
        for child in self.winfo_children():
            child.destroy()

    def _show_start_page(self) -> None:
        self._clear()
        style = ttk.Style(self)
        style.configure("Start.TButton", font=("Segoe UI", 15), padding=(24, 10))
        frame = ttk.Frame(self, padding=60)
        frame.pack(expand=True)
        ttk.Label(frame, text="HM Motion Calculator", font=("Segoe UI", 28, "bold"), justify="center").pack(pady=16)
        ttk.Label(frame, text="请选择项目。没有现有项目时，请先新建项目。", font=("Segoe UI", 14)).pack(pady=12)
        ttk.Button(frame, text="新建项目", style="Start.TButton", command=self.new_project).pack(fill="x", pady=8)
        ttk.Button(frame, text="打开项目", style="Start.TButton", command=self.open_project).pack(fill="x", pady=8)

    def _build_main(self) -> None:
        self._clear()
        self._configure_styles()
        self._build_menu()
        self.columnconfigure(0, weight=0)
        self.columnconfigure(1, weight=1)
        self.columnconfigure(2, weight=0, minsize=420)
        self.rowconfigure(0, weight=2, uniform="main")
        self.rowconfigure(1, weight=1, uniform="main")
        self._build_axis_panel()
        self._build_center()
        self._build_detail_panel()
        self._build_chart()
        self.refresh_all()

    def _configure_styles(self) -> None:
        style = ttk.Style(self)
        style.configure(".", font=("Segoe UI", 11))
        style.configure("TLabel", font=("Segoe UI", 11))
        style.configure("TButton", font=("Segoe UI", 11))
        style.configure("TEntry", font=("Segoe UI", 11))
        style.configure("TCombobox", font=("Segoe UI", 11))
        style.configure("TNotebook.Tab", font=("Segoe UI", 13, "bold"), padding=(18, 12))
        style.configure("Treeview", font=("Segoe UI", 11), rowheight=32)
        style.configure("Treeview.Heading", font=("Segoe UI", 11, "bold"))

    def _create_scrollable_tab(self) -> tuple[ttk.Frame, ttk.Frame]:
        container = ttk.Frame(self.notebook)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(0, weight=1)
        canvas = tk.Canvas(container, highlightthickness=0)
        y_scroll = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        x_scroll = ttk.Scrollbar(container, orient="horizontal", command=canvas.xview)
        canvas.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")

        inner = ttk.Frame(canvas, padding=8)
        window_id = canvas.create_window((0, 0), window=inner, anchor="nw")

        def update_scrollregion(_event: tk.Event | None = None) -> None:
            canvas.configure(scrollregion=canvas.bbox("all"))

        inner.bind("<Configure>", update_scrollregion)
        return container, inner

    def _build_menu(self) -> None:
        menu = tk.Menu(self)
        menu.add_command(label="新建项目", command=self.new_project)
        menu.add_command(label="打开项目", command=self.open_project)
        menu.add_command(label="保存项目", command=self.save_project)
        menu.add_command(label="导出 Excel", command=self.export_project_excel)
        self.config(menu=menu)

    def _build_axis_panel(self) -> None:
        panel = ttk.Frame(self, padding=10)
        panel.grid(row=0, column=0, sticky="nsw")
        ttk.Label(panel, text="运动轴", font=("Segoe UI", 14, "bold")).pack(anchor="w")
        self.axis_list = tk.Listbox(panel, width=28, height=8, font=("Segoe UI", 13), activestyle="dotbox")
        self.axis_list.pack(fill="both", expand=True, pady=8)
        self.axis_list.bind("<<ListboxSelect>>", lambda _e: self.on_axis_select())
        buttons = ttk.Frame(panel)
        buttons.pack(fill="x")
        ttk.Button(buttons, text="新增", command=self.add_axis).pack(side="left", padx=2)
        ttk.Button(buttons, text="编辑", command=self.edit_axis).pack(side="left", padx=2)
        ttk.Button(buttons, text="删除", command=self.delete_axis).pack(side="left", padx=2)
        ttk.Separator(panel).pack(fill="x", pady=8)
        self.axis_info = tk.StringVar(value="")
        ttk.Label(panel, textvariable=self.axis_info, justify="left", wraplength=250, font=("Segoe UI", 10)).pack(anchor="w")

    def _build_center(self) -> None:
        center = ttk.Frame(self, padding=8)
        center.grid(row=0, column=1, sticky="nsew")
        center.columnconfigure(0, weight=1)
        center.rowconfigure(1, weight=1)
        self.center_axis_var = tk.StringVar(value="")
        ttk.Label(center, textvariable=self.center_axis_var, font=("Segoe UI", 12, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 4))
        self.notebook = ttk.Notebook(center)
        self.notebook.grid(row=1, column=0, sticky="nsew")
        self.action_tab = ttk.Frame(self.notebook, padding=8)
        self.composite_tab = ttk.Frame(self.notebook, padding=8)
        self.target_tab = ttk.Frame(self.notebook, padding=8)
        self.notebook.add(self.action_tab, text="单轴时间计算")
        self.notebook.add(self.composite_tab, text="组合动作时间计算")
        self.notebook.add(self.target_tab, text="运动速度反算")
        self._build_action_tab()
        self._build_composite_tab()
        self._build_target_tab()

    def _build_action_tab(self) -> None:
        self.action_tab.columnconfigure(0, weight=1)
        self.action_tab.rowconfigure(0, weight=1)
        self.action_tab.rowconfigure(3, weight=1)
        right = ttk.Frame(self.action_tab)
        right.grid(row=0, column=0, sticky="nsew")
        right.columnconfigure(0, weight=1)
        right.rowconfigure(3, weight=1)
        ttk.Label(right, text="运动参数", font=("Segoe UI", 13, "bold")).grid(row=0, column=0, sticky="w")
        grid = ttk.Frame(right)
        grid.grid(row=1, column=0, sticky="ew", pady=6)
        ttk.Label(grid, text="参数").grid(row=0, column=0, padx=2)
        ttk.Label(grid, text="寄存器").grid(row=1, column=0, padx=2, sticky="e")
        ttk.Label(grid, text="电机端").grid(row=2, column=0, padx=2, sticky="e")
        ttk.Label(grid, text="物理端").grid(row=3, column=0, padx=2, sticky="e")
        velocity_count = list(FIELD_LABELS).index("a1")
        vel_unit_col = 1
        sep_col = vel_unit_col + velocity_count + 1   # col 7
        acc_unit_col = sep_col + 1                    # col 8

        # 速度组左侧单位列
        ttk.Label(grid, text="RPM",   foreground="#555").grid(row=2, column=vel_unit_col, padx=4, sticky="e")
        ttk.Label(grid, text="mm/s",  foreground="#555").grid(row=3, column=vel_unit_col, padx=4, sticky="e")

        # 速度 / 加速度分隔线
        ttk.Separator(grid, orient="vertical").grid(row=0, column=sep_col, rowspan=4, sticky="ns", padx=8)

        # 加速度组左侧单位列
        ttk.Label(grid, text="RPM/s",  foreground="#555").grid(row=2, column=acc_unit_col, padx=4, sticky="e")
        ttk.Label(grid, text="mm/s²", foreground="#555").grid(row=3, column=acc_unit_col, padx=4, sticky="e")

        for col, key in enumerate(FIELD_LABELS):
            if col < velocity_count:
                ui_col = vel_unit_col + 1 + col           # cols 2-6
            else:
                ui_col = acc_unit_col + 1 + (col - velocity_count)  # cols 9-14
            ttk.Label(grid, text=FIELD_LABELS[key]).grid(row=0, column=ui_col, padx=2)
            reg_var = tk.StringVar(value="0")
            motor_var = tk.StringVar(value="0")
            phys_var = tk.StringVar(value="0")
            hint_var = tk.StringVar(value="")
            self.param_reg_vars[key] = reg_var
            self.param_motor_vars[key] = motor_var
            self.param_phys_vars[key] = phys_var
            self.param_reg_hint_vars[key] = hint_var
            reg = ttk.Entry(grid, textvariable=reg_var, width=10)
            reg.grid(row=1, column=ui_col, padx=2, pady=2)
            hint = tk.Label(grid, textvariable=hint_var, foreground="#999", background="white", font=("Segoe UI", 9))
            hint.place(in_=reg, relx=1.0, x=-5, rely=0.5, anchor="e")
            hint.place_forget()
            self.param_reg_entries[key] = reg
            self.param_reg_hint_labels[key] = hint
            motor = ttk.Entry(grid, textvariable=motor_var, width=10)
            phys = ttk.Entry(grid, textvariable=phys_var, width=10)
            motor.grid(row=2, column=ui_col, padx=2, pady=2)
            phys.grid(row=3, column=ui_col, padx=2, pady=2)
            reg.bind("<FocusIn>", lambda _e, field=key: self._update_reg_hint(field))
            reg.bind("<FocusOut>", lambda _e, field=key: self._on_reg_focus_out(field))
            reg.bind("<KeyRelease>", lambda _e, field=key: self._update_reg_hint(field))
            reg.bind("<Return>", lambda _e, field=key: self._on_reg_enter(field))
            motor.bind("<FocusOut>", lambda _e, field=key: self.sync_param_from_motor(field))
            phys.bind("<FocusOut>", lambda _e, field=key: self.sync_param_from_physical(field))
        ttk.Button(right, text="重新计算", command=self.recalculate_all).grid(row=2, column=0, sticky="w", pady=4)

        distance_box = ttk.LabelFrame(right, text="距离列表")
        distance_box.grid(row=3, column=0, sticky="nsew", pady=6)
        distance_box.columnconfigure(0, weight=1)
        distance_box.rowconfigure(1, weight=1)
        controls = ttk.Frame(distance_box)
        controls.grid(row=0, column=0, sticky="ew", pady=4)
        self.distance_mm_var = tk.StringVar(value="")
        self.distance_reg_var = tk.StringVar(value="")
        self.distance_note_var = tk.StringVar(value="")
        ttk.Label(controls, text="mm").pack(side="left")
        mm_entry = ttk.Entry(controls, textvariable=self.distance_mm_var, width=10)
        mm_entry.pack(side="left", padx=3)
        mm_entry.bind("<FocusOut>", lambda _e: self._auto_sync_distance_from_mm())
        ttk.Label(controls, text="X_TARGET").pack(side="left")
        reg_entry = ttk.Entry(controls, textvariable=self.distance_reg_var, width=12)
        reg_entry.pack(side="left", padx=3)
        reg_entry.bind("<FocusOut>", lambda _e: self._auto_sync_distance_from_reg())
        ttk.Label(controls, text="备注").pack(side="left")
        ttk.Entry(controls, textvariable=self.distance_note_var, width=18).pack(side="left", padx=3)
        ttk.Button(controls, text="添加距离", command=self.add_distance).pack(side="left", padx=5)
        ttk.Button(controls, text="删除距离", command=self.delete_distance).pack(side="left", padx=5)

        columns = ("distance", "xtarget", "total", "note")
        self.distance_tree = ttk.Treeview(distance_box, columns=columns, show="headings", height=8)
        headers = {"distance": "距离 mm", "xtarget": "X_TARGET", "total": "运动时间 s", "note": "备注"}
        for col in columns:
            self.distance_tree.heading(col, text=headers[col])
            width = 180 if col == "note" else 120
            self.distance_tree.column(col, width=width, anchor="center")
        self.distance_tree.grid(row=1, column=0, sticky="nsew")
        self.distance_tree.bind("<<TreeviewSelect>>", lambda _e: self.on_distance_select())

    def _build_composite_tab(self) -> None:
        self.composite_tab.columnconfigure(0, weight=0, minsize=230)
        self.composite_tab.columnconfigure(1, weight=1)
        self.composite_tab.rowconfigure(0, weight=1)

        left = ttk.Frame(self.composite_tab)
        left.grid(row=0, column=0, sticky="nsw", padx=(0, 10))
        ttk.Label(left, text="组合动作", font=("Segoe UI", 13, "bold")).pack(anchor="w")
        self.composite_action_list = ttk.Treeview(
            left,
            columns=("number", "name", "total"),
            show="headings",
            height=12,
        )
        for col, text, width in [
            ("number", "编号", 56),
            ("name", "名称", 120),
            ("total", "组合总时间", 120),
        ]:
            self.composite_action_list.heading(col, text=text)
            self.composite_action_list.column(col, width=width, anchor="center", stretch=(col == "name"))
        self.composite_action_list.pack(fill="both", expand=True, pady=6)
        self.composite_action_list.bind("<<TreeviewSelect>>", lambda _e: self.on_composite_action_select())
        action_buttons = ttk.Frame(left)
        action_buttons.pack(fill="x")
        ttk.Button(action_buttons, text="新增", command=self.add_composite_action).pack(side="left", padx=(0, 4))
        ttk.Button(action_buttons, text="删除", command=self.delete_composite_action).pack(side="left", padx=4)
        ttk.Button(action_buttons, text="上移", command=lambda: self.move_current_composite_action(-1)).pack(side="left", padx=4)
        ttk.Button(action_buttons, text="下移", command=lambda: self.move_current_composite_action(1)).pack(side="left", padx=4)

        right = ttk.Frame(self.composite_tab)
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)
        right.rowconfigure(2, weight=1)

        name_bar = ttk.Frame(right)
        name_bar.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        name_bar.columnconfigure(3, weight=1)
        ttk.Label(name_bar, text="编号").grid(row=0, column=0, sticky="w", padx=(0, 6))
        number_entry = ttk.Entry(name_bar, textvariable=self.composite_number_var, width=12)
        number_entry.grid(row=0, column=1, sticky="w", padx=(0, 12))
        number_entry.bind("<FocusOut>", lambda _e: self.rename_current_composite_action())
        number_entry.bind("<Return>", lambda _e: self.rename_current_composite_action())
        ttk.Label(name_bar, text="名称").grid(row=0, column=2, sticky="w", padx=(0, 6))
        name_entry = ttk.Entry(name_bar, textvariable=self.composite_name_var, width=32)
        name_entry.grid(row=0, column=3, sticky="ew")
        name_entry.bind("<FocusOut>", lambda _e: self.rename_current_composite_action())
        name_entry.bind("<Return>", lambda _e: self.rename_current_composite_action())

        top = ttk.Frame(right)
        top.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        ttk.Button(top, text="增加 STEP", command=self.add_composite_step).pack(side="left", padx=(0, 5))
        ttk.Button(top, text="删除选中 STEP", command=self.delete_composite_step).pack(side="left", padx=5)
        ttk.Label(top, textvariable=self.composite_total_var, font=("Segoe UI", 12, "bold")).pack(side="left", padx=16)
        ttk.Label(top, textvariable=self.composite_status_var, foreground="#777").pack(side="left", padx=8)

        body = ttk.Frame(right)
        body.grid(row=2, column=0, sticky="nsew")
        body.columnconfigure(0, weight=1)
        body.rowconfigure(0, weight=1)
        body.rowconfigure(1, weight=1)

        editor_box = ttk.LabelFrame(body, text="组合动作编辑：每一行串行执行，同一行内动作1/动作2/动作3并行执行")
        editor_box.grid(row=0, column=0, sticky="nsew")
        editor_box.columnconfigure(0, weight=1)
        editor_box.rowconfigure(0, weight=1)
        editor_canvas = tk.Canvas(editor_box, height=190, highlightthickness=0)
        editor_y = ttk.Scrollbar(editor_box, orient="vertical", command=editor_canvas.yview)
        editor_x = ttk.Scrollbar(editor_box, orient="horizontal", command=editor_canvas.xview)
        editor_canvas.configure(yscrollcommand=editor_y.set, xscrollcommand=editor_x.set)
        editor_canvas.grid(row=0, column=0, sticky="nsew")
        editor_y.grid(row=0, column=1, sticky="ns")
        editor_x.grid(row=1, column=0, sticky="ew")
        self.composite_editor = ttk.Frame(editor_canvas)
        editor_canvas.create_window((0, 0), window=self.composite_editor, anchor="nw")
        self.composite_editor.bind("<Configure>", lambda _e: editor_canvas.configure(scrollregion=editor_canvas.bbox("all")))
        editor_canvas.bind("<MouseWheel>", lambda e: editor_canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"))

        columns = ("step", "note", "motion", "delay", "total")
        self.composite_result_tree = ttk.Treeview(body, columns=columns, show="headings", height=10)
        headers = {"step": "STEP", "note": "备注", "motion": "并行动作时间 s", "delay": "延时 s", "total": "STEP 总时间 s"}
        for col in columns:
            self.composite_result_tree.heading(col, text=headers[col])
            width = 220 if col == "motion" else 160
            self.composite_result_tree.column(col, width=width, anchor="center")
        self.composite_result_tree.grid(row=1, column=0, sticky="nsew", pady=(8, 0))
        self.composite_result_tree.bind("<Double-1>", self.edit_composite_step_note)

        self.refresh_composite_tab()

    def ensure_current_composite_action(self) -> CompositeAction | None:
        if not self.state_data.composite_actions:
            self.state_data.composite_actions.append(CompositeAction(name="组合动作1"))
        if not self.current_composite_action_id:
            self.current_composite_action_id = self.state_data.composite_actions[0].id
        current = self.get_current_composite_action()
        if current:
            return current
        self.current_composite_action_id = self.state_data.composite_actions[0].id
        return self.state_data.composite_actions[0]

    def get_current_composite_action(self) -> CompositeAction | None:
        return next((item for item in self.state_data.composite_actions if item.id == self.current_composite_action_id), None)

    def refresh_composite_action_list(self) -> None:
        if not hasattr(self, "composite_action_list"):
            return
        current = self.ensure_current_composite_action()
        self._refreshing_composite_action_list = True
        for item in self.composite_action_list.get_children():
            self.composite_action_list.delete(item)
        self._composite_total_by_id = {}
        for idx, item in enumerate(self.state_data.composite_actions, start=1):
            try:
                total = calculate_composite_motion(self.state_data.axes, self.state_data.actions, item.steps, self.state_data.fclk_hz).total_time
                total_text = _fmt(total)
                self._composite_total_by_id[item.id] = total
            except Exception:
                total_text = ""
                self._composite_total_by_id[item.id] = None
            number = item.number.strip() or str(idx)
            self.composite_action_list.insert("", tk.END, iid=item.id, values=(number, item.name, total_text))
        if current:
            if current.id in self.composite_action_list.get_children():
                self.composite_action_list.selection_set(current.id)
                self.composite_action_list.see(current.id)
            self.composite_number_var.set(current.number.strip() or str(next((idx for idx, item in enumerate(self.state_data.composite_actions, start=1) if item.id == current.id), "")))
            self.composite_name_var.set(current.name)
        self._refreshing_composite_action_list = False

    def on_composite_action_select(self) -> None:
        if self._refreshing_composite_action_list:
            return
        selection = self.composite_action_list.selection()
        if not selection:
            return
        if selection[0] == self.current_composite_action_id:
            return
        self.save_all_composite_steps()
        self.current_composite_action_id = selection[0]
        current = self.get_current_composite_action()
        self.composite_number_var.set(current.number if current else "")
        self.composite_name_var.set(current.name if current else "")
        self.refresh_composite_tab()

    def add_composite_action(self) -> None:
        self.save_all_composite_steps()
        next_number = str(len(self.state_data.composite_actions) + 1)
        action = CompositeAction(number=next_number, name=f"组合动作{next_number}")
        self.state_data.composite_actions.append(action)
        self.current_composite_action_id = action.id
        self.refresh_composite_tab()

    def delete_composite_action(self) -> None:
        current = self.get_current_composite_action()
        if not current:
            return
        self.state_data.composite_actions = [item for item in self.state_data.composite_actions if item.id != current.id]
        self.current_composite_action_id = self.state_data.composite_actions[0].id if self.state_data.composite_actions else ""
        self.refresh_composite_tab()

    def rename_current_composite_action(self) -> str:
        current = self.get_current_composite_action()
        if current:
            current.number = self.composite_number_var.get().strip()
            current.name = self.composite_name_var.get().strip() or current.name
            self.refresh_composite_action_list()
        return "break"

    def move_current_composite_action(self, direction: int) -> None:
        current = self.get_current_composite_action()
        if not current:
            return
        self.save_all_composite_steps()
        index = self.state_data.composite_actions.index(current)
        new_index = index + direction
        if new_index < 0 or new_index >= len(self.state_data.composite_actions):
            return
        self.state_data.composite_actions[index], self.state_data.composite_actions[new_index] = (
            self.state_data.composite_actions[new_index],
            self.state_data.composite_actions[index],
        )
        self.current_composite_action_id = current.id
        self.refresh_composite_tab()

    def _composite_distance_options(self) -> list[str]:
        options = [""]
        for action in self.state_data.actions:
            axis = self.get_axis(action.axis_id)
            if not axis:
                continue
            for idx, distance in enumerate(action.distances, start=1):
                note = distance.note.strip() or f"动作{idx}"
                label = f"{axis.name} / {note} / {_fmt(distance.distance)}"
                self._composite_option_by_id[distance.id] = label
                self._composite_id_by_option[label] = distance.id
                options.append(label)
        return options

    def refresh_composite_tab(self) -> None:
        if not self.composite_editor:
            return
        current_action = self.ensure_current_composite_action()
        self.refresh_composite_action_list()
        for child in self.composite_editor.winfo_children():
            child.destroy()
        self.composite_vars = {}
        self._composite_option_by_id = {}
        self._composite_id_by_option = {}
        distance_options = self._composite_distance_options()

        headers = ["STEP", "备注", "延时 ms", "动作1", "动作2", "动作3"]
        for col, text in enumerate(headers):
            ttk.Label(self.composite_editor, text=text, font=("Segoe UI", 11, "bold")).grid(row=0, column=col, sticky="w", padx=4, pady=4)

        steps = current_action.steps if current_action else []
        for row, step in enumerate(steps, start=1):
            for key in ("action1_distance_id", "action2_distance_id", "action3_distance_id"):
                if getattr(step, key) and getattr(step, key) not in self._composite_option_by_id:
                    setattr(step, key, "")
            ttk.Label(self.composite_editor, text=step.name).grid(row=row, column=0, sticky="w", padx=4, pady=3)
            vars_for_step = {
                "note": tk.StringVar(value=step.note),
                "delay_ms": tk.StringVar(value=_fmt(step.delay_ms)),
                "action1_distance_id": tk.StringVar(value=self._composite_option_by_id.get(step.action1_distance_id, "")),
                "action2_distance_id": tk.StringVar(value=self._composite_option_by_id.get(step.action2_distance_id, "")),
                "action3_distance_id": tk.StringVar(value=self._composite_option_by_id.get(step.action3_distance_id, "")),
            }
            self.composite_vars[step.id] = vars_for_step
            note = ttk.Entry(self.composite_editor, textvariable=vars_for_step["note"], width=18)
            note.grid(row=row, column=1, sticky="ew", padx=4, pady=3)
            note.bind("<FocusOut>", lambda _e, step_id=step.id: self.save_composite_step(step_id))
            note.bind("<Return>", lambda _e, step_id=step.id: self.save_composite_step(step_id))
            delay = ttk.Entry(self.composite_editor, textvariable=vars_for_step["delay_ms"], width=10)
            delay.grid(row=row, column=2, sticky="ew", padx=4, pady=3)
            delay.bind("<FocusOut>", lambda _e, step_id=step.id: self.save_composite_step(step_id))
            for col, key, options in [
                (3, "action1_distance_id", distance_options),
                (4, "action2_distance_id", distance_options),
                (5, "action3_distance_id", distance_options),
            ]:
                combo = ttk.Combobox(self.composite_editor, textvariable=vars_for_step[key], values=options, state="readonly", width=28)
                combo.grid(row=row, column=col, sticky="ew", padx=4, pady=3)
                combo.bind("<<ComboboxSelected>>", lambda _e, step_id=step.id: self.save_composite_step(step_id))
        self.composite_editor.columnconfigure(1, weight=1)
        for col in range(3, 6):
            self.composite_editor.columnconfigure(col, weight=1)
        self.calculate_composite_time()

    def save_composite_step(self, step_id: str) -> None:
        current_action = self.get_current_composite_action()
        step = next((item for item in current_action.steps if item.id == step_id), None) if current_action else None
        vars_for_step = self.composite_vars.get(step_id)
        if not step or not vars_for_step:
            return
        step.note = vars_for_step["note"].get().strip()
        step.delay_ms = _float(vars_for_step["delay_ms"].get(), 0)
        step.action1_distance_id = self._composite_id_by_option.get(vars_for_step["action1_distance_id"].get(), "")
        step.action2_distance_id = self._composite_id_by_option.get(vars_for_step["action2_distance_id"].get(), "")
        step.action3_distance_id = self._composite_id_by_option.get(vars_for_step["action3_distance_id"].get(), "")
        self.calculate_composite_time()

    def save_all_composite_steps(self) -> None:
        self.rename_current_composite_action()
        for step_id in list(self.composite_vars):
            self.save_composite_step(step_id)

    def add_composite_step(self) -> None:
        self.save_all_composite_steps()
        current_action = self.ensure_current_composite_action()
        if not current_action:
            return
        current_action.steps.append(CompositeStep(name=f"STEP{len(current_action.steps) + 1}"))
        self.refresh_composite_tab()

    def delete_composite_step(self) -> None:
        selected = self.composite_result_tree.selection()
        if not selected:
            return
        selected_id = selected[0]
        current_action = self.get_current_composite_action()
        if not current_action:
            return
        current_action.steps = [step for step in current_action.steps if step.id != selected_id]
        for idx, step in enumerate(current_action.steps, start=1):
            step.name = f"STEP{idx}"
        self.refresh_composite_tab()

    def edit_composite_step_note(self, event: tk.Event) -> None:
        step_id = self.composite_result_tree.identify_row(event.y)
        if not step_id:
            return
        current_action = self.get_current_composite_action()
        step = next((item for item in current_action.steps if item.id == step_id), None) if current_action else None
        if not step:
            return
        value = simpledialog.askstring("编辑 STEP 备注", f"{step.name} 备注", initialvalue=step.note, parent=self)
        if value is None:
            return
        step.note = value.strip()
        self.refresh_composite_tab()

    def calculate_composite_time(self) -> None:
        if not hasattr(self, "composite_result_tree"):
            return
        for item in self.composite_result_tree.get_children():
            self.composite_result_tree.delete(item)
        try:
            current_action = self.get_current_composite_action()
            steps = current_action.steps if current_action else []
            result = calculate_composite_motion(
                self.state_data.axes,
                self.state_data.actions,
                steps,
                self.state_data.fclk_hz,
            )
            for step_result in result.steps:
                move_text = " / ".join(
                    f"{move.axis_name}:{move.distance_note or _fmt(move.distance)} {_fmt(move.total_time)}s"
                    for move in step_result.move_results
                ) or "-"
                self.composite_result_tree.insert(
                    "",
                    tk.END,
                    iid=step_result.step.id,
                    values=(
                        step_result.step.name,
                        step_result.step.note,
                        move_text,
                        _fmt(step_result.delay_time),
                        _fmt(step_result.total_time),
                    ),
                )
            self.composite_total_var.set(f"组合总时间：{_fmt(result.total_time)} s")
            self.composite_status_var.set("")
            if current_action and hasattr(self, "composite_action_list") and current_action.id in self.composite_action_list.get_children():
                idx = next((i for i, item in enumerate(self.state_data.composite_actions, start=1) if item.id == current_action.id), "")
                number = current_action.number.strip() or str(idx)
                self.composite_action_list.item(current_action.id, values=(number, current_action.name, _fmt(result.total_time)))
        except Exception as exc:
            self.composite_total_var.set("")
            self.composite_status_var.set(str(exc))

    def _build_target_tab(self) -> None:
        fields = [
            ("distance",   "运动距离 mm",            "500"),
            ("xtarget",    "X_TARGET",               ""),
            ("time",       "目标时间 s",              "2"),
            ("strategy",   "策略",                   "平衡策略"),
            ("vmax_reg",   "Vmax 寄存器上限",         ""),
            ("vmax_motor", "Vmax 电机端上限 RPM",     ""),
            ("acc_reg",    "AMAX 寄存器上限",         ""),
            ("acc_motor",  "AMAX 电机端上限 RPM/s",   ""),
            ("dec_reg",    "DMAX 寄存器上限",         ""),
            ("dec_motor",  "DMAX 电机端上限 RPM/s",   ""),
        ]
        _focus_binds = {
            "distance":   lambda _e: self._auto_sync_target_dist_from_mm(),
            "xtarget":    lambda _e: self._auto_sync_target_dist_from_reg(),
            "vmax_motor": lambda _e: self._auto_sync_target_vmax_from_motor(),
            "vmax_reg":   lambda _e: self._auto_sync_target_vmax_from_reg(),
            "acc_motor":  lambda _e: self._auto_sync_target_acc_from_motor(),
            "acc_reg":    lambda _e: self._auto_sync_target_acc_from_reg(),
            "dec_motor":  lambda _e: self._auto_sync_target_dec_from_motor(),
            "dec_reg":    lambda _e: self._auto_sync_target_dec_from_reg(),
        }
        for row, (key, label, default) in enumerate(fields):
            ttk.Label(self.target_tab, text=label).grid(row=row, column=0, sticky="w", pady=4)
            var = tk.StringVar(value=default)
            self.target_vars[key] = var
            if key == "strategy":
                widget = ttk.Combobox(self.target_tab, textvariable=var, values=list(STRATEGY_LABELS), state="readonly", width=24)
            else:
                widget = ttk.Entry(self.target_tab, textvariable=var, width=24)
            widget.grid(row=row, column=1, sticky="w", pady=4)
            if key in _focus_binds:
                widget.bind("<FocusOut>", _focus_binds[key])
        n = len(fields)
        ttk.Button(self.target_tab, text="开始反算", command=self.calculate_target_time).grid(row=n, column=0, columnspan=2, sticky="w", pady=10)
        ttk.Label(self.target_tab, text="六点斜坡低冲击 / 防失步 / 快速节拍策略（待开发）", foreground="#777").grid(row=n + 1, column=0, columnspan=3, sticky="w")
        self.target_result = tk.StringVar(value="")
        ttk.Label(self.target_tab, textvariable=self.target_result, justify="left").grid(row=n + 2, column=0, columnspan=4, sticky="w", pady=8)

    def _build_detail_panel(self) -> None:
        panel = ttk.Frame(self, padding=(14, 10))
        panel.grid(row=0, column=2, sticky="nsew")
        panel.columnconfigure(1, weight=1)
        ttk.Label(panel, text="计算详情", font=("Segoe UI", 14, "bold")).grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 2)
        )
        self.detail_axis_var = tk.StringVar(value="")
        ttk.Label(panel, textvariable=self.detail_axis_var, font=("Segoe UI", 11), foreground="#555").grid(
            row=1, column=0, columnspan=2, sticky="w", pady=(0, 6)
        )
        rows = [
            ("profile",         "曲线类型"),
            ("reached",         "是否达到 VMAX"),
            ("tvmax_reg",       "TVMAX 寄存器"),
            ("tvmax_ms",        "TVMAX ms"),
            None,
            ("vpeak_reg",       "Vpeak 寄存器"),
            ("vpeak_motor",     "Vpeak 电机速度"),
            ("vpeak_phys",      "Vpeak 物理速度"),
            None,
            ("total_time",      "总时间"),
            ("acc_time",        "  加速时间"),
            ("const_time",      "  匀速时间"),
            ("dec_time",        "  减速时间"),
            None,
            ("total_distance",  "总距离"),
            ("acc_distance",    "  加速距离"),
            ("const_distance",  "  匀速距离"),
            ("dec_distance",    "  减速距离"),
            None,
            ("message",         "提示"),
        ]
        for grid_row, item in enumerate(rows, start=2):
            if item is None:
                ttk.Separator(panel, orient="horizontal").grid(
                    row=grid_row, column=0, columnspan=2, sticky="ew", pady=5
                )
            else:
                key, name = item
                self.detail_vars[key] = tk.StringVar(value="")
                ttk.Label(panel, text=name, font=("Segoe UI", 11), foreground="#555").grid(
                    row=grid_row, column=0, sticky="w", padx=(0, 20), pady=3
                )
                lbl = ttk.Label(panel, textvariable=self.detail_vars[key], font=("Segoe UI", 11))
                lbl.grid(row=grid_row, column=1, sticky="w", pady=3)
                self.detail_labels[key] = lbl

    def _build_chart(self) -> None:
        self.chart = tk.Canvas(self, height=420, bg="white", highlightthickness=1, highlightbackground="#ccc")
        self.chart.grid(row=1, column=0, columnspan=3, sticky="nsew", padx=8, pady=(2, 8))

    def new_project(self) -> None:
        dialog = ProjectDialog(self)
        if not dialog.result:
            return
        self.state_data = AppState(project_name=dialog.result["name"], fclk_hz=dialog.result["fclk_hz"])
        self.project_path = ""
        self._build_main()

    def open_project(self) -> None:
        if self.state_data.axes or self.state_data.actions:
            if messagebox.askyesno("打开项目", "是否保存当前项目？"):
                self.save_project()
        path = filedialog.askopenfilename(filetypes=[("Project", "*.json")])
        if path:
            self.state_data = load_project(path)
            self.project_path = path
            self._build_main()

    def save_project(self) -> None:
        if hasattr(self, "composite_vars"):
            self.save_all_composite_steps()
        if not self.project_path:
            path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("Project", "*.json")])
            if not path:
                return
            self.project_path = path
        save_project(self.project_path, self.state_data)

    def export_project_excel(self) -> None:
        if hasattr(self, "composite_vars"):
            self.save_all_composite_steps()
        default_name = f"{self.state_data.project_name or 'MotionProject'}.xlsx"
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel Workbook", "*.xlsx")],
        )
        if not path:
            return
        try:
            export_excel(path, self.state_data)
        except Exception as exc:
            messagebox.showerror("导出 Excel", f"导出失败：{exc}")
            return
        messagebox.showinfo("导出 Excel", f"导出完成：\n{path}")

    def refresh_all(self) -> None:
        self.title(f"HM Motion Calculator - {self.state_data.project_name}")
        self.refresh_axes()
        self.ensure_current_axis_action()
        self.load_current_action_params()
        self.refresh_calculation_views()

    def refresh_axes(self) -> None:
        self.axis_list.delete(0, tk.END)
        for axis in self.state_data.axes:
            self.axis_list.insert(tk.END, axis.name)
        if self.state_data.axes and not self.current_axis_id:
            self.current_axis_id = self.state_data.axes[0].id
        for i, axis in enumerate(self.state_data.axes):
            if axis.id == self.current_axis_id:
                self.axis_list.selection_set(i)
                self.axis_list.see(i)
                break
        self.update_axis_info()

    def ensure_current_axis_action(self) -> None:
        axis = self.get_current_axis()
        if not axis:
            self.current_action_id = ""
            return
        action = next((item for item in self.state_data.actions if item.axis_id == axis.id), None)
        if not action:
            action = MotionAction(name=f"{axis.name} 参数", axis_id=axis.id)
            self.state_data.actions.append(action)
        self.current_action_id = action.id

    def add_axis(self) -> None:
        dialog = AxisDialog(self)
        if dialog.result:
            self.state_data.axes.append(dialog.result)
            self.current_axis_id = dialog.result.id
            action = MotionAction(name=f"{dialog.result.name} 参数", axis_id=dialog.result.id)
            self.state_data.actions.append(action)
            self.current_action_id = action.id
            self.refresh_all()

    def edit_axis(self) -> None:
        axis = self.get_current_axis()
        if not axis:
            return
        dialog = AxisDialog(self, axis)
        if dialog.result:
            idx = self.state_data.axes.index(axis)
            self.state_data.axes[idx] = dialog.result
            self.current_axis_id = dialog.result.id
            action = next((item for item in self.state_data.actions if item.axis_id == dialog.result.id), None)
            if action:
                action.name = f"{dialog.result.name} 参数"
            self.refresh_all()

    def delete_axis(self) -> None:
        axis = self.get_current_axis()
        if not axis:
            return
        self.state_data.axes = [item for item in self.state_data.axes if item.id != axis.id]
        self.state_data.actions = [item for item in self.state_data.actions if item.axis_id != axis.id]
        self.current_axis_id = self.state_data.axes[0].id if self.state_data.axes else ""
        self.current_action_id = ""
        self.refresh_all()

    def add_distance(self) -> None:
        action = self.get_current_action()
        axis = self.get_axis(action.axis_id) if action else None
        if not action or not axis:
            return
        distance = _float(self.distance_mm_var.get(), 0)
        if not self.distance_mm_var.get().strip() and self.distance_reg_var.get().strip():
            distance = register_to_distance(axis, _float(self.distance_reg_var.get(), 0))
        action.distances.append(DistanceCase(distance=distance, note=self.distance_note_var.get()))
        self.refresh_calculation_views()

    def delete_distance(self) -> None:
        action = self.get_current_action()
        if not action or not self.current_distance_id:
            return
        action.distances = [item for item in action.distances if item.id != self.current_distance_id]
        self.current_distance_id = action.distances[0].id if action.distances else ""
        self.refresh_calculation_views()

    def on_axis_select(self) -> None:
        selection = self.axis_list.curselection()
        if selection:
            self.current_axis_id = self.state_data.axes[selection[0]].id
            self.ensure_current_axis_action()
            self.load_current_action_params()
            self.refresh_calculation_views()
            self.update_axis_info()

    def on_distance_select(self) -> None:
        selection = self.distance_tree.selection()
        if selection:
            self.current_distance_id = selection[0]
            self.show_selected_distance_detail()

    def get_axis(self, axis_id: str) -> Axis | None:
        return next((axis for axis in self.state_data.axes if axis.id == axis_id), None)

    def get_current_axis(self) -> Axis | None:
        return self.get_axis(self.current_axis_id)

    def get_current_action(self) -> MotionAction | None:
        return next((action for action in self.state_data.actions if action.id == self.current_action_id), None)

    def update_axis_info(self) -> None:
        axis = self.get_current_axis()
        if not axis:
            self.axis_info.set("暂无运动轴")
            if hasattr(self, "center_axis_var"):
                self.center_axis_var.set("")
            if hasattr(self, "detail_axis_var"):
                self.detail_axis_var.set("")
            return
        if hasattr(self, "center_axis_var"):
            self.center_axis_var.set(axis.name)
        if hasattr(self, "detail_axis_var"):
            self.detail_axis_var.set(axis.name)
        try:
            derived = derive_axis(axis)
            self.axis_info.set(
                f"当前轴: {axis.name}\n"
                f"类型: {MECHANISM_NAMES.get(axis.normalized_type(), axis.normalized_type())}\n"
                f"Count: {_fmt(derived.count, 0)}\n"
                f"每圈位移: {_fmt(derived.unit_per_motor_rev)} {derived.unit_name}\n"
                f"microsteps/{derived.unit_name}: {_fmt(derived.microsteps_per_unit)}\n"
                f"项目主频: {_fmt(self.state_data.fclk_hz / 1_000_000, 1)} MHz"
            )
        except Exception as exc:
            self.axis_info.set(str(exc))

    def load_current_action_params(self) -> None:
        action = self.get_current_action()
        self._loading_action_params = True
        if not action:
            for var in list(self.param_phys_vars.values()) + list(self.param_motor_vars.values()) + list(self.param_reg_vars.values()):
                var.set("")
            self._loading_action_params = False
            return
        for key, var in self.param_phys_vars.items():
            value = getattr(action.params, key)
            var.set(_fmt(value) if (key in ("vstart", "vstop") or value != 0.0) else "")
        self.sync_all_registers()
        self._loading_action_params = False

    def save_current_action_params(self) -> None:
        action = self.get_current_action()
        if not action:
            return
        action.params = self._params_from_inputs(action.params)

    def _params_from_inputs(self, fallback: MotionParams) -> MotionParams:
        return MotionParams(**{
            key: _float(var.get(), getattr(fallback, key))
            for key, var in self.param_phys_vars.items()
        })

    def sync_param_from_physical(self, key: str) -> None:
        action = self.get_current_action()
        axis = self.get_axis(action.axis_id) if action else None
        if not action or not axis:
            return
        phys_str = self.param_phys_vars[key].get().strip()
        if not phys_str and key not in ("vstart", "vstop"):
            self.param_reg_vars[key].set("")
            self.param_motor_vars[key].set("")
            return
        value = _float(phys_str, 0)
        if key in TIME_FIELDS:
            reg = tvmax_to_register(self.state_data.fclk_hz, value)
            motor_value = value
        elif key in VELOCITY_FIELDS:
            reg = velocity_to_register(axis, self.state_data.fclk_hz, value)
            motor_value = velocity_to_motor_rpm(axis, value)
        else:
            reg = acceleration_to_register(axis, self.state_data.fclk_hz, value)
            motor_value = acceleration_to_motor_rpm_s(axis, value)
        self.param_reg_vars[key].set(str(reg.value))
        self.param_motor_vars[key].set(_fmt(motor_value))

    def sync_param_from_motor(self, key: str) -> None:
        action = self.get_current_action()
        axis = self.get_axis(action.axis_id) if action else None
        if not action or not axis:
            return
        motor_value = _float(self.param_motor_vars[key].get(), 0)
        if key in TIME_FIELDS:
            value = motor_value
            reg = tvmax_to_register(self.state_data.fclk_hz, value)
        elif key in VELOCITY_FIELDS:
            value = motor_rpm_to_velocity(axis, motor_value)
            reg = velocity_to_register(axis, self.state_data.fclk_hz, value)
        else:
            value = motor_rpm_s_to_acceleration(axis, motor_value)
            reg = acceleration_to_register(axis, self.state_data.fclk_hz, value)
        self.param_phys_vars[key].set(_fmt(value))
        self.param_reg_vars[key].set(str(reg.value))

    def sync_param_from_register(self, key: str) -> None:
        action = self.get_current_action()
        axis = self.get_axis(action.axis_id) if action else None
        if not action or not axis:
            return
        reg = _float(self.param_reg_vars[key].get(), 0)
        if key in TIME_FIELDS:
            value = register_to_tvmax_ms(self.state_data.fclk_hz, reg)
            motor_value = value
        elif key in VELOCITY_FIELDS:
            value = register_to_velocity(axis, self.state_data.fclk_hz, reg)
            motor_value = register_to_motor_rpm(axis, self.state_data.fclk_hz, reg)
        else:
            value = register_to_acceleration(axis, self.state_data.fclk_hz, reg)
            motor_value = register_to_motor_rpm_s(axis, self.state_data.fclk_hz, reg)
        self.param_phys_vars[key].set(_fmt(value))
        self.param_motor_vars[key].set(_fmt(motor_value))

    def sync_all_registers(self) -> None:
        for key in FIELD_LABELS:
            self.sync_param_from_physical(key)

    def _update_reg_hint(self, key: str) -> None:
        try:
            v = int(float(self.param_reg_vars[key].get()))
            if v >= 1_000_000:
                self.param_reg_hint_vars[key].set("百万")
            elif v >= 100_000:
                self.param_reg_hint_vars[key].set("十万")
            elif v >= 10_000:
                self.param_reg_hint_vars[key].set("万")
            elif v >= 1_000:
                self.param_reg_hint_vars[key].set("千")
            else:
                self.param_reg_hint_vars[key].set("")
        except (ValueError, TypeError):
            self.param_reg_hint_vars[key].set("")
        label = self.param_reg_hint_labels.get(key)
        if label:
            if self.param_reg_hint_vars[key].get():
                label.place(in_=self.param_reg_entries[key], relx=1.0, x=-5, rely=0.5, anchor="e")
            else:
                label.place_forget()

    def _hide_reg_hint(self, key: str) -> None:
        self.param_reg_hint_vars[key].set("")
        label = self.param_reg_hint_labels.get(key)
        if label:
            label.place_forget()

    def _on_reg_focus_out(self, key: str) -> None:
        self.sync_param_from_register(key)
        self._hide_reg_hint(key)

    def _on_reg_enter(self, key: str) -> str:
        self.sync_param_from_register(key)
        self._hide_reg_hint(key)
        fields = list(FIELD_LABELS)
        idx = fields.index(key)
        if idx + 1 < len(fields):
            next_entry = self.param_reg_entries.get(fields[idx + 1])
            if next_entry:
                next_entry.focus_set()
                next_entry.selection_range(0, tk.END)
        return "break"

    def sync_distance_register(self) -> None:
        axis = self.get_axis(self.get_current_action().axis_id) if self.get_current_action() else self.get_current_axis()
        if not axis:
            return
        if self.distance_mm_var.get().strip():
            reg = distance_to_register(axis, _float(self.distance_mm_var.get(), 0))
            self.distance_reg_var.set(str(reg.value))
        elif self.distance_reg_var.get().strip():
            self.distance_mm_var.set(_fmt(register_to_distance(axis, _float(self.distance_reg_var.get(), 0))))

    def _auto_sync_distance_from_mm(self) -> None:
        axis = self.get_axis(self.get_current_action().axis_id) if self.get_current_action() else self.get_current_axis()
        if not axis or not self.distance_mm_var.get().strip():
            return
        try:
            self.distance_reg_var.set(str(distance_to_register(axis, _float(self.distance_mm_var.get(), 0)).value))
        except Exception:
            pass

    def _auto_sync_distance_from_reg(self) -> None:
        axis = self.get_axis(self.get_current_action().axis_id) if self.get_current_action() else self.get_current_axis()
        if not axis or not self.distance_reg_var.get().strip():
            return
        try:
            self.distance_mm_var.set(_fmt(register_to_distance(axis, _float(self.distance_reg_var.get(), 0))))
        except Exception:
            pass

    def _composite_references_action(self, composite: CompositeAction, action: MotionAction) -> bool:
        distance_ids = {distance.id for distance in action.distances}
        if not distance_ids:
            return False
        for step in composite.steps:
            if step.action1_distance_id in distance_ids:
                return True
            if step.action2_distance_id in distance_ids:
                return True
            if step.action3_distance_id in distance_ids:
                return True
        return False

    def _safe_composite_total(self, composite: CompositeAction) -> float | None:
        try:
            return calculate_composite_motion(
                self.state_data.axes,
                self.state_data.actions,
                composite.steps,
                self.state_data.fclk_hz,
            ).total_time
        except Exception:
            return None

    def _composite_time_changes(self, action: MotionAction, new_params: MotionParams) -> list[tuple[str, float | None, float | None]]:
        related = [item for item in self.state_data.composite_actions if self._composite_references_action(item, action)]
        if not related:
            return []
        old_params = action.params
        before = {item.id: self._safe_composite_total(item) for item in related}
        action.params = new_params
        try:
            after = {item.id: self._safe_composite_total(item) for item in related}
        finally:
            action.params = old_params
        changes: list[tuple[str, float | None, float | None]] = []
        for item in related:
            before_total = before[item.id]
            after_total = after[item.id]
            if before_total is None or after_total is None:
                if before_total != after_total:
                    changes.append((item.name, before_total, after_total))
            elif abs(before_total - after_total) > 1e-9:
                changes.append((item.name, before_total, after_total))
        return changes

    def _format_optional_time(self, value: float | None) -> str:
        return "计算错误" if value is None else _fmt(value)

    def _confirm_composite_time_changes(self, changes: list[tuple[str, float | None, float | None]]) -> bool:
        dialog = tk.Toplevel(self)
        dialog.title("确认重新计算")
        dialog.transient(self)
        dialog.grab_set()
        dialog.geometry("620x360")
        dialog.columnconfigure(0, weight=1)
        dialog.rowconfigure(1, weight=1)
        ttk.Label(dialog, text="以下组合动作总时间将变化，确认后写入当前单轴参数。").grid(
            row=0, column=0, sticky="w", padx=12, pady=(12, 8)
        )
        columns = ("name", "before", "after", "delta")
        tree = ttk.Treeview(dialog, columns=columns, show="headings", height=8)
        headers = {
            "name": "组合动作名称",
            "before": "修改前总时间 s",
            "after": "修改后总时间 s",
            "delta": "差值 s",
        }
        for col in columns:
            tree.heading(col, text=headers[col])
            tree.column(col, width=150, anchor="center")
        tree.grid(row=1, column=0, sticky="nsew", padx=12)
        for name, before, after in changes:
            if before is None or after is None:
                delta = ""
            else:
                delta = f"{after - before:+.3f}"
            tree.insert("", tk.END, values=(name, self._format_optional_time(before), self._format_optional_time(after), delta))
        result = tk.BooleanVar(value=False)
        buttons = ttk.Frame(dialog)
        buttons.grid(row=2, column=0, sticky="e", padx=12, pady=12)
        ttk.Button(buttons, text="取消", command=dialog.destroy).pack(side="right", padx=(8, 0))

        def confirm() -> None:
            result.set(True)
            dialog.destroy()

        ttk.Button(buttons, text="确认修改", command=confirm).pack(side="right")
        dialog.wait_window()
        return result.get()

    def refresh_calculation_views(self) -> None:
        self.refresh_distance_table()
        self.refresh_composite_tab()

    def recalculate_all(self) -> None:
        action = self.get_current_action()
        if action:
            old_params = action.params
            new_params = self._params_from_inputs(old_params)
            changes = self._composite_time_changes(action, new_params)
            if changes and not self._confirm_composite_time_changes(changes):
                self.load_current_action_params()
                self.refresh_distance_table()
                self.refresh_composite_tab()
                return
            action.params = new_params
        self.refresh_calculation_views()

    def refresh_distance_table(self) -> None:
        for item in self.distance_tree.get_children():
            self.distance_tree.delete(item)
        action = self.get_current_action()
        axis = self.get_axis(action.axis_id) if action else None
        if not action or not axis:
            return
        for distance in action.distances:
            try:
                result = calculate_motion(axis, self.state_data.fclk_hz, action.params, distance)
                xt = distance_to_register(axis, distance.distance).value
                values = (_fmt(distance.distance), xt, _fmt(result.total_time), distance.note)
            except Exception as exc:
                values = (_fmt(distance.distance), "", str(exc), distance.note)
            self.distance_tree.insert("", tk.END, iid=distance.id, values=values)
        if action.distances and not self.current_distance_id:
            self.current_distance_id = action.distances[0].id
        if self.current_distance_id in self.distance_tree.get_children():
            self.distance_tree.selection_set(self.current_distance_id)
            self.show_selected_distance_detail()

    def show_selected_distance_detail(self) -> None:
        action = self.get_current_action()
        axis = self.get_axis(action.axis_id) if action else None
        distance = next((item for item in action.distances if item.id == self.current_distance_id), None) if action else None
        if not action or not axis or not distance:
            return
        try:
            result = calculate_motion(axis, self.state_data.fclk_hz, action.params, distance)
            self.show_detail(result, axis)
            self.draw_chart(result)
        except Exception as exc:
            self.detail_vars["message"].set(str(exc))

    def show_detail(self, result: MotionResult, axis: Axis) -> None:
        unit = derive_axis(axis).unit_name
        vpeak_reg = velocity_to_register(axis, self.state_data.fclk_hz, result.vpeak)
        vpeak_motor = velocity_to_motor_rpm(axis, result.vpeak)
        total_distance = result.acc_distance + result.const_distance + result.dec_distance
        self.detail_vars["profile"].set(result.profile_type)
        self.detail_vars["reached"].set("是" if result.reached_vmax else "否")
        self.detail_vars["tvmax_reg"].set(str(result.tvmax_register.value))
        self.detail_vars["tvmax_ms"].set(f"{_fmt(register_to_tvmax_ms(self.state_data.fclk_hz, result.tvmax_register.value))} ms")
        self.detail_vars["vpeak_reg"].set(str(vpeak_reg.value))
        self.detail_vars["vpeak_motor"].set(f"{_fmt(vpeak_motor)} RPM")
        self.detail_vars["vpeak_phys"].set(f"{_fmt(result.vpeak)} {unit}/s")
        self.detail_vars["total_time"].set(f"{_fmt(result.total_time)} s")
        self.detail_vars["total_distance"].set(f"{_fmt(total_distance)} {unit}")
        self.detail_vars["acc_time"].set(f"{_fmt(result.acc_time)} s")
        self.detail_vars["const_time"].set(f"{_fmt(result.const_time)} s")
        self.detail_vars["dec_time"].set(f"{_fmt(result.dec_time)} s")
        self.detail_vars["acc_distance"].set(f"{_fmt(result.acc_distance)} {unit}")
        self.detail_vars["const_distance"].set(f"{_fmt(result.const_distance)} {unit}")
        self.detail_vars["dec_distance"].set(f"{_fmt(result.dec_distance)} {unit}")
        self.detail_vars["message"].set("")

    def draw_chart(self, result: MotionResult) -> None:
        canvas = self.chart
        canvas.delete("all")
        w = max(canvas.winfo_width(), 400)
        h = max(canvas.winfo_height(), 150)
        pad = 46
        total = max(result.total_time, 1e-9)
        vmax = max(result.vpeak, 1e-9)

        # Store chart state for interactive marker
        self._chart_result = result
        self._chart_w = w
        self._chart_h = h
        self._chart_pad = pad
        self._chart_total = total
        self._chart_vmax = vmax
        action = self.get_current_action()
        axis = self.get_axis(action.axis_id) if action else self.get_current_axis()
        try:
            self._chart_unit = derive_axis(axis).unit_name if axis else "mm"
        except Exception:
            self._chart_unit = "mm"
        self._marker_x = float(pad)  # reset to t=0 on each new chart

        def x(t: float) -> float:
            return pad + (w - 2 * pad) * t / total

        def y(v: float) -> float:
            return h - pad - (h - 2 * pad) * v / vmax

        canvas.create_line(pad, h - pad, w - pad, h - pad, fill="#999")
        canvas.create_line(pad, pad, pad, h - pad, fill="#999")
        points: list[float] = []
        for i, seg in enumerate(result.segments):
            if i == 0:
                points.extend([x(seg.start_time), y(seg.start_velocity)])
            points.extend([x(seg.end_time), y(seg.end_velocity)])
            canvas.create_line(x(seg.end_time), pad, x(seg.end_time), h - pad, fill="#ddd", dash=(2, 3))
        if len(points) >= 4:
            canvas.create_line(*points, fill="#1f6feb", width=3)
        params_for_chart = action.params if action else None
        if self._last_target_result is result and self._last_target_params:
            params_for_chart = self._last_target_params
        for velocity, label, with_value in [
            (params_for_chart.v1 if params_for_chart else 0.0, "V1", False),
            (params_for_chart.v2 if params_for_chart else 0.0, "V2", False),
            (result.vpeak, "Vpeak", True),
        ]:
            if velocity > 0 and velocity <= vmax * 1.05:
                yy = y(velocity)
                canvas.create_line(pad, yy, w - pad, yy, fill="#e7a400", dash=(4, 3))
                if with_value and axis:
                    reg_val = velocity_to_register(axis, self.state_data.fclk_hz, velocity).value
                    text = f"{label} {reg_val}"
                else:
                    text = label
                canvas.create_text(w - pad - 4, yy - 9, anchor="e", text=text, font=("Segoe UI", 10))
        canvas.create_text(pad, 18, anchor="w", text="速度-时间曲线", font=("Segoe UI", 12, "bold"))
        canvas.create_text(w - pad, h - 18, anchor="e", text=f"总时间 {_fmt(result.total_time)} s", font=("Segoe UI", 11))

        self._redraw_marker()
        canvas.bind("<ButtonPress-1>", self._on_chart_press)
        canvas.bind("<B1-Motion>", self._on_chart_drag)

    def _interp_velocity(self, t: float, result: MotionResult) -> float:
        for seg in result.segments:
            if seg.start_time <= t <= seg.end_time:
                dt = seg.end_time - seg.start_time
                if dt < 1e-12:
                    return seg.start_velocity
                return seg.start_velocity + (t - seg.start_time) / dt * (seg.end_velocity - seg.start_velocity)
        if not result.segments:
            return 0.0
        return result.segments[0].start_velocity if t <= 0 else result.segments[-1].end_velocity

    def _redraw_marker(self) -> None:
        canvas = self.chart
        canvas.delete("marker")
        if not self._chart_result:
            return
        pad = self._chart_pad
        w = self._chart_w
        h = self._chart_h
        total = self._chart_total
        vmax = self._chart_vmax
        mx = max(float(pad), min(float(w - pad), self._marker_x))
        t = (mx - pad) / max(w - 2 * pad, 1) * total
        t = max(0.0, min(total, t))
        v = self._interp_velocity(t, self._chart_result)
        y_c = h - pad - (h - 2 * pad) * v / vmax if vmax > 0 else float(h - pad)
        canvas.create_line(mx, pad, mx, h - pad, fill="#e03030", width=2, dash=(5, 3), tags="marker")
        canvas.create_oval(mx - 5, y_c - 5, mx + 5, y_c + 5,
                           fill="#e03030", outline="white", width=2, tags="marker")
        anchor = "w" if mx < w - pad - 170 else "e"
        lx = mx + 10 if anchor == "w" else mx - 10
        canvas.create_text(
            lx, max(float(pad + 14), y_c - 18), anchor=anchor,
            text=f"t={_fmt(t)} s    v={_fmt(v)} {self._chart_unit}/s",
            font=("Segoe UI", 10, "bold"), fill="#e03030", tags="marker",
        )

    def _on_chart_press(self, event: tk.Event) -> None:
        if not self._chart_result:
            return
        self._marker_x = max(float(self._chart_pad), min(float(self._chart_w - self._chart_pad), float(event.x)))
        self._redraw_marker()

    def _on_chart_drag(self, event: tk.Event) -> None:
        if not self._chart_result:
            return
        self._marker_x = max(float(self._chart_pad), min(float(self._chart_w - self._chart_pad), float(event.x)))
        self._redraw_marker()

    def _get_target_axis(self) -> Axis | None:
        return self.get_current_axis() or (self.state_data.axes[0] if self.state_data.axes else None)

    def _auto_sync_target_dist_from_mm(self) -> None:
        axis = self._get_target_axis()
        if not axis or not self.target_vars["distance"].get().strip():
            return
        try:
            self.target_vars["xtarget"].set(str(distance_to_register(axis, _float(self.target_vars["distance"].get(), 0)).value))
        except Exception:
            pass

    def _auto_sync_target_dist_from_reg(self) -> None:
        axis = self._get_target_axis()
        if not axis or not self.target_vars["xtarget"].get().strip():
            return
        try:
            self.target_vars["distance"].set(_fmt(register_to_distance(axis, _float(self.target_vars["xtarget"].get(), 0))))
        except Exception:
            pass

    def _auto_sync_target_vmax_from_motor(self) -> None:
        axis = self._get_target_axis()
        if not axis or not self.target_vars["vmax_motor"].get().strip():
            return
        try:
            phys = motor_rpm_to_velocity(axis, _float(self.target_vars["vmax_motor"].get(), 0))
            self.target_vars["vmax_reg"].set(str(velocity_to_register(axis, self.state_data.fclk_hz, phys).value))
        except Exception:
            pass

    def _auto_sync_target_vmax_from_reg(self) -> None:
        axis = self._get_target_axis()
        if not axis or not self.target_vars["vmax_reg"].get().strip():
            return
        try:
            self.target_vars["vmax_motor"].set(_fmt(register_to_motor_rpm(axis, self.state_data.fclk_hz, _float(self.target_vars["vmax_reg"].get(), 0))))
        except Exception:
            pass

    def _auto_sync_target_acc_from_motor(self) -> None:
        axis = self._get_target_axis()
        if not axis or not self.target_vars["acc_motor"].get().strip():
            return
        try:
            phys = motor_rpm_s_to_acceleration(axis, _float(self.target_vars["acc_motor"].get(), 0))
            self.target_vars["acc_reg"].set(str(acceleration_to_register(axis, self.state_data.fclk_hz, phys).value))
        except Exception:
            pass

    def _auto_sync_target_acc_from_reg(self) -> None:
        axis = self._get_target_axis()
        if not axis or not self.target_vars["acc_reg"].get().strip():
            return
        try:
            self.target_vars["acc_motor"].set(_fmt(register_to_motor_rpm_s(axis, self.state_data.fclk_hz, _float(self.target_vars["acc_reg"].get(), 0))))
        except Exception:
            pass

    def _auto_sync_target_dec_from_motor(self) -> None:
        axis = self._get_target_axis()
        if not axis or not self.target_vars["dec_motor"].get().strip():
            return
        try:
            phys = motor_rpm_s_to_acceleration(axis, _float(self.target_vars["dec_motor"].get(), 0))
            self.target_vars["dec_reg"].set(str(acceleration_to_register(axis, self.state_data.fclk_hz, phys).value))
        except Exception:
            pass

    def _auto_sync_target_dec_from_reg(self) -> None:
        axis = self._get_target_axis()
        if not axis or not self.target_vars["dec_reg"].get().strip():
            return
        try:
            self.target_vars["dec_motor"].set(_fmt(register_to_motor_rpm_s(axis, self.state_data.fclk_hz, _float(self.target_vars["dec_reg"].get(), 0))))
        except Exception:
            pass

    def sync_target_distance(self) -> None:
        self._auto_sync_target_dist_from_mm()

    def sync_target_limits(self) -> None:
        self._auto_sync_target_vmax_from_motor()
        self._auto_sync_target_acc_from_motor()
        self._auto_sync_target_dec_from_motor()

    def calculate_target_time(self) -> None:
        axis = self._get_target_axis()
        if not axis:
            self.target_result.set("请先新增运动轴。")
            return
        try:
            vmax_motor = _float(self.target_vars["vmax_motor"].get(), 0)
            acc_motor = _float(self.target_vars["acc_motor"].get(), 0)
            dec_motor = _float(self.target_vars["dec_motor"].get(), 0)
            result = recommend_for_target_time(
                axis,
                _float(self.target_vars["distance"].get(), 0),
                _float(self.target_vars["time"].get(), 0),
                STRATEGY_LABELS[self.target_vars["strategy"].get()],
                motor_rpm_to_velocity(axis, vmax_motor) if vmax_motor else None,
                motor_rpm_s_to_acceleration(axis, acc_motor) if acc_motor else None,
                motor_rpm_s_to_acceleration(axis, dec_motor) if dec_motor else None,
                self.state_data.fclk_hz,
            )
            vmax_reg = velocity_to_register(axis, self.state_data.fclk_hz, result.vmax_mm_s).value
            amax_reg = acceleration_to_register(axis, self.state_data.fclk_hz, result.acc_mm_s2).value
            dmax_reg = acceleration_to_register(axis, self.state_data.fclk_hz, result.dec_mm_s2).value
            self.target_result.set(
                f"策略: {self.target_vars['strategy'].get()}\n"
                f"VMAX: {_fmt(result.vmax_mm_s)}，寄存器值: {vmax_reg}\n"
                f"AMAX: {_fmt(result.acc_mm_s2)}，寄存器值: {amax_reg}\n"
                f"DMAX: {_fmt(result.dec_mm_s2)}，寄存器值: {dmax_reg}\n"
                f"总时间: {_fmt(result.forward.total_time)} s，曲线: {result.forward.profile_type}"
            )
            self.show_detail(result.forward, axis)
            self.draw_chart(result.forward)
        except Exception as exc:
            self.target_result.set(str(exc))

    def _build_target_tab(self) -> None:
        self.target_tab.columnconfigure(0, weight=0, minsize=360)
        self.target_tab.columnconfigure(1, weight=1)
        self.target_tab.rowconfigure(0, weight=1)
        left = ttk.Frame(self.target_tab, padding=(4, 2))
        right = ttk.LabelFrame(self.target_tab, text="推荐参数", padding=10)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(1, weight=1)

        fields = [
            ("distance", "运动距离 mm", "500"),
            ("xtarget", "X_TARGET", ""),
            ("time", "目标时间 s", "2"),
            ("strategy", "策略", "平衡（梯形）"),
            ("vmax_reg", "VMAX 上限", "2000000"),
            ("vmax_motor", "VMAX 电机端 RPM", ""),
            ("ad_reg", "A/D 上限", "200000"),
        ]
        focus_binds = {
            "distance": lambda _e: self._auto_sync_target_dist_from_mm(),
            "xtarget": lambda _e: self._auto_sync_target_dist_from_reg(),
            "vmax_reg": lambda _e: self._auto_sync_target_vmax_from_reg(),
            "vmax_motor": lambda _e: self._auto_sync_target_vmax_from_motor(),
        }
        for row, (key, label, default) in enumerate(fields):
            ttk.Label(left, text=label).grid(row=row, column=0, sticky="w", pady=4)
            var = tk.StringVar(value=default)
            self.target_vars[key] = var
            if key == "strategy":
                widget = ttk.Combobox(left, textvariable=var, values=list(STRATEGY_LABELS), state="readonly", width=24)
                widget.bind("<<ComboboxSelected>>", lambda _e: self._on_target_strategy_change())
            else:
                widget = ttk.Entry(left, textvariable=var, width=24)
            widget.grid(row=row, column=1, sticky="ew", pady=4)
            if key in focus_binds:
                widget.bind("<FocusOut>", focus_binds[key])
        left.columnconfigure(1, weight=1)

        self.target_accel_ratio_var = tk.DoubleVar(value=25)
        self.target_decel_ratio_var = tk.DoubleVar(value=25)
        ratio_start = len(fields)
        ttk.Label(left, text="加速时间占比").grid(row=ratio_start, column=0, sticky="w", pady=(12, 4))
        self.target_accel_ratio_label = ttk.Label(left, text="25%")
        self.target_accel_ratio_label.grid(row=ratio_start, column=1, sticky="e", pady=(12, 4))
        tk.Scale(left, from_=5, to=45, orient="horizontal", resolution=1, showvalue=False, variable=self.target_accel_ratio_var, command=lambda _v: self._on_target_ratio_change()).grid(row=ratio_start + 1, column=0, columnspan=2, sticky="ew")
        ttk.Label(left, text="减速时间占比").grid(row=ratio_start + 2, column=0, sticky="w", pady=(10, 4))
        self.target_decel_ratio_label = ttk.Label(left, text="25%")
        self.target_decel_ratio_label.grid(row=ratio_start + 2, column=1, sticky="e", pady=(10, 4))
        tk.Scale(left, from_=5, to=45, orient="horizontal", resolution=1, showvalue=False, variable=self.target_decel_ratio_var, command=lambda _v: self._on_target_ratio_change()).grid(row=ratio_start + 3, column=0, columnspan=2, sticky="ew")

        ttk.Button(left, text="开始计算", command=self.calculate_target_time).grid(row=ratio_start + 4, column=0, sticky="w", pady=12)

        result_rows = [
            ("status", "状态"),
            ("summary", "概要"),
            ("vmax", "VMAX"),
            ("v1", "V1"),
            ("v2", "V2"),
            ("a1", "A1"),
            ("a2", "A2"),
            ("amax", "AMAX"),
            ("dmax", "DMAX"),
            ("d2", "D2"),
            ("d1", "D1"),
            ("tvmax", "TVMAX"),
        ]
        for row, (key, label) in enumerate(result_rows):
            ttk.Label(right, text=label, foreground="#555").grid(row=row, column=0, sticky="w", padx=(0, 12), pady=3)
            var = tk.StringVar(value="")
            self.target_result_vars[key] = var
            ttk.Label(right, textvariable=var, justify="left", wraplength=520).grid(row=row, column=1, sticky="w", pady=3)
        self.target_apply_button = ttk.Button(right, text="应用参数", command=self.apply_target_params, state="disabled")
        self.target_apply_button.grid(row=len(result_rows), column=0, columnspan=2, sticky="w", pady=(12, 0))
        self._on_target_strategy_change(recalculate=False)
        self._auto_sync_target_vmax_from_reg()

    def _on_target_strategy_change(self, recalculate: bool = True) -> None:
        strategy = STRATEGY_LABELS.get(self.target_vars.get("strategy", tk.StringVar(value="平衡（梯形）")).get(), "balanced")
        accel, decel = STRATEGY_DEFAULT_RATIOS[strategy]
        self.target_accel_ratio_var.set(accel)
        self.target_decel_ratio_var.set(decel)
        self._update_target_ratio_labels()
        if recalculate and self._target_has_calculated:
            self.calculate_target_time()

    def _on_target_ratio_change(self) -> None:
        self._update_target_ratio_labels()
        if self._target_has_calculated:
            self.calculate_target_time()

    def _update_target_ratio_labels(self) -> None:
        self.target_accel_ratio_label.configure(text=f"{int(self.target_accel_ratio_var.get())}%")
        self.target_decel_ratio_label.configure(text=f"{int(self.target_decel_ratio_var.get())}%")

    def _target_set_status(self, message: str, ok: bool = False) -> None:
        self.target_result_vars["status"].set(message)
        if self.target_apply_button:
            self.target_apply_button.configure(state="normal" if ok else "disabled")

    def _format_target_param(self, key: str, value: float, axis: Axis) -> str:
        if key in VELOCITY_FIELDS:
            reg = velocity_to_register(axis, self.state_data.fclk_hz, value).value
            motor = velocity_to_motor_rpm(axis, value)
            unit = derive_axis(axis).unit_name
            return f"寄存器 {reg} / 电机 {_fmt(motor)} RPM / 物理 {_fmt(value)} {unit}/s"
        if key in TIME_FIELDS:
            reg = tvmax_to_register(self.state_data.fclk_hz, value).value
            return f"寄存器 {reg} / {_fmt(value)} ms"
        reg = acceleration_to_register(axis, self.state_data.fclk_hz, value).value
        unit = derive_axis(axis).unit_name
        return f"寄存器 {reg} / 物理 {_fmt(value)} {unit}/s²"

    def _show_target_result(self, result) -> None:
        axis = self._get_target_axis()
        if not axis:
            return
        params = result.params
        self.target_result_vars["summary"].set(
            f"总时间 {_fmt(result.forward.total_time)} s / 误差 {_fmt(result.time_error_s)} s / 曲线 {result.forward.profile_type}"
        )
        for key in ("vmax", "v1", "v2", "a1", "a2", "amax", "dmax", "d2", "d1", "tvmax_ms"):
            result_key = "tvmax" if key == "tvmax_ms" else key
            self.target_result_vars[result_key].set(self._format_target_param(key, getattr(params, key), axis))

    def calculate_target_time(self) -> None:
        axis = self._get_target_axis()
        if not axis:
            self._target_set_status("请先新增运动轴。")
            return
        try:
            self._auto_sync_target_dist_from_mm()
            self._auto_sync_target_vmax_from_reg()
            vmax_limit = register_to_velocity(axis, self.state_data.fclk_hz, _float(self.target_vars["vmax_reg"].get(), 0))
            ad_limit = register_to_acceleration(axis, self.state_data.fclk_hz, _float(self.target_vars["ad_reg"].get(), 0))
            result = recommend_for_target_time(
                axis,
                _float(self.target_vars["distance"].get(), 0),
                _float(self.target_vars["time"].get(), 0),
                STRATEGY_LABELS[self.target_vars["strategy"].get()],
                vmax_limit,
                ad_limit,
                ad_limit,
                self.state_data.fclk_hz,
                self.target_accel_ratio_var.get() / 100.0,
                self.target_decel_ratio_var.get() / 100.0,
            )
            self._last_target_params = result.params
            self._last_target_result = result.forward
            self._target_has_calculated = True
            self._show_target_result(result)
            self._target_set_status("计算完成", ok=True)
            self.show_detail(result.forward, axis)
            self.draw_chart(result.forward)
        except Exception as exc:
            self._target_has_calculated = True
            self._target_set_status(str(exc), ok=False)

    def apply_target_params(self) -> None:
        action = self.get_current_action()
        if not action or not self._last_target_params:
            self._target_set_status("没有可应用的推荐参数。", ok=False)
            return
        action.params = self._last_target_params
        self.load_current_action_params()
        self.refresh_distance_table()
        self._target_set_status("已应用到当前动作", ok=True)


def run_app() -> None:
    app = MotionCalculatorApp()
    app.mainloop()
